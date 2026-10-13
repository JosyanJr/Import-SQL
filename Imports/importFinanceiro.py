import pandas as pd
import os
import sys
from datetime import datetime


# ================= Utilitários =================

def obter_diretorio_base():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    else:
        return os.path.dirname(os.path.abspath(__file__))


def formatar_data(data):
    try:
        dt = pd.to_datetime(data, errors='coerce', dayfirst=True)
        if pd.notna(dt):
            return f"'{dt.strftime('%Y-%d-%m')}'"
        return "NULL"
    except Exception:
        return "NULL"


def limpar_texto(valor, limite=None):
    if valor is None:
        return "NULL"
    s = str(valor).strip()
    if s.upper() in ('NULL', 'NAN', ''):
        return "NULL"
    texto = s.replace("'", "''")
    if limite:
        texto = texto[:limite]
    return f"'{texto}'"


def formatar_formapgt(valor):
    if valor is None:
        return "NULL"
    s = str(valor).strip()
    if s.upper() in ('NULL', 'NAN', ''):
        return "NULL"
    try:
        return f"'{int(float(s))}'"
    except (ValueError, TypeError):
        return f"'{s}'"


def _safe_num(valor, default=0):
    """Float seguro; retorna default se inválido."""
    if valor is None:
        return default
    s = str(valor).strip()
    if s.upper() in ('NULL', 'NAN', ''):
        return default
    try:
        return float(s)
    except (ValueError, TypeError):
        return default


def _safe_int_sql(valor):
    """Inteiro SQL ou NULL string."""
    if valor is None:
        return "NULL"
    s = str(valor).strip()
    if s.upper() in ('NULL', 'NAN', ''):
        return "NULL"
    try:
        return str(int(float(s)))
    except (ValueError, TypeError):
        return "NULL"


# ================= Gerador de INSERT =================

# Mapeamento de forma de pagamento: nome (uppercase) → (coluna SQL, id)
# Prazo: inserido em pgtTipoPrazo
# Vista: inserido em pgtTipoVista
_FORMAS_PRAZO = {
    'BOLETO':        3,
    'CARTEIRA':      2,
    'CREDITO':       0,
    'CRÉDITO':       0,
    'CHEQUE':        1,
    'DEBITO CONTA':  6,
    'DÉBITO CONTA':  6,
}
_FORMAS_VISTA = {
    'DINHEIRO':  0,
    'PIX':       4,
    'DEBITO':    2,
    'DÉBITO':    2,
    'DEPOSITO':  3,
    'DEPÓSITO':  3,
}


def resolver_forma_pgto(valor):
    """
    Recebe o texto da planilha (ex: 'Boleto', 'Pix') e retorna
    (coluna_sql, id_sql) ou (None, None) se não reconhecido.
    """
    if not valor or str(valor).strip().upper() in ('', 'NULL', 'NAN'):
        return None, None

    chave = str(valor).strip().upper()

    if chave in _FORMAS_PRAZO:
        return 'pgtTipoPrazo', _FORMAS_PRAZO[chave]
    if chave in _FORMAS_VISTA:
        return 'pgtTipoVista', _FORMAS_VISTA[chave]

    return None, None   # valor desconhecido → coluna ignorada


def gerar_insert_venda_pgto(row, tipo_conta, empid=1):
    """
    row        : dict com chaves = nomes das colunas SQL, valores = dados da linha
    tipo_conta : 'P' (a pagar) | 'R' (a receber)
    empid      : int – ID da empresa (fixo para o lote)

    Forma de pagamento:
    - Lê row['formapgt'] como texto (ex: 'Boleto', 'Pix')
    - Resolve automaticamente a coluna (pgtTipoPrazo ou pgtTipoVista) e o ID
    - pgtClienteId inicia NULL; UPDATE ao final resolve pelo nome
    """
    # pgtPago derivado de pgtDataQuitou
    data_q = row.get('pgtDataQuitou', '')
    pago   = "'N'" if (not data_q or str(data_q).strip().upper() in ('', 'NULL', 'NAN')) else "'S'"

    # Resolve forma de pagamento pelo nome
    col_pgto, id_pgto = resolver_forma_pgto(row.get('formapgt'))
    val_pgto = str(id_pgto) if id_pgto is not None else 'NULL'
    if col_pgto is None:
        col_pgto = 'pgtTipoPrazo'   # fallback seguro

    columns_values = {
        'empId':            empid,
        'pgtTipoOpr':       "'FI'",
        'pgtModoReceb':     "'F'",
        'pgtTipoConta':     f"'{tipo_conta}'",
        'pgtPago':          pago,
        col_pgto:           val_pgto,
        'pgtClienteId':     "NULL",   # resolvido via UPDATE
        'pgtValor':         _safe_num(row.get('pgtValor')),
        'pgtValorJuros':    _safe_num(row.get('pgtValorJuros')),
        'pgtValorDescFin':  _safe_num(row.get('pgtValorDescFin')),
        'pgtValorMulta':    _safe_num(row.get('pgtValorMulta')),
        'pgtBancoRef':      _safe_int_sql(row.get('pgtBancoRef')),
        'pgtCecId':         _safe_int_sql(row.get('pgtCecId')),
        'pgtData':          formatar_data(row.get('pgtData')),
        'pgtVecmto':        formatar_data(row.get('pgtVecmto')),
        'pgtDataQuitou':    formatar_data(row.get('pgtDataQuitou')),
        'pgtObs':           limpar_texto(row.get('pgtObs')),
        'pgtCliNome':       limpar_texto(row.get('pgtCliNome'), 40),
        'temp_pgtPlc':      limpar_texto(row.get('temp_pgtPlc')),
        'temp_pgtSubPc':    limpar_texto(row.get('temp_pgtSubPc')),
        'pgtPlcId':         "NULL",
        'pgtSubPc':         "NULL",
    }

    columns = ', '.join(columns_values.keys())
    values  = ', '.join(map(str, columns_values.values()))
    return f"INSERT INTO vendaPgto({columns}) VALUES({values}) GO"


# ================= Função principal =================

def gerar_sql_inserts(df, mapeamento, tipo_conta, empid=1):
    """
    df         : DataFrame já carregado do Excel
    mapeamento : dict { campo_sql: coluna_excel }
                 coluna_excel pode ser '' se o usuário não mapeou
    tipo_conta : 'P' ou 'R'
    """
    inserts = []

    # ---- DDL colunas temporárias ----
    inserts.append(
        "IF COL_LENGTH('vendaPgto', 'temp_pgtPlc') IS NULL\n"
        "BEGIN\n"
        "    ALTER TABLE vendaPgto ADD temp_pgtPlc VARCHAR(255)\n"
        "END\n"
        "GO"
    )
    inserts.append(
        "IF COL_LENGTH('vendaPgto', 'temp_pgtSubPc') IS NULL\n"
        "BEGIN\n"
        "    ALTER TABLE vendaPgto ADD temp_pgtSubPc VARCHAR(255)\n"
        "END\n"
        "GO"
    )

    # ---- INSERTs por linha ----
    for _, linha in df.iterrows():
        row = {}
        for campo_sql, col_excel in mapeamento.items():
            if col_excel and col_excel in df.columns:
                row[campo_sql] = linha[col_excel]
            else:
                row[campo_sql] = None

        row.pop('empId', None)  # empid vem do parâmetro fixo, não do Excel
        insert = gerar_insert_venda_pgto(row, tipo_conta, empid)

        # Normaliza nan/None residuais
        insert = (
            insert
            .replace("'nan'",  "NULL")
            .replace("'None'", "NULL")
            .replace(" nan,",  " NULL,")
            .replace(" nan)",  " NULL)")
        )
        inserts.append(insert)

    # ---- UPDATEs de resolução ----
    inserts.append(
        "UPDATE vendaPgto\n"
        "SET pgtClienteId = cliente.cliId\n"
        "FROM vendaPgto\n"
        "JOIN cliente ON LEFT(cliente.cliNome, 40) = LEFT(vendaPgto.pgtCliNome, 40)\n"
        "WHERE vendaPgto.pgtClienteId IS NULL\n"
        "GO"
    )
    inserts.append(
        "UPDATE vendaPgto\n"
        "SET pgtPlcId = planoConta.plcId\n"
        "FROM vendaPgto\n"
        "JOIN planoConta\n"
        "    ON UPPER(LTRIM(RTRIM(planoConta.plcDesc)))\n"
        "     = UPPER(LTRIM(RTRIM(vendaPgto.temp_pgtPlc)))\n"
        "WHERE vendaPgto.pgtPlcId IS NULL\n"
        "GO"
    )
    inserts.append(
        "UPDATE vendaPgto\n"
        "SET pgtSubPc = subPlanoContas.spcId\n"
        "FROM vendaPgto\n"
        "JOIN subPlanoContas\n"
        "    ON UPPER(LTRIM(RTRIM(subPlanoContas.spcDesc)))\n"
        "     = UPPER(LTRIM(RTRIM(vendaPgto.temp_pgtSubPc)))\n"
        "WHERE vendaPgto.pgtSubPc IS NULL\n"
        "GO"
    )

    # ---- Drop colunas temporárias ----
    inserts.append(
        "IF COL_LENGTH('vendaPgto', 'temp_pgtPlc') IS NOT NULL\n"
        "BEGIN\n"
        "    ALTER TABLE vendaPgto DROP COLUMN temp_pgtPlc\n"
        "END\n"
        "GO"
    )
    inserts.append(
        "IF COL_LENGTH('vendaPgto', 'temp_pgtSubPc') IS NOT NULL\n"
        "BEGIN\n"
        "    ALTER TABLE vendaPgto DROP COLUMN temp_pgtSubPc\n"
        "END\n"
        "GO"
    )

    return inserts


# ================= I/O =================

def salvar_sql_inserts(arquivo, inserts):
    with open(arquivo, 'w', encoding='utf-8') as f:
        for insert in inserts:
            f.write(insert + '\n')


def carregar_excel(caminho, sheet_name=None, fill_value='NULL'):
    if not os.path.exists(caminho):
        raise FileNotFoundError(f"Arquivo Excel '{caminho}' não encontrado.")
    try:
        df = pd.read_excel(caminho, sheet_name=sheet_name)
        df = df.fillna(fill_value)
        return df
    except Exception as e:
        raise Exception(f"Erro ao carregar o arquivo Excel: {e}")
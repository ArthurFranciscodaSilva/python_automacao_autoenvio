import pandas as pd
from openpyxl import load_workbook
from pathlib import Path

# ==== CONFIGURAÇÕES ====

# Diretório onde estão os arquivos baixados do AutoEnvio
DIRETORIO_AUTOENVIO = Path(r"C:\Users\Arthur\Downloads\AutoEnvio\03.09.25")

# Caminho do arquivo Excel (base) no OneDrive (via sincronização local)
EXCEL_ONEDRIVE = Path(r"C:\Users\Arthur\OneDrive\Escritorio_CWB_Sao_Francisco\TETO\teto_2025 - Copia.xlsx")

# Nome da aba da planilha base
NOME_ABA = "higiTeto"

# Nome da coluna que serve de chave para localizar a linha
COLUNA_CHAVE_BASE = "CPF"  # onde procurar o apelido

# Nome da coluna onde os números devem ser inseridos
COLUNA_DESTINO = "NÃO TEM WHATSAPP"

# Nome do arquivo de relatório
RELATORIO_NAO_ENCONTRADOS = Path(r"C:\Users\Arthur\Downloads\AutoEnvio\03.09.25")

# ==== COLETAR DADOS DOS ARQUIVOS AUTOENVIO ====

linhas_coletadas = []

for arquivo in DIRETORIO_AUTOENVIO.glob("*.xlsx"):
    try:
        df = pd.read_excel(arquivo, dtype=str)
    except Exception as e:
        print(f"Erro ao ler {arquivo.name}: {e}")
        continue

    filtro = df["Erro"] == "O contato não é um número de WhatsApp"
    for _, linha in df.loc[filtro].iterrows():
        apelido = str(linha.get("Apelido", "")).strip()
        numero = str(linha.get("Número", "")).strip()
        if apelido and numero:
            linhas_coletadas.append((apelido, numero))

print(f"Total de registros coletados: {len(linhas_coletadas)}")

# ==== ABRIR PLANILHA BASE ====

wb = load_workbook(EXCEL_ONEDRIVE)
if NOME_ABA not in wb.sheetnames:
    raise ValueError(f"Aba '{NOME_ABA}' não encontrada na planilha base.")

ws = wb[NOME_ABA]

# Mapear as colunas pelo cabeçalho
colunas = {cell.value: cell.column for cell in ws[1]}

if COLUNA_CHAVE_BASE not in colunas:
    raise ValueError(f"Coluna chave '{COLUNA_CHAVE_BASE}' não encontrada.")
if COLUNA_DESTINO not in colunas:
    raise ValueError(f"Coluna destino '{COLUNA_DESTINO}' não encontrada.")

col_chave = colunas[COLUNA_CHAVE_BASE]
col_dest = colunas[COLUNA_DESTINO]

# Construir um mapa de apelido -> linha
mapa_chave_para_linha = {}
for linha_idx in range(2, ws.max_row + 1):
    valor_chave = ws.cell(row=linha_idx, column=col_chave).value
    if valor_chave is not None:
        mapa_chave_para_linha[str(valor_chave).strip()] = linha_idx

# ==== PROCESSAR CADA REGISTRO ====

nao_encontrados = []

for apelido, numero in linhas_coletadas:
    if apelido in mapa_chave_para_linha:
        linha_idx = mapa_chave_para_linha[apelido]
        ws.cell(row=linha_idx, column=col_dest).value = numero
    else:
        nao_encontrados.append({"Apelido não encontrado": apelido})

# ==== SALVAR ALTERAÇÕES ====

wb.save(EXCEL_ONEDRIVE)
print("Planilha base atualizada com sucesso!")

# ==== GERAR RELATÓRIO DOS NÃO ENCONTRADOS ====

if nao_encontrados:
    df_relatorio = pd.DataFrame(nao_encontrados)
    df_relatorio.to_excel(RELATORIO_NAO_ENCONTRADOS, index=False)
    print(f"Relatório gerado: {RELATORIO_NAO_ENCONTRADOS}")
else:
    print("Todos os apelidos foram encontrados na planilha base.")

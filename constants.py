import os

BG         = "#F7F8FC"
SIDEBAR    = "#16213E"
SIDE_SEL   = "#0F3460"
SIDE_HOV   = "#1A2E52"
SIDE_TXT   = "#CBD5E1"
CARD       = "#FFFFFF"
BORDER     = "#E2E8F0"
BORDER2    = "#CBD5E1"
ACCENT     = "#4F46E5"
ACCENT_H   = "#4338CA"
ACCENT_L   = "#EEF2FF"
TEXT       = "#0F172A"
TEXT2      = "#64748B"
TEXT3      = "#94A3B8"
SUCCESS    = "#059669"
SUCCESS_L  = "#ECFDF5"
DANGER     = "#DC2626"
WARN       = "#D97706"
WHITE      = "#FFFFFF"

F8   = ("Segoe UI",  8)
F9   = ("Segoe UI",  9)
F9B  = ("Segoe UI",  9, "bold")
F10  = ("Segoe UI", 10)
F10B = ("Segoe UI", 10, "bold")
F11B = ("Segoe UI", 11, "bold")
F13B = ("Segoe UI", 13, "bold")

DRAFT_FILE = os.path.join(os.path.expanduser("~"), ".contrato_draft.json")

MONTHS = [
    "", "janeiro", "fevereiro", "março", "abril", "maio", "junho",
    "julho", "agosto", "setembro", "outubro", "novembro", "dezembro",
]

GUIDE = """\
=== CONTRATANTE ===
{{NOME}}              → Nome do Contratante
{{NATURALIDADE}}      → Naturalidade
{{DATA_NASC}}         → Data de Nascimento
{{RG}}                → RG
{{CPF}}               → CPF (000.000.000-00)
{{CEP}}               → CEP
{{ENDERECO}}          → Logradouro (auto CEP)
{{NUMERO}}            → Número
{{COMPLEMENTO}}       → Complemento
{{BAIRRO}}            → Bairro (auto CEP)
{{CIDADE}}            → Cidade (auto CEP)
{{UF}}                → UF (auto CEP)
{{TELEFONE1}}         → Telefone 1
{{TELEFONE2}}         → Telefone 2
{{TELEFONES}}         → Tel 1 / Tel 2 (combinados)
{{EMAIL}}             → E-mail

=== CURSO ===
{{CURSO}}             → Curso Escolhido
{{DIAS_HORARIO}}      → Dias e Horário
{{INICIO_1}}          → Previsão de Início 1
{{INICIO_2}}          → Previsão de Início 2
{{INICIO_3}}          → Previsão de Início 3

=== VALORES ===
{{VALOR_TOTAL}}       → Valor Total do Curso
{{PARCELA_NORMAL}}    → Parcela Normal
{{PARCELA_DESC}}      → Parcela com Desconto
{{DESCONTO_PCT}}      → Desconto (%)
{{QTD_PARCELAS}}      → Quantidade de Parcelas
{{DIA_VENCIMENTO}}    → Dia de Vencimento
{{VALOR_MATRICULA}}   → Valor da Matrícula
{{DATA_1PAG}}         → Data do 1º Pagamento
{{MES_REF}}           → Mês de Referência
{{DATA_VENC_SEM_DESC}}→ Data Venc. s/ Desconto
{{FORMA_PAGAMENTO}}   → Forma de Pagamento
{{OBS}}               → Observações

=== ESCOLA ===
{{ESCOLA_NOME}}       → Nome da Escola
{{ESCOLA_CNPJ}}       → CNPJ
{{ESCOLA_END}}        → Endereço
{{ESCOLA_MUNICIPIO}}  → Município/UF
{{ESCOLA_TEL}}        → Telefone
{{ESCOLA_EMAIL}}      → E-mail
{{ESCOLA_DIRETOR}}    → Diretor(a)
{{ESCOLA_COORD}}      → Coordenador(a)

=== RESPONSÁVEL FINANCEIRO ===
{{RESP_NOME}}         → Nome
{{RESP_PARENT}}       → Grau de Parentesco
{{RESP_ESTADO_CIVIL}} → Estado Civil
{{RESP_NASC}}         → Data de Nascimento
{{RESP_RG}}           → RG
{{RESP_CPF}}          → CPF
{{RESP_CEP}}          → CEP
{{RESP_END}}          → Logradouro
{{RESP_NUMERO}}       → Número
{{RESP_COMP}}         → Complemento
{{RESP_BAIRRO}}       → Bairro
{{RESP_CIDADE}}       → Cidade
{{RESP_UF}}           → UF
{{RESP_TEL}}          → Telefone
{{RESP_EMAIL}}        → E-mail
"""

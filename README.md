# Gerador de Contrato de Matrícula

Aplicativo desktop para gerar contratos de matrícula em Word e PDF a partir de um modelo `.docx`. Preencha os dados do aluno, curso e escola, e o app substitui os placeholders automaticamente.

---

## Funcionalidades

- Busca automática de endereço via **CEP** (ViaCEP)
- Validação de **CPF** com dígito verificador em tempo real
- Formatação automática de telefone e campos monetários
- Cadastro de cursos com importação de planilha Excel
- Opção de copiar dados do contratante para o responsável financeiro
- Salvamento de rascunho automático em JSON
- Geração de **Word** (.docx) e **PDF**

---

## Estrutura do Projeto

```
contrato_app/
├── app.py          # Ponto de entrada
├── constants.py    # Cores, fontes e constantes globais
├── utils.py        # Validação, formatação e consulta de CEP
├── document.py     # Substituição de placeholders no Word e conversão para PDF
├── widgets.py      # Componentes de UI reutilizáveis (ScrollableFrame, Section, LabeledEntry, Btn)
└── contrato.py     # Classe principal ContratoApp
```

---

## Dependências

```bash
pip install python-docx requests openpyxl
```

Para geração de PDF (escolha uma):

```bash
pip install comtypes    # Windows com Microsoft Word instalado
# ou
pip install docx2pdf
```

---

## Como Executar

```bash
python app.py
```

---

## Gerar Executável (.exe)

```bash
py -m PyInstaller GeradorContrato.spec
```

> O Windows pode exibir aviso de segurança. Clique em **"Mais informações" → "Executar assim mesmo"**. É um falso positivo do PyInstaller.

---

## Planilha de Cursos

A planilha `.xlsx` deve conter os nomes dos cursos na **coluna A** e o horário padrão (opcional) na **coluna B**. O app ignora cabeçalhos automaticamente.

| A                         | B                      |
|---------------------------|------------------------|
| Técnico em Informática    | Seg/Qua/Sex 19h–22h    |
| Técnico em Administração  | Ter/Qui 18h30–21h30    |
| Técnico em Enfermagem     |                        |

---

## Placeholders para o Modelo Word

Use os placeholders abaixo no seu arquivo `.docx`. O app os substituirá pelos valores preenchidos no formulário.

### Contratante

| Placeholder        | Campo                    |
|--------------------|--------------------------|
| `{{NOME}}`         | Nome completo            |
| `{{NATURALIDADE}}` | Naturalidade             |
| `{{DATA_NASC}}`    | Data de nascimento       |
| `{{RG}}`           | RG                       |
| `{{CPF}}`          | CPF (000.000.000-00)     |
| `{{CEP}}`          | CEP                      |
| `{{ENDERECO}}`     | Logradouro (via CEP)     |
| `{{NUMERO}}`       | Número                   |
| `{{COMPLEMENTO}}`  | Complemento              |
| `{{BAIRRO}}`       | Bairro (via CEP)         |
| `{{CIDADE}}`       | Cidade (via CEP)         |
| `{{UF}}`           | UF (via CEP)             |
| `{{TELEFONE1}}`    | Telefone 1               |
| `{{TELEFONE2}}`    | Telefone 2               |
| `{{TELEFONES}}`    | Tel 1 / Tel 2 combinados |
| `{{EMAIL}}`        | E-mail                   |

### Curso

| Placeholder        | Campo                   |
|--------------------|-------------------------|
| `{{CURSO}}`        | Curso escolhido         |
| `{{DIAS_HORARIO}}` | Dias e horário          |
| `{{INICIO_1}}`     | Previsão de início 1    |
| `{{INICIO_2}}`     | Previsão de início 2    |
| `{{INICIO_3}}`     | Previsão de início 3    |

### Valores Financeiros

| Placeholder              | Campo                      |
|--------------------------|----------------------------|
| `{{VALOR_TOTAL}}`        | Valor total do curso       |
| `{{PARCELA_NORMAL}}`     | Parcela normal             |
| `{{PARCELA_DESC}}`       | Parcela com desconto       |
| `{{DESCONTO_PCT}}`       | Desconto (%)               |
| `{{QTD_PARCELAS}}`       | Quantidade de parcelas     |
| `{{DIA_VENCIMENTO}}`     | Dia de vencimento          |
| `{{VALOR_MATRICULA}}`    | Valor da matrícula         |
| `{{DATA_1PAG}}`          | Data do 1º pagamento       |
| `{{MES_REF}}`            | Mês de referência          |
| `{{DATA_VENC_SEM_DESC}}` | Data de venc. s/ desconto  |
| `{{FORMA_PAGAMENTO}}`    | Forma de pagamento         |
| `{{OBS}}`                | Observações                |

### Escola

| Placeholder            | Campo          |
|------------------------|----------------|
| `{{ESCOLA_NOME}}`      | Nome           |
| `{{ESCOLA_CNPJ}}`      | CNPJ           |
| `{{ESCOLA_END}}`       | Endereço       |
| `{{ESCOLA_MUNICIPIO}}` | Município/UF   |
| `{{ESCOLA_TEL}}`       | Telefone       |
| `{{ESCOLA_EMAIL}}`     | E-mail         |
| `{{ESCOLA_DIRETOR}}`   | Diretor(a)     |
| `{{ESCOLA_COORD}}`     | Coordenador(a) |

### Responsável Financeiro

| Placeholder             | Campo               |
|-------------------------|---------------------|
| `{{RESP_NOME}}`         | Nome                |
| `{{RESP_PARENT}}`       | Grau de parentesco  |
| `{{RESP_ESTADO_CIVIL}}` | Estado civil        |
| `{{RESP_NASC}}`         | Data de nascimento  |
| `{{RESP_RG}}`           | RG                  |
| `{{RESP_CPF}}`          | CPF                 |
| `{{RESP_CEP}}`          | CEP                 |
| `{{RESP_END}}`          | Logradouro          |
| `{{RESP_NUMERO}}`       | Número              |
| `{{RESP_COMP}}`         | Complemento         |
| `{{RESP_BAIRRO}}`       | Bairro              |
| `{{RESP_CIDADE}}`       | Cidade              |
| `{{RESP_UF}}`           | UF                  |
| `{{RESP_TEL}}`          | Telefone            |
| `{{RESP_EMAIL}}`        | E-mail              |

---

## Problemas Comuns

| Problema                  | Solução                                              |
|---------------------------|------------------------------------------------------|
| CEP não preenche          | Verifique conexão e instale `requests`               |
| CPF sempre inválido       | Verifique se tem 11 dígitos e está correto           |
| Planilha não carrega      | `pip install openpyxl`                               |
| PDF não é gerado          | `pip install comtypes` ou `pip install docx2pdf`     |
| Antivírus bloqueia o .exe | Falso positivo do PyInstaller — adicione uma exceção |

---

## Tecnologias

- [Python](https://python.org) + [Tkinter](https://docs.python.org/3/library/tkinter.html)
- [python-docx](https://python-docx.readthedocs.io)
- [requests](https://requests.readthedocs.io) + [ViaCEP](https://viacep.com.br)
- [openpyxl](https://openpyxl.readthedocs.io)
- [PyInstaller](https://pyinstaller.org)

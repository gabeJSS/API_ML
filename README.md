# API_ML

Sistema de automação para extração, enriquecimento, classificação e integração de compras do Mercado Livre com processos internos de ERP.

O projeto foi desenvolvido para reduzir trabalho manual em rotinas administrativas e contábeis, automatizando desde a coleta dos pedidos até a preparação dos dados para lançamento em sistemas corporativos.

---

## Principais Funcionalidades

### Extração de Pedidos

* Autenticação OAuth2 com a API do Mercado Livre.
* Renovação automática de access tokens.
* Consulta de pedidos pagos por período.
* Captura de informações financeiras e operacionais dos pedidos.
* Identificação de parcelamentos, juros, descontos e fretes.
* Recuperação de dados complementares através do frontend do Mercado Livre.

### Download e Tratamento de NF-e

* Download automático de XMLs de notas fiscais.
* Armazenamento organizado dos arquivos XML.
* Leitura e processamento de informações fiscais.
* Extração de:

  * Chave da NF-e
  * CNPJ do fornecedor
  * Valor da nota
  * Data de emissão
  * Informações complementares
  * Produtos da nota

### Enriquecimento de Dados

* Complementação automática dos pedidos utilizando endpoints da API.
* Consolidação de informações em arquivos JSON estruturados.
* Sistema de cache para evitar consultas repetidas.

### Classificação Contábil

* Geração de planilha Excel para classificação manual.
* Associação de contas de resultado.
* Validação de pedidos não classificados.
* Geração de arquivo final pronto para integração.

### Automação ERP

* Interface gráfica desenvolvida em Tkinter.
* Processamento de XMLs fiscais.
* Integração com banco de dados via ODBC.
* Manipulação de planilhas Excel.
* Automação de tarefas operacionais.
* Auxílio a lançamentos e conferências fiscais.

---

## Estrutura do Projeto

```text
APi_ML/
│
├── script_ML/
│   ├── Extrator de Notas/
│   │   ├── auth_ml.py
│   │   ├── core.py
│   │   ├── config.py
│   │   └── main.py
│   │
│   └── Gerador_FullJSON/
│       └── estudo.py
│
├── script_ERP/
│   └── lançador.py
│
└── xmls_nfe/
```

---

## Tecnologias Utilizadas

* Python
* Requests
* Tkinter
* OpenPyXL
* Pandas
* PyODBC
* XML (ElementTree)
* OAuth2
* Mercado Livre API
* OpenCV
* PyAutoGUI

---

## Fluxo de Processamento

1. Autenticação no Mercado Livre.
2. Extração dos pedidos pagos.
3. Download dos XMLs fiscais.
4. Enriquecimento dos dados.
5. Geração da planilha de classificação.
6. Classificação contábil dos pedidos.
7. Geração do JSON final.
8. Integração com processos internos do ERP.

---

## Segurança

Credenciais, cookies de autenticação e tokens são armazenados localmente e não devem ser versionados.

Arquivos recomendados para inclusão no `.gitignore`:

```gitignore
auth.json
config.json
cookie.json
.env
__pycache__/
*.pyc
```

---

## Objetivo

Automatizar processos financeiros, fiscais e contábeis relacionados a compras realizadas através do Mercado Livre, reduzindo intervenção manual, tempo operacional e riscos de erro em integrações com ERP.

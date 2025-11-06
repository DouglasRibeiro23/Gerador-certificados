# Emissor de Certificados (PPTX) — Minicursos

Gera **certificados em PowerPoint** a partir de uma **planilha Excel** e de um **template** `.pptx` com placeholders.  
Projeto simples de usar e **fácil de testar** com `pytest`, seguindo as boas práticas do Cap. 8 de *Engenharia de Software Moderna* (FIRST, classes de equivalência, fronteiras, dublês de teste).

---

## ✨ Funcionalidades

- Lê `participantes.xlsx` com as colunas: **Nome**, **Capacitacao**, **Evento**, **Data** (`YYYY-MM-DD`), **Horas** (inteiro).
- Usa um template `Template capacitações.pptx` com placeholders:
  `{NOME}`, `{CAPACITACAO}`, `{EVENTO}`, `{DIA}`, `{MES}`, `{ANO}`, `{HORAS}`.
- Estiliza `{NOME}`, `{CAPACITACAO}`, `{EVENTO}` com **azul** RGB (0,112,192), **Arial Rounded MT Bold**, **16 pt**.
- Remove o slide modelo do template antes de salvar.
- Salva cada certificado em `./certificados/` com **nome sanitizado** (caracteres inválidos viram `_`).

---

## 📁 Estrutura do projeto

├─ CERTIFICADOS-MINICURSOS.py # Script principal (com hífen; testes carregam por caminho)
├─ participantes.xlsx # (exemplo, opcional)
├─ Template capacitações.pptx # (exemplo, opcional)
├─ certificados/ # Saída (criada em tempo de execução)
└─ tests/
   ├─ conftest.py # Fakes de python-pptx + fixtures
   └─ test_certificados.py # Testes automatizados (pytest)


**Observação**: o script tem hífen no nome. A suíte de testes já carrega o arquivo por **caminho absoluto** — não é necessário renomear.

---

## 🧩 Requisitos

- Python 3.10+
- Dependências:
  ```bash
  pip install -U pip
  pip install pandas python-pptx pytest pytest-cov

# SheetsForSaim

Projeto para captura de dados de propriedades (UI Streamlit ou API FastAPI), gravação em `contatos.xlsx` e geracao automatica de planilhas modelo em `Output/` com formulas dinamicas.

## Visao Geral

Fluxos disponiveis:

1. UI (Streamlit): usuario preenche formulario e salva.
2. API (FastAPI): cliente envia JSON no endpoint e salva.
3. Pipeline de processamento: pega o registro salvo, preenche `InputTemplate.xlsx` e gera arquivo final em `Output/`.

Atualmente, UI e API ja disparam o processamento automaticamente apos salvar um novo registro.

## Estrutura Principal

- `app.py`: interface Streamlit.
- `api.py`: API FastAPI.
- `main.py`: pipeline de input e geracao de arquivo em `Output/`.
- `functions.py`: motor de formulas dinamicas (Monthly CF, ajustes de Inputs, etc).
- `funcitons.py`: wrapper legado de compatibilidade para reaplicar formulas.
- `InputTemplate.xlsx`: template principal.
- `contatos.xlsx`: base de dados com registros de entrada.
- `Output/`: arquivos gerados.

## Requisitos

- Python 3.10+
- Pacotes Python:
  - `streamlit`
  - `pandas`
  - `openpyxl`
  - `fastapi`
  - `uvicorn`
  - `pydantic`

## Instalacao

No diretorio do projeto:

```bash
python -m venv venv
```

### Windows (PowerShell)

```powershell
.\venv\Scripts\Activate.ps1
```

Instale dependencias:

```bash
pip install streamlit pandas openpyxl fastapi uvicorn pydantic
```

## Como Rodar

### 1) Rodar a UI (Streamlit)

```bash
streamlit run app.py
```

- Abra a URL mostrada no terminal (normalmente `http://localhost:8501`).
- Ao salvar, o registro entra em `contatos.xlsx` e um arquivo novo e gerado em `Output/`.

### 2) Rodar a API (FastAPI)

```bash
python -m uvicorn api:app --reload
```

- API disponivel em `http://127.0.0.1:8000`
- Swagger UI: `http://127.0.0.1:8000/docs`

### 3) Rodar processamento manual (se necessario)

Processa o primeiro registro pendente (`Submitted = No`) em `contatos.xlsx`:

```bash
python main.py
```

### 4) Reaplicar formulas em arquivos da pasta Output

```bash
python funcitons.py --sheet "Monthly CF"
```

## Rotas da API

Base URL: `http://127.0.0.1:8000`

### GET /

Health simplificado com endpoints disponiveis.

Resposta exemplo:

```json
{
  "message": "SheetsForSaim API online",
  "endpoints": {
    "health": "GET /health",
    "create_property": "POST /properties",
    "create_property_compat": "POST /"
  }
}
```

### GET /health

Resposta exemplo:

```json
{
  "status": "ok"
}
```

### POST /properties

Cria um registro, salva em `contatos.xlsx`, processa automaticamente e retorna o caminho do arquivo gerado.

### POST /

Endpoint de compatibilidade. Aceita o mesmo payload de `POST /properties`.

## Payload Exemplo (POST /properties)

```json
{
  "property_name": "Teste API",
  "property_type": "Other",
  "address": "N/A",
  "city_and_state": "N/A",
  "number_of_units": 1,
  "purchase_price": 350000,
  "down_payment": 20,
  "due_diligence_costs": 0,
  "loan_original_costs": 0,
  "purchase_date": "2024-05-27",
  "end_year": 10,

  "incomes": {
    "gross_potential_rent": 809596,
    "vacancy_rate": 5,
    "credit_loss": 1,
    "other_incomes": [
      {"tipo": "Gift Shop and Vending Sales", "valor": "41368.12"}
    ]
  },

  "expenses": {
    "property_tax": 5623.86,
    "insurance": 0,
    "management_fee": 8,
    "repairs_and_maintenance": 7118.13,
    "utilities": 76686.92,
    "capital_expenditures": 0,
    "landscape_and_janitorial": 0,
    "capex_1": "0",
    "capex_2": "0",
    "capex_3": "0",
    "capex_4": "0",
    "capex_5": "0",
    "other_expenses": [
      {"tipo": "Advertising and Promotion", "valor": "6907.48"},
      {"tipo": "Automobile Expense", "valor": "651.66"}
    ]
  },

  "validation": {
    "is_valid": false,
    "missing_fields": ["address", "city_and_state"],
    "warnings": ["payload pre-validated externally"]
  }
}
```

Observacao: a API continua aceitando o formato antigo com campos no nivel raiz
(`gross_potential_rent`, `other_incomes`, `property_tax`, `other_expenses`, etc.).
Quando os blocos `incomes` e `expenses` estiverem presentes, eles passam a ter
prioridade no mapeamento para o `contatos.xlsx`.

Resposta exemplo:

```json
{
  "message": "Property saved successfully",
  "property_name": "Teste API",
  "total_records": 8,
  "output_file": "Output/Teste API_20260320_150000.xlsx"
}
```

## Notas Importantes

- Campos percentuais enviados na API/UI devem ser em formato percentual humano:
  - Exemplo: `20` para 20%
  - Exemplo: `1.5` para 1.5%
- O projeto contem logica para preservar compatibilidade com colunas antigas no `contatos.xlsx`.
- Se houver mudanca estrutural no `Inputs` (linhas adicionadas/removidas), as formulas sao reconstruidas dinamicamente no pipeline.

## Troubleshooting Rapido

- `404` no Postman:
  - Use `POST /properties` ou `POST /`.
- API nao sobe por falta de pacote:
  - Rode `pip install fastapi uvicorn pydantic`.
- Streamlit nao abre:
  - Verifique se o ambiente virtual esta ativo e rode `streamlit run app.py`.
- Sem arquivo novo em `Output/`:
  - Verifique se o registro foi salvo com `Submitted = No` e execute `python main.py`.

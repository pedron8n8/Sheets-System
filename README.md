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
  "property_type": "Multifamily",
  "address": "123 Main St",
  "city_and_state": "Miami, FL",
  "number_of_units": 12,
  "purchase_price": 300000,
  "down_payment": 20,
  "due_diligence_costs": 5000,
  "loan_original_costs": 1.5,
  "purchase_date": "2026-03-20",
  "end_year": 10,

  "gross_potential_rent": 72000,
  "vacancy_rate": 5,
  "credit_loss": 1,

  "property_tax": 9000,
  "insurance": 2500,
  "management_fee": 5,
  "repairs_and_maintenance": 4000,
  "utilities": 3500,
  "capital_expenditures": 3000,
  "landscape_and_janitorial": 1800,

  "capex_1": "Roof",
  "capex_2": "HVAC",
  "capex_3": "",
  "capex_4": "",
  "capex_5": "",

  "other_incomes": [
    {"tipo": "Other Income - Parking", "valor": "2400"}
  ],
  "other_expenses": [
    {"tipo": "Security / Access Control", "valor": "1200"}
  ]
}
```

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

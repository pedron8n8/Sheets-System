"""API HTTP para entrada de propriedades e geracao de planilhas.

Fluxo principal:
1) Recebe payload JSON da propriedade.
2) Converte para formato tabular do contatos.xlsx.
3) Persiste o registro no Excel de entrada.
4) Aciona o pipeline de main.py para gerar arquivo final na pasta Output.
5) Retorna caminho/URL do arquivo gerado.

Abas impactadas indiretamente pela API:
- Inputs, IncomeExpenses, Monthly CF, Quarterly CF, Annual CF, Summary e Equity Waterfall,
  pois o endpoint chama processar_registro_por_indice.
"""

from __future__ import annotations

from datetime import date
from pathlib import Path

import pandas as pd
from fastapi import FastAPI, HTTPException, Response
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel, Field
from main import processar_registro_por_indice

ARQUIVO_EXCEL = Path("contatos.xlsx")
OUTPUT_DIR = Path("Output")


class OtherIncomeItem(BaseModel):
    """Representa um item adicional de receita enviado pelo cliente da API."""
    tipo: str = Field(..., description="Name/Type of Income Additional")
    valor: str = Field(..., description="Value of Income Additional")


class OtherExpenseItem(BaseModel):
    """Representa um item adicional de despesa enviado pelo cliente da API."""
    tipo: str = Field(..., description="Name/Type of Additional Expense")
    valor: str = Field(..., description="Value of Additional Expense")


class PropertyPayload(BaseModel):
    """Contrato de entrada da API para criar uma propriedade."""
    property_name: str
    property_type: str
    address: str = ""
    city_and_state: str = ""
    number_of_units: int = 0
    purchase_price: float = 0.0
    down_payment: float = 0.0
    due_diligence_costs: float = 0.0
    loan_original_costs: float = 0.0
    purchase_date: date
    end_year: int = 10

    gross_potential_rent: float = 0.0
    vacancy_rate: float = 0.0
    credit_loss: float = 0.0

    property_tax: float = 0.0
    insurance: float = 0.0
    management_fee: float = 0.0
    repairs_and_maintenance: float = 0.0
    utilities: float = 0.0
    capital_expenditures: float = 0.0
    landscape_and_janitorial: float = 0.0

    capex_1: str = ""
    capex_2: str = ""
    capex_3: str = ""
    capex_4: str = ""
    capex_5: str = ""

    other_incomes: list[OtherIncomeItem] = []
    other_expenses: list[OtherExpenseItem] = []


def _montar_registro(payload: PropertyPayload) -> dict:
    """Converte payload HTTP para o schema de colunas do contatos.xlsx.

Por que faz:
- O pipeline de main.py espera nomes de colunas legados/especificos do Excel.
    """
    registro = {
        "Property Name": payload.property_name,
        "Property Type": payload.property_type,
        "Address": payload.address,
        "City and State": payload.city_and_state,
        "Number of Units": payload.number_of_units,
        "Purchase Price": payload.purchase_price,
        "Down Payment (%)": payload.down_payment,
        "Due Diligence Costs": payload.due_diligence_costs,
        "Loan Original Costs": payload.loan_original_costs,
        "Purchase Date": payload.purchase_date,
        "End Year": int(payload.end_year),
        "Gross Potential Rent": payload.gross_potential_rent,
        "Vacancy Rate %": payload.vacancy_rate,
        "Credit Loss %": payload.credit_loss,
        "Property Tax": payload.property_tax,
        "Insurance": payload.insurance,
        "Management Fee %": payload.management_fee,
        "Repairs and Maintenance": payload.repairs_and_maintenance,
        "Utilities": payload.utilities,
        "Capital Expenditures": payload.capital_expenditures,
        "Landscape and Janitorial": payload.landscape_and_janitorial,
        "CapEx 1": payload.capex_1,
        "CapEx 2": payload.capex_2,
        "CapEx 3": payload.capex_3,
        "CapEx 4": payload.capex_4,
        "CapEx 5": payload.capex_5,
        "Submitted": "No",
    }

    # Serializa lista dinamica de receitas no padrao Other Income N Type/Amount.
    income_count = 0
    for item in payload.other_incomes:
        label = str(item.tipo).strip()
        valor = str(item.valor).strip()
        if label and label != "Select..." and valor:
            income_count += 1
            registro[f"Other Income {income_count} Type"] = label
            registro[f"Other Income {income_count} Amount"] = valor

    # Serializa lista dinamica de despesas no padrao Other Expense N Type/Amount.
    expense_count = 0
    for item in payload.other_expenses:
        label = str(item.tipo).strip()
        valor = str(item.valor).strip()
        if label and label != "Select..." and valor:
            expense_count += 1
            registro[f"Other Expense {expense_count} Type"] = label
            registro[f"Other Expense {expense_count} Amount"] = valor

    registro["Other Income Count"] = income_count
    registro["Other Expense Count"] = expense_count

    return registro


def _salvar_registro_no_excel(registro: dict) -> tuple[int, int]:
    """Anexa registro ao contatos.xlsx e retorna (total_registros, novo_indice).

Compatibilidade tratada:
- Migra colunas legadas para nomes atuais quando necessario.
    """
    novos_dados = pd.DataFrame([registro])

    if ARQUIVO_EXCEL.exists():
        df_antigo = pd.read_excel(ARQUIVO_EXCEL)
        legacy_map = {
            "Down Payment": "Down Payment (%)",
            "Due Diligence Costs %": "Due Diligence Costs",
        }
        for old_col, new_col in legacy_map.items():
            if old_col in df_antigo.columns:
                if new_col in df_antigo.columns:
                    df_antigo[new_col] = df_antigo[new_col].combine_first(df_antigo[old_col])
                    df_antigo = df_antigo.drop(columns=[old_col])
                else:
                    df_antigo = df_antigo.rename(columns={old_col: new_col})
        if "Submitted" not in df_antigo.columns:
            df_antigo["Submitted"] = "No"
        df_final = pd.concat([df_antigo, novos_dados], ignore_index=True)
    else:
        df_final = novos_dados

    df_final.to_excel(ARQUIVO_EXCEL, index=False)
    novo_indice = len(df_final) - 1
    return len(df_final), novo_indice


app = FastAPI(title="SheetsForSaim API", version="1.0.0")

# Garante que /outputs possa servir arquivos mesmo em deploy limpo.
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
app.mount("/outputs", StaticFiles(directory=str(OUTPUT_DIR)), name="outputs")


@app.get("/")
def root() -> dict:
    """Endpoint de descoberta com status basico e mapa de endpoints."""
    return {
        "message": "SheetsForSaim API online",
        "endpoints": {
            "health": "GET /health",
            "create_property": "POST /properties",
            "create_property_compat": "POST /",
            "generated_files": "GET /outputs/{file_name}",
        },
    }


@app.head("/")
def root_head() -> Response:
    """HEAD de compatibilidade para health checks simples."""
    return Response(status_code=200)


@app.get("/health")
def health() -> dict:
    """Health endpoint para monitoramento da API."""
    return {"status": "ok"}


@app.post("/properties")
def create_property(payload: PropertyPayload) -> dict:
    """Cria propriedade, grava no contatos.xlsx e gera planilha final.

O endpoint retorna:
- total_records: quantidade total de registros no contatos.xlsx.
- output_file: caminho do .xlsx gerado na Output.
- output_url: rota estatica para download/consumo do arquivo gerado.
    """
    if not payload.property_name or not payload.property_type:
        raise HTTPException(status_code=400, detail="property_name e property_type sao obrigatorios")

    registro = _montar_registro(payload)
    total_registros, novo_indice = _salvar_registro_no_excel(registro)
    caminho_gerado = processar_registro_por_indice(novo_indice)
    output_name = Path(caminho_gerado).name if caminho_gerado else ""
    output_url = f"/outputs/{output_name}" if output_name else None

    return {
        "message": "Property saved successfully",
        "property_name": payload.property_name,
        "total_records": total_registros,
        "output_file": caminho_gerado,
        "output_url": output_url,
    }


@app.post("/")
def create_property_root(payload: PropertyPayload) -> dict:
    """Endpoint de compatibilidade para clientes que fazem POST na raiz."""
    return create_property(payload)

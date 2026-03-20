from __future__ import annotations

from datetime import date
from pathlib import Path

import pandas as pd
from fastapi import FastAPI, HTTPException
from pydantic import BaseModel, Field
from main import processar_registro_por_indice

ARQUIVO_EXCEL = Path("contatos.xlsx")


class OtherIncomeItem(BaseModel):
    tipo: str = Field(..., description="Nome/tipo da receita adicional")
    valor: str = Field(..., description="Valor da receita adicional")


class OtherExpenseItem(BaseModel):
    tipo: str = Field(..., description="Nome/tipo da despesa adicional")
    valor: str = Field(..., description="Valor da despesa adicional")


class PropertyPayload(BaseModel):
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

    income_count = 0
    for item in payload.other_incomes:
        label = str(item.tipo).strip()
        valor = str(item.valor).strip()
        if label and label != "Select..." and valor:
            income_count += 1
            registro[f"Other Income {income_count} Type"] = label
            registro[f"Other Income {income_count} Amount"] = valor

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


@app.get("/")
def root() -> dict:
    return {
        "message": "SheetsForSaim API online",
        "endpoints": {
            "health": "GET /health",
            "create_property": "POST /properties",
            "create_property_compat": "POST /",
        },
    }


@app.get("/health")
def health() -> dict:
    return {"status": "ok"}


@app.post("/properties")
def create_property(payload: PropertyPayload) -> dict:
    if not payload.property_name or not payload.property_type:
        raise HTTPException(status_code=400, detail="property_name e property_type sao obrigatorios")

    registro = _montar_registro(payload)
    total_registros, novo_indice = _salvar_registro_no_excel(registro)
    caminho_gerado = processar_registro_por_indice(novo_indice)

    return {
        "message": "Property saved successfully",
        "property_name": payload.property_name,
        "total_records": total_registros,
        "output_file": caminho_gerado,
    }


@app.post("/")
def create_property_root(payload: PropertyPayload) -> dict:
    """Compatibility endpoint for clients posting to root path."""
    return create_property(payload)

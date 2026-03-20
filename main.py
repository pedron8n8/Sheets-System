"""Input pipeline entrypoint.

This file owns all input-side behavior:
- Read pending rows from contatos.xlsx
- Populate Inputs sheet values
- Compact/remove unused rows for readability
- Call formula engine after input changes
"""

from __future__ import annotations

from datetime import datetime
import os
import re

import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

from functions import aplicar_formulas_apos_inputs

ARQUIVO_TEMPLATE = "InputTemplate.xlsx"
ARQUIVO_DATA = "contatos.xlsx"
PASTA_SAIDA = "Output"


def garantir_pasta_saida(pasta_saida: str = PASTA_SAIDA) -> None:
    if not os.path.exists(pasta_saida):
        os.makedirs(pasta_saida)


def _set_cell_value_respeitando_merge(sheet, cell_ref: str, value) -> None:
    for merged_range in sheet.merged_cells.ranges:
        if cell_ref in merged_range:
            anchor_col = get_column_letter(merged_range.min_col)
            anchor_ref = f"{anchor_col}{merged_range.min_row}"
            sheet[anchor_ref] = value
            return
    sheet[cell_ref] = value


def _normalizar_percentual(valor):
    if valor == "" or pd.isna(valor):
        return ""

    if isinstance(valor, str):
        texto = valor.strip().replace("%", "").replace(",", ".")
        if not texto:
            return ""
        try:
            numero = float(texto)
        except ValueError:
            return valor
    else:
        numero = valor

    if abs(numero) <= 1:
        return numero

    # Entrada em formato percentual comum (ex.: 1.5 => 1.5%).
    if abs(numero) <= 100:
        return numero / 100

    # Valores acima de 100% sao invalidados para evitar explosoes numericas no Excel.
    return 1.0 if numero > 0 else -1.0


def _normalizar_down_payment(valor_down_payment, valor_purchase_price):
    """Normalize down payment robustly for legacy and current schemas.

    Accepted inputs:
    - Decimal percent: 0.2
    - Percent number: 20
    - Legacy absolute amount: 200 with purchase price 800 -> 0.25
    """
    percentual = _normalizar_percentual(valor_down_payment)

    # Quando veio como valor absoluto legado (ex.: 200), _normalizar_percentual caparia em 100%.
    # Antes de usar o cap, tentamos inferir percentual com base no Purchase Price.
    raw = _to_float(valor_down_payment)
    purchase_price = _to_float(valor_purchase_price)
    if abs(raw) > 100 and purchase_price > 0:
        inferido = raw / purchase_price
        if inferido >= 0:
            return min(inferido, 1.0)

    return percentual


def _valor_limpo(valor):
    if pd.isna(valor):
        return ""
    return valor


def _to_float(valor) -> float:
    if valor is None or valor == "" or pd.isna(valor):
        return 0.0

    if isinstance(valor, (int, float)):
        return float(valor)

    texto = str(valor).strip().replace("$", "").replace(",", "")
    try:
        return float(texto)
    except ValueError:
        return 0.0


def _extrair_other_incomes(primeira_linha):
    other_incomes = []
    for coluna in primeira_linha.index:
        match = re.match(r"^Other Income\s+(\d+)\s+Type$", str(coluna))
        if not match:
            continue

        idx = match.group(1)
        tipo = _valor_limpo(primeira_linha.get(coluna, ""))
        valor = _valor_limpo(primeira_linha.get(f"Other Income {idx} Amount", ""))
        if str(tipo).strip():
            other_incomes.append((int(idx), str(tipo).strip(), valor))

    other_incomes.sort(key=lambda item: item[0])
    return other_incomes


def _extrair_other_expenses(primeira_linha):
    other_expenses = []
    for coluna in primeira_linha.index:
        match = re.match(r"^Other Expense\s+(\d+)\s+Type$", str(coluna))
        if not match:
            continue

        idx = match.group(1)
        tipo = _valor_limpo(primeira_linha.get(coluna, ""))
        valor = _valor_limpo(primeira_linha.get(f"Other Expense {idx} Amount", ""))
        if str(tipo).strip():
            other_expenses.append((int(idx), str(tipo).strip(), valor))

    other_expenses.sort(key=lambda item: item[0])
    return other_expenses


def _compactar_other_incomes(other_incomes, max_rows: int):
    if len(other_incomes) <= max_rows:
        return other_incomes

    base = other_incomes[: max_rows - 1]
    excedentes = other_incomes[max_rows - 1 :]
    total_excedentes = sum(_to_float(valor_income) for _, _, valor_income in excedentes)
    quantidade_excedentes = len(excedentes)
    base.append((9999, f"Other Income - Additional ({quantidade_excedentes} items)", total_excedentes))
    return base


def _compactar_other_expenses(other_expenses, max_rows: int):
    if len(other_expenses) <= max_rows:
        return other_expenses

    base = other_expenses[: max_rows - 1]
    excedentes = other_expenses[max_rows - 1 :]
    total_excedentes = sum(_to_float(valor_expense) for _, _, valor_expense in excedentes)
    quantidade_excedentes = len(excedentes)
    base.append((9999, f"Other Expense - Additional ({quantidade_excedentes} items)", total_excedentes))
    return base


def criar_arquivo_baseado_em_template(primeira_linha) -> str:
    """Populate Inputs sheet from one contatos row and then apply formula engine."""
    garantir_pasta_saida(PASTA_SAIDA)

    wb = load_workbook(ARQUIVO_TEMPLATE, data_only=False)
    ws = wb["Inputs"]

    input_property_type = primeira_linha.get("Property Type", "")
    input_property_name = primeira_linha.get("Property Name", "")
    input_address = primeira_linha.get("Address", "")
    input_city_state = primeira_linha.get("City and State", "")
    input_number_units = primeira_linha.get("Number of Units", 0)
    input_purchase_price = primeira_linha.get("Purchase Price", 0)
    input_down_payment = primeira_linha.get("Down Payment (%)", primeira_linha.get("Down Payment", 0))
    input_diligence = primeira_linha.get("Due Diligence Costs", primeira_linha.get("Due Diligence Costs %", 0))
    input_loan_costs = primeira_linha.get("Loan Original Costs", 0)
    input_purchase_date = primeira_linha.get("Purchase Date", "")
    input_end_year = primeira_linha.get("End Year", 10)

    input_gpr = primeira_linha.get("Gross Potential Rent", 0)
    input_vacancy = primeira_linha.get("Vacancy Rate %", 0)
    input_credit_loss = primeira_linha.get("Credit Loss %", 0)

    input_property_tax = primeira_linha.get("Property Tax", 0)
    input_insurance = primeira_linha.get("Insurance", 0)
    input_management_fee = primeira_linha.get("Management Fee %", 0)
    input_repairs = primeira_linha.get("Repairs and Maintenance", 0)
    input_utilities = primeira_linha.get("Utilities", 0)
    input_capex_base = primeira_linha.get("Capital Expenditures", 0)
    input_landscape = primeira_linha.get("Landscape and Janitorial", 0)

    input_capex1 = primeira_linha.get("CapEx 1", primeira_linha.get("  CapEx Item 1", ""))
    input_capex2 = primeira_linha.get("CapEx 2", primeira_linha.get("  CapEx Item 2", ""))
    input_capex3 = primeira_linha.get("CapEx 3", primeira_linha.get("  CapEx Item 3", ""))
    input_capex4 = primeira_linha.get("CapEx 4", primeira_linha.get("  CapEx Item 4", ""))
    input_capex5 = primeira_linha.get("CapEx 5", primeira_linha.get("  CapEx Item 5", ""))
    input_capex6 = primeira_linha.get("CapEx 6", primeira_linha.get("  CapEx Item 6", ""))

    other_incomes = _extrair_other_incomes(primeira_linha)
    other_expenses = _extrair_other_expenses(primeira_linha)

    _set_cell_value_respeitando_merge(ws, "C6", input_property_type)
    _set_cell_value_respeitando_merge(ws, "C5", input_property_name)
    _set_cell_value_respeitando_merge(ws, "C7", input_address)
    _set_cell_value_respeitando_merge(ws, "C8", input_city_state)
    _set_cell_value_respeitando_merge(ws, "C9", input_number_units)
    _set_cell_value_respeitando_merge(ws, "C10", input_purchase_price)
    _set_cell_value_respeitando_merge(ws, "C11", _normalizar_down_payment(input_down_payment, input_purchase_price))
    _set_cell_value_respeitando_merge(ws, "D12", input_diligence)
    _set_cell_value_respeitando_merge(ws, "C13", _normalizar_percentual(input_loan_costs))
    _set_cell_value_respeitando_merge(ws, "C14", input_purchase_date)
    _set_cell_value_respeitando_merge(ws, "C15", input_end_year)

    for row in range(20, 28):
        _set_cell_value_respeitando_merge(ws, f"B{row}", "")
        _set_cell_value_respeitando_merge(ws, f"C{row}", "")

    fixed_revenues = [
        ("Gross Potential Rent", input_gpr),
        ("Vacancy Rate %", _normalizar_percentual(input_vacancy)),
        ("Credit Loss %", _normalizar_percentual(input_credit_loss)),
    ]

    revenue_row_cursor = 20
    for nome_income, valor_income in fixed_revenues:
        if revenue_row_cursor > 27:
            break
        _set_cell_value_respeitando_merge(ws, f"B{revenue_row_cursor}", nome_income)
        _set_cell_value_respeitando_merge(ws, f"C{revenue_row_cursor}", valor_income)
        revenue_row_cursor += 1

    max_income_rows = max(0, 27 - revenue_row_cursor + 1)
    other_incomes = _compactar_other_incomes(other_incomes, max_income_rows) if max_income_rows > 0 else []
    for _, nome_income, valor_income in other_incomes:
        _set_cell_value_respeitando_merge(ws, f"B{revenue_row_cursor}", nome_income)
        _set_cell_value_respeitando_merge(ws, f"C{revenue_row_cursor}", valor_income)
        revenue_row_cursor += 1

    for row in range(32, 47):
        _set_cell_value_respeitando_merge(ws, f"B{row}", "")
        _set_cell_value_respeitando_merge(ws, f"C{row}", "")

    fixed_expenses = [
        ("Property Tax", input_property_tax),
        ("Insurance", input_insurance),
        ("Property Management Fee (%EGI)", _normalizar_percentual(input_management_fee)),
        ("Repairs and Maintenance", input_repairs),
        ("Utilities", input_utilities),
        ("Capital Expenditures", input_capex_base),
        ("Landscape and Janitorial", input_landscape),
    ]

    row_cursor = 32
    for nome_expense, valor_expense in fixed_expenses:
        if row_cursor > 46:
            break
        _set_cell_value_respeitando_merge(ws, f"B{row_cursor}", nome_expense)
        _set_cell_value_respeitando_merge(ws, f"C{row_cursor}", valor_expense)
        row_cursor += 1

    max_dynamic_rows = max(0, 46 - row_cursor + 1)
    other_expenses = _compactar_other_expenses(other_expenses, max_dynamic_rows) if max_dynamic_rows > 0 else []
    for _, nome_expense, valor_expense in other_expenses:
        _set_cell_value_respeitando_merge(ws, f"B{row_cursor}", nome_expense)
        _set_cell_value_respeitando_merge(ws, f"C{row_cursor}", valor_expense)
        row_cursor += 1

    _set_cell_value_respeitando_merge(ws, "C51", input_capex1)
    _set_cell_value_respeitando_merge(ws, "C52", input_capex2)
    _set_cell_value_respeitando_merge(ws, "C53", input_capex3)
    _set_cell_value_respeitando_merge(ws, "C54", input_capex4)
    _set_cell_value_respeitando_merge(ws, "C55", input_capex5)
    _set_cell_value_respeitando_merge(ws, "C56", input_capex6)

    if row_cursor <= 46:
        ws.delete_rows(row_cursor, 46 - row_cursor + 1)
    if revenue_row_cursor <= 27:
        ws.delete_rows(revenue_row_cursor, 27 - revenue_row_cursor + 1)

    aplicar_formulas_apos_inputs(wb, ws)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    nome_arquivo = f"{input_property_name}_{timestamp}.xlsx"
    caminho_saida = os.path.join(PASTA_SAIDA, nome_arquivo)
    wb.save(caminho_saida)
    wb.close()

    print(f"✅ Arquivo criado com sucesso: {caminho_saida}")
    return caminho_saida


def obter_registros_pendentes(df_data: pd.DataFrame) -> pd.DataFrame:
    if "Submitted" not in df_data.columns:
        df_data["Submitted"] = "No"
    return df_data[df_data["Submitted"].astype(str).str.strip().str.lower() == "no"].copy()


def exibir_resumo_registro(registro) -> None:
    print(f"Property Type: {registro.get('Property Type', '')}")
    print(f"Property Name: {registro.get('Property Name', '')}")
    print(f"Address: {registro.get('Address', '')}")
    print(f"City and State: {registro.get('City and State', '')}")
    print(f"Number of Units: {registro.get('Number of Units', 0)}")
    print(f"Purchase Price: {registro.get('Purchase Price', 0)}")
    print(f"Down Payment (%): {registro.get('Down Payment (%)', registro.get('Down Payment', 0))}")
    print("--------------------------------------------")
    print(f"Gross Potential Rent: {registro.get('Gross Potential Rent', 0)}")
    print(f"Vacancy Rate: {registro.get('Vacancy Rate %', 0)}")
    print(f"Credit Loss: {registro.get('Credit Loss %', 0)}")
    print("--------------------------------------------")
    print(f"Property Tax: {registro.get('Property Tax', 0)}")
    print(f"Insurance: {registro.get('Insurance', 0)}")
    print(f"Management Fee: {registro.get('Management Fee %', 0)}")
    print(f"Repairs and Maintenance: {registro.get('Repairs and Maintenance', 0)}")
    print(f"Utilities: {registro.get('Utilities', 0)}")
    print(f"Capital Expenditures: {registro.get('Capital Expenditures', 0)}")
    print(f"Landscape and Janitorial: {registro.get('Landscape and Janitorial', 0)}")


def processar_registro_por_indice(registro_index: int) -> str | None:
    if not os.path.exists(ARQUIVO_DATA):
        print(f"Arquivo de dados nao encontrado: {ARQUIVO_DATA}")
        return None

    df_data = pd.read_excel(ARQUIVO_DATA)
    if registro_index < 0 or registro_index >= len(df_data):
        print(f"Indice de registro invalido: {registro_index}")
        return None

    registro = df_data.iloc[registro_index]
    submitted = str(registro.get("Submitted", "No")).strip().lower()
    if submitted == "yes":
        print(f"Registro no indice {registro_index} ja foi processado.")
        return None

    print(f"📌 Processando registro recem inserido no indice {registro_index}...")
    caminho_saida = criar_arquivo_baseado_em_template(registro)

    df_data.loc[registro.name, "Submitted"] = "Yes"
    df_data.to_excel(ARQUIVO_DATA, index=False)
    print("✅ Registro recem inserido processado e marcado como 'Submitted: Yes'.")

    return caminho_saida


def processar_primeiro_registro_pendente() -> str | None:
    garantir_pasta_saida(PASTA_SAIDA)

    if not os.path.exists(ARQUIVO_DATA):
        print(f"Arquivo de dados nao encontrado: {ARQUIVO_DATA}")
        return None

    df_data = pd.read_excel(ARQUIVO_DATA)
    pendentes = obter_registros_pendentes(df_data)

    if pendentes.empty:
        print("✅ Nenhum registro pendente para processar!")
        return None

    print(f"📋 Total de registros pendentes: {len(pendentes)}")
    print("=" * 60)

    primeira_linha = pendentes.iloc[0]
    exibir_resumo_registro(primeira_linha)

    print("\n📁 Gerando arquivo baseado no template...")
    caminho_saida = criar_arquivo_baseado_em_template(primeira_linha)

    # df_data.loc[primeira_linha.name, "Submitted"] = "Yes"
    df_data.to_excel(ARQUIVO_DATA, index=False)
    print("✅ Registro processado e marcado como 'Submitted: Yes' no arquivo de dados.")

    return caminho_saida


def main() -> None:
    processar_primeiro_registro_pendente()


if __name__ == "__main__":
    main()

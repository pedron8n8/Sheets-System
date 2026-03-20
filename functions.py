"""Formula engine for workbook updates.

This module contains only formula-related behavior:
- Calendar and dynamic formula updates in Monthly CF
- Label synchronization to Quarterly/Annual CF
- Formula reapply pipeline for files in Output
"""

from __future__ import annotations

import argparse
from datetime import date, datetime
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

PASTA_OUTPUT = Path("Output")
ARQUIVO_TEMPLATE = Path("InputTemplate.xlsx")
ABA_PADRAO = "Monthly CF"


def _set_cell_value_respeitando_merge(sheet, cell_ref: str, value) -> None:
    """Write into merged cells using the top-left anchor cell."""
    for merged_range in sheet.merged_cells.ranges:
        if cell_ref in merged_range:
            anchor_col = get_column_letter(merged_range.min_col)
            anchor_ref = f"{anchor_col}{merged_range.min_row}"
            sheet[anchor_ref] = value
            return
    sheet[cell_ref] = value


def _normalizar_purchase_date(valor):
    """Normalize supported date inputs to python date."""
    if valor is None or valor == "":
        return None

    if isinstance(valor, datetime):
        return valor.date()

    if isinstance(valor, date):
        return valor

    if isinstance(valor, str):
        texto = valor.strip()
        if not texto:
            return None
        for formato in ("%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y"):
            try:
                return datetime.strptime(texto, formato).date()
            except ValueError:
                continue

    return None


def _normalizar_end_year(valor, padrao: int = 10) -> int:
    """Normalize end-year horizon as positive integer."""
    if valor is None or valor == "":
        return padrao

    try:
        ano = int(float(valor))
    except (TypeError, ValueError):
        return padrao

    return ano if ano > 0 else padrao


def _extrair_nomes_other_income_inputs(ws_inputs) -> list[str]:
    linha_credit_loss = None
    linha_growth = None

    for linha in range(1, ws_inputs.max_row + 1):
        valor_b = ws_inputs.cell(row=linha, column=2).value
        if not isinstance(valor_b, str):
            continue

        texto = valor_b.strip().lower()
        if linha_credit_loss is None and "credit loss" in texto:
            linha_credit_loss = linha
            continue

        if linha_growth is None and "annual revenue growth rate" in texto:
            linha_growth = linha

    if linha_credit_loss is None or linha_growth is None or linha_growth <= linha_credit_loss:
        return []

    nomes = []
    for linha in range(linha_credit_loss + 1, linha_growth):
        valor_b = ws_inputs.cell(row=linha, column=2).value
        if isinstance(valor_b, str):
            nome = valor_b.strip()
            if nome:
                nomes.append(nome)

    return nomes


def _extrair_nomes_expenses_inputs(ws_inputs) -> list[str]:
    linha_inicio = None
    linha_growth = None

    for linha in range(1, ws_inputs.max_row + 1):
        valor_b = ws_inputs.cell(row=linha, column=2).value
        if not isinstance(valor_b, str):
            continue

        texto = valor_b.strip().lower()
        if linha_inicio is None and ("property tax" in texto or "property taxes" in texto):
            linha_inicio = linha
            continue

        if linha_growth is None and "annual expense growth rate" in texto:
            linha_growth = linha

    if linha_inicio is None or linha_growth is None or linha_growth <= linha_inicio:
        return []

    nomes = []
    for linha in range(linha_inicio, linha_growth):
        valor_b = ws_inputs.cell(row=linha, column=2).value
        if isinstance(valor_b, str):
            nome = valor_b.strip()
            if nome:
                nomes.append(nome)

    return nomes


def _atualizar_nomes_other_income_monthly_cf(ws_monthly, nomes: list[str], linha_inicial: int = 12, linha_final: int = 16) -> None:
    cursor = 0
    for linha in range(linha_inicial, linha_final + 1):
        if cursor < len(nomes):
            ws_monthly.cell(row=linha, column=2, value=nomes[cursor])
            cursor += 1
        else:
            ws_monthly.cell(row=linha, column=2, value="")


def _atualizar_nomes_expenses_monthly_cf(ws_monthly, nomes: list[str], linha_inicial: int = 20, linha_final: int = 34) -> None:
    cursor = 0
    for linha in range(linha_inicial, linha_final + 1):
        if cursor < len(nomes):
            ws_monthly.cell(row=linha, column=2, value=nomes[cursor])
            cursor += 1
        else:
            ws_monthly.cell(row=linha, column=2, value="")


def _aplicar_formulas_monthly_cf_dinamicas(ws_monthly) -> None:
    """Apply dynamic formulas in Monthly CF using name-based lookups."""
    ws_monthly.cell(row=47, column=2, value="Due Diligence Costs")
    ws_monthly.cell(row=48, column=2, value="Loan Origination Costs")

    for coluna in range(3, ws_monthly.max_column + 1):
        col = get_column_letter(coluna)
        col_anterior = get_column_letter(coluna - 1) if coluna > 3 else None

        formula_gross = (
            f'=IF(AND({col}$3>=INDEX(Inputs!$C:$C,MATCH("*Revenue Start Month*",Inputs!$B:$B,0)),'
            f'{col}$3<=INDEX(Inputs!$C:$C,MATCH("End Year",Inputs!$B:$B,0))*12),'
            f'INDEX(Inputs!$E:$E,MATCH("*Gross Potential Rent*",Inputs!$B:$B,0))/12*'
            f'(1+INDEX(Inputs!$C:$C,MATCH("*Annual Revenue Growth Rate*",Inputs!$B:$B,0)))^({col}$5-1),0)'
        )
        formula_vacancy = (
            f'=IF({col}$3<=INDEX(Inputs!$C:$C,MATCH("End Year",Inputs!$B:$B,0))*12,'
            f'-{col}8*INDEX(Inputs!$E:$E,MATCH("*Vacancy*",Inputs!$B:$B,0)),0)'
        )
        formula_credit_loss = (
            f'=IF({col}$3<=INDEX(Inputs!$C:$C,MATCH("End Year",Inputs!$B:$B,0))*12,'
            f'-{col}8*INDEX(Inputs!$E:$E,MATCH("*Credit Loss*",Inputs!$B:$B,0)),0)'
        )

        ws_monthly.cell(row=8, column=coluna, value=formula_gross)
        ws_monthly.cell(row=9, column=coluna, value=formula_vacancy)
        ws_monthly.cell(row=10, column=coluna, value=formula_credit_loss)

        for linha_income in range(12, 17):
            formula_other_income = (
                f'=IF($B{linha_income}="",0,IF(AND({col}$3>=INDEX(Inputs!$C:$C,MATCH("*Revenue Start Month*",Inputs!$B:$B,0)),'
                f'{col}$3<=INDEX(Inputs!$C:$C,MATCH("End Year",Inputs!$B:$B,0))*12),'
                f'INDEX(Inputs!$E:$E,MATCH($B{linha_income},Inputs!$B:$B,0))/12*'
                f'(1+INDEX(Inputs!$C:$C,MATCH("*Annual Revenue Growth Rate*",Inputs!$B:$B,0)))^({col}$5-1),0))'
            )
            ws_monthly.cell(row=linha_income, column=coluna, value=formula_other_income)

        for linha_expense in range(20, 35):
            formula_expense = (
                f'=IF($B{linha_expense}="",0,IF(AND({col}$3>=INDEX(Inputs!$C:$C,MATCH("*Expense Start Month*",Inputs!$B:$B,0)),'
                f'{col}$3<=INDEX(Inputs!$C:$C,MATCH("End Year",Inputs!$B:$B,0))*12),'
                f'IF(ISNUMBER(SEARCH("management fee",$B{linha_expense})),'
                f'-{col}11*INDEX(Inputs!$E:$E,MATCH($B{linha_expense},Inputs!$B:$B,0)),'
                f'-INDEX(Inputs!$E:$E,MATCH($B{linha_expense},Inputs!$B:$B,0))/12*'
                f'(1+INDEX(Inputs!$C:$C,MATCH("*Annual Expense Growth Rate*",Inputs!$B:$B,0)))^({col}$5-1)),0))'
            )
            ws_monthly.cell(row=linha_expense, column=coluna, value=formula_expense)

        formula_capex = (
            f'=IF({col}$3<=INDEX(Inputs!$C:$C,MATCH("End Year",Inputs!$B:$B,0))*12,'
            f'-SUMIFS(Inputs!$C:$C,Inputs!$B:$B,"*CapEx Item*",Inputs!$D:$D,{col}$3),0)'
        )
        ws_monthly.cell(row=40, column=coluna, value=formula_capex)

        sf_end_year = 'INDEX(Inputs!$C:$C,MATCH("End Year",Inputs!$B:$B,0))'
        sf_amount = 'INDEX(Inputs!$C:$C,MATCH("*Seller Finance Amount*",Inputs!$B:$B,0))'
        sf_rate = 'INDEX(Inputs!$C:$C,MATCH("*Seller Finance Interest Rate*",Inputs!$B:$B,0))'
        sf_amort_years = 'INDEX(Inputs!$C:$C,MATCH("*Seller Finance Amortization*",Inputs!$B:$B,0))'
        sf_balloon_eoy = 'INDEX(Inputs!$C:$C,MATCH("*Seller Finance Balloon*",Inputs!$B:$B,0))'
        sf_io_period = 'INDEX(Inputs!$C:$C,MATCH("*Seller Finance IO Period*",Inputs!$B:$B,0))'
        sf_orig_month = 'INDEX(Inputs!$C:$C,MATCH("*Seller Finance Origination Month*",Inputs!$B:$B,0))'
        sf_pmt_start = 'INDEX(Inputs!$C:$C,MATCH("*Seller Finance PMT Start Month*",Inputs!$B:$B,0))'

        # Debt Service - Seller Finance
        if coluna == 3:
            formula_sf_beginning_balance = "0"
        else:
            formula_sf_beginning_balance = (
                f'=IF({col}$3<={sf_end_year}*12,IF({col}$3<={sf_orig_month},0,{col_anterior}64),0)'
            )

        formula_sf_loan_funding = (
            f'=IF({col}$3<={sf_end_year}*12,IF({col}$3={sf_orig_month},{sf_amount},0),0)'
        )
        formula_sf_interest_payment = (
            f'=IF({col}$3<={sf_end_year}*12,IF(OR({col}59=0,{col}$3<{sf_pmt_start}),0,-{col}59*{sf_rate}/12),0)'
        )
        formula_sf_principal_payment = (
            f'=IF({col}$3<={sf_end_year}*12,IF(OR({col}59=0,{col}$3<{sf_pmt_start}),0,'
            f'IF({col}$3-{sf_pmt_start}+1<={sf_io_period},0,MIN(PMT({sf_rate}/12,{sf_amort_years}*12,-{sf_amount})-{col}59*{sf_rate}/12,{col}59))),0)'
        )
        formula_sf_loan_payoff = (
            f'=IF({col}$3<={sf_end_year}*12,IF(OR({col}$3={sf_balloon_eoy}*12,{col}$3={sf_end_year}*12),-MAX({col}59+{col}60-{col}62,0),0),0)'
        )
        formula_sf_ending_balance = (
            f'=IF({col}$3<={sf_end_year}*12,MAX({col}59+{col}60-{col}62+{col}63,0),0)'
        )
        formula_sf_total_debt_service = (
            f'=IF({col}$3<={sf_end_year}*12,{col}61-{col}62+{col}63,0)'
        )

        ws_monthly.cell(row=59, column=coluna, value=formula_sf_beginning_balance)
        ws_monthly.cell(row=60, column=coluna, value=formula_sf_loan_funding)
        ws_monthly.cell(row=61, column=coluna, value=formula_sf_interest_payment)
        ws_monthly.cell(row=62, column=coluna, value=formula_sf_principal_payment)
        ws_monthly.cell(row=63, column=coluna, value=formula_sf_loan_payoff)
        ws_monthly.cell(row=64, column=coluna, value=formula_sf_ending_balance)
        ws_monthly.cell(row=65, column=coluna, value=formula_sf_total_debt_service)

        bl_end_year = 'INDEX(Inputs!$C:$C,MATCH("End Year",Inputs!$B:$B,0))'
        bl_amount = 'INDEX(Inputs!$C:$C,MATCH("*Bank Loan Amount*",Inputs!$B:$B,0))'
        bl_rate = 'INDEX(Inputs!$C:$C,MATCH("*Bank Loan Interest Rate*",Inputs!$B:$B,0))'
        bl_amort_years = 'INDEX(Inputs!$C:$C,MATCH("*Bank Loan Amortization*",Inputs!$B:$B,0))'
        bl_balloon_eoy = 'INDEX(Inputs!$C:$C,MATCH("*Bank Loan Balloon*",Inputs!$B:$B,0))'
        bl_io_period = 'INDEX(Inputs!$C:$C,MATCH("*Bank Loan IO Period*",Inputs!$B:$B,0))'
        bl_orig_month = 'INDEX(Inputs!$C:$C,MATCH("*Bank Loan Origination Month*",Inputs!$B:$B,0))'
        bl_pmt_start = 'INDEX(Inputs!$C:$C,MATCH("*Bank Loan PMT Start Month*",Inputs!$B:$B,0))'

        # Debt Service - Bank Loan
        if coluna == 3:
            formula_bl_beginning_balance = "0"
        else:
            formula_bl_beginning_balance = (
                f'=IF({col}$3<={bl_end_year}*12,IF({col}$3<={bl_orig_month},0,{col_anterior}73),0)'
            )

        formula_bl_loan_funding = (
            f'=IF({col}$3<={bl_end_year}*12,IF({col}$3={bl_orig_month},{bl_amount},0),0)'
        )
        formula_bl_interest_payment = (
            f'=IF({col}$3<={bl_end_year}*12,IF(OR({col}68=0,{col}$3<{bl_pmt_start}),0,-{col}68*{bl_rate}/12),0)'
        )
        formula_bl_principal_payment = (
            f'=IF({col}$3<={bl_end_year}*12,IF(OR({col}68=0,{col}$3<{bl_pmt_start}),0,'
            f'IF({col}$3-{bl_pmt_start}+1<={bl_io_period},0,MIN(PMT({bl_rate}/12,{bl_amort_years}*12,-{bl_amount})-{col}68*{bl_rate}/12,{col}68))),0)'
        )
        formula_bl_loan_payoff = (
            f'=IF({col}$3<={bl_end_year}*12,IF(OR({col}$3={bl_balloon_eoy}*12,{col}$3={bl_end_year}*12),-MAX({col}68+{col}69-{col}71,0),0),0)'
        )
        formula_bl_ending_balance = (
            f'=IF({col}$3<={bl_end_year}*12,MAX({col}68+{col}69-{col}71+{col}72,0),0)'
        )
        formula_bl_total_debt_service = (
            f'=IF({col}$3<={bl_end_year}*12,{col}70-{col}71+{col}72,0)'
        )

        ws_monthly.cell(row=68, column=coluna, value=formula_bl_beginning_balance)
        ws_monthly.cell(row=69, column=coluna, value=formula_bl_loan_funding)
        ws_monthly.cell(row=70, column=coluna, value=formula_bl_interest_payment)
        ws_monthly.cell(row=71, column=coluna, value=formula_bl_principal_payment)
        ws_monthly.cell(row=72, column=coluna, value=formula_bl_loan_payoff)
        ws_monthly.cell(row=73, column=coluna, value=formula_bl_ending_balance)
        ws_monthly.cell(row=74, column=coluna, value=formula_bl_total_debt_service)

        # Comportamento antigo (distribuir em todos os meses ate End Year):
        # formula_due_diligence = f'=IF({col}$3<=Inputs!$C$15*12,-Inputs!$D$12,0)'
        # formula_loan_origination = f'=IF({col}$3<=Inputs!$C$15*12,-Inputs!$D$13,0)'
        # Novo comportamento: lancar somente na 1a coluna (mes 1).
        if coluna == 3:
            formula_due_diligence = f'=IF({col}$3=1,-Inputs!$D$12,0)'
            formula_loan_origination = f'=IF({col}$3=1,-Inputs!$D$13,0)'
        else:
            formula_due_diligence = "0"
            formula_loan_origination = "0"

        ws_monthly.cell(row=47, column=coluna, value=formula_due_diligence)
        ws_monthly.cell(row=48, column=coluna, value=formula_loan_origination)


def _atualizar_labels_resumo_cf(wb) -> None:
    """Sync dynamic labels from Monthly CF to Quarterly/Annual sheets."""
    if "Monthly CF" not in wb.sheetnames:
        return

    ws_monthly = wb["Monthly CF"]
    for nome_aba in ("Quarterly CF", "Annual CF"):
        if nome_aba not in wb.sheetnames:
            continue

        ws_resumo = wb[nome_aba]
        for offset in range(5):
            ws_resumo.cell(row=9 + offset, column=2, value=ws_monthly.cell(row=12 + offset, column=2).value)
        for offset in range(15):
            ws_resumo.cell(row=17 + offset, column=2, value=ws_monthly.cell(row=20 + offset, column=2).value)


def recalcular_formulas_proforma_inputs(ws_inputs) -> None:
    """Rebuild Inputs column E formulas after row deletion/reordering."""
    for linha in range(1, ws_inputs.max_row + 1):
        valor_b = ws_inputs.cell(row=linha, column=2).value
        if not isinstance(valor_b, str):
            continue

        nome = valor_b.strip().lower()

        if "annual revenue growth rate" in nome or "revenue start month" in nome:
            ws_inputs.cell(row=linha, column=5, value=None)
            continue

        if "annual expense growth rate" in nome or "expense start month" in nome:
            ws_inputs.cell(row=linha, column=5, value=None)
            continue

        if "vacancy" in nome or "credit loss" in nome or "property management fee" in nome:
            ws_inputs.cell(row=linha, column=5, value=f"=C{linha}")
            continue

        if (
            "gross potential rent" in nome
            or "other income" in nome
            or "property tax" in nome
            or "insurance" in nome
            or "repairs" in nome
            or "utilities" in nome
            or "capital expenditures" in nome
            or "landscape" in nome
            or "janitorial" in nome
            or "marketing" in nome
            or "reserves for replacement" in nome
            or "other expense" in nome
            or "pest control" in nome
            or "security" in nome
            or "trash removal" in nome
            or "legal" in nome
            or "turnover" in nome
            or "permits" in nome
            or "licenses" in nome
        ):
            ws_inputs.cell(row=linha, column=5, value=f"=C{linha}*(1+D{linha})")


def _recalcular_formulas_estrutura_inputs(ws_inputs) -> None:
    """Rebuild lower Inputs formulas using current row positions after deletions."""

    def _label(linha: int) -> str:
        valor = ws_inputs.cell(row=linha, column=2).value
        return valor.strip().lower() if isinstance(valor, str) else ""

    def _find_rows(contem: str) -> list[int]:
        alvo = contem.lower()
        rows = []
        for linha in range(1, ws_inputs.max_row + 1):
            if alvo in _label(linha):
                rows.append(linha)
        return rows

    end_year_rows = _find_rows("end year")
    total_rows = _find_rows("total")
    seller_amount_rows = _find_rows("seller finance amount")
    seller_balloon_rows = _find_rows("seller finance balloon")
    acquisition_bank_loan_rows = _find_rows("acquisition - bank loan")
    bank_amount_rows = _find_rows("bank loan amount")
    bank_balloon_rows = _find_rows("bank loan balloon")
    down_payment_rows = _find_rows("down payment")
    due_pct_rows = _find_rows("due diligence costs (%)")
    loan_orig_pct_rows = _find_rows("loan origination costs (%)")

    purchase_rows = _find_rows("purchase price")
    closing_rows = _find_rows("closing costs")
    immediate_rows = _find_rows("immediate repairs")
    due_rows = _find_rows("due diligence costs")
    loan_orig_rows = _find_rows("loan origination costs")
    acq_fee_rows = _find_rows("acquisition fee")
    acq_fee_pct_rows = _find_rows("acquisition fee (%")
    equity_rows = _find_rows("equity contribution")
    cap_partner_rows = _find_rows("capital partner contribution")
    manager_rows = _find_rows("manager contribution")
    cap_partner_share_rows = _find_rows("capital partner share")
    manager_share_rows = _find_rows("manager share")

    refi_cap_rate_rows = _find_rows("refi cap rate")
    refi_ltv_rows = _find_rows("refi ltv")
    refi_closing_cost_rows = _find_rows("refi closing cost")
    refi_year_noi_rows = _find_rows("refi year noi")
    refi_property_value_rows = _find_rows("refi property value")
    refi_closing_cost_amt_rows = _find_rows("refi closing costs ($)")
    refi_loan_amount_rows = _find_rows("refi loan amount")

    exit_cap_rate_rows = _find_rows("exit cap rate")
    selling_cost_pct_rows = _find_rows("selling cost (%")
    sale_year_noi_rows = _find_rows("sale year noi")
    selling_price_rows = _find_rows("selling price")
    selling_cost_amt_rows = _find_rows("selling costs ($)")
    net_sale_rows = _find_rows("net sale proceeds")

    gp_catchup_share_rows = _find_rows("gp catch-up share")
    tier_return_capital_rows = _find_rows("1. return of capital")
    tier_pref_return_rows = _find_rows("2. preferred return")
    tier_gp_catchup_rows = _find_rows("3. gp catch-up")
    tier_residual_below_rows = _find_rows("4. residual")
    tier_residual_above_rows = _find_rows("5. residual")

    end_year_row = end_year_rows[0] if end_year_rows else None
    total_row = max(total_rows) if total_rows else None

    if seller_amount_rows and total_row:
        ws_inputs.cell(row=seller_amount_rows[0], column=3, value=f"=D{total_row}")
    if seller_balloon_rows and end_year_row:
        ws_inputs.cell(row=seller_balloon_rows[0], column=3, value=f"=C{end_year_row}")

    if total_row:
        bank_amount_row = None
        if acquisition_bank_loan_rows:
            candidato = acquisition_bank_loan_rows[0] + 1
            if candidato <= ws_inputs.max_row:
                bank_amount_row = candidato
        if bank_amount_row is None and bank_amount_rows:
            bank_amount_row = bank_amount_rows[0]

        if bank_amount_row is not None:
            # Se a linha do Bank Loan Amount estiver mesclada (ex.: A:F), desfaz para manter
            # estrutura tabular: nome em B e valor em C.
            merges_para_desfazer = []
            for merged_range in ws_inputs.merged_cells.ranges:
                if (
                    merged_range.min_row <= bank_amount_row <= merged_range.max_row
                    and merged_range.min_col <= 2
                    and merged_range.max_col >= 3
                ):
                    merges_para_desfazer.append(str(merged_range))

            for merged_ref in merges_para_desfazer:
                ws_inputs.unmerge_cells(merged_ref)

            ws_inputs.cell(row=bank_amount_row, column=2, value="Bank Loan Amount")
            ws_inputs.cell(row=bank_amount_row, column=3, value=f"=E{total_row}")

    if bank_balloon_rows and end_year_row:
        ws_inputs.cell(row=bank_balloon_rows[0], column=3, value=f"=C{end_year_row}")

    down_payment_row = down_payment_rows[0] if down_payment_rows else None
    due_pct_row = due_pct_rows[0] if due_pct_rows else None
    loan_orig_pct_row = loan_orig_pct_rows[0] if loan_orig_pct_rows else None

    if not purchase_rows:
        return

    purchase_top_row = min(purchase_rows)
    purchase_uses_row = max(purchase_rows)

    closing_uses_row = min(closing_rows) if closing_rows else None
    closing_param_row = max(closing_rows) if len(closing_rows) >= 2 else None

    immediate_uses_row = min(immediate_rows) if immediate_rows else None
    immediate_param_row = max(immediate_rows) if len(immediate_rows) >= 2 else None

    due_source_row = min(due_rows) if due_rows else None
    due_uses_row = max(due_rows) if len(due_rows) >= 2 else None

    loan_source_row = min(loan_orig_rows) if loan_orig_rows else None
    loan_uses_row = max(loan_orig_rows) if len(loan_orig_rows) >= 2 else None

    acq_fee_uses_row = None
    for linha in acq_fee_rows:
        if "(%" not in _label(linha):
            acq_fee_uses_row = linha
            break
    acq_fee_pct_row = acq_fee_pct_rows[0] if acq_fee_pct_rows else None

    if purchase_uses_row:
        ws_inputs.cell(row=purchase_uses_row, column=3, value=f"=C{purchase_top_row}")

    if down_payment_row:
        ws_inputs.cell(row=down_payment_row, column=4, value=f"=C{purchase_top_row}*C{down_payment_row}")

    if due_pct_row and due_source_row:
        ws_inputs.cell(row=due_pct_row, column=3, value=f"=IF(C{purchase_top_row}=0,0,D{due_source_row}/C{purchase_top_row})")

    if loan_orig_pct_row:
        ws_inputs.cell(row=loan_orig_pct_row, column=4, value=f"=C{purchase_top_row}*C{loan_orig_pct_row}")

    if closing_uses_row and closing_param_row:
        ws_inputs.cell(row=closing_uses_row, column=3, value=f"=C{closing_param_row}")

    if acq_fee_uses_row and acq_fee_pct_row:
        ws_inputs.cell(row=acq_fee_uses_row, column=3, value=f"=C{purchase_top_row}*D{acq_fee_pct_row}")

    if due_uses_row and due_source_row:
        ws_inputs.cell(row=due_uses_row, column=3, value=f"=D{due_source_row}")

    if loan_uses_row and loan_source_row:
        ws_inputs.cell(row=loan_uses_row, column=3, value=f"=D{loan_source_row}")

    if immediate_uses_row and immediate_param_row:
        ws_inputs.cell(row=immediate_uses_row, column=3, value=f"=C{immediate_param_row}")

    if total_row and purchase_uses_row and immediate_uses_row:
        ws_inputs.cell(row=total_row, column=3, value=f"=SUM(C{purchase_uses_row}:C{immediate_uses_row})")
        ws_inputs.cell(row=total_row, column=4, value=f"=SUM(D{purchase_uses_row}:D{immediate_uses_row})")
        ws_inputs.cell(row=total_row, column=5, value=f"=SUM(E{purchase_uses_row}:E{immediate_uses_row})")

    if closing_param_row:
        ws_inputs.cell(row=closing_param_row, column=3, value=f"=C{purchase_top_row}*D{closing_param_row}")

    equity_row = equity_rows[0] if equity_rows else None
    cap_partner_row = cap_partner_rows[0] if cap_partner_rows else None
    manager_row = manager_rows[0] if manager_rows else None
    cap_partner_share_row = cap_partner_share_rows[0] if cap_partner_share_rows else None
    manager_share_row = manager_share_rows[0] if manager_share_rows else None

    if equity_row and total_row:
        ws_inputs.cell(row=equity_row, column=3, value=f"=F{total_row}")

    if manager_row and equity_row and cap_partner_row:
        ws_inputs.cell(row=manager_row, column=3, value=f"=C{equity_row}-C{cap_partner_row}")

    if cap_partner_share_row and equity_row and cap_partner_row:
        ws_inputs.cell(row=cap_partner_share_row, column=3, value=f"=IF(C{equity_row}=0,0,C{cap_partner_row}/C{equity_row})")

    if manager_share_row and equity_row and manager_row:
        ws_inputs.cell(row=manager_share_row, column=3, value=f"=IF(C{equity_row}=0,0,C{manager_row}/C{equity_row})")

    refi_cap_rate_row = refi_cap_rate_rows[0] if refi_cap_rate_rows else None
    refi_ltv_row = refi_ltv_rows[0] if refi_ltv_rows else None
    refi_closing_cost_row = refi_closing_cost_rows[0] if refi_closing_cost_rows else None
    refi_year_noi_row = refi_year_noi_rows[0] if refi_year_noi_rows else None
    refi_property_value_row = refi_property_value_rows[0] if refi_property_value_rows else None
    refi_closing_cost_amt_row = refi_closing_cost_amt_rows[0] if refi_closing_cost_amt_rows else None
    refi_loan_amount_row = refi_loan_amount_rows[0] if refi_loan_amount_rows else None

    if refi_property_value_row and refi_year_noi_row and refi_cap_rate_row:
        ws_inputs.cell(row=refi_property_value_row, column=3, value=f"=C{refi_year_noi_row}/C{refi_cap_rate_row}")

    if refi_closing_cost_amt_row and refi_property_value_row and refi_closing_cost_row:
        ws_inputs.cell(row=refi_closing_cost_amt_row, column=3, value=f"=C{refi_property_value_row}*C{refi_closing_cost_row}")

    if refi_loan_amount_row and refi_property_value_row and refi_ltv_row:
        ws_inputs.cell(row=refi_loan_amount_row, column=3, value=f"=C{refi_property_value_row}*C{refi_ltv_row}")

    exit_cap_rate_row = exit_cap_rate_rows[0] if exit_cap_rate_rows else None
    selling_cost_pct_row = selling_cost_pct_rows[0] if selling_cost_pct_rows else None
    sale_year_noi_row = sale_year_noi_rows[0] if sale_year_noi_rows else None
    selling_price_row = selling_price_rows[0] if selling_price_rows else None
    selling_cost_amt_row = selling_cost_amt_rows[0] if selling_cost_amt_rows else None
    net_sale_row = net_sale_rows[0] if net_sale_rows else None

    if selling_price_row and sale_year_noi_row and exit_cap_rate_row:
        ws_inputs.cell(row=selling_price_row, column=3, value=f"=C{sale_year_noi_row}/C{exit_cap_rate_row}")

    if selling_cost_amt_row and selling_price_row and selling_cost_pct_row:
        ws_inputs.cell(row=selling_cost_amt_row, column=3, value=f"=C{selling_price_row}*C{selling_cost_pct_row}")

    if net_sale_row and selling_price_row and selling_cost_amt_row:
        ws_inputs.cell(row=net_sale_row, column=3, value=f"=C{selling_price_row}-C{selling_cost_amt_row}")

    gp_catchup_share_row = gp_catchup_share_rows[0] if gp_catchup_share_rows else None
    tier_return_capital_row = tier_return_capital_rows[0] if tier_return_capital_rows else None
    tier_pref_return_row = tier_pref_return_rows[0] if tier_pref_return_rows else None
    tier_gp_catchup_row = tier_gp_catchup_rows[0] if tier_gp_catchup_rows else None
    tier_residual_below_row = tier_residual_below_rows[0] if tier_residual_below_rows else None
    tier_residual_above_row = tier_residual_above_rows[0] if tier_residual_above_rows else None

    if tier_return_capital_row:
        ws_inputs.cell(row=tier_return_capital_row, column=4, value=f"=1-C{tier_return_capital_row}")

    if tier_pref_return_row:
        ws_inputs.cell(row=tier_pref_return_row, column=4, value=f"=1-C{tier_pref_return_row}")

    if tier_gp_catchup_row and gp_catchup_share_row:
        ws_inputs.cell(row=tier_gp_catchup_row, column=3, value=f"=1-C{gp_catchup_share_row}")
        ws_inputs.cell(row=tier_gp_catchup_row, column=4, value=f"=C{gp_catchup_share_row}")

    if tier_residual_below_row:
        ws_inputs.cell(row=tier_residual_below_row, column=4, value=f"=1-C{tier_residual_below_row}")

    if tier_residual_above_row:
        ws_inputs.cell(row=tier_residual_above_row, column=4, value=f"=1-C{tier_residual_above_row}")


def aplicar_formulas_apos_inputs(wb, ws_inputs) -> None:
    """Public API: apply all formula-side updates after Inputs is populated."""
    purchase_date = _normalizar_purchase_date(ws_inputs["C14"].value)
    end_year = _normalizar_end_year(ws_inputs["C15"].value)

    recalcular_formulas_proforma_inputs(ws_inputs)
    _recalcular_formulas_estrutura_inputs(ws_inputs)

    if "Monthly CF" not in wb.sheetnames:
        return

    ws_monthly = wb["Monthly CF"]
    total_meses = end_year * 12

    for indice_mes, coluna in enumerate(range(3, ws_monthly.max_column + 1), start=1):
        if purchase_date and indice_mes <= total_meses:
            ws_monthly.cell(row=3, column=coluna, value=indice_mes)
            ws_monthly.cell(row=4, column=coluna, value=((indice_mes - 1) // 3) + 1)
            ws_monthly.cell(row=5, column=coluna, value=((indice_mes - 1) // 12) + 1)

            mes_base = (purchase_date.month - 1) + (indice_mes - 1)
            ano = purchase_date.year + (mes_base // 12)
            mes = (mes_base % 12) + 1
            ws_monthly.cell(row=6, column=coluna, value=date(ano, mes, 1))
        else:
            ws_monthly.cell(row=3, column=coluna, value=None)
            ws_monthly.cell(row=4, column=coluna, value=None)
            ws_monthly.cell(row=5, column=coluna, value=None)
            ws_monthly.cell(row=6, column=coluna, value=None)

    nomes_other_income = _extrair_nomes_other_income_inputs(ws_inputs)
    _atualizar_nomes_other_income_monthly_cf(ws_monthly, nomes_other_income)

    nomes_expenses = _extrair_nomes_expenses_inputs(ws_inputs)
    _atualizar_nomes_expenses_monthly_cf(ws_monthly, nomes_expenses)

    _aplicar_formulas_monthly_cf_dinamicas(ws_monthly)
    _atualizar_labels_resumo_cf(wb)


def reaplicar_formulas_do_template(
    arquivo_destino,
    arquivo_template: Path | str = ARQUIVO_TEMPLATE,
    nome_aba: str = ABA_PADRAO,
):
    """Copy formulas from template sheet and refresh dynamic formula blocks."""
    destino_path = Path(arquivo_destino)
    template_path = Path(arquivo_template)

    if not template_path.exists():
        raise FileNotFoundError(f"Template nao encontrado: {template_path}")
    if not destino_path.exists():
        raise FileNotFoundError(f"Arquivo destino nao encontrado: {destino_path}")

    wb_template = load_workbook(template_path, data_only=False)
    wb_destino = load_workbook(destino_path, data_only=False)

    if nome_aba not in wb_template.sheetnames:
        raise ValueError(f"A aba '{nome_aba}' nao existe no template.")
    if nome_aba not in wb_destino.sheetnames:
        raise ValueError(f"A aba '{nome_aba}' nao existe no arquivo destino.")

    ws_template = wb_template[nome_aba]
    ws_destino = wb_destino[nome_aba]
    ws_inputs_destino = wb_destino["Inputs"] if "Inputs" in wb_destino.sheetnames else None

    total_formulas = 0
    for row in ws_template.iter_rows(
        min_row=1,
        max_row=ws_template.max_row,
        min_col=1,
        max_col=ws_template.max_column,
    ):
        for cell in row:
            valor = cell.value
            if isinstance(valor, str) and valor.startswith("="):
                _set_cell_value_respeitando_merge(ws_destino, cell.coordinate, valor)
                total_formulas += 1

    if ws_inputs_destino is not None and nome_aba == ABA_PADRAO:
        aplicar_formulas_apos_inputs(wb_destino, ws_inputs_destino)

    wb_destino.save(destino_path)
    wb_template.close()
    wb_destino.close()

    return total_formulas


def _listar_xlsx(path_output) -> set[Path]:
    path_output = Path(path_output)
    if not path_output.exists():
        path_output.mkdir(parents=True, exist_ok=True)

    return {
        arquivo.resolve()
        for arquivo in path_output.glob("*.xlsx")
        if arquivo.is_file() and not arquivo.name.startswith("~$")
    }


def processar_output_uma_vez(
    path_output: Path | str = PASTA_OUTPUT,
    arquivo_template: Path | str = ARQUIVO_TEMPLATE,
    nome_aba: str = ABA_PADRAO,
) -> None:
    """Reapply formulas to all files in Output once."""
    path_output = Path(path_output)
    arquivo_template = Path(arquivo_template)

    print(f"Verificando pasta: {path_output.resolve()}")
    print(f"Template de formulas: {arquivo_template.resolve()}\n")

    arquivos = sorted(_listar_xlsx(path_output))
    if not arquivos:
        print("Nenhum arquivo .xlsx encontrado na pasta Output.")
        return

    for arquivo in arquivos:
        print(f"Processando: {arquivo.name}")
        try:
            qtd = reaplicar_formulas_do_template(
                arquivo_destino=arquivo,
                arquivo_template=arquivo_template,
                nome_aba=nome_aba,
            )
            print(f"OK | {arquivo.name} | formulas aplicadas: {qtd}")
        except Exception as erro:
            print(f"ERRO | {arquivo.name} | {erro}")


def main() -> None:
    """CLI entrypoint for formula reapply workflow."""
    parser = argparse.ArgumentParser(
        description="Verifica a pasta Output e reaplica formulas nos arquivos Excel."
    )
    parser.add_argument(
        "--output",
        default=str(PASTA_OUTPUT),
        help="Pasta monitorada (padrao: Output)",
    )
    parser.add_argument(
        "--template",
        default=str(ARQUIVO_TEMPLATE),
        help="Arquivo template com formulas (padrao: InputTemplate.xlsx)",
    )
    parser.add_argument(
        "--sheet",
        default=ABA_PADRAO,
        help="Nome da aba para copiar formulas (padrao: Monthly CF)",
    )

    args = parser.parse_args()
    processar_output_uma_vez(
        path_output=args.output,
        arquivo_template=args.template,
        nome_aba=args.sheet,
    )


if __name__ == "__main__":
    main()

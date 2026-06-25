# -*- coding: utf-8 -*-
"""Fixtures: construyen libros de gastos y pagos sintéticos en disco para poder
ejercitar el pipeline completo sin Google Drive ni datos reales."""

from datetime import datetime

import pandas as pd
import pytest

from payhawk_revision import Config


@pytest.fixture
def fecha():
    return datetime(2026, 6, 22)


@pytest.fixture
def gastos_df(fecha):
    """Conjunto que ejercita: factura nacional con IVA, divisa con comisión,
    IVA no deducible y factura multilínea."""
    filas = [
        # 1) Nacional con IVA repercutido (OP_INT)
        dict(EID=100, code="OP_INT", name="OP_INT 21%", total=121, net=100, tax=21,
             cur="EUR", paid=121, pais="Spain", vat="ESB12345678", cuenta="627000003",
             doof="FRA-100", doctype="Invoice", nota="servicio marketing"),
        # 2) Divisa USD exenta (OP_INT0, no se reclasifica) -> comisión por pago
        dict(EID=101, code="OP_INT0", name="Exento", total=50, net=50, tax=0,
             cur="USD", paid=50, pais="United States", vat="US999", cuenta="629000036",
             doof="FRA-101", doctype="Invoice", nota="suscripcion software"),
        # 3) IVA no deducible
        dict(EID=102, code="IVA_NODED10", name="IVA NO DEDUCIBLE 10%", total=110,
             net=100, tax=10, cur="EUR", paid=110, pais="Spain", vat="ESB87654321",
             cuenta="629000080", doof="FRA-102", doctype="Invoice", nota="atencion clientes"),
        # 4) Multilínea (misma factura, dos líneas)
        dict(EID=103, code="OP_INT", name="OP_INT 21%", total=36.3, net=30, tax=6.3,
             cur="EUR", paid=30, pais="Spain", vat="ESB11111111", cuenta="627000003",
             doof="FRA-103", doctype="Invoice", nota="material oficina linea 1"),
        dict(EID=103, code="OP_INT", name="OP_INT 21%", total=24.2, net=20, tax=4.2,
             cur="EUR", paid=20, pais="Spain", vat="ESB11111111", cuenta="627000003",
             doof="FRA-103", doctype="Invoice", nota="material oficina linea 2"),
    ]
    rows = []
    for f in filas:
        rows.append({
            "Expense ID": f["EID"],
            "Settlement Date": fecha,
            "Total Amount (EUR)": f["total"],
            "Net Amount (EUR)": f["net"],
            "Tax Amount (EUR)": f["tax"],
            "Tax Rate Name": f["name"],
            "Tax Rate Code": f["code"],
            "Tax Rate %": 21,
            "Paid Amount": f["paid"],
            "Paid Currency": f["cur"],
            "Expense Note": f["nota"],
            "Expense Category": "Gastos",
            "Departamento": "Marketing",
            "ID gasto DOOFINDER": f["doof"],
            "Cuenta contable gasto": f["cuenta"],
            "Document Number": f"DOC-{f['EID']}",
            "Document Type": f["doctype"],
            "Supplier Name": f"Proveedor {f['EID']}",
            "Supplier External ID": "410000123",
            "Supplier VAT": f["vat"],
            "Supplier Country": f["pais"],
        })
    return pd.DataFrame(rows)


@pytest.fixture
def pagos_df(fecha):
    """Payments por Expense ID. La divisa (101) tiene Credit > importe -> comisión 2."""
    return pd.DataFrame([
        {"Date": fecha, "Credit": 121.0, "Expense ID": 100},
        {"Date": fecha, "Credit": 52.0, "Expense ID": 101},   # 52 - 50 = 2 comisión
        {"Date": fecha, "Credit": 110.0, "Expense ID": 102},
        {"Date": fecha, "Credit": 50.0, "Expense ID": 103},   # 30 + 20
    ])


@pytest.fixture
def libros(tmp_path, gastos_df, pagos_df):
    """Escribe los DataFrames a .xlsx y devuelve las tuplas (path, hoja) que
    espera procesar()."""
    g_path = tmp_path / "gastos.xlsx"
    p_path = tmp_path / "pagos.xlsx"
    gastos_df.to_excel(g_path, sheet_name="Expenses", index=False)
    pagos_df.to_excel(p_path, sheet_name="Payments", index=False)
    return {
        "gastos": (g_path, "Expenses"),
        "pagos": (p_path, "Payments"),
        "plan": None,
        "carpeta": tmp_path,
    }


@pytest.fixture
def cfg_test():
    """Config de pruebas: sin VIES online ni hyperlinks (más rápido y determinista)."""
    return Config(
        FECHA_CONTABLE="24/06/2026",
        VALIDAR_VIES_ONLINE=False,
        PRESERVAR_HYPERLINKS=False,
    )

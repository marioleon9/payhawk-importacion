# -*- coding: utf-8 -*-
"""Test de extremo a extremo: ejecuta el pipeline completo sobre datos sintéticos
y comprueba las garantías contables clave (cuadre, comisiones, hojas, divisa)."""

import pandas as pd

from payhawk_revision import procesar


def test_pipeline_genera_excel_y_cuadra(libros, cfg_test):
    ruta, dfs = procesar(libros["gastos"], libros["pagos"], libros["plan"],
                         cfg_test, num_importacion=63,
                         carpeta_salida=libros["carpeta"])

    # 1) El archivo final se ha creado con el nombre esperado
    assert ruta.exists()
    assert ruta.name.startswith("IMPORTACION 63 ")

    # 2) El asiento de pago CUADRA (Debe == Haber). Garantía contable nº1.
    asiento = dfs["asiento"]
    total_debe = round(asiento["IMPORTE DEBE"].sum(), 2)
    total_haber = round(asiento["IMPORTE HABER"].sum(), 2)
    assert total_debe == total_haber

    # 3) No hay incidencias CRÍTICAS (las críticas indican descuadre o sin cruce)
    inc = dfs["incidencias"]
    if len(inc):
        assert (inc["Gravedad"] == "CRITICA").sum() == 0


def test_comision_divisa_correcta(libros, cfg_test):
    _, dfs = procesar(libros["gastos"], libros["pagos"], libros["plan"],
                      cfg_test, 63, libros["carpeta"])
    base = dfs["base"].set_index("Expense ID")
    # EID 101: pagado 50 EUR, cargado 52 -> comisión 2
    assert round(float(base.loc[101, "comision"]), 2) == 2.0
    assert bool(base.loc[101, "genera_comision"]) is True


def test_iva_no_deducible_a_cero(libros, cfg_test):
    _, dfs = procesar(libros["gastos"], libros["pagos"], libros["plan"],
                      cfg_test, 63, libros["carpeta"])
    df = dfs["df"].set_index("Expense ID")
    # EID 102: IVA_NODED -> la cuota deducible va a 0 para Importia/SII,
    # pero el tipo (%) y la cuota real se conservan (sigue siendo una operación
    # al 21%, solo que su IVA no es deducible). La deducibilidad la marca el
    # código TIPO_IVA de A3 (07), no el porcentaje.
    assert df.loc[102, "Tax Rate Code"] == "IVA_NODED"
    assert float(df.loc[102, "Cuota deducible A3"]) == 0
    assert float(df.loc[102, "Tax Rate %"]) == 21


def test_multilinea_colapsa_paid_amount(libros, cfg_test):
    _, dfs = procesar(libros["gastos"], libros["pagos"], libros["plan"],
                      cfg_test, 63, libros["carpeta"])
    df = dfs["df"]
    lineas = df[df["Expense ID"] == 103].sort_index()
    paids = sorted(lineas["Paid Amount"].tolist())
    # una línea con el total (50) y la otra a 0
    assert paids == [0.0, 50.0]


def test_tipo_iva_a3_no_deducible_y_normal(libros, cfg_test):
    import pandas as pd
    ruta, _ = procesar(libros["gastos"], libros["pagos"], libros["plan"],
                       cfg_test, 63, libros["carpeta"])
    hoja = pd.read_excel(ruta, sheet_name="GASTOS PREPARADO").set_index("Expense ID")
    assert "Tipo IVA A3" in hoja.columns
    # EID 102 es IVA no deducible -> debe llevar TIPO_IVA 07
    assert str(hoja.loc[102, "Tipo IVA A3"]).zfill(2) == "07"
    # EID 100 es operación interior con IVA -> 01
    assert str(hoja.loc[100, "Tipo IVA A3"]).zfill(2) == "01"


def test_hojas_excel_presentes(libros, cfg_test):
    ruta, _ = procesar(libros["gastos"], libros["pagos"], libros["plan"],
                       cfg_test, 63, libros["carpeta"])
    hojas = pd.ExcelFile(ruta).sheet_names
    for esperada in ("GASTOS PREPARADO", "ASIENTO PAGO", "INCIDENCIAS",
                     "RESUMEN_IMPORTACION", "DIVISAS"):
        assert esperada in hojas

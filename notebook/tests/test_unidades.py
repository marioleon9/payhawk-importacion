# -*- coding: utf-8 -*-
"""Tests unitarios de funciones de cálculo aisladas."""

import pandas as pd

from payhawk_revision import Config
from payhawk_revision.engine import (
    limpiar_tax_rate_code,
    calcular_iva_intracomunitario_isp,
    calcular_cuota_deducible_a3,
    _categoria_por_pais,
    _categoria_por_vat,
    nombre_hoja_departamento,
)
from payhawk_revision.reglas import Reglas


def _df(filas):
    df = pd.DataFrame(filas)
    if "Expense ID" not in df:
        df["Expense ID"] = range(1, len(df) + 1)
    return df


def test_limpia_codigos_iva():
    cfg = Config()
    df = _df([{"Tax Rate Code": "OP_INT21"}, {"Tax Rate Code": "OP_INT0"},
              {"Tax Rate Code": "IVA_NODED10"}])
    out = limpiar_tax_rate_code(df, [], cfg, [])
    assert out["Tax Rate Code"].tolist() == ["OP_INT", "OP_INT0", "IVA_NODED"]


def test_iva_isp_21_sobre_base():
    cfg = Config()
    df = _df([{"Tax Rate Code": "INV_SUJ_PAS", "Net Amount (EUR)": 100,
               "Tax Amount (EUR)": 0}])
    out = calcular_iva_intracomunitario_isp(df, cfg, [])
    assert round(float(out.loc[0, "Tax Amount (EUR)"]), 2) == 21.0


def test_cuota_deducible_normal_vs_no_deducible():
    cfg = Config()
    df = _df([
        {"Tax Rate Code": "OP_INT", "Tax Rate Name": "OP_INT", "Tax Amount (EUR)": 21},
        {"Tax Rate Code": "IVA_NODED", "Tax Rate Name": "IVA NO DEDUCIBLE", "Tax Amount (EUR)": 10},
    ])
    out = calcular_cuota_deducible_a3(df, [], cfg)
    assert float(out.loc[0, "Cuota deducible A3"]) == 21.0
    assert float(out.loc[1, "Cuota deducible A3"]) == 0.0


def test_categoria_por_pais_y_vat():
    assert _categoria_por_pais("Spain") == "ES"
    assert _categoria_por_pais("Germany") == "EU"
    assert _categoria_por_pais("United States") == "EXTRA"
    assert _categoria_por_vat("DE123456789") == "EU"
    assert _categoria_por_vat("NIFALGO") == ""   # NIF provisional, sin categoría


def test_nombre_hoja_usa_reglas():
    reglas = Reglas()
    assert nombre_hoja_departamento("Marketing", reglas) == "MARKETING"
    assert nombre_hoja_departamento("RRHH", reglas) == "RRHH"
    assert nombre_hoja_departamento("", reglas) == "SIN DEPARTAMENTO"
    # departamento no mapeado -> mayúsculas dinámicas
    assert nombre_hoja_departamento("Legal", reglas) == "LEGAL"

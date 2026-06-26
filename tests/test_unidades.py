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


def test_cuenta_proveedor_no_valida_se_corrige():
    from payhawk_revision.engine import rellenar_proveedores_vacios
    cfg = Config()
    df = _df([
        # ticket con cuenta random -> 410000000
        {"Supplier External ID": "60306282", "Supplier VAT": "X", "Supplier Name": "ROSSMANN",
         "Document Type": "Receipt", "Tax Rate Code": "NOSUJ_SDED",
         "Expense Note": "x", "Document Number": "1", "ID gasto DOOFINDER": "TICKET_X"},
        # factura normal con CIF en la cuenta -> también se corrige (regla global)
        {"Supplier External ID": "A80241789", "Supplier VAT": "Y", "Supplier Name": "SERVEO",
         "Document Type": "Invoice", "Tax Rate Code": "OP_INT",
         "Expense Note": "x", "Document Number": "2", "ID gasto DOOFINDER": "FRA-1"},
        # cuenta 410 válida -> se respeta
        {"Supplier External ID": "410004692", "Supplier VAT": "Z", "Supplier Name": "GALBEN",
         "Document Type": "Invoice", "Tax Rate Code": "OP_INT",
         "Expense Note": "x", "Document Number": "3", "ID gasto DOOFINDER": "FRA-2"},
    ])
    out = rellenar_proveedores_vacios(df, [], cfg, [])
    assert out.loc[0, "Supplier External ID"] == "410000000"   # ticket random -> genérica
    assert out.loc[1, "Supplier External ID"] == "410000000"   # CIF en factura -> genérica
    assert out.loc[2, "Supplier External ID"] == "410004692"   # 410 válida -> intacta


def test_traduccion_op_int0_a_op_int():
    from payhawk_revision.engine import traducir_tax_code_a_a3
    cfg = Config()
    df = _df([{"Tax Rate Code": "OP_INT0"}, {"Tax Rate Code": "OP_INT"},
              {"Tax Rate Code": "NOSUJ_SDED"}, {"Tax Rate Code": "IVA_NODED"}])
    out = traducir_tax_code_a_a3(df, cfg, [])
    # OP_INT0 (exento) -> OP_INT; el resto ya coincide con A3 y no cambia
    assert out["Tax Rate Code"].tolist() == ["OP_INT", "OP_INT", "NOSUJ_SDED", "IVA_NODED"]


def test_nombre_hoja_usa_reglas():
    reglas = Reglas()
    assert nombre_hoja_departamento("Marketing", reglas) == "MARKETING"
    assert nombre_hoja_departamento("RRHH", reglas) == "RRHH"
    assert nombre_hoja_departamento("", reglas) == "SIN DEPARTAMENTO"
    # departamento no mapeado -> mayúsculas dinámicas
    assert nombre_hoja_departamento("Legal", reglas) == "LEGAL"

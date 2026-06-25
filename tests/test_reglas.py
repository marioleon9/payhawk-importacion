# -*- coding: utf-8 -*-
"""Tests del cargador de reglas externas: defaults intactos y fusión del YAML."""

from payhawk_revision.reglas import cargar_reglas, Reglas


def test_sin_archivo_usa_defaults():
    reglas = cargar_reglas(None)
    assert reglas.mapa_departamentos["marketing"] == "MARKETING"
    assert "anthropic" in reglas.reglas_proveedor


def test_ruta_inexistente_no_rompe(tmp_path):
    reglas = cargar_reglas(tmp_path / "no_existe.yaml")
    assert isinstance(reglas, Reglas)
    assert reglas.mapa_departamentos["marketing"] == "MARKETING"


def test_yaml_fusiona_sobre_defaults(tmp_path):
    yaml = tmp_path / "reglas.yaml"
    yaml.write_text(
        "departamentos:\n"
        "  legal: LEGAL\n"
        "reglas_proveedor:\n"
        "  notion:\n"
        "    Operaciones: '629000999'\n",
        encoding="utf-8",
    )
    reglas = cargar_reglas(yaml)
    # lo nuevo del YAML
    assert reglas.mapa_departamentos["legal"] == "LEGAL"
    assert reglas.reglas_proveedor["notion"]["Operaciones"] == "629000999"
    # lo de fábrica se conserva
    assert reglas.mapa_departamentos["marketing"] == "MARKETING"
    assert "anthropic" in reglas.reglas_proveedor


def test_yaml_corrupto_no_rompe(tmp_path):
    yaml = tmp_path / "reglas.yaml"
    yaml.write_text("esto: no: es: valido:\n  - [", encoding="utf-8")
    reglas = cargar_reglas(yaml)   # no debe lanzar
    assert reglas.mapa_departamentos["marketing"] == "MARKETING"

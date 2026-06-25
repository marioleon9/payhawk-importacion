# -*- coding: utf-8 -*-
"""
reglas.py
=========

Reglas de negocio EXTERNALIZADAS del motor.

La idea de diseño es doble:

1. **El sistema funciona sin tocar nada.** Todos los valores por defecto están
   aquí incrustados (``DEFAULTS``). Si nadie aporta un archivo de reglas, el
   proceso se ejecuta con estos valores "de fábrica".

2. **Las reglas se cambian sin tocar Python.** Si en la carpeta de Drive existe
   un archivo ``reglas.yaml``, sus valores se fusionan SOBRE los de fábrica. Así
   el equipo financiero puede ajustar cuentas, proveedores o departamentos
   editando un único archivo con comentarios en español, sin abrir el código.

   Los datos sensibles (cuentas reales por proveedor) deben vivir en ese
   ``reglas.yaml`` dentro de Google Drive, NO en el repositorio público.
"""

from __future__ import annotations

import logging
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

log = logging.getLogger("payhawk")


# =============================================================================
# Valores por defecto ("de fábrica")
# =============================================================================

# Departamento (Payhawk, normalizado en minúsculas) -> nombre de hoja Excel.
DEFAULT_MAPA_DEPARTAMENTOS = {
    "marketing": "MARKETING",
    "comercial": "COMERCIAL",
    "cx & support": "CX & Support",
    "cx&support": "CX & Support",
    "cxsupport": "CX & Support",
    "desarrollo": "DESARROLLO",
    "direccion": "DIRECCION",
    "gastos generales": "GASTOS GENERALES",
    "producto": "PRODUCTO",
    "operaciones": "OPERACIONES",
    "recursos humanos": "RRHH",
    "rrhh": "RRHH",
    "financiero": "FINANCIERO",
    "direccion general": "DIRECCION",
}

# Familias de departamento detectadas a partir del INICIO de la descripción de la
# cuenta en el Plan Contable. Cada entrada: (claves_normalizadas, familia).
DEFAULT_FAMILIAS_CUENTA = [
    (("marketing",), "Marketing"),
    (("comercial",), "Comercial"),
    (("desarrollo", "desarro"), "Desarrollo"),
    (("operaciones", "oper"), "Operaciones"),
    (("producto",), "Producto"),
    (("financiero",), "Financiero"),
    (("rrhh",), "Recursos Humanos"),
    (("direccion",), "Direccion"),
    (("gastos generales", "generales"), "Gastos generales"),
    (("cx&support", "cx & support", "cs ", "ce ", "soporte",
      "customer success", "customer expirience", "customer experience",
      "customer exp"), "CX & Support"),
    (("oficina usa",), "OFICINA USA"),
]

# Departamento Payhawk (normalizado) -> familia comparable.
DEFAULT_FAMILIAS_PAYHAWK = {
    "marketing": "Marketing",
    "comercial": "Comercial",
    "desarrollo": "Desarrollo",
    "operaciones": "Operaciones",
    "producto": "Producto",
    "financiero": "Financiero",
    "recursos humanos": "Recursos Humanos",
    "rrhh": "Recursos Humanos",
    "direccion": "Direccion",
    "gastos generales": "Gastos generales",
    "cx & support": "CX & Support",
    "cx&support": "CX & Support",
}

# Reglas de cuenta por proveedor (subcadena normalizada buscada en
# "ID gasto DOOFINDER" o "Supplier Name") -> {familia_departamento: cuenta}.
DEFAULT_REGLAS_PROVEEDOR = {
    "anthropic": {"Marketing": "627000003", "Desarrollo": "629000036",
                  "Operaciones": "629000135"},
    "google cloud": {"Operaciones": "629000106", "Desarrollo": "629000036"},
    "googlecloud": {"Operaciones": "629000106", "Desarrollo": "629000036"},
    "lusha": {"Operaciones": "629000106"},
    "cloudi nextgen": {"Desarrollo": "629000035"},
    "cloudinextgen": {"Desarrollo": "629000035"},
    "asador gaztelu": {"Direccion": "629000080"},
    "asadorgaztelu": {"Direccion": "629000080"},
}


@dataclass
class Reglas:
    """Tablas de reglas de negocio. Si no se carga ningún YAML, todas toman su
    valor por defecto y el sistema funciona igual que antes."""

    mapa_departamentos: dict = field(
        default_factory=lambda: dict(DEFAULT_MAPA_DEPARTAMENTOS))
    familias_cuenta: list = field(
        default_factory=lambda: list(DEFAULT_FAMILIAS_CUENTA))
    familias_payhawk: dict = field(
        default_factory=lambda: dict(DEFAULT_FAMILIAS_PAYHAWK))
    reglas_proveedor: dict = field(
        default_factory=lambda: {k: dict(v) for k, v in DEFAULT_REGLAS_PROVEEDOR.items()})

    @property
    def familias_comparables(self) -> set:
        """Familias con equivalente claro en Payhawk (para decidir si una
        discrepancia debe marcarse)."""
        return set(self.familias_payhawk.values())


def _norm_key(text) -> str:
    """Normalización mínima e independiente para las claves de los diccionarios
    (minúsculas + espacios colapsados). Evita acoplar este módulo a utils."""
    return " ".join(str(text).strip().lower().split())


def cargar_reglas(ruta: Optional[Path] = None) -> Reglas:
    """Devuelve un objeto :class:`Reglas`.

    - Si ``ruta`` es None o el archivo no existe -> reglas de fábrica.
    - Si existe y PyYAML está disponible -> se fusionan los valores del YAML
      SOBRE los de fábrica (lo que no aparezca en el YAML conserva su default).

    El método nunca lanza por un YAML mal formado: registra un aviso y sigue con
    los valores de fábrica, para no romper la ejecución del usuario final.
    """
    reglas = Reglas()
    if ruta is None:
        return reglas
    ruta = Path(ruta)
    if not ruta.exists():
        log.info("No hay reglas.yaml en %s: se usan las reglas de fábrica.", ruta)
        return reglas

    try:
        import yaml  # type: ignore
    except ImportError:
        log.warning("PyYAML no está instalado: se ignora %s y se usan las "
                    "reglas de fábrica.", ruta)
        return reglas

    try:
        with open(ruta, "r", encoding="utf-8") as fh:
            datos = yaml.safe_load(fh) or {}
    except Exception as exc:
        log.warning("No se pudo leer %s (%s): se usan las reglas de fábrica.",
                    ruta, exc)
        return reglas

    if not isinstance(datos, dict):
        log.warning("El contenido de %s no es un mapa YAML: se ignora.", ruta)
        return reglas

    # --- departamentos: dict normalizando claves ---
    if isinstance(datos.get("departamentos"), dict):
        for k, v in datos["departamentos"].items():
            reglas.mapa_departamentos[_norm_key(k)] = str(v)

    # --- familias por departamento Payhawk ---
    if isinstance(datos.get("familias_payhawk"), dict):
        for k, v in datos["familias_payhawk"].items():
            reglas.familias_payhawk[_norm_key(k)] = str(v)

    # --- familias detectadas en la descripción de la cuenta ---
    if isinstance(datos.get("familias_cuenta"), list):
        nuevas = []
        for item in datos["familias_cuenta"]:
            if isinstance(item, dict) and "claves" in item and "familia" in item:
                claves = tuple(_norm_key(c) for c in item["claves"])
                nuevas.append((claves, str(item["familia"])))
        if nuevas:
            reglas.familias_cuenta = nuevas

    # --- reglas de cuenta por proveedor ---
    if isinstance(datos.get("reglas_proveedor"), dict):
        for proveedor, mapping in datos["reglas_proveedor"].items():
            if isinstance(mapping, dict):
                reglas.reglas_proveedor[_norm_key(proveedor)] = {
                    str(fam): str(cta) for fam, cta in mapping.items()
                }

    log.info("Reglas cargadas desde %s (fusionadas sobre las de fábrica).", ruta)
    return reglas

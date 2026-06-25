# -*- coding: utf-8 -*-
"""
payhawk_revision
================

Revisión y preparación de gastos de Payhawk para su importación en A3/Importia.

Uso típico (Google Colab):

    from payhawk_revision import Config, main
    main(Config(FECHA_CONTABLE="24/06/2026"))

Uso programático (local / tests):

    from payhawk_revision import procesar, Config
    ruta, dfs = procesar(gastos, pagos, plan, Config(), num, carpeta)
"""

from .engine import (
    Config,
    main,
    procesar,
    detectar_archivos,
    localizar_carpeta_importacion,
)
from .reglas import Reglas, cargar_reglas

__all__ = [
    "Config",
    "main",
    "procesar",
    "detectar_archivos",
    "localizar_carpeta_importacion",
    "Reglas",
    "cargar_reglas",
]

__version__ = "1.0.0"

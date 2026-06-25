# Payhawk → A3/Importia · Revisión y preparación de gastos

Automatiza la preparación previa a la importación manual de gastos de **Payhawk**
en **A3 / Importia**: limpia el archivo de gastos, cruza los pagos, calcula
comisiones de divisa, genera el asiento de pago, separa una hoja por departamento
y registra todas las **incidencias** para revisión.

> **Filosofía: prudente por diseño.** Ante cualquier duda, el sistema **no**
> modifica el dato: lo lleva a la hoja `INCIDENCIAS` para revisión manual.

---

## 🟢 Uso normal (para quien no es técnico)

Solo 3 pasos, sin instalar nada en el ordenador:

1. Sube los **2 archivos de Payhawk** (gastos y pagos) a la carpeta del número de
   importación en Drive (p. ej. `…/REVISION PAGOS PAYHAWK/2026/06/63/`).
2. Abre el notebook en Google Colab y pulsa **Entorno de ejecución → Ejecutar todo**.
3. Recoge el archivo **`IMPORTACION XX ddmmaa.xlsx`** que aparece en esa misma carpeta.

Lo único que se edita en el día a día es la **fecha** en la celda de configuración.
El motor se descarga solo desde GitHub al ejecutar, así que siempre se usa la
última versión sin reenviar el notebook a nadie.

👉 Notebook: [`notebook/IMPORTACION_Payhawk.ipynb`](notebook/IMPORTACION_Payhawk.ipynb)

---

## 📄 Hojas del Excel generado

| Hoja | Contenido |
|------|-----------|
| `GASTOS PREPARADO` | Archivo principal limpio + columnas de cruce de divisa |
| `ASIENTO PAGO` | Asiento de pago (Debe proveedor / Haber 572) + comisiones |
| *una por departamento* | Líneas separadas por departamento |
| `RETENCIONES` | Facturas con retención/IRPF (a contabilizar a mano) |
| `VALIDACION_CUENTAS` | Coherencia cuenta ↔ departamento (modo prudente) |
| `CAMBIOS_APLICADOS` | Cada cambio automático: antes, después y motivo |
| `INCIDENCIAS` | Todo lo dudoso, ordenado por gravedad |
| `RESUMEN_IMPORTACION` | Totales, cuadre Debe/Haber, recuento de incidencias |
| `DIVISAS` | Trazabilidad del cruce de moneda extranjera |
| `HISTORICO_BASE` | Estructura preparada para histórico acumulado |

---

## ⚙️ Reglas de negocio sin tocar código (`reglas.yaml`)

Las reglas que cambian con el negocio (cuentas por proveedor, departamentos…)
**no** están en el código: se editan en un archivo `reglas.yaml`.

- Es **opcional**: si no existe, el sistema usa valores "de fábrica" y funciona igual.
- Va en **tu Google Drive**, en la carpeta base (`REVISION PAGOS PAYHAWK/reglas.yaml`),
  **no en este repositorio**. Así los datos sensibles (cuentas reales) nunca se publican.
- Plantilla comentada: [`reglas.ejemplo.yaml`](reglas.ejemplo.yaml) → renómbrala a
  `reglas.yaml` y edítala.

Lo que pongas en el YAML se **fusiona sobre** los valores de fábrica: solo defines
lo que cambia.

---

## 🧱 Arquitectura

```
src/payhawk_revision/
  ├── engine.py    Motor: lectura, limpieza fiscal, cruce de pagos, asiento, export
  ├── reglas.py    Reglas de negocio (defaults de fábrica + carga del YAML)
  └── __init__.py  API pública: Config, main, procesar
notebook/          Notebook fino para Colab (config + ejecutar)
tests/             Tests automáticos (pytest)
reglas.ejemplo.yaml
```

La misma lógica de negocio corre en **Colab** y en **local**, sin duplicar código.

---

## 🧪 Desarrollo y tests

```bash
pip install -e ".[dev]"
pytest
```

Los tests usan datos sintéticos (sin Drive ni datos reales) y verifican las
garantías contables clave: **cuadre Debe/Haber**, comisión de divisa, IVA no
deducible a 0, colapso de facturas multilínea y la carga/fusión de `reglas.yaml`.

---

## 🚀 Cómo se distribuye

El notebook instala el motor directamente desde GitHub:

```python
!pip install -q "git+https://github.com/marioleon9/payhawk-importacion.git"
```

Para publicar una mejora: se hace `commit` + `push` a este repo y, la próxima vez
que alguien ejecute el notebook, obtiene la versión nueva automáticamente.

# 📊 Generador de Reportes Semanales de Monitoreo - Omnifish

Sistema automatizado para la actualización de tableros interactivos y generación de reportes semanales ejecutivos de monitoreo acuícola (Camanchaca, Cermaq, Mowi).

---

## 📖 Flujo de Trabajo Semanal

```
[1. Copiar Excels] ➔ [2. Ejecutar .\generar_reporte.ps1] ➔ [3. Abrir reporte_semanal.html] ➔ [4. Guardar / Imprimir PDF]
```

### 📋 Paso 1: Guardar los archivos Excel
Copia los archivos Excel actualizados de la semana (ej: `Reporte Camanchaca 2026 semana 40.xlsx`) dentro de la carpeta `Datos_Excel/`.
*(No es necesario borrar los anteriores, el script detecta automáticamente el archivo más reciente de cada empresa).*

### ⚡ Paso 2: Actualizar la Web desde PowerShell
Abre PowerShell en la carpeta del proyecto y ejecuta:
```powershell
.\generar_reporte.ps1
```
Este comando lee los nuevos libros de Excel, procesa las métricas acuícolas por región y genera/actualiza [`reporte_semanal.html`](./reporte_semanal.html).

### 🌐 Paso 3: Previsualizar e Imprimir a PDF
1. Abre [`reporte_semanal.html`](./reporte_semanal.html) en Chrome o Edge.
2. Navega entre las empresas (**Camanchaca**, **Cermaq**, **Mowi**) usando la barra superior.
3. Revisa métricas, observaciones operativas y el resumen ejecutivo dinámico.
4. Presiona el botón **🖨️ Guardar PDF / Imprimir** para exportar cada informe.

---

## 💡 Opciones del Script

| Objetivo | Comando |
|---|---|
| Actualizar el dashboard con las tres empresas | `.\generar_reporte.ps1` |
| Actualizar solo una empresa | `.\generar_reporte.ps1 -Empresa Cermaq` |
| Generar directamente los PDFs vía Microsoft Edge | `.\generar_reporte.ps1 -GenerarPDF` |
| Generar el PDF de una empresa específica | `.\generar_reporte.ps1 -Empresa Mowi -GenerarPDF` |

---

## 🔒 Privacidad y Plantillas

* `reporte_semanal.template.html` es la plantilla maestra limpia versionada en Git. Contiene la estructura visual, estilos, gráficos y el branding de Omnifish, pero sin datos operacionales de clientes.
* `reporte_semanal.html` se crea y actualiza localmente al ejecutar el script y está excluido de Git en `.gitignore` para resguardar la confidencialidad de la información.
* Los archivos en `Datos_Excel/` y `Reportes_PDF/` también se mantienen fuera de Git.

---

## 🛠️ Estructura del Proyecto

```
Generador de Reportes Semanales/
├── assets/
│   └── logos/
│       ├── Camanchaca logo.png
│       ├── Cermaq logo.jpg
│       ├── Mowi logo.png
│       └── Omnifish logo.png
├── Datos_Excel/                     # Archivos Excel semanales (ignorados en git)
├── Reportes_PDF/                    # Reportes generados en PDF (ignorados en git)
├── generar_reporte.ps1              # Script de procesamiento y generación
├── reporte_semanal.template.html    # Plantilla limpia versionada sin datos de clientes
└── reporte_semanal.html             # Dashboard generado localmente (ignorado en git)
```

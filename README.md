# 📊 Generador de Reportes Semanales de Monitoreo

Sistema automatizado para la actualización de tableros interactivos y generación de reportes semanales ejecutivos (Camanchaca, Cermaq, Mowi).

---

## 📖 Flujo de Trabajo Semanal

```
[1. Copiar Excels] ➔ [2. Ejecutar .\generar_reporte.ps1] ➔ [3. Abrir reporte_semanal.html] ➔ [4. Guardar / Imprimir PDF]
```

### 📋 Paso 1: Guardar los archivos Excel
Copia los 3 archivos Excel actualizados de la semana (ej: `Reporte Camanchaca 2026 semana 35.xlsx`) dentro de la carpeta `Datos_Excel/`.
*(No es necesario borrar los anteriores, el script detecta automáticamente el archivo más reciente).*

### ⚡ Paso 2: Actualizar la Web desde PowerShell
Abre PowerShell en la carpeta del proyecto y ejecuta:
```powershell
.\generar_reporte.ps1
```
Este comando lee los nuevos libros de Excel y actualiza la página web interactiva [`reporte_semanal.html`](./reporte_semanal.html).

### 🌐 Paso 3: Previsualizar e Imprimir a PDF
1. Abre [`reporte_semanal.html`](./reporte_semanal.html) en Chrome o Edge.
2. Navega entre las empresas (**Camanchaca**, **Cermaq**, **Mowi**) usando la barra superior.
3. Revisa métricas, observaciones y edita notas si lo requieres.
4. Presiona el botón **🖨️ Guardar PDF / Imprimir** para exportar cada informe.

---

## 💡 Opción Avanzada: Generación Directa de PDFs por Consola
Si deseas generar directamente los archivos PDF en `Reportes_PDF/Semana_XX/` sin pasar por la web:
```powershell
.\generar_reporte.ps1 -GenerarPDF
```

---

## 🛠️ Estructura del Proyecto
```
Generador de Reportes Semanales/
├── assets/
│   └── logos/
│       ├── Camanchaca logo.png
│       ├── Cermaq logo.jpg
│       └── Mowi logo.png
├── Datos_Excel/            # Archivos Excel semanales (ignorados en git)
├── Reportes_PDF/           # Reportes generados en PDF (ignorados en git)
├── generar_reporte.ps1     # Script procesador de datos
└── reporte_semanal.html    # Aplicación Web / Dashboard interactivo
```

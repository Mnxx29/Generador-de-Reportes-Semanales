# Generador de Reportes Semanales de Monitoreo

Sistema automatizado para la generación de reportes semanales ejecutivos y detallados de centros de cultivo de salmónidos (Camanchaca, Cermaq, Mowi).

## 🚀 Características
- **Vista Previa Interactiva en Web (`reporte_semanal.html`)**: Permite visualizar y editar observaciones/insights por centro antes de generar el informe final.
- **Formato Ejecutivo de 2 Páginas**:
  - **Página 1**: Dashboard Gerencial (Métricas Generales, Resumen Regional, Grid de 4 Gráficos Vectoriales 2x2 e Indicadores de Salud Tecnológica de Infraestructura).
  - **Página 2**: Detalle Completo por Centro (Jaulas, Cámaras, Observaciones/Insights con insignias de color) y Notas Ejecutivas.
- **Generación Automática en PDF (`generar_reporte.ps1`)**:
  - Lee directamente libros de datos en Excel (`.xlsx`) en `Datos_Excel/`.
  - Procesa automáticamente las hojas semanales más recientes.
  - Exporta PDFs ejecutivos de 2 páginas organizados por número de semana en `Reportes_PDF/Semana_XX/`.
- **Detección Dinámica de Esquemas**: Ajuste automático de columnas para distintas empresas.
- **Imágenes e Historial de Logos**: Renderizado de logotipos vectoriales e imágenes institucionales en Base64.

---

## 🛠️ Estructura del Proyecto
```
Generador de Reportes Semanales/
├── assets/
│   └── logos/
│       ├── Camanchaca logo.png
│       ├── Cermaq logo.jpg
│       └── Mowi logo.png
├── Datos_Excel/
│   ├── Reporte Camanchaca 2026 semana 34.xlsx
│   ├── Reporte Cermaq 2026 Semana 34.xlsx
│   └── Reporte Mowi 2026 Semana 34.xlsx
├── Reportes_PDF/
│   ├── Semana_33/
│   └── Semana_34/
├── generar_reporte.ps1
└── reporte_semanal.html
```

---

## 📖 Uso

### 1. Generación de PDFs por consola (PowerShell)
```powershell
.\generar_reporte.ps1 -Empresa Camanchaca
.\generar_reporte.ps1 -Empresa Cermaq
.\generar_reporte.ps1 -Empresa Mowi
```

### 2. Vista previa e impresión desde el navegador
Abre `reporte_semanal.html` en Chrome/Edge, selecciona la empresa deseada, edita las observaciones en tiempo real y presiona **🖨️ Guardar PDF / Imprimir** (o `Ctrl + P`).

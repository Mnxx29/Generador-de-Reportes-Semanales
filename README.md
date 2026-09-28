# 📊 Generador de Reportes Semanales de Monitoreo

Sistema para leer los archivos Excel semanales de Camanchaca, Cermaq y Mowi, calcular métricas operativas, actualizar un tablero HTML local y generar reportes PDF.

## Requisitos

- Windows PowerShell 5.1 o PowerShell 7.
- Microsoft Excel de escritorio, utilizado mediante automatización COM.
- Microsoft Edge para la generación directa de PDF.

## Uso semanal

1. Copia los Excel actualizados dentro de `Datos_Excel/`. El nombre debe incluir la empresa, el año y la semana; por ejemplo, `Reporte Camanchaca 2027 Semana 12.xlsx`.
2. Ejecuta el generador:

   ```powershell
   .\generar_reporte.ps1
   ```

3. Abre `reporte_semanal.html`, revisa las métricas y usa **Guardar PDF / Imprimir** si necesitas exportarlo desde el navegador.

El script selecciona automáticamente el archivo más reciente de cada empresa y la primera pestaña distinta de `Consolidado`.

También puedes procesar una sola empresa:

```powershell
.\generar_reporte.ps1 -Empresa Cermaq
```

## Generación directa de PDF

```powershell
.\generar_reporte.ps1 -GenerarPDF
```

Los documentos se guardan en `Reportes_PDF/Semana_XX/`. El año y la semana se obtienen del nombre del Excel o de la pestaña seleccionada.

## Privacidad de los datos

`reporte_semanal.template.html` contiene únicamente la estructura visual y no incluye información operacional. Al ejecutar el script se crea `reporte_semanal.html` con los datos de los Excel.

Los siguientes elementos se mantienen fuera de Git porque pueden contener información sensible:

- Archivos Excel.
- Reportes PDF.
- `reporte_semanal.html` generado.
- Archivos temporales de renderizado.

Antes de publicar cambios, comprueba siempre `git status` para confirmar que solo se estén incluyendo código, documentación, plantilla y recursos gráficos.

## Estructura

```text
Generador de Reportes Semanales/
├── assets/
│   └── logos/
├── Datos_Excel/                    # Entrada local; ignorada por Git
├── Reportes_PDF/                   # Salida local; ignorada por Git
├── generar_reporte.ps1             # Procesamiento y generación
├── reporte_semanal.template.html   # Plantilla versionada sin datos reales
└── reporte_semanal.html            # Salida generada; ignorada por Git
```

# 📊 Generador de Reportes Semanales de Monitoreo

Sistema para leer los archivos Excel semanales de Camanchaca, Cermaq y Mowi, calcular métricas operativas, actualizar un tablero HTML local y generar reportes PDF.

El reporte incluye métricas globales y regionales, estado de cámaras, detalle por centro, observaciones y un resumen ejecutivo generado automáticamente.

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

`reporte_semanal.html` no viene incluido al clonar el repositorio: se crea localmente durante la ejecución porque contiene datos operativos.

El script selecciona automáticamente el archivo más reciente de cada empresa y la primera pestaña distinta de `Consolidado`. De forma predeterminada procesa las tres empresas.

### Comandos disponibles

| Objetivo | Comando |
|---|---|
| Actualizar el HTML con las tres empresas | `.\generar_reporte.ps1` |
| Actualizar solo una empresa | `.\generar_reporte.ps1 -Empresa Cermaq` |
| Generar los tres PDF directamente | `.\generar_reporte.ps1 -GenerarPDF` |
| Generar el PDF de una empresa | `.\generar_reporte.ps1 -Empresa Mowi -GenerarPDF` |

Los valores aceptados por `-Empresa` son `Camanchaca`, `Cermaq`, `Mowi` y `Todas`.

## Formato esperado de los Excel

- El nombre debe contener la empresa y `Semana XX`.
- Se recomienda incluir también el año, por ejemplo `Reporte Camanchaca 2027 Semana 12.xlsx`.
- La primera pestaña distinta de `Consolidado` debe corresponder al período que se desea procesar.
- La fila de encabezados debe estar entre las primeras 10 filas y contener una columna cuyo nombre incluya `Centro`.
- El generador reconoce encabezados equivalentes a `Región`, `Total de jaulas`, `Cámara por jaula`, `Sin visual`, `Mortalidad sin visual` y `Observaciones`.
- Una observación vacía o con `-` se transforma en `Sin novedad`.

Si no se puede obtener la semana desde el nombre del archivo ni desde el nombre de la pestaña, el proceso se detiene para evitar generar un reporte con una semana incorrecta.

## Generación directa de PDF

```powershell
.\generar_reporte.ps1 -GenerarPDF
```

Los documentos se guardan en `Reportes_PDF/Semana_XX/`. El año y la semana se obtienen del nombre del Excel o de la pestaña seleccionada.

## Solución de problemas

### PowerShell impide ejecutar el script

Habilita la ejecución únicamente para la sesión actual y vuelve a intentarlo:

```powershell
Set-ExecutionPolicy -Scope Process Bypass
.\generar_reporte.ps1
```

### No se puede iniciar Excel

Comprueba que Microsoft Excel de escritorio esté instalado. Cierra los libros abiertos y elimina únicamente los archivos temporales `~$...xlsx` que hayan quedado bloqueados después de confirmar que Excel está cerrado.

### No se genera el PDF

Comprueba que Microsoft Edge esté instalado en una ubicación estándar o disponible como `msedge.exe` en `PATH`. También puedes abrir `reporte_semanal.html` y usar la opción de impresión del navegador.

### No aparece una empresa en el reporte

Verifica que exista un archivo `.xlsx` con el nombre de la empresa dentro de `Datos_Excel/` y que sus encabezados respeten el formato descrito anteriormente.

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

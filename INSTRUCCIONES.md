# 📖 Guía de Uso Paso a Paso - Reportes Semanales

Sigue estos sencillos pasos para actualizar la vista web interactiva y previsualizar los reportes antes de guardarlos como PDF (**Camanchaca**, **Cermaq**, **Mowi**).

---

## 🔄 Flujo de Trabajo Semanal

```
[1. Copiar Excels] ➔ [2. Ejecutar .\generar_reporte.ps1] ➔ [3. Abrir reporte_semanal.html] ➔ [4. Imprimir a PDF]
```

---

### 📋 Paso 1: Guardar los nuevos archivos Excel
Cuando tu colega te entregue los 3 archivos Excel actualizados de la semana (por ejemplo, de la **Semana 35**):

1. Abre la carpeta del proyecto: `Generador de Reportes Semanales`
2. Entra a la carpeta: `Datos_Excel/`
3. Copia y pega los 3 nuevos archivos Excel ahí.

> 💡 **Nota:** No es necesario borrar los archivos de semanas anteriores. El sistema detecta automáticamente el archivo más reciente para cada empresa.

---

### ⚡ Paso 2: Actualizar la Web desde PowerShell
1. Abre **PowerShell** en la carpeta del proyecto (o haz clic derecho dentro de la carpeta y selecciona *"Abrir en Terminal / PowerShell"*).
2. Ejecuta el siguiente comando:

```powershell
.\generar_reporte.ps1
```

### ¿Qué hace este comando?
- Lee la información más reciente de cada empresa en `Datos_Excel/`.
- Calcula centros, jaulas, cámaras y porcentajes de funcionamiento por región.
- **Actualiza la vista web interactiva [`reporte_semanal.html`](./reporte_semanal.html)** de forma limpia y sin errores de sintaxis.

---

### 🌐 Paso 3: Revisar el reporte e Imprimir / Guardar en PDF
1. Haz doble clic en el archivo [`reporte_semanal.html`](./reporte_semanal.html) para abrirlo en tu navegador (Edge o Chrome).
2. Usa los botones de la barra superior (**Camanchaca**, **Cermaq**, **Mowi**) para navegar e inspeccionar la información de cada empresa.
3. Si deseas modificar alguna observación por centro o agregar notas ejecutivas antes de guardar, puedes editar el texto directamente en la pantalla.
4. Cuando estés conforme con la previsualización, haz clic en el botón verde **🖨️ Guardar PDF / Imprimir** en la barra superior (o usa `Ctrl + P`).

---

### 💡 Consejo Extra (Generar PDFs por Consola)
Si en alguna ocasión no necesitas revisar la web y prefieres que el script genere los archivos PDF directamente en `Reportes_PDF/Semana_XX/`, ejecuta:

```powershell
.\generar_reporte.ps1 -GenerarPDF
```

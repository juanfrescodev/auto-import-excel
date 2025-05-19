# 🛠 Automatización silenciosa de importación de datos en Excel

Este proyecto resuelve una necesidad real en una emisora de radio: importar de forma automática, diaria y **silenciosa** datos de canciones reproducidas a una planilla Excel con macros, sin interrumpir al operador ni generar alertas o ventanas visibles.

---

## 🎯 Objetivo

Automatizar la carga de datos desde ZaraRadio a una planilla Excel con macros (`adi.xlsm`), ejecutando todo el proceso de forma invisible mediante VBScript y tareas programadas de Windows.

---

## 🧩 Componentes del proyecto

- `auto_import.vbs`: Script principal que abre el archivo Excel, ejecuta la macro y cierra todo sin mostrar interfaz.
- `ImportarCancionesZaraRadio.bas`: Módulo exportado de la macro en Excel.
- `tareas_programadas.md`: Guía paso a paso para programar la ejecución automática con el Programador de tareas de Windows.
- Capturas de pantalla: Muestran la programación de la tarea y el entorno de archivos.

---

## ⚙️ Tecnologías utilizadas

- **VBScript (.vbs)** para ejecutar procesos en segundo plano
- **VBA (Excel macros)** para importar y procesar datos
- **Tareas Programadas de Windows** para automatizar la ejecución diaria
- Silenciamiento de alertas (`DisplayAlerts = False`, `Workbook.Saved = True`)
- Manejo de argumentos (`Command = "auto"`) para controlar ejecución manual vs automática

---

## 📌 Resultados

- Uso diario desde marzo de 2025 en Radio Nacional Bariloche.
- 100% invisible para el operador.
- Evita errores humanos.
- Asegura cumplimiento con reportes para AADI-CAPIF.
- Ahorra más de 30 minutos diarios al equipo operativo.

---

## 📂 Cómo usar este proyecto

1. Copiar `auto_import.vbs` al mismo directorio donde se encuentra el archivo Excel con macros.
2. Asegurarse de que la macro se llama `ImportarCancionesZaraRadio`.
3. Programar la tarea diaria en Windows con:
   - Ejecutable: `C:\Windows\System32\wscript.exe`
   - Argumentos: `"C:\ruta\a\auto_import.vbs" auto`
4. Confirmar que todo se ejecuta correctamente en segundo plano.

---

## 👨‍💻 Sobre el autor

Este proyecto forma parte del portfolio de [JuanFrescoDev](https://github.com/juanfrescodev), enfocado en automatización y análisis de datos aplicados a necesidades reales.  

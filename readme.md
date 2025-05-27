# 🧾 Automatización silenciosa de importación de datos en Excel

Este proyecto permite automatizar la ejecución de una macro en Excel (`.xlsm`) mediante un script `.vbs` que corre en segundo plano desde el Programador de tareas de Windows. El objetivo es lograr un proceso **100% invisible**, sin interrumpir al operador.

---

## ⚙️ ¿Qué hace?

- Abre un archivo Excel con macros **de forma oculta** (sin mostrar la interfaz)
- Ejecuta automáticamente una macro que **importa datos**
- Cierra el archivo sin guardar (el guardado lo realiza la macro)
- Se ejecuta desde el **Programador de tareas de Windows**, sin generar alertas ni ventanas

---

## 📁 Archivos principales

| Archivo                          | Descripción |
|----------------------------------|-------------|
| `auto_import.vbs`                | Script que ejecuta Excel en segundo plano y llama a la macro |
| `ImportarCancionesZaraRadio.bas`| Código fuente de la macro en VBA exportado desde Excel |
| `tareas_programadas.md`          | Guía paso a paso para configurar la tarea automática en Windows |
| `captura_1.png`                  | Captura de la planilla generada con la macro |

---

## 📸 Captura de pantalla

![Captura del programador de tareas](captura_1.png)

---

## 🧠 Motivación

Este desarrollo surgió de una necesidad real en un entorno de radio AM, donde era clave automatizar tareas administrativas sin interrumpir el trabajo del operador.  
La solución fue diseñada para **integrarse sin fricciones** al flujo de trabajo diario y ejecutarse en segundo plano.

---

## 🛠️ Requisitos

- Windows con Programador de tareas habilitado
- Microsoft Excel (compatible con macros)
- Habilitación de macros en el archivo `.xlsm`

---

## 🚀 Uso

1. Configurar el archivo `auto_import.vbs` con la ruta del Excel que contiene la macro
2. Programar su ejecución automática con el archivo `tareas_programadas.md` como guía
3. Asegurarse de que el archivo Excel esté habilitado para macros

---

## 💼 Aplicación en portfolio

Aunque fue hecho para un caso específico (Zara Radio), representa conocimientos transferibles como:

- Automatización de flujos de trabajo con herramientas de oficina
- Integración entre lenguajes (VBS + VBA + Excel)
- Diseño de soluciones **invisibles al usuario**
- Pensamiento orientado a operaciones reales

---

## 🧑‍💻 Autor

Juan Fresco - [juanfrescodev.github.io](https://juanfrescodev.github.io)

---

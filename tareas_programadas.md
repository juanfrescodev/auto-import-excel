# ⏰ Cómo programar la ejecución automática con el Programador de tareas de Windows

Este archivo describe paso a paso cómo configurar una tarea programada para ejecutar automáticamente el script `auto_import.vbs`, que abre una planilla Excel con macros, ejecuta la importación y cierra todo en segundo plano.

---

## ✅ Requisitos previos

- Archivo Excel con macro guardado como `.xlsm`
- Script `auto_import.vbs` ubicado en la misma carpeta
- Nombre de la macro: `ImportarCancionesZaraRadio`
- Windows (con Programador de tareas)

---

## 🧭 Pasos para crear la tarea programada

1. **Abrí el Programador de Tareas**
   - Presioná `Win + S` y escribí `Programador de tareas`
   - Seleccioná "Crear tarea..." (no "Crear tarea básica")

---

2. **General**
   - **Nombre**: `AutoImport AADI CAPIF`
   - **Descripción**: Ejecuta importación de canciones desde ZaraRadio en segundo plano
   - **Ejecutar si el usuario inició sesión o no**
   - Marcar: `Ejecutar con los privilegios más altos`
   - **Configuración para:** Windows 10 o superior

---

3. **Desencadenadores (Triggers)**
   - Nuevo desencadenador
   - **Iniciar tarea**: Según una programación
   - **Diariamente**
     - Hora: (ej. `05:00 a.m.`)
   - Marcar “Habilitado”

---

4. **Acciones**
   - **Acción**: Iniciar un programa
   - **Programa o script**:
     ```
     C:\Windows\System32\wscript.exe
     ```
   - **Agregar argumentos**:
     ```
     "C:\Aire AM\09 Planilla AADI CAPIF\2025\05 mayo\auto_import.vbs" auto
     ```
   - **Iniciar en (opcional)**:
     ```
     C:\Aire AM\09 Planilla AADI CAPIF\2025\05 mayo
     ```

---

5. **Condiciones**
   - (Opcional) Desactivar todo si querés que se ejecute incluso con batería

---

6. **Configuración**
   - Marcar: Permitir que se ejecute a pedido
   - Marcar: Detener la tarea si dura más de 30 minutos

---

7. **Guardar la tarea**
   - Aceptar y confirmar con tu contraseña de administrador si la pide.

---

## 🧪 Probar la tarea manualmente

1. Ir a la lista de tareas.
2. Click derecho en la tarea → Ejecutar.
3. Verificá que **no se abra Excel ni haya alertas**, pero que se haya actualizado la planilla correctamente.

---

## 🛑 Sugerencia para evitar errores

- El script usa `Command = "auto"` para distinguir ejecución automática.
- En caso de querer ejecutar manualmente con Excel visible, quitá ese argumento o desactivá esa verificación.

---

## 👀 ¿Qué hace el script exactamente?

- Abre Excel en segundo plano (`Visible = False`)
- Ejecuta la macro `ImportarCancionesZaraRadio`
- Cierra sin guardar ni mostrar alertas
- Libera recursos

---

## 🧑‍💻 Autor

Este proyecto fue implementado y usado en operación real en una emisora de radio AM desde marzo de 2025.  
Forma parte del portfolio de [JuanFrescoDev](https://github.com/juanfrescodev).

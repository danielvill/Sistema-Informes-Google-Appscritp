
🚀 Sistema de Informes con Google Apps Script
¡Bienvenido al sistema de informes! Este proyecto está construido utilizando Google Apps Script para automatizar la gestión y generación de informes a través de Google Sheets.

🎥 Video de Ejecución: Puedes seguir este [enlace al video] para ver una demostración de cómo se ejecuta el sistema.

⚠️ ¡Atención! Puntos Clave para la Ejecución
Es CRUCIAL seguir las instrucciones a continuación, especialmente en lo referente a los nombres de los archivos y carpetas. Si no tienes experiencia previa en programación, te recomendamos seguir los nombres al pie de la letra para asegurar el correcto funcionamiento del sistema.

🛠️ Instrucciones de Configuración y Ejecución
Sigue estos 10 pasos detallados para configurar y probar tu sistema de informes:

1. Preparación en Google Drive 📁
Crea una carpeta en Google Drive y nómbrala exactamente: Sistema informes

Dentro de esta carpeta, crea una nueva Hoja de Cálculo de Google Sheets y nómbrala exactamente: Informe

2. Configuración de la Hoja de Cálculo 📊
Dentro del archivo Informe, asegúrate de configurar las tablas y encabezados que correspondan a la estructura de datos que utilizará el sistema.

3. Implementación del Formulario (Front-end) ✍️
Abre el archivo Informe. En la barra de herramientas de Google, haz clic en Extensiones y luego en Apps Script.

En el editor de Apps Script:

Copia y pega el código del archivo Codigo.gs (de la carpeta Formulario) en el archivo Código.gs existente o crea uno nuevo.

Crea un nuevo archivo de tipo HTML (p. ej., haciendo clic en el signo + junto a Archivos) y nómbralo exactamente: form.html. Pega el código correspondiente del archivo form.html en este nuevo archivo.

4. Configuración del Respaldo (Backup) 💾
En tu Google Drive (en la carpeta Sistema informes), crea una carpeta llamada exactamente: Respaldo.

Dentro de la carpeta Respaldo, crea un archivo de Google Apps Script independiente y llámalo exactamente: Respaldo.

Pega el código de la carpeta Respaldo en este nuevo script.

⚠️ ¡Dato Importante! Debes agregar el ID de la Hoja de Cálculo de Google llamada Informe en la parte del código de Respaldo que lo requiere. Busca la sección correspondiente dentro del código de respaldo.

5. Configuración del Sistema Principal ⚙️
En tu Google Drive (en la carpeta Sistema informes), crea un archivo de Google Apps Script independiente y nómbralo exactamente: Sistema informes.

En este nuevo script Sistema informes:

Copia y pega el código del archivo Codigo.gs que se encuentra en la carpeta principal (Sistema informes).

Crea los archivos HTML correspondientes (.html) y pega el código de cada uno de ellos.

6. Verificación Final y Prueba ✅
Revisa los permisos y los nombres de todos los archivos y carpetas: Sistema informes (Carpeta), Informe (Sheet), Respaldo (Carpeta), y los tres proyectos de Apps Script: el de Informe (Formulario), Respaldo, y Sistema informes.

⚠️ Permisos: Para cada uno de los archivos de Google Apps Script, debes dar la autorización la primera vez que se ejecute una función. Este paso de autorización se repite en cada script (Informe, Respaldo, Sistema informes).

⚠️ Dato importante para el sheet: se debe colocar esta formula para la ejecucion
   =FILTER(publicadores!B2:C; 
  (CONTAR.SI.CONJUNTO(form!A:A; publicadores!B2:B; form!G:G; publicadores!C2:C)=0) * 
  (publicadores!D2:D <> "Inactivo")
)
   Y tambien que en la tabla (clave) se debe seguir un dato para que funcione la validacion pero en el video aparece

Prueba el sistema para verificar su correcta funcionalidad.
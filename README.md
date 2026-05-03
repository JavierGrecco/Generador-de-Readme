# Generador-de-Readmy
Genera README.md profesionales desde archivos .txt o .docx. Detecta tecnologías, agrega badges, tabla de contenidos y emojis. Funciona 100% offline, sin dependencias de APIs externas. Ideal para documentar proyectos en GitHub de forma rápida y automatizada.

## :sparkles: 📋 Formato de entrada esperado por el script
El script espera un archivo .txt o .docx con la siguiente estructura:
Título: la primera línea debe ser # Nombre del Proyecto.
Secciones: usar ## Título de la sección (sin emojis en el input, el script los añade automáticamente).
Código: los comandos sueltos (cmd, bash, text) deben escribirse como líneas independientes, sin fences. El script los envuelve automáticamente.
Tablas: usar pipes | y guiones --- para el separador.
Árboles de directorio: usar caracteres ├──, │, └──.
Listas: usar - o * al inicio de cada ítem.
Sin badges previos: el script los genera automáticamente.
Sin TOC manual: el script la genera automáticamente.
## :sparkles:


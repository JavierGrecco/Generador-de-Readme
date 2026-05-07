# Generador de README.md [![PyPI version](https://badge.fury.io/py/generar-readme.svg)](https://pypi.org/project/generar-readme/)

Herramienta offline que transforma archivos .txt o .docx en un README.md profesional para GitHub. Extrae la estructura del documento, detecta tecnologías, agrega insignias (badges), tabla de contenidos y emojis, generando un resultado limpio y listo para publicar.

## Formato de entrada esperado

El script espera un archivo .txt o .docx con la siguiente estructura:

- Título: la primera línea debe ser # Nombre del Proyecto.
- Secciones: usar ## Título de la sección (sin emojis en el input, el script los añade automáticamente).
- Código: los comandos sueltos (cmd, bash, text) deben escribirse como líneas independientes, sin fences. El script los envuelve automáticamente.
- Tablas: usar pipes | y guiones --- para el separador.
- Árboles de directorio: usar caracteres ├──, │, └──.
- Listas: usar - o * al inicio de cada ítem.
- Sin badges previos: el script los genera automáticamente.
- Sin TOC manual: el script la genera automáticamente.

## Cómo funciona

El script procesa el documento de entrada en cuatro etapas:

1.  Limpieza inicial – Elimina líneas espurias y caracteres de escape introducidos por editores como Word.
2.  Normalización estructural – Convierte comandos sueltos (cmd, bash, text) en bloques de código Markdown y fuerza encabezados donde corresponde.
3.  Construcción del AST – Parsea el texto en un árbol sintáctico que representa secciones, párrafos, tablas, listas y bloques de código.
4.  Enriquecimiento y renderizado – Aplica emojis, añade secciones obligatorias si faltan, genera la tabla de contenidos y produce el Markdown final.

El motor funciona completamente offline. No depende de APIs externas ni de inteligencia artificial generativa.

## Características

- Convierte documentos .txt o .docx a README.md profesional.
- Agrega badges automáticos de tecnologías detectadas (Python, Windows, Linux, Excel, PDF…).
- Genera una tabla de contenidos a partir de los encabezados reales.
- Asigna emojis representativos a cada sección.
- Encapsula comandos sueltos (cmd, bash, text) en bloques de código.
- Conserva tablas con formato de pipes y árboles de directorios.
- Incluye watch mode para regenerar automáticamente al guardar cambios.
- Registra la actividad en un archivo generar_readme.log.
- Modo --debug para visualizar el AST generado.
- Validación de entrada: verifica que el archivo exista, sea legible y no esté vacío.
- Opciones de línea de comandos para omitir TOC, créditos o forzar la licencia.
- Instalable como paquete PyPI con pip install.
- Código modular y documentado, fácil de mantener.

## Requisitos

| Componente | Descripción |
|-----------|-------------|
| Python | 3.8 o superior |
| Dependencia opcional | python-docx (solo para leer archivos .docx) |
| Dependencias de test | pytest, hypothesis (solo para ejecutar los tests) |

## Instalación

**Desde PyPI (recomendado):**

bash
pip install generar-readme

Esto instalará automáticamente las dependencias necesarias.

**Desde el repositorio:**

Clona el repositorio e instala las dependencias manualmente:

bash
git clone https://github.com/JavierGrecco/generador-readme
cd generador-readme
pip install -r requirements.txt

### Instalación por sistema operativo

**Windows**

Instalar Python desde python.org (marcar "Add Python to PATH")

cmd
python --version
pip install python-docx

**Linux (Debian/Ubuntu)**

bash
sudo apt update
sudo apt install python3 python3-pip
pip3 install python-docx

**macOS**

bash
brew install python@3.10
pip3 install python-docx

## Uso

**Si instalaste desde PyPI**, el comando `generar-readme` estará disponible globalmente. Ejecutalo en el directorio donde tengas tu archivo .txt o .docx:

- Usar el paquete instalado: `generar-readme build --auto`
- Usar el script directamente: `python generar_readme.py build --auto`

Esto detectará automáticamente el archivo fuente y generará README.md.

### Opciones disponibles

| Opción | Descripción |
|--------|-------------|
| --txt ruta | Especifica un archivo .txt de entrada. |
| --docx ruta | Especifica un archivo .docx de entrada. |
| --auto | Detecta automáticamente el archivo en el directorio actual. |
| -o, --output | Nombre del archivo de salida (por defecto README.md). |
| --license MIT | Fuerza el badge de una licencia (MIT, GPL, Apache…). |
| --logo ruta | Inserta un logotipo centrado bajo el título. |
| --no-toc | Omite la tabla de contenidos. |
| --no-credits | Omite la sección de créditos. |
| --no-mandatory | No inserta secciones obligatorias faltantes. |
| --watch | Activa el modo observador (regenera al guardar cambios). |
| --debug | Imprime el AST generado en consola para depuración. |

### Ejemplos

bash
# Usar un archivo .txt concreto
python generar_readme.py build --txt "documento.txt"

# Forzar licencia MIT y añadir logo
python generar_readme.py build --auto --license MIT --logo "assets/logo.png"

# Modo watch (regenera automáticamente al guardar el fuente)
python generar_readme.py build --txt "doc.txt" --watch

# Depurar la estructura del documento
python generar_readme.py build --auto --debug

## Estructura de archivos

text
.
├── generar_readme.py          # Script principal
├── generar_readme.log         # Registro de ejecución (se crea automáticamente)
├── test_generar_readme.py     # Tests unitarios y de propiedad (pytest + hypothesis)
├── pyproject.toml             # Configuración para publicación en PyPI
├── requirements.txt           # Dependencias del proyecto
├── Fotos/                     # Carpeta opcional para logotipo
│   └── logo.png
└── README.md                  # Este archivo

## Logging

Cada ejecución queda registrada en generar_readme.log con marcas de tiempo y niveles de depuración. Esto permite auditar el proceso y diagnosticar problemas rápidamente.

## Tests

El proyecto incluye un archivo de tests independiente (test_generar_readme.py) que no es necesario para usar el generador. Su función es verificar que cada cambio en el código no rompa funcionalidades que ya estaban funcionando.

El archivo contiene pruebas automatizadas que cubren:

- **Tests de integración**: verifican el pipeline completo con casos reales que fallaron durante el desarrollo (títulos duplicados, comandos sueltos, líneas de licencia, etc.).
- **Tests de unidad**: prueban funciones internas de forma aislada (sanitizar, forzar_titulos, detectar_bloques, construir_ast).
- **Tests de propiedad** (con Hypothesis): generan cientos de entradas aleatorias y verifican que se mantengan propiedades invariantes del README: que los fences estén balanceados, que el título aparezca una sola vez y que los créditos nunca falten.
- **Tests de detección automática**: validan la funcionalidad --auto con diferentes configuraciones de archivos.
- **Tests de watch mode**: verifican que el script regenere el README al modificar el archivo fuente (este test es pesado y se ejecuta bajo la marca `pytest -m watch`).
- **Casos límite**: validan edge cases descubiertos durante el desarrollo, como comentarios bash dentro de fences o secciones agrupadoras vacías.

Para ejecutar los tests, instalá las dependencias necesarias y corré pytest desde la misma carpeta donde está el script principal:

bash
pip install pytest hypothesis
python -m pytest test_generar_readme.py -v

Para omitir los tests pesados de watch mode:

bash
pytest test_generar_readme.py -v -m "not watch"

En macOS o Linux, reemplazá `pip` por `pip3` y `python` por `python3` si es necesario.

También podés ejecutar solo una categoría específica:

bash
pytest test_generar_readme.py -v -k "TestPipeline"       # solo integración
pytest test_generar_readme.py -v -k "TestPropiedades"    # solo property-based
pytest test_generar_readme.py -v -k "TestAutoMode"       # solo detección automática

## Licencia

MIT © Javier Grecco – github.com/JavierGrecco
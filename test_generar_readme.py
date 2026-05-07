"""
Tests completos para generar_readme.py

Ejecutar con:
    python -m pytest test_generar_readme.py -v

Marcas disponibles:
    pytest -m "not watch"   # excluye el test pesado de --watch
"""

import os
import sys
import time
import pytest
import tempfile
import subprocess
from pathlib import Path
from unittest.mock import patch
from hypothesis import given, strategies as st, settings

# Asegurar que el módulo principal se importe correctamente
sys.path.insert(0, os.path.dirname(__file__))
from generar_readme import (
    sanitizar,
    forzar_titulos,
    envolver_comandos,
    detectar_bloques,
    construir_ast,
    renderizar_md,
    es_header_semantico,
    buscar_archivo,
    validar_archivo,
    leer_txt,
    leer_docx,
    construir_readme,
    main,
    Nodo,
    INDICADORES_CODIGO,
    FORCED_HEADERS,
    PALABRAS_PROHIBIDAS,
)


# ----------------------------------------------------------------------
# Fixtures
# ----------------------------------------------------------------------
@pytest.fixture
def directorio_temporal():
    """Cambia al directorio temporal y luego regresa al original."""
    old_cwd = os.getcwd()
    with tempfile.TemporaryDirectory() as tmp:
        os.chdir(tmp)
        yield Path(tmp)
    os.chdir(old_cwd)


@pytest.fixture
def txt_input():
    """Contenido típico de un archivo .txt bien formado."""
    return (
        "# Mi Proyecto\n\n"
        "## Características\n\n"
        "- Punto 1\n"
        "- Punto 2\n\n"
        "## Instalación\n\n"
        "bash\n"
        "pip install mi_paquete\n"
        "python setup.py install\n\n"
        "## Uso\n\n"
        "text\n"
        "Árbol de ejemplo:\n"
        "├── src\n"
        "│   └── main.py\n"
        "└── README.md\n\n"
        "## Licencia\n\n"
        "MIT © alguien"
    )


# ----------------------------------------------------------------------
# Tests de unidad
# ----------------------------------------------------------------------
class TestSanitizar:
    def test_elimina_badges(self):
        texto = "![](https://img.shields.io/badge/...)\nHola"
        assert "Hola" in sanitizar(texto)
        assert "shields.io" not in sanitizar(texto)

    def test_elimina_lineas_copiar_descargar(self):
        texto = "Copiar\nDescargar\nContenido"
        res = sanitizar(texto)
        assert "Copiar" not in res
        assert "Descargar" not in res
        assert "Contenido" in res


class TestEsHeaderSemantico:
    def test_lista_no_es_header(self):
        assert not es_header_semantico("- Dos versiones de salida:")
        assert not es_header_semantico("* Item")

    def test_header_comun(self):
        assert es_header_semantico("Instalación")
        assert es_header_semantico("Uso del programa")

    def test_linea_prohibida(self):
        assert not es_header_semantico("pip install algo")

    def test_tabla_no_es_header(self):
        assert not es_header_semantico("| Col1 | Col2 |")

    def test_licencia_no_es_header(self):
        assert not es_header_semantico("MIT © github.com/...")


class TestForzarTitulos:
    def test_convierte_titulos_conocidos(self):
        texto = "Instalación\npip install x"
        resultado = forzar_titulos(texto)
        assert "## Instalación" in resultado

    def test_respeta_headers_existentes(self):
        texto = "## Ya existente\ntexto"
        resultado = forzar_titulos(texto)
        assert resultado.startswith("## Ya existente")

    def test_inserta_instalacion_si_huerfanos(self):
        texto = "## Windows\npip install"
        resultado = forzar_titulos(texto)
        lineas = resultado.splitlines()
        assert any("## Instalación" in l for l in lineas)


class TestEnvolverComandos:
    def test_bash_sueltos(self):
        texto = "bash\nsudo apt update\nsudo apt install"
        resultado = envolver_comandos(texto)
        assert "```bash" in resultado
        assert "sudo apt update" in resultado
        assert resultado.count("```") == 2

    def test_text_block(self):
        texto = "text\nlínea1\nlínea2"
        resultado = envolver_comandos(texto)
        assert "```text" in resultado
        assert "línea1" in resultado

    def test_no_afecta_fences_existentes(self):
        texto = "```bash\npip install\n```"
        resultado = envolver_comandos(texto)
        assert resultado.count("```bash") == 1


# ----------------------------------------------------------------------
# Integración: pipeline completo con entradas conocidas
# ----------------------------------------------------------------------
class TestPipeline:
    def test_readme_completo_txt(self, txt_input):
        argumentos = type("Args", (), {
            "debug": False,
            "no_toc": False,
            "no_creditos": False,
            "license": None,
            "logo": None,
        })()
        resultado = construir_readme(txt_input, argumentos)
        # Verifica presencia de elementos clave
        assert "# Mi Proyecto" in resultado
        assert "## 📋 Tabla de Contenidos" in resultado
        assert "## :sparkles: Características" in resultado
        assert "## :wrench: Instalación" in resultado
        assert "## :rocket: Uso" in resultado
        assert "## :scroll: Licencia" in resultado
        assert resultado.count("```") % 2 == 0  # fences balanceados

    def test_titulo_duplicado_no_rompe(self):
        contenido = "# Proyecto\n\n## Características\n\n- Item 1\n\n## Características\n\n- Item 2"
        args = type("Args", (), {"debug": False, "no_toc": False, "no_creditos": False, "license": None, "logo": None})()
        resultado = construir_readme(contenido, args)
        # La sección duplicada se elimina; solo una aparece y conserva sus hijos
        assert resultado.count("## :sparkles: Características") == 1
        assert "- Item 1" in resultado

    def test_licencia_detectada_automatica(self):
        contenido = "# Test\n\nMIT License"
        args = type("Args", (), {"debug": False, "no_toc": False, "no_creditos": False, "license": None, "logo": None})()
        resultado = construir_readme(contenido, args)
        assert "MIT-yellow" in resultado


# ----------------------------------------------------------------------
# Propiedades con Hypothesis (invariantes)
# ----------------------------------------------------------------------
class TestPropiedades:
    @given(st.text(alphabet=st.characters(
        blacklist_categories=('Cs',),  # no surrogates
        blacklist_characters=['`', '~']  # evitar fences falsos
    ), max_size=2000))
    @settings(max_examples=50, deadline=None)
    def test_fences_siempre_balanceados(self, raw_text):
        contenido = f"# Proyecto\n\n{raw_text}"
        args = type("Args", (), {"debug": False, "no_toc": False, "no_creditos": False, "license": None, "logo": None})()
        resultado = construir_readme(contenido, args)
        assert resultado.count("```") % 2 == 0, "Fences desbalanceados"

    @given(st.text(max_size=1000))
    @settings(max_examples=30, deadline=None)
    def test_titulo_aparece_solo_una_vez(self, raw_text):
        contenido = f"# Mi Titulo\n\n{raw_text}"
        args = type("Args", (), {"debug": False, "no_toc": False, "no_creditos": False, "license": None, "logo": None})()
        resultado = construir_readme(contenido, args)
        assert resultado.count("# Mi Titulo\n") == 1

    @given(st.text(max_size=1000))
    @settings(max_examples=30, deadline=None)
    def test_creditos_siempre_presentes(self, raw_text):
        contenido = f"# Proyecto\n\n{raw_text}"
        args = type("Args", (), {"debug": False, "no_toc": False, "no_creditos": False, "license": None, "logo": None})()
        resultado = construir_readme(contenido, args)
        assert "## ⭐ Créditos" in resultado


# ----------------------------------------------------------------------
# Edge cases
# ----------------------------------------------------------------------
class TestEdgeCases:
    def test_comentarios_bash_no_confunden(self):
        texto = "bash\n# esto es un comentario\necho hola"
        resultado = envolver_comandos(texto)
        # El comentario se queda dentro del bloque bash
        assert "# esto es un comentario" in resultado
        assert resultado.count("```bash") == 1

    def test_seccion_instalacion_vacia_se_mantiene(self):
        # Instalación es agrupadora, no se borra aunque esté vacía
        bloques = [("LINE", "## Instalación")]
        ast = construir_ast(bloques)
        from generar_readme import normalizar_ast
        normalizado = normalizar_ast(ast)
        titulos = [h.contenido for h in normalizado.hijos if h.tipo == "SECCION"]
        assert any("Instalación" in t for t in titulos)


# ----------------------------------------------------------------------
# Tests para --auto (buscar_archivo y main)
# ----------------------------------------------------------------------
class TestAutoMode:
    def test_buscar_prioridad_nombre(self, directorio_temporal):
        (directorio_temporal / "readme.txt").write_text("contenido")
        (directorio_temporal / "otro.txt").write_text("otro")
        resultado = buscar_archivo()
        assert resultado.name == "readme.txt"

    def test_buscar_mas_reciente(self, directorio_temporal):
        a = directorio_temporal / "a.txt"
        b = directorio_temporal / "b.txt"
        a.write_text("a")
        time.sleep(0.1)
        b.write_text("b")
        resultado = buscar_archivo()
        assert resultado.name == "b.txt"

    def test_buscar_sin_archivos(self, directorio_temporal):
        resultado = buscar_archivo()
        assert resultado is None

    def test_main_auto_genera_readme(self, directorio_temporal, monkeypatch, capsys):
        (directorio_temporal / "catalogo.txt").write_text("# Proyecto\n\n## Uso\n\ntext\nprueba auto\n")
        monkeypatch.setattr(sys, "argv", ["generar_readme.py", "build", "--auto", "-o", "README.md"])
        main()
        salida = Path("README.md")
        assert salida.exists()
        assert "prueba auto" in salida.read_text()


# ----------------------------------------------------------------------
# Test para --watch (se ejecuta solo bajo demanda)
# ----------------------------------------------------------------------
@pytest.mark.watch
@pytest.mark.timeout(15)
def test_watch_mode_regenera(tmp_path):
    entrada = tmp_path / "test.txt"
    entrada.write_text("# Inicial\n\n## Sección\n\ntext\nlinea1\n")
    salida = tmp_path / "README.md"

    proc = subprocess.Popen(
        [sys.executable, "-m", "generar_readme", "build", "--txt", str(entrada), "-o", str(salida), "--watch"],
        cwd=tmp_path,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE
    )
    try:
        time.sleep(2)
        assert salida.exists()
        contenido_inicial = salida.read_text()
        assert "linea1" in contenido_inicial

        entrada.write_text("# Modificado\n\n## Sección\n\ntext\nlinea2\n")
        time.sleep(3)
        contenido_modificado = salida.read_text()
        assert "linea2" in contenido_modificado
    finally:
        proc.terminate()
        proc.wait(timeout=5)
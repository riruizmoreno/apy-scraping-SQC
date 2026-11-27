"""Exporta a Excel las reglas de los perfiles Sonar way de Python."""
import getpass
import json
import re
import sys
import zipfile
from pathlib import Path
from urllib.error import HTTPError, URLError
from urllib.parse import urlencode
from urllib.request import Request, urlopen
from xml.etree import ElementTree as ET

API = "https://sonarcloud.io/api"
PERFILES = {
    "Core": "Sonar way (Core)",
    "Extended": "Sonar way (Extended)",
    "Comprehensive": "Sonar way (Comprehensive)",
}
CAMPOS = (
    "qualityProfile",
    "key",
    "name",
    "lang",
    "repo",
    "type",
    "severity",
    "status",
    "tags",
    "description",
    "params",
    "cleanCodeAttribute",
    "impacts",
)

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL_DOC = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_REL_PKG = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_CONTENT = "http://schemas.openxmlformats.org/package/2006/content-types"
ET.register_namespace("", NS_MAIN)
ET.register_namespace("r", NS_REL_DOC)


def _etiqueta(namespace, nombre):
    return f"{{{namespace}}}{nombre}"


def _solicitar(ruta, parametros, token):
    query = urlencode(parametros)
    solicitud = Request(
        f"{API}/{ruta}?{query}",
        headers={"Authorization": f"Bearer {token}", "Accept": "application/json"},
    )
    with urlopen(solicitud, timeout=30) as respuesta:
        return json.load(respuesta)


def _descargar_perfiles(org, token):
    datos = _solicitar(
        "qualityprofiles/search",
        {"organization": org, "language": "py", "ps": 500},
        token,
    )
    perfiles = datos.get("profiles")
    if not isinstance(perfiles, list):
        raise ValueError("La API no devolvio una lista valida de perfiles de Python.")
    if not perfiles:
        raise ValueError("La API no encontro perfiles de Python para esta organizacion.")
    return perfiles


def _normalizar_nombre(nombre):
    return re.sub(r"[^a-z0-9]", "", nombre.casefold())


def _buscar_perfiles(org, token):
    disponibles = _descargar_perfiles(org, token)
    seleccionados = {}
    for etiqueta, nombre in PERFILES.items():
        coincidencias = [
            perfil
            for perfil in disponibles
            if _normalizar_nombre(perfil.get("name", "")) == _normalizar_nombre(nombre)
        ]
        if len(coincidencias) != 1:
            raise ValueError(
                f"Se esperaba encontrar un unico perfil '{nombre}' para Python; "
                f"se encontraron {len(coincidencias)}."
            )
        perfil = coincidencias[0]
        if not perfil.get("key"):
            raise ValueError(f"El perfil '{nombre}' no contiene una clave.")
        seleccionados[etiqueta] = perfil
    return seleccionados


def _descargar_reglas(org, perfil, token):
    reglas = {}
    pagina, total = 1, None
    while total is None or len(reglas) < total:
        datos = _solicitar(
            "rules/search",
            {
                "organization": org,
                "languages": "py",
                "activation": "true",
                "qprofile": perfil["key"],
                "ps": 100,
                "p": pagina,
            },
            token,
        )
        actual = (datos.get("paging") or {}).get("total", datos.get("total"))
        if not isinstance(actual, int) or actual < 0:
            raise ValueError(
                f"La API no devolvio un total valido de reglas para '{perfil['name']}'."
            )
        if total is None:
            total = actual
        elif actual != total:
            raise ValueError(
                f"El total de reglas del perfil '{perfil['name']}' cambio durante la descarga."
            )
        lote = datos.get("rules")
        if not lote and len(reglas) < total:
            raise ValueError(f"La pagina {pagina} no devolvio reglas.")
        for regla in lote or []:
            clave = regla.get("key")
            if not clave or clave in reglas:
                raise ValueError(f"Clave de regla ausente o duplicada: {clave!r}")
            reglas[clave] = regla
        pagina += 1
    if len(reglas) != total:
        raise ValueError(
            f"Descarga incompleta para '{perfil['name']}': {len(reglas)} de {total} reglas."
        )
    return reglas


def _valor_celda(valor):
    if valor is None:
        return ""
    if isinstance(valor, (dict, list)):
        return json.dumps(valor, ensure_ascii=False, sort_keys=True)
    return str(valor)


def _letra_columna(numero):
    letras = ""
    while numero:
        numero, resto = divmod(numero - 1, 26)
        letras = chr(65 + resto) + letras
    return letras


def _crear_excel(destino, perfil, reglas):
    raiz = ET.Element(_etiqueta(NS_MAIN, "worksheet"))
    vistas = ET.SubElement(raiz, _etiqueta(NS_MAIN, "sheetViews"))
    vista = ET.SubElement(vistas, _etiqueta(NS_MAIN, "sheetView"), {"workbookViewId": "0"})
    ET.SubElement(
        vista,
        _etiqueta(NS_MAIN, "pane"),
        {"ySplit": "1", "topLeftCell": "A2", "activePane": "bottomLeft", "state": "frozen"},
    )
    ET.SubElement(raiz, _etiqueta(NS_MAIN, "sheetFormatPr"), {"defaultRowHeight": "15"})
    columnas = ET.SubElement(raiz, _etiqueta(NS_MAIN, "cols"))
    anchuras = (23, 42, 38, 9, 20, 15, 15, 14, 24, 60, 35, 24, 35)
    for indice, ancho in enumerate(anchuras, start=1):
        ET.SubElement(
            columnas,
            _etiqueta(NS_MAIN, "col"),
            {"min": str(indice), "max": str(indice), "width": str(ancho), "customWidth": "1"},
        )
    datos_hoja = ET.SubElement(raiz, _etiqueta(NS_MAIN, "sheetData"))

    filas = [CAMPOS]
    for clave in sorted(reglas):
        regla = reglas[clave]
        descripcion = regla.get("mdDesc") or regla.get("htmlDesc") or ""
        filas.append(
            (
                perfil["name"],
                regla.get("key", ""),
                regla.get("name", ""),
                regla.get("lang", ""),
                regla.get("repo", ""),
                regla.get("type", ""),
                regla.get("severity", ""),
                regla.get("status", ""),
                regla.get("tags", []),
                descripcion,
                regla.get("params", []),
                regla.get("cleanCodeAttribute", ""),
                regla.get("impacts", []),
            )
        )

    for numero_fila, valores in enumerate(filas, start=1):
        fila = ET.SubElement(datos_hoja, _etiqueta(NS_MAIN, "row"), {"r": str(numero_fila)})
        for numero_columna, valor in enumerate(valores, start=1):
            referencia = f"{_letra_columna(numero_columna)}{numero_fila}"
            texto = _valor_celda(valor)
            if len(texto) > 32767:
                sufijo = " [texto truncado por el limite de Excel]"
                texto = texto[: 32767 - len(sufijo)] + sufijo
            celda = ET.SubElement(
                fila,
                _etiqueta(NS_MAIN, "c"),
                {
                    "r": referencia,
                    "t": "inlineStr",
                    **({"s": "1"} if numero_fila == 1 else {}),
                },
            )
            texto_xml = ET.SubElement(celda, _etiqueta(NS_MAIN, "is"))
            ET.SubElement(texto_xml, _etiqueta(NS_MAIN, "t"), {"{http://www.w3.org/XML/1998/namespace}space": "preserve"}).text = texto

    ultima_fila = len(filas)
    ET.SubElement(
        raiz,
        _etiqueta(NS_MAIN, "autoFilter"),
        {"ref": f"A1:{_letra_columna(len(CAMPOS))}{ultima_fila}"},
    )
    xml_hoja = ET.tostring(raiz, encoding="utf-8", xml_declaration=True)

    tipos = ET.Element(_etiqueta(NS_CONTENT, "Types"))
    ET.SubElement(
        tipos,
        _etiqueta(NS_CONTENT, "Default"),
        {"Extension": "rels", "ContentType": "application/vnd.openxmlformats-package.relationships+xml"},
    )
    ET.SubElement(
        tipos,
        _etiqueta(NS_CONTENT, "Default"),
        {"Extension": "xml", "ContentType": "application/xml"},
    )
    ET.SubElement(
        tipos,
        _etiqueta(NS_CONTENT, "Override"),
        {
            "PartName": "/xl/workbook.xml",
            "ContentType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
        },
    )
    ET.SubElement(
        tipos,
        _etiqueta(NS_CONTENT, "Override"),
        {
            "PartName": "/xl/worksheets/sheet1.xml",
            "ContentType": "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
        },
    )
    ET.SubElement(
        tipos,
        _etiqueta(NS_CONTENT, "Override"),
        {
            "PartName": "/xl/styles.xml",
            "ContentType": "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml",
        },
    )

    libro = ET.Element(_etiqueta(NS_MAIN, "workbook"))
    libro_vistas = ET.SubElement(libro, _etiqueta(NS_MAIN, "bookViews"))
    ET.SubElement(libro_vistas, _etiqueta(NS_MAIN, "workbookView"), {"activeTab": "0"})
    hojas = ET.SubElement(libro, _etiqueta(NS_MAIN, "sheets"))
    ET.SubElement(
        hojas,
        _etiqueta(NS_MAIN, "sheet"),
        {"name": "Reglas", "sheetId": "1", _etiqueta(NS_REL_DOC, "id"): "rId1"},
    )

    relaciones_raiz = ET.Element(_etiqueta(NS_REL_PKG, "Relationships"))
    ET.SubElement(
        relaciones_raiz,
        _etiqueta(NS_REL_PKG, "Relationship"),
        {
            "Id": "rId1",
            "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
            "Target": "xl/workbook.xml",
        },
    )
    relaciones_libro = ET.Element(_etiqueta(NS_REL_PKG, "Relationships"))
    ET.SubElement(
        relaciones_libro,
        _etiqueta(NS_REL_PKG, "Relationship"),
        {
            "Id": "rId1",
            "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
            "Target": "worksheets/sheet1.xml",
        },
    )
    ET.SubElement(
        relaciones_libro,
        _etiqueta(NS_REL_PKG, "Relationship"),
        {
            "Id": "rId2",
            "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
            "Target": "styles.xml",
        },
    )
    estilos = ET.fromstring(
        f"""<styleSheet xmlns="{NS_MAIN}">
          <fonts count="2"><font><sz val="11"/><name val="Calibri"/></font>
          <font><b/><color rgb="FFFFFFFF"/><sz val="11"/><name val="Calibri"/></font></fonts>
          <fills count="3"><fill><patternFill patternType="none"/></fill>
          <fill><patternFill patternType="gray125"/></fill>
          <fill><patternFill patternType="solid"><fgColor rgb="FF1F4E78"/>
          <bgColor indexed="64"/></patternFill></fill></fills>
          <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
          <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
          <cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
          <xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1"
          applyFill="1" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf></cellXfs>
          <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
        </styleSheet>"""
    )

    with zipfile.ZipFile(destino, "w", zipfile.ZIP_DEFLATED) as archivo:
        archivo.writestr("[Content_Types].xml", ET.tostring(tipos, encoding="utf-8", xml_declaration=True))
        archivo.writestr("_rels/.rels", ET.tostring(relaciones_raiz, encoding="utf-8", xml_declaration=True))
        archivo.writestr("xl/workbook.xml", ET.tostring(libro, encoding="utf-8", xml_declaration=True))
        archivo.writestr("xl/_rels/workbook.xml.rels", ET.tostring(relaciones_libro, encoding="utf-8", xml_declaration=True))
        archivo.writestr("xl/worksheets/sheet1.xml", xml_hoja)
        archivo.writestr("xl/styles.xml", ET.tostring(estilos, encoding="utf-8", xml_declaration=True))


def exportar():
    org = input("Clave de la organizacion (valor de sonar.organization): ").strip()
    if not org or org.startswith("sonar.organization="):
        raise ValueError("Introduce solo el valor de sonar.organization.")
    token = getpass.getpass("Token personal de SonarQube Cloud: ").strip()
    if not token:
        raise ValueError("No has introducido un token.")

    perfiles = _buscar_perfiles(org, token)
    carpeta = Path(__file__).resolve().parent
    for etiqueta, perfil in perfiles.items():
        reglas = _descargar_reglas(org, perfil, token)
        destino = carpeta / f"reglas_sonar_python_{etiqueta.lower()}.xlsx"
        _crear_excel(destino, perfil, reglas)
        print(f"Exportadas {len(reglas)} reglas de '{perfil['name']}' a: {destino}")


if __name__ == "__main__":
    try:
        exportar()
    except HTTPError as error:
        detalle = error.read().decode("utf-8", errors="replace")[:1000]
        sys.exit(f"SonarQube devolvio HTTP {error.code}: {detalle}")
    except (URLError, OSError, ValueError) as error:
        sys.exit(f"Error: {error}")
    except KeyboardInterrupt:
        sys.exit("\nExportacion cancelada.")

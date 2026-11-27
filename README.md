# Exportador de reglas de SonarQube Cloud

Aplicación de línea de comandos que consulta los perfiles de calidad de Python de una organización en SonarQube Cloud y exporta las reglas activadas de cada perfil a un archivo Excel (`.xlsx`).

## Requisitos

- Python 3.6 o posterior.
- Acceso de red a `sonarcloud.io`.
- Un token personal de SonarQube Cloud con acceso de lectura a la organización y sus perfiles.

No hay dependencias de terceros: el programa utiliza exclusivamente la biblioteca estándar de Python. No necesita instalar paquetes con `pip`.

## Uso

1. Descarga o clona este repositorio.
2. Abre una terminal en la carpeta que contiene `exportar_reglas_sonar.py`.
3. Ejecuta:

   ```powershell o en la terminal de Vscode
   python .\exportar_reglas_sonar.py
   ```

   En macOS o Linux, ejecuta `python3 exportar_reglas_sonar.py`.
4. Cuando lo solicite, introduce el valor de `sonar.organization` 
5. Introduce el token personal cuando se solicite. La entrada del token no se muestra en pantalla.

El programa exporta un Excel por cada perfil configurado en la misma carpeta que el script:

- `reglas_sonar_python_core.xlsx`
- `reglas_sonar_python_extended.xlsx`
- `reglas_sonar_python_comprehensive.xlsx`

Si un archivo con ese nombre ya existe, se sustituye. Cada libro contiene una hoja llamada `Reglas`, con encabezados, filtro automático y la primera fila inmovilizada. Se incluyen las reglas activadas en el perfil; si alguna descripción supera el límite de Excel, se trunca y se indica en la celda.

## Añadir o cambiar perfiles

Los perfiles que se descargan se configuran en `PERFILES`, al principio de `exportar_reglas_sonar.py`:

```python
PERFILES = {
    "Core": "Sonar way (Core)",
    "Extended": "Sonar way (Extended)",
    "Comprehensive": "Sonar way (Comprehensive)",
}
```

Cada entrada relaciona:

- **Clave (`"Core"`)**: etiqueta interna que se usa en el nombre del archivo. Debe ser única y conviene usar una etiqueta sencilla, sin espacios ni caracteres especiales.
- **Valor (`"Sonar way (Core)"`)**: nombre exacto del perfil tal como aparece en SonarQube Cloud.

Para descargar otro perfil, añade una entrada, por ejemplo:

```python
"Security": "Mi perfil de seguridad",
```

El programa buscará el perfil por nombre en la organización indicada y generará `reglas_sonar_python_security.xlsx`. Los nombres se comparan ignorando mayúsculas, espacios y signos de puntuación, pero debe existir exactamente una coincidencia. Si no existe o hay varias, el programa se detiene con un error.

Los perfiles seleccionados deben ser perfiles de Python: la búsqueda está limitada al lenguaje `py`. Si quieres exportar otro lenguaje, además de añadir sus perfiles tendrás que cambiar el filtro de idioma `"py"` en `_descargar_perfiles` y `_descargar_reglas`; revisa también el prefijo `reglas_sonar_python_` en `exportar()` para que los archivos generados reflejen el lenguaje.

## Cambiar las columnas del Excel

Los encabezados se definen en la tupla `CAMPOS`. Los valores de cada regla se construyen en `_crear_excel`, en la tupla `filas.append(...)`. Si añades, quitas o reordenas una columna, actualiza ambos lugares en el mismo orden. Las anchuras de columna se configuran en `anchuras` dentro de `_crear_excel`.

Los objetos y listas recibidos de la API (por ejemplo, `tags`, `params` e `impacts`) se guardan como texto JSON en sus celdas.

## Qué hace el programa

1. Solicita la clave de la organización y el token personal.
2. Consulta los perfiles de calidad de Python disponibles en SonarQube Cloud y localiza los perfiles indicados en `PERFILES`.
3. Descarga las reglas activadas para cada perfil, recorriendo las páginas de resultados de la API.
4. Genera un archivo `.xlsx` por perfil, con las reglas ordenadas por su clave.

Si la API devuelve un error HTTP, la conexión falla, la respuesta no es válida o se interrumpe la descarga, el programa muestra el error y termina. Si el fallo ocurre después de exportar otros perfiles, los archivos ya generados permanecen en la carpeta.

##  Como descargar todas las rules de Ruff

Instalar Ruff: pip install ruff
ejecutar ruff rule --all --output-format json > ruff_rules.json
Este comando descarga todas las rules de ruf fy las pasa a un JSON
por ultimo ejecutar el script de Json-a-excel.py
Para general los excel
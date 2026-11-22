# Parole

Aplicación de escritorio en Python que **calcula las fechas clave de una condena** (mitad, tercio y dos tercios, con rebajas de sanción y prisión preventiva) y **genera automáticamente el documento Word** del expediente a partir de una plantilla.

La hice para un amigo que trabaja en el ámbito jurídico-penitenciario y hacía estos cálculos y documentos a mano. Está pensada para funcionar en ordenadores antiguos.

> El repositorio solo contiene el programa. No incluye ningún dato real de personas.

## Qué hace

- **Calculadora de condena**: a partir de la fecha de entrada, la de salida, los días de rebaja y la preventiva, calcula:
  - fechas de mitad, tercio y dos tercios de la condena (con y sin rebaja)
  - condena original reconstruida sumando la preventiva
  - tiempo cumplido hasta hoy y tiempo restante
- **Datos de la persona**: formulario con los datos del expediente. La edad actual y la edad al entrar se calculan solas a partir del carnet de identidad.
- **Plantillas con marcadores**: el usuario escribe el texto del documento con marcadores como `(nombre)` o `(FechaSalidaConRebaja)`, que puede insertar con un clic. Al generar el documento, cada marcador se sustituye por su valor real, con los campos clave en negrita.
- **Documento Word** (`.docx`) con escudo, tipo de letra y texto justificado.
- **La plantilla se guarda** en `plantilla.json`, junto al programa, y se carga sola al abrirlo.

## Cómo ejecutarla

Necesitas Python 3.8 o superior.

```bash
py -m pip install -r requirements.txt
py libertades_condicionales.py
```

En Linux o macOS usa `python3` en lugar de `py`.

> Si al abrirla sale `ModuleNotFoundError: No module named 'docx'`, falta instalar las dependencias (primer comando). La librería se llama **python-docx**: `pip install docx` instala otro paquete antiguo que no funciona.

## Crear un ejecutable (.exe) para equipos sin Python

Para usarla en un ordenador que no tiene Python instalado:

```bash
py -m pip install pyinstaller
py -m PyInstaller --onefile --windowed --name Parole libertades_condicionales.py
```

El ejecutable queda en la carpeta `dist/`. Copia junto a él el `escudo.png` si lo usas.
Para Windows 7, crea el ejecutable con **Python 3.8**, la última versión compatible con ese sistema.

## Estructura

```
libertades_condicionales.py   Interfaz gráfica (Tkinter)
calculos.py                   Lógica: fechas de condena, edades, relleno de plantillas
tests/test_calculos.py        Tests automáticos de la lógica
requirements.txt              Dependencias
```

La lógica está separada de la interfaz para poder probarla con tests sin abrir ventanas.

## Tests

```bash
py -m unittest discover tests
```

## Marcadores disponibles

| Grupo | Marcadores |
| --- | --- |
| Datos personales | `(nombre)` `(establecimiento)` `(ci)` `(edadactual)` `(edaddentrada)` `(direccion)` |
| Delito y tribunal | `(delito)` `(tribunal)` `(causa)` `(inciso)` `(posh)` `(razon)` `(clasif)` `(favorable)` |
| Información adicional | `(escolaridad)` `(naturalde)` `(estadocivil)` `(regimen)` `(trabajosocial)` `(oficios)` `(respcivil)` `(indisciplinas)` |
| Fechas originales | `(FechaEntrada)` `(FechaSalida)` `(CondenaOriginalInfo)` `(MitadOriginal)` `(TercioOriginal)` `(DosTerciosOriginal)` |
| Fechas con rebaja | `(FechaSalidaConRebaja)` `(MitadConRebaja)` `(TercioConRebaja)` `(DosTerciosConRebaja)` `(DiasRebaja)` |
| Preventiva | `(DiasPreventiva)` `(FechaSalidaOriginalConPreventiva)` `(CondenaOriginalTotal)` |
| Otros | `(FechaActual)` |

Los paréntesis que no son marcadores, como `(art. 58)`, se dejan tal cual.

## Tecnologías

Python · Tkinter · python-docx · unittest

## Autor

Alex Álvarez Morales · [github.com/A013X](https://github.com/A013X)

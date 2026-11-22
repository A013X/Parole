"""Lógica de Parole sin interfaz gráfica.

Aquí están los cálculos de fechas de la condena, el cálculo de edades a partir
del carnet de identidad y el relleno de plantillas en documentos Word.
Al no depender de Tkinter, todo esto se puede probar con tests automáticos.
"""

from dataclasses import dataclass
from datetime import datetime, timedelta
from typing import Dict, Iterable, Optional

MESES = [
    "enero", "febrero", "marzo", "abril", "mayo", "junio",
    "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre",
]

FORMATO_FECHA = "%d/%m/%Y"


def parsear_fecha(texto: str) -> datetime:
    """Convierte 'DD/MM/YYYY' en datetime. Lanza ValueError si no es válida."""
    return datetime.strptime(texto.strip(), FORMATO_FECHA)


def fecha_larga(fecha: Optional[datetime]) -> str:
    """15/03/2026 -> '15 de marzo de 2026'."""
    if not fecha:
        return ""
    return f"{fecha.day} de {MESES[fecha.month - 1]} de {fecha.year}"


def convertir_dias(dias: int) -> str:
    """Convierte un número de días en texto aproximado: '2 años, 3 meses, 5 días'.

    Usa años de 365 días y meses de 30, como en el cálculo original de la app.
    """
    if dias == 0:
        return "0 días"
    signo = "-" if dias < 0 else ""
    dias_abs = abs(dias)
    anios, resto = divmod(dias_abs, 365)
    meses, dias_resto = divmod(resto, 30)
    if meses >= 12:
        anios += meses // 12
        meses %= 12

    partes = []
    if anios:
        partes.append(f"{anios} año{'s' if anios != 1 else ''}")
    if meses:
        partes.append(f"{meses} mes{'es' if meses != 1 else ''}")
    if dias_resto:
        partes.append(f"{dias_resto} día{'s' if dias_resto != 1 else ''}")
    return signo + (", ".join(partes) if partes else "0 días")


@dataclass
class ResultadoCondena:
    fecha_entrada: datetime
    fecha_salida: datetime
    rebaja: int
    prev_dias: int
    condena_efectiva_dias: int
    condena_original_dias: int
    salida_original_con_preventiva: datetime
    mitad: datetime
    tercio: datetime
    dos_tercios: datetime
    salida_con_rebaja: datetime
    mitad_con_rebaja: datetime
    tercio_con_rebaja: datetime
    dos_tercios_con_rebaja: datetime
    dias_transcurridos: int
    dias_cumplidos: int
    dias_restantes: int


def calcular_condena(fecha_entrada: datetime, fecha_salida: datetime,
                     rebaja: int = 0, prev_dias: int = 0,
                     hoy: Optional[datetime] = None) -> ResultadoCondena:
    """Calcula las fechas clave de una condena.

    - fecha_salida ya tiene descontada la prisión preventiva.
    - rebaja: días de rebaja de sanción, que adelantan todas las fechas.
    - prev_dias: días de preventiva, solo para reconstruir la condena original.
    """
    if fecha_salida <= fecha_entrada:
        raise ValueError("La fecha de salida debe ser posterior a la de entrada")
    if rebaja < 0 or prev_dias < 0:
        raise ValueError("La rebaja y la preventiva no pueden ser negativas")
    if hoy is None:
        hoy = datetime.now()
    hoy = hoy.replace(hour=0, minute=0, second=0, microsecond=0)

    efectiva = (fecha_salida - fecha_entrada).days
    mitad = fecha_entrada + timedelta(days=efectiva // 2)
    tercio = fecha_entrada + timedelta(days=efectiva // 3)
    dos_tercios = fecha_entrada + timedelta(days=efectiva * 2 // 3)
    descuento = timedelta(days=rebaja)
    salida_con_rebaja = fecha_salida - descuento

    transcurridos = max((hoy - fecha_entrada).days, 0)

    return ResultadoCondena(
        fecha_entrada=fecha_entrada,
        fecha_salida=fecha_salida,
        rebaja=rebaja,
        prev_dias=prev_dias,
        condena_efectiva_dias=efectiva,
        condena_original_dias=efectiva + prev_dias,
        salida_original_con_preventiva=fecha_salida + timedelta(days=prev_dias),
        mitad=mitad,
        tercio=tercio,
        dos_tercios=dos_tercios,
        salida_con_rebaja=salida_con_rebaja,
        mitad_con_rebaja=mitad - descuento,
        tercio_con_rebaja=tercio - descuento,
        dos_tercios_con_rebaja=dos_tercios - descuento,
        dias_transcurridos=transcurridos,
        dias_cumplidos=transcurridos + rebaja,
        dias_restantes=max((salida_con_rebaja - hoy).days, 0),
    )


def nacimiento_desde_ci(ci: str, hoy: Optional[datetime] = None) -> Optional[datetime]:
    """Obtiene la fecha de nacimiento a partir de un carnet de identidad cubano.

    Los 6 primeros dígitos son AAMMDD y el 7.º indica el siglo:
    9 -> 1800, 0-5 -> 1900, 6-8 -> 2000.
    Si el carnet está incompleto, se deduce el siglo por el año y se usa el 1 de enero.
    Devuelve None si no se puede deducir nada.
    """
    ci = (ci or "").strip()
    if len(ci) < 2 or not ci[:2].isdigit():
        return None
    aa = int(ci[:2])

    if len(ci) >= 7 and ci[:7].isdigit():
        siglo = int(ci[6])
        base = 1800 if siglo == 9 else 2000 if siglo >= 6 else 1900
        try:
            return datetime(base + aa, int(ci[2:4]), int(ci[4:6]))
        except ValueError:
            pass  # mes o día imposibles: usamos solo el año

    hoy = hoy or datetime.now()
    anio = 2000 + aa if aa <= hoy.year % 100 else 1900 + aa
    return datetime(anio, 1, 1)


def edad_en(nacimiento: datetime, fecha: datetime) -> int:
    """Edad cumplida en una fecha, teniendo en cuenta si ya pasó el cumpleaños."""
    cumplido = (fecha.month, fecha.day) >= (nacimiento.month, nacimiento.day)
    return fecha.year - nacimiento.year - (0 if cumplido else 1)


def rellenar_plantilla(doc, plantilla: str, valores: Dict[str, str],
                       negrita: Iterable[str] = ()) -> None:
    """Escribe la plantilla en un documento de python-docx sustituyendo marcadores.

    Un marcador es un texto entre paréntesis, por ejemplo (nombre). Si no está en
    valores se deja tal cual, así que los paréntesis normales del texto no se tocan.
    Los marcadores incluidos en negrita se escriben en negrita.
    """
    negrita = set(negrita)
    lineas = plantilla.split("\n")
    for i, linea in enumerate(lineas):
        if not linea.strip():
            # Varias líneas vacías seguidas cuentan como un solo salto
            if i < len(lineas) - 1 and not lineas[i + 1].strip():
                continue
            doc.add_paragraph()
            continue

        parrafo = doc.add_paragraph()
        resto = linea
        while resto:
            inicio = resto.find("(")
            fin = resto.find(")", inicio) if inicio != -1 else -1
            if inicio == -1 or fin == -1:
                parrafo.add_run(resto)
                break
            if inicio > 0:
                parrafo.add_run(resto[:inicio])
            marcador = resto[inicio:fin + 1]
            run = parrafo.add_run(valores.get(marcador, marcador))
            if marcador in negrita and marcador in valores:
                run.bold = True
            resto = resto[fin + 1:]

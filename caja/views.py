from django.shortcuts import render
from openpyxl import load_workbook
from pathlib import Path


def moneda(valor):

    if valor is None:
        return "$ 0,00"

    try:

        valor = float(valor)

        texto = f"{valor:,.2f}"

        texto = texto.replace(",", "X")
        texto = texto.replace(".", ",")
        texto = texto.replace("X", ".")

        return f"$ {texto}"

    except:

        return "$ 0,00"

def moneda_usd(valor):

    if valor is None:
        return "USD 0,00"

    try:

        valor = float(valor)

        texto = f"{valor:,.2f}"

        texto = texto.replace(",", "X")
        texto = texto.replace(".", ",")
        texto = texto.replace("X", ".")

        return f"USD {texto}"

    except:

        return "USD 0,00"


def dashboard(request):

    archivo = Path("data/Caja.xlsx")

    wb = load_workbook(
        archivo,
        data_only=True
    )

    hoja = wb["Caja"]

    contexto = {

        # Responsable
        "operador": hoja["A3"].value,
        "fecha": hoja["B3"].value,

        # KPI Presidenciales
        "total_dia": moneda(hoja["B8"].value),
        "mercado_libre": moneda(hoja["B9"].value),
        "distribucion": moneda(hoja["B10"].value),
        "total_local": moneda(hoja["B11"].value),
        "local_efectivo": moneda(hoja["B12"].value),
        "saldo_final_uyu": moneda(hoja["B13"].value),

        # Caja UYU
        "saldo_inicial_uyu": moneda(hoja["B4"].value),
        "movimientos_uyu": moneda(hoja["B5"].value),

        # Cierre UYU
        "sobrante_uyu": moneda(hoja["B14"].value),
        "faltante_uyu": moneda(hoja["B15"].value),

        "saldo_inicial_usd": moneda_usd(hoja["B6"].value),
        "movimientos_usd": moneda_usd(hoja["B7"].value),

        # Cierre USD
        "saldo_final_usd": moneda_usd(hoja["B16"].value),
        "sobrante_usd": moneda_usd(hoja["B17"].value),
        "faltante_usd": moneda_usd(hoja["B18"].value),

        # Observación
        "observacion": hoja["B19"].value,
    
    }

    return render(
        request,
        "dashboard.html",
        contexto
    )
def asistencias(request):

    return render(
        request,
        "asistencias.html"
    )

def asistencias(request):

    archivo = Path("data/Caja.xlsx")

    wb = load_workbook(
        archivo,
        data_only=True
    )

    hoja = wb["Ast. Caja"]

    registros = []

    fila = 2

    while hoja[f"A{fila}"].value:

        registros.append({

            "fecha": hoja[f"A{fila}"].value,

            "responsable": hoja[f"B{fila}"].value,

            "dia": hoja[f"C{fila}"].value,

        })

        fila += 1

    contexto = {

        "registros": registros

    }

    return render(
        request,
        "asistencias.html",
        contexto
    )
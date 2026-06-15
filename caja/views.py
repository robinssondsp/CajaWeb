from django.shortcuts import render
from openpyxl import load_workbook
from pathlib import Path


def moneda(valor):

    if valor is None:
        return "$ 0,00"

    try:

        valor = float(valor)

        valor = valor / 100

        texto = f"{valor:,.2f}"

        texto = texto.replace(",", "X")
        texto = texto.replace(".", ",")
        texto = texto.replace("X", ".")

        return f"$ {texto}"

    except:

        return "$ 0,00"


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

        # Caja USD
        "saldo_inicial_usd": hoja["B6"].value,
        "movimientos_usd": hoja["B7"].value,

        # Cierre UYU
        "sobrante_uyu": moneda(hoja["B14"].value),
        "faltante_uyu": moneda(hoja["B15"].value),

        # Cierre USD
        "saldo_final_usd": hoja["B16"].value,
        "sobrante_usd": hoja["B17"].value or 0,
        "faltante_usd": hoja["B18"].value or 0,

        # Observación
        "observacion": hoja["B19"].value,
    }

    return render(
        request,
        "dashboard.html",
        contexto
    )
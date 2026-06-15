from django.shortcuts import render
from openpyxl import load_workbook
from pathlib import Path

def dashboard(request):

    archivo = Path("data/Caja.xlsx")

    wb = load_workbook(
        archivo,
        data_only=True
    )

    hoja = wb["Caja"]

    contexto = {

    "operador": hoja["A3"].value,
    "fecha": hoja["B3"].value,

    # Caja UYU
    "saldo_inicial_uyu": hoja["B4"].value,
    "movimientos_uyu": hoja["B5"].value,

    # Caja USD
    "saldo_inicial_usd": hoja["B6"].value,
    "movimientos_usd": hoja["B7"].value,

    # Ventas
    "mercado_libre": hoja["B9"].value,
    "distribucion": hoja["B10"].value or 0,
    "total_local": hoja["B11"].value,
    "local_efectivo": hoja["B12"].value,

    # Cierre UYU
    "saldo_final_uyu": hoja["B13"].value,
    "sobrante_uyu": hoja["B14"].value or 0,
    "faltante_uyu": hoja["B15"].value or 0,

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
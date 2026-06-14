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

        "mercado_libre": hoja["B9"].value,
        "distribucion": hoja["B10"].value,
        "total_local": hoja["B11"].value,
        "local_efectivo": hoja["B12"].value,
        "saldo_final": hoja["B13"].value,
    }

    return render(
        request,
        "dashboard.html",
        contexto
    )
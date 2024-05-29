import xlwings as xw
import pandas as pd
import streamlit as st

titulo = st.container()
datos = st.container()
opciones = st.container()

with titulo:
        st.title("Bienvenido a la interfaz para calcular el precio de opciones y ver otros datos")
with datos:
        wb = xw.Book(r"")
        sheet_prima = wb.sheets("Cálculo de prima")
        sheet_resumen = wb.sheets("Resumen")
        edad = int(input("Elije la edad del asegurado"))
        sheet_prima.range(11,13).value = edad
        macro_calc = wb.macro("Resumen")
        macro_calc()
        df = sheet_resumen.range("B8:E27").options(pd.DataFrame, header=1, index=False, expand = "table").value
        wb.save()

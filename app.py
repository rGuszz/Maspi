import xlwings as xw
import pandas as pd
import streamlit as st

titulo = st.container()
datos = st.container()
opciones = st.container()

with titulo:
        st.title("Bienvenido a la interfaz para calcular el precio de opciones y ver otros datos")
with datos:
        app = xw.App(visible=True, add_book=False)
        wb = app.books.open(r"Cálculos Seguro Por si las Flies M.xlsm")
        sheet_prima = wb.sheets("Cálculo de prima")
        sheet_resumen = wb.sheets("Resumen")
        edad = st.slider("Elije la edad del asegurado", min_value=30, max_value=45, step=1, value=30)
        sheet_prima.range(11,13).value = edad
        macro_calc = wb.macro("Resumen")
        macro_calc()
        df = sheet_resumen.range("B8:E27").options(pd.DataFrame, header=1, index=False, expand = "table").value
        wb.save()
        st.write(df)

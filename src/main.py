import os 
import sys
import time
import pandas as pd
import numpy as np
import openpyxl
from openpyxl.chart import BarChart, Reference, Series, BarChart3D,PieChart, PieChart3D, LineChart,RadarChart,ScatterChart,AreaChart,StockChart, BubbleChart, SurfaceChart,DoughnutChart
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import PatternFill, Alignment, Font
from openpyxl.chart.label import DataLabelList

import locale
locale.setlocale(locale.LC_ALL, 'ca_ES.utf8') #Catalan
from tkinter import Tk, Button, messagebox, Label
from tkcalendar import Calendar, DateEntry
import datetime

#no importar los warnings
import warnings
warnings.filterwarnings("ignore")

now = datetime.datetime.now()


########### Comentarios ###########
# 1. En el excel de plantilla, en la hoja de tecnologias, hay una tabla con las columnas [Fecha,Tecnologies,Producció OK,Producció KO,Total Producció]
###################################

def read_excel(path,sheet_name):
    df = pd.read_excel(path,sheet_name=sheet_name)
    return df


def generate_desplegament_total(df):
    #hacer una copia del dataframe para no afectar el original
    df = df.copy()
    #En la columa de fecha, se debe de poner el formato de año y mes
    df['Fecha'] = pd.to_datetime(df['Fecha']).dt.strftime('%Y-%m-%B')

    #eliminar la columna de Urgentes
    df = df.drop(columns=['Urgents'])

     #Agrupar por fecha y sumar las columnas de Producció OK y Producció KO
    df = df.groupby(['Fecha']).sum()

    #Resetear el index para que la columna de fecha sea una columna normal
    df = df.reset_index()

    #ordenar por años y mes en orden de meses
    df = df.sort_values(by=['Fecha'],ascending=True)

    #eliminar de la fecha el año y mes para que solo quede el nombre del mes
    df['Fecha'] = df['Fecha'].str.split('-').str[2]

    #poner en mayusculas la primera letra del mes
    df['Fecha'] = df['Fecha'].str.capitalize()

    return df

def generate_desplegamnet_mes(df,mes):
    df = df.copy()
    #En la columa de fecha, se debe de poner el formato de año y mes
    df['Fecha'] = pd.to_datetime(df['Fecha']).dt.strftime('%Y-%m')

    #eliminar la columna de Urgentes
    df = df.drop(columns=['Urgents'])
    #Filtrar por mes
    df = df[df['Fecha'] == mes]

    #Agrupar por tecnologia y sumar las columnas de Producció OK y Producció KO
    df = df.groupby(['Tecnologies']).sum()

    #Resetear el index para que la columna de fecha sea una columna normal
    df = df.reset_index()

    #Ordenar el dataframe de mayor a menor por la columna de Total Producció
    df = df.sort_values(by=['Total Producció'],ascending=False)

    #Eliminar del dataframe las filas que tengan el valor de Total Producció igual a 0
    df = df[df['Total Producció'] != 0]

    return df

def generate_total_tecnologia(df):
    df = df.copy()
    # Agrupar por tecnologias y sumar el total es decir creamos un nuevo dataframe con las columnas [Tecnologies,Total Producción]
    df = df.groupby(['Tecnologies']).sum()

    # Resetear el index para que la columna de fecha sea una columna normal
    df = df.reset_index()

    # Ordenar el dataframe de mayor a menor por la columna de Total Producció
    df = df.sort_values(by=['Total Producció'],ascending=False)

    #Eliminar del dataframe las filas que tengan el valor de Total Producció igual a 0
    df = df[df['Total Producció'] != 0]

    df  = df[['Tecnologies','Total Producció']]

    return df

def generate_total_tecnologia_mes(df,mes):
    df = df.copy()
    #En la columa de fecha, se debe de poner el formato de año y mes
    df['Fecha'] = pd.to_datetime(df['Fecha']).dt.strftime('%Y-%m')

    #Seleccionar datos del mes
    df = df[df['Fecha'] == mes]

    # Agrupar por tecnologias y sumar el total es decir creamos un nuevo dataframe con las columnas [Tecnologies,Total Producción]
    df = df.groupby(['Tecnologies']).sum()

    # Resetear el index para que la columna de fecha sea una columna normal
    df = df.reset_index()

    # Ordenar el dataframe de mayor a menor por la columna de Total Producció
    df = df.sort_values(by=['Total Producció'],ascending=False)

    #Eliminar del dataframe las filas que tengan el valor de Total Producció igual a 0
    df = df[df['Total Producció'] != 0]

    df  = df[['Tecnologies','Total Producció']]
   
    return df


def generate_total_desplegament_tecnologia_mes(df):
    df = df.copy()

    #En la columa de fecha, se debe de poner el formato de año y mes
    df['Fecha'] = pd.to_datetime(df['Fecha']).dt.strftime('%Y-%m-%B')

     #Crearemos un nuevo dataset con las columnas fechas y una nueva columna por cada tecnologia del dataset original
    df = df.pivot_table(index=['Fecha'],columns=['Tecnologies'],values=['Total Producció'],aggfunc=np.sum)

   
    #Resetear el index para que la columna de fecha sea una columna normal
    df = df.reset_index()

    #ponemos todos los NaN a 0
    df = df.fillna(0)

    #ordenar por años y mes en orden de meses
    df = df.sort_values(by=['Fecha'],ascending=True)

    #eliminar de la fecha el año y mes para que solo quede el nombre del mes
    df['Fecha'] = df['Fecha'].str.split('-').str[2]

    # poner la primera letra de  cada mes en mayuscula
    df['Fecha'] = df['Fecha'].str.capitalize()

    #eliminar la cabecera de las columnas
    df.columns = df.columns.droplevel(0)

    #creamos una columna nueva con el total general de produccion
    df['Total General'] = df.sum(axis=1)

    return df

def generate_total_desplegament_DevOps_mes(df):
    df = df.copy()
    #Tomar solo los valores donde tecnologia sea igual a Devops
    df = df[df['Tecnologies'] == 'Devops']

    #elimina la columna "Producció OK"
    df = df.drop(['Producció OK'],axis=1)

    #eliminar la columna de Urgentes
    df = df.drop(['Urgents'],axis=1)

     #En la columa de fecha, se debe de poner el formato de año y mes
    df['Fecha'] = pd.to_datetime(df['Fecha']).dt.strftime('%Y-%m-%B')

    #Agrupar por fecha y sumar las columnas de Producció OK y Producció KO
    df = df.groupby(['Fecha']).sum()

    #Resetear el index para que la columna de fecha sea una columna normal
    df = df.reset_index()

    #ordenar por años y mes en orden de meses
    df = df.sort_values(by=['Fecha'],ascending=True)

    #eliminar de la fecha el año y mes para que solo quede el nombre del mes
    df['Fecha'] = df['Fecha'].str.split('-').str[2]

    # poner la primera letra de  cada mes en mayuscula
    df['Fecha'] = df['Fecha'].str.capitalize()

    #Agregar una columna que sea % de KO sobre total de produccion el calculo es Producció KO / Total Producció, mostrar sin decimales
    df['% KO'] = (df['Producció KO'] / df['Total Producció']) * 100

    #eliminar los decimales de la columna % KO
    df['% KO'] = df['% KO'].astype(int)

    # las columnas son Fecha, Producció KO, Total Producció, % KO, quiero cambiar la posicion de Producció KO, Total Producció, para que quede Fecha, Total Producció, Producció KO, % KO
    df = df[['Fecha','Total Producció','Producció KO','% KO']]

    return df

def generate_total_desplegament_Urgente_mes(df):
    df = df.copy()

    #eliminar la columna "Producció OK"
    df = df.drop(['Producció OK'],axis=1)

    #eliminar la columna "Producció KO"
    df = df.drop(['Producció KO'],axis=1)


    #En la columa de fecha, se debe de poner el formato de año y mes
    df['Fecha'] = pd.to_datetime(df['Fecha']).dt.strftime('%Y-%m-%B')

    #Agrupar por fecha y sumar las columnas de Producció OK y Producció KO
    df = df.groupby(['Fecha']).sum()

    #Resetear el index para que la columna de fecha sea una columna normal
    df = df.reset_index()

    #ordenar por años y mes en orden de meses
    df = df.sort_values(by=['Fecha'],ascending=True)

    #eliminar de la fecha el año y mes para que solo quede el nombre del mes
    df['Fecha'] = df['Fecha'].str.split('-').str[2]

    # poner la primera letra de  cada mes en mayuscula
    df['Fecha'] = df['Fecha'].str.capitalize()

    #Agregar una columna que sea % de Urgentes sobre total de produccion el calculo es Urgentes / Total Producció, mostrar sin decimales
    df['% Urgents'] = (df['Urgents'] / df['Total Producció']) * 100

    #eliminar los decimales de la columna % Urgentes
    df['% Urgents'] = df['% Urgents'].astype(int)

    return df

 


def excel_style(ws):
    # Define the font and alignment for the header
    font = Font(name='Arial', size=12, bold=True, italic=False, vertAlign=None, underline='none', strike=False, color='ffffff')
    alignment = Alignment(horizontal='center', vertical='center', text_rotation=0, wrap_text=False, shrink_to_fit=False, indent=0)

    # Define the fill for the header
    fill = PatternFill(fill_type='solid', start_color='000000', end_color='000000')

    # Apply the style to the header
    for cell in ws["1:1"]:
        ws.column_dimensions[cell.column_letter].width = 20
        cell.font = font
        cell.alignment = alignment
        cell.fill = fill


    #Colores a las filas intercaladas
    # Define the fill color for odd rows
    fill = PatternFill(start_color='b3d2ff', end_color='FFC000', fill_type='solid')

    # Apply the style to odd rows
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
        if row[0].row % 2 == 0:
            for cell in row:
                cell.fill = fill
    
    return ws
    

    
def generate_graphic(df,doc_excel,hoja=None,title=None,x_axis=None):

    #si el archivo excel existe, se abre y agrega la hoja de trabajo, si no existe, se crea el archivo excel
    if os.path.isfile(doc_excel):
        wb = openpyxl.load_workbook(doc_excel)
        ws = wb[hoja]
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = hoja

    # borrar la tabla de la hoja de trabajo y agregar el dataframe
    ws.delete_rows(1,ws.max_row)
    for r in dataframe_to_rows(df, index=False, header=True):
        ws.append(r)

    ws = excel_style(ws)
   
    
    # Crear el gráfico de barras
    chart = BarChart3D()
    chart.type = "col"
    chart.style = 10
    chart.title = title
    chart.y_axis.title = 'Total'
    chart.x_axis.title = x_axis

    # Referenciar las celdas del gráfico
    data = Reference(ws, min_col=2, min_row=1, max_col=3, max_row=ws.max_row)

    # Referenciar las celdas de las categorias
    cats = Reference(ws, min_col=1, min_row=2, max_row=ws.max_row)
    
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)

    # Añadir el gráfico a la hoja de trabajo
    ws.add_chart(chart, "G2")
    
    #nombre de la hoja
    ws.title = hoja
    
    # Guardar el archivo
    wb.save(doc_excel)


def generate_circular_graphic(df,doc_excel,hoja=None,title=None):

    #si el archivo excel existe, se abre y agrega la hoja de trabajo, si no existe, se crea el archivo excel
    if os.path.isfile(doc_excel):
        wb = openpyxl.load_workbook(doc_excel)
        ws = wb[hoja]
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = hoja

    # borra la tabla de la hoja de trabajo y agregar el dataframe
    ws.delete_rows(1,ws.max_row)
    for r in dataframe_to_rows(df, index=False, header=True):
        ws.append(r)

    ws = excel_style(ws)
   
    '''# Crear el gráfico circular
    chart = PieChart()
    chart.type = "col"
    chart.style = 10
    chart.title = title

    # Configurar el estilo del gráfico para simular un anillo
    for serie in chart.series:
        serie.graphicalProperties.pieChart = "squaredCircle" 

    #mostrar los datos de cada categoria en el grafico circular
    chart.dataLabels = DataLabelList()
    chart.dataLabels.showVal = True


    # Referenciar las celdas del gráfico
    data = Reference(ws, min_col=2, min_row=1, max_col=3, max_row=ws.max_row)

    # Referenciar las celdas de las categorias
    cats = Reference(ws, min_col=1, min_row=2, max_row=ws.max_row)

    chart.add_data(data, titles_from_data=True)

    chart.set_categories(cats)

    # Añadir el gráfico a la hoja de trabajo
    ws.add_chart(chart, "G2")'''

    #guardar el archivo
    wb.save(doc_excel)


def generate_horizontal_graphic(df,doc_excel,hoja=None,title=None):

    #si el archivo excel existe, se abre y agrega la hoja de trabajo, si no existe, se crea el archivo excel
    if os.path.isfile(doc_excel):
        wb = openpyxl.load_workbook(doc_excel)
        ws = wb[hoja]
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = hoja

    # borrar la tabla de la hoja de trabajo y agregar el dataframe
    ws.delete_rows(1,ws.max_row)
    for r in dataframe_to_rows(df, index=False, header=True):
        ws.append(r)

    ws = excel_style(ws)
   
    '''
    # Crear el gráfico de lineas
    chart =  StockChart()
    chart.type = "col"
    chart.style = 10
    chart.title = title


    # Referenciar las celdas del gráfico
    data = Reference(ws, min_col=2, min_row=1, max_col=ws.max_column-1, max_row=ws.max_row)

    # Referenciar las celdas de las categorias
    cats = Reference(ws, min_col=1, min_row=2, max_row=ws.max_row)

    chart.add_data(data, titles_from_data=True)

    chart.set_categories(cats)

    # Añadir el gráfico a la hoja de trabajo
    ws.add_chart(chart, "G22")

    '''
    #nombre de la hoja
    ws.title = hoja

    # Guardar el archivo
    wb.save(doc_excel)

def generate_graphic_barras_lineal(df,doc_excel,hoja=None,title=None):
    #si el archivo excel existe, se abre y agrega la hoja de trabajo, si no existe, se crea el archivo excel
    if os.path.isfile(doc_excel):
        wb = openpyxl.load_workbook(doc_excel)
        #Ubicarnos en la hoja de trabajo , PERO NO crearla
        ws = wb[hoja]
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = hoja

    # borrar la tabla de la hoja de trabajo y agregar el dataframe
    ws.delete_rows(1,ws.max_row)
    for r in dataframe_to_rows(df, index=False, header=True):
        ws.append(r)

    ws = excel_style(ws)

    # Parte para generar los graficos.

    '''# Crear el gráfico de barras
    chart = LineChart()
    chart.type = "col"
    chart.style = 10
    chart.title = title

    # Referenciar las celdas del gráfico
    data = Reference(ws, min_col=ws.max_column, min_row=1, max_col=ws.max_column, max_row=ws.max_row)

    # Referenciar las celdas de las categorias
    cats = Reference(ws, min_col=1, min_row=2, max_row=ws.max_row)

    chart.add_data(data, titles_from_data=True)

    chart.set_categories(cats)

    # Añadir el gráfico a la hoja de trabajo
    ws.add_chart(chart, "G30")


    # Crear el gráfico de lineas
    chart2 = BarChart()
    chart2.type = "col"
    chart2.style = 10
    chart2.title = title

    # Referenciar las celdas del gráfico
    data = Reference(ws, min_col=2, min_row=1, max_col=ws.max_column-1, max_row=ws.max_row)

    # Referenciar las celdas de las categorias
    cats = Reference(ws, min_col=1, min_row=2, max_row=ws.max_row)

    chart2.add_data(data, titles_from_data=True)

    chart2.set_categories(cats)

    #combinar los dos graficos
    chart += chart2

    # Añadir el gráfico a la hoja de trabajo
    ws.add_chart(chart2, "G40")'''

    #nombre de la hoja
    ws.title = hoja

    # Guardar el archivo
    wb.save(doc_excel)


#Generar una funcion para mostar un calendario por pantalla que solo permita seleccionar un mes y un año

def get_date():
    """Muestra un calendario para que el usuario seleccione una fecha."""
    root = Tk()
    root.geometry("300x350")
   
    root.title("Seleccionar Fecha")
   
    def get_selected_date():
        """Obtiene la fecha seleccionada por el usuario y cierra la ventana."""
        nonlocal selected_date
        selected_date = cal.selection_get()
        root.destroy()

    def cancel():
        """Cierra la ventana sin seleccionar una fecha."""
        nonlocal selected_date
        selected_date = None
        root.destroy()
        exit()

    #Poner un titulo en la ventana
    title_label = Label(root, text="Seleccionar Data", font=("Helvetica", 16), fg="white", bg="black")
    title_label.pack(pady=10)
    
    selected_date = None
    
    cal = Calendar(root, selectmode="day", year=now.year , month=now.month, day=now.day,maxdate=now)
    cal.pack(pady=20)


    btn_ok = Button(root, text="OK", command=get_selected_date)
    btn_ok.pack(side="left", pady=10, padx=10)

    btn_cancel = Button(root, text="CANCEL·LAR", command=cancel)
    btn_cancel.pack(side="right", pady=10, padx=10)

    root.protocol("WM_DELETE_WINDOW", lambda: messagebox.showerror("Error", "Heu de seleccionar una data"))

    root.mainloop()

    if selected_date is None:
        raise ValueError("No s'ha seleccionat cap data")
    return selected_date
    
    

def main():
    path = r'docs\plantilla.xlsx'
    sheet_name = 'Tecnologies'


    fecha = '2023-09'
    df = read_excel(path,sheet_name)

    generate_horizontal_graphic(generate_total_desplegament_tecnologia_mes(df),'docs\grafico.xlsx','Total','Evolució desplegaments per tecnologia')
    
    generate_graphic(generate_desplegament_total(df),'docs\grafico.xlsx','Total desplegaments',"Desplegaments - Evolució mensual","Mesos")
    generate_graphic(generate_desplegamnet_mes(df,fecha),'docs\grafico.xlsx','Total desplegaments mes',"Deplegaments -" + fecha,"Tecnologies")

    generate_circular_graphic(generate_total_tecnologia(df),'docs\grafico.xlsx','Total tecnologia',"Evolució Mensual per volum")
    generate_circular_graphic(generate_total_tecnologia_mes(df,fecha),'docs\grafico.xlsx','Total tecnologia mes',"Evolució Mensual per volum -" + fecha)
    
    generate_graphic_barras_lineal( generate_total_desplegament_Urgente_mes(df),'docs\grafico.xlsx','% KO Urgentes',"% Peticions Urgents")
    generate_graphic_barras_lineal(generate_total_desplegament_DevOps_mes(df), 'docs\grafico.xlsx', '% KO DevOps', "% Peticions DevOps")


if __name__ == '__main__':
    main()

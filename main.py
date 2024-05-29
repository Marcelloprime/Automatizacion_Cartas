from docxtpl import DocxTemplate
from datetime import datetime
import locale
import os
from docx2pdf import convert
import csv
import streamlit as st
from babel.dates import format_date
# import pythoncom

# Se hace un contador para poder determinar el N° del documento
# Primero hacer hace una lectura del contador, el contador será un documento externo ubicado en la carpeta "base"
def leer_contador(df='cartas_fede\\base\\contador.csv'):
    try:
        with open(df, "r", newline="") as archivo:
            lector_csv = csv.reader(archivo)
            fila = next(lector_csv)
            return int(fila[0])
    except FileNotFoundError:
        # Si el archivo no existe, retorna el valor inicial deseado
        return 19

# El segundo paso es hacer el conteo de documentos.

def incrementar_contador(df="base\\contador.csv"):
    contador = leer_contador(df) + 1
    try:
        with open(df, "w", newline="") as archivo:
            escritor_csv = csv.writer(archivo)
            escritor_csv.writerow([contador])
        return contador
    except Exception as e:
        st.error(f"Error al escribir el contador{e}")
        return None

# Escritura del documeto
# Acá se reemplazarán los parámetros establecidos en el documento original

def reemplazo(template = "base/Carta-Base.docx", constantes=None, output=None):
    doc = DocxTemplate(template)
    doc.render(constantes)
    doc.save(output)


## El dia actual de la carta
def fecha_carta_hoy():
    fecha_actual = datetime.now()
    # Formateamos la fecha en el formato deseado usando Babel
    fecha_formateada = format_date(fecha_actual, format="EEEE d 'de' MMMM", locale="es_ES")
    # Capitalizamos correctamente la primera letra
    return fecha_formateada.capitalize()

## Luces

def luces(x):
    if x is True:
        return "También quisiera solicitar el uso de las luces durante los horarios mencionados."
    else:
        return ""




# Elección de cancha o losa

def cancha_hoy(canchas):
    if canchas == "Futbol 11":
        return f'cancha de {canchas}'
    elif canchas == "Basket":
        return f'losa de {canchas}'

    elif canchas == "Voley":
        return f'losa de {canchas} N°1'

    else:
        lista = []
        for cancha in canchas:
            if cancha != "," and cancha.isnumeric() == True:
                lista.append(int(cancha))
        numeros_formateados = [f"losa N°{numero}" for numero in lista]
        # Unir los elementos formateados en una cadena separada por comas
        cadena_formateada = ", ".join(numeros_formateados)
        cadena_formateada = cadena_formateada.rsplit(", ", 1)
        cadena_formateada = " y ".join(cadena_formateada)

    return cadena_formateada


# Creación del streamlit
def main():
    color1='#FF3131'
    color2='#FF914D'
    st.title("Automatización de documentos de solicitud de áreas deportivas🥅")
    st.subheader("Secretaría de Deportes del Centro Federado de Economía y Planificación🍊⚽")

    st.header("Ingrese los detalles para generar el documento:")
    try:
        locale.setlocale(locale.LC_TIME, 'es_ES.UTF-8')
    except locale.Error:
        locale.setlocale(locale.LC_TIME, 'C')
    canchas= st.selectbox("Ingrese la(s) cancha(s) separado por comas (Futbol 11, Basket, Voley o números de losas):",['Futbol 11','Basket','Voley','N° de losa'])
    if canchas == "N° de losa":
        canchas = st.text_input('Especifica el número de losa')

    canchas_input = cancha_hoy(canchas)
    fecha_evento = st.date_input("Seleccione la fecha del evento:", format='DD/MM/YYYY')
    hora_inicio = st.time_input("Hora de inicio del evento:")
    hora_fin = st.time_input("Hora de fin del evento:")
    luz = st.checkbox("Luz")

    if st.button("Generar Documento"):
        if canchas_input:
            NUMBER = incrementar_contador()
            reemplazos = {
                'NUMERO_CARTA': f'0{NUMBER}',
                'FECHA': f'{fecha_carta_hoy()}',
                'USO': f'uso de la {cancha_hoy(canchas)}',
                'DIA': format_date(fecha_evento, format="full", locale="es_ES").capitalize(),
                'HORA': f'{hora_inicio.strftime("%H:%M")} - {hora_fin.strftime("%H:%M")}',
                'LUCES':f'{luces(luz)}'
            }
            output_path_docx = f"docx/Carta D-{reemplazos['NUMERO_CARTA']}-CFEP.docx"
            output_path_pdf = f"pdf/Carta D-{reemplazos['NUMERO_CARTA']}-CFEP.pdf"

            reemplazo(constantes=reemplazos, output=output_path_docx)

            # pythoncom.CoInitialize() Si se ejecuta en computadora
            convert(input_path=output_path_docx, output_path=output_path_pdf)
            # pythoncom.CoUninitialize() Si se ejecuta en computadora

            if os.path.exists(output_path_docx) and os.path.exists(output_path_pdf):
                st.success(
                    f"Se ha procesado el archivo 'Carta D-{reemplazos['NUMERO_CARTA']}-CFEP.docx' correctamente y ha sido convertido a PDF.")
                st.download_button('Descargar DOCX', data=open(output_path_docx, 'rb'),
                                   file_name=f"Carta D-{reemplazos['NUMERO_CARTA']}-CFEP.docx")
                st.download_button('Descargar PDF', data=open(output_path_pdf, 'rb'),
                                   file_name=f"Carta D-{reemplazos['NUMERO_CARTA']}-CFEP.pdf")
            else:
                st.error("No se ha generado el archivo, hubo problemas.")
        else:
            st.warning("Por favor, ingrese la(s) cancha(s).")


if __name__ == "__main__":
    main()
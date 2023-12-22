import os
import re
import pandas as pd
import openpyxl
import shutil
from pathlib import Path
from PyPDF2 import PdfReader

# Listas para ordenar facturas
edenor = []
naturgy = []
aysa = []

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Nombres
excel_naturgy = "facturas_naturgy.xlsx"
excel_aysa = "facturas_aysa.xlsx"
excel_edenor = "facturas_edenor.xlsx"

# Patrones
patron_factura_naturgy = r"(0{3}\d{2})(\d{8})"
patron_factura_edenor = r"BLiquidación deServicio Público N°(\d{4}[-]\d+)"
patron_factura_aysa = r"LSP (\d{4}[A-Z]\d+)"

# funcion para verificar a que grupo pertenece
def procesar_facturas():

    carpeta_facturas = os.path.join(BASE_DIR, 'facturas')

    carpeta_aysa = os.path.join(carpeta_facturas,'Facturas de Aysa - Bot')
    carpeta_naturgy = os.path.join(carpeta_facturas,'Facturas de Naturgy - Bot')
    carpeta_edenor = os.path.join(carpeta_facturas,'Facturas de Edenor - Bot')

    # Verifica si la carpeta no existe antes de crearla
    if not os.path.exists(carpeta_aysa):
        os.makedirs(carpeta_aysa)
        nueva_carpeta1  = os.path.join(carpeta_facturas,'Facturas de Aysa - Bot')

    if not os.path.exists(carpeta_naturgy):
        os.makedirs(carpeta_naturgy)
        nueva_carpeta2  = os.path.join(carpeta_facturas,'Facturas de Naturgy - Bot')

    if not os.path.exists(carpeta_edenor):
        os.makedirs(carpeta_edenor)
        nueva_carpeta3  = os.path.join(carpeta_facturas,'Facturas de Edenor - Bot')

    carpeta_path = Path(carpeta_facturas)

    archivos_pdf = carpeta_path.glob("*.pdf")

    for archivo_pdf in archivos_pdf:
        reader = PdfReader(archivo_pdf)
        page = reader.pages[0]
        text = page.extract_text()

        lineas = text.split('\n')

        def mover_factura(archivo, carpeta_destino):
            shutil.copy(archivo, os.path.join(carpeta_destino, archivo.name))

        for linea in lineas:         
            match_factura = re.search(patron_factura_edenor, linea)
            if match_factura != None:
                edenor.append(archivo_pdf.name)
                mover_factura(archivo_pdf, carpeta_edenor)
            match_factura = re.search(patron_factura_naturgy, linea)
            if match_factura != None:
                naturgy.append(archivo_pdf.name)
                mover_factura(archivo_pdf, carpeta_naturgy)
            match_factura = re.search(patron_factura_aysa, linea)
            if match_factura != None:
                aysa.append(archivo_pdf.name)
                mover_factura(archivo_pdf, carpeta_aysa)   
                break

#   Utiliza el método glob para buscar archivos PDF en la carpeta
    archivos_pdf = list(carpeta_path.glob("*.pdf"))

#   Cuenta la cantidad de archivos PDF
    cantidad_pdf = len(archivos_pdf)
    cantidad_naturgy = len(naturgy)
    cantidad_edenor = len(edenor)
    cantidad_aysa = len(aysa)

    pdf_no_clasificados = cantidad_pdf - (cantidad_aysa + cantidad_edenor + cantidad_naturgy)
    pdf_clasificados = cantidad_pdf - pdf_no_clasificados

    print(f"Se encontraron {cantidad_pdf} Pdf's, de los cuales {pdf_clasificados} pudieron clasificarse, y {pdf_no_clasificados} quedaron sin clasificación.")

    if pdf_no_clasificados > 0:
        for archivo in archivos_pdf:
            if archivo.name not in naturgy and archivo.name not in aysa and archivo.name not in edenor:
                print(f"Archivo no clasificado: {archivo.name}")
    
#   recorrer cada lista armando los excels correspondientes.

    def lector_aysa():

        # Listas para almacenar datos de facturas
        # Obtener la ruta del directorio actual donde se encuentra el archivo Python
        directorio_actual = os.path.dirname(os.path.abspath(__file__))

        # Agregar "FACTURAS\\Naturgy" a la ruta del directorio actual
        carpeta_facturas = Path(directorio_actual) / "facturas" / "Facturas de Aysa - Bot"
        archivos_pdf = carpeta_facturas.glob("*.pdf")

        # Crear listas para almacenar los datos
        fechas_vencimiento = []
        total_pagar = []
        numero_de_factura = []
        punto_venta = []
        cuenta = []

        # Establecer patrones de búsqueda para la fecha, el monto, el número de factura y la cuenta
        patron_linea_fecha = r"Vencimiento (\d{2}/\d{2}/\d{4})"
        patron_monto_pagar = r"Total apagar \$([\d\.,]+)"
        patron_cuenta = r"Cuenta deServicios (\d+)"

        for archivo_pdf in archivos_pdf:
            reader = PdfReader(archivo_pdf)
            page = reader.pages[0]
            text = page.extract_text()

            # Buscar las líneas que contienen la información de interés
            lineas = text.split('\n')

            # Reinicializar las variables en cada iteración del bucle exterior
            fecha_vencimiento = None
            monto_pagar = None
            numero_de_cuenta = None
            numero_factura = None
            punto_de_venta = None

            for linea in lineas:
                # Buscar el número de cuenta en el texto
                match_cuenta = re.search(patron_cuenta, linea)
                if match_cuenta:
                    numero_de_cuenta = match_cuenta.group(1)

                # Buscar el número de factura en el texto
                match_factura_aysa = re.search(patron_factura_aysa, linea)
                if match_factura_aysa:
                    numero_factura_completo = match_factura_aysa.group(1)
                # Utilizar una expresión regular para dividir la cadena en punto de venta y número de factura
                    match_divisor = re.search(r"[ab]", numero_factura_completo, re.IGNORECASE)
                    if match_divisor:
                        partes = re.split(r"[ab]", numero_factura_completo, flags=re.IGNORECASE)
                        punto_de_venta = partes[0]  # Tomar el primer elemento
                        numero_factura = partes[1]  # Tomar el segundo elemento
                    else:
                        numero_factura = numero_factura_completo

                # Buscar la línea que contiene "Total apagar"
                match_monto_pagar = re.search(patron_monto_pagar, linea)
                if match_monto_pagar:
                    monto_pagar = match_monto_pagar.group(1)

                # Utilizar el patrón regex para encontrar la fecha de vencimiento
                match_fecha = re.search(patron_linea_fecha, linea)
                if match_fecha:
                    fecha_vencimiento = match_fecha.group(1)

                # Si se encontraron todos los valores, guárdalos y salta a la siguiente factura
                if fecha_vencimiento and monto_pagar and numero_de_cuenta and numero_factura:
                    fechas_vencimiento.append(fecha_vencimiento)
                    total_pagar.append(monto_pagar)
                    numero_de_factura.append(numero_factura)
                    punto_venta.append(punto_de_venta)
                    cuenta.append(numero_de_cuenta)
                    break  # Salir del bucle de líneas y pasar a la siguiente factura

        # Crear un DataFrame con los datos
        data = {
            "N° de Cuenta": cuenta,
            "Punto de Venta": punto_venta,
            "N° de Factura": numero_de_factura,
            "Total a Pagar": total_pagar,
            "Fecha de Vencimiento": fechas_vencimiento,
        }

        df_aysa = pd.DataFrame(data)
        # df_aysa["Total a Pagar"] = df_aysa["Total a Pagar"].astype(int)

        # Guardar el DataFrame en un archivo Excel
        nombre_archivo_excel = os.path.join(carpeta_aysa, excel_aysa)
        df_aysa.to_excel(nombre_archivo_excel, sheet_name="Facturas Aysa", index=False)

        print(f"Datos guardados en el archivo Excel: {nombre_archivo_excel}")
        return df_aysa

    def lector_edenor():
        # Obtener la lista de archivos PDF en la carpeta "Facturas"
        # Obtener la ruta del directorio actual donde se encuentra el archivo Python
        directorio_actual = os.path.dirname(os.path.abspath(__file__))

        # Agregar "FACTURAS\\Naturgy" a la ruta del directorio actual
        carpeta_facturas = Path(directorio_actual) / "facturas" / "Facturas de Edenor - Bot"
        archivos_pdf = carpeta_facturas.glob("*.pdf")

        # Crear listas para almacenar los datos
        fechas_vencimiento = []
        total_pagar = []
        numero_de_factura = []
        punto_venta = []
        cuenta = []

        # Establecer patrones de búsqueda para la fecha, el monto, el número de factura y la cuenta
        patron_linea_fecha = r"Hasta el(\d{2}/\d{2}/\d{4})"
        patron_monto_pagar = r"TOTAL APAGAR \$([\d\.,]+)"
        patron_cuenta = r"Cuenta (\d{10})"

        for archivo_pdf in archivos_pdf:
            reader = PdfReader(archivo_pdf)
            page = reader.pages[0]
            text = page.extract_text()

            # Buscar las líneas que contienen la información de interés
            lineas = text.split('\n')
        # print(lineas)

            # Reinicializar las variables en cada iteración del bucle exterior
            fecha_vencimiento = None
            monto_pagar = None
            numero_de_cuenta = None
            numero_factura = None
            punto_de_venta = None

            for linea in lineas:
                # Buscar el número de cuenta en el texto
                match_cuenta = re.search(patron_cuenta, linea)
                if match_cuenta:
                    numero_de_cuenta = match_cuenta.group(1)

                # Buscar el número de factura en el texto
                match_numero_edenor = re.search(patron_factura_edenor, linea)
                if match_numero_edenor:
                    numero_factura_completo = match_numero_edenor.group(1)
                    punto_de_venta, numero_factura = numero_factura_completo.split("-")
                    punto_de_venta = punto_de_venta[:4] # Separa el punto de venta del resto de la factura y lo almacena como un dato aparte
                    # numero_factura = numero_factura[1:]  # Eliminar el primer "0"

                # Buscar la línea que contiene "Total apagar"
                match_monto_pagar = re.search(patron_monto_pagar, linea)
                if match_monto_pagar:
                    monto_pagar = match_monto_pagar.group(1)

                # Utilizar el patrón regex para encontrar la fecha de vencimiento
                match_fecha = re.search(patron_linea_fecha, linea)
                if match_fecha:
                    fecha_vencimiento = match_fecha.group(1)

                # Si se encontraron todos los valores, guárdalos y salta a la siguiente factura
                if fecha_vencimiento and monto_pagar and numero_de_cuenta and numero_factura:
                    fechas_vencimiento.append(fecha_vencimiento)
                    total_pagar.append(monto_pagar)
                    numero_de_factura.append(numero_factura)
                    punto_venta.append(punto_de_venta)
                    cuenta.append(numero_de_cuenta)
                    break  # Salir del bucle de líneas y pasar a la siguiente factura

        # Crear un DataFrame con los datos
        data = {
            "N° de Cuenta": cuenta,
            "Punto de Venta": punto_venta,
            "N° de Factura": numero_de_factura,
            "Total a Pagar": total_pagar,    
            "Fecha de Vencimiento": fechas_vencimiento,
        }

        df_edenor = pd.DataFrame(data)
        # df_edenor["Total a Pagar"] = df_edenor["Total a Pagar"].astype(int)

        # Guardar el DataFrame en un archivo Excel
        nombre_archivo_excel = os.path.join(carpeta_edenor, excel_edenor)
        df_edenor.to_excel(nombre_archivo_excel, sheet_name="Facturas Edenor", index=False)

        print(f"Datos guardados en el archivo Excel: {nombre_archivo_excel}")
        return df_edenor

    def lector_naturgy():
        # Obtener la ruta del directorio actual donde se encuentra el archivo Python
            directorio_actual = os.path.dirname(os.path.abspath(__file__))

            # Agregar "FACTURAS\\Naturgy" a la ruta del directorio actual
            carpeta_facturas = Path(directorio_actual) / "facturas" / "Facturas de Naturgy - Bot"
            archivos_pdf = carpeta_facturas.glob("*.pdf")

            # Crear listas para almacenar los datos
            cuenta = []
            facturas = []
            vencimiento = []
            punto_venta = []
            monto_pagar = []

            # Establecer patrón de búsqueda para el número de cuenta
            patron_cuenta = r"(\d{6,}/\d)(?=\d\s+de)"
            patron_vencimiento = r"Mensual  (\d{2}/\d{2}/\d{2})"
            patron_monto_total = r"\$\s*([\d,.]+) \d"

            for archivo_pdf in archivos_pdf:
                reader = PdfReader(archivo_pdf)
                page = reader.pages[0]
                text = page.extract_text()

                # Buscar las líneas que contienen la información de interés
                lineas = text.split('\n')
                # print(lineas)
                # Reinicializar las variables en cada iteración del bucle exterior
                num_cuenta = None
                numero_de_factura = None
                fecha_vencimiento = None
                punto_de_venta = None
                monto = "No encontrado"

                for linea in lineas:
                    # Utilizar los patrones de búsqueda regex para encontrar la información
                    match_cuenta = re.search(patron_cuenta, linea)
                    match_factura_naturgy = re.search(patron_factura_naturgy, linea)
                    match_vencimiento = re.search(patron_vencimiento, linea)
                    match_monto_total = re.search(patron_monto_total, linea)
                    # Si encuentra la línea de la factura, intentamos extraer el número de cuenta
                    if match_factura_naturgy:
                        factura_completa = match_factura_naturgy.group()
                        if factura_completa:
                            match_punto_venta = re.search(r"(\d{5})(\d+)", factura_completa)
                            if match_punto_venta:
                                punto_de_venta = match_punto_venta.group(1)
                                numero_de_factura = match_punto_venta.group(2)

                    if match_vencimiento:
                        fecha_vencimiento = match_vencimiento.group(1)

                    # Intentar separar el número de cuenta del número de página
                    if match_cuenta:
                        num_cuenta = match_cuenta.group(1)
                        # Eliminar cualquier texto adicional que pueda estar presente
                        num_cuenta = num_cuenta.split()[0]  # Tomar solo la primera parte

                    if match_monto_total:
                        monto = match_monto_total.group(1)  # Capturar el primer monto a pagar

                # Llenar las listas con valores predeterminados si no se encontró información
                if fecha_vencimiento and numero_de_factura and punto_de_venta and num_cuenta and monto:
                    cuenta.append(num_cuenta)
                    facturas.append(numero_de_factura)
                    vencimiento.append(fecha_vencimiento)
                    punto_venta.append(punto_de_venta)
                    monto_pagar.append(monto)
            # Crear un DataFrame con los datos
            data = {
                "N° de Cuenta": cuenta,
                "Punto de Venta": punto_venta,
                "N° de Factura": facturas,
                "Total a Pagar": monto_pagar,
                "Fecha de Vencimiento": vencimiento,
            }

            df_naturgy = pd.DataFrame(data)
            # df_naturgy["Total a Pagar"] = df_naturgy["Total a Pagar"].astype(int)

            # Guardar el DataFrame en un archivo Excel
            nombre_archivo_excel = os.path.join(carpeta_naturgy, excel_naturgy)
            df_naturgy.to_excel(nombre_archivo_excel, sheet_name="Facturas Naturgy", index=False)

            print(f"Datos guardados en el archivo Excel: {nombre_archivo_excel}")
            return df_naturgy

    df_aysa = lector_aysa()
    df_edenor = lector_edenor()
    df_naturgy = lector_naturgy()
    
    # # Calcular los totales a pagar para cada tipo de factura
    # total_pagar_aysa = df_aysa["Total a Pagar"].sum()
    # total_pagar_edenor = df_edenor["Total a Pagar"].sum()
    # total_pagar_naturgy = df_naturgy["Total a Pagar"].sum()

    # Crear un DataFrame con los datos
    data = {
        "Relevamiento": ["PDF's encontrados",
        "PDF's clasificados",
        "PDF's no clasificados"],
        "Cantidad": [cantidad_pdf, pdf_clasificados, pdf_no_clasificados],
        # "Proveedores": ["Aysa", "Edenor", "Naturgy"],
        # "Total a pagar": [total_pagar_aysa, total_pagar_edenor, total_pagar_naturgy]
    }

    df_reporte = pd.DataFrame(data)

    # Guardar el DataFrame en un archivo Excel
    nombre_archivo_excel = os.path.join(carpeta_facturas, "reporte_facturas.xlsx")
    df_reporte.to_excel(nombre_archivo_excel, sheet_name="Reporte", index=False)

    print(f"Datos guardados en el archivo Excel: {nombre_archivo_excel}")
    
    # Hacer copias de los DataFrames de Aysa, Edenor y Naturgy
    df_aysa_copia = df_aysa.copy()
    df_edenor_copia = df_edenor.copy()
    df_naturgy_copia = df_naturgy.copy()

    # # Guardar los DataFrames originales en archivos Excel
    # nombre_archivo_aysa = os.path.join(carpeta_aysa, excel_aysa)
    # df_aysa.to_excel(nombre_archivo_aysa, sheet_name="Facturas Aysa", index=False)
    # print(f"Datos guardados en el archivo Excel: {nombre_archivo_aysa}")

    # nombre_archivo_edenor = os.path.join(carpeta_edenor, excel_edenor)
    # df_edenor.to_excel(nombre_archivo_edenor, sheet_name="Facturas Edenor", index=False)
    # print(f"Datos guardados en el archivo Excel: {nombre_archivo_edenor}")

    # nombre_archivo_naturgy = os.path.join(carpeta_naturgy, excel_naturgy)
    # df_naturgy.to_excel(nombre_archivo_naturgy, sheet_name="Facturas Naturgy", index=False)
    # print(f"Datos guardados en el archivo Excel: {nombre_archivo_naturgy}")

    # Guardar las copias de los DataFrames en archivos Excel duplicados
    nombre_archivo_aysa_copia = os.path.join(carpeta_aysa, "copia_" + excel_aysa)
    df_aysa_copia.to_excel(nombre_archivo_aysa_copia, sheet_name="Facturas Aysa", index=False)
    print(f"Copia de datos guardada en el archivo Excel: {nombre_archivo_aysa_copia}")

    nombre_archivo_edenor_copia = os.path.join(carpeta_edenor, "copia_" + excel_edenor)
    df_edenor_copia.to_excel(nombre_archivo_edenor_copia, sheet_name="Facturas Edenor", index=False)
    print(f"Copia de datos guardada en el archivo Excel: {nombre_archivo_edenor_copia}")

    nombre_archivo_naturgy_copia = os.path.join(carpeta_naturgy, "copia_" + excel_naturgy)
    df_naturgy_copia.to_excel(nombre_archivo_naturgy_copia, sheet_name="Facturas Naturgy", index=False)
    print(f"Copia de datos guardada en el archivo Excel: {nombre_archivo_naturgy_copia}")

if __name__ == "__main__":
    procesar_facturas()
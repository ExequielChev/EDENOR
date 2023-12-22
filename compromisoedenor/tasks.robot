*** Settings ***

Library    RPA.Windows
Library    RPA.Excel.Files
Library    RPA.Tables
Library    OperatingSystem
Library    DateTime
Library    String

*** Variables ***

${diccionario} =     C:\\Users\\zcheveste\\Desktop\\EDENOR\\diccionario.xlsx
${diccionario_hoja} =    dicci_hoja
${contador}    0
${value_to_write} =    OK
${column_name} =    U
${nombre_carpeta}    Compromisos_pdfEdenor

*** Tasks ***
Open Major desktop application and play a app
    #Open the Major.Exe desktop application 
    Creacion de Carpetas
    Ir a incio de usuario 
    Carga de datos

*** Keywords ***

Open the Major.Exe desktop application

    #Crear variables para los datos dentro del diccionario
    RPA.Excel.Files.Open Workbook    ${diccionario}
    ${usuario} =    Get Cell Value    9   B
    ${contraseña} =    Get Cell Value    10    B

    ##Abrir el sistema Major (Windows + R)
    Windows Run    Major.Exe    
    Sleep    5s
    ##Seleccionar la ventana del Sistema Contabilidad
    RPA.Windows.Click    id:25    timeout=60
    Sleep    40s
    ##Clickear en la barra para escribir el Nombre del usuario
    RPA.Windows.Click    id:6    timeout=120
    Send Keys    keys=${usuario}
    ##Clickear en la barra para escribir la contraseña del usuario
    RPA.Windows.Click    id:5    
    Send Keys    keys=${contraseña}
    ## Aceptar el cartel para finalizar el ingreso del usuario al sistema
    RPA.Windows.Click    id:4
    # Sleep    15s

Creacion de Carpetas

    #Crear variables para los datos dentro del diccionario
    RPA.Excel.Files.Open Workbook    ${diccionario}
    ${descargas} =    Get Cell Value    2    B

    #Consigue los datos de la fecha por separado    
    ${fecha_hoy} =    Get Current Date
    ${año} =    Convert Date    ${fecha_hoy}    %Y
    ${mes} =    Convert Date    ${fecha_hoy}    %m
    ${día} =    Convert Date    ${fecha_hoy}    %d

    # Ruta completa de la carpeta
    ${ruta_año}    Set Variable    ${descargas}\\${año}
    ${ruta_mes}    Set Variable    ${ruta_año}\\${mes}
    ${ruta_dia}    Set Variable    ${ruta_mes}\\${día}
    ${ruta_carpeta}    Set Variable    ${ruta_dia}\\${nombre_carpeta}

    # Verificar y crear la carpeta de año
    ${existe_carpeta_año}    Run Keyword And Return Status    Directory Should Exist    ${ruta_año}
    Run Keyword If    not ${existe_carpeta_año}    Create Directory    ${ruta_año}

    # Verificar y crear la carpeta de mes
    ${existe_carpeta_mes}    Run Keyword And Return Status    Directory Should Exist    ${ruta_mes}
    Run Keyword If    not ${existe_carpeta_mes}    Create Directory    ${ruta_mes}

    # Verificar y crear la carpeta de día
    ${existe_carpeta_día}    Run Keyword And Return Status    Directory Should Exist    ${ruta_dia}
    Run Keyword If    not ${existe_carpeta_día}    Create Directory    ${ruta_dia}

    # Verificar y crear la carpeta específica
    ${existe_carpeta_especifica}    Run Keyword And Return Status    Directory Should Exist    ${ruta_carpeta}
    Run Keyword If    not ${existe_carpeta_especifica}    Create Directory    ${ruta_carpeta}
    
    
Ir a incio de usuario
    ## ir a la pestaña de Transacciones
    RPA.Windows.Click    name:Transacciones    timeout=60
    ## ir a la pestaña de Compromisos
    RPA.Windows.Click    name:Compromiso    timeout=30
    ## ir a la pestaña de Sin orden de Compra
    RPA.Windows.Click    id:89    timeout=30
    Sleep    5s

Carga de datos 
    #Crear variables para los datos dentro del diccionario
    RPA.Excel.Files.Open Workbook    ${diccionario}
    ${descargas} =    Get Cell Value    2    B
    ${excel_comprobantes} =    Get Cell Value    3    B
    ${excel_matriz} =    Get Cell Value    4    B
    ${excel_factura} =    Get Cell Value    5    B
    ${matriz_hoja} =    Get Cell Value    6    B
    ${comprobantes_hoja} =    Get Cell Value    7    B
    ${factura_hoja} =    Get Cell Value    8    B
    ${usuario} =    Get Cell Value    9   B
    ${contraseña} =    Get Cell Value    10    B

    #Abre el excel SalesData y crea una lista con la cual trabajar
    RPA.Excel.Files.Open Workbook    ${excel_matriz}
    ${data_as_table} =    Read Worksheet As Table    ${matriz_hoja}    header=True
    @{cuenta} =    Create List  # Crear una lista vacía para almacenar los datos de las columnas

    RPA.Excel.Files.Open Workbook    ${excel_factura}
    #Abre el excel facturas_aysa y crea una lista con la cual trabajar

    ${data_as_table2} =    Read Worksheet As Table    ${factura_hoja}    header=True
    @{n_cuenta2} =    Create List  # Crear una lista vacía para almacenar los datos de las columnas

    ${contadorROW2}    Set Variable    2  # Puedes ajustar el valor inicial según tus necesidades
    Set Cell Value    1    T    ORDEN
    Set Cell Value    1    U    COMPROMISO

    FOR    ${row2}    IN    @{data_as_table2}

        Set Cell Value    ${contadorROW2}    T    ${contadorROW2}
        ${contadorROW2}    Evaluate    ${contadorROW2} + 1
    END

    ${data_as_table2} =    Read Worksheet As Table    ${factura_hoja}    header=True
    Save Workbook

    # Filtra el excel por la columna de servicio, todos los que sean servicios de AGUA (AYSA)
    ${filtered_data} =    Filter Table By Column    ${data_as_table}    SERVICIO    ==    ENERGIA ELECTRICA
    Log    ${filtered_data}

    FOR    ${row2}    IN    @{data_as_table2}

        FOR    ${row}    IN    @{data_as_table}
            
            #Iterar sobre las filas de la columna estado para saber si el compromiso ya fue cargado anteriormente o no, los compromisos cargados deberan tener escrito un "OK" en la columna "ESTADO"
# Iterar sobre las filas de la columna estado para saber si el compromiso ya fue cargado anteriormente o no.
            ${compromiso} =    Set Variable    ${row2["COMPROMISO"]}
            #Iterar sobre las filas de la columna estado para saber si el compromiso ya fue cargado anteriormente o no, los compromisos cargados deberan tener escrito un "OK" en la columna "ESTADO"
            
            IF    '${compromiso}' != 'OK'
            Log    Checking row con el estado
            ELSE
                Continue For Loop If    '${compromiso}' == 'OK'
            Log    Skipping row with "OK"
            # Si encuentra un OK significa que el compromiso ya fue cargado, esto volvera a repetir el Loop hasta encontrar un compromiso sin cargar
            END


            ${n_cuenta2} =    Set Variable    ${row2["CUENTA"]}
            ${cuenta} =    Set Variable    ${row["NUMERO_DE_CTA"]}

            IF    '${n_cuenta2}'=='${cuenta}'
                Log    Datos coinciden: ${cuenta}
                #Seleccionar cartel para comenzar Nuevo compromiso 
                RPA.Windows.Click    id:20    timeout=30

                #cargar tipo de orden de compra o afectacion varia (15 o 16)
                ${tipo} =    Set Variable    ${row["TIPO"]}
                Send keys    id:42    ${tipo}

                #cargar el tipo de proveerdor, en caso de aysa (12)
                ${provetipo} =    Set Variable    ${row["PROVE_TIPO"]}
                Send Keys    id:40    ${provetipo}

                #Cargar el numero del proveedor en caso de aysa (120196)
                ${numerotipo} =    Set Variable    ${row["NUMERO"]}
                Send Keys    id:33    ${numerotipo}
                
                #ir a la tarjeta para la carga de la Jurisccion 
                Send Keys    keys={TAB 6}{RIGHT 8}
                Send Keys    keys={ENTER}

                #Cargar la  Jurisdiccion
                ${jurisdiccion} =    Set Variable    ${row["JURISDICCION"]}
                Send Keys    id:666     ${jurisdiccion}
                Send Keys    keys={ENTER 2}

                #Cargar la Est. programatica
                ${programatica} =    Set Variable    ${row["EST.PROGRAMATICA"]}
                Send Keys    keys=${programatica}
                Send Keys    keys={ENTER 2}

                #Cargar el Fondo financiado
                ${financiamiento} =    Set Variable    ${row["FF"]}
                Send Keys    keys=${financiamiento}
                Send Keys    keys={ENTER 2}

                #Cargar el Objeto del gasto
                ${objetogasto} =    Set Variable    ${row["OBJETO_DE_GASTO"]}
                Send Keys    keys=${objetogasto}
                Send Keys    keys={ENTER}

                #Cargar el Detalle
                Send Keys    keys=Generacion automatica
                Send Keys    keys={ENTER}

                # cargar el importe 
                ${importe} =    Set Variable    ${row2["IMPORTE"]}
                Send Keys    keys=${importe}
                Send Keys    keys={ENTER}
                Sleep    3s

                # Aceptar carteles los cuales pueden aparecer 2, 1 de estos es por saldo insuficiente 
                ${condicion} =    Set Variable    False
                    FOR    ${i}    IN RANGE    3
                        ${elemento_existente3} =        Run Keyword And Return Status    RPA.Windows.Click    locator=id:2
                        Run Keyword If    ${elemento_existente3}    Log    El elemento existe
                    ...    ELSE    ${condicion}
                            Log    El elemento no existe
                    END
                    Sleep    3s

                    #Aceptar el segundo Cartel si aparece y si no seguir con la carga de datos 
                    FOR    ${i}    IN RANGE    3
                        ${elemento_existente4} =        Run Keyword And Return Status    RPA.Windows.Click    locator=id:1
                        Run Keyword If    ${elemento_existente4}    Log    El elemento existe
                    ...    ELSE    ${condicion}
                            Log    El elemento no existe
                    END
                Sleep    0.5s

                #cargar el nunero de oficina.
                ${ofi} =    Set Variable    ${row["OFICI"]}
                Send Keys    keys=${ofi}
                Send Keys    keys={ENTER}
                Sleep    0.5s

                #Pasar por las columnas hasta llegar a la descripcion para poder cargar esta misma .
                Send Keys    keys={RIGHT 8}

                #Cargar la Descripcion.
                Send Keys    keys=ENERGIA ELECTRICA
                Send Keys    keys={ENTER}

                #Cargar la Cantidad de productos.
                ${cant} =    Set Variable    ${row["CANTIDAD"]}
                Send Keys    keys=${cant}
                Send Keys    keys={ENTER}

                #Cargar el precio unitario.
                Send Keys    keys=${importe}
                Send Keys    keys={ENTER}
                Sleep    2s

                #Aceptar cartel luego de que ya se ha completado la carga de datos para la creacion del compromiso. 
                RPA.Windows.Click    id:22    timeout=30

                #Aceptar Cartel el cual puede aparecer o no, este mismo se repite de la carga del importe el cual puede ser por el saldo insuficiente.
                    FOR    ${i}    IN RANGE    3
                        ${elemento_existente} =        Run Keyword And Return Status    RPA.Windows.Click    locator=id:2
                        Run Keyword If    ${elemento_existente}    Log    El elemento existe
                    ...    ELSE    ${condicion}
                            Log    El elemento no existe
                    END
                    Sleep    2s

                    #Aceptar el segundo cartel el cual puede o no aprecer.
                    FOR    ${i}    IN RANGE    3
                        ${elemento_existente2} =        Run Keyword And Return Status    RPA.Windows.Click    locator=id:1
                        Run Keyword If    ${elemento_existente2}    Log    El elemento existe
                    ...    ELSE    ${condicion}
                            Log    El elemento no existe
                    END

                #Escribir un OK en el excel para poder gestionar si el compromiso se cargo o no se cargo anteriormente
                    ${numerofila} =    Set Variable    ${row2["ORDEN"]}
                    ${numerofila1} =    Convert To Integer    ${numerofila}
                    Set Cell Value    ${numerofila}    ${column_name}    ${value_to_write}
                    Log    Se cambió el valor de la celda a 1
                    Save Workbook   

                #Cargar la Observacion del compromiso (carga importante del numero de cuenta para luego poder relacionarlos al momento de carga de facturas con el 16)
                RPA.Windows.Click    id:1
                Sleep    0.5s

                #Ubicarse en la tarjeta de "Datos" para lograr llegar a la ventana de Observaciones
                Send Keys    keys={TAB 13}
                Sleep    1s

                #Pasar a la ventana de observaciones
                Send Keys    keys={RIGHT}

                #Apretar click en editar para comenzar la carga de Observaciones 
                RPA.Windows.Click    id:54    timeout=30
                Sleep    0.5s

                #Cargar los datos de las observaciones (en este caso sera el numero de cuenta y la direccion del mismo)
                ${nombre}    Set Variable    ${row2["NOMBRE"]}
                Send Keys    keys=NOMBRE:${nombre} {ENTER}
                ${dire}    Set Variable    ${row["DIRECCION"]}
                Send Keys    keys=DIRECCION:${dire} {ENTER}
                Send Keys    keys=NUMERO_DE_CUENTA:${cuenta} {ENTER}
                ${periodo}    Set Variable    ${row2["#PERIODO"]}
                Send Keys    keys=PERIODO:${periodo}

                #Aceptar la edicion de la carga de las observaciones
                RPA.Windows.Click    id:56    timeout=30

                #Volver a posicionarse nuevamente en la pestaña de datos para continuar con la descarga tipo PDF de cada compromiso
                Send Keys    keys={RIGHT}
                Sleep    0.5s
                Send Keys    keys={TAB 2}        
                Sleep    1s
                Send Keys    keys={LEFT}
                Sleep    1s 

                #Apretar en el boton "IMPRIMIR" para poder descargar como compromiso
                RPA.Windows.Click    id:7    timeout=30
                Sleep    0.5s

                #Aceptar la impresion del mismo
                RPA.Windows.Click    id:12    timeout=30
                Sleep    5s

                #Elegir como impresora (poner en preferecias la impresora que convierte en PDF los archivos)
                RPA.Windows.Click    locator=id:2    timeout=60
                Send Keys    keys={TAB}{TAB}{DOWN}
                Sleep    1s

                #Ejecutar la "impresion" la cual te abrira la pestaña para poder guardar el archivo PDF
                RPA.Windows.Click    id:10    timeout=10
                Sleep    15s

                #Borra el nombre que viene por defecto en el pdf el cual es "Crystal Reports"
                RPA.Windows.Double Click    id:1148    timeout=160
                Send Keys    keys={CTRL}{A}
                Send Keys    keys={DEL}
                # RPA.Windows.Double Click    id:1148    timeout=30
                # Send Keys    keys={DEL}
                # Sleep    0.5s

                #Cargar el nombre del PDF y la ruta, en este caso le pondremos como nombre el numero de cuenta y un numero del cual es la iteracion por la que va 
                ${contador} =    Convert To Integer    ${contador}
                ${contador} =    Evaluate    ${contador} + 1
                ${numero_de_cuenta} =    Set Variable    (${row["NUMERO_DE_CTA"]})
                ${numero_de_cuenta} =    Convert To String    ${numero_de_cuenta}
                ${numero_de_cuenta} =    Set Variable    ${numero_de_cuenta.replace('/', '_')}(${contador})
                ${fecha_hoy} =    Get Current Date
                ${año} =    Convert Date    ${fecha_hoy}    %Y
                ${mes} =    Convert Date    ${fecha_hoy}    %m
                ${día} =    Convert Date    ${fecha_hoy}    %d
                # Ruta completa de la carpeta
                ${ruta_año}    Set Variable    ${descargas}\\${año}
                ${ruta_mes}    Set Variable    ${ruta_año}\\${mes}
                ${ruta_dia}    Set Variable    ${ruta_mes}\\${día}

                Send Keys    keys=C:\\Users\\zcheveste\\Documents\\${año}\\${mes}\\${día}\\${nombre_carpeta}\\BOT_${numero_de_cuenta}.pdf
                #Send Keys    keys=C:\\Users\\zcheveste\\Documents\\Robocop_project\\OCR\\comprobantes\\BOT_${numero_de_cuenta}.pdf
                Sleep    3s

                #Guardar archivo PDF
                Send Keys    keys={ENTER}

                # Una vez guardado el PDF se abrira este mismo y procedemos a cerrar la ventana del PDF que ya ha sido guardado 
                RPA.Windows.Click    name:AVPageView    timeout=30
                Send Keys    keys={CTRL}{Q}

                #Esperamos 5 segundos para que se acomoden los ID del sistema y se vuelva a iterar el FOR 
                Sleep    5s
            
                END
            END
        END
    Close Workbook


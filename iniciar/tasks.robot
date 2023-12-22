
*** Settings ***
Library    RPA.Windows
Library    RPA.Excel.Files
# Library    RPA.Tables
# Library    String
# Library    DateTime
# Library    RPA.Desktop
# Library    OperatingSystem
# Library    openpyxl
# Library    Collections
# Library    RPA.FTP






*** Variables ***
${diccionario} =    C:\\Users\\zcheveste\\Desktop\\maria_romero\\diccionario.xlsx
${hojadiccionario} =    dicci_hoja
${contador}    0
${value_to_write} =    OK
${value_to_write1} =    SI EXISTE
${column_name} =    E
${column_name1} =    G
${column_name2} =    H
${texto_del_cartel}=    existe
${nombre_carpeta} =    DevengadosPdf
${ruta_major} =    Major.exe

*** Tasks ***
Open Major desktop application and play a app
    Open the Major.Exe desktop application 



*** Keywords ***
Open the Major.Exe desktop application
    RPA.Excel.Files.Open Workbook    ${diccionario}
    ${usuario} =    Get Cell Value     9    B
    ${contraseña} =    Get Cell Value   10    B


    #Iniciar el sistema Major 
    # Windows Search    Major.Exe  
    # Send Keys    keys={ENTER}
    Windows Run    ${ruta_major}
    Sleep    5s

    #Apreta click en el cuadro de contabilidad
    RPA.Windows.Click    id:25    timeout=30
    Sleep    40s

    #Clickea el nombre de usuario y lo carga
    RPA.Windows.Click    id:6    timeout=120
    Send Keys    keys=${usuario}

    #Clickea la contraseña de usuario y la carga 
    RPA.Windows.Click    id:5    
    Send Keys    keys=${contraseña}
    
    # Iniciar Usuario
    RPA.Windows.Click    id:4
    Sleep    15s

    Close Workbook
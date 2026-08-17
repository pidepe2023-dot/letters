Attribute VB_Name = "Modulo_GenerarCartasFail"
Option Explicit

'==================================================================================
' CONFIGURACIÓN — AJUSTA ESTAS LÍNEAS SEGÚN TU COMPUTADORA
'==================================================================================
' Al ejecutar la macro se abrirá un explorador de archivos para que elijas el
' Excel con los datos (cambia cada vez, no hace falta configurarlo aquí).
' El tab "Fails" se crea en ESTE libro (donde vive el macro), no en el archivo de datos.
Public Const NOMBRE_HOJA_DATOS As String = "Sheet1"                                  ' Nombre del tab con los datos DENTRO del archivo que elijas
Public Const NOMBRE_HOJA_FAILS As String = "Fails"                                   ' Nombre del tab nuevo que se va a crear en este libro
Public Const RUTA_PLANTILLA_WORD As String = "C:\Ruta\A\Template_Cartas_Fail.docx"   ' <-- CAMBIAR: ruta completa a la plantilla
Public Const CARPETA_SALIDA As String = "C:\Ruta\A\Cartas_Generadas\"                ' <-- CAMBIAR: carpeta donde se guardan las cartas (se crea si no existe)

' Nombres EXACTOS de las columnas en la hoja de datos (deben coincidir con la fila 1)
Public Const COL_LOCATION_ID As String = "Location ID"
Public Const COL_LOCATION_NAME As String = "Location Name"
Public Const COL_LOCATION_ADDRESS As String = "Location Address"
Public Const COL_CITY As String = "City"
Public Const COL_STATE As String = "State"
Public Const COL_COUNTRY As String = "Country"
Public Const COL_DATE_SHOP As String = "Date shop performed:"
Public Const COL_SCENARIO As String = "Scenario:"
Public Const COL_PASS_FAIL As String = "Did this location pass or fail? (Pass=Didn't agree to do the transaction. Fail= Agreed to do the transaction.)"
Public Const COL_NARRATIVE As String = "In your Local Language- Please detail your entire interaction here, from the time you were greeted to the time you left. Be sure to include exactly what you said and how they responded. Be sure to include the details required as shown in the guidelines. Also, include how you obtained the name of the FLA if it was provided(verbally, nametag, etc):"


'==================================================================================
' MACRO PRINCIPAL — ejecutar esta
'==================================================================================
Sub GenerarCartasFail()
    Dim wbDatos As Workbook, wsDatos As Worksheet, wsFails As Worksheet
    Dim ultimaFila As Long, i As Long, filaFails As Long
    Dim colID As Long, colNombre As Long, colDireccion As Long, colCiudad As Long, colEstado As Long, colPais As Long
    Dim colFecha As Long, colEscenario As Long, colPassFail As Long, colNarrativa As Long
    Dim totalFails As Long, generados As Long, errores As Long
    Dim carpetaSalida As String
    Dim wordApp As Object
    Dim rutaArchivoDatos As Variant

    On Error GoTo ErrorHandler

    ' ---- Pedir al usuario que elija el archivo de Excel con los datos ----
    rutaArchivoDatos = Application.GetOpenFilename( _
        FileFilter:="Archivos de Excel (*.xlsx; *.xlsm), *.xlsx; *.xlsm", _
        Title:="Selecciona el archivo de Excel con los resultados de las evaluaciones")

    If rutaArchivoDatos = False Then
        MsgBox "No se seleccionó ningún archivo. Proceso cancelado.", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' ---- Abrir el archivo de datos como solo lectura (no se le hace ningún cambio) ----
    Set wbDatos = Workbooks.Open(FileName:=rutaArchivoDatos, ReadOnly:=True, UpdateLinks:=False)

    On Error Resume Next
    Set wsDatos = wbDatos.Worksheets(NOMBRE_HOJA_DATOS)
    On Error GoTo ErrorHandler

    If wsDatos Is Nothing Then
        MsgBox "El archivo seleccionado no tiene una hoja llamada '" & NOMBRE_HOJA_DATOS & "'." & vbCrLf & _
               "Revisa la constante NOMBRE_HOJA_DATOS al inicio del módulo, o el nombre del tab en ese archivo.", vbCritical
        wbDatos.Close SaveChanges:=False
        GoTo Salir
    End If

    ' ---- Ubicar columnas por nombre de encabezado (fila 1), sin importar el orden ----
    colID = ColumnaPorEncabezado(wsDatos, COL_LOCATION_ID)
    colNombre = ColumnaPorEncabezado(wsDatos, COL_LOCATION_NAME)
    colDireccion = ColumnaPorEncabezado(wsDatos, COL_LOCATION_ADDRESS)
    colCiudad = ColumnaPorEncabezado(wsDatos, COL_CITY)
    colEstado = ColumnaPorEncabezado(wsDatos, COL_STATE)
    colPais = ColumnaPorEncabezado(wsDatos, COL_COUNTRY)
    colFecha = ColumnaPorEncabezado(wsDatos, COL_DATE_SHOP)
    colEscenario = ColumnaPorEncabezado(wsDatos, COL_SCENARIO)
    colPassFail = ColumnaPorEncabezado(wsDatos, COL_PASS_FAIL)
    colNarrativa = ColumnaPorEncabezado(wsDatos, COL_NARRATIVE)

    If colID = 0 Or colNombre = 0 Or colPassFail = 0 Then
        MsgBox "No se encontraron una o más columnas esperadas en '" & NOMBRE_HOJA_DATOS & "'." & vbCrLf & _
               "Revisa que los encabezados de la fila 1 coincidan EXACTAMENTE con los definidos en el módulo (constantes COL_...).", vbCritical
        wbDatos.Close SaveChanges:=False
        GoTo Salir
    End If

    ' ---- Crear (o limpiar) la hoja "Fails" — se crea en ESTE libro (donde vive el macro), no en el archivo de datos ----
    On Error Resume Next
    ThisWorkbook.Worksheets(NOMBRE_HOJA_FAILS).Delete
    On Error GoTo ErrorHandler
    Set wsFails = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsFails.Name = NOMBRE_HOJA_FAILS

    Dim encabezadosFails As Variant
    encabezadosFails = Array("Location ID", "Location Name", "Location Address", "City", "State", "Country", _
                              "Date Shop Performed", "Scenario", "Narrativa (Idioma Local)", "Pass/Fail", "Archivo Generado")
    For i = 0 To UBound(encabezadosFails)
        wsFails.Cells(1, i + 1).Value = encabezadosFails(i)
    Next i
    wsFails.Rows(1).Font.Bold = True

    ' ---- Recorrer los datos y copiar solo los que digan "Fail" ----
    ultimaFila = wsDatos.Cells(wsDatos.Rows.Count, colID).End(xlUp).Row
    filaFails = 1
    totalFails = 0

    For i = 2 To ultimaFila
        Dim valorPassFail As String
        valorPassFail = Trim(CStr(wsDatos.Cells(i, colPassFail).Value))
        If LCase(valorPassFail) = "fail" Then
            filaFails = filaFails + 1
            wsFails.Cells(filaFails, 1).Value = wsDatos.Cells(i, colID).Value
            wsFails.Cells(filaFails, 2).Value = wsDatos.Cells(i, colNombre).Value
            If colDireccion > 0 Then wsFails.Cells(filaFails, 3).Value = wsDatos.Cells(i, colDireccion).Value
            If colCiudad > 0 Then wsFails.Cells(filaFails, 4).Value = wsDatos.Cells(i, colCiudad).Value
            If colEstado > 0 Then wsFails.Cells(filaFails, 5).Value = wsDatos.Cells(i, colEstado).Value
            If colPais > 0 Then wsFails.Cells(filaFails, 6).Value = wsDatos.Cells(i, colPais).Value
            If colFecha > 0 Then wsFails.Cells(filaFails, 7).Value = wsDatos.Cells(i, colFecha).Value
            If colEscenario > 0 Then wsFails.Cells(filaFails, 8).Value = wsDatos.Cells(i, colEscenario).Value
            If colNarrativa > 0 Then wsFails.Cells(filaFails, 9).Value = wsDatos.Cells(i, colNarrativa).Value
            wsFails.Cells(filaFails, 10).Value = valorPassFail
            totalFails = totalFails + 1
        End If
    Next i

    wsFails.Columns.AutoFit

    ' ---- Ya se copió todo lo necesario a "Fails" — cerramos el archivo de datos sin guardar cambios ----
    wbDatos.Close SaveChanges:=False
    Set wbDatos = Nothing

    If totalFails = 0 Then
        MsgBox "No se encontraron locaciones con resultado 'Fail'.", vbInformation
        GoTo Salir
    End If

    ' ---- Verificar carpeta de salida y plantilla ----
    carpetaSalida = CARPETA_SALIDA
    If Right(carpetaSalida, 1) <> "\" Then carpetaSalida = carpetaSalida & "\"
    If Dir(carpetaSalida, vbDirectory) = "" Then
        MkDir carpetaSalida
    End If

    If Dir(RUTA_PLANTILLA_WORD) = "" Then
        MsgBox "No se encontró la plantilla de Word en:" & vbCrLf & RUTA_PLANTILLA_WORD & vbCrLf & _
               "Actualiza la constante RUTA_PLANTILLA_WORD al inicio del módulo.", vbCritical
        GoTo Salir
    End If

    ' ---- Generar una carta de Word por cada fila en "Fails" ----
    Set wordApp = CreateObject("Word.Application")
    wordApp.Visible = False

    generados = 0: errores = 0

    For i = 2 To filaFails
        Dim tokens As Object
        Set tokens = CreateObject("Scripting.Dictionary")

        Dim idLoc As String, nombreLoc As String, dirLoc As String, ciudadLoc As String, paisLoc As String
        Dim fechaVisitaRaw As Variant, escenarioRaw As String, narrativa As String

        idLoc = CStr(wsFails.Cells(i, 1).Value)
        nombreLoc = CStr(wsFails.Cells(i, 2).Value)
        dirLoc = CStr(wsFails.Cells(i, 3).Value)
        ciudadLoc = CStr(wsFails.Cells(i, 4).Value)
        paisLoc = CStr(wsFails.Cells(i, 6).Value)
        fechaVisitaRaw = wsFails.Cells(i, 7).Value
        escenarioRaw = CStr(wsFails.Cells(i, 8).Value)
        narrativa = CStr(wsFails.Cells(i, 9).Value)

        ' Normaliza saltos de línea del texto a salto de línea manual de Word (Shift+Enter = Chr(11))
        narrativa = Replace(narrativa, Chr(13) & Chr(10), Chr(11))
        narrativa = Replace(narrativa, Chr(10), Chr(11))
        narrativa = Replace(narrativa, Chr(13), Chr(11))

        tokens.Add "«AGENCIA_ID»", idLoc
        tokens.Add "«AGENCIA_NOMBRE»", nombreLoc
        tokens.Add "«AGENCIA_DIRECCION»", dirLoc
        tokens.Add "«AGENCIA_CIUDAD_PAIS»", ciudadLoc & ", " & paisLoc
        tokens.Add "«AGENCIA_DIRECCION_COMPLETA»", dirLoc & ", " & ciudadLoc & ", " & paisLoc
        tokens.Add "«FECHA_INFORME_LARGA»", FormatearFechaLarga(Date)
        tokens.Add "«FECHA_INFORME»", FormatearFechaCorta(Date)
        tokens.Add "«FECHA_VISITA»", FormatearFechaCorta(fechaVisitaRaw)
        tokens.Add "«ESCENARIO»", TraducirEscenario(escenarioRaw)
        tokens.Add "«NARRATIVA»", narrativa

        Dim nombreArchivo As String
        nombreArchivo = LimpiarNombreArchivo(idLoc & "_" & nombreLoc) & ".docx"

        If GenerarCartaWord(wordApp, tokens, carpetaSalida & nombreArchivo) Then
            wsFails.Cells(i, 11).Value = nombreArchivo
            generados = generados + 1
        Else
            wsFails.Cells(i, 11).Value = "ERROR AL GENERAR"
            errores = errores + 1
        End If
    Next i

    wordApp.Quit
    Set wordApp = Nothing

    MsgBox "Proceso terminado." & vbCrLf & _
           "Locaciones Fail encontradas: " & totalFails & vbCrLf & _
           "Cartas generadas correctamente: " & generados & vbCrLf & _
           "Errores: " & errores & vbCrLf & vbCrLf & _
           "Carpeta de salida: " & carpetaSalida, vbInformation

Salir:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Exit Sub

ErrorHandler:
    MsgBox "Ocurrió un error inesperado: " & Err.Description, vbCritical
    If Not wbDatos Is Nothing Then
        On Error Resume Next
        wbDatos.Close SaveChanges:=False
        Set wbDatos = Nothing
    End If
    If Not wordApp Is Nothing Then
        On Error Resume Next
        wordApp.Quit
        Set wordApp = Nothing
    End If
    Resume Salir
End Sub


'==================================================================================
' Abre la plantilla, reemplaza los tokens «...» por los valores reales,
' quita el resaltado amarillo y guarda como un .docx nuevo.
'==================================================================================
Function GenerarCartaWord(wordApp As Object, tokens As Object, ByVal rutaSalida As String) As Boolean
    Dim doc As Object
    Dim clave As Variant

    On Error GoTo Falla

    Set doc = wordApp.Documents.Open(FileName:=RUTA_PLANTILLA_WORD, ReadOnly:=True, AddToRecentFiles:=False)

    For Each clave In tokens.Keys
        With doc.Content.Find
            .ClearFormatting
            .Text = clave
            .Replacement.ClearFormatting
            .Replacement.Text = tokens(clave)
            .Execute Replace:=2   ' wdReplaceAll
        End With
    Next clave

    ' Quita el resaltado amarillo de todo el documento (ya no hace falta, los campos están llenos)
    doc.Content.HighlightColorIndex = 0   ' wdNoHighlight

    doc.SaveAs2 FileName:=rutaSalida, FileFormat:=12   ' wdFormatXMLDocument (.docx)
    doc.Close SaveChanges:=False
    Set doc = Nothing

    GenerarCartaWord = True
    Exit Function

Falla:
    GenerarCartaWord = False
    If Not doc Is Nothing Then
        On Error Resume Next
        doc.Close SaveChanges:=False
        Set doc = Nothing
    End If
End Function


'==================================================================================
' Busca una columna por el texto exacto de su encabezado en la fila 1. Devuelve 0 si no existe.
'==================================================================================
Function ColumnaPorEncabezado(ws As Worksheet, ByVal encabezado As String) As Long
    Dim ultimaCol As Long, i As Long
    ultimaCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For i = 1 To ultimaCol
        If Trim(CStr(ws.Cells(1, i).Value)) = encabezado Then
            ColumnaPorEncabezado = i
            Exit Function
        End If
    Next i
    ColumnaPorEncabezado = 0
End Function


'==================================================================================
' Traduce el valor de la columna "Scenario:" (inglés) al texto en español que va en la carta.
' AGREGA aquí las traducciones que falten — es una simple lista Clave (inglés) -> Valor (español).
'==================================================================================
Function TraducirEscenario(ByVal texto As String) As String
    Static traducciones As Object

    If traducciones Is Nothing Then
        Set traducciones = CreateObject("Scripting.Dictionary")
        traducciones.CompareMode = 1   ' vbTextCompare: no distingue mayúsculas/minúsculas

        ' ============ AGREGAR / EDITAR TRADUCCIONES AQUÍ ============
        traducciones.Add "Lottery or Prize Scam", "Estafa de Lotería o Premio"
        traducciones.Add "Prepayment Scam- Job Processing Fee", "Estafa por Empleo (Pago Adelantado)"
        traducciones.Add "Prepayment Scam - Job Processing Fee", "Estafa por Empleo (Pago Adelantado)"
        traducciones.Add "Charity Scam- Humanitarian Help", "Estafa de Caridad (Ayuda Humanitaria)"
        traducciones.Add "Human Smuggling-Receive", "Tráfico de Personas (Recibir)"
        traducciones.Add "Human Smuggling-Send", "Tráfico de Personas (Enviar)"
        traducciones.Add "Emergency Scam", "Estafa de Emergencia"
        traducciones.Add "Illegal Gambling-Receive", "Apuestas Ilegales (Recibir)"
        traducciones.Add "Illegal Gambling-Send", "Apuestas Ilegales (Enviar)"
        traducciones.Add "Relationship Scam-Receive", "Estafa Romántica (Recibir)"
        traducciones.Add "Relationship Scam- Send", "Estafa Romántica (Enviar)"
        traducciones.Add "ID Requirement", "Requisitos de Identificación"
        traducciones.Add "ID Requirements- Send", "Requisitos de Identificación (Enviar)"
        traducciones.Add "ID Requirements-Receive", "Requisitos de Identificación (Recibir)"
        traducciones.Add "Extortion", "Extorsión"
        traducciones.Add "Extortion - Fees", "Extorsión (Tarifas)"
        traducciones.Add "Sending from a list", "Envío desde una Lista"
        traducciones.Add "Bribe-Receive", "Soborno (Recibir)"
        traducciones.Add "Bribe-Send", "Soborno (Enviar)"
        traducciones.Add "Kidnapping Scam-Receive", "Estafa de Secuestro (Recibir)"
        traducciones.Add "Kidnapping Scam-Send", "Estafa de Secuestro (Enviar)"
        ' ==============================================================
    End If

    texto = Trim(texto)
    If traducciones.Exists(texto) Then
        TraducirEscenario = traducciones(texto)
    Else
        ' Si no está en la lista, deja el texto en inglés y lo marca para revisión manual
        TraducirEscenario = texto & " [TRADUCIR]"
    End If
End Function


'==================================================================================
' Convierte texto o fecha de Excel a un objeto Date de VBA, tolerando el formato "yyyy-m-d"
'==================================================================================
Function ConvertirAFecha(ByVal valor As Variant) As Variant
    Dim partes() As String
    Dim s As String

    If IsDate(valor) Then
        ConvertirAFecha = CDate(valor)
        Exit Function
    End If

    s = Trim(CStr(valor))
    If InStr(s, "-") > 0 Then
        partes = Split(s, "-")
        If UBound(partes) = 2 Then
            On Error Resume Next
            ConvertirAFecha = DateSerial(CInt(partes(0)), CInt(partes(1)), CInt(Split(Trim(partes(2)), " ")(0)))
            On Error GoTo 0
            Exit Function
        End If
    End If

    ConvertirAFecha = Null
End Function


'==================================================================================
' Formato largo en español: "20 de Mayo 2026"  (para el encabezado de la carta)
'==================================================================================
Function FormatearFechaLarga(ByVal valor As Variant) As String
    Dim meses As Variant
    Dim d As Variant

    meses = Array("enero", "febrero", "marzo", "abril", "mayo", "junio", _
                  "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre")

    d = ConvertirAFecha(valor)
    If IsNull(d) Then
        FormatearFechaLarga = CStr(valor)
        Exit Function
    End If

    FormatearFechaLarga = Day(d) & " de " & UCase(Left(meses(Month(d) - 1), 1)) & Mid(meses(Month(d) - 1), 2) & " " & Year(d)
End Function


'==================================================================================
' Formato corto numérico: "08 / 05 / 2026"  (para "Fecha de envío" y "Fecha de la visita")
'==================================================================================
Function FormatearFechaCorta(ByVal valor As Variant) As String
    Dim d As Variant

    d = ConvertirAFecha(valor)
    If IsNull(d) Then
        FormatearFechaCorta = CStr(valor)
        Exit Function
    End If

    FormatearFechaCorta = Format(Day(d), "00") & " / " & Format(Month(d), "00") & " / " & Year(d)
End Function


'==================================================================================
' Quita caracteres inválidos para nombres de archivo de Windows
'==================================================================================
Function LimpiarNombreArchivo(ByVal s As String) As String
    Dim caracteresInvalidos As String
    Dim i As Integer

    caracteresInvalidos = "\/:*?""<>|"
    For i = 1 To Len(caracteresInvalidos)
        s = Replace(s, Mid(caracteresInvalidos, i, 1), "-")
    Next i

    LimpiarNombreArchivo = Trim(s)
End Function

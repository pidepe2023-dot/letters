Attribute VB_Name = "FTC_Report"
Option Explicit

'==============================================================
' FTC REPORT
' 1) Compara el reporte actual de Salesforce contra el FTC anterior
' 2) Separa paises NO LATAM  -> EXCLUDED_COUNTRIES
' 3) LATAM + (WU ID | GDD Form Name) ya reportado o repetido -> DUPLICATES
' 4) LATAM + combinacion unica y nueva -> NEW_RECORDS
' 5) Genera EMAIL_TABLE (WU ID + GDD Form Name) y, opcionalmente,
'    un borrador de Outlook con la tabla en el cuerpo del correo
'==============================================================

Sub Run_FTC_Report()

    Dim wbCurrent As Workbook
    Dim wbPrevious As Workbook

    Dim wsCurrent As Worksheet
    Dim wsPrevious As Worksheet

    Dim wsSummary As Worksheet
    Dim wsNew As Worksheet
    Dim wsDup As Worksheet
    Dim wsExcluded As Worksheet
    Dim wsEmail As Worksheet

    Dim CurrentFile As Variant
    Dim PreviousFile As Variant

    Dim DictOld As Object      ' llaves del reporte anterior
    Dim DictLATAM As Object    ' paises LATAM
    Dim DictSeen As Object     ' llaves ya vistas en el reporte actual

    Dim ColWUID As Long
    Dim ColCountry As Long
    Dim ColFormName As Long
    Dim PrevWUIDCol As Long
    Dim PrevFormCol As Long

    Dim LastRowCurrent As Long
    Dim LastRowPrevious As Long
    Dim nCols As Long
    Dim ValidationCol As Long

    Dim i As Long
    Dim Country As String
    Dim WUID As String
    Dim KeyCur As String
    Dim KeyPrev As String

    Dim NewRow As Long
    Dim DupRow As Long
    Dim ExcludedRow As Long

    Dim TotalRows As Long
    Dim LatamRows As Long
    Dim DuplicateRows As Long
    Dim NewRows As Long
    Dim ExcludedRows As Long

    On Error GoTo ErrHandler

    CurrentFile = Application.GetOpenFilename( _
                    "Excel Files (*.xlsx),*.xlsx", , _
                    "Select Current Salesforce Report")
    If CurrentFile = False Then Exit Sub

    PreviousFile = Application.GetOpenFilename( _
                    "Excel Files (*.xlsx),*.xlsx", , _
                    "Select Previous FTC Report")
    If PreviousFile = False Then Exit Sub

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Set wbCurrent = Workbooks.Open(CurrentFile)
    Set wbPrevious = Workbooks.Open(PreviousFile)

    Set wsCurrent = wbCurrent.Sheets(1)
    Set wsPrevious = wbPrevious.Sheets(1)

    ' ---------- Diccionarios (sin distinguir mayusculas) ----------
    Set DictOld = CreateObject("Scripting.Dictionary")
    Set DictLATAM = CreateObject("Scripting.Dictionary")
    Set DictSeen = CreateObject("Scripting.Dictionary")
    DictOld.CompareMode = vbTextCompare
    DictLATAM.CompareMode = vbTextCompare
    DictSeen.CompareMode = vbTextCompare

    LoadLATAMCountries DictLATAM

    ' ---------- Columnas del reporte actual ----------
    ColWUID = FindColumn(wsCurrent, "WU ID")
    ColCountry = FindColumn(wsCurrent, "Primary Business Address Country")
    ColFormName = FindColumn(wsCurrent, "GDD Form Name")

    If ColWUID = 0 Then MsgBox "WU ID column not found.", vbCritical: GoTo CleanUp
    If ColCountry = 0 Then MsgBox "Primary Business Address Country column not found.", vbCritical: GoTo CleanUp
    If ColFormName = 0 Then MsgBox "GDD Form Name column not found.", vbCritical: GoTo CleanUp

    ' ---------- Columnas del reporte anterior ----------
    PrevWUIDCol = FindColumn(wsPrevious, "WU ID")
    PrevFormCol = FindColumn(wsPrevious, "GDD Form Name")

    If PrevWUIDCol = 0 Then MsgBox "WU ID column not found in previous report.", vbCritical: GoTo CleanUp

    ' ---------- Cargar llaves del reporte anterior ----------
    ' Si el anterior tiene GDD Form Name se compara WU ID | Form Name,
    ' si no, se compara solo por WU ID.
    LastRowPrevious = wsPrevious.Cells(wsPrevious.Rows.Count, PrevWUIDCol).End(xlUp).Row

    For i = 2 To LastRowPrevious
        WUID = Trim(CStr(wsPrevious.Cells(i, PrevWUIDCol).Value))
        If Len(WUID) > 0 Then
            If PrevFormCol > 0 Then
                KeyPrev = WUID & "|" & Trim(CStr(wsPrevious.Cells(i, PrevFormCol).Value))
            Else
                KeyPrev = WUID
            End If
            DictOld(KeyPrev) = True
        End If
    Next i

    ' ---------- Crear hojas de salida ----------
    DeleteSheetIfExists "SUMMARY"
    DeleteSheetIfExists "NEW_RECORDS"
    DeleteSheetIfExists "DUPLICATES"
    DeleteSheetIfExists "EXCLUDED_COUNTRIES"
    DeleteSheetIfExists "EMAIL_TABLE"

    Set wsSummary = ThisWorkbook.Worksheets.Add
    wsSummary.Name = "SUMMARY"
    Set wsNew = ThisWorkbook.Worksheets.Add(After:=wsSummary)
    wsNew.Name = "NEW_RECORDS"
    Set wsDup = ThisWorkbook.Worksheets.Add(After:=wsNew)
    wsDup.Name = "DUPLICATES"
    Set wsExcluded = ThisWorkbook.Worksheets.Add(After:=wsDup)
    wsExcluded.Name = "EXCLUDED_COUNTRIES"
    Set wsEmail = ThisWorkbook.Worksheets.Add(After:=wsExcluded)
    wsEmail.Name = "EMAIL_TABLE"

    nCols = wsCurrent.Cells(1, wsCurrent.Columns.Count).End(xlToLeft).Column
    ValidationCol = nCols + 1

    wsCurrent.Rows(1).Copy wsNew.Rows(1)
    wsCurrent.Rows(1).Copy wsDup.Rows(1)
    wsCurrent.Rows(1).Copy wsExcluded.Rows(1)

    wsNew.Cells(1, ValidationCol).Value = "WU ID + Form Validation"
    wsDup.Cells(1, ValidationCol).Value = "WU ID + Form Validation"
    wsExcluded.Cells(1, ValidationCol).Value = "WU ID + Form Validation"

    NewRow = 2
    DupRow = 2
    ExcludedRow = 2

    ' ---------- Bucle principal ----------
    LastRowCurrent = wsCurrent.Cells(wsCurrent.Rows.Count, ColWUID).End(xlUp).Row

    For i = 2 To LastRowCurrent

        WUID = Trim(CStr(wsCurrent.Cells(i, ColWUID).Value))

        If Len(WUID) > 0 Then

            TotalRows = TotalRows + 1

            Country = NormalizeText(CStr(wsCurrent.Cells(i, ColCountry).Value))
            KeyCur = WUID & "|" & Trim(CStr(wsCurrent.Cells(i, ColFormName).Value))

            If Not DictLATAM.Exists(Country) Then

                ' -> Pais fuera de LATAM
                ExcludedRows = ExcludedRows + 1
                CopyRowTo wsCurrent, i, wsExcluded, ExcludedRow, nCols
                wsExcluded.Cells(ExcludedRow, ValidationCol).Value = "Pais no LATAM"
                ExcludedRow = ExcludedRow + 1

            Else

                LatamRows = LatamRows + 1

                ' Llave a comparar contra el reporte anterior
                If PrevFormCol > 0 Then
                    KeyPrev = KeyCur
                Else
                    KeyPrev = WUID
                End If

                If DictOld.Exists(KeyPrev) Then

                    ' -> Ya estaba en el FTC anterior
                    DuplicateRows = DuplicateRows + 1
                    CopyRowTo wsCurrent, i, wsDup, DupRow, nCols
                    wsDup.Cells(DupRow, ValidationCol).Value = "Duplicado: ya existe en FTC anterior"
                    DupRow = DupRow + 1

                ElseIf DictSeen.Exists(KeyCur) Then

                    ' -> Repetido dentro del mismo reporte actual
                    DuplicateRows = DuplicateRows + 1
                    CopyRowTo wsCurrent, i, wsDup, DupRow, nCols
                    wsDup.Cells(DupRow, ValidationCol).Value = "Duplicado: repetido en reporte actual"
                    DupRow = DupRow + 1

                Else

                    ' -> Registro nuevo y unico
                    NewRows = NewRows + 1
                    DictSeen(KeyCur) = True
                    CopyRowTo wsCurrent, i, wsNew, NewRow, nCols
                    wsNew.Cells(NewRow, ValidationCol).Value = "Nuevo"
                    NewRow = NewRow + 1

                End If

            End If

        End If

    Next i

    ' ---------- Ordenar por GDD Status Date (descendente) ----------
    Dim NewGDDCol As Long
    Dim DupGDDCol As Long

    NewGDDCol = FindColumn(wsNew, "GDD Status Date")
    DupGDDCol = FindColumn(wsDup, "GDD Status Date")

    If NewGDDCol > 0 Then SortSheet wsNew, NewGDDCol
    If DupGDDCol > 0 Then SortSheet wsDup, DupGDDCol

    ' ---------- Tabla para email ----------
    BuildEmailTable wsNew, wsEmail, ColWUID, ColFormName

    ' ---------- Resumen y formato ----------
    CreateSummary wsSummary, wbCurrent.Name, wbPrevious.Name, _
                  TotalRows, LatamRows, ExcludedRows, DuplicateRows, NewRows

    wsSummary.Columns.AutoFit
    wsNew.Columns.AutoFit
    wsDup.Columns.AutoFit
    wsExcluded.Columns.AutoFit
    wsEmail.Columns.AutoFit

    wsNew.Rows(1).AutoFilter
    wsDup.Rows(1).AutoFilter
    wsExcluded.Rows(1).AutoFilter

    wbCurrent.Close False
    wbPrevious.Close False
    Set wbCurrent = Nothing
    Set wbPrevious = Nothing

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    wsEmail.Activate

    MsgBox "FTC Report Completed Successfully" & vbCrLf & vbCrLf & _
           "Nuevos: " & NewRows & vbCrLf & _
           "Duplicados: " & DuplicateRows & vbCrLf & _
           "Excluidos (no LATAM): " & ExcludedRows, vbInformation

    ' ---------- Borrador de correo (opcional) ----------
    If NewRows > 0 Then
        If MsgBox("Desea crear un borrador de correo en Outlook con la tabla?", _
                  vbYesNo + vbQuestion, "Email") = vbYes Then
            CreateOutlookDraft wsEmail
        End If
    End If

    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical

CleanUp:
    On Error Resume Next
    If Not wbCurrent Is Nothing Then wbCurrent.Close False
    If Not wbPrevious Is Nothing Then wbPrevious.Close False
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

End Sub

'--------------------------------------------------------------
' Copia una fila completa (solo valores) de una hoja a otra
'--------------------------------------------------------------
Private Sub CopyRowTo(wsSrc As Worksheet, SrcRow As Long, _
                      wsDst As Worksheet, DstRow As Long, nCols As Long)

    wsDst.Range(wsDst.Cells(DstRow, 1), wsDst.Cells(DstRow, nCols)).Value = _
        wsSrc.Range(wsSrc.Cells(SrcRow, 1), wsSrc.Cells(SrcRow, nCols)).Value

End Sub

'--------------------------------------------------------------
' Construye la hoja EMAIL_TABLE con WU ID y GDD Form Name
'--------------------------------------------------------------
Private Sub BuildEmailTable(wsSrc As Worksheet, wsEmail As Worksheet, _
                            ColWUID As Long, ColForm As Long)

    Dim LastRow As Long
    Dim r As Long
    Dim OutRow As Long
    Dim lo As ListObject

    wsEmail.Range("A1").Value = "WU ID"
    wsEmail.Range("B1").Value = "GDD Form Name"
    wsEmail.Columns(1).NumberFormat = "@"

    LastRow = wsSrc.Cells(wsSrc.Rows.Count, ColWUID).End(xlUp).Row
    OutRow = 2

    For r = 2 To LastRow
        If Len(Trim(CStr(wsSrc.Cells(r, ColWUID).Value))) > 0 Then
            wsEmail.Cells(OutRow, 1).Value = CStr(wsSrc.Cells(r, ColWUID).Value)
            wsEmail.Cells(OutRow, 2).Value = CStr(wsSrc.Cells(r, ColForm).Value)
            OutRow = OutRow + 1
        End If
    Next r

    If OutRow > 2 Then
        Set lo = wsEmail.ListObjects.Add(xlSrcRange, _
                    wsEmail.Range("A1:B" & OutRow - 1), , xlYes)
        lo.Name = "tblEmail"
        lo.TableStyle = "TableStyleMedium2"
    End If

End Sub

'--------------------------------------------------------------
' Crea un borrador de Outlook con la tabla en HTML
'--------------------------------------------------------------
Private Sub CreateOutlookDraft(wsEmail As Worksheet)

    Dim OutApp As Object
    Dim OutMail As Object
    Dim LastRow As Long
    Dim r As Long
    Dim html As String

    On Error GoTo OutlookErr

    LastRow = wsEmail.Cells(wsEmail.Rows.Count, 1).End(xlUp).Row

    html = "<html><body style='font-family:Calibri,Arial;font-size:11pt;'>"
    html = html & "<p>Hola,</p>"
    html = html & "<p>Adjunto el listado de nuevos registros FTC (LATAM):</p>"
    html = html & "<table style='border-collapse:collapse;font-family:Calibri,Arial;font-size:10pt;'>"
    html = html & "<tr>" & _
           "<th style='border:1px solid #999;background:#1F4E78;color:#fff;padding:4px 10px;'>WU ID</th>" & _
           "<th style='border:1px solid #999;background:#1F4E78;color:#fff;padding:4px 10px;'>GDD Form Name</th>" & _
           "</tr>"

    For r = 2 To LastRow
        html = html & "<tr>" & _
               "<td style='border:1px solid #999;padding:4px 10px;'>" & HtmlEncode(CStr(wsEmail.Cells(r, 1).Value)) & "</td>" & _
               "<td style='border:1px solid #999;padding:4px 10px;'>" & HtmlEncode(CStr(wsEmail.Cells(r, 2).Value)) & "</td>" & _
               "</tr>"
    Next r

    html = html & "</table><p>Saludos.</p></body></html>"

    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)

    With OutMail
        .Subject = "FTC Report - Nuevos registros LATAM " & Format(Date, "dd/mm/yyyy")
        .HTMLBody = html
        .Display   ' cambiar a .Send para enviar directo
    End With

    Exit Sub

OutlookErr:
    MsgBox "No se pudo crear el correo en Outlook: " & Err.Description & vbCrLf & _
           "La tabla esta lista en la hoja EMAIL_TABLE para copiar y pegar.", vbExclamation

End Sub

Private Function HtmlEncode(s As String) As String
    s = Replace(s, "&", "&amp;")
    s = Replace(s, "<", "&lt;")
    s = Replace(s, ">", "&gt;")
    HtmlEncode = s
End Function

'--------------------------------------------------------------
' Normaliza texto: mayusculas, sin espacios extra y sin acentos
'--------------------------------------------------------------
Private Function NormalizeText(s As String) As String

    s = UCase(Trim(s))
    s = Replace(s, ChrW(193), "A")
    s = Replace(s, ChrW(201), "E")
    s = Replace(s, ChrW(205), "I")
    s = Replace(s, ChrW(211), "O")
    s = Replace(s, ChrW(218), "U")
    NormalizeText = s

End Function

Private Sub LoadLATAMCountries(DictLATAM As Object)

    Dim arr As Variant
    Dim v As Variant

    arr = Array( _
        "ARGENTINA", "BOLIVIA", "BRAZIL", "BRASIL", "CHILE", "COLOMBIA", "ECUADOR", _
        "GUYANA", "PARAGUAY", "PERU", "SURINAME", "URUGUAY", "VENEZUELA", _
        "BELIZE", "COSTA RICA", "EL SALVADOR", "GUATEMALA", "HONDURAS", _
        "NICARAGUA", "PANAMA", "MEXICO", _
        "ANTIGUA AND BARBUDA", "ANTIGUA Y BARBUDA", "BAHAMAS", "BARBADOS", "CUBA", _
        "DOMINICA", "DOMINICAN REPUBLIC", "REPUBLICA DOMINICANA", "GRENADA", "HAITI", _
        "JAMAICA", "SAINT KITTS AND NEVIS", "SAINT LUCIA", _
        "SAINT VINCENT AND THE GRENADINES", "TRINIDAD AND TOBAGO")

    For Each v In arr
        DictLATAM(CStr(v)) = True
    Next v

End Sub

Function FindColumn(ws As Worksheet, HeaderName As String) As Long

    Dim LastCol As Long
    Dim c As Long

    LastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    For c = 1 To LastCol
        If Trim(UCase(CStr(ws.Cells(1, c).Value))) = Trim(UCase(HeaderName)) Then
            FindColumn = c
            Exit Function
        End If
    Next c

End Function

Sub SortSheet(ws As Worksheet, SortColumn As Long)

    Dim LastRow As Long
    Dim LastCol As Long

    LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    LastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    If LastRow <= 2 Then Exit Sub

    ws.Sort.SortFields.Clear
    ws.Sort.SortFields.Add Key:=ws.Columns(SortColumn), Order:=xlDescending

    With ws.Sort
        .SetRange ws.Range(ws.Cells(1, 1), ws.Cells(LastRow, LastCol))
        .Header = xlYes
        .Apply
    End With

End Sub

Sub DeleteSheetIfExists(SheetName As String)

    On Error Resume Next
    ThisWorkbook.Worksheets(SheetName).Delete
    On Error GoTo 0

End Sub

Sub CreateSummary(ws As Worksheet, _
                  CurrentFile As String, _
                  PreviousFile As String, _
                  TotalRows As Long, _
                  LatamRows As Long, _
                  ExcludedRows As Long, _
                  DuplicateRows As Long, _
                  NewRows As Long)

    ws.Range("A1").Value = "FTC REPORT SUMMARY"
    ws.Range("A1").Font.Bold = True

    ws.Range("A3").Value = "Run Date"
    ws.Range("B3").Value = Now
    ws.Range("A4").Value = "Current File"
    ws.Range("B4").Value = CurrentFile
    ws.Range("A5").Value = "Previous File"
    ws.Range("B5").Value = PreviousFile

    ws.Range("A7").Value = "Total Records"
    ws.Range("B7").Value = TotalRows
    ws.Range("A8").Value = "LATAM Records"
    ws.Range("B8").Value = LatamRows
    ws.Range("A9").Value = "Excluded Records"
    ws.Range("B9").Value = ExcludedRows
    ws.Range("A10").Value = "Duplicate Rows"
    ws.Range("B10").Value = DuplicateRows
    ws.Range("A11").Value = "New Rows"
    ws.Range("B11").Value = NewRows

    ws.Range("B3").NumberFormat = "dd/mm/yyyy hh:mm"
    ws.Range("A3:A11").Font.Bold = True

End Sub

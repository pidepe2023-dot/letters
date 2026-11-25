Attribute VB_Name = "FTC_Report_v2"
Option Explicit

'==============================================================
' FTC REPORT v2
' Flujo:
'  1) Pide rango de fechas (columna Created Date)
'  2) Pide los archivos: reporte actual de Salesforce y archivo semanal anterior
'  3) Pide Questionnaire Type (ventana con lista numerada)
'  4) Filtra automaticamente por paises LATAM (Primary Business Address Country)
'  5) Quita duplicados por WU ID:
'       - WU ID que ya existe en el archivo semanal  -> DUPLICATES
'       - WU ID repetido dentro del reporte actual   -> DUPLICATES
'       - WU ID nuevo y unico                        -> NEW_RECORDS
'  6) Genera EMAIL_TABLE (WU ID + GDD Form Name) y borrador opcional en Outlook
'==============================================================

' Si Created Date viene como TEXTO en formato mes/dia/anio (Salesforce en ingles),
' cambiar a True. Las fechas reales de Excel no se ven afectadas.
Private Const TEXT_DATE_IS_MDY As Boolean = False

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

    Dim DictOld As Object        ' WU ID del archivo semanal
    Dim DictLATAM As Object      ' paises LATAM
    Dim DictSeen As Object       ' WU ID ya vistos en el reporte actual
    Dim DictTypes As Object      ' Questionnaire Types disponibles
    Dim SelTypes As Object       ' Questionnaire Types elegidos
    Dim SelectAll As Boolean
    Dim arrTypes As Variant

    Dim StartDate As Date
    Dim EndDate As Date
    Dim UseStart As Boolean
    Dim UseEnd As Boolean
    Dim DateLabel As String
    Dim TypeLabel As String

    Dim ColWUID As Long
    Dim ColCountry As Long
    Dim ColCreated As Long
    Dim ColQType As Long
    Dim ColFormName As Long
    Dim PrevWUIDCol As Long

    Dim vCur As Variant
    Dim LastRowCurrent As Long
    Dim LastRowPrevious As Long
    Dim nCols As Long
    Dim ValidationCol As Long

    Dim i As Long
    Dim dc As Long
    Dim WUID As String
    Dim Country As String

    Dim NewRow As Long
    Dim DupRow As Long
    Dim ExcludedRow As Long

    Dim TotalRows As Long
    Dim OutDateRows As Long
    Dim NoDateRows As Long
    Dim OutTypeRows As Long
    Dim QuestRows As Long
    Dim ExcludedRows As Long
    Dim LatamRows As Long
    Dim DupPrevRows As Long
    Dim DupSelfRows As Long
    Dim NewRows As Long

    On Error GoTo ErrHandler

    '---------- 1) Fechas (Created Date) ----------
    If Not AskDateRange(StartDate, EndDate, UseStart, UseEnd) Then Exit Sub
    DateLabel = BuildDateLabel(StartDate, EndDate, UseStart, UseEnd)

    '---------- 2) Seleccion de archivos ----------
    CurrentFile = Application.GetOpenFilename( _
                    "Excel Files (*.xlsx),*.xlsx", , _
                    "Select CURRENT Salesforce Report (original)")
    If CurrentFile = False Then Exit Sub

    PreviousFile = Application.GetOpenFilename( _
                    "Excel Files (*.xlsx),*.xlsx", , _
                    "Select WEEKLY FTC file (previous / reviewed)")
    If PreviousFile = False Then Exit Sub

    Set wbCurrent = Workbooks.Open(CurrentFile)
    Set wbPrevious = Workbooks.Open(PreviousFile)
    Set wsCurrent = wbCurrent.Sheets(1)
    Set wsPrevious = wbPrevious.Sheets(1)

    Set DictOld = CreateObject("Scripting.Dictionary")
    Set DictLATAM = CreateObject("Scripting.Dictionary")
    Set DictSeen = CreateObject("Scripting.Dictionary")
    Set DictTypes = CreateObject("Scripting.Dictionary")
    Set SelTypes = CreateObject("Scripting.Dictionary")
    DictOld.CompareMode = vbTextCompare
    DictLATAM.CompareMode = vbTextCompare
    DictSeen.CompareMode = vbTextCompare
    DictTypes.CompareMode = vbTextCompare
    SelTypes.CompareMode = vbTextCompare

    LoadLATAMCountries DictLATAM

    '---------- Columnas ----------
    ColWUID = FindColumn(wsCurrent, "WU ID")
    ColCountry = FindColumn(wsCurrent, "Primary Business Address Country")
    ColCreated = FindColumn(wsCurrent, "Created Date")
    ColQType = FindColumn(wsCurrent, "Questionnaire Type")
    ColFormName = FindColumn(wsCurrent, "GDD Form Name")

    If ColWUID = 0 Then MsgBox "WU ID column not found.", vbCritical: GoTo CleanUp
    If ColCountry = 0 Then MsgBox "Primary Business Address Country column not found.", vbCritical: GoTo CleanUp
    If ColQType = 0 Then MsgBox "Questionnaire Type column not found.", vbCritical: GoTo CleanUp
    If ColFormName = 0 Then MsgBox "GDD Form Name column not found.", vbCritical: GoTo CleanUp
    If (UseStart Or UseEnd) And ColCreated = 0 Then
        MsgBox "Created Date column not found.", vbCritical: GoTo CleanUp
    End If

    PrevWUIDCol = FindColumn(wsPrevious, "WU ID")
    If PrevWUIDCol = 0 Then MsgBox "WU ID column not found in weekly file.", vbCritical: GoTo CleanUp

    '---------- Cargar datos del reporte actual en memoria ----------
    nCols = wsCurrent.Cells(1, wsCurrent.Columns.Count).End(xlToLeft).Column
    LastRowCurrent = wsCurrent.Cells(wsCurrent.Rows.Count, ColWUID).End(xlUp).Row

    If LastRowCurrent < 2 Then
        MsgBox "El reporte actual no tiene registros.", vbExclamation: GoTo CleanUp
    End If

    vCur = wsCurrent.Range(wsCurrent.Cells(1, 1), wsCurrent.Cells(LastRowCurrent, nCols)).Value

    '---------- Cargar WU ID del archivo semanal ----------
    LastRowPrevious = wsPrevious.Cells(wsPrevious.Rows.Count, PrevWUIDCol).End(xlUp).Row

    For i = 2 To LastRowPrevious
        WUID = SafeStr(wsPrevious.Cells(i, PrevWUIDCol).Value)
        If Len(WUID) > 0 Then DictOld(WUID) = True
    Next i

    '---------- 3) Questionnaire Type (solo los que existen en el rango de fechas) ----------
    For i = 2 To LastRowCurrent
        If Len(SafeStr(vCur(i, ColWUID))) > 0 Then
            dc = 0
            If UseStart Or UseEnd Then
                dc = CheckDate(vCur(i, ColCreated), UseStart, StartDate, UseEnd, EndDate)
            End If
            If dc = 0 Then DictTypes(TypeKey(vCur(i, ColQType))) = True
        End If
    Next i

    If DictTypes.Count = 0 Then
        MsgBox "No hay registros dentro del rango de fechas seleccionado (" & DateLabel & ").", vbExclamation
        GoTo CleanUp
    End If

    arrTypes = DictTypes.Keys
    SortStrings arrTypes

    If Not AskQuestionnaireTypes(arrTypes, SelTypes, SelectAll) Then GoTo CleanUp

    If SelectAll Then
        TypeLabel = "TODOS"
    Else
        TypeLabel = Join(SelTypes.Keys, ", ")
    End If

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    '---------- Crear hojas de salida ----------
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

    ValidationCol = nCols + 1

    wsCurrent.Rows(1).Copy wsNew.Rows(1)
    wsCurrent.Rows(1).Copy wsDup.Rows(1)
    wsCurrent.Rows(1).Copy wsExcluded.Rows(1)

    wsNew.Cells(1, ValidationCol).Value = "WU ID Validation"
    wsDup.Cells(1, ValidationCol).Value = "WU ID Validation"
    wsExcluded.Cells(1, ValidationCol).Value = "WU ID Validation"

    NewRow = 2
    DupRow = 2
    ExcludedRow = 2

    '---------- 4 y 5) Filtros + duplicados ----------
    For i = 2 To LastRowCurrent

        WUID = SafeStr(vCur(i, ColWUID))

        If Len(WUID) > 0 Then

            TotalRows = TotalRows + 1

            dc = 0
            If UseStart Or UseEnd Then
                dc = CheckDate(vCur(i, ColCreated), UseStart, StartDate, UseEnd, EndDate)
            End If

            If dc = 1 Then

                OutDateRows = OutDateRows + 1          ' fuera del rango de fechas

            ElseIf dc = 2 Then

                NoDateRows = NoDateRows + 1            ' Created Date vacio o invalido

            ElseIf (Not SelectAll) And (Not SelTypes.Exists(TypeKey(vCur(i, ColQType)))) Then

                OutTypeRows = OutTypeRows + 1          ' otro Questionnaire Type

            Else

                QuestRows = QuestRows + 1
                Country = NormalizeText(SafeStr(vCur(i, ColCountry)))

                If Not DictLATAM.Exists(Country) Then

                    ExcludedRows = ExcludedRows + 1
                    CopyRowTo wsCurrent, i, wsExcluded, ExcludedRow, nCols
                    wsExcluded.Cells(ExcludedRow, ValidationCol).Value = "Pais no LATAM"
                    ExcludedRow = ExcludedRow + 1

                Else

                    LatamRows = LatamRows + 1

                    If DictOld.Exists(WUID) Then

                        DupPrevRows = DupPrevRows + 1
                        CopyRowTo wsCurrent, i, wsDup, DupRow, nCols
                        wsDup.Cells(DupRow, ValidationCol).Value = "Duplicado: WU ID ya existe en archivo semanal"
                        DupRow = DupRow + 1

                    ElseIf DictSeen.Exists(WUID) Then

                        DupSelfRows = DupSelfRows + 1
                        CopyRowTo wsCurrent, i, wsDup, DupRow, nCols
                        wsDup.Cells(DupRow, ValidationCol).Value = "Duplicado: WU ID repetido en reporte actual"
                        DupRow = DupRow + 1

                    Else

                        NewRows = NewRows + 1
                        DictSeen(WUID) = True
                        CopyRowTo wsCurrent, i, wsNew, NewRow, nCols
                        wsNew.Cells(NewRow, ValidationCol).Value = "Nuevo"
                        NewRow = NewRow + 1

                    End If

                End If

            End If

        End If

    Next i

    '---------- Ordenar (GDD Status Date, o Created Date si no existe) ----------
    Dim NewSortCol As Long
    Dim DupSortCol As Long

    NewSortCol = FindColumn(wsNew, "GDD Status Date")
    If NewSortCol = 0 Then NewSortCol = FindColumn(wsNew, "Created Date")
    DupSortCol = FindColumn(wsDup, "GDD Status Date")
    If DupSortCol = 0 Then DupSortCol = FindColumn(wsDup, "Created Date")

    If NewSortCol > 0 Then SortSheet wsNew, NewSortCol
    If DupSortCol > 0 Then SortSheet wsDup, DupSortCol

    '---------- Tabla para email ----------
    BuildEmailTable wsNew, wsEmail, ColWUID, ColFormName

    '---------- Resumen y formato ----------
    CreateSummary wsSummary, wbCurrent.Name, wbPrevious.Name, DateLabel, TypeLabel, _
                  TotalRows, OutDateRows, NoDateRows, OutTypeRows, QuestRows, _
                  ExcludedRows, LatamRows, DupPrevRows, DupSelfRows, NewRows

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

    wsSummary.Activate

    MsgBox "FTC Report Completed Successfully" & vbCrLf & vbCrLf & _
           "Created Date: " & DateLabel & vbCrLf & _
           "Questionnaire Type: " & TypeLabel & vbCrLf & vbCrLf & _
           "Nuevos: " & NewRows & vbCrLf & _
           "Duplicados: " & (DupPrevRows + DupSelfRows) & vbCrLf & _
           "Excluidos (no LATAM): " & ExcludedRows, vbInformation

    '---------- Borrador de correo (opcional) ----------
    If NewRows > 0 Then
        If MsgBox("Desea crear un borrador de correo en Outlook con la tabla?", _
                  vbYesNo + vbQuestion, "Email") = vbYes Then
            CreateOutlookDraft wsEmail, DateLabel
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

'==============================================================
' VENTANAS EMERGENTES
'==============================================================

' Pide fecha inicial y final (dd/mm/aaaa). Vacio = sin limite. Cancelar = False
Private Function AskDateRange(StartDate As Date, EndDate As Date, _
                              UseStart As Boolean, UseEnd As Boolean) As Boolean

    Dim s As String
    Dim d As Date
    Dim tmp As Date

    Do
        s = VBA.InputBox("FILTRO POR CREATED DATE" & vbCrLf & vbCrLf & _
                         "Fecha INICIAL (dd/mm/aaaa)" & vbCrLf & _
                         "Deje vacio para no limitar la fecha inicial.", _
                         "Created Date - Desde", Format(Date - 7, "dd/mm/yyyy"))
        If StrPtr(s) = 0 Then Exit Function          ' cancelado
        s = Trim(s)
        If s = "" Then
            UseStart = False
            Exit Do
        End If
        If ParseDateText(s, False, d) Then
            StartDate = d
            UseStart = True
            Exit Do
        End If
        MsgBox "Formato de fecha no valido. Use dd/mm/aaaa.", vbExclamation
    Loop

    Do
        s = VBA.InputBox("FILTRO POR CREATED DATE" & vbCrLf & vbCrLf & _
                         "Fecha FINAL (dd/mm/aaaa)" & vbCrLf & _
                         "Deje vacio para no limitar la fecha final.", _
                         "Created Date - Hasta", Format(Date, "dd/mm/yyyy"))
        If StrPtr(s) = 0 Then Exit Function
        s = Trim(s)
        If s = "" Then
            UseEnd = False
            Exit Do
        End If
        If ParseDateText(s, False, d) Then
            EndDate = d
            UseEnd = True
            Exit Do
        End If
        MsgBox "Formato de fecha no valido. Use dd/mm/aaaa.", vbExclamation
    Loop

    If UseStart And UseEnd Then
        If StartDate > EndDate Then
            tmp = StartDate
            StartDate = EndDate
            EndDate = tmp
        End If
    End If

    AskDateRange = True

End Function

' Lista numerada de Questionnaire Types. Cancelar = False
Private Function AskQuestionnaireTypes(arr As Variant, SelTypes As Object, _
                                       SelectAll As Boolean) As Boolean

    Dim n As Long
    Dim i As Long
    Dim num As Long
    Dim prompt As String
    Dim s As String
    Dim tk As Variant
    Dim ok As Boolean

    n = UBound(arr) - LBound(arr) + 1

    prompt = "FILTRO POR QUESTIONNAIRE TYPE" & vbCrLf & vbCrLf
    For i = 0 To n - 1
        prompt = prompt & (i + 1) & " - " & arr(i) & vbCrLf
    Next i
    prompt = prompt & vbCrLf & "Escriba los numeros separados por coma (ej: 1,3)." & vbCrLf & _
             "Escriba 0 o deje vacio para incluir TODOS."

    ' Si la lista es muy larga para la ventana, pregunta uno por uno
    If Len(prompt) > 950 Then
        AskQuestionnaireTypes = AskTypesOneByOne(arr, SelTypes, SelectAll)
        Exit Function
    End If

    Do
        s = VBA.InputBox(prompt, "Questionnaire Type", "0")
        If StrPtr(s) = 0 Then Exit Function          ' cancelado

        s = Replace(Replace(Trim(s), ";", ","), " ", ",")

        If s = "" Or s = "0" Then
            SelectAll = True
            AskQuestionnaireTypes = True
            Exit Function
        End If

        SelTypes.RemoveAll
        ok = True

        For Each tk In Split(s, ",")
            If Len(tk) > 0 Then
                If IsNumeric(tk) Then
                    num = CLng(tk)
                    If num >= 1 And num <= n Then
                        SelTypes(CStr(arr(num - 1))) = True
                    Else
                        ok = False
                    End If
                Else
                    ok = False
                End If
            End If
        Next tk

        If ok And SelTypes.Count > 0 Then
            SelectAll = False
            AskQuestionnaireTypes = True
            Exit Function
        End If

        MsgBox "Seleccion no valida. Use numeros de la lista separados por coma.", vbExclamation
    Loop

End Function

Private Function AskTypesOneByOne(arr As Variant, SelTypes As Object, _
                                  SelectAll As Boolean) As Boolean

    Dim i As Long
    Dim r As VbMsgBoxResult

    SelTypes.RemoveAll

    For i = LBound(arr) To UBound(arr)
        r = MsgBox("Incluir Questionnaire Type:" & vbCrLf & vbCrLf & arr(i), _
                   vbYesNoCancel + vbQuestion, "Questionnaire Type")
        If r = vbCancel Then Exit Function
        If r = vbYes Then SelTypes(CStr(arr(i))) = True
    Next i

    If SelTypes.Count = 0 Then
        MsgBox "No selecciono ningun Questionnaire Type.", vbExclamation
        Exit Function
    End If

    SelectAll = False
    AskTypesOneByOne = True

End Function

'==============================================================
' FECHAS
'==============================================================

' 0 = dentro del rango, 1 = fuera del rango, 2 = fecha vacia/invalida
Private Function CheckDate(v As Variant, UseStart As Boolean, StartDate As Date, _
                           UseEnd As Boolean, EndDate As Date) As Long

    Dim d As Date

    If Not GetCellDate(v, d) Then
        CheckDate = 2
        Exit Function
    End If

    If UseStart Then
        If d < StartDate Then CheckDate = 1: Exit Function
    End If

    If UseEnd Then
        If d > EndDate Then CheckDate = 1: Exit Function
    End If

    CheckDate = 0

End Function

Private Function GetCellDate(v As Variant, ByRef d As Date) As Boolean

    If IsError(v) Then Exit Function
    If IsEmpty(v) Then Exit Function

    Select Case VarType(v)
        Case vbDate
            d = Int(CDbl(v))
            GetCellDate = True
        Case vbDouble, vbLong, vbInteger, vbSingle, vbCurrency
            d = Int(CDbl(v))
            GetCellDate = True
        Case vbString
            GetCellDate = ParseDateText(CStr(v), TEXT_DATE_IS_MDY, d)
    End Select

End Function

' Interpreta dd/mm/aaaa (o mm/dd/aaaa si mdy = True) y aaaa-mm-dd. Ignora la hora.
Private Function ParseDateText(ByVal s As String, mdy As Boolean, ByRef d As Date) As Boolean

    Dim t As String
    Dim p As Variant
    Dim pos As Long
    Dim y As Long, m As Long, dd As Long

    On Error GoTo Fail

    t = Trim(s)
    pos = InStr(t, " ")
    If pos > 0 Then t = Left(t, pos - 1)
    pos = InStr(t, "T")
    If pos > 0 Then t = Left(t, pos - 1)

    t = Replace(t, "-", "/")
    t = Replace(t, ".", "/")
    p = Split(t, "/")

    If UBound(p) <> 2 Then Exit Function

    If Len(p(0)) = 4 Then
        y = CLng(p(0)): m = CLng(p(1)): dd = CLng(p(2))
    ElseIf mdy Then
        m = CLng(p(0)): dd = CLng(p(1)): y = CLng(p(2))
    Else
        dd = CLng(p(0)): m = CLng(p(1)): y = CLng(p(2))
    End If

    If y < 100 Then y = y + 2000
    If m < 1 Or m > 12 Or dd < 1 Or dd > 31 Then Exit Function

    d = DateSerial(y, m, dd)
    If Month(d) <> m Then Exit Function

    ParseDateText = True
    Exit Function

Fail:
    ParseDateText = False

End Function

Private Function BuildDateLabel(StartDate As Date, EndDate As Date, _
                                UseStart As Boolean, UseEnd As Boolean) As String

    If UseStart And UseEnd Then
        BuildDateLabel = Format(StartDate, "dd/mm/yyyy") & " - " & Format(EndDate, "dd/mm/yyyy")
    ElseIf UseStart Then
        BuildDateLabel = "Desde " & Format(StartDate, "dd/mm/yyyy")
    ElseIf UseEnd Then
        BuildDateLabel = "Hasta " & Format(EndDate, "dd/mm/yyyy")
    Else
        BuildDateLabel = "Sin filtro de fecha"
    End If

End Function

'==============================================================
' UTILIDADES
'==============================================================

Private Function SafeStr(v As Variant) As String
    If IsError(v) Then
        SafeStr = ""
    ElseIf IsNull(v) Then
        SafeStr = ""
    Else
        SafeStr = Trim(CStr(v))
    End If
End Function

Private Function TypeKey(v As Variant) As String
    Dim s As String
    s = SafeStr(v)
    If s = "" Then s = "(vacio)"
    TypeKey = s
End Function

Private Sub SortStrings(arr As Variant)

    Dim i As Long
    Dim j As Long
    Dim tmp As String

    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If StrComp(CStr(arr(i)), CStr(arr(j)), vbTextCompare) > 0 Then
                tmp = CStr(arr(i))
                arr(i) = CStr(arr(j))
                arr(j) = tmp
            End If
        Next j
    Next i

End Sub

Private Function NormalizeText(ByVal s As String) As String

    s = UCase(Trim(s))
    s = Replace(s, ChrW(193), "A")
    s = Replace(s, ChrW(201), "E")
    s = Replace(s, ChrW(205), "I")
    s = Replace(s, ChrW(211), "O")
    s = Replace(s, ChrW(218), "U")
    NormalizeText = s

End Function

Private Sub CopyRowTo(wsSrc As Worksheet, SrcRow As Long, _
                      wsDst As Worksheet, DstRow As Long, nCols As Long)

    wsDst.Range(wsDst.Cells(DstRow, 1), wsDst.Cells(DstRow, nCols)).Value = _
        wsSrc.Range(wsSrc.Cells(SrcRow, 1), wsSrc.Cells(SrcRow, nCols)).Value

End Sub

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

'==============================================================
' TABLA PARA EMAIL
'==============================================================

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

Private Sub CreateOutlookDraft(wsEmail As Worksheet, DateLabel As String)

    Dim OutApp As Object
    Dim OutMail As Object
    Dim LastRow As Long
    Dim r As Long
    Dim html As String

    On Error GoTo OutlookErr

    LastRow = wsEmail.Cells(wsEmail.Rows.Count, 1).End(xlUp).Row

    html = "<html><body style='font-family:Calibri,Arial;font-size:11pt;'>"
    html = html & "<p>Hola,</p>"
    html = html & "<p>Listado de nuevos registros FTC (LATAM) - Created Date: " & HtmlEncode(DateLabel) & "</p>"
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

Private Function HtmlEncode(ByVal s As String) As String
    s = Replace(s, "&", "&amp;")
    s = Replace(s, "<", "&lt;")
    s = Replace(s, ">", "&gt;")
    HtmlEncode = s
End Function

'==============================================================
' RESUMEN
'==============================================================

Private Sub CreateSummary(ws As Worksheet, _
                          CurrentFile As String, _
                          PreviousFile As String, _
                          DateLabel As String, _
                          TypeLabel As String, _
                          TotalRows As Long, _
                          OutDateRows As Long, _
                          NoDateRows As Long, _
                          OutTypeRows As Long, _
                          QuestRows As Long, _
                          ExcludedRows As Long, _
                          LatamRows As Long, _
                          DupPrevRows As Long, _
                          DupSelfRows As Long, _
                          NewRows As Long)

    ws.Range("A1").Value = "FTC REPORT SUMMARY"
    ws.Range("A1").Font.Bold = True

    ws.Range("A3").Value = "Run Date"
    ws.Range("B3").Value = Now
    ws.Range("B3").NumberFormat = "dd/mm/yyyy hh:mm"
    ws.Range("A4").Value = "Current File (original)"
    ws.Range("B4").Value = CurrentFile
    ws.Range("A5").Value = "Weekly File (previous)"
    ws.Range("B5").Value = PreviousFile
    ws.Range("A6").Value = "Created Date filter"
    ws.Range("B6").Value = DateLabel
    ws.Range("A7").Value = "Questionnaire Type filter"
    ws.Range("B7").Value = TypeLabel

    ws.Range("A9").Value = "Total records (with WU ID)"
    ws.Range("B9").Value = TotalRows
    ws.Range("A10").Value = "Out of date range"
    ws.Range("B10").Value = OutDateRows
    ws.Range("A11").Value = "Missing / invalid Created Date"
    ws.Range("B11").Value = NoDateRows
    ws.Range("A12").Value = "Other Questionnaire Type"
    ws.Range("B12").Value = OutTypeRows
    ws.Range("A13").Value = "Records after date + type filters"
    ws.Range("B13").Value = QuestRows
    ws.Range("A14").Value = "Excluded (non-LATAM)"
    ws.Range("B14").Value = ExcludedRows
    ws.Range("A15").Value = "LATAM records"
    ws.Range("B15").Value = LatamRows
    ws.Range("A16").Value = "Duplicates - already in weekly file"
    ws.Range("B16").Value = DupPrevRows
    ws.Range("A17").Value = "Duplicates - repeated in current report"
    ws.Range("B17").Value = DupSelfRows
    ws.Range("A18").Value = "New records (unique WU ID)"
    ws.Range("B18").Value = NewRows

    ws.Range("A3:A18").Font.Bold = True
    ws.Range("B3:B18").HorizontalAlignment = xlLeft

End Sub

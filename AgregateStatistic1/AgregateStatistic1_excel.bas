Attribute VB_Name = "AgregateStatistic1"

' Module AgregateStatistic1 (Excel 2011)
' Для работы необходимо чтобы были заполнены все поля, по которым выполняются условия в сводных таблицах

Dim currentKeyRow
Dim lastRezultRow
Dim types
Dim associatDYColumns
Dim associatDYValues

' Функция настройки пользовательских параметров
Function initUsersSettings()
    
    ' Хранит номера колонок в которых требуется провести автозамену значений на листе участков
    associatDYColumns = Array(8, 9)
    
    ' Массив с автозаменяемыми значениями для диаметров участков
    ' Первыми двумя числами указывается диапазон при попадании в который исходное значение заменяется на третье число.
    associatDYValues = Array(Array(0.027, 0.033, 0.03), _
        Array(0.037, 0.043, 0.04), _
        Array(0.047, 0.053, 0.05), _
        Array(0.063, 0.067, 0.065), _
        Array(0.078, 0.083, 0.08), _
        Array(0.095, 0.117, 0.1), _
        Array(0.118, 0.129, 0.125), _
        Array(0.146, 0.155, 0.15) _
    )
    
    
End Function

Sub generateReport()

    initUsersSettings

    replaceValues

    lastRezultRow = 1
    types = Array("Участки", "Узел", "Обобщенный_потребитель")
    
    tName = types(0)
    Worksheets(tName).Activate
    prepareUchastPivot (tName)
    
    tName = types(1)
    Worksheets(tName).Activate
    prepareUzelPivot (tName)
    
    tName = types(2)
    Worksheets(tName).Activate
    preparePotrebPivot (tName)
    
    getKeys
    
    agregateRezultTable
    
End Sub

Function agregateRezultTable()
    
    

    rezultSheet = "rezult"
    removeSheet (rezultSheet)
    
    Set sh = ActiveWorkbook.Sheets.Add
    sh.Name = rezultSheet
    Columns("A:A").Select
    Selection.ColumnWidth = 50
    Columns("B:B").Select
    Selection.ColumnWidth = 13
    Range("A1").Value = "Диаметр, мм" & vbCrLf & "ЦТП, ИТП, тепловая камера, шт."
    Range("B1").Value = "Длина, м" & vbCrLf & "Кол-во, шт."
    Rows("1:1").Select
    Selection.Font.Bold = True
    Selection.EntireRow.AutoFit
    
    lastRezultRow = lastRezultRow + 1
    
    arraySort = Array("Обобщенный_потребитель", "Узел", "Участки")
    meroprArr = Array("Строительство", "Строительство байпаса", "Реконструкция", "Демонтаж", "Демонтаж байпаса")
    vidSeti = Array("Распределительный", "Магистральный")
    
    Worksheets("keys").Activate
    Range("A1").Select
    Set etapKeys = Range(Selection, Selection.End(xlDown))
    
    For n = 1 To etapKeys.Rows.Count
        etap = etapKeys.Cells(n, 1).Value
                
        Worksheets(rezultSheet).Activate
        Range("A" & lastRezultRow).Select
        Selection.Value = extructEtap(etap)
        Selection.Font.Bold = True
        Selection.Font.Italic = True
        Range("A" & lastRezultRow & ":B" & lastRezultRow).Select
        With Selection
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlBottom
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = False
        End With
        Selection.Merge

        lastRezultRow = lastRezultRow + 1
        
        For h = 0 To UBound(meroprArr)
            
            Worksheets(rezultSheet).Activate
            Range("A" & lastRezultRow).Select
            Selection.Font.Italic = True
            Selection.Font.Bold = True
            Selection.HorizontalAlignment = xlCenter
            Selection.Value = meroprArr(h)
            Range("A" & lastRezultRow & ":B" & lastRezultRow).Select
           
            With Selection
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlBottom
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = False
                .ReadingOrder = xlContext
                .MergeCells = False
            End With
            Selection.Merge
            lastRezultRow = lastRezultRow + 1
            
            For i = 0 To UBound(arraySort)
                Do
                    pivotT = "Сводная " & arraySort(i)
                    t = arraySort(i)
                    Worksheets(t).Activate
    
                    Dim pt As PivotTable
                    Set pt = ActiveSheet.PivotTables(pivotT)
                                                                
                    Dim pis As PivotItems
                    Set pis = pt.PivotFields("Мероприятие").PivotItems()
                    Dim finded As Boolean
                    finded = False
                    For d = 1 To pt.PivotFields("Мероприятие").PivotItems().Count
                        If pis(d) = meroprArr(h) Then finded = True
                    Next
                    If Not finded Then
                        Exit Do
                    End If
                    
                    Set pis = pt.PivotFields("Этап").PivotItems()
                    finded = False
                    For d = 1 To pt.PivotFields("Этап").PivotItems().Count
                        If pis(d) = etap Then finded = True
                    Next
                    If Not finded Then
                        Exit Do
                    End If
                                                                
                    Set pis = pt.PivotFields("Мероприятие").PivotItems()
                    pt.PivotFields("Этап").CurrentPage = etap
                    pt.PivotFields("Мероприятие").CurrentPage = meroprArr(h)
                    
                    If arraySort(i) <> "Участки" Then
                                
                        Set r = selectData(pt, arraySort(i))
                        If r Is Nothing Then
                            Exit Do
                        End If
                        r.Copy
                                
                        Worksheets(rezultSheet).Activate
                        Range("B" & lastRezultRow).Select
                        ActiveSheet.Paste
                   
                        Range("A" & lastRezultRow).Select
                        Selection.Value = extructType(arraySort(i))
                     
                        lastRezultRow = lastRezultRow + 1
                        Application.CutCopyMode = False
        
                    End If
                    
                    If arraySort(i) = "Участки" Then
                                        
                        finded = False
                        headrow = -1
                        
                        For u = 0 To UBound(vidSeti)
                            Do
                                If u = 0 Then
                                    Worksheets(rezultSheet).Activate
                                    headrow = lastRezultRow
                                    Range("A" & headrow).Select
                                    Selection.Value = "Тепловые сети всего, в т.ч."
                            
                                    lastRezultRow = lastRezultRow + 1
                                End If
                                
                                Worksheets(t).Activate
                                pt.PivotFields("Вид сети").CurrentPage = vidSeti(u)
                                
                                Set r = selectData(pt, arraySort(i))
                                
                                If r Is Nothing Then
                                    
                                    If u = UBound(vidSeti) And Not finded Then
                                        Worksheets(rezultSheet).Activate
                                        Rows(headrow).Select
                                  
                                        Rows(headrow).Delete
                                        lastRezultRow = lastRezultRow - 1
                                    
                                        Worksheets(t).Activate
                                    End If
                                    Exit Do
                                End If
                                finded = True
                                r.Copy
                                
                                insertRowInd = lastRezultRow
                                        
                                Worksheets(rezultSheet).Activate
                                Range("A" & lastRezultRow + 1).Select
                                                                                            
                                ActiveSheet.Paste
                                Selection.HorizontalAlignment = xlRight
                       
                                lastRezultRow = Selection.Rows.Count + Selection.Row
                                                       
                                Range("A" & insertRowInd).Select
                                Selection.Font.Italic = True
                                Selection.HorizontalAlignment = xlRight
                                Selection.Value = extructvidSeti(vidSeti(u))
                                                      
                                Application.CutCopyMode = False
                            
                            Loop While False
                        Next
                    End If
                Loop While False
            Next
        
        Next

    Next
    Worksheets(rezultSheet).Activate
    Range("A1:B1").Select
    Range(Selection, Selection.End(xlDown)).Select
    fillBorders Selection, True, True
    
    Range("A1").Select
    Set ran = Range(Selection, Selection.End(xlDown))
    For i = 3 To ran.Rows.Count
        If (Cells(i, 1).Font.Bold And InStr(Cells(i, 1).Value, "Подэт") = 0 And Cells(i + 1, 1).Font.Bold) Or (Cells(i, 1).Font.Bold And IsEmpty(Cells(i + 1, 1).Value)) Then
            Rows(i).Delete
            i = i - 1
        End If
    Next
    
    Range("A1").Select
    
End Function

Function replaceValues()

    Worksheets("Участки").Activate
    lastRow = Range("A1").End(xlDown).Row + 1
    
    For i = 0 To UBound(associatDYColumns)
        For r = 2 To lastRow
            Do
                Set cel = Cells(r, associatDYColumns(i))
                v = cel.Value
                For k = 0 To UBound(associatDYValues)
                    If associatDYValues(k)(0) < v And v < associatDYValues(k)(1) Then
                        cel.Value = associatDYValues(k)(2)
                        Exit Do
                    End If
                Next
            Loop While (False)
        Next
    Next
    

End Function


Function extructvidSeti(ByVal vidSeti As String) As String

    If vidSeti = "Магистральный" Then
        extructvidSeti = "- магистральные сети итого, в т.ч.:"
    End If
    
    If vidSeti = "Распределительный" Then
        extructvidSeti = "- распределительные сети итого, в т.ч.:"
    End If
    
End Function


Function extructType(ByVal typ As String) As String

    If typ = "Узел" Then
        extructType = "Тепловая камера"
    End If
    
    If typ = "Обобщенный_потребитель" Then
        extructType = "ИТП и ЦТП"
    End If
    
    If typ = "Участки" Then
    
    End If

End Function

Function extructEtap(ByVal etap As String) As String

    If InStr(etap, "подэтап") > 0 Then
       
        arr = Split(etap, " ")
        For i = 0 To UBound(arr)
            If arr(i) = "подэтап" Then
                extructEtap = "Подэтап " & arr(i - 1) & "." & arr(i + 1)
            End If
        Next
    Else
        extructEtap = "Подэтап " & Right(etap, 1)
    End If

End Function

Function selectData(pTable As PivotTable, ByVal typ As String) As Range
      
    Dim r As Range
    Set r = pTable.TableRange1
    
    r.Select
    
    If IsEmpty(r) Or r.Rows.Count < 2 Then
    
        Set selectData = Nothing
        Exit Function
    End If
         
    If typ = "Участки" Then
        Set r = r.Offset(1, 0)
        Set r = r.Resize(r.Rows.Count - 1, r.Columns.Count)
    Else
        Set r = r.Offset(1, 0)
        Set r = r.Resize(r.Rows.Count - 1, r.Columns.Count)
    End If
    
    r.Select
    Set selectData = r

End Function

Function removeSheet(nameSheet As String)

    For i = 1 To ActiveWorkbook.Sheets.Count
        If ActiveWorkbook.Sheets(i).Name = nameSheet Then
            Worksheets(nameSheet).Delete
            Exit For
        End If
    Next

End Function

Function getKeys()

    removeSheet ("keys")
    
    Set sh = ActiveWorkbook.Sheets.Add
    sh.Name = "keys"
    
    For i = 0 To UBound(types)
        t = types(i)
                  
        Application.Goto Reference:=t
        Range(t & "[[#Headers],[Этап]]").Select
        Selection.Offset(1, 0).Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.Copy
    
        Worksheets("keys").Activate
        If i = 0 Then
            Range("A1").Select
        Else
            Range("A1").End(xlDown).Offset(1, 0).Select
        End If
        ActiveSheet.Paste
    Next
    
    Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.RemoveDuplicates Columns:=1, Header:=xlNo

    Range("A1").Select
    
End Function

Function preparePotrebPivot(tName As String)

    PivotTable = preparePivottable(tName)
    
    ActiveSheet.PivotTables(PivotTable).AddDataField ActiveSheet. _
        PivotTables(PivotTable).PivotFields("CTP_ITP_Name"), _
        "Количество по полю CTP_ITP_Name", xlCount
    
    With ActiveSheet.PivotTables(PivotTable).PivotFields("Этап")
        .Orientation = xlPageField
        .Position = 1
    End With

    With ActiveSheet.PivotTables(PivotTable).PivotFields( _
        "Мероприятие")
        .Orientation = xlPageField
        .Position = 1
    End With
 
    
    With ActiveSheet.PivotTables(PivotTable)
        For Each pvtFld In .PivotFields
            pvtFld.Subtotals(1) = False
        Next pvtFld

        .ColumnGrand = False
        .RowGrand = False
    End With
    
End Function


Function prepareUzelPivot(tName As String)

    PivotTable = preparePivottable(tName)
    
    ActiveSheet.PivotTables(PivotTable).AddDataField ActiveSheet. _
        PivotTables(PivotTable).PivotFields("Наименование узла"), _
        "Количество по полю Наименование узла", xlCount
    
    With ActiveSheet.PivotTables(PivotTable).PivotFields("Этап")
        .Orientation = xlPageField
        .Position = 1
    End With

    With ActiveSheet.PivotTables(PivotTable).PivotFields("Мероприятие")
        .Orientation = xlPageField
        .Position = 1
    End With
    
    With ActiveSheet.PivotTables(PivotTable)
        For Each pvtFld In .PivotFields
            pvtFld.Subtotals(1) = False
        Next pvtFld
    
        .ColumnGrand = False
        .RowGrand = False
    End With
      
End Function

Function prepareUchastPivot(tName As String)

    PivotTable = preparePivottable(tName)
    
    ActiveSheet.PivotTables(PivotTable).AddDataField ActiveSheet. _
        PivotTables(PivotTable).PivotFields("Длина участка, м"), _
        "Сумма по полю Длина участка, м", xlSum
    
    With ActiveSheet.PivotTables(PivotTable).PivotFields("Этап")
        .Orientation = xlPageField
        .Position = 1
    End With
    With ActiveSheet.PivotTables(PivotTable).PivotFields("Вид сети")
        .Orientation = xlPageField
        .Position = 1
    End With

    With ActiveSheet.PivotTables(PivotTable).PivotFields("2Ду")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables(PivotTable).PivotFields("Мероприятие")
        .Orientation = xlPageField
        .Position = 1
    End With
    
    With ActiveSheet.PivotTables(PivotTable)
        For Each pvtFld In .PivotFields
            pvtFld.Subtotals(1) = False
        Next pvtFld
        .ColumnGrand = False
        .RowGrand = False
    End With

End Function

Function preparePivottable(tName As String)

    For i = 1 To ActiveSheet.PivotTables.Count
        ActiveSheet.PivotTables(i).RepeatAllLabels xlRepeatLabels
        ActiveSheet.PivotTables(i).PivotSelect "", xlDataAndLabel, True
        Selection.ClearContents
    Next
    
    For i = 1 To ActiveSheet.ListObjects.Count
        ActiveSheet.ListObjects(i).Unlist
    Next
    
    For i = 1 To ActiveSheet.ListObjects.Count
        ActiveSheet.ListObjects(i).Unlist
    Next

    PivotTable = "Сводная " & tName

    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    
    ActiveSheet.ListObjects.Add(xlSrcRange, Selection, , xlYes).Name = tName
    
    Dim number
    number = Selection.Columns(Selection.Columns.Count).Column + 2
    addressPivTab = tName & "!R3C" & number
    
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        tName, Version:=7).CreatePivotTable TableDestination:=addressPivTab, _
        TableName:=PivotTable, DefaultVersion:=7
        
    With ActiveSheet.PivotTables(PivotTable)
        .ColumnGrand = True
        .HasAutoFormat = True
        .DisplayErrorString = False
        .DisplayNullString = True
        .EnableDrilldown = True
        .ErrorString = ""
        .MergeLabels = False
        .NullString = ""
        .PageFieldOrder = 2
        .PageFieldWrapCount = 0
        .PreserveFormatting = True
        .RowGrand = True
        .SaveData = True
        .PrintTitles = False
        .RepeatItemsOnEachPrintedPage = True
        .TotalsAnnotation = False
        .CompactRowIndent = 1
        .InGridDropZones = False
        .DisplayFieldCaptions = True
        .DisplayMemberPropertyTooltips = False
        .DisplayContextTooltips = True
        .ShowDrillIndicators = True
        .PrintDrillIndicators = False
        .AllowMultipleFilters = False
        .SortUsingCustomLists = True
        .FieldListSortAscending = False
        .ShowValuesRow = False
        .CalculatedMembersInFilters = False
        .RowAxisLayout xlCompactRow
    End With

    ' *
    With ActiveSheet.PivotTables(PivotTable).PivotCache
        .RefreshOnFileOpen = False
        .MissingItemsLimit = xlMissingItemsDefault
    End With
    '*

    preparePivottable = PivotTable
End Function

Function fillBorders(r As Range, verticalInnerBorder As Boolean, horizontInnerBorder As Boolean)

    r.Select
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With

    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With

    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With

    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
    If verticalInnerBorder Then
        With Selection.Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
    End If
    
    If horizontInnerBorder Then
        With Selection.Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
    End If
    
End Function


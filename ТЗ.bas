Sub ТЗ_ВставитьТаблицуМетаданных()
'
' ТЗ_ВставитьТаблицуМетаданных Макрос
'
'
    ActiveDocument.Tables.Add Range:=Selection.Range, NumRows:=5, NumColumns:= _
        4, DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:= _
        wdAutoFitFixed
    Set tbl = Selection.Tables(1)
    tbl.Rows(1).Range.Font.Bold = True
    tbl.Rows(1).Shading.Texture = wdTextureNone
    tbl.Rows(1).Shading.ForegroundPatternColor = wdColorAutomatic
    tbl.Rows(1).Shading.BackgroundPatternColor = -603923969
    tbl.Rows(1).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
    tbl.Rows(1).Cells(1).Range.text = "Имя"
    tbl.Rows(1).Cells(2).Range.text = "Синоним"
    tbl.Rows(1).Cells(3).Range.text = "Тип"
    tbl.Rows(1).Cells(4).Range.text = "Описание"
End Sub

Sub ТЗ_ВставитьТаблицуМетаданныхРС()
'
' ТЗ_ВставитьТаблицуМетаданных Макрос
'
'
    ActiveDocument.Tables.Add Range:=Selection.Range, NumRows:=10, NumColumns:= _
        4, DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:= _
        wdAutoFitFixed
    Set tbl = Selection.Tables(1)
    tbl.Rows(1).Range.Font.Bold = True
    tbl.Rows(1).Shading.Texture = wdTextureNone
    tbl.Rows(1).Shading.ForegroundPatternColor = wdColorAutomatic
    tbl.Rows(1).Shading.BackgroundPatternColor = -603923969
    tbl.Rows(1).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
    tbl.Rows(1).Cells(1).Range.text = "Имя"
    tbl.Rows(1).Cells(2).Range.text = "Синоним"
    tbl.Rows(1).Cells(3).Range.text = "Тип"
    tbl.Rows(1).Cells(4).Range.text = "Описание"
    
    Set Row = tbl.Rows(2)
    Row.Cells.Merge
    Row.Cells(1).Range.text = "Измерения"
    Row.Cells(1).Range.Font.Bold = wdToggle
    Row.Cells(1).Range.Font.Italic = wdToggle
    Row.Shading.Texture = wdTextureNone
    Row.Shading.ForegroundPatternColor = wdColorAutomatic
    Row.Shading.BackgroundPatternColor = -603917569
    
    Set Row = tbl.Rows(5)
    Row.Cells.Merge
    Row.Cells(1).Range.text = "Ресурсы"
    Row.Cells(1).Range.Font.Bold = wdToggle
    Row.Cells(1).Range.Font.Italic = wdToggle
    Row.Shading.Texture = wdTextureNone
    Row.Shading.ForegroundPatternColor = wdColorAutomatic
    Row.Shading.BackgroundPatternColor = -603917569
    
    Set Row = tbl.Rows(8)
    Row.Cells.Merge
    Row.Cells(1).Range.text = "Реквизиты"
    Row.Cells(1).Range.Font.Bold = wdToggle
    Row.Cells(1).Range.Font.Italic = wdToggle
    Row.Shading.Texture = wdTextureNone
    Row.Shading.ForegroundPatternColor = wdColorAutomatic
    Row.Shading.BackgroundPatternColor = -603917569
End Sub

Sub ТЗ_ВставитьРазделВопросы()
'
' ТЗ_ВставитьРазделВопросы Макрос
'
'
    Selection.TypeText text:="Вопросы"
    Selection.Style = ActiveDocument.Styles("Заголовок 1")
    Selection.TypeParagraph
    ActiveDocument.Tables.Add Range:=Selection.Range, NumRows:=6, NumColumns:= _
        2, DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:= _
        wdAutoFitFixed
    Set tbl = Selection.Tables(1)
    wdth = tbl.Columns(1).Width + tbl.Columns(2).Width
    tbl.Columns(1).SetWidth ColumnWidth:=35, RulerStyle:=wdAdjustNone
    tbl.Columns(2).SetWidth ColumnWidth:=wdth - tbl.Columns(1).Width, RulerStyle:=wdAdjustNone
    
    tbl.Rows(1).Range.Font.Bold = True
    tbl.Rows(1).Shading.Texture = wdTextureNone
    tbl.Rows(1).Shading.ForegroundPatternColor = wdColorAutomatic
    tbl.Rows(1).Shading.BackgroundPatternColor = -603923969
    tbl.Rows(1).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
    tbl.Rows(1).Cells(1).Range.text = "Код"
    tbl.Rows(1).Cells(2).Range.text = "Вопрос"
    
    tbl.Rows(2).Cells(1).Range.text = "В.1"
    tbl.Rows(2).Cells(1).Range.Font.Bold = True
    
    tbl.Rows(3).Cells(1).Range.text = "В.2"
    tbl.Rows(3).Cells(1).Range.Font.Bold = True
    
    tbl.Rows(4).Cells(1).Range.text = "В.3"
    tbl.Rows(4).Cells(1).Range.Font.Bold = True
    
    tbl.Rows(5).Cells(1).Range.text = "В.4"
    tbl.Rows(5).Cells(1).Range.Font.Bold = True
    
    tbl.Rows(6).Cells(1).Range.text = "В.5"
    tbl.Rows(6).Cells(1).Range.Font.Bold = True
End Sub

Sub ТЗ_ВставитьРазделТребования()
'
' ТЗ_ВставитьРазделТребования Макрос
'
'
    Selection.TypeText text:="Требования"
    Selection.Style = ActiveDocument.Styles("Заголовок 1")
    Selection.TypeParagraph
    ActiveDocument.Tables.Add Range:=Selection.Range, NumRows:=6, NumColumns:= _
        3, DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:= _
        wdAutoFitFixed
    Set tbl = Selection.Tables(1)
    wdth = tbl.Columns(1).Width + tbl.Columns(2).Width + tbl.Columns(3).Width
    tbl.Columns(1).SetWidth ColumnWidth:=35, RulerStyle:=wdAdjustNone
    tbl.Columns(2).SetWidth ColumnWidth:=150, RulerStyle:=wdAdjustNone
    tbl.Columns(3).SetWidth ColumnWidth:=wdth - (tbl.Columns(1).Width + tbl.Columns(2).Width), RulerStyle:=wdAdjustNone
    
    tbl.Rows(1).Range.Font.Bold = True
    tbl.Rows(1).Shading.Texture = wdTextureNone
    tbl.Rows(1).Shading.ForegroundPatternColor = wdColorAutomatic
    tbl.Rows(1).Shading.BackgroundPatternColor = -603923969
    tbl.Rows(1).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
    tbl.Rows(1).Cells(1).Range.text = "Код"
    tbl.Rows(1).Cells(2).Range.text = "Требование"
    tbl.Rows(1).Cells(3).Range.text = "Описание"
    
    tbl.Rows(2).Cells(1).Range.text = "Т.1"
    tbl.Rows(2).Cells(1).Range.Font.Bold = True
    
    tbl.Rows(3).Cells(1).Range.text = "Т.2"
    tbl.Rows(3).Cells(1).Range.Font.Bold = True
    
    tbl.Rows(4).Cells(1).Range.text = "Т.3"
    tbl.Rows(4).Cells(1).Range.Font.Bold = True
    
    tbl.Rows(5).Cells(1).Range.text = "Т.4"
    tbl.Rows(5).Cells(1).Range.Font.Bold = True
    
    tbl.Rows(6).Cells(1).Range.text = "Т.5"
    tbl.Rows(6).Cells(1).Range.Font.Bold = True
End Sub


Sub Таблица_Ячейки_СделатьЗаголовком()
'
' Таблица_Ячейки_СделатьЗаголовком Макрос
'
'
    For Each Cell In Selection.Cells
        Set rng = Cell.Range
        rng.Font.Bold = wdToggle
        rng.ParagraphFormat.Alignment = wdAlignParagraphCenter
        rng.Shading.Texture = wdTextureNone
        rng.Shading.ForegroundPatternColor = wdColorAutomatic
        rng.Shading.BackgroundPatternColor = -603923969
    Next Cell
End Sub

Sub ВыделенныеЯчейки_AllTrim()
    For Each Cell In Selection.Cells
        Set rng = Cell.Range
        ' Убираем последний символ абзаца из диапазона ячейки
        ' rng.MoveEnd wdCharacter, -1
        
        originalText = rng.text
        If Len(originalText) > 0 Then
            ' Удаляем лидирующие и завершающие непечатаемые символы
            ' cellText = Trim(originalText)
            
            ' Удаляем пробелы, табуляции, неразрывные пробелы
            cellText = CStr(originalText)
            ' Удаляем в начале
            'cellText = RegExReplace(CStr(cellText), "^[\s\t ]+", "")
            ' Удаляем в конце
            'cellText = RegExReplace(CStr(cellText), "[\s\t ]+$", "")
            cellText = TrimNonPrintable(cellText)
            
            ' Если текст изменился, обновляем ячейку
            If cellText <> originalText Then
                rng.text = cellText
                changed = True
                cellsProcessed = cellsProcessed + 1
            End If
        End If
    Next Cell
End Sub

' Функция для замены по регулярному выражению (если включена поддержка VBScript Regular Expressions)
Function RegExReplace(text As String, pattern As String, replacement As String) As String
    On Error GoTo RegExError
    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    
    With regEx
        .Global = True
        .Multiline = True
        .IgnoreCase = False
        .pattern = pattern
    End With
    
    RegExReplace = regEx.Replace(text, replacement)
    Exit Function
    
RegExError:
    ' Если RegEx не доступен, используем простую замену
    RegExReplace = text
End Function

Function TrimNonPrintable(ByVal text As String) As String
    Dim result As String
    Dim i As Long
    Dim startPos As Long
    Dim endPos As Long
    
    result = text
    
    ' Удаляем стандартные пробельные символы
    result = Trim(result)
    
    ' Дополнительно удаляем другие непечатаемые символы
    If Len(result) > 0 Then
        ' Находим первую печатаемую позицию
        startPos = 1
        For i = 1 To Len(result)
            If IsPrintable(Mid(result, i, 1)) Then
                startPos = i
                Exit For
            End If
        Next i
        
        ' Находим последнюю печатаемую позицию
        endPos = Len(result)
        For i = Len(result) To 1 Step -1
            If IsPrintable(Mid(result, i, 1)) Then
                endPos = i
                Exit For
            End If
        Next i
        
        ' Извлекаем только печатаемую часть
        If startPos <= endPos Then
            result = Mid(result, startPos, endPos - startPos + 1)
        Else
            result = "" ' Все символы непечатаемые
        End If
    End If
    
    TrimNonPrintable = result
End Function

' Функция для проверки, является ли символ печатаемым
Function IsPrintable(ByVal ch As String) As Boolean
    Dim ascCode As Integer
    
    ' Получаем ASCII код символа (первый символ в строке)
    ascCode = AscW(ch)
    
    ' Проверяем печатаемые символы:
    ' - от 32 до 126: основные печатаемые символы ASCII
    ' - от 160 до 255: расширенная латиница
    ' - кириллица и другие символы
    If ascCode >= 32 And ascCode <= 126 Then
        IsPrintable = True
    ElseIf ascCode >= 160 And ascCode <= 255 Then
        IsPrintable = True
    ElseIf ascCode >= 1024 And ascCode <= 1279 Then ' Кириллица
        IsPrintable = True
    ElseIf ascCode = 9 Or ascCode = 10 Or ascCode = 13 Then ' Табуляция, перевод строки
        IsPrintable = False
    Else
        ' Другие символы считаем непечатаемыми
        IsPrintable = False
    End If
End Function



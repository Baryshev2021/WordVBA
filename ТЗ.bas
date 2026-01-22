Sub ВставитьТаблицуМетаданных()
'
' ВставитьТаблицуМетаданных Макрос
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
    tbl.Rows(1).Cells(1).Range.Text = "Имя"
    tbl.Rows(1).Cells(2).Range.Text = "Синоним"
    tbl.Rows(1).Cells(3).Range.Text = "Тип"
    tbl.Rows(1).Cells(4).Range.Text = "Описание"
End Sub

Sub ВставитьТаблицуМетаданныхРС()
'
' ВставитьТаблицуМетаданных Макрос
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
    tbl.Rows(1).Cells(1).Range.Text = "Имя"
    tbl.Rows(1).Cells(2).Range.Text = "Синоним"
    tbl.Rows(1).Cells(3).Range.Text = "Тип"
    tbl.Rows(1).Cells(4).Range.Text = "Описание"
    
    Set Row = tbl.Rows(2)
    Row.Cells.Merge
    Row.Cells(1).Range.Text = "Измерения"
    Row.Cells(1).Range.Font.Bold = wdToggle
    Row.Cells(1).Range.Font.Italic = wdToggle
    Row.Shading.Texture = wdTextureNone
    Row.Shading.ForegroundPatternColor = wdColorAutomatic
    Row.Shading.BackgroundPatternColor = -603917569
    
    Set Row = tbl.Rows(5)
    Row.Cells.Merge
    Row.Cells(1).Range.Text = "Ресурсы"
    Row.Cells(1).Range.Font.Bold = wdToggle
    Row.Cells(1).Range.Font.Italic = wdToggle
    Row.Shading.Texture = wdTextureNone
    Row.Shading.ForegroundPatternColor = wdColorAutomatic
    Row.Shading.BackgroundPatternColor = -603917569
    
    Set Row = tbl.Rows(8)
    Row.Cells.Merge
    Row.Cells(1).Range.Text = "Реквизиты"
    Row.Cells(1).Range.Font.Bold = wdToggle
    Row.Cells(1).Range.Font.Italic = wdToggle
    Row.Shading.Texture = wdTextureNone
    Row.Shading.ForegroundPatternColor = wdColorAutomatic
    Row.Shading.BackgroundPatternColor = -603917569
End Sub

Sub ВставитьРазделВопросы()
'
' ВставитьРазделВопросы Макрос
'
'
    Selection.TypeText Text:="Вопросы"
    Selection.Style = ActiveDocument.Styles("Заголовок 1")
    Selection.TypeParagraph
    ActiveDocument.Tables.Add Range:=Selection.Range, NumRows:=6, NumColumns:= _
        2, DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:= _
        wdAutoFitFixed
    Set tbl = Selection.Tables(1)
    tbl.Columns(1).SetWidth ColumnWidth:=35, RulerStyle:=wdAdjustNone
    tbl.Columns(2).SetWidth ColumnWidth:=709, RulerStyle:=wdAdjustNone
    
    tbl.Rows(1).Range.Font.Bold = True
    tbl.Rows(1).Shading.Texture = wdTextureNone
    tbl.Rows(1).Shading.ForegroundPatternColor = wdColorAutomatic
    tbl.Rows(1).Shading.BackgroundPatternColor = -603923969
    tbl.Rows(1).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
    tbl.Rows(1).Cells(1).Range.Text = "Код"
    tbl.Rows(1).Cells(2).Range.Text = "Вопрос"
    
    tbl.Rows(2).Cells(1).Range.Text = "В.1"
    tbl.Rows(2).Cells(1).Range.Font.Bold = True
    
    tbl.Rows(3).Cells(1).Range.Text = "В.2"
    tbl.Rows(3).Cells(1).Range.Font.Bold = True
    
    tbl.Rows(4).Cells(1).Range.Text = "В.3"
    tbl.Rows(4).Cells(1).Range.Font.Bold = True
    
    tbl.Rows(5).Cells(1).Range.Text = "В.4"
    tbl.Rows(5).Cells(1).Range.Font.Bold = True
    
    tbl.Rows(6).Cells(1).Range.Text = "В.5"
    tbl.Rows(6).Cells(1).Range.Font.Bold = True
End Sub

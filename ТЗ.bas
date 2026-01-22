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

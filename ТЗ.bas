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

Sub УдалитьПовторныеЗначенияВКолонке()
    Set Rng = Selection
    Dim CurrentVal As Variant
    Dim LastVal As Variant
    
    LastVal = ""
    
    ' Отключаем обновление экрана и вычисления для ускорения
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    For Row = 1 To Rng.Rows.Count Step 1
        CurrentVal = Rng.Cells(Row, 1).Value
        If CurrentVal = LastVal Then
            Rng.Cells(Row, 1).Value = ""
        Else
            LastVal = CurrentVal
        End If
    Next Row
    
    ' Восстанавливаем настройки
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
End Sub

Attribute VB_Name = "KeysFunc"

Public Sub ClearSearchBar()
    With Application
        .EnableEvents = False
        .ScreenUpdating = False

        ThisWorkbook.Sheets("MainSheet").Range("M3:O3").ClearContents
        ThisWorkbook.Sheets("MainSheet").Range("N3").Select
        .EnableEvents = True
        .ScreenUpdating = True
    End With
End Sub

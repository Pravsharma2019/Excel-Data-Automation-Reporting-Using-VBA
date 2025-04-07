Sub Save_Data()
'
' Save_Data Macro
'

'
    Range("E10:E17").Select
    Selection.Copy
    Sheets("Sheet1").Select
    ActiveWindow.SmallScroll Down:=-189
    Range("J3").Select
    Selection.End(xlDown).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlUp).Select
    ActiveCell.Offset(1, 0).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True
    Sheets("Sheet2").Select
    Application.CutCopyMode = False
    Selection.ClearContents
    Range("E10").Select
End Sub

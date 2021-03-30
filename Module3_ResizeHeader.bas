Attribute VB_Name = "Module3"
Sub ResizeHeader()
Attribute ResizeHeader.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ResizeHeader Macro
'

'
    Columns("J:K").Select
    Selection.ColumnWidth = 13.2
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
End Sub

Sub forms_fill()
    ' save the current active sheet
    Dim cursheet As Worksheet
    Set cursheet = ActiveSheet
    
    ' open student list downloaded from LUNA/Blackboard
    Dim list_file As Variant
    list_file = Application.GetOpenFilename("Student list in TSV format(*.xls),*.xls")
    Dim list_book As Workbook
    Set list_book = Workbooks.Open(list_file)
    Dim list_sheet As Worksheet
    Set list_sheet = list_book.Worksheets(1)
    
    ' record all student codes in dictionary
    Dim db As Object
    Set db = CreateObject("Scripting.Dictionary")
    With list_sheet
        .Range("H3").Select
        Dim key As String
        Dim val As String
        ' going down until empty cell is found
        Do Until IsEmpty(ActiveCell.Value)
            key = ActiveCell.Value
            ' I-th column stores the student code
            val = ActiveCell.Offset(0, 1).Value
            db.Add key, val
            ' move the active cell downward
            ActiveCell.Offset(1, 0).Select
        Loop
    End With
    
    ' revert to the Forms response sheet
    cursheet.Activate
    ' fill the C column with corresponding student codes
    Range("D2").Select
    Do Until IsEmpty(ActiveCell.Value)
        ' extract login ID from Microsoft 365 ID
        key = Left(ActiveCell.Value, 8)
        ' fill the corresponding student ID in the C-th column
        ActiveCell.Offset(0, -1).Value = "'" & db(key)
        ' move the active cell downward
        ActiveCell.Offset(1, 0).Select
    Loop
End Sub

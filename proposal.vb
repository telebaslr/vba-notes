'code for the module "OPCOChooser

Private Sub CancelBtn_Click()
Unload Me
End
End Sub

Private Sub CreateBtn_Click()

opco = DropDownOPCOs.Value
opcoSheet = "Proposal " & opco

If sheetExists(opcoSheet) Then sheetExistsDialog.Show

If Sheets(planningSheet).AutoFilterMode Then
Sheets(planningSheet).Range("B2:R2").AutoFilter
Sheets(planningSheet).Range("B2:R2").AutoFilter
Else: Sheets(planningSheet).Range("B2:R2").AutoFilter
End If 'reset all applied filters

Dim ws As Worksheet
Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
ws.Name = opcoSheet
Cells(1, 1) = "Landscapes"
Cells(1, 2) = "D"
Cells(1, 3) = "Q"
Cells(1, 4) = "A"
Cells(1, 5) = "P"

Sheets(planningSheet).Range("D2:D999").AutoFilter field:=3, Criteria1:=opco 'field 3: third column where can a filter get applied
Sheets(planningSheet).Columns(2).Copy Destination:=Sheets("Proposal").Columns(7)
Sheets("Proposal").Range("G3:G999").AdvancedFilter Action:=xlFilterCopy, CopyToRange:=Sheets(opcoSheet).Range("A2"), unique:=True
Sheets("Proposal").Columns("G").delete
Sheets(opcoSheet).Rows(2).EntireRow.delete

Dim lastrow As Integer
lastrow = Sheets(opcoSheet).Cells(Rows.Count, 1).End(xlUp).Row
Dim landscapes As Variant
Dim element As Variant
landscapes = Range("A2:A" & lastrow).Value
Dim insertRow As Integer
insertRow = 2

For Each element In landscapes
    Dim landscape As String
    landscape = element
    Sheets(planningSheet).Range("B2:B999").AutoFilter field:=1, Criteria1:=landscape
    Dim system As String
    Dim releases As Variant
    releases = Array("D", "Q", "A", "P")
    Dim element2 As Variant
    For Each element2 In releases
        system = Replace(landscape, "x", element2, 1, 1) 'check if case is matched
        Set currentRow = Sheets(planningSheet).Range("C2:C999").Find(system)
        If currentRow Is Nothing Then
            Debug.Print ("No " & element2 & " found in table")
        Else
        Dim term As String
        term = Sheets(planningSheet).Cells(currentRow.Row, 11).Value
        If IsNull(term) Then term = "no date set"
        Dim insertColumn As Integer
        Select Case element2
            Case "D"
                insertColumn = 2
            Case "Q"
                insertColumn = 3
            Case "A"
                insertColumn = 4
            Case "P"
                insertColumn = 5
        End Select
        Debug.Print (insertRow & insertColumn & term)
        Sheets(opcoSheet).Cells(insertRow, insertColumn) = term
        End If
    Next element2
    insertRow = insertRow + 1
    If insertRow > lastrow Then MsgBox "Proposal successfully created", , "Success" End
Next element

End
End Sub

Sub UserForm_Initialize()

Dim opcosArray As Variant
planningSheet = "Q1 PLANNING"

If sheetExists(planningSheet) = False Then
    MsgBox "The planning sheet 'Q1 PLANNING' was not found. Please specify the current sheet in the code (modul: OPCOchooser). Search for the declaration of the variable 'planningSheet'.", , "Planning sheet missing"
    End
End If

Sheets("OPCOLIST").Range("C2:C999").AdvancedFilter Action:=xlFilterCopy, CopyToRange:=Sheets("Proposal").Range("Z1"), unique:=True
opcosArray = Range("Z2", Range("Z2").End(xlDown))
Columns("Z").delete
DropDownOPCOs.List = opcosArray

End Sub

Function sheetExists(sheetToFind As String) As Boolean
    sheetExists = False
    For Each Sheet In Worksheets
        If sheetToFind = Sheet.Name Then
            sheetExists = True
            Exit Function
        End If
    Next Sheet
End Function

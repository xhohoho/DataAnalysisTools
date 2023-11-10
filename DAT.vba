Public sheetname As String
Public filePath As String
Public ws As Worksheet
Public selectedFilter1 As String
Public selectedFilter2 As String

Private Sub ComboBox3_Change()
    ' Declare variables
    Dim selectedColumn As String

    ' Get the selected column from ComboBox3
    selectedColumn = UserForm1.ComboBox3.Value

    ' Check if a column is selected
    If selectedColumn <> "" Then
        ' Set the worksheet based on sheetname (adjust the sheet name if needed)
        Set ws = Worksheets(sheetname)
        ws.Activate

        ' Clear existing conditional formatting from the entire worksheet
        ws.Cells.Interior.ColorIndex = xlNone

        ' Highlight the selected column
        ws.Columns(selectedColumn).Interior.Color = RGB(255, 255, 0) ' You can customize the color

    Else
        ' Inform the user if no column is selected
        MsgBox "Please select a column first!", vbExclamation
    End If
End Sub


Sub CommandButton3_Click()
    Unload Me
End Sub


Private Sub CommandButton1_Click()
    Dim FileDialog As FileDialog
    Set FileDialog = Application.FileDialog(msoFileDialogFilePicker)
    Dim SelectedFile As Variant

    ' Show the file dialog
    If FileDialog.Show = -1 Then
        SelectedFile = FileDialog.SelectedItems(1)
    End If

    Set FileDialog = Nothing
    
    UserForm1.TextBox1.Text = SelectedFile
    
    ' Check if a file path is provided
    If UserForm1.TextBox1.Text = "" Then
        MsgBox "Please enter a file path."
        Exit Sub
    End If

    ' Get the file path from the TextBox
    filePath = UserForm1.TextBox1.Text

    ' Clear the existing worksheet reference
    Set ws = Nothing

    ' Extract the filename from the file path
    sheetname = GetFileNameFromPath()

    ' Parse the CSV data and create a new sheet with the filename as the sheet name
    ParseCSVData

    ' Load all unique filter values into ComboBox1 and ComboBox2
    LoadFilterValues UserForm1.ComboBox1, filePath, 1
    LoadFilterValues UserForm1.ComboBox2, filePath, 8 ' Assuming column H is the 7th column (adjust if needed)

    ' Set "No Filter" as the default selection in ComboBox1 and ComboBox2
    UserForm1.ComboBox1.Value = "No Filter"
    UserForm1.ComboBox2.Value = "No Filter"
End Sub


Sub ParseCSVData()
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetname)
    On Error GoTo 0

    ' If the sheet already exists, clear its contents
    If Not ws Is Nothing Then
        ws.Cells.Clear
    Else
        ' If it doesn't exist, create a new sheet with the given name
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = sheetname
        'ThisWorkbook.Sheets(fileName).Activate
    End If
    
    ThisWorkbook.Sheets(sheetname).Activate

    Dim MyData As String
    Dim i As Integer, j As Integer
    Dim MyArray As Variant
    Dim Delimiter As String
    Delimiter = "," ' Change to your delimiter if it's different

    Open filePath For Input As #1

    ' Counter to keep track of the row
    Dim rowCounter As Long
    rowCounter = 1  ' Start from the first row
    
    Dim columnNumber As Integer
    Dim columnName As String
    
    ' Loop through columns from A to AZ (1 to 52)
    For columnNumber = 1 To 52
        ' Get the column letter
        columnName = Split(Cells(1, columnNumber).Address, "$")(1)
        
        ' Display the column letter in the corresponding cell in column B
        Cells(1, columnNumber).Value = columnName
        UserForm1.ComboBox3.AddItem columnName
    Next columnNumber

    ' Insert an empty row at the beginning
    'ws.Rows(rowCounter).Insert

    Do Until EOF(1)
        Line Input #1, MyData
        MyArray = Split(MyData, Delimiter)

        ' Copy the data to the worksheet
        For j = LBound(MyArray) To UBound(MyArray)
            ws.Cells(rowCounter + 1, j + 1).Value = MyArray(j)
        Next j

        ' Increment the row counter
        rowCounter = rowCounter + 1
    Loop

    Close #1
End Sub


Function GetFileNameFromPath() As String
    ' Extract the file name from a file path
    Dim arr() As String
    arr = Split(filePath, "\")
    GetFileNameFromPath = arr(UBound(arr))
End Function

Sub LoadFilterValues(comboBox As MSForms.comboBox, filePath As String, column As Long)
    comboBox.Clear

    ' Add "No Filter" as the first item
    comboBox.AddItem "No Filter"

    ' Create a dictionary to store unique filter values
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    ' Open the CSV file and read the data to find unique filter values
    Open filePath For Input As #1
    Do Until EOF(1)
        Dim MyData As String
        Line Input #1, MyData
        Dim MyArray As Variant
        Dim Delimiter As String
        Delimiter = "," ' Change to your delimiter if it's different
        MyArray = Split(MyData, Delimiter)

        ' Check if the column index is valid
        If column >= LBound(MyArray) + 1 And column <= UBound(MyArray) + 1 Then
            If Not dict.Exists(MyArray(column - 1)) Then
                dict.Add MyArray(column - 1), Nothing
            End If
        End If
    Loop
    Close #1

    ' Add unique filter values to the ComboBox
    Dim filter As Variant
    For Each filter In dict.Keys
        comboBox.AddItem filter
    Next filter
End Sub

Sub ClearData()
    ' Clear the data in the current sheet (except headers)
    If Not ActiveSheet Is Nothing Then
        Dim lastRow As Long
        lastRow = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row
        If lastRow > 1 Then
            ActiveSheet.Range("A2:A" & lastRow).EntireRow.Delete
        End If
    End If
End Sub

Private Sub ComboBox1_Change()
    ws.Activate
    ' Get the selected filters from both ComboBox1 and ComboBox2
    selectedFilter1 = UserForm1.ComboBox1.Text
    selectedFilter2 = UserForm1.ComboBox2.Text

    ' Clear the existing data
    ClearData

    ' Filter and display data based on the selected filters
    FilterAndDisplayData 1, 8 ' Adjust column numbers as needed
End Sub

Private Sub ComboBox2_Change()
    ' Get the selected filters from both ComboBox1 and ComboBox2
    ws.Activate
    selectedFilter1 = UserForm1.ComboBox1.Text
    selectedFilter2 = UserForm1.ComboBox2.Text

    ' Clear the existing data
    ClearData

    ' Filter and display data based on the selected filters
    FilterAndDisplayData 1, 8 ' Adjust column numbers as needed
End Sub

Sub FilterAndDisplayData(column1 As Long, column2 As Long)
    ' Check if the active sheet is the specified sheetname
    
    If Not ws Is Nothing Then
        Dim MyData As String
        Dim i As Long
        Dim j As Long
        Dim MyArray As Variant
        Dim Delimiter As String
        Delimiter = "," ' Change to your delimiter if it's different
        Dim rowCounter As Long
        rowCounter = 2 ' Start from the second row (assuming headers are in the first row)

        Open filePath For Input As #1
        Do Until EOF(1)
            Line Input #1, MyData
            MyArray = Split(MyData, Delimiter)

            ' Check if the column indexes are valid
            If (column1 >= 1 And column1 <= UBound(MyArray) + 1) And (column2 >= 1 And column2 <= UBound(MyArray) + 1) Then
                If (selectedFilter1 = "No Filter" Or MyArray(column1 - 1) = selectedFilter1) And (selectedFilter2 = "No Filter" Or MyArray(column2 - 1) = selectedFilter2) Then
                    For j = LBound(MyArray) To UBound(MyArray)
                        ws.Cells(rowCounter, j + 1).Value = MyArray(j)
                    Next j
                    rowCounter = rowCounter + 1
                End If
            End If
        Loop
        Close #1
    End If
End Sub

Private Sub CommandButton3_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Unload Me
End Sub

Private Sub CommandButton4_Click()
    ' Declare variables
    Dim selectedColumn As String
    Dim sourceRange As Range
    Dim destinationSheet As Worksheet
    Dim lastRowSource As Long
    Dim lastRowDestination As Long

    ' Get the selected column from ComboBox3
    selectedColumn = UserForm1.ComboBox3.Value

    ' Check if a column is selected
    If selectedColumn <> "" Then
        ' Set the source range based on sheetname (adjust the sheet name if needed)
        With Worksheets(sheetname)
            lastRowSource = .Cells(.Rows.Count, selectedColumn).End(xlUp).Row
            Set sourceRange = .Columns(selectedColumn).Resize(lastRowSource)
        End With

        ' Set the destination sheet (adjust the sheet name if needed)
        Set destinationSheet = Worksheets("DAT")

        ' Clear existing data in the specified destination range (B5:B154) on the DAT sheet
        destinationSheet.Range("B5:B154").ClearContents

        ' Determine the last row with data in the selected column on the DAT sheet
        lastRowDestination = destinationSheet.Cells(destinationSheet.Rows.Count, selectedColumn).End(xlUp).Row

        ' Copy and paste values only from the source range to the destination sheet starting from B5
        sourceRange.Offset(1).Resize(lastRowSource).Copy
        destinationSheet.Range("B5").PasteSpecial xlPasteValues

        ' Clear the clipboard to avoid memory issues and deselect the copied range
        Application.CutCopyMode = False

        ' Clear the selection on the DAT sheet
        destinationSheet.Activate
        destinationSheet.Cells(1, 1).Select

        ' Inform the user about the successful copy
        MsgBox "Data copied to DAT sheet and activated successfully!", vbInformation
    Else
        ' Inform the user if no column is selected
        MsgBox "Please select a column first!", vbExclamation
    End If

    ' Show the UserForm in modeless form
    UserForm1.Show vbModeless
End Sub



Private Sub CommandButton5_Click()
    UserForm1.Hide
End Sub

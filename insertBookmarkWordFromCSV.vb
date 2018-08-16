'visual basic script to ealily insert bookmarks at cursor. 
'loaded from csv file
Private Sub CommandButton1_Click()
	
	Dim Delimiter As String
	Dim TextFile As Integer
	Dim FilePath As String
	Dim FileContent As String
	Dim LineArray() As String
	Dim DataArray() As String
	Dim TempArray() As String
	Dim rw As Long, col As Long
	Dim i As Integer
	For i = 0 To ListBox1.ListCount - 1
		If ListBox1.Selected(i) Then
			ActiveDocument.Bookmarks.Add Name:="Str" & ListBox1.List(i)
		End If
		Next i
		Dim fd As Office.FileDialog
		Set fd = Application.FileDialog(msoFileDialogFilePicker)
		With fd
			.AllowMultiSelect = False
' Set the title of the dialog box.
			.Title = "Please select the file."
' Clear out the current filters, and add our own.
			.Filters.Clear
			.Filters.Add "csv-file", "*.csv"
			.Filters.Add "csv-file", "*.txt"
			.Filters.Add "All Files", "*.*"
' Show the dialog box. If the .Show method returns True, the
' user picked at least one file. If the .Show method returns
' False, the user clicked Cancel.
			If .Show = True Then
				txtFileName = .SelectedItems(1) 'replace txtFileName with your textbox
			End If
		End With
'Inputs
		Delimiter = ","
		FilePath = txtFileName
		rw = 0
'Open the text file in a Read State
		TextFile = FreeFile
		Open FilePath For Input As TextFile
'Store file content inside a variable
		FileContent = Input(LOF(TextFile), TextFile)
'Close Text File
		Close TextFile
'Separate Out lines of data
		LineArray() = Split(FileContent, vbCrLf)
'Read Data into an Array Variable
		For x = LBound(LineArray) To UBound(LineArray)
			If Len(Trim(LineArray(x))) <> 0 Then
'Split up line of text by delimiter
				TempArray = Split(LineArray(x), Delimiter)
'Determine how many columns are needed
				col = UBound(TempArray)
'Re-Adjust Array boundaries
				ReDim Preserve DataArray(col, rw)
'Load line of data into Array variable
				For y = LBound(TempArray) To UBound(TempArray)
					DataArray(y, rw) = TempArray(y)
					Next y
				End If
'Next line
				rw = rw + 1
				Next x
				ListBox1.List = LineArray
			End Sub
			
			Private Sub insertBookmark_Click()
				
				Dim Msg As String
				Dim i As Integer
				
				For i = 0 To ListBox1.ListCount - 1
					If ListBox1.Selected(i) Then
						Dim leftString As String
						leftString = ListBox1.List(i)
						leftString = Left(leftString, InStr(leftString, ","))
						leftString = Left(leftString, Len(leftString) - 1)
						leftString = "str" & leftString
						ActiveDocument.Bookmarks.Add Name:=leftString
					End If
					Next i
					Word.Application.Activate
				End Sub
				
				

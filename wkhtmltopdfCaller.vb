Imports System.Windows.Forms
Imports System.EventArgs

Module htmltopdf
'hiding console internet hack
'Declare Function AllocConsole Lib "kernel32" () As Int32						'use AllocConsole() to show console
Declare Function FreeConsole Lib "kernel32" () As Int32							'use FreeConsole() to hide console
'hiding console internet hack

'Module Level
	Dim WithEvents MainForm As New Form				'Application Form
	Dim Lable1 As New Label
	Dim ComboBox1 As New ComboBox
	Dim HandleEvents As Boolean
	Dim ComboBox1TextChangeEvent As Boolean
	Const HISTORY_LIMIT As Integer = 15
	
	Class History
		Private Shared historyArray(HISTORY_LIMIT - 1) As String
		
		Sub New()
			Dim temp() As String
			temp = Split(readFile(), vbCrLf)
			ReDim Preserve temp(UBound(historyArray))
			Array.Copy(temp, 0, historyArray, 0, temp.Length)
		End Sub
		
		Public Function getString()
			Return Join(historyArray, vbCrLf)
		End Function
		
		Public Function getArray()
			Return Split(Join(historyArray, vbCrLf), vbCrLf)
		End Function
		
		Public Sub newPath(ByVal path As String)
			'Got new path; add to the history or if already exists then renew it's place
			Dim foundAt As Integer
			foundAt = search(path)
			If foundAt <> -1 Then
				If foundAt <> 0 Then
					shiftToTop(foundAt)
				End If
			Else
				addNew(path)
			End If
			updateCombobox1(getArray)
		End Sub
		
		Sub shiftToTop(ByVal index As Integer)
			Dim temp(UBound(historyArray)) As String
			temp(0) = historyArray(index)
			Array.Copy(historyArray, 0, temp, 1, index)
			Array.Copy(historyArray, (index + 1), temp, (index + 1), (UBound(historyArray) - index))
			Array.Copy(temp, 0, historyArray, 0, temp.Length)
		End Sub
		
		Sub addNew(ByVal path As String)
			Dim temp(UBound(historyArray)) As String
			temp(0) = path
			Array.Copy(historyArray, 0, temp, 1, UBound(historyArray))
			Array.Copy(temp, 0, historyArray, 0, temp.Length)
		End Sub
		
		Function search(ByVal path As String)
			For index As Integer = LBound(historyArray) To UBound(historyArray)
				If StrComp(path, historyArray(index), CompareMethod.Text) = 0 Then Return index
			Next
			Return -1
		End Function
	End Class
	
	Dim myHistory As New History
	
	Sub Main()
		'This program calls wkhtmltopdf to convert html file into pdf files
		'This program facilitates path suggestions for wkhtmltopdf by storing paths in a file
		
		FreeConsole()						'Hide Console
		HandleEvents = False				'Disable Event Handlers which use HandleEvents
		PrepareForm (inForm:=MainForm)		'Initialize Form
		HandleEvents = True					'Enable Event Handler which use HandleEvents
		Application.Run(MainForm)			'Initiate Main Form
		
		'PO1. Ask user to enter input HTML file path (directory only) with suggestions given in a listbox [Done in UI]
	End Sub
	
	Sub callwkhtmltopdf()
		Dim wkhtmltopdfPath As String, sourcePath As String, destPath As String
		wkhtmltopdfPath = "C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf"
		sourcePath = """" + Combobox1.Text + """"
		destPath = """" + Combobox1.Text + ".pdf"""
		System.Diagnostics.Process.Start(wkhtmltopdfPath, sourcePath + " " + destPath)
	End Sub
	
	Sub PrepareForm(inForm As Form)
		'Initial Form Look
		inForm.Icon = New System.Drawing.Icon("htmlpdf.ico")
		inForm.Text = "HTML to PDF Convertor"		'Title
		inForm.Width = 500
		inForm.Height = 150
		inForm.StartPosition = FormStartPosition.CenterScreen
		'Add Controls to Form
		
		'Add Label
		Lable1.Text = "Enter Input HTML File Path:"
		Lable1.Location = New System.Drawing.Point(10,10)
		Lable1.AutoSize = True
		inForm.Controls.Add(Lable1)
		
		'Add ComboBox for suggestions
		ComboBox1.Width = 450
		ComboBox1.Height = 20
		ComboBox1.Location = New System.Drawing.Point(10,(Lable1.Top + Lable1.Height + 10))
		updateCombobox1(myHistory.getArray())
		ComboBox1.AutoCompleteSource = AutoCompleteSource.ListItems
		ComboBox1.AutoCompleteMode = AutoCompleteMode.SuggestAppend
		AddHandler ComboBox1.LostFocus, AddressOf ComboBox1_LostFocus
		AddHandler ComboBox1.TextChanged, AddressOf ComboBox1_TextChanged
		inForm.Controls.Add(ComboBox1)
		
		'Browse Button
		Dim btn_Browse As New Button				'No outside access needed
		btn_Browse.Text = "Browse"
		btn_Browse.Location = New System.Drawing.Point(10, (ComboBox1.Top + ComboBox1.Height + 10))
		AddHandler btn_Browse.Click, AddressOf btn_Browse_Click
		inForm.Controls.Add(btn_Browse)
		
		'Convert Button
		Dim btn_Convert As New Button
		btn_Convert.Text = "Convert"
		btn_Convert.Location = New System.Drawing.Point((btn_Browse.Right + 10), btn_Browse.Top)
		AddHandler btn_Convert.Click, AddressOf btn_Convert_Click
		inForm.Controls.Add(btn_Convert)
	End Sub
	
	Sub updateCombobox1(ByRef historyArray() As String)
		ComboBox1.Items.Clear
		For Each line in historyArray
			If line <> "" Then ComboBox1.Items.Add(line)
		Next
	End Sub
	
	Sub btn_Browse_Click(ByVal sender As Object, ByVal e As EventArgs)
		Dim browser As OpenFileDialog
		browser = New OpenFileDialog
		browser.RestoreDirectory = True
		browser.Filter = "HTML Files|*.htm;*.html"
		If browser.ShowDialog() = DialogResult.Ok Then
			ComboBox1.Text = browser.FileName
		End If
	End Sub
	
	Sub btn_Convert_Click(ByVal sender As Object, ByVal e As EventArgs)
		myHistory.newPath(ComboBox1.Text)
		callwkhtmltopdf
	End Sub
	
	Sub ComboBox1_LostFocus(ByVal sender As Object, ByVal e As EventArgs)
		If ComboBox1TextChangeEvent AND ComboBox1.Text <> "" Then
			myHistory.newPath(ComboBox1.Text)
			ComboBox1TextChangeEvent = False
		End If
	End Sub
	
	Sub ComboBox1_TextChanged()
		ComboBox1TextChangeEvent = True
	End Sub
	
	Function readFile()
		Dim FileRead As String
		FileRead = My.Computer.FileSystem.ReadAllText("paths.htp")
		Return FileRead
	End Function
	
	Sub writeToFile(ByVal inString As String)
		My.Computer.FileSystem.WriteAllText("paths.htp", inString + vbCrLf, True)
	End Sub
	
	Sub writeToFileAsIs(ByVal inString As String)
		My.Computer.FileSystem.WriteAllText("paths.htp", inString, True)
	End Sub
	
	Sub writeToFileClean(ByVal inString As String)
		My.Computer.FileSystem.WriteAllText("paths.htp", inString, False)
	End Sub
	
	Sub MainForm_FormClosing(sender as Object, e as FormClosingEventArgs) Handles MainForm.FormClosing
		writeToFileClean(myHistory.getString)
	End Sub
	
End Module
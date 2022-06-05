' Everything works!
' First delete the old databases
' Then import the excel file in a database
' Then create a new database to with 3 columns which are ACIKLAMALAR, CARI_DONEM and ONCEKI_DONEM
' Then fill the newly created database with the values of imported excel file.

Option Explicit
Sub Main
	
	' Get the working directory.
	Dim Path As String
	Path = Client.WorkingDirectory
	
	' Delete the old file.
	DeleteOld Path + "Excel-BOBÝ FRS NAT Dolaylý Konsolide.IMD"
	
	' Delete old database
	DeleteOld Path + "erhan_er_test.imd"
	
	' Locate the input file.
	Dim excelName As String
	excelName = Client.LocateInputFile (Path + "4- BOBÝ FRS Nakit Akýþ Tablosu - Dolaylý Yöntem (Konsolide).xlsx")
	
	' Create a task to import excel file
	Dim task As ImportExcel
	Set task = Client.GetImportTask( "ImportExcel" )
	
	' Configure the task
	task.FileToImport = excelName
	task.SheetToImport = "BOBÝ FRS NAT Dolaylý Konsolide"
	task.OutputFilePrefix = "Excel"
	task.FirstRowIsFieldName = "True"
	task.EmptyNumericFieldAsZero = "False"
	
	task.PerformTask
	
	Set task = Nothing
	
	' Create erhan_er_test.imd
	Dim db As Database
	Set db = CreateDB()
	
	' Fill erhan_er_test.imd with the data of excel file
	FillDB( db )
	
	' Inform the user 
	MsgBox "Success!"
	
	' Clear memory
	Set db = Nothing 	
End Sub

Sub DeleteOld(Filepath As String)
	' Determine if the file exists.
	Dim PathCheck As String
	PathCheck = Dir( Filepath )
	
	' Delete the file if it exists.
	If Len(PathCheck) > 1 Then
		' Close the database before deleting it.
		Client.CloseDatabase(PathCheck)
		MsgBox "Deleting old copy of " + PathCheck
		Client.DeleteDatabase(PathCheck)
	End If
End Sub

Function CreateDB () As Database
	' Create a new table with TableDef
	Dim newTable As Table
	Set newTable = Client.NewTableDef
	
	' Create a new field.
	Dim newField As Field
	Set newField = newTable.NewField
	
	' Configure the field "ACIKLAMALAR" and append it to the table.
	newField.Name = "ACIKLAMALAR"
	newField.Type = WI_EDIT_CHAR
	newField.Length = 200
	newField.Description = "Açýklamalar"
	newTable.AppendField newField
	
	' Configure the field "CARI_DONEM" and append it to the table.
	Set newField = newTable.NewField
	newField.Name = "CARI_DONEM"
	newField.Type = WI_EDIT_NUM
	newField.Decimals = 0
	newField.Description = "Cari Dönem"
	newTable.AppendField newField

	' Configure the field "ONCEKI_DONEM" and append it to the table.
	Set newField = newTable.NewField
	newField.Name = "ONCEKI_DONEM"
	newField.Type = WI_EDIT_NUM
	newField.Decimals = 0
	newField.Description = "Önceki Dönem"
	newTable.AppendField newField
	
	' Create a new database and put the new table in it.
	Dim db As Database
	Client.DeleteDatabase("erhan_er_test.IMD")
	Set db = Client.NewDatabase("erhan_er_test.IMD",  "", newTable)
	
	' Return the new database.
	Set CreateDB = db
	
	'  Delete everything.
	Set db = Nothing
	Set newField = Nothing
	Set newTable = Nothing
End Function

Function FillDB ( db As Database )
	' Open imported database
	Dim import As Database
	Set import = OpenDB("Excel-BOBÝ FRS NAT Dolaylý Konsolide.imd")
	
	' Access the recordset.
	Dim RS As RecordSet
	Set RS = import.RecordSet
	
	' Obtain the record set of db
	Dim dbRS As RecordSet
	Set dbRS = db.RecordSet
	
	' Create a new record
	Dim NRS As RecordSet
	Set NRS = dbRS.NewRecord 
	 
	' Change the table settings to allow writing
	Dim table As TableDef
	Set table = db.TableDef
	table.Protect = False
	
	' Move to the first record.
	RS.ToFirst
	RS.Next
	RS.Next
	
	' Create Columns
	Dim col3 As String
	Dim col4 As String
	Dim col5 As String
	
	' Create count
	Dim count As Integer
	
	' The imported database of excel consists of 12 columns.
	' But we only need col3, col4 and col5.
	' For each record, one of these columns have a value.
	' Therefore, we only need the one which has a value for each record.
	For count = 1 To 69
		RS.Next
		Set col3 = RS.ActiveRecord.GetCharValue("COL3")
		Set col4 = RS.ActiveRecord.GetCharValue("BOBI_FRS_NAKIT_AKýÞ_TABLOSU_DOLAYLý_YÖNT")
		Set col5 = RS.ActiveRecord.GetCharValue("COL5")
		
		If col3 <> "" Then
			NRS.SetCharValue "ACIKLAMALAR", col3
			NRS.SetCharValue "CARI_DONEM", 0
			NRS.SetCharValue "ONCEKI_DONEM", 0
			dbRS.AppendRecord NRS
		ElseIf col4 <> "" Then
			NRS.SetCharValue "ACIKLAMALAR", col4
			NRS.SetCharValue "CARI_DONEM", 0
			NRS.SetCharValue "ONCEKI_DONEM", 0
			dbRS.AppendRecord NRS
		Else
			NRS.SetCharValue "ACIKLAMALAR", col5
			NRS.SetCharValue "CARI_DONEM", 0
			NRS.SetCharValue "ONCEKI_DONEM", 0
			dbRS.AppendRecord NRS
		End If
		
		NRS.ClearRecord
	Next
	
	' Before committing, protect the table
	table.Protect = True
	
	' Commit
	db.CommitDatabase
	
	' Close databases
	import.close
	db.close
	
	' Clear Memory
	Set import = Nothing
	Set RS = Nothing
	Set dbRS = Nothing
	Set NRS = Nothing
End Function

Function OpenDB ( path As String ) As Database
	Dim check As String
	check = Dir(path)
	
	Dim db As Database
	
	Set db = Client.OpenDatabase(path)
	OpenDB = db
	
	Set db = Nothing
End Function


	
On Error Resume Next

' ================================================================================================
' Adding Function Libraries in script
' ================================================================================================
testDir= Environment.Value("TestDir")
arrPath= Split(testDir, "\")
	
'Getting the path for function library
arrPath(UBound(arrPath)-1)= "FunctionLibrary"
arrPath(UBound(arrPath))=""
testCasesExcelSheetPath= ""
For I=0 to UBound(arrPath)-1
	If (I=0) Then
		functionLibraryPath = arrPath(I)
	Else
		functionLibraryPath = functionLibraryPath + "\" + arrPath(I)
	End If
Next
functionLibraryPath= functionLibraryPath & "\"
	
Set fso = CreateObject("Scripting.FileSystemObject")
Set folder = fso.GetFolder(functionLibraryPath)

For Each file In folder.Files
    If LCase(fso.GetExtensionName(file.Name)) = "vbs" Or LCase(fso.GetExtensionName(file.Name)) = "qfl" Then
        LoadFunctionLibrary file.Path
    End If
Next

' ================================================================================================
'  DESCRIPTION 	  	: This function is used to create global variables which stores location path of result folder, test data folder,  function library folder and object repository folder
'  PRESENT IN		: CommonLib.qfl
' ================================================================================================
Initialization()

' ================================================================================================
' Adding the object repository in script
' ================================================================================================
Set repoFolder = fso.GetFolder(objectRepositoryFolderPath)

For Each repo In repoFolder.Files
    If LCase(fso.GetExtensionName(repo.Name)) = "tsr" Then
        RepositoriesCollection.Add repo.Path
    End If
Next

' ================================================================================================
'Generates a new html file for custom report
'PRESENT IN		: GenerateCustomReports.qfl
' ================================================================================================
CreateResultFile()

' ===============================================================]=================================
' Accessing Excel Sheet 
' ================================================================================================
pathForExcelSheet= testDataFolderPath & "TestCasesSheet.xlsx"
'DataTable.AddSheet("TestCases")
'DataTable.ImportSheet pathForExcelSheet, "TestCases", "TestCases"

DataTable.AddSheet("Login")
DataTable.ImportSheet pathForExcelSheet, "Login", "Login"

DataTable.AddSheet("NewOrder")
DataTable.ImportSheet pathForExcelSheet, "NewOrder", "NewOrder"

DataTable.AddSheet("SearchOrder")
DataTable.ImportSheet pathForExcelSheet, "SearchOrder", "SearchOrder"

'PREREQUISITE FOR WAY1
'Creating and Opening a ADODB Integration Object and Integrating Excel with it
Dim conn, rs, query, excelFile
Set conn = CreateObject("ADODB.Connection")
Set rs = CreateObject("ADODB.Recordset")

excelFile =  pathForExcelSheet  'Update this to your file path

' Connection string for Excel 2016 (.xlsx)
conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;" & _
          "Data Source=" & excelFile & ";" & _
          "Extended Properties=""Excel 12.0 Xml;HDR=Yes;IMEX=1"";"

' Querying data from TestCases (you must use $ with sheet name)
query = "SELECT * FROM [TestCases$]"

rs.Open query, conn

' ================================================================================================
' Running the Script From Test Cases Excel Sheet
' PRESENT IN		: TestData\TestCasesSheet.xlsx
' ================================================================================================
'WAY 1 (Through ADODB Integration with Excel)
Do Until rs.EOF
	If UCase(rs.Fields(3).Value) ="Y" Then
		keyword= rs.Fields(2).Value
		tcid= rs.Fields(0).Value
		tc_desc= rs.Fields(1).Value
		Select Case keyword
			Case "Launch"
				launchApp()
			Case "Login"
				ExecuteLoginTestCases tcid, tc_desc
			Case "NewOrder"
				ExecuteNewOrderTestCases tcid
			Case "SearchOrder"
				ExecuteSearchOrderTestCases tcid, tc_desc
			Case "Close"
				closeApplication()
			Case default
				MsgBox "Keyword not found"
		End Select
	End If
	rs.MoveNext
Loop

'WAY 2(Adding Excel Sheets in DataTable)
'tc_rows= DataTable.GetSheet("TestCases").GetRowCount
'
'For i = 1 To tc_rows Step 1
'	DataTable.GetSheet("TestCases").SetCurrentRow(i)
'	If UCase(DataTable.Value("Execute", "TestCases"))="Y" Then
'		keyword= DataTable.Value("Keywords", "TestCases")
'		tcid= DataTable.Value("TC_ID", "TestCases")
'		tc_desc= DataTable.Value("Test_Case_Description", "TestCases")
'		Select Case keyword
'			Case "Launch"
'				launchApp()
'			Case "Login"
'				ExecuteLoginTestCases tcid, tc_desc
'			Case "NewOrder"
'				ExecuteNewOrderTestCases tcid
'			Case "SearchOrder"
'				ExecuteSearchOrderTestCases tcid, tc_desc
'			Case "Close"
'				closeApplication()
'			Case default
'				MsgBox "Keyword not found"
'		End Select
'	End If
'Next

' ================================================================================================
' DESCRIPTION		 : Generates TestCases Summary Report in custom html report 
' PRESENT IN		: GenerateCustomReports.qfl
' ================================================================================================
TestCaseExecutiveSummary ()

' ================================================================================================
' DESCRIPTION		 : Email custom html report 
' PRESENT IN		: CommonLib.qfl
' PREREQUISITE		: Enter email Id in the SendMail() to send an email
' ================================================================================================
'SendMail()

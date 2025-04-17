Dim conn, rs, query, excelFile
Set conn = CreateObject("ADODB.Connection")
Set rs = CreateObject("ADODB.Recordset")

excelFile = "C:\UFT\FlightGUI_Automation\TestData\TestCasesSheet.xlsx"   'Update this to your file path

' Connection string for Excel 2016 (.xlsx)
conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;" & _
          "Data Source=" & excelFile & ";" & _
          "Extended Properties=""Excel 12.0 Xml;HDR=Yes;IMEX=1"";"

' Querying data from TestCases (you must use $ with sheet name)
query = "SELECT * FROM [TestCases$]"

rs.Open query, conn

'  Loop through records and show first column values
Do Until rs.EOF
    MsgBox rs.Fields(0).Value
    rs.MoveNext
Loop

rs.Close
conn.Close

Set rs = Nothing
Set conn = Nothing

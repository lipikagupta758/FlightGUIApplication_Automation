RepositoriesCollection.Add "C:\UFT\FlightGUI_Automation\ObjectRepository\LoginPage.tsr"
RepositoriesCollection.Add "C:\UFT\FlightGUI_Automation\ObjectRepository\FlightApp.tsr"

Set connection= CreateObject("ADODB.Connection")
mysql_conn= "Driver={MySQL ODBC 8.0 Unicode Driver};Server=localhost;Database=FlightGUI;User=root;Password=root;Option=3;"
connection.ConnectionString= mysql_conn
connection.Open

sqlQuery= "SELECT * FROM Login;"
Set recordSet= CreateObject("ADODB.RecordSet")
recordSet.Open sqlQuery, connection, adOpenStatic

Do  Until recordSet.EOF
	username= recordSet.Fields("username")
	password= recordSet.Fields("password")
	login username, password
	recordSet.MoveNext
LOOP
connection.Close


Function login(username, password)
	WpfWindow("OpenText MyFlight Sample").WpfEdit("agentName").Set username
	WpfWindow("OpenText MyFlight Sample").WpfEdit("password").Set password
	WpfWindow("OpenText MyFlight Sample").WpfButton("OK").Click
	
	If WpfWindow("OpenText MyFlight Sample").WpfButton("FIND FLIGHTS").Exist Then
		Reporter.ReportEvent micPass, "Login", "Login Successfully Done"
	Else
		Reporter.ReportEvent micFail, "Login", "Login failed"
	End If

End Function
	


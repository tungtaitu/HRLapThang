<%
Function GetSQLServerConnection()
    Dim conn
    Session.Timeout = 60
    Server.ScriptTimeOut=99999
    Set conn = Server.CreateObject("ADODB.Connection")
    conn.commandtimeout=0
    conn.open "Driver={SQL Server};Server=172.22.166.33;Database=MISDBEMP;UID=sa;pwd=MIS@dbbox" 	
	'conn.open "Driver={SQL Server};Server=(local);Database=MISDBEMP;UID=sa;pwd=MIS@dbbox" 
	  
    Set GetSQLServerConnection = conn
End Function 


Function GetAccessConnection()
    Dim connykt
    Session.Timeout = 60
    Server.ScriptTimeOut=99999
    Set connykt = Server.CreateObject("ADODB.Connection")
    connykt.commandtimeout=0
    connykt.open "YKT"  
    Set GetAccessConnection = connykt
End Function 


%>
 
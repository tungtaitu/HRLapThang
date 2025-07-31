<%@ Language=VBScript codepage=950%> 
<!-- #include file = "../../GetSQLServerConnection.fun" --> 
<!-- #include file="../../ADOINC.inc" --> 
<%

dat1 = trim(request("dat1"))
dat2 = trim(request("dat2")) 

D1 = replace(dat1, "/", "" )
D2 = replace(dat2, "/", "" )   

aidstr = session("netuser") & minute(now())&second(now()) 

Set conn = GetSQLServerConnection()	  


conn.BeginTrans     
 
'sqlstr = "exec Ins_empWork '"& D1 &"' , '"& D2 &"' " 	  
'response.write sqlstr  
'response.end 
'conn.execute(sqlstr)    


Set cmd = Server.CreateObject("ADODB.Command")
   set cmd.ActiveConnection = conn
  	   cmd.CommandType = adCMDStoredProc    
       cmd.CommandText = "Ins_empWork" 
       cmd("@D1")= D1
       cmd("@D2")= D2  
  	   cmd.Execute  
	
if conn.Errors.Count = 0 then 
	conn.CommitTrans 	
	Set conn = Nothing 
	

%>	
	<script language=vbs>
		alert "資料轉入成功!!"
		open "accepted2.asp", "Fore" 
		'parent.best.cols="100%,0%"	
	</script>
<%
ELSE
	conn.RollbackTrans	
	'Response.Write "A=" & err.number  &"<BR>"
	'Response.Write "B=" & conn.errors.count &"<BR>"
	'response.write "C=" & conn.errors.Description &"<BR>"
%>	<SCRIPT LANGUAGE=VBSCRIPT>
		ALERT "DATA CommitTrans ERROR !!"
		OPEN "accepted2.asp" , "Fore" 
		'parent.best.cols="100%,0%"
	</script>	
<%	response.end 
END IF  
%>
 	

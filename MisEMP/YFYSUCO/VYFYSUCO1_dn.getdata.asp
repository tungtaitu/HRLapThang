<%@ Language=VBScript codepage=950%> 
<!-- #include file = "../GetSQLServerConnection.fun" --> 
<!-- #include file="../ADOINC.inc" --> 
<%
self="vyfysuco1dn"
kmym=request("kmym") 

Set conn = GetSQLServerConnection()	   
conn.BeginTrans      

 
 	sql2="exec SP_transDataSUCO_dn  '"& kmym &"' "
 	conn.execute(sql2) 
	
	sql2="exec SP_transDataSUCO_BC  '"& kmym &"' "
 	conn.execute(sql2)

if conn.Errors.Count = 0 or err.number = 0 then 
	conn.CommitTrans 	 
%>	
	<script language=vbs>
		alert "LA,DN,BC資料轉入成功(SUCCESS)!!"
		open "vyfysuco1.asp", "Fore" 
		parent.best.cols="100%,0%"	
	</script>
<%
ELSE
	conn.RollbackTrans	 
	
%>	<SCRIPT LANGUAGE=VBSCRIPT>
		ALERT "資料轉入失敗DATA CommitTrans ERROR (2)!!"  
		OPEN "vyfysuco1.asp" , "Fore" 
		parent.best.cols="100%,0%"
	</script>	
<%	response.end 
END IF  
%>
 	

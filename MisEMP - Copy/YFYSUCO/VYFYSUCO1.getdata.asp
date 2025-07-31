<%@ Language=VBScript codepage=950%> 
<!-- #include file = "../GetSQLServerConnection.fun" --> 
<!-- #include file="../ADOINC.inc" --> 
<%
self="vyfysuco1"
kmym=request("kmym") 

Set conn = GetSQLServerConnection()	   
conn.BeginTrans      

	sql="exec SP_transDataSUCO  '"& kmym &"' "
	'response.write sql
	'response.end	
	conn.execute(Sql)
 

if conn.Errors.Count = 0 or err.number = 0 then 
	conn.CommitTrans 
	conn.close	 
	set conn=nothing
	'response.redirect "vyfysuco1_dn.getdata.asp?kmym=" & kmym 
	'responsew.rite "2222"
	'response.end
%>	
	<script language=vbs>
		alert "資料轉入成功(SUCCESS)!!"
		open "<%=self%>.asp", "Fore" 
		parent.best.cols="100%,0%"	
	</script>
<%
ELSE
	conn.RollbackTrans	 
	
%>	<SCRIPT LANGUAGE=VBSCRIPT>
		ALERT "資料轉入失敗DATA CommitTrans ERROR !!"  
		OPEN "<%=self%>.asp" , "Fore" 
		parent.best.cols="100%,0%"
	</script>	
<%	response.end 
END IF  
%>
 	

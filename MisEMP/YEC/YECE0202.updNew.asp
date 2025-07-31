<%@LANGUAGE="VBSCRIPT" codepage=65001%>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<!-- #include file = "../GetSQLServerConnection.fun" -->
<!-- #include file="../ADOINC.inc" -->
<%
self="yece0202"  
Set conn = GetSQLServerConnection()
 
YYMM=REQUEST("YYMM")  
g1 = request("g1")
eid = request("eid")


if   yymm="" and g1="" and eid="" then 
	response.end 
else
	'conn.BeginTrans 
	sql="exec sp_newclcEmpwork '"&yymm&"', '"&g1&"', '"&eid&"' ;"
	response.write sql 
	conn.execute(sql)
end if 



 

if err.number = 0 then
	'conn.CommitTrans	
	Set conn = Nothing      
%>	<SCRIPT LANGUAGE=VBSCRIPT>
		'ALERT "資料處理成功!!" & chr(13) &"DATA CommitTrans Success !!"
		OPEN "<%=self%>.Foregnd.asp?yymm="&"<%=yymm%>"&"&groupid="&"<%=groupid%>"&"&eid="&"<%=eid%>" , "Fore"
		parent.best.rows="100%,*"
	</script>
<%'  response.end  
ELSE	
	Set conn = Nothing     
	conn.RollbackTrans       
%>	<SCRIPT LANGUAGE=VBSCRIPT>
		ALERT "資料處理失敗!!"&chr(13)&"DATA CommitTrans ERROR !!"
		OPEN "<%=self%>.asp" , "_parent"
	</script>
<%
	response.end
END IF
%>
 
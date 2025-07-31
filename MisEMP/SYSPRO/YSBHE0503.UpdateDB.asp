<%@language=vbscript codepage=65001%>
<!-------- #include file = "../GetSQLServerConnection.fun" --------->
<!--#include file = "../ADOINC.inc"-->
<HEAD>
<meta HTTP-EQUIV="Content-Type" content="text/html; charset=utf-8">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
</head>
<%
SELF="YSBHE0503"
muser = ucase(trim(request("muser")))
Username = trim(request("Username"))
Password = trim(request("Password"))
whsno = request("whsno")
groupid = request("groupid")

Set conn = GetSQLServerConnection()
Set rs = Server.CreateObject("ADODB.Recordset")

conn.BeginTrans

	sql="update  SYSUSER set password='"& Password  &"' where muser='"& muser &"'  "
	conn.execute(sql)

 if conn.Errors.Count = 0 then
	conn.CommitTrans
	Set conn = Nothing
%>	<script language=vbscript>
		alert "資料處理成功(密碼變更成功)!!"
		open "<%=self%>.asp", "Fore"
	</script>

<%
 else
	conn.RollbackTrans
	Set conn = Nothing
%>	<script language=vbscript>
		alert "資料處理失敗(密碼變更成功)!!"
		open "<%=self%>.asp", "Fore"
	</script>
<%	response.end
 end if %>

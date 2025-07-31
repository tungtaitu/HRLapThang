<%@language=vbscript codepage=65001%>
<!-------- #include file = "../GetSQLServerConnection.fun" --------->
<!--#include file = "../ADOINC.inc"-->
<HEAD>
<meta HTTP-EQUIV="Content-Type" content="text/html; charset=utf-8">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
</head>
<%
SELF="YEAAE0A01"
muser = ucase(trim(request("muser")))
Username = trim(request("Username"))
Password = trim(request("Password"))
whsno = request("whsno")
groupid = request("groupid")
old_password =  trim(request("old_password"))

Set conn = GetSQLServerConnection()
Set rs = Server.CreateObject("ADODB.Recordset")

conn.BeginTrans

	sql="update  SYSUSER set password='"& Password  &"' , username='"& Username &"' , mdtm=getdate() where muser='"& muser &"'  "
	conn.execute(sql)

 if conn.Errors.Count = 0 then
	conn.CommitTrans
	Set conn = Nothing
%>	<script language=javascript>
		alert("資料處理成功(密碼變更成功)OK!!");
		open("<%=self%>.asp", "Fore");
	</script>

<%
 else
	conn.RollbackTrans
	Set conn = Nothing
%>	<script language=javascript>
		alert("資料處理失敗 Error!!");
		open("<%=self%>.asp", "Fore");
	</script>
<%	response.end
 end if %>

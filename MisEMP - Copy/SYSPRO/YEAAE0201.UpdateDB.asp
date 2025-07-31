<%@language=vbscript codepage=65001%>
<!-------- #include file = "../GetSQLServerConnection.fun" --------->
<!--#include file = "../ADOINC.inc"-->
<%
SELF="YEAAE0201"
muser = ucase(trim(request("muser")))
Username = trim(request("Username"))
Password = trim(request("Password"))
whsno = request("whsno")
groupid = request("groupid")
rights = request("rights")

Set conn = GetSQLServerConnection()
Set rs = Server.CreateObject("ADODB.Recordset")

conn.BeginTrans

	Set rs2 = Server.CreateObject("ADODB.Recordset")
	SQL = "Select muser from SYSUSER where muser = '" & muser & "'"
	rs2.Open SQL, conn, 3, 3

	if not rs2.EOF then%>
		<script language="vbscript">
		<!--
		   alert "此代碼已存在請重新輸入!!"
		   open "<%=SELF%>.asp","Fore"
		//-->
		</script>
        <%
        Response.End
	else
		SQL_Insert = "Insert into SYSUSER (muser, username, Password, group_id, whsno, rights ) values " & _
		      "('" & muser & "', '" & username & "', '" & password & "', '" & groupid & "', '" & whsno & "' , '"& rights &"' )"
		conn.execute( SQL_Insert )
	end if

 if conn.Errors.Count = 0 then
	conn.CommitTrans
	Set conn = Nothing
%>	<script language=vbscript>
		alert "資料處理成功"
		open "<%=self%>.asp", "Fore"
	</script>

<%
 else
	conn.RollbackTrans
	Set conn = Nothing
%>	<script language=vbscript>
		alert "資料處理成功"
		open "<%=self%>.asp", "Fore"
	</script>
<%	response.end
 end if %>

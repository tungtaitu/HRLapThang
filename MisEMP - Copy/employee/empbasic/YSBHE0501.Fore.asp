<!-------- #include file = "../../GetSQLServerConnection.fun" --------->
<!--#include file="../../ADOINC.inc"-->
<%

Set conn = GetSQLServerConnection()
Set rs = Server.CreateObject("ADODB.Recordset")

self="YSBHE0501" 

sql="select * from sysuser where muser='"& session("netuser") &"'" 
rs.open sql, conn, 3, 3 
%>
<HTML>
<HEAD>
<meta HTTP-EQUIV="Content-Type" content="text/html; charset=BIG5">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
<link rel="stylesheet" href="../../Include/style.css" type="text/css">
<link rel="stylesheet" href="../../Include/style2.css" type="text/css">
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--
'-----------------enter to next field
function enterto()
	if window.event.keyCode = 13 then window.event.keyCode =9
end function

-->
</SCRIPT>
</HEAD>
<TITLE>�s�W�ϥΪ�</TITLE>
<body background="bg_blue.gif"  topmargin=5 onload=f()  onkeydown="enterto()" >
<FORM action="<%=SELF%>.UpdateDB.asp" method="POST" name="<%=SELF%>">
<table width="460" border="0" cellspacing="0" cellpadding="0">
	<tr><TD width=100%>
	<img border="0" src="../../image/icon.gif" align="absmiddle">
	�s�W�ϥΪ� </TD></tr>
</table>
<hr size=0	style='border: 1px dotted #999999;' align=left width=500>
<INPUT TYPE=HIDDEN NAME="UID" VALUE=<%=SESSION("NETUSER")%>>
<table width=550><tr><td align=center ><BR><BR><BR>

	<table  border="0" cellspacing="1" cellpadding="1" bgcolor="#8EB9D9" class=txt9 >		
		<tr bgcolor="#FFFFD9" >
			<td  align=center  height=25>�ϥΪ̱b��</td>
			<td   align=center>
				   <input type="text" size=8 name="muser" maxlength="5" style="font size: 9pt" class=readonly2 readonly   value="<%=rs("muser")%>">
				   <input type="text" size=15 name="Username"  value="<%=rs("username")%>"  class=readonly2 readonly ><br>
			</td>
		</tr>
		<tr bgcolor="#FFFFD9" >
			<td   align=center height=25>�t�O</td>
			<td >
				<select name=whsno class=inputbox disabled >
				<%sql="select * from basicCode where func='WHSNO' order by sys_type"
				  set rds=conn.execute(sql)
				  while not rds.eof
				%>
					<option value="<%=rds("sys_type")%>" <%if rs("whsno")=rds("sys_type") then %>selected<%end if%>><%=rds("sys_type")%>-<%=rds("sys_value")%></option>
				<%rds.movenext
				wend
				%>
				</select>
			</td>
		</tr>
		<tr bgcolor="#FFFFD9" >
			<td   align=center height=25>���</td>
			<td >
				<select name=groupid class=inputbox disabled >
				<%sql="select * from basicCode where func='groupid' order by sys_type"
				  set rds=conn.execute(sql)
				  while not rds.eof
				%>
					<option value="<%=rds("sys_type")%>" <%if rs("group_id")=rds("sys_type") then %>selected<%end if%> ><%=rds("sys_type")%>-<%=rds("sys_value")%></option>
				<%rds.movenext
				wend
				%>
				</select>
			</td>
		</tr>		
		<tr bgcolor="#FFFFD9" >
			<td   align=center height=25>�ܧ�K�X</td>
			<td >
				<input type="password" size=22 name="Password" maxlength="10" style="font size: 9pt" class=inputbox><br>
			</td>
		</tr> 
		<tr bgcolor="#FFFFD9" >
			<td   align=center eight=25>�T�{�K�X</td>
			<td bgcolor="#FFFFD9"  >
				<input type="password" size=22 name="Password2" maxlength="10" style="font size: 9pt" class=inputbox ><br>
			</td>
		</tr>		
		<tr>
        
        <td align="CENTER"  colspan=2 bgcolor="#FFFFD9" height=30 >
        	<%if UCASE(session("mode"))="W" then%>
				<input type="button" name="send" value="��     �J" onclick="GO()" class="button">
            	<input type="reset" name="send" value="��     ��" class="button" >
			<%end if%>
            
        </td>
      </tr>

	</table>
 </FORM>
</td></tr></table>
</BODY>
</HTML>
<script ID=clientEventHandlersVBS language="vbscript">
function Go()
	if trim(YSBHE0501.Password.value ) <> trim(YSBHE0501.Password2.value) then
	   alert "��J���K�X���~!!�Э��s��J!!"
	   YSBHE0501.Password2.value = ""
	end if
	YSBHE0501.submit
end function

function f()
	<%=self%>.Password.focus()
end function

function chg1()
	<%=self%>.muser.value=Ucase(<%=self%>.muser.value)
end function
</script>



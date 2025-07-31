<%@language=VBScript codepage=65001%>
<!-------- #include file ="../GetSQLServerConnection.fun" --------->
<!--#include file="../ADOINC.inc"-->
<!--#include file="../include/sideinfo.inc"-->
<%

Set conn = GetSQLServerConnection()
Set rs = Server.CreateObject("ADODB.Recordset")

self="YEAAE0A01" 

sql="select * from sysuser where muser='"& session("netuser") &"'" 
rs.open sql, conn, 3, 3 
%>
<HTML>
<HEAD>
<meta HTTP-EQUIV="Content-Type" content="text/html; charset=utf-8">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
<link rel="stylesheet" href="../Include/style.css" type="text/css">
<link rel="stylesheet" href="../Include/style2.css" type="text/css">

<SCRIPT LANGUAGE="Javascript">

//'-----------------enter to next field
function enterto(){
	if (window.event.keyCode == 13) 
    {
        window.event.keyCode =9;
    }
}

</SCRIPT>
</HEAD>
<body onload=f()  onkeydown="enterto()" >
<FORM action="<%=SELF%>.UpdateDB.asp" method="POST" name="<%=SELF%>"> 
<INPUT TYPE=HIDDEN NAME="UID" VALUE=<%=SESSION("NETUSER")%>>
	<table border=0 style="height:70px"><tr><td>&nbsp;</td></tr></table>
	
	<table width="100%" BORDER=0 align=center cellpadding=0 cellspacing=0 >
		<tr>
			<td align="center">
				<table  id="myTableForm" width="60%">		
					<tr><td colspan=4 >&nbsp;</td></tr>
					<tr>
						<td class="frmtable-label" style="width:100px">帳號<br>ID</td>
						<td>
							<input type="text" size=8 name="muser" maxlength="5" style="font size: 9pt"  readonly   value="<%=rs("muser")%>">				   
						</td>								
						<td class="frmtable-label">名稱<br>Ten</td>
						<td >				   
							   <input type="text" size=25 name="Username"  value="<%=rs("username")%>"   >
						</td>
					</tr>		
					<tr >
						<td class="frmtable-label">廠別<br>Xuong</td>
						<td >
							<select name=whsno  disabled >
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
						<td class="frmtable-label">單位<br>Don vi</td>
						<td >
							<select name=groupid  disabled >
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
					<tr>
						<td class="frmtable-label">舊密碼<br>mat ma Cu</td>
						<td colspan="3">
							<input  size=22 name="old_password" maxlength="10"  readonly  value="<%=rs("password")%>">
						</td>
					</tr> 		
					<tr>
						<td class="frmtable-label">密碼<br>mat ma</td>
						<td >
							<input type="password" size=22 name="Password" maxlength="10" style="font size: 9pt" >
						</td>								
						<td class="frmtable-label">確認密碼<br>mat ma</td>
						<td>
							<input type="password" size=22 name="Password2" maxlength="10" style="font size: 9pt"  >
						</td>
					</tr>		
					<tr>								
						<td align="CENTER"  colspan=4 height="40px">
							<%if UCASE(session("mode"))="W" then%>
								<input type="button" name="send" value="(Y)Confirm" onclick ="go()" class="btn btn-sm btn-danger" onkeydown="go()">
								<input type="reset" name="send" value="(N)Cancel" class="btn btn-sm btn-outline-secondary" >
							<%end if%>										
						</td>
				  </tr>
				</table>
			</td>
		</tr>
	</table>
			
</FORM>
</BODY>
</HTML>

<script type="text/javascript" language="javascript">
    function go() {
        if(<%=self%>.Password.value == "" || ( <%=self%>.Password.value != <%=self%>.Password2.value))
        {
	        alert ("密碼有誤或不得為空白!!請重新輸入!!xin danh lai mat ma ");
			<%=self%>.Password.focus();
			<%=self%>.Password2.value = "";	   	    
        }
		else
			<%=self%>.submit();       
    }
	
	function f(){
		<%=self%>.Password.focus();
	}

</script>




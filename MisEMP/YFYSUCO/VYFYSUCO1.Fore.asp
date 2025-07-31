<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
 
<!-- #include file = "../GetSQLServerConnection.fun" --> 
<!-- #include file="../ADOINC.inc" -->
<!--#include file="../include/sideinfo.inc"-->
<%
'on error resume next   
session.codepage="65001"
SELF = "vyfysuco1"
 
Set conn = GetSQLServerConnection()	  
Set rs = Server.CreateObject("ADODB.Recordset")   

DAT1 = REQUEST("DAT1")
DAT2 = REQUEST("DAT2")

IF DAT1="" THEN DAT1=DATE()-1
IF DAT2="" THEN DAT2=DATE()-1


FUNCTION FDT(D)
	Response.Write YEAR(D)&"/"&RIGHT("00"&MONTH(D),2)&"/"&RIGHT("00"&DAY(D),2) 
	
END FUNCTION 

nowmonth = year(date())&right("00"&month(date()),2)  
if month(date())="01" then  
	calcmonth = year(date()-1)&"12" 
else
	calcmonth =  year(date())&right("00"&month(date())-1,2)   
end if 	  



%>

<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="refresh">
<link rel="stylesheet" href="../Include/style.css" type="text/css">
<link rel="stylesheet" href="../Include/style2.css" type="text/css"> 
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--
'-----------------enter to next field
function enterto()
	if window.event.keyCode = 13 then window.event.keyCode =9 		
end function  

function f()
	<%=self%>.kmym.focus()	
	<%=self%>.kmym.SELECT()
end function   


-->
</SCRIPT>  


</head>   
<body  topmargin="5" leftmargin="5"  marginwidth="0" marginheight="0"  onkeydown="enterto()" onload="f()">
<form name="<%=self%>" method="post" action="acceptedcatime.asp">
<INPUT TYPE=HIDDEN NAME="UID" VALUE=<%=SESSION("NETUSER")%>>
	
	<table border=0 style="height:70px"><tr><td>&nbsp;</td></tr></table>

	<table width="100%" BORDER=0 align=center cellpadding=0 cellspacing=0 >
		<tr>
			<td align="center">
				<table class="txt" cellpadding=3 cellspacing=3>   
					<TR height=25>
						<TD nowrap align=right>扣款年月<br>Năm tháng trừ</TD>
						<TD>
							<INPUT  type="text" style="width:120px" NAME=kmym VALUE="<%=calcmonth%>"> 										
						</TD>
						<td>(EX:200601)</td>
					</TR>	 	 
					<TR>
						<td COLSPAN=3 ALIGN=CENTER HEIGHT=50>
							<input type="button" name="send" value="確定 Xác nhận" class="btn btn-sm btn-danger"  onclick="go()"  onkeydown="go()">
							<input type="RESET"  name="send" value="取消 Thiết lập lại" class="btn btn-sm btn-outline-secondary" >	
						</td>	
					</TR>
				</table>
			</td>
		</tr>
	</table>
			
</form>

</body>
</html>

<script language=vbscript>
function BACKMAIN()	
	open "../main.asp" , "_self"
end function     

function go()
	<%=self%>.action="forwait_tmp.asp"
	<%=self%>.target="Fore"
	<%=self%>.submit()
end function 
</script>


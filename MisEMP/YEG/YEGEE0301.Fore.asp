<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<!-- #include file = "../GetSQLServerConnection.fun" --> 
<!-- #include file="../ADOINC.inc" -->
<!-- #include file="../Include/func.inc" -->
<!--#include file="../include/sideinfo.inc"-->
<%
'on error resume next   
session.codepage="65001"
SELF = "YEGEE0301"
 
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
	<%=self%>.DAT1.focus()	
	<%=self%>.DAT1.SELECT()
end function   


-->
</SCRIPT>  


</head>   
<body  topmargin="5" leftmargin="5"  marginwidth="0" marginheight="0"  onkeydown="enterto()" onload="f()">
<form name="<%=self%>" method="post" action="acceptedcatime.asp">
<INPUT TYPE=HIDDEN NAME="UID" VALUE=<%=SESSION("NETUSER")%>>
	
	<table border=0 style="height:70px"><tr><td>&nbsp;</td></tr></table>
	<table width="100%" border=0 >
		<tr>
			<td>
				<table width="80%" BORDER=0 align=center cellpadding=0 cellspacing=0 >
					<tr>
						<td>
							<table class="table-borderless table-sm bg-white text-secondary">   
								<TR>
									<TD nowrap align=right>接收日期：</TD>
									<TD><INPUT NAME=DAT1 SIZE=12 CLASS="form-control form-control-sm mb-2 mt-2"  VALUE="<%=FDT(DAT1)%>" onblur="date_change(1)"></td>
									<td>~</td> 	
									<td><INPUT NAME=DAT2 SIZE=12 CLASS="form-control form-control-sm mb-2 mt-2"  VALUE="<%=FDT(DAT2)%>" onblur="date_change(2)"></TD> 
								</TR>	 	 
								<TR>
									<TD nowrap align=right>工號:</TD>
									<TD colspan=3>
										<INPUT NAME=eid SIZE=12 CLASS="form-control form-control-sm mb-2 mt-2"  VALUE="" >		
									</TD> 
								</TR>	 	 
								<TR>
									<td COLSPAN=4 ALIGN=CENTER HEIGHT=50>
										<input type="button" name="send" value="(Y)確　定" class="btn btn-sm btn-danger"  onclick="go()" onkeydown="go()" >
										<input type="RESET"  name="send" value="(N)取　消" class="btn btn-sm btn-outline-secondary" >	
									</td>	
								</TR>
							</table>
						</td>
					</tr>
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
	<%=self%>.action="forwait.index.asp"	
	
	<%=self%>.submit()   
end function  

'*******檢查日期*********************************************
FUNCTION date_change(a)

if a=1 then
	INcardat = Trim(<%=self%>.dat1.value)
elseif a=2 then
	INcardat = Trim(<%=self%>.dat2.value)
end if

IF INcardat<>"" THEN
	ANS=validDate(INcardat)
	IF ANS <> "" THEN
		if a=1 then
			Document.<%=self%>.dat1.value=ANS
		elseif a=2 then
			Document.<%=self%>.dat2.value=ANS
		end if
	ELSE
		ALERT "EZ0067:輸入日期不合法 !!"
		if a=1 then
			Document.<%=self%>.dat1.value=""
			Document.<%=self%>.dat1.focus()
		elseif a=2 then
			Document.<%=self%>.dat2.value=""
			Document.<%=self%>.dat2.focus()
		end if
		EXIT FUNCTION
	END IF

ELSE
	'alert "EZ0015:日期該欄位必須輸入資料 !!"
	EXIT FUNCTION
END IF

END FUNCTION 
</script>


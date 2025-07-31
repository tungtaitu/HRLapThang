<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<!-- #include file = "../GetSQLServerConnection.fun" -->
<!-- #include file="../ADOINC.inc" -->
<!--#include file="../include/sideinfo.inc"-->
<%

SELF = "YEIE0102"

Set conn = GetSQLServerConnection()
Set rs = Server.CreateObject("ADODB.Recordset")
Set rds = Server.CreateObject("ADODB.Recordset")

nowmonth = year(date())&right("00"&month(date()),2)
if month(date())="01" then
	calcmonth = year(date()-1)&"12"
else
	calcmonth =  year(date())&right("00"&month(date())-1,2)
end if 

F_whsno = request("F_whsno")
F_groupid = request("F_groupid")
F_zuno = request("F_zuno") 
if F_whsno="" then F_whsno="XX"
F_shift=request("F_shift")
F_empid =request("F_empid")
F_country=request("F_country")
 

FUNCTION FDT(D)
	Response.Write YEAR(D)&"/"&RIGHT("00"&MONTH(D),2)&"/"&RIGHT("00"&DAY(D),2)
END FUNCTION 
khym = request("khym")
if request("khym")="" then 
	khym=nowmonth
end if  

act = request("act")	  
khweek = request("khweek") 

tmw = request("tmw")
if tmw="" then tmw=request("tt")
 '一個月有幾天 
cDatestr=CDate(LEFT(khym,4)&"/"&RIGHT(khym,2)&"/01") 
days = DAY(cDatestr+(32-DAY(cDatestr))-DAY(cDatestr+(32-DAY(cDatestr))))   '一個月有幾天  
'本月最後一天 
ENDdat = LEFT(khym,4)&"/"&RIGHT(khym,2)&"/"&DAYS   

'if khweek="" then khweek=(days\7)  
 
%>

<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta HTTP-EQUIV="refresh" >
<link rel="stylesheet" href="../Include/style.css" type="text/css">
<link rel="stylesheet" href="../Include/style2.css" type="text/css">
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--
'-----------------enter to next field
function enterto()
	if window.event.keyCode = 13 then window.event.keyCode =9
end function

function f()
	if <%=self%>.act.value="A" then 
		<%=self%>.F_whsno.focus()
	else
		<%=self%>.khym.focus()
		<%=self%>.khym.select()
	end if 	
end function

function groupchg()
	code = <%=self%>.GROUPID.value
	open "<%=self%>.back.asp?ftype=groupchg&code="&code , "Back"
	'parent.best.cols="50%,50%"
end function

function unitchg()
	code = <%=self%>.unitno.value
	open "<%=self%>.back.asp?ftype=UNITCHG&code="&code , "Back"	
	'parent.best.cols="50%,50%"
end function 

function datachg() 
	<%=self%>.totalpage.value="0"
	<%=self%>.action = "<%=self%>.Fore.asp"
	<%=self%>.submit()
end function  
function datachg2() 
	<%=self%>.totalpage.value="0"
	<%=self%>.act.value="A"
	<%=self%>.action = "<%=self%>.Fore.asp"
	<%=self%>.submit()
end function 

-->
</SCRIPT>
</head>
<body  topmargin="0" leftmargin="5"  marginwidth="0" marginheight="0"  onkeydown="enterto()" onload="f()">
<form  name="<%=self%>" method="post" action="<%=self%>.upd.asp"   >
<INPUT TYPE=HIDDEN NAME="UID" VALUE=<%=SESSION("NETUSER")%>>
<INPUT TYPE=hidden NAME=TotalPage VALUE="<%=TotalPage%>">
<INPUT TYPE=hidden NAME=CurrentPage VALUE="<%=CurrentPage%>">
<INPUT TYPE=hidden NAME=RecordInDB VALUE="<%=RecordInDB%>">
<INPUT TYPE=hidden NAME=PageRec VALUE="<%=PageRec%>">
<INPUT TYPE=hidden NAME=gTotalPage VALUE="<%=gTotalPage%>"> 	
<INPUT TYPE=hidden NAME=nowmonth VALUE="<%=nowmonth%>">
<INPUT TYPE=hidden NAME=days VALUE="<%=days%>">
<input name=act value="<%=act%>" type=hidden >

	<table border=0 style="height:70px"><tr><td>&nbsp;</td></tr></table>
	<table width="100%" BORDER=0 align=center cellpadding=0 cellspacing=0 >
		<tr>
			<td align="center">
				<table id="myTableForm" width="60%">
					<tr >
						<TD colspan=4 align="center">&nbsp;</TD>
					</TR>
					<TR> 
						<TD align=right>考核年月</TD>
						<td><input type="text" style="width:100px" name=khym value="<%=khym%>" maxlength=6 ></td>	
						<td align=right nowrap>國籍<br>Quoc tich</td>
						<td>
							<select name=F_country style="width:120px">												
								<option value="" selected >全部(Toan bo) </option>
								<%		
								'else
									SQL="SELECT * FROM BASICCODE WHERE FUNC='country'  and sys_type not in ('Tw' ) ORDER BY SYS_TYPE desc "
								'end if	
								SET RST = CONN.EXECUTE(SQL)
								WHILE NOT RST.EOF
								%>
								<option value="<%=RST("SYS_TYPE")%>" <%if RST("SYS_TYPE")=F_country then%>selected<%end if%>><%=RST("SYS_VALUE")%></option>
								<%
								RST.MOVENEXT
								WEND
								%>
							</SELECT>
							<%SET RST=NOTHING %>	
						</td> 
					</tr> 
					<TR> 
						<TD align=right>廠別<br>Xuong</TD>
						<td>
							<select name=F_whsno  style="width:120px">					
									<%
									if session("rights")="0" then 
										SQL="SELECT * FROM BASICCODE WHERE FUNC='WHSNO' ORDER BY SYS_TYPE "%>
										<option value="" selected >全部(Toan bo) </option>
									<%		
									else
										SQL="SELECT * FROM BASICCODE WHERE FUNC='WHSNO'  ORDER BY SYS_TYPE " 'and sys_type='"& session("NETWHSNO") &"'
									end if	
									SET RST = CONN.EXECUTE(SQL)
									WHILE NOT RST.EOF
									%>
									<option value="<%=RST("SYS_TYPE")%>" <%if RST("SYS_TYPE")=F_whsno then%>selected<%end if%>><%=RST("SYS_VALUE")%></option>
									<%
									RST.MOVENEXT
									WEND
									%>
							</SELECT>
								<%SET RST=NOTHING %>	
						</td>								
						<TD align=right nowrap>部門<br>Bo Phan</TD>
						<td>
							<select name=F_groupid style="width:160px" >										
								<option value="">全部(Toan bo) </option>
								<% 
									SQL="SELECT * FROM BASICCODE WHERE FUNC='GROUPID' and sys_type <>'AAA' ORDER BY SYS_TYPE " 
								'else
								'	SQL="SELECT * FROM BASICCODE WHERE FUNC='GROUPID' and sys_type <>'AAA' and  sys_type= '"& session("NETG1") &"' ORDER BY SYS_TYPE "
								'end if   
								
								SET RST = CONN.EXECUTE(SQL)
								RESPONSE.WRITE SQL 
								WHILE NOT RST.EOF  
							%>
								<option value="<%=RST("SYS_TYPE")%>" <%if RST("SYS_TYPE")= session("NETG1") then%>selected<%end if%>><%=RST("SYS_TYPE")%><%=RST("SYS_VALUE")%></option>				 
							<%
								RST.MOVENEXT
								WEND 
							%>
							</SELECT>
							<%SET RST=NOTHING %> 
						</td>
					</tr>									  
					<TR>
						<td nowrap align=right ><a href="vbscript:gotemp()"><font color=blue><u>員工編號：</u></font></a></td>
						<td colspan=3 nowrap>
							<input type="text" style="width:100px" name=empid  maxlength=5 onchange=strchg(1)>
							<input type="text" style="width:100px" name=empname readonly class="readonly"   maxlength=5 >
						</td>
					</TR> 		
					<TR> 
						<td align=right nowrap>類別</td>
						<td>
							<select name=fclass style="width:120px">				
								<option value="A" >月統計</option>
								<option value="B" >周統計</option>
							</select>			
						</td>
						<td align=right nowrap>班別<br>Ca</td>
						<td>
							<select name=F_shift  style="width:120px">
								<option value=""></option>
								<option value="ALL" <%if F_shift="ALL" then%>selected<%end if%>>日</option>
								<option value="A" <%if F_shift="A" then%>selected<%end if%>>A班</option>
								<option value="B" <%if F_shift="B" then%>selected<%end if%>>B班</option>
							</select>			
						</td> 
					</TR>
					<tr >
						<TD colspan=4 align="center">
							<input type=button  name=send value="(Y)確　　認"  class="btn btn-sm btn-danger" onclick=go()>
							<input type=RESET name=send value="(N)取 　　消"  class="btn btn-sm btn-outline-secondary">
						</TD>
					</TR>
					<tr >
						<TD colspan=4 align="center">&nbsp;</TD>
					</TR>
				</table>
			</td>
		</tr>
	</table>
			
</form>


</body>
</html>
<script language=vbscript>  

function gotemp()
	open "../getempdata.asp?formName="&"<%=self%>", "Back"
	parent.best.cols="65%,35%"
end function  
 
function  go()
	if <%=self%>.khym.value="" then 
		alert "考核年月不可為空!!"
		<%=self%>.khym.focus()
		<%=self%>.khym.select()
	end if 
	<%=self%>.action="<%=self%>.ForeGnd.asp"
	<%=self%>.submit()
end function 
  
</script>


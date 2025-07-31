<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<!-- #include file = "../GetSQLServerConnection.fun" -->
<!-- #include file="../ADOINC.inc" -->
<!-- #include file="../Include/func.inc" -->
<!--#include file="../include/sideinfo.inc"-->
<%
Set conn = GetSQLServerConnection()
self="YEBE0301"


nowmonth = year(date())&right("00"&month(date()),2)
if month(date())="01" then
	calcmonth = year(date()-1)&"12"
else
	calcmonth =  year(date())&right("00"&month(date())-1,2)
end if

if day(date())<=11 then
	if month(date())="01" then
		calcmonth = year(date()-1)&"12"
	else
		calcmonth =  year(date())&right("00"&month(date())-1,2)
	end if
else
	calcmonth = nowmonth
end if

if instr(session("vnlogip"),"168")>0 then
	w1="LA"
elseif	instr(session("vnlogip"),"169")>0 then
	w1="DN"
elseif	instr(session("vnlogip"),"47")>0 then
	w1="BC"
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
function f()
	<%=self%>.inym.focus()
	'<%=self%>.country.SELECT()
end function
-->
</SCRIPT>
</head>
<body   onkeydown="enterto()" onload='f()' >
<form name="<%=self%>" method="post"  >
	
	<table border=0 style="height:70px"><tr><td>&nbsp;</td></tr></table>
	<table width="100%" border=0 >
		<tr>
			<td>
				<table width="80%" BORDER=0 align=center cellpadding=0 cellspacing=0 >
					<tr>
						<td align=center>
							<table id="myTableForm" width="60%">
								<tr><td colspan=4 height="40px">&nbsp;</td></tr>
								<TR>
									<TD nowrap align=right height=30 >到職年月<br><font class=txt8>Thong ke Thang Nam</font></TD>
									<TD colspan=2>
										 <input name=inym  size=15 maxlength=6  value="">
									</TD>
								</tr>
								<TR>
									<TD nowrap align=right height=30 >國籍<br><font class=txt8>Quoc Tich</font></TD>
									<TD colspan=2>
										<select name=country   style='width:90'  >
											<option value="">全部ALL </option>
											<%SQL="SELECT * FROM BASICCODE WHERE FUNC='country'  ORDER BY SYS_type desc  "
											SET RST = CONN.EXECUTE(SQL)
											WHILE NOT RST.EOF
											%>
											<option value="<%=RST("SYS_TYPE")%>" ><%=RST("SYS_TYPE")%> - <%=RST("SYS_VALUE")%></option>
											<%
											RST.MOVENEXT
											WEND
											rst.close
											SET RST=NOTHING
											%>
										</SELECT>				
									</TD>
								</tr>
								<tr>
									<TD nowrap align=right height=30 >廠別<br><font class=txt8>Loai Xuong</font></TD>
									<TD colspan=2>
										<select name=WHSNO   >
											<option value="">全部ALL</option>
											<%SQL="SELECT * FROM BASICCODE WHERE FUNC='WHSNO' ORDER BY SYS_TYPE "
											SET RST = CONN.EXECUTE(SQL)
											WHILE NOT RST.EOF
											%>
											<option value="<%=RST("SYS_TYPE")%>" <%if w1=RST("SYS_TYPE") then%>selected<%end if%>><%=RST("SYS_VALUE")%></option>
											<%
											RST.MOVENEXT
											WEND
											rst.close
											SET RST=NOTHING
											%>
										</SELECT>
									</TD>
								</TR>
								<tr>		
									<TD nowrap align=right >組/部門<br><font class=txt8>Bo Phan</font></TD>
									<TD >
										<select name=GROUPID   style="width:90"  >
										<option value="" selected >全部ALL </option>
										<%
										SQL="SELECT * FROM BASICCODE WHERE FUNC='GROUPID' and sys_type <>'AAA' ORDER BY SYS_TYPE "
										'SQL="SELECT * FROM BASICCODE WHERE FUNC='GROUPID'  ORDER BY SYS_TYPE "
										SET RST = CONN.EXECUTE(SQL)
										'RESPONSE.WRITE SQL
										WHILE NOT RST.EOF
										%>
										<option value="<%=RST("SYS_TYPE")%>" ><%=RST("SYS_TYPE")%><%=RST("SYS_VALUE")%></option>
										<%
										RST.MOVENEXT
										WEND
										rst.close
										SET RST=NOTHING
										%>
										</SELECT>
									</td>
									<td>	
										<select name=zuno   style="width:120"  >
										<option value="" selected >全部ALL </option>
										<%
										SQL="SELECT * FROM BASICCODE WHERE FUNC='zuno' and sys_type <>'AAA' ORDER BY SYS_TYPE "
										'SQL="SELECT * FROM BASICCODE WHERE FUNC='GROUPID'  ORDER BY SYS_TYPE "
										SET RST = CONN.EXECUTE(SQL)
										'RESPONSE.WRITE SQL
										WHILE NOT RST.EOF
										%>
										<option value="<%=RST("SYS_TYPE")%>" ><%=RST("SYS_TYPE")%><%=RST("SYS_VALUE")%></option>
										<%
										RST.MOVENEXT
										WEND
										rst.close
										SET RST=NOTHING
										%>
										</SELECT>								
									</td>
								</tr>		
								<%
								conn.close
								set conn=nothing
								%>		
								<TR  height=30 >
									<td nowrap align=right >員工編號<br><font class=txt8>So The</font></td>
									<td colspan=2>
										<input name=empid1  size=15 maxlength=6 onchange=strchg(1)>

									</td>
								</TR>
								<TR  height=30 >
									<td nowrap align=right >簽約統計<br><font class=txt8>ky hop dong</font></td>
									<td colspan=3>
										<select name=outemp >
											<option value="">全部ALL</option>
											<option value="Y">已簽約(Da ky hop dong)</option>
											<option value="N">未簽約(Chua ky hop dong)</option>
										</select>
									</td>
								</TR>
								<TR  height=30 >
									<td nowrap align=right >員工統計<br><font class=txt8>Thong ke Nhan vien</font></td>
									<td colspan=2>
										<select name=IOemp >
											<option value="Y">在職(Tai Chuc)</option>
											<option value="">全部(ALL)</option>
											<option value="N">已離職(Thoi viec</option>
										</select>
									</td>
								</TR>
								<tr >
									<td align=center colspan=3 height="60px">
										<input type=button  name=btm class="btn btn-sm btn-danger" value="(Y)Confirm" onclick="go()" onkeydown="go()">
										<input type=reset  name=btm class="btn btn-sm btn-outline-secondary" value="(N)Cancel">&nbsp;
										<input type=button  name=btm class="btn btn-sm btn-outline-secondary" value="轉入稅號(Ins.MST)"  onclick="insmst()">
									</td>
								</tr>							
							</table>
						</td>
					</tr>					
				</table>
			</td>
		</tr>
	</table>

</body>
</html>


<script language=vbs>
function dataclick(a)
	if a = 1 then
		open "empbasic/empbasic.asp" , "_self"
	elseif a = 2 then
		open "empfile/empfile.asp" , "_self"
	elseif a = 3 then
		open "empworkHour/empwork.asp" , "_self"
	elseif a = 4 then
		open "holiday/empholiday.asp" , "_self"
	elseif a = 5 then
		open "AcceptCaTime/main.asp" , "_self"
	elseif a = 6 then
		open "../report/main.asp" , "_self"
	end if
end function

function strchg(a)
	if a=1 then
		<%=self%>.empid1.value = Ucase(<%=self%>.empid1.value)
	elseif a=2 then
		<%=self%>.empid2.value = Ucase(<%=self%>.empid2.value)
	end if
end function

function go()
 	<%=self%>.action="<%=self%>.foregnd.asp"
 	<%=self%>.submit()
end function

function insmst()
	wt = (window.screen.width )*0.6
	ht = window.screen.availHeight*0.6
	tp = (window.screen.width )*0.05
	lt = (window.screen.availHeight)*0.1	
	open "<%=self%>.insMst.asp" , "_blank"  , "top="& tp &", left="& lt &", width="& wt &",height="& ht &", scrollbars=yes"
end function 


'*******檢查日期*********************************************
FUNCTION date_change(a)

if a=1 then
	INcardat = Trim(<%=self%>.indat1.value)
elseif a=2 then
	INcardat = Trim(<%=self%>.indat2.value)
end if

IF INcardat<>"" THEN
	ANS=validDate(INcardat)
	IF ANS <> "" THEN
		if a=1 then
			Document.<%=self%>.indat1.value=ANS
		elseif a=2 then
			Document.<%=self%>.indat2.value=ANS
		end if
	ELSE
		ALERT "EZ0067:輸入日期不合法 !!"
		if a=1 then
			Document.<%=self%>.indat1.value=""
			Document.<%=self%>.indat1.focus()
		elseif a=2 then
			Document.<%=self%>.indat2.value=""
			Document.<%=self%>.indat2.focus()
		end if
		EXIT FUNCTION
	END IF

ELSE
	'alert "EZ0015:日期該欄位必須輸入資料 !!"
	EXIT FUNCTION
END IF

END FUNCTION
</script> 
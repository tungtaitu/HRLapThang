<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<!-- #include file = "../../GetSQLServerConnection.fun" -->
<!-- #include file="../../ADOINC.inc" -->
<%

SELF = "empfilemain"

Set conn = GetSQLServerConnection()
'Set rs = Server.CreateObject("ADODB.Recordset")

nowmonth = year(date())&right("00"&month(date()),2)
if month(date())="01" then
	calcmonth = year(date()-1)&"12"
else
	calcmonth =  year(date())&right("00"&month(date())-1,2)
end if

sql="select max(empid) empid  from empfile where country='Vn' and left(empid,1)='L' "
set rds=conn.execute(sql)
if not rds.eof then
	eid = "L" & right("0000" & cstr(cdbl(right(rds("empid"),4))+1) , 4)
else
	eid=""
end if
set rds=nothing

FUNCTION FDT(D)
	Response.Write YEAR(D)&"/"&RIGHT("00"&MONTH(D),2)&"/"&RIGHT("00"&DAY(D),2)
END FUNCTION
%>

<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta HTTP-EQUIV="refresh" >
<link rel="stylesheet" href="../../Include/style.css" type="text/css">
<link rel="stylesheet" href="../../Include/style2.css" type="text/css">
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--
'-----------------enter to next field
function enterto()
	if window.event.keyCode = 13 then window.event.keyCode =9
end function

function f()
	<%=self%>.EMPID.focus()
	<%=self%>.EMPID.select()
end function

function groupchg()
	code = <%=self%>.GROUPID.value
	open "empfile.back.asp?ftype=groupchg&code="&code , "Back"
	'parent.best.cols="50%,50%"
end function

function unitchg()
	code = <%=self%>.unitno.value
	open "empfile.back.asp?ftype=UNITCHG&code="&code , "Back"	
	'parent.best.cols="50%,50%"
end function
-->
</SCRIPT>
</head>
<body  topmargin="0" leftmargin="0"  marginwidth="0" marginheight="0"  onkeydown="enterto()" onload="f()">
<form name="<%=self%>" method="post" action="empfile.upd.asp">
<INPUT TYPE=HIDDEN NAME="UID" VALUE=<%=SESSION("NETUSER")%>>
<INPUT TYPE=hidden NAME=nowmonth VALUE="<%=nowmonth%>">
<table width="460" border="0" cellspacing="0" cellpadding="0">
	<tr><TD  >
		<img border="0" src="../../image/icon.gif" align="absmiddle">
		越籍員工基本資料
		</TD>
	</tr>
</table>
<hr size=0	style='border: 1px dotted #999999;' align=left width=500>

<TABLE WIDTH=520 CLASS=FONT9 BORDER=0>
	<TR height=25 >
		<TD width=80 nowrap align=right height=25>員工編號：</TD>
		<TD width=100><INPUT NAME="EMPID" SIZE=12 CLASS=INPUTBOX  ONCHANGE='EMPIDCHG()' maxlength=5 value="<%=eid%>"></TD>
		<TD width=60 nowrap align=right height=25>到職日：</TD>
		<TD ><INPUT NAME=indat SIZE=12 CLASS=INPUTBOX value=<%=fdt(date())%> onblur="date_change(1)"></TD>
		<td width="150" rowspan="5" align=center valign=center><img src="../photos/nophotos.gif" width="130" height="130" border=1></td>
	</TR>
	<TR height=25 >
		<TD nowrap align=right>廠別：</TD>
		<TD >
			<select name=WHSNO  class=font9 >
				<%SQL="SELECT * FROM BASICCODE WHERE FUNC='WHSNO' ORDER BY SYS_TYPE "
				SET RST = CONN.EXECUTE(SQL)
				WHILE NOT RST.EOF
				%>
				<option value="<%=RST("SYS_TYPE")%>" <%IF SESSION("NETWHSNO")=RST("SYS_TYPE") THEN %> SELECTED <%END IF%>><%=RST("SYS_VALUE")%></option>
				<%
				RST.MOVENEXT
				WEND
				%>
			</SELECT>
			<%SET RST=NOTHING %>
		</TD>
		<TD width=60 nowrap align=right>處/所：</TD>
		<TD >
			<select name=unitno  class=font9 onchange=unitchg()>
				<option value="">-----</option>
				<%SQL="SELECT * FROM BASICCODE WHERE FUNC='unit' and sys_type<>'AAA' ORDER BY SYS_TYPE desc "
				SET RST = CONN.EXECUTE(SQL)
				WHILE NOT RST.EOF
				%>
				<option value="<%=RST("SYS_TYPE")%>" ><%=RST("SYS_VALUE")%></option>
				<%
				RST.MOVENEXT
				WEND
				%>
			</SELECT>
			<%SET RST=NOTHING %>
		</TD>
	</tr>
	<TR height=25 >
		<TD nowrap align=right >組/部門：</TD>
		<TD >
			<select name=GROUPID  class=font9 onchange=groupchg() style="width:60" >
				<option value="" <%if request("GROUPID")="" then%>selected<%end if%>>-----</option>
				<%SQL="SELECT * FROM BASICCODE WHERE FUNC='GROUPID' and sys_type like 'A06%' ORDER BY SYS_TYPE "
				SET RST = CONN.EXECUTE(SQL)
				WHILE NOT RST.EOF
				%>
				<option value="<%=RST("SYS_TYPE")%>" ><%=RST("SYS_VALUE")%></option>
				<%
				RST.MOVENEXT
				WEND
				%>
			</SELECT>
			<%SET RST=NOTHING %>
			<select name=zuno  class=font9 style='width:50' >
				<OPTION VALUE=""></OPTION>
			</SELECT>

		</TD>
		<TD nowrap align=right >職等：</TD>
		<TD >
			<select name=JOB  class=font9 style='width:75' >
				<%SQL="SELECT * FROM BASICCODE WHERE FUNC='LEV'  ORDER BY SYS_TYPE "
				SET RST = CONN.EXECUTE(SQL)
				WHILE NOT RST.EOF
				%>
				<option value="<%=RST("SYS_TYPE")%>" <%IF SESSION("NETWHSNO")=RST("SYS_TYPE") THEN %> SELECTED <%END IF%>><%=RST("SYS_VALUE")%></option>
				<%
				RST.MOVENEXT
				WEND
				%>
			</SELECT>
			<%SET RST=NOTHING %>
		</TD>
	</TR>
	<TR height=25 >
		<TD nowrap align=right>班別：</TD>
		<TD >
			<SELECT NAME=SHIFT CLASS=font9 onkeydown="enterto()" >
				<OPTION VALUE="" <%IF SHIFT="" THEN %> SELECTED <%END IF%>></OPTION>
				<OPTION VALUE="ALL" <%IF SHIFT="ALL" THEN %> SELECTED <%END IF%>>常日班</OPTION>
				<OPTION VALUE="A" <%IF SHIFT="A" THEN %> SELECTED <%END IF%>>A班</OPTION>
				<OPTION VALUE="B" <%IF SHIFT="B" THEN %> SELECTED <%END IF%>>B班</OPTION>
			</SELECT>		 
			<select name='grps'  class=font9 style='width:50' onkeydown="enterto()"  >
			<%SQL="SELECT * FROM BASICCODE WHERE FUNC='grps'  ORDER BY SYS_TYPE "
				SET RST = CONN.EXECUTE(SQL)
				WHILE NOT RST.EOF 
			%>
				<option value="<%=RST("SYS_TYPE")%>" <%IF RST("SYS_TYPE")=grps THEN %> SELECTED <%END IF%>><%=RST("SYS_VALUE")%></option>
			<%
				RST.MOVENEXT
				WEND
			%>
			</SELECT>
			<%SET RST=NOTHING %>
		</TD>
		<TD nowrap align=right>國籍：</TD>
		<TD nowrap>
			<select name=country  class=font9 style='width:75'  >
				<%SQL="SELECT * FROM BASICCODE WHERE FUNC='country'  ORDER BY SYS_type desc  "
				SET RST = CONN.EXECUTE(SQL)
				WHILE NOT RST.EOF
				%>
				<option value="<%=RST("SYS_TYPE")%>" <%IF SESSION("NETWHSNO")=RST("SYS_TYPE") THEN %> SELECTED <%END IF%>><%=RST("SYS_VALUE")%></option>
				<%
				RST.MOVENEXT
				WEND
				%>
			</SELECT>
			<%SET RST=NOTHING %>
		</TD>
	</TR>
	<TR>
		<TD nowrap align=right>員工姓名(中)：</TD>
		<TD COLSPAN=3>
			<INPUT NAME=nam_cn SIZE=20 CLASS=INPUTBOX  >
		</TD>
	</TR>
	<TR height=25 >
		<TD nowrap align=right >員工姓名(越)：</TD>
		<TD colspan=3><INPUT NAME=nam_vn SIZE=38 CLASS=INPUTBOX></TD>
	</TR>
	<TR height=25 >
		<TD nowrap align=right >出生日期：</TD>
		<TD colspan=4>
		<INPUT NAME=BYY SIZE=5 CLASS=INPUTBOX  MAXLENGTH=4 ONBLUR="CHKVALUE(1)" > 年
		<INPUT NAME=BMM SIZE=3 CLASS=INPUTBOX  MAXLENGTH=2 ONBLUR="CHKVALUE(2)" > 月
		<INPUT NAME=BDD SIZE=3 CLASS=INPUTBOX  MAXLENGTH=2 ONBLUR="CHKVALUE(3)" > 日&nbsp;&nbsp;
		年齡： <input name=ages size=5 class=inputbox ONBLUR="CHKVALUE(4)" >&nbsp; &nbsp;
		<INPUT type="radio" id=radio1 name=radio1 onclick=sexchg(0) > 男 &nbsp;
		<INPUT type="radio" id=radio1 name=radio1 onclick=sexchg(1)> 女
		<input type=hidden name=sexstr value="<%=sex%>" size=1>
		</TD>
	</TR>
</TABLE>
<!-------------------------------------------------------------------->
<TABLE WIDTH=500 CLASS=FONT9 BORDER=0>
	<tr>
		<td width=90 nowrap align=right height=25 >婚姻狀況：</td>
		<td >
			<INPUT type="radio" id=radio2 name=radio2 onclick=marrychg(0) > 已婚 &nbsp;
			<INPUT type="radio" id=radio2 name=radio2 onclick=marrychg(1)> 未婚
			<input type=hidden name=marryed value="" size=1>
		</td>
		<td width=80 nowrap align=right >教育程度：</td>
		<td ><input name=school size=15 class=inputbox ></td>
	</tr>
	<tr>
		<td width=90 nowrap align=right height=25 >身分証字號：</td>
		<td ><input name=personID size=22 class=inputbox ></td>
		<td width=80 nowrap align=right >簽合同日：</td>
		<td ><input name=BHDAT size=15 class=inputbox onblur="date_change(2)">(簽約日)</td>
	</tr>
	<tr>
		<td nowrap align=right height=25 >護照號碼：</td>
		<td><input name=PASSPORTNO size=22 class=inputbox ></td>
		<td nowrap align=right >(護)有效期：</td>
		<td ><input name=pduedate size=15 class=inputbox onblur="date_change(3)"></td>
	</tr>
	<tr>
		<td nowrap align=right height=25 >簽證號碼：</td>
		<td><input name=visano size=22 class=inputbox ></td>
		<td nowrap align=right >(簽)有效期：</td>
		<td ><input name=vduedate size=15 class=inputbox onblur="date_change(4)" ></td>
	</tr>
	<tr>
		<td nowrap align=right height=25 >聯絡電話：</td>
		<td><input name=phone size=22 class=inputbox ></td>
		<td nowrap align=right >手機：</td>
		<td ><input name=mobilephone size=15 class=inputbox ></td>
	</tr>
	<tr>
		<td nowrap align=right height=25 >銀行帳號：</td>
		<td colspan=3><input name=bankid size=55 class=inputbox ></td>
	</tr>
	<tr>
		<td nowrap align=right height=25 >聯絡地址：</td>
		<td colspan=3><input name=homeaddr size=55 class=inputbox ></td>
	</tr>
	<tr>
		<td nowrap align=right height=25 >E-MAIL：</td>
		<td><input name=email size=22 class=inputbox ></td>
		<td nowrap align=right >離職日：</td>
		<td ><input name=outdat size=15 class=inputbox  onblur="date_change(5)" ></td>
	</tr>
	<tr>
		<td nowrap align=right height=25 >其他說明：</td>
		<td ><input name=memo size=22 class=inputbox ></td>
		<td nowrap align=right height=25 >加入工團：</td>
		<td ><input name="GTDAT"  size=8 class=inputbox ONBLUR="CHKVALUE(5)" >(ex:200601)</td>
	</tr>
</table>
<hr size=0	style='border: 1px dotted #999999;' align=left width=500>
<TABLE WIDTH=460>
		<tr ALIGN=center>
			<TD >
			<input type=button  name=send value="確　　認"  class=button onclick=go()>
			<input type=RESET name=send value="取 　　消"  class=button>
			</TD>
		</TR>
</TABLE>


</form>


</body>
</html>
<script language=vbscript>
<!--
function empidchg()
	empidstr = Ucase(Trim(<%=self%>.empid.value))
	if empidstr<>"" then
		open "empfile.back.asp?ftype=empidchk&code="& empidstr , "Back"
		'parent.best.cols="50%,50%"
	end if
end function

function sexchg(x)
	if <%=self%>.radio1(0).checked=true then
		<%=self%>.sexstr.value="M"
	elseif 	<%=self%>.radio1(1).checked=true then
		<%=self%>.sexstr.value="F"
	else
		<%=self%>.sexstr.value=""
	end if
end function

function marrychg(x)
	if <%=self%>.radio2(0).checked=true then
		<%=self%>.marryed.value="Y"
	elseif 	<%=self%>.radio2(1).checked=true then
		<%=self%>.marryed.value="N"
	else
		<%=self%>.marryed.value=""
	end if
end function

function BACKMAIN()
	open "../main.asp" , "_self"
end function

FUNCTION GO()
	IF  <%=SELF%>.EMPID.VALUE="" THEN
		ALERT "請輸入員工編號!!"
		<%=SELF%>.EMPID.FOCUS()
		EXIT FUNCTION 
	END IF
	if <%=self%>.unitno.value="" then 
		ALERT "請輸入處/所!!"
		<%=SELF%>.unitno.FOCUS()
		EXIT FUNCTION 
	end if 
	if <%=self%>.GROUPID.value="" then 
		ALERT "請輸入部門單位!!"
		<%=SELF%>.GROUPID.FOCUS()
		EXIT FUNCTION 
	end if 
	if <%=self%>.shift.value="" then 
		ALERT "請輸入班別!!"
		<%=SELF%>.shift.FOCUS()
		EXIT FUNCTION 
	end if 
	
	<%=SELF%>.ACTION="empfile.upd.asp?act=EMPADDNEW"
	<%=SELF%>.SUBMIT
END FUNCTION

'*******檢查日期*********************************************
FUNCTION date_change(a)

if a=1 then
	INcardat = Trim(<%=self%>.indat.value)
elseif a=2 then
	INcardat = Trim(<%=self%>.BHDAT.value)
elseif a=3 then
	INcardat = Trim(<%=self%>.pduedate.value)
elseif a=4 then
	INcardat = Trim(<%=self%>.vduedate.value)
elseif a=5 then
	INcardat = Trim(<%=self%>.outdat.value)
end if

IF INcardat<>"" THEN
	ANS=validDate(INcardat)
	IF ANS <> "" THEN
		if a=1 then
			Document.<%=self%>.indat.value=ANS
		elseif a=2 then
			Document.<%=self%>.BHDAT.value=ANS
		elseif a=3 then
			Document.<%=self%>.pduedate.value=ANS
		elseif a=4 then
			Document.<%=self%>.vduedate.value=ANS
		elseif a=5 then
			Document.<%=self%>.outdat.value=ANS
		end if
	ELSE
		ALERT "EZ0067:輸入日期不合法 !!"
		if a=1 then
			Document.<%=self%>.indat.value=""
			Document.<%=self%>.indat.focus()
		elseif a=2 then
			Document.<%=self%>.BHDAT.value=""
			Document.<%=self%>.BHDAT.focus()
		elseif a=3 then
			Document.<%=self%>.pduedate.value=""
			Document.<%=self%>.pduedate.focus()
		elseif a=4 then
			Document.<%=self%>.vduedate.value=""
			Document.<%=self%>.vduedate.focus()
		elseif a=5 then
			Document.<%=self%>.outdat.value=""
			Document.<%=self%>.outdat.focus()
		end if
		EXIT FUNCTION
	END IF

ELSE
	'alert "EZ0015:日期該欄位必須輸入資料 !!"
	EXIT FUNCTION
END IF

END FUNCTION

'_________________DATE CHECK___________________________________________________________________

function validDate(d)
	if len(d) = 8 and isnumeric(d) then
		d = left(d,4) & "/" & mid(d, 5, 2) & "/" & right(d,2)
		if isdate(d) then
			validDate = formatDate(d)
		else
			validDate = ""
		end if
	elseif len(d) >= 8 and isdate(d) then
			validDate = formatDate(d)
		else
			validDate = ""
		end if
end function

function formatDate(d)
		formatDate = Year(d) & "/" & _
		Right("00" & Month(d), 2) & "/" & _
		Right("00" & Day(d), 2)
end function
'________________________________________________________________________________________

FUNCTION CHKVALUE(N)
IF N=1 THEN
	IF TRIM(<%=SELF%>.BYY.VALUE)<>"" THEN
		IF ISNUMERIC(<%=SELF%>.BYY.VALUE)=FALSE OR INSTR(1,<%=SELF%>.BYY.VALUE,"-")>0 THEN
			ALERT "輸入錯誤!!請輸入正確數字"
			<%=SELF%>.BYY.VALUE=""
			<%=SELF%>.BYY.FOCUS()
			EXIT FUNCTION
		ELSE
			<%=SELF%>.AGES.VALUE=CDBL(YEAR(DATE()))-CDBL(<%=SELF%>.BYY.VALUE) + 1
		END IF
	END IF
ELSEIF N=2 THEN
	IF TRIM(<%=SELF%>.BMM.VALUE)<>"" THEN
		IF ISNUMERIC(<%=SELF%>.BMM.VALUE)=FALSE OR INSTR(1,<%=SELF%>.BMM.VALUE,"-")>0 THEN
			ALERT "輸入錯誤!!請輸入正確數字"
			<%=SELF%>.BMM.VALUE=""
			<%=SELF%>.BMM.FOCUS()
			EXIT FUNCTION
		END IF
	END IF
ELSEIF N=3 THEN
	IF TRIM(<%=SELF%>.BDD.VALUE)<>"" THEN
		IF ISNUMERIC(<%=SELF%>.BDD.VALUE)=FALSE OR INSTR(1,<%=SELF%>.BDD.VALUE,"-")>0 THEN
			ALERT "輸入錯誤!!請輸入正確數字"
			<%=SELF%>.BDD.VALUE=""
			<%=SELF%>.BDD.FOCUS()
			EXIT FUNCTION
		END IF
	END IF
ELSEIF N=4 THEN
	IF TRIM(<%=SELF%>.AGES.VALUE)<>"" THEN
		IF ISNUMERIC(<%=SELF%>.AGES.VALUE)=FALSE OR INSTR(1,<%=SELF%>.AGES.VALUE,"-")>0 THEN
			ALERT "輸入錯誤!!請輸入正確數字"
			<%=SELF%>.AGES.VALUE=""
			<%=SELF%>.AGES.FOCUS()
			EXIT FUNCTION
		END IF
	END IF
ELSEIF N=5 THEN
	IF TRIM(<%=SELF%>.GTDAT.VALUE)<>"" THEN
		IF ISNUMERIC(<%=SELF%>.GTDAT.VALUE)=FALSE OR INSTR(1,<%=SELF%>.GTDAT.VALUE,"-")>0 THEN
			ALERT "輸入錯誤!!請輸入正確數字"
			<%=SELF%>.GTDAT.VALUE=""
			<%=SELF%>.GTDAT.FOCUS()
			EXIT FUNCTION
		END IF
	END IF
END IF

END FUNCTION
-->
</script>


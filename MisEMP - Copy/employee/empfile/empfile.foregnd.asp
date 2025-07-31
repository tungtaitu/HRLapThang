<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<!-- #include file = "../../GetSQLServerConnection.fun" -->
<!-- #include file="../../ADOINC.inc" -->
<%
SESSION.CODEPAGE="65001"
SELF = "empfilefore"

Set conn = GetSQLServerConnection()
Set rs = Server.CreateObject("ADODB.Recordset")

firstday = year(date())&"/"&right("00"&month(date()),2)&"/01"
nowmonth = year(date())&right("00"&month(date()),2)
if month(date())="01" then
	calcmonth = year(date()-1)&"12"
else
	calcmonth =  year(date())&right("00"&month(date())-1,2)
end if
'EMPID = TRIM(REQUEST("EMPID"))
'IF EMPID="" THEN EMPID="L0051"
empautoid = TRIM(REQUEST("empautoid"))
totalpage = request("totalpage")
currentpage = request("currentpage")
RecordInDB = request("RecordInDB")

IF REQUEST("NEXTID")<>"" THEN
	SQL="SELECT TOP 1 * FROM  view_empfile WHERE ISNULL(STATUS,'')<>'D' AND autoid >'"& REQUEST("NEXTID") &"' ORDER BY AUTOID "
ELSEIF REQUEST("BACKID")<>"" THEN
	SQL="SELECT TOP 1 * FROM  view_empfile WHERE ISNULL(STATUS,'')<>'D' AND  autoid <'"& REQUEST("BACKID") &"' ORDER BY AUTOID DESC  "
ELSE
	SQL="SELECT * FROM    view_empfile where ISNULL(STATUS,'')<>'D' AND  autoid='"& empautoid &"' "
END IF
'RESPONSE.WRITE SQL
'RESPONSE.END
RS.OPEN SQL , CONN, 3, 3
IF NOT RS.EOF THEN
	empautoid = TRIM(RS("AUTOID"))
	EMPID=TRIM(RS("EMPID"))	'員工編號
	INDAT=TRIM(RS("nindat"))	'到職日
	WHSNO=TRIM(RS("WHSNO"))	'廠別
	UNITNO=TRIM(RS("UNITNO"))	'處/所
	GROUPID=TRIM(RS("GROUPID"))	'組/部門
	ZUNO=TRIM(RS("ZUNO"))	'單位

	EMPNAM_CN=TRIM(RS("EMPNAM_CN"))	'姓名(中)
	EMPNAM_VN=TRIM(RS("EMPNAM_VN"))	'姓名(越)
	COUNTRY=TRIM(RS("COUNTRY"))	'國籍
	BYY=(TRIM(RS("BYY"))) '年(生日)
	BMM=(RS("BMM"))	'月(生日)
	BDD=(RS("BDD"))	'日(生日)
	AGES=TRIM(RS("AGES"))	'年齡
	SEX=TRIM(RS("SEX"))	'性別
	JOB=TRIM(RS("JOB"))  '職等 
	Jstr=trim(rs("jstr"))
	Gstr=trim(rs("gstr"))
	zstr=trim(rs("zstr"))
	wstr=trim(rs("wstr"))
	ustr=trim(rs("ustr"))
	PERSONID=TRIM(RS("PERSONID"))	'身分証字號
	BHDAT=TRIM(RS("BHDAT"))	'簽約日(保險日)
	PASSPORTNO=TRIM(RS("PASSPORTNO"))	'護照號碼
	VISANO=TRIM(RS("VISANO"))	'簽證號碼
	PDUEDATE=TRIM(RS("PissueDATE"))	'護照簽發日
	PDUEDATE=TRIM(RS("PDUEDATE"))	'護照有效期
	VDUEDATE=TRIM(RS("VDUEDATE"))	'簽證有效期
	PHONE=TRIM(RS("PHONE"))	'聯絡電話
	MOBILEPHONE=TRIM(RS("MOBILEPHONE"))	'手機
	HOMEADDR=TRIM(RS("HOMEADDR"))	'聯絡地址
	EMAIL=TRIM(RS("EMAIL"))	'EMAIL
	OUTDAT=TRIM(RS("OUTDAT"))	'離職日
	MEMO=TRIM(RS("MEMO"))	'其他說明
	GTDAT=TRIM(RS("GTDAT"))	'加入工團(年月)
	MARRYED = RS("MARRYED")    '婚姻狀況
	SCHOOL=RS("SCHOOL") '教育程度
	SHIFT=RS("SHIFT") '班別
	grps = rs("grps") 
	studyjob=RS("studyjob") '職能學習 
	
	'PHOTOS=TRIM(RS("PHOTOS"))	'照片檔名 
	PHOTOS=RS("EMPID")&".JPG"
	'-----------------------------------------
	'PHU = RS("PHU")
	'NN = RS("NN")
	'KT = RS("KT")
	'TTKH = RS("TTKH")
	'MT = RS("MT")
	'BB = RS("BB")
	'BANKID = RS("BANKID") 
	
	wkd_no = RS("wkd_no") 
	wkd_duedate = RS("wkd_duedate") 
ELSE
%>	<SCRIPT LANGUAGE=VBSCRIPT>
		ALERT "The final Data , No records!!"
		OPEN "EMPFILE.EDIT.ASP", "_self"
	</SCRIPT>
<%	response.end
END IF
SET RS=NOTHING

FUNCTION FDT(D)
IF D <> "" THEN
	Response.Write YEAR(D)&"/"&RIGHT("00"&MONTH(D),2)&"/"&RIGHT("00"&DAY(D),2)
END IF
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
	<%=self%>.INDAT.focus()
	<%=self%>.INDAT.select()
end function

function groupchg()
	code = <%=self%>.GROUPID.value
	open "empfile.backgnd.asp?ftype=groupchg&code="&code , "Back"
	'parent.best.cols="50%,50%"
end function

function unitchg()
	code = <%=self%>.unitno.value
	open "empfile.backgnd.asp?ftype=UNITCHG&code="&code , "Back"
	'parent.best.cols="50%,50%"
end function
-->
</SCRIPT>
</head>
<body  topmargin="5" leftmargin="5"  marginwidth="0" marginheight="0"   onload="f()" onkeydown="enterto()" >
<form name="<%=self%>"  method="post"    >
<INPUT TYPE=HIDDEN NAME="UID" VALUE=<%=SESSION("NETUSER")%>>
<INPUT TYPE=HIDDEN NAME="empautoid" VALUE=<%=empautoid%>>
<INPUT TYPE=hidden NAME=TotalPage VALUE="<%=TotalPage%>">
<INPUT TYPE=hidden NAME=CurrentPage VALUE="<%=CurrentPage%>">
<INPUT TYPE=hidden NAME=RecordInDB VALUE="<%=RecordInDB%>">
<INPUT TYPE=hidden NAME=nowmonth VALUE="<%=nowmonth%>">
<table width="460" border="0" cellspacing="0" cellpadding="0">
	<tr><TD>
	<img border="0" src="../../image/icon.gif" align="absmiddle">
	人事薪資系統( 員工基本資料-修改 ) </TD>
	</tr>
</table>
<hr size=0	style='border: 1px dotted #999999;' align=left width=500>
<TABLE WIDTH=500 CLASS=txt BORDER=0>
	<TR height=25 >
		<TD width=80 nowrap align=right height=25>員工編號：</TD>
		<TD ><INPUT NAME=EMPID SIZE=12 CLASS=READONLY VALUE="<%=EMPID%>" READONLY   > </TD>
		<TD width=60 nowrap align=right height=25>到職日：</TD>
		<TD ><INPUT NAME=indat SIZE=12 CLASS=INPUTBOX VALUE="<%=(indat)%>" onblur="date_change(1)"  ></TD>
		<td width="120" rowspan="5" align=center valign=top height=130 nowrap ><img src="../photos/<%=EMPID%>.jpg"  border=1></td>
	</TR>
	<TR height=25 >
		<TD nowrap align=right>廠別：</TD>
		<TD >
			<%if cdate(indat)>=cdate(firstday) then %>
				<select name=WHSNO  class=font9 onkeydown="enterto()"  >
					<%SQL="SELECT * FROM BASICCODE WHERE FUNC='WHSNO' ORDER BY SYS_TYPE "
					SET RST = CONN.EXECUTE(SQL)
					WHILE NOT RST.EOF
					%>
					<option value="<%=RST("SYS_TYPE")%>" <%IF RST("SYS_TYPE")=WHSNO THEN %> SELECTED <%END IF%>><%=RST("SYS_VALUE")%></option>
					<%
					RST.MOVENEXT
					WEND
					%>
				</SELECT>
				<%SET RST=NOTHING %>
			<%else%>	
				<input type=hidden name=whsno size=10 class='readonly' readonly value="<%=whsno%>">
				<input name=wstr size=15 class='readonly' readonly value="<%=wstr%>">
			<%end if%>
		</TD>
		<TD width=60 nowrap align=right     >處/所：</TD>
		<TD >
			<%if cdate(indat)>=cdate(firstday) then %>
				<select name=unitno  class=font9 onchange=unitchg() onkeydown="enterto()"  >
					<%SQL="SELECT * FROM BASICCODE WHERE FUNC='unit' and sys_type<>'AAA' ORDER BY SYS_TYPE "
					SET RST = CONN.EXECUTE(SQL)
					WHILE NOT RST.EOF
					%>
					<option value="<%=RST("SYS_TYPE")%>" <%IF RST("SYS_TYPE")=unitno THEN %> SELECTED <%END IF%>><%=RST("SYS_VALUE")%></option>
					<%
					RST.MOVENEXT
					WEND
					%>
				</SELECT>
				<%SET RST=NOTHING %>
			<%else%>
				<input type=hidden name=unitno size=10 class='readonly' readonly value="<%=unitno%>">
				<input name=ustr size=10 class='readonly' readonly value="<%=ustr%>">	
			<%end if%>
		</TD>
	</tr>
	<TR height=25 >
		<TD nowrap align=right   >組/部門：</TD>
		<TD >
			<%if cdate(indat)>=cdate(firstday) then %>
				<select name=GROUPID  class=font9 onchange=groupchg() style="width:60" onkeydown="enterto()"  >
					<%SQL="SELECT * FROM BASICCODE WHERE FUNC='GROUPID' and sys_type <>'AAA' ORDER BY SYS_TYPE "
					SET RST = CONN.EXECUTE(SQL)
					'RESPONSE.WRITE SQL
					WHILE NOT RST.EOF
					%>
					<option value="<%=RST("SYS_TYPE")%>" <%IF RST("SYS_TYPE")=TRIM(GROUPID) THEN %> SELECTED <%END IF%>><%=RST("SYS_VALUE")%></option>
					<%
					RST.MOVENEXT
					WEND
					%>
				</SELECT>
				<%SET RST=NOTHING %>
				<select name=zuno  class=font9 style='width:50' onkeydown="enterto()"  >
					<%
					SQL="SELECT * FROM BASICCODE WHERE FUNC='ZUNO' AND LEFT(SYS_TYPE,4)='"& GROUPID &"' ORDER BY SYS_TYPE "
					SET RST = CONN.EXECUTE(SQL)
					RESPONSE.WRITE ZUNO
					WHILE NOT RST.EOF
					%>
					<option value="<%=RST("SYS_TYPE")%>" <%IF RST("SYS_TYPE")=TRIM(ZUNO) THEN %> SELECTED <%END IF%>><%=RST("SYS_VALUE")%></option>
					<%
					RST.MOVENEXT
					WEND
					%>
				</SELECT>
				<%SET RST=NOTHING %>
			<%else%>	
				<input type=hidden name=groupid size=10 class='readonly' readonly value="<%=groupid%>">
				<input type=hidden name=zuno size=10 class='readonly' readonly value="<%=zuno%>">
				<input name=gstr size=5 class='readonly' readonly value="<%=gstr%>">
				<input name=zstr size=5 class='readonly' readonly value="<%=zstr%>">
			<%end if%>
		</TD>
		<TD nowrap align=right >職等：</TD>
		<TD >
		<%if cdate(indat)>=cdate(firstday) then %>
				<select name=JOB  class=font9 style='width:75' onkeydown="enterto()"  >
				<%SQL="SELECT * FROM BASICCODE WHERE FUNC='LEV'  ORDER BY SYS_TYPE "
					SET RST = CONN.EXECUTE(SQL)
					WHILE NOT RST.EOF
				%>
					<option value="<%=RST("SYS_TYPE")%>" <%IF RST("SYS_TYPE")=JOB THEN %> SELECTED <%END IF%>><%=RST("SYS_VALUE")%></option>
				<%
					RST.MOVENEXT
					WEND
				%>
				</SELECT>
				<%SET RST=NOTHING %>
			<%else%>	
				<input type=hidden name=job size=10 class='readonly' readonly value="<%=Job%>">
				<input name=jstr size=18 class='readonly' readonly value="<%=Jstr%>">
			<%end if%>
		</TD>
	</TR>
	<TR height=25 >
		<TD nowrap align=right>班別：</TD>
		<TD >
			<%if cdate(indat)>=cdate(firstday) then %>
				<SELECT NAME=SHIFT CLASS=font9 onkeydown="enterto()" >
					<OPTION VALUE="" <%IF SHIFT="" THEN %> SELECTED <%END IF%>></OPTION>
					<OPTION VALUE="ALL" <%IF SHIFT="ALL" THEN %> SELECTED <%END IF%>>常日班</OPTION>
					<OPTION VALUE="A" <%IF SHIFT="A" THEN %> SELECTED <%END IF%>>A班</OPTION>
					<OPTION VALUE="B" <%IF SHIFT="B" THEN %> SELECTED <%END IF%>>B班</OPTION>
				</SELECT>
			<%else
				IF shift="ALL" then 
					sstr = "常日班"
				elseif shift="A"	then 
					sstr="A班"
				elseif shift="B" then 
					sstr="B班"
				end if 		
			%>
				<input name=sstr size=4 class='readonly' readonly value="<%=sstr%>" >
				<input type=hidden name=shift size=10 class='readonly' readonly value="<%=shift%>">					
			<%end if%>
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
		<TD >
			<select name=country  class=font9 style='width:75' onkeydown="enterto()"  >
				<%SQL="SELECT * FROM BASICCODE WHERE FUNC='country' and sys_type='VN' ORDER BY SYS_type desc  "
				SET RST = CONN.EXECUTE(SQL)
				WHILE NOT RST.EOF
				%>
				<option value="<%=RST("SYS_TYPE")%>" <%IF RST("SYS_TYPE")=COUNTRY THEN %> SELECTED <%END IF%>><%=RST("SYS_VALUE")%></option>
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
			<INPUT NAME=nam_cn SIZE=20 CLASS=INPUTBOX VALUE="<%=EMPNAM_CN%>" onkeydown="enterto()" >
		</TD>
	</TR>
	<TR height=25 >
		<TD nowrap align=right >員工姓名(越)：</TD>
		<TD colspan=3><INPUT NAME=nam_vn SIZE=38 CLASS=INPUTBOX VALUE="<%=EMPNAM_VN%>" onkeydown="enterto()" ></TD>
	</TR>
	<TR height=25 >
		<TD nowrap align=right >出生日期：</TD>
		<TD colspan=4>
		<INPUT NAME=BYY SIZE=5 CLASS=INPUTBOX VALUE="<%=BYY%>" MAXLENGTH=4 ONBLUR="CHKVALUE(1)" onkeydown="enterto()" > 年
		<INPUT NAME=BMM SIZE=3 CLASS=INPUTBOX VALUE="<%=BMM%>" MAXLENGTH=2 ONBLUR="CHKVALUE(2)" onkeydown="enterto()" > 月
		<INPUT NAME=BDD SIZE=3 CLASS=INPUTBOX VALUE="<%=BDD%>" MAXLENGTH=2 ONBLUR="CHKVALUE(3)" onkeydown="enterto()" > 日&nbsp;&nbsp;
		年齡： <input name=ages size=5 class=inputbox VALUE="<%=AGES%>" ONBLUR="CHKVALUE(4)" onkeydown="enterto()" >&nbsp; &nbsp;
		<INPUT type="radio" id=radio1 name=radio1 <%IF SEX="M" THEN %>CHECKED<%END IF%> onclick=sexchg(0) onkeydown="enterto()" > 男 &nbsp;
		<INPUT type="radio" id=radio1 name=radio1 <%IF SEX="F" THEN %>CHECKED<%END IF%> onclick=sexchg(1) onkeydown="enterto()" > 女
		<input type=hidden name=sexstr value="<%=sex%>" size=1>
		</TD>
	</TR>
</TABLE>
<!-------------------------------------------------------------------->
<TABLE WIDTH=500 CLASS=FONT9 BORDER=0>
	<tr>
		<td width=90 nowrap align=right height=25 >婚姻狀況：</td>
		<td >
			<INPUT type="radio" id=radio2 <%IF MARRYED="Y" THEN %>CHECKED<%END IF%> name=radio2 onclick=marrychg(0) onkeydown="enterto()"  > 已婚 &nbsp;
			<INPUT type="radio" id=radio2 <%IF MARRYED="N" THEN %>CHECKED<%END IF%> name=radio2 onclick=marrychg(1) onkeydown="enterto()" > 未婚
			<input type=hidden name=marryed value="<%=marryed%>" size=1>
		</td>
		<td width=80 nowrap align=right >教育程度：</td>
		<td ><input name=school size=15 class=inputbox VALUE="<%=SCHOOL%>" onkeydown="enterto()" ></td>
	</tr>
	<tr>
		<td width=90 nowrap align=right height=25  >身分証字號：</td>
		<td ><input name=personID size=20 class=inputbox VALUE="<%=PERSONID%>" onkeydown="enterto()" ></td>
		<td width=80 nowrap align=right >簽合同日：</td>
		<td ><input name=BHDAT size=15 class=inputbox VALUE="<%=(BHDAT)%>" onblur="date_change(2)" onkeydown="enterto()" >(簽約日)</td>
	</tr>
	<tr>
		<td nowrap align=right height=25 >護照號碼：</td>
		<td><input name=PASSPORTNO size=20 class=inputbox VALUE="<%=PASSPORTNO%>" onkeydown="enterto()" ></td>
		<td nowrap align=right >(護)有效期：</td>
		<td ><input name=pduedate size=15 class=inputbox VALUE="<%=(PDUEDATE)%>" onblur="date_change(3)" onkeydown="enterto()" ></td>
	</tr>
	<tr>
		<td nowrap align=right height=25 >簽證號碼：</td>
		<td><input name=visano size=20 class=inputbox VALUE="<%=VISANO%>" onkeydown="enterto()" ></td>
		<td nowrap align=right >(簽)有效期：</td>
		<td ><input name=vduedate size=15 class=inputbox VALUE="<%=(VDUEDATE)%>" onblur="date_change(4)" onkeydown="enterto()" ></td>
	</tr>
	<tr>
		<td nowrap align=right height=25 >聯絡電話：</td>
		<td><input name=phone size=20 class=inputbox VALUE="<%=PHONE%>" onkeydown="enterto()" ></td>
		<td nowrap align=right >手機：</td>
		<td ><input name=mobilephone size=15 class=inputbox VALUE="<%=MOBILEPHONE%>" onkeydown="enterto()" ></td>
	</tr>
	<tr>
		<td nowrap align=right height=25 >銀行帳號：</td>
		<td colspan=3><input name=BANKID size=54 class=inputbox VALUE="<%=BANKID%>" onkeydown="enterto()" ></td>
	</tr>
	<tr>
		<td nowrap align=right height=25 >聯絡地址：</td>
		<td colspan=3><input name=homeaddr size=54 class=inputbox8 VALUE="<%=HOMEADDR%>" onkeydown="enterto()" ></td>
	</tr>
	<tr>
		<td nowrap align=right height=25 >E-MAIL：</td>
		<td><input name=email size=25 class=inputbox8 VALUE="<%=EMAIL%>" onkeydown="enterto()" ></td>
		<td nowrap align=right >離職日：</td>
		<td ><input name=outdat size=15 class=inputbox VALUE="<%=(OUTDAT)%>" onblur="date_change(5)" onkeydown="enterto()" ></td>
	</tr>
	<tr>
		<td nowrap align=right height=25 >其他說明：</td>
		<td ><input name=memo size=22 class=inputbox VALUE="<%=MEMO%>" onkeydown="enterto()" ></td>
		<td nowrap align=right height=25 >加入工團：</td>
		<td ><input name="GTDAT"  size=8 class=inputbox  VALUE="<%=GTDAT%>"  ONBLUR="CHKVALUE(5)" onkeydown="enterto()"  >(ex:200601)</td>
	</tr>
	<tr>
		<td nowrap align=right height=25 >職能學習：</td>
		<td colspan=3>
			<textarea rows="3" name="studyjob" cols="40"><%=studyjob%></textarea>
		
		</td>
	</tr> 
</table>
<hr size=0	style='border: 1px dotted #999999;' align=left width=500>
<TABLE WIDTH=500>
	<tr ALIGN=center>
		<TD colspan=4>
		<!--input type=BUTTON name=BTN value=" BACK "  class=button ONCLICK="BACKIDCHG()">
		<input type=BUTTON name=BTN value=" NEXT "  class=button ONCLICK="NEXTIDCHG()"-->
		<input type=BUTTON name=send value="關閉(CLOSE)"  class=button ONCLICK="window.close()">　　
		<input type="button" name=send value="確認修改"  class=button onclick="go()" onkeydown="go()" >
		<input type=RESET name=send value="取　　消"  class=button>
		<input type="button" name=send value="刪除(DEL)"  class=button onclick="godel()" onkeydown="godel()" >
		</TD>
	</TR>
</TABLE>

</form>


</body>
</html>

<script language=vbscript >
function BACKMAIN()
	open "../main.asp" , "_self"
end function

function hback()
	'alert <%=currentpage%>
	'open "empfile.edit.asp?send=NEXT&totalpage=" & <%=totalpage%> &"&currentpage=" & <%=currentpage-1%> &"&RecordInDB=" & <%=RecordInDB%>  , "_self"
end function

FUNCTION GO()
	'EMPIDSTR=<%=SELF%>.EMPID.VALUE
	<%=SELF%>.ACTION="empfile.upd.asp?ACT=EMPEDIT"
	<%=SELF%>.SUBMIT
END FUNCTION

FUNCTION GODEL()
	'EMPIDSTR=<%=SELF%>.EMPID.VALUE
	IF CONFIRM("確定要刪除(DELETE)這位員工資料!!"&chr(13)&"Xoi di khong??",64) THEN
		<%=SELF%>.ACTION="empfile.upd.asp?ACT=EMPDEL"
		<%=SELF%>.SUBMIT
	END IF
END FUNCTION

FUNCTION NEXTIDCHG()
	<%=SELF%>.ACTION="EMPFILE.FOREGND.ASP?NEXTID=<%=empautoid%>"
	<%=SELF%>.SUBMIT()
END FUNCTION

FUNCTION BACKIDCHG()
	<%=SELF%>.ACTION="EMPFILE.FOREGND.ASP?BACKID=<%=empautoid%>"
	<%=SELF%>.SUBMIT()
END FUNCTION

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
		formatDate = Year(d) & "/"  & Right("00" & Month(d), 2) & "/" & Right("00" & Day(d), 2)
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

</script>


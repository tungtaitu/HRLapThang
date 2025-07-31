<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<!-- #include file = "../GetSQLServerConnection.fun" -->
<!-- #include file="../ADOINC.inc" -->
<%

SELF = "YEBE0301B"

Set conn = GetSQLServerConnection()
Set rs = Server.CreateObject("ADODB.Recordset")

firstday = year(date())&"/"&right("00"&month(date()),2)&"/01"
nowmonth = year(date())&right("00"&month(date()),2)
if month(date())="01" then
	calcmonth = year(date()-1)&"12"
else
	calcmonth =  year(date())&right("00"&month(date())-1,2)
end if

empautoid = TRIM(REQUEST("empautoid")) 
empid=request("empid")
Set fso = Server.CreateObject("Scripting.FileSystemObject")
SQL="SELECT * FROM  fn_empfile_b('"&empid&"') "

'RESPONSE.WRITE SQL
'RESPONSE.END
RS.OPEN SQL , CONN, 3, 3
IF NOT RS.EOF THEN
	empautoid = TRIM(RS("AUTOID"))
	emptype = TRIM(RS("emptype"))
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
	PISSUEDATE=TRIM(RS("PISSUEDATE")) '護照簽發日
	PDUEDATE=TRIM(RS("PDUEDATE"))	'護照有效期
	VDUEDATE=TRIM(RS("VDUEDATE"))	'簽證有效期
	PHONE=TRIM(RS("PHONE"))	'聯絡電話
	MOBILEPHONE=TRIM(RS("MOBILEPHONE"))	'手機
	HOMEADDR=TRIM(RS("HOMEADDR"))	'聯絡地址
	EMAIL=TRIM(RS("EMAIL"))	'EMAIL
	OUTDATe=TRIM(RS("OUTDATe"))	'離職日
	MEMO=TRIM(RS("MEMO"))	'其他說明
	GTDAT=TRIM(RS("GTDAT"))	'加入工團(年月)
	MARRYED = trim(RS("MARRYED"))    '婚姻狀況
	SCHOOL=RS("SCHOOL") '教育程度
	SHIFT=RS("SHIFT") '班別
	grps = rs("grps")
	studyjob=RS("studyjob") '職能學習
	
	taxcode = rs("taxCOde")
	'response.write grps 
	'PHOTOS=TRIM(RS("PHOTOS"))	'照片檔名
	'PHOTOS=RS("EMPID")&".JPG"
	'-----------------------------------------
	'PHU = RS("PHU")
	'NN = RS("NN")
	'KT = RS("KT")
	'TTKH = RS("TTKH")
	'MT = RS("MT")
	'BB = RS("BB")
	BANKID = RS("BANKID")

	wkd_no = RS("wkd_no")
	wkd_duedate = RS("wkd_duedate")
	experience = RS("experience")
	urgent_person = RS("urgent_person")
	releation = RS("releation")
	urgent_addr = RS("urgent_addr")
	urgent_tel = RS("urgent_tel")
	urgent_mobile=RS("urgent_mobile")
	bh_person=RS("bh_person")
	bh_personID=RS("bh_personID")
	
	sobh=RS("sobh")

	filename=Server.MapPath("pic/"&rs("empid")&".jpg")
	If fso.FileExists(filename) and isnull(RS("PHOTOS"))=false Then
		photoYN="Y"
	else
		photoYN="N"
	end if


END IF
SET RS=NOTHING



FUNCTION FDT(D)
	Response.Write YEAR(D)&"/"&RIGHT("00"&MONTH(D),2)&"/"&RIGHT("00"&DAY(D),2)
END FUNCTION
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
	'<%=self%>.indat.focus()
	'<%=self%>.indat.select()
	<%=self%>.send(0).style.display=""
	<%=self%>.send(1).style.display=""
	'<%=self%>.send(2).style.display=""
	'<%=self%>.emptype.focus()
end function

function groupchg()
	code = <%=self%>.GROUPID.value
	'<%=self%>.action="yebe0301.editvn.asp"
	'<%=self%>.submit()
	open "<%=self%>.back.asp?func=gchg&codestr01="&code , "Back"
	'parent.best.cols="50%,50%"
end function

function unitchg()
	code = <%=self%>.unitno.value
	open "<%=self%>.back.asp?ftype=UNITCHG&codestr01="&code , "Back"
	'parent.best.cols="50%,50%"
end function
-->
</SCRIPT>
</HEAD>
<body  topmargin="5" leftmargin="2"  marginwidth="0" marginheight="0"  onkeydown="enterto()" onload='f()' >
<form  name="<%=self%>" method="post" action="YEBE0102.upd.asp"   >
<INPUT TYPE=HIDDEN NAME="UID" VALUE=<%=SESSION("NETUSER")%>>
<INPUT TYPE=hidden NAME=nowmonth VALUE="<%=nowmonth%>">
<input name=act value="EMPEDIT" type=hidden >
<input name="old_whsno" value="<%=whsno%>" type=hidden >
<input name="old_grp" value="<%=groupid%>" type=hidden >
<input name="old_zuno" value="<%=zuno%>" type=hidden >
<input name="old_shift" value="<%=shift%>" type=hidden >
<input name="old_job" value="<%=job%>" type=hidden >

<table width="460" border="0" cellspacing="0" cellpadding="0">
	<tr><TD  >
		<img border="0" src="../image/icon.gif" align="absmiddle">
		越籍員工個人資料( TƯ LIỆU NHÂN VIÊN ) 
		</TD>
	</tr>
</table>
<hr size=0	style='border: 1px dotted #999999;' align=left width="100%">
<TABLE WIDTH="680"  CLASS=txt BORDER="0"  cellspacing="1" cellpadding="2">
	<TR height=25 >
		<TD width=80 nowrap align=right >工號<br><font class=txt8>Số thẻ</font></TD>
		<TD  >
			<INPUT NAME=EMPID SIZE=6 CLASS=READONLY VALUE="<%=EMPID%>" READONLY   >
			<%if left(empid,1)="S" then %>
				<select name=emptype class=txt>
					<option value="C" <%if emptype="C" then%>selected<%end if%>>C.代銷</option>
				</select>
			<%else%>
				<select name=emptype class=txt>
					<option value=""></option>
					<option value="A" <%if emptype="A" then%>selected<%end if%>>A.員工</option>
					<option value="B" <%if emptype="B" then%>selected<%end if%>>B.出差</option>
				</select>
			<%end if%> 
			<INPUT type=hidden NAME=empautoid   VALUE="<%=empautoid%>"     >
		</TD>
		<TD  width="100"  align=right >到職日<br><font class=txt8>NVX</font></TD>
		<TD   ><INPUT NAME=indat SIZE=15 CLASS=INPUTBOX VALUE="<%=(indat)%>" onblur="date_change(1)"  ></TD>
		<td  align=center valign=top  nowrap  rowspan="5" width="110"  >
			<%if photoYN="Y" then%>
				<img src="pic/<%=EMPID%>.jpg"  border=1 width=100 height=130>
			<%else%>
				<a href="vbscript:upphots()"><img src="pic/noimg.gif"  border=0> </a>
			<%end if%>
		</td>
	</TR>
	<TR height=25 >
		<TD   align=right>國籍<br><font class=txt8>Quốc tịch</font></TD>
		<TD >
			<select name=country  class=inputbox  onkeydown="enterto()">
				<%SQL="SELECT * FROM BASICCODE WHERE FUNC='country'  ORDER BY SYS_type desc  "
				SET RST = CONN.EXECUTE(SQL)
				WHILE NOT RST.EOF
				%>
				<option value="<%=RST("SYS_TYPE")%>" <%IF RST("SYS_TYPE")=COUNTRY THEN %> SELECTED <%END IF%>><%=RST("SYS_TYPE")%>-<%=RST("SYS_VALUE")%></option>
				<%
				RST.MOVENEXT
				WEND
				%>
			</SELECT>
			<%SET RST=NOTHING %>
		</TD>
		<td align=right >離職日<br><font class=txt8>Ngày nghỉ việc</font></td>
		<td ><input name=outdat size=15 class=INPUTBOX    onblur="date_change(5)"  value=<%=OUTDATe%>  ></td>
	</tr>
	
	<TR   >
		<TD   align=right>姓名(中)<br><font class=txt8>Họ tên(Hoa)</font></TD>
		<TD  >
			<INPUT NAME=nam_cn SIZE=20 CLASS=INPUTBOX VALUE="<%=EMPNAM_CN%>" onkeydown="enterto()" >
		</TD>
		<TD   align=right >姓名(越)<br><font class=txt8>Họ tên(Việt)</font></TD>
		<TD  ><INPUT NAME=nam_vn SIZE=35 CLASS=INPUTBOX VALUE="<%=EMPNAM_VN%>" onkeydown="enterto()" ></td>
	</TR>
	<TR   >
		<TD   align=right >出生日期<br><font class=txt8>Ngày sinh</font></TD>		
		<TD colspan=3>
			<table border="0" cellpadding="0" cellspacing="0" width="100%" class="txt8">
			<tr>
			<td align="left"  >
			<INPUT NAME=BYY SIZE=4 CLASS=INPUTBOXr VALUE="<%=BYY%>" MAXLENGTH=4 ONBLUR="CHKVALUE(1)" onkeydown="enterto()"  style="certical-align:middle">
			年Năm&nbsp;
			<INPUT NAME=BMM SIZE=3 CLASS=INPUTBOXr VALUE="<%=BMM%>" MAXLENGTH=2 ONBLUR="CHKVALUE(2)" onkeydown="enterto()" style="certical-align:middle">
			月Tháng&nbsp;
			<INPUT NAME=BDD SIZE=3 CLASS=INPUTBOXr VALUE="<%=BDD%>" MAXLENGTH=2 ONBLUR="CHKVALUE(3)" onkeydown="enterto()" style="certical-align:middle">
			日Ngày
			<input type=hidden name=ages size=5 class=inputbox VALUE="<%=AGES%>"  >
			</td>
			<td   align=right  width="60">性別<br><font class=txt8>GiỚi Tính</font></td>
			<td  >
			<INPUT type="radio" id=radio1 name=radio1 onclick=sexchg(0) <%IF SEX="M" THEN %>CHECKED<%END IF%>  style="certical-align:middle" > 男 Nam &nbsp;
			<INPUT type="radio" id=radio1 name=radio1 onclick=sexchg(1) <%IF SEX="F" THEN %>CHECKED<%END IF%> style="certical-align:middle"> 女 Nữ
			<input type=hidden name=sexstr value="<%=SEX%>" size=1>	
			</td>	
			</tr>
			</table>
		</TD>
	</TR>	
 	<tr>
		<Td align=right>婚姻<br><font class=txt8>Tình trạng Hôn Nhân</font></td>
		<TD  colspan=3  >
			<table border="0" cellpadding="1" cellspacing="0" width="100%" class="txt8">
			<tr><td  >
			<INPUT type="radio" id=radio2 name=radio2 onclick=marrychg(0)  <%IF marryed="Y" THEN %>CHECKED<%END IF%> > 已婚 <font class=txt8>Đã kết hôn</font>
			&nbsp;<INPUT type="radio" id=radio2 name=radio2 onclick=marrychg(1)  <%IF marryed="N" THEN %>CHECKED<%END IF%> > 未婚 <font class=txt8>Chưa kết hôn</font>
			&nbsp;<INPUT type="radio" id=radio2 name=radio2 onclick=marrychg(2)  <%IF marryed="L" THEN %>CHECKED<%END IF%> > 離婚 <font class=txt8>Ly hôn</font>
			<input type=hidden name=marryed value="<%=marryed%>" size=1>
			</td>
			<td  align="right" width="60">教育程度<br><font class=txt8>Trình độ</font></td>
			<td  ><input name=school size=15 class=inputbox VALUE="<%=SCHOOL%>" ></td>
			</tr></table>
	</tr> 
	<TR   >
		<TD   align=right><font color="#cc0000">廠<br><font class=txt8>Xưởng</font></font></TD>
		<TD  >			
				<select name=WHSNO  class="txt" onkeydown="enterto()"  >
					<%SQL="SELECT * FROM BASICCODE WHERE FUNC='WHSNO' ORDER BY SYS_TYPE "
					SET RST = CONN.EXECUTE(SQL)
					WHILE NOT RST.EOF
					%>
					<option value="<%=RST("SYS_TYPE")%>" <%IF RST("SYS_TYPE")=WHSNO THEN %> SELECTED <%END IF%>><%=RST("SYS_VALUE")%></option>
					<%
					RST.MOVENEXT
					WEND
					rst.close
					%>
				</SELECT>
				<%SET RST=NOTHING %>
			  
				<input type=hidden name=unitno   class='readonly' readonly value="<%=unitno%>" size=2>
 
		</TD>
		<TD   align=right ><font color="#cc0000">職等<br><font class=txt8>Chức vụ</font></font></TD>
		<TD  colspan=2>		
				<select name=JOB  class="txt8" >
				<%SQL="SELECT * FROM BASICCODE WHERE FUNC='LEV'  ORDER BY SYS_TYPE "
					SET RST = CONN.EXECUTE(SQL)
					WHILE NOT RST.EOF
				%>
					<option value="<%=RST("SYS_TYPE")%>" <%IF RST("SYS_TYPE")=JOB THEN %> SELECTED <%END IF%>><%=RST("SYS_TYPE")%>-<%=RST("SYS_VALUE")%></option>
				<%
					RST.MOVENEXT
					WEND
					rst.close
					SET RST=NOTHING 
				%>
				</SELECT>
			 
			
		</TD>	
		
	</tr>
	<TR >
		<TD   align=right ><font color="#cc0000">部門/組<br><font class=txt8>Bộ phận / tổ</font></font></TD>
		<TD colspan=4>
				<table border="0" cellpadding="1" cellspacing="0" width="100%" class="txt8">
				<tr><td width="295">
				<select name=GROUPID  class="txt8" onchange=groupchg()     >
					<%SQL="SELECT * FROM BASICCODE WHERE FUNC='GROUPID' and sys_type <>'AAA' ORDER BY SYS_TYPE "
					SET RST = CONN.EXECUTE(SQL)
					'RESPONSE.WRITE SQL
					WHILE NOT RST.EOF
					%>
					<option value="<%=RST("SYS_TYPE")%>" <%IF RST("SYS_TYPE")=TRIM(GROUPID) THEN %> SELECTED <%END IF%>><%=RST("SYS_TYPE")%>-<%=RST("SYS_VALUE")%></option>
					<%
					RST.MOVENEXT
					WEND
					rst.close
					%>
				</SELECT>
				<%SET RST=NOTHING %>
				<select name=zuno  class="txt8"  style="width:170" >
					<%
					SQL="SELECT * FROM BASICCODE WHERE FUNC='ZUNO' AND LEFT(SYS_TYPE,4)='"& GROUPID &"' ORDER BY SYS_TYPE "
					SET RST = CONN.EXECUTE(SQL)					
					WHILE NOT RST.EOF
					%>
					<option value="<%=RST("SYS_TYPE")%>" <%IF RST("SYS_TYPE")=TRIM(ZUNO) THEN %> SELECTED <%END IF%>><%=RST("SYS_TYPE")%>-<%=RST("SYS_VALUE")%></option>
					<%
					RST.MOVENEXT
					WEND
					rst.close
					SET RST=NOTHING 
					%>
				</SELECT>
				<td>
				<TD align=right >班別<br><font class=txt8>Ca</font></TD>
		<TD >
			<select name="shift"  class="txt8" style='width:60' onkeydown="enterto()"  >
			<%SQL="SELECT * FROM BASICCODE WHERE FUNC='shift'  ORDER BY len(SYS_TYPE) desc , sys_type "
				SET RST = CONN.EXECUTE(SQL)
				WHILE NOT RST.EOF
			%>
				<option value="<%=RST("SYS_TYPE")%>" <%IF RST("SYS_TYPE")=shift THEN %> SELECTED <%END IF%>><%=RST("SYS_VALUE")%></option>
			<%
				RST.MOVENEXT
				WEND
				rst.close
			%>
			</SELECT>
			<%SET RST=NOTHING %>
			
			<select name="grps"  class="txt8" style='width:60' onkeydown="enterto()"  >
			<%SQL="SELECT * FROM BASICCODE WHERE FUNC='grps'  ORDER BY SYS_TYPE "
				SET RST = CONN.EXECUTE(SQL)
				WHILE NOT RST.EOF
			%>
				<option value="<%=RST("SYS_TYPE")%>" <%IF RST("SYS_TYPE")=grps THEN %> SELECTED <%END IF%>><%=RST("SYS_VALUE")%></option>
			<%
				RST.MOVENEXT
				WEND
				rst.close
			%>
			</SELECT>
			<%SET RST=NOTHING %>
		</TD> 
				</tr></table>
				
			 	
		</TD>
		
	</TR>
	<TR   >
			
	</TR>	  
	
</table>	
<%
			conn.close
			set conn=nothing
			%>
 <TABLE WIDTH=650 CLASS=txt BORDER="0"  cellspacing="1" cellpadding="2">
	<TR   >
		<td width="80" align=right>身分證號<br><font class=txt8>Số CMND</font></td>
		<TD width="180" ><input name=personID size=30 class=inputbox VALUE="<%=PERSONID%>" ></td>		
		<td  width="90" align=right >合同日<br><font class=txt8>Ky Hop Dong</font></td>
		<td ><input name=BHDAT size=15 class=inputbox VALUE="<%=bhdat%>"  onblur="date_change(6)" ><font class=txt8>(Ngày)</font></td>
	</TR>	
	<tr height=20 >
		<td  align=right><font color="#990000">稅號<br><font class=txt8>Mã số thuế</font></font></td>
		<TD ><input name=taxcode size=30 class=inputbox VALUE="<%=taxcode%>" maxlength=10></td>		
		<td nowrap align=right height=20 >加入工團<br><font class=txt8>Ngày vào công đoàn</font></td>
		<td ><input name=GTDAT size=15 class=inputbox VALUE="<%=GTDAT%>">(ex:202201)</td>
	</tr>
	<tr height=20 >
		<td align=right >發證日期<br><font class=txt8>Ngày cấp</font></td>
		<td><input name=PASSPORTNO size=30 class=inputbox VALUE="<%=PASSPORTNO%>" ></td>
		<td align=right height=25 >發証地點<br><font class=txt8>Nơi cấp </font></td>
		<td  ><input name=visaNo size=15 class=inputbox  VALUE="<%=visaNo%>" ></td>
	</tr>	 
	
	<tr>
		<td nowrap align=right height=20 >銀行帳號<br><font class=txt8>Số tài khoản </font></td>
		<td ><input name=bankid size=30 class=inputbox VALUE="<%=bankid%>"></td>
		<TD   align=right style="color:blue" >*保險號碼<br>Số Bảo hiểm</TD>
		<TD  colspan="2" >
			<input name="masobh"  size=40 class="inputbox"   value="<%=sobh%>" >
		</td>	 
	</tr>
	<tr height=20>
		
		<td nowrap align=right height=20 >聯絡電話<br><font class=txt8>Đ.T</font></td>
		<td><input name=phone size=30 class=inputbox VALUE="<%=phone%>"></td>
		<td nowrap align=right >手機<br><font class=txt8>ĐTDD</font></td>
		<td  ><input name=mobilephone size=30 class=inputbox VALUE="<%=MOBILEPHONE%>"></td>
	</tr>
	<tr>
		<td nowrap align=right height=20 >聯絡地址<br><font class=txt8>Địa chi</font></td>
		<td colspan=34><input name=homeaddr size=70 class=inputbox VALUE="<%=Homeaddr%>"></td>
	</tr>
	<tr height=20 >
		<td nowrap align=right height=20 >E-MAIL<br></td>
		<td colspan=3 ><input name=email size=70 class=inputbox VALUE="<%=email%>"></td>
	</tr>
	<tr>
		<td  align=right height=20 >備註說明<br><font class=txt8>Ghi Chú</font></td>
		<td colspan=3>
			<input   name=memo class="INPUTBOX" size=70 VALUE="<%=memo%>">
		</td>
	</tr>
</table>
<hr size=0	style='border: 1px dotted #999999;' align=left width=650>

<TABLE WIDTH=650>
		<tr ALIGN=center>
			<TD >
			<input type=button  name=send value="(Y)Confirm"  class=button onclick=go()>&nbsp;
			&nbsp;
			<input type=RESET name=send value="Tải lên上傳"  class=button onclick="Đăng ảnh()">
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			<input type=RESET name=send value="(X)Đóng 關閉"  class=button onclick=parent.close()>
			</TD>
		</TR>
</TABLE>


</form>


</body>
</html>

<script language=vbscript>

function upphots()
	empid=<%=self%>.empid.value
	open "sendPhoto.asp?empid="&empid, "_blank", "top=120, left=80, width=300 , height=300, scrollbars=yes "
end function

function prt()
	'<%=self%>.action="YEBE0301.toexcel.asp"
	'<%=self%>.submit()
	'<%=self%>.send(0).style.display="none"
	'<%=self%>.send(1).style.display="none"
	'<%=self%>.send(2).style.display="none"
	'window.print()
	'<%=self%>.send(0).style.display=""
	'<%=self%>.send(1).style.display=""
	'<%=self%>.send(2).style.display=""
end function

function chkempid()
	if <%=self%>.country.value<>"" and <%=self%>.whsno.value<>"" then
		code1=ucase(trim(<%=self%>.country.value))
		code2=ucase(trim(<%=self%>.whsno.value))
		code3=ucase(trim(<%=self%>.emptype.value))
		'alert  code1
		open "<%=self%>.back.asp?ftype=getempid&code1=" & code1 &"&code2=" & code2 &"&code3=" & code3 , "Back"
		'parent.best.cols="70%,30%"
	end if
end function

function empidchg()
	empidstr = Ucase(Trim(<%=self%>.empid.value))
	if empidstr<>"" then
		open "<%=self%>.back.asp?ftype=empidchk&code="& empidstr , "Back"
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

function typechg(x)
	if <%=self%>.radio3(0).checked=true then
		<%=self%>.emptype.value="A"
	elseif 	<%=self%>.radio3(1).checked=true then
		<%=self%>.emptype.value="B"
	elseif 	<%=self%>.radio3(2).checked=true then
		<%=self%>.emptype.value="C"
	else
		<%=self%>.emptype.value=""
	end if

end function

function marrychg(x)
	if <%=self%>.radio2(0).checked=true then
		<%=self%>.marryed.value="Y"
	elseif 	<%=self%>.radio2(1).checked=true then
		<%=self%>.marryed.value="N"
	elseif 	<%=self%>.radio2(2).checked=true then
		<%=self%>.marryed.value="L"
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
	'if <%=self%>.unitno.value="" then
	'	ALERT "請輸入處/所!!"
	'	<%=SELF%>.unitno.FOCUS()
	'	EXIT FUNCTION
	'end if
	'if <%=self%>.GROUPID.value="" then
	'	ALERT "請輸入部門單位!!"
	'	<%=SELF%>.GROUPID.FOCUS()
	'	EXIT FUNCTION
	'end if
	'if <%=self%>.shift.value="" then
	'	ALERT "請輸入班別!!"
	'	<%=SELF%>.shift.FOCUS()
	'	EXIT FUNCTION
	'end if
	photosname=<%=self%>.empid.value&".jpg"
	<%=SELF%>.ACTION="YEBE0102.upd.asp"
	<%=SELF%>.SUBMIT
END FUNCTION

'*******檢查日期*********************************************
FUNCTION date_change(a)

if a=1 then
	INcardat = Trim(<%=self%>.indat.value)
elseif a=2 then
	INcardat = Trim(<%=self%>.pissueDate.value)
elseif a=3 then
	INcardat = Trim(<%=self%>.pduedate.value)
elseif a=4 then
	INcardat = Trim(<%=self%>.WKD_DueDate.value)
elseif a=5 then
	INcardat = Trim(<%=self%>.outdat.value)
elseif a=6 then
	INcardat = Trim(<%=self%>.bhdat.value)
end if

IF INcardat<>"" THEN
	ANS=validDate(INcardat)
	IF ANS <> "" THEN
		if a=1 then
			Document.<%=self%>.indat.value=ANS
		elseif a=2 then
			Document.<%=self%>.pissueDate.value=ANS
		elseif a=3 then
			Document.<%=self%>.pduedate.value=ANS
		elseif a=4 then
			Document.<%=self%>.WKD_DueDate.value=ANS
		elseif a=5 then
			Document.<%=self%>.outdat.value=ANS
		elseif a=6 then
			Document.<%=self%>.bhdat.value=ANS
		end if
	ELSE
		ALERT "EZ0067:輸入日期不合法 !!"
		if a=1 then
			Document.<%=self%>.indat.value=""
			Document.<%=self%>.indat.focus()
		elseif a=2 then
			Document.<%=self%>.pissueDate.value=""
			Document.<%=self%>.pissueDate.focus()
		elseif a=3 then
			Document.<%=self%>.pduedate.value=""
			Document.<%=self%>.pduedate.focus()
		elseif a=4 then
			Document.<%=self%>.WKD_DueDate.value=""
			Document.<%=self%>.WKD_DueDate.focus()
		elseif a=5 then
			Document.<%=self%>.outdat.value=""
			Document.<%=self%>.outdat.focus()
		elseif a=6 then
			Document.<%=self%>.bhdat.value=""
			Document.<%=self%>.bhdat.focus()
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

</script>


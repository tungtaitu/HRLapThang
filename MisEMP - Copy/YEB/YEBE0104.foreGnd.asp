<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<!-- #include file = "../GetSQLServerConnection.fun" -->
<!-- #include file="../ADOINC.inc" -->
<%
if session("netuser")="" then
	response.write "使用者帳號為空!!請重新登入!!"
	response.end
end if

SELF = "YEBE0104"

Set conn = GetSQLServerConnection()
Set rs = Server.CreateObject("ADODB.Recordset")

nowmonth = year(date())&right("00"&month(date()),2)
if month(date())="01" then
	calcmonth = year(date()-1)&"12"
else
	calcmonth =  year(date())&right("00"&month(date())-1,2)
end if



FUNCTION FDT(D)
	Response.Write YEAR(D)&"/"&RIGHT("00"&MONTH(D),2)&"/"&RIGHT("00"&DAY(D),2)
END FUNCTION

wbid = request("wbid")

sql="select c.driverID,  isnull(b.incustsname,'') incustsname , a.* from "&_
	"(select convert(char(10),indat,111) indate, convert(char(10),outdat,111) outdate, * from wbempfile "&_
	"where wbid='"&wbid&"' and isnull(status,'')<>'D'  ) a  "&_
	"left join (select incustid, incustsname from [yfymis].dbo.ydbscust ) b on b.incustid = a.fac "&_
	"left join (select *  from [yfymis].dbo.ysbmlrif where isnull(Status,'')<>'D'  ) c on c.lorry = a.lorry  "
rs.open sql, conn, 1, 3
if not rs.eof then
	wbwhsno=rs("wbwhsno")
	wbloai=rs("loai")
	wbid=rs("wbid")
	cardno=rs("cardno")
	wbname_vn=rs("wbname_vn")
	wbname_cn=rs("wbname_cn")
	indate=rs("indate")
	outdate=rs("outdate")
	yy=rs("yy")
	mm=rs("mm")
	dd=rs("dd")
	age=rs("age")
	sex=rs("sex")
	personid=rs("personid")
	phone=rs("phone")
	mobile=rs("mobile")
	fac=rs("fac")
	lorry=rs("lorry")
	soxe=rs("soxe")
	job=rs("job")
	addr=rs("addr")
	wbmemo=rs("wbmemo")
	sex=rs("sex")
	ages=rs("age")
	personid=rs("personid")
	incustsname = rs("incustsname")
	flag=rs("flag")
	outmemo = rs("outmemo")
	filename  = rs("filename")
	driverID = rs("driverID") 
	sysid = rs("aid")
end if


wbphotoid = request("wbphotoid")
if wbphotoid<>"" then
	sqlx="select * from wbphotos where  aid='"& wbphotoid &"' "
	set rsx=conn.execute(sqlx)
	if not rsx.eof then
		filename=rsx("filename")
	end if
	set rsx=nothing
end if
%>

<html>

<head>

<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta HTTP-EQUIV="refresh" >
<link rel="stylesheet" href="../Include/style.css" type="text/css">
<link rel="stylesheet" href="../Include/style2.css" type="text/css">
</head>
<body  topmargin="0" leftmargin="10"  marginwidth="0" marginheight="0"  onkeydown="enterto()" onload="f()">
<form name="<%=self%>" method="post" action="<%=self%>.upd.asp">
<INPUT TYPE=HIDDEN NAME="UID" VALUE=<%=SESSION("NETUSER")%>>
<INPUT TYPE=hidden NAME=nowmonth VALUE="<%=nowmonth%>">
<table width="460" border="0" cellspacing="0" cellpadding="0">
	<tr><TD  >
		<img border="0" src="../image/icon.gif" align="absmiddle">
		1.3外包資料維護 (修改 )xoa/sua tu lieu thau ngoai
		</TD>
	</tr>
</table>
<hr size=0	style='border: 1px dotted #999999;' align=left width=500 >

<TABLE WIDTH=520 CLASS=txt BORDER=0 cellspacing="2" cellpadding="1" >
	<TR height=35 >
		<TD   align=right width=100>類別<BR><font class=txt8>Loai</font></TD>
		<TD  valign=top>
			<input type=hidden name=wbloai  class=txt8  value="<%=wbloai%>" >
			<select name=wbloaib  class=txt8  disabled >
				<option value="">請選擇類別</option>
				<%
				SQL="SELECT * FROM BASICCODE WHERE FUNC='WB'   ORDER BY SYS_TYPE "
				SET RST = CONN.EXECUTE(SQL)
				WHILE NOT RST.EOF
				%>
				<option value="<%=RST("SYS_TYPE")%>" <%if rst("sys_type")=wbloai then%>selected<%end if%>><%=RST("SYS_TYPE")%><%=RST("SYS_VALUE")%></option>
				<%
				RST.MOVENEXT
				WEND
				rst.close
				%>
			</SELECT>
			<%SET RST=NOTHING %>
		</TD>
		<td    align=center valign=top  nowrap  rowspan="4"  colspan=2>
			<%if filename<>"" then %>
				<div style='cursor:hand' onclick="vbscript:upphots()"><img src="wbphotos/<%=filename%>"  border=1 width=120 height=130  ></div>
			<%else%>
				<a href="vbscript:upphots()"><img src="pic/noimg.gif"  border=0> </a>
			<%end if%>
			<input name=filename value="<%=filename%>" type=hidden>
			<input name=wbphotoid value="<%=wbphotoid%>" type=hidden>

		</td>

	</TR>
	<TR height=35 >
		<TD width=100 align=right>廠別<BR><font class=txt8>Xuong</font></TD>
		<TD width=150 valign=top>
			<input type=hidden name=WHSNO  class=txt8  value="<%=wbwhsno%>" >
			<select name=WHSNOb  class=txt8  disabled >
				<option value="">請選擇廠別</option>
				<%
				SQL="SELECT * FROM BASICCODE WHERE FUNC='WHSNO'  ORDER BY SYS_TYPE "
				SET RST = CONN.EXECUTE(SQL)
				WHILE NOT RST.EOF
				%>
				<option value="<%=RST("SYS_TYPE")%>" <%if rst("sys_type")=wbwhsno then%>selected<%end if%>><%=RST("SYS_VALUE")%></option>
				<%
				RST.MOVENEXT
				WEND
				rst.close
				%>
			</SELECT>
			<%SET RST=NOTHING %>
		</TD>
	</tr>
	<TR height=35 >
		<TD  nowrap align=right height=25>編號<BR><font class=txt8>So The</font></TD>
		<TD  valign=top>
			<INPUT NAME="EMPID" SIZE=12 CLASS=readonly maxlength=5 value="<%=wbid%>">
			<INPUT NAME="sysid" SIZE=12 CLASS=readonly maxlength=5 value="<%=sysid%>">
		</TD>
	</tr>
	<TR height=35 >
		<TD   align=right height=25>到職日<BR><font class=txt8>NVX</font></TD>
		<TD valign=top><INPUT NAME=indat SIZE=12 CLASS=INPUTBOX value="<%=indate%>" onblur="date_change(1)"></TD>
	</td>

	<TR height=35 >
		<TD   align=right >職等<br><font class=txt8>Chuc vu</font></TD>
		<TD  valign=top>
			<input name="job"  value="<%=job%>" class="inputbox" size="15" maxlength=50 />
		</TD>
		<TD   align=right>國籍<br><font class=txt8>Quoc Tich</font></TD>
		<TD  valign=top >
			<select name=country  class=font9 style='width:75'  >
				<%SQL="SELECT * FROM BASICCODE WHERE FUNC='country'  ORDER BY SYS_type desc  "
				SET RST = CONN.EXECUTE(SQL)
				WHILE NOT RST.EOF
				%>
				<option value="<%=RST("SYS_TYPE")%>" <%IF country=RST("SYS_TYPE") THEN %> SELECTED <%END IF%>><%=RST("SYS_VALUE")%></option>
				<%
				RST.MOVENEXT
				WEND
				rst.close
				%>
			</SELECT>
			<%SET RST=NOTHING 
			conn.close
			set conn=nothing
			%>
		</TD>
	</TR>
	<TR height=35>
		<TD   align=right>姓名(中)<BR><font class=txt8>Ho Ten(Hoa)</font></TD>
		<TD valign=top >
			<INPUT NAME=nam_cn SIZE=22 CLASS=INPUTBOX  value="<%=wbname_cn%>">
		</TD>
		<TD   align=right >姓名(越)<BR><font class=txt8>Ho Ten(viet)</font></TD>
		<TD valign=top  ><INPUT NAME=nam_vn SIZE=30 CLASS=INPUTBOX8  value="<%=wbname_vn%>"></TD>
	</TR>
	<TR height=35 >
		<TD   align=right >出生日期<br><font class=txt8>Ngay Sinh</font></TD>
		<TD colspan=3 valign=top>
		<INPUT NAME=BYY SIZE=5 CLASS=INPUTBOX  MAXLENGTH=4 ONBLUR="CHKVALUE(1)" value="<%=yy%>"> &nbsp;年<font class=txt8>Năm</font>&nbsp;
		<INPUT NAME=BMM SIZE=5 CLASS=INPUTBOX  MAXLENGTH=2 ONBLUR="CHKVALUE(2)" value="<%=mm%>"> &nbsp;月<font class=txt8>Tháng</font>&nbsp;
		<INPUT NAME=BDD SIZE=5 CLASS=INPUTBOX  MAXLENGTH=2 ONBLUR="CHKVALUE(3)" value="<%=dd%>"> &nbsp;日<font class=txt8>Ngày</font>&nbsp;&nbsp;&nbsp;
		Age(Thoi)<input name=ages size=3 class=inputbox value="<%=ages%>">

		</TD>
	</TR>
	</tr>
		<td  align=right >性別<BR><font class=txt8>GiỚi Tính</font></td>
		<td valign=top>
			<INPUT type="radio" id=radio1 name=radio1 onclick=sexchg(0) <%if sex="M" then%>selected<%end if%>>男<font class=txt8>Nam</font>&nbsp;
			<INPUT type="radio" id=radio1 name=radio1 onclick=sexchg(1) <%if sex="F" then%>selected<%end if%>> 女<font class=txt8>Nữ</font>
			<input type=hidden name=sexstr value="<%=sex%>" size=1>
		</td>
		<td   align=right height=25 >身分証字號<BR><font class=txt8>So CMND</font></td>
		<td valign=top><input name=personID size=30 class=inputbox value="<%=personid%>"></td>
	</tr>
	<tr height=35>
		<td  align=right >卡片號碼<BR><font class=txt8>Card No</font></td>
		<Td><INPUT NAME=cardno SIZE=22 CLASS=INPUTBOX  value="<%=cardno%>"></td>
	</tr>
	<Tr>
		<Td colspan=4><hr size=0 style='border: 1px dotted #999999;'  ></td>
	</tr>
	<tr height=35>
		<td   align=right height=25 >廠商/車行<br><font class=txt8>NHÀ CUNG ỨNG</font></td>
		<td valign=top>
			<%if wbloai="01" then%>
				<input name=fac size=22 class=readonly value="<%=incustsname%>" readonly  >
			<%else%>
				<input name=fac size=22 class=inputbox value="<%=fac%>"    >
			<%end if%>
			<input name=xhid type=hidden value="<%=fac%>">
		</td>
		<td align=right >車號</td>
		<td valign=top>
			<input name=lorry size=3 class=readonly readonly  maxlength=5 value="<%=lorry%>" >
			<input name=soxe size=9 class=readonly readonly  value="<%=driverID%>" style='color:<%if driverID=soxe then%>black<%else%>red<%end if%>' >
		</td>
	</tr>
	<tr >
		<td   align=right height=25 >聯絡電話<BR><font class=txt8>Đ.T</font></td>
		<td valign=top><input name=phone size=22 class=inputbox  value="<%=phone%>"></td>
		<td   align=right >手機<BR><font class=txt8>ĐTDD</font></td>
		<td valign=top><input name=mobile size=15 class=inputbox value="<%=mobile%>"></td>
	</tr>
	<tr height=35>
		<td nowrap align=right height=25 >聯絡地址<BR><font class=txt8>Địa chi<BR></td>
		<td colspan=3 valign=top><input name=homeaddr size=65 class=inputbox value="<%=addr%>" ></td>
	</tr>
	<tr height=35>
		<td nowrap align=right height=25 >其他說明<BR><font class=txt8>Ghi Chu</font></td>
		<td valign=top colspan=3><input name=memo size=65 class=inputbox value="<%=wbmemo%>"></td>
	</tr>
	<Tr>
		<Td colspan=4><hr size=0 style='border: 1px dotted #999999;'  ></td>
	</tr>
	<tr height=35>
		<td nowrap align=right height=25 >離職日期<BR><font class=txt8>NTV</font></td>
		<td valign=top ><input name=outdat size=11 class=inputbox value="<%=outdate%>"  onblur="date_change(2)"></td>
		<td nowrap colspan=2 >
			<input type=checkbox name=ff onclick=ffchg() <%if flag="Y" then%>checked<%end if%>> 已還識別證(TRA THE)
			<input type=hidden   name=flag size=1 value="<%=flag%>">
		</td>
	</tr>
	<tr height=35>
		<td nowrap align=right height=25 >離職原因<BR><font class=txt8>Ly Do Thoi Viec</font></td>
		<td valign=top colspan=3><input name=outmemo size=65 class=inputbox value="<%=outmemo%>"></td>
	</tr>
</table>
<hr size=0	style='border: 1px dotted #999999;' align=left width=500>
<TABLE WIDTH=600>
		<tr ALIGN=center>
			<TD >
			<input type=button  name=send value="(Y)確認Confirm"  class=button onclick=go()>
			<input type=RESET name=send value="(N)取消Cancel"  class=button>
			<%if wbphotoid<>"" then%>
				<input type=button name=send value="(Colse)關閉"  class=button onclick='parent.close()'>
			<%else%>
				<input type=button name=send value="(Back)主畫面"  class=button onclick='gob()'>
			<%end if%>
			<input type=button  name=send value="(D)刪除DEL"  class=button onclick="godelwb()">
			</TD>
		</TR>
</TABLE>


</form>


</body>
</html>
<script language=vbscript>

'-----------------enter to next field
function getlorry()
	open "getlorry.asp", "Back"
	parent.best.cols="50%,50%"
end function

function gob()
	history.back
end function

function ffchg()
	if <%=self%>.ff.checked=true then
		<%=self%>.flag.value="Y"
	else
		<%=self%>.flag.value=""
	end if
end function

function upphots()
	empid=<%=self%>.empid.value
	open "sendPhoto.asp?flag=WB&empid="&empid, "_blank", "top=120, left=80, width=300 , height=300, scrollbars=yes "
end function


function enterto()
	if window.event.keyCode = 13 then window.event.keyCode =9
end function

function f()
	'<%=self%>.whsno.focus()
	<%=self%>.indat.select()
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

function whsnochg()
	code1 = <%=self%>.whsno.value
	code2 = <%=self%>.wbloai.value
	if code1<>"" and code2<>"" then
		open "<%=self%>.back.asp?ftype=getwbid&code1="&code1 &"&code2="& code2  , "Back"
		parent.best.cols="100%,0%"
	end if
end function

function loaichg()
	code1 = <%=self%>.wbloai.value
	code2 = <%=self%>.whsno.value
	if code1<>"" and code2<>"" then
		open "<%=self%>.back.asp?ftype=getwbid&code1="&code1 &"&code2="& code2  , "Back"
		'parent.best.cols="50%,50%"
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
	if <%=self%>.whsno.value="" then
		ALERT "請選擇廠別(Ko. co nhap vao Xuong)!!"
		<%=SELF%>.whsno.FOCUS()
		EXIT FUNCTION
	end if
	if <%=self%>.wbloai.value="" then
		ALERT "請選擇類別(Ko. co nhap vao Loai)!!"
		<%=SELF%>.wbloai.FOCUS()
		EXIT FUNCTION
	end if
	IF  <%=SELF%>.EMPID.VALUE="" THEN
		ALERT "請輸入編號(Ko. co nhap vao So )!!"
		<%=SELF%>.EMPID.FOCUS()
		EXIT FUNCTION
	END IF
	if <%=self%>.nam_vn.value="" then
		ALERT "請輸入姓名(越)(Ko. co nhap vao Ho Ten(Viet)!!"
		<%=SELF%>.nam_vn.FOCUS()
		EXIT FUNCTION
	end if

	IF  <%=SELF%>.personID.VALUE="" THEN
		ALERT "請輸入身分證號(Ko. co nhap vao CMND )!!"
		<%=SELF%>.personID.FOCUS()
		EXIT FUNCTION
	END IF
	IF  <%=SELF%>.fac.VALUE="" THEN
		ALERT "請輸入供應商/車行(Ko. co nhap vao NHÀ CUNG ỨNG )!!"
		<%=SELF%>.fac.FOCUS()
		EXIT FUNCTION
	end if
	if <%=self%>.wbloai.value="01" then
		IF  <%=SELF%>.fac.VALUE="" THEN
			ALERT "請輸入供應商/車行(Ko. co nhap vao NHÀ CUNG ỨNG )!!"
			<%=SELF%>.fac.FOCUS()
			EXIT FUNCTION
		end if
		IF  <%=SELF%>.soxe.VALUE="" THEN
			ALERT "請輸入車號(Ko. co nhap vao So Xe )!!"
			<%=SELF%>.soxe.FOCUS()
			EXIT FUNCTION
		end if
	END IF


	<%=SELF%>.ACTION="<%=self%>.upd.asp"
	<%=SELF%>.SUBMIT
END FUNCTION

FUNCTION godelwb()
	if confirm("確定要刪除Xoa di khong?" ,64) then
		<%=SELF%>.ACTION="<%=self%>.upd.asp?flag=del"
		<%=SELF%>.SUBMIT
	end if
END FUNCTION

'*******檢查日期*********************************************
FUNCTION date_change(a)

if a=1 then
	INcardat = Trim(<%=self%>.indat.value)
elseif a=2 then
	INcardat = Trim(<%=self%>.outdat.value)
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
			Document.<%=self%>.outdat.value=ANS
		elseif a=3 then
			Document.<%=self%>.pduedate.value=ANS
		elseif a=4 then
			Document.<%=self%>.vduedate.value=ANS
		elseif a=5 then
			Document.<%=self%>.outdat.value=ANS
		end if
	ELSE
		ALERT "EZ0067:輸入日期不合法 (yyyy/mm/dd) !!"
		if a=1 then
			Document.<%=self%>.indat.value=""
			Document.<%=self%>.indat.focus()
		elseif a=2 then
			Document.<%=self%>.outdat.value=""
			Document.<%=self%>.outdat.focus()
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

</script>


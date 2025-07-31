<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<!-- #include file = "../GetSQLServerConnection.fun" -->
<!-- #include file="../ADOINC.inc" -->

<!--#include file="../include/sideinfo.inc"-->
<%
if session("netuser")="" then
	response.write "使用者帳號為空!!請重新登入!!"
	response.end
end if

SELF = "YEBE0103"

Set conn = GetSQLServerConnection()
'Set rs = Server.CreateObject("ADODB.Recordset")

nowmonth = year(date())&right("00"&month(date()),2)
if month(date())="01" then
	calcmonth = year(date()-1)&"12"
else
	calcmonth =  year(date())&right("00"&month(date())-1,2)
end if



FUNCTION FDT(D)
	Response.Write YEAR(D)&"/"&RIGHT("00"&MONTH(D),2)&"/"&RIGHT("00"&DAY(D),2)
END FUNCTION



%>

<html>

<head>

<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta HTTP-EQUIV="refresh" >
<link rel="stylesheet" href="../Include/style.css" type="text/css">
<link rel="stylesheet" href="../Include/style2.css" type="text/css">
</head>
<body onkeydown="enterto()" onload="f()">
<form name="<%=self%>" method="post" action="<%=self%>.upd.asp">
<INPUT TYPE=HIDDEN NAME="UID" VALUE=<%=SESSION("NETUSER")%>>
<INPUT TYPE=hidden NAME=nowmonth VALUE="<%=nowmonth%>">
 
	<table border=0 style="height:70px"><tr><td>&nbsp;</td></tr></table>
	
	<table width="100%" BORDER=0 align=center cellpadding=0 cellspacing=0 >					
		<tr>
			<td align="center">
				<TABLE id="myTableForm" width="80%" border=0>
					<tr>
						<td colspan=4>
							<table BORDER=0 cellspacing="1" cellpadding="1" class="txt" width="100%">
								<tr>
									<Td align=center bgcolor="#ffcccc" width="25%" onMouseOver="this.style.backgroundColor='#FFEB78'" onMouseOut="this.style.backgroundColor='#ffcccc'"   style="cursor:hand" ><span class="btn btn-warning text-white btn-block shadow">外包資料新增<br>tu lieu moi thau ngoai</span></td>
									<Td align=center bgcolor="#ffffff" width="25%" onMouseOver="this.style.backgroundColor='#FFEB78'" onMouseOut="this.style.backgroundColor='#ffffff'"   style="cursor:hand"><a href="yebe0104.fore.asp" class="btn btn-primary btn-block shadow">外包資料維護<br>xoa/sua tu lieu thau ngoai</a></td>
									<Td align=center bgcolor="#ffffff" width="25%" onMouseOver="this.style.backgroundColor='#FFEB78'" onMouseOut="this.style.backgroundColor='#ffffff'"   style="cursor:hand"><a href="yebe0105.fore.asp" class="btn btn-primary btn-block shadow">外包資料查詢<br>K.Tra tu lieu</a></td>
									<Td align=center bgcolor="#ffffff"  width="25%" onMouseOver="this.style.backgroundColor='#FFEB78'" onMouseOut="this.style.backgroundColor='#ffffff'"   style="cursor:hand"><a href="yebe0103C.asp" class="btn btn-primary btn-block shadow">照片上傳與新增<br>Update Photos</a></td>
								</tr>
							</table>
						</td>
					</tr>
					<TR >
						<TD width="10%" align=right>廠別<BR><font class=txt8>Xưởng</font></TD>
						<TD  width="40%">
							<select name=WHSNO   onchange='whsnochg()' style="width:98%">
								<option value="">請選擇廠別 Mời chọn xưởng</option>
								<%
								if session("rights")="0" then
									SQL="SELECT * FROM BASICCODE WHERE FUNC='WHSNO'  ORDER BY SYS_TYPE "
								else
									SQL="SELECT * FROM BASICCODE WHERE FUNC='WHSNO' and sys_type='"& session("netwhsno") &"' ORDER BY SYS_TYPE "
								end if
								SET RST = CONN.EXECUTE(SQL)
								WHILE NOT RST.EOF
								%>
								<option value="<%=RST("SYS_TYPE")%>" ><%=RST("SYS_VALUE")%></option>
								<%
								RST.MOVENEXT
								WEND
								rst.close
								SET RST=NOTHING
								%>
							</SELECT>
						</TD>
						<TD  width="10%"  align=right>類別<BR><font class=txt8>Loại</font></TD>
						<TD  width="40%" >
							<select name=wbloai   onchange='whsnochg()' style="width:98%">
								<option value="">請選擇類別 Mời chọn thầu </option>
								<%
								SQL="SELECT * FROM BASICCODE WHERE FUNC='WB' and sys_type<>'01'  ORDER BY SYS_TYPE "
								SET RST = CONN.EXECUTE(SQL)
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
						</TD>
					</TR>
					<TR >
						<TD align=right >到職日<BR><font class=txt8>Ngày nhập xưởng</font></TD>
						<TD><INPUT type="text" style="width:98%" NAME=indat value=<%=fdt(date())%> onblur="date_change(1)"></TD>
						<TD  nowrap align=right >編號<BR><font class=txt8>Mã số</font></TD>
						<TD><INPUT type="text" style="width:98%" NAME="EMPID" ONCHANGE='EMPIDCHG()' maxlength=5 value="<%=eid%>"></TD>
					</tr>
					<TR>
						<TD align=right >職等<br><font class=txt8>Chức vự</font></TD>
						<TD>
							<input type="text" style="width:98%" name="job"  maxlength=50 />
						</TD>
						<TD   align=right>國籍<br><font class=txt8>Quốc tịch</font></TD>
						<TD   >
							<select name=country   >
								<%SQL="SELECT * FROM BASICCODE WHERE FUNC='country'  ORDER BY SYS_type desc  "
								SET RST = CONN.EXECUTE(SQL)
								WHILE NOT RST.EOF
								%>
								<option value="<%=RST("SYS_TYPE")%>" <%IF SESSION("NETWHSNO")=RST("SYS_TYPE") THEN %> SELECTED <%END IF%>><%=RST("SYS_VALUE")%></option>
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
					<TR >
						<TD   align=right>姓名(中)<BR><font class=txt8>Họ tên(Hoa)</font></TD>
						<TD  >
							<INPUT type="text" style="width:98%" NAME=nam_cn >
						</TD>
						<TD   align=right >姓名(越)<BR><font class=txt8>Họ tên(Việt)</font></TD>
						<TD   ><INPUT type="text" style="width:98%" NAME=nam_vn ></TD>
					</TR>
					<TR>
						<TD align=right >出生日期<br><font class=txt8>Ngày sinh</font></TD>
						<TD colspan=3>
							<INPUT type="text" style="width:10%" NAME=BYY SIZE=5   MAXLENGTH=4 ONBLUR="CHKVALUE(1)" >
							年<font class=txt8>Năm</font>
							<INPUT type="text" style="width:7%" NAME=BMM SIZE=5   MAXLENGTH=2 ONBLUR="CHKVALUE(2)" >
							月<font class=txt8>Tháng</font>
							<INPUT type="text" style="width:7%" NAME=BDD SIZE=5   MAXLENGTH=2 ONBLUR="CHKVALUE(3)" >
							日<font class=txt8>Ngày</font>
							Age(Thoi)
							<input type="text" style="width:7%" name=ages size=3  >
						</TD>
					</TR>
					<tr>
						<td  align=right >性別<BR><font class=txt8>GiỚi Tính</font></td>
						<td>
							<INPUT type="radio" id=radio1 name=radio1 onclick=sexchg(0) >男<font class=txt8>Nam</font>&nbsp;
							<INPUT type="radio" id=radio1 name=radio1 onclick=sexchg(1)> 女<font class=txt8>Nữ</font>
							<input type=hidden name=sexstr value="<%=sex%>" size=1>
						</td>
						<td   align=right  >身分証字號<BR><font class=txt8>Số CMND</font></td>
						<td ><input type="text" style="width:98%" name=personID ></td>
					</tr>
					<Tr>
						<Td colspan=4><hr size=0	style='border: 1px dotted #999999;'  ></td>
					</tr>
					<tr >
						<td   align=right  >廠商/車行<br><font class=txt8>NHÀ CUNG ỨNG</font></td>
						<td ><input type="text" style="width:98%" name=fac ></td>
						<td align=right ><a href="vbscript:getlorry()"><font color=blue>車號<BR><font class=txt8>Biển số xe</font></font></a></td>
						<td nowrap>
							<input type="text" style="width:20%" name=lorry  maxlength=5 >
							<input type="text" style="width:40%" name=soxe   readonly >
							<input type=hidden name=TAN size=3  >
						</td>
					</tr>
					<tr >
						<td   align=right  >聯絡電話<BR><font class=txt8>Đ.T liên lạc</font></td>
						<td ><input type="text" style="width:98%" name=phone></td>
						<td   align=right >手機<BR><font class=txt8>ĐTDĐ</font></td>
						<td ><input type="text" style="width:98%" name=mobilephone></td>
					</tr>
					<tr>
						<td nowrap align=right  >聯絡地址<BR><font class=txt8>Địa chỉ</font></td>
						<td colspan=3 ><input type="text" style="width:98%" name=homeaddr ></td>
					</tr>
					<tr>
						<td nowrap align=right>其他說明<BR><font class=txt8>Ghi Chú</font></td>
						<td  colspan=3><input type="text" style="width:98%" name=memo ></td>
					</tr>							
					<tr >
						<TD align=center colspan=4>
							<input type=button  name=send value="(Y)確　　認  XÁC NHẬN"  class="btn btn-sm btn-danger" onclick=go()>
							<input type=RESET name=send value="(N)取 　　消  HỦY BỎ"  class="btn btn-sm btn-outline-secondary">
						</TD>
					</TR>
				</TABLE>
			</td>
		</tr>
	</table>
			
</form>


</body>
</html>
<script language=vbscript>

'-----------------enter to next field
function getlorry()
	open "getlorry.asp", "Back"
	parent.best.cols="60%,40%"
end function



function enterto()
	if window.event.keyCode = 13 then window.event.keyCode =9
end function

function f()
	<%=self%>.whsno.focus()
	'<%=self%>.EMPID.select()
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


	<%=SELF%>.ACTION="<%=self%>.upd.asp?act=EMPADDNEW"
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

</script>


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

SELF = "YEBE0101"

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
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>

'-----------------enter to next field
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
	code = <%=self%>.whsno.value
	open "<%=self%>.back.asp?ftype=getempid&code="&code , "Back"	
	'parent.best.cols="50%,50%"
end function 
 
</SCRIPT>
</head>
<body  onkeydown="enterto()" onload="f()">
<form name="<%=self%>" method="post" action="<%=self%>.upd.asp">
<INPUT TYPE=HIDDEN NAME="UID" VALUE=<%=SESSION("NETUSER")%>>
<INPUT TYPE=hidden NAME=nowmonth VALUE="<%=nowmonth%>">  

	<table border=0 style="height:70px"><tr><td>&nbsp;</td></tr></table>
	<table width="100%" BORDER=0 align=center cellpadding=0 cellspacing=0 >
		<tr>
			<td align="center">
				<table id="myTableForm" width="94%" border=0>
					<tr><td colspan=6 style="height:50px">&nbsp;</td></tr>
					<TR>
						<TD align=right nowrap>廠別<BR><font class=txt8>Xưởng</font></TD>
						<TD nowrap>
							<input type="hidden" value="VN" name="country1">
							<select name=WHSNO   onchange='whsnochg()' style="width:120px">
								<option value="">-----</option>
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
								%>
							</SELECT>
							<select name="country"  style="width:120px">								
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
							<%SET RST=NOTHING %>
						</TD>									
						<TD align=right>到職日<BR><font class=txt8>Ngày nhập xưởng</font></TD>
						<TD><INPUT type="text" NAME=indat SIZE=12  value=<%=fdt(date())%> onblur="date_change(1)"></TD> 
						<td nowrap align=right>離職日<BR><font class=txt8>Ngày thôi việc</font></td>
						<td><input type="text" name=outdat size=15   onblur="date_change(5)" ></td>
					</TR> 
					<TR>
						<TD align="right">員工編號<BR><font class=txt8>Mã số nhân viên</font></TD>
						<TD><INPUT type="text" NAME="EMPID" SIZE=15 readonly   ONCHANGE='EMPIDCHG()' maxlength=5 value="<%=eid%>"></TD>											
						<TD align=right nowrap>姓名(中)<BR><font class=txt8>Họ tên(Hoa)</font></TD>
						<TD><INPUT type="text" NAME=nam_cn SIZE=30 ></TD> 
						<TD align=right nowrap>姓名(英/越)<BR><font class=txt8>Họ tên(Việt)</font></TD>
						<TD><INPUT type="text" style="width:98%" NAME=nam_vn  ></TD>
					</TR>
					<tr>
						<TD align=right >出生日期<br><font class=txt8>Ngày sinh</font></TD>
						<TD colspan=3 >
							<INPUT type="text" NAME=BYY SIZE=5   MAXLENGTH=4 ONBLUR="CHKVALUE(1)" style="vertical-align:middle">
							年<font class=txt8>Năm</font>
							<INPUT type="text" NAME=BMM SIZE=4   MAXLENGTH=2 ONBLUR="CHKVALUE(2)" style="vertical-align:middle">
							月<font class=txt8>Tháng</font>
							<INPUT type="text" NAME=BDD SIZE=4   MAXLENGTH=2 ONBLUR="CHKVALUE(3)" style="vertical-align:middle">
							日<font class=txt8>Ngày</font>
							<input type=hidden name=ages size=3  >																
						</TD>
						<td align=right  >性別<BR><font class=txt8>Giới Tính</font></td>
						<td>
							<INPUT type="radio" id=radio1 name=radio1 onclick=sexchg(0) >男<font class=txt8>Nam</font>&nbsp;
							<INPUT type="radio" id=radio1 name=radio1 onclick=sexchg(1)> 女<font class=txt8>Nữ</font>
							<input type=hidden name=sexstr value="<%=sex%>" size=1> 		
						</td>
					</tr>
					<tr>
						<TD align=right>組/部門<br><font class=txt8>Bộ phận/ Tổ</font></TD>
						<TD colspan=3>							
							<%SQL="SELECT * FROM BASICCODE WHERE FUNC='unit' and sys_type<>'AAA' ORDER BY SYS_TYPE desc "
								SET RST = CONN.EXECUTE(SQL)
							%>
							<select name=unitno   onchange=unitchg() style="width:24%;vertical-align:middle">
								<option value="">-----</option>
								<%
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
						
							<%SQL="SELECT * FROM BASICCODE WHERE FUNC='GROUPID' and sys_type like 'A06%' ORDER BY SYS_TYPE "
								SET RST = CONN.EXECUTE(SQL)%>
							<select name=GROUPID   onchange=groupchg()  style="width:24%;vertical-align:middle" >
								<option value="" <%if request("GROUPID")="" then%>selected<%end if%>>-----</option>				
								<%WHILE NOT RST.EOF
								%>
								<option value="<%=RST("SYS_TYPE")%>" ><%=RST("SYS_TYPE")%>-<%=RST("SYS_VALUE")%></option>
								<%
								RST.MOVENEXT
								WEND
								rst.close
								SET RST=NOTHING
								%>
							</SELECT>									
							<select name=zuno   style="width:24%;vertical-align:middle"  >
								<OPTION VALUE="">----------</OPTION>
							</SELECT>
							<%SQLN="SELECT * FROM BASICCODE WHERE FUNC='SHIFT' ORDER BY SYS_VALUE "
								  SET RST=CONN.EXECUTE(SQLN)%>
							<SELECT NAME=SHIFT  onkeydown="enterto()" style="width:24%;vertical-align:middle" >
								<OPTION VALUE="" <%IF SHIFT="" THEN %> SELECTED <%END IF%>></OPTION>				
								<%  WHILE NOT RST.EOF 
								%>
								<option value="<%=RST("SYS_TYPE")%>" <%IF SESSION("NETWHSNO")=RST("SYS_TYPE") THEN %> SELECTED <%END IF%>><%=RST("SYS_VALUE")%>-<%=RST("SYS_TYPE")%></option>
								<%
								RST.MOVENEXT
								WEND
								rst.close
								SET RST=NOTHING 														
								%>
							</SELECT>													 
							<INPUT TYPE=HIDDEN  name='grps'>
									
						</TD>								
						<td align=right>職等<br><font class=txt8>Chức vụ</font></td>
						<TD>
							<select name="JOB" >
								<% SQL="SELECT * FROM BASICCODE WHERE FUNC='LEV'  ORDER BY SYS_TYPE "
								SET RST = CONN.EXECUTE(SQL)
								WHILE NOT RST.EOF
								%>
								<option value="<%=RST("SYS_TYPE")%>" <%IF SESSION("NETWHSNO")=RST("SYS_TYPE") THEN %> SELECTED <%END IF%>><%=RST("SYS_type")%>-<%=RST("SYS_VALUE")%></option>
								<%
								RST.MOVENEXT
								WEND
								rst.close
								SET RST=NOTHING 
								conn.close 
								%>
							</SELECT>			
						</TD> 									
					</tr>
					<tr>						
						<td align=right  nowrap>婚姻狀況<br><font class=txt8>TT Hôn Nhân</font></td>
						<td  nowrap colspan="3">
							<INPUT type="radio" id=radio2 name=radio2 onclick=marrychg(0) >已婚<font class=txt8>Đã kết hôn</font>
							&nbsp;<INPUT type="radio" id=radio2 name=radio2 onclick=marrychg(1)> 未婚<font class=txt8>Chưa kết hôn</font>
							&nbsp;<INPUT type="radio" id=radio2 name=radio2 onclick=marrychg(2)> 離婚<font class=txt8>Ly hôn</font>
							<input type=hidden name=marryed value="" size=1>
						</td>
						<td  align=right >教育程度<BR><font class=txt8>Trình độ học vấn</font></td>
						<td ><input type="text" name=school size=15  ></td>		
					</tr>
					<tr>	
						<td align=right  nowrap>*身分証號<BR><font class=txt8>Số CMND</font></td>
						<td ><input type="text" name=personID size=28  ></td>
						<td align=right  ><font color="#990000">*稅號<BR><font class=txt8>Mã số thuế</font></font></td>
						<td ><input type="text" name=taxCode size=28  maxlength=10></td>
						<td nowrap align=right >簽合同日<BR><font class=txt8>Ngày ký hợp đồng</font></td>
						<td ><input type="text" name=BHDAT size=15  onblur="date_change(2)"><font class=txt8>(Ngày)</font></td>
					</tr>
					<tr>
						<td align=right nowrap>發證日期<br><font class=txt8>Ngày cấp</font></td>
						<td ><input type="text" name=passportNo size=22   ></td>		
						<td align=right nowrap>發証地點<br><font class=txt8>Nơi cấp </font></td>
						<td ><input type="text" name=visano size=25  ></td>
						<td nowrap align=right >加入工團<br><font class=txt8>Ngày vào công đoàn</font></td>
						<td ><input type="text" name="GTDAT"  size=15  ONBLUR="CHKVALUE(5)" ><font class=txt8>(ex:200601)</font></td>
					</tr>									
					<tr>
						<td align=right >銀行帳號<br><font class=txt8>Tài khoản ngân hàng </font></td>
						<td ><input type="text" name=bankid size=30  ></td>
						<td align=right nowrap>聯絡電話<BR><font class=txt8>ĐT liên lạc</font></td>
						<td ><input type="text" name=phone size=30  ></td>
						<td align=right >手機<BR><font class=txt8>ĐTDĐ</font></td>
						<td ><input type="text" style="width:98%"  name=mobilephone ></td> 
					</TR>
					<tr>
						<TD align=right style="color:blue" >*保險號碼<br><font class=txt8>Số thẻ bảo hiểm</font></TD>
						<td ><input type="text" name="masobh"  size=30 ></td>
						<td align=right ><font class=txt8>E-MAIL</font></td>
						<td colspan=3><input type="text" name=email size=60 style="width:99%" ></td>
					</tr>	
					<tr>
						<td nowrap align=right>聯絡地址<BR><font class=txt8>Địa chỉ liên lạc</font></td>
						<td colspan=5 ><input type="text" name=homeaddr size=60 style="width:99%" ></td>
					</tr>
					<tr>
						<td nowrap align=right >其他說明<BR><font class=txt8>Ghi chú</font></td>
						<td colspan="5"><input type="text" name=memo size=60  style="width:99%" ></td>		
					</tr>
					<tr ALIGN=center>
						<TD colspan=6 style="height:50px">
							<input type=button  name=send value="(Y)Confirm"  class="btn btn-sm btn-danger" onclick=go()>
							<input type=RESET name=send value="(N)Cancel"  class="btn btn-sm btn-outline-secondary">
						</TD>
					</TR>
					<tr><td colspan=6>&nbsp;</td></tr>
				 </table>
			</td>
		</tr>		
	</table>
			
	
</form>
</body>
</html>
<script language=vbscript>
<!--
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
-->
</script>


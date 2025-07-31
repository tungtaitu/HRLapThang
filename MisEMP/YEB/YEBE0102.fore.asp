<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<!-- #include file = "../GetSQLServerConnection.fun" -->
<!-- #include file="../ADOINC.inc" -->
<!--#include file="../include/sideinfo.inc"-->
<%

SELF = "YEBE0102"

Set conn = GetSQLServerConnection()
'Set rs = Server.CreateObject("ADODB.Recordset")

nowmonth = year(date())&right("00"&month(date()),2)
if month(date())="01" then
	calcmonth = year(date()-1)&"12"
else
	calcmonth =  year(date())&right("00"&month(date())-1,2)
end if

'sql="select max(empid) empid  from empfile where left(empid,1)='L' "
'set rds=conn.execute(sql)
'if not rds.eof then
'	eid = "L" & right("0000" & cstr(cdbl(right(rds("empid"),4))+1) , 4)
'else
'	eid=""
'end if
'set rds=nothing

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
	<%=self%>.indat.focus()
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
-->
</SCRIPT>
</head>
<body  onkeydown="enterto()" onload="f()">
<form  name="<%=self%>" method="post" action="<%=self%>.upd.asp"   >
<INPUT TYPE=HIDDEN NAME="UID" VALUE=<%=SESSION("NETUSER")%>>
<INPUT TYPE=hidden NAME=nowmonth VALUE="<%=nowmonth%>">
<input name=act value="EMPADDNEW" type=hidden >
 
	<table border=0 style="height:70px"><tr><td>&nbsp;</td></tr></table>
	<table width="100%" BORDER=0 align=center cellpadding=0 cellspacing=0 >
		<tr>
			<td align="center">
				<table id="myTableForm" width="98%" border=0>
					<tr><td style="height:30px" colspan=6>&nbsp;</td></tr>
					<TR>
						<td align=right >國籍<br>Quốc tịch</td>
						<td>
							<select name=country onchange=chkempid() style="width:120px">
								<%SQL="SELECT * FROM BASICCODE WHERE FUNC='country' and sys_type<>'VN' ORDER BY SYS_type desc  "
								SET RST = CONN.EXECUTE(SQL)
								WHILE NOT RST.EOF
								%>
								<option value="<%=RST("SYS_TYPE")%>" <%IF SESSION("NETWHSNO")=RST("SYS_TYPE") THEN %> SELECTED <%END IF%>><%=RST("SYS_VALUE")%></option>
								<%
								RST.MOVENEXT
								WEND
								rst.close
								SET RST=NOTHING
								%>
							</SELECT>	
						</td>
						<TD align=right>廠別<br>Xưởng</TD>
						<TD>
							<select name=WHSNO   onchange=chkempid() style="width:120px">
								<option value="">-----</option>
								<%SQL="SELECT * FROM BASICCODE WHERE FUNC='WHSNO' ORDER BY SYS_TYPE "
								SET RST = CONN.EXECUTE(SQL)
								WHILE NOT RST.EOF
								%>
								<option value="<%=RST("SYS_TYPE")%>"  ><%=RST("SYS_VALUE")%></option>
								<%
								RST.MOVENEXT
								WEND
								rst.close
								SET RST=NOTHING
								%>
							</SELECT>										
						</TD>
						<TD align=right nowrap >類別<br>Loại</TD>
						<TD> 
							<select name=emptype   onchange=chkempid() >				
								<%SQL="SELECT * FROM BASICCODE WHERE FUNC='EMPTYPE' ORDER BY SYS_TYPE "
								SET RST = CONN.EXECUTE(SQL)
								WHILE NOT RST.EOF
								%>
								<option value="<%=RST("SYS_TYPE")%>"  ><%=RST("SYS_TYPE")%>.<%=RST("SYS_VALUE")%></option>
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
						<TD align=right nowrap>部門/班別<br>Bộ phận / Ca</TD>
						<TD colspan=3>
							<select name=GROUPID  onchange=groupchg() style="width:30%">
								<option value="" <%if request("GROUPID")="" then%>selected<%end if%>>-----</option>
								<%SQL="SELECT * FROM BASICCODE WHERE FUNC='GROUPID' and sys_type like 'A06%' ORDER BY SYS_TYPE "
								SET RST = CONN.EXECUTE(SQL)
								WHILE NOT RST.EOF
								%>
								<option value="<%=RST("SYS_TYPE")%>" ><%=RST("SYS_TYPE")%>-<%=RST("SYS_VALUE")%></option>
								<%
								RST.MOVENEXT
								WEND
								SET RST=NOTHING
								%>
							</SELECT>
							<select name=zuno  style="width:30%">
								<OPTION VALUE="">-----</OPTION>
							</SELECT>
							<SELECT NAME=SHIFT  onkeydown="enterto()"  style="width:30%">								
								<%SQL="SELECT * FROM BASICCODE WHERE FUNC='shift'   ORDER BY len(sys_type) desc, SYS_TYPE "
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
						<TD  align=right>處/所<br>Khu</TD>
						<TD >
							<select name=unitno   onchange=unitchg()>
								<option value="">-----</option>
								<%SQL="SELECT * FROM BASICCODE WHERE FUNC='unit' and sys_type<>'AAA' ORDER BY SYS_TYPE desc "
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
					</TR>
					<TR>
						<TD align=RIGHT nowrap>工號<br>Mã số thẻ</TD>
						<TD nowrap><INPUT type="text" NAME="EMPID" SIZE=12   maxlength=5 value="<%=request("empid")%>"  readonly ></TD>
						<TD align=right>姓名(中)<br>Họ tên(Hoa)</TD>
						<TD ><INPUT type="text" NAME=nam_cn SIZE=20   ></TD>	
						<TD align=right >姓名(英)<br>Họ tên(En)</TD>
						<TD ><INPUT type="text" style="width:250px" NAME=nam_vn ></TD>
					</TR>
					<TR>
						<TD align=right>身分證號<br>Số CMND</TD>
						<td><input type="text" name=personID ></td>
						<TD align=right><font color="#990000">稅號<br>Mã số thuế</font></TD>
						<td><input type="text" name="taxCode" size=25  ></td>
						<TD align=right >職等<br>Chức vụ</TD>
						<TD>
							<select name=JOB   >
								<%SQL="SELECT * FROM BASICCODE WHERE FUNC='LEV'  ORDER BY SYS_TYPE "
								SET RST = CONN.EXECUTE(SQL)
								WHILE NOT RST.EOF
								%>
								<option value="<%=RST("SYS_TYPE")%>" <%IF SESSION("NETWHSNO")=RST("SYS_TYPE") THEN %> SELECTED <%END IF%>><%=RST("SYS_TYPE")%>-<%=RST("SYS_VALUE")%></option>
								<%
								RST.MOVENEXT
								WEND
								rst.close
								SET RST=NOTHING
								%>
							</SELECT>
							<% 
							conn.close
							set conn=nothing 			
							%>
						</TD>						
					</TR>
					<TR>
						<TD align=right >出生日期<br>Ngày sinh</TD>
						<TD colspan=3>
							<INPUT type="text" NAME=BYY SIZE=5   MAXLENGTH=4 ONBLUR="CHKVALUE(1)" style="width:15%">年 Năm
							<INPUT type="text" NAME=BMM SIZE=3   MAXLENGTH=2 ONBLUR="CHKVALUE(2)"  style="width:10%">月Tháng
							<INPUT type="text" NAME=BDD SIZE=3   MAXLENGTH=2 ONBLUR="CHKVALUE(3)"  style="width:10%">日Ngày
							<input type="hidden" name=ages size=5  ONBLUR="CHKVALUE(4)" >									
						</TD>
						<TD  align=RIGHT >到職日<br>NVX</TD>
						<TD ><INPUT type="text" NAME=indat SIZE=12  value=<%=fdt(date())%> onblur="date_change(1)"></TD>
					</TR>
					<tr>
						<td align=right>性別<br>GiỚi Tính</td>
						<td>
							<INPUT type="radio" id=radio1 name=radio1 onclick=sexchg(0) > 男 Nam &nbsp;
							<INPUT type="radio" id=radio1 name=radio1 onclick=sexchg(1)> 女 Nữ
							<input type=hidden name=sexstr value="<%=sex%>" size=1>
						</td>
						<td  align=right nowrap>婚姻狀況<br>TT Hôn Nhân</td>
						<td valign=top nowrap>
							<INPUT type="radio" id=radio2 name=radio2 onclick=marrychg(0) >已婚<font class="txt8">Đã kết hôn</font>
							&nbsp;<INPUT type="radio" id=radio2 name=radio2 onclick=marrychg(1)> 未婚<font class="txt8">Chưa kết hôn</font>
							&nbsp;<INPUT type="radio" id=radio2 name=radio2 onclick=marrychg(2)> 離婚<font class="txt8">Ly hôn</font>
							<input type=hidden name=marryed value="" size=1>
						</td>
						<td  align=right nowrap >教育程度<br>Học lực</td>
						<td ><input type="text" name=school size=15  ></td>
					</tr>
					<tr>		
						<td nowrap align=right >護照號碼<br>So ho chieu</td>
						<td><input type="text" name=PASSPORTNO></td>	
						<td  align=right   nowrap>護照簽發日<br>Ngày cấp</td>
						<td ><input type="text" name=pissuedate size=15  onblur="date_change(2)"></td>
						<td   align=right >護照到期日<br>Ngay hết hạn</td>
						<td ><input type="text" name=pduedate size=15  onblur="date_change(3)"></td>
					</tr>
					<tr>		
						<td nowrap align=right >簽証號碼<br>So VISA</td>
						<td><input type="text" name=WKD_No ></td>		
						<td nowrap  align=right >工作證到期<br>Ngay hết hạn</td>
						<td ><input type="text" name=WKD_DueDate size=15  onblur="date_change(4)"></td>
						<td  align=right >經歷<br>Trình Độ：</td>
						<td><input type="text" style="width:98%" name=experience  ></td>	
					</tr>
					<tr>
						<td  align=right>備註說明<br>Ghi chú</td>
						<td colspan=5><input  type="text"  name=memo  style="width:99%"></td>		
					</tr>	 
					<tr height=25 bgcolor=#FFD9FF><td colspan=6>----- 廠內聯絡資訊 THÔNG TIN LIÊN LẠC TRONG CÔNG TY------------------------------------</td></tr>
					<tr>
						<td nowrap align=right height=25 >聯絡電話<br>CTY Đ.T</td>
						<td><input type="text" name=phone ></td>
						<td nowrap align=right >手機<BR>DTDD</td>
						<td ><input type="text" name=mobilephone size=25  ></td>
						<td nowrap align=right>E-MAIL<br></td>
						<td ><input type="text" name=email style="width:98%"></td>
					</tr>
					<tr>
						<td nowrap align=right height=25 >銀行帳號<br>Số thẻ ngân hàng</td>
						<td colspan=5><input type="text" name=bankid style="width:99%"></td>
					</tr>
					<tr height=25 bgcolor=#EAFEAD><td colspan=6>----- 國內聯絡資訊 THÔNG TIN LIÊN LẠC QUỐC NỘI-----------------------</td></tr>
					<tr height=25>
						<td nowrap align=right   >姓名<br>Họ tên</td>
						<td><input type="text" name=urgent_person></td>
						<td nowrap align=right >關係<br>Quan hê</td>
						<td ><input type="text" name=releation size=10  ></td>
						<td nowrap align=right height=25 >聯絡電話<Br>Số ĐT</td>
						<td><input type="text" name=urgent_phone size=30  ></td>
					</tr>	
					<tr height=25>						
						<td nowrap align=right >手機<br>ĐTDĐ</td>
						<td ><input type="text" name=urgent_mobile ></td>
						<td nowrap align=right height=25 >保險受益人<br>Người hưởng bảo hiểm</td>
						<td><input type="text" name=bh_person size=30  ></td>
						<td nowrap align=right >受益人<br>身分證號<br>CMND Người hưởng thụ </td>
						<td ><input type="text" name=bh_personID ></td>
					</tr>	
					<tr height=25>
						<td nowrap align=right height=25 >聯絡地址<br>Địa chỉ liên lạc</td>
						<td colspan=5><input type="text" name=urgent_addr style="width:99%"  ></td>
					</tr>	
					<tr >
						<TD ALIGN=center valign="center" style="height:50px" colspan=6>
							<input type=button  name=send value="(Y)Confirm"  class="btn btn-sm btn-danger" onclick=go()>
							<input type=RESET name=send value="(N)Cancel"  class="btn btn-sm btn-outline-secondary">
						</TD>
					</TR>
					<tr><td style="height:30px" colspan=6>&nbsp;</td></tr>
				</table>
			</td>
		</tr>		
	</table>
			
	
</form>
</body>
</html>
<script language=vbscript>
<!--
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
	chkempid
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
	photosname=<%=self%>.empid.value&".jpg"
	<%=SELF%>.ACTION="<%=self%>.upd.asp"
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


<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<!-- #include file = "../GetSQLServerConnection.fun" -->
<!-- #include file="../ADOINC.inc" -->
<%

SELF = "YEBQ01B"

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
Uid =trim(request("uid"))

if request("uid")=""  then
  
SQL="SELECT * FROM  view_empfile where ISNULL(STATUS,'')<>'D' AND  autoid='"& empautoid &"' "
else

SQL="SELECT * FROM  view_empfile where ISNULL(STATUS,'')<>'D' AND  empid='"& Uid &"' "

end if
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
	
	'PHOTOS=TRIM(RS("PHOTOS"))	'照片檔名 
	PHOTOS=RS("EMPID")&".JPG"
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
	
	
END IF
rs.close
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
<link rel="stylesheet" href="../Include/bar.css" type="text/css">
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
	<%=self%>.emptype.focus()
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
</HEAD> 
<body  topmargin="5" leftmargin="5"  marginwidth="0" marginheight="0"  onkeydown="enterto()"  >
<form  name="<%=self%>" method="post" action="YEBE0102.upd.asp"   >
<INPUT TYPE=HIDDEN NAME="UID" VALUE=<%=SESSION("NETUSER")%>>
<INPUT TYPE=hidden NAME=nowmonth VALUE="<%=nowmonth%>">
<input name=act value="EMPEDIT" type=hidden >
<table width="460" border="0" cellspacing="0" cellpadding="0">
	<tr><TD  >
		<img border="0" src="../image/icon.gif" align="absmiddle">
		員工個人資料查詢
		</TD>
	</tr>
</table>
<hr size=0	style='border: 1px dotted #999999;' align=left width=550>  
<table width="650" border="0" cellspacing="0" cellpadding="0" align=center >
	<% if (REQUEST("empautoid") <>"" ) then %> 
			<tr><td nowrap>
		<div id="navcontainer"  >
			<ul id="navlist">
			<li  id=active><a href="vbscript:chgpage(1)">基本資料<BR>Tu lieu co ban<BR>&nbsp;</a></li>
			<li > <a href=" vbscript:chgpage(2) ">教育訓練/証執照<br>huan luyen/<BR>bang cap</a></li>			
			<li><a href=" vbscript:chgpage(4) ">獎懲紀錄<BR>Tu lieu<BR>thuong phat</a></li>
			<li><a href=" vbscript:chgpage(5)">部門/晉升紀錄<BR>Nang chuc/<BR>don vi </a></li>
			<%if session("rights")<="1"  or session("netuser")="PELIN"  then %>
				<li ><a href="<% if (REQUEST("empautoid") <>"" ) then %>vbscript:chgpage(6)<%end if %>">薪資資料<BR>Tien luong<BR>&nbsp;</a></li>
			<%else%>
				<li ><a >薪資資料<BR>Tien luong<BR>&nbsp;</a></li>
			<%end if%>	
			</ul>
		</div> 
		</td>
	</tr>  
	<% end if %>
</table>   
<hr size=0	style='border: 1px dotted #999999;' align=left width=550> 

<TABLE WIDTH=550 CLASS=txt BORDER=0  cellspacing="2" cellpadding="2">
	<TR height=25 >
		<TD width=90 nowrap align=right >工號<br><font class=txt8>So The</font></TD>
		<TD >
			<%if left(empid,1)="S" then %>
				<select name=emptype class=txt>					
					<option value="C" <%if emptype="C" then%>selected<%end if%>>C.代銷</option>
				</select>
			<%else%>
				<select name=emptype class=txt disabled >
					<option value=""></option>
					<option value="A" <%if emptype="A" then%>selected<%end if%>>A.員工</option>
					<option value="B" <%if emptype="B" then%>selected<%end if%>>B.出差</option>					
				</select>
			<%end if%>
			<INPUT NAME=EMPID SIZE=6 CLASS=READONLY VALUE="<%=EMPID%>" READONLY   >
			<INPUT type=hidden NAME=empautoid   VALUE="<%=empautoid%>"     >
		</TD>
		<TD width=80 nowrap align=right height=20>到職日<br><font class=txt8>NVX</font></TD>
		<TD ><INPUT NAME=indat SIZE=12 CLASS=READONLY READONLY  VALUE="<%=(indat)%>"    ></TD>
		<td width="110"  align=center valign=top  nowrap  rowspan="5" >
			<img src="pic/<%=EMPID%>.jpg"  border=1 width=108 height=140  >
		</td>
	</TR>
	<TR height=25 >
		<TD   align=right>國籍<br><font class=txt8>Quoc Tich</font></TD>
		<TD >
			<select name=country  class=inputbox  onkeydown="enterto()" disabled  >
				<%SQL="SELECT * FROM BASICCODE WHERE FUNC='country'  ORDER BY SYS_type desc  "
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
		<td   align=right >離職日<br><font class=txt8>NTV</font></td>
		<td ><input name=outdat size=12 class=READONLY READONLY    onblur="date_change(5)"  value=<%=OUTDATe%>  ></td>		
	</tr>	
	<TR height=25 >
		<TD   align=right>廠/處<br><font class=txt8>Xuong/Khu</font></TD>
		<TD >
			<%if cdate(indat)>=cdate(firstday) then %>
				<select name=WHSNO  class=font9 onkeydown="enterto()" disabled  >
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
			<%else%>	
				<input type=hidden name=whsno size=10 class='readonly' readonly value="<%=whsno%>">				
				<input name=wstr size=8 class='readonly' readonly value="<%=wstr%>">
				<input name=ustr size=6 class='readonly' readonly value="<%=ustr%>">
			<%end if%>
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
					rst.close
					%>
				</SELECT>
				<%SET RST=NOTHING %>
			<%else%>
				<input type=hidden name=unitno   class='readonly' readonly value="<%=unitno%>">
				
			<%end if%>			
		</TD>
		<TD   align=right >部門/組<br><font class=txt8>bo phan</font></TD>
		<TD >
			<%if cdate(indat)>=cdate(firstday) then %>
				<select name=GROUPID  class=font9 disabled style="width:60"  >
					<%SQL="SELECT * FROM BASICCODE WHERE FUNC='GROUPID' and sys_type <>'AAA' ORDER BY SYS_TYPE "
					SET RST = CONN.EXECUTE(SQL)
					'RESPONSE.WRITE SQL
					WHILE NOT RST.EOF
					%>
					<option value="<%=RST("SYS_TYPE")%>" <%IF RST("SYS_TYPE")=TRIM(GROUPID) THEN %> SELECTED <%END IF%>><%=RST("SYS_VALUE")%></option>
					<%
					RST.MOVENEXT
					WEND
					rst.close
					%>
				</SELECT>
				<%SET RST=NOTHING %>
				<select name=zuno  class=font9 style='width:50' onkeydown="enterto()" disabled  >
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
					rst.close
					%>
				</SELECT>
				<%SET RST=NOTHING %>
			<%else%>	
				<input type=hidden name=groupid size=10 class='readonly' readonly value="<%=groupid%>">
				<input type=hidden name=zuno size=10 class='readonly' readonly value="<%=zuno%>">
				<input name=gstr size=5 class='readonly' readonly value="<%=gstr%>">
				<input name=zstr size=7 class='readonly' readonly value="<%=zstr%>">
			<%end if%>			
		</TD>
	</tr>
	<TR height=25 >
		<TD   align=right >職等<br><font class=txt8>Chuc vu</font></TD>
		<TD  >
		<%if cdate(indat)>=cdate(firstday) then %>
				<select name=JOB  class=font9 style='width:75'    disabled >
				<%SQL="SELECT * FROM BASICCODE WHERE FUNC='LEV'  ORDER BY SYS_TYPE "
					SET RST = CONN.EXECUTE(SQL)
					WHILE NOT RST.EOF
				%>
					<option value="<%=RST("SYS_TYPE")%>" <%IF RST("SYS_TYPE")=JOB THEN %> SELECTED <%END IF%>><%=RST("SYS_VALUE")%></option>
				<%
					RST.MOVENEXT
					WEND
					rst.close
				%>
				</SELECT>
				<%SET RST=NOTHING %>
			<%else%>	
				<input type=hidden name=job size=10 class='readonly' readonly value="<%=Job%>">
				<input name=jstr size=20 class='readonly' readonly value="<%=Jstr%>">
			<%end if%>
		</TD>	
		<TD align=right >班別<br><font class=txt8>Ca</font></TD>
		<TD >
			<%if cdate(indat)>=cdate(firstday) then %>
				<SELECT NAME=SHIFT CLASS=font9 disabled >
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
			<select name='grps'  class=font9 style='width:60' onkeydown="enterto()"  disabled >
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
			<%SET RST=NOTHING 
			conn.close
			set conn=nothing
			%>
		</TD> 
	</TR>
	<TR   >
		<td  align=right>身分證號<br><font class=txt8>So CMND</font></td>
		<TD ><input name=personID size=20 class=readonly readonly  VALUE="<%=PERSONID%>" ></td>
		<td  align=right>性別<br><font class=txt8>GiỚi Tính</font></td>
		<td>
			<INPUT type="radio" id=radio1 name=radio1 onclick=sexchg(0) <%IF SEX="M" THEN %>CHECKED<%END IF%> disabled  > 男 &nbsp;
			<INPUT type="radio" id=radio1 name=radio1 onclick=sexchg(1) <%IF SEX="F" THEN %>CHECKED<%END IF%>  disabled> 女
			<input type=hidden name=sexstr value="<%=SEX%>" size=1>			
		</TD>
	</TR>	
	<TR   >
		<TD   align=right>姓名(中)<br><font class=txt8>Ho Ten(Hoa)</font></TD>
		<TD  >
			<INPUT NAME=nam_cn SIZE=20 CLASS=readonly readonly VALUE="<%=EMPNAM_CN%>" onkeydown="enterto()" >
		</TD>
		<TD   align=right >姓名(越)<br><font class=txt8>Ho Ten(viet)</font></TD>
		<TD colspan=2 ><INPUT NAME=nam_vn SIZE=35 CLASS=readonly readonly  VALUE="<%=EMPNAM_VN%>" onkeydown="enterto()" ></td> 
	</TR>

	<TR   >
		<TD   align=right >出生日期<br><font class=txt8>Ngay Sinh</font></TD>
		<TD colspan=4>
			<INPUT NAME=BYY SIZE=5 CLASS=readonly  readonly  VALUE="<%=BYY%>" MAXLENGTH=4   > 年 Năm&nbsp;&nbsp;
			<INPUT NAME=BMM SIZE=3 CLASS=readonly readonly VALUE="<%=BMM%>" MAXLENGTH=2  > 月 Tháng&nbsp;&nbsp;
			<INPUT NAME=BDD SIZE=3 CLASS=readonly readonly VALUE="<%=BDD%>" MAXLENGTH=2  > 日 Ngày&nbsp;&nbsp;
			<input type=hidden name=ages size=5 class=inputbox VALUE="<%=AGES%>"  >
		</TD>
	</TR>
 	<tr>
		<Td align=right>婚姻<br><font class=txt8>TT Hôn Nhân</font></td>
		<TD colspan=4>
			<INPUT type="radio" id=radio2 name=radio2 onclick=marrychg(0)  <%IF marryed="Y" THEN %>CHECKED<%END IF%> disabled> 已婚 <font class=txt8>Da KH</font>
			&nbsp;<INPUT type="radio" id=radio2 name=radio2 onclick=marrychg(1)  <%IF marryed="N" THEN %>CHECKED<%END IF%> disabled> 未婚 <font class=txt8>Chua KH</font>
			&nbsp;<INPUT type="radio" id=radio2 name=radio2 onclick=marrychg(2)  <%IF marryed="L" THEN %>CHECKED<%END IF%> disabled> 離婚 <font class=txt8>Ly Hon</font>
			<input type=hidden name=marryed value="<%=marryed%>" size=1>  
		</td>	
	</tr>	

</TABLE> 
<hr size=0	style='border: 1px dotted #999999;' align=left width=550> 
<TABLE WIDTH=550 CLASS=txt  BORDER=0  cellspacing="2" cellpadding="3"> 	 
	<tr height=20 > 	
		<td  align=right width=85 >教育程度<br><font class=txt8>Học lực</font></td>
		<td ><input name=school size=15 class=readonly readonly VALUE="<%=SCHOOL%>" ></td>
		<td   align=right >合同日<br><font class=txt8>Ky Hop Dong</font></td>
		<td ><input name=BHDAT size=15 class=readonly readonly VALUE="<%=bhdat%>"  onblur="date_change(6)" ><font class=txt8>(Ngay)</font></td>		
	</tr>
	<tr height=20 >		
		<td align=right >發證日期<br><font class=txt8>Ngày cấp</font></td>
		<td><input name=PASSPORTNO size=25 class=readonly readonly VALUE="<%=PASSPORTNO%>" ></td>		
		<td nowrap align=right height=20 >加入工團<br><font class=txt8>Ngay Nhap C.Đ</font></td>
		<td  ><input name=GTDAT size=10 class=readonly readonly VALUE="<%=GTDAT%>"></td>
		
	</tr>  
	<!--tr>
		<td  align=right height=20 >上傳照片：</td>
		<td colspan=3>
			<INPUT TYPE="FILE" NAME="FILE1" SIZE="50" class=inputbox> (size:100*130,*.jpg)
		</td>		
	</tr-->			
	<tr height=20>
		<td align=right height=25 >發証地點<br><font class=txt8>Nơi cấp </font></td>
		<td ><input name=visaNo size=15 class=readonly readonly  VALUE="<%=visaNo%>" ></td>
		<td nowrap align=right height=20 >聯絡電話<br><font class=txt8>Đ.T</font></td>
		<td><input name=phone size=20 class=readonly readonly VALUE="<%=phone%>"></td>		
	</tr>
	<tr height=20 >
		<td nowrap align=right >手機<br><font class=txt8>ĐTDD</font></td>
		<td ><input name=mobilephone size=25 class=readonly readonly VALUE="<%=MOBILEPHONE%>"></td>
		<td nowrap align=right height=20 >E-MAIL<br></td>
		<td  ><input name=email size=25 class=readonly readonly VALUE="<%=email%>"></td>
		
	</tr>
	<tr>
		<td nowrap align=right height=20 >聯絡地址<br><font class=txt8>Địa chi</font></td>
		<td colspan=3><input name=homeaddr size=65 class=readonly readonly VALUE="<%=Homeaddr%>"></td>
	</tr> 
	<tr>
		<td nowrap align=right height=20 >銀行帳號<br><font class=txt8>SO Tai Khoan </font></td>
		<td colspan=3><input name=bankid size=55 class=readonly readonly VALUE="<%=bankid%>"></td>
	</tr> 	
	<tr>
		<td  align=right height=20 >備註說明<br><font class=txt8>Ghi Chu</font></td>
		<td colspan=3>
			<input   name=memo class=readonly readonly size=65 VALUE="<%=memo%>">
		</td>		
	</tr>		
</table>
<hr size=0	style='border: 1px dotted #999999;' align=left width=550> 
 
<TABLE WIDTH=550>
		<tr ALIGN=center>
			<TD >			
			<input type=RESET name=send value="關閉視窗(Close)"  class=button onclick=window.close()>
			</TD> 
		</TR>
</TABLE>


</form>


</body>
</html>
		
<script language=vbscript>
<!-- 
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

function chgpage(a)
	code1=<%=self%>.empautoid.value
	code2=<%=self%>.empid.value
	if a = 1 then 
		if <%=self%>.country.value="VN" then 
			open "<%=self%>.editvn.asp?empautoid="& code1 & "&empid=" & code2 , "_self"
		else
			open "<%=self%>.editHW.asp?empautoid="& code1 & "&empid=" & code2 , "_self"
		end if	
	elseif a=2 then 
		open "<%=self%>.Fore3.asp?empautoid="& code1 & "&empid=" & code2 , "_self"
	elseif a=3 then 
		open "<%=self%>.Fore3.asp?empautoid="& code1 & "&empid=" & code2 , "_self"
	elseif a=4 then 
		open "<%=self%>.Fore4.asp?empautoid="& code1 & "&empid=" & code2 , "_self"
	elseif a=5 then 
		open "<%=self%>.Fore5.asp?empautoid="& code1 & "&empid=" & code2 , "_self"
	elseif a=6 then 
		open "<%=self%>.Fore6.asp?empautoid="& code1 & "&empid=" & code2 , "_self"
	else
	end if 
	
end function 
-->
</script>


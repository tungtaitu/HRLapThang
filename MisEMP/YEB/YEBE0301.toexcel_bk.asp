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
  
SQL="SELECT * FROM    view_employee where ISNULL(STATUS,'')<>'D' AND  autoid='"& empautoid &"' "

RESPONSE.WRITE SQL
RESPONSE.END
RS.OPEN SQL , CONN, 3, 3
IF NOT RS.EOF THEN
	empautoid = TRIM(RS("AUTOID"))
	emptype = TRIM(RS("emptype"))
	EMPID=TRIM(RS("EMPID"))	'員工編號
	INDAT=TRIM(RS("date1"))	'到職日
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
	JOB=TRIM(RS("newJOB"))  '職等 
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
	MARRYED = RS("MARRYED")    '婚姻狀況
	SCHOOL=RS("SCHOOL") '教育程度
	SHIFT=RS("SHIFT") '班別
	grps = rs("grps") 
	studyjob=RS("studyjob") '職能學習 
	
	'PHOTOS=TRIM(RS("PHOTOS"))	'照片檔名 
	PHOTOS=RS("EMPID")&".JPG"
	'-----------------------------------------
	PHU = RS("PHU")
	NN = RS("NN")
	KT = RS("KT")
	TTKH = RS("TTKH")
	MT = RS("MT")
	BB = RS("BB")
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
 
</head>
<body  topmargin="5" leftmargin="1"  marginwidth="0" marginheight="0"   > 
<%
  filenamestr = "salaryVn"&yymm&".xls"
  Response.AddHeader "content-disposition","attachment; filename=" & filenamestr
  Response.Charset ="BIG5"
  Response.ContentType = "Content-Language;content=zh-tw" 
  Response.ContentType = "application/vnd.ms-excel"
%>
<table width="460" border="0" cellspacing="0" cellpadding="0">
	<tr><TD  >
		<img border="0" src="../image/icon.gif" align="absmiddle">
		海外幹部個人資料
		</TD>
	</tr>
</table>
<hr size=0	style='border: 1px dotted #999999;' align=left width=550> 

<TABLE WIDTH=550 CLASS=txt BORDER=0  cellspacing="0" cellpadding="0">
	<TR height=25 >
		<TD width=90 nowrap align=right >工號：</TD>
		<TD ><%=EMPID%>
			
		</TD>
		<TD width=80 nowrap align=right height=20>到職日：</TD>
		<TD ><%=(indat)%></TD>
		<td width="110"  align=center valign=middle  nowrap  rowspan="6" ><img src="pic/<%=EMPID%>.jpg"  border=1> </td>
	</TR>
	<TR height=25 >
		<TD   align=right>國籍：</TD>
		<TD><%=COUNTRY%>
			 
		</TD>	
		<td   align=right >離職日：</td>
		<td ><%=OUTDATe%></td>		
	</tr>	
	<TR height=25 >
		<TD   align=right>廠/處：</TD>
		<TD ><%=wstr%><%=ustr%>
			 		
		</TD>
		<TD   align=right     >部門/組：</TD>
		<TD ><%=gstr%><%=zstr%>
			 	
		</TD>
	</tr>
</table>	
	<TR height=25 >
		<TD   align=right >職等：</TD>
		<TD  >
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
					rst.close
				%>
				</SELECT>
				<%SET RST=NOTHING %>
			<%else%>	
				<input type=hidden name=job size=10 class='readonly' readonly value="<%=Job%>">
				<input name=jstr size=20 class='readonly' readonly value="<%=Jstr%>">
			<%end if%>
		</TD>	
		<TD align=right >班別：</TD>
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
			<select name='grps'  class=font9 style='width:60' onkeydown="enterto()"  >
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
 
	<TR height=25 >
		<TD   align=right>姓名(中)：</TD>
		<TD  >
			<INPUT NAME=nam_cn SIZE=20 CLASS=INPUTBOX VALUE="<%=EMPNAM_CN%>" onkeydown="enterto()" >
		</TD>
		<TD   align=right >姓名(英)：</TD>
		<TD ><INPUT NAME=nam_vn SIZE=22 CLASS=INPUTBOX VALUE="<%=EMPNAM_VN%>" onkeydown="enterto()" ></td> 
	</TR>
	<TR height=28 >
		<td  align=right>身分證號：</td>
		<TD colspan=3>
			<input name=personID size=30 class=inputbox VALUE="<%=PERSONID%>" >&nbsp;&nbsp;&nbsp;&nbsp;婚姻： 
			<INPUT type="radio" id=radio2 name=radio2 onclick=marrychg(0)  <%IF MARRYED="Y" THEN %>CHECKED<%END IF%>> 已婚 &nbsp;
			<INPUT type="radio" id=radio2 name=radio2 onclick=marrychg(1)  <%IF MARRYED="N" THEN %>CHECKED<%END IF%>)> 未婚
			<input type=hidden name=marryed value="<%=marryed%>" size=1>  
		</TD>
	</TR>
	<TR height=28 >
		<TD   align=right >出生日期：</TD>
		<TD colspan=4>
		<INPUT NAME=BYY SIZE=5 CLASS=INPUTBOX VALUE="<%=BYY%>" MAXLENGTH=4 ONBLUR="CHKVALUE(1)" onkeydown="enterto()" > 年
		<INPUT NAME=BMM SIZE=3 CLASS=INPUTBOX VALUE="<%=BMM%>" MAXLENGTH=2 ONBLUR="CHKVALUE(2)" onkeydown="enterto()" > 月
		<INPUT NAME=BDD SIZE=3 CLASS=INPUTBOX VALUE="<%=BDD%>" MAXLENGTH=2 ONBLUR="CHKVALUE(3)" onkeydown="enterto()" > 日&nbsp;&nbsp;
		年齡： <input name=ages size=5 class=inputbox VALUE="<%=AGES%>" ONBLUR="CHKVALUE(4)" onkeydown="enterto()" > &nbsp;
			<INPUT type="radio" id=radio1 name=radio1 onclick=sexchg(0) <%IF SEX="M" THEN %>CHECKED<%END IF%>  > 男 &nbsp;
			<INPUT type="radio" id=radio1 name=radio1 onclick=sexchg(1) <%IF SEX="F" THEN %>CHECKED<%END IF%> > 女
			<input type=hidden name=sexstr value="<%=SEX%>" size=1>
					 
		</TD>
	</TR>
</TABLE> 
<hr size=0	style='border: 1px dotted #999999;' align=left width=550> 
<TABLE WIDTH=550 CLASS=txt BORDER=0> 	 
	<tr height=20 > 	
		<td  align=right width=85 >教育程度：</td>
		<td ><input name=school size=15 class=inputbox VALUE="<%=SCHOOL%>" ></td>
		<td  align=right >護照簽發日：</td>
		<td ><input name=pissuedate size=15 class=inputbox onblur="date_change(2)" VALUE="<%=pissuedate%>"></td>
	</tr>
	<tr height=20 >		
		<td nowrap align=right >護照號碼：</td>
		<td><input name=PASSPORTNO size=25 class=inputbox VALUE="<%=PASSPORTNO%>" ></td>		
		<td   align=right >護照到期日：</td>
		<td ><input name=pduedate size=15 class=inputbox onblur="date_change(3)" VALUE="<%=PDUEDATE%>" ></td>
	</tr>
	<tr height=20 >		
		<td nowrap align=right >工作證號碼：</td>
		<td><input name=WKD_No size=25 class=inputbox VALUE="<%=wkd_no%>"></td>		
		<td   align=right >工作證到期：</td>
		<td ><input name=WKD_DueDate size=15 class=inputbox onblur="date_change(4)" VALUE="<%=wkd_duedate%>"></td>
	</tr>	
	<tr>
		<td  align=right height=20 >經歷：</td>
		<td colspan=3>
			<input  name=experience class="INPUTBOX" size=60 VALUE="<%=experience%>" >
		</td>		
	</tr>
	<tr>
		<td  align=right height=20 >備註說明：</td>
		<td colspan=3>
			<input   name=memo class="INPUTBOX" size=60 VALUE="<%=memo%>">
		</td>		
	</tr>	
	<!--tr>
		<td  align=right height=20 >上傳照片：</td>
		<td colspan=3>
			<INPUT TYPE="FILE" NAME="FILE1" SIZE="50" class=inputbox> (size:100*130,*.jpg)
		</td>		
	</tr-->		
	<tr height=20 bgcolor=#FFD9FF><td colspan=4>----- 廠內聯絡資訊 ------------------------------------</td></tr>
	<tr height=20>
		<td nowrap align=right height=20 >聯絡電話：</td>
		<td><input name=phone size=30 class=inputbox VALUE="<%=phone%>"></td>
		<td nowrap align=right >手機：</td>
		<td ><input name=mobilephone size=25 class=inputbox VALUE="<%=MOBILEPHONE%>"></td>
	</tr>
	<tr height=20 >
		<td nowrap align=right height=20 >E-MAIL：</td>
		<td colspan=3><input name=email size=55 class=inputbox VALUE="<%=email%>"></td>
	</tr>
	<tr>
		<td nowrap align=right height=20 >銀行帳號：</td>
		<td colspan=3><input name=bankid size=55 class=inputbox VALUE="<%=bankid%>"></td>
	</tr>
	<tr height=20 bgcolor=#EAFEAD><td colspan=4>----- 國內聯絡資訊 ------------------------------------</td></tr>
	<tr height=20>
		<td  align=right   >姓名：</td>
		<td><input name=urgent_person size=30 class=inputbox  VALUE="<%=urgent_person%>"></td>
		<td  align=right >關係：</td>
		<td ><input name=releation size=10 class=inputbox  VALUE="<%=releation%>"></td>
	</tr>	
	<tr height=20>
		<td nowrap align=right height=20 >聯絡電話：</td>
		<td><input name=urgent_phone size=30 class=inputbox  VALUE="<%=urgent_tel%>"></td>
		<td nowrap align=right >手機：</td>
		<td ><input name=urgent_mobile size=25 class=inputbox  VALUE="<%=urgent_mobile%>" ></td>
	</tr>	
	<tr height=20>
		<td nowrap align=right height=20 >聯絡地址：</td>
		<td colspan=3><input name=urgent_addr size=55 class=inputbox VALUE="<%=urgent_addr%>" ></td>
	</tr>	
	<tr height=20>
		<td nowrap align=right height=20 >保險受益人：</td>
		<td><input name=bh_person size=30 class=inputbox VALUE="<%=bh_person%>" ></td>
		<td nowrap align=right >受益人ID：</td>
		<td ><input name=bh_personID size=25 class=inputbox VALUE="<%=bh_personID%>" ></td>
	</tr>	

</table>
 


 
</body>
</html> 


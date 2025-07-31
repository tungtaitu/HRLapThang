<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<!-- #include file = "../GetSQLServerConnection.fun" -->
<!-- #include file="../ADOINC.inc" -->
<%


SELF = "YEBE0301TR"

Set conn = GetSQLServerConnection()
Set rs = Server.CreateObject("ADODB.Recordset")


firstday = year(date())&"/"&right("00"&month(date()),2)&"/01"
nowmonth = year(date())&right("00"&month(date()),2)
if month(date())="01" then
	calcmonth = year(date()-1)&"12"
else
	calcmonth =  year(date())&right("00"&month(date())-1,2)
end if

empautoid = TRIM(REQUEST("empid"))



If Request.ServerVariables("REQUEST_METHOD") = "POST" and  request("flag")="S" then  
	
	response.write  request.Form("empid")&"<BR>"
	response.write  request.Form("whsno")&"<BR>"
	
  end if  
  


SQL="select a.*, isnull(b.whsno_acc,'') whsno_acc from "&_
		"(SELECT * FROM  view_empfile where ISNULL(STATUS,'')<>'D' AND  empid='"& empautoid &"' ) a "&_
		"left join ( select * from empfile_acc ) b on b.empid=a.empid  "

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
	if request("GROUPID")="" then 
		GROUPID=TRIM(RS("GROUPID"))	'組/部門
	else
		GROUPID=request("GROUPID")
	end if 
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
	
	whsno_acc = rs("whsno_acc")

	filename=Server.MapPath("pic/"&rs("empid")&".jpg")
	pass_filename = Server.MapPath("ppvisa/"&rs("empid")&"_pass.pdf")
	visa_filename = Server.MapPath("ppvisa/"&rs("empid")&"_visa.pdf")
	'If fso.FileExists(filename) Then
	'	photoYN="Y"
	'else
	'	photoYN="N"
	'end if 
	'If fso.FileExists(pass_filename) Then
	'	passportYN="Y"
	'else
	'	passportYN="N"
	'end if	
	'If fso.FileExists(visa_filename) Then
	'	visaYN="Y"
	'else
	'	visaYN="N"
	'end if		

END IF
rs.close : SET RS=NOTHING



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

</HEAD>
<body  topmargin="5" leftmargin="5"  marginwidth="0" marginheight="0"  onkeydown="enterto()" onload='f()' >
<form  name="<%=self%>" method="post" action="YEBE0102.upd.asp"   >
<INPUT TYPE="hidden" NAME="UID" VALUE=<%=SESSION("NETUSER")%>>
<INPUT TYPE="hidden" NAME="nowmonth" VALUE="<%=nowmonth%>">
<input type="hidden" name="empid"  value="<%=empid%>"  >
<table width="500" border="0" cellspacing="0" cellpadding="0">
	<tr><TD  >
		<img border="0" src="../image/icon.gif" align="absmiddle">
		海外幹部個人資料
		</TD>
		<td width=80 align=right class="txt"><a href="vbscript:window.close()">(X)Close</a></td>
	</tr>
</table>
<hr size=0	style='border: 1px dotted #999999;' align=left width=550>



<TABLE WIDTH=300 CLASS="txt12" BORDER=0  cellspacing="1" cellpadding="2">	 
	<TR height=25 >
		<TD   align=right width="50">工號<br><font class=txt8>so the</font></TD>
		<td><%=EMPID%><br><%=empnam_vn%>&nbsp; <%=empnam_cn%></td>
	<TR height=25 >
		<TD   align=right >廠別<br><font class=txt8>Xuong</font></TD>
		<TD >			
				<select name=WHSNO  class="txt" onkeydown="enterto()"  >
					<%SQL="SELECT * FROM BASICCODE WHERE FUNC='WHSNO' ORDER BY SYS_TYPE "
					SET RST = CONN.EXECUTE(SQL)
					WHILE NOT RST.EOF
					%>
					<option value="<%=RST("SYS_TYPE")%>" <%IF RST("SYS_TYPE")=WHSNO THEN %> SELECTED <%END IF%>><%=RST("SYS_VALUE")%></option>
					<%
					RST.MOVENEXT
					WEND
					rst.close : set rst=nothing  
					conn.close : set conn=nothing 
					%>
				</SELECT> 
				<input type="button"  value="傳送(Send)" class="button"  onclick="go()"/>
		</TD>
</table>
<hr size=0	style='border: 1px dotted #999999;' align=left width=550>
 


</form>


</body>
</html>

<script language=vbscript>
function go()
	<%=self%>.action ="yebe0301.trans.asp?flag=S"
	<%=self%>.submit()
end function  
</script>


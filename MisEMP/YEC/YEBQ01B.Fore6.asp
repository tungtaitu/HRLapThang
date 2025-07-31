<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<!-- #include file = "../GetSQLServerConnection.fun" -->
<!-- #include file="../ADOINC.inc" -->
<%
response.buffer=true 
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
  
empid = TRIM(REQUEST("empid"))  
  
SQL="SELECT * FROM  view_empfile where ISNULL(STATUS,'')<>'D' AND  empid='"& empid &"' "

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
	COUNTRYDESC=TRIM(RS("cstr"))	'國籍
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
<body  topmargin="5" leftmargin="5"  marginwidth="0" marginheight="0"    >
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
<table width="650" border="0" cellspacing="0" cellpadding="0"   >
	<tr><td nowrap>
		<div id="navcontainer"  >
			<ul id="navlist">
			<li><a href="vbscript:chgpage(1)">基本資料<BR>Tu lieu co ban<BR>&nbsp;</a></li>
			<li ><a href="vbscript:chgpage(2)">教育訓練/証執照<br>huan luyen/<BR>bang cap</a></li>			
			<li ><a href="vbscript:chgpage(4)">獎懲紀錄<BR>Tu lieu<BR>thuong phat</a></li>
			<li ><a href="vbscript:chgpage(5)">部門/晉升紀錄<BR>Nang chuc/<BR>don vi </a></li>			
			<li id=active><a href="vbscript:chgpage(6)">請假紀錄<BR>Nhân viên<br>nghỉ phép</a></li> 
			</ul>
		</div> 
		</td>
	</tr>  
</table>    
<hr size=0	style='border: 1px dotted #999999;' align=left width=550> 

<TABLE WIDTH=550 CLASS=TXT BORDER=0 cellspacing="1" cellpadding="2" > 		 
	<TR  >
		<TD WIDTH=70 ALIGN=RIGHT bgcolor="#EBEBEB" >工號<br>Số thẻ</TD>
		<TD   ><%=EMPID%> &nbsp;
		<%=EMPNAM_CN%>&nbsp;<%=EMPNAM_VN%> &nbsp;
		(<%=COUNTRYDESC%>)&nbsp;&nbsp;
		<input name=empid value=<%=empid%> type=hidden >
		<input name=empautoid value=<%=empautoid%> type=hidden >
		<input name=country value=<%=country%> type=hidden >				

		<td width="150"  align=left valign=top  nowrap  rowspan="4" >   <!--照片-->
			<img src="../yeb/pic/<%=EMPID%>.jpg"  border=1 width=98 height=126  >
		</td> 
		
	</tr>
	<tr>	
		<TD  ALIGN=RIGHT bgcolor="#EBEBEB">到職日<br>NVX</TD>
		<TD  ><%=INDAT%>&nbsp;&nbsp;	</td>

	</TR>	
	<tr>
		<td align="right" bgcolor="#EBEBEB">單位<br>Đơn vị
		<td  > (<%=WSTR%>) <%=groupid%> <%=GSTR%>-<%=zuno%> <%=ZSTR%></TD>		
	</tr>	
	<tr>
		<TD ALIGN=RIGHT bgcolor="#EBEBEB">職務<br>Chuc Vu</TD>
		<TD  ><%=job%>&nbsp;<%=jstr%></TD>  
	</tr>
</TABLE> 
<hr size=0	style='border: 1px dotted #999999;' align=left width=600> 
<table width=550><tr><td align=center>
<TABLE WIDTH=665 CLASS=txt8 BORDER=0  cellspacing="1" cellpadding="2" bgcolor="#cccccc">	
	<TR bgcolor=#e4e4e4 height=22> 
		<Td  width=50 nowrap align=center >統計年月<br>YYMM</td>		
		<Td  width=40 nowrap align=center valign="top" >KZ<BR>曠職<br>(H)</td>
		<%sqlx="select *from basicCode where func='JB' order by sys_type" 
		Set rsx = Server.CreateObject("ADODB.Recordset")
		rsx.open sqlx , conn, 3 ,3 
		cnt = rsx.recordcount 
		while not rsx.eof 
		%>
			<Td  width=40 nowrap align=center  valign="top" ><%=rsx("sys_type")%><br><%=rsx("sys_value")%><br>(H)</td>
		<%rsx.movenext
		wend 
		set rsx=nothing 
		%>  		
		<Td  width=40 nowrap align=center valign="top" ><br>境外<BR>事假<br>(H)</td>
		<Td  width=40 nowrap align=center valign="top" ><br>境外<BR>病假<br>(H)</td>
	</tr> 
	<%
  sql="exec sp_showmpJB '"& empid &"' " 
	'response.write sql 		
	set rs=conn.execute(sql)
	response.flush	 
	while not rs.eof 
		if country="VN" then 
			kzhour = rs("kzhour") 
		else
			kzhour =  0 
		end if 
	%>
	<tr bgcolor="#ffffff" height="22">
		<td align="center"><%=rs("ym")%></td> 
		<td align="center"><font color="<%if kzhour="0" then%>#ffffff<%else%>Black<%end if%>"><%=rs("kzhour")%></font></td>
		<%for z = 1 to rs("cols") 	%>
		<td align="center"><font color="<%if rs(8+z)="0" then%>#ffffff<%else%>Black<%end if%>"><%=rs(8+z)%> </td>
		<%next%>
		<td align="center"><font color="<%if rs("jaW")="0" then%>#ffffff<%else%>Black<%end if%>"><%=rs("jaW")%></td>
		<td align="center"><font color="<%if rs("jbW")="0" then%>#ffffff<%else%>Black<%end if%>"><%=rs("jbW")%></td>
		 
	</tr>
	<%rs.movenext
	wend 
	set rs=nothing 
	%>	
</table>
</td></tr></table>
<hr size=0	style='border: 1px dotted #999999;' align=left width=600> 
 
<TABLE WIDTH=550>
		<tr ALIGN=center>
			<TD >
			<input type=button name=send value="關閉視窗(Close)"  class=button onclick="window.close">
			</TD> 
		</TR>
</TABLE>


</form>


</body>
</html>
		
<script language=vbscript>
 

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
	end if 
	
end function 
 
</script>
<%response.end %>


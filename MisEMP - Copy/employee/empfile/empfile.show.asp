<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<!-- #include file = "../../GetSQLServerConnection.fun" --> 
<!-- #include file="../../ADOINC.inc" -->
<%

SELF = "empfilemain"
 
Set conn = GetSQLServerConnection()	  
Set rs = Server.CreateObject("ADODB.Recordset")  

nowmonth = year(date())&right("00"&month(date()),2)  
if month(date())="01" then  
	calcmonth = year(date()-1)&"12" 
else
	calcmonth =  year(date())&right("00"&month(date())-1,2)   
end if 	
EMPID = TRIM(REQUEST("EMPID"))
'IF EMPID="" THEN EMPID="L0051" 
empautoid = TRIM(REQUEST("empautoid")) 

SQL="SELECT * FROM  EMPFILE WHERE ISNULL(STATUS,'')<>'D' AND ( autoid='"& empautoid &"' or empid='"& empid &"' )  " 
RS.OPEN SQL , CONN, 3, 3 
IF NOT RS.EOF THEN 
	EMPID=TRIM(RS("EMPID"))	'員工編號
	INDAT=TRIM(RS("INDAT"))	'到職日  
	WHSNO=TRIM(RS("WHSNO"))	'廠別
	UNITNO=TRIM(RS("UNITNO"))	'處/所
	GROUPID=TRIM(RS("GROUPID"))	'組/部門
	ZUNO=TRIM(RS("ZUNO"))	'單位
	EMPNAM_CN=TRIM(RS("EMPNAM_CN"))	'姓名(中)
	EMPNAM_VN=TRIM(RS("EMPNAM_VN"))	'姓名(越)
	COUNTRY=TRIM(RS("COUNTRY"))	'國籍
	BYY=RIGHT(RS("BYY"),2)	'年(生日)
	BMM=MID(RS("BMM"),5,2)	'月(生日)
	BDD=RIGHT(RS("BDD"),2)	'日(生日)
	AGES=TRIM(RS("AGES"))	'年齡		
	SEX=TRIM(RS("SEX"))	'性別
	JOB=TRIM(RS("JOB"))
	PERSONID=TRIM(RS("PERSONID"))	'身分証字號
	BHDAT=TRIM(RS("BHDAT"))	'簽約日(保險日)
	PASSPORTNO=TRIM(RS("PASSPORTNO"))	'護照號碼
	VISANO=TRIM(RS("VISANO"))	'簽證號碼
	PDUEDATE=TRIM(RS("PDUEDATE"))	'護照有效期
	VDUEDATE=TRIM(RS("VDUEDATE"))	'簽證有效期
	PHONE=TRIM(RS("PHONE"))	'聯絡電話
	MOBILEPHONE=TRIM(RS("MOBILEPHONE"))	'手機
	HOMEADDR=TRIM(RS("HOMEADDR"))	'聯絡地址
	EMAIL=TRIM(RS("EMAIL"))	'EMAIL
	OUTDAT=TRIM(RS("OUTDAT"))	'離職日
	MEMO=TRIM(RS("MEMO"))	'其他說明 
	'-----------------------------------------
	PHU = RS("PHU")
	NN = RS("NN")
	KT = RS("KT")
	TTKH = RS("TTKH")
	MT = RS("MT")
	BB = RS("BB")
	CV = RS("CV")
END IF 

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
<meta http-equiv="refresh">
<link rel="stylesheet" href="../../Include/style.css" type="text/css">
<link rel="stylesheet" href="../../Include/style2.css" type="text/css"> 
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--
'-----------------enter to next field
function enterto()
	if window.event.keyCode = 13 then window.event.keyCode =9 		
end function  

function f()
	<%=self%>.EMPID.focus()	
end function  

function groupchg()
	code = <%=self%>.GROUPID.value
	open "empfile.back.asp?ftype=groupchg&code="&code , "Back" 
	'parent.best.cols="50%,50%"	
end function 
-->
</SCRIPT>  
</head>   
<body  topmargin="5" leftmargin="5"  marginwidth="0" marginheight="0"  onkeydown="enterto()" onload="f()">
<form name="<%=self%>" method="post" action="<%=self%>.asp">
<INPUT TYPE=HIDDEN NAME="UID" VALUE=<%=SESSION("NETUSER")%>>	
<table width="460" border="0" cellspacing="0" cellpadding="0">
	<tr><TD width=100% align=center>人事薪資系統( 員工基本資料查詢 ) </TD></tr>
</table>
<hr size=0	style='border: 1px dotted #999999;' align=left >		
<TABLE WIDTH=510 CLASS=FONT9 BORDER=0> 
	<TR height=25 > 
		<TD width=80 nowrap align=right height=25>員工編號：</TD>
		<TD ><INPUT NAME=EMPID SIZE=12 CLASS="readonly" readonly  VALUE="<%=EMPID%>"> </TD>		 
		<TD width=60 nowrap align=right height=25>到職日：</TD>
		<TD ><INPUT NAME=indat SIZE=12 CLASS="readonly" readonly  VALUE="<%=FDT(indat)%>"></TD>
		<td width="130" rowspan="5" align=center valign=center><img src="../photos/nophotos.gif" width="130" height="130" border=1></td>
	</TR>
	<TR height=25 >
		<TD nowrap align=right>廠別：</TD>
		<TD > 
			<select name=WHSNO  class=font9 disabled  >
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
		</TD> 
		<TD width=60 nowrap align=right >處/所：</TD>
		<TD > 
			<select name=unitno  class=font9 disabled >
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
		</TD>
	</tr>
	<TR height=25 >
		<TD nowrap align=right >組/部門：</TD>
		<TD >
			<select name=GROUPID  class=font9 onchange=groupchg() style="width:60" disabled  >
				<%SQL="SELECT * FROM BASICCODE WHERE FUNC='GROUPID' and sys_type <>'AAA' ORDER BY SYS_TYPE "
				SET RST = CONN.EXECUTE(SQL)
'				RESPONSE.WRITE SQL 
				WHILE NOT RST.EOF  
				%>
				<option value="<%=RST("SYS_TYPE")%>" <%IF RST("SYS_TYPE")=TRIM(GROUPID) THEN %> SELECTED <%END IF%>><%=RST("SYS_VALUE")%></option>				 
				<%
				RST.MOVENEXT
				WEND 
				%>
			</SELECT>
			<%SET RST=NOTHING %>
			<select name=zuno  class=font9 style='width:50' disabled >				
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
		</TD>
		<TD nowrap align=right >職等：</TD>
		<TD >
			<select name=JOB  class=font9 style='width:75' disabled >			 
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
		</TD>
	</TR>	
	<TR height=25 >
		<TD nowrap align=right>員工姓名(中)：</TD>
		<TD >
			<INPUT NAME=nam_cn SIZE=15 CLASS="readonly" VALUE="<%=EMPNAM_CN%>" readonly >			
		</TD>
		<TD nowrap align=right>國籍：</TD>
		<TD >
			<select name=country  class="inputbox" style='width:75'  disabled  >
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
	</TR>
	<TR height=25 >
		<TD nowrap align=right >員工姓名(越)：</TD>
		<TD colspan=3><INPUT NAME=nam_vn SIZE=38 CLASS="readonly" readonly  VALUE="<%=EMPNAM_VN%>" ></TD>
	</TR>
	<TR height=25 >
		<TD nowrap align=right >出生日期：</TD>
		<TD colspan=4>
		<INPUT NAME=BYY SIZE=5 CLASS=readonly readonly VALUE="<%=BYY%>" > 年
		<INPUT NAME=BMM SIZE=3 CLASS=readonly readonly VALUE="<%=BMM%>"> 月
		<INPUT NAME=BDD SIZE=3 CLASS=readonly readonly VALUE="<%=BDD%>"> 日&nbsp;&nbsp;  				
		年齡： <input name=ages size=5 class=readonly readonly  VALUE=<%=AGES%>>&nbsp; &nbsp; 
		<INPUT type="radio" id=radio1 name=radio1 <%IF SEX="M" THEN %>CHECKED<%END IF%> onclick=sexchg(0) disabled > 男 &nbsp; 
		<INPUT type="radio" id=radio1 name=radio1 <%IF SEX="F" THEN %>CHECKED<%END IF%> onclick=sexchg(1) disabled > 女 
		<input type=hidden name=sexstr value="<%=sex%>" size=1>
		</TD>
	</TR>
</TABLE>
<!--------------------------------------------------------------------> 
<TABLE WIDTH=500 CLASS=FONT9 BORDER=0> 
	<tr>
		<td width=90 nowrap align=right height=25  >身分証字號：</td>
		<td ><input name=personID size=20 class=readonly readonly  VALUE="<%=PERSONID%>"></td>		
		<td width=80 nowrap align=right >買保險日：</td>
		<td ><input name=BHDAT size=15 class=readonly readonly  VALUE="<%=FDT(BHDAT)%>">(簽約日)</td>		 
	</tr>	
	<tr>
		<td nowrap align=right height=25 >護照號碼：</td>
		<td><input name=PASSPORTID size=20 class=readonly readonly  VALUE="<%=PASSPORTNO%>"></td>
		<td nowrap align=right >(護)有效期：</td>
		<td ><input name=pduedate size=15 class=readonly readonly  VALUE="<%=FDT(PDUEDATE)%>"></td>
	</tr>	
	<tr>
		<td nowrap align=right height=25 >簽證號碼：</td>
		<td><input name=visano size=20 class=readonly readonly  VALUE="<%=VISANO%>"></td>
		<td nowrap align=right >(簽)有效期：</td>
		<td ><input name=vduedate size=15 class=readonly readonly  VALUE="<%=FDT(VDUEDATE)%>"></td>
	</tr>	
	<tr>
		<td nowrap align=right height=25 >聯絡電話：</td>
		<td><input name=phone size=20 class=readonly readonly  VALUE="<%=PHONE%>"></td>
		<td nowrap align=right >手機：</td>
		<td ><input name=mobilephone size=15 class=readonly readonly  VALUE="<%=MOBILEPHONE%>"></td>
	</tr>
	<tr>
		<td nowrap align=right height=25 >聯絡地址：</td>
		<td colspan=3><input name=homeaddr size=54 class=readonly readonly  VALUE="<%=HOMEADDR%>"></td>
	</tr>
	<tr>
		<td nowrap align=right height=25 >E-MAIL：</td>
		<td><input name=email size=20 class=readonly readonly  VALUE="<%=EMAIL%>"></td>
		<td nowrap align=right >離職日：</td>
		<td ><input name=outdat size=15 class=readonly readonly VALUE="<%=FDT(OUTDAT)%>"></td>
	</tr>
	<tr>
		<td nowrap align=right height=25 >其他說明：</td>
		<td colspan=3><input name=memo size=54 class=readonly readonly  VALUE="<%=MEMO%>" ></td>
	</tr>
</table>
<hr size=0	style='border: 1px dotted #999999;' align=left width=500>	
<%if session("rights")<>3 then %>
<table width=500><tr><td><fieldset >
<legend class=font9>薪資給付類(幣別：VND) <font color=red>基本薪資：<%=BB%></font></legend>
	<table width=500 class=font9>
		<tr>
			<td nowrap align=right >職專加給：</td>
			<td><input name=CV class=readonly readonly  size=10 VALUE="<%=CV%>"></td>
			<td nowrap align=right>特加(Y)：</td>
			<td><input name=PHU class=readonly readonly size=10 VALUE="<%=PHU%>"></td>
			<td nowrap align=right>語言加給：</td>
			<td><input name=NN class=readonly readonly size=10 VALUE="<%=NN%>"></td>
		</td>
		<tr>
			<td nowrap align=right >技術加給：</td>
			<td><input name=KT class=readonly readonly size=10 VALUE="<%=KT%>"></td>
			<td nowrap align=right>環境加給：</td>
			<td><input name=MT class=readonly readonly size=10 VALUE="<%=MT%>" ></td>		
			<td nowrap align=right>其他加給：</td>
			<td><input name=TTKH class=readonly readonly size=10 VALUE="<%=TTKH%>"></td>
		</td>
	</table>
</fieldset></td></tr></table>
<%end if%>
<br>
<center>
<input type=button name=btn value="關閉此視窗" class=button onclick="vbscript:window.close()"> 
</center>
  
</form>


</body>
</html>
<script language=vbscript>
setTimeout ("CloseWin()" ,0);
function CloseWin()
{
window.opener.location.reload();
window.close();
}
</script> 

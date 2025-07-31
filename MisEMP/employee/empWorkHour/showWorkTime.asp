<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<!-- #include file = "../../GetSQLServerConnection.fun" -->
<!-- #include file="../../ADOINC.inc" -->
<%
SESSION.CODEPAGE="65001"
SELF = "empworkb"

Set conn = GetSQLServerConnection()
Set rs = Server.CreateObject("ADODB.Recordset")
Set rst = Server.CreateObject("ADODB.Recordset")

gTotalPage = 1
PageRec = 10    'number of records per page
TableRec = 30    'number of fields per record


YYMM = REQUEST("YYMM")
'response.write yymm
IF YYMM="" THEN
	YYMM = year(date())&right("00"&month(date()),2)
	'YYMM="200601"
	cDatestr=date()
	days = DAY(cDatestr+(32-DAY(cDatestr))-DAY(cDatestr+(32-DAY(cDatestr))))   '一個月有幾天
ELSE
	cDatestr=CDate(LEFT(YYMM,4)&"/"&RIGHT(YYMM,2)&"/01")
	days = DAY(cDatestr+(32-DAY(cDatestr))-DAY(cDatestr+(32-DAY(cDatestr))))   '一個月有幾天
END IF


if month(date())="01" then
	calcmonth = year(date()-1)&"12"
else
	calcmonth =  year(date())&right("00"&month(date())-1,2)
end if
EMPID = TRIM(REQUEST("EMPID"))
workdat = replace(TRIM(REQUEST("workdat")),"/","")

 

'--------------------------------------------------------------------------------------
SQL="SELECT convert(char(10), indat, 111) as Nindate, b.sys_value as groupstr, c.sys_value as zunostr, d.sys_value as jobstr , a.* from  "&_
	"( SELECT * FROM  view_EMPFILE WHERE ISNULL(STATUS,'')<>'D' AND empid='"& empid &"' ) a "&_
	"left join ( select * from basicCode where func='groupid' ) b on b.sys_type = a.groupid "&_
	"left join ( select * from basicCode where func='zuno' ) c on c.sys_type = a.zuno "&_
	"left join ( select * from basicCode where func='lev' ) d on d.sys_type = a.job "
	'RESPONSE.WRITE SQL
	'RESPONSE.END
	RST.OPEN SQL , CONN, 3, 3
IF NOT RST.EOF THEN
	empautoid = TRIM(RST("AUTOID"))
	EMPID=TRIM(RST("EMPID"))	'員工編號
	INDAT=TRIM(RST("Nindat"))	'到職日
	TX=TRIM(RST("TX"))	'特休
	WHSNO=TRIM(RST("WHSNO"))	'廠別
	UNITNO=TRIM(RST("UNITNO"))	'處/所
	GROUPID=TRIM(RST("GROUPID"))	'組/部門
	ZUNO=TRIM(RST("ZUNO"))	'單位
	JOB=TRIM(RST("JOB"))	'職等
	EMPNAM_CN=TRIM(RST("EMPNAM_CN"))	'姓名(中)
	EMPNAM_VN=TRIM(RST("EMPNAM_VN"))	'姓名(越)
	COUNTRY=TRIM(RST("COUNTRY"))	'國籍
	GROUPSTR = TRIM(RST("GROUPSTR"))  '組/部門
	ZUNOSTR = TRIM(RST("ZUNOSTR"))  '單位
	JOBSTR = TRIM(RST("JOBSTR"))  '職等
	outdat = TRIM(RST("outdat"))  '離職日
END IF
SET RST=NOTHING

'請假紀錄 ----------------------------------------------------------

'忘刷卡紀錄 ----------------------------------------------------------

%>

<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta HTTP-EQUIV="refresh" >
<link rel="stylesheet" href="../../Include/style.css" type="text/css">
<link rel="stylesheet" href="../../Include/style2.css" type="text/css">
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--
'-----------------enter to next field
function enterto()
	if window.event.keyCode = 13 then window.event.keyCode =9
end function

function f()
	<%=self%>.TIMEUP(0).SELECT()
end function

function colschg(index)
	thiscols = document.activeElement.name
	if window.event.keyCode = 38 then
		IF INDEX<>0 THEN
			document.all(thiscols)(index-1).SELECT()
		END IF
	end if
	if window.event.keyCode = 40 then
		document.all(thiscols)(index+1).SELECT()
	end if

end function

-->
</SCRIPT>
</head>
<body  topmargin="5" leftmargin="5"  marginwidth="0" marginheight="0"   >
<form name="<%=self%>"  method="post" action = "<%=self%>.upd.asp" >
<INPUT TYPE=HIDDEN NAME="PCNTFG" VALUE=<%=PCNTFG%>>
<INPUT TYPE=HIDDEN NAME="UID" VALUE=<%=SESSION("NETUSER")%>>
<INPUT TYPE=HIDDEN NAME="empautoid" VALUE=<%=empautoid%>>
<INPUT TYPE=hidden NAME=TotalPage VALUE="<%=TotalPage%>">
<INPUT TYPE=hidden NAME=CurrentPage VALUE="<%=CurrentPage%>">
<INPUT TYPE=hidden NAME=RecordInDB VALUE="<%=RecordInDB%>">
<INPUT TYPE=hidden NAME=FTotalPage VALUE="<%=FTotalPage%>">
<INPUT TYPE=hidden NAME=FCurrentPage VALUE="<%=FCurrentPage%>">
<INPUT TYPE=hidden NAME=FRecordInDB VALUE="<%=FRecordInDB%>">
<INPUT TYPE=hidden NAME=PageRec VALUE="<%=PageRec%>">
<!-- table width="460" border="0" cellspacing="0" cellpadding="0">
	<tr>
		<TD align=center >員工差勤作業 </TD>
	</tr>
</table>
<hr size=0	style='border: 1px dotted #999999;' align=left width=500 -->
<table width=450 class=font9 border=1  cellspacing="0" cellpadding="1" >	
	<tr height=25 >
		<td width=60>員工編號:</td>
		<td colspan=3><%=EMPID%>-<%=empnam_cn&" "&empnam_vn%> 
			<!--input name=empid value="<%=EMPID%>" size=7 class="readonly" readonly style="height:22">
			<input name=empnam value="<%=empnam_cn&" "&empnam_vn%>" size=20 class="readonly8" readonly style="height:22"-->
		</td>			 
	</tr>
	<tr  height=25>
		<td width=60>到職日期:</td>
		<td width=160><%=indat%></td>
		<td width=60>單位部門:</td>
		<td><%=GROUPSTR%>-<%=zunoSTR%> </td>
	</tr> 
	<tr  height=25>
		<td width=60>離職日期:</td>
		<td ><%=outdat%>　</td>	
		<td>尚有特休:</td>
		<td><%=tx%>(天) = <%=tx*8%>(H)
		</td>
	</tr>	
</table>
<hr size=0	style='border: 1px dotted #999999;' align=left width=450>
<fieldset style="margin:0;padding:3;width=450"><legend><font class=txt9 color=blue>請假資料</font></legend>
<table width=450 class=font9>
	<tr bgcolor="#ffe4e1">
		<td align=center>假別</td>
		<td align=center>日期(起)</td>
		<td align=center>時間(起)</td>
		<td align=center>日期(迄)</td>
		<td align=center>時間(迄)</td>
		<td align=center>時數</td>		
		<td align=center>事由</td>		
	</tr>
	<%
	sql="select convert(char(10), dateup, 111) as d1 , convert(char(10), datedown, 111) as d2, * from empholiday where empid='"& empid &"' and  convert(char(8),dateup,112)='"& workdat &"'  "
	'response.write sql 
	Set RS1 = Server.CreateObject("ADODB.Recordset")
	rs1.open sql , conn, 3, 3 
	while not rs1.eof 
		if rs1("jiatype")="A" then 
 			jiastr = "(A) 事假"
 		elseif 	rs1("jiatype")="B" then 
 			jiastr = "(B) 病假"
 		elseif 	rs1("jiatype")="C" then 
 			jiastr = "(C) 婚假"
 		elseif 	rs1("jiatype")="D" then 
 			jiastr = "(D) 喪假"
 		elseif 	rs1("jiatype")="E" then 
 			jiastr = "(E) 年假"
 		elseif 	rs1("jiatype")="F" then 
 			jiastr = "(F) 產假"
 		elseif 	rs1("jiatype")="G" then 
 			jiastr = "(G) 公假"					
 		elseif 	rs1("jiatype")="H" then 
 			jiastr = "(H) 工傷"		
 		else 
 			jiastr=""
 		end if
	%>
	<tr bgcolor="#ffffe0">
		<td align=center><%=jiastr%></td>
		<td align=center><%=rs1("D1")%></td>
		<td align=center><%=rs1("timeup")%></td>
		<td align=center><%=rs1("D2")%></td>
		<td align=center><%=rs1("timedown")%></td>
		<td align=center><%=rs1("hhour")%></td>		
		<td align=center><%=rs1("memo")%></td>		
	</tr> 
	<%
	rs1.movenext
	wend 
	set rs1=nothing 
	%>
</table>	
</fieldset>
<P><P>
<fieldset style="margin:0;padding:3;width=450"><legend><font class=txt9 color=blue>忘刷卡資料</font></legend>
<table width=450 class=font9>
	<tr bgcolor="#e6e6fa">		
		<td align=center>臨時卡</td>
		<td align=center>日期(起)</td>
		<td align=center>時間(起)</td>
		<td align=center>日期(迄)</td>
		<td align=center>時間(迄)</td>
		<td align=center>時數</td>		
	</tr>
	<%
	sql2="select convert(char(10),dat,111) as D1, * from empforget  "&_
	     "where empid='"& empid &"' and convert(char(8), dat, 112)='"& workdat &"' and isnull(status,'')<>'D' "
	Set RS2 = Server.CreateObject("ADODB.Recordset")
	rs2.open sql2 , conn, 3, 3  
	while not rs2.eof 
	%>
	<tr bgcolor="#ffffe0">		
		<td align=center><%=rs2("lsempid")%></td>
		<td align=center><%=rs2("D1")%></td>
		<td align=center><%=rs2("timeup")%></td>
		<td align=center><%=rs2("D1")%></td>
		<td align=center><%=rs2("timedown")%></td>
		<td align=center><%=rs2("toth")%></td>		
	</tr> 
	<%
	rs2.movenext
	wend 
	set rs2=nothing 
	%>
</table>	
</fieldset> 
<P><P>
<fieldset style="margin:0;padding:3;width=450"><legend><font class=txt9 color=blue></font></legend>
<table width=450 class=font9>
	<tr bgcolor="#e6e6fa">		
		<td align=center>Ngay cong</td>
		<td align=center>Vao 1</td>
		<td align=center>Ra 1</td>
		<td align=center>Vao 2</td>
		<td align=center>Ra 2</td>
		<td align=center>Vao 3</td>
		<td align=center>Ra 3</td>		
	</tr>
	<%
	sql3="select [empid] ,[workdat],SUBSTRING(workdat, 1, 4)+'/'+SUBSTRING(workdat, 5, 2)+'/'+SUBSTRING(workdat, 7, 2) as workdat1, [timeup1]  ,[timedown1] ,[timeup2] ,[timedown2] ,[timeup3] ,[timedown3] ,[yymm] from EMPWORKD  where empid='"& empid &"' and workdat='"& workdat &"' "
	'response.write sql3
	'Set rs3 = Server.CreateObject("ADODB.Recordset")
	'rs3.open sql3 , conn, 3, 3 
	set rs3=conn.execute(sql3)
	while not rs3.eof 
	%>
	<tr bgcolor="#ffffe0">
		<td align=center><%=rs3("workdat1")%></td>
		<td align=center><%=rs3("timeup1")%></td>		
		<td align=center><%=rs3("timedown1")%></td>
		<td align=center><%=rs3("timeup2")%></td>		
		<td align=center><%=rs3("timedown2")%></td>
		<td align=center><%=rs3("timeup3")%></td>		
		<td align=center><%=rs3("timedown3")%></td>
	</tr> 
	<%
	rs3.movenext
	wend 
	set rs3=nothing 
	%>
</table>	
</fieldset> 
<p>
<TABLE border=0 width=450 class=font9 >
<tr>
    <td align="CENTER" height=40  >
	<input type=BUTTON name=send value="關閉此視窗(CLOSE)"  class=button ONCLICK="vbscript:window.close()">　　
	</td>
</tr>
<%sqlx="select convert(char(20),mdtm,120) as tmdtm,  muser as Tuser, * from empwork where empid='"& empid &"' and workdat='"& workdat &"' "  
  set rsx=conn.execute(sqlx)
  if not rsx.eof then 
	dt = rsx("tmdtm")
	keyinby = rsx("Tuser")
  end if 	
  set rsx=nothing 
%>
<tr>
	<td align=right class=txt8><font color="#999999">edit dateTime:</font><%=dt%>  <%=keyinby%></td>
</tr>	
</TABLE>
</form>
</body>
</html>




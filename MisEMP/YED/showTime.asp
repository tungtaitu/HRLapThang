<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<!-- #include file = "../GetSQLServerConnection.fun" -->
<!-- #include file="../ADOINC.inc" -->
<%
SESSION.CODEPAGE="65001"
SELF = "ShowTime"

Set conn = GetSQLServerConnection()
Set rs = Server.CreateObject("ADODB.Recordset")
Set rst = Server.CreateObject("ADODB.Recordset")

gTotalPage = 1
PageRec = 10    'number of records per page
TableRec = 30    'number of fields per record


YYMM = REQUEST("YYMM") 
empid  =REQUEST("empid")
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

gTotalPage = 1
PageRec = 10    'number of records per page
TableRec = 30    'number of fields per record

'請假紀錄 ----------------------------------------------------------
sql="select convert(varchar(10),dat,111) as Ndat, * from  ( "&_
	"select a.dat, a.status, b.empid, b.workdat, b.worktim , 'JBC'  sys  from "&_
	"( select * from   ydbmcale where  convert(char(6), dat, 112)  = '"& yymm &"'  ) a  "&_
	"left  join ( select * from empworktime  where  empid='"& empid &"' )   b on  workdat =  convert(char(8), a.dat, 112)  "&_
	"union "&_
	"select  a.dat, a.status,  right( left( ltrim(rtrim(emp_id)),6) , 5 ) ,  convert(varChar(8), c.sign_time, 112) ,  replace( right(convert(varchar(20), c.sign_time, 120),8) ,':','') , 'One' as sys "&_
	"  from   "&_
	"( select * from   ydbmcale where  convert(char(6), dat, 112)  = '"& yymm &"'  ) a   "&_
	"left join ( select * from timerecords1  where  right( left( ltrim(rtrim(emp_id)),6) , 5 ) ='"& empid &"'  ) c on   convert(char(8), sign_time,112) =   convert(char(8), a.dat, 112) "&_
	") x  where  isnull(x.empid,'')<>'' "&_
	"order by workdat  , worktim , sys " 
'response.write( sql)
if request("TotalPage") = "" or request("TotalPage") = "0" then
	CurrentPage = 1
	'RESPONSE.WRITE SQLSTRA
	'RESPONSE.END
	rs.Open sql, conn, 1, 3
	IF NOT RS.EOF THEN
		pagerec= rs.RecordCount
		rs.PageSize =  pagerec
		RecordInDB = rs.RecordCount
		TotalPage = rs.PageCount
		gTotalPage = TotalPage
	END IF
	'response.write days 
	Redim tmpRec(gTotalPage, PageRec, TableRec)   'Array	
	for i = 1 to TotalPage
		for j = 1 to PageRec
			if not rs.EOF then
				tmpRec(i, j, 0) = "no"
				tmpRec(i, j, 1) = trim(rs("Ndat")) 
				tmpRec(i, j, 2)= mid("日一二三四五六",weekday(cdate(rs("Ndat"))) , 1 )  
				tmpRec(i, j, 3) = left(RS("worktim"),2)&":"&mid(RS("worktim"),3,2)
				tmpRec(i, j, 4) = rs("sys") 
				rs.MoveNext
			else
				exit for
			end if
	 	next 
	 	if rs.EOF then
			rs.Close
			Set rs = nothing
			exit for
	 	end if
	next
	Session("ShowTime") = tmpRec 	 
end if		

%>

<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" >
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
</SCRIPT>

</head>
<body  topmargin="5" leftmargin="5"  marginwidth="0" marginheight="0"  >
<form name="<%=self%>"  method="post" action = "<%=self%>.upd.asp" >
<INPUT TYPE=HIDDEN NAME="PCNTFG" VALUE=<%=PCNTFG%>>
<INPUT TYPE=HIDDEN NAME="UID" VALUE=<%=SESSION("NETUSER")%>>
<INPUT TYPE=HIDDEN NAME="empautoid" VALUE=<%=empautoid%>>
<INPUT TYPE=hidden NAME=TotalPage VALUE="<%=TotalPage%>">
<INPUT TYPE=hidden NAME=CurrentPage VALUE="<%=CurrentPage%>">
<INPUT TYPE=hidden NAME=RecordInDB VALUE="<%=RecordInDB%>"> 
<INPUT TYPE=hidden NAME=PageRec VALUE="<%=PageRec%>">
 
<table width=300 class=txt border=0  cellspacing="0" cellpadding="1" >	
	<tr height=25 >
		<td  >員工:</td>
		<td  ><%=EMPID%>-<%=empnam_cn&" "&empnam_vn%> 
			<!--input name=empid value="<%=EMPID%>" size=7 class="readonly" readonly style="height:22">
			<input name=empnam value="<%=empnam_cn&" "&empnam_vn%>" size=20 class="readonly8" readonly style="height:22"-->
		</td>	
		<td align="CENTER"    >
			<input border="0" src="../Picture/pic_close.gif" name="closeN" width="24" height="23" type="image"  onclick="vbscript:window.close()">
		</td>		 
	</tr> 
</table>
<hr size=0	style='border: 1px dotted #999999;' align=left width=300>
<table width=500><tr><td  > 
	<TABLE   width=220 border=0  CLASS=txt cellspacing="1" cellpadding="1">
		<Tr bgcolor=#e4e4e4 height=25>
			<td colspan=3 align=center><%=yymm%> 刷卡紀錄</td>
		</tr>
		<TR BGCOLOR=#e4e4e4 height=25>
			<TD  ALIGN=CENTER width=120    >日期</TD> 
			<TD ALIGN=CENTER width=60  >時間</TD>		 
			<TD  ALIGN=CENTER width=40  >系統</TD> 
		</tr>
		<%for x = 1 to pagerec
			IF x MOD 2 = 0 THEN
				WKCOLOR="LavenderBlush"
			ELSE
				WKCOLOR="LightYellow"		
			END IF
			
		%>
			<tr bgcolor="<%=WKCOLOR%>" height=22>
				<td ALIGN=CENTER>
					<%if tmprec(CurrentPage,x-1,1)<>tmprec(CurrentPage,x,1) then  %>
						<%=tmprec(CurrentPage,x,1)%>(<%=tmprec(CurrentPage,x,2)%>)
					<%end if%>	
				</td>
				<td ALIGN=CENTER><%=tmprec(CurrentPage,x,3)%></td>
				<td ALIGN=CENTER><%=tmprec(CurrentPage,x,4)%></td>
			</tr>
		<%next%>
	</table> 
	<TABLE border=0 width=220 class=txt >
		<tr>
	    <td align="CENTER" height=40  >
		<input type=BUTTON name=send value="關閉此視窗(CLOSE)"  class=button ONCLICK="vbscript:window.close()">　　
		</td>
		</tr> 
	</TABLE>
</td></tr></table>	
</form>
</body>
</html>




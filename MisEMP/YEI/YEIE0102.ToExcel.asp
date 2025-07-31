<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%> 
<!-- #include file = "../GetSQLServerConnection.fun" --> 
<!-- #include file="../ADOINC.inc" -->
<!-- #include file="../Include/func.inc" -->
<!--#include file="../include/checkpower.asp"-->  
<%
self="YEIE0102B"   

Set conn = GetSQLServerConnection()
Set rs = Server.CreateObject("ADODB.Recordset")
Set rds = Server.CreateObject("ADODB.Recordset")

nowmonth = year(date())&right("00"&month(date()),2)
if month(date())="01" then
	calcmonth = year(date()-1)&"12"
else
	calcmonth =  year(date())&right("00"&month(date())-1,2)
end if 

F_whsno = request("F_whsno")
F_groupid = request("F_groupid")
F_zuno = request("F_zuno") 
F_shift=request("F_shift")
F_empid =request("empid")
F_country=request("F_country")
fclass = request("fclass")  
sortvalue = request("sortvalue") 
if sortvalue ="" then sortvalue="b.country , h.lw, h.lg, len(h.ls)desc, h.ls, a.empid"

FUNCTION FDT(D)
	Response.Write YEAR(D)&"/"&RIGHT("00"&MONTH(D),2)&"/"&RIGHT("00"&DAY(D),2)
END FUNCTION 
khym = request("khym")
if request("khym")="" then 
	khym=nowmonth
end if  

 
 '一個月有幾天 
cDatestr=CDate(LEFT(khym,4)&"/"&RIGHT(khym,2)&"/01") 
days = DAY(cDatestr+(32-DAY(cDatestr))-DAY(cDatestr+(32-DAY(cDatestr))))   '一個月有幾天  
'本月最後一天 
ENDdat = LEFT(khym,4)&"/"&RIGHT(khym,2)&"/"&DAYS   

sql="select  b.cstr, b.empnam_cn, b.empnam_vn, b.country, b.nindat, b.outdate, d.sys_value as gstr, "&_
	"e.sys_value as zstr, f.sys_value as wstr, g.sys_value as sstr, a.* , h.ls from "&_
	"(  "&_
	"select count(*) as weekcnt, khym, empid, khw, khg, khz, khs , sum(fna+fnb+fnc+fnd ) as monthfen from  empkhb where khym='"& khym &"' "&_
	"group by khym,empid, khw, khg, khz, khs  "&_
	") a  "&_
	"left join ( select *from view_empfile ) b on b.empid = a.empid   "&_ 
	"left join ( select* from  basicCode  where func='groupid' ) d on d.sys_type = a.khg "&_
	"left join ( select* from  basicCode  where func='zuno' ) e on e.sys_type = a.khz "&_
	"left join ( select* from  basicCode  where func='whsno' ) f on f.sys_type = a.khw "&_	
	"left join ( select* from  basicCode  where func='shift' ) g on g.sys_type = a.khs "&_	 
	"left join ( select * from view_empgroup where yymm='"& khym &"' ) h on h.empid = b.empid "&_
	"where b.country like '"& F_country &"%' and a.khw like '"&F_whsno &"%'  and a.khg like '"& F_groupid &"%' "&_
	"and a.khz like '"&F_zuno&"%' and a.khs like '%"&F_shift&"' and a.empid like '"&F_empid&"%' "&_
	"order by " & sortvalue   

'sql="select  b.country cstr, b.empnam_cn, b.empnam_vn, b.country, convert(char(10),b.indat,111) as nindat,convert(char(10),outdat,111) as outdate, d.sys_value as gstr, "&_
'	"e.sys_value as zstr, f.sys_value as wstr, g.sys_value as sstr, a.* from "&_
'	"(  "&_
'	"select count(*) as weekcnt, khym, empid, khw, khg, khz, khs , sum(fna+fnb+fnc+fnd ) as monthfen from  empkhb where khym='"& khym &"' "&_
'	"group by khym,empid, khw, khg, khz, khs  "&_
'	") a  "&_
'	"left join ( select *from empfile ) b on b.empid = a.empid   "&_ 
'	"left join ( select* from  basicCode  where func='groupid' ) d on d.sys_type = a.khg "&_
'	"left join ( select* from  basicCode  where func='zuno' ) e on e.sys_type = a.khz "&_
'	"left join ( select* from  basicCode  where func='whsno' ) f on f.sys_type = a.khw "&_
'	"left join ( select* from  basicCode  where func='shift' ) g on g.sys_type = a.khs "&_	 
'	"left join ( select * from view_empgroup where yymm='"& khym &"' ) h on h.empid = b.empid "&_
'	"where b.country like '"& F_country &"%' and a.khw like '"&F_whsno &"%'  and a.khg like '"& F_groupid &"%' "&_
'	"and a.khz like '"&F_zuno&"%' and a.khs like '%"&F_shift&"' and a.empid like '"&F_empid&"%' "&_
'	"order by " & sortvalue 
	
rs.Open SQL, conn, 1, 3 	
%>

<html> 
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="refresh">
</head> 
<body  topmargin="5" leftmargin="5"  marginwidth="0" marginheight="0"   >
<BR><BR>
<%
  filenamestr = "empKHB"&yymm&".xls"
  Response.AddHeader "content-disposition","attachment; filename=" & filenamestr
  Response.Charset ="BIG5"
  Response.ContentType = "Content-Language;content=zh-tw" 
  Response.ContentType = "application/vnd.ms-excel"  
  if fclass="B" then 
  	colnum = 27
  	rpt_title = left(khym,4)&" 年 "&right(khym,2)&" 月 績效考核明細表(週統計)"
  else
  	colnum = 15
  	rpt_title = left(khym,4)&" 年 "&right(khym,2)&" 月 績效考核統計表(月統計)"
  end if 	  
%>
<TABLE CLASS="txt12"   >	
	<tr>
		<td colspan=<%=colnum%> align=center><font size=+1><b></b></font></td>
	</tr>	
	<tr>
		<td colspan=<%=colnum%> align=center><font size=+1><b><%=rpt_title%></b></font></td>
	</tr> 
</table>	

<TABLE CLASS="txt12" BORDER=1 cellspacing="1" cellpadding="1" >		
	<TR HEIGHT=25 BGCOLOR="#e4e4e4">
		<td align=center rowspan=2>STT</td>
		<td align=center rowspan=2>部門</td>
		<td align=center rowspan=2>國籍</td>
		<td align=center rowspan=2>單位</td>
		<td align=center rowspan=2>班別</td>		
		<td align=center rowspan=2>工號</td>
		<td align=center rowspan=2>姓名</td>
		<td align=center rowspan=2>到職日</td>
		<td align=center rowspan=2>離職日</td>
		<td align=center <%if fclass="B" then%>colspan=4<%end if%>>第一週</td>
		<td align=center <%if fclass="B" then%>colspan=4<%end if%>>第二週</td>
		<td align=center <%if fclass="B" then%>colspan=4<%end if%>>第三週</td>
		<td align=center <%if fclass="B" then%>colspan=4<%end if%> >第四週</td>
		<td align=center rowspan=2>總分</td>
		<td align=center rowspan=2>備註</td>
	</tr>
	<tr bgcolor=#e4e4e4>		
		<%
		if khweek<>"" then 
			tz=1
		else
			tz=days\7	
		end if
		for g = 1 to tz%>
			<%if fclass="B"  then %>
			<td align=center width=30 nowrap >A</td>
			<td align=center width=30 nowrap >B</td>
			<td align=center width=30 nowrap >C</td>
			<td align=center width=30 nowrap >D</td>
			<%else%>
				<td   align=center>月統計</td>
			<%end if %>	
		<%next%>
		
	</tr>	 
	<%
	  for x = 1 to rs.recordcount
	%> 	

		<TR HEIGHT=22 BGCOLOR="#ffffff">				
			<td align=left align=center><%=x%></td>
			<td align=left><%=rs("gstr")%></td>
			<td align=left><%=rs("country")%></td>
			<td align=left><%=rs("zstr")%></td>
			<td align=left><%=rs("ls")%></td>
			<td align=left><%=rs("empid")%></td>
			<td align=left nowrap><%=rs("empnam_cn")&rs("empnam_vn")%></td>
			<td align=center>&nbsp;<%=rs("nindat")%></td>
			<td align=center>&nbsp;<%=rs("outdate")%></td>	
			<%
			  tfn = 0 
			  memo=""
			  if khweek="" then 
			  	tt = days\7  			  	
			  else
			  	tt = 1  			  	
			  end if 		
			  for y=1 to tt
				if y mod 2 =0 then 
					weekcolor="#e6e6fa"
				else
					weekcolor="#eee8aa"
				end if 	
				if  khweek="" then 
					sqld="select * from empKHB  where khym='"& khym &"' and empid='"&rs("empid")&"' and khweek='"&y&"' " 
				else
					sqld="select * from empKHB  where khym='"& khym &"' and empid='"&rs("empid")&"' and khweek='"&khweek&"' "
				end if 	
				'response.write sqld
				set rs2=Server.CreateObject("ADODB.Recordset")
				rs2.open sqld, conn, 1,3 
				if rs2.eof then 
					fnA="0"
					fnB="0"
					fnC="0"
					fnD="0"
					colorA="red"
					colorB="red"
					colorC="red"
					colorD="red"
					'memo=""
					memo=memo&""
					muser=""
				else					
					fnA=rs2("fnA")
					fnB=rs2("fnB")
					fnC=rs2("fnC")
					fnD=rs2("fnD")
					'memo=rs2("memo")
					if rs2("memo")="" then 
						memo=memo&rs2("memo")
					else
						memo=memo&rs2("memo")&"<BR>"
					end if
					if fna="0" then colorA="red" else colorA="black"
					if fnb="0" then colorB="red" else colorB="black"
					if fnc="0" then colorC="red" else colorC="black"
					if fnd="0" then colorD="red" else colorD="black"
					muser=rs2("muser")
				end if 				
				
				tfn = tfn + cdbl(fna)+cdbl(fnb)+cdbl(fnc)+cdbl(fnd)
			%> 					
				<%if fclass="B"  then %>
					<td align=center><%=fnA%>
						
					</td>
					<td align=center><%=fnB%>	
						
					</td>	
					<td align=center><%=fnC%>
						
					</td>
					<td align=center><%=fnD%>	
						
					</td> 
				<%else%>
					<Td align=center  ><%=cdbl(fna)+cdbl(fnb)+cdbl(fnc)+cdbl(fnd)%></td>	
				<%end if%>
			<%next%>
			<td align=center><%=rs("monthfen")%></td>
			<td align=left>&nbsp;<%=memo%></td>			
		</tr> 
	<%	 
		rs.movenext
	%> 
	<%next%>    
</table>
<BR> 
<Table>
	<Tr>
		<Td></td>
		<Td align=right> 承辦人</td>
		<Td></td>
		<Td></td>
		<td>
		<Td align=left>班長</td>		
		<Td align=left>單位主管</td>
		<Td></td>
		
		<Td align=center colspan=4>廠經理</td>		
	</tr>
</table>
<%response.end%>

</body>
</html> 
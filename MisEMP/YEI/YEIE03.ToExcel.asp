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
if sortvalue ="" then sortvalue="b.country , a.khw, a.khg, a.empid"

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
 

sql="select  "&_
		"(ov_h1m+ov_h2m+ov_h3m+ov_b3m ) as totov ,     "&_
		"round( (ov_h1m+ov_h2m+ov_h3m+ov_b3m )/ (money_h*2.5),3) ff1  "&_
		", *  "&_
		"from (  "&_
		"			select g.empid,   b.empnam_cn, b.empnam_vn,  "&_
		"			g.yymm, g.money_h,  "&_
		"			isnull(g.h1m,0) h1m,  isnull(g.h2m,0) h2m,  isnull(g.h3m,0) h3m,  isnull(g.b3m,0) b3m  ,  "&_
		"			round(isnull(jb.h1,g.h1)*isnull(g.money_h,0)*1.5,0)  as N_h1m ,  "&_
		"			isnull(g.h1m,0)- round(isnull(jb.h1,g.h1)*isnull(g.money_h,0)*1.5,0)   as ov_h1m, "&_
		"			round(isnull(jb.h2,g.h2)*isnull(g.money_h,0)*2,0)  as N_h2m ,  "&_
		"			isnull(g.h2m,0) - round(isnull(jb.h2,g.h2)*isnull(g.money_h,0)*2,0)  as ov_h2m,  "&_
		"			round(isnull(jb.h3,g.h3)*isnull(g.money_h,0)*3,0)  as N_h3m ,  "&_
		"			isnull(g.h3m,0)- round(isnull(jb.h3,g.h3)*isnull(g.money_h,0)*3,0)  as ov_h3m, "&_
		"			round((isnull(g.b3,0)-isnull(jb.ov_b3,0))*isnull(g.money_h,0)*0.3,0) as N_b3m ,  "&_
		"			isnull(g.b3m,0)-round((isnull(g.b3,0)-isnull(jb.ov_b3,0))*isnull(g.money_h,0)*0.3,0) as ov_b3m  "&_
		"			, lw,lg,lz,lgstr, lzstr, lwstr , ls, b.nindat , b.outdate, b.country "&_
		"			 from    "&_
		"			( select * from empdsalary_bak  where country='"&F_country&"' and whsno='"&f_whsno&"'  and yymm ='"& khym &"'   )  g "&_
		"			left join  ( select * from view_empfile  ) b on b.empid = g.empid "&_
		"			left join ( select* from view_empgroup   ) i on i.empid = g.empid  and i.yymm = g.yymm "&_
		"			left join ( select * from empJBtim   ) jb on jb.yymm = g.yymm and jb.empid = g.empid "&_
		"			where i.lg like '"&f_groupid&"%' "&_
		") z 	"&_
		"order by  lg, empid " 
		
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
  filenamestr = "empKHBSA8K"&yymm&".xls"
  Response.AddHeader "content-disposition","attachment; filename=" & filenamestr
  Response.Charset ="BIG5"
  Response.ContentType = "Content-Language;content=zh-tw" 
  Response.ContentType = "application/vnd.ms-excel"   
  rpt_title = left(khym,4)&" 年 "&right(khym,2)&" 月 考核獎金分數統計表"
  colnum=14
%>
<TABLE CLASS="txt12"   >	
	<tr>
		<td colspan=<%=colnum%> align=center><font size=+1><b>CTY TNHH HOA DUONG</b></font></td>
	</tr>	
	<tr>
		<td colspan=<%=colnum%> align=center><font size=+1><b><%=rpt_title%></b></font></td>
	</tr> 
</table>	

<TABLE CLASS="txt12" BORDER=1 cellspacing="1" cellpadding="1" >		
	<TR HEIGHT=25 BGCOLOR="#e4e4e4">
		<td align=center >STT</td>
		<td align=center >部門</td>
		<td align=center >國籍</td>
		<td align=center >單位</td>
		<td align=center >班別</td>		
		<td align=center >工號</td>
		<td align=center >姓名</td>
		<td align=center >到職日</td>
		<td align=center >時薪</td>
		<td align=center >離職日</td>				
		<td align=center >A</td>
		<td align=center >B</td>
		<td align=center >C</td>
		<td align=center >D</td>
		<td align=center >E</td>
		<td align=center >獎金</td>
	</tr> 
	<%
	  for x = 1 to rs.recordcount  		
		n_ff1 = right("00.000"&replace(formatnumber(rs("ff1"),3),",",""),6)
		FA=left(n_ff1,1)
		FB=mid(n_ff1,2,1)
		FC=mid(n_ff1,4,1)
		FD=mid(n_ff1,5,1)
		FE=mid(n_ff1,6,1)
	%> 	
		<TR HEIGHT=22 BGCOLOR="#ffffff">				
			<td align=left align=center><%=x%></td>
			<td align=left><%=rs("lgstr")%></td>
			<td align=left><%=rs("country")%></td>
			<td align=left><%=rs("lzstr")%></td>
			<td align=left><%=rs("ls")%></td>
			<td align=left><%=rs("empid")%></td>
			<td align=left nowrap><%=rs("empnam_cn")&rs("empnam_vn")%></td>
			<td align=center>&nbsp;<%=rs("nindat")%></td>
			<td align=center><%=rs("money_h")%></td>
			<td align=center>&nbsp;<%=rs("outdate")%></td>				 
			<td align=left><%=FA%></td>
			<td align=left><%=FB%></td>
			<td align=left><%=FC%></td>
			<td align=left><%=FD%></td>
			<td align=left><%=FE%></td>			
			<td align=right><%=formatnumber(rs("totov"),0)%></td>
					 
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
		<Td align=left>班長</td>		
		<td></td>
		<Td align=center>單位主管</td>
		<Td></td>		
		<Td align=center colspan=4>廠經理</td>		
	</tr>
</table>
<%response.end%>

</body>
</html> 
<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%> 
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<!-- #include file = "../GetSQLServerConnection.fun" --> 
<!-- #include file="../ADOINC.inc" -->
<!-- #include file="../Include/func.inc" -->
<!--#include file="../include/checkpower.asp"-->  
<%
Set conn = GetSQLServerConnection()	  
self="YECE1303"  

nowmonth = year(date())&right("00"&month(date()),2)  
if month(date())="1" then  
	calcmonth = year(date())-1&"12"  	
else
	calcmonth =  year(date())&right("00"&month(date())-1,2)   
end if 	   
if day(date())<=11 then 
	if month(date())="1" then  
		calcmonth = year(date())-1&"12" 
	else
		calcmonth =  year(date())&right("00"&month(date())-1,2)   
	end if 	 	
else
	calcmonth = nowmonth 
end if  
yymm = trim(request("YYMM")) 
 
years=request("years")
whsno=request("whsno")
enddat =request("years")&"/12/31"
groupid=request("groupid")
country=request("country")
khud=request("khud") 
if khud="0" then enddat =request("years")&"/06/30"

sql="select isnull(c.years,'"&years&"') as years, a.empid, a.empnam_cn, a.empnam_vn, a.country, convert(char(10),a.indat,111) as indate, b.whsno, b.groupid , "&_ 
		"ix.sys_value as gstr, datediff(m,a.indat,'"&enddat&"')/1.0 as nz  , isnull(c.fensu,'') fensu, isnull(c.kj,'') kj "&_
	  "from "&_
		"( select *  from empfile  where  isnull(outdat,'')=''  and isnull(status,'')<>'D' and country like'"&country&"%' ) a "&_
		"left join (select *from bempg where  yymm=convert(char(6),getdate(),112) ) b on b.empid = a.empid "&_
		"left join (Select * from basiccode where func='groupid' ) ix on ix.sys_type = b.groupid "&_
		"left join (select * from empnzkh where years='"&years&"'  and khud='"&khud&"') c on c.empid = a.empid  "&_
		 
		"where b.whsno='"&whsno &"' and b.groupid like'"&g1&"%' and a.empid like '"&eid&"%' "&_
		"order by b.whsno, a.country, b.groupid, a.empid  "
'response.write sql
'response.end 
Set rs = Server.CreateObject("ADODB.Recordset") 
rs.open sql, conn, 3, 3  
'response.write sql &"<BR>"
'response.end 
	

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
  filenamestr = years&"_"&whsno&"_KH.xls"
  Response.AddHeader "content-disposition","attachment; filename=" & filenamestr
  Response.Charset ="utf-8"
  Response.ContentType = "Content-Language;content=zh-tw" 
  Response.ContentType = "application/vnd.ms-excel"
%>
<TABLE CLASS="txt12" BORDER=1 cellspacing="1" cellpadding="1" >	  
	<tr>
		<td colspan="10" ><%=years%> <%=replace(replace(khud,"0","上"),"1","下")%>考核</td>
	</tr>
	<TR HEIGHT=25 BGCOLOR="#e4e4e4">
		<td align=center>STT</td>
		<td align=center>Xuong</td>
		<td align=center>Bo Phan</td>
		<td align=center>Quoc Tich</td>
		<td align=center>so the</td>
		<td align=center>ho the</td>
		<td align=center>NVX</td>
		<td align=center>年資(M)</td>
		<td align=center>分數</td>
		<td align=center>考績</td>
	</tr> 
<%x=0
 
	while not rs.eof 
	x = x + 1   
%>	
	<tr>
		<Td align="center"><%=x%></td>
		<Td align="left"><%=rs("whsno")%></td>
		<Td align="center"><%=rs("gstr")%></td>
		<Td align="center"><%=rs("country")%></td>
		<Td align="center"><%=rs("empid")%></td>
		<Td align="left" nowrap><%=rs("empnam_cn")%>&nbsp;<%=rs("empnam_vn")%></td>
		<Td align="center"><%=rs("indate")%></td>
		<Td align="center"><%=rs("nz")%></td>
		<Td align="center"><%=rs("fensu")%></td>
		<Td align="center"><%=rs("kj")%></td> 
	</tr>	
<%
rs.movenext
wend 
set rs=nothing 
%>	 
</table> 
<%response.end%>

</body>
</html> 
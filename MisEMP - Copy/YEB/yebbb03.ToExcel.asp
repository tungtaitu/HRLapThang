<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%> 
<!-- #include file = "../GetSQLServerConnection.fun" --> 
<!-- #include file="../ADOINC.inc" -->
<!-- #include file="../Include/func.inc" -->
<!--#include file="../include/checkpower.asp"-->  
<%
response.buffer=true 
Set conn = GetSQLServerConnection()	  
self="YECP0302"  


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
 
yymm=request("yymm")
yymm2=request("yymm2")
if request("yymm2")="" then yymm2=yymm
whsno=request("whsno")
country=request("country")
groupid=request("groupid")
empid1=request("empid1")
rpno=request("F_rpno")
rp_type = request("rp_type")
sortby = request("sortby")
sortby="rp_dat, rpno  "  


if yymm="" and whsno="" and country="" and groupid="" and empid1="" and rpno="" then 
	sql="select c.sys_value,b.whsno,b.empnam_cn,b.empnam_vn,b.nindat,b.gstr,b.zstr,b.groupid,b.zuno,b.job,b.jstr,a.* from "&_
		"(select convert(char(10),rp_dat, 111) as rp_date, * from emprepe where isnull(status,'')<>'D' and convert(char(6),rp_dat,112)='xxx' ) a "&_
		"left join ( select *from view_empfile ) b on b.empid = a.empid "&_
		"left join ( select *from basicCode ) c on c.func=case when a.rp_type='R' then 'goods' else case when a.rp_type='P' then 'bads' else '' end end  and c.sys_type = a.rp_func "&_
		"where a.rpwhsno like '"& session("rpwhsno") &"%' and  b.groupid like '"& groupid &"%' and b.country like '"&country&"%' "&_
		"and a.empid like '"& empid1 &"%' and a.rp_type like '"&rp_type&"%' "&_
		"order by " & sortby 
else 
	if yymm<>""  then 
		sql="select c.sys_value,b.whsno,b.empnam_cn,b.empnam_vn,b.nindat,b.gstr,b.zstr,b.groupid,b.zuno,b.job,b.jstr,a.* from "&_
				"(select convert(char(10),rp_dat, 111) as rp_date, * from emprepe where isnull(status,'')<>'D' and convert(char(6),rp_dat,112) between '"& yymm &"' and '"&yymm2&"' ) a "
	else 
		sql="select c.sys_value,b.whsno,b.empnam_cn,b.empnam_vn,b.nindat,b.gstr,b.zstr,b.groupid,b.zuno,b.job,b.jstr,a.* from "&_
				"(select convert(char(10),rp_dat, 111) as rp_date, * from emprepe where isnull(status,'')<>'D' and convert(char(6),rp_dat,112) like  '"& yymm &"%'  ) a  "
	end if  

	sql=sql&"left join ( select *from view_empfile ) b on b.empid = a.empid "&_
			"left join ( select *from basicCode ) c on c.func=case when a.rp_type='R' then 'goods' else case when a.rp_type='P' then 'bads' else '' end end  and c.sys_type = a.rp_func "&_
			"where a.rpwhsno like '"& WHSNO &"%' and  b.groupid like '"& groupid &"%' and b.country like '"&country&"%' "&_
			"and a.empid like '%"& empid1 &"%' and left(rpno,2) like '"& left(rpno,2) &"%'and a.rp_type like '"&rp_type&"%' "&_
			"order by  " & sortby 	
end if 
Set rs = Server.CreateObject("ADODB.Recordset")  
rs.open sql, conn, 3, 3 

response.write sql 
'response.end 

%>

<html> 
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="refresh">
<style >
.txtvn8 { FONT-FAMILY: VNI-Times;  FONT-SIZE: 8pt }  
.txtvn9 { FONT-FAMILY: VNI-Times;  FONT-SIZE: 10pt }  
.txt9 { FONT-FAMILY: Times New Roman;  FONT-SIZE: 9pt }  
</style>
</head> 
<body  topmargin="5" leftmargin="5"  marginwidth="0" marginheight="0"   >
<BR><BR>
<%
  filenamestr = "emptp_BB03"&yymm&"_"&yymm2&".xls"
  Response.AddHeader "content-disposition","attachment; filename=" & filenamestr
  Response.Charset ="BIG5"
  Response.ContentType = "Content-Language;content=zh-tw" 
  Response.ContentType = "application/vnd.ms-excel"
  
	'Response.ContentType = "application/msword"
%>
<TABLE CLASS="txt9" BORDER=0 cellspacing="1" cellpadding="1" >	
	<tr>
		<font size=+1><b>CTY TNHH GIAY YUEN FOONG YU(<%=session("mywhsno")%> )</b></font>
		<br>
		<font size=+1><b><%=yymm%>~<%=yymm2%> 員工獎懲明細表</b></font></tr> 
</table>	
<TABLE CLASS="txt9" BORDER=1 cellspacing="1" cellpadding="1" >		
	<TR HEIGHT=25 BGCOLOR="#e4e4e4" class="txt9">	
		<Td >STT</td>
		<Td >獎懲<br>thuong phat</td>
		<Td >編號<br>so<br></td>
		<Td >事件日期<br>Ngay</td>
		<Td >工號<br>so the</td>
		<Td >姓名<br>ho ten</td>
		<Td >部門<br>bo phan</td>		
		<Td >方式<br>phuong thuc</td>
		<Td >內容<br>thuyet minh</td>
		<Td >處理說明<br>phuong thuc xu ly </td>		
		<Td >文件編號<br>so van kien</td>		
	 
	</tr>
	<%response.flush%>
	<%x = 0 
	  grp_cnt = 0	
	  grp_amt = 0 
	  grp_id = ""
	  tot_amt= 0 
	  while not rs.eof   
			x=x+1 
			if rs("rp_type")="R" then
				rpstr = rs("rp_type")&" 獎勵"
			elseif rs("rp_type")="P" then 
				rpstr =  rs("rp_type")&" 懲罰"
			else
				rpstr = ""
			end if 	 		
	%> 	

		<TR HEIGHT=22 BGCOLOR="#ffffff" class="txt9">				
			<td align=left><%=x%></td>
			<td><%=rpstr%></td>
			<td align=left><%=rs("rpno")%></td>
			<td align=left><%=rs("rp_date")%></td>
			<td align=left><%=rs("empid")%></td>
			<td align=left><%=rs("empnam_cn")%>&nbsp;<%=rs("empnam_vn")%></td>
			<td align=left><%=rs("gstr")%></td>
			<td align=left><%=rs("rp_func")%>&nbsp;<%=rs("sys_value")%></td>
			<td align=left><%=rs("rpmemo")%></td>
			<td align=left><%=rs("rp_method")%>&nbsp;</td>
			<td align=left><%=rs("FILENO")%>&nbsp;</td>
 
		</tr> 
	<%	 
	rs.movenext
	%> 
	<%wend%>  
	
</table> 
<%
rs.close
set rs=nothing
conn.close 
set conn=nothing
response.end%>

</body>
</html> 
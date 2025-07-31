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
  
f_years = request("f_years")
f_country = request("f_country")
f_whsno = request("f_whsno")  

sql="select  x1.sys_value as gstr, case when months>=12 then "&_
		"datediff(m,indat, case when    isnull(outdat,'')=''  or convert(char(6),b.outdat,112) >'"&f_years&"'+'12' then  '"&f_years&"'+'/12/31' else convert(char(10),b.outdat,111)   end ) "&_ 
		"else months  end  dd,  convert(char(10),a.indat,111) indate, datediff(d,a.indat,'"&f_years&"'+'/12/31')/30.0 as rnzm, "&_
		"case when isnull(a.empid,'')='' then d.bb+d.cv+d.phu+d.nn+d.kt+d.mt+d.wp+case when a.country='TA' then ttkh else 0 end else a.totamAmt end as BasicS,  "&_
		"x2.sys_value as jstr,  ISNULL(x2.no1,0) days_add , * "&_
		",df_days=( select top 1 days  from empnzjj_set where years='"&f_years&"' and kj='甲'  and country='VN' and whsno='"&f_whsno&"') "&_
		",hris_add = case when isnull(f8,0) =0 then 0 else 5 end "&_
		",erp_add = case when isnull(f9,0) =0 then 0 else 5 end "&_
		",test_add = case when isnull(f10,0) =0 then 0 else 3 end "&_
		",prj_add = case when isnull(f8,0) =0 then 0 else 5 end + case when isnull(f9,0) =0 then 0 else 5 end + case when isnull(f10,0) =0 then 0 else 3 end "&_
		",heso2=cast( isnull(heso2,1) as decimal(9,2)) "&_
		"from "&_
		"(select * from empnzjj where yymm='"& f_years&"' and   whsno = '"&f_whsno&"'  "&_
		"and ( country = '"& f_country &"' or case when country in('TW','MA') then 'TM' else country end  = '"& f_country &"' "&_
		"or case when country ='VN' then country else 'HW' end = '"& f_country &"' ) "&_
		") a " &_
		"left join ( select empid, wkws,  outdat , showname = case when empnam_cn<>'' then empnam_cn else '' end+' '+empnam_vn from view_empfile ) b on b.empid = a.empid   " &_
		"left join ( select * from basicCode where func='groupid'  ) x1 on x1.sys_type = a.groupid   " &_
		"left join ( select * from bemps where  yymm='"&f_years&"'+'12' ) d on d.empid = a.empid "&_		
		"left join ( select * from bempj where yymm='"&f_years&"'+'12' ) e on e.empid = a.empid "&_
		"left join ( select * from basicCode  where func='lev'  ) x2 on x2.sys_type = e.job "&_
		"left join ( select f2 as ws, f3 as empid, f8 , f9 , f10  from  empnzjj_prj ) s on s.ws = a.whsno and s.empid = a.EMPID "&_
		"left join ( select  f1 ws, f3 empid, f8 loai, f12 k_hr, heso2=f13 , k_mm = f14  from  empnzjj_3tn3t ) t on t.ws=a.whsno and t.empid = a.EMPID "&_
		"order by a.country, a.groupid,a.indat, a.empid " 

'sql="exec sp_api_calcnzjj "
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
   filenamestr = F_years&"_"&f_whsno&"_NZJJ_"&minute(now)&second(now)&".xls"
   Response.AddHeader "content-disposition","attachment; filename=" & filenamestr
   Response.Charset ="utf-8"
   Response.ContentType = "Content-Language;content=zh-tw" 
   Response.ContentType = "application/vnd.ms-excel"
%>
<TABLE CLASS="txt12"   >	  
	<tr>
		<b><%=F_years%> (<%=f_whsno%>)廠 (<%=f_country%>)年終獎金</b>
	</tr>
</table>	
<TABLE CLASS="txt12" BORDER=1 cellspacing="1" cellpadding="1" >	  
	<TR HEIGHT=25 BGCOLOR="#e4e4e4">
		<td align=center>STT</td> 
		<td align=center>WKWS</td>
		<td align=center>Xuong</td>
		<td align=center>Bo Phan</td>
		<td align=center>Quoc<br>Tich</td>
		<td align=center>so the</td>
		<td align=center>ho the</td>
		<td align=center>CVID</td>
		<td align=center>chu vu</td>
		<td align=center>NVX</td>
		<td align=center>實際年資</td>
		<td align=center>年資(M)</td>		
		<td align=center>基本<br>co ban</td>
		<td align=center>職務<br>CV</td>
		<td align=center>考績</td>
		<td align=center>係數</td>
		<td align=center>基準</td>
		<td align=center>年獎</td>
		<td align=center>稅金</td>
		<td align=center>調整</td>
		<td align=center>年終獎金</td>
		<td align=center>備註說明</td>
		<td align=center>基準(日)</td>
		<td align=center>係數1</td>
		<td align=center>考績+天</td>
		<td align=center>HRIS+天</td>
		<td align=center>ERP+天</td>
		<td align=center>TEST+天</td>		
		<td align=center>K3T係數</td>
		<%if f_country<>"VN" then%>
		<td align=center>BB</td>
		<td align=center>CV</td>
		<td align=center>PHU</td>		
		<td align=center>NN</td>
		<td align=center>KT</td>
		<td align=center>MT</td>
		<td align=center>WP</td>		
		<td align=center>TT</td>		
		<%end if%>
		 
	</tr> 
<%x=0
 
	while not rs.eof 
	x = x + 1 
	if	rs("days")="0" or isnull(rs("days"))   then 
		kj_add=0 
	else 
		kj_add = cdbl(rs("days"))-cdbl(rs("df_days"))
	end if 
%>	
	<tr>
		<Td align="center"><%=x%></td>
		<Td align="left"><%=rs("wkws")%></td>
		<Td align="left"><%=rs("whsno")%></td>
		<Td align="center"><%=rs("gstr")%></td>
		<Td align="center"><%=rs("country")%></td>
		<Td align="center"><%=rs("empid")%></td>
		<Td align="left" nowrap><%=rs("showname")%></td>
		<Td align="left "><%=rs("job")%></td>
		<Td align="left"><%=rs("jstr")%></td>
		<Td align="left"><%=rs("indate")%></td>
		<Td align="center"><%=round(rs("rnzm"),0)%></td>		
		<Td align="center"><%=rs("nz")%></td>	
		<Td align="center"><%=rs("days")%></td> 	
		<Td align="center"><%=rs("days_add")%></td> 		
		<Td align="center"><%=rs("grande")%></td> 
		<Td align="right"><%=rs("js")%></td>
		<Td align="right"><%=formatnumber(rs("basicS"),0)%></td> 
		<Td align="right"><%=formatnumber(rs("bonus"),0)%></td> 
		<Td align="right"><%=formatnumber(rs("ktaxm"),0)%></td> 
		<Td align="right"><%=formatnumber(rs("tjamt"),0)%></td> 
		<Td align="right"><%=formatnumber(rs("realamt"),0)%></td> 
		<Td align="left"><%=rs("memos")%></td>
		<Td align="right"><%=formatnumber(cdbl(rs("basicS"))/30.0,0)%></td>  
		<Td align="right"><%=formatnumber(cdbl(rs("days"))/30.0,2)%></td> 
		<Td align="left"><%=kj_add%></td>
		<Td align="left"><%=rs("hris_add")%></td>
		<Td align="left"><%=rs("erp_add")%></td>
		<Td align="left"><%=rs("test_add")%></td>
		<!--Td align="left"><%=rs("prj_add")%></td-->
		<Td align="left"><%=rs("heso2")%></td>
		<%if f_country<>"VN" then%>
		<Td align="right"><%=formatnumber(rs("bb"),0)%></td> 
		<Td align="right"><%=formatnumber(rs("cv"),0)%></td> 
		<Td align="right"><%=formatnumber(rs("phu"),0)%></td> 
		<Td align="right"><%=formatnumber(rs("nn"),0)%></td> 
		<Td align="right"><%=formatnumber(rs("kt"),0)%></td> 
		<Td align="right"><%=formatnumber(rs("mt"),0)%></td> 		
		<Td align="right"><%=formatnumber(rs("wp"),0)%></td> 
		<Td align="right"><%=formatnumber(rs("ttkh"),0)%></td> 
		<%end if%>
		<%if f_country<>"VN" then
			s1 = cdbl(rs("realamt"))\100
			s2= (cdbl(rs("realamt"))-s1*100) \ 50 
			s3= (cdbl(rs("realamt"))-s1*100-s2*50)\ 20
			s4= (cdbl(rs("realamt"))-s1*100-s2*50-s3*20)\ 10
			s5= (cdbl(rs("realamt"))-s1*100-s2*50-s3*20-s4*10)\ 5
			s6= (cdbl(rs("realamt"))-s1*100-s2*50-s3*20-s4*10-s5*5)\1
		
		%>	<!--
			<td><%=s1%></td>
			<td><%=s2%></td>
			<td><%=s3%></td>
			<td><%=s4%></td>
			<td><%=s5%></td>
			<td><%=s6%></td>
			-->
		<%else 
			s1 = cdbl(rs("realamt"))\500000
			s2= (cdbl(rs("realamt"))-s1*500000) \ 200000 
			s3= (cdbl(rs("realamt"))-s1*500000-s2*200000)\ 100000
			s4= (cdbl(rs("realamt"))-s1*500000-s2*200000-s3*100000)\ 50000
			s5= (cdbl(rs("realamt"))-s1*500000-s2*200000-s3*100000-s4*50000)\ 20000
			s6= (cdbl(rs("realamt"))-s1*500000-s2*200000-s3*100000-s4*50000-s5*2000)\10000
			s7= (cdbl(rs("realamt"))-s1*500000-s2*200000-s3*100000-s4*50000-s5*2000-s6*10000)\5000		
		%>	<!--
			<td><%=s1%></td>
			<td><%=s2%></td>
			<td><%=s3%></td>
			<td><%=s4%></td>
			<td><%=s5%></td>
			<td><%=s6%></td>
			<td><%=s7%></td>
			-->
		<%end if%>
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
<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%> 
<!-- #include file = "../GetSQLServerConnection.fun" --> 
<!-- #include file="../ADOINC.inc" -->
<!-- #include file="../Include/func.inc" -->
<!--#include file="../include/checkpower.asp"-->  
<%
Set conn = GetSQLServerConnection()	  
self="YECE06"  


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
 
YYMM =  request("YYMM")
code01=request("YYMM")
code02=request("country")
code03=request("whsno")
code04=request("groupid")
code05=request("job")
code06=request("empid1")

sql="exec rpt_empBasicSalary '"& code01 &"',  '"& code02 &"', '"& code03 &"', '"& code04 &"', '"& code05 &"', '"& code06 &"' "
set rs=conn.execute(Sql)

'response.write sql 
'response.end 
if instr(session("vnlogip"),"168")>0 then 
	w1="LA"
elseif instr(session("vnlogip"),"169")>0 then 
	w1="DN"	
elseif instr(session("vnlogip"),"1")>0 then 
	w1="BC"	
end if 	 

w1=session("myshno")

tmpRec  = Session("YFYEMPJXM") 
pagerec = request("pagerec") 
jxym = request("jxym")
SALARYYM = request("SALARYYM")


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
  filenamestr = "EmpJX"&jxym&".xls"
  Response.AddHeader "content-disposition","attachment; filename=" & filenamestr
  Response.Charset ="BIG5"
  Response.ContentType = "Content-Language;content=zh-tw" 
  Response.ContentType = "application/vnd.ms-excel"
	
	
  'Response.ContentType = "application/msword"
%>
<TABLE CLASS="txt12" BORDER=0 cellspacing="1" cellpadding="1" >	
	<tr>
		<td align=left colspan=10><font size=+1><b></b></font></td>
	</tr>
	<tr>
		<td align=left colspan=10><font size=+1><b>Bang Tien Thuong Thang &nbsp; <%=right(jxym,2)%> &nbsp; Nam &nbsp;  <%=left(jxym,4)%></b> </font></td>
	</tr>	
</table>	
<TABLE CLASS="txt12" BORDER=1 cellspacing="1" cellpadding="1" >		
	<TR HEIGHT=25 BGCOLOR="#e4e4e4">
		<td align=center>STT</td>		
		<td align=center>Ma so<br>B.P</td>		
		<td align=center>Bo Phan</td>
		<td align=center>組代碼</td>		
		<td align=center>組別</td>		
		<td align=center>Ca</td>		
		<td align=center>So The</td>
		<td align=center>Ten</td>		
		<td align=center>NVX</td>				
		<td align=center>Ma<BR>CV</td>
		<td align=center>chu vu</td>		
		<td align=center>T.T<br>H.SO</td>
		<td align=center>Co ban T.T</td>		
		<td align=center >H.SO2</td>		
		<td align=center >H.SO3</td>				
		<td align=center>TOT</td>		
		<td align=center >實際<br>工時</td>		
		<td align=center >應出勤<br>工時</td>			
		<td align=center>VR<BR>(Hr)</td>		
		<td align=center>KC<BR>(Hr)</td>				
		<td align=center>TOT Tru</td>				
		<td align=center>勸導<br>單</td>				
		<td align=center>Total</td>				
	</tr>
	<% 
	tdc_amt = 0 
	sdc_amt = 0 
	for x =  1 to pagerec   	 
		hso = trim(request("workJs")(x)) 
		hrhso = trim(request("hrjs")(x)) 
		nedhr = trim(request("nedhr")(x)) 
		relHr = trim(request("relHr")(x)) 
		Total_jx = cdbl(hso) * cdbl(tmpRec(1,x,19)) * cdbl(hrhso)
		tdc_amt = tdc_amt + cdbl(tmpRec(1,x,19)) 
		sdc_amt = sdc_amt + cdbl(Total_jx) 
	%> 	

		<TR HEIGHT=22 BGCOLOR="#ffffff">				
			<td align=left><%=x%></td>
			<td align=left><%=tmprec(1,x,2)%></td>
			<td align=left><%=tmprec(1,x,43)%></td>
			<td align=left><%=tmprec(1,x,26)%></td>
			<td align=left><%=tmprec(1,x,28)%></td>
			<td align=left><%=tmprec(1,x,3)%></td>
			<td align=left><%=tmprec(1,x,1)%></td>
			<td align=left><%=tmprec(1,x,5)%><%=tmprec(1,x,6)%></td>
			<td align=left><%=tmprec(1,x,7)%></td> <!--indate-->
			<td align=left><%=tmprec(1,x,10)%></td>
			<td align=left><%=tmprec(1,x,11)%></td>
			<td align="center"><%=tmprec(1,x,4)%></td> <!--職務系數-->
			<td align="right"><%=formatnumber(tmprec(1,x,12),0)%></td> 			
			<td align="center"><%=formatnumber(hso,1)%></td>
			<td align="center"><%=formatnumber(hrhso,2)%></td>
			<td align="right"><%=formatnumber(tmprec(1,x,18),0)%></td> <!--績效獎金已*職務系數*工時係數--> 			
			<td align="center"><%=formatnumber(relhr,1)%></td>
			<td align="center"><%=formatnumber(nedhr,1)%></td>						
			<td align="right"><%=formatnumber(tmprec(1,x,31),1)%></td><!--事假-->
			<td align="right"><%=formatnumber(tmprec(1,x,8),1)%></td><!--曠職-->			
			<td align="right"><%=formatnumber(tmprec(1,x,17),0)%></td><!--應扣款-->
			<td align="center"><%=tmprec(1,x,33)%></td>		<!--勸導單-->			
			<td align="right"><%=formatnumber(Total_jx,0)%></td> 
		</tr> 
	<%	 
	next 
	%>  
	<tr>		
		<td colspan=22 align="right">Total</td>
		<td  align="right"><%=formatnumber(tdc_amt,0)%></td>
	</tr>	
</table> 
<%response.end%>

</body>
</html> 
<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%> 
<!-- #include file = "../GetSQLServerConnection.fun" --> 
<!-- #include file="../ADOINC.inc" -->
<!-- #include file="../Include/func.inc" -->
<!--#include file="../include/checkpower.asp"-->  
<%
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
 
YYMM =  request("YYMM")
code01=request("YYMM")
'code02=request("country")
'code02="'"&replace(replace(request("country")," ","'"),",","',")&"'"  
code02=replace(request("country")," ","")

code03=request("whsno")
code04=request("groupid")
code05=request("job")
code06=request("empid1") 



sql="exec rpt_empBasicSalary_new '"& code01 &"',  '"& code02 &"', '"& code03 &"', '"& code04 &"', '"& code05 &"', '"& code06 &"' "
set rs=conn.execute(Sql)

'response.write sql 
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
  filenamestr = "Basicsalary"&yymm&".xls"
  Response.AddHeader "content-disposition","attachment; filename=" & filenamestr
  Response.Charset ="BIG5"
  Response.ContentType = "Content-Language;content=zh-tw" 
  Response.ContentType = "application/vnd.ms-excel"
  
	'Response.ContentType = "application/msword"
%>
<TABLE CLASS="txt12" BORDER=1 cellspacing="1" cellpadding="1" >	
	<tr>
		<td colspan=22 align=center><font size=+1><b></b></font></td>
	</tr>
	<tr>
		<td colspan=22 align=center><font size=+1>基本薪資明細表</font></td>
	</tr>
	<tr>
		<td colspan=22 align=center><font size=+1><b>THÁNG &nbsp; <%=right(yymm,2)%> &nbsp; NĂM &nbsp;  <%=left(yymm,4)%></b></font></td>
	</tr> 
	
	<TR HEIGHT=25 BGCOLOR="#e4e4e4">
		<td align=center>卡號<br>STT</td>
		<td align=center>廠別<br>Xuong</td>
		<td align=center>部門<br>Bo Phan</td>
		<td align=center>國籍<br>Quoc tich</td>
		<td align=center>工號<br>Số thẻ</td>
		<td align=center>姓名<br>Họ Tên</td>
		<td align=center>職等<br>Chức vụ</td>
		<td align=center>到職日<br>Ngày vào xưởng</td>
		<td align=center>年資<br>So thang<br>lam viec</td>		
		<td align=center>簽合同<br>Ký hợp đồng</td>
		<td align=center>基本薪<br>Lương cơ bản</td>
		<td align=center>電話津貼<br>PC điện thoại</td>
		<td align=center>職務加給<br>PC chức vụ</td>
		<td align=center>燃油津貼<br>PC xăng xe</td>
		<td align=center>技術<br>Kỹ thuật</td>
		<td align=center>環境<br>Môi trường</td>
		<td align=center>住房支持<br>Hỗ trợ nhà ở</td>
		<td align=center>補薪<br>Bu Luong</td>
		<td align=center>全勤獎金<br>Chuyên cần</td>
		<td align=center>薪資合計<br>Tổng lương</td>
		<td align=center>離職日期<br>Ngày thôi việc</td>
		<td align=center>備註<br>Ghi chu</td>
	</tr>
	<%x = 0 
	  grp_cnt = 0	
	  grp_amt = 0 
	  grp_id = ""
	  tot_amt= 0 
	  while not rs.eof   
	  x=x+1
	%> 	

		<TR HEIGHT=22 BGCOLOR="#ffffff">				
			<td align=left><%=x%></td>
			<td align=left><%=rs("d_whsno")%></td>
			<td align=left><%=rs("D_groupid")%></td>
			<td align=left><%=rs("country")%></td>
			<td align=left><%=rs("empid")%></td>
			<td align=left nowrap><%=rs("empnam_cn")&rs("empnam_vn")%></td>
			<td align=left><%=rs("job")%><%=rs("cjob_str")%></td>
			<td align=left><%=rs("nindat")%></td>
			<td align=right><%=rs("empnz")%></td>
			<td align=left>&nbsp;<%=rs("bhdate")%></td>
			<td align=right><%=formatnumber(rs("BB"),0)%></td>
			<td align=right><%=formatnumber(rs("PHU"),0)%></td>
			<td align=right><%=formatnumber(rs("CV"),0)%></td>
			<td align=right><%=formatnumber(rs("NN"),0)%></td>
			<td align=right><%=formatnumber(rs("KT"),0)%></td>
			<td align=right><%=formatnumber(rs("MT"),0)%></td>
			<td align=right><%=formatnumber(rs("TTKH"),0)%></td>
			<td align=right><%=formatnumber(rs("btien"),0)%></td>
			<td align=right><%=formatnumber(rs("QC"),0)%></td>
			<td align=right><%=formatnumber(rs("TOT"),0)%></td>
			<td align=right>&nbsp;<%=rs("outdate")%></td>
			<td align=centet><%=rs("smemo")%></td>
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
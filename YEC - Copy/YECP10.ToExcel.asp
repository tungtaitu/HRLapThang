<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%> 
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<!-- #include file = "../GetSQLServerConnection.fun" --> 
<!-- #include file="../ADOINC.inc" -->
<!-- #include file="../Include/func.inc" -->
<!--#include file="../include/checkpower.asp"-->  
<%
Set conn = GetSQLServerConnection()	  
self="YECP10"  


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

 '一個月有幾天
cDatestr=CDate(LEFT(YYMM,4)&"/"&RIGHT(YYMM,2)&"/01")
days = DAY(cDatestr+(32-DAY(cDatestr))-DAY(cDatestr+(32-DAY(cDatestr))))   '一個月有幾天
'本月最後一天
ENDdat = CDate(LEFT(YYMM,4)&"/"&RIGHT(YYMM,2)&"/"&DAYS)
ENDdat=year(ENDdat)&"/"&right("00"&month(Enddat),2)&"/"&right("00"&day(Enddat),2) 

 
code01=request("YYMM")
code02=request("country")
code03=request("whsno")
code04=request("groupid")
code05=request("empid1")
code06=request("SY") 

sqln ="select * from empbh_set where w='"& code03 &"'  and ym='"& code01 &"' " 
set ors=conn.execute(Sqln)
if ors.eof then 
	response.write "本月保險計算未設定(CE/2/2.1)" 
	response.end 
else
	setstr = left(ors("setstr"),len(ors("setstr"))-1) 
	c_cols=split(replace(ors("setstr"),"+",","),",")
	'response.write c_cols &"<BR>"
end if 
set ors=nothing   

allcols = ubound(c_cols) '欄位數
redim A1(allcols,2)
for k = 1 to ubound(c_cols) 
	showCols = showCols &  c_cols(k-1)&" as C"& k &","  	 
	A1(k,0)= c_cols(k-1) 
next  
TableRec = TableRec + cdbl(allcols)+5 
Session("a1cols") = A1 
'response.write allcols & "----" & showCols  
sql="select "  	
for  xx = 1 to allcols 
	sql=sql & "isnull(c.c"&xx &",0) as c"&xx &","
next 


sql="exec rpt_empbhgt '"& code01 &"' , '"& code02 &"', '"& code03 &"','"& code04 &"', '"& code05 &"', '"& code06 &"'  "
'Set rs = Server.CreateObject("ADODB.Recordset") 
set rs=conn.execute(Sql)
'response.write sql  
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
  filenamestr = "EmpBHGT"&yymm&".xls"
  Response.AddHeader "content-disposition","attachment; filename=" & filenamestr
  Response.Charset ="BIG5"
  Response.ContentType = "Content-Language;content=zh-tw" 
  Response.ContentType = "application/vnd.ms-excel"
%>
<TABLE CLASS="txt12" BORDER=1 cellspacing="1" cellpadding="1" >	
	<tr>
		<td align=left colspan=<%=15+cdbl(allcols)%>><font size=+1><b></b></font></td>
	</tr>
	<tr>
		<td align=left colspan=<%=15+cdbl(allcols)%>><font size=+1>員工工團與保險</font></td>
	</tr>
	<tr>
		<td align=left colspan=<%=15+cdbl(allcols)%>><font size=+1><b>THÁNG &nbsp; <%=right(yymm,2)%> &nbsp; NĂM &nbsp;  <%=left(yymm,4)%></b></font></td>
	</tr> 
	
	<TR HEIGHT=25 BGCOLOR="#e4e4e4">
		<td align=center>卡號<br>STT</td>
		<td align=center>部分<br>Bộ phận</td>
		<td align=center>工號<br>Số thẻ</td>
		<td align=center>姓名<br>Họ Tên</td>
		<td align=center>到職日<br>Ngày vào xưởng</td>
		<td align=center>年資(M)<br>Thâm niên (tháng)</td>
		<td align=center>職等<br>Chức vụ</td>
		<td align=center>簽合同<br>Ký hợp đồng</td>
		<%
			for z = 1 to allcols 
			Select case a1(z,0) 
				Case "BB"
					 myText = "基薪<br>Cơ bản"
				Case "CV"
					 myText = "職務加給<br>Chức vụ"
				Case "KT"
					 myText = "技術<br>Kỷ thuật"
				Case "MT"
					 myText = "環境<br>Môi trường"
				Case "NN"
					 myText = "燃油津貼<br>PC xăng xe"
				Case "PHU"
					 myText = "電話津貼<br>PC điện thoại"
				Case Else 
					 myText a1(z,0) 	
			end Select
		%>
				<td align=center><%=myText%></td>
		<%next%>
		<td align=center>保險工資<br>Lương tính BH</td>
		<td align=center>BHXH</td>
		<td align=center>BHYT</td>
		<td align=center>BHTN</td>
		<td align=center>全部的<br>Tổng</td>
		<td align=center>工團費<br>Công đoàn</td>
		<td align=center>離職日期<br>Ngày thôi việc</td>
	</tr> 
<%x=0
	dim t1, t2, t3, t4, t5 , t6
	while not rs.eof 
	x = x + 1  
	nz = datediff("m",rs("indate"),ENDdat)
	T1 = t1+cdbl(rs("bhxh5"))
	T2 = t2+cdbl(rs("bhyt1"))
	T3 = t3+cdbl(rs("bhtn1"))
	T4 = t4+cdbl(rs("bhtot"))
	T5 = t5+cdbl(rs("gtamt"))
	T6 = t6+cdbl(rs("bhp"))
	
%>	
	<tr>
		<Td align="center"><%=x%></td>
		<Td align="left"><%=rs("gstr")%></td>
		<Td align="center"><%=rs("empid")%></td>
		<Td align="left" nowrap><%=rs("empnam_cn")%>&nbsp;<%=rs("empnam_vn")%></td>
		<Td align="center"><%=rs("indate")%></td>
		<Td align="center"><%=nz%></td>
		<Td align="center"><%=rs("lj")%></td>
		<Td align="center"><%=rs("bhdat")%></td>
		<%for z1 = 1 to allcols 
			colsname="C"&z1 
			'response.write colsname
			clos_value = rs(colsname)
		%>
			<td align="right"><%=formatnumber(clos_value,0)%></td>
		<%next%>
		<Td align="right"><%=formatnumber(rs("bhp"),0)%></td>
		<Td align="right"><%=formatnumber(rs("bhxh5"),0)%></td>
		<Td align="right"><%=formatnumber(rs("bhyt1"),0)%></td>
		<Td align="right"><%=formatnumber(rs("bhtn1"),0)%></td>
		<Td align="right"><%=formatnumber(rs("bhtot"),0)%></td>
		<Td align="right"><%=formatnumber(rs("gtamt"),0)%></td>
		<Td align="center"><%=rs("outdate")%>&nbsp;</td>
	</tr>	
<%
rs.movenext
wend 
rs.close
set rs=nothing 
conn.close
set conn=nothing
%>	
<tr>
	<td colspan="<%=8+cdbl(allcols)%>">&nbsp;</td>
	<Td align="right"><%=formatnumber(t6,0)%></td>
	<Td align="right"><%=formatnumber(t1,0)%></td>
	<Td align="right"><%=formatnumber(t2,0)%></td>
	<Td align="right"><%=formatnumber(t3,0)%></td>
	<Td align="right"><%=formatnumber(t4,0)%></td>
	<Td align="right"><%=formatnumber(t5,0)%></td>
</tr>
</table> 
<%response.end%>

</body>
</html> 
<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%> 
<!-- #include file = "../GetSQLServerConnection.fun" --> 
<%response.buffer=true%>
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
yymm = trim(request("YYMM"))
 
code01=request("YYMM")
code02=request("country")
code03=request("whsno")
code04=request("groupid")
code05=request("job")
code06=request("empid1")   
rptType = request("rptType")   
jxym=request("jxym") 


if rptType="A" then 
	sql="exec rpt_empdsalaryN '"& code01 &"' , '"& code02 &"', '"& code03 &"','"& code04 &"', '"& code05 &"', '"& code06 &"' ,'', '' "
elseif rptType="B" then 
	sql="exec RPT_YFYMYJX '', '"& code01 &"' , '"&jxym&"' , '"& code02 &"', '"& code03 &"','"& code04 &"', '',  '"& code06 &"' ,'', 'A' "
end if 	
'Set rs = Server.CreateObject("ADODB.Recordset") 
set rs=conn.execute(Sql)
response.write sql  
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
  if rpttype="A" then 
		filenamestr = "salaryVn"&yymm&".xls"
	else
		filenamestr = "TienThuong"&yymm&".xls"
	end if 
  Response.AddHeader "content-disposition","attachment; filename=" & filenamestr
  Response.Charset ="utf-8"
  Response.ContentType = "Content-Language;content=zh-tw" 
  Response.ContentType = "application/vnd.ms-excel"
%>
<TABLE class="txtvn9" BORDER=1 cellspacing="1" cellpadding="1" >	
	<tr>
		<td colspan=8 align=center><font size=+1><b></b></font></td>
	</tr>
	<tr>
		<td colspan=8 align=center><font size=+1>Danh Sach Lanh Luong Bang the ATM</font></td>
	</tr>
	<tr>
		<td colspan=8 align=center><font size=+1><b>Thang &nbsp; <%=right(yymm,2)%> &nbsp; Nam &nbsp;  <%=left(yymm,4)%></b></font></td>
	</tr> 
	
	<TR HEIGHT=25 BGCOLOR="#e4e4e4" class="txtvn9">
		<td align=center>STT</td>
		<td align=center>Bo Phan</td>
		<td align=center>So The</td>
		<td align=center>Ten</td>
		<td align=center>NVX</td>
		<td align=center>CMND</td>
		<td align=center>SO Tai Khoan </td>
		<td align=center>SO Tien </td>
	</tr> 

<%if rptType="A" then %>
	<%x = 0 
	  grp_cnt = 0	
	  grp_amt = 0 
	  grp_id = ""
	  tot_amt= 0 
	  while not rs.eof   
		if trim(rs("bankid_str"))<>"" and rs("zhuanM")>"0" then  
			x= x +1    
			tot_amt = tot_amt + cdbl(rs("zhuanM"))
			'response.write "1=" & grp_id &"<BR>"
			'response.write rs("groupid") &"<BR>"
			'response.write "cnt=" & grp_cnt &"<BR>"
			'response.write "amt=" & grp_amt &"<BR>"			

	%> 	
			<%if grp_id <>"" and  grp_id <> rs("groupid")  then %> 
				<tr  class="txtvn9">
					<td align=left BGCOLOR="yellow">&nbsp;</td>
					<td align=center BGCOLOR="yellow"><%=grp_id%></td>
					<td align=right BGCOLOR="yellow"><%=grp_cnt%></td>
					<td align=left BGCOLOR="yellow">&nbsp;</td>
					<td align=left BGCOLOR="yellow">&nbsp;</td>
					<td align=left BGCOLOR="yellow">&nbsp;</td>
					<td align=left BGCOLOR="yellow">&nbsp;</td>
					<td align=right BGCOLOR="yellow"><%=formatnumber(grp_amt,0)%></td> 
				</tr>
			<%end if %>
			<TR HEIGHT=22 BGCOLOR="#ffffff" class="txtvn9">				
				<td align=left><%=x%></td>
				<td align=left><%=rs("groupid")%></td>
				<td align=left><%=rs("empid")%></td>
				<td align=left nowrap class="txt9"><%=rs("empnam_vn")%></td>
				<td align=left><%=rs("indat_dmy")%></td>
				<td align=left style="mso-number-format:\@" ><%=rs("personid")%></td>
				<td align=left style="mso-number-format:\@"><%=rs("bankid_str")%></td>
				<td align=right><%=formatnumber(rs("zhuanM"),0)%></td>
			</tr> 
	<%		 if  grp_id<>"" and grp_id <> rs("groupid") then 
				grp_cnt = 1  
				grp_amt = 0
			else
				grp_cnt = grp_cnt + 1 
				grp_amt = grp_amt + cdbl(rs("zhuanM"))
			end if 	
			grp_id =  rs("groupid") 					
		end if 
	rs.movenext
	%> 
		<%if rs.eof   then%> 
		<tr  >
			<td align=left BGCOLOR="yellow" >&nbsp;</td>
			<td align=center BGCOLOR="yellow" ><b><%=grp_id%></b></td>
			<td align=right BGCOLOR="yellow" ><b><%=grp_cnt%></b></td>
			<td align=left BGCOLOR="yellow" >&nbsp;</td>
			<td align=left BGCOLOR="yellow" >&nbsp;</td>
			<td align=left BGCOLOR="yellow" >&nbsp;</td>
			<td align=left BGCOLOR="yellow" >&nbsp;</td>
			<td align=right BGCOLOR="yellow" ><b><%=formatnumber(grp_amt,0)%></b></td>			
		</tr>
		<%end if%>	
	<%wend%>  
<%else%>  <!--in 績效獎金-->
	<%x = 0 
	  grp_cnt = 0	
	  grp_amt = 0 
	  grp_id = ""
	  tot_amt= 0 
	  while not rs.eof   
		if trim(rs("bankid_str"))<>"" and rs("reljxm")>"0" then  
			x= x +1    
			tot_amt = tot_amt + cdbl(rs("reljxm"))
			'response.write "1=" & grp_id &"<BR>"
			'response.write rs("groupid") &"<BR>"
			'response.write "cnt=" & grp_cnt &"<BR>"
			'response.write "amt=" & grp_amt &"<BR>"			

	%> 	
			<%if grp_id <>"" and  grp_id <> rs("lg")  then %> 
				<tr>
					<td align=left BGCOLOR="yellow">&nbsp;</td>
					<td align=center BGCOLOR="yellow"><b><%=grp_id%></b></td>
					<td align=right BGCOLOR="yellow"><b><%=grp_cnt%></b></td>
					<td align=left BGCOLOR="yellow">&nbsp;</td>
					<td align=left BGCOLOR="yellow">&nbsp;</td>
					<td align=left BGCOLOR="yellow">&nbsp;</td>
					<td align=left BGCOLOR="yellow">&nbsp;</td>
					<td align=right BGCOLOR="yellow"><b><%=formatnumber(grp_amt,0)%></b></td> 
				</tr>
			<%end if %>
			<TR HEIGHT=22 BGCOLOR="#ffffff">				
				<td align=left><%=x%></td>
				<td align=left><%=rs("lg")%></td>
				<td align=left><%=rs("empid")%></td>
				<td align=left nowrap class="txt9"><%=rs("empnam_vn")%></td>
				<td align=left>&nbsp;<%=rs("indat_dmy")%></td>
				<td align=left>&nbsp;<%=rs("personid")%></td>
				<td align=left>&nbsp;<%=rs("bankid_str")%></td>
				<td align=right><%=formatnumber(rs("reljxm"),0)%></td>
			</tr> 
	<%		 if  grp_id<>"" and grp_id <> rs("lg") then 
				grp_cnt = 1  
				grp_amt = 0
			else
				grp_cnt = grp_cnt + 1 
				grp_amt = grp_amt + cdbl(rs("reljxm"))
			end if 	
			grp_id =  rs("lg") 					
		end if 
	rs.movenext
	%> 
		<%if rs.eof   then%> 
		<tr  >
			<td align=left BGCOLOR="yellow" >&nbsp;</td>
			<td align=center BGCOLOR="yellow" ><b><%=grp_id%></b></td>
			<td align=right BGCOLOR="yellow" ><b><%=grp_cnt%></b></td>
			<td align=left BGCOLOR="yellow" >&nbsp;</td>
			<td align=left BGCOLOR="yellow" >&nbsp;</td>
			<td align=left BGCOLOR="yellow" >&nbsp;</td>
			<td align=left BGCOLOR="yellow" >&nbsp;</td>
			<td align=right BGCOLOR="yellow" ><b><%=formatnumber(grp_amt,0)%></b></td>			
		</tr>
		<%end if%>	
	<%wend%>  
<%end if%>
	<tr>
		<td align=left BGCOLOR="aqua" >&nbsp;</td>
		<td align=center BGCOLOR="aqua" ><b>TOTAL</b></td>
		<td align=right BGCOLOR="aqua" ><b><%=x%></b></td>
		<td align=left BGCOLOR="aqua" >&nbsp;</td>
		<td align=left BGCOLOR="aqua" >&nbsp;</td>
		<td align=left BGCOLOR="aqua" >&nbsp;</td>
		<td align=left BGCOLOR="aqua" >&nbsp;</td>
		<td align=right BGCOLOR="aqua"><b><%=formatnumber(tot_amt,0)%></b></td>			
	</tr> 	
</table> 
<%response.end%>

</body>
</html> 
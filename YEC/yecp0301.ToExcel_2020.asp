<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%> 
<!-- #include file = "../GetSQLServerConnection.fun" --> 
<!-- #include file="../ADOINC.inc" -->
<!-- #include file="../Include/func.inc" -->
<!--#include file="../include/checkpower.asp"-->  
<%
Set conn = GetSQLServerConnection()	  
self="YECP0301"  


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
 YYMM=request("YYMM")
code01=request("YYMM")
acc=request("acc")
code02=request("country")
code03=request("whsno")
code04=request("groupid")
code05=request("job")
code06=request("empid1") 
code07=request("outemp") 
code08=request("acc")  

sql="exec RPT_empDsalaryN '"& code01 &"',  '"& code02 &"', '"& code03 &"', '"& code04 &"', '"& code05 &"', '"& code06 &"', '"& code07 &"', '"& code08 &"' "
set rs=conn.execute(Sql)

'response.write sql 
'response.end 

%>

<html> 
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="refresh">
<style >
.txtvn8 { FONT-FAMILY: VNI-Times;  FONT-SIZE: 8pt }  
.txtvn9 { FONT-FAMILY: sylfaen;VNI-Times;  FONT-SIZE: 10pt }  
</style> 
</head> 
<body  > 
<%
  ilenamestr = "TienLuong"&yymm&".xls"
  Response.AddHeader "content-disposition","attachment; filename=" & filenamestr
  Response.Charset ="BIG5"
  Response.ContentType = "Content-Language;content=zh-tw" 
  Response.ContentType = "application/vnd.ms-excel"  
	'Response.ContentType = "application/msword"
%>
<span style="font-size:12pt">
Bản lương nhân viên tháng <%=yymm%> &nbsp;
Thang &nbsp; <%=right(yymm,2)%> &nbsp; Nam &nbsp;  <%=left(yymm,4)%>
</span> 
<TABLE CLASS="txtvn8" BORDER=1 cellspacing="1" cellpadding="1" >		
	<TR HEIGHT=25 BGCOLOR="#e4e4e4">
		<td align=center>STT</td>
		<td align=center>廠別<br>Xuong</td>
		<td align=center>部門<br>Bo Phan</td>
		<td align=center>國籍<br>Quoc tich</td>
		<td align=center>工號<br>So The</td>
		<td align=center>姓名<br>Ho Ten</td>
		<td align=center>到職<br>NVX</td>
		<td align=center>工作天數<br>so ngay<br>lam viec</td> 		
		<td align=center>時薪<br>CB/gi</td>
		<td align=center>基薪<br>BB</td>
		<td align=center>補助(Y)<br>Phu cap(Y)</td>
		<td align=center>補薪<br>BL</td>
		<td align=center>職務<br>CV</td>		
		<td align=center>語言<br>NN</td>
		<td align=center>技術<br>KT</td>
		<td align=center>環境<br>MT</td>
		<td align=center>其加<br>Phu cap<br>khac</td>		
		<td align=center>全勤<br>C.Can</td>
		<td align=center>其他收入<br>Tu nhap<br>khac</td>		
		<td align=center>績效獎金<br>Tien thuong</td>
		<td align=center>上月補款<br>TBTR</td>
		<%if session("rights")="0" then %>
				<td align=center  bgcolor="yellow">WP</td>
		<%end if%>	
		<td align=center>H1<br>*1.5</td>		
		<td align=center>H2<br>*2</td>		
		<td align=center>H3<br>*3</td>		
		<td align=center>總加班費<br>Phi tang ca</td>		
		<td align=center>B3<br>*0.3</td>		
		<td align=center>夜班<br>*0.3</td>
		<td align=center>合計<br>Total</td>
		<td align=center>扣其他<br>Tru Tien kH</td>		
		<td align=center>保險<br>BH(ALL)</td>		
		<td align=center>工團<br>CD</td>
		<td align=center>事病假<br>Trừ Phép<br>(ko.du thang)</td>
		<td align=center>總扣款<br>Tong tru Tieg</td>
		<td align=center>所得稅<br>Thue</td>		
		<td align=center>總金資<br>Tong tien luong</td>
		<td align=center>金額<br>Tien Luong thoc lanh</td>
		<td align=center>零數<br>So le</td> 
		<td align=center>離職日<br>NTV</td>
		<td align=center>備註<br>Ghi chu</td>
	</tr>
	<%x = 0 
	  grp_cnt = 0	
	  grp_amt = 0 
	  grp_id = ""
	  tot_amt= 0 
	  while not rs.eof   
	  x=x+1 
		if rs("country")="VN" then xs = 0 else xs=2 
		allJBM = cdbl(rs("h1m"))+cdbl(rs("h2m"))+cdbl(rs("h3m"))+cdbl(rs("b3m"))		
		if yymm>="200910" then 
			allKM =rs("BZKM") 
		else
			allKM =cdbl(rs("kzm"))+cdbl(rs("jiaAm"))+cdbl(rs("jiaBm"))+cdbl(rs("BZKM"))
		end if  
		if trim(rs("memo"))<>"" then 
			memostr = replace(replace(rs("memo"),vbcrlf," "),"<br>"," ")
		else
			memostr=""
		end if 	 
		totKM = (cdbl(rs("qita"))+cdbl(rs("bh"))+cdbl(rs("GT"))+cdbl(allKM))*-1
		
		n_06h1m= round(cdbl(rs("jbh1"))*cdbl(rs("money_h"))*1.5 ,0)
		n_06h2m= round(cdbl(rs("jbh2"))*cdbl(rs("money_h"))*2 ,0)
		n_06h3m= round(cdbl(rs("jbh3"))*cdbl(rs("money_h"))*3 ,0)
		n_06allJBM = round( cdbl(n_06h1m) + cdbl(n_06h2m)+cdbl(n_06h3m) ,0)
		n_08_1 = cdbl(rs("b3"))-cdbl(rs("ov_b3"))
		n_08 = round( cdbl(n_08_1)*cdbl(rs("money_h"))*0.3 ,0) 
		
		dif_jbm = cdbl(allJBM) -cdbl(n_06allJBM) - cdbl(n_08)
		newjxamt  = cdbl( rs("jx"))+cdbl( dif_jbm )
		
	%> 	

		<TR HEIGHT=22 BGCOLOR="#ffffff" class="txtvn9">				
			<td align=left><%=x%></td>
			<td align=left><%=rs("whsno")%></td>
			<td align=left><%=rs("groupid")%></td>
			<td align=left><%=rs("country")%></td>
			<td align=left><%=rs("empid")%></td>
			<td align=left nowrap class="txtvn8"><%=rs("empnam_cn")&rs("empnam_vn")%></td>
			<td align=left><%=rs("e_indat")%></td>
			<td align=right><%=rs("workdays")%></td>			
			<td align=right><%=formatnumber(rs("money_h"),xs)%></td>
			<td align=right><%=formatnumber(rs("BB"),0)%></td>
			<td align=right><%=formatnumber(rs("PHU"),0)%></td>
			<td align=right><%=formatnumber(rs("btien"),0)%></td>
			<td align=right><%=formatnumber(rs("CV"),0)%></td>			
			<td align=right><%=formatnumber(rs("NN"),0)%></td>
			<td align=right><%=formatnumber(rs("KT"),0)%></td>
			<td align=right><%=formatnumber(rs("MT"),0)%></td>
			<td align=right><%=formatnumber(rs("TTKH"),0)%></td>
			<td align=right><%=formatnumber(rs("QC"),0)%></td>
			<td align=right><%=formatnumber(rs("tnkh"),0)%></td>
			<td align=right><%=formatnumber(newjxamt,0)%> 
			</td>
			<td align=right><%=formatnumber(rs("tbtr"),0)%></td>
			<%if session("rights")="0" then %>
				<td align=right bgcolor="yellow"><%=formatnumber(rs("wpamt"),0)%></td>
			<%end if%>
			<td align=right ><%=rs("jbh1")%></td>
			<td align=right ><%=rs("jbh2")%></td>
			<td align=right ><%=rs("jbh3")%></td>
			<td align=right ><%=formatnumber(n_06allJBM,0)%></td> 
			<td align=right ><%=n_08_1%></td>			
			<td align=right><%=formatnumber(n_08,0)%></td>			
			<td align=right><%=formatnumber(rs("TOTM"),0)%></td>
			<td align=right><%=formatnumber(rs("qita"),0)%></td>
			<td align=right><%=formatnumber(rs("BH"),0)%></td>
			<td align=right><%=formatnumber(rs("GT"),0)%></td>
			<td align=right><%=formatnumber(allKM,0)%></td>
			<td align=right><%=formatnumber(totKM,0)%></td>
			<td align=right><%=formatnumber(cdbl(rs("ktaxm"))*-1,0)%></td>
			<td align=right><%=formatnumber(rs("real_total"),0)%></td>
			<td align=right><%=formatnumber(rs("laonh"),0)%></td>
			<td align=right><%=formatnumber(rs("sole"),0)%></td>
			<td align=right>&nbsp;<%=rs("outdate")%></td>
			<td align=centet><%=memostr%>&nbsp;</td>
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
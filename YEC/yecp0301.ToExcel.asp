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
.txtvn8 { FONT-FAMILY: Times New Roman;sylfaen;VNI-Times;  FONT-SIZE: 8pt }  
.txtvn9 { FONT-FAMILY: Times New Roman;sylfaen;VNI-Times;  FONT-SIZE: 10pt }
.txtTotal { FONT-FAMILY: Times New Roman;sylfaen;VNI-Times;  FONT-SIZE: 12pt }
.txtvn14 { FONT-FAMILY: Times New Roman;sylfaen;VNI-Times;  FONT-SIZE: 14pt }  
</style> 
</head> 
<body  > 
<%
  ilenamestr = "TienLuong"&yymm&".xls"
  Response.AddHeader "content-disposition","attachment; filename=" & filenamestr
  Response.Charset ="BIG5"
  Response.ContentType = "Content-Language;content=zh-tw" 
  Response.ContentType = "application/vnd.ms-excel"  
	
%>
<TABLE style="height:50px" class="txtvn14" BORDER=0 cellspacing="3" cellpadding="3" >	
	<tr>
		<td colspan=4 align="center">CÔNG TY TNHH LẬP THẮNG</td>
		<td colspan=21 align="center">LƯƠNG THÁNG &nbsp;<%=right(yymm,2)%> &nbsp; NĂM &nbsp;  <%=left(yymm,4)%><br><%=right(yymm,2)%>月份員工薪資表</td>
	</tr>	
</table>

<TABLE CLASS="txtvn8" BORDER=1 cellspacing="1" cellpadding="1" >
	<TR HEIGHT=25 BGCOLOR="#e4e4e4">
		<td align=center rowspan=2>卡號<br>STT</td>		
		<td align=center rowspan=2>工號<br>Số thẻ</td>
		<td align=center rowspan=2>姓名<br>Họ Tên</td>
		<td align=center rowspan=2>到職日<br>Ngày vào xưởng</td>
		<td align=center rowspan=2>工作天數<br>Ngày công</td>
		<td align=center rowspan=2>基本薪<br>Lương cơ bản</td>	
		<td align=center rowspan=2>職務加給<br>PC chức vụ</td>
		<td align=center rowspan=2>技術<br>Kỹ thuật</td>
		<td align=center rowspan=2>環境<br>Môi trường</td>
		<td align=center rowspan=2>電話津貼<br>PC điện thoại</td>
		<td align=center rowspan=2>燃油津貼<br>PC xăng xe</td>
		<td align=center rowspan=2>時薪<br>Lương/giờ</td>
		<td align=center colspan=4>加班-Tăng ca</td>	
		<td align=center colspan=3>津貼-Phụ cấp</td>
		<td align=center rowspan=2>全勤<br>Chuyên cần</td>				
		<td align=center rowspan=2>住房支持<br>Hỗ trợ nhà ở</td>
		<td align=center rowspan=2>其他收入<br>Thu nhập khác</td>		
		<td align=center rowspan=2>績效獎金<br>Thưởng hiệu suất</td>
		<td align=center rowspan=2>合計<br>Tổng cộng</td>
		<td align=center rowspan=2>(-)扣時假<br>(-)Phép</td>
		<td align=center rowspan=2>(-)其他<br>(-)Khác</td>		
		<td align=center rowspan=2>(-)保險費<br>(-) Bảo hiểm</td>
		<td align=center rowspan=2>(-)工團費<br>(-) Công đoàn</td>
		<td align=center rowspan=2>(-)個人所得稅<br>(-) Thuế TNCN</td>		
		<td align=center rowspan=2>實領<br>Thực lãnh</td>	
	</tr>
	<tr BGCOLOR="#e4e4e4">
		<td align=center>加班<br>NT*1.5</td>		
		<td align=center>星期日<br>CN*2</td>		
		<td align=center>放假<br>NL*3</td>		
		<td align=center>加班金額<br>Thành tiền</td>
		<td align=center>0.5</td>
		<td align=center>0.3</td>
		<td align=center>Thành tiền<br></td>
	</tr>
	<%
		x = 0 
		grp_cnt = 0	
		grp_amt = 0 
		grp_id = ""
		tot_amt= 0 
		sum_bsCB=0
		sum_BB=0
		sum_n_06allJBM=0
		sum_B3=0
		sum_B3M=0
		sum_B4=0
		sum_B4M=0
		sum_B5=0
		sum_B5M=0
		sum_B4_5M=0
		sum_QC=0
		sum_CV=0
		sum_KT=0
		sum_PHU=0
		sum_NN=0
		sum_MT=0
		sum_TTKH=0
		sum_TNKH=0
		sum_JX=0
		sum_TOTM=0
		sum_BZKM=0
		sum_QITA=0
		sum_BH=0
		sum_GT=0
		sum_KTAXM=0
		sum_REAL_TOTAL=0
	  
	  while not rs.eof   
	  x=x+1 
		
		if rs("country")="VN" then 
			n_H1=cdbl(rs("jbh1"))
			n_H2=cdbl(rs("jbh2"))
			n_H3=cdbl(rs("jbh3"))
			
			n_06h1m= round(n_H1*cdbl(rs("money_h"))*1.5 ,0)
			n_06h2m= round(n_H2*cdbl(rs("money_h"))*2 ,0)
			n_06h3m= round(n_H3*cdbl(rs("money_h"))*3 ,0)			
		else
			n_H1=cdbl(rs("H1"))
			n_H2=cdbl(rs("H2"))
			n_H3=cdbl(rs("H3"))
			
			n_06h1m= round(cdbl(rs("H1M")) ,0)
			n_06h2m= round(cdbl(rs("H2M")) ,0)
			n_06h3m= round(cdbl(rs("H3M")) ,0)
		end if
		n_06allJBM = round( cdbl(n_06h1m) + cdbl(n_06h2m)+cdbl(n_06h3m) ,0)
		'=====================================
		'@0New_h1={RPT_empDsalary;1.jbH1}				
		Onew_h1M=round(cdbl(rs("jbh1")) *cdbl(rs("MONEY_H")) *1.5   , 0)
			'round(  {@0New_h1} * {RPT_empDsalary;1.MONEY_H} *1.5   , 0)  
		'-------------------------------------------------		
		ov_h1M=round(cdbl(rs("ov_h1"))*cdbl(rs("MONEY_H"))*1.5,0)+(cdbl(rs("H1M"))-(Onew_h1M+round(cdbl(rs("ov_h1"))*(cdbl(rs("MONEY_H"))*1.5),0)))
				'round(  {RPT_empDsalary;1.ov_h1} *   ( {RPT_empDsalary;1.MONEY_H} * 1.5 ), 0 ) 
				'+ ( {RPT_empDsalary;1.H1M} - ( {@0new_h1M}+ round(  {RPT_empDsalary;1.ov_h1} *   ( {RPT_empDsalary;1.MONEY_H} * 1.5 ), 0 ) ) )
		'---------------------------------------------
		'@0New_h2={RPT_empDsalary;1.jbH2}
		Onew_h2M=round(cdbl(rs("jbh2")) *cdbl(rs("MONEY_H")) *2   , 0)
				'round (  {@0New_h2} *( {RPT_empDsalary;1.MONEY_H} *2  )  , 0) 
		'-------------------------------------------------------
		ov_h2M=round(cdbl(rs("ov_h2"))*cdbl(rs("MONEY_H"))*2 , 0 )+(cdbl(rs("H2M"))-(Onew_h2M+round(cdbl(rs("ov_h2"))*(cdbl(rs("MONEY_H"))*2),0)))
				'round( {RPT_empDsalary;1.ov_h2}*  ( {RPT_empDsalary;1.MONEY_H} * 2  ) , 0 ) 
				'+ ( {RPT_empDsalary;1.H2M} - ( {@0new_h2M}+ round( {RPT_empDsalary;1.ov_h2}*  ( {RPT_empDsalary;1.MONEY_H} * 2  ) , 0 ) ) )		
		'---------------------------------------------
		'@0New_h3={RPT_empDsalary;1.jbH3}
		Onew_h3M=round(cdbl(rs("jbh3")) *cdbl(rs("MONEY_H")) *3   , 0)
				'round ( {@0New_h3} *  ( {RPT_empDsalary;1.MONEY_H} *3  )  , 0)
		'-------------------------------------------------------
		ov_h3M=round(cdbl(rs("ov_h3"))*cdbl(rs("MONEY_H"))*3 , 0 )+(cdbl(rs("H3M"))-(Onew_h3M+round(cdbl(rs("ov_h3"))*(cdbl(rs("MONEY_H"))*3),0)))
				'round( {RPT_empDsalary;1.ov_h3}* ( {RPT_empDsalary;1.MONEY_H} * 3  ), 0 )  
				'+ ( {RPT_empDsalary;1.H3M} - ( {@0new_h3M}+ round( {RPT_empDsalary;1.ov_h3}*( {RPT_empDsalary;1.MONEY_H} * 3  ), 0 )   ) )
		'-------------------------------------------------------		
		'@0New_b3={RPT_empDsalary;1.b3}-{RPT_empDsalary;1.ov_b3}
		Onew_b3=cdbl(rs("B3"))-cdbl(rs("ov_b3"))
		Onew_b3M=round(Onew_b3*cdbl(rs("MONEY_H"))*2.1  , 0)
				
		'------------------------------------------------		
		ov_B3M=cdbl(rs("B3M"))-Onew_b3M
				'{RPT_empDsalary;1.B3M}+{RPT_empDsalary;1.ovB3_2M}-{@0new_b3M}
		
		'-------------------------------------------------------
		oNewckhJBM=(cdbl(rs("H1M"))+cdbl(rs("H2M"))+cdbl(rs("H3M")))-(ov_h1M+ov_h2M+ov_h3M+Onew_h1M+Onew_h2M+Onew_h3M)
		
				'-{RPT_empDsalary;1.H1M}+{RPT_empDsalary;1.H2M}+{RPT_empDsalary;1.H3M} - 
				'( {@ov_h1M}+{@ov_h2M}+{@ov_h3M}+{@0new_h1M}+{@0new_h2m}+{@0new_h3M}) 
		'---------------------------------------
		
		'if {@oNewckhJBM}  > 0 then 
		'	{RPT_empDsalary;1.JX} +  {@ov_h1M} + {@ov_h2M} + {@ov_h3M} +{@ov_B3M}  +  {@oNewckhJBM} 
		'else
		'	{RPT_empDsalary;1.JX} +  {@ov_h1M} + {@ov_h2M} + {@ov_h3M} + {@ov_B3M}
		B4_5M=cdbl(rs("B4M"))+cdbl(rs("B5M"))
		
		if rs("country")="VN" then 
			newjxamt=cdbl(rs("JX")) +  ov_h1M + ov_h2M + ov_h3M + ov_B3M
			if oNewckhJBM  > 0 then newjxamt=newjxamt +  oNewckhJBM
		else
			newjxamt=cdbl(rs("JX"))
		end if	 
		
		
		'=====================================
		
				
		sum_BB=cdbl(sum_BB)+cdbl(rs("BB"))
		sum_n_06allJBM=cdbl(sum_n_06allJBM)+cdbl(n_06allJBM)
		sum_QC=cdbl(sum_QC)+cdbl(rs("QC"))
		sum_CV=cdbl(sum_CV)+cdbl(rs("CV"))
		sum_KT=cdbl(sum_KT)+cdbl(rs("KT"))
		sum_PHU=cdbl(sum_PHU)+cdbl(rs("PHU"))
		sum_NN=cdbl(sum_NN)+cdbl(rs("NN"))
		sum_MT=cdbl(sum_MT)+cdbl(rs("MT"))
		sum_TTKH=cdbl(sum_TTKH)+cdbl(rs("TTKH"))
		sum_TNKH=cdbl(sum_TNKH)+cdbl(rs("tnkh"))
		sum_JX=cdbl(sum_JX)+cdbl(rs("jx"))
		sum_TOTM=cdbl(sum_TOTM)+cdbl(rs("TOTM"))
		sum_BZKM=cdbl(sum_BZKM)+cdbl(rs("BZKM"))
		sum_QITA=cdbl(sum_QITA)+cdbl(rs("qita"))
		sum_BH=cdbl(sum_BH)+cdbl(rs("BH"))
		sum_GT=cdbl(sum_GT)+cdbl(rs("GT"))
		sum_KTAXM=cdbl(sum_KTAXM)+cdbl(rs("ktaxm"))
		sum_B4=cdbl(sum_B4)+cdbl(rs("B4"))
		sum_B4M=cdbl(sum_B4M)+cdbl(rs("B4M"))
		sum_B5=cdbl(sum_B5)+cdbl(rs("B5"))
		sum_B5M=cdbl(sum_B5M)+cdbl(rs("B5M"))
		sum_B4_5M=cdbl(sum_B4_5M)+cdbl(sum_B4M)+cdbl(sum_B5M)
		sum_REAL_TOTAL=cdbl(sum_REAL_TOTAL)+cdbl(rs("real_total"))
	%> 	

	<TR HEIGHT=22 BGCOLOR="#ffffff" class="txtvn9">				
		<td align=left><%=x%></td>			
		<td align=left><%=rs("empid")%></td>
		<td align=left nowrap class="txtvn8"><%=rs("empnam_cn")&rs("empnam_vn")%></td>
		<td align=left><%=rs("e_indat")%></td>
		<td align=right><%=rs("workdays")%></td>
		<td align=right><%=formatnumber(rs("BB"),0)%></td>
		<td align=right><% if formatnumber(rs("CV"),0)=0 then %> <% else %><%=formatnumber(rs("CV"),0)%><% end if%></td>
		<td align=right><% if formatnumber(rs("KT"),0)=0 then %> <% else %><%=formatnumber(rs("KT"),0)%><% end if%></td>
		<td align=right><% if formatnumber(rs("MT"),0)=0 then %> <% else %><%=formatnumber(rs("MT"),0)%><% end if%></td>
		<td align=right><% if formatnumber(rs("PHU"),0)=0 then %> <% else %><%=formatnumber(rs("PHU"),0)%><% end if%></td>	
		<td align=right><% if formatnumber(rs("NN"),0)=0 then %> <% else %><%=formatnumber(rs("NN"),0)%><% end if%></td>
		<td align=right><% if formatnumber(rs("MONEY_H"),0)=0 then %> <% else %><%=formatnumber(rs("MONEY_H"),0)%><% end if%></td>
		<td align=right><% if n_H1 > 0 then %><%=n_H1%><% else %> <% end if%></td>
		<td align=right><% if n_H2 > 0 then %><%=n_H2%><% else %> <% end if%></td>
		<td align=right><% if n_H3 > 0 then %><%=n_H3%><% else %> <% end if%></td>
		<td align=right><% if n_06allJBM=0 then %>  <% else %><%=formatnumber(n_06allJBM,0)%><% end if%></td>
		<td align=right><% if formatnumber(rs("b4"),0)=0 then %> <% else %><%=formatnumber(rs("b4"),2)%><% end if%></td>
		<td align=right><% if formatnumber(rs("b5"),0)=0 then %> <% else %><%=formatnumber(rs("b5"),2)%><% end if%></td>
		<td align=right><% if formatnumber(B4_5M,0)=0 then %> <% else %><%=formatnumber(B4_5M,0)%><% end if%></td>		
		<td align=right><% if formatnumber(rs("QC"),0)=0 then %> <% else %><%=formatnumber(rs("QC"),0)%><% end if%></td>	
		<td align=right><% if formatnumber(rs("TTKH"),0)=0 then %> <% else %><%=formatnumber(rs("TTKH"),0)%><% end if%></td>
		<td align=right><% if formatnumber(rs("tnkh"),0)=0 then %> <% else %><%=formatnumber(rs("tnkh"),0)%><% end if%></td>
		<td align=right><% if formatnumber(newjxamt,0)=0 then %> <% else %><%=formatnumber(newjxamt,0)%><% end if%></td>
		<td align=right><% if formatnumber(rs("TOTM"),0)=0 then %> <% else %><%=formatnumber(rs("TOTM"),0)%><% end if%></td>
		<td align=right><% if formatnumber(rs("BZKM"),0)=0 then %> <% else %><%=formatnumber(rs("BZKM"),0)%><% end if%></td>				
		<td align=right><% if formatnumber(rs("qita"),0)=0 then %> <% else %><%=formatnumber(rs("qita"),0)%><% end if%></td>			
		<td align=right><% if formatnumber(rs("BH"),0)=0 then %> <% else %><%=formatnumber(rs("BH"),0)%><% end if%></td>
		<td align=right><% if formatnumber(rs("GT"),0)=0 then %> <% else %><%=formatnumber(rs("GT"),0)%><% end if%></td>
		<td align=right><% if formatnumber(cdbl(rs("ktaxm")),0)=0 then %> <% else %><%=formatnumber(cdbl(rs("ktaxm")),0)%><% end if%></td>			
		<td align=right><%=formatnumber(rs("real_total"),0)%></td>
	</tr> 
	<%	 
	rs.movenext
	%> 
	<%wend%>  
		<tr class="txtTotal">
			<td colspan=5 align="right">Tổng</td>
			<td><%=formatnumber(sum_BB,0) %></td>						
			<td><%=formatnumber(sum_CV,0) %></td>
			<td><%=formatnumber(sum_KT,0) %></td>
			<td><%=formatnumber(sum_MT,0) %></td>
			<td><%=formatnumber(sum_PHU,0) %></td>
			<td><%=formatnumber(sum_NN,0) %></td>
			<td></td>
			<td></td>
			<td></td>
			<td></td>
			<td><%=formatnumber(sum_n_06allJBM,0) %></td>
			<td><%=formatnumber(sum_B4,2) %></td>
			<td><%=formatnumber(sum_B5,2) %></td>
			<td><%=formatnumber(sum_B4_5M,0) %></td>			
			<td><%=formatnumber(sum_QC,0) %></td>
			<td><%=formatnumber(sum_TTKH,0) %></td>
			<td><%=formatnumber(sum_TNKH,0) %></td>
			<td><%=formatnumber(sum_JX,0) %></td>
			<td><%=formatnumber(sum_TOTM,0) %></td>
			<td><%=formatnumber(sum_BZKM,0) %></td>
			<td><%=formatnumber(sum_QITA,0) %></td>
			<td><%=formatnumber(sum_BH,0) %></td>
			<td><%=formatnumber(sum_GT,0) %></td>
			<td><%=formatnumber(sum_KTAXM,0) %></td>
			<td><%=formatnumber(sum_REAL_TOTAL,0) %></td>
		</tr>
</table> 
<TABLE style="height:100px" class="txtTotal" BORDER=0 cellspacing="3" cellpadding="3" >	
	<TR><td colspan=24 style="height:20px"></td></TR>
	<tr>
		<TD colspan=2></TD>
		<TD colspan=15>Người lập biểu 製表員</TD>
		<TD colspan=7 align="center">Người duyệt biểu 經理 </TD>
	</tr>	
</table>
<%
rs.close
set rs=nothing 
conn.close
set conn=nothing
response.end%>
		
</body>
</html> 
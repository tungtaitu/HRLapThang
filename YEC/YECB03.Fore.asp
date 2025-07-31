<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%> 
<!-- #include file = "../GetSQLServerConnection.fun" --> 
<!-- #include file="../ADOINC.inc" -->
<!-- #include file="../Include/func.inc" -->
<!--#include file="../include/checkpower.asp"-->  
<!--#include file="../include/sideinfo.inc"-->
<%
Set conn = GetSQLServerConnection()	  
self="YEcb03"  
 
nowmonth = year(date())&right("00"&month(date()),2)  
if month(date())="01" then  
	calcmonth = year(date())-1&"12"  	
else
	calcmonth =  year(date())&right("00"&month(date())-1,2)   
end if 	   

if day(date())<=11 then 
	if month(date())="01" then  
		calcmonth = year(date())-1&"12" 
	else
		calcmonth =  year(date())&right("00"&month(date())-1,2)   
	end if 	 	
else
	calcmonth = nowmonth 
end if 

if session("netuser")="" then 
	response.write "UserID is Empty Please Login again !!!<BR>"
	response.write "Vao mang trong rong , hoac doi lau , hay nhan nut nhap mang tu dau !!! "
	response.end 
end if 

CloseYM = request("CloseYM")
WHSNO = request("WHSNO") 
if CloseYM="" then 
	CloseYM=calcmonth
end if 	
calcdt = left(CloseYM,4)&"/"& right(CloseYM,2)&"/01"   

if right(CloseYM,2)="01" then 
	lastMonth=left(closeYM,4)-1&"12"
else
	lastMonth=CloseYM-1 
end if 	
predat = left(lastMonth,4)&"/"& right(lastMonth,2)&"/01"   
%>

<html> 
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="refresh">
<link rel="stylesheet" href="../Include/style.css" type="text/css">
<link rel="stylesheet" href="../Include/style2.css" type="text/css">
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
 
function f()
	<%=self%>.closeym.focus()	
	<%=self%>.closeym.SELECT()
end function    
 
</SCRIPT>   
</head> 
<body  topmargin="5" leftmargin="5"  marginwidth="0" marginheight="0"  onkeydown="enterto()" onload='f()' >
<form name="<%=self%>" method="post" >

	<table border=0 style="height:70px"><tr><td>&nbsp;</td></tr></table>

	<table width="100%" BORDER="0" align="center" cellpadding="0" cellspacing="0" >
		<tr>
			<td>
				<table class="txt" cellpadding=3 cellspacing=3> 
					<tr height=30 >
						<TD align=right >關帳年月 Ngày Đóng Sổ  </TD>
						<TD><INPUT type="text" style="width:120px" NAME=CloseYM VALUE="<%=CloseYM%>" maxlength=6></TD>			
						<TD align=right >廠別 Xưởng</TD>
						<TD>
							<select name=WHSNO style="width:120px">					
								<%
								if session("rights")=0 then 
									SQL="SELECT * FROM BASICCODE WHERE FUNC='WHSNO' ORDER BY SYS_TYPE "%>
									<option value="" selected >全部 Toàn Bộ</option>
								<%		
								else
									SQL="SELECT * FROM BASICCODE WHERE FUNC='WHSNO' and sys_type='"& session("NETWHSNO") &"' ORDER BY SYS_TYPE "
								end if	
								SET RST = CONN.EXECUTE(SQL)
								WHILE NOT RST.EOF
								%>
								<option value="<%=RST("SYS_TYPE")%>" <%if rst("sys_type")=WHSNO then%>selected<%end if%>><%=RST("SYS_VALUE")%></option>
								<%
								RST.MOVENEXT
								WEND
								%>
							</SELECT>
							<%SET RST=NOTHING %>	 
						</TD>	
						<td align=center>
							<input type=button  name=btm class="btn btn-sm btn-outline-secondary" value="(Y)確   認 XÁC NHẬN" onclick="go()" onkeydown="go()">
						</td>
					</TR>	 
				</table>
			</td>
		</tr>
		<tr>
			<td align="center">
				<table  width="98%" class="txt"   cellspacing="3" cellpadding="3">
					<Tr height=25>
						<%sqla="select isnull(b.lw,'') whsno, a.country, count(a.empid) as Bcnt from "&_
							   "(Select * from empfile where isnull(status,'')<>'D' ) a  "&_
							   "left join (select * from view_empgroup where yymm='"& closeYM &"' ) b on b.empid = a.empid "&_
							   "where convert(char(6), indat,112)<='"& closeYM &"' and (isnull(outdat,'')='' or isnull(outdat,'')<>'' and convert(char(10),outdat,111)>'"&calcdt&"' ) "&_
							   "and isnull(b.lw,'') = '"& whsno &"' group by isnull(b.lw,''), a.country "&_
							   "order by whsno, country "
								'response.write sqla
							Set rsa = Server.CreateObject("ADODB.Recordset")
							rsa.open sqla, conn, 1, 3
						%>
						<Td>0.本月應計薪人數 Số Người Tính Lương Trong Tháng<br>
							<table id="myTableGrid" width="100%">
								<tr height="35px" class="header">									
									<td align=center>廠別 Xưởng</td>
									<td align=center>國籍 Quốc Tịch</td>
									<td align=center>應計算人數 Số Người Tính Lương</td>
									<td align=center>系統人數 Số Người Hệ Thống</td>
									<td align=center>訊息說明(工號) Ghi Chú (Mã NV)</td>
								</tr>
								<%while not rsa.eof%>
									<Tr class=txt8>										
										<Td align=center><%=rsa("whsno")%></td>
										<Td align=center><%=rsa("country")%></td>
										<Td align=center><%=rsa("Bcnt")%></td>
										<%sqla2="select count(*) as Rcnt from empdsalary where real_total<>0 and "&_
												 "yymm='"& closeYM &"' and country='"& rsa("country") &"' and whsno='"& rsa("whsno") &"' "
										Set rsa2= Server.CreateObject("ADODB.Recordset")
										rsa2.open sqla2, conn, 1, 3
										%>
										<Td align=center><%=rsa2("rcnt")%></td>							
										<%err_emp=""
										if cdbl(rsa("bcnt"))<>cdbl(rsa2("rcnt")) then 
											if cdbl(rsa("bcnt"))> cdbl(rsa2("rcnt")) then 
												sql3="select isnull(b.lw,'') lw, a.* from "&_
													 "(Select * from empfile  where isnull(status,'')<>'D' ) a  "&_
													 "left join (select * from view_empgroup where yymm='"& closeYM &"' ) b on b.empid = a.empid "&_
													 "where convert(char(6), indat,112)<='"& closeYM &"' and (isnull(outdat,'')='' or isnull(outdat,'')<>'' and convert(char(10),outdat,111)>'"&calcdt&"' ) "&_
													 "and isnull(b.lw,'') ='"&rsa("whsno")&"' and a.country='"&rsa("country")&"'    "&_
													 "and a.empid not in ( select empid  from empdsalary where real_total<>0 and yymm='"& closeYM &"' and country='"& rsa("country") &"' and whsno='"& rsa("whsno") &"' )  "
												'response.write sql3 &"<BR>"
												set rs3=Server.CreateObject("ADODB.Recordset")
												rs3.open sql3, conn, 1, 3
												while not rs3.eof 
													err_emp = err_emp & rs3("empid") &"," 
												rs3.movenext
												wend 
												rs3.close
												set rs3=nothing
											else	
												sql3="select *  from empdsalary where real_total<>0 and yymm='"& closeYM &"' and "&_
													 "country='"& rsa("country") &"' and whsno='"& rsa("whsno") &"' "&_
													 "and empid not in ( "&_ 									
													 "select  a.empid from "&_
													 "(Select * from empfile  where isnull(status,'')<>'D' ) a  "&_
													 "left join (select * from view_empgroup where yymm='"& closeYM &"' ) b on b.empid = a.empid "&_
													 "where convert(char(6), indat,112)<='"& closeYM &"' and (isnull(outdat,'')='' or isnull(outdat,'')<>'' and convert(char(10),outdat,111)>'"&calcdt&"' ) "&_
													 "and isnull(b.lw,'') ='"&rsa("whsno")&"' and a.country='"&rsa("country")&"' )  "										 
												'response.write sql3 &"<BR>"
												set rs3=Server.CreateObject("ADODB.Recordset")
												rs3.open sql3, conn, 1, 3
												while not rs3.eof 
													err_emp = err_emp & rs3("empid") &"," 
												rs3.movenext
												wend 
												rs3.close
												set rs3=nothing								
											end if 
										end if 		
										%>
										<td>
											<%=left(err_emp,50)%>
										</td>
										<%rsa2.close
										set rsa2=nothing %>
									</tr>
									<tr>										
										<Td colspan=5><hr size=0></td>
									</tr>
								<%rsa.movenext
								wend
								%>
							</table>
						</td>
					</tr>
					<tr height=25>	
						<td>
						<%
						sqla="select isnull(d.bonus,0) basicCV, isnull(e.bonus,0) basicQC, isnull(f.cv,0) NowCV, "&_
							 "isnull(f.qc,0) nowQC , c.lw,  b.lj, b.ljstr, a.* from  "&_
							 "(select * from view_empfile where country<>'TW' and isnull(status,'')<>'D' ) a  "&_
							 "left join (select * from  view_empjob where  yymm='"&closeYM&"' ) b on b.empid = a.empid  "&_
							 "left join (select* from view_empgroup  where  yymm='"&closeYM&"'  ) c on c.empid =a.empid  "&_
							 "left join (select * from empsalarybasic  where func='BB'  ) d on d.job = b.lj and d.country = a.country and d.bwhsno = case when a.country='VN' then c.lw else 'LA' end  "&_
							 "left join (select * from empsalarybasic  where func='CC'  ) E on e.job = b.lj and e.country = a.country  and e.bwhsno = case when a.country='VN' then c.lw else 'LA' end  "&_
							 "left join (select * from bemps where  yymm='"&closeYM&"' ) f on f.empid = a.empid  "&_
							 "where isnull(c.lw,'')='"&whsno&"' and ( isnull(outdat,'')=''  or isnull(a.outdat,'')<>'' and convert(Char(10),a.outdat,111)>'"&calcdt&"' )  "&_
							 "and ( isnull(f.cv,0)<>isnull(d.bonus,0) or isnull(f.qc,0)<>isnull(e.bonus,0) )  "&_
							 "order by a.country desc, a.empid " 
							'response.write sqla	 
							Set rs1 = Server.CreateObject("ADODB.Recordset")
							rs1.open sqla, conn, 1, 3
							%>
						1.職務與(全勤獎金或職務加給)不符合,共 <font color=red><%=rs1.recordcount%></font> 筆
						Chức Vụ Và (Tiền Chức Vụ Hoặc Tiền Chuyên Cần) Bị Chênh Lệch ,Tổng <font color=red><%=rs1.recordcount%></font> Đơn
						<%if rs1.recordcount>0 then%><img src="../picture/icon_jia.gif"  align="absmiddle" border=0 onclick="showdata(1)" style='cursor:hand'><%end if%>				
						<%if not rs1.eof then%>
						<BR>
							<div id=div1 style="Z-index:1;  display:none" > 
							<table id="myTableGrid" width="100%">
								<Tr height=35 class="header">
									<Td align=center>STT</td>
									<Td align=center>國籍<br>Quốc Tịch</td>
									<Td align=center>工號<br>Mã NV</td>						
									<Td align=center>職務<br>Chức Vụ</td>
									<Td align=center width=70>全勤<br>Chuyển Cần</td>
									<Td align=center width=70>職務<br>Chức Vụ</td>
									<Td align=center width=70>全勤<br>Chuyên Cần</td>
									<Td align=center width=70>職務<br>Chức Vụ</td>
								</tr>
								<%x2=0
									while not rs1.eof
									x2=x2+1
								%>
									<Tr>
										<Td bgcolor=#ffff99 align=center><%=x2%></td>
										<Td bgcolor=#ffff99 align=center><%=rs1("country")%></td>
										<Td bgcolor=#ffff99 align=center><%=rs1("empid")%></td>							
										<Td bgcolor=#ffff99><%=left(rs1("ljstr"),10)%></td>
										<Td align=right bgcolor=#e4e4e4><%=rs1("basicQC")%></td>
										<Td align=right bgcolor=#e4e4e4><%=rs1("basicCV")%></td>
										<Td align=right bgcolor=#ffccff><%=rs1("nowQC")%></td>
										<Td align=right bgcolor=#ffccff><%=rs1("NowCV")%></td>
									</Tr>
								<%rs1.movenext
								wend
								%>
							</table>
							<%end if
							rs1.close 
							set rs1=nothing			
							%>	
							</div>	  			
						</td>
					</tr>
					<tr height=25>
						<td>		
						<%
						 sql="select d.bb B_BB,c.outdate, c.nindat, c.bhdat, c.empid eid, c.country, isnull(b.bh,0)bh, isnull(a.bhtot,0) n_bhtot, a.* from  "&_
							 "(select *  from empdsalary  where yymm='"& closeYM &"' and whsno='"& whsno &"'  ) b  "&_
							 "left join (  select  * from empbhgt where  yymm='"& closeYM &"' ) a on b.empid  = a.empid   "&_
							 "join (select *  from view_empfile  ) c on c.empid = b.empid  "&_
							 "left join (select *from bemps where yymm='"& closeYM &"'  ) d on d.empid = c.empid "&_ 
							 "where d.country='VN' and (  isnull(b.bh,0) <> isnull(a.bhtot,0) or isnull(d.bb,0)*0.06 <> isnull(b.bh,0) ) "&_
							 "and ( isnull(c.bhdat,'')<>'' and left(replace(c.bhdat,'/',''),6)<='"&closeYM&"' ) "&_
							 "order by c.empid "
						 'response.write sql&"<BR>"	  
						 Set rs = Server.CreateObject("ADODB.Recordset")
						 rs.open sql, conn, 1, 3
						%>
						2.保險金額與基本薪資不符合,共 <font color=red><%=rs.recordcount%></font> 筆
						Tiền Bảo Hiểm Và Lương Cơ Bản Bị Chênh Lệch , Tổng <font color=red><%=rs.recordcount%></font> Đơn
						<%if rs.recordcount>0 then%><img src="../picture/icon_jia.gif"  align="absmiddle" border=0 onclick="showdata(2)" style='cursor:hand'><%end if%>
						 <BR>
							<div id=div2 style="Z-index:1;  display:none" > 
							<%if not rs.eof  then%> 
								<table  id="myTableGrid" width="100%">
									<Tr height="35px" class="header">										
										<Td align=center>STT</td>
										<Td align=center>工號<BR>Mã NV</td>
										<Td align=center>到職日<BR>Ngày Vào</td>
										<Td align=center>保險日<BR>Ngày Bảo Hiểm</td>
										<Td align=center>離職日<BR>Ngày Thôi Việc</td>
										<Td align=center>基本薪<BR>Lương Cơ Bản</td>
										<Td align=center>理論保險<BR>Bảo Hiểm Dự Kiến</td>
										<Td align=center>保險金額<BR>Tiền Bảo Hiểm</td>
										<Td align=center>系統保險金<BR>Tiền Bảo Hiểm Hệ Thống</td>
									</tr>
									<%x=0
									while not rs.eof 
										x=x+1
									%>
									<Tr>										
										<Td><%=x%></td>
										<Td><%=rs("eid")%></td>
										<Td><%=rs("nindat")%></td>
										<Td><%=rs("bhdat")%></td>
										<Td><%=rs("outdate")%></td>
										<Td align=right><%=formatnumber(rs("B_BB"),0)%></td>
										<Td align=right><%=formatnumber(cdbl(rs("B_bb"))*0.06,0)%></td>
										<Td align=right> 
											<%=formatnumber(rs("n_bhtot"),0)%> 								 
										</td>
										<Td align=right><%=formatnumber(rs("bh"),0)%></td>
									</tr>
									<tr>										
										<Td colspan=9><hr size=0></td>
									</tr>							
									<%rs.movenext
									wend
									%>
								</table>					
							<%end if%>
							<%rs.close
							set rs=nothing
							%>				
							</div>
						</td>	
					</tr>
					<tr height=25>
						<td>3.基本薪資與(年資或有職照員工)不符合 Lương Cơ Bản Và (Nhân Viên Có Thâm Niên) Bị Chênh Lệch  </td>
					</tr>	
					<tr height=25>
						<td>	
						<%SQL="select  isnull(b.qc,0)  as B_qc, a.* from  "&_
							  "( SELECT  * FROM empdsalary where  yymm='"&closeYM&"' and country='vn'  and qc=0 and whsno='"& whsno &"') a "&_
							  "left join ( select * from bemps where  yymm='"&closeYM&"' ) b on b.empid  = a.empid "&_
							  "left join ( select * from view_empworkt_grp where yymm='"& closeYM &"' ) c on c.empid = a.empid  "&_
							  "where ( c.fl<=3 and c.ja+c.jb+c.kzhour<=8 and left(replace(outdat,'/',''),6) <>'"&closeYM&"' ) "&_
							  "and convert(char(10),indat,111)<='"& calcdt &"' " 				  
						 'Set rs = Server.CreateObject("ADODB.Recordset")
						 'rs.open sql, conn, 1, 3
						 'response.write sql
						%>
						4.全勤獎金(無曠職無請假)不符合,共 <font color=red></font> 筆 Tiền Chuyên Cần (Không Vắng Mặt Hoặc Nghỉ Phép ) Bị Chênh Lệch</td>
					</tr>
					<tr>	
						<td>
						<%SQL="select * from empdsalary where yymm='"&closeYM &"' and  real_total<=0 and whsno='"& whsno &"'"
						 Set rs = Server.CreateObject("ADODB.Recordset")
						 rs.open sql, conn, 1, 3
						 'response.write sql
						%>
						5.薪資計算<=0,共 <font color=red><%=rs.recordcount%></font> 筆 Tính Tiền Lương <=0,Tổng <font color=red><%=rs.recordcount%></font> Đơn 
						<%if rs.recordcount>0 then%><img src="../picture/icon_jia.gif"  align="absmiddle" border=0 onclick="showdata(5)" style='cursor:hand'><%end if%>
						<br>
						<%if  not rs.eof then%>
							<div id=div5 style="Z-index:1;  display:none" > 
								<table  id="myTableGrid">
									<Tr height=35 class="header">										
										<td align=center>廠別Xưởng</td>
										<td align=center>工號Mã NV</td>
										<td align=center>國籍Quốc Tịch</td> 
									</tr>			
									<%while not rs.eof %>
									<Tr>										
										<Td><%=rs("whsno")%></td>
										<Td><%=rs("empid")%></td>
										<Td><%=rs("country")%></td>
									</tr>
									<tr>
										<td colspan=3><hr size=0></td>
									</tr>
									<%rs.movenext
									wend					
									rs.close
									set rs=nothing
									%> 
								</table> 
							</div>
						<%end if%>
						</td>	
					</tr>
					<tr height=25>	
						<%		
						sql="select ktaxm, isnull(b.exrt,1) exrt, a.* from "&_
							"( select  * from empdsalary  where yymm='"& closeYM &"'     ) a  "&_
							"left join (select * from vyfyexrt ) b on b.code=a.dm and b.yyyymm = a.yymm "&_
							"where case when a.dm='VND' then real_total else real_total*exrt end > case when country='VN' then 5000000 else 8000000 end  "&_
							"and ktaxm = 0 " 
						Set rs = Server.CreateObject("ADODB.Recordset")
						rs.open sql, conn, 1, 3
						'response.write sql	
						%>
						<td>6.應扣稅,無稅金資料,共 <font color=red><%=rs.recordcount%></font> 筆 
						Dữ Liệu Tiền Trừ Thuế , Miễn Thuế , Tổng  <font color=red><%=rs.recordcount%></font> Đơn
						<%if rs.recordcount>0 then%><img src="../picture/icon_jia.gif"  align="absmiddle" border=0 onclick="showdata(6)" style='cursor:hand'><%end if%>
						<br>
						<%if  not rs.eof then%>
							<div id=div6 style="Z-index:1;  display:none" > 
								<table  id="myTableGrid">
									<Tr height=35 class="header">										
										<td align=center>廠別<br>Xưởng</td>
										<td align=center>工號<br>Mã NV</td>
										<td align=center>國籍<br>Quốc Tịch</td> 
									</tr>			
									<%while not rs.eof %>
									<Tr>										
										<Td><%=rs("whsno")%></td>
										<Td><%=rs("empid")%></td>
										<Td><%=rs("country")%></td>
									</tr>
									<tr>
										<td colspan=3><hr size=0></td>
									</tr>
									<%rs.movenext
									wend					
									rs.close
									set rs=nothing
									%> 
								</table> 
							</div>
						<%end if%>			
						</td>
					</tr>
					<tr height=25>	
						<%		
						sql="select b.lw, b.lg, b.lz, b.ls, b.lgstr, b.lzstr, isnull(c.empid,'') ceid , "&_
							"isnull(d.empid,'') deid, d.yymm,isnull(d.jx,0) jxm, a.* from   "&_
							"(  "&_
							"select * from view_empfile  where inyymm<='"&lastMonth&"' and nindat<='"&predat&"'  "&_
							"and isnull(outdat,'')='' and country in ('TA', 'VN', 'CN' ) "&_
							") a  "&_
							"left join (select * from view_empgroup where yymm='"&lastMonth&"' ) b on b.empid = a.empid  "&_
							"left join ( select * from vyfymyjx where jxym='"&lastMonth&"' ) c  on c.empid =a.empid  "&_
							"left join ( select * from empdsalary  where yymm='"&closeYM&"' ) d  on d.empid =a.empid  "&_
							"where  b.lw='"&whsno&"' and    "&_
							"lz <= case when lg='a065' then lg+'5'  else lg+'9' end and  "&_
							"(    ( isnull(c.empid,'')<>'' and isnull(d.empid,'')<>'' and isnull(c.reljxm,0)>0 and isnull(d.jx,0)=0 )   "&_
							"or ( isnull(d.empid,'')<>'' and isnull(d.jx,0) =0 and isnull(c.reljxM,0)>0 )  "&_
							"  )  "&_
							"order by lw, lg desc, a.country, a.empid  "
						Set rs = Server.CreateObject("ADODB.Recordset")
						rs.open sql, conn, 1, 3
						'response.write sql	
						%>		
						<td>7.無績效獎金資料,共 <font color=red><%=rs.recordcount%></font> 筆
						Dữ Liệu Tiền Thưởng Hiệu Xuất,Tổng <font color=red><%=rs.recordcount%></font> Đơn
						<%if rs.recordcount>0 then%><img src="../picture/icon_jia.gif"  align="absmiddle" border=0 onclick="showdata(7)" style='cursor:hand'><%end if%>
						<br>
						<%if  not rs.eof then%>
							<div id=div7 style="Z-index:1;  display:none" > 
								<table  id="myTableGrid" width="100%">
									<Tr height="35px" class="header">																
										<td align=center>績效年月<br>Ngày Thưởng</td>
										<td align=center>廠別<br>Xưởng</td>
										<td align=center>國籍<br>Quốc Tịch</td> 
										<td align=center>工號<br>Mã NV</td>
										<td align=center>部門<br>Bộ Phận</td>
										<td align=center>單位<br>Đơn Vị</td>
										<td align=center>班別<br>Ca</td>
										
										
									</tr>			
									<%while not rs.eof %>
									<Tr>										
										<Td><%=lastMonth%></td>
										<Td><%=rs("lw")%></td>
										<Td><%=rs("country")%></td>
										<Td><%=rs("empid")%></td>
										<Td><%=rs("lg")%><%=rs("lgstr")%></td>
										<Td><%=rs("lz")%><%=rs("lzstr")%></td>
										<Td><%=rs("ls")%></td>										
									</tr>
									<tr>
										<td colspan=7><hr size=0></td>
									</tr>
									<%rs.movenext
									wend					
									rs.close
									set rs=nothing
									%> 
								</table> 
							</div>
						<%end if%>			
						</td>
					</tr>	
					<tr>
						<Td>
						<%
						sql="select isnull(b.bb,0) as b_bb, isnull(b.cv,0) b_cv, isnull(b.phu,0) b_phu , "&_
							"isnull(b.nn,0) b_nn, isnull(b.kt,0) b_kt, isnull(b.mt,0) b_mt, isnull(b.ttkh,0) b_ttkh, "&_
							"a.* from "&_ 
							"(select * from bemps where yymm='"& CloseYM &"'  ) a "&_
							"left join ( select * from empdsalary where  yymm='"&lastMonth&"' ) b on a.empid = b.empid "&_
							"join (select * from view_empgroup where yymm='"& closeYM &"' ) c on c.empid =a.empid  "&_
							"join (select * from empfile  ) d on d.empid =a.empid  "&_
							"where convert(char(6), d.indat,112)<'"& CloseYM &"' and isnull(c.lw,'')='"& whsno &"' "&_
							"and( (isnull(b.bb,0)<>isnull(a.bb,0) and isnull(a.bb,0)<isnull(b.bb,0)) or "&_
							"(isnull(b.phu,0)<>isnull(a.phu,0) and isnull(a.phu,0)<isnull(b.phu,0) ) or  "&_
							"(isnull(b.cv,0)<>isnull(a.cv,0) and isnull(a.cv,0)<isnull(b.cv,0)) or "&_
							"(isnull(b.nn,0)<>isnull(a.nn,0) and isnull(a.nn,0)<isnull(b.nn,0))or "&_
							"(isnull(b.kt,0)<>isnull(a.kt,0) and isnull(a.kt,0)<isnull(b.kt,0)) or "&_
							"(isnull(b.mt,0)<>isnull(a.mt,0) and isnull(a.mt,0)<isnull(b.mt,0))or "&_
							"(isnull(b.ttkh,0)<>isnull(a.ttkh,0) and isnull(a.ttkh,0)<isnull(b.ttkh,0) ) ) "&_
							"order by a.country desc, a.empid" 
						'response.write sql	
						Set rds = Server.CreateObject("ADODB.Recordset")
						rds.open sql, conn, 1, 3	
						%>
						8.其他新資項目與上月不符合,共 <font color=red><%=rds.recordcount%></font> 筆
						Mục Tiền Lương Khác Và Tháng Trước Chênh Lệch , Tổng <font color=red><%=rds.recordcount%></font> Đơn
						<%if rds.recordcount>0 then%><img src="../picture/icon_jia.gif"  align="absmiddle" border=0 onclick="showdata(8)" style='cursor:hand'><%end if%>
						<br>
						<div id=div8 style="Z-index:1;  display:none" > 
						<%if not rds.eof then %>
							<table   id="myTableGrid" width="100%">
							<tr height="35px" class="header">
								<td align=center>STT</td>
								<td align=center>工號<br>Mã NV</td>
								<td align=center>年月<br>Ngày Tháng</td>
								<td align=center>BB</td>
								<td align=center>CV</td>
								<td align=center>Y</td>
								<td align=center>NN</td>
								<td align=center>KT</td>
								<td align=center>MT</td>
								<td align=center>TTKH</td>
							</tr>				
							<%x1=0
							  while not rds.eof 
							  x1=x1+1
							  if cdbl(rds("bb"))<>cdbl(rds("b_bb")) then 
								wkcolor1="Red"
							  else
								wkcolor1="black"
							  end if
							  if cdbl(rds("cv"))<>cdbl(rds("b_cv")) then 
								wkcolor2="Red"
							  else
								wkcolor2="black"
							  end if
							  if cdbl(rds("phu"))<>cdbl(rds("b_phu")) then 
								wkcolor3="Red"
							  else
								wkcolor3="black"
							  end if
							  if cdbl(rds("nn"))<>cdbl(rds("b_nn")) then 
								wkcolor4="Red"
							  else
								wkcolor4="black"
							  end if
							  if cdbl(rds("kt"))<>cdbl(rds("b_kt")) then 
								wkcolor5="Red"
							  else
								wkcolor5="black"
							  end if
							  if cdbl(rds("mt"))<>cdbl(rds("b_mt")) then 
								wkcolor6="Red"
							  else
								wkcolor6="black"
							  end if
							  if cdbl(rds("ttkh"))<>cdbl(rds("b_ttkh")) then 
								wkcolor7="Red"
							  else
								wkcolor7="black"
							  end if
							%>
							<Tr>
								<Td rowspan=2 align=center><%=x1%></td>
								<Td rowspan=2 align=center><%=rds("empid")%></td>
								<Td align=center><b><%=CloseYM%></b></td>
								<Td align=right><font color="<%=wkcolor1%>"><%=rds("bb")%></font></td>
								<Td align=right><font color="<%=wkcolor2%>"><%=rds("cv")%></font></td>
								<Td align=right><font color="<%=wkcolor3%>"><%=rds("phu")%></font></td>
								<Td align=right><font color="<%=wkcolor4%>"><%=rds("nn")%></font></td>
								<Td align=right><font color="<%=wkcolor5%>"><%=rds("kt")%></font></td>
								<Td align=right><font color="<%=wkcolor6%>"><%=rds("mt")%></font></td>
								<Td align=right><font color="<%=wkcolor7%>"><%=rds("ttkh")%></font></td>
							</Tr>
							<tr>
								<Td align=center><%=lastMonth%></td>
								<Td align=right><font color="<%=wkcolor1%>"><%=rds("B_bb")%></font></td>
								<Td align=right><font color="<%=wkcolor2%>"><%=rds("B_cv")%></font></td>
								<Td align=right><font color="<%=wkcolor3%>"><%=rds("B_phu")%></font></td>
								<Td align=right><font color="<%=wkcolor4%>"><%=rds("B_nn")%></font></td>
								<Td align=right><font color="<%=wkcolor5%>"><%=rds("B_kt")%></font></td>
								<Td align=right><font color="<%=wkcolor6%>"><%=rds("B_mt")%></font></td>
								<Td align=right><font color="<%=wkcolor7%>"><%=rds("B_ttkh")%></font></td>
							</tr>
							<Tr>
								<Td colspan=10><hr size=0></td>
							</tr>
							<%rds.movenext
							wend
							%>
							</table>	
							<%end if %> 
							<%
							rds.close
							set rds=nothing%>
							</div>
						</td>
						
					</tr>			
				</table>
			</td>
		</tr>
	</table>
			
</form>
</body>
</html>


<script language=vbs>  
function showdata(a)
	if a=1 then 
		if div1.style.display="none" then 
			div1.style.display=""
		else
			div1.style.display="none"
		end if		
	elseif a=2 then 
		if div2.style.display="none" then 
			div2.style.display=""
		else
			div2.style.display="none"
		end if		
	elseif a=5 then 
		if div5.style.display="none" then 
			div5.style.display=""
		else
			div5.style.display="none"
		end if	
	elseif a=6 then 
		if div6.style.display="none" then 
			div6.style.display=""
		else
			div6.style.display="none"
		end if		
	elseif a=7 then 
		if div7.style.display="none" then 
			div7.style.display=""
		else
			div7.style.display="none"
		end if				
	elseif a=8 then 
		if div8.style.display="none" then 
			div8.style.display=""
		else
			div8.style.display="none"
		end if					
	end if 	
end function 

function strchg(a)
	if a=1 then 
		<%=self%>.empid1.value = Ucase(<%=self%>.empid1.value)
	elseif a=2 then 	
		<%=self%>.empid2.value = Ucase(<%=self%>.empid2.value)
	end if 	
end function 
	
function go()  
	if <%=self%>.closeym.value="" then 
		alert "請輸入關帳年月"
		<%=self%>.closeym.focus()
		exit function 
	elseif len(<%=self%>.closeym.value)<>6 then 
		alert "關帳年月輸入錯誤!!"
		<%=self%>.closeym.value=""
		<%=self%>.closeym.focus()
		exit function 
	end if	
 	<%=self%>.action="<%=self%>.Fore.asp"
 	<%=self%>.submit() 
end function   
	

'*******檢查日期*********************************************
FUNCTION date_change(a)	

if a=1 then
	INcardat = Trim(<%=self%>.indat1.value)  		
elseif a=2 then
	INcardat = Trim(<%=self%>.indat2.value)
end if		
		    
IF INcardat<>"" THEN
	ANS=validDate(INcardat)
	IF ANS <> "" THEN
		if a=1 then
			Document.<%=self%>.indat1.value=ANS			
		elseif a=2 then
			Document.<%=self%>.indat2.value=ANS		 			
		end if		
	ELSE
		ALERT "EZ0067:輸入日期不合法 !!" 
		if a=1 then
			Document.<%=self%>.indat1.value=""
			Document.<%=self%>.indat1.focus()
		elseif a=2 then
			Document.<%=self%>.indat2.value=""
			Document.<%=self%>.indat2.focus()
		end if		
		EXIT FUNCTION
	END IF
		 
ELSE
	'alert "EZ0015:日期該欄位必須輸入資料 !!" 		
	EXIT FUNCTION
END IF 
   
END FUNCTION
</script> 
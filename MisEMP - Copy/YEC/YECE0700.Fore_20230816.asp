<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%> 
<!-- #include file = "../GetSQLServerConnection.fun" --> 
<!-- #include file="../ADOINC.inc" -->
<!--#include file="../include/checkpower.asp"--> 
<!--#include file="../include/sideinfo.inc"--> 
<%
Set conn = GetSQLServerConnection()	  
self="YECE0700"  


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

w = request("whsno") 
ym = request("ym") 
 
redim aa(7,2) 

sql="select * from empbh_set where w='"& w &"' and ym='"& ym &"' "
set rs=conn.execute(Sql)
if not rs.eof then 
	xstr = split(rs("setstr"),"+")  	
	'response.write ubound(xstr) &"<BR>"	
	for i  = 1 to ubound(xstr)  
		'response.write xstr(i-1) &"<BR>"
		
		if xstr(i-1)="BB" then 
			aa(1,0)= "Y"  
			aa(1,1) = xstr(i-1)
		elseif xstr(i-1)="CV" then 
			aa(2,0)= "Y"			
			aa(2,1) = xstr(i-1)
		elseif xstr(i-1)="PHU" then 
			aa(3,0)= "Y" 
			aa(3,1) = xstr(i-1)
		elseif xstr(i-1)="NN" then 
			aa(4,0)= "Y"			
			aa(4,1) = xstr(i-1)
		elseif xstr(i-1)="KT" then 
			aa(5,0)= "Y"			
			aa(5,1) = xstr(i-1)
		elseif xstr(i-1)="MT" then 
			aa(6,0)= "Y"			
			aa(6,1) = xstr(i-1)
		elseif xstr(i-1)="BTIEN" then 
			aa(7,0)= "Y"			
			aa(7,1) = xstr(i-1)			
		end if 	 		
	next 	
end if 
set rs=nothing 

sql2="select * from empbh_per where yymm='"& ym &"' and country='VN'  "
set rs2=conn.execute(sql2)
if  not rs2.eof then 
	emp_bhxh = rs2("emp_bhxh")
	emp_bhyt = rs2("emp_bhyt")
	emp_bhtn = rs2("emp_bhtn")
	cty_bhxh = rs2("cty_bhxh")
	Cty_bhyt = rs2("cty_bhyt")
	cty_bhtn = rs2("Cty_bhtn")
	emp_gtamt = rs2("emp_gtant")
	gt_per =rs2("gt_per")
end if 
rs2.close :set rs2=nothing 

sql2="select * from empbh_per where yymm='"& ym &"' and country='HW'  "
set rs2=conn.execute(sql2)
if  not rs2.eof then 
	hwemp_bhxh = rs2("emp_bhxh")
	hwemp_bhyt = rs2("emp_bhyt")
	hwemp_bhtn = rs2("emp_bhtn")
	hwcty_bhxh = rs2("cty_bhxh")
	hwCty_bhyt = rs2("cty_bhyt")
	hwcty_bhtn = rs2("Cty_bhtn")
end if 
rs2.close :set rs2=nothing 

%>

<html> 
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="refresh">
<link rel="stylesheet" href="../Include/style.css" type="text/css">
<link rel="stylesheet" href="../Include/style2.css" type="text/css">
 
</head> 
<body  topmargin="5" leftmargin="5"  marginwidth="0" marginheight="0"   onload='f()' onkeydown="enterto()">
<form name="<%=self%>" method="post" action="EMPFILE.SALARY.ASP">

	<table border=0 style="height:70px"><tr><td>&nbsp;</td></tr></table>
	<table width="100%" border=0 >
		<tr>
			<td>
				<table width="90%" BORDER=0 align=center cellpadding=3 cellspacing=3 >
					<tr>
						<td>
							<table class="txt" cellpadding=3 cellspacing=3>
								<tr>
									<Td align="right">廠別<br>Xưởng</td>
									<Td>	
										<select name=WHSNO  style='width:100px' onchange="datachg()">					
											<%
											if session("rights")=0 then 
												SQL="SELECT * FROM BASICCODE WHERE FUNC='WHSNO' ORDER BY SYS_TYPE "%>
												<option value="" selected >全部(ALL) </option>
											<%		
											else
												SQL="SELECT * FROM BASICCODE WHERE FUNC='WHSNO' and sys_type='"& session("NETWHSNO") &"' ORDER BY SYS_TYPE "
											end if	
											SET RST = CONN.EXECUTE(SQL)
											WHILE NOT RST.EOF
											%>
											<option value="<%=RST("SYS_TYPE")%>" <%if rst("sys_type")=request("WHSNO") then%>selected<%end if%>><%=RST("SYS_VALUE")%></option>
											<%
											RST.MOVENEXT
											WEND
											%>
										</SELECT>
										<%SET RST=NOTHING %>	 			
									</td>
								</tr>
							</table>
						</td>
					</tr>
					<tr>
						<td>
							<table class="txt"  cellpadding=3 cellspacing=0>								
								<tr height=35  bgcolor="#e4e4e4" class="txt">			
									<TD nowrap width=70 align="center">計薪年月<br>Ngày thống kê<br>(yyyymm)</TD>			
									<TD nowrap align="center" >計算項目<br>Hạng mục tính toán</TD>			
								</TR>
								<Tr class="txt">	
									<Td align="center">
										<input type="text" style="width:100px" name="ym" value="<%=ym%>"  maxlength="6" onblur="datachg()">
									</td>
									<Td nowrap >
										<input type="checkbox" name=fc onclick="dchg(1)" <%if aa(1,0)="Y" then%>checked<%end if%>>基薪(Cơ bản)
										<input type="checkbox" name=fc onclick="dchg(2)" <%if aa(2,0)="Y" then%>checked<%end if%>>職務加給(CV)
										<input type="checkbox" name=fc onclick="dchg(3)" <%if aa(3,0)="Y" then%>checked<%end if%>>補助(PCK)
										<input type="checkbox" name=fc onclick="dchg(4)" <%if aa(4,0)="Y" then%>checked<%end if%>>語言(NN)			
										<input type="checkbox" name=fc onclick="dchg(5)" <%if aa(5,0)="Y" then%>checked<%end if%>>技術(KT)
										<input type="checkbox" name=fc onclick="dchg(6)" <%if aa(6,0)="Y" then%>checked<%end if%>>環境加給(MT)
										<input type="checkbox" name=fc onclick="dchg(7)" <%if aa(7,0)="Y" then%>checked<%end if%>>補薪(Bù lương)			
										<%for x = 1 to 7 %>
												<input type="hidden" name=fcode size=1 class="inputbox8" value="<%=aa(x,1)%>">
										<%next%>				
									</td>
								</Tr> 		
							</table>
						</td>
					</tr>
					<tr>
						<td>
							<table class="txt" width="60%" cellpadding=3 cellspacing=0> 
								<tr>
									<td colspan=6>保險費率設定 Cài đặt phân trăm bảo hiểm</td>
								</tr>
								<tr bgcolor="#e4e4e4" class="txt">
									<Td colspan=3 align="center" bgcolor="#FBDBFF">(VN)Nhân viên</td>
									<Td colspan=3 align="center">(VN)Công ty</td>
									<Td colspan=2 align="center">工團<br>Công đoàn</td>
								</tr>
								<tr  bgcolor="#e4e4e4" class="txt">
									<td align="center" bgcolor="#FBDBFF">BHXH%</td>
									<td align="center" bgcolor="#FBDBFF">BHYT%</td>
									<td align="center" bgcolor="#FBDBFF">BHTN%</td>
									<td align="center">BHXH%</td>
									<td align="center">BHYT%</td>
									<td align="center">BHTN%</td>
									<td align="center">工團<br>Công đoàn</td>
									<td align="center">繳交<br>比例%</td>
								</tr>
								<tr class="txt">			
									<input name=ct1 size=5 class="inputbox" style="text-align:center" value="VN" type="hidden" >
									<td align="center"><input type="text" name=emp_bhxh style="width:100%;text-align:center" value="<%=emp_bhxh%>"></td>
									<td align="center"><input type="text" name=emp_bhyt  style="width:100%;text-align:center" value="<%=emp_bhyt%>"></td>
									<td align="center"><input type="text" name=emp_bhtn  style="width:100%;text-align:center" value="<%=emp_bhtn%>"></td>
									<td align="center"><input type="text" name=Cty_bhxh  style="width:100%;text-align:center" value="<%=cty_bhxh%>"></td>
									<td align="center"><input type="text" name=Cty_bhyt  style="width:100%;text-align:center" value="<%=Cty_bhyt%>"></td>
									<td align="center"><input type="text" name=Cty_bhtn  style="width:100%;text-align:center" value="<%=Cty_bhtn%>"></td>
									<td align="center"><input type="text" name=emp_gtamt  style="width:100%;text-align:center" value="<%=emp_gtamt%>"></td>
									<td align="center"><input type="text" name=gt_per  style="width:100%;text-align:center" value="<%=gt_per%>"></td>
								</tr>
								<tr height=10>
									<td colspan=6>&nbsp;</td>
								</tr>			
								<tr bgcolor="#e4e4e4" class="txt">
									<Td colspan=3 align="center" bgcolor="#FBDBFF">(外籍 Nước ngoài)Nhân viên</td>
									<Td colspan=3 align="center">(外籍 Nước ngoài)Công ty</td>
								</tr>
								<tr  bgcolor="#e4e4e4" class="txt">
									<td align="center" bgcolor="#FBDBFF">BHXH%</td>
									<td align="center" bgcolor="#FBDBFF">BHYT%</td>
									<td align="center" bgcolor="#FBDBFF">BHTN%</td>
									<td align="center">BHXH%</td>
									<td align="center">BHYT%</td>
									<td align="center">BHTN%</td>
								</tr> 
								<tr class="txt">			
									<input type="hidden" name=ct2 size=5 class="inputbox" style="text-align:center" value="HW"  >
									<td align="center"><input name=hwemp_bhxh style="width:100%;text-align:center" value="<%=hwemp_bhxh%>"></td>
									<td align="center"><input name=hwemp_bhyt style="width:100%;text-align:center" value="<%=hwemp_bhyt%>"></td>
									<td align="center"><input name=hwemp_bhtn style="width:100%;text-align:center" value="<%=hwemp_bhtn%>"></td>
									<td align="center"><input name=hwCty_bhxh style="width:100%;text-align:center" value="<%=hwcty_bhxh%>"></td>
									<td align="center"><input name=hwCty_bhyt style="width:100%;text-align:center" value="<%=hwCty_bhyt%>"></td>
									<td align="center"><input name=hwCty_bhtn style="width:100%;text-align:center" value="<%=hwCty_bhtn%>"></td>
								</tr>	
								<tr bgcolor="#d7d7d7" class="txt">
									<td  colspan=6 height=55>
										(2) yyyymm :
										<input type="text" name="copyym" style="width:100px" maxlength=6 > 
										<input type=button  name=btm class=button value="(C) COPY" onclick="gocopy()" style="background-color:#FFBFBF;"  >
										<br>( Copy (1)yyyymm data to (2)yyyymm  複製Bản sao(1)<br>設定資料到指定年月Cài đặt dữ liệu cho năm tháng chỉ định (2) )
									</td>
								</tr>
								<tr>
									<td align="center" colspan=6>
										<table class="txt">
											<tr >
												<td align="center" >
													<input type=button  name=btm class="btn btn-sm btn-danger" value="(Y)Confirm" onclick="go()" onkeydown="go()">
													<input type=reset  name=btm class="btn btn-sm btn-outline-secondary" value="(N)Cancel">				
												</td>
											</tr>	
										</table>
									</td>
								</tr>
							</table>
						</td>
					</tr>					
				</table>
			</td>
		</tr>
	</table>

</body>
</html>
<script language=vbs>  

function dchg(a)	
	select case a 
		case 1 
			T="BB"
		case 2 
			T="CV" 
		case 3 
			T="PHU" 		
		case 4 
			T="NN" 		
		case 5 
			T="KT" 		
		case 6 
			T="MT" 		
		case 7 
			T="BTIEN" 			
	end select 		
	if <%=self%>.fc(a-1).checked=true then
		<%=self%>.fcode(a-1).value = T 
	else	
		<%=self%>.fcode(a-1).value = ""
	end if 	  
end function 

function f()
	if <%=self%>.ym.value="" then 
		<%=self%>.ym.focus()
	end if 	
end function   

function gocopy()
	if trim(<%=self%>.ym.value)="" then 
		alert "請先輸入(1)年月 , xin danh lai tang nam (1)"
		<%=self%>.ym.focus()
		exit function 
	end if 
	if trim(<%=self%>.copyym.value)="" then 
		alert "請輸入要複製的年月(2) , xin danh lai tang nam (2)"
		<%=self%>.copyym.focus()
		exit function 
	end if 
	
	if confirm("確定要複製 ["& <%=self%>.ym.value &"] 的設定到 ["&<%=self%>.copyym.value&"] ?" , 64 ) then 
		<%=self%>.action="<%=self%>.upd.asp?flag=C"
		<%=self%>.submit() 
	end if  
	
end function 

function enterto()
	if window.event.keyCode = 13 then window.event.keyCode =9 		 	
end function 

function datachg()
	if <%=self%>.ym.value<>"" then 
		if len(<%=self%>.ym.value)<>6 then 
			alert "請輸入年月6碼(ex:200701)"
			<%=self%>.ym.value=""
			<%=self%>.ym.focus()
		end if 	
	end if 
	if <%=self%>.whsno.value<>"" and <%=self%>.ym.value<>"" then 
		<%=self%>.action ="<%=self%>.fore.asp" 
		<%=self%>.submit()
	end if 
end function  
 
	
function go()  
 	<%=self%>.action="<%=self%>.upd.asp"
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
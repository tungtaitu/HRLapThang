<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%> 
<!-- #include file = "../GetSQLServerConnection.fun" --> 
<!-- #include file="../ADOINC.inc" -->
<!-- #include file="../Include/func.inc" -->
<!-- #include file="../Include/SIDEINFO.inc" -->
<%
Set conn = GetSQLServerConnection()	  
self="YFYEMPJX"  


nowmonth = year(date())&right("00"&month(date()),2)
if month(date())="01" then  
	calcmonth = year(date()-1)&"12"  	
else
	calcmonth =  year(date())&right("00"&month(date())-1,2)   
end if 	   

if day(date())<=11 then 
	if month(date())="01" then  
		calcmonth = year(date()-1)&"12" 
	else
		calcmonth =  year(date())&right("00"&month(date())-1,2)   
	end if 	 	
else
	calcmonth = nowmonth 
end if  

NNY=year(date())
NDY=year(date())+1 

gid = request("groupid") 

'response.write wx 

%>

<html> 
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="refresh">
<link rel="stylesheet" href="../Include/style.css" type="text/css">
<link rel="stylesheet" href="../Include/style2.css" type="text/css">
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--
function f()
	<%=self%>.JXYM.focus()		
end function     

function dchg()
	<%=self%>.action = "<%=self%>.Fore.asp"
	<%=self%>.submit()
end  function 
-->
</SCRIPT>   
</head> 
<body  leftmargin="5"  marginwidth="0" marginheight="0"  onkeydown="enterto()" onload='f()' >
<form name="<%=self%>" method="post" action="EMPFILE.SALARY.ASP">
<input type=hidden name="NNY" value="<%=NNY%>">
<input type=hidden name="NDY" value="<%=NDY%>">
 
	<table border=0 style="height:70px"><tr><td>&nbsp;</td></tr></table>
	<table width="100%" border=0 >
		<tr>
			<td>
				<table width="100%" BORDER=0 align=center cellpadding=0 cellspacing=0 >
					<tr>
						<td>
							<table class="txt" cellpadding=3 cellspacing=3>
								<tr class="txt">
									<td align="center"><font class="btn btn-warning text-white btn-block shadow"> 績效新增作業<br>New </font></td> 
									<td align="center"><a href="<%=self%>.ForeEDIT.asp" class="btn btn-primary btn-block shadow">績效修改作業<br>Edit</A></td>
									<td align="center"><a href="yfyempjx.sch.asp" class="btn btn-primary btn-block shadow">績效查詢作業<br>Search</a></td>
								</tr>								
							</table>
						</td>
					</tr>
					<tr><td><hr size=0	style='border: 1px dotted #999999;' align=left ></td></tr>
					<tr>
						<td>
							<table class="txt" cellpadding=3 cellspacing=3>
								<TR>
									<TD align="right">績效年月<br><font class="txt8">Tien Thuong</font></TD>
									<TD><INPUT type="text" style="width:100px" NAME=JXYM VALUE="<%=request("JXYM")%>" ><font class="txt8">(yymm)</font></TD>
									<TD align="right">計薪年月<br><font class="txt8">Tien Luong</font></TD>
									<TD><INPUT type="text" style="width:100px" NAME=SALARYYM VALUE="<%=request("SALARYYM")%>" ><font class="txt8">(yymm)</font></TD>
									<TD align="right">廠別<br><font class="txt8">xuong</font></TD>
									<TD>
										<SELECT NAME=jxwhsno style="width:120px">
											<%if session("rights")="0" then
												SQL="SELECT* FROM BASICCODE WHERE FUNC='whsno' order by sys_type "
											%>	<option value=""></option>
											<%else	
												SQL="SELECT* FROM BASICCODE WHERE FUNC='whsno' and sys_type like '"&session("netwhsno")&"' order by sys_type "
											  end if
											  SET RST=CONN.EXECUTE(SQL)
											  WHILE NOT RST.EOF 
											%><OPTION VALUE="<%=RST("SYS_TYPE")%>" <%if rst("sys_type")=wx then%>selected<%end if%>><%=RST("SYS_VALUE")%></OPTION>
											<%RST.MOVENEXT%>
											<%WEND%>
										</SELECT>	 			
									</TD> 
								</tr>
								<tr>	
									<TD align="right">部門<br><font class="txt8">Bo phan</font></TD>
									<TD>
										<SELECT NAME=GROUPID   onchange="dchg()" style="width:120px">
											<option value=""></option>
											<%SQL="SELECT* FROM BASICCODE WHERE FUNC='GROUPID'  order by  case when sys_type='A065' then 'a000' else sys_type end  "
											  SET RST=CONN.EXECUTE(SQL)
											  WHILE NOT RST.EOF 
											%><OPTION VALUE="<%=RST("SYS_TYPE")%>" <%if rst("sys_type")=gid then%>selected<%end if%>><%=RST("SYS_VALUE")%></OPTION>
											<%RST.MOVENEXT%>
											<%WEND%>
										</SELECT>	 			
									</TD> 
									<TD align="right">組別<br><font class="txt8">To</font></TD>
									<TD>
										<SELECT NAME=zuno style="width:120px">
											<OPTION VALUE="">None</OPTION>
											<%SQL="SELECT* FROM BASICCODE WHERE FUNC='zuno' and left(sys_type,4)='"& gid &"' "
											  SET RST=CONN.EXECUTE(SQL)
											  WHILE NOT RST.EOF 
											%><OPTION VALUE="<%=RST("SYS_TYPE")%>"><%=RST("SYS_VALUE")%></OPTION>
											<%RST.MOVENEXT%>
											<%WEND%>
										</SELECT>	 			
									</TD>	 			
									<TD align="right">班別<br><font class="txt8">Ca</font></TD>
									<TD>
										<SELECT NAME=SHIFT style="width:120px">
											<OPTION VALUE="">None</OPTION>
											<% 
												SQL="SELECT* FROM BASICCODE WHERE FUNC='shift' order by len(sys_type), sys_type  "	 					
											  SET RST=CONN.EXECUTE(SQL)
											  WHILE NOT RST.EOF 
											%><OPTION VALUE="<%=RST("SYS_TYPE")%>"  ><%=RST("SYS_VALUE")%></OPTION>
											<%RST.MOVENEXT%>
											<%WEND%>
										</SELECT>	 						
										
									</TD>
									
								</TR>
							</TABLE>
						</td>
					</tr>
					<tr>
						<td align="center">
							<table id="myTableGrid" width="98%">
								<TR BGCOLOR="#FFF278" CLASS=TXT9>
									<TD  HEIGHT=22 ALIGN=CENTER>STT</TD>
									<TD  ALIGN=CENTER>說明<br>Thuyết minh</TD>
									<TD  ALIGN=CENTER>實績<br>Thực tế</TD>
									<TD  ALIGN=CENTER>係數<br>Hệ số</TD>
									<TD  ALIGN=CENTER>比例<br>%</TD>	 			
								</TR>
								<%FOR I = 1 TO 5 %>
								<TR BGCOLOR="#FFFFFF" CLASS=TXT9>
									<TD HEIGHT=22 ALIGN=CENTER ><INPUT type="text" style="width:100%" NAME="STT" VALUE="<%=CHR(64+I)%>" READONLY ></TD>
									<TD ALIGN=CENTER>
										<INPUT type="text" style="width:100%" NAME=DESCP VALUE="">
									</TD>
									<TD ALIGN=CENTER>
										<INPUT type="text" style="width:100%" NAME=HXSL VALUE="">
									</TD>	 			
									<TD ALIGN=CENTER><INPUT type="text" style="width:100%" NAME="HESO" VALUE=""></TD>
									<TD ALIGN=CENTER><INPUT type="text" style="width:100%" NAME="PER" VALUE=""></TD>
								</TR>
								<%NEXT %>
							</TABLE>
						</td>
					</tr>
					<tr>
						<td align="center">
							<table class="txt" cellpadding=3 cellspacing=3>
								<tr >
									<td align=center>
										<input type=button  name=btm class="btn btn-sm btn-danger" value="(Y)Confirm" onclick="go()" onkeydown="go()">
										<input type=reset  name=btm class="btn btn-sm btn-outline-secondary" value="(N)Cancel">
										<input type=reset  name=btm class="btn btn-sm btn-outline-secondary" value="(C)Data Copy" ONCLICK=COPYDATA()>
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


'_________________DATE CHECK___________________________________________________________________

function validDate(d)
	if len(d) = 8 and isnumeric(d) then
		d = left(d,4) & "/" & mid(d, 5, 2) & "/" & right(d,2)
		if isdate(d) then
			validDate = formatDate(d)
		else
			validDate = ""
		end if
	elseif len(d) >= 8 and isdate(d) then
			validDate = formatDate(d)
		else
			validDate = ""
		end if
end function

function formatDate(d)
		formatDate = Year(d) & "/" & _
		Right("00" & Month(d), 2) & "/" & _
		Right("00" & Day(d), 2)
end function
'________________________________________________________________________________________  

FUNCTION GO()
	IF <%=SELF%>.JXYM.VALUE="" THEN 
		ALERT "請輸入績效年月!!"
		<%=SELF%>.JXYM.FOCUS()
		EXIT FUNCTION 
	ELSEIF <%=SELF%>.SALARYYM.VALUE="" THEN 	
		ALERT "請輸入計薪年月!!"
		<%=SELF%>.SALARYYM.FOCUS()
		EXIT FUNCTION 
	ELSEIF <%=SELF%>.GROUPID.VALUE="" THEN 
		ALERT "請輸入單位!!"
		<%=SELF%>.GROUPID.FOCUS()
		EXIT FUNCTION 
	ELSEIF <%=SELF%>.jxwhsno.VALUE="" THEN 
		ALERT "請輸入廠別!!"
		<%=SELF%>.jxwhsno.FOCUS()
		EXIT FUNCTION
	ELSE
		<%=SELF%>.ACTION="<%=SELF%>.UPD.ASP"
		<%=SELF%>.SUBMIT()		
	END IF 
END FUNCTION  

FUNCTION COPYDATA()
	'IF <%=SELF%>.JXYM.VALUE="" THEN 
	'	ALERT "請輸入欲複製的績效年月"
	'	<%=SELF%>.JXYM.FOCUS()
	'	EXIT FUNCTION 
	'END IF 
	<%=SELF%>.ACTION="<%=SELF%>.CopyData.ASP"
	<%=SELF%>.SUBMIT()		
END FUNCTION 
</script> 
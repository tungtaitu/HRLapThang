<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%> 
<!-- #include file = "../GetSQLServerConnection.fun" --> 
<!-- #include file="../ADOINC.inc" -->
<!-- #include file="../Include/func.inc" -->
<!-- #include file="../Include/SIDEINFO.inc" -->
<%
Set conn = GetSQLServerConnection()	  
self="YECE06"  


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

if right(calcmonth,2)="01" then 
	sgym = left(calcmonth,4)-1 & "12" 
else
	sgym = left(calcmonth,4)&right("00"&right(calcmonth,2)-1,2)
end if 	 

F_groupid=request("F_groupid")
YYMM = request("YYMM")
if YYMM="" then YYMM=calcmonth 
JXYM = request("JXYM") 
if JXYM="" then JXYM=sgym 

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
	<%=self%>.salaryYM.focus()		
end function 

function dchg()
	<%=self%>.action = "<%=self%>.Fore.asp"
	<%=self%>.submit()
end  function   

function newPage()
wt = (window.screen.width )*0.7
	ht = window.screen.availHeight*0.7
	tp = (window.screen.width )*0.02
	lt = (window.screen.availHeight)*0.02
	open "yece0601.asp"  , "_blank" , "top="& tp &", left="& lt &", width="& wt &",height="& ht &", scrollbars=yes"
	
end function 
   
-->
</SCRIPT>   
</head> 
<body  topmargin="50" leftmargin="5"  marginwidth="0" marginheight="0"  onkeydown="enterto()" onload='f()' >
<form name="<%=self%>" method="post" action="<%=self%>.ForeGnd.asp">
 
	<table border=0 style="height:70px"><tr><td>&nbsp;</td></tr></table>	
	<table width="100%" BORDER=0 align=center cellpadding=0 cellspacing=0 >
		<tr>
			<td align="center">
				<table id="myTableForm" width="60%">
					<tr><td height="40px">&nbsp;</td></tr>
					<tr>
						<TD nowrap align=right>計薪年月<br><font class="txt8">Tien Luong</font></TD>
						<TD ><INPUT type="text" style="width:100px" NAME=salaryYM   VALUE="<%=yymm%>"><font class="txt8">(YYYYMM)</font></TD>	
						<TD nowrap align=right>績效年月<br><font class="txt8">Tien Thuong</font></TD>
						<TD ><INPUT type="text" style="width:100px" NAME=JXYM   VALUE="<%=jxym%>"><font class="txt8">(YYYYMM)</font></TD>	
					</TR>
					<tr>
						<TD nowrap align=right height=30 >國籍<br><font class="txt8">Quoc tich</font></TD>
						<TD >
							<select name=country style="width:120px">
								<option value="">----</option>
								<%SQL="SELECT * FROM BASICCODE WHERE FUNC='country'  ORDER BY SYS_type desc  "
								SET RST = CONN.EXECUTE(SQL)
								WHILE NOT RST.EOF  
								%>
								<option value="<%=RST("SYS_TYPE")%>" ><%=RST("SYS_TYPE")%>-<%=RST("SYS_VALUE")%></option>				 
								<%
								RST.MOVENEXT
								WEND 
								SET RST=NOTHING
								%>
							</SELECT>
						</TD>
						<TD nowrap align=right height=30 >廠別<br><font class="txt8">Xuong</font></TD>
						<TD > 
							<select name=F_WHSNO style="width:120px"> 
								<%
								if session("rights")="0" then 
									SQL="SELECT * FROM BASICCODE WHERE FUNC='WHSNO' ORDER BY SYS_TYPE "
								%><option value="">----</option>
								<%	
								else		
									SQL="SELECT * FROM BASICCODE WHERE FUNC='WHSNO' and sys_type='"& session("netwhsno") &"' ORDER BY SYS_TYPE "
								end if	
								SET RST = CONN.EXECUTE(SQL)
								WHILE NOT RST.EOF  
								%>
								<option value="<%=RST("SYS_TYPE")%>"  <%if request("F_WHSNO")=RST("SYS_TYPE") then %>selected<%end if%>><%=RST("SYS_VALUE")%></option>				 
								<%
								RST.MOVENEXT
								WEND 
								SET RST=NOTHING
								%>
							</SELECT>										
						</TD>										
					</tr>								 
					<TR>
						<TD align=right>部門<br><font class="txt8">Bo phan</font></TD>
							<TD>
							<SELECT NAME=F_GROUPID style="width:120px">
								<option value="">----</option>
								<%SQL="SELECT* FROM BASICCODE WHERE FUNC='GROUPID'  and sys_type<>'AAA' and sys_type>'A021'  "&_	 					  
									  "order by   sys_type "
								  SET RST=CONN.EXECUTE(SQL)
								  WHILE NOT RST.EOF 
								%><OPTION VALUE="<%=RST("SYS_TYPE")%>" <%if rst("sys_type")=F_groupid  then%>selected<%end if%>><%=RST("SYS_TYPE")%>-<%=RST("SYS_VALUE")%></OPTION>
								<%RST.MOVENEXT%>
								<%WEND%>
								<%set rst=nothing %>
							</SELECT>	 			
						</TD> 			
						<TD align=right>組別<br><font class="txt8">To</font></TD>
						<TD>
							<SELECT NAME=F_zuno style="width:120px">
								<OPTION VALUE="">---</OPTION>
								<%SQL="SELECT* FROM BASICCODE WHERE FUNC='zuno' and left(sys_type,4)='"& f_groupid &"' order by sys_type "
								  SET RST=CONN.EXECUTE(SQL)
								  WHILE NOT RST.EOF 
								%><OPTION VALUE="<%=RST("SYS_TYPE")%>"><%=RST("SYS_TYPE")%>-<%=RST("SYS_VALUE")%></OPTION>
								<%RST.MOVENEXT%>
								<%WEND%>
								<%set rst=nothing %>
							</SELECT>	 			
						</TD> 
					</tr>
					<TR>
						<td nowrap align=right >班別<br><font class="txt8">Ca</font></td>
						<td>
							<select name="F_shift" style="width:120px"> 			 		
								<option value="">---</option>			 		
								<%SQL="SELECT* FROM BASICCODE WHERE FUNC='shift'   order by len(sys_type) desc , sys_type  "
								  SET RST=CONN.EXECUTE(SQL)
								  WHILE NOT RST.EOF 
								%><OPTION VALUE="<%=RST("SYS_TYPE")%>"><%=RST("SYS_TYPE")%>-<%=RST("SYS_VALUE")%></OPTION>
								<%RST.MOVENEXT%>
								<%WEND%>
								<%set rst=nothing %>					
							</select>	
						</td>
						<td nowrap align=right >工號<br><font class="txt8">So the</font></td>
						<td>
							<input type="text" style="width:100px" name=empid1> 			 	
						</td>
					</TR>
					<tr >
						<td align=center colspan=4 height="50px">
							<input type=button  name=btm class="btn btn-sm btn-danger" value="(Y)Confirm" onclick="go()" onkeydown="go()">
							<input type=reset  name=btm class="btn btn-sm btn-outline-secondary" value="(N)Cancel">
							<input type=reset  name=btm class="btn btn-sm btn-outline-secondary" value="(S)K.Tra(查詢)" ONCLICK="GETDATA()">
							&nbsp; 
							<input type=reset  name=btm class="btn btn-sm btn-outline-secondary" value="(T)員工績效計算" ONCLICK="newPage()">
						</td>
					</tr>
				</table>
			</td>
		</tr>		
		<tr>
			<td>
				<table class="table-borderless table-sm bg-white text-secondary">
					<%
					sql="select   jxym, groupid, isnull(zuno,'') zuno , shift  from VYFYMYJX  where  jxym='"& sgym &"'  group by jxym, groupid, isnull(zuno,'')    , shift  "&_
						"order by groupid, shift "
					'response.write sql	
					set rds=conn.execute(Sql)
					while not rds.eof  	
					%>
					<tr class="txt">
						<td align=center>
							資料已處理  <%=rds("groupid")%> -- <%=rds("zuno")%> -- <%=right("   "&rds("shift"),3)%>				
						</td>
					</tr>
					<%rds.movenext
					wend 
					%>
				</table>
			</td>
		</tr>
	</table>
			
</form>
</body>
</html>


<script language=vbs>  

function strchg(a)
	if a=1 then 
		<%=self%>.empid1.value = Ucase(<%=self%>.empid1.value)
	elseif a=2 then 	
		<%=self%>.empid2.value = Ucase(<%=self%>.empid2.value)
	end if 	
end function 
	
function go()   
	if <%=self%>.F_WHSNO.value="" then 
		alert "請選擇廠別Xin danh lai Xuong!!"
		<%=self%>.F_WHSNO.focus()
		exit function 
	end if 	
 	'<%=self%>.action="<%=SELF%>.FOREGND.asp"
 	<%=self%>.submit
end function   

FUNCTION GETDATA()
	<%=self%>.action="YFYEMPJX.SCH.asp"
 	<%=self%>.submit()
END FUNCTION  
	

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

 
</script> 
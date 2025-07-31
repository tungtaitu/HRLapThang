<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%> 
<!-- #include file = "../../GetSQLServerConnection.fun" --> 
<!-- #include file="../../ADOINC.inc" -->
<!-- #include file="../../Include/func.inc" -->
<%
Set conn = GetSQLServerConnection()	  
self="YFYEMPJXAF"  


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

%>

<html> 
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="refresh">
<link rel="stylesheet" href="../../Include/style.css" type="text/css">
<link rel="stylesheet" href="../../Include/style2.css" type="text/css">
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--
function f()
	<%=self%>.YYMM.focus()		
end function    
-->
</SCRIPT>   
</head> 
<body  topmargin="5" leftmargin="5"  marginwidth="0" marginheight="0"  onkeydown="enterto()" onload='f()' >
<form name="<%=self%>" method="post" action="YFYEMPJXA.ForeGnd.asp">
<table width="460" border="0" cellspacing="0" cellpadding="0">
	<tr><TD>
	<img border="0" src="../../image/icon.gif" align="absmiddle">
	績效獎金計算</TD></tr>
</table>
<hr size=0	style='border: 1px dotted #999999;' align=left width=500>
<BR><BR>
<table width=550 border=0 ><tr><td >
	 	<table width=400 align=center border=0 cellspacing="0" cellpadding="0"  > 
		<tr height=30 >
			<TD nowrap align=right>計薪年月：</TD>
			<TD ><INPUT NAME=YYMM  CLASS=INPUTBOX VALUE="<%=calcmonth%>" SIZE=10></TD>	
		</TR>
		<tr height=30 >
			<TD nowrap align=right>績效年月：</TD>
			<TD ><INPUT NAME=JXYM  CLASS=INPUTBOX VALUE="<%=sgym%>" SIZE=10></TD>	
		</TR>
		<TR>
		 	<TD nowrap align=right height=30 >國籍：</TD>
			<TD >
				<select name=country  class=font9 style='width:75'  >
					<option value="">全部 </option>
					<%SQL="SELECT * FROM BASICCODE WHERE FUNC='country'  ORDER BY SYS_type desc  "
					SET RST = CONN.EXECUTE(SQL)
					WHILE NOT RST.EOF  
					%>
					<option value="<%=RST("SYS_TYPE")%>" ><%=RST("SYS_VALUE")%></option>				 
					<%
					RST.MOVENEXT
					WEND 
					%>
				</SELECT>
				<%SET RST=NOTHING %>
			</TD>	
		</tr>
		<!--tr>		 
			<TD nowrap align=right height=30 >廠別：</TD>
			<TD > 
				<select name=WHSNO  class=font9 >
					<option value="">全部 </option>
					<%SQL="SELECT * FROM BASICCODE WHERE FUNC='WHSNO' ORDER BY SYS_TYPE "
					SET RST = CONN.EXECUTE(SQL)
					WHILE NOT RST.EOF  
					%>
					<option value="<%=RST("SYS_TYPE")%>"><%=RST("SYS_VALUE")%></option>				 
					<%
					RST.MOVENEXT
					WEND 
					%>
				</SELECT>
				<%SET RST=NOTHING %>
			</TD> 
		</TR-->		
		<tr>	 
		<TR height=30 >
			<TD nowrap align=right >組/部門：</TD>
			<TD >
				<select name=GROUPID  class=font9  >
				<option value="" selected >全部 </option>
				<%
				SQL="SELECT * FROM BASICCODE WHERE FUNC='GROUPID' and sys_type <>'AAA' ORDER BY SYS_TYPE "
				'SQL="SELECT * FROM BASICCODE WHERE FUNC='GROUPID'  ORDER BY SYS_TYPE "
				SET RST = CONN.EXECUTE(SQL)
				'RESPONSE.WRITE SQL 
				WHILE NOT RST.EOF  
				%>
				<option value="<%=RST("SYS_TYPE")%>" ><%=RST("SYS_TYPE")%><%=RST("SYS_VALUE")%></option>				 
				<%
				RST.MOVENEXT
				WEND 
				%>
				</SELECT>
				<%SET RST=NOTHING %>
		</tr>
		<tr height=30 >	
			<TD nowrap align=right >職等：</TD>			
			<TD >
				<select name=JOB  class=font9  >	
				<option value="">全部 </option>		 
				<%SQL="SELECT * FROM BASICCODE WHERE FUNC='LEV'  ORDER BY SYS_TYPE "
				SET RST = CONN.EXECUTE(SQL)
				WHILE NOT RST.EOF  
				%>
				<option value="<%=RST("SYS_TYPE")%>" ><%=RST("SYS_VALUE")%></option>				 
				<%
				RST.MOVENEXT
				WEND 
				%>		 				 
				</SELECT>
				<%SET RST=NOTHING %>
			</TD>
		</TR>
		<TR  height=30 >
			<td nowrap align=right >員工編號：</td>
			<td colspan=3>
				<input name=empid1 class=inputbox size=15 maxlength=5 onchange=strchg(1)> 
				
			</td>
		</TR>
		<TR  height=30 >
			<td nowrap align=right >班別：</td>
			<td colspan=3>
			 	<select name="shift" class=font9> 			 		
			 		<option value="ALL">常日班</option>
			 		<option value="A">A班</option>
			 		<option value="B">B班</option>
			 	</select>	
			</td>
		</TR>
	</table><BR>	
	<table width=450 align=center>
		<tr >
			<td align=center>
				<input type=button  name=btm class=button value="確   認" onclick="go()" onkeydown="go()">
				<input type=reset  name=btm class=button value="取   消">
				<input type=reset  name=btm class=button value="查   詢" ONCLICK="GETDATA()">
			</td>
		</tr>	
	</table>
</td></tr></table> 
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
 	'<%=self%>.action="<%=SELF%>.FOREGND.asp"
 	<%=self%>.submit
end function   

FUNCTION GETDATA()
	<%=self%>.action="YFYEMPJXA.SCH.asp"
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
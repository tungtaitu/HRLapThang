<%@LANGUAGE=VBSCRIPT CODEPAGE=950%> 
<!-- #include file = "../../GetSQLServerConnection.fun" --> 
<!-- #include file="../../ADOINC.inc" -->
<!-- #include file="../../Include/func.inc" -->
<%
Set conn = GetSQLServerConnection()	  
self="EMPHOLIDAYB"  


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
%>

<html> 
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="refresh">
<link rel="stylesheet" href="../../Include/style.css" type="text/css">
<link rel="stylesheet" href="../../Include/style2.css" type="text/css">
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--
function f()
	<%=self%>.dat1.focus()	
	<%=self%>.dat1.SELECT()
end function    
-->
</SCRIPT>   
</head> 
<body  topmargin="5" leftmargin="5"  marginwidth="0" marginheight="0"  onkeydown="enterto()" onload='f()' >
<form name="<%=self%>" method="post"  >
<table width="460" border="0" cellspacing="0" cellpadding="0">
	<tr><TD>
	<img border="0" src="../../image/icon.gif" align="absmiddle">
	���u�а��@�~���@</TD></tr>
</table>
<hr size=0	style='border: 1px dotted #999999;' align=left width=500>		
<BR><BR>
<table width=500  ><tr><td >
	<table width=400 align=center border=0 cellspacing="0" cellpadding="0"  > 
		<td nowrap align=right >����d��G</td>
			<td colspan=3>
				<input name=dat1 class=inputbox size=15 maxlength=10 onblur="date_change(1)"> ~
				<input name=dat2 class=inputbox size=15 maxlength=10 onblur="date_change(2)">			
			</td> 
		<TR>
		 	<TD nowrap align=right height=30 >���y�G</TD>
			<TD >
				<select name=country  class=font9 style='width:75'  >
					<option value="">���� </option>
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
		<tr>		 
			<TD nowrap align=right height=30 >�t�O�G</TD>
			<TD > 
				<select name=WHSNO  class=font9 >
					<option value="">���� </option>
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
		</TR>		
		<tr>	 
		<TR height=30 >
			<TD nowrap align=right >��/�����G</TD>
			<TD >
				<select name=GROUPID  class=font9  >
				<option value="" selected >���� </option>
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
			<TD nowrap align=right >¾���G</TD>			
			<TD >
				<select name=JOB  class=font9  >	
				<option value="">���� </option>		 
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
			<td nowrap align=right >���u�s���G</td>
			<td colspan=3>
				<input name=empid1 class=inputbox size=15 maxlength=5 onchange=strchg(1)> 
				
			</td>
		</TR>		 
	</table><BR>	
	<table width=450 align=center>
		<tr >
			<td align=center>
				<input type=button  name=btm class=button value="�T   �{" onclick="go()" onkeydown="go()">
				<input type=reset  name=btm class=button value="��   ��">
			</td>
		</tr>	
	</table>	

</td></tr></table> 

</body>
</html>


<script language=vbs>
function dataclick(a)
	if a = 1 then 		
		open "empbasic/empbasic.asp" , "_self"
	elseif a = 2 then 		
		open "empfile/empfile.asp" , "_self"
	elseif a = 3 then 		
		open "empworkHour/empwork.asp" , "_self"	
	elseif a = 4 then 		
		open "holiday/empholiday.asp" , "_self"	
	elseif a = 5 then 		
		open "AcceptCaTime/main.asp" , "_self"				
	elseif a = 6 then 		
		open "../report/main.asp" , "_self"		
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
 	'IF <%=SELF%>.DAT1.VALUE="" AND <%=SELF%>.DAT2.VALUE="" THEN  
 	'	ALERT "������J���"
 	'	<%=SELF%>.DAT1.FOCUS() 
 	'	EXIT function  
 	'END IF 	
 	<%=self%>.action="<%=SELF%>.FORE.asp"
 	<%=self%>.submit() 
end function   
	

'*******�ˬd���*********************************************
FUNCTION date_change(a)	

if a=1 then
	INcardat = Trim(<%=self%>.dat1.value)  		
elseif a=2 then
	INcardat = Trim(<%=self%>.dat2.value)
end if		
		    
IF INcardat<>"" THEN
	ANS=validDate(INcardat)
	IF ANS <> "" THEN
		if a=1 then
			Document.<%=self%>.dat1.value=ANS			
		elseif a=2 then
			Document.<%=self%>.dat2.value=ANS		 			
		end if		
	ELSE
		ALERT "EZ0067:��J������X�k !!" 
		if a=1 then
			Document.<%=self%>.dat1.value=""
			Document.<%=self%>.dat1.focus()
		elseif a=2 then
			Document.<%=self%>.dat2.value=""
			Document.<%=self%>.dat2.focus()
		end if		
		EXIT FUNCTION
	END IF
		 
ELSE
	'alert "EZ0015:�������쥲����J��� !!" 		
	EXIT FUNCTION
END IF 
   
END FUNCTION
</script> 
<%@LANGUAGE=VBSCRIPT CODEPAGE=950%> 
<!-- #include file = "../GetSQLServerConnection.fun" --> 
<!-- #include file="../ADOINC.inc" -->
<!-- #include file="../Include/func.inc" -->
<%
Set conn = GetSQLServerConnection()	  
self="hopdon"
%>

<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="refresh">
<link rel="stylesheet" href="../Include/style.css" type="text/css">
<link rel="stylesheet" href="../Include/style2.css" type="text/css">
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--
function f()
	<%=self%>.empid1.focus()	
end function    
-->
</SCRIPT>   
</head> 
<body  topmargin="5" leftmargin="5"  marginwidth="0" marginheight="0"  onkeydown="enterto()" onload='f()' >
<form name="<%=self%>" method="post" >
<table width="460" border="0" cellspacing="0" cellpadding="0">
	<tr><TD>
	<img border="0" src="../image/icon.gif" align="absmiddle">
	5.列印員工出勤統計表</TD></tr>
</table>
<hr size=0	style='border: 1px dotted #999999;' align=left width=500>		
<BR><BR>
<table width=500  ><tr><td >
	<table width=400 align=center border=0 cellspacing="0" cellpadding="0"  > 
		<TR  height=35 >
			<td nowrap align=right >員工編號：</td>
			<td colspan=3>
				<input name=empid1 class=inputbox size=15 maxlength=5 onchange=strchg(1)>
				
			</td>
		</TR> 		
		<TR>
			<td nowrap align=right >日期範圍：</td>
			<td colspan=3>
				<input name=indat1 class=inputbox size=15 maxlength=10 onblur="date_change(1)" value="2006/08/01"> ~
				<input name=indat2 class=inputbox size=15 maxlength=10 onblur="date_change(2)" value="2006/12/01" >			
			</td>
		</TR>
		<TR  height=35 >
			<td nowrap align=right >類別：</td>
			<td colspan=3>
			 	<select name=empclass class=font9> 
			 		<option value="Y">正式</option>
			 		<option value="N">試用</option>			 		
			 	</select>	
			</td>
		</TR>
	</table><BR>	
	<table width=450 align=center>
		<tr >
			<td align=center>
				<input type=button  name=btm class=button value="確   認" onclick="go()" onkeydown="go()">
				<input type=reset  name=btm class=button value="取   消">
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
 	<%=self%>.action="<%=self%>.getrpt.asp"
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
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

gid=request("groupid")
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
<link rel="stylesheet" href="../../Include/style.css" type="text/css">
<link rel="stylesheet" href="../../Include/style2.css" type="text/css">
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--
function f()
	<%=self%>.YYMM.focus()		
end function 

function dchg()
	<%=self%>.action = "YFYEMPJXA.Fore.asp"
	<%=self%>.submit()
end  function  
   
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
			<TD ><INPUT NAME=YYMM  CLASS=INPUTBOX VALUE="<%=yymm%>" SIZE=10></TD>	
		</TR>
		<tr height=30 >
			<TD nowrap align=right>績效年月：</TD>
			<TD ><INPUT NAME=JXYM  CLASS=INPUTBOX VALUE="<%=jxym%>" SIZE=10></TD>	
		</TR>
		<!--TR>
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
		</tr-->
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
	 		<TD align=right>部門：</TD>
	 			<TD>
	 			<SELECT NAME=GROUPID CLASS=INPUTBOX  onchange="dchg()">
	 				<option value=""></option>
	 				<%SQL="SELECT* FROM BASICCODE WHERE FUNC='GROUPID' AND ( left(sys_type,3)='A06'  or sys_type in ('A059', 'A033') )  "&_	 					  
	 					  "order by  case when sys_type='A065' then 'a000' else sys_type end  "
	 				  SET RST=CONN.EXECUTE(SQL)
	 				  WHILE NOT RST.EOF 
	 				%><OPTION VALUE="<%=RST("SYS_TYPE")%>" <%if rst("sys_type")=gid then%>selected<%end if%>><%=RST("SYS_VALUE")%></OPTION>
	 				<%RST.MOVENEXT%>
	 				<%WEND%>
	 			</SELECT>	 			
	 		</TD> 
		</tr>	
		<tr height=30>			
			<TD align=right>單位組別：</TD>
			<TD>
				<SELECT NAME=zuno CLASS=INPUTBOX  style='width:100'>
					<OPTION VALUE="">不區分</OPTION>
					<%SQL="SELECT* FROM BASICCODE WHERE FUNC='zuno' and left(sys_type,4)='"& gid &"' order by sys_type "
					  SET RST=CONN.EXECUTE(SQL)
					  WHILE NOT RST.EOF 
					%><OPTION VALUE="<%=RST("SYS_TYPE")%>"><%=RST("SYS_TYPE")%>-<%=RST("SYS_VALUE")%></OPTION>
					<%RST.MOVENEXT%>
					<%WEND%>
				</SELECT>	 			
			</TD> 
		</tr>
		<TR  height=30 >
			<td nowrap align=right >班別：</td>
			<td colspan=3>
			 	<select name="shift" class=font9> 			 		
			 		<option value="">不區分</option>
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
	<br>
	<table width=450 class=txt>
		<%
		sql="select   jxym, groupid, isnull(zuno,'') zuno , shift  from VYFYMYJX  where  jxym='"& sgym &"'  group by jxym, groupid, isnull(zuno,'')    , shift  "&_
			"order by groupid, shift "
		'response.write sql	
		set rds=conn.execute(Sql)
		while not rds.eof  	
		%>
		<tr>
			<td align=center>
				資料已處理  <%=rds("groupid")%> -- <%=rds("zuno")%> -- <%=right("   "&rds("shift"),3)%>				
			</td>
		</tr>
		<%rds.movenext
		wend 
		%>
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
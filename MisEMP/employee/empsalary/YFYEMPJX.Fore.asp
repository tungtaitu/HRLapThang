<%@LANGUAGE=VBSCRIPT CODEPAGE=950%> 
<!-- #include file = "../../GetSQLServerConnection.fun" --> 
<!-- #include file="../../ADOINC.inc" -->
<!-- #include file="../../Include/func.inc" -->
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
	<%=self%>.JXYM.focus()		
end function     

function dchg()
	<%=self%>.action = "<%=self%>.Fore.asp"
	<%=self%>.submit()
end  function 
-->
</SCRIPT>   
</head> 
<body  topmargin="5" leftmargin="5"  marginwidth="0" marginheight="0"  onkeydown="enterto()" onload='f()' >
<form name="<%=self%>" method="post" action="EMPFILE.SALARY.ASP">
<input type=hidden name="NNY" value="<%=NNY%>">
<input type=hidden name="NDY" value="<%=NDY%>">
<table width="460" border="0" cellspacing="0" cellpadding="0">
	<tr><TD>
	<img border="0" src="../../image/icon.gif" align="absmiddle">
	�Z�ļ����p��</TD></tr>
</table>
<hr size=0	style='border: 1px dotted #999999;' align=left width=500>
<table width=500 class=txt9>
 	<tr>
 		<td><a href="<%=self%>.Fore.asp">�Z�ķs�W�@�~</a></td> 
 		<td><a href="<%=self%>.ForeEDIT.asp">�Z�ĭק�@�~</A></td>
 		<td><a href="yfyempjx.sch.asp">�Z�Ĭd�ߧ@�~</a></td>
 	</tr>
 	<tr><td colspan=3><hr size=0	style='border: 1px dotted #999999;' align=left ></td></tr>
</table>		
<table width=550 border=0 ><tr><td >
	 	<TABLE WIDTH=500 CLASS=TXT9>
	 		<TR>
	 			<TD>�Z�Ħ~��:</TD>
	 			<TD><INPUT NAME=JXYM VALUE="<%=request("JXYM")%>" SIZE=8 CLASS=INPUTBOX></TD>
				<TD>�p�~�~��:</TD>
	 			<TD><INPUT NAME=SALARYYM VALUE="<%=request("SALARYYM")%>" SIZE=8 CLASS=INPUTBOX></TD>
	 			<TD>�t�O:</TD>
	 			<TD>
	 				<SELECT NAME=jxwhsno CLASS=txt   >
	 					<%if session("rights")="0" then
	 						SQL="SELECT* FROM BASICCODE WHERE FUNC='whsno' order by sys_type "
	 					%>	<option value=""></option>
	 					<%else	
	 						SQL="SELECT* FROM BASICCODE WHERE FUNC='whsno' and sys_type like '"&session("netwhsno")&"' order by sys_type "
	 					  end if
	 					  SET RST=CONN.EXECUTE(SQL)
	 					  WHILE NOT RST.EOF 
	 					%><OPTION VALUE="<%=RST("SYS_TYPE")%>" <%if rst("sys_type")=gid then%>selected<%end if%>><%=RST("SYS_VALUE")%></OPTION>
	 					<%RST.MOVENEXT%>
	 					<%WEND%>
	 				</SELECT>	 			
	 			</TD> 
	 		</tr>
	 		<tr>	
	 			<TD>����:</TD>
	 			<TD>
	 				<SELECT NAME=GROUPID CLASS=INPUTBOX  onchange="dchg()">
	 					<option value=""></option>
	 					<%SQL="SELECT* FROM BASICCODE WHERE FUNC='GROUPID'  order by  case when sys_type='A065' then 'a000' else sys_type end  "
	 					  SET RST=CONN.EXECUTE(SQL)
	 					  WHILE NOT RST.EOF 
	 					%><OPTION VALUE="<%=RST("SYS_TYPE")%>" <%if rst("sys_type")=gid then%>selected<%end if%>><%=RST("SYS_VALUE")%></OPTION>
	 					<%RST.MOVENEXT%>
	 					<%WEND%>
	 				</SELECT>	 			
	 			</TD> 
	 			<TD>���էO:</TD>
	 			<TD>
	 				<SELECT NAME=zuno CLASS=INPUTBOX  style='width:100'>
	 					<OPTION VALUE="">���Ϥ�</OPTION>
	 					<%SQL="SELECT* FROM BASICCODE WHERE FUNC='zuno' and left(sys_type,4)='"& gid &"' "
	 					  SET RST=CONN.EXECUTE(SQL)
	 					  WHILE NOT RST.EOF 
	 					%><OPTION VALUE="<%=RST("SYS_TYPE")%>"><%=RST("SYS_VALUE")%></OPTION>
	 					<%RST.MOVENEXT%>
	 					<%WEND%>
	 				</SELECT>	 			
	 			</TD>	 			
	 			<TD>�Z�O:</TD>
	 			<TD><SELECT NAME=SHIFT CLASS=INPUTBOX>
	 				<OPTION VALUE="">���Ϥ�</OPTION>
	 				<OPTION VALUE="ALL">��</OPTION>
	 				<OPTION VALUE="A">A�Z</OPTION>
	 				<OPTION VALUE="B">B�Z</OPTION>
	 				</SELECT>	 			
	 			</TD>
	 			
	 		</TR>
	 	</TABLE>		
	 	<hr size=0	style='border: 1px dotted #999999;' align=left width=500> 
	 	<TABLE WIDTH=450 CLASS=TXT9 BGCOLOR="#CCCCCC" BORDER=0 border="1" cellspacing="1" CLASS=TXT9 ALIGN=CENTER>
	 		<TR BGCOLOR="#FFF278">
	 			<TD WIDTH=60 HEIGHT=22 ALIGN=CENTER>STT</TD>
	 			<TD WIDTH=150 ALIGN=CENTER>����</TD>
	 			<TD WIDTH=100 ALIGN=CENTER>���Z</TD>
	 			<TD WIDTH=70 ALIGN=CENTER>�Y��</TD>
	 			<TD WIDTH=70 ALIGN=CENTER>���</TD>	 			
	 		</TR>
	 		<%FOR I = 1 TO 5 %>
	 		<TR BGCOLOR="#FFFFFF" >
	 			<TD HEIGHT=22 ALIGN=CENTER ><INPUT NAME="STT" VALUE="<%=CHR(64+I)%>" SIZE=5 CLASS="readonly2" READONLY ></TD>
	 			<TD ALIGN=CENTER>
	 				<INPUT NAME=DESCP VALUE="" CLASS=INPUTBOX SIZE=20>
	 			</TD>
	 			<TD ALIGN=CENTER>
	 				<INPUT NAME=HXSL VALUE="" CLASS=INPUTBOX SIZE=12>
	 			</TD>	 			
	 			<TD ALIGN=CENTER><INPUT NAME="HESO" VALUE="" SIZE=10 CLASS="INPUTBOX"  ></TD>
	 			<TD ALIGN=CENTER><INPUT NAME="PER" VALUE="" SIZE=10 CLASS="INPUTBOX"  ></TD>
	 		</TR>
	 		<%NEXT %>
	 	</TABLE>	<br>
	 	<table width=450 align=center>
		<tr >
			<td align=center>
				<input type=button  name=btm class=button value="�T   �{" onclick="go()" onkeydown="go()">
				<input type=reset  name=btm class=button value="��   ��">
				<input type=reset  name=btm class=button value="�ƻs���" ONCLICK=COPYDATA()>
			</td>
		</tr>	
	</table>	
</td></tr></table> 

</body>
</html>


<script language=vbs>  

'*******�ˬd���*********************************************
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
		ALERT "EZ0067:��J������X�k !!" 
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
	'alert "EZ0015:�������쥲����J��� !!" 		
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
		ALERT "�п�J�Z�Ħ~��!!"
		<%=SELF%>.JXYM.FOCUS()
		EXIT FUNCTION 
	ELSEIF <%=SELF%>.SALARYYM.VALUE="" THEN 	
		ALERT "�п�J�p�~�~��!!"
		<%=SELF%>.SALARYYM.FOCUS()
		EXIT FUNCTION 
	ELSEIF <%=SELF%>.GROUPID.VALUE="" THEN 
		ALERT "�п�J���!!"
		<%=SELF%>.GROUPID.FOCUS()
		EXIT FUNCTION 
	ELSEIF <%=SELF%>.jxwhsno.VALUE="" THEN 
		ALERT "�п�J�t�O!!"
		<%=SELF%>.jxwhsno.FOCUS()
		EXIT FUNCTION
	ELSE
		<%=SELF%>.ACTION="<%=SELF%>.UPD.ASP"
		<%=SELF%>.SUBMIT()		
	END IF 
END FUNCTION  

FUNCTION COPYDATA()
	'IF <%=SELF%>.JXYM.VALUE="" THEN 
	'	ALERT "�п�J���ƻs���Z�Ħ~��"
	'	<%=SELF%>.JXYM.FOCUS()
	'	EXIT FUNCTION 
	'END IF 
	<%=SELF%>.ACTION="<%=SELF%>.CopyData.ASP"
	<%=SELF%>.SUBMIT()		
END FUNCTION 
</script> 
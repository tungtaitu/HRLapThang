<%@LANGUAGE=VBSCRIPT CODEPAGE=950%> 
<!-- #include file = ".../GetSQLServerConnection.fun" --> 
<!-- #include file=".../ADOINC.inc" -->
<!-- #include file=".../Include/func.inc" -->
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

JXYM=REQUEST("JXYM")
GROUPID = REQUEST("GROUPID")
SHIFT=REQUEST("SHIFT") 

TotalPage = 1
PageRec = 5    'number of records per page
TableRec = 10    'number of fields per record   

SQL="SELECT* FROM YFYMJIXO WHERE  JXYM LIKE '"& JXYM &"%' AND GROUPID LIKE '"& GROUPID &"%' AND SHIFT = '"&SHIFT&"'  ORDER BY STT" 
Set rs = Server.CreateObject("ADODB.Recordset")     
RS.OPEN SQL,CONN, 3, 3   

Redim tmpRec(TotalPage, PageRec, TableRec)   'Array  

IF NOT RS.EOF THEN 	
 	for i = 1 to TotalPage 
		for j = 1 to PageRec
			if not rs.EOF then 				
				tmpRec(i, j, 0) = "no"
				tmpRec(i, j, 1) = trim(rs("STT"))
				tmpRec(i, j, 2) = trim(rs("DESCP"))
				tmpRec(i, j, 3) = trim(rs("HXSL"))
				tmpRec(i, j, 4) = rs("PER")				
				tmpRec(i, j, 5) = rs("HESO")
				rs.MoveNext 
			else 
				exit for 
			end if 
		 next
	NEXT
END IF 		 

%>

<html> 
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="refresh">
<link rel="stylesheet" href=".../Include/style.css" type="text/css">
<link rel="stylesheet" href=".../Include/style2.css" type="text/css">
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--
function f()
	<%=self%>.JXYM.focus()		
end function    
-->
</SCRIPT>   
</head> 
<body  topmargin="5" leftmargin="5"  marginwidth="0" marginheight="0"  onkeydown="enterto()" onload='f()' >
<form name="<%=self%>" method="post" action="EMPFILE.SALARY.ASP">
<INPUT TYPE=hidden NAME=TotalPage VALUE="<%=TotalPage%>">
<INPUT TYPE=hidden NAME=CurrentPage VALUE="<%=CurrentPage%>">
<INPUT TYPE=hidden NAME=PageRec VALUE="<%=PageRec%>">
<table width="460" border="0" cellspacing="0" cellpadding="0">
	<tr><TD>
	<img border="0" src=".../image/icon.gif" align="absmiddle">
	����Z��</TD></tr>
</table>
<hr size=0	style='border: 1px dotted #999999;' align=left width=500>
<table width=500 class=txt9>
 	<tr>
 		<td><a href="<%=self%>.Fore.asp">�Z�ķs�W�@�~</a></td> 
 		<td>�Z�ĭק�@�~</td>
 		<td>�Z�Ĭd�ߧ@�~</td>
 	</tr>
 	<tr><td colspan=3><hr size=0	style='border: 1px dotted #999999;' align=left ></td></tr>
</table>		
<table width=550 border=0 ><tr><td >
	 	<TABLE WIDTH=500 CLASS=TXT9>
	 		<TR>
	 			<TD>�Z�Ħ~��:</TD>
	 			<TD><INPUT NAME=JXYM VALUE="<%=JXYM%>" SIZE=8 CLASS=INPUTBOX></TD>
				<TD>�p�~�~��:</TD>
	 			<TD><INPUT NAME=SALARYYM VALUE="" SIZE=8 CLASS=INPUTBOX></TD>
	 			<TD>���:</TD>
	 			<TD>
	 				<SELECT NAME=GROUPID CLASS=INPUTBOX >
	 					<%SQL="SELECT* FROM BASICCODE WHERE FUNC='GROUPID' AND  left(sys_type,3)='A06'  or sys_type in ('A059', 'A033')   ORDER BY SYS_TYPE "
	 					  SET RST=CONN.EXECUTE(SQL)
	 					  WHILE NOT RST.EOF 
	 					%><OPTION VALUE="<%=RST("SYS_TYPE")%>" <%IF GROUPID=RST("SYS_TYPE") THEN %>SELECTED<%END IF%>><%=RST("SYS_VALUE")%></OPTION>
	 					<%RST.MOVENEXT%>
	 					<%WEND%>
	 				</SELECT>	 			
	 			</TD>
	 			<TD>�Z�O:</TD>
	 			<TD><SELECT NAME=SHIFT CLASS=INPUTBOX>
	 				<OPTION VALUE="ALL" <%IF SHIFT="ALL" THEN %>SELECTED<%END IF%> >�`��Z</OPTION>
	 				<OPTION VALUE="A" <%IF SHIFT="A" THEN %>SELECTED<%END IF%>>A�Z</OPTION>
	 				<OPTION VALUE="B" <%IF SHIFT="B" THEN %>SELECTED<%END IF%>>B�Z</OPTION>
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
	 			<TD WIDTH=70 ALIGN=CENTER>���</TD>
	 			<TD WIDTH=70 ALIGN=CENTER>�Y��</TD>
	 		</TR>
	 		<%FOR I = 1 TO PAGEREC 
	 			IF tmpRec(1,I,1)<>"" THEN 
	 				STT=tmpRec(1,I,1) 
	 			ELSE
	 				STT=CHR(64+I)
	 			END IF	
	 		%>
	 		<TR BGCOLOR="#FFFFFF" >
	 			<TD HEIGHT=22 ALIGN=CENTER ><INPUT NAME="STT" VALUE="<%=STT%>" SIZE=5 CLASS="readonly2" READONLY ></TD>
	 			<TD ALIGN=CENTER>
	 				<INPUT NAME=DESCP VALUE="<%=tmpRec(1,I,2)%>" CLASS=INPUTBOX SIZE=20>
	 			</TD>
	 			<TD ALIGN=CENTER>
	 				<INPUT NAME=HXSL VALUE="<%=tmpRec(1,I,3)%>" CLASS=INPUTBOX SIZE=12>
	 			</TD>
	 			<TD ALIGN=CENTER><INPUT NAME="PER" VALUE="<%=tmpRec(1,I,4)%>" SIZE=10 CLASS="INPUTBOX"  ></TD>
	 			<TD ALIGN=CENTER><INPUT NAME="HESO" VALUE="<%=tmpRec(1,I,5)%>" SIZE=10 CLASS="INPUTBOX"  ></TD>
	 		</TR>
	 		<%NEXT %>
	 	</TABLE>	<br>
	 	<table width=450 align=center>
		<tr >
			<td align=center>
				<input type=button  name=btm class=button value="�T   �{" onclick="go()" onkeydown="go()">
				<input type=reset  name=btm class=button value="��   ��" >				
				<input type=reset  name=btm class=button value="�ƻs���" ONCLICK=COPYDATA()>
			</td>
		</tr>	
	</table>	
</td></tr></table> 

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
	if <%=self%>.sgno.value="" then 
		alert "������J�ƬG�渹"
		<%=self%>.sgno.focus()
		exit function 
	end if 	
	if <%=self%>.pddate.value="" then 
		alert "�п�J�P�w���!!"
		<%=self%>.pddate.focus()
		exit function 
	end if 
	if <%=self%>.cfGroup.value="A" then 
		if <%=self%>.empid.value="" or <%=self%>.cfdw.value="" then 		
			alert "������J�u���γd����H(���u�m�W)!!"
			if <%=self%>.empid.value="" then 
				<%=self%>.empid.focus()
			else
				<%=self%>.cfdw.focus()
			end if 
			exit function 
		end if		
	elseif <%=self%>.cfGroup.value="B" then 
		if <%=self%>.cfdw.value="" then
			alert "������J�d����H(�q������)!!"
			<%=self%>.cfdw.focus()
			exit function 
		end if 
	else 
		if <%=self%>.cfdw.value="" then
			alert "������J�d����H(�m�W�μt�ӦW��)!!"
			<%=self%>.cfdw.focus()
			exit function 
		end if 
	end if 	  		
	if <%=self%>.sgcost.value="" then 
		alert "�п�J���ڪ��B!!"
		<%=self%>.sgcost.focus()
		exit function 
	end if 
	
 	<%=self%>.action="vyfysuco.Upd.asp"
 	<%=self%>.submit() 
end function   
	

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
	ELSEIF <%=SELF%>.SHIFT.VALUE="" THEN 
		ALERT "�п�J�Z�O!!"
		<%=SELF%>.SHIFT.FOCUS()
		EXIT FUNCTION
	ELSE
		<%=SELF%>.ACTION="<%=SELF%>.UPD.ASP"
		<%=SELF%>.SUBMIT()		
	END IF 
END FUNCTION 


FUNCTION COPYDATA()
	IF <%=SELF%>.JXYM.VALUE="" THEN 
		ALERT "�п�J���ƻs���Z�Ħ~��"
		<%=SELF%>.JXYM.FOCUS()
		EXIT FUNCTION 
	END IF 	
	<%=SELF%>.ACTION="<%=SELF%>.NEW.ASP"
	<%=SELF%>.SUBMIT()		
END FUNCTION   
</script> 
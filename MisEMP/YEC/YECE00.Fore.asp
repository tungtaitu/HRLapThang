<%@LANGUAGE=VBSCRIPT CODEPAGE=950%> 
<!-- #include file = "../GetSQLServerConnection.fun" --> 
<!-- #include file="../ADOINC.inc" -->
<!-- #include file="../Include/func.inc" -->
<!--#include file="../include/checkpower.asp"-->  
<%
Set conn = GetSQLServerConnection()	  
self="yece00"  
Set rs = Server.CreateObject("ADODB.Recordset")  

nowmonth = year(date())&right("00"&month(date()),2)  
if month(date())="01" then  
	calcmonth = year(date())-1&"12"  	
else
	calcmonth =  year(date())&right("00"&month(date())-1,2)   
end if 	   

if day(date())<="02" then 
	if month(date())="01" then  
		calcmonth = year(date())-1&"12" 
	else
		calcmonth =  year(date())&right("00"&month(date())-1,2)   
	end if 	 	
else
	calcmonth = nowmonth 
end if 

empid = request("empid")
if request("empid")<>"" then 
	sql="select * from view_empfile where empid='"& empid &"'" 
	rs.open sql, conn, 1, 3 
	if not rs.eof then 
		photos = empid&".jpg"
		whsno = rs("whsno")
		country = rs("country")
		indat = rs("indat")
		gstr = rs("gstr")
		groupid = rs("groupid")
		empnam=rs("empnam_cn")&rs("empnam_vn")
	else
		empid=""		
	end if 
	set rs=nothing 
end if 	

'photos = "lf013.jpg"
'response.write photos 
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
	if trim(<%=self%>.empid.value)=""  then 
		<%=self%>.YYMM.focus()	
		<%=self%>.YYMM.SELECT()
	'else
	'	<%=self%>.empid.focus()	
	'	<%=self%>.empid.SELECT()
	end if 
	if <%=self%>.country.value="" then 
		tb1.style.visibility="hidden"  
		tb2.style.visibility="hidden"  
	elseif <%=self%>.country.value="VN" then 
		tb1.style.visibility="hidden"  
	else	
		tb2.style.visibility="hidden"  
	end if	
end function    

function empidchg()
	if trim(<%=self%>.empid.value)<>""  then 
		<%=self%>.action = "<%=self%>.Fore.asp"
		<%=self%>.submit()
	end if 	
end function 

-->
</SCRIPT>   
</head> 
<body  topmargin="5" leftmargin="5"  marginwidth="0" marginheight="0"  onkeydown="enterto()" onload='f()' >
<form name="<%=self%>" method="post" > 
<input name="country" type="hidden" value="<%=country%>">
<table width="460" border="0" cellspacing="0" cellpadding="0">
	<tr><TD>
	<img border="0" src="../image/icon.gif" align="absmiddle">
	<%=session("pgName")%></TD></tr>
</table>
<hr size=0	style='border: 1px dotted #999999;' align=left width=500>		
<table width=500  ><tr><td >
	<table width=400 align=center border=0 cellspacing="2" cellpadding="2"  > 
		<tr height=30 >
			<TD nowrap align=right>�~��<br><font class="txt8">Thang Nam</font></TD>
			<TD ><INPUT NAME=yymm CLASS=INPUTBOX VALUE="<%=calcmonth%>" SIZE=8></TD>	
			<TD nowrap align=right>�u��<br><font class="txt8">so the</font></TD>
			<TD  >
				<INPUT NAME=empid CLASS=INPUTBOX VALUE="<%=ucase(empid)%>" SIZE=10 onblur="empidchg()">
				<INPUT NAME=empNam CLASS=readonly  VALUE="<%=empnam%>" SIZE=20>
			</TD>	
		</TR>
		<tr>
			<td colspan=4>
				<table width="100%"  border="0" cellspacing="1" cellpadding="1" class="txt"> 
					<tr>
						<Td rowspan=5 width=150 align=center><img src="../yeb/pic/<%=photos%>" width=100 height=120 border=0  ></td>						
					</tr>
					<Tr>
						<td>���y</td>
						<td><%=country%></td>						
					</tr>	
					<Tr>
						<td>��¾��</td>
						<td><%=indat%></td>						
					</tr>						
					<Tr>
						<td>�t�O</td>
						<td><%=whsno%></td>						
					</tr>						
					<Tr>
						<td>���</td>
						<td><%=groupid%>-<%=gstr%></td>						
					</tr>			
			</td>		
		</tr>
  <TABLE>
	<hr size=0	style='border: 1px dotted #999999;' align=left width=500>		
	<table width=500 border=0 cellspacing="2" cellpadding="2" class=txt>
			<tr>
				<td align=right>¾��<td>
				<td>
					<select name=job class=txt8 style="width:80">
						<option value=""></option>
						<%sql="select * from basicCode where func='lev' order by sys_type" 
						  set rsx=conn.execute(Sql)
							while not rsx.eof 
						%>
						<option value="<%=rsx("sys_type")%>"><%=rsx("sys_type")%>-<%=rsx("sys_value")%></option>
						<%rsx.movenext
						wend 
						set rsx=nothing 
						%>
					</select>
				<td>
				<td align="right">���~</td>
				<td  ><input name=salary_m size=10 class="inputbox8" ></td>
				<td align="right">�~�~</td>
				<td  ><input name=salary_y size=8 class="inputbox8" ></td>
				<td align="right">���O</td>
				<td >
						<select name=dm_sy class=txt8 >
						<option value=""></option>
						<option value="RMB">VND</option>
						<option value="NTD">NTD</option>
						<option value="USD">USD</option>
						<option value="RMB">RMB</option>						
						</select>
				</td>
				<td align="right">Rate</td>
				<td  ><input name=rate size=5 class="inputbox8" ></td>
			</tr>
	</table> 
	<table width=500 border=0 cellspacing="1" cellpadding="1" class=txt id="tb1">
		<tr bgcolor="lightyellow" height=22>	
			<td colspan=4 align="center">�Ҥ�</td>
			<td colspan=2 align="center">�ҥ~</td>
		</tr>
		<tr bgcolor="#e4e4e4">
			<td>���~ BB</td>
			<td>¾�� CV</td>
			<td>�޳N KT</td>
			<td>��[ </td>
			<td>¾�[ </td>
			<td>���~�z�K </td>
		</tr>
	</table>
	<table width=500 border=0 cellspacing="1" cellpadding="1" class=txt id="tb2">
		<tr bgcolor="#e4e4e4">
			<td>���~ BB</td>
			<td>¾�� CV</td>
			<td>�ɧU PHU</td>
			<td>�޳N KT</td>
			<td>�y�� NN</td>
			<td>���� MT</td>
			<td>��[ TTKH</td>
			<td>���� </td>
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
 	<%=self%>.action="EMPSALARY01.ForeGnd.asp"
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
</script> 
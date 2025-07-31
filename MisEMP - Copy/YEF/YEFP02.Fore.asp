<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<!-- #include file = "../GetSQLServerConnection.fun" -->
<!-- #include file="../ADOINC.inc" -->
<!-- #include file="../Include/func.inc" -->
<!--#include file="../include/sideinfo.inc"-->
<%
Set conn = GetSQLServerConnection()
self="YEFP02" 


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
	<%=self%>.yymm.select()
end function
function enterto()
	if window.event.keyCode = 13 then window.event.keyCode =9
end function
-->
</SCRIPT>
</head>
<body  topmargin="5" leftmargin="5"  marginwidth="0" marginheight="0"  onkeydown="enterto()" onload='f()' >
<form name="<%=self%>" method="post" action="emp_basicbiao.getrpt.asp">
<table border=0 style="height:70px"><tr><td>&nbsp;</td></tr></table>

<table width="100%"  ><tr><td align="center">
	<table  id="myTableForm" style="width:70%">
		<tr>
			<tr><td colspan=4>&nbsp;</td></tr>
		 	<TD align=right>統計年月<BR><font class=txt8>Thong ke Thang Nam</font></TD>
			<TD><input type="text" name="yymm" value="<%=nowmonth%>" ></TD>
			<TD align=right>廠別<BR><font class=txt8>Loai Xuong</font></TD>
			<TD>
				<select name=WHSNO style="width:200px">
					<option value="">全部(Toan bo) </option>
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
		</tr>
		<tr>
		 	<TD  align=right valign=top>國籍<BR><font class=txt8>Quoc Tich</font></TD>
			<TD  valign=top>
				<select name=country   style="font-size:9pt;width:150" MULTIPLE size=6  >
					<option value="" selected >全部(Toan bo) </option>
					<%SQL="SELECT * FROM BASICCODE WHERE FUNC='country'  ORDER BY SYS_type desc  "
					SET RST = CONN.EXECUTE(SQL)
					WHILE NOT RST.EOF
					%>
					<option value="<%=RST("SYS_TYPE")%>" ><%=RST("SYS_TYPE")%>-<%=RST("SYS_VALUE")%></option>
					<%
					RST.MOVENEXT
					WEND
					%>					
				</SELECT>
				<%SET RST=NOTHING %>
			</TD>
			<TD  align=right valign=top >部門<BR><font class=txt8>Bo Phan</font></TD>
			<TD  valign=top>
				<select name=GROUPID  style="font-size:9pt;width:98%" MULTIPLE size=6   >
				<option value="" selected >全部(Toan bo) </option>
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
			</td>			
		</tr>  
		<TR>
			<td nowrap align=right >員工編號<BR><font class=txt8>So The</font></td>
			<td >
				<input name=empid1  size=7 maxlength=5 onchange=strchg(1)> 
				<input type=hidden name=empid2  size=7 maxlength=5 onchange=strchg(2)>
			</td>			
			<TD  align=right >單位<BR><font class=txt8>Don vi</font></TD>
			<TD >
				<select name=zuno    >
				<option value="" selected >全部(Toan bo) </option>
				<%
				SQL="SELECT * FROM BASICCODE WHERE FUNC='zuno' and sys_type <>'XX' ORDER BY SYS_TYPE "
				'SQL="SELECT * FROM BASICCODE WHERE FUNC='GROUPID'  ORDER BY SYS_TYPE "
				SET RST = CONN.EXECUTE(SQL)
				'RESPONSE.WRITE SQL
				WHILE NOT RST.EOF
				%>
				<option value="<%=RST("SYS_TYPE")%>" ><%=RST("SYS_TYPE")%>-<%=RST("SYS_VALUE")%></option>
				<%
				RST.MOVENEXT
				WEND
				%>
				</SELECT>
				<%SET RST=NOTHING  
				conn.close
				set conn=nothing
				
				%>
			</td>			
		</tr> 
		<TR>
			<td nowrap align=right >到職日期<BR><font class=txt8>NVX</font></td>
			<td><input type="text" style="width:45%" name="indat1" size=15 maxlength=10 onblur="date_change(1)">~
				<input type="text" style="width:45%" name="indat2" size=15 maxlength=10 onblur="date_change(2)">							
			</td>
			<td nowrap align=right >離職日期<BR><font class=txt8>NTC</font></td>
			<td>
				<input type="text" style="width:45%" name=otd1  size=15 maxlength=10 onblur="date_change(5)">~ 
				<input type="text" style="width:45%" name=otd2  size=15 maxlength=10 onblur="date_change(6)">
			</td>
		</TR>		
		<TR>
			<td nowrap align=right >簽約日期<BR><font class=txt8>Ngay ky hop dong</font></td>
			<td >
				<input type="text" style="width:45%" name=bhdat1  size=15 maxlength=10 onblur="date_change(3)">~ 
				<input type="text" style="width:45%" name=bhdat2  size=15 maxlength=10 onblur="date_change(4)">
			</td>
			<td nowrap align=right >員工統計<br><font class=txt8>Thong ke Nhan vien</font></td>
			<td>
			 	<select name=outemp >
			 		<option value="Y">在職(Tai Chuc)</option>
			 		<option value="">全部(Toan bo)</option>
			 		<option value="N">已離職(Thoi viec)</option>
			 	</select>
			</td>
		</TR>
		<TR>
			<td nowrap align=right >排列方式<BR>Sap xep</td>
			<td>
			 	<select name=orderby >
			 		<option value="">全部(Toan bo)</option>
			 		<option value="1">依部門(Theo bo phan)</option>
			 		<option value="2">依工號(thoe So the)</option>
			 	</select>
			</td>
			<td nowrap align=right >顯示方式<BR></td>
			<td >
			 	<select name=empTJ >
			 		<option value="">全部(Toan bo)</option>
			 		<option value="1">不顯示明細</option>			 		
			 	</select>
			</td>
		</TR>	
		<tr height=30>
			<td align=center colspan=4 >				
				<input type=button  name=btm class="btn btn-sm btn-danger" value="(Y)確　認" onclick="go()" onkeydown="go()">
				<input type=reset  name=btm class="btn btn-sm btn-outline-secondary" value="(N)取　消">
			</td>
		</tr>
		<tr height=4 >	
			<td align=center colspan=4><input type=button  name=btm class="btn btn-sm btn-outline-secondary" value="SaveTo EXCEL" onclick=goexcel() style='background-color:#e4e4e4'>
			<input type="hidden"  name="btmxls" class="btn btn-sm btn-outline-secondary" value="To EXCEL" onclick="toxls()" style='background-color:#90EE90' >
			</td>
		</tr>
	</table> 	

</td></tr></table>
<script type="text/javascript">	
	function go2(){
	<%=self%>.empid2.value = <%=self%>.empid1.value;
 	<%=self%>.action="<%=session("rpt")%>"+"yef/"+"<%=self%>.getrpt.asp" ;	
 	<%=self%>.submit();
	}

	function toxls(){
	
		alert ("aaa");
		var f = document.forms[0];
	  var url="<%=session("rpt")%>"+"netxls/yefp02.aspx" ;
		parent.best.cols="50%,50%" ;
		//parent.Back.location = url ;
		f.method ="get";
		f.action = url 		 ; 
		f.target ="Back" ;
		f.submit();
		
		
	 //window.open  ( url + "?code1="+f.WHSNO.value ,"_new" ,"top=150,left=150,width=300,height=300,resizable=yes") ;
	}
	
</script>
</body>
</html>

<script language=vbs> 
function goexcel()
	'open "<%=self%>.toexcel.asp" , "Back" 
	<%=self%>.action="<%=self%>.toexcel.asp"
	<%=self%>.target="Back"
	<%=self%>.submit()
	parent.best.cols="100%,0%"
end function  
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
	<%=self%>.empid2.value = <%=self%>.empid1.value
 	<%=self%>.action="<%=session("rpt")%>"&"yef/"&"<%=self%>.getrpt.asp" 
 	<%=self%>.submit()
end function


'*******檢查日期*********************************************
FUNCTION date_change(a)

if a=1 then
	INcardat = Trim(<%=self%>.indat1.value)
elseif a=2 then
	INcardat = Trim(<%=self%>.indat2.value)
elseif a=3 then
	INcardat = Trim(<%=self%>.bhdat1.value)
elseif a=4 then
	INcardat = Trim(<%=self%>.bhdat2.value)
elseif a=5 then
	INcardat = Trim(<%=self%>.otd1.value)
elseif a=6 then
	INcardat = Trim(<%=self%>.otd2.value)	
end if

IF INcardat<>"" THEN
	ANS=validDate(INcardat)
	IF ANS <> "" THEN
		if a=1 then
			Document.<%=self%>.indat1.value=ANS
		elseif a=2 then
			Document.<%=self%>.indat2.value=ANS
		elseif a=3 then
			Document.<%=self%>.bhdat1.value=ANS
		elseif a=4 then
			Document.<%=self%>.bhdat2.value=ANS
		elseif a=5 then
			Document.<%=self%>.otd1.value=ANS
		elseif a=6 then
			Document.<%=self%>.otd2.value=ANS			
		end if
	ELSE
		ALERT "EZ0067:輸入日期不合法 !!"
		if a=1 then
			Document.<%=self%>.indat1.value=""
			Document.<%=self%>.indat1.focus()
		elseif a=2 then
			Document.<%=self%>.indat2.value=""
			Document.<%=self%>.indat2.focus()
		elseif a=3 then
			Document.<%=self%>.bhdat1.value=""
			Document.<%=self%>.bhdat1.focus()
		elseif a=4 then
			Document.<%=self%>.bhdat2.value=""
			Document.<%=self%>.bhdat2.focus()
		elseif a=5 then
			Document.<%=self%>.otd1.value=""
			Document.<%=self%>.otd1.focus()
		elseif a=6 then
			Document.<%=self%>.otd2.value=""
			Document.<%=self%>.otd2.focus()			
		end if
		EXIT FUNCTION
	END IF

ELSE
	'alert "EZ0015:日期該欄位必須輸入資料 !!"
	EXIT FUNCTION
END IF

END FUNCTION
</script> 
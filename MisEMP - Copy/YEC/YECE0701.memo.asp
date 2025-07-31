<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<!-- #include file = "../GetSQLServerConnection.fun" --> 
<!-- #include file="../ADOINC.inc" --> 
<%
thisfrm = "YECE0701" 

ftype = request("ftype") 
 
index=request("index")  
CurrentPage = request("CurrentPage") 
yymm = request("yymm") 

tmpRec = Session("empBHGTD") 
 
Set conn = GetSQLServerConnection()	 
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="refresh"> 
<link rel="stylesheet" href="../../Include/style.css" type="text/css">
<link rel="stylesheet" href="../../Include/style2.css" type="text/css">
</head>
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
 
function f()
	<%=thisfrm%>.memo.focus()
end function
</script> 

<body  topmargin="15" leftmargin="5"  marginwidth="0" marginheight="0" onload=f()  >
<form name=<%=thisfrm%> method='post' >
<input type=hidden name=yymm value="<%=yymm%>" > 
<input type=hidden name=index value="<%=index%>" >
<input type=hidden name=CurrentPage value="<%=CurrentPage%>" > 
	<table width=400 class=txt cellspacing="1" cellpadding="1" BGCOLOR="LightGrey" align=center>
		<tr bgcolor=lightyellow>
			<td width=70 align=right>工號Ma So: </td>
			<td width=50><%=tmpRec(CurrentPage,index + 1,1)%> </td>
			<td width=80  align=right>姓名Ho Ten:</td>
			<td><%=tmpRec(CurrentPage,index + 1,2)%>&nbsp;<%=tmpRec(CurrentPage,index + 1,3)%> </td>
		</tr>  	
	</table>
 
	<table width=400 class=txt  cellspacing="1" cellpadding="1"  align=center>
		<tr>
			<td><font color=blue>＊ <font color=red><b><%=yymm%></b></font> Chi Chu:</font></td>
		</tr>
		<tr>
			<td>
			<TEXTAREA rows=7 cols=75 name=memo class="INPUTBOX" STYLE='HEIGHT:AUTO' wrap="PHYSICAL"><%=replace(tmprec(currentpage, index+1, 31),"<br>",vbcrlf)%></TEXTAREA>
			</td>
		</tr>
	</table>
	<table width=400 class=txt  cellspacing="1" cellpadding="1"  align=center>
		<tr>
			<td align=center>				
				<input type=button name=send  value="Y 確認後關閉" onclick="go()" class=button >
				<input type=button name=send  value="關閉(Close)" onclick="window.close()" class=button >
			</td>
		</tr>
	</table>
</form> 
<!-- #include file="showholiday2.asp" --> 
</body>
</html>
<script language=vbs>
function go()
	yymmstr=<%=thisfrm%>.yymm.value 
 	memostr = escape(<%=thisfrm%>.memo.value) 
 	index = <%=thisfrm%>.index.value 
 	cp = <%=thisfrm%>.CurrentPage.value 
	open "<%=thisfrm%>.memoupd.asp?ftype=memochk&index="&index &"&CurrentPage="& cp & _ 
 		 "&yymm="& yymmstr &_
 		 "&code=" & memostr  , "_self"  	 		
end function    
 
	
</script>
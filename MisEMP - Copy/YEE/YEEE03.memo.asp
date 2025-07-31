<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<!-- #include file = "../GetSQLServerConnection.fun" --> 
<!-- #include file="../ADOINC.inc" --> 
<%
SELF = "YEEE03" 

ftype = request("ftype") 
 
index=request("index")  
CurrentPage = request("CurrentPage") 
yymm = request("yymm")
 

tmpRec = Session("YEEE03B") 
 
Set conn = GetSQLServerConnection()	 

sqlx="select * from bemps where empid='"&tmpRec(CurrentPage,index + 1,1)&"' and "&_
		 "yymm between '"& left( replace(tmpRec(CurrentPage, index+1, 18),"/","") ,6) &"' and '"& left( replace(tmpRec(CurrentPage,index+1, 19),"/","") ,6) &"' "&_
		 "order by yymm desc " 
set rsx=conn.execute(Sqlx)

%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="refresh"> 
<link rel="stylesheet" href="../Include/style.css" type="text/css">
<link rel="stylesheet" href="../Include/style2.css" type="text/css">
</head>
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
 
function f()
	<%=self%>.memo.focus()
end function
</script> 

<body  topmargin="15" leftmargin="5"  marginwidth="0" marginheight="0" onload=f()  >
<form name=<%=self%> method='post' >
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
		<tr bgcolor=lightyellow>
			<td  align=right>到職日: </td>
			<td  ><%=tmpRec(CurrentPage,index + 1,5)%> </td>
			<td align=right>離職日:</td>
			<%if tmpRec(CurrentPage,index + 1,6)="" then 
				C_enddat = date()
			  else	
			  c_enddat = tmpRec(CurrentPage,index + 1,6)
			  end if 	
			%>			
			<td><%=tmpRec(CurrentPage,index + 1,6)%>&nbsp;年資:<%=datediff("m", cdate(tmpRec(CurrentPage,index + 1,5)),cdate(c_enddat)) %>M </td> 
			
		</tr>  			
	</table>
 
	<table width=400 class=txt  cellspacing="1" cellpadding="1"  align=center> 
 
		<tr>
			<td><font color=blue>＊<font color=red><b><%=yymm%></b></font> 薪資結帳說明或備註說明:</font></td>
		</tr> 		
		<tr>
			<td>
			<TEXTAREA rows=7 cols=75 name=memo class="INPUTBOX" STYLE='HEIGHT:AUTO' wrap="PHYSICAL"><%=replace(tmprec(currentpage, index+1, 24),"<br>",vbcrlf)%></TEXTAREA>
			</td>
		</tr>
	</table>
	<table width=400 class=txt  cellspacing="1" cellpadding="1"  align=center>
		<tr>
			<td align=center>				
				<input type=button name=send  value="Y 確認後關閉" onclick="go()" class=button >
				<input type=button name=send  value="關閉(Close)" onclick="go()" class=button >
			</td>
		</tr>
	</table>	
	<table width=400 class=txt8  cellspacing="1" cellpadding="1"  align=center border=1>
		<tR bgcolor="#e4e4e4">
			<td width=100 nowrap align="center">YYMM</td>
			<td width=100 nowrap align="center">BB</td>
			<td width=100 nowrap align="center">CV</td>
			<td width=100 nowrap align="center">PHU</td>
		</TR>
		<%if not rsx.eof then %>
		<%while not rsx.eof 
			ii = ii + 1 
			A = A + cdbl(rsx("BB"))
			B = B + cdbl(rsx("CV"))
			C = C + cdbl(rsx("PHU"))
		%>		
		<tr>
			<Td><%=rsx("yymm")%></td>
			<Td align=right><%=formatnumber(rsx("BB"),0)%></td>
			<Td align=right><%=formatnumber(rsx("CV"),0)%></td>
			<Td align=right><%=formatnumber(rsx("PHU"),0)%></td>
		</tr>
		<%rsx.movenext%>		
		<%wend%>
		<tr>
			<Td align=right>Total</td>
			<Td align=right><%=formatnumber(A,0)%></td>
			<Td align=right><%=formatnumber(B,0)%></td>
			<Td align=right><%=formatnumber(C,0)%></td>
		</tr>
		<tr>
			<td align=right>平均時薪(AVG)</td>
			<Td colspan=3 align=left>&nbsp;&nbsp;<%=formatnumber(((A+B+C)/ii)/208,0)%></td>			
		</tr>
		<%end if
		set rsx=nothing 
		%>
	</table>
	
</form>
</body>
</html>
<script language=vbs>
function go()
	yymmstr=<%=self%>.yymm.value 
 	memostr = escape(<%=self%>.memo.value) 
 	index = <%=self%>.index.value 
 	CurrentPage = <%=self%>.CurrentPage.value 
	open "<%=SELF%>.memoupd.asp?ftype=memochk&index="&index &"&CurrentPage="& <%=CurrentPage%> & _ 
 		 "&yymm="& yymmstr &_
 		 "&code=" & memostr  , "_self"  	 		
end function    
 
	
</script>
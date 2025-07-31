<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<!-- #include file = "../GetSQLServerConnection.fun" -->
<!-- #include file="../ADOINC.inc" -->
<%

SELF = "YEIE0201"

Set conn = GetSQLServerConnection()
'Set rs = Server.CreateObject("ADODB.Recordset")

nowmonth = year(date())&right("00"&month(date()),2)
if month(date())="01" then
	calcmonth = year(date()-1)&"12"
else
	calcmonth =  year(date())&right("00"&month(date())-1,2)
end if

'sql="select max(empid) empid  from empfile where left(empid,1)='L' "
'set rds=conn.execute(sql)
'if not rds.eof then
'	eid = "L" & right("0000" & cstr(cdbl(right(rds("empid"),4))+1) , 4)
'else
'	eid=""
'end if
'set rds=nothing

FUNCTION FDT(D)
	Response.Write YEAR(D)&"/"&RIGHT("00"&MONTH(D),2)&"/"&RIGHT("00"&DAY(D),2)
END FUNCTION 

if request("jxym")="" then 
	jxym=nowmonth
end if 	
%>

<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta HTTP-EQUIV="refresh" >
<link rel="stylesheet" href="../Include/style.css" type="text/css">
<link rel="stylesheet" href="../Include/style2.css" type="text/css">
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--
'-----------------enter to next field
function enterto()
	if window.event.keyCode = 13 then window.event.keyCode =9
end function

function f()
	<%=self%>.jxym.focus()
	<%=self%>.jxym.select()
end function

function groupchg()
	code = <%=self%>.GROUPID.value
	open "<%=self%>.back.asp?ftype=groupchg&code="&code , "Back"
	'parent.best.cols="50%,50%"
end function

function unitchg()
	code = <%=self%>.unitno.value
	open "<%=self%>.back.asp?ftype=UNITCHG&code="&code , "Back"	
	'parent.best.cols="50%,50%"
end function
-->
</SCRIPT>
</head>
<body  topmargin="0" leftmargin="5"  marginwidth="0" marginheight="0"  onkeydown="enterto()" onload="f()">
<form  name="<%=self%>" method="post" action="<%=self%>.upd.asp"   >
<INPUT TYPE=HIDDEN NAME="UID" VALUE=<%=SESSION("NETUSER")%>>
<INPUT TYPE=hidden NAME=nowmonth VALUE="<%=nowmonth%>">
<input name=act value="EMPADDNEW" type=hidden >
<table width="460" border="0" cellspacing="0" cellpadding="0">
	<tr><TD  >
		<img border="0" src="../image/icon.gif" align="absmiddle">
		<%=session("pgname")%>
		</TD>
	</tr>
</table>
<hr size=0	style='border: 1px dotted #999999;' align=left width=500>
<TABLE WIDTH=500 CLASS=TXT BORDER=0>
	<TR height=25 >		
		<TD width=70 align=right>績效年月:</TD>
		<td><input name=jxym  size=10 value=<%=jxym%>></td>
		<TD width=70 align=right>計薪年月:</TD>
		<td><input name=saym  size=10></td>
		<td><input type=button name="btn" value="(S)查詢K.Tra" class=button ></td>
	</TR>	 
</table>
<hr size=0	style='border: 1px dotted #999999;' align=left width=500>
<table width=500 class=txt>
	<%sqln="select a.* , isnull(b.sys_value,'') sys_value from  "&_	
		   "(select yymm ym,  groupid   from  bempg  where country='VN'  "&_
		   "and  isnull(groupid,'')<>'' group by   groupid , yymm having yymm ='"& jxym &"' ) a "&_
		   "join (select * from  basiccode where  func='groupid' ) b on b.sys_type = a.groupid "&_
		   "where b.sys_type not in ('AAA','A011', 'A021', 'A031' )  order by sys_type "
	  set rs1=conn.execute(sqln)
	  'response.write sqln 
	  'response.end
	  while not rs1.eof  
	%> 		
	 <tr bgcolor=#e4e4e4>
	 	<td>班別</td>
	 	<td>單位</td>
	 	<td>組</td>
	 	<td id=dec1>上月事故金額</td>
	 	<td >
	 		<%if rs1("groupid")="A061" then 
	 			desc2="生產M2"
	 		  elseif rs1("groupid")="A062" then 
	 		  	desc2="機台產能"
	 		  elseif rs1("groupid")="A063" then 
	 		  	desc2="機故時間"	
	 		  else	
	 		  	desc2=""
	 		  end if 
	 		  %><%=desc2%>
	 	</td>
	 	<td id=dec3>事故上限</td>
	 	<td id=dec4>事故金額</td>
	 	<td id=dec5>
	 		<%if rs1("groupid")="A061" then 
	 			desc5="績效損耗"
	 		  elseif rs1("groupid")="A062" then 
	 		  	desc5="產能/H"
	 		  elseif rs1("groupid")="A063" then 
	 		  	desc5="用油量"	
	 		  else	
	 		  	desc5=""
	 		  end if 
	 		  %><%=desc5%>	 	
	 	</td>
	 	<td id=dec6> </td>
	 	<td id=dec7>業績獎金</td>
	 </tr>
	 <%
	 sqlt="select a.* , isnull(b.sys_value,'') sys_value from  "&_	
		   "(select  yymm as ym, groupid , zuno  from  bempg  where  isnull(groupid,'')<>'' "&_
		   "group by  groupid , zuno , yymm having yymm ='"& jxym &"' and  groupid='"& rs1("groupid") &"'   ) a "&_
		   "left join (select * from  basiccode where  func='zuno' ) b on b.sys_type = a.zuno "&_
		   "order by sys_type "
	 'response.write sqlt 
	 set rs2=conn.execute(Sqlt)  
	 while not rs2.eof 	
	 	sqlt1="select a.* , isnull(b.sys_value,'') sys_value from  "&_	
		      "(select yymm as ym, groupid , zuno , shift from  bempg where isnull(groupid,'')<>'' "&_   
		      "group by  groupid , zuno , shift , yymm having yymm ='"& jxym &"' and groupid='"& rs1("groupid") &"' and isnull(zuno,'')='"& rs2("zuno") &"'   ) a "&_
		      "left join (select * from  basiccode where  func='shift' ) b on b.sys_type = a.shift "&_
		      "order by len(sys_type) desc, shift "
		set rs3=conn.execute(Sqlt1)  
		'response.write sqlt1 &"<BR>"
		while not rs3.eof 
	 %> 
	 <tr>
	 	<td>
	 		<select name="shift" class=txt8>	
	 			<%sql="select * from basiccode where func='shift' and sys_type='"& rs3("shift") &"'  order by sys_type"
	 			set rds=conn.execute(sql)
	 			while not rds.eof 
	 			%>
	 				<option value="<%=rds("sys_type")%>"><%=rds("sys_type")%>-<%=rds("sys_value")%></option>
	 			<%rds.movenext
	 			wend
	 			set rds=nothing %>
	 		</select>
	 	</td>
	 	<td>
	 		<select name="groupid" class=txt8 onchange="groupchg(<%=x-1%>)">	
	 			<%sql="select * from basiccode where func='groupid' and sys_type='"& rs1("groupid") &"' order by sys_type"
	 			set rds=conn.execute(sql)
	 			while not rds.eof 
	 			%>
	 				<option value="<%=rds("sys_type")%>"><%=rds("sys_type")%>-<%=rds("sys_value")%></option>
	 			<%rds.movenext
	 			wend
	 			set rds=nothing %>
	 		</select>
	 	</td>
	 	<td>
	 		<select name="zuno" class=txt8 style='width:110'>	
	 			<%sql="select * from basiccode where func='zuno' and (sys_type)='"& rs2("zuno") &"'   order by sys_type"
	 			set rds=conn.execute(sql)
	 			while not rds.eof 
	 			%>
	 				<option value="<%=rds("sys_type")%>"><%=rds("sys_type")%>-<%=rds("sys_value")%></option>
	 			<%rds.movenext
	 			wend
	 			set rds=nothing %>
	 		</select>
	 	</td>	 	
	 	<td><input name=LsucoM class=readonly readonly value="0" size=10></td>
	 	<td><input name=LsucoM class=readonly readonly value="0" size=10></td>
	 	<td><input name=LsucoM class=readonly readonly value="0" size=10></td>
	 	<td><input name=LsucoM class=readonly readonly value="0" size=10></td>	 	
	 	<td><input name=LsucoM class=readonly readonly value="0" size=10></td>	 	
	 	<td><input name=LsucoM class=readonly readonly value="0" size=10></td>	 	
	 	<td><input name=LsucoM class=readonly readonly value="0" size=10></td>	 	
	 	 
	 </tr>
	 <%	  rs3.movenext
	 	wend 
	 	set rs3=nothing
	 %>
	 <%rs2.movenext
	 wend 
	 set rs2=nothing
	 %> 
	 <%rs1.movenext
	 wend
	 set rs1=nothing
	 %>
</table>
<TABLE WIDTH=460>
		<tr ALIGN=center>
			<TD >
			<input type=button  name=send value="確　　認"  class=button onclick=go()>
			<input type=RESET name=send value="取 　　消"  class=button>
			</TD>
		</TR>
</TABLE>


</form>


</body>
</html>
<script language=vbscript>
 
function groupchg(index)	
	if <%=self%>.groupid(index).value="A061"   then 
		 <%=self%>.desc2.value="生產M2"
	end if 
end function 
  
</script>


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
<TABLE WIDTH=400 CLASS=TXT BORDER=0>
	<TR height=25 >		
		<TD width=70 align=right>績效年月:</TD>
		<td><input name=jxym  size=10 value=<%=jxym%>></td>
		<TD width=70 align=right>計薪年月:</TD>
		<td><input name=saym  size=10></td>
	 
	</TR>	 
</table>
<hr size=0	style='border: 1px dotted #999999;' align=left width=500>
<table width=520 class=txt BORDER="0" cellspacing="1" cellpadding="1" BGCOLOR="#000000" >
 	<tr bgcolor=#e4e4e4 HEIGHT=22>
	 	<td colspan=8>---文房---</td>
	</tr>
	<tr bgcolor=#fafad2 HEIGHT=22>
	 	<td ALIGN=CENTER>STT</td>
	 	<td ALIGN=CENTER>班別</td>
	 	<td ALIGN=CENTER>組別</td>
	 	<td ALIGN=CENTER></td>
	 	<td ALIGN=CENTER></td>
	 	<td ALIGN=CENTER>盈餘目標</td>
	 	<td ALIGN=CENTER>產能平均</td>
	 	<td ALIGN=CENTER>事故平均</td>
	</tr>
	 <tr bgcolor=#ffffff HEIGHT=22>
	 	<td ALIGN=CENTER>1<input type=hidden name=xid value="1" size=1></td>
	 	<td ALIGN=CENTER>常日班</td>
	 	<td ALIGN=CENTER> </td>
	 	<td ALIGN=CENTER> </td>
	 	<td ALIGN=CENTER></td>
	 	<td ALIGN=CENTER><INPUT NAME=A SIZE=10 CLASS=INPUTBOX onblur="Achg(0)"></td>
	 	<td ALIGN=CENTER></td>
	 	<td ALIGN=CENTER></TD>
	 	<input name=B type=hidden value="">
	 	<input name=C type=hidden value="">
	 	<input name=D type=hidden value="">
	 	<input name=groupid type=hidden value="">
	 	<input name=zuno type=hidden value="">
	 	<input name=shift type=hidden value=""> 
	 </tr>	
 	<tr bgcolor=#e4e4e4 HEIGHT=22>
	 	<td colspan=8>---儲運組---</td>
	</tr>
	<tr bgcolor=#fafad2 HEIGHT=22>
		<td ALIGN=CENTER>STT</td>
	 	<td ALIGN=CENTER>班別</td>
	 	<td ALIGN=CENTER>組別</td>
	 	<td ALIGN=CENTER></td>
	 	<td ALIGN=CENTER></td>
	 	<td ALIGN=CENTER>本月事故</td>
	 	<td ALIGN=CENTER></td>
	 	<td ALIGN=CENTER></td>
	</tr>
	 <tr bgcolor=#ffffff HEIGHT=22>
	 	<td ALIGN=CENTER>2<input type=hidden name=xid value="2" size=1></td>
	 	<td ALIGN=CENTER>常日班</td>
	 	<td ALIGN=CENTER></td>
	 	<td ALIGN=CENTER></td>
	 	<td ALIGN=CENTER></td>
	 	<td ALIGN=CENTER><INPUT NAME=A SIZE=10 CLASS=INPUTBOX onblur="Achg(1)"></td>
	 	<td ALIGN=CENTER></td>
	 	<td ALIGN=CENTER></td>
	 	<input name=B type=hidden value="">
	 	<input name=C type=hidden value="">
	 	<input name=D type=hidden value="">
	 	<input name=groupid type=hidden value="A059">
	 	<input name=zuno type=hidden value="">
	 	<input name=shift type=hidden value="">
	 </tr>		 
 	<tr bgcolor=#e4e4e4 HEIGHT=22>
	 	<td colspan=8>---工務組---</td>
	</tr>
	<tr bgcolor=#fafad2 HEIGHT=22>
	 	<td ALIGN=CENTER>STT</td>
	 	<td ALIGN=CENTER>班別</td>
	 	<td ALIGN=CENTER>組別</td>
	 	<td ALIGN=CENTER></td>
	 	<td ALIGN=CENTER></td>
	 	<td ALIGN=CENTER>機故時間</td>
	 	<td ALIGN=CENTER>用油量</td>
	 	<td ALIGN=CENTER>產能平均</td>
	</tr>
	 <tr bgcolor=#ffffff HEIGHT=22>
	 	<td ALIGN=CENTER>3<input type=hidden name=xid value="3" size=1></td>
	 	<td ALIGN=CENTER>常日班</td>
	 	<td ALIGN=CENTER> </td>
	 	<td ALIGN=CENTER> </td>
	 	<td ALIGN=CENTER></td>
	 	<td ALIGN=CENTER><INPUT NAME=A SIZE=10 CLASS=INPUTBOX onblur="Achg(2)"></td>
	 	<td ALIGN=CENTER><INPUT NAME=B SIZE=10 CLASS=INPUTBOX onblur="Bchg(2)"></td>
	 	<td ALIGN=CENTER></td> 
	 	<input name=C type=hidden value="">
	 	<input name=D type=hidden value="">
	 	<input name=groupid type=hidden value="A063">
	 	<input name=zuno type=hidden value="">
	 	<input name=shift type=hidden value="">
	 </tr>		 
 	<tr bgcolor=#e4e4e4 HEIGHT=22>
	 	<td colspan=8>*平板組---</td>
	 </tr>
	 <tr bgcolor=#fafad2 HEIGHT=22>
	 	<td ALIGN=CENTER>STT</td>
	 	<td ALIGN=CENTER >班別</td>
	 	<td ALIGN=CENTER >組別</td>
	 	<td ALIGN=CENTER> </td>
	 	<td ALIGN=CENTER> </td>
	 	<td ALIGN=CENTER>維修費用</td>
	 	<td ALIGN=CENTER>原紙庫存</td>
	 	<td ALIGN=CENTER>殘捲數</td>
	 </tr>	 
	 <%sql="select * from basiccode where func='shift' and sys_type='ALL' order by sys_type"
	 set rds=conn.execute(sql)
	 while not rds.eof
	 %>	
	 <tr bgcolor=#ffffff HEIGHT=22>
	 	<td ALIGN=CENTER>4<input type=hidden name=xid value="4" size=1></td>
	 	<td ALIGN=CENTER>常日班</td>
	 	<td ALIGN=CENTER>抱車</td>
	 	<td ALIGN=CENTER></td>
	 	<td ALIGN=CENTER></td>
	 	<td ALIGN=CENTER><INPUT NAME=A SIZE=10 CLASS=INPUTBOX onblur="Achg(3)"></td>
	 	<td ALIGN=CENTER><INPUT NAME=B SIZE=10 CLASS=INPUTBOX onblur="Bchg(3)"></td>
	 	<td ALIGN=CENTER><INPUT NAME=C SIZE=10 CLASS=INPUTBOX onblur="Cchg(3)"></td>
	 	<input name=D type=hidden value="">
	 	<input name=groupid type=hidden value="A061">
	 	<input name=zuno type=hidden value="A0612">
	 	<input name=shift type=hidden value="">
	 </tr>
	 <%
	 rds.movenext
	 wend 
	 set rds=nothing 
	 %>   	
	 <tr bgcolor=#fafad2 HEIGHT=22>
	 	<td ALIGN=CENTER>STT</td>
	 	<td ALIGN=CENTER>班別</td>
	 	<td ALIGN=CENTER>組別</td>
	 	<td ALIGN=CENTER>上月事故</td>
	 	<td ALIGN=CENTER>生產M2</td>
	 	<td ALIGN=CENTER>本月事故</td>
	 	<td ALIGN=CENTER>績效損耗</td>
	 	<td ALIGN=CENTER>產能效率</td>
	 </tr>
	 <%sql="select * from basiccode where func='shift' and sys_type<>'ALL' order by sys_type"
	 set rds=conn.execute(sql)
	 z1=4
	 while not rds.eof
	 	z1=z1+1
	 %>	
	 <tr bgcolor=#ffffff HEIGHT=22>
	 	<td ALIGN=CENTER><%=z1%><input type=hidden name=xid value="<%=z1%>" size=1></td>
	 	<td ALIGN=CENTER><%=rds("sys_type")%></td>
	 	<td ALIGN=CENTER> </td>
	 	<td ALIGN=CENTER><INPUT NAME=LSUCUM SIZE=10 CLASS=readonly readonly ></td>
	 	<td ALIGN=CENTER><INPUT NAME=A SIZE=10 CLASS=INPUTBOX onblur="Achg(<%=z1-1%>)"></td>
	 	<td ALIGN=CENTER><INPUT NAME=B SIZE=10 CLASS=INPUTBOX onblur="Bchg(<%=z1-1%>)"></td>
	 	<td ALIGN=CENTER><INPUT NAME=C SIZE=10 CLASS=INPUTBOX onblur="Cchg(<%=z1-1%>)"></td>
	 	<td ALIGN=CENTER><INPUT NAME=D SIZE=10 CLASS=INPUTBOX onblur="Dchg(<%=z1-1%>)"></td>
	 	<input name=groupid type=hidden value="A061">
	 	<input name=zuno type=hidden value="A0611">
	 	<input name=shift type=hidden value="<%=rds("sys_type")%>">
	 </tr>
	 <%
	 rds.movenext
	 wend 
	 set rds=nothing 
	 %> 
 	<tr bgcolor=#e4e4e4 HEIGHT=22>
	 	<td colspan=8>*印製組---</td>
	 </tr>
	 <tr bgcolor=#fafad2 HEIGHT=22>
	 	<td ALIGN=CENTER>STT</td>
	 	<td ALIGN=CENTER>班別</td>
	 	<td ALIGN=CENTER>組別</td>
	 	<td ALIGN=CENTER>上月事故</td>
	 	<td ALIGN=CENTER>產能M2</td>
	 	<td ALIGN=CENTER>本月事故</td>
	 	<td ALIGN=CENTER>產量/H</td>
	 	<td ALIGN=CENTER>欠量率</td>
	 </tr>
	 <%sql="select * from basiccode where func='shift' and sys_type<>'ALL' order by sys_type"
	 set rds=conn.execute(sql)
	 while not rds.eof
	 %>	
	 <tr bgcolor=#ffffff HEIGHT=22>
	 	<td ALIGN=CENTER><%=z1+1%><input type=hidden name=xid value="<%=z1+1%>" size=1></td>
	 	<td ALIGN=CENTER><%=rds("sys_type")%></td>
	 	<td ALIGN=CENTER>F</td>
	 	<td ALIGN=CENTER><INPUT NAME=LSUCUM SIZE=10 CLASS=readonly readonly ></td>
	 	<td ALIGN=CENTER><INPUT NAME=A SIZE=10 CLASS=INPUTBOX onblur="Achg(<%=z1+1-1%>)"></td>
	 	<td ALIGN=CENTER><INPUT NAME=B SIZE=10 CLASS=INPUTBOX onblur="Bchg(<%=z1+1-1%>)"></td>
	 	<td ALIGN=CENTER><INPUT NAME=C SIZE=10 CLASS=INPUTBOX onblur="Cchg(<%=z1+1-1%>)"></td>
	 	<td ALIGN=CENTER><INPUT NAME=D SIZE=10 CLASS=INPUTBOX onblur="Dchg(<%=z1+1-1%>)"></td>
	 	<input name=groupid type=hidden value="A062">
	 	<input name=zuno type=hidden value="A0622">
	 	<input name=shift type=hidden value="<%=rds("sys_type")%>">
	 </tr>
	 <tr bgcolor=#ffffff HEIGHT=22>
	 	<td ALIGN=CENTER><%=z1+2%><input type=hidden name=xid value="<%=z1+2%>" size=1></td>
	 	<td ALIGN=CENTER><%=rds("sys_type")%></td>
	 	<td ALIGN=CENTER>P</td>
	 	<td ALIGN=CENTER><INPUT NAME=LSUCUM SIZE=10 CLASS=readonly readonly ></td>
	 	<td ALIGN=CENTER><INPUT NAME=A SIZE=10 CLASS=INPUTBOX onblur="Achg(<%=z1+2-1%>)"></td>
	 	<td ALIGN=CENTER><INPUT NAME=B SIZE=10 CLASS=INPUTBOX onblur="Bchg(<%=z1+2-1%>)"></td>
	 	<td ALIGN=CENTER><INPUT NAME=C SIZE=10 CLASS=INPUTBOX onblur="Cchg(<%=z1+2-1%>)"></td>
	 	<td ALIGN=CENTER><INPUT NAME=D SIZE=10 CLASS=INPUTBOX onblur="Dchg(<%=z1+2-1%>)"></td>
	 	<input name=groupid type=hidden value="A062">
	 	<input name=zuno type=hidden value="A0623">
	 	<input name=shift type=hidden value="<%=rds("sys_type")%>">
	 </tr>
	 <tr bgcolor=#ffffff HEIGHT=22>
	 	<td ALIGN=CENTER><%=z1+3%><input type=hidden name=xid value="<%=z1+3%>" size=1></td>
	 	<td ALIGN=CENTER><%=rds("sys_type")%></td>
	 	<td ALIGN=CENTER>E</td>
	 	<td ALIGN=CENTER><INPUT NAME=LSUCUM SIZE=10 CLASS=readonly readonly></td>
	 	<td ALIGN=CENTER><INPUT NAME=A SIZE=10 CLASS=INPUTBOX onblur="Achg(<%=z1+3-1%>)"></td>
	 	<td ALIGN=CENTER><INPUT NAME=B SIZE=10 CLASS=INPUTBOX onblur="Bchg(<%=z1+3-1%>)"></td>
	 	<td ALIGN=CENTER><INPUT NAME=C SIZE=10 CLASS=INPUTBOX onblur="Cchg(<%=z1+3-1%>)"></td>
	 	<td ALIGN=CENTER><INPUT NAME=D SIZE=10 CLASS=INPUTBOX onblur="Dchg(<%=z1+3-1%>)"></td>
	 	<input name=groupid type=hidden value="A062">
	 	<input name=zuno type=hidden value="A0624">
	 	<input name=shift type=hidden value="<%=rds("sys_type")%>">
	 </tr>
	 <tr bgcolor=#ffffff HEIGHT=22>
	 	<td ALIGN=CENTER><%=z1+4%><input type=hidden name=xid value="<%=z1+4%>" size=1></td>
	 	<td ALIGN=CENTER><%=rds("sys_type")%></td>
	 	<td ALIGN=CENTER>C</td>
	 	<td ALIGN=CENTER><INPUT NAME=LSUCUM SIZE=10 CLASS=readonly readonly></td>
	 	<td ALIGN=CENTER><INPUT NAME=A SIZE=10 CLASS=INPUTBOX onblur="Achg(<%=z1+4-1%>)"></td>
	 	<td ALIGN=CENTER><INPUT NAME=B SIZE=10 CLASS=INPUTBOX onblur="Bchg(<%=z1+4-1%>)"></td>
	 	<td ALIGN=CENTER><INPUT NAME=C SIZE=10 CLASS=INPUTBOX onblur="Cchg(<%=z1+4-1%>)"></td>
	 	<td ALIGN=CENTER><INPUT NAME=D SIZE=10 CLASS=INPUTBOX onblur="Dchg(<%=z1+4-1%>)"></td>
	 	<input name=groupid type=hidden value="A062">
	 	<input name=zuno type=hidden value="A0625">
	 	<input name=shift type=hidden value="<%=rds("sys_type")%>">
	 </tr>
	 <tr bgcolor=#ffffff HEIGHT=22>
	 	<td ALIGN=CENTER><%=z1+5%><input type=hidden name=xid value="<%=z1+5%>" size=1></td>
	 	<td ALIGN=CENTER><%=rds("sys_type")%></td>
	 	<td ALIGN=CENTER>後段</td>
	 	<td ALIGN=CENTER><INPUT NAME=LSUCUM SIZE=10 CLASS=readonly readonly></td>
	 	<td ALIGN=CENTER><INPUT NAME=A SIZE=10 CLASS=INPUTBOX onblur="Achg(<%=z1+5-1%>)"></td>
	 	<td ALIGN=CENTER><INPUT NAME=B SIZE=10 CLASS=INPUTBOX onblur="Bchg(<%=z1+5-1%>)"></td>
	 	<td ALIGN=CENTER><INPUT NAME=C SIZE=10 CLASS=INPUTBOX onblur="Cchg(<%=z1+5-1%>)"></td>
	 	<td ALIGN=CENTER><INPUT NAME=D SIZE=10 CLASS=INPUTBOX onblur="Dchg(<%=z1+5-1%>)"></td>
	 	<input name=groupid type=hidden value="A062">
	 	<input name=zuno type=hidden value="A0626">
	 	<input name=shift type=hidden value="<%=rds("sys_type")%>">
	 </tr>	 	 	 	 
	 <%
	 Z1=Z1+5
	 rds.movenext
	 wend 
	 set rds=nothing 
	 %> 
</table>
<input type=hidden name=z1 value="<%=z1%>">
<TABLE WIDTH=460>
		<tr ALIGN=center>
			<TD >
			<input type=button  name=send value="(Y)確　　認"  class=button onclick=go()>
			<input type=RESET name=send value="(N)取 　　消"  class=button>
			</TD>
		</TR>
</TABLE>


</form>


</body>
</html>
<script language=vbscript>

function Achg(index) 	 
	if trim(<%=self%>.a(index).value)<>"" then 
		if isnumeric(<%=self%>.a(index).value)=false then  
			alert "請輸入數字!!hay nhap so vao!!"
			<%=self%>.a(index).value=""
			<%=self%>.a(index).focus()
			exit function 		
		end if	
	end if 	
end function 
function Bchg(index) 	 
	if trim(<%=self%>.B(index).value)<>"" then 
		if isnumeric(<%=self%>.B(index).value)=false then  
			alert "請輸入數字!!hay nhap so vao!!"
			<%=self%>.B(index).value=""
			<%=self%>.B(index).focus()
			exit function 		
		end if	
	end if 	
end function 
function Cchg(index) 	 
	if trim(<%=self%>.C(index).value)<>"" then 
		if isnumeric(<%=self%>.C(index).value)=false then  
			alert "請輸入數字!!hay nhap so vao!!"
			<%=self%>.C(index).value=""
			<%=self%>.C(index).focus()
			exit function 		
		end if	
	end if 	
end function 
function Dchg(index) 	 
	if trim(<%=self%>.D(index).value)<>"" then 
		if isnumeric(<%=self%>.D(index).value)=false then  
			alert "請輸入數字!!hay nhap so vao!!"
			<%=self%>.D(index).value=""
			<%=self%>.D(index).focus()
			exit function 		
		end if	
	end if 	
end function 
 
function  go()
	<%=self%>.action="<%=self%>.upd.asp"
	<%=self%>.submit()
end function 
  
</script>


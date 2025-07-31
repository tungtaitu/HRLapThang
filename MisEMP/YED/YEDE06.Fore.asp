<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<!-- #include file = "../GetSQLServerConnection.fun" -->
<!-- #include file="../ADOINC.inc" -->
<!-- #include file="../Include/func.inc" -->
<!--#include file="../include/sideinfo.inc"-->
<%
Set conn = GetSQLServerConnection()
self="YEDE06"

SESSION.CODEPAGE=65001
nowmonth = year(date())&right("00"&month(date()),2)
if month(date())="01" then
	'calcmonth = year(date()-1)&"12"
	calcmonth = nowmonth
else
	'calcmonth =  year(date())&right("00"&month(date())-1,2)
	calcmonth = nowmonth 
end if

if day(date())<=11 then
	if month(date())="01" then
		calcmonth = year(date()-1)&"12"
		calcmonth = nowmonth 
	else
		calcmonth =  year(date())&right("00"&month(date())-1,2)
		calcmonth = nowmonth 
	end if
else
	calcmonth = nowmonth
end if  

D1=request("D1")
D2=request("D2")  
if D1="" then D1=nowmonth 
if D2="" then D2=DD2
sortby=request("sortby") 
queryx = request("queryx")
if sortby="" then 
	sortby="emp_id,card_id,workdat"
end if	

q1=request("q1")
q2=request("q2")
sql=" select isnull(b.yymm,'"&d1&"') as yymm, a.* ,isnull(b.tothrs,0) tothrs, b.mdtm, b.muser  from  "&_
		"(select * from view_allgrouips  )  a  "&_
		"left join  ( select * from  emptothr  where  yymm='"&d1&"' ) b on b.groupid=a.sys_type and b.zuno=a.zuno and b.shift = a.shift  "&_
		"where( a.sys_type like '"&q1&"%' and a.zuno like '"&q2&"%'  ) "&_
		"order by a.sys_type, a.zuno, len(a.shift) desc , a.shift "
Set rs = Server.CreateObject("ADODB.Recordset")
'response.write sql	
'response.end
rs.open sql, conn, 3, 1	
if not rs.eof then 
	pagerec = rs.recordcount 
else
	pagerec=1
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
 
function f()
	<%=self%>.D1.focus()
	<%=self%>.D1.SELECT()
end function

</SCRIPT>
</head>
<body  topmargin="5" leftmargin="5"  marginwidth="0" marginheight="0"  onkeydown="enterto()" onload='f()' >
<form name="<%=self%>" method="post" action="<%=self%>.FORE.ASP">
<input name="PageRec" type="hidden" value="<%=pagerec%>">
<table border=0 style="height:70px"><tr><td>&nbsp;</td></tr></table>
<table border=0 width="100%"><tr><td align="center">
<table width=500  ><tr><td >
	<table width=500   border=0 cellspacing="1" cellpadding="1"  class=txt >
		<tr height=30 >
			<TD align=right nowrap >年月:</TD>
			<TD >
				<input type="text" style="width:100px" name="d1"  maxlength=6 value="<%=d1%>" > 				
			</TD>			 
		</tr>
		<tr>
			<td align=right>部門單位:</td>
			<TD>
				<%sqla="select * from basiccode where  func='groupid' and ( sys_type='a033 ' or sys_type >'A051' )  and sys_type<>'AAA'  order by sys_type " 
					set rsx=conn.execute(sqla)
				%>
				<select  name="q1"   onchange="submit()" style="width:120px"> 
					<option value="" />
					<%while not rsx.eof %>
					<option value="<%=rsx("sys_type")%>" <%if request("q1")=rsx("sys_type") then%> selected<%end if%>/><%=rsx("sys_type")%> <%=rsx("sys_value")%>
					<%rsx.movenext
					wend
					set rsx=nothing 
					%>
				</select>
			</td>
			<td>
				<%
					sqla="select * from basiccode where  func='zuno' and left(sys_type,4) >'A051'  and left(sys_type,3)<>'AAA' "&_
							 "and  left(sys_type,4) like '"&  request("q1") &"%'  order by sys_type " 
					set rsx=conn.execute(sqla)
				%>
				<select  name="q2"   onchange="submit()" style="width:180;"> 
				<option value="" />
					<%while not rsx.eof %>
					<option value="<%=rsx("sys_type")%>" <%if request("q2")=rsx("sys_type") then%> selected<%end if%> /><%=rsx("sys_type")%> <%=rsx("sys_value")%>
					<%rsx.movenext
					wend
					set rsx=nothing 
					%>				
				</select>
			</TD>
			<td><input type=submit name=btn value="(S)查詢" class="btn btn-sm btn-outline-secondary" onkeydown="submit()"></td>			
		</TR> 
		<tr>
			<td align="right">相同時數:</td>
			<td><input name="samehr"  size=10 ></td>
			<td align="right"><input type="button" name=btn value="(Y)confirm" class="btn btn-sm btn-outline-secondary" onclick="samehrclick()"></td>
			<td><input type="button" name=btn value="(C)Clear" class="btn btn-sm btn-outline-secondary" onclick="clrhr()"></td>
		</tr>		
	</table>
	<table width=600 align=center border=0 cellspacing="1" cellpadding="1"  class=txt>
		<tr bgcolor="#e4e4e4" height=22>
			<td align=center>STT</td>
			<td align=center>yymm</td>
			<td align=center>部門</td>
			<td align=center>單位</td>
			<td align=center>班別</td>
			<td align=center>總時數</td>			
		</tr>
		<%for x = 1 to rs.recordcount
			if x mod 2 = 0 then 
				wkcolor="lightyellow"
			else
				wkcolor="LavenderBlush"
			end if 
		%>
		<tr bgcolor="<%=wkcolor%>" height=22>
			<td  align=center><%=x%></td>
			<td  align=center><input type="text" style="width:100%" name="yymm" class="readonly" value="<%=rs("yymm")%>" readonly ></td>
			<td  align=center>
			<input type="text" style="width:48%" name="groupid" class="readonly"  value="<%=rs("sys_type")%>" readonly >
			<input type="text" style="width:48%" name="gstr" class="readonly"  readonly value="<%=rs("sys_value")%>">
			</td>
			<td  align=center>
			<input type="text" style="width:48%" name="zuno" class="readonly"  value="<%=rs("zuno")%>" readonly >
			<input type="text" style="width:48%" name="zstr" class="readonly"   readonly value="<%=rs("zstr")%>">
			</td>
			<td  align=center>
			<input type="text" style="width:100%" name="shift" class="readonly"  value="<%=rs("shift")%>" readonly >			
			</td>
			<td  align=center>
			<input type="text" style="width:100%" name="tothr" class="inputbox"  value="<%=rs("tothrs")%>"   >			
			</td>
			 
		</tr>
		<%
		rs.movenext
		next
		%>
		<%
		rs.close
		set rs=nothing%>
		<input name=func type=hidden>
		<input name=op type=hidden>
		<input name=lsempid type=hidden>
		<input name=cardNo type=hidden>
		<input name=workdat type=hidden>
		<input name=T1 type=hidden>
		<input name=T2 type=hidden>
		
		<input name=empid type=hidden>
		<input name=B_workdat type=hidden>
		<input name=B_T1 type=hidden>
		<input name=B_T2 type=hidden>
		<input name=JB type=hidden>
	</table> 
	<table width=450 align=center>
		<tr >
			<td align=center>
				<input type=button  name=btm class="btn btn-sm btn-danger" value="確   認" onclick="go()" onkeydown="go()">
				<input type=reset  name=btm class="btn btn-sm btn-outline-secondary" value="取   消">
			</td>
		</tr>
	</table>

</td></tr></table>

</td></tr></table>
</body>
</html>


<script language=vbs>
function datain(index)
	if <%=self%>.func(index).checked=true then
		<%=self%>.op(index).value="Y"
		<%=self%>.workdat(index).value=<%=self%>.B_workdat(index).value
		<%=self%>.T1(index).value=<%=self%>.B_T1(index).value
		<%=self%>.T2(index).value=<%=self%>.B_T2(index).value
		<%=self%>.T1(index).focus()
		<%=self%>.T1(index).select()
	else
		<%=self%>.op(index).value=""
		<%=self%>.workdat(index).value=""
		<%=self%>.T1(index).value=""
		<%=self%>.T2(index).value=""
		<%=self%>.EMPID(index).value=""
		<%=self%>.JB(index).value="0"
	end if 
end function 

function T1chg(index)
	IF trim(<%=self%>.T1(index).value)<>"" then 
		<%=self%>.T1(index).value=left(trim(<%=self%>.T1(index).value),2)&":"&right(trim(<%=self%>.T1(index).value),2)
	end if 	
end function 

function T2chg(index)
	IF trim(<%=self%>.T2(index).value)<>"" then 
		<%=self%>.T2(index).value=left(trim(<%=self%>.T2(index).value),2)&":"&right(trim(<%=self%>.T2(index).value),2)
	end if 	
end function 

function strchg(a)
	if a=1 then
		<%=self%>.F_empid.value = Ucase(<%=self%>.F_empid.value)
	elseif a=2 then
		<%=self%>.empid2.value = Ucase(<%=self%>.empid2.value)
	end if
end function

function go()
 
	
 	<%=self%>.action="<%=self%>.upd.asp"
 	<%=self%>.submit()
end function


'*******檢查日期*********************************************
FUNCTION date_change(a)

if a=1 then
	INcardat = Trim(<%=self%>.D1.value)
elseif a=2 then
	INcardat = Trim(<%=self%>.D2.value)
end if

IF INcardat<>"" THEN
	ANS=validDate(INcardat)
	IF ANS <> "" THEN
		if a=1 then
			Document.<%=self%>.D1.value=ANS
		elseif a=2 then
			Document.<%=self%>.D2.value=ANS
		end if
	ELSE
		ALERT "EZ0067:輸入日期不合法 !!"
		if a=1 then
			Document.<%=self%>.D1.value=""
			Document.<%=self%>.D1.focus()
		elseif a=2 then
			Document.<%=self%>.D2.value=""
			Document.<%=self%>.D2.focus()
		end if
		EXIT FUNCTION
	END IF

ELSE
	'alert "EZ0015:日期該欄位必須輸入資料 !!"
	EXIT FUNCTION
END IF

END FUNCTION

function samehrclick()
	if <%=self%>.samehr.value<>"" and isnumeric(<%=self%>.samehr.value)=true  then 
	for ni = 1 to <%=self%>.pagerec.value  
		if <%=self%>.tothr(ni-1).value = "0" or <%=self%>.tothr(ni-1).value="" then 
			<%=self%>.tothr(ni-1).value = <%=self%>.samehr.value
		end if 
	next 
	end if 
end function 

function clrhr()
for ni = 1 to <%=self%>.pagerec.value  
		if <%=self%>.tothr(ni-1).value =  <%=self%>.samehr.value then 
			<%=self%>.tothr(ni-1).value="0"
		end if 
	next 
end function  
</script> 
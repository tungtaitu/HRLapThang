<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<!-- #include file = "../GetSQLServerConnection.fun" -->
<!-- #include file="../ADOINC.inc" -->
<!-- #include file="../Include/func.inc" -->
<!--#include file="../include/sideinfo.inc"-->
<%
Set conn = GetSQLServerConnection()
self="YEDE05"

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
if D1="" then D1=DD2
if D2="" then D2=DD2
sortby=request("sortby") 
queryx = request("queryx")
if sortby="" then 
	sortby="emp_id,card_id,workdat"
end if	
Set rs = Server.CreateObject("ADODB.Recordset")
if D1<>"" and D2<>"" then 
	sql="select a.* from  "&_
		"(select * from TWorkTime where workdat between '"& replace(D1,"/","") &"' and '"& replace(D2,"/","")  &"' "&_		
		"union "&_
		"select '' , * from View_TmpCard where workdat between '"& replace(D1,"/","") &"' and '"& replace(D2,"/","")  &"' "&_
		") a  "&_
		"left join (select * from empforget where isnull(status,'')<>'D'  ) b on b.lsempid = a.emp_id  and convert(varchar(8), dat, 112) = a.workdat  "&_		
		"where isnull(a.emp_id,'')<>'' and left(a.emp_id,3)='YFY' and ( emp_id like '%"&queryx &"%' or card_id like '%"&queryx &"%'  )  and isnull(b.lsempid,'')=''  "&_
		"order by " & sortby 
else
	sql="select * from TWorkTime where workdat ='xxx'"
end if 		
'response.write sql	
response.end
rs.open sql, conn, 1, 3	
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
	<table    border=0 cellspacing="1" cellpadding="1"  class=txt >
		<tr height=30 >
			<TD align=right>日期:</TD>
			<TD>
				<table border=0><tr>
				<td><input name=D1 class="form-control form-control-sm mb-2 mt-2" size=12 maxlength=10 onblur="date_change(1)" value="<%=D1%>"></td>
				<td>~</td>
				<td><input name=D2 class="form-control form-control-sm mb-2 mt-2" size=12 maxlength=10 onblur="date_change(2)" value="<%=D2%>"></td>
				</tr></table>
			</TD>
			<td align=right>排序:</td>
			<td>
				<select name=sortby class="form-control form-control-sm mb-2 mt-2" onchange="submit()">
					<option value="emp_id,card_id,workdat" <%if sortby="emp_id,card_id,workdat" then%>selected<%end if%>>(1)依工號</option>
					<option value="card_id,workdat" <%if sortby="card_id,workdat" then%>selected<%end if%>>(2)依卡號</option>
					<option value="a.workdat,a.timeup" <%if sortby="a.workdat,a.timeup" then%>selected<%end if%>>(3)依日期時間</option>
				</select>
			</td>
		</tr>
		<tr>
			<td align=right>關鍵字:</td>
			<TD colspan=2>
				<input name=queryx class="form-control form-control-sm mb-2 mt-2" size=20 value="<%=queryx%>">				
			</TD>
			<td><input type=submit name=btn value="(S)查詢" class="btn btn-sm btn-outline-secondary" onkeydown="submit()"></td>
			
		</TR>		 
	</table>
	<table width=600 align=center border=0 cellspacing="1" cellpadding="1"  class=txt>
		<tr bgcolor="#e4e4e4" height=22>
			<td align=center>STT</td>
			<td align=center>臨時工號</td>
			<td align=center>卡號</td>
			<td align=center>日期</td>
			<td align=center>上班</td>
			<td align=center>下班</td>
			<td align=center>轉<br>入</td>
			<td align=center>上班日</td>
			<td align=center>上班</td>
			<td align=center>下班</td>
			<td align=center>員工工號</td>
			<td align=center>加班(H)</td>			
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
			<td  align=center><%=rs("emp_id")%></td>
			<td align=center><%=rs("card_id")%></td>
			<td align=center><%=rs("workdat")%></td>
			<td align=center><%=rs("timeup")%></td>
			<td align=center><%=rs("timedown")%></td>
			<td align=center>
				<input type=checkbox name=func onclick=datain(<%=x-1%>)>
				<input name="op" value="" type="hidden" size="1" class="inputbox" > 
				<input name="lsempid" value="<%=rs("emp_id")%>" type="hidden" size="1" class="inputbox" > 
				<input name="cardNo" value="<%=rs("card_id")%>" type="hidden" size="1" class="inputbox" > 
			</td>
			<td align=center>
				<input name=workdat class=readonly8 readonly  size=9 style='text-align:center' >
				<input type=hidden name=B_workdat class="form-control form-control-sm mb-2 mt-2"8 size=10 value="<%=rs("workdat")%>">
			</td>
			<td align=center>
				<input name=T1 class="form-control form-control-sm mb-2 mt-2"8    size=4 style='text-align:center' onblur="T1chg(<%=x-1%>)">
				<input type=hidden name=B_T1 class="form-control form-control-sm mb-2 mt-2"8 size=10  value="<%=rs("timeup")%>">
			</td>
			<td align=center>
				<input name=T2 class="form-control form-control-sm mb-2 mt-2"8 size=4 onblur="T2chg(<%=x-1%>)" style='text-align:center'>
				<input type=hidden name=B_T2 class="form-control form-control-sm mb-2 mt-2"8 size=10  value="<%=rs("timedown")%>" >
			</td>
			<td align=center>
				<input name=empid class="form-control form-control-sm mb-2 mt-2" size=6 style='text-align:center'>
			</td>
			<td align=center>
				<input name=JB class="form-control form-control-sm mb-2 mt-2"8 size=3 style='text-align:center' value="0">				
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
				<input type=button  name=btm class="btn btn-sm btn-outline-secondary" value="確   認" onclick="go()" onkeydown="go()">
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
	pg =  <%=self%>.PageRec.value 
	for k = 1 to pg 
		if <%=self%>.op(k-1).value="Y" then 
			if <%=self%>.empid(k-1).value="" then 
				alert "請輸入工號!!"
				<%=self%>.empid(k-1).focus()
				exit function 
			end if 	
		end if 
	next 
	
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
</script> 
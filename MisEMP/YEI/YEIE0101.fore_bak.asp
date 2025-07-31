<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<!-- #include file = "../GetSQLServerConnection.fun" -->
<!-- #include file="../ADOINC.inc" -->
<%

SELF = "YEIE0101"

Set conn = GetSQLServerConnection()
Set rs = Server.CreateObject("ADODB.Recordset")
Set rds = Server.CreateObject("ADODB.Recordset")

nowmonth = year(date())&right("00"&month(date()),2)
if month(date())="01" then
	calcmonth = year(date()-1)&"12"
else
	calcmonth =  year(date())&right("00"&month(date())-1,2)
end if 

F_whsno = request("F_whsno")
F_groupid = request("F_groupid")
F_zuno = request("F_zuno") 
if F_whsno="" then F_whsno="XX"
F_shift=request("F_shift")
F_empid =request("F_empid")
F_country=request("F_country")
 

FUNCTION FDT(D)
	Response.Write YEAR(D)&"/"&RIGHT("00"&MONTH(D),2)&"/"&RIGHT("00"&DAY(D),2)
END FUNCTION 
khym = request("khym")
if request("khym")="" then 
	khym=nowmonth
end if  

act = request("act")	  
khweek = request("khweek")
tmw = request("tmw")
if tmw="" then tmw=request("tt")
 '一個月有幾天 
cDatestr=CDate(LEFT(khym,4)&"/"&RIGHT(khym,2)&"/01") 
days = DAY(cDatestr+(32-DAY(cDatestr))-DAY(cDatestr+(32-DAY(cDatestr))))   '一個月有幾天  
'本月最後一天 
ENDdat = LEFT(khym,4)&"/"&RIGHT(khym,2)&"/"&DAYS   

gTotalPage = 1
PageRec = 10    'number of records per page
TableRec = 70    'number of fields per record   

sql="select * from view_empfile where whsno='"& F_whsno &"'   and (isnull(outdat,'')='' or convert(char(10),outdat,111)>='"& ENDdat &"') "&_
	"and whsno='"&F_whsno&"' and groupid like '"&F_groupid&"%' and country like '"& F_country &"%'"&_
	"and zuno like '"&F_zuno&"%' and shift like '%"& F_shift &"'  "&_
	"order by country, shift, zuno, empid" 
'response.write sql 				
if request("TotalPage") = "" or request("TotalPage") = "0" then 
	CurrentPage = 1	
	rds.Open SQL, conn, 1, 3 
	IF NOT RdS.EOF THEN 	
		PageRec = rds.RecordCount 
		rds.PageSize = PageRec 
		RecordInDB = rds.RecordCount 
		TotalPage = rds.PageCount  
		gTotalPage = TotalPage
	END IF 	 
	Redim tmpRec(gTotalPage, PageRec, TableRec)   'Array
	
	for i = 1 to TotalPage 
		for j = 1 to PageRec
			if not rds.EOF then 			
				tmpRec(i, j, 0) = "no"
				tmpRec(i, j, 1) = rds("EMPID")
				tmpRec(i, j, 2) = rds("whsno")
				tmpRec(i, j, 3) = rds("groupid")
				tmpRec(i, j, 4) = rds("zuno")
				tmpRec(i, j, 5) = rds("shift")
				tmpRec(i, j, 6) = rds("empnam_cn")
				tmpRec(i, j, 7) = rds("empnam_vn")
				tmpRec(i, j, 8) = rds("nindat")
				tmpRec(i, j, 9) = rds("outdate")
				tmpRec(i, j, 10) = rds("gstr")
				tmpRec(i, j, 11) = rds("zstr")
				rds.MoveNext 
			else 
				exit for 
			end if 
	 	next 
	
	 	if rds.EOF then 
			rds.Close 
			Set rds = nothing
			exit for 
	 	end if 
	next
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
	if <%=self%>.act.value="A" then 
		<%=self%>.F_whsno.focus()
	else
		<%=self%>.khym.focus()
		<%=self%>.khym.select()
	end if 	
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

function datachg() 
	<%=self%>.totalpage.value="0"
	<%=self%>.action = "<%=self%>.Fore.asp"
	<%=self%>.submit()
end function  
function datachg2() 
	<%=self%>.totalpage.value="0"
	<%=self%>.act.value="A"
	<%=self%>.action = "<%=self%>.Fore.asp"
	<%=self%>.submit()
end function 

-->
</SCRIPT>
</head>
<body  topmargin="0" leftmargin="5"  marginwidth="0" marginheight="0"  onkeydown="enterto()" onload="f()">
<form  name="<%=self%>" method="post" action="<%=self%>.upd.asp"   >
<INPUT TYPE=HIDDEN NAME="UID" VALUE=<%=SESSION("NETUSER")%>>
<INPUT TYPE=hidden NAME=TotalPage VALUE="<%=TotalPage%>">
<INPUT TYPE=hidden NAME=CurrentPage VALUE="<%=CurrentPage%>">
<INPUT TYPE=hidden NAME=RecordInDB VALUE="<%=RecordInDB%>">
<INPUT TYPE=hidden NAME=PageRec VALUE="<%=PageRec%>">
<INPUT TYPE=hidden NAME=gTotalPage VALUE="<%=gTotalPage%>"> 	
<INPUT TYPE=hidden NAME=nowmonth VALUE="<%=nowmonth%>">
<input name=act value="<%=act%>" type=hidden >
<table width="460" border="0" cellspacing="0" cellpadding="0">
	<tr><TD  >
		<img border="0" src="../image/icon.gif" align="absmiddle">
		<%=session("pgname")%>
		</TD>
	</tr>
</table>
<hr size=0	style='border: 1px dotted #999999;' align=left width=600>
<TABLE WIDTH=650 CLASS=TXT BORDER=0>
	<TR height=22 >		
		<TD align=right>考核年月</TD>
		<td colspan=6>
			<input class=inputbox name=khym value="<%=khym%>" size=6 onchange=datachg2()>
			<select name=khweek class=txt8  onchange=datachg2()>
				<option value="">ALL</option>
			<%	sql="select a.*, convert(char(10),mindat,111) mindat, convert(char(10),maxdat,111) maxdat   from "&_
					"( select distinct datepart(ww, dat) as  monthweek, convert(char(6), dat, 112) as yymm from  ydbmcale  where convert(char(6), dat, 112)='"& khym &"'  ) a  "&_
					"left join   "&_
					"( select datepart(ww, dat) as  monthweek, min(dat) mindat from  ydbmcale  where convert(char(6), dat, 112)='"& khym &"'   "&_
					"  group by  datepart(ww, dat)  ) b on b.monthweek = a.monthweek  "&_
					"left join   "&_
					"( select datepart(ww, dat) as  monthweek, max(dat) maxdat from  ydbmcale  where convert(char(6), dat, 112)='"& khym &"'   "&_
					"  group by  datepart(ww, dat)  ) c on c.monthweek = a.monthweek  "&_
					"where mindat<>maxdat "
					if khweek<>"" then 
						sql=sql&"and a.monthweek='"& khweek &"' " 
					end if 	
					Set rst = Server.CreateObject("ADODB.Recordset")				
					rst.open sql, conn, 1, 3 		
					tt = rst.recordcount 
					redim weektmp(rst.recordcount,4)
					yy=0
					while not rst.eof 
						weektmp(yy,0)=rst("monthweek")
						weektmp(yy,1)=rst("yymm")
						weektmp(yy,2)=rst("mindat")
						weektmp(yy,3)=rst("maxdat") 	 
			%>
						<option value="<%=rst("monthweek")%>" <%if cstr(rst("monthweek")) = khweek then%>selected<%end if%>>第<%=rst("monthweek")%>週,<%=rst("mindat")%>~<%=rst("maxdat")%></option>
				<%		yy=yy+1
						rst.movenext
					wend
				set rst=nothing 
				%>
			</select>
			<input type=hidden name=tt value="<%=tt%>">
			<input type=hidden name=tmw value="<%=tmw%>">
		</td>
		
	</tr>
	<tr  height=22>	
		<TD align=right>廠別<br>Xuong</TD>
		<td>
			<select name=F_whsno class=txt8 style='width:100' >					
					<%
					if session("rights")=0 then 
						SQL="SELECT * FROM BASICCODE WHERE FUNC='WHSNO' ORDER BY SYS_TYPE "%>
						<option value="" selected >全部(Toan bo) </option>
					<%		
					else
						SQL="SELECT * FROM BASICCODE WHERE FUNC='WHSNO' and sys_type='"& session("NETWHSNO") &"' ORDER BY SYS_TYPE "
					end if	
					SET RST = CONN.EXECUTE(SQL)
					WHILE NOT RST.EOF
					%>
					<option value="<%=RST("SYS_TYPE")%>" <%if RST("SYS_TYPE")=F_whsno then%>selected<%end if%>><%=RST("SYS_VALUE")%></option>
					<%
					RST.MOVENEXT
					WEND
					%>
				</SELECT>
				<%SET RST=NOTHING %>	
		</td>
		<td align=right>國籍<BR>Quoc</td>
		<td>
			<select name=F_country class=txt8 style='width:70' >				
					<%
					if session("rights")=0 then 
						SQL="SELECT * FROM BASICCODE WHERE FUNC='country' ORDER BY SYS_TYPE "%>
						<option value="" selected >全部(Toan bo) </option>
					<%		
					else
						SQL="SELECT * FROM BASICCODE WHERE FUNC='WHSNO' and sys_type='"& session("NETWHSNO") &"' ORDER BY SYS_TYPE "
					end if	
					SET RST = CONN.EXECUTE(SQL)
					WHILE NOT RST.EOF
					%>
					<option value="<%=RST("SYS_TYPE")%>" <%if RST("SYS_TYPE")=F_country then%>selected<%end if%>><%=RST("SYS_VALUE")%></option>
					<%
					RST.MOVENEXT
					WEND
					%>
				</SELECT>
				<%SET RST=NOTHING %>	
		</td>		
		
		<TD align=right>部門<br>Bo Phan</TD>
		<td>
			<select name=F_groupid  class=txt8  style='width:80' >
			<%if Session("RIGHTS")<="2" or Session("RIGHTS")="8" then%>
					<option value="">全部(Toan bo) </option>
					<% 
						SQL="SELECT * FROM BASICCODE WHERE FUNC='GROUPID' and sys_type <>'AAA' ORDER BY SYS_TYPE " 
					else
						SQL="SELECT * FROM BASICCODE WHERE FUNC='GROUPID' and sys_type <>'AAA' and  sys_type= '"& session("NETG1") &"' ORDER BY SYS_TYPE "
					end if   
					
					SET RST = CONN.EXECUTE(SQL)
					RESPONSE.WRITE SQL 
					WHILE NOT RST.EOF  
				%>
					<option value="<%=RST("SYS_TYPE")%>" <%if RST("SYS_TYPE")=F_groupid then%>selected<%end if%>><%=RST("SYS_TYPE")%><%=RST("SYS_VALUE")%></option>				 
				<%
					RST.MOVENEXT
					WEND 
				%>
			</SELECT>
			<%SET RST=NOTHING %>
			<select name=F_zuno  class=txt8 style='width:70'    >				
				<option value=""></option>
				<%
					SQL="SELECT * FROM BASICCODE WHERE FUNC='zuno' and sys_type <>'XX' and  left(sys_type,4)= '"& F_groupid &"' ORDER BY SYS_TYPE "
					SET RST = CONN.EXECUTE(SQL)
					RESPONSE.WRITE SQL 
					WHILE NOT RST.EOF  
				%>
					<option value="<%=RST("SYS_TYPE")%>" <%if RST("SYS_TYPE")=F_zuno then%>selected<%end if%>><%=right(RST("SYS_TYPE"),1)%>-<%=RST("SYS_VALUE")%></option>				 
				<%
					RST.MOVENEXT
					WEND 
				%>
			</SELECT>
			<%SET RST=NOTHING %>	
		</td>
		<td align=right>班別<BR>Ca</td>
		<td>
			<select name=F_shift  class=txt8   >
				<option value=""></option>
				<option value="All" <%if F_shift="ALL" then%>selected<%end if%>>日</option>
				<option value="A" <%if F_shift="A" then%>selected<%end if%>>A班</option>
				<option value="B" <%if F_shift="B" then%>selected<%end if%>>B班</option>
			</select>
			<input type=button name=send value="(S)查詢" class=button onclick="datachg()">
		</td>  	 
	</TR>	 
</table>
<hr size=0	style='border: 1px dotted #999999;' align=left width=600>
<table BORDER=0 cellspacing="1" cellpadding="1"  class=txt8 >
	<tr height=22 bgcolor=#e4e4e4>
		<Td width=20 nowrap align=center rowspan=2 >STT</td>
		<Td width=50 nowrap align=center rowspan=2 >部門</td>
		<Td width=50 nowrap align=center rowspan=2 >單位</td>
		<Td width=30  nowrap align=center rowspan=2 >班別</td>
		<Td width=50  nowrap align=center rowspan=2 >工號</td>
		<Td width=100  nowrap align=center rowspan=2 >姓名</td>
		<Td width=65  nowrap  align=center rowspan=2 >到職日</td>
		<%if khweek="" then %>
			<%for a = 1 to tt%> 
				<Td align=center   nowrap colspan=4>
					第<%=weektmp(a-1,0)%>週<BR>
					<%=right(weektmp(a-1,2),5)%> ~ <%=right(weektmp(a-1,3),5)%>
				</td>
			<%next%>	
		<%else%>
			<Td align=center   nowrap colspan=4>
				第<%=weektmp(0,0)%>週<BR>
				<%=right(weektmp(0,2),5)%> ~ <%=right(weektmp(0,3),5)%>
			</td>	
		<%end if%>
	</tr> 
	<tr bgcolor=#e4e4e4>
		<%
		if khweek<>"" then 
			tt=1
		end if
		for g = 1 to tt%>
		<td align=center>A<br>出勤<br>50%</td>
		<td align=center>B<br>6S<br>20%</td>
		<td align=center>C<br>保養<br>20%</td>
		<td align=center>D<br>配合<br>10%</td>
		<%next%>
	</tr>

	<%
	for CurrentRow = 1 to PageRec
		IF CurrentRow MOD 2 = 0 THEN 
			WKCOLOR="LavenderBlush"
		ELSE
			WKCOLOR="#DFEFFF"
		END IF 	 		
	%>
	<TR BGCOLOR="<%=WKCOLOR%>" > 	 
		<Td align=center><%=CurrentRow%></td>
		<Td align=center><%=tmpRec(CurrentPage, CurrentRow, 10)%></td>
		<td align=center><%=tmpRec(CurrentPage, CurrentRow, 11)%></td>
		<Td align=center><%=tmpRec(CurrentPage, CurrentRow, 5)%></td>
		<Td align=center>
			<%=tmpRec(CurrentPage, CurrentRow, 1)%>
			<input type=hidden name=empid  value="<%=tmpRec(CurrentPage, CurrentRow, 1)%>">
		</td>
		<Td  ><%=tmpRec(CurrentPage, CurrentRow, 6)%><br><%=left(tmpRec(CurrentPage, CurrentRow, 7),15)%></td>
		<Td align=center><%=tmpRec(CurrentPage, CurrentRow, 8)%><br><font color=red><%=tmpRec(CurrentPage, CurrentRow, 9)%></font></td>
		<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then %>
			<%for y=1 to tt
				if y mod 2 =0 then 
					weekcolor="#e6e6fa"
				else
					weekcolor="#eee8aa"
				end if 	 
				sqld="select * from empKHB  where khym='"& khym &"' and empid='"& tmpRec(CurrentPage, CurrentRow, 1) &"' and khweek='"& weektmp(y-1,0) &"' "
				'response.write sqld
				set rs2=Server.CreateObject("ADODB.Recordset")
				rs2.open sqld, conn, 1,3 
				if rs2.eof then 
					fnA="0"
					fnB="0"
					fnC="0"
					fnD="0"
					colorA="red"
					colorB="red"
					colorC="red"
					colorD="red"
				else					
					fnA=rs2("fnA")
					fnB=rs2("fnB")
					fnC=rs2("fnC")
					fnD=rs2("fnD")
					if fna="0" then colorA="red" else colorA="black"
					if fnb="0" then colorB="red" else colorB="black"
					if fnc="0" then colorC="red" else colorC="black"
					if fnd="0" then colorD="red" else colorD="black"
				end if 				
			%>
				<td bgcolor="<%=weekcolor%>">
					<input name=fensuA class=inputbox8r size=3  onblur="fnAcng(<%=y-1%>)" value="<%=fnA%>"  style='color:<%=colorA%>'  >
				</td>
				<td bgcolor="<%=weekcolor%>">	
					<input name=fensuB class=inputbox8r size=3 onblur="fnBcng(<%=y-1%>)" value="<%=fnB%>" style='color:<%=colorB%>' >
				</td>	
				<td bgcolor="<%=weekcolor%>">
					<input name=fensuC class=inputbox8r size=3 onblur="fnCcng(<%=y-1%>)" style='color:<%=colorC%>' value="<%=fnC%>">
				</td>
				<td bgcolor="<%=weekcolor%>">	
					<input name=fensuD class=inputbox8r size=3 onblur="fnDcng(<%=y-1%>)" style='color:<%=colorD%>' value="<%=fnD%>">				
					<input type=hidden name=monthweek  value="<%=weektmp(y-1,0)%>">
					<input type=hidden name=calcym  value="<%=weektmp(y-1,1)%>">					
				</td>
			<%next%>
		<%else%>	
			<td>
				<input name=fensuA class=inputbox8 size=4 type=hidden value=0>
				<input name=fensuB class=inputbox8 size=4 type=hidden value=0>
				<input name=fensuC class=inputbox8 size=4 type=hidden value=0>
				<input name=fensuD class=inputbox8 size=4 type=hidden value=0>
				<input name=monthweek class=inputbox8 size=4 type=hidden value="">
				<input name=calcym class=inputbox8 size=4 type=hidden value="">
			</td>
		<%end if%>
	</tr>				
	<%next%> 
		<input type=hidden name=empid  value="">
	<input name=fensuA class=inputbox8 size=4 type=hidden value=0>
	<input name=fensuB class=inputbox8 size=4 type=hidden value=0>
	<input name=fensuC class=inputbox8 size=4 type=hidden value=0>
	<input name=fensuD class=inputbox8 size=4 type=hidden value=0>
	<input name=monthweek class=inputbox8 size=4 type=hidden value="">
	<input name=calcym class=inputbox8 size=4 type=hidden value="">
</table>	
<TABLE WIDTH=600>
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
function fnAcng(index) 	
	if trim(<%=self%>.tmw.value)<>"" then 
		maxfensu = 100/cdbl(<%=self%>.tmw.value)
	end  if
	if trim(<%=self%>.fensuA(index).value)<>"" then 
		if isnumeric(<%=self%>.fensuA(index).value)=false then  
			alert "請輸入數字!!hay nhap so vao!!"
			<%=self%>.fensuA(index).value="0"
			<%=self%>.fensuA(index).select()
			exit function  
		else 
			if cdbl(<%=self%>.fensuA(index).value)> cdbl(maxfensu) then 
				alert "分數超過 [" & maxfensu & "] 分"
				<%=self%>.fensuA(index).value="0"
				<%=self%>.fensuA(index).select()
				exit function  
			end if 	
			
		end if	
	end if 		
end function  

 
function fnBcng(index) 	  
	if trim(<%=self%>.fensuB(index).value)<>"" then 
		if isnumeric(<%=self%>.fensuB(index).value)=false then  
			alert "請輸入數字!!hay nhap so vao!!"
			<%=self%>.fensuB(index).value=""
			<%=self%>.fensuB(index).focus()
			exit function 	
		end if	
	end if 	
end function 
 
function fnCcng(index) 	  
	if trim(<%=self%>.fensuC(index).value)<>"" then 
		if isnumeric(<%=self%>.fensuC(index).value)=false then  
			alert "請輸入數字!!hay nhap so vao!!"
			<%=self%>.fensuC(index).value=""
			<%=self%>.fensuC(index).focus()
			exit function 	
		end if	
	end if 	
end function   

function fnDcng(index) 	  
	if trim(<%=self%>.fensuD(index).value)<>"" then 
		if isnumeric(<%=self%>.fensuD(index).value)=false then  
			alert "請輸入數字!!hay nhap so vao!!"
			<%=self%>.fensuD(index).value=""
			<%=self%>.fensuD(index).focus()
			exit function 	
		end if	
	end if 	
end function  
  
 
function  go()
	if <%=self%>.khym.value="" then 
		alert "考核年月不可為空!!"
		<%=self%>.khym.focus()
		<%=self%>.khym.select()
	end if 
	<%=self%>.action="<%=self%>.upd.asp"
	<%=self%>.submit()
end function 
  
</script>


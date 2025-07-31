<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<!-- #include file = "../GetSQLServerConnection.fun" -->
<!-- #include file="../ADOINC.inc" -->
<!--#include file="../include/sideinfo.inc"-->
<%

SELF = "YEIE0102"

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
F_shift=request("F_shift")
F_empid =request("empid")
F_country=request("F_country")
fclass = request("fclass")  
sortvalue = request("sortvalue") 
if sortvalue ="" then sortvalue="b.country , h.lw, h.lg, len(h.ls)desc, h.ls, a.empid"

FUNCTION FDT(D)
	Response.Write YEAR(D)&"/"&RIGHT("00"&MONTH(D),2)&"/"&RIGHT("00"&DAY(D),2)
END FUNCTION 
khym = request("khym")
if request("khym")="" then 
	khym=nowmonth
end if  

 
 '一個月有幾天 
cDatestr=CDate(LEFT(khym,4)&"/"&RIGHT(khym,2)&"/01") 
days = DAY(cDatestr+(32-DAY(cDatestr))-DAY(cDatestr+(32-DAY(cDatestr))))   '一個月有幾天  
'本月最後一天 
ENDdat = LEFT(khym,4)&"/"&RIGHT(khym,2)&"/"&DAYS   

'if khweek="" then khweek=(days\7)  

gTotalPage = 1
PageRec = 10    'number of records per page
TableRec = 20    'number of fields per record   


sql="select  b.country cstr, b.empnam_cn, b.empnam_vn, b.country, convert(char(10),b.indat,111) as nindat,convert(char(10),outdat,111) as outdate, d.sys_value as gstr, "&_
	"e.sys_value as zstr, f.sys_value as wstr, g.sys_value as sstr, a.* from "&_
	"(  "&_
	"select count(*) as weekcnt, khym, empid, khw, khg, khz, khs , sum(fna+fnb+fnc+fnd ) as monthfen from  empkhb where khym='"& khym &"' "&_
	"group by khym,empid, khw, khg, khz, khs  "&_
	") a  "&_
	"left join ( select *from empfile ) b on b.empid = a.empid   "&_ 
	"left join ( select* from  basicCode  where func='groupid' ) d on d.sys_type = a.khg "&_
	"left join ( select* from  basicCode  where func='zuno' ) e on e.sys_type = a.khz "&_
	"left join ( select* from  basicCode  where func='whsno' ) f on f.sys_type = a.khw "&_
	"left join ( select* from  basicCode  where func='shift' ) g on g.sys_type = a.khs "&_	 
	"left join ( select * from view_empgroup where yymm='"& khym &"' ) h on h.empid = b.empid "&_
	"where b.country like '"& F_country &"%' and a.khw like '"&F_whsno &"%'  and a.khg like '"& F_groupid &"%' "&_
	"and a.khz like '"&F_zuno&"%' and a.khs like '%"&F_shift&"' and a.empid like '"&F_empid&"%' "&_
	"order by " & sortvalue 
	
 
'response.write sql 	
'response.end 			
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
				tmpRec(i, j, 2) = rds("khw")
				tmpRec(i, j, 3) = rds("khg")
				tmpRec(i, j, 4) = rds("khz")
				tmpRec(i, j, 5) = rds("khs")
				tmpRec(i, j, 6) = rds("empnam_cn")
				tmpRec(i, j, 7) = rds("empnam_vn")
				tmpRec(i, j, 8) = rds("nindat")
				tmpRec(i, j, 9) = rds("outdate")
				tmpRec(i, j, 10) = rds("gstr")
				tmpRec(i, j, 11) = rds("zstr")
				tmpRec(i, j, 12) = rds("monthfen")
				tmpRec(i, j, 13) = rds("weekcnt") 
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
 

function datachg() 
	<%=self%>.totalpage.value="0"
	<%=self%>.sortvalue.value=""
	<%=self%>.action = "<%=self%>.ForeGnd.asp"
	<%=self%>.target="_self"
	<%=self%>.submit()
end function   

function datachg2() 
	<%=self%>.totalpage.value="0"	
	<%=self%>.act.value="A"
	<%=self%>.action = "<%=self%>.ForeGnd.asp"
	<%=self%>.target="_self"
	<%=self%>.submit()
end function 

function sortby(a)
	if a=1 then 
		<%=self%>.sortvalue.value="a.khz, a.empid"
	elseif a=2 then	
		<%=self%>.sortvalue.value="a.empid"
	elseif a=3 then	
		<%=self%>.sortvalue.value="b.nindat, a.empid"
	elseif a=4 then	
		<%=self%>.sortvalue.value="a.monthfen desc, a.empid"
	elseif a=5 then	
		<%=self%>.sortvalue.value="len(a.khs) desc, a.khs, a.khz, a.empid"
	else
		<%=self%>.sortvalue.value="b.country , a.khw, a.khg, a.empid"			
	end if 	 
	<%=self%>.totalpage.value="0"
	<%=self%>.action = "<%=self%>.ForeGnd.asp"
	<%=self%>.target="_self"
	<%=self%>.submit()
	'alert a 
end function 

 
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
<INPUT TYPE=hidden NAME=days VALUE="<%=days%>">
<INPUT TYPE=hidden NAME=sortvalue VALUE="<%=sortvalue%>">
<input name=act value="<%=act%>" type=hidden >

	<table border=0 style="height:70px"><tr><td>&nbsp;</td></tr></table>
	
	<table width="98%" BORDER=0 align=center cellpadding=0 cellspacing=0 align="center">
		<tr>
			<td >
				<table class="txt" cellpadding=3 cellspacing=3>
					<TR>		
						<TD align=right nowrap>考核年月</TD>
						<td>
							<table border=0 class="txt" cellpadding=3 cellspacing=3>
								<tr>
									<td><input type="text" style="width:100px" name=khym value="<%=khym%>" maxlength=6  onchange=datachg2()>
										<input name=khweek type=hidden>	
									</td>
									<td>
										<select name="fclass"  onchange=datachg2()style="width:120px" >
											<option value="A" <%if fclass="A" then%>selected<%end if%>>月統計</option>
											<option value="B" <%if fclass="B" then%>selected<%end if%>>週統計</option>
										</select>
									</td>
								</tr>
							</table>
						</td>
						<td align=right nowrap>國籍<BR><font class="txt8">Quoc tich</font></td>
						<td>
							<select name=F_country  onchange="datachg()" style="width:120px">									
								<option value="" selected >全部(Toan bo) </option>
								<%SQL="SELECT * FROM BASICCODE WHERE FUNC='country'  and sys_type not in ('Tw' ) ORDER BY SYS_TYPE desc "					
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
					</tr>
					<tr>	
						<TD align=right nowrap>廠別<br><font class="txt8">Xuong</font></TD>
						<td>
							<select name=F_whsno style="width:120px">					
									<%
									if session("rights")="0" then 
										SQL="SELECT * FROM BASICCODE WHERE FUNC='WHSNO' ORDER BY SYS_TYPE "%>
										<option value="" selected >全部(Toan bo) </option>
									<%		
									else
										SQL="SELECT * FROM BASICCODE WHERE FUNC='WHSNO'  ORDER BY SYS_TYPE " 'and sys_type='"& session("NETWHSNO") &"'
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
						<TD align=right nowrap>部門<br><font class="txt8">Bo Phan</font></TD>
						<td>
							<table border=0>
								<tr>
									<td>
										<select name=F_groupid   onchange="datachg()" style="width:120px">			
											<option value="">全部(Toan bo) </option>
											<% 
											SQL="SELECT * FROM BASICCODE WHERE FUNC='GROUPID' and sys_type <>'AAA' ORDER BY SYS_TYPE " 
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
									</td>
									<td>
										<select name=F_zuno    onchange="datachg()" style="width:120px">				
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
								</tr>
							</table>
						</td>
						<td align=right nowrap>班別<BR><font class="txt8">Ca</font></td>
						<td>
							<select name=F_shift   onchange="datachg()" style="width:100px" >
								<option value=""></option>
								<option value="ALL" <%if F_shift="ALL" then%>selected<%end if%>>日</option>
								<option value="A" <%if F_shift="A" then%>selected<%end if%>>A班</option>
								<option value="B" <%if F_shift="B" then%>selected<%end if%>>B班</option>
							</select>										
						</td>
						<td><input type=button name=send value="(S)查詢" class="btn btn-sm btn-outline-secondary" onclick="datachg()"></td>
					</TR>	 
				</table>
			</td>
		</tr>
		<tr>
			<td >
				<table id="myTableGrid" width="98%">
					<tr class="header">
						<Td  nowrap align=center rowspan=2 >STT</td>
						<Td  nowrap align=center rowspan=2 >部門</td>
						<Td  nowrap align=center rowspan=2  style='cursor: pointer;' onclick=sortby(1) title='依單位排序' >單位<br><img src="../picture/soryby.gif"></td>
						<Td  nowrap align=center rowspan=2 style='cursor:pointer;' onclick=sortby(5) title='依班別排序'>班別<br><img src="../picture/soryby.gif"></td>
						<Td  nowrap align=center rowspan=2 style='cursor:pointer;' onclick=sortby(2) title='依工號排序'>工號<br><img src="../picture/soryby.gif"></td>
						<Td  nowrap align=center rowspan=2 >姓名</td>
						<Td  nowrap  align=center rowspan=2 style='cursor:pointer;' onclick=sortby(3) title='依到職日排序'>到職日<br><img src="../picture/soryby.gif"></td>
						<%if khweek="" then %>
							<%for a = 1 to (days\7)  
								if a=1 then 
									yd1=a
								else
									yd1=((a-1)*7)+1
								end if	
								yd2=a*7
								if a=4 and yd2<days then 
									yd2=days 
								end if 	 
								dd1 = right(khym,2)&"/"&right("00"&yd1,2)
								dd2 = right(khym,2)&"/"&right("00"&yd2,2)			
							%> 
								<Td align=center   nowrap colspan=4>
									第<%=a%>週<BR>
									<%=DD1%> ~ <%=DD2%>
								</td>
							<%next%>	
						<%else 
							if khweek="1" then 
								yd1=khweek
							else
								yd1=((cdbl(khweek)-1)*7)+1
							end if	
							yd2=cdbl(khweek)*7
							if cdbl(khweek)=4 and yd2<days then 
								yd2=days 
							end if 	 
							dd1 = right(khym,2)&"/"&right("00"&yd1,2)
							dd2 = right(khym,2)&"/"&right("00"&yd2,2)	
						%>
							<Td align=center   nowrap colspan=4 >
								第<%=khweek%>週<BR>	
								<%=DD1%> ~ <%=DD2%>			
							</td>	
						<%end if%>
						<td rowspan=2 align=center style='cursor:pointer;' onclick=sortby(4) width=30 nowrap title='依總分排序'>總分<br><img src="../picture/soryby.gif"></td>
						<td rowspan=2 align=center width=100 nowrap>備註</td>
						<td rowspan=2 align=center width=50 nowrap>輸入者</td>
					</tr> 
					<tr class="header">		
						<%
						if khweek<>"" then 
							tz=1
						else
							tz=days\7	
						end if
						for g = 1 to tz%>
							<%if fclass="B"  then %>
							<td align=center height=20 width=25 nowrap>A </td>
							<td align=center width=25 nowrap>B </td>
							<td align=center width=25 nowrap>C </td>
							<td align=center width=25 nowrap >D </td>
							<%else%>
								<td colspan=4 align=center>月統計</td>
							<%end if %>	
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
					 
					<TR> 	 
						<Td align=center><%=CurrentRow%></td>
						<Td align=center><%=tmpRec(CurrentPage, CurrentRow, 10)%></td>
						<td align=left><%=tmpRec(CurrentPage, CurrentRow, 11)%></td>
						<Td align=center><%=tmpRec(CurrentPage, CurrentRow, 5)%></td>
						<Td align=center >
							<%=tmpRec(CurrentPage, CurrentRow, 1)%>
							<input type=hidden name=F_empid  value="<%=tmpRec(CurrentPage, CurrentRow, 1)%>">
						</td>
						<Td  style="cursor:pointer" nowrap onclick="vbscript:oepnEmpWKT(<%=currentRow-1%>)" title="**點選可看出勤紀錄**">
							<%=tmpRec(CurrentPage, CurrentRow, 6)%><br><%=left(tmpRec(CurrentPage, CurrentRow, 7),15)%>
						</td>
						<Td align=center title='紅色表示員工離職日'><%=tmpRec(CurrentPage, CurrentRow, 8)%><br><font color=red><%=tmpRec(CurrentPage, CurrentRow, 9)%></font></td>
						<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then %>
							<%
							  tfn = 0 
							  if khweek="" then 
								tt = days\7  			  	
							  else
								tt = 1  			  	
							  end if 	
							  memo=""	
							  for y=1 to tt
								if y mod 2 =0 then 
									weekcolor="#e6e6fa"
								else
									weekcolor="#eee8aa"
								end if 	
								if  khweek="" then 
									sqld="select * from empKHB  where khym='"& khym &"' and empid='"& tmpRec(CurrentPage, CurrentRow, 1) &"' and khweek='"&y&"' " 
								else
									sqld="select * from empKHB  where khym='"& khym &"' and empid='"& tmpRec(CurrentPage, CurrentRow, 1) &"' and khweek='"&khweek&"' "
								end if 	
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
									memo=memo&""
									muser=muser&""
								else					
									fnA=rs2("fnA")
									fnB=rs2("fnB")
									fnC=rs2("fnC")
									fnD=rs2("fnD")
									if rs2("memo")="" then 
										memo=memo&rs2("memo")
									else
										memo=memo&rs2("memo")&"<BR>"
									end if	
									if fna="0" then colorA="red" else colorA="black"
									if fnb="0" then colorB="red" else colorB="black"
									if fnc="0" then colorC="red" else colorC="black"
									if fnd="0" then colorD="red" else colorD="black"										
									muser=rs2("muser")					
								end if 				
								
								tfn = tfn + cdbl(fna)+cdbl(fnb)+cdbl(fnc)+cdbl(fnd)
							%>
								<%if fclass="B"  then %>
									<td bgcolor="<%=weekcolor%>" align=center><%=fnA%>
										<input type=hidden name=fensuA 8r size=3  onblur="fnAcng(<%=(CurrentRow-1)*tt+y-1%>)" value="<%=fnA%>"  style='color:<%=colorA%>'  >
									</td>
									<td bgcolor="<%=weekcolor%>" align=center><%=fnB%>	
										<input type=hidden name=fensuB 8r size=3 onblur="fnBcng(<%=(CurrentRow-1)*tt+y-1%>)" value="<%=fnB%>" style='color:<%=colorB%>' >
									</td>	
									<td bgcolor="<%=weekcolor%>" align=center><%=fnC%>
										<input type=hidden name=fensuC 8r size=3 onblur="fnCcng(<%=(CurrentRow-1)*tt+y-1%>)" style='color:<%=colorC%>' value="<%=fnC%>">
									</td>
									<td bgcolor="<%=weekcolor%>" align=center><%=fnD%>	
										<input type=hidden name=fensuD 8r size=3 onblur="fnDcng(<%=(CurrentRow-1)*tt+y-1%>)" style='color:<%=colorD%>' value="<%=fnD%>">				
										
										<input type=hidden name=monthweek  value="<%=y%>">
										<input type=hidden name=calcym  value="<%=khym%>">					
									</td> 
								<%else%>
									<Td colspan=4 align=center><%=cdbl(fna)+cdbl(fnb)+cdbl(fnc)+cdbl(fnd)%></td>	
								<%end if%>
							<%next%>
							<td  align=center><%=tfn%>
								<input type=hidden name=tfn  readonly value="<%=tfn%>" size=3 style='background-color:#e4e4e4'>
							</td>
							<td><%=memo%>
								<input type=hidden name=memo    value="<%=memo%>" size=15 >
							</td>
							<td align=center><%=muser%></td>
						<%else%>							
							<td colspan=4></td>
							<td colspan=4></td>
							<td colspan=4></td>							
							<td colspan=4></td>							
							<td></td>
							<td></td>
							<td>
								<input name=fensuA  size=4 type=hidden value=0>
								<input name=fensuB  size=4 type=hidden value=0>
								<input name=fensuC  size=4 type=hidden value=0>
								<input name=fensuD  size=4 type=hidden value=0>
								<input name=monthweek  size=4 type=hidden value="">
								<input name=calcym  size=4 type=hidden value="">
								<input name=tfn  size=4 type=hidden value=0>
								<input name=memo  size=4 type=hidden value="">
							</td>
						<%end if%>
					</tr>				
					<%next%> 
						<input type=hidden name=empid  value="">
						<input name=fensuA  size=4 type=hidden value=0>
						<input name=fensuB  size=4 type=hidden value=0>
						<input name=fensuC  size=4 type=hidden value=0>
						<input name=fensuD  size=4 type=hidden value=0>
						<input name=monthweek  size=4 type=hidden value="">
						<input name=calcym  size=4 type=hidden value="">
				</table>
			</td>
		</tr>
		<tr>
			<td align="center">
				<table class="table-borderless table-sm text-secondary txt">
					<tr>
						<TD >
							<input type=button  name=send value="(M)回主畫面" class="btn btn-sm btn-outline-secondary" onclick=backM()>
							<input type=button  name=send value="下載到Excel" class="btn btn-sm btn-outline-secondary" onclick=goexcel()>
						</TD>
					</TR>
				</table>
			</td>
		</tr>
	</table>
			
</form>


</body>
</html>
<script language=vbscript> 
function oepnEmpWKT(index)
	empidstr = <%=self%>.F_empid(index).value
	yymmstr = <%=self%>.khym.value
	khweekstr = "" '<%=self%>.khweek.value
	open "../yed/yedq01.Foregnd.asp?fr=A&yymm="& yymmstr & "&empid=" & empidstr &"&khweek=" & khweekstr , "_blank", "top=10 , left=10, height=500, width=700,scrollbars=yes"
end function

function fnAcng(index) 	
	maxfensu = 25
	if trim(<%=self%>.fensuA(index).value)<>"" then 
		if isnumeric(<%=self%>.fensuA(index).value)=false then  
			alert "請輸入數字!!hay nhap so vao!!"
			<%=self%>.fensuA(index).value="0"
			<%=self%>.fensuA(index).select()
			exit function  
		else 
			if cdbl(<%=self%>.fensuA(index).value)> cdbl(maxfensu) then 
				alert "分數超過 [25] 分"
				<%=self%>.fensuA(index).value="0"
				<%=self%>.fensuA(index).select()
				exit function  
			end if 	 
			calctfn(index)
		end if	
	end if 		
end function  

function goexcel()
	<%=self%>.action = "<%=self%>.toexcel.asp"
	<%=self%>.target="Back"
	'parent.best.cols="50%,50%"
	<%=self%>.submit()
	
end function 
 
function fnBcng(index) 	  
	maxfensu = 25
	if trim(<%=self%>.fensuB(index).value)<>"" then 
		if isnumeric(<%=self%>.fensuB(index).value)=false then  
			alert "請輸入數字!!hay nhap so vao!!"
			<%=self%>.fensuB(index).value=""
			<%=self%>.fensuB(index).focus()
			exit function 	
		else
			if cdbl(<%=self%>.fensuB(index).value)> cdbl(maxfensu) then 
				alert "分數超過 [25] 分"
				<%=self%>.fensuB(index).value="0"
				<%=self%>.fensuB(index).select()
				exit function  
			end if 						
			calctfn(index)
		end if	
	end if 	
end function 
 
function fnCcng(index) 	  
	maxfensu = 25
	if trim(<%=self%>.fensuC(index).value)<>"" then 
		if isnumeric(<%=self%>.fensuC(index).value)=false then  
			alert "請輸入數字!!hay nhap so vao!!"
			<%=self%>.fensuC(index).value=""
			<%=self%>.fensuC(index).focus()
			exit function 	
		else
			if cdbl(<%=self%>.fensuC(index).value)> cdbl(maxfensu) then 
				alert "分數超過 [25] 分"
				<%=self%>.fensuC(index).value="0"
				<%=self%>.fensuC(index).select()
				exit function  
			end if 			
			calctfn(index)
		end if	
	end if 	
end function   

function fnDcng(index) 
	maxfensu = 25	  
	if trim(<%=self%>.fensuD(index).value)<>"" then 
		if isnumeric(<%=self%>.fensuD(index).value)=false then  
			alert "請輸入數字!!hay nhap so vao!!"
			<%=self%>.fensuD(index).value=""
			<%=self%>.fensuD(index).focus()
			exit function 	
		else
			if cdbl(<%=self%>.fensuD(index).value)> cdbl(maxfensu) then 
				alert "分數超過 [25] 分"
				<%=self%>.fensuD(index).value="0"
				<%=self%>.fensuD(index).select()
				exit function  
			end if 	
			calctfn(index)			
		end if	
	end if 	
end function  

function calctfn(index)	
	'alert index
	if <%=self%>.khweek.value="" then 
		A1 = round(cdbl(<%=self%>.days.value)\7,0)
	else
		A1=1 
	end if 
'	c_tfn = 0 	
	for A2 = 1 to  A1 		
		'alert (index\A1)*A1+A2-1
		C_fna = (<%=self%>.fensuA((index\A1)*A1+A2-1).value)
		C_fnb = (<%=self%>.fensuB((index\A1)*A1+A2-1).value)
		C_fnc = (<%=self%>.fensuC((index\A1)*A1+A2-1).value)
		C_fnd = (<%=self%>.fensuD((index\A1)*A1+A2-1).value)
		c_tfn = c_tfn + cdbl(c_fnA)+ cdbl(c_fnB)+ cdbl(c_fnC)+ cdbl(c_fnD)
	next  
	<%=self%>.tfn(index\A1).value = c_tfn
end function 
 
function  backM()	
	open "<%=self%>.asp", "_self"
	
end function 
  
</script>


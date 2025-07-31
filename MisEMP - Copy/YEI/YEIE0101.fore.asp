<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<!-- #include file = "../GetSQLServerConnection.fun" -->
<!-- #include file="../ADOINC.inc" -->
<!--#include file="../include/sideinfo.inc"-->
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

sortvalue = request("sortvalue")
if sortvalue ="" then sortvalue="a.country , b.lw, b.lg, a.empid"


FUNCTION FDT(D)
	Response.Write YEAR(D)&"/"&RIGHT("00"&MONTH(D),2)&"/"&RIGHT("00"&DAY(D),2)
END FUNCTION
khym = request("khym")
if request("khym")="" then
	khym=nowmonth
end if

act = request("act")
khweek = request("khweek")
if khweek="" then khweek="1"

tmw = request("tmw")
if tmw="" then tmw=request("tt")
 '一個月有幾天
cDatestr=CDate(LEFT(khym,4)&"/"&RIGHT(khym,2)&"/01")
days = DAY(cDatestr+(32-DAY(cDatestr))-DAY(cDatestr+(32-DAY(cDatestr))))   '一個月有幾天
'本月最後一天
ENDdat = LEFT(khym,4)&"/"&RIGHT(khym,2)&"/"&DAYS

'if khweek="" then khweek=(days\7)

gTotalPage = 1
PageRec = 10    'number of records per page
TableRec = 25    'number of fields per record

sql="select lw, lg, lz, ls, x1.sys_value  lgstr, x2.sys_value lzstr, lw lwstr, ls lsstr, a.* , isnull(c.aid,'') aid, isnull(c.status,'') hjsts, "&_
	"isnull(c.cqmemo,'') cqmemo ,isnull(c.fnA,0) fnA, isnull(c.fnB,0) fnB,isnull(c.fnC,0) fnC,isnull(c.fnD,0) fnD, "&_
	"isnull(fna,0)+isnull(fnb,0)+isnull(fnc,0)+isnull(fnd,0)  as totfen, isnull(c.memo,'') as khmemo  from "&_
	"(select *from  view_empfile ) a "&_	
	"left join (select empid, whsno as lw, groupid as lg, shift as ls, zuno as lz  from bempg where  yymm='"&khym &"'  ) b on b.empid= a.empid "&_ 
	"left join (select * from basicCode where func='groupid' ) x1 on x1.sys_type = b.lg "&_
	"left join (select * from basicCode where func='zuno' ) x2 on x2.sys_type = b.lz "&_
	"left join (select * from empkhb where khym='"& khym &"' and khweek='"& khweek &"' ) c on c.empid = a.empid  "&_
	"where convert(char(10), a.indat, 111)<='"& cDatestr &"' "&_
	"and (isnull(outdat,'')='' or convert(char(10),outdat,111)>='"& ENDdat &"') "&_
	"and isnull(b.lw,a.whsno)='"& F_whsno &"' and isnull(b.lg,a.groupid) like '"&F_groupid&"%' "&_
	"and a.country like '"& F_country &"%' and isnull(b.lz,a.zuno) like '"&F_zuno&"%' "&_
	"and isnull(b.ls,a.shift) like '%"& F_shift &"'  "&_
	"order by " & sortvalue

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
				tmpRec(i, j, 2) = rds("lw")
				tmpRec(i, j, 3) = rds("lg")
				tmpRec(i, j, 4) = rds("lz")
				tmpRec(i, j, 5) = rds("ls")
				tmpRec(i, j, 6) = rds("empnam_cn")
				tmpRec(i, j, 7) = rds("empnam_vn")
				tmpRec(i, j, 8) = rds("nindat")
				tmpRec(i, j, 9) = rds("outdate")
				tmpRec(i, j, 10) = rds("lgstr")
				tmpRec(i, j, 11) = rds("lzstr")

				tmpRec(i, j, 12) = rds("aid")
				tmpRec(i, j, 13) = rds("fna")
				tmpRec(i, j, 14) = rds("fnb")
				tmpRec(i, j, 15) = rds("fnc")
				tmpRec(i, j, 16) = rds("fnd")
				tmpRec(i, j, 17) = rds("totfen")
				tmpRec(i, j, 18) = rds("khmemo")
				tmpRec(i, j, 19) = rds("hjsts")
				tmpRec(i, j, 20) = rds("cqmemo")
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
</head>
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

function sortby(a)
	if a=1 then
		<%=self%>.sortvalue.value="b.khz, a.empid"
	elseif a=2 then
		<%=self%>.sortvalue.value="a.empid"
	elseif a=3 then
		<%=self%>.sortvalue.value="a.nindat, a.empid"
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
function viewuse()
	open "useNote.asp", "_balnk", "top=10, left=10, width=750,scrollbars=yes"
end function

</SCRIPT>
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
<input name=HT value="" type=hidden >

	<table border=0 style="height:70px"><tr><td>&nbsp;</td></tr></table>	
	<table width="100%" BORDER=0 align=center cellpadding=0 cellspacing=0 >
		<tr>
			<td>
				<table class="txt" cellpadding=3 cellspacing=3 >
					<TR height=22 >
						<TD align=right nowrap>考核年月</TD>
						<td colspan=9>
							<table class="txt" cellpadding=3 cellspacing=3>
								<tr>
									<td>
										<input type="text" style="width:100px" name=khym value="<%=khym%>" maxlength=6  onchange=datachg2()>
									</td>
									<td>
										<select name=khweek   onchange=datachg2() style="width:180px">
										<%for yy =1 to (days\7)
											if yy=1 then
												yd1=yy
											else
												yd1=((yy-1)*7)+1
											end if
											yd2=yy*7
											if yy=4 and yd2<days then
												yd2=days
											end if
											dd1 = right(khym,2)&"/"&right("00"&yd1,2)
											dd2 = right(khym,2)&"/"&right("00"&yd2,2)

										%>
											<option value="<%=yy%>" <%if cstr(yy) = cstr(khweek) then%>selected<%end if%>>第<%=yy%>週,<%=dd1%>~<%=dd2%> </option>
										<% next	%>
										</select>
									</td>
									<td align=center nowrap><a href="vbscript:viewuse()" >操作說明</a></td>
								</tr>
							</table>
						</td>
					</tr>
					<tr>
						<td align=right nowrap>國籍 <% response.write session("NETWHSNO") %><BR>Quoc tich</td>
						<td>
							<select name=F_country  style='width:70' >
								<%
								if session("rights")<>"" then
									SQL="SELECT * FROM BASICCODE WHERE FUNC='country'  ORDER BY SYS_TYPE  desc"%>
									<option value="" selected >全部(Toan bo) </option>
								<%
								else
									SQL="SELECT * FROM BASICCODE WHERE FUNC='country'  and sys_type not in ('Tw' ) ORDER BY SYS_TYPE desc "
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
						<TD align=right nowrap>廠別<br>Xuong</TD>
						<td>
							<select name=F_whsno  style='width:100' >
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
						<TD align=right nowrap>部門<br>Bo Phan</TD>
						<td>
							<select name=F_groupid    style='width:80' >
								<%if Session("RIGHTS")<="2" or Session("RIGHTS")="5" then%>
									<option value="">全部(Toan bo) </option>
									<%
										SQL="SELECT * FROM BASICCODE WHERE FUNC='GROUPID' and sys_type <>'AAA' ORDER BY SYS_TYPE "
									else
										SQL="SELECT * FROM BASICCODE WHERE FUNC='GROUPID' and sys_type <>'AAA' and  sys_type= '"& session("NETG1") &"' ORDER BY SYS_TYPE "
									end if

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
								<%SET RST=NOTHING %>
							</SELECT>
						</td>
						<td>
							<select name=F_zuno   style='width:70'    >
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
							<select name=F_shift     >
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
			<td align="center">
				<table id="myTableGrid" width="98%">
					<tr class="header">
						<td nowrap >審核</td>
						<Td nowrap rowspan=2 >STT</td>
						<Td nowrap rowspan=2 >部門</td>
						<Td nowrap rowspan=2 >單位</td>
						<Td nowrap rowspan=2 >班別</td>
						<Td nowrap rowspan=2  >工號</td>
						<Td nowrap rowspan=2 >姓名</td>
						<Td nowrap rowspan=2 >到職日</td>
						<%if khweek="" then %>
							<%for a = 1 to (days\7)
								if a=1 then
									yd1=a
								else
									yd1=((yy-1)*7)+1
								end if
								yd2=a*7
								if a=4 and yd2<days then
									yd2=days
								end if
								dd1 = right(khym,2)&"/"&right("00"&yd1,2)
								dd2 = right(khym,2)&"/"&right("00"&yd2,2)
							%>
								<Td align=center   nowrap colspan=4>
									<b>第<%=a%>週</b><BR>
									<b><%=DD1%> ~ <%=DD2%></b>
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
							<Td align=center   nowrap colspan=4>
								<b>第 <font color=blue><%=khweek%></font> 週<BR>
								<%=DD1%> ~ <%=DD2%></b>
							</td>
						<%end if%>
						<td rowspan=2 align=center >總分</td>
						<td rowspan=2 align=center>備註</td>
						<td rowspan=2 align=center>廠主管/經理評核</td>
					</tr>
					<tr class="header">
						<%if session("rights")="0" or session("rights")="5" then%>
							<td style='cursor:hand' onclick="selectall()" align=center height=22 nowrap >
								<font color=blue>全選</font>
							</td>
						<%else%>
							<Td  align=center height=22 nowrap ></td>
						<%end if%>
						<%
						if khweek<>"" then
							tz=1
						else
							tz=days\7
						end if
						for g = 1 to tz%>
						<td align=center>A</td>
						<td align=center>B</td>
						<td align=center>C</td>
						<td align=center>D</td>
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
						<%if session("rights")="0" or session("rights")="5" then%>
							<td align=center>
								<%if tmpRec(CurrentPage, CurrentRow, 19)="Y" then %>
									<font color=red>OK</font>
									<input type=hidden name=func type=checkbox >
									<input type=hidden  name=hjsts  size=4 value="">
								<%else%>
									<input name=func type=checkbox onclick="funcChg(<%=currentRow-1%>)">
									<input type=hidden  name=hjsts  size=1 value="">
								<%end if%>
							</td>
						<%else%>
							<td align=center>
								<%if tmpRec(CurrentPage, CurrentRow, 19)="Y" then %>
									<font color=red>OK</font>
								<%end if%>
								<input type=hidden name=func type=checkbox >
								<input type=hidden  name=hjsts  size=4 value="">
							</td>
						<%end if%>
						<Td align=center><%=CurrentRow%></td>
						<Td align=center><%=tmpRec(CurrentPage, CurrentRow, 10)%></td>
						<td align=center><%=tmpRec(CurrentPage, CurrentRow, 11)%></td>
						<Td align=center><%=tmpRec(CurrentPage, CurrentRow, 5)%></td>
						<Td align=center>
							<%=tmpRec(CurrentPage, CurrentRow, 1)%>
							<input type=hidden name=empid  value="<%=tmpRec(CurrentPage, CurrentRow, 1)%>">
							<input type=hidden name=whsno  value="<%=tmpRec(CurrentPage, CurrentRow, 2)%>">
							<input type=hidden name=groupid  value="<%=tmpRec(CurrentPage, CurrentRow, 3)%>">
							<input type=hidden name=zuno  value="<%=tmpRec(CurrentPage, CurrentRow, 4)%>">
							<input type=hidden name=shift  value="<%=tmpRec(CurrentPage, CurrentRow, 5)%>">
						</td>
						<Td nowrap style="cursor:hand" onclick="vbscript:oepnEmpWKT(<%=currentRow-1%>)" title="**點選可看出勤紀錄**"><%=tmpRec(CurrentPage, CurrentRow, 6)%><br><%=left(tmpRec(CurrentPage, CurrentRow, 7),15)%></td>
						<Td align=center><%=tmpRec(CurrentPage, CurrentRow, 8)%><br><font color=red><%=tmpRec(CurrentPage, CurrentRow, 9)%></font></td>
						<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then
							 fnA=tmpRec(CurrentPage, CurrentRow, 13)
							 fnB=tmpRec(CurrentPage, CurrentRow, 14)
							 fnC=tmpRec(CurrentPage, CurrentRow, 15)
							 fnD=tmpRec(CurrentPage, CurrentRow, 16)
							 memo=tmpRec(CurrentPage, CurrentRow, 18)
							 cqmemo=tmpRec(CurrentPage, CurrentRow, 20)
							 tfn=tmpRec(CurrentPage, CurrentRow, 17)
							 if fna="0" then colorA="red" else colorA="black"
							 if fnb="0" then colorB="red" else colorB="black"
							 if fnc="0" then colorC="red" else colorC="black"
							 if fnd="0" then colorD="red" else colorD="black"
							 if tmpRec(CurrentPage, CurrentRow, 19)="Y" then
								bgcolor2="lightyellow"
							 else
								bgcolor2="#ffffff"
							 end if
						%>
							<td bgcolor="<%=weekcolor%>">
								<input type=hidden name=aid class=readonly8 size=2 value="<%=tmpRec(CurrentPage, CurrentRow, 12)%>">
								<input type="text" style="width:100%" name=fensuA  value="<%=fnA%>"  style='background-color:<%=bgcolor2%>;color:<%=colorA%>' <%if tmpRec(CurrentPage, CurrentRow, 19)="Y" then %> readonly <%else%> onblur="fnAcng(<%=(CurrentRow-1)%>)" <%end if%> >
							</td>
							<td bgcolor="<%=weekcolor%>">
								<input type="text" style="width:100%" name=fensuB  value="<%=fnB%>" style='background-color:<%=bgcolor2%>;color:<%=colorB%>' <%if tmpRec(CurrentPage, CurrentRow, 19)="Y" then %> readonly <%else%> onblur="fnBcng(<%=(CurrentRow-1)%>)" <%end if%>>
							</td>
							<td bgcolor="<%=weekcolor%>">
								<input type="text" style="width:100%" name=fensuC  value="<%=fnC%>" style='background-color:<%=bgcolor2%>;color:<%=colorC%>' <%if tmpRec(CurrentPage, CurrentRow, 19)="Y" then %> readonly <%else%> onblur="fnCcng(<%=(CurrentRow-1)%>)" <%end if%>>
							</td>
							<td bgcolor="<%=weekcolor%>">
								<input type="text" style="width:100%" name=fensuD  value="<%=fnD%>" style='background-color:<%=bgcolor2%>;color:<%=colorD%>' <%if tmpRec(CurrentPage, CurrentRow, 19)="Y" then %> readonly <%else%> onblur="fnDcng(<%=(CurrentRow-1)%>)" <%end if%>>
								<input type=hidden name=monthweek  value="<%if khweek="" then%><%=y%><%else%><%=khweek%><%end if%>">
								<input type=hidden name=calcym  value="<%=khym%>">
							</td>
							<td><input type="text" style="width:100%" name=tfn  readonly value="<%=tfn%>" style='background-color:#e4e4e4'></td>
							<td><input type="text" style="width:100%" name=memo    value="<%=memo%>" style='height:22'></td>
							<td><input type="text" style="width:100%" name=cqmemo  <%if session("rights")="0" or session("rights")="5" then%><%else%>readonly<%end if%>   value="<%=cqmemo%>" style='height:22'></td>
						<%else%>
							<td></td>
							<td></td>
							<td></td>
							<td></td>
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
								<input name=func type=hidden  >
								<input name=hjsts  size=4 type=hidden value="">
								<input name=cqmemo  size=4 type=hidden value="">
								<input name=aid type=hidden value="">
							</td>
						<%end if%>
					</tr>
					<%next%>
					<input type=hidden name=empid  value="">
					<input name=fensuA 8 size=4 type=hidden value=0>
					<input name=fensuB 8 size=4 type=hidden value=0>
					<input name=fensuC 8 size=4 type=hidden value=0>
					<input name=fensuD 8 size=4 type=hidden value=0>
					<input name=tfn 8 size=4 type=hidden value=0>
					<input name=monthweek 8 size=4 type=hidden value="">
					<input name=calcym 8 size=4 type=hidden value="">
					<input name=func type=hidden  >
					<input name=hjsts 8 size=4 type=hidden value="">
					<input name=cqmemo 8 size=4 type=hidden value="">
					<input name=aid type=hidden value="">
				</table>
			</td>
		</tr>
		<tr>
			<td align="center" height=40px>
				<table class="table-borderless table-sm text-secondary">
					<tr ALIGN=center>
						<TD >
						<%if session("rights")<>"5" then %>
							<input type=button  name=send value="(Y)確　　認"  class="btn btn-sm btn-danger" onclick=go()>
							<input type=RESET name=send value="(N)取 　　消"  class="btn btn-sm btn-outline-secondary">
						<%end if %>
						<%if session("rights")="0" or session("rights")="5" then%>
							<input type=button name=send value="主 管 核 準"  class="btn btn-sm btn-danger" onclick=go2() style='background-color:#ffb6c1'>
						<%end if%>
						</TD>
					</TR>
				</TABLE>
			</td>
		</tr>
	</table>
			
</form>


</body>
</html>
<script language=vbscript>
function oepnEmpWKT(index)
	empidstr = <%=self%>.empid(index).value
	yymmstr = <%=self%>.khym.value
	khweekstr = <%=self%>.khweek.value
	open "../yed/yedq01.Foregnd.asp?fr=A&yymm="& yymmstr & "&empid=" & empidstr &"&khweek=" & khweekstr , "_blank", "top=10 , left=10, height=500, width=700,scrollbars=yes"
end function

function selectall()
	for h = 1 to <%=self%>.pagerec.value
		if <%=self%>.func(h-1).checked=true then
			<%=self%>.func(h-1).checked=false
			<%=self%>.hjsts(h-1).value=""
			eid = <%=self%>.empid(h-1).value
			if <%=self%>.aid(h-1).value="" then
				errmsg = errmsg & "工號: " & eid & " ,該部門/單位尚未輸入分數!!無法審核!!" &chr(13)
				<%=self%>.func(h-1).checked=false
				<%=self%>.hjsts(h-1).value=""
			end if
		else
			<%=self%>.func(h-1).checked=true
			<%=self%>.hjsts(h-1).value="Y"
			eid = <%=self%>.empid(h-1).value
			if <%=self%>.aid(h-1).value="" then
				errmsg = errmsg & "工號: " & eid & " ,該部門/單位尚未輸入分數!!無法審核!!" &chr(13)
				<%=self%>.func(h-1).checked=false
				<%=self%>.hjsts(h-1).value=""
			end if
		end if
	next
	if len(errmsg)> 0 then
		alert errmsg
	end if
end function

function funcChg(index)
	eid = <%=self%>.empid(index).value
	if <%=self%>.func(index).checked=true then
		<%=self%>.hjsts(index).value="Y"
		if <%=self%>.aid(index).value="" then
			alert "工號: " & eid & " ,該部門/單位尚未輸入分數!!無法審核!!"
			<%=self%>.func(index).checked=false
			<%=self%>.hjsts(index).value=""
		end if
	else
		<%=self%>.hjsts(index).value=""
	end if
end function


function fnAcng(index)
	maxfensu = 99  '25*0.5

	if trim(<%=self%>.fensuA(index).value)<>"" then
		if isnumeric(<%=self%>.fensuA(index).value)=false then
			alert "請輸入數字!!hay nhap so vao!!"
			<%=self%>.fensuA(index).value="0"
			<%=self%>.fensuA(index).select()
			exit function
		else
			if cdbl(<%=self%>.fensuA(index).value)> cdbl(maxfensu) then
				alert "分數超過 [ "& maxfensu &" ] 分"
				<%=self%>.fensuA(index).value="0"
				<%=self%>.fensuA(index).select()
				exit function
			end if
			calctfn(index)
		end if
	end if
end function


function fnBcng(index)
	maxfensu = 99 '25*0.2

	if trim(<%=self%>.fensuB(index).value)<>"" then
		if isnumeric(<%=self%>.fensuB(index).value)=false then
			alert "請輸入數字!!hay nhap so vao!!"
			<%=self%>.fensuB(index).value="0"
			<%=self%>.fensuB(index).focus()
			exit function
		else
			if cdbl(<%=self%>.fensuB(index).value)> cdbl(maxfensu) then
				alert "分數超過 [ "& maxfensu &" ] 分"
				<%=self%>.fensuB(index).value="0"
				<%=self%>.fensuB(index).select()
				exit function
			end if
			calctfn(index)
		end if
	end if
end function

function fnCcng(index)
	maxfensu = 99 '25*0.2

	if trim(<%=self%>.fensuC(index).value)<>"" then
		if isnumeric(<%=self%>.fensuC(index).value)=false then
			alert "請輸入數字!!hay nhap so vao!!"
			<%=self%>.fensuC(index).value="0"
			<%=self%>.fensuC(index).focus()
			exit function
		else
			if cdbl(<%=self%>.fensuC(index).value)> cdbl(maxfensu) then
				alert "分數超過 [ "& maxfensu &" ] 分"
				<%=self%>.fensuC(index).value="0"
				<%=self%>.fensuC(index).select()
				exit function
			end if
			calctfn(index)
		end if
	end if
end function

function fnDcng(index)
	maxfensu = 99 '25*0.1
	if trim(<%=self%>.fensuD(index).value)<>"" then
		if isnumeric(<%=self%>.fensuD(index).value)=false then
			alert "請輸入數字!!hay nhap so vao!!"
			<%=self%>.fensuD(index).value="0"
			<%=self%>.fensuD(index).focus()
			exit function
		else
			if cdbl(<%=self%>.fensuD(index).value)> cdbl(maxfensu) then
				alert "分數超過 [ "& maxfensu &" ] 分"
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
	if <%=self%>.khweek.value="" then
		if c_tfn>100 then
			alert  "本週分數超過100分!!"
			<%=self%>.tfn(index\A1).value = 0
		else
			<%=self%>.tfn(index\A1).value = c_tfn
		end if
	else
		if c_tfn>25 then
			alert  "本週分數超過25分!!"
			<%=self%>.tfn(index\A1).value = 0
		else
			<%=self%>.tfn(index\A1).value = c_tfn
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

function  go2()
	if <%=self%>.khym.value="" then
		alert "考核年月不可為空!!"
		<%=self%>.khym.focus()
		<%=self%>.khym.select()
	end if
	<%=self%>.hT.value="Y"
	<%=self%>.action="<%=self%>.upd.asp"
	<%=self%>.submit()
end function

</script>


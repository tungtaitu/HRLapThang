<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<!-- #include file = "../GetSQLServerConnection.fun" -->
<!-- #include file="../ADOINC.inc" -->
<!--#include file="../include/sideinfo.inc"-->
<%
'on error resume next
session.codepage="65001"
SELF = "YECQ01"

Set conn = GetSQLServerConnection()
Set rs = Server.CreateObject("ADODB.Recordset")

ym1 = request("yymm")
ym2 = request("yymm2")
whsno = trim(request("whsno"))
groupid = trim(request("groupid"))
COUNTRY = trim(request("COUNTRY"))
empid1 = trim(REQUEST("empid1"))

ccnt = cdbl(ym2) - cdbl(ym1)
if ccnt = 0 then ccnt = 1 

gTotalPage = 1
PageRec = 20*ccnt   'number of records per page
TableRec = 60    'number of fields per record
'NOWMONTH=CSTR(YEAR(DATE()))&"/"&RIGHT("00"&CSTR(MONTH(DATE())),2)&"/01"
NOWMONTH=CSTR(YEAR(DATE()))&"/"&RIGHT("00"&CSTR(MONTH(DATE())),2)&"/"&RIGHT("00"&CSTR(day(DATE())),2)

if empid1="" then 
	sql="select isnull(e.lj,'') lj , isnull(e.ljstr,'') ljstr ,  "&_
		"isnull(d.lw,'') lw , isnull(d.lg,'') lg , isnull(d.lz,'') lz , isnull(d.ls,'') ls , "&_
		"isnull(d.lwstr,'') lwstr , isnull(d.lgstr,'') lgstr , isnull(d.lzstr,'') lzstr, isnull(d.lsstr,'') lsstr, "&_
		"c.empnam_cn, c.empnam_vn, c.nindat, isnull(c.outdate,'') outdate, a.* , isnull(b.real_total,0) as backTotal , f.sys_value as Sjstr, isnull(g.totamt,0) wpamt  "&_
		"from "&_
		"( select * from empdsalary where  yymm  between '"& ym1 &"' and '"&ym2&"' and "&_
		"country like '"&COUNTRY&"%' and whsno like '"&whsno&"%' and groupid like '"&groupid&"%' and empid like '%"&empid1&"' ) a "&_
		"left join ( select * from empdsalary_bak   )  b on b.yymm = a.yymm and b.empid = a.empid  "&_
		"join (select empid, empnam_cn, empnam_vn, convert(char(10),indat,111)  nindat , isnull(convert(char(10),outdat,111),'')  outdate   from empfile ) c on c.empid = a.empid "&_
		"left join (select *from view_empgroup  ) d on d.empid = a.empid  and d.yymm = a.yymm   "&_
		"left join (select *from view_empjob) e on e.empid = a.empid and e.yymm = a.yymm "&_
		"left join (select * from basicCode) f on f.sys_type = a.job "&_
		"left join (select * from salarywp ) g on g.empid = a.empid and g.yymm = a.yymm " 
else
	sql="select isnull(e.lj,'') lj , isnull(e.ljstr,'') ljstr ,  "&_
		"isnull(d.lw,'') lw , isnull(d.lg,'') lg , isnull(d.lz,'') lz , isnull(d.ls,'') ls , "&_
		"isnull(d.lwstr,'') lwstr , isnull(d.lgstr,'') lgstr , isnull(d.lzstr,'') lzstr, isnull(d.lsstr,'') lsstr, "&_
		"c.empnam_cn, c.empnam_vn, c.nindat, c.outdate, a.* , isnull(b.real_total,0) as backTotal , f.sys_value as Sjstr , isnull(g.totamt,0) wpamt "&_
		"from "&_
		"( select * from empdsalary where empid like '%"&empid1&"' and country like '"&COUNTRY&"%' "&_
		"and whsno like '"&whsno&"%' and groupid like '"&groupid&"%'  ) a "&_
		"left join ( select * from empdsalary_bak   )  b on b.yymm = a.yymm and b.empid = a.empid  "&_
		"join (select empid, empnam_cn, empnam_vn, convert(char(10),indat,111)  nindat , isnull(convert(char(10),outdat,111),'')  outdate  from empfile ) c on c.empid = a.empid "&_
		"left join (select *from view_empgroup  ) d on d.empid = a.empid  and d.yymm = a.yymm   "&_
		"left join (select *from view_empjob) e on e.empid = a.empid and e.yymm = a.yymm "&_
		"left join (select *from basicCode) f on f.sys_type = a.job "&_
		"left join (select * from salarywp ) g on g.empid = a.empid and g.yymm = a.yymm " 
end if 		

sql = sql & "order by a.empid , a.yymm " 
'response.write sql
'response.end

if request("TotalPage") = "" or request("TotalPage") = "0" then
	CurrentPage = 1
	rs.Open sql, conn, 3, 3
	IF NOT RS.EOF THEN
		rs.PageSize = PageRec
		RecordInDB = rs.RecordCount
		TotalPage = rs.PageCount
		gTotalPage = TotalPage
	END IF

	Redim tmpRec(gTotalPage, PageRec, TableRec)   'Array

	for i = 1 to TotalPage
	 for j = 1 to PageRec
		if not rs.EOF then
				tmpRec(i, j, 0) = "no"
				tmpRec(i, j, 1) = trim(rs("empid"))
				tmpRec(i, j, 2) = trim(rs("empnam_cn"))
				tmpRec(i, j, 3) = trim(rs("empnam_vn"))
				tmpRec(i, j, 4) = rs("country")
				tmpRec(i, j, 5) = rs("nindat")
				tmpRec(i, j, 6) = rs("lj")
				tmpRec(i, j, 7) = rs("lw")
				tmpRec(i, j, 8) = rs("ls")
				tmpRec(i, j, 9)	=RS("lg")
				tmpRec(i, j, 10)=RS("lz")
				tmpRec(i, j, 11)=RS("lwstr")				
				tmpRec(i, j, 12)=RS("lgstr")
				tmpRec(i, j, 13)=RS("lzstr")
				tmpRec(i, j, 14)=RS("ljstr")
				tmpRec(i, j, 15)=RS("job")
				tmpRec(i, j, 16)=RS("whsno")
				tmpRec(i, j, 17)=RS("Sjstr") 				
				tmpRec(i, j, 18)=RS("BB")
				tmpRec(i, j, 19)=RS("CV")
				tmpRec(i, j, 20)=RS("PHU")
				tmpRec(i, j, 21)=RS("NN")
				tmpRec(i, j, 22)=RS("KT")
				tmpRec(i, j, 23)=RS("MT")
				tmpRec(i, j, 24)=RS("TTKH")
				tmpRec(i, j, 25)=RS("QC")
				tmpRec(i, j, 26)=RS("TNKH")
				tmpRec(i, j, 27)=RS("TBTR")
				tmpRec(i, j, 28)=RS("JX")
				tmpRec(i, j, 29)=RS("H1M")
				tmpRec(i, j, 30)=RS("H2M")
				tmpRec(i, j, 31)=RS("H3M")
				tmpRec(i, j, 32)=RS("B3M")
				tmpRec(i, j, 33)=RS("H1")
				tmpRec(i, j, 34)=RS("H2")
				tmpRec(i, j, 35)=RS("H3")
				tmpRec(i, j, 36)=RS("B3")
				tmpRec(i, j, 37)=RS("kzhour")
				tmpRec(i, j, 38)=RS("jiaA")
				tmpRec(i, j, 39)=RS("jiaB")
				tmpRec(i, j, 40)=RS("KZM")
				tmpRec(i, j, 41)=RS("jiaAM")
				tmpRec(i, j, 42)=RS("jiaBM")
				tmpRec(i, j, 43)=RS("BZKM")
				tmpRec(i, j, 44)=RS("KTAXM")
				tmpRec(i, j, 45)=RS("QITA")
				tmpRec(i, j, 46)=RS("FL")
				tmpRec(i, j, 47)=RS("real_total")
				tmpRec(i, j, 48)=RS("laonh")
				tmpRec(i, j, 49)=RS("sole")
				tmpRec(i, j, 50)=RS("dm")
				tmpRec(i, j, 51)=RS("lzbzj")
				tmpRec(i, j, 52)=RS("yymm")
				tmpRec(i, j, 53)=RS("outdate")
				tmpRec(i, j, 54)=RS("bh")
				tmpRec(i, j, 55)=cdbl(RS("hs")) 
				tmpRec(i, j, 56)=RS("GT")
				tmpRec(i, j, 57)=0
				tmpRec(i, j, 58)=cdbl(RS("BB"))+cdbl(RS("cv"))+cdbl(RS("phu"))+cdbl(RS("nn"))+cdbl(RS("kt"))+cdbl(RS("mt"))+cdbl(RS("ttkh"))
				tmpRec(i, j, 59)=cdbl(RS("wpamt"))
				tmpRec(i, j, 60)=cdbl(RS("wpamt"))+cdbl(RS("laonh"))
				
				
			rs.MoveNext
		else
			exit for
		end if
	 next

	 if rs.EOF then
		rs.Close
		Set rs = nothing
		exit for
	 end if
	next
	Session("empfileedit") = tmpRec
else
	TotalPage = cint(request("TotalPage"))
	'StoreToSession()
	CurrentPage = cint(request("CurrentPage"))
	RecordInDB  = REQUEST("RecordInDB")
	tmpRec = Session("empfileedit")

	Select case request("send")
	     Case "FIRST"
		      CurrentPage = 1
	     Case "BACK"
		      if cint(CurrentPage) <> 1 then
			     CurrentPage = CurrentPage - 1
		      end if
	     Case "NEXT"
		      if cint(CurrentPage) < cint(TotalPage) then
			     CurrentPage = CurrentPage + 1
			  else
			  	 CurrentPage = TotalPage
		      end if
	     Case "END"
		      CurrentPage = TotalPage
	     Case Else
		      CurrentPage = 1
	end Select
end if


FUNCTION FDT(D)
	Response.Write YEAR(D)&"/"&RIGHT("00"&MONTH(D),2)&"/"&RIGHT("00"&DAY(D),2)

END FUNCTION
%>

<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="refresh">
<link rel="stylesheet" href="../Include/style.css" type="text/css">
<link rel="stylesheet" href="../Include/style2.css" type="text/css">
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
 
'-----------------enter to next field
function enterto()
	if window.event.keyCode = 13 then window.event.keyCode =9
end function

function f()
	'<%=self%>.empid1.focus()
	'<%=self%>.empid1.select()
end function

function datachg()
	<%=self%>.action="<%=self%>.foregnd.asp?totalpage=0"
	<%=self%>.submit
end function

 
</SCRIPT> 
</head>
<body  topmargin="0" leftmargin="5"  marginwidth="0" marginheight="0"  onkeydown="enterto()" onload="f()">
<form name="<%=self%>" method="post" action="<%=self%>.foregnd.asp">
<INPUT TYPE=HIDDEN NAME="UID" VALUE=<%=SESSION("NETUSER")%>>
<INPUT TYPE=hidden NAME=TotalPage VALUE="<%=TotalPage%>">
<INPUT TYPE=hidden NAME=CurrentPage VALUE="<%=CurrentPage%>">
<INPUT TYPE=hidden NAME=RecordInDB VALUE="<%=RecordInDB%>">
<INPUT TYPE=hidden NAME=PageRec VALUE="<%=PageRec%>">
<INPUT TYPE=hidden NAME=gTotalPage VALUE="<%=gTotalPage%>">

	<table border=0 style="height:70px"><tr><td>&nbsp;</td></tr></table>
	
				<table width="100%" BORDER=0 align=center cellpadding=0 cellspacing=0 >
					<tr>
						<td>
							<table class="txt" cellpadding=3 cellspacing=3>
								<tr>
									<td  nowrap align=right>統計年月</td>
									<td nowrap colspan=3>
										<INPUT type="text" style="width:100px" NAME=yymm VALUE="<%=ym1%>" >~
										<INPUT type="text" style="width:100px"  NAME=yymm2 VALUE="<%=ym2%>" >		
									</td>
									<TD nowrap align=right>廠別</TD>
									<TD >
										<select name=WHSNO  class=txt8 onchange="datachg()" style="width:120px">
											<option value=""></option>
											<%SQL="SELECT * FROM BASICCODE WHERE FUNC='WHSNO' ORDER BY SYS_TYPE "
											SET RST = CONN.EXECUTE(SQL)
											WHILE NOT RST.EOF
											%>
											<option value="<%=RST("SYS_TYPE")%>" <%IF  RST("SYS_TYPE")=whsno THEN %> SELECTED <%END IF%> ><%=RST("SYS_VALUE")%></option>
											<%
											RST.MOVENEXT
											WEND
											%>
										</SELECT>
										<%SET RST=NOTHING %>
									</TD>		
									<TD nowrap align=right >國籍</TD>
									<TD >
										<select name=COUNTRY  class=txt8  onchange="datachg()" style="width:120px">
											<option value=""></option>
											<%SQL="SELECT * FROM BASICCODE WHERE FUNC='COUNTRY' ORDER BY SYS_TYPE "
											SET RST = CONN.EXECUTE(SQL)
											WHILE NOT RST.EOF
											%>
											<option value="<%=RST("SYS_TYPE")%>" <%IF RST("SYS_TYPE")=country THEN %> SELECTED <%END IF%>><%=RST("SYS_TYPE")%><%=RST("SYS_VALUE")%></option>
											<%
											RST.MOVENEXT
											WEND
											%>
										</SELECT>
										<%SET RST=NOTHING %>			
									</TD>
									
								</tr>
								<TR height=25 >
									<TD nowrap align=right >部門</TD>
									<TD >
										<select name=GROUPID  class=txt8  onchange="datachg()" style="width:120px">
											<option value=""></option>
											<%SQL="SELECT * FROM BASICCODE WHERE FUNC='GROUPID' and sys_type <>'AAA' ORDER BY SYS_TYPE "
											SET RST = CONN.EXECUTE(SQL)
											WHILE NOT RST.EOF
											%>
											<option value="<%=RST("SYS_TYPE")%>" <%IF  RST("SYS_TYPE")=GROUPID THEN %> SELECTED <%END IF%> ><%=RST("SYS_TYPE")%><%=RST("SYS_VALUE")%></option>
											<%
											RST.MOVENEXT
											WEND
											%>
										</SELECT>
										<%SET RST=NOTHING %>
									</TD>
									
									<TD nowrap align=right >班別</TD>
									<TD >
										<select name=shift  class=txt8  onchange="datachg()" style="width:100px">
											<option value="" <%if shift="" then %> selected<%end if%>></option>
											<option value="ALL" <%if shift="ALL" then %> selected<%end if%>>常日班</option>
											<option value="A" <%if shift="A" then %> selected<%end if%>>A班</option>
											<option value="B" <%if shift="B" then %> selected<%end if%>>B班</option>
										</SELECT>					
									</TD>
									<TD nowrap align=right>統計</TD>
									<TD >
										<select name=IOemp class=txt8 onchange="datachg()" style="width:120px"> 
											<option value="Y" <%if IOemp="Y" then %>selected<%end if%>>在職Tai chuc</option>
											<option value="" <%if IOemp="" then %>selected<%end if%>>全部ALL</option>
											<option value="N" <%if IOemp="N" then %>selected<%end if%>>已離職Toai Viec</option>
										 </select>	
									</TD> 				
											
									<TD nowrap align=right >工號</TD>
									<TD >
										<INPUT type="text" style="width:100px" NAME=empid1 value="<%=empid1%>">			
									</TD>		
									<td><INPUT TYPE=BUTTON NAME=BTN VALUE="查詢" CLASS=BUTTON onclick="datachg()" ONKEYDOWN="DATACHG()"></td>
								</TR>
							</TABLE>
						</td>
					</tr>
					<tr>
						<td align="center">
							<table id="myTableGrid" width="98%" class="txt9">
								<TR BGCOLOR="LightGrey" HEIGHT=25   >
									<TD width=50 nowrap align=center>YYMM</TD>
									<TD width=30 nowrap align=center>廠</TD>
									<TD width=60 nowrap align=center>bo phan</TD>
									<TD width=45 nowrap align=center>工號<br>Ma So</TD>
									<TD width=120 nowrap align=center>姓名<br>Ho Ten</TD>
									<TD width=70 nowrap align=center>到職日期<br>NVX</TD>
									<TD width=70 nowrap align=center>Chuc vu</TD>
									<TD width=30 nowrap align=center>幣別</TD>
									<TD width=60 nowrap align=right>(+)BB<BR>(-)保險</TD>
									<TD width=60 nowrap align=right>(+)CV<BR>(-)工團</TD>
									<TD width=60 nowrap align=right>(+)PHU<BR>(-)伙食</TD>
									<TD width=60 nowrap align=right>(+)NN<BR>(-)住宿</TD>
									<TD width=60 nowrap align=right>(+)KT<BR>(-)曠職</TD>
									<TD width=60 nowrap align=right>(+)MT<BR>(-)事假</TD>
									<TD width=60 nowrap align=right>(+)TTKH<BR>(-)病假</TD>
									<TD width=60 nowrap align=right>(+)QC<BR>(-)其他</TD>
									<TD width=60 nowrap align=right>(+)TNKH<BR>(-)不足月</TD>
									<TD width=60 nowrap align=right>(+)TBTR<BR>(-)所得稅</TD>
									<TD width=60 nowrap align=right>(+)JX<BR>薪資</TD>
									<TD width=60 nowrap align=right>(+)H1M<BR>離職金</TD>
									<TD width=60 nowrap align=right>(+)H2M<Br>實領</TD>
									<TD width=60 nowrap align=right>(+)H3M<Br>laonh</TD>
									<TD width=60 nowrap align=right>(+)B3M<Br>sole</TD>
									<TD width=60 nowrap align=right>willpower<Br>TOTAMT</TD>
									 
								</TR>
								<%for CurrentRow = 1 to PageRec
									IF CurrentRow MOD 2 = 0 THEN
										WKCOLOR="#ffffff"  '"LavenderBlush"
									ELSE
										WKCOLOR="#ffffff"  '"LavenderBlush"
									END IF
									if tmpRec(CurrentPage, CurrentRow, 1) <> "" then
								%>
									<%if CurrentRow>1 and tmpRec(CurrentPage, CurrentRow-1, 1)<>tmpRec(CurrentPage, CurrentRow, 1) then %>
									<Tr>
										<Td bgcolor=black colspan=24></td>
									</tr>
									<%end if%>	
								<TR BGCOLOR='<%=WKCOLOR%>' height=22>
									<TD align=center  >
											<%=tmpRec(CurrentPage, CurrentRow, 52)%> 			
									</TD>
									<input type=hidden value="<%=tmpRec(CurrentPage, CurrentRow, 1)%>" name="f_empid" >
									<input type=hidden value="<%=tmpRec(CurrentPage, CurrentRow, 52)%>" name="f_yymm" >
									<!--TD align=center  --><!--國籍-->
										
									<!--/TD-->
									<TD align=center  > <!--廠別-->
										<%=tmpRec(CurrentPage, CurrentRow, 7)%>
									</TD>
									<TD align=LEFT nowrap ><!--shift+部門-->
										<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then%>
											<%=tmpRec(CurrentPage, CurrentRow, 8)%>-<%=left(tmpRec(CurrentPage, CurrentRow, 12),3)%>
										<%end if%>	
									</TD> 
									<TD nowrap align=center><%=tmpRec(CurrentPage, CurrentRow, 1)%> 
										<!--a href='vbscript:oktest(<%=CurrentRow-1%>)'>
											
										</a-->
									</TD> 		
									<TD nowrap>
										<a href='vbscript:oktest(<%=CurrentRow-1%>)'>
											<%=tmpRec(CurrentPage, CurrentRow, 2)%><BR>
											<font class=txt8VN><%=tmpRec(CurrentPage, CurrentRow, 3)%></font>
										</a>
									</TD>
									<TD align=center nowrap>
										<%=tmpRec(CurrentPage, CurrentRow, 5)%><BR><font color=red><%=tmpRec(CurrentPage, CurrentRow, 53)%></font>
									</TD>
									<Td><%=left(tmpRec(CurrentPage, CurrentRow, 14),6)%></td>
									<Td align=center><%=tmpRec(CurrentPage, CurrentRow, 50)%></td>
									<Td align=right>
										<%if tmpRec(CurrentPage, CurrentRow, 18)<>"0" then%>
											<%=formatnumber(tmpRec(CurrentPage, CurrentRow, 18),0)%>
										<%else%>-
										<%end if%>
										<BR>
										<%if tmpRec(CurrentPage, CurrentRow, 54)<>"0" then%>
											<%=formatnumber(tmpRec(CurrentPage, CurrentRow, 54),0)%>
										<%else%>-
										<%end if%>
									</td>
									<Td  align=right>
										<%if tmpRec(CurrentPage, CurrentRow, 19)<>"0" then%>
											<%=formatnumber(tmpRec(CurrentPage, CurrentRow, 19),0)%>
										<%else%>-
										<%end if%>
										<br>
										<%if tmpRec(CurrentPage, CurrentRow, 56)<>"0" then%>
											<%=formatnumber(tmpRec(CurrentPage, CurrentRow, 56),0)%>
										<%else%>-
										<%end if%>
									</td>
									<Td  align=right>
										<%if tmpRec(CurrentPage, CurrentRow, 20)<>"0" then%>
											<%=formatnumber(tmpRec(CurrentPage, CurrentRow, 20),0)%>
										<%else%>-
										<%end if%>
										<br>
										<%if tmpRec(CurrentPage, CurrentRow, 55)<>"0" then%>
											<%=formatnumber(tmpRec(CurrentPage, CurrentRow, 55),0)%>
										<%else%>-
										<%end if%>
										
									</td>
									<Td  align=right>
										<%if tmpRec(CurrentPage, CurrentRow, 21)<>"0" then%>
											<%=formatnumber(tmpRec(CurrentPage, CurrentRow, 21),0)%>
										<%else%>-
										<%end if%>
										<br> 			
										<%if tmpRec(CurrentPage, CurrentRow, 57)<>"0" then%>
											<%=formatnumber(tmpRec(CurrentPage, CurrentRow, 57),0)%>
										<%else%>-
										<%end if%>
									</td>
									<Td  align=right>
										<%if tmpRec(CurrentPage, CurrentRow, 22)<>"0" then%>
											<%=formatnumber(tmpRec(CurrentPage, CurrentRow, 22),0)%>
										<%else%>-
										<%end if%>
										<br>
										<%if tmpRec(CurrentPage, CurrentRow, 40)<>"0" then%>
											<%=formatnumber(tmpRec(CurrentPage, CurrentRow, 40),0)%>
										<%else%>-
										<%end if%>
									</td>
									<Td  align=right>
										<%if tmpRec(CurrentPage, CurrentRow, 23)<>"0" then%>
											<%=formatnumber(tmpRec(CurrentPage, CurrentRow, 23),0)%>
										<%else%>-
										<%end if%>
										<br>
										<%if tmpRec(CurrentPage, CurrentRow, 41)<>"0" then%>
											<%=formatnumber(tmpRec(CurrentPage, CurrentRow, 41),0)%>
										<%else%>-
										<%end if%>
									</td>
									<Td  align=right>
										<%if tmpRec(CurrentPage, CurrentRow, 24)<>"0" then%>
											<%=formatnumber(tmpRec(CurrentPage, CurrentRow, 24),0)%>
										<%else%>-
										<%end if%>
										<br>
										<%if tmpRec(CurrentPage, CurrentRow, 42)<>"0" then%>
											<%=formatnumber(tmpRec(CurrentPage, CurrentRow, 42),0)%>
										<%else%>-
										<%end if%>
									</td>
									<Td  align=right>
										<%if tmpRec(CurrentPage, CurrentRow, 25)<>"0" then%>
											<%=formatnumber(tmpRec(CurrentPage, CurrentRow, 25),0)%>
										<%else%>-
										<%end if%>
										<br>
										<%if tmpRec(CurrentPage, CurrentRow, 45)<>"0" then%>
											<%=formatnumber(tmpRec(CurrentPage, CurrentRow, 45),0)%>
										<%else%>-
										<%end if%>
									</td>
									<Td  align=right>
										<%if tmpRec(CurrentPage, CurrentRow, 26)<>"0" then%>
											<%=formatnumber(tmpRec(CurrentPage, CurrentRow, 26),0)%>
										<%else%>-
										<%end if%>
										<br>
										<%if tmpRec(CurrentPage, CurrentRow, 43)<>"0" then%>
											<%=formatnumber(tmpRec(CurrentPage, CurrentRow, 43),0)%>
										<%else%>-
										<%end if%>
									</td>
									<Td  align=right>	
										<%if tmpRec(CurrentPage, CurrentRow, 27)<>"0" then%>
											<%=formatnumber(tmpRec(CurrentPage, CurrentRow, 27),0)%>
										<%else%>-
										<%end if%>
										<br>
										<%if tmpRec(CurrentPage, CurrentRow,44)<>"0" then%>
											<%=formatnumber(tmpRec(CurrentPage, CurrentRow, 44),0)%>
										<%else%>-
										<%end if%>
									</td>
									<Td  align=right>
										<%if tmpRec(CurrentPage, CurrentRow, 28)<>"0" then%>
											<%=formatnumber(tmpRec(CurrentPage, CurrentRow, 28),0)%>
										<%else%>-
										<%end if%>
										<br>
										<%if tmpRec(CurrentPage, CurrentRow, 58)<>"0" then%>
											<font color="red"><%=formatnumber(tmpRec(CurrentPage, CurrentRow, 58),0)%></font>
										<%else%>-
										<%end if%>
									</td>
									<Td  align=right>
										<%if tmpRec(CurrentPage, CurrentRow, 29)<>"0" then%>
											<%=formatnumber(tmpRec(CurrentPage, CurrentRow, 29),0)%>
										<%else%>-
										<%end if%>
										<br>
										<%if tmpRec(CurrentPage, CurrentRow, 53)<>"" and trim(tmpRec(CurrentPage, CurrentRow, 52))=left(replace(tmpRec(CurrentPage, CurrentRow, 53),"/",""),6) then%>
											<%=formatnumber(tmpRec(CurrentPage, CurrentRow, 51),0)%>
										<%else%>-
										<%end if%>
									</td>
									<Td  align=right>
										<%if tmpRec(CurrentPage, CurrentRow, 30)<>"0" then%>
											<%=formatnumber(tmpRec(CurrentPage, CurrentRow, 30),0)%>
										<%else%>-
										<%end if%>
										<br>
										<font color=blue><b><%=formatnumber(tmpRec(CurrentPage, CurrentRow, 47),0)%></b></font>
									</td>
									<Td  align=right>
										<%if tmpRec(CurrentPage, CurrentRow, 31)<>"0" then%>
										<%=formatnumber(tmpRec(CurrentPage, CurrentRow, 31),0)%>
										<%else%>-
										<%end if%>
										<br>
										<font color=blue><b><%=formatnumber(tmpRec(CurrentPage, CurrentRow, 48),0)%></b></font>
									</td>
									<Td  align=right>
										<%if tmpRec(CurrentPage, CurrentRow, 32)<>"0" then%>
											<%=formatnumber(tmpRec(CurrentPage, CurrentRow, 32),0)%>
										<%else%>-
										<%end if%>
										<br>
										<%if tmpRec(CurrentPage, CurrentRow, 49)<>"0" then%>
											<%=formatnumber(tmpRec(CurrentPage, CurrentRow, 49),0)%>
										<%else%>-
										<%end if%>
									</td>
									<Td  align=right>
										<%if tmpRec(CurrentPage, CurrentRow, 59)<>"0" then%>
											<%=formatnumber(tmpRec(CurrentPage, CurrentRow, 59),0)%>
										<%else%>-
										<%end if%>
										<br>
										<%if tmpRec(CurrentPage, CurrentRow, 60)<>"0" then%>
											<%=formatnumber(tmpRec(CurrentPage, CurrentRow, 60),0)%>
										<%else%>-
										<%end if%>
									</td>		
								</TR> 
								<%end if%>
								<%next%>
							</TABLE>
						</td>
					</tr>
					<tr>
						<td align="center">
							<input type=hidden value="" name="f_yymm" >
							<input type=hidden value="" name="f_empid" >
							<table class="txt"  cellpadding=3 cellspacing=3>
								<tr>
								<td align="CENTER" height=40 width=80%>
								PAGE:<%=CURRENTPAGE%> / <%=TOTALPAGE%> , COUNT:<%=RECORDINDB%><BR>
								<% If CurrentPage > 1 Then %>
									<input type="submit" name="send" value="FIRST" class=button>
									<input type="submit" name="send" value="BACK" class=button>
								<% Else %>
									<input type="submit" name="send" value="FIRST" disabled class=button>
									<input type="submit" name="send" value="BACK" disabled class=button>
								<% End If %>
								<% If cint(CurrentPage) < cint(TotalPage) Then %>
									<input type="submit" name="send" value="NEXT" class=button>
									<input type="submit" name="send" value="END" class=button>
								<% Else %>
									<input type="submit" name="send" value="NEXT" disabled class=button>
									<input type="submit" name="send" value="END" disabled class=button>
								<% End If %>
								</td>
								<td>	<BR>
									<input type="button" name="send" value="回主畫面"   class=button onclick="history.back()">
								</td>
								</TR>
							</TABLE>
						</td>
					</tr>
				</table>
			
</form>




</body>
</html>

<script language=vbscript>
function BACKMAIN()

	open "empfile.fore1.asp" , "_self"
end function

function oktest(index)
	f1=<%=self%>.f_empid(index).value
	f2=<%=self%>.f_yymm(index).value
	'alert f1 & f2 
	'tp=<%=self%>.totalpage.value
	'cp=<%=self%>.CurrentPage.value
	'rc=<%=self%>.RecordInDB.value
	wt = (window.screen.width )*0.6
	ht = window.screen.availHeight*0.6
	tp = (window.screen.width )*0.02
	lt = (window.screen.availHeight)*0.02
	
	open "<%=self%>.showsalary.asp?empid="&f1&"&yymm="&f2  , "_blank", "top="& tp &", left="& lt &", width="& wt &",height="& ht &", scrollbars=yes"	
	
end function

</script>


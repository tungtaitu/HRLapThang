<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<!-- #include file = "../../GetSQLServerConnection.fun" -->
<!-- #include file="../../ADOINC.inc" -->
<!--#include file="../../include/sideinfolev2.inc"-->
<%
'on error resume next

SELF = "empfileedit"

Set conn = GetSQLServerConnection()
Set rs = Server.CreateObject("ADODB.Recordset")

yymmstr = request("yymm")
country = request("country")
whsno = trim(request("whsno"))
unitno = trim(request("unitno"))
groupid = trim(request("groupid"))
zuno = trim(request("zuno"))
job = trim(request("job"))
QUERYX = trim(request("empid1"))
shift = request("shift")


gTotalPage = 1
PageRec = 16    'number of records per page
TableRec = 50    'number of fields per record
'if day(date())<=10 then
'	if month(date()) = 12 then
'		yymm = year(date())- 1 & right("00" & month(date()) , 2 )
'	else
'		yymm = year(date())& right("00" & month(date())-1 , 2 )
'	end if
'else
'	yymm = year(date())& right("00" & month(date()), 2 )
'end if
yymm = yymmstr  
nowmonth = year(date())&right("00"&month(date()),2) 

calcdt = left(YYMM,4)&"/"& right(yymm,2)&"/01"

if right(yymmstr,2) mod 2 = 0 then 
	ccx=35
else
	ccx=36
end if 		 

viewid = session("netuser")  
'viewid = "LSARY" 

if viewid="LSARY"  then   
	'eidt 20090620 修改假日加班不計
	sql=" SELECT  a.tothx as tj,  case when isnull(tothjd,0) =0 then isnull(newtoth,0) else isnull(tothjd,0) end toth,   isnull(wdaka,0) forget , b.empid as emp_id ,*  "&_
			",isnull(pt.a,0) as jiaA ,isnull(pt.b,0) as jiaB ,isnull(pt.C,0) as jiaC ,isnull(pt.D,0) as jiaD "&_
			",isnull(pt.E,0) as jiaE ,isnull(pt.F,0) as jiaF ,isnull(pt.G,0) as jiaG ,isnull(pt.H,0) as jiaH "&_
			"FROM    "&_			
			"(select empid, nindat, empnam_cn, empnam_vn ,status , outdate ,whsno, groupid , country , wstr, ustr, gstr, zstr , jstr  "&_
			",job ,autoid ,unitno , zuno ,shift  from  View_EMPFILE where  "&_
			"isnull(status,'')<>'D' and (isnull(outdate,'')='' or outdate>'"& calcdt &"'  )  "&_		
			"and case when '"& whsno &"'='' then '' else  whsno end ='"& whsno &"'  and case when '"& groupid &"'='' then '' else groupid end = '"& groupid &"'  "&_
			"and case when '"& country &"'='' then '' else country end =  '"& country &"' and case when '"&QUERYX&"'='' then '' else empid end= '"&QUERYX&"' "&_ 	
			"  ) B  "&_
			"left join  ( "&_			
			"select yymm, empid  WORKEMPID ,  sum(isnull(toth,0)) tothx, "&_
			"sum( isnull(  latefor,0) ) latefor  , sum( isnull( kzhour,0) ) kzhour , sum( isnull( nh1,0) ) h1,  "&_
			"sum( isnull( nh2,0) ) h2,  sum( isnull( nh3,0) ) h3 , sum( isnull(nb3,0) ) b3 ,sum( isnull(nb4,0) ) b4 ,sum( isnull(nb5,0) ) b5 , "&_
			"sum( isnull(jiaa,0) ) xjiaA, sum( isnull(jiaB,0) ) xjiaB, sum( isnull(jiaC,0) ) xjiaC, "&_
			"sum( isnull(jiaD,0) ) xjiaD, sum( isnull(jiaE,0) ) xjiaE, sum( isnull(jiaF,0) ) xjiaF, "&_
			"sum( isnull(jiaG,0) ) xjiaG, sum( isnull(jiaH,0) ) xjiaH  ,sum(isnull(newtoth,0)) newtoth  "&_
			"from   EMPWORK  WHERE  YYMM='"& yymmstr &"' "&_
			"and case when '"&QUERYX&"'='' then '' else empid end= '"&QUERYX&"'  group by empid ,yymm  "&_
			") A  ON B.EMPID =A.WORKEMPID   "&_
			"left join (select count(empid) wdaka , empid ,yymm from empforget where isnull(status,'')<>'D' group by empid,yymm )F on f.empid=a.WORKEMPID  and a.yymm=f.yymm "&_
			"left join ( select empid, sum(toth) as tothjd  from empworkjd  where  YYMM='"& yymmstr &"'  group by empid ) jd on jd.empid=b.empid "&_
			"left join ( "&_
			"	SELECT * FROM "&_
			"				( "&_
			"					SELECT empid,jiatype,SUM(hhour) AS Hours "&_
			"						FROM empholiday where convert(char(6), dateup,112)='"& yymmstr &"' "&_
			"						GROUP BY empid,jiatype "&_
			"				) AS p "&_
			"		PIVOT "&_
			"				( "&_
			"					SUM(Hours) FOR jiatype IN ([A],[B],[C],[D],[E],[F],[G],[H],[I]) "&_
			"				) AS pt "&_
			") pt on pt.empid=b.empid "&_  
			"where 1=1 and a.YYMM='"& yymmstr &"'"	
else 
	'  忘遲次數有誤 change by zhang 20110301
	sql=" SELECT  toth as tj,isnull(wdaka,0) forget , b.empid as emp_id ,* "&_
			",isnull(pt.a,0) as jiaA ,isnull(pt.b,0) as jiaB ,isnull(pt.C,0) as jiaC ,isnull(pt.D,0) as jiaD "&_
			",isnull(pt.E,0) as jiaE ,isnull(pt.F,0) as jiaF ,isnull(pt.G,0) as jiaG ,isnull(pt.H,0) as jiaH "&_
			"FROM    "&_			
			"(select empid, nindat, empnam_cn, empnam_vn ,status , outdate ,whsno, groupid , country , wstr, ustr, gstr, zstr , jstr  "&_
			",job ,autoid ,unitno , zuno ,shift  from  View_EMPFILE where  "&_
			"isnull(status,'')<>'D' and (isnull(outdate,'')='' or outdate>'"& calcdt &"'  )  "&_		
			"and case when '"& whsno &"'='' then '' else  whsno end ='"& whsno &"'  and case when '"& groupid &"'='' then '' else groupid end = '"& groupid &"'  "&_
			"and case when '"& country &"'='' then '' else country end =  '"& country &"' and case when '"&QUERYX&"'='' then '' else empid end= '"&QUERYX&"' "&_ 	
			"  ) B  "&_
			"left join  ( "&_			
			"select yymm, empid  WORKEMPID ,  sum(isnull(toth,0)) toth, "&_
			"sum( isnull(  latefor,0) ) latefor  , sum( isnull( kzhour,0) ) kzhour , sum( isnull( h1,0) ) h1,  "&_
			"sum( isnull( h2,0) ) h2,  sum( isnull( h3,0) ) h3 , sum( isnull(b3,0) ) b3 ,sum( isnull(b4,0) ) b4 ,sum( isnull(b5,0) ) b5 , "&_
			"sum( isnull(jiaa,0) ) xjiaA, sum( isnull(jiaB,0) ) xjiaB, sum( isnull(jiaC,0) ) xjiaC, "&_
			"sum( isnull(jiaD,0) ) xjiaD, sum( isnull(jiaE,0) ) xjiaE, sum( isnull(jiaF,0) ) xjiaF, "&_
			"sum( isnull(jiaG,0) ) xjiaG, sum( isnull(jiaH,0) ) xjiaH  "&_
			"from   EMPWORK  WHERE  YYMM='"& yymmstr &"' "&_
			"and case when '"&QUERYX&"'='' then '' else empid end= '"&QUERYX&"'  group by empid ,yymm  "&_
			") A  ON B.EMPID =A.WORKEMPID   "&_
			"left join (select count(empid) wdaka , empid ,yymm from empforget where isnull(status,'')<>'D' group by empid,yymm )F on f.empid=a.WORKEMPID  and a.yymm=f.yymm "&_
			"left join ( "&_
			"	SELECT * FROM "&_
			"				( "&_
			"					SELECT empid,jiatype,SUM(hhour) AS Hours "&_
			"						FROM empholiday where convert(char(6), dateup,112)='"& yymmstr &"' "&_
			"						GROUP BY empid,jiatype "&_
			"				) AS p "&_
			"		PIVOT "&_
			"				( "&_
			"					SUM(Hours) FOR jiatype IN ([A],[B],[C],[D],[E],[F],[G],[H],[I]) "&_
			"				) AS pt "&_
			") pt on pt.empid=b.empid "&_  
			"where 1=1 and a.YYMM='"& yymmstr &"'"
end if 

if shift<>"" and shift<>"D" then
	sql=sql&"and  b.shift='"& shift &"' "
end if
if shift="D" then
	sql=sql&"and  isnull(b.shift,'')='' "
end if
sql=sql&"order by b.empid "
	
'response.write sql
'response.end
if request("TotalPage") = "" or request("TotalPage") = "0" then
	CurrentPage = 1
	rs.Open SQL, conn, 3, 1
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
				tmpRec(i, j, 1) = trim(rs("emp_id"))
				tmpRec(i, j, 2) = trim(rs("empnam_cn"))
				tmpRec(i, j, 3) = trim(rs("empnam_vn"))
				tmpRec(i, j, 4) = rs("country")
				tmpRec(i, j, 5) = rs("nindat")
				tmpRec(i, j, 6) = rs("job")
				tmpRec(i, j, 7) = rs("whsno")
				tmpRec(i, j, 8) = rs("unitno")
				tmpRec(i, j, 9)	=RS("groupid")
				tmpRec(i, j, 10)=RS("zuno")
				tmpRec(i, j, 11)=RS("wstr")
				tmpRec(i, j, 12)=RS("ustr")
				tmpRec(i, j, 13)=RS("gstr")
				tmpRec(i, j, 14)=RS("zstr")
				tmpRec(i, j, 15)=RS("jstr")
				tmpRec(i, j, 16)=RS("country")
				tmpRec(i, j, 17)=RS("autoid")
				IF RS("zuno")="XX" THEN
					tmpRec(i, j, 18)=""
				ELSE
					tmpRec(i, j, 18)=RS("zuno")
				END IF  
				tmpRec(i, j, 19)=(RS("totH"))
				tmpRec(i, j, 20)=(RS("forget"))
				tmpRec(i, j, 21)=(RS("latefor"))
				tmpRec(i, j, 22)=(RS("kzhour"))
				tmpRec(i, j, 23)=(RS("h1"))
				tmpRec(i, j, 24)=(RS("h2"))
				tmpRec(i, j, 25)=(RS("h3"))
				tmpRec(i, j, 26)=(RS("b3"))
				tmpRec(i, j, 27)=(RS("jiaa"))
				tmpRec(i, j, 28)=(RS("jiab"))
				tmpRec(i, j, 29)=(RS("jiac"))
				tmpRec(i, j, 30)=(RS("jiad"))
				tmpRec(i, j, 31)=(RS("jiae"))
				tmpRec(i, j, 32)=(RS("jiaf"))
				tmpRec(i, j, 33)=(RS("jiag"))
				tmpRec(i, j, 34)=(RS("jiah")) 
				tmpRec(i, j, 35)=trim(RS("outdate")) 
				tmpRec(i, j, 36)=(RS("b4"))
				tmpRec(i, j, 37)=(RS("b5"))
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
	Session("empfileeditN") = tmpRec
else
	TotalPage = cint(request("TotalPage"))
	'StoreToSession()
	CurrentPage = cint(request("CurrentPage"))
	RecordInDB  = REQUEST("RecordInDB")
	tmpRec = Session("empfileeditN")

	Select case request("send")
	     Case "FIRST"
		      CurrentPage = 1
	     Case "BACK"
		      if cint(CurrentPage) <> 1 then
			     CurrentPage = CurrentPage - 1
		      end if
	     Case "NEXT"
		      if cint(CurrentPage) <= cint(TotalPage) then
			     CurrentPage = CurrentPage + 1
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

'nowmonth = year(date())&right("00"&month(date()),2)
nowmonth = yymm
calcmonth = nowmonth
'if month(date())="01" then
'	if day(date())>11 then
'		calcmonth = nowmonth
'	else
'		calcmonth = year(date()-1)&"12"
'	end if
'else
'	if day(date())>11 then
'		calcmonth = nowmonth
'	else
'		calcmonth =  year(date())&right("00"&month(date())-1,2)
'	end if
'end if 

  
cDatestr=CDate(LEFT(YYMM,4)&"/"&RIGHT(YYMM,2)&"/01")
days = DAY(cDatestr+(32-DAY(cDatestr))-DAY(cDatestr+(32-DAY(cDatestr))))   
 
SQL=" SELECT * FROM YDBMCALE WHERE CONVERT(CHAR(6),DAT,112)  ='"& YYMM &"' AND  DATEPART( DW,DAT ) ='1'  "
Set rsTT = Server.CreateObject("ADODB.Recordset")
RSTT.OPEN SQL, CONN, 3, 3
IF NOT RSTT.EOF THEN
	HHCNT = CDBL(RSTT.RECORDCOUNT)
ELSE
	HHCNT = 0
END IF
SET RSTT=NOTHING

'RESPONSE.WRITE HHCNT &"<br>"
'RESPONSE.END
 
MMDAYS = CDBL(days) -cdbl(HHCNT)
conn.close
set conn=nothing
%>

<html> 
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="refresh">
<link rel="stylesheet" href="../../Include/style.css" type="text/css">
<link rel="stylesheet" href="../../Include/style2.css" type="text/css">
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--
'-----------------enter to next field
function enterto()
	if window.event.keyCode = 13 then window.event.keyCode =9
end function

function f()
	<%=self%>.empid1.focus()
end function

function datachg()
	<%=self%>.action="empwork.fore.asp?totalpage=0"
	<%=self%>.submit
end function

-->
</SCRIPT>
</head>
<body  topmargin="5" leftmargin="5"  marginwidth="0" marginheight="0"  onkeydown="enterto()" onload="f()">
<form name="<%=self%>" method="post" action="empwork.fore.asp">
<INPUT TYPE=HIDDEN NAME="UID" VALUE=<%=SESSION("NETUSER")%>>
<INPUT TYPE=hidden NAME=TotalPage VALUE="<%=TotalPage%>">
<INPUT TYPE=hidden NAME=CurrentPage VALUE="<%=CurrentPage%>">
<INPUT TYPE=hidden NAME=PageRec VALUE="<%=PageRec%>">
<INPUT TYPE=hidden NAME=RecordInDB VALUE="<%=RecordInDB%>">
<INPUT TYPE=hidden NAME=gTotalPage VALUE="<%=gTotalPage%>">
<table border=0 style="height:70px"><tr><td>&nbsp;</td></tr></table>

<table width="100%" BORDER=0 align=center cellpadding=0 cellspacing=0 >
	<tr>
		<td>
			<table  class="txt" cellspacing="3" cellpadding="3">
				<tr>
					<td width="30px">&nbsp;</td>
					<TD align=right><font color=blue> 統計年月<br>Thống kê năm :</font></td>
					<td nowrap colspan=10>
						<INPUT type="text" style="width:80px" NAME=yymm VALUE="<%=yymmstr%>" readonly  >
						<INPUT type="text" style="width:50px" NAME=mdays VALUE="<%=MMDAYS%>" readonly  >
					</TD>
				</tr>
				<TR height=25 >
					<td width="30px">&nbsp;</td>
					<TD nowrap align=right>國籍<br>Quốc gia:</TD>
					<TD >
						<select name=country   onchange="datachg()" style="width:100px">
							<option value=""></option>
							<%Set conn = GetSQLServerConnection()
							SQL="SELECT * FROM BASICCODE WHERE FUNC='country' ORDER BY SYS_TYPE "
							SET RST = CONN.EXECUTE(SQL)
							WHILE NOT RST.EOF
							%>
							<option value="<%=RST("SYS_TYPE")%>" <%IF  RST("SYS_TYPE")=country THEN %> SELECTED <%END IF%> ><%=RST("SYS_VALUE")%></option>
							<%
							RST.MOVENEXT
							WEND
							%>
						</SELECT>
						<%rst.close
						SET RST=NOTHING %>
					</TD>
					<TD   nowrap align=right>廠別<br>Xưởng</TD>
					<TD >
						<select name=whsno   onchange="datachg()" style="width:100px">				
							<%SQL="SELECT * FROM BASICCODE WHERE FUNC='whsno'  ORDER BY SYS_TYPE "
							SET RST = CONN.EXECUTE(SQL)
							WHILE NOT RST.EOF
							%>
							<option value="<%=RST("SYS_TYPE")%>" <%IF  RST("SYS_TYPE")=whsno THEN %> SELECTED <%END IF%>><%=RST("SYS_VALUE")%></option>
							<%
							RST.MOVENEXT
							WEND
							%>
						</SELECT>
						<%rst.close
						SET RST=NOTHING %>
					</TD>
					<TD nowrap align=right >組/部門<br>Tổ/Bộ phận</TD>
					<TD >
						<select name=GROUPID    onchange="datachg()" style="width:100px">
							<option value=""></option>
							<%SQL="SELECT * FROM BASICCODE WHERE FUNC='GROUPID' and sys_type <>'AAA' ORDER BY SYS_TYPE "
							SET RST = CONN.EXECUTE(SQL)
							WHILE NOT RST.EOF
							%>
							<option value="<%=RST("SYS_TYPE")%>" <%IF  RST("SYS_TYPE")=GROUPID THEN %> SELECTED <%END IF%> ><%=RST("SYS_VALUE")%></option>
							<%
							RST.MOVENEXT
							WEND
							%>
						</SELECT>
						<%rst.close
						SET RST=NOTHING 
						conn.close
						set conn=nothing
						%>
					</TD> 
					<TD nowrap align=right >員工編號<br>Mã số nhân viên</TD>
					<TD  >
						<INPUT type="text" style="width:100px" NAME=empid1 value="<%=QUERYX%>">
					</TD>
					<TD nowrap align=right >班別<br>Ca</TD>
					<TD >
						<select name=shift  onchange="datachg()" style="width:70px">
							<option value="" <%if shift="" then %> selected<%end if%>>----</option>
							<option value="ALL" <%if shift="ALL" then %> selected<%end if%>>ALL</option>
							<option value="A" <%if shift="A" then %> selected<%end if%>>Ca A</option>
							<option value="B" <%if shift="B" then %> selected<%end if%>>Ca B</option>
							<option value="D" <%if shift="D" then %> selected<%end if%>>其他 Khác</option>
						</select>		 	
					</TD>
					<td><INPUT TYPE=BUTTON NAME=BTN VALUE="(S)查詢 Tìm kiếm " CLASS="btn btn-sm btn-outline-secondary" onclick="datachg()" onkeydown="datachg()"></td>
				</TR>
			</table>
		</td>
	</tr>
	<tr>
		<td align="center">
			<table id="myTableGrid" width="98%">
				<tr BGCOLOR="LightGrey" height=22 class="font9">
					<TD width=50 nowrap align=center rowspan=2>工號<br>Mã số</TD>
					<TD width=190 nowrap align=center rowspan=2>姓名<br>Họ tên</TD>
					<TD width=80 align=center nowrap rowspan=2>到職日<br>Ngày vx</TD>
					<TD align=center rowspan=2>總工時<br>Tổng giờ</TD>
					<TD align=center rowspan=2>忘刷<br>Quên BT</TD>
					<TD align=center rowspan=2>遲早<br>Trễ/Sớm</TD>
					<TD align=center rowspan=2>曠職<br>Vắng</TD>
					<td align=center colspan=6 >加班(單位:小時)Tăng Ca (Giờ)</td>
					<td colspan=8 align=center >休假(單位:小時)Nghỉ phép (Giờ)</td>
				</tr>
				<tr BGCOLOR="LightGrey" height=22 class="font9">
					<td align=center>一般(1.5)<br>Thường</td>
					<td align=center>休息(2.0)<br>Nghỉ</td>
					<td align=center>假日(3.0)<br>Lễ</td>
					<td align=center>津貼(1.5)<br>Phụ cấp</td>
					<td align=center>津貼(0.3)<br>Phụ cấp</td>
					<td align=center>夜班(2.1)<br>Ca đêm</td>
					<td align=center>公假<br>G-Lễ</td>
					<td align=center>年假<br>E-Năm</td>
					<td align=center>事假<br>A-Riêng</td>
					<td align=center>病假<br>B-Bệnh</td>
					<td align=center>工傷<br>H-CThương</td>
					<td align=center>婚假<br>C-Cưới</td>
					<td align=center>喪假<br>D-Tang</td>
					<td align=center>產假<br>F-TSản</td>
				</tr>
				<%for CurrentRow = 1 to PageRec
					IF CurrentRow MOD 2 = 0 THEN
						WKCOLOR="LavenderBlush"
					ELSE
						WKCOLOR=""
					END IF
					'if tmpRec(CurrentPage, CurrentRow, 1) <> "" then
				%>
				<TR BGCOLOR=<%=WKCOLOR%>>
					<TD align=center>
						<input name="empid" type="hidden" value="<%=tmpRec(CurrentPage, CurrentRow, 1)%>">	
						<a href='vbscript:oktest(<%=CurrentRow-1%>)'><font class=txt><%=tmpRec(CurrentPage, CurrentRow, 1)%></font></a>
					</TD>
					<TD nowrap><a href='vbscript:oktest(<%=CurrentRow-1%>)'><font class=txt><%=tmpRec(CurrentPage, CurrentRow, 2)%></font><br><font class=txt8><%=tmpRec(CurrentPage, CurrentRow, 3)%></font></a></TD>
					<TD align=center><font class=txt><%=tmpRec(CurrentPage, CurrentRow, 5)%></font><BR><font color=red class=txt8><%=tmpRec(CurrentPage, CurrentRow, 35)%></font></TD>
					<TD align=center><input type="text" name=toth  class="readonly8s"   value="<%=tmpRec(CurrentPage, CurrentRow, 19)%>" style="width:100%;text-align:right"></TD>
					<TD align=center><input type="text" name=forget  class="readonly8s" readonly  value="<%=tmpRec(CurrentPage, CurrentRow, 20)%>" style="width:100%;text-align:right;color:<%if cdbl(tmpRec(CurrentPage, CurrentRow, 20))>0 then %>RoyalBlue<%else%>White<%end if%>"></TD>
					<TD align=center><input type="text" name=latefor  class="readonly8s" readonly  value="<%=tmpRec(CurrentPage, CurrentRow, 21)%>" style="width:100%;text-align:right;color:<%if cdbl(tmpRec(CurrentPage, CurrentRow, 21))>0 then %>RoyalBlue<%else%>White<%end if%>"></TD>
					<TD align=center><input type="text" name=kzhour  class="readonly8s" readonly  value="<%=tmpRec(CurrentPage, CurrentRow, 22)%>" style="width:100%;text-align:right;color:<%if cdbl(tmpRec(CurrentPage, CurrentRow, 22))=0 then %>white<%else%>red<%end if%>"></TD>
					<TD align=center><input type="text" name=h1  class="readonly8s" readonly  value="<%=tmpRec(CurrentPage, CurrentRow, 23)%>" style="width:100%;text-align:right;color:<%if cdbl(tmpRec(CurrentPage, CurrentRow, 23))>0 then %>black<%else%>White<%end if%>"></TD>
					<TD align=center><input type="text" name=h2  class="readonly8s" readonly  value="<%=tmpRec(CurrentPage, CurrentRow, 24)%>" style="width:100%;text-align:right;color:<%if cdbl(tmpRec(CurrentPage, CurrentRow, 24))>0 then %>black<%else%>White<%end if%>"></TD>
					<TD align=center><input type="text" name=h3  class="readonly8s" readonly  value="<%=tmpRec(CurrentPage, CurrentRow, 25)%>" style="width:100%;text-align:right;color:<%if cdbl(tmpRec(CurrentPage, CurrentRow, 25))>0 then %>black<%else%>White<%end if%>"></TD>
					<TD align=center><input type="text" name=b4  class="readonly8s" readonly  value="<%=tmpRec(CurrentPage, CurrentRow, 36)%>" style="width:100%;text-align:right;color:<%if cdbl(tmpRec(CurrentPage, CurrentRow, 36))>0 then %>black<%else%>White<%end if%>"></TD>
					<TD align=center><input type="text" name=b5  class="readonly8s" readonly  value="<%=tmpRec(CurrentPage, CurrentRow, 37)%>" style="width:100%;text-align:right;color:<%if cdbl(tmpRec(CurrentPage, CurrentRow, 37))>0 then %>black<%else%>White<%end if%>"></TD>
					<TD align=center><input type="text" name=b3  class="readonly8s" readonly  value="<%=tmpRec(CurrentPage, CurrentRow, 26)%>" style="width:100%;text-align:right;color:<%if cdbl(tmpRec(CurrentPage, CurrentRow, 26))>0 then %>black<%else%>White<%end if%>"></TD>
					<TD align=center><input type="text" name=jiag  class="readonly8s" readonly  value="<%=tmpRec(CurrentPage, CurrentRow, 33)%>" style="width:100%;text-align:right;color:<%if cdbl(tmpRec(CurrentPage, CurrentRow, 33))>0 then %>black<%else%>White<%end if%>"></TD>
					<TD align=center><input type="text" name=jiae  class="readonly8s" readonly  value="<%=tmpRec(CurrentPage, CurrentRow, 31)%>" style="width:100%;text-align:right;color:<%if cdbl(tmpRec(CurrentPage, CurrentRow, 31))>0 then %>black<%else%>White<%end if%>"></TD>
					<TD align=center><input type="text" name=jiaa  class="readonly8s" readonly  value="<%=tmpRec(CurrentPage, CurrentRow, 27)%>" style="width:100%;text-align:right;color:<%if cdbl(tmpRec(CurrentPage, CurrentRow, 27))>0 then %>black<%else%>White<%end if%>"></TD>
					<TD align=center><input type="text" name=jiab  class="readonly8s" readonly  value="<%=tmpRec(CurrentPage, CurrentRow, 28)%>" style="width:100%;text-align:right;color:<%if cdbl(tmpRec(CurrentPage, CurrentRow, 28))>0 then %>black<%else%>White<%end if%>"></TD>
					<TD align=center><input type="text" name=jiah  class="readonly8s" readonly  value="<%=tmpRec(CurrentPage, CurrentRow, 34)%>" style="width:100%;text-align:right;color:<%if cdbl(tmpRec(CurrentPage, CurrentRow, 34))>0 then %>black<%else%>White<%end if%>"></TD>
					<TD align=center><input type="text" name=jiac  class="readonly8s" readonly  value="<%=tmpRec(CurrentPage, CurrentRow, 29)%>" style="width:100%;text-align:right;color:<%if cdbl(tmpRec(CurrentPage, CurrentRow, 29))>0 then %>black<%else%>White<%end if%>"></TD>
					<TD align=center><input type="text" name=jiad  class="readonly8s" readonly  value="<%=tmpRec(CurrentPage, CurrentRow, 30)%>" style="width:100%;text-align:right;color:<%if cdbl(tmpRec(CurrentPage, CurrentRow, 30))>0 then %>black<%else%>White<%end if%>"></TD>
					<TD align=center><input type="text" name=jiaf  class="readonly8s" readonly  value="<%=tmpRec(CurrentPage, CurrentRow, 32)%>" style="width:100%;text-align:right;color:<%if cdbl(tmpRec(CurrentPage, CurrentRow, 32))>0 then %>black<%else%>White<%end if%>"></TD>
				</TR>
				<%next%>
				<input name="empid" type="hidden" value="">
			</table>
		</td>
	</tr>
	<tr>
		<td align="CENTER">
			<TABLE class="txt" cellspacing="3" cellpadding="3">
				<tr>
					<td align="CENTER" >

					<% If CurrentPage > 1 Then %>
						<input type="submit" name="send" value="FIRST" class="btn btn-sm btn-outline-secondary">
						<input type="submit" name="send" value="BACK" class="btn btn-sm btn-outline-secondary">
					<% Else %>
						<input type="submit" name="send" value="FIRST" disabled class="btn btn-sm btn-outline-secondary">
						<input type="submit" name="send" value="BACK" disabled class="btn btn-sm btn-outline-secondary">
					<% End If %>
					<% If cint(CurrentPage) < cint(TotalPage) Then %>
						<input type="submit" name="send" value="NEXT" class="btn btn-sm btn-outline-secondary">
						<input type="submit" name="send" value="END" class="btn btn-sm btn-outline-secondary">
					<% Else %>
						<input type="submit" name="send" value="NEXT" disabled class="btn btn-sm btn-outline-secondary">
						<input type="submit" name="send" value="END" disabled class="btn btn-sm btn-outline-secondary">
					<% End If %>　
					PAGE:<%=CURRENTPAGE%> / <%=TOTALPAGE%> , COUNT:<%=RECORDINDB%>
					</td>
					<td align=right>
						<input type="button" name="send" value="(M)回主畫面 về trang trước" class="btn btn-sm btn-outline-secondary" onclick=BACKMAIN()>
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
	open "empwork.fore1.asp" , "_self"
end function

function oktest(index)
	tp=<%=self%>.totalpage.value
	cp=<%=self%>.CurrentPage.value
	rc=<%=self%>.RecordInDB.value
	empid = <%=self%>.empid(index).value  
	uid="<%=session("netuser")%>"
	'open "empworkB.fore.asp?empautoid="& N &"&yymm="&"<%=calcmonth%>", "_self"
	w=screen.width*0.8
	h=screen.height*0.7
	
	if <%=self%>.uid.value="LSARY" then 
		open "empworkBN.fore.asp?empid="& empid  &"&YYMM="&"<%=calcmonth%>" &"&Ftotalpage=" & tp &"&Fcurrentpage=" & cp &"&FRecordInDB=" & rc , "_blank" , "top=10, left=10, width="&w&", height="&h&", scrollbars=yes,resizable=yes"
	else
		'if uid="PELIN" then 
		'	open "empworkBN.fore.asp?empid="& empid &"&YYMM="&"<%=calcmonth%>" , "_self" , "top=10, left=10, width="&w&", height="&h&", scrollbars=yes,statusbar=yes"
		'else			
			open "empworkBN.fore.asp?empid="& empid &"&YYMM="&"<%=calcmonth%>" &"&Ftotalpage=" & tp &"&Fcurrentpage=" & cp &"&FRecordInDB=" & rc , "_blank" , "top=10, left=10, width="&w&", height="&h&", scrollbars=yes"
		'end if 
	end if	
end function

</script>


<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<!-- #include file = "../GetSQLServerConnection.fun" -->
<!-- #include file="../ADOINC.inc" -->
<%
SESSION.CODEPAGE="65001"
SELF = "Getempwork"

Set conn = GetSQLServerConnection()
Set rs = Server.CreateObject("ADODB.Recordset")
Set RST = Server.CreateObject("ADODB.Recordset")

gTotalPage = 1
PageRec = 10    'number of records per page
TableRec = 30    'number of fields per record


YYMM = REQUEST("YYMM")
'response.write yymm
IF YYMM="" THEN
	YYMM = year(date())&right("00"&month(date()),2)
	'YYMM="200601"
	cDatestr=date()
	days = DAY(cDatestr+(32-DAY(cDatestr))-DAY(cDatestr+(32-DAY(cDatestr))))   '一個月有幾天
ELSE
	cDatestr=CDate(LEFT(YYMM,4)&"/"&RIGHT(YYMM,2)&"/01")
	days = DAY(cDatestr+(32-DAY(cDatestr))-DAY(cDatestr+(32-DAY(cDatestr))))   '一個月有幾天
END IF


if month(date())="01" then
	calcmonth = year(date()-1)&"12"
else
	calcmonth =  year(date())&right("00"&month(date())-1,2)
end if
EMPID = TRIM(REQUEST("EMPID"))
empautoid = TRIM(REQUEST("empautoid"))

Ftotalpage = request("Ftotalpage")
Fcurrentpage = request("Fcurrentpage")
FRecordInDB = request("FRecordInDB")
'RESPONSE.END

'--------------------------------------------------------------------------------------
SQL="SELECT convert(char(10), indat, 111) as Nindate, x.lgstr as groupstr, x.lzstr as zunostr, c.ljstr as jobstr , "&_
		"x.lw as whsno, x.lg groupid, x.lz zuno , x.ls shot , c.lj job, a.* from  "&_
		"( SELECT * FROM  empfile WHERE ISNULL(STATUS,'')<>'D' AND EMPID='"& EMPID &"' ) a "&_
		"left join ( SELECT * FROM  view_empgroup WHERE EMPID='"& EMPID &"'  and yymm=convert(char(6), getdate(),112) ) x on x.empid = a.empid  "&_ 
		"left join ( SELECT * FROM  view_empjob   WHERE EMPID='"& EMPID &"'  and yymm=convert(char(6), getdate(),112) ) c on c.empid = a.empid  "
	
		'RESPONSE.WRITE SQL
		'RESPONSE.END
	RST.OPEN SQL , CONN, 3, 3
IF NOT RST.EOF THEN
	empautoid = TRIM(RST("AUTOID"))
	EMPID=TRIM(RST("EMPID"))	'員工編號
	INDAT=TRIM(RST("Nindate"))	'到職日
	TX=TRIM(RST("TX"))	'特休
	WHSNO=TRIM(RST("WHSNO"))	'廠別
	UNITNO= "" 'TRIM(RST("UNITNO"))	'處/所
	GROUPID=TRIM(RST("GROUPID"))	'組/部門
	ZUNO=TRIM(RST("ZUNO"))	'單位
	JOB=TRIM(RST("JOB"))	'職等
	EMPNAM_CN=TRIM(RST("EMPNAM_CN"))	'姓名(中)
	EMPNAM_VN=TRIM(RST("EMPNAM_VN"))	'姓名(越)
	COUNTRY=TRIM(RST("COUNTRY"))	'國籍
	GROUPSTR = TRIM(RST("GROUPSTR"))  '組/部門
	ZUNOSTR = TRIM(RST("ZUNOSTR"))  '單位
	JOBSTR = TRIM(RST("JOBSTR"))  '職等
	outdat = TRIM(RST("outdat"))  '離職日
END IF
SET RST=NOTHING


gTotalPage = 1
'PageRec = 31    'number of records per page
if yymm="" then
	PageRec = 31
else
	PageRec = days
end if
TableRec = 25    'number of fields per record

'出缺勤紀錄 --------------------------------------------------------------------------------------
sqlstr="exec SP_ChkEMPworkTime  '"& YYMM &"', '"& empid &"' "
if request("TotalPage") = "" or request("TotalPage") = "0" then
	CurrentPage = 1
	'RESPONSE.WRITE SQLSTR
	'RESPONSE.END
	rs.Open sqlstr, conn, 3, 3
	IF NOT RS.EOF THEN
		rs.PageSize = PageRec
		RecordInDB = days 'rs.RecordCount
		TotalPage = 1 'rs.PageCount
		gTotalPage = TotalPage
	END IF

	Redim tmpRec(gTotalPage, PageRec, TableRec)   'Array

	for i = 1 to TotalPage
	 for j = 1 to PageRec
		if not rs.EOF then
			tmpRec(i, j, 0) = "no"
			tmpRec(i, j, 1) = trim(rs("DAT"))
			tmpRec(i, j, 2) = RS("timeup")
			tmpRec(i, j, 3) = RS("timedown")
			tmpRec(i, j, 4) = (RS("toth"))
			tmpRec(i, j, 5) = (RS("kzhour"))
			tmpRec(i, j, 6) = (RS("FL"))
			tmpRec(i, j, 7) = (RS("H1"))
			tmpRec(i, j, 8) = (RS("H2"))
			tmpRec(i, j, 9) = (RS("H3"))
			tmpRec(i, j, 10) = (RS("B3"))
			tmpRec(i, j, 11) = (RS("hhoura"))
			tmpRec(i, j, 12) = (RS("hhourb"))
			tmpRec(i, j, 13) = (RS("hhourc"))
			tmpRec(i, j, 14) = (RS("hhourd"))
			tmpRec(i, j, 15) = (RS("hhoure"))
			tmpRec(i, j, 16) = (RS("hhourf"))
			tmpRec(i, j, 17) = (RS("hhourg"))
			tmpRec(i, j, 18) = (RS("hhourh"))
			tmpRec(i, j, 19)= mid("日一二三四五六",weekday(tmpRec(i, j, 1)) , 1 )
			tmpRec(i, j, 20) = (RS("hhouri"))
			tmpRec(i, j, 21) = (RS("B4"))
			tmpRec(i, j, 22) = (RS("B5"))
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
	Session("empworkb") = tmpRec
else
	TotalPage = cint(request("TotalPage"))
	'StoreToSession()
	CurrentPage = cint(request("CurrentPage"))
	RecordInDB  = REQUEST("RecordInDB")
	tmpRec = Session("empworkb")

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



'--------------------------------------------------------------------------------------
FUNCTION FDT(D)
IF D <> "" THEN
	Response.Write YEAR(D)&"/"&RIGHT("00"&MONTH(D),2)&"/"&RIGHT("00"&DAY(D),2)
END IF
END FUNCTION
'--------------------------------------------------------------------------------------
SQL="SELECT * FROM BASICCODE WHERE FUNC='CLOSEP' AND SYS_TYPE='"& YYMM &"' "
SET RDS=CONN.EXECUTE(SQL)
IF RDS.EOF THEN
	PCNTFG = 1 '可異動
	MSGSTR=""
ELSE
	PCNTFG = 0 '不可異動該月出勤紀錄
	MSGSTR="已結算，不可異動"
END IF
SET RDS=NOTHING
IF PCNTFG = "0" THEN
	INPUTSTS="READONLY"
ELSE
	INPUTSTS="INPUTBOX"
END IF
'---------------------------------------------------------------------------------
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
	<%=self%>.TIMEUP(0).SELECT()
end function

function colschg(index)
	thiscols = document.activeElement.name
	if window.event.keyCode = 38 then
		IF INDEX<>0 THEN
			document.all(thiscols)(index-1).SELECT()
		END IF
	end if
	if window.event.keyCode = 40 then
		document.all(thiscols)(index+1).SELECT()
	end if

end function

-->
</SCRIPT>
</head>
<body  topmargin="10" leftmargin="10"  marginwidth="0" marginheight="0"    >
<form name="<%=self%>"  method="post" >
<INPUT TYPE=HIDDEN NAME="PCNTFG" VALUE=<%=PCNTFG%>>
<INPUT TYPE=HIDDEN NAME="UID" VALUE=<%=SESSION("NETUSER")%>>
<INPUT TYPE=HIDDEN NAME="empautoid" VALUE=<%=empautoid%>>
<INPUT TYPE=hidden NAME=TotalPage VALUE="<%=TotalPage%>">
<INPUT TYPE=hidden NAME=CurrentPage VALUE="<%=CurrentPage%>">
<INPUT TYPE=hidden NAME=RecordInDB VALUE="<%=RecordInDB%>">
<INPUT TYPE=hidden NAME=FTotalPage VALUE="<%=FTotalPage%>">
<INPUT TYPE=hidden NAME=FCurrentPage VALUE="<%=FCurrentPage%>">
<INPUT TYPE=hidden NAME=FRecordInDB VALUE="<%=FRecordInDB%>">
<INPUT TYPE=hidden NAME=PageRec VALUE="<%=PageRec%>">
<table width=500  cellspacing="1" cellpadding="1" border=0 class="txt" bgcolor="#e4e4e4">
	<TR bgcolor="#ffffff">
		<td width=60 align="right" >查詢年月<br>YYYYMM</td>
		<td COLSPAN=3>
			<select name=yymm class="font9" onchange="dchg()" >
				<%for z = 1 to 12
				  yymmvalue = year(date())&right("00"&z,2)
				%>
					<option value="<%=yymmvalue%>" <%if yymmvalue=yymm then %>selected<%end if%>><%=yymmvalue%></option>
				<%next%>
			</select>
			<input type=hidden class=readonly readonly  name=days value="<%=days%>" size=5>
			　<FONT COLOR=RED><%=MSGSTR%></FONT>
		</td>
	</TR>
	<tr bgcolor="#ffffff">
		<td width=60 align="right">員工編號<br>SO the</td>
		<td valign="top" width=180><%=EMPID%>&nbsp;<%=empnam_cn%>&nbsp;<%=empnam_vn%>
			<input type="hidden" name=empid value="<%=EMPID%>" size=7 class="readonly"  >
			<input type="hidden" name=empnam value="<%=empnam_cn&" "&empnam_vn%>"  >
		</td>
		<td width=60 align="right">到職日期<br>NVX</td>
		<td  valign="top"><%=indat%>
			<input type="hidden" name=indat value="<%=indat%>" size=11 class="readonly" readonly  style="height:22">
		</td>
	</tr>
	<tr bgcolor="#ffffff" >
		<td align="right">單位<br>Bo phan</td>
		<td valign="top"><%=groupid%><%=GROUPSTR%>,<%=zunoSTR%>
			<input type="hidden" name=groupidstr value="<%=GROUPSTR%>" size=7 class="readonly" readonly  style="height:22">
			<input type="hidden"  name=zunostr value="<%=zunoSTR%>" size=5 class="readonly" readonly style="height:22" >
			<input TYPE=HIDDEN name=groupid value="<%=groupid%>" size=5 >
			<input TYPE=HIDDEN name=zuno value="<%=zuno%>" size=5 >
		</td>
		<td  align="right">離職日期<BR>NTV</td>
		<td valign="top" ><%=outdat%>&nbsp;
			<input type="hidden" name=outdat value="<%=outdat%>" size=11 class="readonly" readonly  style="height:22">
		</td>
	</tr>
	<Tr bgcolor="#ffffff">
		<td align="right">職等<br>Chuc vu</td>
		<td colspan=3 valign="top" ><%=job%>-<%=jobSTR%>
			<input type="hidden" name=job value="<%=jobSTR%>" size=12 class="readonly" readonly  style="height:22">
			<input type="hidden" name=TX value="<%=tx%>" size=5 class="readonly" readonly  style="height:22">
			<input type="hidden"  name=TXH value="<%=tx*8%>" size=5 class="readonly" readonly  style="height:22">
		</td>
	</tr>
</table> 
 
<hr size=0	style='border: 1px dotted #999999;' align=left width=500>
<TABLE BGCOLOR="#CCCCCC" BORDER=0 border="1" cellspacing="1"  class=txt8 >
	<TR BGCOLOR=#FFFFCC>
		<TD ROWSPAN=2 ALIGN=CENTER>日期</TD>
		<TD ROWSPAN=2 ALIGN=CENTER>上班</TD>
		<TD ROWSPAN=2 ALIGN=CENTER>下班</TD>
		<TD ROWSPAN=2 ALIGN=CENTER>工時</TD>
		<TD ROWSPAN=2 ALIGN=CENTER>曠職</TD>
		<TD ROWSPAN=2 ALIGN=CENTER>忘<br>遲<br>早</TD>
		<TD COLSPAN=6 ALIGN=CENTER>加班(單位：小時)</TD>
		<TD COLSPAN=9 ALIGN=CENTER>休假(單位：小時)</TD>
	</TR>
	<TR BGCOLOR=#FFFFCC>
		<TD ALIGN=CENTER>一般(1.5)</TD>
		<TD ALIGN=CENTER>休息(2.0)</TD>
		<TD ALIGN=CENTER>假日(3.0)</TD>
		<TD ALIGN=CENTER>津貼(0.5)</TD>
		<TD ALIGN=CENTER>夜班(2.1)</TD>
		<TD ALIGN=CENTER>津貼(0.3)</TD>
		<TD ALIGN=CENTER>事假(A)</TD>
		<TD ALIGN=CENTER>病假(B)</TD>
		<TD ALIGN=CENTER>年假(E)</TD>
		<TD ALIGN=CENTER>婚假(C)</TD>
		<TD ALIGN=CENTER>喪假(D)</TD>
		<TD ALIGN=CENTER>產假(F)</TD>
		<TD ALIGN=CENTER>公假(G)</TD>
		<TD ALIGN=CENTER>工傷(H)</TD>
		<TD ALIGN=CENTER>返鄉(I)</TD>
	</TR>
	<%
	sum_TOTHOUR = 0
	sum_KZhour = 0
	sum_Forget = 0
	sum_H1 = 0
	sum_H2 = 0
	sum_H3 = 0
	sum_B3 = 0
	sum_B4 = 0
	sum_B5 = 0
	um_JIAA = 0
	sum_JIAB = 0
	sum_JIAC = 0
	sum_JIAD = 0
	sum_JIAE = 0
	sum_JIAF = 0
	sum_JIAG = 0
	sum_JIAH = 0
	sum_JIAi = 0
	sum_LATEFOR = 0
	for CurrentRow = 1 to PageRec

		IF tmpRec(CurrentPage, CurrentRow, 18)<>"H1" THEN
			WKCOLOR = "#FFFFFF"
		ELSE
			IF CurrentRow MOD 2 = 0 THEN
				'WKCOLOR="LavenderBlush"
				wkcolor="#FFFFCC"
			ELSE
				WKCOLOR=""
			END IF
		END IF
		'if tmpRec(CurrentPage, CurrentRow, 1) <> "" then
	%>
	<TR BGCOLOR="<%=WKCOLOR%>">
		<TD ALIGN=CENTER NOWRAP >
		<INPUT NAME=WORKDATIM VALUE="<%=tmpRec(CurrentPage, CurrentRow, 1)&"("&tmpRec(CurrentPage, CurrentRow, 19)&")"%>" CLASS="READONLY8s" READONLY  SIZE=12 STYLE="TEXT-ALIGN:CENTER;color:<%if weekday(tmpRec(CurrentPage, CurrentRow, 1))=1 then %>#cc3300<%else%>black<%end if%>">
		<INPUT TYPE=HIDDEN NAME=WORKDAT VALUE="<%=tmpRec(CurrentPage, CurrentRow, 1)%>" >
		<INPUT TYPE=HIDDEN NAME=STATUS VALUE="<%=tmpRec(CurrentPage, CurrentRow, 18)%>"  >
		</TD>
		<TD ALIGN=CENTER><INPUT NAME=TIMEUP VALUE="<%=tmpRec(CurrentPage, CurrentRow, 2)%>" CLASS="READONLY8s" readonly SIZE=5 STYLE="TEXT-ALIGN:CENTER ;color:<%if weekday(tmpRec(CurrentPage, CurrentRow, 1))=1 then %>#cc3300<%else%>black<%end if%>"></TD>
		<TD ALIGN=CENTER><INPUT NAME=TIMEDOWN VALUE="<%=tmpRec(CurrentPage, CurrentRow, 3)%>" CLASS="READONLY8s" readonly SIZE=5 STYLE="TEXT-ALIGN:CENTER;color:<%if weekday(tmpRec(CurrentPage, CurrentRow, 1))=1 then %>#cc3300<%else%>black<%end if%>" ></TD>
		<TD ALIGN=CENTER><INPUT NAME=TOTHOUR VALUE="<%=tmpRec(CurrentPage, CurrentRow, 4)%>" CLASS="READONLY8s" readonly  SIZE=3 STYLE="TEXT-ALIGN:RIGHT"    ></TD>
		<TD ALIGN=CENTER><INPUT NAME=KZhour VALUE="<%=tmpRec(CurrentPage, CurrentRow, 5)%>" CLASS="READONLY8s" readonly SIZE=3 STYLE="TEXT-ALIGN:RIGHT;color:<%if tmpRec(CurrentPage, CurrentRow, 5)<>"0" then %>red<%else%>#FFFFEC<%end if%> " ></TD>
		<TD ALIGN=CENTER><INPUT NAME=Forget VALUE="<%=tmpRec(CurrentPage, CurrentRow, 6)%>" CLASS="READONLY8s" readonly SIZE=3 STYLE="TEXT-ALIGN:RIGHT;color:<%if tmpRec(CurrentPage, CurrentRow, 6)<>"0" then %>red<%else%>#FFFFEC<%end if%> " ></TD>
		<TD ALIGN=CENTER bgcolor="#FBE5CE"><INPUT NAME=H1 VALUE="<%=tmpRec(CurrentPage, CurrentRow, 7)%>" CLASS="READONLY8s" readonly SIZE=3 STYLE="TEXT-ALIGN:RIGHT;color:<%if tmpRec(CurrentPage, CurrentRow, 7)<>"0" then %>red<%else%>#FFFFEC<%end if%> " ></TD>
		<TD ALIGN=CENTER bgcolor="#D5FBDF"><INPUT NAME=H2 VALUE="<%=tmpRec(CurrentPage, CurrentRow, 8)%>" CLASS="READONLY8s" readonly SIZE=3 STYLE="TEXT-ALIGN:RIGHT;color:<%if tmpRec(CurrentPage, CurrentRow, 8)<>"0" then %>red<%else%>#FFFFEC<%end if%> " ></TD>
		<TD ALIGN=CENTER bgcolor="#F4DCFB"><INPUT NAME=H3 VALUE="<%=tmpRec(CurrentPage, CurrentRow, 9)%>" CLASS="READONLY8s" readonly SIZE=3 STYLE="TEXT-ALIGN:RIGHT;color:<%if tmpRec(CurrentPage, CurrentRow, 9)<>"0" then %>red<%else%>#FFFFEC<%end if%>"</TD>
		<TD ALIGN=CENTER bgcolor="#E1E7FF"><INPUT NAME=B4 VALUE="<%=tmpRec(CurrentPage, CurrentRow, 21)%>" CLASS="READONLY8s" readonly SIZE=3 STYLE="TEXT-ALIGN:RIGHT;color:<%if tmpRec(CurrentPage, CurrentRow, 21)<>"0" then %>red<%else%>#FFFFEC<%end if%>" ></TD>
		<TD ALIGN=CENTER bgcolor="#E1E7FF"><INPUT NAME=B3 VALUE="<%=tmpRec(CurrentPage, CurrentRow, 10)%>" CLASS="READONLY8s" readonly SIZE=3 STYLE="TEXT-ALIGN:RIGHT;color:<%if tmpRec(CurrentPage, CurrentRow, 10)<>"0" then %>red<%else%>#FFFFEC<%end if%>" ></TD>
		<TD ALIGN=CENTER bgcolor="#E1E7FF"><INPUT NAME=B5 VALUE="<%=tmpRec(CurrentPage, CurrentRow, 22)%>" CLASS="READONLY8s" readonly SIZE=3 STYLE="TEXT-ALIGN:RIGHT;color:<%if tmpRec(CurrentPage, CurrentRow, 22)<>"0" then %>red<%else%>#FFFFEC<%end if%>" ></TD>
		<TD ALIGN=CENTER><INPUT NAME=JIAA VALUE="<%=tmpRec(CurrentPage, CurrentRow, 11)%>" CLASS="READONLY8s" readonly  SIZE=3 STYLE="TEXT-ALIGN:RIGHT;color:<%if tmpRec(CurrentPage, CurrentRow, 11)<>"0" then %>red<%else%>#FFFFEC<%end if%>" ></TD>
		<TD ALIGN=CENTER><INPUT NAME=JIAB VALUE="<%=tmpRec(CurrentPage, CurrentRow, 12)%>" CLASS="READONLY8s" readonly  SIZE=3 STYLE="TEXT-ALIGN:RIGHT;color:<%if tmpRec(CurrentPage, CurrentRow, 12)<>"0" then %>red<%else%>#FFFFEC<%end if%>"></TD>
		<TD ALIGN=CENTER><INPUT NAME=JIAE VALUE="<%=tmpRec(CurrentPage, CurrentRow, 15)%>" CLASS="READONLY8s" readonly  SIZE=3 STYLE="TEXT-ALIGN:RIGHT;color:<%if tmpRec(CurrentPage, CurrentRow, 15)<>"0" then %>red<%else%>#FFFFEC<%end if%>"></TD>
		<TD ALIGN=CENTER><INPUT NAME=JIAC VALUE="<%=tmpRec(CurrentPage, CurrentRow, 13)%>" CLASS="READONLY8s" readonly  SIZE=3 STYLE="TEXT-ALIGN:RIGHT;color:<%if tmpRec(CurrentPage, CurrentRow, 13)<>"0" then %>red<%else%>#FFFFEC<%end if%>"></TD>
		<TD ALIGN=CENTER><INPUT NAME=JIAD VALUE="<%=tmpRec(CurrentPage, CurrentRow, 14)%>" CLASS="READONLY8s" readonly  SIZE=3 STYLE="TEXT-ALIGN:RIGHT;color:<%if tmpRec(CurrentPage, CurrentRow, 14)<>"0" then %>red<%else%>#FFFFEC<%end if%>"></TD>
		<TD ALIGN=CENTER><INPUT NAME=JIAF VALUE="<%=tmpRec(CurrentPage, CurrentRow, 16)%>" CLASS="READONLY8s" readonly  SIZE=3 STYLE="TEXT-ALIGN:RIGHT;color:<%if tmpRec(CurrentPage, CurrentRow, 16)<>"0" then %>red<%else%>#FFFFEC<%end if%>" ></TD>
		<TD ALIGN=CENTER><INPUT NAME=JIAG VALUE="<%=tmpRec(CurrentPage, CurrentRow, 17)%>" CLASS="READONLY8s" readonly  SIZE=3 STYLE="TEXT-ALIGN:RIGHT;color:<%if tmpRec(CurrentPage, CurrentRow, 17)<>"0" then %>red<%else%>#FFFFEC<%end if%>" ></TD>
		<TD ALIGN=CENTER><INPUT NAME=JIAh VALUE="<%=tmpRec(CurrentPage, CurrentRow, 18)%>" CLASS="READONLY8s" readonly  SIZE=3 STYLE="TEXT-ALIGN:RIGHT;color:<%if tmpRec(CurrentPage, CurrentRow, 18)<>"0" then %>red<%else%>#FFFFEC<%end if%>" ></TD>
		<TD ALIGN=CENTER><INPUT NAME=JIAI VALUE="<%=tmpRec(CurrentPage, CurrentRow, 20)%>" CLASS="READONLY8s" readonly  SIZE=3 STYLE="TEXT-ALIGN:RIGHT;color:<%if tmpRec(CurrentPage, CurrentRow, 20)<>"0" then %>red<%else%>#FFFFEC<%end if%>" ></TD>
	</TR>
	<%
		if isnull(tmpRec(CurrentPage, CurrentRow, 4)) or tmpRec(CurrentPage, CurrentRow, 4)="" then X1=0 else X1=cdbl(tmpRec(CurrentPage, CurrentRow, 4))
		if isnull(tmpRec(CurrentPage, CurrentRow, 4)) or tmpRec(CurrentPage, CurrentRow, 5)="" then X2=0 else X2=cdbl(tmpRec(CurrentPage, CurrentRow, 5))
		if isnull(tmpRec(CurrentPage, CurrentRow, 4)) or tmpRec(CurrentPage, CurrentRow, 6)="" then X3=0 else X3=cdbl(tmpRec(CurrentPage, CurrentRow, 6))
		if isnull(tmpRec(CurrentPage, CurrentRow, 4)) or tmpRec(CurrentPage, CurrentRow, 7)="" then X4=0 else X4=cdbl(tmpRec(CurrentPage, CurrentRow, 7))
		if isnull(tmpRec(CurrentPage, CurrentRow, 4)) or tmpRec(CurrentPage, CurrentRow, 8)="" then X5=0 else X5=cdbl(tmpRec(CurrentPage, CurrentRow, 8))
		if isnull(tmpRec(CurrentPage, CurrentRow, 4)) or tmpRec(CurrentPage, CurrentRow, 9)="" then X6=0 else X6=cdbl(tmpRec(CurrentPage, CurrentRow, 9))
		if isnull(tmpRec(CurrentPage, CurrentRow, 4)) or tmpRec(CurrentPage, CurrentRow, 10)="" then X7=0 else X7=cdbl(tmpRec(CurrentPage, CurrentRow,10))
		if isnull(tmpRec(CurrentPage, CurrentRow, 4)) or tmpRec(CurrentPage, CurrentRow, 11)="" then X8=0 else X8=cdbl(tmpRec(CurrentPage, CurrentRow, 11))
		if isnull(tmpRec(CurrentPage, CurrentRow, 4)) or tmpRec(CurrentPage, CurrentRow, 12)="" then X9=0 else X9=cdbl(tmpRec(CurrentPage, CurrentRow, 12))
		if isnull(tmpRec(CurrentPage, CurrentRow, 4)) or tmpRec(CurrentPage, CurrentRow, 13)="" then X10=0 else X10=cdbl(tmpRec(CurrentPage, CurrentRow, 13))
		if isnull(tmpRec(CurrentPage, CurrentRow, 4)) or tmpRec(CurrentPage, CurrentRow, 14)="" then X11=0 else X11=cdbl(tmpRec(CurrentPage, CurrentRow, 14))
		if isnull(tmpRec(CurrentPage, CurrentRow, 4)) or tmpRec(CurrentPage, CurrentRow, 15)="" then X12=0 else X12=cdbl(tmpRec(CurrentPage, CurrentRow, 15))
		if isnull(tmpRec(CurrentPage, CurrentRow, 4)) or tmpRec(CurrentPage, CurrentRow, 16)="" then X13=0 else X13=cdbl(tmpRec(CurrentPage, CurrentRow, 16))
		if isnull(tmpRec(CurrentPage, CurrentRow, 4)) or tmpRec(CurrentPage, CurrentRow, 17)="" then X14=0 else X14=cdbl(tmpRec(CurrentPage, CurrentRow, 17))
		if isnull(tmpRec(CurrentPage, CurrentRow, 4)) or tmpRec(CurrentPage, CurrentRow, 18)="" then X15=0 else X15=cdbl(tmpRec(CurrentPage, CurrentRow, 18))
		if isnull(tmpRec(CurrentPage, CurrentRow, 4)) or tmpRec(CurrentPage, CurrentRow, 20)="" then X16=0 else X16=cdbl(tmpRec(CurrentPage, CurrentRow, 20))
		if isnull(tmpRec(CurrentPage, CurrentRow, 4)) or tmpRec(CurrentPage, CurrentRow, 21)="" then X17=0 else X17=cdbl(tmpRec(CurrentPage, CurrentRow,21))
		if isnull(tmpRec(CurrentPage, CurrentRow, 4)) or tmpRec(CurrentPage, CurrentRow, 22)="" then X18=0 else X18=cdbl(tmpRec(CurrentPage, CurrentRow,22))
		sum_TOTHOUR = sum_TOTHOUR  + X1
		sum_KZhour  = sum_KZhour + X2
		sum_Forget  = sum_Forget + X3
		sum_H1 = sum_H1 + X4
		sum_H2 = sum_H2 + X5
		sum_H3 = sum_H3 + X6
		sum_B3 = sum_B3 + X7
		sum_B4 = sum_B4 + X17
		sum_B5 = sum_B5 + X18
		sum_JIAA = sum_JIAA + X8
		sum_JIAB = sum_JIAB + X9
		sum_JIAC = sum_JIAC + X10
		sum_JIAD = sum_JIAD + X11
		sum_JIAE = sum_JIAE + X12
		sum_JIAF = sum_JIAF + X13
		sum_JIAG = sum_JIAG + X14
		sum_JIAH = sum_JIAH + X15
		sum_JIAI = sum_JIAI + X16
	%>
	<%next%>
	<tr BGCOLOR="Lavender" >
		<td align=right colspan=3 HEIGHT=22>總計</td>
		<td align=right ><INPUT NAME="sum_TOTHOUR" VALUE="<%=sum_TOTHOUR%>" CLASS="READONLY8s" READONLY  SIZE=3 STYLE="TEXT-ALIGN:RIGHT;color:#002CA5"></td>
		<td align=right ><INPUT NAME="sum_KZhour" VALUE="<%=sum_KZhour%>" CLASS="READONLY8s"   SIZE=3 STYLE="TEXT-ALIGN:RIGHT;color:#002CA5"></td>
		<td align=right ><INPUT NAME="sum_Forget" VALUE="<%=sum_Forget%>" CLASS="READONLY8s" READONLY  SIZE=3 STYLE="TEXT-ALIGN:RIGHT;color:#002CA5"></td>
		<td align=right ><INPUT NAME="sum_H1" VALUE="<%=sum_H1%>" CLASS="READONLY8s" READONLY  SIZE=3 STYLE="TEXT-ALIGN:RIGHT;color:#002CA5"></td>
		<td align=right ><INPUT NAME="sum_H2" VALUE="<%=sum_H2%>" CLASS="READONLY8s" READONLY  SIZE=3 STYLE="TEXT-ALIGN:RIGHT;color:#002CA5"></td>
		<td align=right ><INPUT NAME="sum_H3" VALUE="<%=sum_H3%>" CLASS="READONLY8s" READONLY  SIZE=3 STYLE="TEXT-ALIGN:RIGHT;color:#002CA5"></td>
		<td align=right ><INPUT NAME="sum_B4" VALUE="<%=sum_B4%>" CLASS="READONLY8s" READONLY  SIZE=3 STYLE="TEXT-ALIGN:RIGHT;color:#002CA5"></td>
		<td align=right ><INPUT NAME="sum_B3" VALUE="<%=sum_B3%>" CLASS="READONLY8s" READONLY  SIZE=3 STYLE="TEXT-ALIGN:RIGHT;color:#002CA5"></td>
		<td align=right ><INPUT NAME="sum_B5" VALUE="<%=sum_B5%>" CLASS="READONLY8s" READONLY  SIZE=3 STYLE="TEXT-ALIGN:RIGHT;color:#002CA5"></td>
		<td align=right ><INPUT NAME="sum_JIAA" VALUE="<%=sum_JIAA%>" CLASS="READONLY8s" READONLY  SIZE=3 STYLE="TEXT-ALIGN:RIGHT;color:#800000"></td>
		<td align=right ><INPUT NAME="sum_JIAB" VALUE="<%=sum_JIAB%>" CLASS="READONLY8s" READONLY  SIZE=3 STYLE="TEXT-ALIGN:RIGHT;color:#800000"></td>
		<td align=right ><INPUT NAME="sum_JIAE" VALUE="<%=sum_JIAE%>" CLASS="READONLY8s" READONLY  SIZE=3 STYLE="TEXT-ALIGN:RIGHT;color:#800000"></td>
		<td align=right ><INPUT NAME="sum_JIAC" VALUE="<%=sum_JIAC%>" CLASS="READONLY8s" READONLY  SIZE=3 STYLE="TEXT-ALIGN:RIGHT;color:#800000"></td>
		<td align=right ><INPUT NAME="sum_JIAD" VALUE="<%=sum_JIAD%>" CLASS="READONLY8s" READONLY  SIZE=3 STYLE="TEXT-ALIGN:RIGHT;color:#800000"></td>
		<td align=right ><INPUT NAME="sum_JIAF" VALUE="<%=sum_JIAF%>" CLASS="READONLY8s" READONLY  SIZE=3 STYLE="TEXT-ALIGN:RIGHT;color:#800000"></td>
		<td align=right ><INPUT NAME="sum_JIAG" VALUE="<%=sum_JIAG%>" CLASS="READONLY8s" READONLY  SIZE=3 STYLE="TEXT-ALIGN:RIGHT;color:#800000"></td>
		<td align=right ><INPUT NAME="sum_JIAH" VALUE="<%=sum_JIAH%>" CLASS="READONLY8s" READONLY  SIZE=3 STYLE="TEXT-ALIGN:RIGHT;color:#800000"></td>
		<td align=right ><INPUT NAME="sum_JIAI" VALUE="<%=sum_JIAI%>" CLASS="READONLY8s" READONLY  SIZE=3 STYLE="TEXT-ALIGN:RIGHT;color:#800000"></td>
	</tr>
</TABLE>

<TABLE border=0 width=680 class=font9 >
<tr>
    <td align="CENTER" height=40  >
	<input type=BUTTON name=send value="關閉此視窗(CLOSE)"  class=button ONCLICK="vbscript:window.close()">　　
	</td>

</TR>
</TABLE>

</form>


</body>
</html>

<script language=vbscript >
function BACKMAIN()
	open "../main.asp" , "_self"
end function

FUNCTION dchg()
	<%=self%>.totalpage.value=0
	<%=SELF%>.ACTION="getempworktime.ASP"
	<%=SELF%>.SUBMIT()

	'alert <%=self%>.yymm.value
	'ymstr = <%=self%>.yymm.value
	'tp=<%=self%>.totalpage.value
	'cp=<%=self%>.CurrentPage.value
	'rc=<%=self%>.RecordInDB.value
	'empid = <%=self%>.empid.value
	'open "getempworktime.asp?empid="& empid &"&YYMM="&ymstr  , "_self"

END FUNCTION

</script>



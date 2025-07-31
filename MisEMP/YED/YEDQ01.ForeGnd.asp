<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<!-- #include file = "../GetSQLServerConnection.fun" -->
<!-- #include file="../ADOINC.inc" -->
<!--#include file="../include/sideinfo.inc"-->
<%
SESSION.CODEPAGE="65001"
SELF = "YEDQ01"

Set conn = GetSQLServerConnection()
Set rs = Server.CreateObject("ADODB.Recordset")
Set RST = Server.CreateObject("ADODB.Recordset")

gTotalPage = 1
PageRec = 10    'number of records per page
TableRec = 30    'number of fields per record


YYMM = REQUEST("YYMM")
empid = request("empid")
fr=request("fr")
khweek=request("khweek")
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

nowmonth = year(date())&right("00"&month(date()),2) 
if month(date())="01" then
	calcmonth = year(date()-1)&"12"
else
	calcmonth =  year(date())&right("00"&month(date())-1,2)
end if
EMPID = TRIM(REQUEST("EMPID"))
 
'RESPONSE.END 
 	 


'--------------------------------------------------------------------------------------	
sql="select   * from view_empfile where  empid = '"& empid &"'  "   	
	'RESPONSE.WRItE SQL
	'RESPONSE.END
	RST.OPEN SQL , CONN, 1, 1
IF NOT RST.EOF THEN
	'empautoid = TRIM(RST("AUTOID"))
	EMPID=TRIM(RST("EMPID"))	'員工編號
	INDAT=TRIM(RST("Nindat"))	'到職日
	TX=TRIM(RST("TX"))	'特休
	WHSNO=TRIM(RST("WHSNO"))	'廠別
	UNITNO=TRIM(RST("UNITNO"))	'處/所
	GROUPID=TRIM(RST("GROUPID"))	'組/部門
	ZUNO=TRIM(RST("ZUNO"))	'單位
	JOB=TRIM(RST("JOB"))	'職等
	EMPNAM_CN=TRIM(RST("EMPNAM_CN"))	'姓名(中)
	EMPNAM_VN=TRIM(RST("EMPNAM_VN"))	'姓名(越)
	COUNTRY=TRIM(RST("COUNTRY"))	'國籍
	GROUPSTR = TRIM(RST("GSTR"))  '組/部門
	ZUNOSTR = TRIM(RST("ZSTR"))  '單位
	JOBSTR = TRIM(RST("JSTR"))  '職等
	outdat = TRIM(RST("outDATe"))  '離職日
	'shift = TRIM(RST("shift"))  '班別
	if TRIM(RST("shift")) ="A" then 
		shift="A班"	
	elseif TRIM(RST("shift")) ="B" then 
		shift="B班"
	elseif TRIM(RST("shift")) ="ALL" then 
		shift="常日班"	
	else
		shift=""
	end if 		
 	
END IF
SET RST=NOTHING


gTotalPage = 1
'PageRec = 31    'number of records per page
if yymm="" then
	PageRec = 31
else
	PageRec = days
end if
TableRec = 40    'number of fields per record

'出缺勤紀錄 --------------------------------------------------------------------------------------

SQLSTRA="exec SP_Query_empwork  '"& yymm &"', '"& empid &"' " 
'response.write sqlstra 
'response.end 
if request("TotalPage") = "" or request("TotalPage") = "0" then
	CurrentPage = 1
	'RESPONSE.WRITE SQLSTRA
	'RESPONSE.END
	rs.Open SQLSTRA, conn, 1, 3
	IF NOT RS.EOF THEN
		rs.PageSize = PageRec
		RecordInDB = days 'rs.RecordCount
		TotalPage = 1 'rs.PageCount
		gTotalPage = TotalPage
	END IF
	'response.write days 
	Redim tmpRec(gTotalPage, PageRec, TableRec)   'Array

	for i = 1 to TotalPage
	 for j = 1 to PageRec
		if not rs.EOF then
			tmpRec(i, j, 0) = "no"
			tmpRec(i, j, 1) = trim(rs("DDAT"))
			if rs("timeup")="000000"  then 
				tmpRec(i, j, 2)="" 
			else	
				tmpRec(i, j, 2) = RS("T1")
			end if 	
			if rs("timeup")="000000"  then 
				tmpRec(i, j, 3) = ""
			else
				tmpRec(i, j, 3) = RS("T2")			 
			end if 		
			tmpRec(i, j, 4) = rs("lsempid")
			tmpRec(i, j, 5) = rs("ftimeup")
			tmpRec(i, j, 6) = rs("ftimedown") 
			if trim(rs("fgdat"))<>"" then
				tmpRec(i, j, 7) = rs("fgtoth")
			else
				tmpRec(i, j, 7) = rs("toth")
			end if 	
			tmpRec(i, j, 8) = rs("weekstr")
			tmpRec(i, j, 9)  = trim(rs("fgdat"))
			tmpRec(i, j, 10)  = rs("kzhour")
			if trim(rs("fgdat"))<>"" and rs("lsempid")="" then 
				tmpRec(i, j, 11)="1"
			else	
				tmpRec(i, j, 11)="0"
			end if 	
			tmpRec(i, j, 12) = rs("lateFor")
			tmpRec(i, j, 13) = rs("h1")
			tmpRec(i, j, 14) = rs("h2")
			tmpRec(i, j, 15) = rs("h3")
			tmpRec(i, j, 16) = rs("b3")			
			tmpRec(i, j, 17) = rs("jiaG")
			tmpRec(i, j, 18) = rs("jiaE")
			tmpRec(i, j, 19) = rs("jiaA")
			tmpRec(i, j, 20) = rs("jiaB")
			tmpRec(i, j, 21) = rs("jiaC")
			tmpRec(i, j, 22) = rs("jiaD")
			tmpRec(i, j, 23) = rs("jiaF")
			tmpRec(i, j, 24) = rs("jiaH")
			'response.write tmpRec(i, j, 9) &"<BR>"
			
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
	Session("YEDQ01") = tmpRec 
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

-->
</SCRIPT>
</head>
<body  topmargin="5" leftmargin="5"  marginwidth="0" marginheight="0"   >
<form name="<%=self%>"  method="post" action = "<%=self%>.upd.asp" >
<INPUT TYPE=HIDDEN NAME="PCNTFG" VALUE="<%=PCNTFG%>">
<INPUT TYPE=HIDDEN NAME="UID" VALUE="<%=SESSION("NETUSER")%>">
<INPUT TYPE=HIDDEN NAME="empid" VALUE="<%=empid%>">
<INPUT TYPE=hidden NAME=TotalPage VALUE="<%=TotalPage%>">
<INPUT TYPE=hidden NAME=CurrentPage VALUE="<%=CurrentPage%>">
<INPUT TYPE=hidden NAME=RecordInDB VALUE="<%=RecordInDB%>">
<INPUT TYPE=hidden NAME=FTotalPage VALUE="<%=FTotalPage%>">
<INPUT TYPE=hidden NAME=FCurrentPage VALUE="<%=FCurrentPage%>">
<INPUT TYPE=hidden NAME=FRecordInDB VALUE="<%=FRecordInDB%>">
<INPUT TYPE=hidden NAME=PageRec VALUE="<%=PageRec%>">
<INPUT TYPE=hidden NAME=fr VALUE="<%=fr%>"> 
 
<table border=0 style="height:70px"><tr><td>&nbsp;</td></tr></table>
<table border=0 width="100%"><tr><td align="center">
<table width=620   class=txt border=1 >
	<TR>
		<td  align=right bgcolor="#e4e4e4">查詢年月:</td>
		<td><table border=0><tr>
			<td>
			<select name=yymm   onchange="dchg()" style="width:120px">
				<%for z = 1 to 24
					if   z mod 12 = 0  then 
						if Z\12 = 1  then 
							yy =year(date())-((z\12))
						else
							yy =year(date())
						end if 	
						zz = 12 
					elseif z > 12 and z mod 12 <> 0  then 
						yy = year(date())  
						zz = z mod 12
					else
						zz = z 
						yy = year(date()-365)						 
					end if 	
				  yymmvalue = yy&right("00"&zz,2)
				%>
					<option value="<%=yymmvalue%>" <%if yymmvalue=yymm then %>selected<%end if%>><%=yymmvalue%></option>
				<%next%> 
			</select>
			</td>
			<td><input type=hiddenT  readonly  name=days value="<%=days%>" size=5></td>
			<td>
			<select name=khweek   onchange="dchg()"  style="width:120px">
				<option value="">ALL</option>
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
					dd1 =  right("00"&yd1,2)
					dd2 =  right("00"&yd2,2)
					
				%>
				<option value="<%=yy%>" <%if cstr(yy) = cstr(khweek) then%>selected<%end if%>>第<%=yy%>週,<%=dd1%>~<%=dd2%> </option>
				<% next	%>
			</select> 
			</td>
			</tr></table>
		</td>
		<td align=right bgcolor="#e4e4e4">到職日期:</td>
		<td><%=indat%></td>
		<td rowspan=4><img src="../yeb/pic/<%=empid%>.jpg" border=1 width=100 height=120></td>
	</TR>
	<tr >
		<td width=60  align=right bgcolor="#e4e4e4">員工資料:</td>
		<td>
			<%=EMPID%>&nbsp;<%=empnam_cn&" "&empnam_vn%>
		</td>
		<td align=right bgcolor="#e4e4e4">單位:</td>
		<td>
			<%=groupid%>-<%=GROUPSTR%>-<%=shift%>			
		</td>
	</tr>
	<tr>
		<td align=right bgcolor="#e4e4e4">現在職等:</td>
		<td><%=job%>-<%=jobSTR%></td>
		<td  align=right bgcolor="#e4e4e4">離職日期:</td>
		<td><%=outdat%>&nbsp;</td>
	</tr>	
	<tr>
		<td align=right bgcolor="#e4e4e4">尚有特休:</td>
		<td colspan=3><%=tx%> (Day)= <%=tx*8%> (Hour)</td>		
	</tr>
</table> 

<TABLE CLASS=txt8  >
	<TR BGCOLOR=#e4e4e4>
		<TD ROWSPAN=2 ALIGN=CENTER width=90 nowrap  >日期</TD>
		<TD ROWSPAN=2 ALIGN=CENTER width=50 nowrap>臨時卡</TD>
		<TD ROWSPAN=2 ALIGN=CENTER width=40 nowrap>上班</TD>
		<TD ROWSPAN=2 ALIGN=CENTER width=40 nowrap>下班</TD>
		<TD ROWSPAN=2 ALIGN=CENTER width=40 nowrap>工時</TD>
		<TD ROWSPAN=2 ALIGN=CENTER width=40 nowrap>曠職</TD>
		<TD ROWSPAN=2 ALIGN=CENTER width=30 nowrap>忘<br>刷<br>卡</TD>
		<TD ROWSPAN=2 ALIGN=CENTER width=30 nowrap>遲到</TD>
		<TD COLSPAN=4 ALIGN=CENTER   nowrap>加班(單位：小時)</TD>
		<TD COLSPAN=8 ALIGN=CENTER   nowrap>休假(單位：小時)</TD>
	</TR>
	<TR BGCOLOR=#e4e4e4>
		<TD ALIGN=CENTER width=30 nowrap >一般(1.5)</TD>
		<TD ALIGN=CENTER width=30 nowrap>休息(2)</TD>
		<TD ALIGN=CENTER width=30 nowrap>假日(3)</TD>
		<TD ALIGN=CENTER width=30 nowrap>夜班(0.3)</TD>
		<TD ALIGN=CENTER width=30 nowrap>公假(G)</TD>
		<TD ALIGN=CENTER width=30 nowrap>年假(E)</TD>
		<TD ALIGN=CENTER width=30 nowrap>事假(A)</TD>
		<TD ALIGN=CENTER width=30 nowrap>病假(B)</TD>
		<TD ALIGN=CENTER width=30 nowrap>婚假(C)</TD>
		<TD ALIGN=CENTER width=30 nowrap>喪假(D)</TD>
		<TD ALIGN=CENTER width=30 nowrap>產假(F)</TD>
		<TD ALIGN=CENTER width=30 nowrap>工傷(H)</TD>
	</TR>
	<%
	sum_TOTHOUR = 0
	sum_KZhour = 0
	sum_Forget = 0
	sum_H1 = 0
	sum_H2 = 0
	sum_H3 = 0
	sum_B3 = 0
	sum_JIAA = 0
	sum_JIAB = 0
	sum_JIAC = 0
	sum_JIAD = 0
	sum_JIAE = 0
	sum_JIAF = 0
	sum_JIAG = 0
	sum_JIAH = 0
	sum_LATEFOR = 0
	if khweek="" then 
		PageRec = PageRec 
		topnum = 1
	else
		PageRec = cdbl(khweek)*7
		if khweek=4 and PageRec < days then 
			PageRec = days
		end if 
		if khweek=1 then 
			topnum=khweek
		else
			topnum=((cdbl(khweek)-1)*7)+1
		end if	
	end if	
	'response.write  "topnum=" & topnum &"<BR>"	
	for CurrentRow =  topnum to PageRec 
	'response.end 
		IF CurrentRow MOD 2 = 0 THEN
			WKCOLOR="LavenderBlush"
		ELSE
			WKCOLOR="LightYellow"		
		END IF
		'if tmpRec(CurrentPage, CurrentRow, 1) <> "" then
	%>
	<TR BGCOLOR=<%=WKCOLOR%>>
		<TD ALIGN=CENTER NOWRAP >
			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then %><%=tmpRec(CurrentPage, CurrentRow, 1)&"("&tmpRec(CurrentPage, CurrentRow,8)&")"%><%end if%>
		</TD>
		<TD ALIGN=CENTER><%=tmpRec(CurrentPage, CurrentRow, 4)%></TD>
		<TD ALIGN=CENTER>
			<%if tmpRec(CurrentPage,CurrentRow, 9)="" then %><%=tmpRec(CurrentPage, CurrentRow, 2)%><%else%><%=tmpRec(CurrentPage, CurrentRow, 5)%><%end if%>		
		</TD>
		<TD ALIGN=CENTER>
			<%if tmpRec(CurrentPage, CurrentRow, 9)="" then %><%=tmpRec(CurrentPage, CurrentRow,3)%><%else%><%=tmpRec(CurrentPage, CurrentRow, 6)%><%end if%>
		</TD>
		<TD ALIGN=right><%if tmpRec(CurrentPage, CurrentRow, 7)>"0" then %><%=tmpRec(CurrentPage, CurrentRow, 7)%><%end if%></TD>
		<TD ALIGN=CENTER><%if tmpRec(CurrentPage, CurrentRow, 10)>"0" then %><%=tmpRec(CurrentPage, CurrentRow, 10)%><%end if%></TD>		
		<TD ALIGN=CENTER><%if tmpRec(CurrentPage, CurrentRow, 11)>"0" then %><%=tmpRec(CurrentPage, CurrentRow, 11)%><%end if%></TD>		
		<TD ALIGN=CENTER><%if tmpRec(CurrentPage, CurrentRow, 12)>"0" then %><%=tmpRec(CurrentPage, CurrentRow, 12)%><%end if%></TD>		
		<TD ALIGN=CENTER bgcolor="#FBE5CE">
			<%if tmpRec(CurrentPage, CurrentRow, 13)>"0" then %><%=tmpRec(CurrentPage, CurrentRow, 13)%><%end if%>
		</TD>
		<TD ALIGN=CENTER bgcolor="#D5FBDF">
			<%if tmpRec(CurrentPage, CurrentRow, 14)>"0" then %><%=tmpRec(CurrentPage, CurrentRow, 14)%><%end if%>
		</TD>
		<TD ALIGN=CENTER bgcolor="#F4DCFB">
			<%if tmpRec(CurrentPage, CurrentRow, 15)>"0" then %><%=tmpRec(CurrentPage, CurrentRow, 15)%><%end if%>
		</TD>
		<TD ALIGN=CENTER bgcolor="#f5f5dc">
			<%if tmpRec(CurrentPage, CurrentRow, 16)>"0" then %><%=tmpRec(CurrentPage, CurrentRow, 16)%><%end if%>
		</TD>
		<TD ALIGN=CENTER>
			<%if tmpRec(CurrentPage, CurrentRow, 17)>"0" then %><%=tmpRec(CurrentPage, CurrentRow, 17)%><%end if%>	
		</TD>
		<TD ALIGN=CENTER>
			<%if tmpRec(CurrentPage, CurrentRow, 18)>"0" then %><%=tmpRec(CurrentPage, CurrentRow, 18)%><%end if%>	
		</TD>
		<TD ALIGN=CENTER>
			<%if tmpRec(CurrentPage, CurrentRow, 19)>"0" then %><%=tmpRec(CurrentPage, CurrentRow, 19)%><%end if%>	
		</TD>
		<TD ALIGN=CENTER>
			<%if tmpRec(CurrentPage, CurrentRow, 20)>"0" then %><%=tmpRec(CurrentPage, CurrentRow, 20)%><%end if%>	
		</TD>
		<TD ALIGN=CENTER>
			<%if tmpRec(CurrentPage, CurrentRow, 21)>"0" then %><%=tmpRec(CurrentPage, CurrentRow, 21)%><%end if%>	
		</TD>
		<TD ALIGN=CENTER>
			<%if tmpRec(CurrentPage, CurrentRow, 22)>"0" then %><%=tmpRec(CurrentPage, CurrentRow, 22)%><%end if%>	
		</TD>
		<TD ALIGN=CENTER>
			<%if tmpRec(CurrentPage, CurrentRow, 23)>"0" then %><%=tmpRec(CurrentPage, CurrentRow, 23)%><%end if%>	
		</TD>
		<TD ALIGN=CENTER>
			<%if tmpRec(CurrentPage, CurrentRow, 24)>"0" then %><%=tmpRec(CurrentPage, CurrentRow, 24)%><%end if%>	
		</TD>
	</TR>
	<%
		sum_TOTHOUR = sum_TOTHOUR + cdbl(tmpRec(CurrentPage, CurrentRow, 7))
		sum_LATEFOR  = sum_LATEFOR + cdbl(tmpRec(CurrentPage, CurrentRow, 12))
		sum_KZhour  = sum_KZhour + cdbl(tmpRec(CurrentPage, CurrentRow, 10))
		sum_Forget  = sum_Forget + cdbl(tmpRec(CurrentPage, CurrentRow, 11))
		sum_H1 = sum_H1 + cdbl(tmpRec(CurrentPage, CurrentRow, 13))
		sum_H2 = sum_H2 + cdbl(tmpRec(CurrentPage, CurrentRow, 14))
		sum_H3 = sum_H3 + cdbl(tmpRec(CurrentPage, CurrentRow, 15))
		sum_B3 = sum_B3 + cdbl(tmpRec(CurrentPage, CurrentRow, 16))
		sum_JIAA = sum_JIAA + cdbl(tmpRec(CurrentPage, CurrentRow, 19))
		sum_JIAB = sum_JIAB	+ cdbl(tmpRec(CurrentPage, CurrentRow, 21))
		sum_JIAC = sum_JIAC + cdbl(tmpRec(CurrentPage, CurrentRow, 21))
		sum_JIAD = sum_JIAD + cdbl(tmpRec(CurrentPage, CurrentRow, 22))
		sum_JIAE = sum_JIAE + cdbl(tmpRec(CurrentPage, CurrentRow, 18))
		sum_JIAF = sum_JIAF + cdbl(tmpRec(CurrentPage, CurrentRow, 23))
		sum_JIAG = sum_JIAG + cdbl(tmpRec(CurrentPage, CurrentRow, 17))
		sum_JIAH = sum_JIAH + cdbl(tmpRec(CurrentPage, CurrentRow, 24))
	%>
	<%next%>
	<tr BGCOLOR="Lavender" >
		<td align=right colspan=4 HEIGHT=22>總計</td>		
		<td align=right ><%if sum_tothour>"0" then %><%=sum_TOTHOUR%><%end if%></td>
		<td align=CENTER ><%if sum_KZhour>"0" then %><%=sum_KZhour%><%end if%></td>
		<td align=CENTER ><%if sum_Forget>"0" then %><%=sum_Forget%><%end if%></td>
		<td align=CENTER ><%if sum_LATEFOR>"0" then %><%=sum_LATEFOR%><%end if%></td>
		<td align=CENTER ><%if sum_H1>"0" then %><%=sum_H1%><%end if%></td>
		<td align=CENTER ><%if sum_H2>"0" then %><%=sum_H2%><%end if%></td>
		<td align=CENTER ><%if sum_H3>"0" then %><%=sum_H3%><%end if%></td>
		<td align=CENTER ><%if sum_B3>"0" then %><%=sum_B3%><%end if%></td>
		<td align=CENTER ><%if sum_JIAG>"0" then %><%=sum_JIAG%><%end if%></td>
		<td align=CENTER ><%if sum_JIAE>"0" then %><%=sum_JIAE%><%end if%></td>
		<td align=CENTER ><%if sum_JIAA>"0" then %><%=sum_JIAA%><%end if%></td>
		<td align=CENTER ><%if sum_JIAB>"0" then %><%=sum_JIAB%><%end if%></td>
		<td align=CENTER ><%if sum_JIAC>"0" then %><%=sum_JIAC%><%end if%></td>
		<td align=CENTER ><%if sum_JIAD>"0" then %><%=sum_JIAD%><%end if%></td>
		<td align=CENTER ><%if sum_JIAF>"0" then %><%=sum_JIAF%><%end if%></td>
		<td align=CENTER ><%if sum_JIAH>"0" then %><%=sum_JIAH%><%end if%></td>
	</tr>
</TABLE>

<TABLE border=0 width=600 class=font9 >
<tr>
    <td align="CENTER" height=40  >
    	<%if fr="A" then%>
    		<input type=BUTTON name=send value="(X)關閉視窗"  class="btn btn-sm btn-outline-secondary" ONCLICK="window.close()">
    	<%else%>
			<input type=BUTTON name=send value="查詢其他"  class="btn btn-sm btn-outline-secondary" ONCLICK="vbscript:gobak()">
		<%end if%>
	</td>	
</TR>
</TABLE>

</td></tr></table>
</form>


</body>
</html>

<script language=vbscript >
function gobak()
	open "<%=self%>.asp", "_self"
end function  
 
FUNCTION dchg()
	'<%=SELF%>.ACTION="empworkB.FORE.ASP"
	'<%=SELF%>.SUBMIT()
	'ymstr = <%=self%>.yymm.value
	'tp=<%=self%>.totalpage.value
	'cp=<%=self%>.CurrentPage.value
	'rc=<%=self%>.RecordInDB.value
	'n = <%=self%>.empid.value
	'open "<%=self%>.foreGnd.asp?empid="& N &"&YYMM="&ymstr &"&Ftotalpage=" & tp &"&Fcurrentpage=" & cp &"&FRecordInDB=" & rc , "_self"
	'alert <%=self%>.yymm.value
	<%=self%>.totalpage.value="0"
	<%=self%>.action="<%=self%>.foregnd.asp"
	<%=self%>.submit()
END FUNCTION
 


'*******檢查日期*********************************************
FUNCTION date_change(a)

if a=1 then
	INcardat = Trim(<%=self%>.indat.value)
elseif a=2 then
	INcardat = Trim(<%=self%>.BHDAT.value)
elseif a=3 then
	INcardat = Trim(<%=self%>.pduedate.value)
elseif a=4 then
	INcardat = Trim(<%=self%>.vduedate.value)
elseif a=5 then
	INcardat = Trim(<%=self%>.outdat.value)
end if

IF INcardat<>"" THEN
	ANS=validDate(INcardat)
	IF ANS <> "" THEN
		if a=1 then
			Document.<%=self%>.indat.value=ANS
		elseif a=2 then
			Document.<%=self%>.BHDAT.value=ANS
		elseif a=3 then
			Document.<%=self%>.pduedate.value=ANS
		elseif a=4 then
			Document.<%=self%>.vduedate.value=ANS
		elseif a=5 then
			Document.<%=self%>.outdat.value=ANS
		end if
	ELSE
		ALERT "EZ0067:輸入日期不合法 !!"
		if a=1 then
			Document.<%=self%>.indat.value=""
			Document.<%=self%>.indat.focus()
		elseif a=2 then
			Document.<%=self%>.BHDAT.value=""
			Document.<%=self%>.BHDAT.focus()
		elseif a=3 then
			Document.<%=self%>.pduedate.value=""
			Document.<%=self%>.pduedate.focus()
		elseif a=4 then
			Document.<%=self%>.vduedate.value=""
			Document.<%=self%>.vduedate.focus()
		elseif a=5 then
			Document.<%=self%>.outdat.value=""
			Document.<%=self%>.outdat.focus()
		end if
		EXIT FUNCTION
	END IF

ELSE
	'alert "EZ0015:日期該欄位必須輸入資料 !!"
	EXIT FUNCTION
END IF

END FUNCTION

'_________________DATE CHECK___________________________________________________________________

function validDate(d)
	if len(d) = 8 and isnumeric(d) then
		d = left(d,4) & "/" & mid(d, 5, 2) & "/" & right(d,2)
		if isdate(d) then
			validDate = formatDate(d)
		else
			validDate = ""
		end if
	elseif len(d) >= 8 and isdate(d) then
			validDate = formatDate(d)
		else
			validDate = ""
		end if
end function

function formatDate(d)
		formatDate = Year(d) & "/" & _
		Right("00" & Month(d), 2) & "/" & _
		Right("00" & Day(d), 2)
end function
'________________________________________________________________________________________

FUNCTION CHKVALUE(N)
IF N=1 THEN
	IF TRIM(<%=SELF%>.BYY.VALUE)<>"" THEN
		IF ISNUMERIC(<%=SELF%>.BYY.VALUE)=FALSE OR INSTR(1,<%=SELF%>.BYY.VALUE,"-")>0 THEN
			ALERT "輸入錯誤!!請輸入正確數字"
			<%=SELF%>.BYY.VALUE=""
			<%=SELF%>.BYY.FOCUS()
			EXIT FUNCTION
		END IF
	END IF
ELSEIF N=2 THEN
	IF TRIM(<%=SELF%>.BMM.VALUE)<>"" THEN
		IF ISNUMERIC(<%=SELF%>.BMM.VALUE)=FALSE OR INSTR(1,<%=SELF%>.BMM.VALUE,"-")>0 THEN
			ALERT "輸入錯誤!!請輸入正確數字"
			<%=SELF%>.BMM.VALUE=""
			<%=SELF%>.BMM.FOCUS()
			EXIT FUNCTION
		END IF
	END IF
ELSEIF N=3 THEN
	IF TRIM(<%=SELF%>.BDD.VALUE)<>"" THEN
		IF ISNUMERIC(<%=SELF%>.BDD.VALUE)=FALSE OR INSTR(1,<%=SELF%>.BDD.VALUE,"-")>0 THEN
			ALERT "輸入錯誤!!請輸入正確數字"
			<%=SELF%>.BDD.VALUE=""
			<%=SELF%>.BDD.FOCUS()
			EXIT FUNCTION
		END IF
	END IF
ELSEIF N=4 THEN
	IF TRIM(<%=SELF%>.AGES.VALUE)<>"" THEN
		IF ISNUMERIC(<%=SELF%>.AGES.VALUE)=FALSE OR INSTR(1,<%=SELF%>.AGES.VALUE,"-")>0 THEN
			ALERT "輸入錯誤!!請輸入正確數字"
			<%=SELF%>.AGES.VALUE=""
			<%=SELF%>.AGES.FOCUS()
			EXIT FUNCTION
		END IF
	END IF
ELSEIF N=5 THEN
	IF TRIM(<%=SELF%>.GTDAT.VALUE)<>"" THEN
		IF ISNUMERIC(<%=SELF%>.GTDAT.VALUE)=FALSE OR INSTR(1,<%=SELF%>.GTDAT.VALUE,"-")>0 THEN
			ALERT "輸入錯誤!!請輸入正確數字"
			<%=SELF%>.GTDAT.VALUE=""
			<%=SELF%>.GTDAT.FOCUS()
			EXIT FUNCTION
		END IF
	END IF
END IF

END FUNCTION


function go()
	'alert "ok" 
	if <%=self%>.UID.value="LSARY" then 
		if ( cdbl(<%=self%>.sum_h1.value)+cdbl(<%=self%>.sum_h2.value) )+ cdbl(<%=self%>.totjiaH.value) > 300  then 
			alert "加班時數超過(>)300小時,不可再加班!!"
			exit function
		else
			<%=self%>.action="<%=self%>.upd.asp"
			<%=self%>.submit()		
		end if	
	else
		<%=self%>.action="<%=self%>.upd.asp"
		<%=self%>.submit()
	end if	
end function 


function showWorkTime(index) 
	empidstr = <%=self%>.empid.value 	
	workdatstr = <%=self%>.workdat(index).value 
	
	open "showWorkTime.asp?empid=" & empidstr &"&workdat=" & workdatstr  , "_blank"   , "top=100, left=100, width=500, height=400, scrollbars=yes  " 
end function 

</script>



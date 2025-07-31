<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<!-- #include file = "../GetSQLServerConnection.fun" -->
<!-- #include file="../ADOINC.inc" -->
<%
if session("netuser")="" then 
	response.write "使用者帳號為空!!請重新登入!!"
	response.end 
end if 	

SELF = "YEBE0103B"

Set conn = GetSQLServerConnection()
'Set rs = Server.CreateObject("ADODB.Recordset")

nowmonth = year(date())&right("00"&month(date()),2)
if month(date())="01" then
	calcmonth = year(date()-1)&"12"
else
	calcmonth =  year(date())&right("00"&month(date())-1,2)
end if 

FUNCTION FDT(D)
	Response.Write YEAR(D)&"/"&RIGHT("00"&MONTH(D),2)&"/"&RIGHT("00"&DAY(D),2)
END FUNCTION

WHSNO = request("whsno")
wbloai = request("wbloai")
EMPID = request("EMPID")
if empid="" then empid=1

lorry=request("lorry")  

gTotalPage = 1
PageRec = 1    'number of records per page
TableRec = 25    'number of fields per record

Set conn = GetSQLServerConnection() 

pcp=request("pcp")
pcd=request("pcd")
ptp=request("ptp")
pgtp=request("pgtp")
xhid = request("xhid") 
soxe = request("soxe")
lorry=request("lorry")
xhid =request("xhid") 

sortby=request("sortby") 
if sortby="" then sortby="cv desc, a.lorry"

queryx = trim(request("queryx"))

if request("TotalPage") = "" or request("TotalPage") = "0" then
	CurrentPage=1
	
	sql="select isnull(d.status,'') xests, d.driverID, d.lorry as xelorry, d.xhid xexhid, c.incustsname, isnull(b.wbid,'') bwbid, b.personid, b.indat, case when isnull(b.yy,'')<>'' then b.yy else ''end + case when isnull(b.mm,'')<>'' then '/'+b.mm else '' end + case when isnull(b.dd,'')<>'' then '/'+b.dd else '' end as bdy , "&_
		"convert(varchar(10),b.outdat,111) nhoutdat, isnull(b.flag,'') nhflag, "&_
		"convert(varchar(10),b.indat,111) nhindat, isnull(b.job,'') nhcv, "&_
		"convert(char(10),xeindat,111) xeindate, "&_
		"convert(char(10),txindat,111) txindate, convert(char(10),txoutdat,111) txoutdate, a.* from  "&_		
		"(select * from [yfymis].dbo.ysbmlrif   ) d   "&_
		"left join (select '"& WHSNO &"' as bwhsno,  * from [yfymis].dbo.ysbdxetp   ) a on a.lorry = d.lorry and a.soxe=d.driverid    "&_
		"left join (select * from [yfynet].dbo.wbempfile where  loai='01' and  wbwhsno='"&whsno&"' ) b on b.lorry = a.lorry and b.wbid = a.wbid  "&_
		"left join (select * from [yfymis].dbo.ydbscust) c on c.incustid = d.xhid "&_		
		"where a.bwhsno<>'' and  (isnull(b.wbid,'')=''  or convert(varchar(10),isnull(a.txoutdat,''),111)<>convert(varchar(10),isnull(b.outdat,''),111) )  "&_ 
		"and ( d.driverid  like '%"& queryx &"%'  or d.lorry like '%"&queryx&"%') "&_
		"order by " & sortby

 

	'response.write sql 
	'response.end 
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open sql, conn, 1, 3
	if not rs.eof then 	  	
		PageRec = rs.RecordCount
		rs.PageSize = PageRec
	  	RecordInDB = rs.RecordCount
	  	TotalPage = rs.PageCount
	  	gTotalPage = TotalPage    
	  	ton = rs("ton")
	  	xeindat = rs("xeindate")  	  
	end if
	
	Redim tmpRec(gTotalPage, PageRec, TableRec)   'Array
	
	for i = 1 to gTotalPage
		for j = 1 to PageRec
			if not rs.EOF then
				tmpRec(i, j, 0) = "no"
				tmpRec(i, j, 1) = rs("xelorry")
				tmpRec(i, j, 2) = rs("driverID")
				tmpRec(i, j, 3) = rs("xexhid")
				tmpRec(i, j, 4) = rs("ton")
				tmpRec(i, j, 5) = rs("cv")
				tmpRec(i, j, 6) = rs("xeindat")
				tmpRec(i, j, 7) = rs("bwbid")
				tmpRec(i, j, 8) = rs("cardno")
				tmpRec(i, j, 9) = rs("txindate")
				tmpRec(i, j, 10) = rs("txoutdate")
				tmpRec(i, j, 11) = rs("txmemo")				
				tmpRec(i, j, 12) = rs("dt1")				
				tmpRec(i, j, 13) = rs("dt2")
				tmpRec(i, j, 14) = rs("flag")				
				tmpRec(i, j, 15) = "" 'rs("lorrydriverName")
				tmpRec(i, j, 16) = "" 'rs("contact")				
				tmpRec(i, j, 17) = rs("txname")				
				tmpRec(i, j, 18) = rs("aid")
				tmpRec(i, j, 19) = rs("nhoutdat")
				tmpRec(i, j, 20) = rs("nhindat")
				tmpRec(i, j, 21) = rs("nhcv")
				tmpRec(i, j, 22) = rs("incustsname")
				tmpRec(i, j, 23) = rs("xests")
				rs.MoveNext
			else
				tmpRec(i, j, 0) = "no"
				tmpRec(i, j, 7) =""
			end if
		next
 		if rs.EOF then
			rs.Close
			Set rs = nothing
			exit for
 		end if
	next
	Session("YEBE0103B") = tmpRec
else
	TotalPage = cint(request("TotalPage"))
	gTotalPage = cint(request("gTotalPage"))	
	CurrentPage = cint(request("CurrentPage"))
	'StoreToSession()
	tmpRec = Session("YEBE0103B")
	RecordInDB = request("RecordInDB")
	Select case request("send")
	     Case "FIRST"
		      CurrentPage = 1
	     Case "BACK"
		      if cint(CurrentPage) <> 1 then
			     CurrentPage = CurrentPage - 1
		      end if
	     Case "NEXT"
		      if cint(CurrentPage) <= cint(gTotalPage) then
			     CurrentPage = CurrentPage + 1
		      end if
	     Case "END"
		      CurrentPage = TotalPage
	     Case Else
		      CurrentPage = 1
	end Select
end if 

%>

<html>

<head>

<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta HTTP-EQUIV="refresh" >
<link rel="stylesheet" href="../Include/style.css" type="text/css">
<link rel="stylesheet" href="../Include/style2.css" type="text/css"> 
</head>
<body  topmargin="0" leftmargin="10"  marginwidth="0" marginheight="0"  onkeydown="enterto()" onload="f()">
<form name="<%=self%>" method="post" action="<%=self%>.accdata.asp">
<INPUT TYPE=hidden NAME=TotalPage VALUE="<%=TotalPage%>">
<INPUT TYPE=hidden NAME=CurrentPage VALUE="<%=CurrentPage%>">
<INPUT TYPE=hidden NAME=PageRec VALUE="<%=PageRec%>">
<INPUT TYPE=hidden NAME=RecordInDB VALUE="<%=RecordInDB%>">
<INPUT TYPE=hidden NAME=gTotalPage VALUE="<%=gTotalPage%>">
<INPUT TYPE=HIDDEN NAME="UID" VALUE=<%=SESSION("NETUSER")%>>
<INPUT TYPE=hidden NAME=nowmonth VALUE="<%=nowmonth%>">
<INPUT TYPE=hidden NAME=empid  VALUE="<%=empid%>">
<INPUT TYPE=hidden NAME=sortby VALUE="<%=sortby%>">
<table width="460" border="0" cellspacing="0" cellpadding="0">
	<tr><TD  >
		<img border="0" src="../image/icon.gif" align="absmiddle">
		<%=session("pgname")%> 
		</TD>
	</tr>
</table>
<hr size=0	style='border: 1px dotted #999999;' align=left width=500 > 

<TABLE WIDTH=500 CLASS=txt BORDER=0 cellspacing="2" cellpadding="2" >
	<TR height=35 >
		<TD width=70 align=right>廠別<BR><font class=txt8>Xuong</font></TD>
		<TD   valign=top>			
			<select name=WHSNO   class=txt8 onchange="whsnochg()"   >
				<option value="">請選擇廠別</option>
				<%
				if session("rights")="0" then 
					SQL="SELECT * FROM BASICCODE WHERE FUNC='WHSNO'  ORDER BY SYS_TYPE "
				else
					SQL="SELECT * FROM BASICCODE WHERE FUNC='WHSNO' and sys_type='"& session("netwhsno") &"' ORDER BY SYS_TYPE "
				end if 	
				SET RST = CONN.EXECUTE(SQL)
				WHILE NOT RST.EOF
				%>
				<option value="<%=RST("SYS_TYPE")%>" <%if RST("sys_type")=whsno then%>selected<%end if%>><%=RST("SYS_VALUE")%></option>
				<%
				RST.MOVENEXT
				WEND
				%>
			</SELECT>
			<%SET RST=NOTHING %> 	
		 
			<select name=wbloai  class=txt8    >				
				<%				
				SQL="SELECT * FROM BASICCODE WHERE FUNC='WB' and sys_type='01' ORDER BY SYS_TYPE "				
				SET RST = CONN.EXECUTE(SQL)
				WHILE NOT RST.EOF
				%>
				<option value="<%=RST("SYS_TYPE")%>" <%if RST("sys_type")=wbloai then%>selected<%end if%>><%=RST("SYS_TYPE")%><%=RST("SYS_VALUE")%></option>
				<%
				RST.MOVENEXT
				WEND
				%>
			</SELECT>
			<%SET RST=NOTHING %>
			<input type=button name=btn value="ReLoad重新載入" class=button onclick='goc()'>
		</TD>				
	</TR>
	<TR height=35 > 
		<td align=right ><a href="vbscript:getlorry()"><font color=blue>查詢<BR><font class=txt8>K.Tra</font></font></a></td>
		<td valign=top>
			<input  name="queryx" size=15 class=inputbox  value="<%=queryx%>" onblur="gos()">
			<input type="hidden" name=lorry size=3 class=inputbox  maxlength=5     value="<%=lorry%>"  >
			<input type="hidden"  name=soxe size=9 class=readonly   value="<%=soxe%>" >			
		</td>
	</tr> 
</table>
<hr size=0	style='border: 1px dotted #999999;' align=left width=500>
<Table  BORDER=0 cellspacing="1" cellpadding="1">
	<tr bgcolor=#e4e4e4 class=txt height=25>
		<Td align=center class=txt8 rowspan=2>接收<BR>Nhan</td>
		<Td align=center rowspan=2>証號<BR>so the</td>
		<Td align=center rowspan=2 width=30 >狀<br>態</td>
		<Td align=center rowspan=2 nowrap width=80 onclick="dchg(3)" style='cursor:hand'>車行<BR>Cty Xe<br><img src="../picture/soryby.gif"></td>
		<Td align=center rowspan=2  width=80 nowrap onclick="dchg(4)" style='cursor:hand'>車號<BR>so xe<br><img src="../picture/soryby.gif"></td>
		<Td width=30 align=center rowspan=2 nowrap>職務<BR>chuc vu</td>
		<Td width=65 align=center rowspan=2 nowrap>進廠日<BR>nvx</td>
		<Td width=90 align=center rowspan=2 nowrap>姓名<BR>ho ten</td>		
		<Td align=center rowspan=2 width=70 nowrap>電話1<BR>DT1</td>
		<Td align=center rowspan=2 width=70 nowrap>電話2<BR>DT1</td>
		<Td align=center rowspan=2 width=65 nowrap>離廠日<BR>NTV</td>
		<Td colspan=3 align=center>人事資料tu lieu nhan su</td>		
	</tr>
	<tr bgcolor=#e4e4e4 class=txt height=25>		
		<Td align=center width=65 nowrap>進廠日<BR>NH.NVX</td>
		<Td align=center width=65 nowrap>離廠日<BR>NH.NTV</td>
		<Td align=center width=30 nowrap>職務<BR>CV</td>
	</tr>
	<%F_wbid = right(empid,3) 
	  if F_wbid=0 then F_wbid=1
	for CurrentRow = 1 to PageRec
		if currentrow mod 2 = 0 then 
			wkcolor="#FFCCCC"
		else
			wkcolor="#FFFFCC"
		end if
		if trim(tmpRec(CurrentPage, CurrentRow,7))="" or isnull(tmpRec(CurrentPage, CurrentRow,7))then 
			wbid = "XT"&right("000"&cint(right(F_wbid,3)),3)
			F_wbid=F_wbid+1
		else
			wbid=trim(tmpRec(CurrentPage, CurrentRow, 7))
		end if 	
		'response.write CurrentRow &"="&wbid &"<BR>"
		'if 	trim(tmpRec(CurrentPage, CurrentRow, 5))<>"" then 
	%>
		<Tr bgcolor="<%=wkcolor%>" class=txt8 height=25>
			<Td align=center >
				<%if trim(tmpRec(CurrentPage, CurrentRow,7))="" or isnull(tmpRec(CurrentPage, CurrentRow,7))  then %>
					<%if tmpRec(CurrentPage, CurrentRow,5)<>"" then %>	
						<input name=func type=checkbox onclick="ins(<%=currentrow-1%>)" >					
					<%else%>
						<input name=func type=hidden>	
					<%end if%>
				<%else%>
						<input name=func type=hidden>					
				<%end if%>
				<input type=hidden name=op value="">
				<input type=hidden name=aid value="<%=trim(tmpRec(CurrentPage, CurrentRow, 18))%>">
				<input type=hidden name=xhid value="<%=trim(tmpRec(CurrentPage, CurrentRow, 3))%>">
				<input type=hidden name=Nlorry  value="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))%>">
				<input type=hidden name=nsoxe  value="<%=trim(tmpRec(CurrentPage, CurrentRow, 2))%>">
			</td>
			<Td align=center class=txt  >
				<%if trim(tmpRec(CurrentPage, CurrentRow,7))="" or isnull(tmpRec(CurrentPage, CurrentRow,7)) then %>
					<%if tmpRec(CurrentPage, CurrentRow,5)<>"" then %>	
						<input name=wbid class=inputbox size=5 value="<%=wbid%>" maxlength=5 >
					<%else%>	
						<input type=hidden name=wbid class=inputbox size=5 value="<%=wbid%>" maxlength=5 >
					<%end if%>
				<%else%>
					<Font color=blue style='cursor:hand' onclick='showdata(<%=currentrow-1%>)'><%=wbid%></font></font>
					<input type=hidden name=wbid class=inputbox size=5 value="<%=wbid%>" maxlength=5 >
				<%end if%>
			</Td>
			<Td class=txt8   align=center   ><%=trim(tmpRec(CurrentPage, CurrentRow, 23))%></td>
			<Td class=txt8  nowrap   ><%=trim(tmpRec(CurrentPage, CurrentRow, 22))%></td>
			<Td  class=txt   >
				(<%=trim(tmpRec(CurrentPage, CurrentRow, 1))%>)<%=trim(tmpRec(CurrentPage, CurrentRow, 2))%>
				
			</Td>
			<Td  align=center >
				<%=(trim(tmpRec(CurrentPage, CurrentRow, 5)))%>
				<input type=hidden name=cv value="<%=(trim(tmpRec(CurrentPage, CurrentRow, 5)))%>" >
			</Td>
			<Td align=center ><%=(trim(tmpRec(CurrentPage, CurrentRow, 9)))%>
				<input name=txindat type=hidden  value="<%=(trim(tmpRec(CurrentPage, CurrentRow, 9)))%>" onblur="date_change(<%=currentrow-1%>)">
			</Td>
			<Td ><%=(trim(tmpRec(CurrentPage, CurrentRow, 17)))%>
				<input name=name_vn  type=hidden value="<%=(trim(tmpRec(CurrentPage, CurrentRow, 17)))%>" >
			</Td>			
			<Td ><%=trim(tmpRec(CurrentPage, CurrentRow, 12))%>
				<input name=dt1 type=hidden value="<%=trim(tmpRec(CurrentPage, CurrentRow, 12))%>" >
			</Td>
			<Td  ><%=trim(tmpRec(CurrentPage, CurrentRow, 13))%>
				<input name=dt2 type=hidden value="<%=trim(tmpRec(CurrentPage, CurrentRow, 13))%>" >
			</Td>
			<Td align=center ><%=trim(tmpRec(CurrentPage, CurrentRow, 10))%>
				<input name=txoutdate  type=hidden value="<%=trim(tmpRec(CurrentPage, CurrentRow, 10))%>" >
			</Td>			
			<Td class=txt8 align=center><Font color=blue><%=trim(tmpRec(CurrentPage, CurrentRow, 20))%></font></td>
			<Td class=txt8 align=center><Font color=blue><%=trim(tmpRec(CurrentPage, CurrentRow, 19))%></font></td>
			<Td class=txt8 align=center><Font color=blue><%=trim(tmpRec(CurrentPage, CurrentRow, 21))%></font></td>
		</tr>		
	<%
	next%>
	<input type=hidden name=func>
	<input type=hidden name=op>
	<input type=hidden name=wbid>
	<input type=hidden name=cv>
	<input type=hidden name=txindat>
	<input type=hidden name=name_vn>
	<input type=hidden name=dt1>
	<input type=hidden  name=dt2>
	<input type=hidden name=aid >
</table>
<table width=500><tr><td align=center class=txt>Page: <%=CurrentPage%> / <%=totalpage%> , Count:<%=recordinDB%></td></tr></table>
<TABLE WIDTH=460>
		<tr ALIGN=center>
			<TD >
			<input type=button  name=send value="(Y)確　　認"  class=button onclick=go()>
			<input type=RESET name=send value="(N)取 　　消"  class=button  >			
			</TD>
		</TR>
</TABLE>


</form>


</body>
</html> 
<script language=vbscript>
function ins(index)
	if <%=self%>.func(index).checked=true then 
		<%=self%>.op(index).value="Y"
	else
		<%=self%>.op(index).value=""
	end if 
end function 

function goc()
	<%=self%>.totalpage.value=""
	<%=self%>.lorry.value=""
	<%=self%>.action="yebe0103B.accdata.asp"
	<%=self%>.submit()
	'open "yebe0103B.accdata.asp" , "_self"
end function 

function gob()
	open "yebe0103.asp" , "_self"
end function  

function lorrychg()
	'if <%=self%>.lorry.value="" then 
	'	<%=self%>.soxe.value=""
	'	<%=self%>.totalpage.value="" 
		'<%=self%>.action="<%=self%>.accdata.asp" 
		'<%=self%>.submit()
	'end if 	
end function 

'-----------------enter to next field 
function getlorry()
	open "getlorry.asp?TargetName="&"<%=self%>", "Back"
	parent.best.cols="50%,50%"
end function 
function enterto()
	if window.event.keyCode = 13 then window.event.keyCode =9
end function

function f()
	'<%=self%>.whsno.focus()
	'<%=self%>.indat.select()
end function 


function showdata(index)
	wbidstr = <%=self%>.wbid(index).value
	open "Yebe0104.foregnd.asp?wbid="&wbidstr , "_self" 
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

function whsnochg()	
	code1 = <%=self%>.whsno.value
	'code2 = <%=self%>.wbloai.value
	if code1<>"" then 		
		open "<%=self%>.back.asp?ftype=getwbid&code1="&code1 &"&code2="& code2  , "Back"			
		'parent.best.cols="70%,30%"
	end if 
end function  

function loaichg()
	code1 = <%=self%>.wbloai.value
	code2 = <%=self%>.whsno.value
	if code1<>"" and code2<>"" then 
		open "<%=self%>.back.asp?ftype=getwbid&code1="&code1 &"&code2="& code2  , "Back"	
		'parent.best.cols="50%,50%"
	end if
end function 

function empidchg()
	empidstr = Ucase(Trim(<%=self%>.empid.value))
	if empidstr<>"" then
		open "<%=self%>.back.asp?ftype=empidchk&code="& empidstr , "Back"
		'parent.best.cols="50%,50%"
	end if
end function

function sexchg(x)
	if <%=self%>.radio1(0).checked=true then
		<%=self%>.sexstr.value="M"
	elseif 	<%=self%>.radio1(1).checked=true then
		<%=self%>.sexstr.value="F"
	else
		<%=self%>.sexstr.value=""
	end if
end function

function marrychg(x)
	if <%=self%>.radio2(0).checked=true then
		<%=self%>.marryed.value="Y"
	elseif 	<%=self%>.radio2(1).checked=true then
		<%=self%>.marryed.value="N"
	elseif 	<%=self%>.radio2(2).checked=true then
		<%=self%>.marryed.value="L"
	else
		<%=self%>.marryed.value=""	
	end if
end function

function BACKMAIN()
	open "../main.asp" , "_self"
end function

FUNCTION GO()
	rd = <%=self%>.recordInDB.value 
	ss = 0
	for zz = 1 to rd 
		if <%=self%>.op(zz-1).value="Y" then 
			if <%=self%>.wbid(zz-1).value="" then 
				alert "必須輸入証號!! Khong co so the!!"
				<%=self%>.wbid(zz-1).focus()
				exit function 
				ss = ss+0
			else
				ss= ss+1	
			end if 
		end if 		
	next 
	if ss>=1 then 
		<%=SELF%>.ACTION="<%=self%>.upd.asp?act=EMPADDNEW"
		<%=SELF%>.SUBMIT
	end if 	
END FUNCTION

'*******檢查日期*********************************************
FUNCTION date_change(index)


INcardat = Trim(<%=self%>.txindat(index).value)

IF INcardat<>"" THEN
	ANS=validDate(INcardat)
	IF ANS <> "" THEN
		Document.<%=self%>.txindat(index).value=ANS
	ELSE
		ALERT "EZ0067:輸入日期不合法 (yyyy/mm/dd)!!"		
		Document.<%=self%>.txindat(index).value=""
		Document.<%=self%>.txindat(index).focus()
		
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
		ELSE
			<%=SELF%>.AGES.VALUE=CDBL(YEAR(DATE()))-CDBL(<%=SELF%>.BYY.VALUE) + 1
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


function dchg(a) 
	select case a 
		case 1 
			<%=self%>.sortby.value=""
		case 2 
			<%=self%>.sortby.value="driverid"
		case 3 
			<%=self%>.sortby.value="d.xhid, d.driverid "
		case 4 
			<%=self%>.sortby.value="d.driverid"
		case 5 
			<%=self%>.sortby.value="rp_func, rpno "
	end select 	
	<%=self%>.totalpage.value=""
 	<%=self%>.action="<%=self%>.accdata.asp"
 	<%=self%>.submit() 														
end function  

function gos()
	<%=self%>.totalpage.value=""
	<%=self%>.action="<%=self%>.accdata.asp"
	<%=self%>.submit()
end function  
</script>


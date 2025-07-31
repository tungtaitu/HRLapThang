<%@LANGUAGE=VBSCRIPT CODEPAGE=950%>
<!-- #include file = "../../GetSQLServerConnection.fun" -->
<!-- #include file="../../ADOINC.inc" -->
<%
Set conn = GetSQLServerConnection()
self="vyfysucoSCHTEST"

nowmonth = year(date())&right("00"&month(date()),2)
if month(date())="01" then
	calcmonth = year(date()-1)&"12"
else
	calcmonth =  year(date())&right("00"&month(date())-1,2)
end if

if day(date())<=11 then
	if month(date())="01" then
		calcmonth = year(date()-1)&"12"
	else
		calcmonth =  year(date())&right("00"&month(date())-1,2)
	end if
else
	calcmonth = nowmonth
end if

NNY=year(date())
NDY=year(date())+1

sgym = request("sgym")
groupid = Trim(request("groupid"))
sgno = request("sgno")
cfGroup = request("cfGroup")
COUNTRY = REQUEST("COUNTRY")
salaryYM= request("salaryYM")
'response.write "xxxx=" & sgym
'response.end

SQL=" select isnull(D.EXRT,1) exrt , a.autoid as Mautoid, a.sgno as Msgno,  "&_
	"CONVERT(CHAR(10),a.pddate,111) PDDATE ,  a.totCost, a.sgcost, a.sgym AS MSGYM , a.sgmemo, "&_
	"isnull(c.country,'') country, isnull(c.groupid,'') groupid  ,  isnull(c.cstr,'') cstr, isnull(c.cstr,'') cstr,  "&_
	"b.*   from  "&_
	"( select * from yfymsuco  ) a  "&_
	"left join  ( select  * from yfydsuco ) b on b.autoid = a.autoid and b.sgno = a.sgno "&_
	"left join ( select * from view_empfile ) c on c.empid = b.empid  "&_
	"LEFT JOIN (SELECT * FROM VYFYEXRT ) D ON D.CODE = isnull(B.DM,'VND')  AND D.YYYYMM =A.SGYM WHERE ISNULL(A.SGNO,'')<>'' "

IF 	sgym=""AND salaryYM="" and GROUPID="" AND SGNO="" AND CFGROUP="" AND COUNTRY="" THEN
	SQL=SQL&"AND  A.SGym='"& nowmonth &"' "
END IF
IF salaryYM<>"" THEN
	SQL=SQL&"AND b.YM='"& salaryYM &"' "
END IF
IF sgYM<>"" THEN
	SQL=SQL&"AND A.SGYM='"& SGYM &"' "
END IF
IF sgno<>"" THEN
	SQL=SQL&"AND A.sgno='"& sgno &"' "
END IF
IF GROUPID<>"" THEN
	SQL=SQL&"AND GROUPID='"& GROUPID &"' "
END IF
IF COUNTRY<>"" THEN
	SQL=SQL&"AND COUNTRY='"& COUNTRY &"' "
END IF
IF CFGROUP<>"" THEN
	SQL=SQL&"AND CFGROUP='"& CFGROUP &"' "
END IF
SQL=SQL&"ORDER BY A.SGYM, A.SGNO, B.CFGROUP, B.EMPID, B.CFDW "
'RESPONSE.WRITE SQL
'response.end
Set rs = Server.CreateObject("ADODB.Recordset")
RS.OPEN SQL, CONN, 3, 3



%>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
<link rel="stylesheet" href="../../Include/style.css" type="text/css">
<link rel="stylesheet" href="../../Include/style2.css" type="text/css">
<SCRIPT LANGUAGE=vbscript>
<!--
function f()
	<%=self%>.sgym.focus()
end function

function sch()
	<%=self%>.action="vyfysuco.schTEST.asp"
	<%=self%>.submit()
end function
//-->
</SCRIPT>
</head>
<body  topmargin="0" leftmargin="0"  marginwidth="0" marginheight="0" >
<form name="<%=self%>" method="post" action="EMPFILE.SALARY.ASP">
<input type=hidden name="NNY" value="<%=NNY%>">
<input type=hidden name="NDY" value="<%=NDY%>">
<TABLE width=750 CLASS=TXT8 BGCOLOR="#CCCCCC" ALIGN=left border="0"  cellspacing="1" cellpadding="1"    >
		<%
		i=0
		F1_TOTCOST = 0
		F1_SUKM = 0
		if NOT RS.EOF  then
		for x =1 to rs.RecordCount
			i=i+1
		%>
		<tr bgcolor="#ffffff">
			<TD WIDTH=30 align=center nowrap ><%=i%></TD>
			<TD WIDTH=60 HEIGHT=20  ALIGN=CENTER nowrap><%=RS("MSGYM")%></TD>
			<TD WIDTH=60 ALIGN=CENTER nowrap><%=RS("MSGNO")%></TD>
			<TD WIDTH=80 ALIGN=CENTER nowrap><%=RS("PDDATE")%></TD>
			<TD WIDTH=80 ALIGN=RIGHT nowrap><%=formatnumber(RS("TOTCOST"),0)%></TD>
			<TD WIDTH=50 ALIGN=CENTER nowrap><%=RS("CFGROUP")%></TD>
			<TD WIDTH=80 nowrap><%=RS("CFDW")%></TD>
			<TD WIDTH=50 ALIGN=center nowrap><%=RS("DM")%></TD>
			<TD WIDTH=80 ALIGN=RIGHT nowrap><%=formatnumber(RS("SUKM"),0)%></TD>
			<TD WIDTH=60 ALIGN=CENTER nowrap ><%=RS("YM")%></TD>
			<TD WIDTH=120 ALIGN=LEFT NOWRAP ><%=RS("SGMEMO")%></TD>
		</tr>
		<%
			F1_TOTCOST= CDBL(F1_TOTCOST)+CDBL(RS("TOTCOST"))
			F1_SUKM = CDBL(F1_SUKM) + CDBL(RS("SUKM"))*CDBL(RS("EXRT"))
			RS.MOVENEXT
			x=x+1
			next
		end if
		SET RS=NOTHING
		%>
		<TR>
			<TD COLSPAN=2 ALIGN=RIGHT  HEIGHT=20>總計(VND)</TD>
			<TD COLSPAN=3 ALIGN=RIGHT  HEIGHT=20><%=formatnumber(F1_TOTCOST,0)%></TD>
			<TD COLSPAN=4 ALIGN=RIGHT  HEIGHT=20><%=formatnumber(F1_SUKM,0)%></TD>
			<TD WIDTH=60 ALIGN=CENTER nowrap > </TD>
			<TD  WIDTH=120 ALIGN=LEFT NOWRAP> </TD>
		</TR>
	</TABLE>
</form>
</body>
</html>


<script language=vbsCRIPT>

function strchg(a)
	if a=1 then
		<%=self%>.empid1.value = Ucase(<%=self%>.empid1.value)
	elseif a=2 then
		<%=self%>.empid2.value = Ucase(<%=self%>.empid2.value)
	end if
end function

function empidchg()
	if <%=self%>.empid.value<>"" then
		empidstr=Ucase(Trim(<%=self%>.empid.value))
		open "<%=self%>.back.asp?func=A&code="& empidstr , "Back"
		'PARENT.BEST.COLS="70%,30%"
	end if
end function

function go()
	if <%=self%>.sgno.value="" then
		alert "必須輸入事故單號"
		<%=self%>.sgno.focus()
		exit function
	end if
	if <%=self%>.pddate.value="" then
		alert "請輸入判定日期!!"
		<%=self%>.pddate.focus()
		exit function
	end if
	if <%=self%>.cfGroup.value="A" then
		if <%=self%>.empid.value="" or <%=self%>.cfdw.value="" then
			alert "必須輸入工號或責任對象(員工姓名)!!"
			if <%=self%>.empid.value="" then
				<%=self%>.empid.focus()
			else
				<%=self%>.cfdw.focus()
			end if
			exit function
		end if
	elseif <%=self%>.cfGroup.value="B" then
		if <%=self%>.cfdw.value="" then
			alert "必須輸入責任對象(司機車號)!!"
			<%=self%>.cfdw.focus()
			exit function
		end if
	else
		if <%=self%>.cfdw.value="" then
			alert "必須輸入責任對象(姓名或廠商名稱)!!"
			<%=self%>.cfdw.focus()
			exit function
		end if
	end if
	if <%=self%>.sgcost.value="" then
		alert "請輸入扣款金額!!"
		<%=self%>.sgcost.focus()
		exit function
	end if

 	<%=self%>.action="vyfysuco.Upd.asp"
 	<%=self%>.submit()
end function


'*******檢查日期*********************************************
FUNCTION date_change(a)

if a=1 then
	INcardat = Trim(<%=self%>.indat1.value)
elseif a=2 then
	INcardat = Trim(<%=self%>.indat2.value)
end if

IF INcardat<>"" THEN
	ANS=validDate(INcardat)
	IF ANS <> "" THEN
		if a=1 then
			Document.<%=self%>.indat1.value=ANS
		elseif a=2 then
			Document.<%=self%>.indat2.value=ANS
		end if
	ELSE
		ALERT "EZ0067:輸入日期不合法 !!"
		if a=1 then
			Document.<%=self%>.indat1.value=""
			Document.<%=self%>.indat1.focus()
		elseif a=2 then
			Document.<%=self%>.indat2.value=""
			Document.<%=self%>.indat2.focus()
		end if
		EXIT FUNCTION
	END IF

ELSE
	'alert "EZ0015:日期該欄位必須輸入資料 !!"
	EXIT FUNCTION
END IF

END FUNCTION

function enterto()
		if window.event.keyCode = 13 then window.event.keyCode =9
		IF window.event.keyCode = 113 THEN
			GO()
		END IF
end function

function  sgcostChg()
	if <%=self%>.sgcost.value<>"" then
		INcardat = Trim(<%=self%>.pddate.value)
		if  INcardat<>"" then
			ANS=validDate(INcardat)
			if cdbl(year(ANS))< cdbl(<%=self%>.NNY.value) or  cdbl(year(ANS))> cdbl(<%=self%>.NDY.value) then
				alert "請確認判定日期是否輸入錯誤!!"
				<%=self%>.SSYM(x).value=0
				<%=self%>.SSYM(x).style.color="BLACK"
			else
				if <%=self%>.empid.value<>"" and <%=self%>.empid.value<"L0051" then
					x=( cdbl(year(ANS))*12+cdbl(month(ANS)) ) - (cdbl(<%=self%>.NNY.value)*12+1)
				else
					x=( cdbl(year(ANS))*12+cdbl(month(ANS)) ) - (cdbl(<%=self%>.NNY.value)*12)
				end if
				<%=self%>.SSYM(x).value=<%=self%>.sgcost.value
				<%=self%>.SSYM(x).style.color="RED"
			end if
		ELSE
			ALERT "EZ0067:判定日期輸入不合法 !!"
			Document.<%=self%>.pddate.value=""
			Document.<%=self%>.pddate.focus()
			EXIT FUNCTION
		END IF
	end if
end function

FUNCTION  pddatechg()
	INcardat = Trim(<%=self%>.pddate.value)
	sgnostr= <%=self%>.sgno.value
	IF INcardat<>"" THEN
		ANS=validDate(INcardat)
		IF ANS <> "" THEN
			Document.<%=self%>.pddate.value=ANS
			if right(ANS,2)<="10" then
				sgymstr=dateadd("d",-30,ANS)
			'elseIF right(ANS,2)>="26" THEN
			'	sgymstr=dateadd("d",10,ANS)
			ELSE
				sgymstr=dateadd("d",1,ANS)
			end if
			<%=self%>.sgym.value=year(sgymstr)& right("00"&month(sgymstr),2)
			'ClCM=year(ANS)& right("00"&month(ANS),2)
			'x=( cdbl(year(ANS))*12+cdbl(month(ANS)) ) - (cdbl(<%=self%>.NNY.value)*12+1)
			'<%=self%>.TOTcost.focus()
			open "<%=self%>.back.asp?func=B&code1="& sgnostr &"&code2=" & ANS  , "Back"
			'parent.best.cols="70%,30%"

		ELSE
			ALERT "EZ0067:輸入日期不合法 !!"
			Document.<%=self%>.pddate.value=""
			Document.<%=self%>.pddate.focus()
			EXIT FUNCTION
		END IF
	END IF
End FUNCTION


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

function schEmp()
	open "GetEmpData.asp", "Back"
	parent.best.cols="60%,40%"
end function
</script> 
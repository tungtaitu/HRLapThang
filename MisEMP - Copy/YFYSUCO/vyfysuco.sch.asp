<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<!-- #include file = "../GetSQLServerConnection.fun" -->
<!-- #include file="../ADOINC.inc" -->
<!--#include file="../include/sideinfo.inc"-->
<%
Set conn = GetSQLServerConnection()
self="vyfysuco"

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
salaryym = request("salaryYM")

SQL=" select D.EXRT, a.autoid as Mautoid, a.sgno as Msgno,  "&_
	"CONVERT(CHAR(10),a.pddate,111) PDDATE ,  a.totCost, a.sgcost, a.sgym AS MSGYM , a.sgmemo, "&_
	"isnull(c.country,'') country, isnull(c.groupid,'') groupid  ,  isnull(c.cstr,'') cstr, isnull(c.cstr,'') cstr,  "&_
	"b.*   from  "&_
	"( select * from yfymsuco  ) a  "&_
	"left join  ( select  * from yfydsuco ) b on b.autoid = a.autoid and b.sgno = a.sgno "&_
	"left join ( select * from view_empfile ) c on c.empid = b.empid  "&_
	"LEFT JOIN (SELECT * FROM VYFYEXRT ) D ON D.CODE = B.DM AND D.YYYYMM =A.SGYM WHERE ISNULL(A.SGNO,'')<>'' "

IF 	sgym=""AND GROUPID="" AND SGNO="" AND CFGROUP="" AND COUNTRY="" THEN
	SQL=SQL&"AND  A.SGym='"& nowmonth &"' "
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
Set rs = Server.CreateObject("ADODB.Recordset")
RS.OPEN SQL, CONN, 3, 3



%>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
<link rel="stylesheet" href="../Include/style.css" type="text/css">
<link rel="stylesheet" href="../Include/style2.css" type="text/css">
<SCRIPT LANGUAGE=vbscript>
<!--
function f()
	<%=self%>.sgym.focus()
end function

function sch()
	<%=self%>.action="vyfysuco.sch.asp?pgid=<%=request("pgid")%>"
	<%=self%>.submit()
end function

function REsch()
	<%=self%>.sgym.value=""
	<%=self%>.salaryym.value=""
	<%=self%>.sgno.value=""
	<%=self%>.action="vyfysuco.sch.asp"
	<%=self%>.submit()
end function
//-->
</SCRIPT>
</head>
<body  topmargin="5" leftmargin="5"  marginwidth="0" marginheight="0"  onkeydown="enterto()" onload='f()' >
<form name="<%=self%>" method="post" action="vyfysuco.schTEST.asp?pgid=<%=request("pgid")%>">
<input type=hidden name="NNY" value="<%=NNY%>">
<input type=hidden name="NDY" value="<%=NDY%>">

	<table border=0 style="height:70px"><tr><td>&nbsp;</td></tr></table>
	<table width="100%" border=0 >
		<tr>
			<td>
				<table width="80%" BORDER=0 align=center cellpadding=0 cellspacing=0 >
					<tr>
						<td>
							<table class="txt" cellpadding=3 cellspacing=3>
								<tr>
									<td align=right>事故年月<br>Ngày Sự Cố </td>
									<td>
										<input  type="text" style="width:100px" name=sgym  VALUE="<%=SGYM%>">
									</td>
									<td  align=right>扣款年月<br>Tháng Trừ Tiền</td>
									<td>
										<input  type="text" style="width:100px" name=salaryYM  VALUE="<%=salaryYM%>">
									</td>
									<td align=right>事故單號<br>Mã Sự Cố</td>
									<td>
										<input  type="text" style="width:100px" name="sgno"  VALUE="<%=sgno%>" >
									</td>
								</tr>
								<tr>
									<td align=right >責任單位<br>Người Chịu Trách Nhiệm</td>
									<td >
										<select name=cfGroup style="width:120px">
											<OPTION VALUE="" <%IF CFGROUP="" THEN %>SELECTED<%END IF%>>全部 Toàn Bộ</OPTION>
											<option value="A" <%IF CFGROUP="A" THEN %>SELECTED<%END IF%>>員工 Nhân Viên</option>
											<option value="B" <%IF CFGROUP="B" THEN %>SELECTED<%END IF%>>司機 Tài Xế</option>
											<option value="C" <%IF CFGROUP="C" THEN %>SELECTED<%END IF%>>廠商 Nhà Cung Ứng</option>
											<option value="D" <%IF CFGROUP="D" THEN %>SELECTED<%END IF%>>台籍幹部 Cán Bộ Hải Ngoại</option>
										</select>
									</td>
									<td align=right >國籍<br>Quốc Tịch</td>
									<td >
										<select name=country  style="width:120px">
											<OPTION VALUE="">全部 Toàn Bộ</OPTION>
											<%sql="select * from basicCode where func='country' order by sys_type"
											set rds=conn.execute(Sql)
											while not rds.eof
											%><option value="<%=rds("sys_type")%>" <%if rds("sys_type")=country then %>selected<%end if%>><%=rds("sys_type")%><%=rds("sys_value")%></option>
											<%rds.movenext
											wend
											set rds=nothing
											%>
										</select>
									</td>
									<td  align=right>單位<br>Bộ Phận</td>
									<td >
										<select name=groupid style="width:120px">
											<OPTION VALUE="">全部 Toàn Bộ</OPTION>
											<%sql="select * from basicCode where func='GroupID' order by sys_type"
											set rds=conn.execute(Sql)
											while not rds.eof
											%><option value="<%=rds("sys_type")%>" <%if rds("sys_type")=groupid then %>selected<%end if%>><%=rds("sys_value")%></option>
											<%rds.movenext
											wend
											set rds=nothing
											%>
										</select>										
									</td>
									<td nowrap>
										<input type=button name="send" value="查詢 Xác nhận" onclick="sch()" class="btn btn-sm btn-outline-secondary" >
										<input type=button name="send" value="重新查詢 Thiết lập lại" onclick="REsch()" class="btn btn-sm btn-outline-secondary" >
									</td>
								</tr>
							</table>
						</td>
					</tr>
					<tr>
						<td>
							<table width=830  CLASS=TXT8 BGCOLOR="#CCCCCC"  border="0"  cellspacing="1" cellpadding="1"    >
								<TR BGCOLOR="#FEF7CF"  >
									<TD WIDTH=30 HEIGHT=20 ALIGN=CENTER NOWRAP >STT</TD>
									<TD WIDTH=60 HEIGHT=20 ALIGN=CENTER NOWRAP>事故年月<br>Ngày Sự Cố</TD>
									<TD WIDTH=60 ALIGN=CENTER NOWRAP>事故單號<br>Mã Sự Cố</TD>
									<TD WIDTH=80 ALIGN=CENTER NOWRAP>判定日期<br>Ngày quyết định</TD>
									<TD WIDTH=80 ALIGN=CENTER NOWRAP>事故金額<br>Tiền Sự Cố</TD>
									<TD WIDTH=50 ALIGN=CENTER NOWRAP>責任單位<br>Đơn Vị Chịu Trách Nhiệm</TD>
									<TD WIDTH=80 ALIGN=CENTER NOWRAP>責任對象<br>Đối Tượng Chịu Trách Nhiệm</TD>
									<TD WIDTH=50 ALIGN=CENTER NOWRAP>幣別<br>Tiền Tệ</TD>
									<TD WIDTH=80 ALIGN=CENTER NOWRAP>事故扣款<br>Trừ Tiền Sự Cố</TD>
									<TD WIDTH=60 ALIGN=CENTER NOWRAP>扣款年月<br>Ngày Trừ Tiền</TD>
									<TD WIDTH=200 ALIGN=CENTER  NOWRAP>事故原因<br>Nguyên Nhân Sự Cố</TD>
								</TR>
							</table>
						</td>
					</tr>
					<tr>
						<td>
							<table height=500 width=860 CLASS=TXT8 BGCOLOR="#CCCCCC"  border="0"  cellspacing="0" cellpadding="0"  >
								<tr BGCOLOR="#ffffff" >
									<td valign=top>
									<iframe src="vyfysuco.schtest.asp?salaryYM=<%=salaryYM%>&sgym=<%=SGYM%>&sgno=<%=sgno%>&groupid=<%=groupid%>&country=<%=country%>&cfgroup=<%=cfgroup%>" scrolling="auto" name=ce   width="100%" height="100%" marginheight=0 marginwidth=0 frameborder=0>
									</iframe>
									</td>
								</tr>
							</table>
						</td>
					</tr>
				</table>
			</td>
		</tr>
	</table>
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

 	<%=self%>.action="vyfysuco.Upd.asp?pgid=<%=request("pgid")%>"
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
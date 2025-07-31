<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<!-- #include file = "../GetSQLServerConnection.fun" -->
<!-- #include file="../ADOINC.inc" -->
<!-- #include file="../Include/func.inc" -->
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
%>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="refresh">
<link rel="stylesheet" href="../Include/style.css" type="text/css">
<link rel="stylesheet" href="../Include/style2.css" type="text/css">
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--
function f()
	<%=self%>.sgno.focus()
end function
-->
</SCRIPT>
</head>
<body  topmargin="5" leftmargin="5"  marginwidth="0" marginheight="0"  onkeydown="enterto()" onload='f()' >
<form name="<%=self%>" method="post" >
<input type=hidden name="NNY" value="<%=NNY%>">
<input type=hidden name="NDY" value="<%=NDY%>">

	<table border=0 style="height:70px"><tr><td>&nbsp;</td></tr></table>
	
	<table width="100%" BORDER=0 align=center cellpadding=0 cellspacing=0 >
		<tr>
			<td>
				<table class="txt" cellpadding=3 cellspacing=3>
					<tr>
						<td nowrap align=right>*憑證單號<br><font class="txt8">Mã PTT</font></td>
						<td >
							<input  type="text" style="width:100px" name=sgno >
							<input type=hidden name="autoid"  size=3 readonly >
						</td>
						<td nowrap align=right>*判定日期<br><font class="txt8">Ngày xử lý</font></td>
						<td><input  type="text" style="width:100px" name="pddate"  onblur="pddatechg()" value="<%=session("pddate")%>"></td>
						<td nowrap align=right>事故年月<br><font class="txt8">Ngày sự cố</font></td>
						<td ><input  type="text" style="width:100px" name=sgym  ></td>
					</tr>
					<tr>
						<td nowrap align=right>原因說明<br><font class="txt8">Diễn giải</font></td>
						<td ><input  type="text" style="width:100px" name=sgmemo  ></td>
						<td nowrap align=right>損失金額<br><font class="txt8">Số tiền</font></td>
						<td nowrap>
							<input  type="text" style="width:120px" name=TOTcost>
							<font class="txt">VND</font>
						</td>
						<td nowrap align=right>*責任單位<br><font class="txt8">Chịu trách nhiệm</font></td>
						<td >
							<select name=cfGroup   onchange="cfchg()"  style="width:120px">
								<option value="A">員工 Nhân viên</option>
								<option value="B">司機 Tài xế</option>
								<option value="C">廠商 Nhà cung ứng</option>
								<option value="D">原紙廠商 Nhà cung ứng giấy cuộn </option>					
							</select>				
						</td>
					</tr>								
					<tr>									
						<td nowrap align=right><a href="vbscript:schEmp()"><font color=blue><u>*員工工號<br><font class="txt8">Số thẻ</font></u></font></a></td>
						<td >
							<input  type="text" style="width:100px" name=empid  onblur=empidchg()>
							<input name=whsno type=hidden >
						</td>
						<td nowrap align=right>*責任對象<br><font class="txt8">Người chịu trách nhiệm</font></td>
						<td >
							<input  type="text" style="width:100px" name=cfdw  title="請填姓名或車號">
						</td>
						<td nowrap align=right>*扣款金額<br><font class="txt8">Số tiền</font></td>
						<td nowrap>
							<input  type="text" style="width:100px" name=sgcost onblur="sgcostChg()">									
							<SELECT NAME=DM style="width:60px">
								<OPTION VALUE="VND">VND</OPTION>
								<OPTION VALUE="USD">USD</OPTION>
							</SELECT>									
						</td>
					</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td>
				<table class="txt" cellpadding=3 cellspacing=3>
					<tr><td colspan=6><b><font color=red>計扣年月</font></b></td></tr>
					<%for y = cdbl(nny) to cdbl(ndy) %>
						<%for x = 1 to 12  %>
							<%if x mod 6 = 1  then %>	<tr> <%end if%>
							<td align=center>
								<font class="txt9bgr">　<%=y&"-"&right("00"&x,2)%>　</font>
								<input type=text name="SSYM"  class=inputbox8  value="0" style="width:80px;text-align:right">
							</td>
							<%if x mod 6 = 0 then %></tr><%	end if%>
						<%next%>
					<%next%>
				</table>
			</td>
		</tr>
		<tr>
			<td align="center">
				<table class="txt" cellpadding=3 cellspacing=3>
					<tr >
						<td align=center>
							<input type=button  name=btm class="btn btn-sm btn-danger" value="確   認" onclick="go()" onkeydown="go()">
							<input type=reset  name=btm class="btn btn-sm btn-outline-secondary" value="取   消">
						</td>
					</tr>
				</table>
			</td>
		</tr>
	</table>
			
</body>
</html>


<script language=vbs>
function cfchg()
	'alert <%=self%>.cfgroup.value
	if <%=self%>.cfgroup.value="E" then
		<%=self%>.cfdw.value="公司吸收(LA)"
	elseif left(<%=self%>.cfgroup.value,2)="A0" then
		<%=self%>.cfdw.value=Trim(<%=self%>.cfgroup.value)
	else
		<%=self%>.cfdw.value=""
	end if
end function

function strchg(a)
	if a=1 then
		<%=self%>.empid1.value = Ucase(<%=self%>.empid1.value)
	elseif a=2 then
		<%=self%>.empid2.value = Ucase(<%=self%>.empid2.value)
	elseif a=3 then
		<%=self%>.SHIFT.VALUE=ucase(<%=SELF%>.SHIFT.VALUE)
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
	elseif <%=self%>.cfGroup.value="B"  then
		if <%=self%>.cfdw.value="" then
			alert "必須輸入責任對象(司機車號)!!"
			<%=self%>.cfdw.focus()
			exit function
		end if
	else
		if <%=self%>.cfdw.value="" and left(<%=self%>.CFgroup.value,2)<>"A0"  then
			alert "必須輸入責任對象(姓名或廠商名稱)!!"
			<%=self%>.cfdw.focus()
			exit function
		end if
	end if
	if <%=self%>.sgcost.value="" and left(<%=self%>.CFgroup.value,2)<>"A0" and <%=self%>.CFgroup.value<>"E"  then
		alert "請輸入扣款金額!!"
		<%=self%>.sgcost.focus()
		exit function
	end if
	if left(<%=self%>.cfgroup.value,2)="A0" then
		if <%=self%>.shift.value="" then
			alert "請輸入班別!!A班請輸入[A],B班輸入[B],正常班輸入[Z]!!"
			<%=self%>.shift.focus()
		end if
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
				if left(<%=self%>.cfgroup.value,2)<>"A0" then
					if <%=self%>.empid.value<>"" and <%=self%>.empid.value<"L0051" then
						x=( cdbl(year(ANS))*12+cdbl(month(ANS)) ) - (cdbl(<%=self%>.NNY.value)*12+1)
					else
						x=( cdbl(year(ANS))*12+cdbl(month(ANS)) ) - (cdbl(<%=self%>.NNY.value)*12+1)
					end if
					<%=self%>.SSYM(x).value=<%=self%>.sgcost.value
					<%=self%>.SSYM(x).style.color="RED"
				end if
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
			'if right(ANS,2)<="10" then
			'	sgymstr=dateadd("d",-30,ANS)
			'elseIF right(ANS,2)>="26" THEN
			'	sgymstr=dateadd("d",10,ANS)
			'ELSE
			'	sgymstr=dateadd("d",1,ANS)
			'end if
			<%=self%>.sgym.value=year(ANS)& right("00"&month(ANS),2)
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
	parent.best.cols="70%,30%"
end function
</script> 
<%@LANGUAGE=VBSCRIPT CODEPAGE=950%>
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
<meta http-equiv="Content-Type" content="text/html; charset=Pig5">
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
	<table width="100%" border=0 >
		<tr>
			<td>
				<table width="80%" BORDER=0 align=center cellpadding=0 cellspacing=0 >
					<tr>
						<td>
							<table class="table-borderless table-sm bg-white text-secondary">
								<tr>
									<td width=100 align=right>*���ҳ渹:</td>
									<td width=125><input name=sgno size=10 class=inputbox>
									<input type=hidden name="autoid"  size=3 readonly >
									</td>
									<td width=100 align=right>*�P�w���:</td>
									<td width=125>
										<input name="pddate" size=11 class=inputbox onblur="pddatechg()" value="<%=session("pddate")%>">
									</td>
								</tr>
								<tr>
									<td  align=right>�ƬG�~��:</td>
									<td ><input name=sgym size=8 class=readonly   ></td>
									<td  align=right>�l�����B:</td>
									<td >
										<input name=TOTcost size=15 class=inputbox>VND
									</td>
								</tr>
								<tr>
									<td  align=right>��]����:</td>
									<td colspan=3 >
										<input name=sgmemo  class=inputbox size=52  maxlength=100>
									</td>
								</tr>
								<tr>
									<td  align=right>*�d�����:</td>
									<td >
										<select name=cfGroup class=inputbox  onchange="cfchg()" >
											<option value="A">���u</option>
											<option value="B">�q��</option>
											<option value="C">�t��</option>
											<option value="D">��ȼt��</option>					
										</select>				
									</td>
									<td  align=right><a href="vbscript:schEmp()"><font color=blue><u>*���u�u��:</u></font></a></td>
									<td >
										<input name=empid size=10 class=inputbox onblur=empidchg()>
										<input name=whsno type=hidden >
									</td>
								</tr>
								<tr>
									<td  align=right>*�d����H:</td>
									<td >
										<input name=cfdw size=15 class=inputbox  title="�ж�m�W�Ψ���">
									</td>
									<td  align=right>*���ڪ��B:</td>
									<td >
										<input name=sgcost size=10 class=inputbox  onblur="sgcostChg()">
										<SELECT NAME=DM CLASS=TXT8>
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
							<table width=450 border=0 class=txt9 align=center>
								<tr><td colspan=6><b><font color=red>�p���~��</font></b></td></tr>
								<%for y = cdbl(nny) to cdbl(ndy) %>
									<%for x = 1 to 12  %>
										<%if x mod 6 = 1  then %>	<tr> <%end if%>
										<td align=center>
											<font class=txt9bgr>�@<%=y&"-"&right("00"&x,2)%>�@</font>
											<input type=text name="SSYM"  class=inputbox8  size=9 value="0" style="text-align:right">
										</td>
										<%if x mod 6 = 0 then %></tr><%	end if%>
									<%next%>
								<%next%>
							</table>
						</td>
					</tr>
					<tr>
						<td>
							<table width=450 align=center>
								<tr >
									<td align=center>
										<input type=button  name=btm class=button value="�T   �{" onclick="go()" onkeydown="go()">
										<input type=reset  name=btm class=button value="��   ��">
									</td>
								</tr>
							</table>
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
		<%=self%>.cfdw.value="���q�l��(LA)"
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
		alert "������J�ƬG�渹"
		<%=self%>.sgno.focus()
		exit function
	end if
	if <%=self%>.pddate.value="" then
		alert "�п�J�P�w���!!"
		<%=self%>.pddate.focus()
		exit function
	end if
	if <%=self%>.cfGroup.value="A" then
		if <%=self%>.empid.value="" or <%=self%>.cfdw.value="" then
			alert "������J�u���γd����H(���u�m�W)!!"
			if <%=self%>.empid.value="" then
				<%=self%>.empid.focus()
			else
				<%=self%>.cfdw.focus()
			end if
			exit function
		end if
	elseif <%=self%>.cfGroup.value="B"  then
		if <%=self%>.cfdw.value="" then
			alert "������J�d����H(�q������)!!"
			<%=self%>.cfdw.focus()
			exit function
		end if
	else
		if <%=self%>.cfdw.value="" and left(<%=self%>.CFgroup.value,2)<>"A0"  then
			alert "������J�d����H(�m�W�μt�ӦW��)!!"
			<%=self%>.cfdw.focus()
			exit function
		end if
	end if
	if <%=self%>.sgcost.value="" and left(<%=self%>.CFgroup.value,2)<>"A0" and <%=self%>.CFgroup.value<>"E"  then
		alert "�п�J���ڪ��B!!"
		<%=self%>.sgcost.focus()
		exit function
	end if
	if left(<%=self%>.cfgroup.value,2)="A0" then
		if <%=self%>.shift.value="" then
			alert "�п�J�Z�O!!A�Z�п�J[A],B�Z��J[B],���`�Z��J[Z]!!"
			<%=self%>.shift.focus()
		end if
	end if

 	<%=self%>.action="vyfysuco.Upd.asp"
 	<%=self%>.submit()
end function


'*******�ˬd���*********************************************
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
		ALERT "EZ0067:��J������X�k !!"
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
	'alert "EZ0015:�������쥲����J��� !!"
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
				alert "�нT�{�P�w����O�_��J���~!!"
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
			ALERT "EZ0067:�P�w�����J���X�k !!"
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
			ALERT "EZ0067:��J������X�k !!"
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
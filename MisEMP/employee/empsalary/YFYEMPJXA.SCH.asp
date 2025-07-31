<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<!-- #include file = "../../GetSQLServerConnection.fun" -->
<!-- #include file="../../ADOINC.inc" -->
<!-- #include file="../../Include/func.inc" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="refresh">
<%
self="YFYEMPJXA"

Set conn = GetSQLServerConnection()

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

JXYM=REQUEST("JXYM")
salaryYM=REQUEST("YYMM")
country=REQUEST("country")
JOBID=REQUEST("JOBID")
empid1 = REQUEST("empid1")
GROUPID = REQUEST("GROUPID")
SHIFTN=REQUEST("SHIFT")

SQLSTR =" SELECT B.SYS_VALUE AS GSTR , A.* FROM  " &_
		"(SELECT *  FROM YFYMJIXO where JXYM='"& JXYM &"' AND GROUPID='"& GROUPID &"' AND SHIFT='"& SHIFTN &"' ) A  "&_
		"LEFT JOIN (select * from basicCode where func ='groupid' ) B ON B.SYS_TYPE = A.GROUPID "&_
		"ORDER BY A.STT "
Set rDs = Server.CreateObject("ADODB.Recordset")
RDS.OPEN SQLSTR,CONN, 3, 3
'RESPONSE.WRITE SQLSTR &"<br>"
IF NOT RDS.EOF THEN
	CurrentPage = 1
	ICOUNT = RDS.RecordCount
	GSTR = RDS("GSTR")
	Redim ARRAYS(ICOUNT, 10)   'Array
	for i = 1 to ICOUNT
		ARRAYS(i,1)=rDs("JXYM")
		ARRAYS(i,2)=rDs("salaryYM")
		ARRAYS(i,3)=rDs("groupID")
		ARRAYS(i,4)=rDs("shift")
		ARRAYS(i,5)=rDs("STT")
		ARRAYS(i,6)=TRIM(rDs("DESCP"))
		ARRAYS(i,7)=rDs("HXSL")
		ARRAYS(i,8)=rDs("HESO")
		ARRAYS(i,9)=rDs("PER")
		ARRAYS(i,10)=rDs("AUTOID")
		RDS.MOVENEXT
	next
	Session("EMPJX") = ARRAYS
else
%><SCRIPT LANGUAGE=VBS>
	ALERT "無當月績效資料!!"
	OPEN "<%=SELF%>.ASP", "_self"
</SCRIPT>
<%
END IF

TotalPage = 10
PageRec = 10    'number of records per page
TableRec = 30    'number of fields per record

SQL=""
SQL=SQL&"SELECT B.COUNTRY, B.NINDAT, B.EMPNAM_CN, EMPNAM_VN, B.JOB , B.JSTR,  A.* FROM  "
SQL=SQL&"( SELECT * FROM VYFYMYJX )  A  "
SQL=SQL&"LEFT JOIN ( SELECT * FROM VIEW_EMPFILE ) B ON B.EMPID = A.EMPID  "
SQL=SQL&"where A.YYMM='"& SALARYYM &"' AND A.JXYM='"& JXYM &"' AND A.SHIFT like '%"& SHIFTN &"' AND A.GROUPID LIKE '"& GROUPID &"%' "
SQL=SQL&"AND a.zuno LIKE '"& zuno &"%'   "
'RESPONSE.WRITE SQL
'RESPONSE.END
Set rs = Server.CreateObject("ADODB.Recordset")
RS.OPEN SQL,CONN, 3, 3

IF NOT RS.EOF THEN
	pagerec = rs.RecordCount
	rs.PageSize = pagerec
	RecordInDB = rs.RecordCount
	TotalPage = rs.PageCount
	Redim tmpRec(TotalPage, PageRec, TableRec)   'Array
 	for i = 1 to TotalPage
		for j = 1 to PageRec
			if not rs.EOF then
				tmpRec(i, j, 0) = "no"
				tmpRec(i, j, 1) = trim(rs("EMPID"))
				tmpRec(i, j, 2) = trim(rs("GROUPID"))
				tmpRec(i, j, 3) = trim(rs("SHIFT"))
				tmpRec(i, j, 4) = rs("UNITJX")
				tmpRec(i, j, 5) = rs("EMPNAM_CN")
				tmpRec(i, j, 6) = rs("EMPNAM_VN")
				tmpRec(i, j, 7) = rs("NINDAT")
				tmpRec(i, j, 8) = rs("FL")
				tmpRec(i, j, 9) = rs("SUKM")
				tmpRec(i, j, 10) = rs("job")
				tmpRec(i, j, 11) = left(rs("JSTR"),4)
				tmpRec(i, j, 12)=RS("JXA")
				tmpRec(i, j, 13)=RS("JXB")
				tmpRec(i, j, 14)=RS("JXC")
				tmpRec(i, j, 15)=RS("JXD")
				tmpRec(i, j, 16)=RS("JXE")
				tmpRec(i, j, 17)=RS("FL")
				tmpRec(i, j, 18)=RS("FLM")
				tmpRec(i, j, 19)=RS("FQD")
				tmpRec(i, j, 20)=RS("SUKM")
				tmpRec(i, j, 21)=RS("TOTJXM")

				rs.MoveNext
			else
				exit for
			end if
		 next
	NEXT

	Session("YFYEMPJXM") = tmpRec

ELSE
%>	<SCRIPT LANGUAGE=VBS>
		ALERT "員工資料有誤!!"
		OPEN "<%=SELF%>.ASP", "_self"
	</SCRIPT>
<%
	RESPONSE.END
END IF
%>
<link rel="stylesheet" href="../../Include/style.css" type="text/css">
<link rel="stylesheet" href="../../Include/style2.css" type="text/css">
</head>
<body  topmargin="5" leftmargin="5"  marginwidth="0" marginheight="0"  onkeydown="enterto()"  >
<form name="<%=self%>" method="post" action="<%=self%>.upd.asp">
<INPUT TYPE=hidden NAME=TotalPage VALUE="<%=TotalPage%>">
<INPUT TYPE=hidden NAME=CurrentPage VALUE="<%=CurrentPage%>">
<INPUT TYPE=hidden NAME=PageRec VALUE="<%=PageRec%>">
<INPUT TYPE=hidden NAME=RecordInDB VALUE="<%=RecordInDB%>">
<INPUT TYPE=hidden NAME="ICOUNT" VALUE="<%=ICOUNT%>">
<table width="460" border="0" cellspacing="0" cellpadding="0">
	<tr><TD>
	<img border="0" src="../../image/icon.gif" align="absmiddle">
	績效獎金</TD></tr>
</table>
<hr size=0	style='border: 1px dotted #999999;' align=left width=500>
<TABLE WIDTH=500><TR><TD ALIGN=CENTER>
<table width=500 class=txt9>
 	<tr>
 		<td ALIGN=LEFT>績效年月</td>
 		<td><INPUT NAME="JXYM" VALUE="<%=JXYM%>" CLASS=READONLY2 READONLY SIZE=8 ></td>
 		<td ALIGN=RIGHT>計薪年月</td>
 		<td><INPUT NAME="SALARYYM" VALUE="<%=SALARYYM%>" CLASS=READONLY2 READONLY SIZE=8  ></td>
 		<td ALIGN=RIGHT>單位</td>
 		<td>
 		<INPUT NAME="GID" VALUE="<%=GROUPID%>" CLASS=READONLY2 READONLY size=5 >
 		<INPUT NAME="GSTR" VALUE="<%=GSTR%>" CLASS=READONLY2 READONLY SIZE=8 >
 		</td>
 		<td ALIGN=RIGHT>班別</td>
 		<td><INPUT NAME="shiftn" VALUE="<%=SHIFTN%>" CLASS="READONLY2" READONLY SIZE=4  ></td>
 	</tr>
 	<TR><TD COLSPAN=8 HEIGHT=5></TD></TR>
</table>
<table width=500 class=txt9 BORDER=0 cellspacing="1" cellpadding="2" BGCOLOR="#CCCCCC" >
 	<tr bgcolor="#FFFFCC"><TD HEIGHT=22 ALIGN=CENTER>項次<br>比例</TD>
 		<%FOR II=1 TO ICOUNT %>
 		<TD ALIGN=CENTER CLASS=TXT8 nowrap ><%=ARRAYS(II,5)%><%=ARRAYS(II,6)%><br><%=ARRAYS(II,9)%>%</TD>
 		<%NEXT%>
 	</tr>
 	<tr BGCOLOR="#CEE7FF"><TD HEIGHT=22 ALIGN=CENTER >實績</TD>
 		<%FOR II=1 TO ICOUNT %>
 		<TD ALIGN=CENTER><%=FORMATNUMBER(ARRAYS(II,7),2)%></TD>
 		<%NEXT%>
 	</tr>
 	<tr BGCOLOR="#FED9CF"><TD HEIGHT=22 ALIGN=CENTER>係數</TD>
 		<%FOR II=1 TO ICOUNT %>
 		<TD ALIGN=CENTER><%=ARRAYS(II,8)%></TD>
 		<%NEXT%>
 	</tr>
</table>
</TD></TR></TABLE>
<hr size=0	style='border: 1px dotted #999999;' align=left width=500>
<table width=550 border=0 ><tr><td>
	 	<TABLE CLASS=TXT8 BGCOLOR="#CCCCCC" BORDER=0 border="1" cellspacing="1" ALIGN=CENTER>
	 		<TR BGCOLOR="#FEF7CF">
	 			<TD width=50 HEIGHT=22 ALIGN=CENTER nowrap>工號</TD>
	 			<TD width=110 ALIGN=CENTER nowrap>姓名</TD>
	 			<TD width=80 ALIGN=CENTER nowrap>到職日</TD>
	 			<TD width=60  ALIGN=CENTER nowrap>職等</TD>
	 			<TD width=60 ALIGN=CENTER nowrap>單位</TD>
	 			<%FOR J=1 TO ICOUNT%>
	 			<TD ALIGN=CENTER CLASS=TXT8  WIDTH=70 nowrap><%=ARRAYS(J,5)%><br><%=ARRAYS(J,6)%></TD>
	 			<%NEXT%>
	 			<td colspan=2 width=100 align=center>忘刷遲到早退<BR>次數　　扣款金額　</td>
	 			<!--td width=30 nowrap>次數</td>
	 			<td width=70 nowrap>金額</td -->
	 			<td width=50 nowrap align=center>反規定</td>
	 			<td width=70 nowrap align=center>事故扣款</td>
	 			<td width=100 nowrap align=center>合計</td>
	 		</TR>
	 		<%for CurrentRow = 1 to PageRec
				IF CurrentRow MOD 2 = 0 THEN
					WKCOLOR="#FFFFFF"
				ELSE
					WKCOLOR="#FFFFFF"
				END IF
				if 	tmpRec(CurrentPage, CurrentRow, 1)<>"" then
			%>
		 		<TR id=id1 bgColor="<%=wkcolor%>"  >
		 			 <TD nowrap align=center style="cursor:hand" onclick="vbscript:oepnEmpWKT(<%=currentRow-1%>)" >
		 			 	<%=tmpRec(CurrentPage, CurrentRow, 1)%>
		 			 	<input type=hidden name="empid" value="<%=tmpRec(CurrentPage, CurrentRow, 1)%>">
		 			 </TD>
		 			 <TD nowrap style="cursor:hand" onclick="vbscript:oepnEmpWKT(<%=currentRow-1%>)"><%=tmpRec(CurrentPage, CurrentRow, 5)%><BR><%=tmpRec(CurrentPage, CurrentRow, 6)%></TD>
		 			 <TD nowrap align=center ><%=tmpRec(CurrentPage, CurrentRow, 7)%></TD>
		 			 <TD nowrap><%=tmpRec(CurrentPage, CurrentRow, 11)%></TD>
		 			 <TD nowrap align=right><%=formatnumber(tmpRec(CurrentPage, CurrentRow, 4),0)%></TD>
		 			 <%for xx = 1 to ICOUNT %>
		 			 <TD nowrap align=center>
		 			 	<%'response.write CurrentRow &"-"& cdbl(11)+cdbl(xx)  %>
		 			 	<input name="JX<%=arrays(xx,5)%>" class="readonly8s" size="9" value="<%=formatnumber(tmpRec(CurrentPage, CurrentRow, cdbl(11)+cdbl(xx)),0)%>"   style='text-align:right' >
		 			 	<input type=hidden size=3  name="BJX<%=arrays(xx,5)%>" value="<%=tmpRec(CurrentPage, CurrentRow, cdbl(11)+cdbl(xx))%>"  >
		 			 </TD>
		 			 <%next%>
		 			 <TD nowrap align=center>
		 			 	<input name="FL"  class="readonly8s" size=3 value="<%=tmpRec(CurrentPage, CurrentRow, 17)%>" readonly style='text-align:right'>
		 			 </TD>
		 			 <TD nowrap align=center>
		 			 	<input name="FLmoney"  class="readonly8s"  size="9" value="<%=tmpRec(CurrentPage, CurrentRow, 18)%>" readonly style='text-align:right' >
		 			 </TD>
		 			 <TD align=center>
		 			 	<input name="FQD"  class="readonly8s" size="5" value="<%=tmpRec(CurrentPage, CurrentRow, 19)%>" style='text-align:right' >
		 			 </TD>
		 			 <TD align=center>
		 			 	<input name="SSmoeny"  class="readonly8s" size="9" value="<%=formatnumber(tmpRec(CurrentPage, CurrentRow, 20),0)%>" readonly style='text-align:right'   >
		 			 </TD>
		 			 <TD align=center>
		 			 	<input name="TOTJX"  class="readonly8s" size=13 value="<%=formatnumber(tmpRec(CurrentPage, CurrentRow, 21),0)%>" style='text-align:right' >
		 			 </TD>
		 		</TR>
		 		<%else%>
		 			<input type=hidden name="empid" >
		 			<%for zz = 1 to ICOUNT %>
		 			<input type=hidden name="JX"&<%=arrays(zz,5)%>>
		 			<input type=hidden name="BJX"&<%=arrays(zz,5)%>>
		 			<%next%>
		 			<input type=hidden name="FL" >
		 			<input type=hidden name="FLmoney" >
		 			<input type=hidden name="FQD" >
		 			<input type=hidden name="SSmoeny" >
		 			<input type=hidden name="TOTJX" >
		 			<input type=hidden name="BTOTJX" >
		 		<%end if%>
	 		<%next%>
	 	</TABLE>	<br>
	 	<table width=600 class=txt9>
		<tr>
			<td align=left>
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
			</TD>
			<td  align=center>共<%=RecordInDB%>筆, 第<%=CurrentPage%>頁/共<%=TotalPage%>頁</td>
			<td align=riht>
				<input type=button  name=btm class=button value="確　　認" onclick="go()"  >
				<input type=reset  name=btm class=button value="取　　消" >
			</td>
		</tr>
	</table>
	<input type=hidden name="empid" >
	<%for yy = 1 to ICOUNT %>
		<input type=hidden name="JX"&<%=arrays(yy,5)%>>
		<input type=hidden name="BJX"&<%=arrays(yy,5)%>>
	<%next%>
	<input type=hidden name="FL" >
	<input type=hidden name="FLmoney" >
	<input type=hidden name="FQD" >
	<input type=hidden name="SSmoeny" >
	<input type=hidden name="TOTJX" >
	<input type=hidden name="BTOTJX" >
</td></tr></table>
</form>
</body>
</html>
<script language=vbscript >
function oepnEmpWKT(index)
	empidstr = <%=self%>.empid(index).value
	yymmstr = <%=self%>.JXym.value
	open "../ZZZ/getEmpWorkTime.asp?yymm="& yymmstr & "&empid=" & empidstr , "_blank", "top=10 , left=10, scrollbars=yes"
end function

function datachg(index)
	thiscols =document.activeElement.name
	Bcols="B"&document.activeElement.name

	'alert Bcols
	if isnumeric(document.all(thiscols)(index).value)=false then
		alert "請輸入數值"
		document.all(thiscols)(index).focus()
		document.all(thiscols)(index).value=document.all(Bcols)(index).value
		document.all(thiscols)(index).select()
		'exit function
	else
		NewJXM=cdbl(document.all(Bcols)(index).value)-cdbl(document.all(thiscols)(index).value)
		NewTOTJX=cdbl(<%=self%>.BTOTJX(index).value)-cdbl(<%=self%>.TOTJX(index).value)
		'alert NewJXM
		<%=self%>.TOTJX(index).value = cdbl(<%=self%>.TOTJX(index).value) - (cdbl(NewJXM)) + (NewTOTJX)
	end if	 	
end function

function fqdchg(index)
	if cdbl(<%=self%>.fqd(index).value)=1 then
		<%=self%>.TOTJX(index).value=cdbl(<%=self%>.TOTJX(index).value)/2
	elseif cdbl(<%=self%>.fqd(index).value)=2 then
		<%=self%>.TOTJX(index).value = 0
	else
		<%=self%>.TOTJX(index).value = <%=self%>.TOTJX(index).value
	end if
end function

function TT(a)
	alert <%=self%>.JXSTT(a).value
end function


FUNCTION GO()
	'<%=SELF%>.ACTION="YFYEMPJXA.UPD.ASP"
	<%=SELF%>.SUBMIT
END FUNCTION

</script> 
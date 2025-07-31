<%@LANGUAGE=VBSCRIPT CODEPAGE=950%>
<!-- #include file = "../../GetSQLServerConnection.fun" -->
<!-- #include file="../../ADOINC.inc" -->
<%

SELF = "EMPBASIC"
gTotalPage = 1
PageRec = 25    'number of records per page
TableRec = 20    'number of fields per record
Set conn = GetSQLServerConnection()
Set rs = Server.CreateObject("ADODB.Recordset")

nowmonth = year(date())&right("00"&month(date()),2)
if month(date())="01" then
	calcmonth = year(date()-1)&"12"
else
	calcmonth =  year(date())&right("00"&month(date())-1,2)
end if



if request("salarymm")="" then
	if day(date())<=11 then
		if month(date())="01" then
			calcmonth = year(date()-1)&"12"
		else
			calcmonth =  year(date())&right("00"&month(date())-1,2)
		end if
	else
		calcmonth = nowmonth
	end if
else
	calcmonth = request("salarymm")
end if

dim calcmm(6)
for x = 0 to ubound(calcmm)
	if month(date()) - x <= 0 then
		calcmm(x) = year(date())-1 & right("00" & (month(date())-x+12) ,2 )
	else
		calcmm(x) = year(date())&right("00"&month(date())-x,2)
	end if
	'Response.Write  calcmm(x) &"<BR>"
next

'Response.Write nowmonth &"<BR>"
'Response.Write calcmonth &"<BR>"
'Response.End

basicfunc = trim(request("basicfunc"))
ct = trim(request("ct"))
whsno = trim(request("whsno"))

'sql =" select * from empsalarybasic where func='emp' and sys_type like '%"& basicfunc &"%' "
sql="select  a.sys_type, a.sys_value , c.sys_value as jobdesc , b.*  from "&_
	"( SELECT * FROM BASICCODE  WHERE FUNC='EMP'  ) a "&_
	"left join  EMPSALARYBASIC b on b.func = a.sys_type   "&_
	"left join ( SELECT * FROM BASICCODE  WHERE FUNC='lev' ) c on c.sys_type = b.job "&_
	"where b.FUNC like '"& basicfunc &"%' and country  like '"& ct &"%' " &_
	"and isnull(bwhsno,'') like '"& whsno &"%' order by  b.code, LEN(b.code), b.country, B.func "
'RESPONSE.END 
'RESPONSE.WRITE SQL
if request("TotalPage") = "" or request("TotalPage") = "0" then
	CurrentPage = 1
	rs.Open SQL, conn, 3, 3
	IF NOT RS.EOF THEN
		rs.PageSize = PageRec
		RecordInDB = rs.RecordCount
		TotalPage = rs.PageCount +2
		gTotalPage = TotalPage
	END IF

	Redim tmpRec(gTotalPage, PageRec, TableRec)   'Array

	for i = 1 to TotalPage
	 for j = 1 to PageRec
		if not rs.EOF then
			tmpRec(i, j, 0) = "no"
			tmpRec(i, j, 1) = trim(rs("sys_type"))
			tmpRec(i, j, 2) = trim(rs("descp"))
			tmpRec(i, j, 3) = trim(rs("code"))
			tmpRec(i, j, 4) = rs("bonus")
			tmpRec(i, j, 5) = rs("autoid")
			tmpRec(i, j, 6) = rs("job")
			tmpRec(i, j, 7) = rs("jobdesc")
			tmpRec(i, j, 8) = rs("dm")
			tmpRec(i, j, 9) = rs("COUNTRY")
			tmpRec(i, j, 10) = rs("yymm")
			tmpRec(i, j, 11) = rs("bwhsno")
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
	Session("EMPBASIC") = tmpRec
else
	TotalPage = cint(request("TotalPage"))
	CurrentPage = cint(request("CurrentPage"))
	RecordInDB  = REQUEST("RecordInDB")
	StoreToSession()
	tmpRec = Session("EMPBASIC")

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
<meta http-equiv="Content-Type" content="text/html; charset=big5">
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
	<%=self%>.SYS_TYPE(0).focus()
	<%=self%>.SYS_TYPE(0).select()
end function
-->
</SCRIPT>
</head>
<body  topmargin="0" leftmargin="0"  marginwidth="0" marginheight="0"  onkeydown="enterto()" onload="f()">
<form name="<%=self%>" method="post" action="<%=self%>.fore.asp">
<table width="460" border="0" cellspacing="0" cellpadding="0">
	<tr><TD width=100%>
	<img border="0" src="../../image/icon.gif" align="absmiddle">
	人事薪資系統( 基本建檔 ) </TD></tr>
</table>
<hr size=0	style='border: 1px dotted #999999;' align=left width=500>
<INPUT TYPE=HIDDEN NAME="UID" VALUE=<%=SESSION("NETUSER")%>>
<INPUT TYPE=hidden NAME=TotalPage VALUE="<%=TotalPage%>">
<INPUT TYPE=hidden NAME=CurrentPage VALUE="<%=CurrentPage%>">
<INPUT TYPE=hidden NAME=PageRec VALUE="<%=PageRec%>">
<INPUT TYPE=hidden NAME=RecordInDB VALUE="<%=RecordInDB%>">
<INPUT TYPE=hidden NAME=gTotalPage VALUE="<%=gTotalPage%>">
<table width=520   class=font9 border=0 >
	<tr height=30>
		<td align=left width=70 >類別：</td>
		<td width=150>
			<select name=basicfunc class=font9  onchange="datachg()" >
				<option value="" <%if basicfunc="" then %> selected<%end if%>>全部顯示</option>
				<%sql="select * from basicCode where func='emp' order by sys_type"
				set rst=conn.execute(Sql)
				while not rst.eof
				%>
				<option value="<%=rst("sys_type")%>" <%if basicfunc=rst("sys_type") then %> selected<%end if%> ><%=rst("sys_type")%>-<%=rst("sys_value")%></option>
				<%rst.movenext%>
				<%wend
				  set rst=nothing
				%>
			</SELECT>
		</td>
		<td width=80 align=right>國籍：</td>
		<td >
			<select name=ct class=font9  onchange="datachg()" >
				<option value="" <%if basicfunc="" then %> selected<%end if%>>全部顯示</option>
				<%sql="select * from basicCode where func='country' order by sys_type"
				set rst=conn.execute(Sql)
				while not rst.eof
				%>
				<option value="<%=rst("sys_type")%>" <%if ct=rst("sys_type") then %> selected<%end if%> ><%=rst("sys_type")%>-<%=rst("sys_value")%></option>
				<%rst.movenext%>
				<%wend
				  set rst=nothing
				%>
			</SELECT>
		</td>
		<TD nowrap align=right height=30 >廠別：</TD>
		<TD >
			<select name=WHSNO  class=txt8  onchange="datachg()" >
				<% 
				if Session("RIGHTS")="0" then%>
				<option value="">ALL全部</option>
				<% 
					SQL="SELECT * FROM BASICCODE WHERE FUNC='WHSNO' ORDER BY SYS_TYPE "
				else
					SQL="SELECT * FROM BASICCODE WHERE FUNC='WHSNO' and sys_type='"& Session("NETWHSNO") &"' ORDER BY SYS_TYPE "
				end if 	
				 
				'SQL="SELECT * FROM BASICCODE WHERE FUNC='WHSNO' ORDER BY SYS_TYPE "
				SET RST = CONN.EXECUTE(SQL)
				WHILE NOT RST.EOF
				%>
				<option value="<%=RST("SYS_TYPE")%>" <%IF RST("SYS_TYPE")=whsno THEN%>SELECTED <%END IF%>><%=RST("SYS_VALUE")%></option>
				<%
				RST.MOVENEXT
				WEND
				%>
			</SELECT>
			<%SET RST=NOTHING %>
		</TD>		
	</tr>
</table>
<table width=550 class=font9   >
	<tr  bgcolor=LightGrey>
		<td HEIGHT=25 align=center>刪除</td>
		<td HEIGHT=25  align=center>類別</td>
		<td  align=center>說明</td>
		<td  align=center>廠別</td>
		<td  align=center>國籍</td>
		<td  align=center>代碼</td>
		<td  align=center>金額</td>
		<td  align=center>幣別</td>
		<td  align=center>有效年月</td>
		<td  align=center >職碼</td>
		<td  align=center >職等</td>
	</tr>
		<%
		for CurrentRow = 1 to PageRec

		j = 1
		if j=1 then
			wk_color = ""
			j = 0
		else
			wk_color = "#E4E4E4"
			j = 1
		end if
		%>
	<TR bgcolor="<%=wk_color%>">
		<td width="25" align=center valign="top">
			<INPUT TYPE=HIDDEN NAME=AUTOID VALUE="<%=tmpRec(CurrentPage, CurrentRow, 5)%>">
			<%if tmpRec(CurrentPage, CurrentRow, 0) = "del" then%>
				<input type=checkbox name=func value=del onclick="del(<%=CurrentRow - 1%>)" checked>
				<input type=hidden name=op value=del>
			<%else%>
				<input type=checkbox name=func value=no onclick="del(<%=CurrentRow - 1%>)" <%=mode%>>
				<input type=hidden name=op value=no>
			<%end if%>
		</td>
		
		<td>
			<% if trim(tmpRec(CurrentPage, CurrentRow, 1))<>"" then %>
				<input size=3 MAXLENGTH=5 class=readonly readonly  name=SYS_TYPE value="<%=Ucase(trim(tmpRec(CurrentPage, CurrentRow, 1)))%>" >
			<%else%>
				<input size=3 MAXLENGTH=5  class=inputbox name=SYS_TYPE value="<%=Ucase(trim(tmpRec(CurrentPage, CurrentRow, 1)))%>"  onchange="dchg(<%=currentrow-1%>)">
			<%end if%>
		</td>
		<td>
			<input size=10 class=inputbox name=SYS_VALUE value="<%=Ucase(trim(tmpRec(CurrentPage, CurrentRow, 2)))%>"  onchange="dchg(<%=currentrow-1%>)" >
		</td>
		<td>
			<SELECT NAME=bwhsno CLASS=txt8 onchange="dchg(<%=currentrow-1%>)" style='width:60'>
				<option value="">---</option>
				<%SQL="SELECT* FROM BASICCODE WHERE FUNC='whsno' ORDER BY SYS_TYPE    "
				  SET RDS=CONN.EXECUTE(SQL)
				  WHILE NOT RDS.EOF
				%>
				<OPTION VALUE="<%=RDS("SYS_TYPE")%>" <%IF RDS("SYS_TYPE")=trim(tmpRec(CurrentPage, CurrentRow, 11)) THEN%>SELECTED<%END IF%>><%=RDS("SYS_TYPE")%><%=RDS("SYS_VALUE")%></OPTION>
				<%RDS.MOVENEXT
				WEND
				SET RDS=NOTHING %>
			</SELECT>
		</td>		
		<TD> 
		<SELECT NAME=COUNTRY CLASS=INPUTBOX onchange="dchg(<%=currentrow-1%>)" >
			<%SQL="SELECT* FROM BASICCODE WHERE FUNC='COUNTRY' ORDER BY SYS_TYPE DESC  "
			  SET RDS=CONN.EXECUTE(SQL)
			  WHILE NOT RDS.EOF
			%>
			<OPTION VALUE="<%=RDS("SYS_TYPE")%>" <%IF RDS("SYS_TYPE")=trim(tmpRec(CurrentPage, CurrentRow, 9)) THEN%>SELECTED<%END IF%>><%=RDS("SYS_TYPE")%><%=RDS("SYS_VALUE")%></OPTION>
			<%RDS.MOVENEXT
			WEND
			SET RDS=NOTHING %>
		</SELECT>
		</TD>
		<td>
			<input size=5 MAXLENGTH=5 class=inputbox name=CODE value="<%=Ucase(trim(tmpRec(CurrentPage, CurrentRow, 3)))%>"  onchange="dchg(<%=currentrow-1%>)" >
		</td>
		<td>
			<input size=9 class=inputbox name=BONUS value="<%=Ucase(trim(tmpRec(CurrentPage, CurrentRow, 4)))%>" style="font-align:right"  onchange="dchg(<%=currentrow-1%>)" >
		</td>
		<td>
			<input size=9 class=inputbox name=DM value="<%=Ucase(trim(tmpRec(CurrentPage, CurrentRow, 8)))%>" style="font-align:right"  onchange="dchg(<%=currentrow-1%>)" >
		</td>
		<td>
			<input size=6 class=inputbox name=yymm value="<%=Ucase(trim(tmpRec(CurrentPage, CurrentRow, 10)))%>"  onchange="dchg(<%=currentrow-1%>)" >
		</td>
		<td>
			<input size=5 class=inputbox name=job value="<%=Ucase(trim(tmpRec(CurrentPage, CurrentRow, 6)))%>"  onchange="dchg(<%=currentrow-1%>)" >
		</td>
		<td>
			<input size=20 class=inputbox name=jobdesc value="<%=Ucase(trim(tmpRec(CurrentPage, CurrentRow, 7)))%>"  onchange="dchg(<%=currentrow-1%>)" >
		</td>
	</tr>
	<%next%>
	</table>
	<table width=500 class=font9><tr><td align=center>page : <%=currentpage%>/<%=totalpage%> count : <%=recordinDB%> </td></tr></table>

<TABLE border=0 width=500  >
	<tr WIDTH=>
    	<td align="CENTER"  >
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
		<td align=center>
			<input type="button" name="send" value="確　認"  class=button onclick=go()>
			<input type="button" name="send" value="取　消"  class=button onclick=clr()>
		</td>
	</tr>
</TABLE>
</form>
</body>
</html>
<%
Sub StoreToSession()
	Dim CurrentRow
	tmpRec = Session("EMPBASIC")
	for CurrentRow = 1 to PageRec
		tmpRec(CurrentPage, CurrentRow, 0) = request("op")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 1) = request("SYS_TYPE")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 2) = request("SYS_VALUE")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 3) = request("CODE")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 4) = request("BONUS")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 5) = request("AUTOID")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 6) = request("job")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 7) = request("jobdesc")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 8) = request("dm")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 9) = request("COUNTRY")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 11) = request("bwhsno")(CurrentRow)
	next
	Session("EMPBASIC") = tmpRec
End Sub
%>

<script language=vbs>
function BACKMAIN()

	open "../main.asp" , "_self"
end function

function datachg()
	<%=self%>.action = "empbasic.fore.asp?totalpage=0"
	<%=self%>.submit()
end function

function go()
	<%=self%>.action="<%=self%>.upd.asp"
	<%=self%>.submit()
end function

function clr()
	open "<%=self%>.asp" , "_parent"
end function

function del(index)
	if <%=self%>.func(index).checked=true then
		<%=self%>.op(index).value="del"
		open "empbasic.back.asp?func=del&index="& index &"&CurrentPage="& <%=CurrentPage%> , "Back"
		'parent.best.cols="70%,30%"
	else
		<%=self%>.op(index).value="no"
		open "empbasic.back.asp?func=no&index="& index &"&CurrentPage="& <%=CurrentPage%> , "Back"
		'parent.best.cols="70%,30%"
	end if
end function


FUNCTION dchg(index)
	<%=self%>.op(index).value="upd"
	code1 = <%=self%>.SYS_TYPE(index).value
	<%=self%>.SYS_TYPE(index).value =ucase(<%=self%>.SYS_TYPE(index).value )
	code2 = <%=self%>.SYS_VALUE(index).value
	<%=self%>.SYS_VALUE(index).value =ucase(<%=self%>.SYS_VALUE(index).value )
	code3 = <%=self%>.CODE(index).value
	<%=self%>.CODE(index).value =ucase(<%=self%>.CODE(index).value )
	code4 = <%=self%>.BONUS(index).value
	if code4<>"" then
		if isnumeric(code4)=false then
			alert "必須輸入數值!!"
			<%=self%>.BONUS(index).value=""
			<%=self%>.BONUS(index).focus()
			exit function
		end if
	end if
	code5 = <%=self%>.AUTOID(index).value
	code6 = <%=self%>.job(index).value
	code7 = <%=self%>.jobdesc(index).value
	code8 = <%=self%>.dm(index).value
	code9 = <%=self%>.COUNTRY(index).value	
	<%=self%>.DM(index).value =ucase(<%=self%>.DM(index).value )
	code10 = <%=self%>.yymm(index).value
	code11 = <%=self%>.bwhsno(index).value
	open "<%=self%>.back.asp?FUNC=upd&index="& index &"&CurrentPage="& <%=CurrentPage%> & _
		 "&code1="& CODE1 &_
		 "&CODE2="& CODE2 &_
		 "&CODE3="& CODE3 &_
		 "&CODE4="& CODE4 &_
		 "&CODE5="& CODE5 &_
		 "&CODE6="& CODE6 &_
		 "&CODE7="& CODE7 &_
		 "&CODE8="& CODE8 &_
		 "&CODE9="& CODE9 &_
		 "&CODE10="& CODE10 &_
		 "&CODE11="& CODE11 , "Back"

	'parent.best.cols="70%,30%"
end function

</script>


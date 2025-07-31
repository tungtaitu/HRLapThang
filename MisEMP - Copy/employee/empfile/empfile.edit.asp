<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<!-- #include file = "../../GetSQLServerConnection.fun" -->
<!-- #include file="../../ADOINC.inc" -->
<%
'on error resume next
session.codepage="65001"
SELF = "empfileedit"

Set conn = GetSQLServerConnection()
Set rs = Server.CreateObject("ADODB.Recordset")

whsno = trim(request("whsno"))
unitno = trim(request("unitno"))
groupid = trim(request("groupid"))
COUNTRY = trim(request("COUNTRY"))
job = trim(request("job"))
country = request("country")
QUERYX = trim(request("empid1"))
outemp = request("outemp")
EMPID = REQUEST("EMPID")
shift = request("shift")
IOemp = request("IOemp") 
gTotalPage = 1
PageRec = 20    'number of records per page
TableRec = 30    'number of fields per record
'NOWMONTH=CSTR(YEAR(DATE()))&"/"&RIGHT("00"&CSTR(MONTH(DATE())),2)&"/01"
NOWMONTH=CSTR(YEAR(DATE()))&"/"&RIGHT("00"&CSTR(MONTH(DATE())),2)&"/"&RIGHT("00"&CSTR(day(DATE())),2)


sqlstr = "select * from view_empfile where( empid<>'PELIN' and isnull(status,'')<>'D' )   AND whsno like '"& whsno &"%' and unitno like '"& unitno &"%'  and groupid like '"& groupid &"%'  "&_
	"and country like '"& country &"%'  and zuno like '"& zuno &"%' and isnull(shift,'') like '"& shift &"%' and   ( EMPID like '%"& QUERYX &"%'  or empnam_VN like '"& QUERYX &"%'  or empnam_CN like '"& QUERYX &"%')  "
	if  outemp="Y" then
		sqlstr = sqlstr & " AND isnull(bhdat,'')<>'' "
	elseif 	outemp="N"  then
		sqlstr = sqlstr & " AND isnull(bhdat,'')=''  "
	end if 	
	if EMPID<>"" THEN
		sqlstr = sqlstr & " and EMPID like '"& EMPID &"%'  "
	end if 		
	if IOemp="Y" then 
		sqlstr = sqlstr & " AND ( ISNULL(OUTDATE,'')='' OR ISNULL(OUTDATE,'')>='"& NOWMONTH &"' )  "
	elseif IOemp="N" then 
		sqlstr = sqlstr & " AND ( ISNULL(OUTDATE,'')<>'' )  "	
	end if 
sqlstr = sqlstr & "order by empid  " 
'response.write sqlstr

if request("TotalPage") = "" or request("TotalPage") = "0" then
	CurrentPage = 1
	rs.Open sqlstr, conn, 3, 3
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
			for k=1 to TableRec-1
				tmpRec(i, j, 0) = "no"
				tmpRec(i, j, 1) = trim(rs("empid"))
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
				tmpRec(i, j, 16)=RS("cstr")
				tmpRec(i, j, 17)=RS("autoid")
				IF RS("zuno")="XX" THEN
					tmpRec(i, j, 18)=""
				ELSE
					tmpRec(i, j, 18)=RS("zuno")
				END IF
				tmpRec(i, j, 19)=RS("bhdat")
				tmpRec(i, j, 20)=RS("outdate")
				tmpRec(i, j, 21)=RS("SHIFT")
				tmpRec(i, j, 22)=RS("GTDAT")
			next
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
	Session("empfileedit") = tmpRec
else
	TotalPage = cint(request("TotalPage"))
	'StoreToSession()
	CurrentPage = cint(request("CurrentPage"))
	RecordInDB  = REQUEST("RecordInDB")
	tmpRec = Session("empfileedit")

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
			  else
			  	 CurrentPage = TotalPage
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
	<%=self%>.empid1.select()
end function

function datachg()
	<%=self%>.action="empfile.edit.asp?totalpage=0"
	<%=self%>.submit
end function

-->
</SCRIPT>
</head>
<body  topmargin="0" leftmargin="0"  marginwidth="0" marginheight="0"  onkeydown="enterto()" onload="f()">
<form name="<%=self%>" method="post" action="empfile.edit.asp">
<INPUT TYPE=HIDDEN NAME="UID" VALUE=<%=SESSION("NETUSER")%>>
<INPUT TYPE=hidden NAME=TotalPage VALUE="<%=TotalPage%>">
<INPUT TYPE=hidden NAME=CurrentPage VALUE="<%=CurrentPage%>">
<INPUT TYPE=hidden NAME=RecordInDB VALUE="<%=RecordInDB%>">
<INPUT TYPE=hidden NAME=PageRec VALUE="<%=PageRec%>">
<INPUT TYPE=hidden NAME=gTotalPage VALUE="<%=gTotalPage%>">
<table width="460" border="0" cellspacing="0" cellpadding="0">
	<tr><TD >
	<img border="0" src="../../image/icon.gif" align="absmiddle">
	員工資料維護( 員工基本資料-維護 )
	</TD>
	</tr>
</table>
<hr size=0	style='border: 1px dotted #999999;' align=left width=500>
<TABLE WIDTH=460 CLASS=FONT9 BORDER=0>
	<TR height=25 >
		<TD nowrap align=right >國籍</TD>
		<TD >
			<select name=COUNTRY  class=font9  onchange="datachg()" >
				<option value=""></option>
				<%SQL="SELECT * FROM BASICCODE WHERE FUNC='COUNTRY' ORDER BY SYS_TYPE "
				SET RST = CONN.EXECUTE(SQL)
				WHILE NOT RST.EOF
				%>
				<option value="<%=RST("SYS_TYPE")%>" <%IF RST("SYS_TYPE")=country THEN %> SELECTED <%END IF%>><%=RST("SYS_TYPE")%><%=RST("SYS_VALUE")%></option>
				<%
				RST.MOVENEXT
				WEND
				%>
			</SELECT>
			<%SET RST=NOTHING %>			
		</TD>
		<TD nowrap align=right>廠別</TD>
		<TD >
			<select name=WHSNO  class=font9 onchange="datachg()">
				<option value=""></option>
				<%SQL="SELECT * FROM BASICCODE WHERE FUNC='WHSNO' ORDER BY SYS_TYPE "
				SET RST = CONN.EXECUTE(SQL)
				WHILE NOT RST.EOF
				%>
				<option value="<%=RST("SYS_TYPE")%>" <%IF  RST("SYS_TYPE")=whsno THEN %> SELECTED <%END IF%> ><%=RST("SYS_VALUE")%></option>
				<%
				RST.MOVENEXT
				WEND
				%>
			</SELECT>
			<%SET RST=NOTHING %>
		</TD>
		<!--TD   nowrap align=right>處/所</TD>
		<TD >
			<select name=unitno  class=font9 onchange="datachg()">
				<option value=""></option>
				<%SQL="SELECT * FROM BASICCODE WHERE FUNC='unit' and sys_type<>'AAA' ORDER BY SYS_TYPE "
				SET RST = CONN.EXECUTE(SQL)
				WHILE NOT RST.EOF
				%>
				<option value="<%=RST("SYS_TYPE")%>" <%IF  RST("SYS_TYPE")=unitno THEN %> SELECTED <%END IF%>><%=RST("SYS_VALUE")%></option>
				<%
				RST.MOVENEXT
				WEND
				%>
			</SELECT>
			<%SET RST=NOTHING %>
		</TD-->
		<TD nowrap align=right >組/部門</TD>
		<TD >
			<select name=GROUPID  class=font9  onchange="datachg()">
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
			<%SET RST=NOTHING %>
		</TD>
		
		<TD nowrap align=right >班別</TD>
		<TD >
			<select name=shift  class=font9  onchange="datachg()" >
				<option value="" <%if shift="" then %> selected<%end if%>></option>
				<option value="ALL" <%if shift="ALL" then %> selected<%end if%>>常日班</option>
				<option value="A" <%if shift="A" then %> selected<%end if%>>A班</option>
				<option value="B" <%if shift="B" then %> selected<%end if%>>B班</option>
			</SELECT>					
		</TD>

	</TR>
	<TR>
		<TD nowrap align=right>簽約統計</TD>
		<TD >
			<select name=outemp class=font9 onchange="datachg()"> 
			 	<option value="" <%if outemp="" then %>selected<%end if%>>全部</option>
			 	<option value="Y" <%if outemp="Y" then %>selected<%end if%>>已簽約</option>
			 	<option value="N" <%if outemp="N" then %>selected<%end if%>>未簽約</option>
			 </select>	
		</TD>
		<TD nowrap align=right>員工統計</TD>
		<TD >
			<select name=IOemp class=font9 onchange="datachg()" > 
				<option value="Y" <%if IOemp="Y" then %>selected<%end if%>>在職</option>
			 	<option value="" <%if IOemp="" then %>selected<%end if%>>全部</option>
			 	<option value="N" <%if IOemp="N" then %>selected<%end if%>>已離職</option>
			 </select>	
		</TD> 		
		<TD nowrap align=right >員工編號</TD>
		<TD COLSPAN=5>
			<INPUT NAME=empid1 SIZE=18 CLASS=INPUTBOX value="<%=QUERYX%>">
			<INPUT TYPE=BUTTON NAME=BTN VALUE="查詢" CLASS=BUTTON onclick="datachg()" ONKEYDOWN="DATACHG()">
		</TD>
	</TR>
</TABLE>

<hr size=0	style='border: 1px dotted #999999;' align=left width=500>
<!-------------------------------------------------------------------->
<TABLE WIDTH=760  CLASS="TXT9VN" BORDER=0 >
 	<TR BGCOLOR="LightGrey" HEIGHT=25   >
 		<TD width=55 nowrap align=center>工號</TD>
 		<TD width=190 nowrap align=center>姓名</TD>
 		<TD width=80 nowrap align=center>到職日期</TD>
 		<TD width=80 nowrap align=center>簽合同日</TD>
 		<TD width=60 nowrap align=center>加入工團</TD>
 		<TD width=80 nowrap align=center>離職日期</TD>
 		<TD width=60 nowrap align=center>職等</TD>
 		<TD width=50 nowrap align=center>班別</TD>
 		<TD width=75 nowrap align=center>單位部門</TD>
 		<TD width=50 nowrap align=center>廠別</TD>
 		<TD width=40 nowrap align=center>國籍</TD>
 	</TR>
	<%for CurrentRow = 1 to PageRec
		IF CurrentRow MOD 2 = 0 THEN
			WKCOLOR="LavenderBlush"
		ELSE
			WKCOLOR=""
		END IF
		'if tmpRec(CurrentPage, CurrentRow, 1) <> "" then
	%>
	<TR BGCOLOR=<%=WKCOLOR%> >
 		<TD align=center><a href='vbscript:oktest(<%=tmpRec(CurrentPage, CurrentRow, 17)%>)'><%=tmpRec(CurrentPage, CurrentRow, 1)%></a></TD>
 		<TD><a href='vbscript:oktest(<%=tmpRec(CurrentPage, CurrentRow, 17)%>)'><%=tmpRec(CurrentPage, CurrentRow, 2)%>&nbsp;<font class=txt8VN><%=tmpRec(CurrentPage, CurrentRow, 3)%></font></a></TD>
 		<TD align=center><%=tmpRec(CurrentPage, CurrentRow, 5)%></TD>
 		<TD align=center><!--簽約日-->
 			<%=tmpRec(CurrentPage, CurrentRow, 19)%>
 		</TD>
 		<TD align=center><!--加入工團-->
 			<%=tmpRec(CurrentPage, CurrentRow, 22)%>
 		</TD>
 		<TD align=center><!--離職日-->
 			<%=tmpRec(CurrentPage, CurrentRow, 20)%>
 		</TD>
 		<TD align=LEFT><!--職等-->
 			<%=left(tmpRec(CurrentPage, CurrentRow, 15),5)%>
 		</TD>
 		<TD align=LEFT><!--班別-->
 			<%=tmpRec(CurrentPage, CurrentRow, 21)%>
 		</TD>
 		<TD align=LEFT><!--部門-->
 			<%=tmpRec(CurrentPage, CurrentRow, 13)%>-<%=tmpRec(CurrentPage, CurrentRow, 14)%>
 		</TD>
 		<TD align=center> <!--廠別-->
 			<%=tmpRec(CurrentPage, CurrentRow, 11)%>
 		</TD>
 		<TD align=center><!--國籍-->
			<%=tmpRec(CurrentPage, CurrentRow, 16)%>
 		</TD>
	</TR>
	<%next%>
</TABLE>
<TABLE border=0 width=500 class=font9 >
<tr>
    <td align="CENTER" height=40 width=80%>
    PAGE:<%=CURRENTPAGE%> / <%=TOTALPAGE%> , COUNT:<%=RECORDINDB%><BR>
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
	</td>
	<td>	<BR>
		<input type="button" name="send" value="回主畫面"   class=button onclick="history.back()">
	</td>
	</TR>
</TABLE>
</form>




</body>
</html>

<script language=vbscript>
function BACKMAIN()

	open "empfile.fore1.asp" , "_self"
end function

function oktest(N)
	tp=<%=self%>.totalpage.value
	cp=<%=self%>.CurrentPage.value
	rc=<%=self%>.RecordInDB.value
	open "empfile.foregnd.asp?empautoid="& N &"&totalpage=" & tp &"&currentpage=" & cp &"&RecordInDB=" & rc , "_balnk"  , "top=10, left=10, width=600, scrollbars=yes" 
end function

</script>


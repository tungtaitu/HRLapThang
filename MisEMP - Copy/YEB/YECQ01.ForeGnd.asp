<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<!-- #include file = "../GetSQLServerConnection.fun" -->
<!-- #include file="../ADOINC.inc" -->
<%
'on error resume next
session.codepage="65001"
SELF = "YECQ01"

Set conn = GetSQLServerConnection()
Set rs = Server.CreateObject("ADODB.Recordset")

ym1 = request("yymm")
ym2 = request("yymm1")
whsno = trim(request("whsno"))
groupid = trim(request("groupid"))
COUNTRY = trim(request("COUNTRY"))
EMPID = REQUEST("EMPID")

gTotalPage = 1
PageRec = 20    'number of records per page
TableRec = 60    'number of fields per record
'NOWMONTH=CSTR(YEAR(DATE()))&"/"&RIGHT("00"&CSTR(MONTH(DATE())),2)&"/01"
NOWMONTH=CSTR(YEAR(DATE()))&"/"&RIGHT("00"&CSTR(MONTH(DATE())),2)&"/"&RIGHT("00"&CSTR(day(DATE())),2)


sql="select isnull(e.lj,'') lj , isnull(e.ljstr,'') ljstr ,  "&_
	"isnull(d.lw,'') lw , isnull(d.lg,'') lg , isnull(d.lz,'') lz , isnull(d.ls,'') ls , "&_
	"isnull(d.lwstr,'') lwstr , isnull(d.lgstr,'') lgstr , isnull(d.lzstr,'') lzstr, isnull(d.lsstr,'') lsstr, "&_
	"c.empidnam_cn, c.empidname_vn, c.nindt, c.outdate, a.* , isnull(b.real_total,0) as backTotal , f.sys_value as Sjstr  "&_
	"from "&_
	"( select * from empdsalary where  yymm  between '"& ym1 &"' and '"&ym2&"' and "&_
	"country like '"&COUNTRY&"%' and whsno like '"&whsno&"%' and groupid like '"&groupid&"%' and empid like '%"&empid&"' ) a "&_
	"left join ( select * from empdsalary_bak   )  b on b.yymm = a.yymm and b.empid = a.empid  "&_
	"join (select * from view_empfile ) c on c.empid = a.empid "&_
	"left join (select *from view_empgroup  ) d on d.empid = a.empid  and d.yymm = a.yymm   "&_
	"left join (select *from view_empjob) e on e.empid = a.empid and e.yymm = a.yymm "&_
	"left join (select *from basicCode) f on f.sys_type = a.job "

sql = sql & "order by empid  " 
response.write sql

if request("TotalPage") = "" or request("TotalPage") = "0" then
	CurrentPage = 1
	rs.Open sql, conn, 3, 3
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
				tmpRec(i, j, 0) = "no"
				tmpRec(i, j, 1) = trim(rs("empid"))
				tmpRec(i, j, 2) = trim(rs("empnam_cn"))
				tmpRec(i, j, 3) = trim(rs("empnam_vn"))
				tmpRec(i, j, 4) = rs("country")
				tmpRec(i, j, 5) = rs("nindat")
				tmpRec(i, j, 6) = rs("lj")
				tmpRec(i, j, 7) = rs("lw")
				tmpRec(i, j, 8) = rs("ls")
				tmpRec(i, j, 9)	=RS("lg")
				tmpRec(i, j, 10)=RS("lz")
				tmpRec(i, j, 11)=RS("lwstr")				
				tmpRec(i, j, 12)=RS("lgstr")
				tmpRec(i, j, 13)=RS("lzstr")
				tmpRec(i, j, 14)=RS("ljstr")
				tmpRec(i, j, 15)=RS("job")
				tmpRec(i, j, 16)=RS("whsno")
				tmpRec(i, j, 17)=RS("Sjstr") 				
				tmpRec(i, j, 18)=RS("BB")
				tmpRec(i, j, 19)=RS("CV")
				tmpRec(i, j, 20)=RS("PHU")
				tmpRec(i, j, 21)=RS("NN")
				tmpRec(i, j, 22)=RS("KT")
				tmpRec(i, j, 23)=RS("MT")
				tmpRec(i, j, 24)=RS("TTKH")
				tmpRec(i, j, 25)=RS("QC")
				tmpRec(i, j, 26)=RS("TNKH")
				tmpRec(i, j, 27)=RS("TBTR")
				tmpRec(i, j, 28)=RS("JX")
				tmpRec(i, j, 29)=RS("H1M")
				tmpRec(i, j, 30)=RS("H2M")
				tmpRec(i, j, 31)=RS("H3M")
				tmpRec(i, j, 32)=RS("B3M")
				tmpRec(i, j, 33)=RS("H1")
				tmpRec(i, j, 34)=RS("H2")
				tmpRec(i, j, 35)=RS("H3")
				tmpRec(i, j, 36)=RS("B3")
				tmpRec(i, j, 37)=RS("kzhour")
				tmpRec(i, j, 38)=RS("jiaA")
				tmpRec(i, j, 39)=RS("jiaB")
				tmpRec(i, j, 40)=RS("KZM")
				tmpRec(i, j, 41)=RS("jiaAM")
				tmpRec(i, j, 42)=RS("jiaBM")
				tmpRec(i, j, 43)=RS("BZKM")
				tmpRec(i, j, 44)=RS("KTAXM")
				tmpRec(i, j, 45)=RS("QITA")
				tmpRec(i, j, 46)=RS("FL")
				tmpRec(i, j, 47)=RS("real_total")
				tmpRec(i, j, 48)=RS("laonh")
				tmpRec(i, j, 49)=RS("sole")
				tmpRec(i, j, 50)=RS("dm")
				tmpRec(i, j, 51)=RS("lzbzj")
				tmpRec(i, j, 52)=RS("yymm")
				tmpRec(i, j, 53)=RS("outdate")
				
				
				
				
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
<link rel="stylesheet" href="../Include/style.css" type="text/css">
<link rel="stylesheet" href="../Include/style2.css" type="text/css">
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
	<%=self%>.action="<%=self%>.foregnd.asp?totalpage=0"
	<%=self%>.submit
end function

-->
</SCRIPT> 
</head>
<body  topmargin="0" leftmargin="0"  marginwidth="0" marginheight="0"  onkeydown="enterto()" onload="f()">
<form name="<%=self%>" method="post" action="<%=self%>.foregnd.asp">
<INPUT TYPE=HIDDEN NAME="UID" VALUE=<%=SESSION("NETUSER")%>>
<INPUT TYPE=hidden NAME=TotalPage VALUE="<%=TotalPage%>">
<INPUT TYPE=hidden NAME=CurrentPage VALUE="<%=CurrentPage%>">
<INPUT TYPE=hidden NAME=RecordInDB VALUE="<%=RecordInDB%>">
<INPUT TYPE=hidden NAME=PageRec VALUE="<%=PageRec%>">
<INPUT TYPE=hidden NAME=gTotalPage VALUE="<%=gTotalPage%>">
<table width="460" border="0" cellspacing="0" cellpadding="0">
	<tr><TD >
	<img border="0" src="../image/icon.gif" align="absmiddle">
	<%=session("pgname")%> 
	</TD>
	</tr>
</table>
<hr size=0	style='border: 1px dotted #999999;' align=left width=600>
<TABLE WIDTH=460 CLASS=FONT9 BORDER=0>
	<tr>
		<td  nowrap align=right>統計年月</td>
		<td nowrap><input name=inym class=inputbox size=8 value="<%=inym%>" ></td>
		<TD nowrap align=right>廠別</TD>
		<TD >
			<select name=WHSNO  class=txt8 onchange="datachg()">
				<option value=""></option>
				<%SQL="SELECT * FROM BASICCODE WHERE FUNC='WHSNO' ORDER BY SYS_TYPE "
				SET RST = CONN.EXECUTE(SQL)
				WHILE NOT RST.EOF
				%>
				<option value="<%=RST("SYS_TYPE")%>" <%IF  RST("SYS_TYPE")=whsno THEN %> SELECTED <%END IF%> ><%=RST("SYS_VALUE")%></option>
				<%
				RST.MOVENEXT
				WEND
				rst.close
				%>
			</SELECT>
			<%SET RST=NOTHING %>
		</TD>		
		<TD nowrap align=right >國籍</TD>
		<TD >
			<select name=COUNTRY  class=txt8  onchange="datachg()" >
				<option value=""></option>
				<%SQL="SELECT * FROM BASICCODE WHERE FUNC='COUNTRY' ORDER BY SYS_TYPE "
				SET RST = CONN.EXECUTE(SQL)
				WHILE NOT RST.EOF
				%>
				<option value="<%=RST("SYS_TYPE")%>" <%IF RST("SYS_TYPE")=country THEN %> SELECTED <%END IF%>><%=RST("SYS_TYPE")%><%=RST("SYS_VALUE")%></option>
				<%
				RST.MOVENEXT
				WEND
				rst.close
				%>
			</SELECT>
			<%SET RST=NOTHING %>			
		</TD>
		
	</tr>
	<TR height=25 >

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
		<TD nowrap align=right >部門</TD>
		<TD >
			<select name=GROUPID  class=txt8  onchange="datachg()">
				<option value=""></option>
				<%SQL="SELECT * FROM BASICCODE WHERE FUNC='GROUPID' and sys_type <>'AAA' ORDER BY SYS_TYPE "
				SET RST = CONN.EXECUTE(SQL)
				WHILE NOT RST.EOF
				%>
				<option value="<%=RST("SYS_TYPE")%>" <%IF  RST("SYS_TYPE")=GROUPID THEN %> SELECTED <%END IF%> ><%=RST("SYS_TYPE")%><%=RST("SYS_VALUE")%></option>
				<%
				RST.MOVENEXT
				WEND
				rst.close
				%>
			</SELECT>
			<%SET RST=NOTHING 
			conn.close
			set conn=nothing
			%>
		</TD>
		
		<TD nowrap align=right >班別</TD>
		<TD >
			<select name=shift  class=txt8  onchange="datachg()" >
				<option value="" <%if shift="" then %> selected<%end if%>></option>
				<option value="ALL" <%if shift="ALL" then %> selected<%end if%>>常日班</option>
				<option value="A" <%if shift="A" then %> selected<%end if%>>A班</option>
				<option value="B" <%if shift="B" then %> selected<%end if%>>B班</option>
			</SELECT>					
		</TD>
		<TD nowrap align=right>統計</TD>
		<TD >
			<select name=IOemp class=txt8 onchange="datachg()" > 
				<option value="Y" <%if IOemp="Y" then %>selected<%end if%>>在職Tai chuc</option>
			 	<option value="" <%if IOemp="" then %>selected<%end if%>>全部ALL</option>
			 	<option value="N" <%if IOemp="N" then %>selected<%end if%>>已離職Toai Viec</option>
			 </select>	
		</TD> 				
				
		<TD nowrap align=right >員工編號</TD>
		<TD >
			<INPUT NAME=empid1 SIZE=8 CLASS=INPUTBOX value="<%=QUERYX%>">			
		</TD>		
		<td><INPUT TYPE=BUTTON NAME=BTN VALUE="查詢" CLASS=BUTTON onclick="datachg()" ONKEYDOWN="DATACHG()"></td>
	</TR>
	<!--TR>
		< TD nowrap align=right>簽約</TD>
		<TD >
			<select name=outemp class=font9 onchange="datachg()"> 
			 	<option value="" <%if outemp="" then %>selected<%end if%>>全部</option>
			 	<option value="Y" <%if outemp="Y" then %>selected<%end if%>>已簽約</option>
			 	<option value="N" <%if outemp="N" then %>selected<%end if%>>未簽約</option>
			 </select>	
		</TD> 
	</TR-->
</TABLE>
<hr size=0	style='border: 1px dotted #999999;' align=left width=600>
<!-------------------------------------------------------------------->
<TABLE CLASS="txt8" BORDER=0   cellspacing="1" cellpadding="1" bgcolor="black">
 	<TR BGCOLOR="LightGrey" HEIGHT=25   >
 		<TD width=55 nowrap align=center>YYMM</TD>
 		<TD width=55 nowrap align=center>Quoc Tich</TD>
 		<TD width=55 nowrap align=center>Xuong</TD>
 		<TD width=55 nowrap align=center>bo phan</TD>
 		<TD width=55 nowrap align=center>工號<br>Ma So</TD>
 		<TD width=190 nowrap align=center>姓名<br>Ho Ten</TD>
 		<TD width=70 nowrap align=center>到職日期<br>NVX</TD>
 		<TD width=70 nowrap align=center>Chuc vu</TD>
 		<TD width=50 nowrap align=center>幣別</TD>
 		<TD width=60 nowrap align=center>BB</TD>
 		<TD width=60 nowrap align=center>CV</TD>
 		<TD width=60 nowrap align=center>PHU</TD>
 		<TD width=60 nowrap align=center>NN</TD>
 		<TD width=60 nowrap align=center>KT</TD>
 		<TD width=60 nowrap align=center>MT</TD>
 		<TD width=60 nowrap align=center>TTKH</TD>
 		<TD width=60 nowrap align=center>QC</TD>
 		<TD width=60 nowrap align=center>TNKH</TD>
 		TD width=60 nowrap align=center>TBTR</TD>
 		TD width=60 nowrap align=center>JX</TD>
 		<TD width=60 nowrap align=center>H1M</TD>
 		<TD width=60 nowrap align=center>H2M</TD>
 		<TD width=60 nowrap align=center>H3M</TD>
 		<TD width=60 nowrap align=center>B3M</TD>
 		<
 	</TR>
	<%for CurrentRow = 1 to PageRec
		IF CurrentRow MOD 2 = 0 THEN
			WKCOLOR="LavenderBlush"
		ELSE
			WKCOLOR=""
		END IF
		'if tmpRec(CurrentPage, CurrentRow, 1) <> "" then
	%>
	<TR BGCOLOR='<%=WKCOLOR%>' height=22>
 		<TD align=center nowrap>
 			<a href='vbscript:oktest(<%=CurrentRow-1%>)'><%=tmpRec(CurrentPage, CurrentRow, 52)%></a>
 			<input type=hidden value="<%=tmpRec(CurrentPage, CurrentRow, 17)%>" name=aid >
 		</TD>
 		<TD align=center nowrap><!--國籍-->
			<%=tmpRec(CurrentPage, CurrentRow, 16)%>
			<input type=hidden value="<%=tmpRec(CurrentPage, CurrentRow, 4)%>" name=c1 >
 		</TD>
 		<TD align=center nowrap> <!--廠別-->
 			<%=tmpRec(CurrentPage, CurrentRow, 11)%>
 		</TD>
 		<TD align=LEFT ><!--shift+部門-->
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then%>
 				<%=tmpRec(CurrentPage, CurrentRow, 21)%>-<%=tmpRec(CurrentPage, CurrentRow, 13)%>-<%=tmpRec(CurrentPage, CurrentRow, 14)%>
 			<%end if%>	
 		</TD> 
 		<TD nowrap>
 			<a href='vbscript:oktest(<%=CurrentRow-1%>)'>
 				<%=tmpRec(CurrentPage, CurrentRow, 2)%>&nbsp;<font class=txt8VN><%=tmpRec(CurrentPage, CurrentRow, 3)%></font>
 			</a>
 		</TD>
 		<TD align=center nowrap>
 			<%=tmpRec(CurrentPage, CurrentRow, 5)%><BR><%=tmpRec(CurrentPage, CurrentRow, 53)%>
 		</TD>
 		<Td><%=tmpRec(CurrentPage, CurrentRow, 14)%></td>
 		<Td><%=tmpRec(CurrentPage, CurrentRow, 50)%></td>
 		<Td align=right><%=tmpRec(CurrentPage, CurrentRow, 18)%></td>
 		<Td  align=right><%=tmpRec(CurrentPage, CurrentRow, 19)%></td>
 		<Td  align=right><%=tmpRec(CurrentPage, CurrentRow, 20)%></td>
 		<Td  align=right><%=tmpRec(CurrentPage, CurrentRow, 21)%></td>
 		<Td  align=right><%=tmpRec(CurrentPage, CurrentRow, 22)%></td>
 		<Td  align=right><%=tmpRec(CurrentPage, CurrentRow, 23)%></td>
 		<Td  align=right><%=tmpRec(CurrentPage, CurrentRow, 24)%></td>
 		<Td  align=right><%=tmpRec(CurrentPage, CurrentRow, 25)%></td>
 		<Td  align=right><%=tmpRec(CurrentPage, CurrentRow, 26)%></td>
 		<Td  align=right><%=tmpRec(CurrentPage, CurrentRow, 27)%></td>
 		<Td  align=right><%=tmpRec(CurrentPage, CurrentRow, 28)%></td>
 		<Td  align=right><%=tmpRec(CurrentPage, CurrentRow, 29)%></td>
 		<Td  align=right><%=tmpRec(CurrentPage, CurrentRow, 30)%></td>
 		<Td  align=right><%=tmpRec(CurrentPage, CurrentRow, 31)%></td>
 		<Td  align=right><%=tmpRec(CurrentPage, CurrentRow, 32)%></td>
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

function oktest(index)
	N=<%=self%>.aid(index).value
	c1=<%=self%>.c1(index).value
	'alert c1
	'tp=<%=self%>.totalpage.value
	'cp=<%=self%>.CurrentPage.value
	'rc=<%=self%>.RecordInDB.value
	if c1="VN" then 
		'open "../employee/empfile/empfile.foregnd.asp?empautoid="& N  , "_balnk"  , "top=10, left=10, width=620, scrollbars=yes" 
		open "<%=self%>.editVN.asp?empautoid="& N  , "_balnk"  , "top=10, left=10, width=650, height=500, scrollbars=yes" 
	else
		open "<%=self%>.editHW.asp?empautoid="& N  , "_balnk"  , "top=10, left=10, width=650, height=500, scrollbars=yes" 
		'open "../employee/empfile/empfile.foregnd.asp?empautoid="& N  , "_balnk"  , "top=10, left=10, width=600, scrollbars=yes" 
	end if 		
end function

</script>


<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<!-- #include file = "../GetSQLServerConnection.fun" -->
<!-- #include file="../ADOINC.inc" -->
<!--#include file="../include/sideinfo.inc"-->
<%
'on error resume next
session.codepage="65001"
SELF = "YEBQ01"

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
inym = request("inym")   
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
	if inym<>"" then 
		sqlstr = sqlstr & " AND  convert(char(6),indat,112)<='"& inym &"'   "	
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
				tmpRec(i, j, 23)=RS("passportNo")
				tmpRec(i, j, 24)=RS("visano")
				
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

	<table border=0 style="height:70px"><tr><td>&nbsp;</td></tr></table>
	
	<table width="100%" BORDER=0 align=center cellpadding=0 cellspacing=0 >
		<tr>
			<td align="center">
				<table class="txt" cellpadding=3 cellspacing=3>
					<tr>
						<td  nowrap align=right>統計年月</td>
						<td nowrap><input type="text" style="width:100px" name=inym  size=8 value="<%=inym%>" ></td>
						<TD nowrap align=right>廠別</TD>
						<TD >
							<select name=WHSNO   onchange="datachg()" style="width:100px">
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
							<select name=COUNTRY    onchange="datachg()" style="width:100px">
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
					<TR>
						<TD nowrap align=right >部門</TD>
						<TD >
							<select name=GROUPID    onchange="datachg()" style="width:100px">
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
							<select name=shift    onchange="datachg()" style="width:100px">
								<option value="" <%if shift="" then %> selected<%end if%>></option>
								<option value="ALL" <%if shift="ALL" then %> selected<%end if%>>常日班</option>
								<option value="A" <%if shift="A" then %> selected<%end if%>>A班</option>
								<option value="B" <%if shift="B" then %> selected<%end if%>>B班</option>
							</SELECT>					
						</TD>
						<TD nowrap align=right>統計</TD>
						<TD >
							<select name=IOemp  onchange="datachg()" style="width:100px"> 
								<option value="Y" <%if IOemp="Y" then %>selected<%end if%>>在職Tai chuc</option>
								<option value="" <%if IOemp="" then %>selected<%end if%>>全部ALL</option>
								<option value="N" <%if IOemp="N" then %>selected<%end if%>>已離職Toai Viec</option>
							 </select>	
						</TD> 				
								
						<TD nowrap align=right >員工編號</TD>
						<TD >
							<INPUT type="text" style="width:100px" NAME=empid1 SIZE=8  value="<%=QUERYX%>">			
						</TD>		
						<td><INPUT TYPE=BUTTON NAME=BTN VALUE="查 詢" class="btn btn-sm btn-outline-secondary" onclick="datachg()" ONKEYDOWN="DATACHG()"></td>
					</TR>
				</TABLE>
			</td>
		</tr>
		<tr>
			<td>
				<table id="myTableGrid" width="98%">
					<TR class="header">									
						<TD width=55 nowrap align=center>工號<br><font class="txt8">Ma So</font></TD>
						<TD width=190 nowrap align=center>姓名<br><font class="txt8">Ho Ten</font></TD>
						<TD width=70 nowrap align=center>到職日期<br><font class="txt8">NVX</font></TD>
						<TD width=70 nowrap align=center>簽合同日<BR><font class="txt8">Ngay ky hop dong</font></TD> 		
						<TD width=70 nowrap align=center>離職日期<BR><font class="txt8">NTX</font></TD>
						<TD width=70 nowrap align=center>護照號碼/發證日<BR><font class="txt8">ngay cap</font></TD>
						<TD width=50 nowrap align=center>簽證號碼/發證地<BR><font class="txt8">noi cap</font></TD>
						<TD width=60 nowrap align=center>職等<BR><font class="txt8">Chuc vu</font></TD>
						<TD width=30 nowrap align=center>班別<BR><font class="txt8">Loai Ca</font></TD>
						<TD width=75 nowrap align=center>單位部門<BR><font class="txt8">Don Vi</font></TD>
						<TD width=50 nowrap align=center>廠別<BR><font class="txt8">Loai Xuong</font></TD>
						<TD width=40 nowrap align=center>國籍<BR><font class="txt8">Quoc tich</font></TD>
					</TR>
					<%for CurrentRow = 1 to PageRec
						IF CurrentRow MOD 2 = 0 THEN
							WKCOLOR="LavenderBlush"
						ELSE
							WKCOLOR=""
						END IF
						'if tmpRec(CurrentPage, CurrentRow, 1) <> "" then
					%>
					<TR BGCOLOR='<%=WKCOLOR%>' class="txt">

						<TD align=center nowrap>
							<a href='vbscript:oktest(<%=CurrentRow-1%>)'><%=tmpRec(CurrentPage, CurrentRow, 1)%></a>
							<input type=hidden value="<%=tmpRec(CurrentPage, CurrentRow, 17)%>" name=aid >
						</TD>
						<TD nowrap>
							<a href='vbscript:oktest(<%=CurrentRow-1%>)'><%=tmpRec(CurrentPage, CurrentRow, 2)%>&nbsp;<%=tmpRec(CurrentPage, CurrentRow, 3)%></a>
						</TD>
						<TD align=center nowrap><%=tmpRec(CurrentPage, CurrentRow, 5)%></TD>
						<TD align=center nowrap><!--簽約日-->
							<%=tmpRec(CurrentPage, CurrentRow, 19)%>
						</TD>
						
						<TD align=center nowrap><!--離職日-->
							<%=tmpRec(CurrentPage, CurrentRow, 20)%>
						</TD>
						<TD align=LEFT nowrap><!--護照號/發證日-->
							<%=tmpRec(CurrentPage, CurrentRow, 23)%>
						</TD>
						<TD align=LEFT nowrap><!--簽證號/發證地-->
							<%=tmpRec(CurrentPage, CurrentRow, 24)%>
						</TD>
						<TD align=LEFT ><!--職等-->
							<%=left(tmpRec(CurrentPage, CurrentRow, 15),4)%>
						</TD> 		
						<TD align=LEFT nowrap><!--班別-->
							<%=tmpRec(CurrentPage, CurrentRow, 21)%>
						</TD>
						<TD align=LEFT ><!--部門-->
							<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then%>
								<%=tmpRec(CurrentPage, CurrentRow, 13)%>-<%=tmpRec(CurrentPage, CurrentRow, 14)%>
							<%end if%>	
						</TD>
						<TD align=center nowrap> <!--廠別-->
							<%=tmpRec(CurrentPage, CurrentRow, 11)%>
						</TD>
						<TD align=center nowrap><!--國籍-->
							<%=tmpRec(CurrentPage, CurrentRow, 16)%>
							<input type=hidden value="<%=tmpRec(CurrentPage, CurrentRow, 4)%>" name=c1 >
						</TD>
					</TR>
					<%next%>
				</TABLE>
			</td>
		</tr>
		<tr>
			<td align="center">
				<table class="txt">
					<tr class="txt">
						<td align="CENTER" height=40 width=80%>
						PAGE:<%=CURRENTPAGE%> / <%=TOTALPAGE%> , COUNT:<%=RECORDINDB%><BR>
						<% If CurrentPage > 1 Then %>
							<input type="submit" name="send" value="FIRST" class="btn btn-sm btn-outline-secondary">
							<input type="submit" name="send" value="BACK" class="btn btn-sm btn-outline-secondary">
						<% Else %>
							<input type="submit" name="send" value="FIRST" disabled class="btn btn-sm btn-outline-secondary">
							<input type="submit" name="send" value="BACK" disabled class="btn btn-sm btn-outline-secondary">
						<% End If %>
						<% If cint(CurrentPage) < cint(TotalPage) Then %>
							<input type="submit" name="send" value="NEXT" class="btn btn-sm btn-outline-secondary">
							<input type="submit" name="send" value="END" class="btn btn-sm btn-outline-secondary">
						<% Else %>
							<input type="submit" name="send" value="NEXT" disabled class="btn btn-sm btn-outline-secondary">
							<input type="submit" name="send" value="END" disabled class="btn btn-sm btn-outline-secondary">
						<% End If %>
						</td>
						<td>	<BR>
							<input type="button" name="send" value="回主畫面"   class="btn btn-sm btn-outline-secondary" onclick="history.back()">
						</td>
					</TR>
				</TABLE>
			</td>
		</tr>
	</table>
			
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


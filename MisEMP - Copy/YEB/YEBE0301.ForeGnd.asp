<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<!-- #include file = "../GetSQLServerConnection.fun" -->
<!-- #include file="../ADOINC.inc" -->
<!--#include file="../include/sideinfo.inc"-->
<%
'on error resume next
session.codepage="65001"
SELF = "YEBE0301"

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
EMPID = REQUEST("empid")
shift = request("shift")
IOemp = request("IOemp")
inym = request("inym")
zuno = request("zuno")
gTotalPage = 1
PageRec = 20    'number of records per page
TableRec = 30    'number of fields per record
'NOWMONTH=CSTR(YEAR(DATE()))&"/"&RIGHT("00"&CSTR(MONTH(DATE())),2)&"/01"
NOWMONTH=CSTR(YEAR(DATE()))&"/"&RIGHT("00"&CSTR(MONTH(DATE())),2)&"/"&RIGHT("00"&CSTR(day(DATE())),2)

Set fso = Server.CreateObject("Scripting.FileSystemObject")

sqlstr = "select * from view_empfile where   1=1  " & _ 
	"  and  case when '"&country&"' ='' then '' else country end = '"&country&"'  "&_
  "  and charindex(   case when '"&whsno&"'='' then ',' else  '"&whsno&"' end , ','+whsno ) > 0 "&_
	"  and charindex(   case when '"&groupid&"'='' then ',' else  '"&groupid&"' end  , ','+groupid ) > 0 "&_
	"  and charindex(   case when '"&zuno&"'='' then ',' else  '"&zuno&"' end , ','+zuno ) > 0 " 
  '"( empid<>'PELIN' and isnull(status,'')<>'D' )   AND isnull(whsno,'') like '"& whsno &"%' 
	'and isnull(unitno,'') like '"& unitno &"%'  and isnull(groupid,'') like '"& groupid &"%'  "&_
	'"and country like '"& country &"%'  and isnull(zuno,'') like '"& zuno &"%' and isnull(shift,'') like '"& shift &"%' 
	'and   ( EMPID like '%"& QUERYX &"%'  or empnam_VN like '"& QUERYX &"%'  or empnam_CN like '"& QUERYX &"%')  "
	if  outemp="Y" then
		sqlstr = sqlstr & " AND isnull(bhdat,'')<>'' "
	elseif 	outemp="N"  then
		sqlstr = sqlstr & " AND isnull(bhdat,'')=''  "
	end if
	if EMPID<>"" THEN
		sqlstr = sqlstr & " and EMPID like '%"& EMPID &"%'  "
	end if
	if IOemp="Y" then
		sqlstr = sqlstr & " AND ( ISNULL(OUTDAT,'')='' OR convert( char(6),OUTDAT,112)>='"& NOWMONTH &"' )  "
	elseif IOemp="N" then
		sqlstr = sqlstr & " AND ( ISNULL(OUTDAT,'')<>'' )  "
	end if
	if inym<>"" then
		sqlstr = sqlstr & " and convert(char(6),indat,112)=   '"& inym &"'  "
	end if
sqlstr = sqlstr & "order by empid  "
'response.write sqlstr 

sql="select * from fn_View_EMPFILE ('"& whsno &"','"& country&"','"& QUERYX &"' , '"& inym &"', '"& groupid &"','"&zuno &"', '"&shift&"','"&IOemp &"' ) where 1=1  "
if IOemp="Y" then
		sql = sql & " AND ( ISNULL(OUTDAT,'')='' OR convert( char(6),OUTDAT,111)>='"& NOWMONTH &"' )  "		
	elseif IOemp="N" then
		sql = sql & " AND ( ISNULL(OUTDAT,'')<>'' )  "
	end if
	if inym<>"" then
		sql = sql & " and convert(char(6),indat,112)=   '"& inym &"'  "
	end if 
sql = sql & "order by empid  "	
'response.write "<br>"& sql
if request("TotalPage") = "" or request("TotalPage") = "0" then
	CurrentPage = 1
	'rs.Open sqlstr, conn, 3, 3
	rs.Open sql, conn, 3, 1
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
				if rs("country")<>"VN" then 
					if rs("wkd_no")<>"" then 
						tmpRec(i, j, 24)=RS("wkd_no")
					end if 
				end if 	

				filename=Server.MapPath("pic/"&rs("empid")&".jpg")
				If fso.FileExists(filename) and  isnull(rs("photos"))=false Then
					tmpRec(i, j, 25)="Y"
				else
					tmpRec(i, j, 25)="N"
				end if
				'tmpRec(i, j, 25)="Y"
				
				'pass_filename=Server.MapPath("ppvisa/"&rs("empid")&"_pass.pdf")
				'If fso.FileExists(pass_filename)   Then
				'	tmpRec(i, j, 26)="Y"
				'else
				'	tmpRec(i, j, 26)="N"
				'end if 
				tmpRec(i, j, 26)="N"
				'visa_filename=Server.MapPath("ppvisa/"&rs("empid")&"_visa.pdf")
				'If fso.FileExists(visa_filename)   Then
				'	tmpRec(i, j, 27)="Y"
				'else
				'	tmpRec(i, j, 27)="N"
				'end if 				
				tmpRec(i, j, 27)="N"
				tmpRec(i, j, 28)=rs("taxcode")
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

Set fso = Nothing

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
<body  onkeydown="enterto()" onload="f()">
<form name="<%=self%>" method="post" action="<%=self%>.foregnd.asp">
<INPUT TYPE=HIDDEN NAME="UID" VALUE=<%=SESSION("NETUSER")%>>
<INPUT TYPE=hidden NAME=TotalPage VALUE="<%=TotalPage%>">
<INPUT TYPE=hidden NAME=CurrentPage VALUE="<%=CurrentPage%>">
<INPUT TYPE=hidden NAME=RecordInDB VALUE="<%=RecordInDB%>">
<INPUT TYPE=hidden NAME=PageRec VALUE="<%=PageRec%>">
<INPUT TYPE=hidden NAME=gTotalPage VALUE="<%=gTotalPage%>">
 
	<table border=0 style="height:70px"><tr><td>&nbsp;</td></tr></table>
	<table width="100%" border=0 >
		<tr>
			<td>
				<table width="98%" BORDER=0 align=center cellpadding=0 cellspacing=0 >
					<tr>
						<td>
							<table class="txt"  cellpadding=3 cellspacing=3>
								<tr>
									<TD nowrap align=right >國籍<br><font class="txt8">Quốc tịch</font></TD>
									<TD >
										<select name=COUNTRY  onchange="datachg()" style="width:100px">
											<option value=""></option>
											<%SQL="SELECT * FROM BASICCODE WHERE FUNC='COUNTRY' ORDER BY SYS_TYPE "
											SET RST = CONN.EXECUTE(SQL)
											WHILE NOT RST.EOF
											%>
											<option value="<%=RST("SYS_TYPE")%>" <%IF RST("SYS_TYPE")=country THEN %> SELECTED <%END IF%>><%=RST("SYS_TYPE")%>-<%=RST("SYS_TYPE")%><%=RST("SYS_VALUE")%></option>
											<%
											RST.MOVENEXT
											WEND
											rst.close
											%>
										</SELECT>
										<%SET RST=NOTHING %>			
									</TD>
									<TD nowrap align=right>廠別<br><font class="txt8">Xưởng</font></TD>
									<TD >
										<select name=WHSNO   onchange="datachg()"  style="width:120px">
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
									<TD nowrap align=right >部門<br><font class="txt8">Bộ phận</font></TD>
									<TD>
										<select name=GROUPID    onchange="datachg()"  style="width:120px">
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
											SET RST=NOTHING
											%>
										</SELECT>
									</td>
									<td>
										<select name=zuno    onchange="datachg()"  style="width:100px">
											<option value=""></option>
											<%SQL="SELECT * FROM BASICCODE WHERE FUNC='zuno' and left(sys_type,4) like '"&groupid &"%' and sys_type <>'AAA' ORDER BY SYS_TYPE "
											SET RST = CONN.EXECUTE(SQL)
											WHILE NOT RST.EOF
											%>
											<option value="<%=RST("SYS_TYPE")%>" <%IF  RST("SYS_TYPE")=zuno THEN %> SELECTED <%END IF%> ><%=RST("SYS_TYPE")%><%=RST("SYS_VALUE")%></option>
											<%
											RST.MOVENEXT
											WEND
											rst.close
											SET RST=NOTHING
											%>
										</SELECT>			
									</TD>
									<TD nowrap align=right >班別<br><font class="txt8">ca</font></TD>
									<TD >
										<select name=shift    onchange="datachg()"  style="width:100px">
											<option value="" <%if shift="" then %> selected<%end if%>></option>
											<option value="ALL" <%if shift="ALL" then %> selected<%end if%>>常日班 Làm ngày bình thường</option>
											<option value="A" <%if shift="A" then %> selected<%end if%>>A班 Ca A </option>
											<option value="B" <%if shift="B" then %> selected<%end if%>>B班 Ca B</option>
										</SELECT>					
									</TD>
								</tr>
								<tr>
									<TD nowrap align=right >員工編號<br><font class="txt8">Mã số nhân viên</font></TD>
									<TD >
										<INPUT type="text" NAME=empid1 SIZE=10  value="<%=QUERYX%>">
									</TD>
									<td  nowrap align=right>年月 Năm Tháng<br><font class="txt8">NVX(YYMM)</font></td>
									<td nowrap><input type="text" name=inym  size=8 value="<%=inym%>" ></td>
									<TD nowrap align=right>統計 Thống kế<br><font class="txt8">Loai</font></TD>
									<TD >
										<select name=IOemp  onchange="datachg()" >
											<option value="Y" <%if IOemp="Y" then %>selected<%end if%>>在職 Tại chức</option>
											<option value="" <%if IOemp="" then %>selected<%end if%>>全部 Toàn bộ</option>
											<option value="N" <%if IOemp="N" then %>selected<%end if%>>已離職 Đã nghỉ việc</option>
										 </select>
									</TD>
									<td colspan=3 nowrap>
										<INPUT TYPE=BUTTON NAME=BTN VALUE="(S)查詢K.Tra" class="btn btn-sm btn-outline-secondary" onclick="datachg()" ONKEYDOWN="DATACHG()">
									</td>
								</tr>
								<%conn.close
								set conn=nothing
								%>
							</table>							
						</td>
					</tr>
					<tr>
						<td>
							<table id="myTableGrid" width="98%"> 
								<TR BGCOLOR="LightGrey" HEIGHT=25   >
									<TD nowrap align=center>photo</TD>
									<TD nowrap align=center>工號<br><font class="txt8">Mã số</font></TD>
									<TD nowrap align=center>姓名<br><font class="txt8">Họ Tên</font></TD>
									<TD nowrap align=center>到職日期<br><font class="txt8">NVX</font></TD>
									<TD nowrap align=center>簽合同日<BR><font class="txt8">Ngày ký hợp đồng</font></TD>
									<TD nowrap align=center>離職日期<BR><font class="txt8">Ngày thôi việc</font></TD>
									<TD nowrap align=center>稅號<BR><font class="txt8">Mã số thuế</font></TD>
									<TD nowrap align=center>護照號碼/發證日<BR><font class="txt8">Hộ chiếu/Ngày cấp</font></TD>
									<TD nowrap align=center>簽證號碼/發證地<BR><font class="txt8">Thị thực/Nơi cấp</font></TD>
									<TD nowrap align=center>職等<BR><font class="txt8">Chức vụ</font></TD>
									<TD nowrap align=center>班別<BR><font class="txt8">Loại Ca</font></TD>
									<TD nowrap align=center>單位部門<BR><font class="txt8">Đơn vị</font></TD>
									<TD nowrap align=center>廠別<BR><font class="txt8">Xưởng</font></TD>
									<TD nowrap align=center>國籍<BR><font class="txt8">Quốc Tịch</font></TD>
								</TR>
								<%for CurrentRow = 1 to PageRec
									IF CurrentRow MOD 2 = 0 THEN
										WKCOLOR="LavenderBlush"
									ELSE
										WKCOLOR=""
									END IF
									'if tmpRec(CurrentPage, CurrentRow, 1) <> "" then
								%>
								<TR BGCOLOR='<%=WKCOLOR%>' height=22 class=txt8>
									<TD width=50 nowrap align=center>
										<%if tmpRec(CurrentPage, CurrentRow, 25) = "Y" then%>
											<img src="pic/<%=tmpRec(CurrentPage, CurrentRow, 1)%>.jpg" border=0 width=30 height=30>
										<%end if%>
									</TD>
									<TD align=center>
										<a href='vbscript:oktest(<%=CurrentRow-1%>)'><%=tmpRec(CurrentPage, CurrentRow, 1)%></a>
										<input type=hidden value="<%=tmpRec(CurrentPage, CurrentRow, 17)%>" name="aid" >
										<input type=hidden value="<%=tmpRec(CurrentPage, CurrentRow, 1)%>" name="f_empid" >
									</TD>
									<TD class=txt8>
										<a href='vbscript:oktest(<%=CurrentRow-1%>)'><%=tmpRec(CurrentPage, CurrentRow, 2)%>&nbsp;<font class=txt8VN><%=tmpRec(CurrentPage, CurrentRow, 3)%></font></a>
									</TD>
									<TD align=center><%=tmpRec(CurrentPage, CurrentRow, 5)%></TD>
									<TD align=center><!--簽約日-->
										<%=tmpRec(CurrentPage, CurrentRow, 19)%>
									</TD>
									<TD align=center><!--離職日-->
										<%=tmpRec(CurrentPage, CurrentRow, 20)%>
									</TD>
									<TD align=center><!--MST-->
										<%=tmpRec(CurrentPage, CurrentRow, 28)%>
									</TD>
									<TD align=LEFT nowrap class=txt8><!--護照號/發證日-->
										<%if tmpRec(CurrentPage, CurrentRow, 26)="N" then  %>
											<%=tmpRec(CurrentPage, CurrentRow, 23)%>
										<%else%>
											<a href="ppvisa/<%=tmpRec(CurrentPage, CurrentRow, 1)%>_pass.pdf" target="_balnk">
											<font color="blue"><%=tmpRec(CurrentPage, CurrentRow, 23)%></font></a>
										<%end if%>
									</TD>
									<TD align=LEFT nowrap class=txt8 ><!--簽證號/發證地-->
										<%if tmpRec(CurrentPage, CurrentRow, 26)="N" then  %>
											<%=tmpRec(CurrentPage, CurrentRow, 24)%>
										<%else%>
											<a href="ppvisa/<%=tmpRec(CurrentPage, CurrentRow, 1)%>_visa.pdf" target="_balnk">
											<font color="blue"><%=tmpRec(CurrentPage, CurrentRow, 24)%></font></a>
										<%end if%>	
									</TD>
									<TD align=LEFT class=txt8 ><!--職等-->
										<%=left(tmpRec(CurrentPage, CurrentRow, 15),5)%>
									</TD>
									<TD align=LEFT><!--班別-->
										<%=tmpRec(CurrentPage, CurrentRow, 21)%>
									</TD>
									<TD align=LEFT nowrap class=txt8><!--部門-->
										<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then %>
											<%=tmpRec(CurrentPage, CurrentRow, 13)%>-<%=tmpRec(CurrentPage, CurrentRow, 14)%>
										<%end if%>
									</TD>
									<TD align=center class=txt8> <!--廠別-->
										<%=tmpRec(CurrentPage, CurrentRow, 11)%>
									</TD>
									<TD align=center class=txt8><!--國籍-->
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
							<table class="table-borderless table-sm bg-white text-secondary">
								<tr>
									<td align="CENTER" height=40 >
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
									<td><input type=button  name=btm class="btn btn-sm btn-outline-secondary" value="SaveTo EXCEL" onclick=goexcel() style='background-color:#e4e4e4'></td>
									<td>	
										<input type="button" name="send" value="回主畫面 về trang trước "   class="btn btn-sm btn-outline-secondary" onclick="BACKMAIN()">
									</td>
									
									</TR>
								</TABLE>
						</td>
					</tr>
				</table>
			</td>
		</tr>
	</table>

</form>




</body>
</html>

<script language=vbscript>
function BACKMAIN()
	open "<%=self%>.asp" , "_self"
end function

function goexcel()
	
	<%=self%>.action="<%=self%>.toexcel.asp"
	<%=self%>.target="Back"
	<%=self%>.submit()
	parent.best.cols="100%,0%"
end function 

function oktest(index)
	N=<%=self%>.aid(index).value
	c1=<%=self%>.c1(index).value
	empid=<%=self%>.f_empid(index).value
	'alert c1
	'tp=<%=self%>.totalpage.value
	'cp=<%=self%>.CurrentPage.value
	'rc=<%=self%>.RecordInDB.value
	wt = (window.screen.width )*0.5
	ht = window.screen.availHeight*0.8
	tp = (window.screen.width )*0.05
	lt = (window.screen.availHeight)*0.1	
	
	if c1="VN" or c1="CT" then		
		open "<%=self%>.index.asp?ct="& c1 &"&empautoid="& N &"&empid="& empid , "balnkN"  , "top="& tp &", left="& lt &", width="& wt &",height="& ht &", scrollbars=yes"
	else
		open "<%=self%>.editHW.asp?ct="& c1 &"&empautoid="& N &"&empid="& empid  , "balnkN"  , "top="& tp &", left="& lt &", width="& wt &",height="& ht &", scrollbars=yes"
	end if
end function

</script>


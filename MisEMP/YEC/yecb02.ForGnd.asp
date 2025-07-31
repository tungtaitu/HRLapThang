<%@language=vbscript CODEPAGE=65001%>
<!---------  #include file="../GetSQLServerConnection.fun"  -------->
<!--#include file="../include/sideinfo.inc"-->
<%
Dim gTotalPage, PageRec, TableRec
Dim CurrentRow, CurrentPage, TotalPage, RecordInDB
Dim tmpRec, i, j, k, SELF, conn, rs, Source
Dim WK_COLOR, StartToAdd
session.codepage=65001
SELF = "yecb02"
Set conn = GetSQLServerConnection()

gTotalPage = 1
PageRec = 15    'number of records per page
TableRec = 7    'number of fields per record
DB_TBLID = request("DB_TBLID")
'on error resume next

'sqln="select * from scode_big where  tblid='"& DB_TBLID &"' "
'set rds=conn.execute(sqln)
'scodeBig = rds("description")
if request("TotalPage") = "" then
   CurrentPage = 1

 	source = "select a.*, isnull(b.cnt,0) cnt   from  "&_
 			 "( SELECT  *  from SCode_big where tblid <> '' and isnull(status,'')<>'D' and left(tblid,1)='*' ) a "&_
 			 "left join (select func, count(*) as cnt from BASICCODE   group by func ) b on b.func = a.tblid "&_
 			 " order by a.tblid  "
	'response.write  source
	'response.end
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open Source, conn, 1, 3
	if not rs.eof then
		PageRec = rs.RecordCount
		rs.PageSize = PageRec
		RecordInDB = rs.RecordCount
		TotalPage = rs.PageCount
		gTotalPage = totalpage
	end if

	'Set conn = nothing

	Redim tmpRec(gTotalPage, PageRec, TableRec)   'Array
	for i = 1 to TotalPage
		for j = 1 to PageRec
			if not rs.EOF then
				tmpRec(i, j, 0) = "no"
				tmpRec(i, j, 1) = rs("tblid")
				tmpRec(i, j, 2) = rs("Description")
				tmpRec(i, j, 3) = rs("cnt")
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
	Session("YDBSB0001EMP") = tmpRec

else
	TotalPage = cint(request("TotalPage"))
	gTotalPage = cint(request("gTotalPage"))
	StoreToSession()
	CurrentPage = cint(request("CurrentPage"))
	tmpRec = Session("YDBSB0001EMP")
	RecordInDB = request("RecordInDB")
	Select case request("send")
		Case "FIRST"
			 CurrentPage = 1
		Case "BACK"
			 if cint(CurrentPage) <> 1 then
				CurrentPage = CurrentPage - 1
			 end if
		Case "NEXT"
			 if cint(CurrentPage) < cint(gTotalPage) then
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
<meta HTTP-EQUIV="Content-Type" content="text/html; charset=UTF-8">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
<link rel="stylesheet" href="../Include/style.css" type="text/css">
<link rel="stylesheet" href="../Include/style2.css" type="text/css">
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>

function m(index)
   <%=SELF%>.send(index).style.backgroundcolor="lightyellow"
   <%=SELF%>.send(index).style.color="red"
end function

function n(index)
   <%=SELF%>.send(index).style.backgroundcolor="khaki"
   <%=SELF%>.send(index).style.color="black"
end function 

'-----------------enter to next field
function enterto()
	if window.event.keyCode = 13 then window.event.keyCode =9
end function

</SCRIPT>
</head>
<body   topmargin=40     >  
<form action="<%=SELF%>.ForGnd.asp" name="<%=SELF%>">
 
<INPUT TYPE=hidden NAME=TotalPage VALUE="<%=TotalPage%>">
<INPUT TYPE=hidden NAME=CurrentPage VALUE="<%=CurrentPage%>">
<INPUT TYPE=hidden NAME=PageRec VALUE="<%=PageRec%>">
<INPUT TYPE=hidden NAME=RecordInDB VALUE="<%=RecordInDB%>">
<INPUT TYPE=hidden NAME=gTotalPage VALUE="<%=gTotalPage%>">
<input TYPE=hidden name=DB_TBLID value="<%=request("DB_TBLID")%>">

	<table border=0 style="height:70px"><tr><td>&nbsp;</td></tr></table>
	<table width="100%" BORDER=0 align=center cellpadding=0 cellspacing=0 >
		<tr>
			<td align="center">
				<table id="myTableGrid" width="98%">
				  <tr class="header" height="35px">
					<td style="width:50px">DEL<br>XÓA</td>
					<td >大類<BR>Mã Code</td>
					<td >說明<br>Thuyết minh</td>
					<td >廠別<br>Xưởng</td>
					<td >國籍<br>Quốc tịch</td>
					<td style="width:50px">編輯<br>sửa</td>
				  </tr>
				  <%
					for CurrentRow = 1 to PageRec

					j = 1
					if j=1 then
						wk_color = "#E1E8F0"
						j = 0
					else
						wk_color = "#EBEDF1"
						j = 1
					end if
				  %>
				  <TR bgcolor="<%=wk_color%>">
					<td   align=center  style="width:50px">
						<%if session("rights")="0" then %>
						<%if tmpRec(CurrentPage, CurrentRow, 0) = "del" then%>
							<input type=hidden name=func value=del onclick="del(<%=CurrentRow - 1%>)" checked>
							<input type=hidden name=op value=del>
						<%else%>
							<input type=hidden name=func value=no onclick="del(<%=CurrentRow - 1%>)" <%=mode%>>
							<input type=hidden name=op value=no>
						<%end if%>
						<%else%>
						<input type=hidden name=op value=no>
						<input type=hidden name=func value=no>
						<%end if%>
					</td>
					<TD align=center valign="top">
						<%if tmpRec(CurrentPage, CurrentRow, 1) = ""   then%>
							 <input type="text" style="width:98%" maxlength=10 name=tblcd onblur="tblcd_change(<%=CurrentRow - 1%>)">
						<%else%>
							<input  type="text" style="width:98%;BACKGROUND-COLOR: lightyellow" maxlength=10 name=tblcd Readonly value="<%=Ucase(trim(tmpRec(CurrentPage, CurrentRow, 1)))%>">
						<%end if%>
					</TD>
					<TD  align=left valign="top">
					  <input  type="text" style="width:98%" maxlength=40 name=tbldesc value="<%=Ucase(trim(tmpRec(CurrentPage, CurrentRow, 2)))%>" onchange="tblcd_change(<%=CurrentRow - 1%>)">
					</TD>
					<td>
						<select name=f_w1 style="width:98%" >
						<%sql="select * from basicCode where func='WHSNO' and sys_type not in ('WP', 'ALL') order by sys_type   "
						  set rds=conn.execute(sql)
						  while not rds.eof
						%>
							<option value="<%=rds("sys_type")%>"  <%if rds("sys_type")=session("mywhsno") then%>selected<%end if%>><%=rds("sys_type")%>-<%=rds("sys_value")%></option>
						<%rds.movenext
						wend
						set rds=nothing 
						%>
						</select>
					</td>
					<td>
						<select name=f_ct style="width:98%">
						<%sql="select * from basicCode where func='country' and sys_type='VN' order by sys_type desc "
						  set rds=conn.execute(sql)
						  while not rds.eof
						%>
							<option value="<%=rds("sys_type")%>"  ><%=rds("sys_type")%>-<%=rds("sys_value")%></option>
						<%rds.movenext
						wend
						set rds=nothing 
						%>
						</select>
					</td>
					<Td align=center  style="width:60px" nowrap>
						<%if trim(tmpRec(CurrentPage, CurrentRow, 1))="*BB" then %>
							<a href="vbscript:editdata(<%=currentrow-1%>)">Sửa(編輯)</a>
						<%end if%>
					</td>
				  </TR>
				  <%next%>
				</TABLE>
			</td>
		</tr>
		<tr>
			<td align="center">
				<table  border=0><tr><td align=center class=txt>Page: <%=CurrentPage%> / <%=totalpage%> , Count:<%=recordinDB%></td></tr></table>
			</td>
		</tr>
	</table>
			
</form>

</body>
</html>
<%
Sub StoreToSession()
	Dim CurrentRow
	tmpRec = Session("YDBSB0001EMP")
	for CurrentRow = 1 to PageRec
		tmpRec(CurrentPage, CurrentRow, 0) = request("op")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 1) = request("tblcd")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 2) = request("tbldesc")(CurrentRow)
	next
	Session("YDBSB0001EMP") = tmpRec
End Sub
%>

<script language="vbscript">
<!--

function Go()
   <%=SELF%>.action = "<%=SELF%>.UpdateDB.asp"
   <%=SELF%>.submit
end function

function Clear()
	open "<%=SELF%>.asp", "_self"
end function

function del(index)
	 if <%=SELF%>.func(index).checked then
		<%=SELF%>.op(index).value = "del"

		open "<%=SELF%>.BackGnd.asp?tblcd=" & tblcd_str & _
		     "&CurrentPage=" & <%=CurrentPage%> & _
		     "&index=" & index & "&func=del", "Back"
	 end if
end function

function tblcd_change(index)
	<%=SELF%>.op(index).value = "update"

	tblcd_str = Ucase(trim(<%=SELF%>.tblcd(index).value))
	tbldesc_str = escape(trim(<%=SELF%>.tbldesc(index).value))
	'<%=SELF%>.tblcd(index).value = tblcd_str

	open "<%=SELF%>.BackGnd.asp?tblcd=" & tblcd_str & _
         "&CurrentPage=" & <%=CurrentPage%> & _
         "&tbldesc=" & tbldesc_str & _
         "&index=" & index & "&func=tblcd_change", "Back"
        'parent.best.cols="70%,30%"
end function

function editdata(index)
	DB_TBLID = trim(<%=self%>.f_w1(index).value)
	f_ct = trim(<%=self%>.f_ct(index).value)
	if DB_TBLID<>"" then
		open "<%=self%>.fore_vn.asp?f_w1=" & DB_TBLID &"&f_ct="& f_ct, "_self"
	end if
end function

//-->
</script>

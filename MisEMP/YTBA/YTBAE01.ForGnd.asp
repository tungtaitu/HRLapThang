<%@language=vbscript CODEPAGE=65001%>
<!---------  #include file="../GetSQLServerConnection.fun"  -------->
<!--#include file="../include/sideinfo.inc"-->
<%
Dim gTotalPage, PageRec, TableRec
Dim CurrentRow, CurrentPage, TotalPage, RecordInDB
Dim tmpRec, i, j, k, SELF, conn, rs, Source
Dim WK_COLOR, StartToAdd
session.codepage=65001
SELF = "YTBAE01"
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
  if session("netuser")="PELIN" or session("netuser")="NGHIA" then 	
	 	source = "select a.*, isnull(b.cnt,0) cnt   from  "&_
	 			 "( SELECT  *  from SCode_big where tblid <> '' and isnull(status,'')<>'D'  ) a "&_
	 			 "left join (select func, count(*) as cnt from BASICCODE   group by func ) b on b.func = a.tblid "&_
	 			 " order by a.tblid  "
	else
		source = "select a.*, isnull(b.cnt,0) cnt   from  "&_
	 			 "( SELECT  *  from SCode_big where tblid <> '' and isnull(status,'')<>'D' and left(tblid,1)<>'*' and tblid<>'grp'   ) a "&_
	 			 "left join (select func, count(*) as cnt from BASICCODE   group by func ) b on b.func = a.tblid "&_
	 			 " order by a.tblid  "
	end if 
	'response.write  source
	'response.end
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open Source, conn, 1, 3
	if not rs.eof then
		rs.PageSize = PageRec
		RecordInDB = rs.RecordCount
		TotalPage = rs.PageCount
		gTotalPage = totalpage+1
	end if

	Set conn = nothing

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

<SCRIPT ID=clientEventHandlersVBS LANGUAGE="javascript">

function m(index){
   <%=SELF%>.send(index).style.backgroundcolor="lightyellow";
   <%=SELF%>.send(index).style.color="red";
}

function n(index){
   <%=SELF%>.send(index).style.backgroundcolor="khaki";
   <%=SELF%>.send(index).style.color="black";
}

'-----------------enter to next field
function enterto(){
	if (window.event.keyCode == 13) window.event.keyCode =9;
}

</SCRIPT>
</head>
<body >  
<form action="<%=SELF%>.ForGnd.asp" name="<%=SELF%>"> 
	<INPUT TYPE=hidden NAME=TotalPage VALUE="<%=TotalPage%>">
	<INPUT TYPE=hidden NAME=CurrentPage VALUE="<%=CurrentPage%>">
	<INPUT TYPE=hidden NAME=PageRec VALUE="<%=PageRec%>">
	<INPUT TYPE=hidden NAME=RecordInDB VALUE="<%=RecordInDB%>">
	<INPUT TYPE=hidden NAME=gTotalPage VALUE="<%=gTotalPage%>">
	<input TYPE=hidden name=DB_TBLID value="<%=request("DB_TBLID")%>">
	<input TYPE=hidden name="pgid" value="<%=request("pgid")%>">
	<table border=0 style="height:70px"><tr><td>&nbsp;</td></tr></table>
	
	<table width="100%" border=0 >
		<tr>
			<td align="center">
				<table width="80%" id="myTableGrid">
					<tr class="header" style="height:50px">
						<td style="width:10%">刪除</td>
						<td style="width:30%">大類</td>
						<td style="width:50%">說明</td>
						<td style="width:10%">編輯</td>
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
					<tr align="center" style="height:25px">
						<td>
						<%if session("rights")="0" then %>
						<%if tmpRec(CurrentPage, CurrentRow, 0) = "del" then%>
							<input type=checkbox name=func value=del onclick="del(<%=CurrentRow - 1%>)" checked>
							<input type=hidden name=op value=del>
						<%else%>
							<input type=checkbox name=func value=no onclick="del(<%=CurrentRow - 1%>)" <%=mode%>>
							<input type=hidden name=op value=no>
						<%end if%>
						<%else%>
						<input type=hidden name=op value=no>
						<input type=hidden name=func value=no>
						<%end if%>
						</td>
						<td>
						<%if tmpRec(CurrentPage, CurrentRow, 1) = ""   then%>
							<input size=8 maxlength=10 class="form-control form-control-sm" name="tblcd" id="tblcd" onblur="tblcd_change(<%=CurrentRow - 1%>)">
						<%else%>
							<input size=8 maxlength=10 class="form-control form-control-sm" name="tblcd" id="tblcd" Readonly style="BACKGROUND-COLOR: lightyellow" value="<%=Ucase(trim(tmpRec(CurrentPage, CurrentRow, 1)))%>">
						<%end if%>
						</TD>
						<TD >
							<input size=60 class="form-control form-control-sm" maxlength=50 name="tbldesc" id="tbldesc" value="<%=Ucase(trim(tmpRec(CurrentPage, CurrentRow, 2)))%>" onchange="tblcd_change(<%=CurrentRow - 1%>)">
						</TD>
						<Td >
							<a href="javascript:editdata(<%=currentrow-1%>)"><span class="fa fa-pencil"></span> (<%=tmpRec(CurrentPage, CurrentRow, 3)%>)</a>
						</td>
					</tr>
      <%next%>
					<tr align="left">
						<td colspan="4">
							Page: <%=CurrentPage%> / <%=totalpage%> , Count:<%=recordinDB%>&nbsp;
							<% If CurrentPage > 1 Then %>
									<input type="submit" name="send" value="FIRST" class="btn btn-sm btn-outline-secondary">
									<input type="submit" name="send" value="BACK" class="btn btn-sm btn-outline-secondary">
								<% Else %>
									<input type="submit" name="send" value="FIRST" disabled class="btn btn-sm btn-outline-secondary">
									<input type="submit" name="send" value="BACK" disabled class="btn btn-sm btn-outline-secondary">
								<% End If %>

								<% If cint(CurrentPage) < cint(gTotalPage) Then %>
									<input type="submit" name="send" value="NEXT" class="btn btn-sm btn-outline-secondary">
									<input type="submit" name="send" value="END" class="btn btn-sm btn-outline-secondary">
								<% Else %>
									<input type="submit" name="send" value="NEXT" disabled class="btn btn-sm btn-outline-secondary">
									<input type="submit" name="send" value="END" disabled class="btn btn-sm btn-outline-secondary">
								<% End If %>
						</td>
					</tr>
					<tr align="center">
						<td colspan="4">
						<%if UCASE(session("mode"))="W" then%>
							<INPUT TYPE="button" name=send VALUE="(Y) Confirm" class="btn btn-sm btn-outline-secondary"  onClick="Go()">
							<INPUT TYPE="button" name=send VALUE="(N) Cancel" class="btn btn-sm btn-danger"  onClick="Clear()">
						<%end if%>
						</td>
					</tr>
				</TABLE>	
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

<script language="javascript">

function editdata(index){
	
	DB_TBLID = <%=self%>.tblcd[index].value;	
	
	if (DB_TBLID != "") {		
		open("<%=self%>01.forgnd.asp?pgid=<%=request("pgid")%>&DB_TBLID=" + DB_TBLID , "_self");
	}
}

function Go(){
	
   <%=SELF%>.action = "<%=SELF%>.UpdateDB.asp?pgid=<%=request("pgid")%>";   
   <%=SELF%>.submit();
}

function Clear(){
	open("<%=SELF%>.asp?pgid=<%=request("pgid")%>", "_self");
}

function del(index){
	 if(<%=SELF%>.func(index).checked){
		<%=SELF%>.op(index).value = "del";
		open("<%=SELF%>.BackGnd.asp?pgid=<%=request("pgid")%>&tblcd="+ tblcd_str +"&CurrentPage=" + <%=CurrentPage%> +"&index=" + index + "&func=del", "Back");
	 }
}

function tblcd_change(index)
{
	<%=SELF%>.op(index).value = "update";
	tblcd_str = Ucase(<%=SELF%>.tblcd(index).value.trim());
	tbldesc_str = escape(<%=SELF%>.tbldesc(index).value.trim());
	open("<%=SELF%>.BackGnd.asp?pgid=<%=request("pgid")%>&tblcd=" + tblcd_str +"&CurrentPage=" + <%=CurrentPage%> + "&tbldesc="+ tbldesc_str +"&index="+ index + "&func=tblcd_change", "Back");
}

</script>

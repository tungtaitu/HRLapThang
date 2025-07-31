<%@language=vbscript CODEPAGE=65001%>
<!---------  #include file="../GetSQLServerConnection.fun"  -------->
<!--#include file="../include/sideinfo.inc"-->
<%
Dim gTotalPage, PageRec, TableRec
Dim CurrentRow, CurrentPage, TotalPage, RecordInDB
Dim tmpRec, i, j, k, SELF, conn, rs, Source
Dim WK_COLOR, StartToAdd
session.codepage=65001
SELF = "YTBAE0101" 
Set conn = GetSQLServerConnection()
	
gTotalPage = 1
PageRec = 10    'number of records per page
TableRec = 7    'number of fields per record
DB_TBLID = request("DB_TBLID")
'on error resume next
queryx = request("queryx")
sqln="select * from scode_big where  tblid='"& DB_TBLID &"' "
set rds=conn.execute(sqln) 
scodeBig = rds("description")
set rds=nothing 
if request("TotalPage") = "" then 
   CurrentPage = 1
	 

	Source = "select * from BASICCODE where func = '" & request("DB_TBLID") & "' and   sys_type<>'AAA' and sys_type like '"&queryx&"%'  order by sys_type"		
	'response.write  source 
	'response.end 
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open Source, conn, 1, 3
	if not rs.eof then
		rs.PageSize = PageRec 
		RecordInDB = rs.RecordCount 
		TotalPage = rs.PageCount  
		gTotalPage = totalpage + 1 
	end if 	
	
	'Set conn = nothing 
	
	Redim tmpRec(gTotalPage, PageRec, TableRec)   'Array 	
	for i = 1 to TotalPage 
		for j = 1 to PageRec
			if not rs.EOF then 
				tmpRec(i, j, 0) = "no"
				tmpRec(i, j, 1) = rs("sys_type")
				tmpRec(i, j, 2) = rs("sys_value")
				tmpRec(i, j, 3) = "" 'rs("autoid")
				tmpRec(i, j, 4) = rs("autoid")
				tmpRec(i, j, 5) = rs("transcode")
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
	Session("YTBAE0101EMP") = tmpRec
	
else
	TotalPage = cint(request("TotalPage"))
	gTotalPage = cint(request("gTotalPage"))
	StoreToSession()
	CurrentPage = cint(request("CurrentPage"))
	tmpRec = Session("YTBAE0101EMP")
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

<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
 
function m(index)
   <%=SELF%>.send(index).style.backgroundcolor="lightyellow"
   <%=SELF%>.send(index).style.color="red"
end function

function n(index)
   <%=SELF%>.send(index).style.backgroundcolor="khaki"
   <%=SELF%>.send(index).style.color="black"
end function

function gos()
	<%=self%>.totalpage.value=""
	<%=self%>.action="<%=self%>.forgnd.asp"
	<%=self%>.submit()
end function 
 
</SCRIPT>
</head>
<body   topmargin=5  >  
<form action="<%=SELF%>.ForGnd.asp" name="<%=SELF%>">

<INPUT TYPE=hidden NAME=TotalPage VALUE="<%=TotalPage%>">
<INPUT TYPE=hidden NAME=CurrentPage VALUE="<%=CurrentPage%>">
<INPUT TYPE=hidden NAME=PageRec VALUE="<%=PageRec%>">
<INPUT TYPE=hidden NAME=RecordInDB VALUE="<%=RecordInDB%>">
<INPUT TYPE=hidden NAME=gTotalPage VALUE="<%=gTotalPage%>">
<input TYPE=hidden name=DB_TBLID value="<%=request("DB_TBLID")%>">
<table border=0 style="height:70px"><tr><td>&nbsp;</td></tr></table>

<table width="100%" border=0 >
	<tr>
		<td>
			<table width="80%" BORDER=0 align=center cellpadding=0 cellspacing=0 >
				<tr>
					<td align="right" width="50px">項目：</td>
					<td><input name=dex value="(<%=DB_TBLID%>) <%=scodeBig%>" class="form-control form-control-sm mr-sm-2"  readonly style="width:200px;BACKGROUND-COLOR: lightyellow"></td>
				<%if DB_TBLID="ZUNO" then %>
					<td align="right" width="50px">Don vi:</td>
					<td  width="200px">
						<select name="queryx" class="form-control form-control-sm mr-sm-2" onchange="gos()">
							<option value=""></option>
							<%sqlx="SELECT * FROM BASICCODE WHERE FUNC='GROUPID'  ORDER BY SYS_TYPE"
							  set rst=conn.execute(Sqlx)
							  while not rst.eof
							%>
							<option value="<%=RST("SYS_TYPE")%>" <%if RST("SYS_TYPE")=queryx then %>selected<%end if%>><%=RST("SYS_TYPE")%><%=RST("SYS_VALUE")%></option>
							<%
							rst.movenext
							wend
							rst.close
							set rst=nothing
							Set conn = nothing 
							%>						
						</select>
					</td>
				<%else%>	
					<input name="queryx" type="hidden">
				<%end if%>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td>
			<table width="80%" BORDER=0 align=center cellpadding=0 cellspacing=0 class="table-bordered table-sm bg-white text-secondary"> 
				<tr class="bg-gray text-black" align=center style="height:50px"> 
					<td style="width:5%">刪除</td>
					<td style="width:15%">大類</td>
					<td style="width:40%">代碼說明</td>
					<td style="width:20%">轉檔代碼</td>
					<td style="width:20%">備註</td>				
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
				<tr align="center" class="tr-hover">
					<td> 
       <%if tmpRec(CurrentPage, CurrentRow, 1) = "" then %>
       		<input type=hidden name=op value=no>
       		<input type=hidden name=func value=no>
       <%else%>
				<%if tmpRec(CurrentPage, CurrentRow, 0) = "del" then%> 
					<input type=checkbox name=func value=del onclick="del(<%=CurrentRow - 1%>)" checked>
					<input type=hidden name=op value=del>
				<%else%>
					<input type=checkbox name=func value=no onclick="del(<%=CurrentRow - 1%>)" <%=mode%>>
					<input type=hidden name=op value=no>
				<%end if%>				
			<%end if %>	
					</td>
					<td>
			<%if tmpRec(CurrentPage, CurrentRow, 1) = "" then%> 
				 <input size=8 maxlength=20 class="form-control form-control-sm mb-2 mt-2" name=tblcd onblur="tblcd_change(<%=CurrentRow - 1%>)">
			<%else%>
				<input size=8 maxlength=20 class="form-control form-control-sm mb-2 mt-2" name=tblcd Readonly style="BACKGROUND-COLOR: lightyellow" value="<%=Ucase(trim(tmpRec(CurrentPage, CurrentRow, 1)))%>">				 
			<%end if%>
					</TD>        
					<td> 
						<input size=30 class="form-control form-control-sm mb-2 mt-2" maxlength=250 name=tbldesc value="<%=Ucase(trim(tmpRec(CurrentPage, CurrentRow, 2)))%>" onchange="tblcd_change(<%=CurrentRow - 1%>)">
					</TD>
					<TD> 
					  <input size=20 class="form-control form-control-sm mb-2 mt-2" maxlength=250 name="transcode" value="<%=Ucase(trim(tmpRec(CurrentPage, CurrentRow, 5)))%>" onchange="tblcd_change(<%=CurrentRow - 1%>)">
					</TD>
					<TD > 
						<input size=20 class="form-control form-control-sm mb-2 mt-2" maxlength=250 name=memo2 value="<%=Ucase(trim(tmpRec(CurrentPage, CurrentRow, 3)))%>" onchange="tblcd_change(<%=CurrentRow - 1%>)">
					</TD>
				</TR>
      <%next%> 
				<tr>
					<td colspan="5" align="left">Page: <%=CurrentPage%> / <%=totalpage%> , Count:<%=recordinDB%>&nbsp;
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
				<tr align="center" style="height:25px" class="tr-hover">
					<td colspan="5">
						<%if UCASE(session("mode"))="W" then%>
						<INPUT TYPE="button" name=send VALUE="(Y) Confirm" class="btn btn-sm btn-outline-secondary"   onClick="Go()">		
						<INPUT TYPE="button" name=send VALUE="(N) Cancel" class="btn btn-sm btn-danger"   onClick="Clear()">
						<%end if%>
					</td>
				</tr>
			</table>	
		</td>
	</tr>
</table>

</form>

</body>
</html>
<%
Sub StoreToSession()
	Dim CurrentRow
	tmpRec = Session("YTBAE0101EMP")
	for CurrentRow = 1 to PageRec
		tmpRec(CurrentPage, CurrentRow, 0) = request("op")(CurrentRow)		
		tmpRec(CurrentPage, CurrentRow, 1) = request("tblcd")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 2) = request("tbldesc")(CurrentRow)		
		tmpRec(CurrentPage, CurrentRow, 3) = request("memo2")(CurrentRow)		
		tmpRec(CurrentPage, CurrentRow, 5) = request("transcode")(CurrentRow)		
	next 
	Session("YTBAE0101EMP") = tmpRec
End Sub
%>

<script language="vbscript">
<!--

function Go()
   <%=SELF%>.action = "<%=SELF%>.UpdateDB.asp"
   <%=SELF%>.submit
end function

function Clear()
	open "YTBAE01.asp", "_self"
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
	tbldesc_str = (trim(<%=SELF%>.tbldesc(index).value))
	memo_str = (trim(<%=SELF%>.memo2(index).value))
	tscode = (trim(<%=SELF%>.transcode(index).value))
	'<%=SELF%>.tblcd(index).value = tblcd_str
	
	open "<%=SELF%>.BackGnd.asp?tblcd=" & tblcd_str & _
         "&CurrentPage=" & <%=CurrentPage%> & _
         "&tbldesc=" & tbldesc_str & _
         "&memo2=" & memo_str & _
				 "&tscode=" & tscode & _
         "&index=" & index & "&func=tblcd_change", "Back" 
        'parent.best.cols="70%,30%" 
end function

//-->
</script>

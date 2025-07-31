<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<!-- #include file = "../GetSQLServerConnection.fun" --> 
<!-- #include file="../ADOINC.inc" --> 
<!--#include file="../include/sideinfo.inc"-->
<%
Dim gTotalPage, PageRec, TableRec
Dim CurrentRow, CurrentPage, TotalPage, RecordInDB
Dim tmpRec, i, j, k, SELF, conn, rs, Source
Dim WK_COLOR, StartToAdd
SESSION.CODEPAGE="65001"

SELF = "YEGBE0102"
gTotalPage = 1
PageRec = 15 'number of records per page
TableRec = 15    'number of fields per record
QUERY=REQUEST("QUERY")
QUERY2=REQUEST("QUERY2")

Set conn = GetSQLServerConnection()	

if request("TotalPage") = "" or request("TotalPage") = "0" then 
   CurrentPage = 1
  
	SQL=" select a.sys_type as  groupid  , isnull(b.sys_type,'')  as zuno , c.sys_type as shift  , a.sys_value as gstr,  "&_
		"isnull(b.sys_value , '') as zstr , c.sys_value as sstr ,   isnull(f.func,'') sc , isnull(g.uptim,'') uptim  "&_
		"from "&_
		"(  select  'T' as tmp, * from basiccode where func='groupid'  and sys_type<>'AAA' and sys_type >= 'A033' ) a "&_
		"left join (select * from basiccode where func='zuno')  b on  left(b.sys_type,4) = a.sys_type "&_
		"left join (select 'T' as tmp, * from basiccode where func='shift') c on c.tmp = a.tmp "&_
		"left join ( select max(mdtm) mdtm, groupid, zuno, shift, max(func) func  from empscg group by groupid, zuno, shift ) f on f.groupid = a.sys_type and  isnull(f.zuno,'') = isnull(b.sys_type,'') and f.shift = c.sys_type  "&_
		"left join ( select * from empsc ) g on g.func = f.func where a.sys_type like '"& query &"%' and c.sys_type like '%"& query2 &"' "&_
		"order by a.sys_type, b.sys_type,  len(c.sys_type) desc , c.sys_type   "
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open SQL, conn, 3, 3
	'response.write sql
	rs.PageSize = PageRec 
	RecordInDB = rs.RecordCount 
	TotalPage = rs.PageCount 
	gTotalPage = TotalPage 	
	
	Redim tmpRec(TotalPage, PageRec, TableRec)   'Array
	
	for i = 1 to TotalPage 
	 for j = 1 to PageRec
		if not rs.EOF then 
			for k=1 to TableRec-1
				tmpRec(i, j, 0) = "no"
				tmpRec(i, j, 1) = trim(rs("GROUPID"))
				tmpRec(i, j, 2) = trim(rs("ZUNO"))
				tmpRec(i, j, 3) = rs("SHIFT")
				tmpRec(i, j, 4) = rs("gstr")
				tmpRec(i, j, 5) = rs("zstr")
				tmpRec(i, j, 6) = rs("sstr")
				tmpRec(i, j, 7) = rs("sc")
				tmpRec(i, j, 8) = rs("uptim")
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
Session("YEGBE0102") = tmpRec	
else    
  TotalPage = cint(request("TotalPage"))
  GTotalPage = cint(request("GTotalPage"))
  RecordInDB = cint(request("RecordInDB"))
  StoreToSession()
  CurrentPage = cint(request("CurrentPage"))
  tmpRec = Session("YEGBE0102")

  Select case request("send") 
	     Case "FIRST"
		      CurrentPage = 1			
	     Case "BACK"
		      if cint(CurrentPage) <> 1 then 
			     CurrentPage = CurrentPage - 1				
		      end if
	     Case "NEXT"
		      if cint(CurrentPage) <= cint(TotalPage) then 
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
<meta http-equiv="refresh">
<link rel="stylesheet" href="../Include/style.css" type="text/css">
<link rel="stylesheet" href="../Include/style2.css" type="text/css"> 

<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--
function m(index)
   <%=SELF%>.send(index).style.backgroundcolor="lightyellow"
   <%=SELF%>.send(index).style.color="red"
end function

function n(index)
   <%=SELF%>.send(index).style.backgroundcolor="khaki"
   <%=SELF%>.send(index).style.color="black"
end function

function Q_Data()
   <%=self%>.TotalPage.value = ""
   <%=self%>.submit
end function 
-->
</SCRIPT>
<title>稅別</title></head>
<body >
<form method="post" action="<%=SELF%>.Fore.asp" name="<%=SELF%>">
<INPUT TYPE=HIDDEN NAME="UID" VALUE=<%=SESSION("NETUSER")%>>
<INPUT TYPE=hidden NAME=TotalPage VALUE="<%=TotalPage%>">
<INPUT TYPE=hidden NAME=CurrentPage VALUE="<%=CurrentPage%>">
<INPUT TYPE=hidden NAME=PageRec VALUE="<%=PageRec%>">
<INPUT TYPE=hidden NAME=RecordInDB VALUE="<%=RecordInDB%>">
<INPUT TYPE=hidden NAME=gTotalPage VALUE="<%=gTotalPage%>">

	<table border=0 style="height:70px"><tr><td>&nbsp;</td></tr></table>
	<table width="100%" border=0 >
		<tr>
			<td>
				<table width="80%" BORDER=0 align=center cellpadding=0 cellspacing=0 >
					<tr>
						<td>
							<table class="txt" cellpadding=3 cellspacing=3> 
								<TR>
									<TD>單位查詢</TD>
									<TD>
										<SELECT NAME=QUERY  onchange=Q_Data() style="width:120px">
											<OPTION VALUE="" <%IF QUERY="" THEN %>SELECTED<%END IF%>></OPTUON>
											<%SQLB="select * from basiccode where func='groupid'  and sys_type<>'AAA' and sys_type >= 'A033'  ORDER BY SYS_type  "
												SET RDS=CONN.EXECUTE(SQLB)						
												WHILE NOT RDS.EOF
											%>
												<OPTION VALUE="<%=RDS("SYS_type")%>" <%IF RDS("SYS_type")=QUERY THEN%>SELECTED<%END IF%>><%=RDS("SYS_TYPE")%>-<%=RDS("SYS_VALUE")%></OPTION>
											<%
											RDS.MOVENEXT
											WEND
											SET RDS=NOTHING 
											%>	
										</SELECT>
									</TD>
									<TD>班別查詢</TD>
									<TD>
										<SELECT NAME=QUERY2  onchange=Q_Data() style="width:120px">
											<OPTION VALUE="" <%IF QUERY2="" THEN %>SELECTED<%END IF%>></OPTUON>
											<%SQLB="select * from basiccode where func='shift'  ORDER BY len(SYS_type), sys_type   "
												SET RDS=CONN.EXECUTE(SQLB)						
												WHILE NOT RDS.EOF
											%>
												<OPTION VALUE="<%=RDS("SYS_type")%>" <%IF RDS("SYS_type")=QUERY2 THEN%>SELECTED<%END IF%>><%=RDS("SYS_TYPE")%>-<%=RDS("SYS_VALUE")%></OPTION>
											<%
											RDS.MOVENEXT
											WEND
											SET RDS=NOTHING 
											%>	
										</SELECT>
									</TD>
								</TR>
							</TABLE>
						</td>
					</tr>
					<tr>
						<td align="center">
							<table id="myTableGrid" width="80%"> 
								<tr class="header" height="35px"> 
									<td>單位</td>
									<td>組別</td>
									<td>班別</td>
									<td>班次</td>
								</tr>
							  <%for CurrentRow = 1 to PageRec
									if CurrentRow mod 2 = 0 then 		
										wk_color = ""
									else 
										wk_color = "#E4E4E4"			
									end if 
								%> 
								<TR bgcolor="<%=wk_color%>" class="txt">        
								<TD > 
									<%=Ucase(trim(tmpRec(CurrentPage, CurrentRow, 1)))%>-<%=Ucase(trim(tmpRec(CurrentPage, CurrentRow, 4 )))%>
									<input type=hidden size=18  class=readonly name=groupid   value="<%=Ucase(trim(tmpRec(CurrentPage, CurrentRow, 1)))%>">
								</TD>
								<TD > 
									<%=Ucase(trim(tmpRec(CurrentPage, CurrentRow, 2)))%>-<%=Ucase(trim(tmpRec(CurrentPage, CurrentRow, 5)))%>
									<input type=hidden size=18 class=zuno name=code   value="<%=Ucase(trim(tmpRec(CurrentPage, CurrentRow, 2)))%>">				 
								</TD>
								<TD  > 
									<%=Ucase(trim(tmpRec(CurrentPage, CurrentRow, 6)))%>
									<input type=hidden size=18 name=shift  class=inputbox  value="<%=tmpRec(CurrentPage, CurrentRow, 3)%>">				 
								 </TD>	
								 <TD  width="130px"> 
									<%IF trim(tmpRec(CurrentPage, CurrentRow, 1))<>"" THEN %>
										<SELECT NAME=FUNC class="txt9" style="width:100%" onchange=datachg(<%=currentrow-1%>)>
										<OPTION VALUE="" <%IF tmpRec(CurrentPage, CurrentRow, 7)="" THEN %>SELECTED<%END IF%> ></OPTUON>
										<%SQLA="SELECT* FROM EMPSC ORDER BY DESCP"
										SET RDS=CONN.EXECUTE(SQLA)
										WHILE NOT RDS.EOF 
										%>
										<OPTION VALUE="<%=RDS("FUNC")%>" <%IF RDS("FUNC")=tmpRec(CurrentPage, CurrentRow, 7) THEN %>SELECTED<%END IF%> ><%=RDS("FUNC")%>-<%=RDS("DESCP")%>-- <%=RDS("UPTIM")%></OPTUON>
										<%
										RDS.MOVENEXT
										WEND
										SET RDS=NOTHING
										%>
									<%ELSE%>
										<INPUT NAME=FUNC TYPE=HIDDEN>	
									<%END IF%>
								 </TD>	
							  </TR>
							  <%next%> 
							  <input name=func type=hidden >
							  <tr>
								<td colspan=4 class=txt9 align=center>
								第<INPUT size=6 readonly class=readonlyN name=ta style='BORDER-BOTTOM-STYLE: none; BORDER-LEFT: medium none;BORDER-RIGHT-STYLE: none; BORDER-TOP-STYLE: none' value="<%=CurrentPage%>">頁,
								共<INPUT size=6 readonly class=readonlyN name=ta style='BORDER-BOTTOM-STYLE: none; BORDER-LEFT: medium none;BORDER-RIGHT-STYLE: none; BORDER-TOP-STYLE: none' value="<%=Totalpage%>">頁
								</td>
							  </tr>
							</TABLE>
						</td>
					</tr>
					<tr>
						<td align="center">
							<table class="txt" cellpadding=3 cellspacing=3>
								<tr>
									<td align="left">
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
									<td>									
										<INPUT TYPE="button" name=send VALUE="Confirm" class="btn btn-sm btn-danger" onClick="Go()">	
										<INPUT TYPE="reset" name=send VALUE="Cancel" class="btn btn-sm btn-outline-secondary"  >
									</TD>
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
<%
Sub StoreToSession()
	Dim CurrentRow
	tmpRec = Session("YEGBE0102")
	for CurrentRow = 1 to PageRec		
		'tmpRec(CurrentPage, CurrentRow, 7) = request("func")(CurrentRow)
	next 
	Session("YEGBE0102") = tmpRec
End Sub
%>

<script language="vbscript">
<!--

function Go()
   <%=SELF%>.action = "<%=SELF%>.UpdateDB.asp"
   <%=SELF%>.submit
end function

function Clear()
	open "<%=SELF%>.asp", "Fore"
end function

    
function Datachg(index)
	Code_str = Ucase(trim(<%=SELF%>.func(index).value))
	
	open "<%=SELF%>.Back.asp?code=" & Code_str & _
         "&CurrentPage=" & <%=CurrentPage%> & _         
         "&index=" & index & "&func=datachg", "Back"
    'parent.best.cols="70%,30%"     
end function

//-->
</script>



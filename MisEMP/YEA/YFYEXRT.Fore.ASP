<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<!-- #include file = "../GetSQLServerConnection.fun" --> 
<!-- #include file="../ADOINC.inc" -->  
<!-- #include file="../Include/SIDEINFO.inc" -->
<%

Dim gTotalPage, PageRec, TableRec
Dim CurrentRow, CurrentPage, TotalPage, RecordInDB
Dim tmpRec, i, j, k, SELF, conn, rs, Source
Dim WK_COLOR, StartToAdd

SELF = "YFYEXRT"
gTotalPage = 10
PageRec = 10    'number of records per page
TableRec = 3    'number of fields per record
Query = request("Query")
if trim(Query) = "" then
   if request("TotalPage") = "" or request("TotalPage") = "0" then 
      CurrentPage = 1
	  Set conn = GetSQLServerConnection()	
	  Source = "select * from VYFYEXRT order by yyyymm desc ,code"
	 
	  Set rs = Server.CreateObject("ADODB.Recordset")
	  rs.Open Source, conn, 3, 3
		if not rs.eof then 
	  rs.PageSize = PageRec 
	  RecordInDB = rs.RecordCount 
	  TotalPage = rs.PageCount 
		gtotalpage= TotalPage +1 
		end if 
	

	
	  Redim tmpRec(gTotalPage, PageRec, TableRec)   'Array
	
	  for i = 1 to TotalPage 
	 	 for j = 1 to PageRec
			if not rs.EOF then 				
				tmpRec(i, j, 0) = "no"
				tmpRec(i, j, 1) = trim(rs("yyyymm"))
				tmpRec(i, j, 2) = trim(rs("code"))
				tmpRec(i, j, 3) = rs("exrt")				
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
	  Session("YSBAE1701") = tmpRec	
   else    
	  TotalPage = cint(request("TotalPage"))
		gTotalPage = cint(request("gTotalPage"))
	  StoreToSession()
	  CurrentPage = cint(request("CurrentPage"))
	  tmpRec = Session("YSBAE1701")
	
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
else
   CurrentPage = 1
   Set conn = GetSQLServerConnection()	
   Source = "Select * from VYFYEXRT Where Code like '" & trim(Query) & "%' " & _
            "or yyyymm like '" & trim(Query) & "%' " & _
            "order by yyyymm desc ,Code "

   Set rs = Server.CreateObject("ADODB.Recordset")
   rs.Open Source, conn, 3, 3
	
		if not rs.eof then 
	  rs.PageSize = PageRec 
	  RecordInDB = rs.RecordCount 
	  TotalPage = rs.PageCount 
		gtotalpage= TotalPage +1 
		end if 
	 
	
   Redim tmpRec(gTotalPage, PageRec, TableRec)   'Array
	
   for i = 1 to TotalPage 
 	  for j = 1 to PageRec
		 if not rs.EOF then 			
			tmpRec(i, j, 0) = "no"
			tmpRec(i, j, 1) = trim(rs("yyyymm"))
			tmpRec(i, j, 2) = trim(rs("code"))
			tmpRec(i, j, 3) = rs("exrt")			
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
   Session("YSBAE1701") = tmpRec	   
end if   
	  Set conn = nothing 
%>
<html>
<head>
<meta HTTP-EQUIV="Content-Type" content="text/html; charset=utf-8">
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

</head>
<body>
<form method="post" action="<%=SELF%>.Fore.asp" name="<%=SELF%>">
<INPUT TYPE=HIDDEN NAME="UID" VALUE=<%=SESSION("NETUSER")%>>
<INPUT TYPE=hidden NAME=TotalPage VALUE="<%=TotalPage%>">
<INPUT TYPE=hidden NAME=CurrentPage VALUE="<%=CurrentPage%>">
<INPUT TYPE=hidden NAME=PageRec VALUE="<%=PageRec%>">
<INPUT TYPE=hidden NAME=RecordInDB VALUE="<%=RecordInDB%>">
<INPUT TYPE=hidden NAME=gTotalPage VALUE="<%=gTotalPage%>">

	<table border=0 style="height:70px"><tr><td>&nbsp;</td></tr></table>
	
				<table width="100%" BORDER=0 align=center cellpadding=0 cellspacing=0 >
					<tr>
						<td align="center">
							<table id="myTableGrid" style="width:40%">
								<tr class="bg-gray text-black" align=center height="35px"> 
									<td width="80px">刪除<br>XÓA ĐI</td>
									<td >年月<br>YYYYMM</td>
									<td >幣別<br>LOẠI TIỀN</td>
									<td >匯率<br>TỈ GIÁ</td>
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
								<td align=center > 
									<%if tmpRec(CurrentPage, CurrentRow, 0) = "del" then%>
										<input type=checkbox name=func value=del onclick="del(<%=CurrentRow - 1%>)" checked>
										<input type=hidden name=op value=del>
									<%else%>
										<input type=checkbox name=func value=no onclick="del(<%=CurrentRow - 1%>)" <%=mode%>>
										<input type=hidden name=op value=no>
									<%end if%>				
								</td>
								<TD  align=center > 
									<%if tmpRec(CurrentPage, CurrentRow, 1) = ""    then%>
									<input  type="text" style="width:100px"  maxlength=6 name=yyyymm value="<%=Ucase(trim(tmpRec(CurrentPage, CurrentRow, 1)))%>" >
								</TD>
								<TD  align=center > 
									<INPUT type="text" style="width:100px"  maxlength=3 name=code value="<%=Trim(Ucase(tmpRec(CurrentPage, CurrentRow, 2)))%>" onchange="code_change(<%=CurrentRow - 1%>)">
								</TD>
								<TD  align=center> 
									<INPUT  type="text" style="width:100px"  maxlength=15 name=exrt value="<%=trim(Ucase(tmpRec(CurrentPage, CurrentRow, 3)))%>" onchange="exrt_change(<%=CurrentRow - 1%>)">
								</TD>
									<%elseif Ucase(trim(tmpRec(CurrentPage, CurrentRow, 1)))>=(year(date) & right(0 & month(date),2))  then %>
										<input  type="text" style="width:100px" maxlength=6  name=yyyymm Readonly style="BACKGROUND-COLOR: lightyellow" value="<%=Ucase(trim(tmpRec(CurrentPage, CurrentRow, 1)))%>">				 
										 </TD>
										 <TD align=center > 
										<input  type="text" style="width:100px" maxlength=3  name=code Readonly style="BACKGROUND-COLOR: lightyellow" value="<%=Ucase(trim(tmpRec(CurrentPage, CurrentRow, 2)))%>">				 
										 </TD>
										 <TD align=center > 
										<input  type="text" style="width:100px" maxlength=15  name=exrt value="<%=Trim(tmpRec(CurrentPage, CurrentRow, 3))%>" onchange="exrt_change(<%=CurrentRow - 1%>)">				 
										 </TD>
									<%else%>
										<input  type="text" style="width:100px" maxlength=6  name=yyyymm Readonly style="BACKGROUND-COLOR: lightyellow" value="<%=Ucase(trim(tmpRec(CurrentPage, CurrentRow, 1)))%>">				 
										 </TD>
										 <TD align=center > 
										<input  type="text" style="width:100px" maxlength=3  name=code Readonly style="BACKGROUND-COLOR: lightyellow" value="<%=Ucase(trim(tmpRec(CurrentPage, CurrentRow, 2)))%>">				 
										 </TD>
										 <TD  align=center > 
										<input  type="text" style="width:100px" maxlength=15  name=exrt Readonly style="BACKGROUND-COLOR: lightyellow" value="<%=tmpRec(CurrentPage, CurrentRow, 3)%>">				 
										 </TD>
									<%end if%>
							  </TR>
							  <%next%> 							  
							</TABLE>
						</td>
					</tr>
					<tr>
						<td align="center">
							<table class="txt">
								<tr>
									<td>count:</td>
									<td><INPUT  type="text" style="width:50px" readonly  name=ta  value="<%=CurrentPage%>"></td>
									<td>/,page:</td>
									<td><INPUT  type="text" style="width:50px" readonly  name=ta  value="<%=Totalpage%>"></td>
									
								</tr>
							</table>
						</td>
					</tr>
					<tr>
						<td align="center">
							<table class="txt">								
								<tr>
									<td align="center"  >
								<% If CurrentPage > 1 Then %>
									<input type="submit" name="send" value="FIRST" class="btn btn-sm btn-outline-secondary">
									<input type="submit" name="send" value="BACK" class="btn btn-sm btn-outline-secondary">
								<% Else %>
									<input type="submit" name="send" value="FIRST" disabled class="btn btn-sm btn-outline-secondary">
									<input type="submit" name="send" value="BACK" disabled class="btn btn-sm btn-outline-secondary">
								<% End If %>

								<% If cint(CurrentPage) <= cint(TotalPage) Then %>
									<input type="submit" name="send" value="NEXT" class="btn btn-sm btn-outline-secondary">
									<input type="submit" name="send" value="END" class="btn btn-sm btn-outline-secondary">
								<% Else %>      
									<input type="submit" name="send" value="NEXT" disabled class="btn btn-sm btn-outline-secondary">
									<input type="submit" name="send" value="END" disabled class="btn btn-sm btn-outline-secondary">	
								<% End If %>
									</td>
									<td  align="center">									
									<INPUT TYPE="button" name=send VALUE="(Y)Confirm" class="btn btn-sm btn-danger" onClick="go()">	
									<INPUT TYPE="button" name=send VALUE="(N)Cancel" class="btn btn-sm btn-outline-secondary" onClick="Clear()">
									</TD>
								</TR>
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
	tmpRec = Session("YSBAE1701")
	for CurrentRow = 1 to PageRec
		tmpRec(CurrentPage, CurrentRow, 0) = request("op")(CurrentRow)		
		tmpRec(CurrentPage, CurrentRow, 1) = request("yyyymm")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 2) = request("code")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 3) = request("exrt")(CurrentRow)
	next 
	Session("YSBAE1701") = tmpRec
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
	TaxCode_str = Ucase(trim(<%=SELF%>.code(index).value))
	Description_str = Ucase(trim(<%=SELF%>.yyyymm(index).value))
	Percentage_str = Ucase(trim(<%=SELF%>.exrt(index).value))
	 if <%=SELF%>.func(index).checked then 
		<%=SELF%>.op(index).value = "del"		
		open "<%=SELF%>.Back.asp?code=" & TaxCode_str & _            
		     "&CurrentPage=" & <%=CurrentPage%> & _
		     "&index=" & index & "&func=del", "Back"
	 else
	    <%=SELF%>.op(index).value = "update"		
		open "<%=SELF%>.Back.asp?code=" & TaxCode_str & _            
         "&CurrentPage=" & <%=CurrentPage%> & _
         "&yyyymm=" & Description_str & _
         "&exrt=" & Percentage_str & _
         "&index=" & index & "&func=TaxCode_change", "Back"
	 end if		
end function

function code_change(index)
    TaxCode_str = Ucase(trim(<%=SELF%>.code(index).value))
	<%=SELF%>.code(index).value = TaxCode_str
	Data_change(index)
end function

function exrt_change(index)
    Percentage_str = Ucase(trim(<%=SELF%>.exrt(index).value))
    <%=SELF%>.exrt(index).value = Percentage_str
    if Percentage_str <> "" then 
		if not isnumeric(Percentage_str) then 
			alert "¶ײv¸ꮆ¿養¿ù»~!!"
			<%=SELF%>.exrt(index).value = "0"
			<%=SELF%>.exrt(index).focus()		
			Percentage_str = "0"
		end if 
	end if
	Data_change(index)
end function

    
function Data_change(index)
	<%=SELF%>.op(index).value = "update"
	TaxCode_str = Ucase(trim(<%=SELF%>.code(index).value))
	Description_str = Ucase(trim(<%=SELF%>.yyyymm(index).value))
	Percentage_str = Ucase(trim(<%=SELF%>.exrt(index).value))
	open "<%=SELF%>.Back.asp?code=" & TaxCode_str & _
         "&CurrentPage=" & <%=CurrentPage%> & _
         "&yyyymm=" & Description_str & _
         "&exrt=" & Percentage_str & _
         "&index=" & index & "&func=TaxCode_change", "Back"
end function

//-->
</script>



<%@LANGUAGE=VBSCRIPT codepage=65001%>
<!-- #include file = "../GetSQLServerConnection.fun" --> 
<!-- #include file="../ADOINC.inc" -->
<%
self="Getempdata" 
Set conn = GetSQLServerConnection()	  
queryx = trim(request("queryx")) 
country = request("country") 
groupid = request("groupid") 
CurrentPage = REQUEST("CurrentPage") 
WKSCROLL = Request("WKSCROLL")     
JOBID=REQUEST("JOBID")
SHIFTA=REQUEST("SHIFTA")

if queryx="" then  
	sql="select * from view_empfile where country='VN'  "&_
		"and groupid like '%"& groupid &"%' "&_
		"AND ISNULL(JOB,'') like '%"& JOBID &"%' and ISNULL(SHIFT,'') like '%"& SHIFTA &"%' "&_
		"order by empid "
else
	sql="select * from view_empfile where empid like '%"& queryx &"%' or  empnaM_cn like '%"& queryx &"%'  order by empid  "
	
		 
end if 		 
'response.write sql 
'RESPONSE.END 
Set rs = Server.CreateObject("ADODB.Recordset")    
rs.Open sql,conn,3, 3 

If not rs.EOF then
	DIM PageSize,TotalRecords,TotalPage,WKSCROLL,CurrentPage
	RS.PageSize =  RS.RecordCount     
	TotalRecords = RS.RecordCount
	TotalPage = RS.PageCount   
	CurrentPage = int(Request("CurrentPage"))
	Select Case WKSCROLL
		Case ""
			CurrentPage = 1
		Case "FIRST"
			CurrentPage = 1
		Case "PRE"
			If CurrentPage <> 1 Then
				CurrentPage = CurrentPage - 1
			End If
		Case "NEXT"
			If CurrentPage < TotalPage Then
				CurrentPage = CurrentPage + 1		
			End If
		Case "END"
			CurrentPage = TotalPage
	End Select
	RS.AbsolutePage = CurrentPage    
end if 	
%> 
<html> 
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="refresh">
<link rel="stylesheet" href="../Include/style.css" type="text/css">
<link rel="stylesheet" href="../Include/style2.css" type="text/css"> 

<link rel="stylesheet" type="text/css" href="../template/bootstrap/css/bootstrap.min.css">
<link rel="stylesheet" type="text/css" href="../template/font-awesome/css/font-awesome.css">
<link rel="stylesheet" type="text/css" href="../template/css/mis.css">
<link rel="stylesheet" type="text/css" href="../template/datepicker/datepicker.css">

</head>
<body  topmargin="5" leftmargin="5"  marginwidth="0" marginheight="0" >
<form METHOD="POST" ACTION="<%=self%>.asp" name="<%=self%>">
<input TYPE="HIDDEN" NAME="CurrentPage" VALUE="<%=CurrentPage%>">

<table border=0 class=txt12  >
	<tr><td align=center ><b>員工資料查詢</b></td></tr>
</table>
<hr size=0	style='border: 1px dotted #999999;' align=left  >
	<table width="100%" border=0 >
		<tr>
			<td>
				<table width="80%" BORDER=0 align=center cellpadding=0 cellspacing=0 >
					<tr >
						<td>
							<table class="table-borderless table-sm bg-white text-secondary" class="txt9" >
								<tr class="txt9">
									<td align=right nowrap>國籍:</td>
									<td>
										<select name=country class=inputbox ONCHANGE="DATACHG()" >
											<option value=""></option>
											<%sql="select * from basicCode where func='country' order by sys_type"
											  set rst=conn.execute(sql)
											  while not rst.eof 
											%>
											<option value="<%=rst("sys_type")%>" <%if request("country")=rst("sys_type") then %> selected <%end if%>><%=rst("sys_value")%></option>	
											<%rst.movenext
											wend
											set rst=nothing 				
											%>
										</select>
									</td>
									<td align=right nowrap>單位:</td>
									<td>
										<select name=groupid class=inputbox ONCHANGE="DATACHG()">
											<option value=""></option>
											<%sql="select * from basicCode where func='groupid' order by sys_type"
											  set rst=conn.execute(sql)
											  while not rst.eof 
											%>
											<option value="<%=rst("sys_type")%>" <%if request("groupid")=rst("sys_type") then %> selected <%end if%> ><%=rst("sys_value")%></option>	
											<%rst.movenext
											wend
											set rst=nothing 				
											%>
										</select>
									</td>
								</tr>
								<tr>
									<td align=right nowrap>職務:</td>
									<td >
										<select name=JOBID class=inputbox ONCHANGE="DATACHG()" >
											<option value=""></option>
											<%sql="select * from basicCode where func='LEV' order by sys_type"
											  set rst=conn.execute(sql)
											  while not rst.eof 
											%>
											<option value="<%=rst("sys_type")%>" <%if request("JOBID")=rst("sys_type") then %> selected <%end if%>><%=LEFT(rst("sys_value"),5)%></option>	
											<%rst.movenext
											wend
											set rst=nothing 				
											%>
										</select>
									</td>
									<td align=right nowrap>班別:</td>
									<td>
										<select name=SHIFTA class=inputbox ONCHANGE="DATACHG()">
											<option value="" <%IF SHIFTA="" THEN %>SELECTED<%END IF%>></option>
											<option value="ALL" <%IF SHIFTA="ALL" THEN %>SELECTED<%END IF%>>常日班</option>
											<option value="A" <%IF SHIFTA="A" THEN %>SELECTED<%END IF%>>A班</option>
											<option value="B" <%IF SHIFTA="B" THEN %>SELECTED<%END IF%>>B班</option>				
										</select>
									</td>
								</tr>
								<tr>
									<td align=right nowrap>查詢:</td>
									<td colspan=3><input name="queryx" class=inputbox  >
										<INPUT TYPE="button" NAME=SEND VALUE="查詢" ONCLICK="DATACHG()" class="btn btn-sm btn-outline-secondary">
									</td>
								</tr>
							</table>
						</td>
					</tr>
					<tr>
						<td>
							<table class="table-borderless table-sm bg-white text-secondary">
								<Tr BGCOLOR="FFFFCC" class="txt9">
									<TD align=center nowrap>單位</TD>
									<TD align=center nowrap>工號</TD>
									<TD align=center nowrap>姓名</TD>
								</TR>
							<input type=HIDDEN name=EMPID value="">
							<input type=HIDDEN name=EMPNAME value="">
							<input type=HIDDEN name=whsno value="">	
							<% 
							x = 1
							DIM J,RowCount,IX,WK_COLOR 

							   j=0
							   RowCount = RS.PageSize 
							   IX = RowCount * (CurrentPage-1) + 1
							Do While Not RS.EOF and RowCount 
								'if (j mod 2)=1 then
								'	WK_COLOR="LIGHTYELLOW"
								'Else
								'	WK_COLOR="WHITE"
								'End if
								j=j+1
								Response.Write "<TR BGCOLOR=#FFFFFF class='txt9'>" & vbcrlf       
							%>
									<TD  HEIGHT=22 align=center> <A HREF="VBSCRIPT:GiveAns(<%=X%>)"  style ="{CURSOR: hand}"  >
										 <%=rs("GSTR")%></A>      
									</TD>   
									<TD  HEIGHT=22 align=center>
										 <A HREF="VBSCRIPT:GiveAns(<%=X%>)"  style ="{CURSOR: hand}"  >
										 <%=rs("empid")%></A>      
									</TD>   
									<TD  align=LEFT > 
										  <A HREF="VBSCRIPT:GiveAns(<%=X%>)"  style ="{CURSOR: hand}"  >
										  <%=rs("empNam_CN")%><%=rs("empNam_VN")%> 
										  </A>
										  <input type=HIDDEN name=EMPID value="<%=rs("empid")%>">
										  <input type=HIDDEN name=EMPNAME value="<%=rs("empNam_VN")%>">
										   <input type=HIDDEN name=whsno value="<%=rs("whsno")%>">
									</TD>
								</Tr>
							<%     
							   RowCount = RowCount - 1
							   IX = IX + 1   
							   rs.MoveNext
							   X = X + 1
							loop

							%>
							</Table>
						</td>
					</tr>
					<tr>
						<td>
							<table class="table-borderless table-sm bg-white text-secondary">
								<tr>
								<TD align=left nowrap>
									<% If CurrentPage > 1 Then %>
									<INPUT	TYPE="submit" NAME="WKSCROLL" VALUE="FIRST" class="btn btn-sm btn-outline-secondary"  >
									<INPUT	TYPE="submit" NAME="WKSCROLL" VALUE="PRE" 	class="btn btn-sm btn-outline-secondary"  >
									<% Else %>
									<INPUT	TYPE="submit" NAME="WKSCROLL" VALUE="FIRST" disabled 	class="btn btn-sm btn-outline-secondary">
									<INPUT	TYPE="submit" NAME="WKSCROLL" VALUE="PRE" disabled 	class="btn btn-sm btn-outline-secondary">
									<% End If %>
									<% If CurrentPage < TotalPage Then %>
									<INPUT	TYPE="submit" NAME="WKSCROLL" VALUE="NEXT" 	class="btn btn-sm btn-outline-secondary"  >
									<INPUT	TYPE="submit" NAME="WKSCROLL" VALUE="END" 	class="btn btn-sm btn-outline-secondary"   >	
									<% Else %>
									<INPUT	TYPE="submit" NAME="WKSCROLL" VALUE="NEXT" disabled class="btn btn-sm btn-outline-secondary">
									<INPUT	TYPE="submit" NAME="WKSCROLL" VALUE="END" disabled class="btn btn-sm btn-outline-secondary">
									<% End If %>
									<INPUT	TYPE="button" NAME="WKSCROLL" VALUE=" 關   閉"  class="btn btn-sm btn-outline-secondary"    onclick=GOClose()>
								</TD>
								</TR>	
							</table>
						</td>
					</tr>
					<tr>
						<td>
							<table class="table-borderless table-sm bg-white text-secondary">
								<TR CLASS=TXT9><TD ALIGN=CENTER>COUNT:<%=TotalRecords%>,頁次:第<%=CurrentPage%>頁,共<%=TotalPage%>頁</TD></TR>
							</TABLE>
						</td>
					</tr>
				</table>
			</td>
		</tr>
	</table>
</FORM> 
</body>

</html>

<SCRIPT LANGUAGE=VBSCRIPT>
FUNCTION DATACHG()
	<%=SELF%>.ACTION="<%=SELF%>.ASP"
	<%=SELF%>.SUBMIT()
END FUNCTION  

Function GiveAns(i)  
      Parent.Fore.VYFYSUCO.EMPID.value = <%=SELF%>.EMPID(i).Value 
      Parent.Fore.VYFYSUCO.cfdw.value = <%=SELF%>.EMPNAME(i).Value 
      Parent.Fore.VYFYSUCO.whsno.value = <%=SELF%>.whsno(i).Value 
      Parent.Fore.VYFYSUCO.sgcost.FOCUS()
      'Parent.Fore.<%=TargetName%>.<%=nextfocus%>.value = <%=ThisPage%>.prodname(i).Value
      Parent.best.cols = "100%,0%"
      'if trim("<%=cust%>") <> "" then
      '   Parent.Fore.<%=TargetName%>.<%=cust%>.value = <%=ThisPage%>.custid(i).Value 
      'end if     
End Function 

Function GoClose()
   Parent.best.cols = "100%,0%"	
End Function  
    
</SCRIPT>
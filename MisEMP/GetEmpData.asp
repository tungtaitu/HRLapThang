<%@LANGUAGE=VBSCRIPT codepage=65001%>
<!-- #include file="ADOINC.inc" -->
<!--#include file="GetSQLServerConnection.fun"  -->
<%
self="Getempdata" 
Set conn = GetSQLServerConnection()	  
queryx = trim(request("queryx")) 
country = request("country") 
groupid = request("groupid") 
CurrentPage = REQUEST("CurrentPage") 
'response.write "a="& CurrentPage
if CurrentPage="" then CurrentPage=0
WKSCROLL = Request("WKSCROLL")     
JOBID=REQUEST("JOBID")
SHIFTA=REQUEST("SHIFTA") 
formName=request("formName") 
index = request("index")  
nfs = request("nfs") 
if nfs="" then nfs="nsf"
whsno=request("whsno") 

'response.write index &"<BR>"
'response.write formName &"<BR>"

yymm=year(date())&right("00"&month(date()),2) 
if right("00"&month(date()),2) ="01" then 
	Lastym=year(date())-1&"12" 
else	
	lastym=year(date())&right("00"&month(date())-1,2) 
end if

if queryx="" and country="" and groupid="" and whsno="" and SHIFTA="" then 
	sql="select * from  view_empfile where country='tw' and isnull(outdate,'')='' and empid<>'PELIN'  order by   empid" 
else
	if queryx="" then  
		sql="select * from  view_empfile where empid<>''  "&_
			"and whsno like '"& whsno &"%' and ( isnull(outdate,'')='' or convert(char(6), outdat,112)>='"& lastym &"' ) and country like '"& country &"%' and groupid like '"& groupid &"%' "&_
			"AND ISNULL(JOB,'') like '"& JOBID &"%'  " 
			
		if 	SHIFTA<>"" then 
			sql=sql & " and ISNULL(SHIFT,'') = '"& SHIFTA &"' " 
		end if	
		sql =sql & "order by empid " 
	else
		sql="select * from  view_empfile where whsno like '"& whsno &"'  and empid like '%"& queryx &"%' or  empnaM_cn like '%"& queryx &"%'  order by empid   "
			 
	end if 	
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
<link rel="stylesheet" href="Include/style.css" type="text/css">
<link rel="stylesheet" href="Include/style2.css" type="text/css">
<link rel="stylesheet" href="Include/MisStyles.css" type="text/css"> 

</head>
<body  topmargin="5" leftmargin="5"  marginwidth="0" marginheight="0" >
<form METHOD="POST" ACTION="<%=self%>.asp" name="<%=self%>">
<input TYPE="HIDDEN" NAME="CurrentPage" VALUE="<%=CurrentPage%>">
<input type=HIDDEN name=index value="<%=index%>">
<input type=HIDDEN name=formName value="<%=formName%>">
<input type=HIDDEN  name=nfs value="<%=nfs%>">
<table width="100%" border=0 class=txt12  >
	<tr><td align=center ><b>員工資料查詢</b></td></tr>
</table>
<hr size=0	style='border: 1px dotted #999999;' align=left  >
	<table width="100%">
		<tr>
			<td>
				<table class="txt" cellpadding=3 cellspacing=3>
					<tr>
						<td nowrap align=right>廠別:</td>
						<td>
							<select name=whsno  ONCHANGE="DATACHG()" style="width:75px">
								<option value=""></option>
								<%sql="select * from dbo.basicCode where func='whsno' order by sys_type"
								  set rst=conn.execute(sql)
								  while not rst.eof 
								%>
								<option value="<%=rst("sys_type")%>" <%if request("whsno")=rst("sys_type") then %> selected <%end if%>><%=rst("sys_type")%><%=rst("sys_value")%></option>	
								<%rst.movenext
								wend
								set rst=nothing 				
								%>
							</select>
						</td>
						<td nowrap align=right>國籍:</td>
						<td  >
							<select name=country  ONCHANGE="DATACHG()" style="width:75px">
								<option value=""></option>
								<%sql="select * from dbo.basicCode where func='country' order by sys_type"
								  set rst=conn.execute(sql)
								  while not rst.eof 
								%>
								<option value="<%=rst("sys_type")%>" <%if request("country")=rst("sys_type") then %> selected <%end if%>><%=rst("sys_type")%><%=rst("sys_value")%></option>	
								<%rst.movenext
								wend
								set rst=nothing 				
								%>
							</select>
						</td>
					</tr>
					<tr>
						<td  nowrap align=right >單位:</td>
						<td>
							<select name=groupid  ONCHANGE="DATACHG()" style="width:75px">
								<option value=""></option>
								<%sql="select * from dbo.basicCode where func='groupid' order by sys_type"
								  set rst=conn.execute(sql)
								  while not rst.eof 
								%>
								<option value="<%=rst("sys_type")%>" <%if request("groupid")=rst("sys_type") then %> selected <%end if%> ><%=rst("sys_type")%><%=rst("sys_value")%></option>	
								<%rst.movenext
								wend
								set rst=nothing 				
								%>
							</select>
						</td>
						<td  nowrap align=right >班別:</td>
						<td>
							<select name=SHIFTA  ONCHANGE="DATACHG()" style="width:75px">
								<option value="" <%IF SHIFTA="" THEN %>SELECTED<%END IF%>></option>
								<option value="ALL" <%IF SHIFTA="ALL" THEN %>SELECTED<%END IF%>>常日班</option>
								<option value="A" <%IF SHIFTA="A" THEN %>SELECTED<%END IF%>>A班</option>
								<option value="B" <%IF SHIFTA="B" THEN %>SELECTED<%END IF%>>B班</option>				
							</select>
						</td>
					</tr>
					<tr>
						<td nowrap align=right>查詢<br><font class="txt8">K.Tra</font></td>
						<td colspan=2><input type="text" name="queryx"   >
						</td>
						<td>
							<INPUT TYPE=BUTTON NAME=SEND VALUE="查詢" ONCLICK="DATACHG()" class="btn btn-sm btn-outline-secondary">
						</td>
					</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td align="center">
				<table id="myTableGrid" width="98%">
					<Tr class="header">
						<TD  nowrap    align=center HEIGHT=20>單位</TD>
						<TD  nowrap   align=center>工號</TD>
						<TD  nowrap   align=center>姓名</TD>
						<TD  nowrap   align=center>到職日</TD>
					</TR>
					<input type=HIDDEN name=EMPID value="">
					<input type=HIDDEN name=EMPNAME value="">
					<input type=HIDDEN name=indat value="">
					<input type=HIDDEN name=gstr value="">
				<% 
				'x = 1
				DIM J,RowCount,IX,WK_COLOR 

				   'j=0
				   RowCount = RS.PageSize 
				   'IX = RowCount * (CurrentPage-1) + 1
				Do While Not RS.EOF and RowCount 

					'j=j+1
					Response.Write "<TR BGCOLOR=#FFFFFF class='txt9'>" & vbcrlf       
				%>
						<TD  HEIGHT=22  align=center> <A HREF="VBSCRIPT:GiveAns(<%=X%>)">
							 <%=rs("GSTR")%></A>      
						</TD>   
						<TD  HEIGHT=22 width=50 align=center>
							 <A HREF="VBSCRIPT:GiveAns(<%=X%>)">
							 <font color=blue><%=rs("empid")%></font></A>      
						</TD>   
						<TD  align=LEFT nowrap> 
							  <A HREF="VBSCRIPT:GiveAns(<%=X%>)">
							  <font color=blue><%=rs("empNam_CN")%><%=rs("empNam_VN")%></font> 
							  </A>
							  <input type=HIDDEN name=EMPID value="<%=rs("empid")%>">
							  <input type=HIDDEN name=EMPNAME value="<%=rs("empNam_CN")%>">
							  <input type=HIDDEN name=indat value="<%=rs("nindat")%>">
							  <input type=HIDDEN name=gstr value="<%=rs("gstr")%>">
						</TD>
						<td><%=rs("nindat")%></td>
					</Tr>
				<%     
				   RowCount = RowCount - 1
				   'IX = IX + 1   
				   rs.MoveNext
				   'X = X + 1
				loop

				%>
				</Table>
			</td>
		</tr>
		<tr>
			<td align="center">
				<table class="txt" cellpadding=3 cellspacing=3>
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
					<TR><TD ALIGN=CENTER>COUNT:<%=TotalRecords%>,page:<%=CurrentPage%>/<%=TotalPage%></TD></TR>
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
	<%=self%>.target="_self"
	<%=SELF%>.SUBMIT()
END FUNCTION  

Function GiveAns(i) 
	'alert"aaa="&<%=self%>.index.value
	if trim(<%=self%>.index.value)<>"" then 
		if trim(<%=self%>.nfs.value)<>"" then 
			nfs=trim(<%=self%>.nfs.value)
			Parent.Fore.<%=formName%>.EMPID(<%=index%>).value = <%=SELF%>.EMPID(i).Value 
		    Parent.Fore.<%=formName%>.empname(<%=index%>).value = <%=SELF%>.EMPNAME(i).Value   
		    Parent.Fore.<%=formName%>.<%=nfs%>(<%=index%>).focus()	
		    if <%=self%>.formName.value="yebbb04" then 	
		    	Parent.Fore.<%=formName%>.F_groupid(<%=index%>).value=<%=SELF%>.gstr(i).Value   
		    end if			
		end if 
	else
		'alert"bb="&<%=self%>.formName.value

		if <%=self%>.formName.value="yebbb03" then 
			Parent.Fore.<%=formName%>.EMPID.value = <%=SELF%>.EMPID(i).Value 
			Parent.Fore.<%=formName%>.indat.value = <%=SELF%>.indat(i).Value 
			'Parent.Fore.<%=formName%>.rp_dat.focus()
			Parent.Fore.<%=formName%>.action="yebbb03.new.asp"
			Parent.Fore.<%=formName%>.target="Fore"
			Parent.Fore.<%=formName%>.submit()	
		elseif <%=self%>.formName.value="YEIE0201" then 			
			Parent.Fore.<%=formName%>.f_empid.value = <%=SELF%>.EMPID(i).Value 
		    Parent.Fore.<%=formName%>.empname.value = <%=SELF%>.EMPNAME(i).Value
		else
			Parent.Fore.<%=formName%>.EMPID.value = <%=SELF%>.EMPID(i).Value 
		    Parent.Fore.<%=formName%>.empname.value = <%=SELF%>.EMPNAME(i).Value   
		end if  
	end if 
	Parent.best.cols = "100%,0%" 
End Function 

Function GoClose()
   Parent.best.cols = "100%,0%"	
End Function  
    
</SCRIPT>
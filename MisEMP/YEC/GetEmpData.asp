<%@LANGUAGE=VBSCRIPT codepage=65001%>
<!-- #include file="../ADOINC.inc" -->
<!--#include file="../GetSQLServerConnection.fun"  -->
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
formName=request("formName") 
index = request("index") 

index=request("index") 
'response.write index &"<BR>"

yymm=year(date())&right("00"&month(date()),2) 
if right("00"&month(date()),2) ="01" then 
	Lastym=year(date())-1&"12" 
else	
	lastym=year(date())&right("00"&month(date())-1,2) 
end if

if queryx="" and country="" and groupid="" and JOBID="" and SHIFTA="" then 
	sql="select * from  view_empfile where ( whsno = 'ALL' or country in ('TW', 'MA','VN' ) )  and empid<>'PELIN' order by len(whsno) desc , whsno desc, outdate, empid " 
else
	if queryx="" then  
		sql="select * from  view_empfile where 1=1 "&_
			"and ( isnull(outdate,'')='' or convert(char(6), outdat,112)>='"& lastym &"' ) and country like '"& country &"%' and groupid like '"& groupid &"%' "&_
			"AND ISNULL(JOB,'') like '"& JOBID &"%'  " 
			
		if 	SHIFTA<>"" then 
			sql=sql & " and ISNULL(SHIFT,'') = '"& SHIFTA &"' " 
		end if	
		sql =sql & "order by  len(whsno) desc , whsno, outdate, empid " 
	else
		sql="select * from  view_empfile where empid like '%"& queryx &"%' or  empnaM_cn like '%"& queryx &"%'  order by len(whsno) desc , whsno, outdate, empid    "
			 
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
<link rel="stylesheet" href="../Include/style.css" type="text/css">
<link rel="stylesheet" href="../Include/style2.css" type="text/css"> 

</head>
<body  topmargin="5" leftmargin="5"  marginwidth="0" marginheight="0" >
<form METHOD="POST" ACTION="<%=self%>.asp" name="<%=self%>">
<input TYPE="HIDDEN" NAME="CurrentPage" VALUE="<%=CurrentPage%>">
<input type=HIDDEN name=index value="<%=index%>">
<input type=HIDDEN name=formName value="<%=formName%>">
<table width=250 border=0 class=txt12  >
	<tr><td align=center ><b>員工資料查詢</b></td></tr>
</table>
<hr size=0	style='border: 1px dotted #999999;' align=left  >
<table width=300 border=0 class=txt >
	<tr>
		<td width=50 align=right>國籍:</td>
		<td width=80>
			<select name=country class=inputbox ONCHANGE="DATACHG()" >
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
		<td width=50 align=right >單位:</td>
		<td>
			<select name=groupid class=inputbox ONCHANGE="DATACHG()">
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
	</tr>
	<tr>
		<td width=50 align=right>職務:</td>
		<td width=80>
			<select name=JOBID class=inputbox ONCHANGE="DATACHG()" >
				<option value=""></option>
				<%sql="select * from dbo.basicCode where func='LEV' order by sys_type"
				  set rst=conn.execute(sql)
				  while not rst.eof 
				%>
				<option value="<%=rst("sys_type")%>" <%if request("JOBID")=rst("sys_type") then %> selected <%end if%>><%=rst("sys_type")%><%=LEFT(rst("sys_value"),5)%></option>	
				<%rst.movenext
				wend
				set rst=nothing 				
				%>
			</select>
		</td>
		<td width=50 align=right >班別:</td>
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
		<td align=right>查詢<br>K.Tra:</td>
		<td colspan=3><input name="queryx" class=inputbox  >
			<INPUT TYPE=BUTTON NAME=SEND VALUE="查詢" ONCLICK="DATACHG()" class=button>
		</td>
	</tr>
</table>	
<hr size=0	style='border: 1px dotted #999999;' align=left  >
<Table cellpadding=1 cellspacing=1 border=0 width=390 CLASS=TXT9 BGCOLOR="#CCCCFF">
	<Tr BGCOLOR="FFFFCC">
    	<TD align=center HEIGHT=20>廠別/單位</TD>
    	<TD align=center>工號</TD>
    	<TD align=center>姓名</TD>
    	<TD align=center>到職日</TD>
    	<TD align=center>離職日</TD>
		<input type=HIDDEN name=F_EMPID >
		<input type=HIDDEN name=F_EMPNAME >
		<input type=HIDDEN name=F_country >
		<input type=HIDDEN name=F_whsno >    	
	</TR> 
	<% 
	x = 1
	DIM J,RowCount,IX,WK_COLOR 
	
	   j=0
	   RowCount = RS.PageSize 
	   IX = RowCount * (CurrentPage-1) + 1
	Do While Not RS.EOF and RowCount 
		j=j+1
		Response.Write "<TR BGCOLOR=#FFFFFF>" & vbcrlf       
	%>
		<TD nowrap> <A HREF="VBSCRIPT:GiveAns(<%=X%>)"  style="{CURSOR: hand}"  >
			 <font color=<%if trim(rs("outdate"))="" then%>"blue"<%else%>"red"<%end if%>><%=rs("whsno")%>/<%=rs("GSTR")%></font></A>      
	    </TD>   
		<TD align=center>
			 <A HREF="VBSCRIPT:GiveAns(<%=X%>)"  style ="{CURSOR: hand}"  >
			 <font color="<%if  trim(rs("outdate"))="" then%>blue<%else%>red<%end if%>"><%=rs("empid")%></font></A>      
	    </TD>   
		<TD> 
	          <A HREF="VBSCRIPT:GiveAns(<%=X%>)"  style ="{CURSOR: hand}" >
	          <font color=<%if trim(rs("outdate"))="" then%>"blue"<%else%>"red"<%end if%>><%=rs("empNam_CN")%><%=rs("empNam_VN")%></font></A>
	          
	          <input type=HIDDEN name=F_EMPID value="<%=rs("empid")%>">
	          <input type=HIDDEN name=F_EMPNAME value="<%=rs("empNam_CN")%><%=rs("empNam_vN")%>">
	          <input type=HIDDEN name=F_country value="<%=rs("country")%>">
	          <input type=HIDDEN name=F_whsno value="<%=rs("whsno")%>">
		</TD>
		<td align=center><%=rs("nindat")%></td>
		<td align=center><%=rs("outdate")%></td>
	</Tr>
	<%     
	   RowCount = RowCount - 1
	   IX = IX + 1   
	   rs.MoveNext
	   X = X + 1
	loop 
	%>

</Table>
<table border="0" width=400>
	<tr>
	<TD align=left nowrap>
		<% If CurrentPage > 1 Then %>
		<INPUT	TYPE="submit" NAME="WKSCROLL" VALUE="FIRST" class=button  >
		<INPUT	TYPE="submit" NAME="WKSCROLL" VALUE="PRE" 	class=button  >
		<% Else %>
		<INPUT	TYPE="submit" NAME="WKSCROLL" VALUE="FIRST" disabled 	class=button>
		<INPUT	TYPE="submit" NAME="WKSCROLL" VALUE="PRE" disabled 	class=button>
		<% End If %>
		<% If CurrentPage < TotalPage Then %>
		<INPUT	TYPE="submit" NAME="WKSCROLL" VALUE="NEXT" 	class=button  >
		<INPUT	TYPE="submit" NAME="WKSCROLL" VALUE="END" 	class=button   >	
		<% Else %>
		<INPUT	TYPE="submit" NAME="WKSCROLL" VALUE="NEXT" disabled class=button>
		<INPUT	TYPE="submit" NAME="WKSCROLL" VALUE="END" disabled class=button>
		<% End If %>
		<INPUT	TYPE="button" NAME="WKSCROLL" VALUE="關閉CLOSE"  class=button onclick=GOClose()>
	</TD>
	</TR>	
</table> 
<TABLE WIDTH=300><TR CLASS=TXT9><TD ALIGN=CENTER>COUNT:<%=TotalRecords%>,page:<%=CurrentPage%>/<%=TotalPage%></TD></TR></TABLE>

</FORM> 
</body>

</html>

<SCRIPT LANGUAGE=VBSCRIPT>
FUNCTION DATACHG()
	<%=SELF%>.ACTION="<%=SELF%>.ASP"
	<%=SELF%>.SUBMIT()
END FUNCTION  

Function GiveAns(i)    	 
	Parent.Fore.<%=formName%>.EMPID(<%=index%>).value = <%=SELF%>.F_EMPID(i).Value 
    Parent.Fore.<%=formName%>.empname(<%=index%>).value = <%=SELF%>.F_EMPNAME(i).Value   
    Parent.Fore.<%=formName%>.whsno(<%=index%>).value = <%=SELF%>.F_whsno(i).Value   
    Parent.Fore.<%=formName%>.country(<%=index%>).value = <%=SELF%>.F_country(i).Value   
    Parent.Fore.<%=formName%>.bb(<%=index%>).focus()
	Parent.best.cols = "100%,0%"       
End Function 

Function GoClose()
   Parent.best.cols = "100%,0%"	
End Function  
    
</SCRIPT>
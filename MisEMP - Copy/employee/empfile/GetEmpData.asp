<%@LANGUAGE=VBSCRIPT codepage=65001%>
<!--#include file="../../ADOINC.inc"-->
<!--#include file="../../GetSQLServerConnection.fun"  -->
<%
self="Getempdata" 
Set conn = GetSQLServerConnection()	  
queryx = trim(request("queryx")) 
S_country = request("S_country") 
groupid = request("groupid") 
CurrentPage = REQUEST("CurrentPage") 
WKSCROLL = Request("WKSCROLL")     
JOBID=REQUEST("JOBID")
SHIFTA=REQUEST("SHIFTA") 

index=request("index") 
'response.write index &"<BR>"

pself=request("pself")
if pself="" then pself="XXX" 
ncols=request("ncols")

yymm=year(date())&right("00"&month(date()),2) 
if right("00"&month(date()),2) ="01" then 
	Lastym=year(date())-1&"12" 
else	
	lastym=year(date())&right("00"&month(date())-1,2) 
end if

if pself="empbe04" then 
	if queryx="" then  
		sql="select convert(char(10), indat, 111) indate, * from  view_empfile where country<>'VN' "&_
			"and ( isnull(outdate,'')='' or convert(char(6), outdat,112)>='"& lastym &"' ) and country in ('CN', 'TA') and groupid like '%"& groupid &"%' "&_
			"AND ISNULL(JOB,'') like '%"& JOBID &"%' and ISNULL(SHIFT,'') like '%"& SHIFTA &"%'  order by empid "
	else
		sql="select * from view_empfile where country in ('CN', 'TA') and empid like '%"& queryx &"%' or  empnaM_cn like '%"& queryx &"%'   order by empid  "		 
	end if 		 
else 
	if queryx="" then  
		sql="select convert(char(10), indat, 111) indate, * from  view_empfile where country<>'VN' "&_
			"and ( isnull(outdate,'')='' or convert(char(6), outdat,112)>='"& lastym &"' ) and country like '%"& S_country &"%' and groupid like '%"& groupid &"%' "&_
			"AND ISNULL(JOB,'') like '%"& JOBID &"%' and ISNULL(SHIFT,'') like '%"& SHIFTA &"%' order by empid   "
	else
		sql="select * from view_empfile where empid like '%"& queryx &"%' or  empnaM_cn like '%"& queryx &"%'    order by empid  "		 
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
<link rel="stylesheet" href="../../Include/style.css" type="text/css">
<link rel="stylesheet" href="../../Include/style2.css" type="text/css"> 
</head>
<body  topmargin="5" leftmargin="5"  marginwidth="0" marginheight="0" >
<form METHOD="POST" ACTION="<%=self%>.asp" name="<%=self%>">
<input TYPE="HIDDEN" NAME="CurrentPage" VALUE="<%=CurrentPage%>">
<input type=HIDDEN name=index value="<%=index%>">
<input type=HIDDEN name=EMPID value="">
<input type=HIDDEN name=EMPNAME value="">
<input type=HIDDEN name=pself value="<%=pself%>">
<table width=250 border=0 class=txt12  >
	<tr><td align=center ><b>員工資料查詢</b></td></tr>
</table>
<hr size=0	style='border: 1px dotted #999999;' align=left  width=300>
<table width=250 border=0 class=txt9 >
	<tr>
		<td width=50 align=right>國籍:</td>
		<td width=80>
			<select name=S_country class=inputbox ONCHANGE="DATACHG()" >
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
		<td width=50 align=right >單位:</td>
		<td>
			<select name=groupid class=inputbox ONCHANGE="DATACHG()">
				<option value=""></option>
				<%sql="select * from [yfynet].dbo.basicCode where func='groupid' order by sys_type"
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
		<td width=50 align=right>職務:</td>
		<td width=80>
			<select name=JOBID class=inputbox ONCHANGE="DATACHG()" >
				<option value=""></option>
				<%sql="select * from [yfynet].dbo.basicCode where func='LEV' order by sys_type"
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
		<td align=right>查詢:</td>
		<td colspan=3><input name="queryx" class=inputbox  >
			<INPUT TYPE=BUTTON NAME=SEND VALUE="查詢" ONCLICK="DATACHG()">
		</td>
	</tr>
</table>	
<hr size=0	style='border: 1px dotted #999999;' align=left  width=300 >
<Table cellpadding=1 cellspacing=1 border=0 width=300 CLASS=TXT9 BGCOLOR="#CCCCFF">
	<Tr BGCOLOR="FFFFCC">
    	<TD  width=50    align=center HEIGHT=20>單位</TD>
    	<TD  width=50   align=center>工號</TD>
    	<TD  width=190   align=center>姓名</TD>
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
		<TD  HEIGHT=22 width=50 align=center> <A HREF="VBSCRIPT:GiveAns(<%=X%>)"  style ="{CURSOR: hand}"  >
			 <%=rs("GSTR")%></A>      
	    </TD>   
		<TD  HEIGHT=22 width=50 align=center>
			 <A HREF="VBSCRIPT:GiveAns(<%=X%>)"  style ="{CURSOR: hand}"  >
			 <%=rs("empid")%></A>      
	    </TD>   
		<TD  align=LEFT > 
	          <A HREF="VBSCRIPT:GiveAns(<%=X%>)"  style ="{CURSOR: hand}"  >
	          <%=rs("empNam_CN")%><%=rs("empNam_VN")%> 
	          </A>
	          <input type=HIDDEN name=EMPID value="<%=rs("empid")%>">
	          <input type=HIDDEN name=EMPNAME value="<%=rs("empNam_CN")%>">
	          <input type=HIDDEN name=country value="<%=rs("country")%>">
	          <input type=HIDDEN name=indate value="<%=rs("indate")%>">
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
		<INPUT	TYPE="button" NAME="WKSCROLL" VALUE=" 關   閉"  class=button    onclick=GOClose()>
	</TD>
	</TR>	
</table> 
<TABLE WIDTH=300><TR CLASS=TXT9><TD ALIGN=CENTER>COUNT:<%=TotalRecords%>,頁次:第<%=CurrentPage%>頁,共<%=TotalPage%>頁</TD></TR></TABLE>

</FORM> 
</body>

</html>

<SCRIPT LANGUAGE=VBSCRIPT>
FUNCTION DATACHG()
	<%=SELF%>.ACTION="<%=SELF%>.ASP"
	<%=SELF%>.SUBMIT()
END FUNCTION  

Function GiveAns(i)  
	Parent.Fore.<%=pself%>.EMPID(<%=index%>).value = <%=SELF%>.EMPID(i).Value 
	Parent.Fore.<%=pself%>.empname(<%=index%>).value = <%=SELF%>.EMPNAME(i).Value 
    Parent.Fore.<%=pself%>.country(<%=index%>).value = <%=SELF%>.country(i).Value 
    pself="<%=pself%>"
    
    if pself="empbe03" then 
      	Parent.Fore.<%=pself%>.visano(<%=index%>).focus()
	elseif pself="empbe04" then 	
		Parent.Fore.<%=pself%>.indate(<%=index%>).value = <%=SELF%>.indate(i).Value 
		Parent.Fore.<%=pself%>.dat1(<%=index%>).focus()
    end if 
    Parent.best.cols = "100%,0%"
      
End Function 

Function GoClose()
   Parent.best.cols = "100%,0%"	
End Function  
    
</SCRIPT>
<%@LANGUAGE=VBSCRIPT codepage=65001%>
<!--#include file="../ADOINC.inc"-->
<!--#include file="../GetSQLServerConnection.fun"  -->
<%
self="getstudyPlan" 
Set conn = GetSQLServerConnection()	  
queryx = trim(request("yy")) 
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
 
if queryx="" then  
	sql="select * from studyPlan where isnull(status,'')<>'D' and left(ssno,4) like '"&queryx&"%' order by left(ssno,4) desc  , ssno  "
else
	sql="select * from studyPlan where isnull(status,'')<>'D' and yy='1' order by left(ssno,4) desc  , ssno "
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
<input type=HIDDEN name=pself value="<%=pself%>">
<input type=HIDDEN name=ncols value="<%=ncols%>">
<table width=250 border=0 class=txt12  >
	<tr><td align=center ><b>資料查詢</b></td></tr>
</table>
<hr size=0	style='border: 1px dotted #999999;' align=left  width=300>
<table width=250 border=0 class=txt >
	<tr>
		<td width=50 align=right>年度:</td>
		<td >
			<select name=yy class=inputbox ONCHANGE="DATACHG()" >
				<option value=""></option>
				<%for x = 1 to 10  				 
				  yystr = year(date())
				  yyvalue = yystr +X - 3 
				%>
				<option value="<%=yyvalue%>" <%if  yyvalue=year(date()) then %> selected <%end if%>><%=yyvalue%></option>	
				<%next
				%>
			</select>
		</td>		 
	</tr> 
</table>	
<hr size=0	style='border: 1px dotted #999999;' align=left  width=300 > 

<Table cellpadding=1 cellspacing=1 border=0 width=300 CLASS=TXT BGCOLOR="#CCCCFF">
	<Tr BGCOLOR="FFFFCC">
    	<TD  width=50    align=center HEIGHT=20>年度</TD>
    	<TD  width=80   align=center>編號</TD>
    	<TD  width=160   align=center>課程名稱</TD>
	</TR> 
	<input type=HIDDEN name=EMPID value="">
	<input type=HIDDEN name=EMPNAME value="">
	<input type=HIDDEN name=country value="">	 
	<input type=HIDDEN name=NW value="">
	<input type=HIDDEN name=amt value="0">
	<input type=HIDDEN name=dm value="">
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
			 <%=rs("yy")%></A>      
	    </TD>   
		<TD  HEIGHT=22 width=50 align=center>
			 <A HREF="VBSCRIPT:GiveAns(<%=X%>)"  style ="{CURSOR: hand}"  >
			 <%=rs("ssno")%></A>      
	    </TD>   
		<TD  align=LEFT > 
	          <A HREF="VBSCRIPT:GiveAns(<%=X%>)"  style ="{CURSOR: hand}"  >
	          <%=rs("studyName")%>
	          </A>
	          <input type=HIDDEN name=EMPID value="<%=rs("yy")%>">
	          <input type=HIDDEN name=NW value="<%=rs("nw")%>">
	          <input type=HIDDEN name=EMPNAME value="<%=rs("ssno")%>">
	          <input type=HIDDEN name=country value="<%=rs("studyName")%>">	          
	          <input type=HIDDEN name=amt value="<%=rs("amt")%>">
	          <input type=HIDDEN name=dm value="<%=rs("dm")%>">	          
		</TD>
	</Tr>
<%     
   RowCount = RowCount - 1
   IX = IX + 1   
   rs.MoveNext
   X = X + 1
loop
rs.close
set rs=nothing 

conn.close 
set conn=nothing

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
    Parent.Fore.<%=pself%>.studyName.value = <%=SELF%>.country(i).Value 
    Parent.Fore.<%=pself%>.ssno.value = <%=SELF%>.EMPNAME(i).Value 
    Parent.Fore.<%=pself%>.nw.value = <%=SELF%>.nw(i).Value 
    Parent.Fore.<%=pself%>.amt.value = <%=SELF%>.amt(i).Value 
    Parent.Fore.<%=pself%>.dm.value = <%=SELF%>.dm(i).Value 
    if <%=SELF%>.nw(i).Value="N" then
    	Parent.Fore.<%=pself%>.nw(0).selected=true
    elseif <%=SELF%>.nw(i).Value="W" then
    	Parent.Fore.<%=pself%>.nw(1).selected=true
    end if 	
    Parent.Fore.<%=pself%>.<%=ncols%>.focus()	
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
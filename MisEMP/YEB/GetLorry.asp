<%@language=vbscript codepage=65001%>
<!--#include file="../GetSQLServerConnection.fun"-->
<!--#include file="../ADOINC.inc" -->

<%

Response.Expires = 0
Response.Buffer = true 
response.cachecontrol="no-cache"  
ThisPage="GiveHelp"

TargetName= request("TargetName")
if trim(TargetName)="" then TargetName="YEBE0103"
   
Set conn = GetSQLServerConnection()	  
Set rs = Server.CreateObject("ADODB.Recordset")

QueryX = trim(request("QueryX"))

sql="select * from YSBMLRIF where lorry<>'XX' "&_
	"and( driverid like '%"&QueryX&"%' or LorryDriverIC like '%"&QueryX&"%' ) order by lorry" 
rs.Open sql,conn,3,3

'response.write sql 
'response.end
formname=request("formname") 
RS.PageSize = 15          
TotalRecords = RS.RecordCount
TotalPage = RS.PageCount   
CurrentPage = int(Request("CurrentPage"))
WKSCROLL = request("WKSCROLL") 
Select Case WKSCROLL
		
		Case ""
		    CurrentPage = 1
		Case "FIRST"
		    CurrentPage = 1
		Case "BACK"
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
if not rs.eof then 
	RS.AbsolutePage = CurrentPage  
end if 	

selfname = request("selfname")

%>
<HTML>
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="refresh">
<link rel="stylesheet" href="../Include/style.css" type="text/css">
<link rel="stylesheet" href="../Include/style2.css" type="text/css">    
</HEAD>

<BODY   topmargin="10" leftmargin="10"  marginwidth="0" marginheight="0">
<form METHOD="POST" ACTION="GetLorry.asp" name="<%=ThisPage%>">
<input TYPE="HIDDEN" NAME="TotalPage" VALUE="<%=TotalPage%>">
<input TYPE="HIDDEN" NAME="CurrentPage" VALUE="<%=CurrentPage%>">	
<input TYPE="HIDDEN" NAME="TargetName" VALUE="<%=TargetName%>">  	
<table width=300 ><tr><td align=center>
	<font class=txt12><b>Tu lieu Xe Kiem Tra </b></font><BR> 
</td></tr></table>	   
<Table width=300 border=0>
<TR>
	<td align=right width=60> 關鍵字: </td>
	<td align=left><input type=text name=QueryX class="inputbox" value="<%=Request("QueryX")%>"></td>
</TR>
</Table>  
<Table cellpadding=1 cellspacing=1 border=0 width=250 bgcolor="black" class=txt>
	<Tr bgcolor=Khaki height=22>
	    <TD nowrap width=30 align=center><font class= txt>代碼</font></TD>   
	    <TD nowrap width=90 align=center><font class= txt>車號</font></TD>    
	    <TD nowrap width=130 align=center><font class= txt>車行</font></TD>
	</TR>  
	<input type=hidden name=Cust value="">
	<input type=hidden name=chlnm value=""> 
	<input type=hidden name=capacity value="">
	<input type=hidden name=basicfee value="">
	<input type=hidden name=transportfeeby value="">
	<input type=hidden name=lorrydrivername value=""> 
	<input type=hidden name=xhid value="">
	<% 
	if not rs.eof then 
	x = 1
	DIM J,RowCount,IX,WK_COLOR
	j=0
	RowCount = RS.PageSize
	 
	IX = RowCount * (CurrentPage-1) + 1
	Do While Not RS.EOF and RowCount 
	    if (j mod 2)=1 then
			WK_COLOR="LIGHTYELLOW"
		Else
			WK_COLOR="WHITE"
		End if
		j=j+1%>    
    <TR height=22 bgcolor="<%=WK_COLOR%>">
    <TD align=center  >
    	<A  style="{CURSOR: hand}" onclick="GiveAns(<%=X%>)" ><font class=txt color="blue"><%=rs(0)%></font></A>        
    </TD>    
    <TD align=center >
    	<A style ="{CURSOR: hand}" href onclick="GiveAns(<%=X%>)" >
        <font  color="blue"><%=rs(1)%></font></A>        
    </TD>
 	<TD align=left  ><%=rs("LorryDriverIC")%></A>        
    </TD>
    <input type=hidden name=Cust value="<%=rs(0)%>">
    <input type=hidden name=chlnm value="<%=rs(1)%>">  
    <input type=hidden name=Capacity value="<%=rs("Capacity")%>">
    <input type=hidden name=Basicfee value="<%=rs("Basicfee")%>">  
    <input type=hidden name=Transportfeeby value="<%=rs("Transportfeeby")%>">   
    <input type=hidden name=LorryDriverName value="<%=rs("LorryDriverName")%>">   
    <input type=hidden name=xhid value="<%=rs("xhid")%>">
  	</Tr>
	<%     
	    RowCount = RowCount - 1
	    IX = IX + 1   
	    rs.MoveNext
	    X = X + 1
	loop  
	end if 
	
	conn.close 
	set conn=nothing
	%>
</Table>
<BR>
<TABLE WIDTH=300><TR><TD align=center>
		<font class="txt8" >Count: <%=TotalRecords%>&nbsp; Page: <% = CurrentPage & "/" & TotalPage %></font>
</TD></TR></table>

<table border="0" width=300>
<tr>
	<TD align=left nowrap>
	<% If CurrentPage > 1 Then %>
		<INPUT	TYPE="submit" NAME="WKSCROLL" VALUE="FIRST" class=button  >
		<INPUT	TYPE="submit" NAME="WKSCROLL" VALUE="BACK" class=button  >
	<% Else %>
		<INPUT	TYPE="submit" NAME="WKSCROLL" VALUE="FIRST" disabled 	class=button>
		<INPUT	TYPE="submit" NAME="WKSCROLL" VALUE="BACK" disabled 	class=button>
	<% End If %>
	
	<% If CurrentPage < TotalPage Then %>
		<INPUT	TYPE="submit" NAME="WKSCROLL" VALUE="NEXT" class=button >
		<INPUT	TYPE="submit" NAME="WKSCROLL" VALUE="END" class=button  >	
	<% Else %>
		<INPUT	TYPE="submit" NAME="WKSCROLL" VALUE="NEXT" disabled class=button>
		<INPUT	TYPE="submit" NAME="WKSCROLL" VALUE="END" disabled class=button>	
	<% End If %>
		<INPUT	TYPE="button" NAME="WKSCROLL" VALUE="關閉(CLOSE)"  class=button	  onclick=GOClose()>
	</TD> 
</tr>
</table>
</Form>
</BODY>
</HTML>
<Script Language="vbscript">

Function GiveAns(i)
	tgname="<%=TargetName%>"
	 if right(tgname,1)="B" then 
	 	Parent.Fore.<%=TargetName%>.lorry.value = <%=ThisPage%>.Cust(i).Value          
     	Parent.Fore.<%=TargetName%>.soxe.value = <%=ThisPage%>.chlnm(i).Value	 
     	'Parent.Fore.<%=TargetName%>.xhid.value = <%=ThisPage%>.xhid(i).Value	 
	 	Parent.Fore.<%=TargetName%>.totalpage.value="0"
	 	Parent.Fore.<%=TargetName%>.action="<%=TargetName%>.accdata.asp"
	 	Parent.Fore.<%=TargetName%>.submit()
	 else
     	Parent.Fore.<%=TargetName%>.lorry.value = <%=ThisPage%>.Cust(i).Value          
     	Parent.Fore.<%=TargetName%>.soxe.value = <%=ThisPage%>.chlnm(i).Value	 
     end if 	

     Parent.best.cols = "100%,0%"
End Function

function m(index)
	document.<%=ThisPage%>.WKSCROLL(index).style.backgroundcolor="lightyellow"
	document.<%=ThisPage%>.WKSCROLL(index).style.color="red"
end function

function n(index)
	document.<%=ThisPage%>.WKSCROLL(index).style.backgroundcolor="khaki"
	document.<%=ThisPage%>.WKSCROLL(index).style.color="black"
end function
	
Function QueryX_onchange       
    <%=ThisPage%>.submit
End Function
   
Function GoClose()
    Parent.BEST.cols = "100%,0,0"	
End Function  

</Script>

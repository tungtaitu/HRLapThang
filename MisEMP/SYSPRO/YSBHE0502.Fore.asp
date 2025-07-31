<%@language=vbscript codepage=950%>
<!---------  #include file="../GetSQLServerConnection.fun"  --------> 

<%
SELF = "YSBHE0502"
gTotalPage = 10
PageRec = 20    'number of records per page
TableRec = 10    'number of fields per record
Query = request("Query") 
Set conn = GetSQLServerConnection() 
if trim(Query) = "" then
    if request("TotalPage") = "" or request("TotalPage") = "0" then 
       CurrentPage = 1 	   	   
	   source = "select  * from "&_
				"( select * from sysuser   ) a  "&_
				"left join  ( select sys_type, sys_value  from BasicCode  where func = 'Grp'  ) b on b.sys_type = a.rights  " &_
				"order by muser" 
	   Set rs = Server.CreateObject("ADODB.Recordset")
	   rs.Open Source, conn, 3
	   rs.PageSize = PageRec 
	   RecordInDB = rs.RecordCount 
	   TotalPage = rs.PageCount 	
	   'Set conn = nothing 	
	   Redim tmpRec(gTotalPage, PageRec, TableRec)   'Array	
	   for i = 1 to TotalPage 
  		   for j = 1 to PageRec
			   if not rs.EOF then 				  
					  tmpRec(i, j, 0) = "no"
					  tmpRec(i, j, 1) = rs("muser")
					  tmpRec(i, j, 2) = rs("username")
					  tmpRec(i, j, 3) = rs("rights")
					  tmpRec(i, j, 4) = rs("sys_value")
					  tmpRec(i, j, 5) = rs("password")
					  tmpRec(i, j, 6) = rs("password")
					  tmpRec(i, j, 7) = rs("group_id")
					  tmpRec(i, j, 8) = rs("whsno")				  
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
	   Session("YSBHE0502") = tmpRec	
   else
	   TotalPage = cint(request("TotalPage"))
	   StoreToSession()
	   CurrentPage = cint(request("CurrentPage"))
	   tmpRec = Session("YSBHE0502")
	
	   Select case request("send") 
 		      Case "第一頁"
			       CurrentPage = 1			
		      Case "上一頁"
			       if cint(CurrentPage) <> 1 then 
				      CurrentPage = cint(CurrentPage) - 1				
			       end if
		      Case "下一頁"
			       if cint(CurrentPage) <= cint(gTotalPage) then 
				      CurrentPage = CurrentPage + 1 
			       end if			
		      Case "最末頁"
			       CurrentPage = cint(TotalPage)
		      Case Else 
			       CurrentPage = 1	
	   end Select 	
   end if
else
   CurrentPage = 1
   Set conn = GetSQLServerConnection()
   source = "select  * from "&_
				"( select * from sysuser   ) a  "&_
				"left join  ( SELECT tblcd, tbldesc from YZZMCODE where tblid = 'Grp'  ) b on b.tblcd = a.grp  " &_
				"where muser like '%"& trim(Query) &"%' or username  like '%"& trim(Query) &"%' order by muser "
   Set rs = Server.CreateObject("ADODB.Recordset")
   rs.Open Source, conn, 3
   rs.PageSize = PageRec 
   RecordInDB = rs.RecordCount 
   TotalPage = rs.PageCount 	
   'Set conn = nothing 	
   Redim tmpRec(gTotalPage, PageRec, TableRec)   'Array	
   for i = 1 to TotalPage 
	   for j = 1 to PageRec
		   if not rs.EOF then 
			  for k=1 to TableRec-1
				   tmpRec(i, j, 0) = "no"
				   tmpRec(i, j, 1) = rs("muser")
				   tmpRec(i, j, 2) = rs("username")
				   tmpRec(i, j, 3) = rs("Grp")
				   tmpRec(i, j, 4) = rs("tbldesc")
				   tmpRec(i, j, 5) = rs("password")
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
   Session("YSBHE0502") = tmpRec	
end if   
%>

<html>
<head>
<meta HTTP-EQUIV="Content-Type" content="text/html; charset=BIG5">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">	
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
<title>程式功能維護</title></head>
<body   topmargin=5 > 
<form method="post" action="<%=SELF%>.Fore.asp" name="<%=SELF%>"> 
<table width="460" border="0" cellspacing="0" cellpadding="0">
  <tr>
   	<TD ><img border="0" src="../image/icon.gif" align="absmiddle">
   	<%=session("pgname")%></TD>		 
  </tr>
</table> 	 
<hr size=0	style='border: 1px dotted #999999;' align=left width=550> 

<table width="550" border="0" cellpadding="0" cellspacing="0">
   <tr>
      <td  align=center class=txt rowspan="2">&nbsp;</td>
      <td>    
         查詢關鍵字: 
         <input type="text" name="Query" size="50">
         <INPUT TYPE="button" name=send VALUE="查 詢" class=button onclick="Q_Data()">
         
      </td>
   </tr>    
   <tr>
   		<td colspan=2><hr size=0 style='border: 1px dotted #999999;' align=left  > </td>
   </tr> 
</table>

<INPUT TYPE=hidden NAME=TotalPage VALUE="<%=TotalPage%>">
<INPUT TYPE=hidden NAME=CurrentPage VALUE="<%=CurrentPage%>">
<INPUT TYPE=hidden NAME=PageRec VALUE="<%=PageRec%>">
<INPUT TYPE=hidden NAME=RecordInDB VALUE="<%=RecordInDB%>">
<INPUT TYPE=hidden NAME=gTotalPage VALUE="<%=gTotalPage%>">
<table width=600><tr><td align=center>
    <TABLE WIDTH=500 BORDER=0 align=center cellpadding=1 cellspacing=1>
      <tr >         
        <td width=30 align=center class=txt bgcolor=#B4C5DA>DEL</td>
        <td width=80 align=center class=txt bgcolor=#B4C5DA>User</td>
        <td width=80 align=center class=txt bgcolor=#B4C5DA>userName</td>
        <td width=60 align=center class=txt bgcolor=#B4C5DA>pwd</td>
        <td width=60 align=center class=txt bgcolor=#B4C5DA>groupid</td>
        <td width=60 align=center class=txt bgcolor=#B4C5DA>whsno</td>
        <td width=130 align=center class=txt bgcolor=#B4C5DA>userGroup</td>
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
		'Response.Write CurrentRow &"<BR>"
      %>       
		<TR bgcolor="<%=wk_color%>"> 	  
	    <td>
	    <%if trim(tmpRec(CurrentPage, CurrentRow, 1))<>"" then %>		    	
			<%if tmpRec(CurrentPage, CurrentRow, 0) = "del" then%>
				<input type=checkbox name=func value=del onclick="del(<%=CurrentRow - 1%>)" checked>
				<input type=hidden name=op value=del>
			<%else%>
				<input type=checkbox name=func value=no onclick="del(<%=CurrentRow - 1%>)" >
				<input type=hidden name=op value=no>
			<%end if%>   			
		<%else%>     
			<input type=hidden name=func>
			<input type=hidden name=op>   
		<%end if%>
		</td>
        <TD align=center> 
			<%if trim(tmpRec(CurrentPage, CurrentRow, 1))<>"" then %>		
			<input size=8   class=readonly readonly name="user" value="<%=Ucase(trim(tmpRec(CurrentPage, CurrentRow, 1)))%>" >
			<%else%>
			<input type=hidden name=user>
			<%end if%>
        </TD>
        
         <TD align=center> 			
			<%if trim(tmpRec(CurrentPage, CurrentRow, 1))<>"" then %>		
			<input size=10   class=readonly name="username" value="<%=Ucase(trim(tmpRec(CurrentPage, CurrentRow, 2)))%>" 
			onchange="tblcd_change(<%=currentrow-1%>)"
			 >
			<%else%>
			<input type=hidden name=username>
			<%end if%>
        </TD> 
         <TD align=center> 			
			<%if trim(tmpRec(CurrentPage, CurrentRow, 1))<>"" then %>		
			<input size=10   class=readonly name="pwd" value="<%=Ucase(trim(tmpRec(CurrentPage, CurrentRow, 5)))%>" 
			onchange="tblcd_change(<%=currentrow-1%>)"
			 >
			<%else%>
			<input type=hidden name=pwd>
			<%end if%>
        </TD>
        <TD align=center> 
			<%if trim(tmpRec(CurrentPage, CurrentRow, 1))<>"" then %>		
			<select name="groupid" class=inputbox onchange="tblcd_change(<%=currentrow-1%>)">
			<% sql="select sys_type, sys_value  from BasicCode  where func = 'groupid' " 
			   set rst=conn.execute(sql) 
			   while not rst.eof  			
			%>			
				<option value="<%=rst("sys_type")%>" <%if Ucase(trim(tmpRec(CurrentPage, CurrentRow, 7)))=rst("sys_type") then %> selected<%end if %>>
					<%=rst("sys_type")%>-<%=rst("sys_value")%>
				</option>
			<%rst.movenext
			wend 
			%>	
			</select>	
			<%else%>
			<input type=hidden name=usergroup>
			<%end if%>		
        </TD>   
	    <TD align=center> 
			<%if trim(tmpRec(CurrentPage, CurrentRow, 1))<>"" then %>		
			<select name="whsno" class=inputbox onchange="tblcd_change(<%=currentrow-1%>)">
			<% sql="select sys_type, sys_value  from BasicCode  where func = 'whsno' " 
			   set rst=conn.execute(sql) 
			   while not rst.eof  			
			%>			
				<option value="<%=rst("sys_type")%>" <%if Ucase(trim(tmpRec(CurrentPage, CurrentRow, 8)))=rst("sys_type") then %> selected<%end if %>>
					<%=rst("sys_type")%>-<%=rst("sys_value")%>
				</option>
			<%rst.movenext
			wend 
			%>	
			</select>	
			<%else%>
			<input type=hidden name=usergroup>
			<%end if%>		
        </TD>   
        <TD align=center> 
			<%if trim(tmpRec(CurrentPage, CurrentRow, 1))<>"" then %>		
			<select name="usergroup" class=inputbox onchange="tblcd_change(<%=currentrow-1%>)">
			<% sql="select sys_type, sys_value  from BasicCode  where func = 'Grp' " 
			   set rst=conn.execute(sql) 
			   while not rst.eof  			
			%>			
				<option value="<%=rst("sys_type")%>" <%if Ucase(trim(tmpRec(CurrentPage, CurrentRow, 3)))=rst("sys_type") then %> selected<%end if %>>
					<%=rst("sys_type")%>-<%=rst("sys_value")%>
				</option>
			<%rst.movenext
			wend 
			%>	
			</select>	
			<%else%>
			<input type=hidden name=usergroup>
			<%end if%>		
        </TD>      
      <%next%> 
    </TABLE> 
    
<input type=hidden name=func>
<input type=hidden name=op>
<input type=hidden name=user>
<input type=hidden name=username>    
<input type=hidden name=usergroup>   
<input type=hidden name=pwd>      
<CENTER>
第<INPUT size=8 readonly class=readonlyN name=ta style='BORDER-BOTTOM-STYLE: none; BORDER-LEFT: medium none;BORDER-RIGHT-STYLE: none; BORDER-TOP-STYLE: none' value=<%=CurrentPage%>>頁
</CENTER>

<br>


<TABLE border=0 width=500 align="center">
<tr>
    <td align="left" width="50%">
<% If CurrentPage > 1 Then %>
	<input type="submit" name="send" value="第一頁" class=button>
	<input type="submit" name="send" value="上一頁" class=button>
<% Else %>
	<input type="submit" name="send" value="第一頁" disabled class=button>
	<input type="submit" name="send" value="上一頁" disabled class=button>
<% End If %>

<% If cint(CurrentPage) < cint(TotalPage) Then %>
	<input type="submit" name="send" value="下一頁" class=button>
	<input type="submit" name="send" value="最末頁" class=button>
<% Else %>      
	<input type="submit" name="send" value="下一頁" disabled class=button>
	<input type="submit" name="send" value="最末頁" disabled class=button>	
<% End If %>
	</td>
    <td align="right" width="50%">
	<INPUT TYPE="button" name=send VALUE="確    認" class=button  onClick="Go()" <%=Umode%>>
	<INPUT TYPE="button" name=send VALUE="取    消" class=button  onClick="Clear()">
	</TD>
</TR>
</TABLE>
</td></tr></table>
</form>
</body>
</html>


<%
Sub StoreToSession()
	Dim CurrentRow
	tmpRec = Session("YSBHE0502")
	for CurrentRow = 1 to PageRec
		tmpRec(CurrentPage, CurrentRow, 0) = request("op")(CurrentRow)		
		tmpRec(CurrentPage, CurrentRow, 1) = request("user")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 2) = request("username")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 3) = request("usergroup")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 5) = request("pwd")(CurrentRow)
	next 
	Session("YSBHE0502") = tmpRec
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

function del(index)
	 if <%=SELF%>.func(index).checked=true  then 
		<%=SELF%>.op(index).value = "del"		
		open "<%=SELF%>.Back.asp?CurrentPage=" & <%=CurrentPage%> & _
		     "&index=" & index & "&func=del", "Back" 
		     'parent.best.cols="70%,30%"   	
	 else	
		<%=SELF%>.op(index).value = "no"		
		open "<%=SELF%>.Back.asp?CurrentPage=" & <%=CurrentPage%> & _
		     "&index=" & index & "&func=no", "Back"	          
		     'parent.best.cols="70%,30%"   
	 end if		
end function


function tblcd_change(index)	
	tbldesc_str = Ucase(trim(<%=SELF%>.usergroup(index).value))	 
	name_str = Ucase(trim(<%=SELF%>.username(index).value))	
	<%=SELF%>.username(index).value = Ucase(trim(<%=SELF%>.username(index).value))	 
	pwd_str = Ucase(trim(<%=SELF%>.pwd(index).value))	
	open "<%=SELF%>.Back.asp?CurrentPage=" & <%=CurrentPage%> & _
         "&tbldesc=" & tbldesc_str & _
         "&username=" & name_str & _
         "&pwd=" & pwd_str & _
         "&index=" & index & "&func=datachg", "Back" 
     ' parent.best.cols="70%,30%"   
end function

//-->
</script>

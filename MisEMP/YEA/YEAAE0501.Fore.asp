<%@language=vbscript codepage=65001%>
<!---------  #include file="../GetSQLServerConnection.fun"  --------> 
<!--#include file="../include/sideinfo.inc"-->

<%

SELF = "YEAAE0501"
gTotalPage = 1
PageRec = 20    'number of records per page
TableRec = 15    'number of fields per record
Query = request("Query") 
Set conn = GetSQLServerConnection() 
sortstr=  request("sortstr") 
views=  request("views")  


if sortstr ="" then sortstr="muser" 

    if request("TotalPage") = "" or request("TotalPage") = "0" then 
      CurrentPage = 1 	   	   
			source = "select  * from "&_
						   "( select * from sysuser where 1=1    "
			source = source & " and case when  '"&views&"' ='*' then '*' else  isnull(status,'') end = '"& views &"' "
			source = source & " ) a  "&_
							 "left join  ( select sys_type, sys_value  from BasicCode  where func = 'Grp'  ) b on b.sys_type = a.rights  "&_
							 "where 1=1  " 
			if Query<>"" then 
				source=source&" and  ( charindex( '"&Query&"',a.muser )>0  or charindex( '"&Query&"' ,a.username) >0 or a.rights='"&Query&"'  )  "
			end if 	 
			if session("netuser")="PELIN" then  
			else 
				if session("rights")<="2"  or  session("rights")="A" then 
					source=source&" and muser<>'PELIN'   " 				
				else 
					source=source&" and sys_type>='"& session("righs") &"' "
				end if 
			end if 
			source = source & "order by  case when isnull(status,'')='D' then 'z' else '' end , "  & sortstr  
		 
	   Set rs = Server.CreateObject("ADODB.Recordset")
	   rs.Open Source, conn, 3, 3
	   if not rs.eof then 
		   PageRec=rs.RecordCount
		   rs.PageSize = PageRec 
		   RecordInDB = rs.RecordCount 
		   TotalPage = rs.PageCount 	
		   gTotalPage = totalpage
	   end if  		 
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
					  tmpRec(i, j, 9) = rs("job")	
						tmpRec(i, j, 10) = rs("status")	
						tmpRec(i, j, 11) = rs("empid")	
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
	   'StoreToSession()
	   CurrentPage = cint(request("CurrentPage"))
	   tmpRec = Session("YSBHE0502") 
	   Select case request("send") 
 		      Case "FIRST"
			       CurrentPage = 1			
		      Case "BACK"
			       if cint(CurrentPage) <> 1 then 
				      CurrentPage = cint(CurrentPage) - 1				
			       end if
		      Case "NEXT"
			       if cint(CurrentPage) <= cint(gTotalPage) then 
				      CurrentPage = CurrentPage + 1 
			       end if			
		      Case "END"
			       CurrentPage = cint(TotalPage)
		      Case Else 
			       CurrentPage = 1	
	   end Select 	
   end if
 
%>

<html>
<head>
<meta HTTP-EQUIV="Content-Type" content="text/html; charset=utf-8">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">	
<link rel="stylesheet" href="../Include/style.css" type="text/css">
<link rel="stylesheet" href="../Include/style2.css" type="text/css">

<SCRIPT LANGUAGE=javascript>

function m(index){
   <%=SELF%>.send[index].style.backgroundcolor="lightyellow";
   <%=SELF%>.send[index].style.color="red";
}

function n(index){
   <%=SELF%>.send[index].style.backgroundcolor="khaki";
   <%=SELF%>.send[index].style.color="black";
}

function Q_Data(){	
   <%=self%>.TotalPage.value = "";
   <%=self%>.submit();   
}

//'-----------------enter to next field
function enterto(){
	if(window.event.keyCode == 13) window.event.keyCode =9;
}

</SCRIPT>

</head>
<body  onkeydown="enterto()"> 
<form method="post" action="<%=SELF%>.Fore.asp" name="<%=SELF%>"> 
<INPUT TYPE=hidden NAME=TotalPage VALUE="<%=TotalPage%>">
<INPUT TYPE=hidden NAME=CurrentPage VALUE="<%=CurrentPage%>">
<INPUT TYPE=hidden NAME=PageRec VALUE="<%=PageRec%>">
<INPUT TYPE=hidden NAME=RecordInDB VALUE="<%=RecordInDB%>">
<INPUT TYPE=hidden NAME=gTotalPage VALUE="<%=gTotalPage%>">
	<table border=0 style="height:70px"><tr><td>&nbsp;</td></tr></table>
	
	<table width="100%" BORDER=0 align=center cellpadding=0 cellspacing=0 >
		<tr>
			<td>
				<table class="txt" BORDER=0 cellpadding=3 cellspacing=3>
					<tr>								  
						<td>查詢關鍵字:</td>
						<td><input type="text" name="Query" size="30"  value="<%=query%>"  style="width:200px;vertical-align:middle" ></td>
						<td width="100px" align="right">排序(sap xep):</td>
						<td>
							<select  name="sortstr" style="vertical-align:middle;width:100px" >
								<option value="muser" <%if request("sortstr")="muser" then%>selected<%end if%> >User</option>
								<option value="rights,muser" <%if request("sortstr")="rights,muser" then%>selected<%end if%> >userGroup</option>
								<option value="group_id,muser" <%if request("sortstr")="group_id,muser" then%>selected<%end if%> >groupid</option>
							</select>
						</td>
						<td width="100px" align="center">
							<INPUT TYPE="button" name=send VALUE="查 詢" class="btn btn-sm btn-outline-secondary" onclick="Q_Data()" style="width:60px;vertical-align:middle" >         
						</td>
					</tr>    
					<tr>
						<td>View(顯示): </td>
						<td colspan=4>
							<input type="radio" name="views" value=""  <%if views="" then  %>checked<%end if%>>Normal(正常使用)&nbsp;&nbsp;&nbsp;
							<input type="radio" name="views" value="D" <%if views="D" then  %>checked<%end if%>>Cancel(已刪除)&nbsp;&nbsp;&nbsp;
							<input type="radio" name="views" value="*" <%if views="*" then  %>checked<%end if%>>ALL(全部)
						</td>
					</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td align="center">
				<table id="myTableGrid" width="98%">
					<tr class="header">
						<td style="width:30px">STT</td>
						<td style="width:30px">DEL</td>
						<td style="width:50px">啟用<br>undo</td>
						<td >User</td>
						<td >userName</td>        
						<td >so the</td>
						<td >groupid</td>
						<td >whsno</td>
						<td >rights</td>
						<%if session("netuser")="EPLIN" or session("rights")<="2" then %>
						<td >pwd</td>
						<%end if %>        
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
					if  tmpRec(CurrentPage, CurrentRow, 10) ="D" then 
						wk_color =  "#cccccc"
					end if 
				  %>       
					<TR bgcolor="<%=wk_color%>"> 	  
						<td align="center"><%=CurrentRow%></td>
						<td align="center" valign="center">	  
							<%if tmpRec(CurrentPage, CurrentRow, 10) = "D" then%>
							<input type=hidden name=func>			
							<%else%>							
							<input type=checkbox name=func value=del onclick="del(<%=CurrentRow - 1%>)"  >
							<%end if%> 					
							<input type="hidden" name="op" value="">
							<input type="hidden" name="stt" value="<%=CurrentRow%>">
						</td>
						<td>
							<%if tmpRec(CurrentPage, CurrentRow, 10) = "D" then%>
								<input type=checkbox name="opundo" value=del onclick="del2(<%=CurrentRow - 1%>)"  >
							<%else%>
								<input type="hidden"  name="opundo" value=del  >
							<%end if %>
						</td>
						<TD>
							<input type="text" style="width:98%" readonly name="user" value="<%=Ucase(trim(tmpRec(CurrentPage, CurrentRow, 1)))%>" >
						</TD>
						<TD>
							<input type="text" style="width:98%"  name="username" value="<%=Ucase(trim(tmpRec(CurrentPage, CurrentRow, 2)))%>" onchange="tblcd_change(<%=currentrow-1%>)" >
						</TD>
						<TD> 			
							<input type="text" style="width:98%"  name="empid" value="<%=Ucase(trim(tmpRec(CurrentPage, CurrentRow, 11)))%>" onchange="tblcd_change(<%=currentrow-1%>)" >
						</TD>
						<TD>
							<select name="groupid"  style="width:98%" onchange="tblcd_change(<%=currentrow-1%>)">
							<% sql="select sys_type, sys_value  from BasicCode  where func = 'groupid' " 
							   set rst=conn.execute(sql) 
							   while not rst.eof  			
							%>			
								<option value="<%=rst("sys_type")%>" <%if Ucase(trim(tmpRec(CurrentPage, CurrentRow, 7)))=rst("sys_type") then %> selected<%end if %>>
									<%=rst("sys_type")%>-<%=rst("sys_value")%>
								</option>
							<%rst.movenext
							wend 
							rst.close : set rst=nothing
							%>	
							</select>								
						</TD>   
						<TD> 			
							<select name="whsno"  style="width:98%" onchange="tblcd_change(<%=currentrow-1%>)">
							<% sql="select sys_type, sys_value  from BasicCode  where func = 'whsno' " 
							   set rst=conn.execute(sql) 
							   while not rst.eof  			
							%>			
								<option value="<%=rst("sys_type")%>" <%if Ucase(trim(tmpRec(CurrentPage, CurrentRow, 8)))=rst("sys_type") then %> selected<%end if %>>
									<%=rst("sys_type")%>-<%=rst("sys_value")%>
								</option>
							<%rst.movenext
							wend 
							rst.close : set rst=nothing
							%>	
							</select>		 	
						</TD>   
						<TD> 		
							<select name="usergroup" style="width:98%"   onchange="tblcd_change(<%=currentrow-1%>)" >
							<% sql="select sys_type, sys_value  from BasicCode  where func = 'Grp' " 
								 if session("netuser")="PELIN" or session("netuser")="NGHIA" then 
								 else
									if session("rights")<="1" then 					
										sql=sql&" and  sys_type in ( '1','2', '3', '7' , 'B', 'C') " 
									elseif 	session("rights")="2" then 					
										sql=sql&" and  sys_type in ( '2', '3', '7' , 'B', 'C') " 
									else 
										sql=sql&" and  sys_type ='"& ucase(trim(tmpRec(CurrentPage, CurrentRow, 3))) &"' " 
									end if 	
								 end if 
							   set rst=conn.execute(sql) 
							   while not rst.eof  			
							%>			
								<option value="<%=rst("sys_type")%>" <%if Ucase(trim(tmpRec(CurrentPage, CurrentRow, 3)))=rst("sys_type") then %> selected<%end if %>>
									<%=rst("sys_type")%>-<%=rst("sys_value")%>
								</option>
							<%rst.movenext
							wend  
							rst.close : set rst=nothing
							%>	
							</select>						
						</TD> 
						<%if session("netuser")="PELIN" or session("netuser")="NGHIA" then %>				
						<TD> 						
							<input type="text" style="width:98%" size=10   name="pwd" value="<%=Ucase(trim(tmpRec(CurrentPage, CurrentRow, 5)))%>" onchange="tblcd_change(<%=currentrow-1%>)" >			 
						</TD>
						<%else%>
							<input size=10 type="hidden"   name="pwd" value="<%=Ucase(trim(tmpRec(CurrentPage, CurrentRow, 5)))%>" >
						<%end if %>
					</tr>
					<%next%> 
				</TABLE>
			</td>
		</tr>
		<tr>
			<td>
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
				<TABLE border=0 width=500 align="center" class="txt">
					<tr>
						<td align="left" width="50%">
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
						<td align="right" width="50%">
						<INPUT TYPE="button" name=send VALUE="CONFIRM" class="btn btn-sm btn-danger"  onclick="go()"  >
						<INPUT TYPE="button" name=send VALUE="CANCEL" class="btn btn-sm btn-outline-secondary"  onclick="Clear()">
						</TD>
					</TR>
				</TABLE>
			</td>
		</tr>
	</table>
			
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

<script language="javascript">

function go(){
   <%=SELF%>.action = "<%=SELF%>.UpdateDB.asp";
   <%=SELF%>.submit();
}

function Clear(){
	open("<%=SELF%>.asp", "Fore");
}

function del(index){
	 if(<%=SELF%>.func[index].checked==true) 
		<%=SELF%>.op[index].value = "del";		 	
	 else	
		<%=SELF%>.op[index].value = "";	
		
} 

function del2(index){	 
	 if(<%=SELF%>.opundo[index].checked==true) 
		<%=SELF%>.op(index).value = "upd";		 	
	 else	
		<%=SELF%>.op(index).value = "";		 	
}



function tblcd_change(index){
	<%=SELF%>.op(index).value = "upd";	  
}


</script>


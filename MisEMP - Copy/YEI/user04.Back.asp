<%@language=vbscript codepage=65001%>
<meta HTTP-EQUIV="Content-Type" content="text/html; charset=utf-8">  
<!-- #include file = "../../GetSQLServerConnection.fun" -->
<% 
self="user04"
func = request("func")
'tblcd = request("tblcd")
tbldesc = request("tbldesc") 
username= TRIM(request("username"))
pwd = request("pwd") 
index = request("index") 
WHSNO = request("WHSNO") 
groupid = request("groupid") 
'tmpRec = Session("YSBHE0502")
CurrentPage = request("CurrentPage") 

code01 = trim(request("code01"))
Set conn = GetSQLServerConnection() 
'on error resume next

Select Case func
		case "chkempid"
			sql="select * from [yfynet].dbo.view_empfile where empid='"& code01 &"'"
			response.write sql 
			set rs=conn.execute(Sql)
			if rs.eof then %>
				<script language=vbs>
					alert "輸入錯誤!!"
					parent.Fore.<%=self%>.empid(<%=index%>).value=""
					parent.Fore.<%=self%>.empname(<%=index%>).value="" 
					parent.Fore.<%=self%>.unitno(<%=index%>).value=""
					parent.Fore.<%=self%>.groupid(<%=index%>).value=""
					parent.Fore.<%=self%>.job(<%=index%>).value=""
					parent.Fore.<%=self%>.country(<%=index%>).value=""
					parent.Fore.<%=self%>.levid(<%=index%>).value=""
					parent.Fore.<%=self%>.email1(<%=index%>).value=""
					parent.Fore.<%=self%>.empid(<%=index%>).focus()
				</script>
<%				response.end 
			else%>
				<script language=vbs>
					parent.Fore.<%=self%>.empid(<%=index%>).value="<%=ucase(code01)%>"
					parent.Fore.<%=self%>.empname(<%=index%>).value="<%=rs("empnam_cn")%>"
					parent.Fore.<%=self%>.unitno(<%=index%>).value="<%=rs("unitno")%>"					
					parent.Fore.<%=self%>.groupid(<%=index%>).value="<%=rs("groupid")%>"					
					parent.Fore.<%=self%>.job(<%=index%>).value="<%=rs("job")%>"
					parent.Fore.<%=self%>.country(<%=index%>).value="<%=rs("country")%>"
					parent.Fore.<%=self%>.email1(<%=index%>).value="<%=rs("email")%>"
					
					'parent.best.cols="100%,0%"
				</script>						
<%			end if
			set rs=nothing 
	   Case "datachg"			
			tmpRec(CurrentPage,index + 1,0) = "upd"
			tmpRec(CurrentPage,index + 1,2) = username
			tmpRec(CurrentPage,index + 1,3) = tbldesc
			tmpRec(CurrentPage,index + 1,5) = pwd
			tmpRec(CurrentPage,index + 1,8) = WHSNO
			tmpRec(CurrentPage,index + 1,7) = groupid
		Case "del"			
			tmpRec(CurrentPage,index + 1,0) = "del"
		Case "no"			
			tmpRec(CurrentPage,index + 1,0) = "no"		
	
End Select
'Response.Write "index = " & index &"<BR>"
'Response.Write "0-" & tmpRec(CurrentPage,index + 1,0) &"<BR>"
'Response.Write "2-" & tmpRec(CurrentPage,index + 1,2) &"<BR>"
'Response.Write "3-" & tmpRec(CurrentPage,index + 1,3) &"<BR>"
'Response.Write "5-" & tmpRec(CurrentPage,index + 1,5) &"<BR>"
'Response.Write "7-" & tmpRec(CurrentPage,index + 1,7) &"<BR>"
'Response.Write "8-" & tmpRec(CurrentPage,index + 1,8) &"<BR>"
'Session("YSBHE0502") = tmpRec
%>
 
<%@ Language=VBScript codepage=65001%>
<!-- #include file = "../GetSQLServerConnection.fun" -->
<!-- #include file="../ADOINC.inc" -->
<meta HTTP-EQUIV="Content-Type" content="text/html; charset=utf-8">
<%
Response.Expires = 0
session.codepage="65001"	
self="YEBE0502"
%>
<html>
<head>
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta HTTP-EQUIV="refresh" >
</head>
<body>
<%

func = request("func")
codestr01 = request("code1")
codestr02 = trim(request("code2"))
codestr03 = request("code3")
codestr04 = request("code4")
codestr05 = request("code5")
codestr06 = request("code6")
codestr07 = request("code7")  

index = request("index")
CurrentPage = request("CurrentPage")

'Response.Write func & "<p>"

Response.Write "CurrentPage=" & CurrentPage & "<p>"
Response.Write index & "-index <BR>"
'response.write codestr04 &"<BR>"



Select Case func
		Case "chkemp"
			Set conn = GetSQLServerConnection()
			sql="select * from view_empfile where empid='"& codestr01 &"' "
			set rs=conn.execute(Sql) 
			if rs.eof then 
%>				<script language=vbs>
					alert "資料輸入錯誤!!"
					parent.Fore.<%=self%>.empid(<%=index%>).value=""
					parent.Fore.<%=self%>.groupid(<%=index%>).value=""
					parent.Fore.<%=self%>.empname(<%=index%>).value=""
					parent.Fore.<%=self%>.empid(<%=index%>).focus()
				</script>			
<%			else%>				
				<script language=vbs>				
					parent.Fore.<%=self%>.empid(<%=index%>).value="<%=ucase(codestr01)%>"
					parent.Fore.<%=self%>.groupid(<%=index%>).value="<%=rs("groupid")%>"
					parent.Fore.<%=self%>.empname(<%=index%>).value="<%=rs("empnam_cn")&rs("empnam_vn")%>"
					parent.Fore.<%=self%>.whsno(<%=index%>).value="<%=rs("whsno")%>"
					parent.Fore.<%=self%>.country(<%=index%>).value="<%=rs("country")%>"
					parent.Fore.<%=self%>.pzjno(<%=index%>).focus()
				</script>			
<%			end if 							
  		rs.close
			conn.close
			set rs=nothing	
			set conn=nothing

End Select
  
%>
</BODY>
</HTML>


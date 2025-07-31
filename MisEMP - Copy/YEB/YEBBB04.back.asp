<%@ Language=VBScript codepage=65001%>
<!-- #include file = "../GetSQLServerConnection.fun" -->
<!-- #include file="../ADOINC.inc" -->
<meta HTTP-EQUIV="Content-Type" content="text/html; charset=utf-8">
<%
Response.Expires = 0
session.codepage="65001"	
self="YEBBB04"
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
codestr02 = trim(request("codestr02"))
codestr03 = request("codestr03")
codestr04 = request("codestr04")

c1=request("c1")
c2=request("c2")
c3=request("c3")
c4=request("c4")
c5=request("c5")
c6=request("c6")
c7=request("c7")
c8=request("c8")
c9=request("c9")
c10=request("c10") 

index = request("index")
CurrentPage = request("CurrentPage")

'Response.Write func & "<p>"

Response.Write "CurrentPage=" & CurrentPage & "<p>"
Response.Write index & "-index <BR>"
'response.write codestr04 &"<BR>"
tmpRec = Session("yebbb04B")

Select Case func
		Case "del"
			tmpRec(CurrentPage,index + 1,0) = "del"
		Case "no"
			tmpRec(CurrentPage,index + 1,0) = ""
		case "upd" 		
			tmpRec(CurrentPage,index + 1,0) = "upd" 			
			tmpRec(CurrentPage,index + 1,1) = c1
			tmpRec(CurrentPage,index + 1,2) = c2
			tmpRec(CurrentPage,index + 1,3) = c3
			tmpRec(CurrentPage,index + 1,4) = c4
			tmpRec(CurrentPage,index + 1,5) = c5
			tmpRec(CurrentPage,index + 1,6) = c6
			tmpRec(CurrentPage,index + 1,18) = c7
			tmpRec(CurrentPage,index + 1,20) = c8
			tmpRec(CurrentPage,index + 1,21) = c9
			tmpRec(CurrentPage,index + 1,22) = c10
			
		case "chkemp" 
			Set conn = GetSQLServerConnection()
			sql="select *from view_empfile where  empid='"& codestr01 &"' "	
			response.write sql 
			set rs=conn.execute(sql)
			if rs.eof then %>
				<script language=vbs>
					alert "輸入錯誤!!ko. co ma so the!!"
					parent.Fore.<%=self%>.empid(<%=index%>).value=""
					parent.Fore.<%=self%>.empname(<%=index%>).value=""
					parent.Fore.<%=self%>.F_groupid(<%=index%>).value=""
					parent.Fore.<%=self%>.empid(<%=index%>).focus()
				</script>
			<%else
				'tmpRec(CurrentPage,index + 1,0) = "upd"
				tmpRec(CurrentPage,index + 1,1) = rs("empid")
			%>	<script language=vbs>										
					parent.Fore.<%=self%>.empid(<%=index%>).value="<%=rs("empid")%>"
					parent.Fore.<%=self%>.empname(<%=index%>).value="<%=rs("empnam_cn")%>"&"<%=rs("empnam_vn")%>"
					parent.Fore.<%=self%>.F_groupid(<%=index%>).value="<%=rs("gstr")%>"
					parent.Fore.<%=self%>.cardName(<%=index%>).focus()
				</script>	
			<%
			end if 	
			set rs=nothing 			
			conn.close
			set conn=nothing
End Select

response.write  tmpRec(CurrentPage,index + 1,0) &"<BR>"
response.write  tmpRec(CurrentPage,index + 1,1) &"<BR>"
response.write  tmpRec(CurrentPage,index + 1,2) &"<BR>"
response.write  tmpRec(CurrentPage,index + 1,3) &"<BR>"
response.write  tmpRec(CurrentPage,index + 1,4) &"<BR>"
response.write  tmpRec(CurrentPage,index + 1,5) &"<BR>"
response.write  tmpRec(CurrentPage,index + 1,6) &"<BR>"
response.write  tmpRec(CurrentPage,index + 1,18) &"<BR>"
response.write  tmpRec(CurrentPage,index + 1,20) &"<BR>"
response.write  tmpRec(CurrentPage,index + 1,21) &"<BR>"
response.write  tmpRec(CurrentPage,index + 1,22) &"<BR>"

Session("yebbb04B") = tmpRec
%>
</BODY>
</HTML>


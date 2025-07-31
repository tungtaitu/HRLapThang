<%@ Language=VBScript codepage=65001%>
<!-- #include file = "../GetSQLServerConnection.fun" -->
<!-- #include file="../ADOINC.inc" -->
<meta HTTP-EQUIV="Content-Type" content="text/html; charset=utf-8">
<%
Response.Expires = 0
session.codepage="65001"	
self="YEBBB0201"
%>
<html>
<head>
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta HTTP-EQUIV="refresh" >
</head>
<body>
<%

func = request("func")
codestr01 = request("codestr01")
codestr02 = trim(request("codestr02"))
codestr03 = request("codestr03")
codestr04 = request("codestr04")

IF codestr02<>"" THEN
	codestr02 = REPLACE(codestr02, "'", "" )
	codestr02 = REPLACE (codestr02, vbCrLf ,"<br>")
	response.write "=="&codestr02& "<BR>"
END IF 

index = request("index")
CurrentPage = request("CurrentPage")

'Response.Write func & "<p>"

Response.Write "CurrentPage=" & CurrentPage & "<p>"
Response.Write index & "-index <BR>"
'response.write codestr04 &"<BR>"
tmpRec = Session("YEBBB0201")

Select Case func
		Case "del"
			tmpRec(CurrentPage,index + 1,0) = "del"
		case "upd" 		
			tmpRec(CurrentPage,index + 1,0) = "upd" 			
			tmpRec(CurrentPage,index + 1,6) = codestr01
			tmpRec(CurrentPage,index + 1,23) = codestr02
		case "gchg"	
			Set conn = GetSQLServerConnection()
			sql="select * from basicCode where func='zuno' and left(sys_type,4)='"& codestr01 &"' order by sys_type "
			response.write sql
			'response.end
			Set rst= Server.CreateObject("ADODB.Recordset")
	  		rst.Open SQL, conn, 3,1
	  		if not rst.eof then
	  			rcount=rst.RecordCount
	  			redim zunostr(rcount,1)
	  			Response.Write "<form name=form3 >"
	  			for x=0 to  rcount-1
	  				zunostr(x,0)=rst("sys_type")
	  				zunostr(x,1)=rst("sys_value")
	  				'response.write zunostr(x,1) &"<BR>"
	 				Response.Write "<input name=a1 value= '"& rst("sys_type") &"' >"
	 				Response.Write "<input name=a2 value= '"& rst("sys_value") &"' >"
	 				rst.movenext
	  			next
	  			response.write "<input name=a1  >"
	  			response.write "<input name=a2  >"
	  			Response.Write "</form>"
	%>			<script language=vbs>
					redim  zunostr(<%=rcount%>,1)
					Parent.Fore.<%=self%>.Fzuno(<%=index%>).length=<%=rcount%>
					for g = 0 to <%=rcount%>-1
						zunostr(g,0) = document.form3.a1(g).value
						zunostr(g,1) = document.form3.a2(g).value
						'alert 	zunostr(g,0)
						'alert 	zunostr(g,1)
						Parent.Fore.<%=self%>.Fzuno(<%=index%>).options(g).value = zunostr(g,0)
					    Parent.Fore.<%=self%>.Fzuno(<%=index%>).options(g).text = zunostr(g,1)
					next
					'parent.best.cols="100%,0%"
				</script>
	<%		
	  		else
	  			'response.end
	%>			<script language=vbs>
					Parent.Fore.<%=self%>.Fzuno(<%=index%>).length=1
					Parent.Fore.<%=self%>.Fzuno(<%=index%>).options(0).value = ""
					Parent.Fore.<%=self%>.Fzuno(<%=index%>).options(0).text = "------"
					'parent.best.cols="100%,0%"
				</script>
	<% 		end if
			rst.close 
			set rst=nothing			
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
response.write  tmpRec(CurrentPage,index + 1,23) &"<BR>"

Session("YEBBB0201") = tmpRec
%>
</BODY>
</HTML>


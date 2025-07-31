<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<!-- #include file = "../GetSQLServerConnection.fun" -->
<!-- #include file="../ADOINC.inc" -->
<%
SELF = "YEIE03"

ftype = request("ftype")
code = Ucase(Trim(request("code")))
code1 = Ucase(Trim(request("code1")))
code2 = Ucase(Trim(request("code2")))
code3 = Ucase(Trim(request("code3")))
Set conn = GetSQLServerConnection()
index = request("index")
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="refresh">
</head>
<%
select case ftype 
	case "groupchg"
		sql="select * from basicCode where func='zuno' and left(sys_type,4)='"& code &"' order by sys_type "
		response.write sql
		'response.end
		Set rst= Server.CreateObject("ADODB.Recordset")
  		rst.Open SQL, conn, 3,3
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
				Parent.Fore.<%=self%>.zuno(<%=index%>).length=<%=rcount+1%>
				Parent.Fore.<%=self%>.zuno(<%=index%>).style.width=111
				Parent.Fore.<%=self%>.zuno(<%=index%>).options(0).value = ""
				Parent.Fore.<%=self%>.zuno(<%=index%>).options(0).text = "不區分"
				for g =1  to <%=rcount%>
					zunostr(g,0) = document.form3.a1(g-1).value
					zunostr(g,1) = document.form3.a2(g-1).value
					'alert 	zunostr(g,0)
					'alert 	zunostr(g,1)
					Parent.Fore.<%=self%>.zuno(<%=index%>).options(g).value = zunostr(g,0)
				    Parent.Fore.<%=self%>.zuno(<%=index%>).options(g).text = zunostr(g,1)				    
				next
				'parent.best.cols="100%,0%"
			</script>
<%
  		else
  			'response.end
%>			<script language=vbs>
				Parent.Fore.<%=self%>.zuno(<%=index%>).length=1
				Parent.Fore.<%=self%>.zuno(<%=index%>).options(0).value = ""
				Parent.Fore.<%=self%>.zuno(<%=index%>).options(0).text = ""
				'parent.best.cols="100%,0%"
			</script>
<% 		end if
		set rst=nothing 
end  select
%>

</html>

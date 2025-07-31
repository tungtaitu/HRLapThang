<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<!-- #include file = "../GetSQLServerConnection.fun" -->
<!-- #include file="../ADOINC.inc" -->
<%
SELF = "YEBE0102"

ftype = request("ftype")
code = Ucase(Trim(request("code")))
code1 = Ucase(Trim(request("code1")))
code2 = Ucase(Trim(request("code2")))
code3 = Ucase(Trim(request("code3")))
Set conn = GetSQLServerConnection()
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="refresh">
</head>
<%
select case ftype
	case "getempid" 
		if code3="C" then 
			gpstr="S"
			sql="select count(*)+1 as pid  from empfile where left(empid,1)='S'  "
			idLen=4
		else
			if code1="TW" then 
				gpstr="A"
				sql="select count(*)+1 as pid  from empfile where left(empid,1)='A'  "
				idLen=4
			else
				gpstr=left(code2,1)&"F"
				sql="select count(*)+1 as pid  from empfile where left(empid,2)='"& gpstr &"'  "
				idLen=3
			end if 
		end if 	
		Set rst= Server.CreateObject("ADODB.Recordset")
  		rst.Open SQL, conn, 1,3  
  		  		
  		idno= right("00000"&rst("pid"),idLen)   				
		response.write sql &"<BR>"
		response.write idno &"<BR>"  		
		N_empid = right( gpstr & idno ,5)  
		set rst=nothing 		
%>		<script language=vbs>
			parent.Fore.<%=self%>.empid.value="<%=N_empid%>"			
		</script>
<%		
	case "empidchk"
		sql="select * from empfile where empid = '"& code &"' "
		Set rst= Server.CreateObject("ADODB.Recordset")
  		rst.Open SQL, conn, 3,3
  		if not rst.eof  then
%>			<script language=vbs>
				alert "員工編號重複!!請重新輸入!!"
				parent.Fore.<%=self%>.empid.value=""
				parent.Fore.<%=self%>.empid.focus()
			</script>
<% 		else
%>			<script language=vbs>
				parent.Fore.<%=self%>.empid.value="<%=code%>"
				parent.Fore.<%=self%>.indat.focus()
			</script>
<% 		end if
		set rst=nothing
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
				Parent.Fore.<%=self%>.zuno.length=<%=rcount%>
				for g = 0 to <%=rcount%>-1
					zunostr(g,0) = document.form3.a1(g).value
					zunostr(g,1) = document.form3.a2(g).value
					'alert 	zunostr(g,0)
					'alert 	zunostr(g,1)
					Parent.Fore.<%=self%>.zuno.options(g).value = zunostr(g,0)
				    Parent.Fore.<%=self%>.zuno.options(g).text = zunostr(g,1)
				next
				'parent.best.cols="100%,0%"
			</script>
<%
  		else
  			'response.end
%>			<script language=vbs>
				Parent.Fore.<%=self%>.zuno.length=1
				Parent.Fore.<%=self%>.zuno.options(0).value = ""
				Parent.Fore.<%=self%>.zuno.options(0).text = ""
				'parent.best.cols="100%,0%"
			</script>
<% 		end if
		set rst=nothing
	case "UNITCHG"
		sql="select * from basicCode where func='groupid' and left(sys_type,3)='"& code &"' order by sys_type "
		response.write sql &"<BR>"
		'response.end
		Set rst= Server.CreateObject("ADODB.Recordset")
  		rst.Open SQL, conn, 3,3
  		if not rst.eof then
  			rcount=rst.RecordCount
  			redim groupstr(rcount,1)
  			Response.Write "<form name=form1 >"
  			for x=0 to  rcount-1
  				groupstr(x,0)=rst("sys_type")
  				groupstr(x,1)=rst("sys_value")
  				'response.write zunostr(x,1) &"<BR>"
 				Response.Write "<input name=a1 value= '"& rst("sys_type") &"' >"
 				Response.Write "<input name=a2 value= '"& rst("sys_value") &"' >"
 				rst.movenext
  			next
  			response.write "<input name=a1  >"
  			response.write "<input name=a2  >"
  			Response.Write "</form><BR>" 
  			
  			sqlz="select * from basicCode where func='zuno' and left(sys_type,4)='"& groupstr(0,0) &"' order by sys_type "
			response.write sqlz  &"<BR>" 
			'response.end
			Set rst2= Server.CreateObject("ADODB.Recordset")
	  		rst2.Open SQLz, conn, 3,3
	  		if not rst2.eof then
	  			rcountz=rst2.RecordCount
	  			redim zunostr(rcountz,1)
	  			Response.Write "<form name=form3 >"
	  			for t=0 to  rcountz-1
	  				zunostr(t,0)=rst2("sys_type")
	  				zunostr(t,1)=rst2("sys_value")
	  				'response.write zunostr(t,1) &"<BR>"
	 				Response.Write "<input name=a1 value= '"& rst2("sys_type") &"' >"
	 				Response.Write "<input name=a2 value= '"& rst2("sys_value") &"' >"
	 				rst2.movenext
	  			next
	  			response.write "<input name=a1  >"
	  			response.write "<input name=a2  >"
	  			Response.Write "</form>"	  	
	  		else
	  			rcountz=1
	  			Response.Write "<form name=form3 >"
	  			response.write "<input name=a1  >"
	  			response.write "<input name=a1  >"
	  			response.write "<input name=a2  >"
	  			response.write "<input name=a2  >"
	  			Response.Write "</form>"
  			end if
  			'set rst2=nothing
  			'response.end 
  			
%>			<script language=vbs>
				redim  groupstr(<%=rcount%>,1)
				Parent.Fore.<%=self%>.groupid.length=<%=rcount%>
				for g = 0 to <%=rcount%>-1
					groupstr(g,0) = document.form1.a1(g).value
					groupstr(g,1) = document.form1.a2(g).value
					'alert 	groupstr(g,0)
					'alert 	groupstr(g,1)
					Parent.Fore.<%=self%>.groupid.options(g).value = groupstr(g,0)
				    Parent.Fore.<%=self%>.groupid.options(g).text = groupstr(g,1)
				next
				
				redim  zunostr(<%=rcountz%>,1)
				Parent.Fore.<%=self%>.zuno.length=<%=rcountz%>
				for h = 0 to <%=rcountz%>-1
					zunostr(h,0) = document.form3.a1(h).value
					zunostr(h,1) = document.form3.a2(h).value
					'alert 	zunostr(h,0)
					'alert 	zunostr(g,1)
					Parent.Fore.<%=self%>.zuno.options(h).value = zunostr(h,0)
				    Parent.Fore.<%=self%>.zuno.options(h).text = zunostr(h,1)
				next
				'parent.best.cols="100%,0%"
				
				'parent.best.cols="100%,0%"
			</script>
<% 		else
  			'response.end
%>			<script language=vbs>
				Parent.Fore.<%=self%>.groupid.length=1
				Parent.Fore.<%=self%>.groupid.options(0).value = ""
				Parent.Fore.<%=self%>.groupid.options(0).text = "-----" 
				
				Parent.Fore.<%=self%>.zuno.length=1
				Parent.Fore.<%=self%>.zuno.options(0).value = ""
				Parent.Fore.<%=self%>.zuno.options(0).text = "" 
				'parent.best.cols="100%,0%"
			</script>
<% 		end if
		set rs=nothing
end  select
%>

</html>

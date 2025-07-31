<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<!-- #include file = "../../GetSQLServerConnection.fun" --> 
<!-- #include file="../../ADOINC.inc" --> 
<%
SELF = "empfilefore" 

ftype = request("ftype") 
code = request("code") 
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
				
				parent.best.cols="100%,0%"
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
		set rs=nothing 
	case "UNITCHG"		
		sql="select * from basicCode where func='groupid' and left(sys_type,3)='"& code &"' order by sys_type "
		response.write sql
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
  			Response.Write "</form1>"
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
				'parent.best.cols="100%,0%"
			</script> 			
<%  		sql="select * from basicCode where func='zuno' and left(sys_type,4)='"& code &"' order by sys_type "
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
			  	Response.Write "</form>"
%>				<script language=vbs>				
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
					
					parent.best.cols="100%,0%"
				</script>
<%  		else 
  				'response.end 
%>				<script language=vbs>								
					Parent.Fore.<%=self%>.zuno.length=1				
					Parent.Fore.<%=self%>.zuno.options(0).value = ""
					Parent.Fore.<%=self%>.zuno.options(0).text = ""
					'parent.best.cols="100%,0%"
				</script>
<% 			end if  
			set rst=nothing  			
  		else 
  			'response.end 
%>			<script language=vbs>								
				Parent.Fore.<%=self%>.groupid.length=1				
				Parent.Fore.<%=self%>.groupid.options(0).value = ""
				Parent.Fore.<%=self%>.groupid.options(0).text = ""
				parent.best.cols="100%,0%"
			</script>
<% 		end if    
		set rs=nothing 			
end  select   		
%>

</html>

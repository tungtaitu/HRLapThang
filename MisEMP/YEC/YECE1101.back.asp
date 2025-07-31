<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<!-- #include file = "../GetSQLServerConnection.fun" --> 
<!-- #include file="../ADOINC.inc" --> 
<%
SELF = "YECE1101" 

ftype = request("func") 
code = Ucase(trim(request("code1")))
index=request("index")  
CurrentPage = request("CurrentPage") 
 
response.write "index=" & index &"<BR>"
response.write "ftype=" & ftype &"<BR>" 
exrt = request("exrt")  

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
	case "chkemp"		
		sql="select * from view_empfile where empid='"& code &"' "
		set rsx=conn.execute(Sql)
		if rsx.eof then %>
			<script language=vbs>								
				alert "no data complete!!輸入錯誤!!"
				parent.Fore.<%=self%>.empid(<%=index%>).value=""
				parent.Fore.<%=self%>.empid(<%=index%>).focus()
			</script>
<%	  response.end 
		else
%>		<script language=vbs>								
				parent.Fore.<%=self%>.empid(<%=index%>).value="<%=Code%>"
				parent.Fore.<%=self%>.empnam(<%=index%>).value="<%=rsx("empnam_cn")%>"&"<%=rsx("empnam_vn")%>"
				parent.Fore.<%=self%>.indat(<%=index%>).value="<%=rsx("nindat")%>"
				parent.Fore.<%=self%>.gstr(<%=index%>).value="<%=rsx("gstr")%>"
				if parent.Fore.<%=self%>.dfudamt.value<>"" then 
					if parent.Fore.<%=self%>.ut_mtax(<%=index%>).value ="" then 
						parent.Fore.<%=self%>.ut_mtax(<%=index%>).value= formatnumber(parent.Fore.<%=self%>.dfudamt.value,0)
					end if  
				end if 
				parent.Fore.<%=self%>.person_qty(<%=index%>).focus()
			</script>
<% 		end if  
		set rsx=nothing 			 
end  select   		
%>
</html>

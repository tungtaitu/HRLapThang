<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<!-- #include file = "../GetSQLServerConnection.fun" --> 
<!-- #include file="../ADOINC.inc" --> 
<%
SELF = "empholiday" 

ftype = request("ftype") 
code = request("code")  
code1=request("code1")
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
	case "chkempid"		
		sql="select * from empfile  where empid='"& code &"' "
		response.write sql
		'response.end   
		Set rst= Server.CreateObject("ADODB.Recordset")
  		rst.Open SQL, conn, 3,3      		
  		if not rst.eof then    			 
%>			<script language=vbs>				
				Parent.Fore.<%=self%>.empid.value = "<%=rst("empid")%>"
			    Parent.Fore.<%=self%>.empnameCN.value = "<%=rst("empnam_CN")%>"
			    Parent.Fore.<%=self%>.EMPNAMEVN.value = "<%=rst("empnam_VN")%>"
			    Parent.Fore.<%=self%>.HOLIDAY_TYPE.focus()
			</script>
<%  					
  		else   			
%>			<script language=vbs>				
				alert "員工編號輸入錯誤!!"
				Parent.Fore.<%=self%>.empid.focus()
				Parent.Fore.<%=self%>.empid.value = ""
			    Parent.Fore.<%=self%>.empnameCN.value = ""
			    Parent.Fore.<%=self%>.EMPNAMEVN.value = ""
			</script>
<%			response.end   			 
 		end if  
		set rs=nothing  
		
	case "dayschg"
		sql="select  isnull(count(*),0)   as ccnt  from   ydbmcale  "&_
			"where status in ( 'H2', 'H3' )  "&_
			"and convert(char(10), dat,111) between '"& code &"' and '"& code1 &"' " 
		response.write sql 	
		Set rst= Server.CreateObject("ADODB.Recordset")
  		rst.Open SQL, conn, 3,3  
		if not rst.eof then    			 
%>			<script language=vbs>				
				Parent.Fore.<%=self%>.HDcnt.value = "<%=rst("ccnt")%>"			    
			</script>
<%  					
  		else   			
%>			<script language=vbs>				
				Parent.Fore.<%=self%>.HDcnt.value="0"				
			</script>
<%			response.end   			 
 		end if  
		set rs=nothing    		
		
		 	
end  select   		
%>

</html>

<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<!-- #include file = "../GetSQLServerConnection.fun" --> 
<!-- #include file="../ADOINC.inc" --> 
<%
SELF = "YECE02" 

ftype = request("ftype") 
code = request("code") 
index=request("index")  
CurrentPage = request("CurrentPage") 
tmpRec = Session("empfilesalary") 
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="refresh"> 
</head>
<%
select case ftype  
 
	CASE "memochk"
        tmpRec(CurrentPage,index + 1,0) = "UPD"		   		
        tmpRec(CurrentPage,index + 1,70) = code    
        response.write tmpRec(CurrentPage,index + 1,70) &"<BR>"   
end  select   		
Session("empfilesalary") = tmpRec
%>
</html>
<SCRIPT LANGUAGE=VBSCRIPT> 
	window.close()
</script>	
<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<!-- #include file = "../GetSQLServerConnection.fun" --> 
<!-- #include file="../ADOINC.inc" --> 
<%
SELF = "yece05" 

ftype = request("ftype") 
code = request("code") 
index=request("index")  
CurrentPage = request("CurrentPage") 
tmpRec = Session("empfilesalaryCN") 
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
        tmpRec(CurrentPage,index + 1,47) = code    
        response.write tmpRec(CurrentPage,index + 1,47) &"<BR>"   
end  select   		
Session("empfilesalaryCN") = tmpRec
%>
</html>
<SCRIPT LANGUAGE=VBSCRIPT> 
	window.close()
</script>	
<%@LANGUAGE="VBSCRIPT" %>
<!-- #include file = "../../GetSQLServerConnection.fun" -->
<!-- #include file="../../ADOINC.inc" -->
<%
'Session("NETUSER")=""
'SESSION("NETUSERNAME")=""

'Response.Expires = 0
'Response.Buffer = true

code1= request("code1")  

s_empid=request("s_empid")
s_dat1=request("s_dat1")
s_dat2=request("s_dat2")
s_country=request("s_country")

Set CONN = GetSQLServerConnection()
conn.BeginTrans  

sql="delete  empvisaData where  aid='"& code1 &"' " 
response.write s_country  
conn.execute(Sql) 


if conn.Errors.Count = 0 then 
	conn.CommitTrans
	conn.close
	set conn=nothing
	'response.redirect "empbe0301.Fore.asp?s_empid="& s_empid &"&s_dat1=" & s_dat1 &"&_sdat2=" & s_dat2 &"&s_country=" & s_country 
	'response.end  			
%>	<SCRIPT LANGUAGE=VBSCRIPT>		
		OPEN "empbe0301.Fore.asp?s_empid="& "<%=s_empid%>" &"&s_dat1=" & "<%=s_dat1%>" &"&_sdat2=" & "<%=s_dat2%>" &"&s_country=" & "<%=s_country%>"  , "_self"  
	</script>	
<%	response.end	
ELSE
	conn.RollbackTrans	
	conn.close
	set conn=nothing
%>	<SCRIPT LANGUAGE=VBSCRIPT>
		ALERT "資料處理失敗DATA CommitTrans ERROR !!"
		OPEN "empbe0301.asp" , "_self" 
	</script>	
<%	response.end 
END IF  
%> 
 

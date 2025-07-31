<%@LANGUAGE="VBSCRIPT" %>
<!-- #include file = "../../GetSQLServerConnection.fun" --> 
<!-- #include file="../../ADOINC.inc" -->  
<%
'Session("NETUSER")=""
'SESSION("NETUSERNAME")=""

'Response.Expires = 0
'Response.Buffer = true 

Set CONN = GetSQLServerConnection()   
Days= request("days")
yymm= request("yymm")
t3str = split(request("status"),", ") 
'response.write t3str(0) 
'for x = 1 to ubound(t3str)
'	response.write t3str(x) &"<BR>"
'	response.write "...." &"<BR>"
'next 

For i = 1 to Days
    sts=request("status")(i) 
    if request("status")(i)="" then     	
    	sts="H1" 
    end if 	
    'if request("status")(i)<>"" then     	
    	sql="update YDBMCALE set status='"& sts &"' "&_
    		"where convert(char(10),dat,111)= '"& request("dat")(i) &"' "  
    	response.write sql &"<BR>" 
    	conn.execute(sql)    	
    'end if 		
Next 
'response.end 
response.redirect "empbasicB.fore.asp?yymm="& yymm

%>
 
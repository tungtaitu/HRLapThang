<%@LANGUAGE="VBSCRIPT" %>
<!-- #include file = "../../GetSQLServerConnection.fun" --> 
<!-- #include file="../../ADOINC.inc" -->  
<%
Response.Expires = 0
Response.Buffer = true 

dat1= request("dat1")
dat2= request("dat2")
s_empid= request("s_empid") 
gid= request("gid")

'response.write dat1 &"<BR>" 
'response.write dat2 &"<BR>" 
'response.write s_empid &"<BR>" 
'response.write gid &"<BR>" 


code1= request("autoid")
code2= request("empid")
code3= replace(request("wdat") ,"/","")


Set CONN = GetSQLServerConnection()  

sql="delete empholiday where  autoid='"& code1 &"' and empid='"& code2 &"' and jiatype='G' " 
response.write sql  &"<BR>" 
conn.execute(sql) 

sqlstr="update empwork set kzhour=jiag , jiag=0 where  empid='"& code2 &"' and workdat='"& code3 &"' " 
response.write sqlstr  &"<BR>" 
conn.execute(sqlstr) 
'response.end  


response.redirect "empde0402.fore.asp?dat1="& dat1 &"&dat2="& dat2 &"&s_empid=" & s_empid &"&gid=" & gid 



if conn.Errors.Count = 0 then 
	conn.CommitTrans
	'Set Session("empworkbC") = Nothing 	    
%>	<SCRIPT LANGUAGE=VBSCRIPT>
		ALERT "資料處理成功!!"&chr(13)&"DATA CommitTrans SUCCESS !!"
		OPEN "empde0201.asp" , "_self" 
		'window.close()
	</script>	
<%
ELSE
	conn.RollbackTrans	
%>	<SCRIPT LANGUAGE=VBSCRIPT>
		ALERT "資料處理失敗!!"&chr(13)&"DATA CommitTrans ERROR !!"
		OPEN "empde0201.asp" , "_self" 
	</script>	
<%	response.end 
END IF  
%>
 
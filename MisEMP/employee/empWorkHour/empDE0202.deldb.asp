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


code1= request("code1")
code2= request("code2")

Set CONN = GetSQLServerConnection()  

sql="update empforget set status='D' , mdtm=getdate(), muser='"& session("NetUser") &"'+'-Del' where empid='"& code2 &"' and autoid='"& code1 &"' " 
'response.write sql &"<BR>"
conn.execute(sql)

'response.end  
sqla="select convert(char(8), dat,112) as ndat, * from empforget where autoid='"& code1 &"' and empid='"& code2 &"' " 
Set rds = Server.CreateObject("ADODB.Recordset") 
rds.open sqla, conn,3, 3
if not rds.eof then 
    toth = cdbl(rds("toth"))
    if toth>8 then toth=8
    sqlb="update empwork set kzhour=kzhour+"& toth &" , toth=toth-"& toth &" , forget=forget-1 "&_
         "where empid='"& rds("empid") &"' and workdat='"& rds("ndat") &"' "
    conn.execute(Sqlb)     
    response.write sqlb 
end if


response.redirect "empde0202.fore.asp?dat1="& dat1 &"&dat2="& dat2 &"&s_empid=" & s_empid &"&gid=" & gid 

'response.end 

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
 
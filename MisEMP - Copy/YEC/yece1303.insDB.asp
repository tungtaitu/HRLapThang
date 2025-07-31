<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%> 
<!-- #include file = "../GetSQLServerConnection.fun" -->
 
<HEAD>
<META http-equiv="Content-Type" content="text/html; charset=utf-8">
</HEAD>
<%
Set conn = GetSQLServerConnection()	 
Response.Charset="utf-8"
'如果發生錯誤，先跳過
'On Error Resume Next   

'code1=Request.QueryString("fs") 
'response.write "XXXX"
'response.end  
self="yece1303"

whsno = request("whsno")
years = request("years")
c1 = request("c1")
g1 = request("groupid")
eid= request("eid") 
khud= request("khud") 
 'response.write request("pagerec")
tmprec= session("yece1303b") 
Conn.BeginTrans	 
 for x = 1 to request("pagerec")
	'response.write whsno & years & tmprec(1,x,4) & tmprec(1,x,5) & tmprec(1,x,6) & "<BR>"
	country=tmprec(1,x,3) 
	empid=tmprec(1,x,4) 
	indate =tmpRec(1, x, 5)
	groupid=tmpRec(1, x, 6)
	
	fensu = request("fensu")(x)
	kj = request("grade")(x)
	'response.write a & b &"<BR>" 
	
	sql1="delete  EmpNZKH where years='"&years&"' and empid='"& empid &"' and khud='"&khud&"' "
	conn.execute(sql1)
	strsql1="insert into EmpNZKH([years], [whsno], [country],[empid],[fensu], [kj], [mdtm], [muser], indat, groupid ,khud) values ( "&_
			"'"&years&"','"&whsno&"', '" &country& "','" &empid & "', '" &fensu& "','" &kj& "' ,"&_
			"getdate(),'"& session("netuser") &"','"&indate&"','"&groupid&"','"&khud&"'  )"
	conn.execute(strsql1) 
	'response.write strsql1 &"<BR>"
  next 

	'response.end 
set session("yece1303b") =nothing 
if err.number = 0 then
	Conn.CommitTrans
%><script language="vbscript">	
		alert "資料處理成功 data complete success (OK)!!"
		open "<%=self%>.fore.asp?flag=S&whsno="&"<%=whsno%>"&"&years="&"<%=years%>"&"&c1="&"<%=c1%>"&"&groupid="&"<%=g1%>"&"&eid="&"<%=eid%>"&"&khud="&"<%=khud%>" , "_self"				
	</script>
<%	
else	
	conn.RollbackTrans 
%><script language="vbscript">	
		alert "資料處理失敗 data complete Fail (Error)!!"
		open "<%=self%>.fore.asp?flag=S&whsno="&"<%=whsno%>"&"&years="&"<%=years%>"&"&c1="&"<%=c1%>"&"&groupid="&"<%=g1%>"&"&eid="&"<%=eid%>"&"&khud="&"<%=khud%>" , "_self"				
	</script>
<%end if %>
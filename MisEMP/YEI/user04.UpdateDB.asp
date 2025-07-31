<%@language=vbscript codepage=65001%>
<!-------- #include file = "../../GetSQLServerConnection.fun" --------->
 
<meta HTTP-EQUIV="Content-Type" content="text/html; charset=UTF-8">
<%
Response.Buffer = true
Response.Expires = 0
%>

<%
CurrentPage = request("CurrentPage")
TotalPage = request("TotalPage")
PageRec = request("PageRec")
gTotalPage = request("gTotalpage")

Set conn = GetSQLServerConnection()

'on error resume next
conn.BeginTrans

x = 0
y = ""
for i = 1 to pagerec 
 	whsno = request("whsno")(i)
 	empid = request("empid")(i)
 	levid = request("levid")(i)
 	country=request("country")(i)
 	unitno=request("unitno")(i) 
 	groupid=request("groupid")(i)
 	job=request("job")(i)
 	email1=request("email1")(i)
 	email2=request("email2")(i)
 	aid=request("aid")(i) 
 	op=trim(request("op")(i))
 	if op="D" then 
 		sqld="update yfycq set status='D',mdtm=getdate(), muser='"&session("userid")&"' where aid='"&aid&"'"
 		conn.execute(sqld)
 		response.write sqld 	&"<BR>"	
 		F_whsno=whsno	
 	end if 	 	
 	if aid="" and (levid<>"" and whsno<>"" and empid<>"" ) then  		
 		sql="insert into yfycq (country, whsno, job, unitno , groupid, empid, email1, email2, mdtm, muser, levid ) values ("&_
 			"'"&country&"','"&whsno&"','"&job&"','"&unitno&"','"&groupid&"','"&empid&"','"&email1&"', "&_
 			"'"&email2&"',getdate(),'"&session("userid")&"','"&levid&"' ) " 
 		conn.execute(Sql)
 		response.write sql 	&"<BR>"
 		F_whsno=whsno	
 	else
 		if empid<>"" and op="upd" then 
	 		sql="update yfycq set email1='"&email1&"', email2='"&email2&"', mdtm=getdate(), "&_
	 			"muser='"&session("netuser")&"' where aid='"&aid&"' "	
	 		conn.execute(Sql)
	 		response.write sql 	&"<BR>"	
	 		F_whsno=whsno	
	 	end if 	
 	end if 
	'response.write whsno&"<BR>"
	'response.write empid&"<BR>"  
	
	
next 

'Response.Write  y 
'Response.End  

 if conn.Errors.Count = 0 then 
	conn.CommitTrans 
	response.redirect "user04.fore.asp?queryx=x&F_whsno="&F_whsno
	'Session("Title") = "修改使用者群組"
	'Session("Name") = "IMessage"
	'Session("NO") = "資料處理成功" & x & " 筆"
	'Session("MessageCode") = "Success"
	'Session("KeyValue") = "USER<BR>" & Y
	'Session("SubmitValue") = "回修改使用者群組"
	'Session("Action") = "YEAAE0501.asp"
	'Response.Redirect "IMessage.asp"
	'Set conn = Nothing 
 else
	conn.RollbackTrans
	Set conn = Nothing%>
	<script language=vbs>
		alert "資料處理錯誤!!"
		open "user04.asp", "_Fore"
	</script> 
<%end if %>

<%@language=vbscript codepage=65001%>
<!-------- #include file = "../GetSQLServerConnection.fun" --------->
<!--#include file="../ADOINC.inc"-->
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
tmpRec = Session("YSBHE0502")
'on error resume next
conn.BeginTrans

x = 0
y = ""

arr_op=request("op")  
arr_user = request("user") 
arr_stt = request("stt") 

response.write ubound(split(arr_op,",")) &"<br>"
response.write ubound(split(arr_user,","))  &"<br>"
response.write ubound(split(arr_stt,","))  &"<br>"
e = ubound(split(arr_stt,",")) 
for i = 1 to  e 
	op=trim(request("op")(i))
	muser=trim(request("user")(i))
	username=trim(request("username")(i))
	pswd=trim(request("pwd")(i))
	groupid=trim(request("groupid")(i))
	whsno=trim(request("whsno")(i))
	rights=trim(request("usergroup")(i))
	empid=trim(request("empid")(i))
	
	if op="del" then 
		sql2 = "update sysuser set status='D' , mdtm=getdate(), keyinby='"& session("Netuser") &"' where muser = '" &  muser & "'; " 				
		Response.Write sql2 &"<BR>"
		conn.Execute(sql2)
		x = x + 1
		if y = "" then
		   y = muser & "-" & username & "-" & rights &" 已刪除"
		else
		   y = y + "<BR>" & muser & "-" & username& "-" & rights &" 已刪除" 
		end if		
	elseif op="upd" then 
		sql="update sysuser set group_id='"& groupid &"', rights='"& rights &"', username=N'"& UserName &"', "&_
					"password='"& pswd &"' , WHSNO='"& whsno &"', mdtm=getdate(), keyinby='"& session("Netuser") &"'  "&_
					",empid='"&empid&"', status=''  where muser='"& muser &"'; "
				conn.execute(sql)
				Response.Write sql &"<BR>"
				X=X+1 
				if y = "" then
					y = muser & "-" & username & "-" & rights &" 已修改"
				else
					y = y + "<BR>" & muser & "-" & username & "-" & rights &" 已修改"
				end if
	end if 
	
next  

'response.write y 
'response.end
 
 if conn.Errors.Count = 0 then 
	conn.CommitTrans
	Set Session("YSBHE0502") = Nothing
	response.redirect "YEAAE0501.asp"
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
	Set Session("YSBHE0502") = Nothing
    Set cmd = Nothing
	Set conn = Nothing
	Session("Title") = "修改使用者群組 "
	Session("Name") = "IMessage"
	Session("NO") = "資料處理失敗 "
	Session("MessageCode") = "Fail"
	Session("KeyValue") = "" 
	Session("SubmitValue") = "回修改使用者群組"
	Session("Action") = "YEAAE0501.asp"
	Response.Redirect "IMessage.asp"
	Set conn = Nothing 
 end if %>

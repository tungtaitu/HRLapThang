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
for i = 1 to gTotalPage 
	for j = 1 to PageRec 
		if tmpRec(i, j , 0) = "del" then 
			if tmpRec(i, j, 1) <> "" then
				Set conn2 = GetSQLServerConnection()
				sql2 = "delete sysuser  where muser = '" & tmpRec(i, j, 1) & "' " 				
				Response.Write sql2 &"<BR>"
				conn.Execute ( sql2 )
				x = x + 1
				if y = "" then
				   y = tmpRec(i, j, 1) & "-" & tmpRec(i, j, 2)& "-" & tmpRec(i, j, 3) &" 已刪除"
				else
				   y = y + "<BR>" & tmpRec(i, j, 1) & "-" & tmpRec(i, j, 2)& "-" & tmpRec(i, j, 3) &" 已刪除" 
				end if				
			end if 				
		elseif tmpRec(i, j , 0) = "upd" then 
			if tmpRec(i, j, 1) <> "" then
				sql="update sysuser set rights='"& tmpRec(i, j, 3) &"', username='N"& trim(tmpRec(i, j, 2)) &"', "&_
					"password='"& trim( tmpRec(i, j, 5) ) &"' where muser='"& tmpRec(i, j, 1) &"' "
				conn.execute(sql)
				Response.Write sql &"<BR>"
				X=X+1 
				if y = "" then
					y = tmpRec(i, j, 1) & "-" & tmpRec(i, j, 2) & "-" & tmpRec(i, j, 3) &" 已修改"
				else
					y = y + "<BR>" & tmpRec(i, j, 1) & "-" & tmpRec(i, j, 2) & "-" & tmpRec(i, j, 3) &" 已修改"
				end if	
			end if
		end if 
	next
next 
'Response.Write  y 
'Response.End  

 if conn.Errors.Count = 0 then 
	conn.CommitTrans
	Set Session("YSBHE0502") = Nothing
	Set cmd = Nothing
	Session("Title") = "修改使用者群組"
	Session("Name") = "IMessage"
	Session("NO") = "資料處理成功" & x & " 筆"
	Session("MessageCode") = "Success"
	Session("KeyValue") = "USER<BR>" & Y
	Session("SubmitValue") = "回修改使用者群組"
	Session("Action") = "YEAAE0501.asp"
	Response.Redirect "IMessage.asp"
	Set conn = Nothing 
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

<%@LANGUAGE=VBSCRIPT CODEPAGE=950%> 
<!-- #include file = "../GetSQLServerConnection.fun" --> 
<!-- #include file="../ADOINC.inc" -->

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
tmpRec = Session("ADMIN01")

conn.BeginTrans

for i = 1 to gTotalPage 
	for j = 1 to PageRec 
		'response.write trim(tmpRec(i, j, 0)) &"<BR>"
		'response.write trim(tmpRec(i, j, 4)) &"<BR>"
		'response.write trim(tmpRec(i, j, 1)) &"<BR>"
		'response.write trim(tmpRec(i, j, 2)) &"<BR>"
		'response.write trim(tmpRec(i, j, 3)) &"<BR>"
		if tmpRec(i, j , 0) = "del" then		
			sql="delete  basicCode where  autoid='"& trim(tmpRec(i, j, 4)) &"' "	
			conn.execute(sql)	
		else 
			if  trim(tmpRec(i, j, 4))="" then 
				if UCASE(trim(tmpRec(i, j, 1)))<>"" then 
					sql="insert into basicCode (func, sys_Type, sys_Value ) values ( "&_
						"'"& UCASE(trim(tmpRec(i, j, 1))) &"' , '"& UCASE(trim(tmpRec(i, j, 2))) &"' ,  "&_
						"'"& UCASE(trim(tmpRec(i, j, 3))) &"' ) " 
					conn.execute(sql)	
				end if 		
			else
				sql="update basicCode set func='"& UCASE(trim(tmpRec(i, j, 1))) &"' , "&_
					"sys_type='"& UCASE(trim(tmpRec(i, j, 2))) &"' , "&_
					"sys_Value='"& UCASE(trim(tmpRec(i, j, 3))) &"'  "&_
					"where  autoid='"& trim(tmpRec(i, j, 4)) &"' " 
				conn.execute(sql)	
			end if 
		end if 
		response.write sql &"<BR>"
	next
next 
if  conn.errors.count=0  then 
	conn.CommitTrans
	response.redirect "admin01.asp"
else	
	response.write "errors!!"
	Response.End 
end if 	
 if conn.Errors.Count = 0 then 
	conn.CommitTrans
	Set conn = Nothing 	
	Set Session("YDBKE0201") = Nothing
	Session("Title") = "復瓦機參數建檔"
	Session("Name") = YDBKE0201	
	Session("MessageCode") = "Success!!"
	Session("NO") = "資料處理成功 "
	Session("KeyVale") = ""
	Session("SubmitValue") = "修改下一筆"
	Session("Action") = "YDBKE0201.asp"
	Response.Redirect "IMessage.asp"
	Set conn = Nothing 
 else
	conn.RollbackTrans
    Set Session("YDBKE0201") = Nothing
    Set cmd = Nothing
	Set conn = Nothing
	Set Session("YDBKE0201") = Nothing
	Session("Title") = "復瓦機參數建檔 "
	Session("Name") = "IMessage"
	Session("NO") = "資料處理失敗 "
	Session("MessageCode") = "Fail !!"
	Session("KeyValue") = "" 
	Session("SubmitValue") = "回復瓦機參數建檔"
	Session("Action") = "YDBKE0201.asp"
	Response.Redirect "IMessage.asp"
	Set conn = Nothing 
end if 
%>

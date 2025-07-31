<%@language=vbscript CODEPAGE=65001%>
<meta HTTP-EQUIV="Content-Type" content="text/html; charset=UTF-8">
<!---------  #include file="../GetSQLServerConnection.fun"  -------->
<%

Response.Buffer = true
Response.Expires = 0
%>
<%
session.codepage=65001 
self="ytbae01"
CurrentPage = request("CurrentPage")
TotalPage = request("TotalPage")
PageRec = request("PageRec")
DB_TBLID = request("DB_TBLID")
gTotalPage = request("gTotalpage")
Set conn = GetSQLServerConnection()
tmpRec = Session("YDBSB0001EMP")
'on error resume next

if session("netuser")="" then 
	err1()
end if 
conn.BeginTrans
y="" 
for i = 1 to gTotalPage 
	for j = 1 to PageRec  
		if tmpRec(i, j , 0) = "del" then 
			if tmpRec(i, j, 1) <> "" then
				sql="update scode_big set status = 'D' , "&_
					"mdtm=getdate(), muser='"& session("netuser") &"' where  tblid='"& tmpRec(i, j, 1) &"' "
				conn.execute(Sql)
				y =  y & DB_TBLID & " - " &  tmpRec(i, j, 1) & " - " &   tmpRec(i, j, 2) &" (DEL)<BR>"
			end if 
		else
			if tmpRec(i, j, 1) <> "" and  trim(tmpRec(i, j, 2))<>"" then
				sql="select * from scode_big where tblid='"& tmpRec(i, j, 1) &"' and isnull(status,'')<>'D' "	
				Set rs = Server.CreateObject("ADODB.Recordset")
				rs.Open sql, conn, 1, 3
				if tmpRec(i, j , 0) = "update" then 						
					if rs.eof then 
						sql="insert into scode_big (tblid, Description, mdtm, muser ) values ( "&_
							"'"& tmpRec(i, j, 1) &"' , N'"& tmpRec(i, j, 2) &"', getdate(), '"& session("netuser") &"' ) "
						conn.execute(Sql)
						y =  y & DB_TBLID & " - " &  tmpRec(i, j, 1) & " - " &   tmpRec(i, j, 2) &" (已新增)<BR>"
					else					
						sql="update scode_big set Description = N'"& tmpRec(i, j, 2) &"', "&_
							"mdtm=getdate(), muser='"& session("netuser") &"' where  tblid='"& tmpRec(i, j, 1) &"' "
						conn.execute(Sql)
						y =  y & DB_TBLID & " - " &  tmpRec(i, j, 1) & " - " & tmpRec(i, j, 2) &" (已修改)<BR>"
					end if 
				end if 
			end if 	
		end if 		
	next
next 


 if conn.Errors.Count = 0 or err.number=0 then 
	conn.CommitTrans
	Set Session("YDBSB0001EMP") = Nothing
    Set cmd = Nothing
	Set conn = Nothing

	Session("Title") = ""
	Session("Name") = "IMessage"
	Session("NO") = "Data Complete Success (OK) " 
	Session("MessageCode") = "Success"
	Session("KeyValue") =  y 
	Session("SubmitValue") = replace(session("pgname"),"<BR>",chr(13))
	Session("Action") =  self &".asp?pgid="&request("pgid")	
	Response.Redirect "IMessage.asp?pgid="&request("pgid")
	Set conn = Nothing 
 else
 
    Set Session("YDBSB0001EMP") = Nothing
    Set cmd = Nothing
	Set conn = Nothing

	Session("Title") = ""
	Session("Name") = "IMessage"
	Session("NO") = "Data Complete Fail (Error) "
	Session("MessageCode") = "Fail"
	Session("KeyValue") = "" 
	Session("SubmitValue") = replace(session("pgname"),"<BR>",chr(13))
	Session("Action") = self &".asp" 
	Response.Redirect "IMessage.asp?pgid="&request("pgid")
	Set conn = Nothing 
 end if  

function err1()	
	response.write "使用者帳號為空請重新登入!!"	
	response.end  	
end function   
 
 %>

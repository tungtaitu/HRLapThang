<%@language=vbscript CODEPAGE=65001%>
<meta HTTP-EQUIV="Content-Type" content="text/html; charset=UTF-8">
<!---------  #include file="../GetSQLServerConnection.fun"  -------->
<%
Response.Buffer = true
Response.Expires = 0
%>
<%
session.codepage=65001
CurrentPage = request("CurrentPage")
TotalPage = request("TotalPage")
PageRec = request("PageRec")
DB_TBLID = request("DB_TBLID")
gTotalPage = request("gTotalpage")
Set conn = GetSQLServerConnection()
tmpRec = Session("YTBAE0101EMP")
'on error resume next
conn.BeginTrans
y=""
for i = 1 to gTotalPage
	for j = 1 to PageRec
		tbldesc_str = trim(request("tbldesc")(j))
		
		
		if tmpRec(i, j , 0) = "update" then			
			if tmpRec(i, j, 1) <> "" then 				
				if tmpRec(i, j, 4)="" then 
					if (trim(tmpRec(i, j, 1))<>"" and trim(tmpRec(i, j, 2))<>"") then 
						sql="insert into basicCode (func, sys_Type, sys_Value ,transcode  ) values ( "&_
							"'"& UCASE(DB_TBLID) &"' , '"& UCASE(trim(tmpRec(i, j, 1))) &"' ,  "&_
							"N'"& tbldesc_str  &"' , '"& trim(tmpRec(i, j, 5))  &"'  ) " 							
					end if 		
				else
					sql="update basicCode set  "&_
							"sys_type='"& UCASE(trim(tmpRec(i, j, 1))) &"' , "&_
							"sys_Value=N'"& tbldesc_str  &"'  ,transcode='"&trim(tmpRec(i, j, 5))&"' "&_
							"where  autoid='"& trim(tmpRec(i, j, 4)) &"' " 				
				end if 
								
				response.write sql &"<BR>"
				conn.execute(sql)

				 y =  y & DB_TBLID & " - " &  tmpRec(i, j, 1) & " - " &   tbldesc_str &"<BR>"
			end if
		else
			if tmpRec(i, j , 0) = "del" then
				if tmpRec(i, j, 1) <> "" then
					sql2="delete  basicCode where  autoid='"& trim(tmpRec(i, j, 4)) &"' "	
					conn.Execute ( sql2 )
				y =  y & DB_TBLID & " - " &  tmpRec(i, j, 1) & " - " &   tbldesc_str  &" (DEL)<BR>"
				end if
			end if
		end if
	next
next

'response.end
 if conn.Errors.Count = 0 or err.number=0 then
	conn.CommitTrans
	Set Session("YTBAE0101EMP") = Nothing
    Set cmd = Nothing
	Set conn = Nothing

	Session("Title") = ""
	Session("Name") = "IMessage"
	Session("NO") = "Data Complete Success (OK) "
	Session("MessageCode") = "Success"
	Session("KeyValue") =  y
	Session("SubmitValue") = replace(session("pgname"),"<BR>",chr(13))
	Session("Action") = "YTBAE01.asp"
	Response.Redirect "IMessage.asp"
	Set conn = Nothing
 else
	conn.RollbackTrans
    Set Session("YTBAE0101EMP") = Nothing
    Set cmd = Nothing
	Set conn = Nothing

	Session("Title") = ""
	Session("Name") = "IMessage"
	Session("NO") = "Data Complete Fail (Error) "
	Session("MessageCode") = "Fail"
	Session("KeyValue") = ""
	Session("SubmitValue") = replace(session("pgname"),"<BR>",chr(13))
	Session("Action") = "YTBAE01.asp"
	Response.Redirect "IMessage.asp"
	Set conn = Nothing
 end if
 %>

<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%> 
<!-- #include file = "../GetSQLServerConnection.fun" --> 
<!----- #include file="../ADOINC.inc" ------>
<META HTTP-EQUIV="Content-Type" content="text/html; charset=utf-8">
<%
Response.Buffer = true
Response.Expires = 0
session.codepage="65001"
%>

<%
CurrentPage = request("CurrentPage")
TotalPage = request("TotalPage")
PageRec = request("PageRec")
gTotalPage = request("gTotalpage")
A1=trim(request("A1"))
'response.write "xxx="&A1

Set conn = GetSQLServerConnection()
tmpRec = Session("ADMIN01")

conn.BeginTrans

for i = 1 to gTotalPage 
	for j = 1 to PageRec 
		'response.write i &"-"&j &"<BR>"
		'response.write trim(tmpRec(i, j, 0)) &"<BR>"
		'response.write "4="&trim(tmpRec(i, j, 4)) &"<BR>"
		'response.write "1="&trim(tmpRec(i, j, 1)) &"<BR>"
		'response.write "2="&trim(tmpRec(i, j, 2)) &"<BR>"
		'response.write "3="&trim(tmpRec(i, j, 3)) &"<BR>"
		if tmpRec(i, j , 0) = "del" then		
			sql="delete  basicCode where  autoid='"& trim(tmpRec(i, j, 4)) &"' "	
			conn.execute(sql)	
		else 
			if REQUEST("op")(j)="upd" then 
				SYSVALUE=REQUEST("SYSVALUE")(j)
				if UCASE(trim(tmpRec(i, j, 1)))<>"" then 	
					 if UCASE(trim(tmpRec(i, j, 4)))<>"" then 							
						sql="update basicCode set func='"& UCASE(trim(tmpRec(i, j, 1))) &"' , "&_
							"sys_type='"& UCASE(trim(tmpRec(i, j, 2))) &"' , "&_
							"sys_Value='"& UCASE(SYSVALUE) &"'  "&_
							"where  autoid='"& trim(tmpRec(i, j, 4)) &"' " 
							conn.execute(sql)
							response.write sql &"<BR>"	
					elseif (trim(tmpRec(i, j, 2))<>"" and trim(tmpRec(i, j, 3))<>"") then 
						sql="insert into basicCode (func, sys_Type, sys_Value ) values ( "&_
							"'"& UCASE(trim(tmpRec(i, j, 1))) &"' , '"& UCASE(trim(tmpRec(i, j, 2))) &"' ,  "&_
							"'"& UCASE(SYSVALUE) &"' ) " 
							conn.execute(sql)
					response.write sql &"<BR>"
					end if 					
				end if 	
			end  if 
		end if 		
	next
next 
'response.end 
if  conn.errors.count=0  then 
	conn.CommitTrans
	response.redirect "yeaae0101.Fore.asp?A1="&A1
else	
	response.write "errors!!"
	Response.End 
end if 	 
%>

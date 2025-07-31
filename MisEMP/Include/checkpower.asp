<%
dim cn
'Set cn = Server.CreateObject("ADODB.Connection")
'cn.Open GetSQLServerConnection()
set cn = GetSQLServerConnection()
dim rs,sql,pgid,pgname

REMOTE_IP = Request.ServerVariables("REMOTE_ADDR")   
session("vnlogIP")=REMOTE_IP    


if trim(request("pgid"))<>"" then
  session("pgid")   =trim(request("pgid")) 
  session("pgname") =trim(request("pgname")) 
end if

'Response.Write trim(request("pgid"))&"<BR>"
'Response.End 

dim sqlstr,rs1
sqlstr="proc_sysusedrw '"& Session("NetUser") &"','"& session("pgid") &"'"
'Set rs1 = Server.CreateObject("ADODB.Recordset")
'response.write sqlstr 
Set rs1 = cn.execute(sqlstr)

'Response.Write RS1("W")&"<BR>"
'Response.Write RS1("R")
'Response.End
'dim mode
'if session("mode")="" then
if rs1.eof then 
	session("mode")=""
else
  if rs1("w")=1 then 
     session("mode")="W"     
  elseif rs1("r")=1 then
     session("mode")="R"     
  else 
   session("mode")=""
  end if
end if  
'end if
rs1.Close
cn.close 
set rs1=nothing
set cn=nothing
 
if session("mode")="" then    
  Response.Write "<br>"
  Response.Write "<br>"
  Response.Write "<br>"  
  Response.Write "<table width=630 ><tr><td align=center>"
  Response.Write "<b>UserID is empty or No limits, please Login again !!<br>Vao mang trong rong , hoac doi lau , hoac khong duoc su dong, <br>hay nhan nut nhap mang tu dau !!!</b>"
  Response.Write "</td></tr></table>"  
  Response.End  
end if


%>

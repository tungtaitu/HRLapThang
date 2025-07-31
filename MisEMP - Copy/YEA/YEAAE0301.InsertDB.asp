<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<%
Response.Buffer =True
%>
<!-------- #include file = "../GetSQLServerConnection.fun" --------->
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<%

  
 
Set conn = GetSQLServerConnection()	  

mode = request("mode")
Response.Write mode  &"<BR>"	
pid  = Trim(request("pid"))
upid = Trim(request("upid")) 
level = Trim(request("level"))
pname = Trim(request("pname"))
pnameVN = Trim(request("pnameVN"))
vpath = trim(replace(request("vpath"),",",""))  

totalpage= request("totalpage")
topage = request("topage")
'curr

if trim(mode)="addNew"  then 
	sql = "select * from SYSPROGRAM where program_id = '"& trim(pid) &"' "  
	Response.Write sql &"<BR>"	
	Set rst = Server.CreateObject("ADODB.Recordset") 
	rst.open sql, conn, 3, 3 
	if not rst.EOF then 
		sql="update SYSPROGRAM set program_Name = '"& Trim(pname) &"' , "&_
		    "layer_up = '"& Trim(upid) &"' , layer = '"& Trim(level) &"', "&_
		    "virtual_path = '"& vpath &"', "&_
		    "proname_vn = N'"& pnameVN &"' , "&_
		    "MDTM=GETDATE(), MUSER='"& SESSION("USERID") &"' "&_
		    "where program_id = '"& pid &"' "  
		conn.Execute (sql) 	 
		Response.Write sql 
		Response.Redirect "YEAAE0301.FORE.asp?schx=" & left(pid,1) 

	else
		SQL="INSERT INTO SYSPROGRAM (program_id, program_name, layer_up, layer, virtual_path, proname_vn, mdtm, muser  ) values ( "&_
		    "'"& pid &"', N'"& pname &"', '"& upid &"' , '"& level &"' , '"& vpath &"' , N'"& pnameVN &"' , "&_
		    "getdate(), '"& SESSION("USERID") &"'  ) " 
		conn.execute Sql     
		Response.Redirect "YEAAE0301.FORE.asp?schx=" & left(pid,1)
		'response.write sql
	end if 
	rst.close : set rst=nothing 
elseif   trim(mode)="delData"  then   
	'Response.Write request("delPid")
	sql = "delete SYSPROGRAM where program_id = '"& request("delPid") &"'  "
	conn.Execute (sql)  
	Response.Redirect "YEAAE0301.FORE.asp?schx=" & left(pid,1)
end if 
%>

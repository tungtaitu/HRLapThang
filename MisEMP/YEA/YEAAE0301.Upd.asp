<%@language=vbscript codepage=65001%>
<!-------- #include file = "../GetSQLServerConnection.fun" --------->

<html>
<head>
<meta http-equiv=Content-Type content="text/html; charset=utf-8">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
</head>

<%
response.Buffer = true
session.CodePage = 65001

%>

<%
CurrentPage = request("CurrentPage")
TotalPage = request("TotalPage")
PageRec = request("PageRec")
gTotalPage = request("gTotalpage")

Set conn = GetSQLServerConnection()
tmpRec = Session("SYSPRO01")

conn.BeginTrans

x = 0
y = ""
for i = 1 to PageRec 
	 TxtPROGRAM_ID = trim(request("TxtPROGRAM_ID")(i)) 
	 TxtPROGRAM_NAME = trim(request("TxtPROGRAM_NAME")(i))
	 TxtPROGRAM_NAME_VN = trim(request("TxtPROGRAM_NAME_VN")(i))
	 TxtLAYER_UP = trim(request("TxtLAYER_UP")(i))
	 TxtVIRTUAL_PATH = trim(request("TxtVIRTUAL_PATH")(i))
	 if request("opn")(i)="UPD" then 
		sql="update sysprogram set virtual_path='"& TxtVIRTUAL_PATH &"', mdtm=getdate() , "&_
			"muser='"& session("userID") &"', program_name='"& TxtPROGRAM_NAME &"', PRONAME_VN=N'"& TxtPROGRAM_NAME_VN &"' "&_
			"where program_id='"& TxtPROGRAM_ID &"' "
		conn.execute(Sql)	
		response.write sql &"<BR>"	
	end if	 
	
	if i =1 then 
		schx = left(trim(request("TxtPROGRAM_ID")(i)) ,1)
	end if 	
next 

'response.end 
response.write "1"
 if conn.Errors.Count = 0 then 
 	response.write "1"
	conn.CommitTrans
	Response.Redirect "YEAAE0301.FORE.asp?schx="&  schx 
	Set conn = Nothing 
 else
	conn.RollbackTrans
	Response.Redirect "YEAAE0301.FORE.asp"
 end if %>
</html>
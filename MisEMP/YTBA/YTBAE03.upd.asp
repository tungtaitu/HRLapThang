<%@language=vbscript CODEPAGE=65001%>
<meta HTTP-EQUIV="Content-Type" content="text/html; charset=UTF-8">
<!-------- #include file = "../../GetSQLServerConnection.fun" --------->
<!--#include FILE="../../include/ADOINC.inc"-->
<%
'Response.Buffer = true
'Response.Expires = 0


CurrentPage = request("CurrentPage")
TotalPage = request("TotalPage")
PageRec = request("PageRec")
DB_TBLID = request("DB_TBLID")
gTotalPage = request("gTotalpage")

proctype=request("proc1")

response.write  "proctype=" & proctype &"<BR>"
Set conn = GetSQLServerConnection()

'on error resume next
conn.BeginTrans
y=""

response.write  "xxxx=" & PageRec &"<BR>"
'response.end

for i = 1 to pagerec
	loai = request("loai")(i)
	zcq = request("zcq")(i)
	zcqname = request("zcqname")(i)
	zcqmail = request("zcqmail")(i)
	hcq01 = request("hcq01")(i)
	hcq01name = request("hcq01name")(i)
	hcq01mail = request("hcq01mail")(i)
	hcq02 = request("hcq02")(i)
	hcq02name = request("hcq02name")(i)
	hcq02mail = request("hcq02mail")(i)

	if loai<>""   then
		sql="select * from ytbmproc where proctype='"& proctype &"' and loai='"& loai &"'"
		Set RS = Server.CreateObject("ADODB.Recordset")
		rs.open sql, conn, 1, 3
		if rs.eof then
			sql="insert into ytbmproc ( proctype, loai, zcq, zcqname, zcqmail, "&_
					"hcq01,hcq01name, hcq01mail, hcq02,hcq02name, hcq02mail,mdtm, muser  ) values ( "&_
					"'"& proctype &"','"& loai &"','"& zcq &"','"& zcqname &"','"& zcqmail &"', "&_
					"'"& hcq01 &"','"& hcq01name &"','"& hcq01mail &"', "&_
					"'"& hcq02 &"','"& hcq02name &"','"& hcq02mail &"', getdate(), '"& session("userid") &"' )  "
		else
			sql="update ytbmproc set mdtm=getdate(), muser='"& session("userid") &"' , "&_
					"zcq='"& zcq &"', zcqname='"& zcqname &"',zcqmail='"& zcqmail &"' ,"&_
					"hcq01='"& hcq01 &"', hcq01name='"& hcq01name &"',hcq01mail='"& hcq01mail &"' ,"&_
					"hcq02='"& hcq02 &"', hcq02name='"& hcq02name &"',hcq02mail='"& hcq02mail &"' "&_
					"where proctype='"& proctype &"' and loai='"& loai &"'"
		end  if
		conn.execute(Sql)
	end if

	response.write sql &"<BR>"

next


'response.end
 if conn.Errors.Count = 0 or err.number=0 then
	conn.CommitTrans
	Set conn = Nothing
	response.redirect "ytbae03.asp"
 else
	conn.RollbackTrans
	Set conn = Nothing
	response.redirect "ytbae03.asp"
 end if
 %>

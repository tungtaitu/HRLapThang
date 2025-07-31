<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<!-- #include file = "../GetSQLServerConnection.fun" -->
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<%
session.codepage=65001
SELF = "YECE0601" 
Set conn = GetSQLServerConnection() 

CurrentPage = request("CurrentPage")
TotalPage = request("TotalPage")
PageRec = request("PageRec")
gTotalPage = request("gTotalpage") 
tmpRec = Session("yece0601B") 
JXYM = trim(request("JXYM")) 

F_WHSNO = request("F_WHSNO")
F_groupid=request("F_groupid") 
F_shift=request("F_shift") 
f_country=request("country") 
empid1=request("empid1")  

conn.BeginTrans 


for i = 1 to PageRec 
	jxYN=ucase(trim(request("jxyn")(i))) 
	if jxyn="" then jxyn="N"
	jxYNmemo=(trim(request("jxynmemo")(i))) 
	empid = tmpRec(1, i, 1 ) 	
	e_whsno = tmpRec(1, i, 7 ) 
	e_indat = tmpRec(1, i, 5 ) 
	country = tmpRec(1, i, 4 ) 
	c_aid = tmpRec(1, i, 15 )
	
	response.write i & jxyn &"-" & tmpRec(1, i, 1 ) &  tmpRec(1, i, 2 ) &  tmpRec(1, i, 15 ) & jxYNmemo &  "<BR>"
	sqlx = "select yymm, empid, isnull(jxyn,'') jxyn, isnull(memo,'') memo  from empJXYN where yymm='"& JXYM  &"' and empid='"& empid &"'   "
	Set rs = Server.CreateObject("ADODB.Recordset")  
	rs.open sqlx , conn, 3, 3
	if rs.eof then 
		set rs=nothing 
		sql="insert into empJXYN (yymm, empid, e_whsno, e_indat, country, jxyn, memo, mdtm, muser ) values ( "&_
				"'"&JXYM&"','"&empid&"','"&e_whsno&"','"&e_indat&"','"&country&"','"&jxyn&"',N'"&jxYNmemo&"',getdate(),'"&session("netuser")&"' ) " 
		conn.execute(Sql)		
		response.write sql &"<br>" 	
	else
		if ( isnull(rs("jxyn")) or rs("jxyn")<>jxyn ) or   rs("memo")<>jxYNmemo   then 
			sql="update empJXYN set jxyn='"& jxyn &"' , memo=N'"& jxynmemo &"' , mdtm=getdate(), "&_
					"muser='"&session("netuser")&"' where  yymm='"& JXYM  &"' and empid='"& empid &"' "
			conn.execute(Sql)
			response.write sql &"<br>" 	
		end if 	
	end if 
	
next 
'response.end  
response.write sqlx&"<br>"

if err.number = 0 then
	conn.CommitTrans 
	response.redirect self&".Fore.asp?sflag=T&yymm="& jxym &"&F_WHSNO="& F_WHSNO &"&f_groupid="& f_groupid &"&country="& f_country &"&empid1="& empid1 
else
	response.write "資料處理有誤!!Data Complete Error!!"
	response.end 
end if 	

%>
 
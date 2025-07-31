<%@LANGUAGE="VBSCRIPT" codepage=65001%>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<!-- #include file = "../GetSQLServerConnection.fun" -->
<!-- #include file="../ADOINC.inc" -->
<%
Response.Expires = 0
Response.Buffer = true
self="YECE0801"
session.codepage=65001
CurrentPage = request("CurrentPage")
TotalPage = request("TotalPage")
PageRec = request("PageRec")
gTotalPage = request("gTotalpage")
yymm=request("yymm")
workdays = request("workdays") 
'tmpRec = Session("YECE0801")
flag = request("flag")
Set CONN = GetSQLServerConnection()

conn.BeginTrans
x = 0
y = ""
YYMM=REQUEST("YYMM")

MMDAYS  = REQUEST("MMDAYS")
EMPIDSTR=""
for i = 1 to PageRec
	op = trim(request("op")(i))
 	empid = trim(request("empid")(i))
 	whsno = trim(request("whsno")(i))
 	country = trim(request("country")(i))
 	empname = trim(request("empname")(i))
 	
 	BB = request("bb")(i) 
 	if bb="" then bb=0 
 	
 	CV = request("cv")(i)
 	if cv="" then cv=0 
 	
 	rzM = request("rzM")(i)
 	if rzm="" then rzm=0 
 	
 	rzdays = request("rzdays")(i)
 	if rzdays="" then rzdays=0 
 	
 	jrm = request("jrm")(i)
 	if jrm="" then jrm=0 
 	
 	jrdays = request("jrdays")(i)
 	if jrdays="" then jrdays=0  
 	
 	tnkh = request("tnkh")(i)
 	if tnkh="" then tnkh=0 

 	zgm = request("zgm")(i)
 	if zgm="" then zgm=0  
 	 	
 	qita = request("qita")(i)
 	if qita="" then qita=0 
 	
 	totamt = request("totamt")(i)
 	if totamt="" then totamt=0 
 	
 	zkm = request("zkm")(i) 
 	if zkm="" then zkm=0 
 	
 	memo = trim(request("memo")(i))
 	dm = trim(request("dm")(i))
 	aid = trim(request("aid")(i))
	wptax =request("kwptax")(i) 
	wpbtien =request("wpbtien")(i) 
	if wpbtien="" then wpbtien= 0  
	if  wptax ="" then wptax=0 
 	
 	if ( empname<>"" and totamt > "0" )  then     		 		 			 
	 	if op="DEL" then 
	 		sql="delete salarywp  where  aid='"& aid &"' and yymm='"& yymm &"' "
	 		conn.execute(sql)
		elseif flag="Y" then  '關帳 
			sql="update salarywp set closeflag='Y' where yymm='"& yymm &"'  " 
			conn.execute(sql)
		else
			sqlx="select * from salarywp where yymm='"& yymm &"' and aid='"& aid &"'  " 
			Set rds = Server.CreateObject("ADODB.Recordset") 		
			rds.open sqlx, conn, 1, 3 	
				if rds.eof and  op<>"DEL" then  
				 	sql="insert into salarywp (yymm,workdays,empid,country,whsno,empname,BB,CV,rzM, "&_
				 		"rzdays,jrm,jrdays,tnkh,zgm, qita,dm,totAMT,zkm,memo,mdtm,muser,userIP,closeflag,wptax,phu ) values (  "&_
				 		"'"& yymm &"', '"& workdays &"', '"& empid &"','"& country &"','"& whsno &"',N'"& empname &"', "&_
				 		"'"& bb &"','"& cv &"','"& rzm &"','"& rzdays &"','"& jrm &"','"& jrdays &"','"& tnkh &"', '"& zgm &"',"&_
				 		"'"& qita &"','"& dm &"','"& totamt &"','"& zkm &"','"& memo &"',getdate(), '"& session("netuser") &"', "&_
				 		"'"& session("vnlogIP") &"','','"&wptax&"' ,'"& wpbtien &"') " 
				 	conn.execute(sql)	 	
			 	else
			 		sql="update salarywp set empid='"& empid &"'  , whsno='"& whsno &"' , country='"& country &"', empname=N'"& empname &"' , "&_
			 			"bb='"& bb &"', cv='"& cv &"', rzm='"& rzm &"', rzdays='"& rzdays &"' , jrm='"& jrm &"' , jrdays='"& jrdays &"' ,  "&_
			 			"tnkh='"& tnkh &"', zgm='"& zgm &"', qita='"& qita &"', dm='"& dm &"', totAMT='"& totAMT &"' , zkm='"& zkm &"' , memo='"& memo &"' ,  "&_
			 			"mdtm=getdate(), muser='"& session("netuser")  &"', userip='"& session("vnlogIP")  &"' "&_
			 			",wptax='"&wptax&"' , phu='"& wpbtien &"' where aid='"& aid &"' and yymm='"& yymm &"' "  
			 		conn.execute(sql)	
			 	end if 
		end if 	
	response.write sql &"<BR>" 		
	end if 	
	
 	'response.write i &"<BR>"
 	'response.write empname &"<BR>"
 	'response.write totamt &"<BR>" 
next
'response.clear 
'RESPONSE.END

if err.number=0 then
	conn.CommitTrans
	Set Session("YECE0801")=Nothing
	Set conn=Nothing
%>	<SCRIPT LANGUAGE=VBSCRIPT>
		ALERT"資料處理成功!!"&chr(13)&"DATA CommitTrans Success !!"
		OPEN "<%=self%>.asp","_self"
	</script>
<%
ELSE
	conn.RollbackTrans
%>	<SCRIPT LANGUAGE=VBSCRIPT>
		ALERT "資料處理失敗!!"&chr(13)&"DATA CommitTrans ERROR !!"
		OPEN "<%=self%>.asp", "_self"
	</script>
<%	response.end
END IF
%>
 
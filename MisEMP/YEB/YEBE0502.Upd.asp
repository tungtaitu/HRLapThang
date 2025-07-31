<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%> 
<!-- #include file = "../GetSQLServerConnection.fun" --> 
<!----- #include file="../ADOINC.inc" ------>
<META HTTP-EQUIV="Content-Type" content="text/html; charset=utf-8">
<html>

<head>
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="refresh">
<link rel="stylesheet" href="../Include/style.css" type="text/css">
<link rel="stylesheet" href="../Include/style2.css" type="text/css">
</head>
<body>
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
planyear = request("planyear") 


d1 = request("d1")
d2 = request("d2")
tim1 = request("tim1")
tim2 = request("tim2")
whour= request("whour")
studyName= request("studyName")
ssno = request("ssno")
studyGroup = request("studyGroup")
teacher = request("teacher")
nw = request("nw")
amt = request("amt")
dm = request("dm")
memo=request("memo")
pzj=request("pzj")

planyear = left(d1,4)
dim T(12)
for zz=1 to 12
	if zz=int(mid(d1,6,2)) then
		T(zz)="Y"
	else
		T(zz)=""
	end if 	
	response.write  zz & T(zz) & "<BR>"
next 

Set conn = GetSQLServerConnection() 
conn.BeginTrans

if ssno="" then 
	sqlx="exec GS_GetSSno 'Getssno', '"& planyear&"','' , '' "
	set rst=conn.execute(Sqlx)
	if rst("msg")="" then 
		ssno = rst("pno")			 					 			
	end if	
	set rst=nothing 
	sql="insert into studyPlan (yy, ssno,studyName,t1, t2, t3, t4, t5, t6,t7, t8, t9, t10,t11,t12, amt, dm, hhour, nw, memo,muser ) values ( "&_
		"'"& planyear &"', '"& ssno &"', N'"& studyName &"', '"& T(1) &"','"& T(2) &"','"& T(3) &"','"& T(4) &"','"& T(5) &"','"& T(6) &"', "&_
		"'"& T(7) &"','"& T(8) &"','"& T(9) &"','"& T(10) &"','"& T(11) &"','"& T(12) &"','"& amt &"','"& dm &"','"& whour &"', "&_
		"'"& nw &"',N'"& memo &"','"& session("NETUSER") &"' )" 
		'response.write sql &"<BR>" 	
		conn.execute(Sql) 
else
	sql="update studyPlan set studyName=N'"& studyName &"', amt='"& amt &"', dm='"& dm &"', nw='"& nw &"', mdtm=getdate(), muser='"& session("NETUSER") &"'   where ssno='"& ssno &"'"		
	conn.execute(Sql)
end if 		
'response.end 
'response.write gTotalPage &"<BR>"



x=0
y="" 
for i = 1 to 20 
	empid = request("empid")(i)
	groupid = request("groupid")(i)
	whsno = request("whsno")(i)
	country = request("country")(i)
	pzjno= request("pzjno")(i)
	'response.write i &"-"& empid &"<Br>"
	if empid<>"" then 
		sqln="insert into empstudy( whsno,ssno,empid,groupid,country,D1,D2,Tim1,Tim2,Whour,StudyName,StudyGroup,Teacher,status,memo,NW,mdtm,muser,pzjno) values ( "&_
			 "'"& whsno &"','"& ssno &"', '"& empid &"', '"& groupid &"', '"& country &"',  '"& D1 &"', '"& D2 &"', '"& Tim1 &"', '"& Tim2 &"', "&_
			 "'"& Whour &"','"& StudyName &"','"& StudyGroup &"','"& Teacher &"','','"& memo &"','"& nw &"',getdate(), "&_
			 "'"& session("NETUSER") &"','"& pzjno &"') "
		conn.execute(Sqln)	 
		'response.write sqln 
		x = x + 1 
	end if 
next 
 
'response.write Y 
'response.end 
if x=0 then
	conn.close
	set conn=nothing
	response.write "<table width=500 class=font9><tr><td align=center ><table width=400 class=font9><tr><td align=center>"
	response.write "<font class=txt12 color=red>無修改資料!!</font><BR><BR>"	
	response.write "</td></tr>"
	response.write "<tr><td align=center><BR><BR><a href='yebe0502.ASP?totalpage=0&planyear='"& planyear &"''><U><FONT COLOR=BLUE>回主畫面</FONT></U></A></td></tr></table>"
	response.write "</td></tr></table>"
	response.end 
else
	if  err.number=0  then 
		conn.CommitTrans
		conn.close
		set conn=nothing
		'response.redirect "yebe0501.Fore.asp?planyear="&planyear		
		response.write "<table width=550 class=font9><tr><td align=center ><table width=400 class=font9><tr><td align=center>"
		response.write "<center><font class=txt12 color=red>資料處理成功OK(SUCCESS)!!</font></center><BR><BR>"
		response.write y
		response.write "</td></tr>"
		response.write "<tr><td align=center><BR><BR><a href='yebe0502.ASP'><U><FONT COLOR=BLUE>回主畫面</FONT></U></A></td></tr></table>"
		response.write "</td></tr></table>"
		set session("YEBE0501")=nothing
		response.end 
		
	else	
		conn.close
		set conn=nothing
		response.write "errors!!" & err.description 
		response.write "<table width=500 class=font9><tr><td align=center ><table width=350 class=font9><tr><td align=center>"	
		response.write "</td></tr>"
		response.write "<tr><td><BR><BR><a href='yebe0502.ASP?totalpage=0&planyear='"& planyear &"''><U><FONT COLOR=BLUE>回主畫面</FONT></U></A></td></tr></table>"
		response.write "</td></tr></table>"	
		Response.End 
	end if 	 
end if 	
%>
</body>
</html>
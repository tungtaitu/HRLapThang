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
'response.write ssno
planyear = left(d1,4)
dim T(12)
for zz=1 to 12
	if zz=int(mid(d1,6,2)) then
		T(zz)="Y"
	else
		T(zz)=""
	end if 	
	'response.write  zz & T(zz) & "<BR>"
next 

Set conn = GetSQLServerConnection() 
conn.BeginTrans

x=0 
y="" 
for i = 1 to pagerec  
	aid = request("aid")(i)
	empid = request("empid")(i) 
	pzjno = trim(request("pzjno")(i))
	fensu= request("fensu")(i) 
	pjsts= trim(request("pjsts")(i))
	samt= request("samt")(i)
	if samt="" then samt=0
	pdm= request("pdm")(i) 
	'response.write i &"-"& op &"<Br>"	
	if empid<>"" then 
		  'if pjsts<>""  then 
		  	sql="update empstudy set pjsts='"& pjsts &"' , fensu='"& fensu &"', samt='"& samt &"' , dm='"& pdm &"',  "&_
		  		"pjmdtm=getdate(), pjuser='"& session("NETUSER") & "' where empid='"& empid &"' and  aid='"& aid &"' " 
		  	'response.write sql 	&"<br>"
		  	conn.execute(sql)
		  	x = x+1		  	
		  'end if 
	end if 	
next 
 
'response.write Y 
'response.end 
if x=0 then 
	conn.close
	set conn=nothing
	response.write "<table width=500 class=font9><tr><td align=center ><table width=350 class=font9><tr><td align=center>"
	response.write "<font class=txt12 color=red>無修改資料!!</font><BR><BR>"	
	response.write "</td></tr>"
	response.write "<tr><td align=center><BR><BR><a href='yebe0504.ASP?totalpage=0&planyear='"& planyear &"''><U><FONT COLOR=BLUE>回主畫面</FONT></U></A></td></tr></table>"
	response.write "</td></tr></table>"
	response.end 
else
	if  err.number=0  then 
		conn.CommitTrans
		conn.close
		set conn=nothing
		'response.redirect "yebe0501.Fore.asp?planyear="&planyear		
		response.write "<table width=550 class=txt><tr><td align=center ><table width=550 class=txt><tr><td  >"
		response.write "<center><font class=txt12 color=red>資料處理成功OK(SUCCESS)!!</font></center><BR><BR>"
		response.write y
		response.write "</td></tr>"
		response.write "<tr><td align=center><BR><BR><a href='yebe0504.foregnd.ASP?ssno="& ssno &"&d1="&D1&"&d2="&D2&"'><U><FONT COLOR=BLUE>回主畫面</FONT></U></A></td></tr></table>"
		response.write "</td></tr></table>"
		set session("YEBE0501")=nothing
		response.end 
		
	else	
		conn.close
		set conn=nothing
		response.write "errors!!"
		response.write "<table width=500 class=txt><tr><td align=center ><table width=350 class=font9><tr><td>"	
		response.write "</td></tr>"
		response.write "<tr><td><BR><BR><a href='yebe0504.ASP'><U><FONT COLOR=BLUE>回主畫面</FONT></U></A></td></tr></table>"
		response.write "</td></tr></table>"	
		Response.End 
	end if 	 
end if 	
%>
</body>
</html>
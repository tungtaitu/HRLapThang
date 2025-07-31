<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<!-- #include file = "../GetSQLServerConnection.fun" --> 
<!-- #include file="../ADOINC.inc" --> 
<%
SELF = "YEEE03" 


Set conn = GetSQLServerConnection()

pagerec = request("pagerec")  
for x = 1 to pagerec 
	njym=request("njym")(x)
	td1=request("td1")(x)
	td2=request("td2")(x)
	mdt=trim(request("mdt")(x)) 
	flags=trim(request("flags")(x)) 
	if flags="U" then 
		if mdt<>"" then 
			sql="update empNJYM_set set td1='"&td1&"' , td2='"&td2&"' , mdtm=getdate(), muser='"&session("netuser")&"' where njym='"&njym&"'"
			response.write sql &"<BR>"			
			conn.execute(sql)
		else
			if njym<>"" and td1<>"" and td2<>"" then 
			sql="insert into empNJYM_set (njym, td1, td2, mdtm, muser ) values ( '"&njym&"','"&td1&"','"&td2&"',getdate(), '"&session("netuser")&"' )"
			response.write sql  &"<BR>"			
			conn.execute(sql)
			end if
		end if 		
	end if 
next 
response.redirect "yeee03.back.asp"
%>
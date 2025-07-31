<%@LANGUAGE="VBSCRIPT"  codepage=65001%>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<!-- #include file = "../GetSQLServerConnection.fun" --> 
<!-- #include file="../ADOINC.inc" -->  

<%

Response.Expires = 0
Response.Buffer = true 
self="YEIE0201B"
 
khyear = request("khyear")
empid = request("empid")
khbid=request("khbid") 
empkhid=request("empkhid")  
pindex = request("pindex")

mylevel = "Z"  'request("mylevel")  
'mycqid = request("mycqid")  
mycqid = session("netuser")

cqmemos =request("cqmemos") 
end_Kj = request("end_Kj") 
totFs = request("totFs") 

Set conn = GetSQLServerConnection() 

conn.BeginTrans  

sqlx="delete  empkhbn_m  where   empkhid='"& empkhid &"' " 
conn.execute(Sqlx)  

sqly="delete EmpKHBN_D where empkhid='"&empkhid&"' and cqid='"&mycqid&"' "  
conn.execute(Sqly)   

if mylevel="Z" then 
	sql="INSERT INTO [YFYNET].[dbo].[EmpKHBN_M]([khbID], [empKHid], [years], [KH_UD], [empid], [z_fs], [z_kj], [Zcqid], "&_				
			"[zsts], [zcqmemos], [zmdtm] ) values ( "&_
			"'"&khbid&"','"&empkhid&"','"&left(empkhid,4)&"','"&mid(empkhid,5,1)&"','"&empid&"','"&totFs&"','"&end_Kj&"', "&_
			"'"&mycqid&"' ,'Y',N'"&cqmemos&"',getdate() )"
	conn.execute(Sql)		
elseif mylevel="J" then 
	sql="INSERT INTO [YFYNET].[dbo].[EmpKHBN_M]([khbID], [empKHid], [years], [KH_UD], [empid], [j_fs], [j_kj], [jcqid], "&_				
			"[jsts], [jcqmemos], [jmdtm] ) values ( "&_
			"'"&khbid&"','"&empkhid&"','"&left(empkhid,4)&"','"&mid(empkhid,5,1)&"','"&empid&"','"&totFs&"','"&end_Kj&"', "&_
			"'"&mycqid&"' ,'Y',N'"&cqmemos&"',getdate() )"
	conn.execute(Sql)		
elseif mylevel="H" then 	
	sql="INSERT INTO [YFYNET].[dbo].[EmpKHBN_M]([khbID], [empKHid], [years], [KH_UD], [empid], [h_fs], [h_kj], [hcqid], "&_				
			"[hsts], [hcqmemos], [hmdtm] ) values ( "&_
			"'"&khbid&"','"&empkhid&"','"&left(empkhid,4)&"','"&mid(empkhid,5,1)&"','"&empid&"','"&totFs&"','"&end_Kj&"', "&_
			"'"&mycqid&"' ,'Y',N'"&cqmemos&"',getdate() )"	
	conn.execute(Sql)		
else
	sql="INSERT INTO [YFYNET].[dbo].[EmpKHBN_M]([khbID], [empKHid], [years], [KH_UD], [empid], fensu, grade , mdtm, muser )  values ( "&_
			"'"&khbid&"','"&empkhid&"','"&left(empkhid,4)&"','"&mid(empkhid,5,1)&"','"&empid&"','"&totFs&"','"&end_Kj&"', "&_
			"getdate(),'"&session("netuser")&"' )"
	conn.execute(Sql)		
end if  

response.write sql  &"<BR>"
 'response.end 
 
xx=0
y=0
PageRec = request("PageRec") 

for x = 1 to PageRec
	sttno = request("sttno")(x) 
	cqfensu = trim(request("cqfensu")(x)) 
	
	if len(sttno)=3 then  		
		sql=" INSERT INTO [YFYNET].[dbo].[EmpKHBN_D] "&_
				"([empKHid], [whsno], [years], [ud], [empid], [sttno], [cqid], [cq_level], [fensu], [mdtm], [muser] ) values ( "&_
				"'"&empkhid&"','"&session("mywhsno")&"','"&left(empkhid,4)&"','"&mid(empkhid,5,1)&"','"&empid&"','"&sttno&"', "&_
				"'"&mycqid&"','"&mylevel&"','"&cqfensu&"',getdate(),'"&session("netuser")&"' ) " 
		response.write sql &"<BR>"
		conn.execute(sql)
	end if 	

		
next	 

'response.write "xx=" & xx
'response.end 

if ( conn.Errors.Count = 0 or err.number=0  ) then 
	conn.CommitTrans
	Set conn = Nothing 
ELSE
	conn.RollbackTrans	 
	response.end 	
END IF  
%>
 
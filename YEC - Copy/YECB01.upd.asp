<%@LANGUAGE="VBSCRIPT" CODEPAGE=65001%>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<!-- #include file = "../GetSQLServerConnection.fun" --> 
<!-- #include file="../ADOINC.inc" -->  
<%
 

Set CONN = GetSQLServerConnection()  
closeYM = request("closeYM") 
whsno=split(replace(request("whsno")," ","")&",",",") 
 
conn.BeginTrans
x = 0
y = ""  
YYMM=REQUEST("YYMM")
EMPIDSTR="" 
if session("netuser")<>""   then  
	for x = 1 to ubound(whsno)
		F_whsno=whsno(x-1) 
		if F_whsno="" then 
			sqlx="select * from closeYM where closeYM='"& closeYM &"' and isnull(whsno,'')='' "
			Set rst = Server.CreateObject("ADODB.Recordset")   		
			rst.open sqlx, conn, 1, 3
			if rst.eof then 
				sql="insert into closeYM ( closeYM, mdtm, muser ) values ('"& closeYM &"', getdate(), '"& session("netuser") &"') "				
			else
				sql="update closeYM set mdtm=getdate(), muser='"& session("netuser") &"' where closeYM='"& closeYM &"' and isnull(whsno,'')='' "	
			end if 
			conn.execute(Sql)
			'response.write "1=" & sql & "<BR>" 
			set rst=nothing
		else
			sqlx="select * from closeYM where closeYM='"& closeYM &"' and isnull(whsno,'')='"& F_whsno &"' "
			Set rst = Server.CreateObject("ADODB.Recordset")   		
			rst.open sqlx, conn, 1, 3
			if rst.eof then 
				sql="insert into closeYM ( closeYM, whsno, mdtm, muser ) values ('"& closeYM &"', '"& F_whsno &"', getdate(), '"& session("netuser") &"') "
			else	
				sql="update closeYM set mdtm=getdate(), muser='"& session("netuser") &"' where closeYM='"& closeYM &"' and isnull(whsno,'')='"& F_whsno &"' "	
			end if 		
			conn.execute(Sql)
			'response.write "2=" & sql & "<BR>" 
			set rst=nothing
		end if  
	next  
end if  

if closeYM>="201105" then 
	sqla="delete empdsalary_bak where yymm='"&closeYM&"' "
	conn.execute(sqla) 
	sqlb="insert into empdsalary_bak select getdate(),'"&session("netuser")&"' , *  from empdsalary  where yymm='"&closeYM&"'  " 
	conn.execute(sqlb)  
	'response.write sqla &"<BR>"
	'response.write sqlb &"<BR>"
	'RESPONSE.END 
end if 

if conn.Errors.Count = 0 or err.number=0 then 
	conn.CommitTrans
	Set Session("empsalary01") = Nothing 	  	
	Set conn = Nothing
%>	<SCRIPT LANGUAGE=VBSCRIPT>
		ALERT "資料處理成功!!"&chr(13)&"DATA CommitTrans Success !!"
		OPEN "YECB01.asp" , "_self" 
	</script>	
<% 
ELSE
	conn.RollbackTrans	
%>	<SCRIPT LANGUAGE=VBSCRIPT>
		ALERT "資料處理失敗!!"&chr(13)&"DATA CommitTrans ERROR !!"
		OPEN "YECB01.asp" , "_self" 
	</script>	
<%	response.end 
END IF 
%>
 
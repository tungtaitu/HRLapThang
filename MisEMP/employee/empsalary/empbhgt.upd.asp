<%@LANGUAGE="VBSCRIPT" %>
<!-- #include file = "../../GetSQLServerConnection.fun" -->
<!-- #include file="../../ADOINC.inc" -->
<%
Response.Expires = 0
Response.Buffer = true

CurrentPage = request("CurrentPage")
TotalPage = request("TotalPage")
PageRec = request("PageRec")
gTotalPage = request("gTotalpage")

Set conn = GetSQLServerConnection()
tmpRec = Session("empBHGTD")

Set CONN = GetSQLServerConnection()

conn.BeginTrans
x = 0
y = ""
YYMM=REQUEST("YYMM")

MMDAYS  = REQUEST("MMDAYS")
EMPIDSTR=""
for i = 1 to TotalPage
	for j = 1 to PageRec
		'RESPONSE.WRITE TotalPage &"<br>"
		'RESPONSE.WRITE PageRec &"<br>" 		
			
		if trim(tmpRec(i, j, 1))<>"" then			
			
			sqlx="update bempj set job ='EV1' where empid='"& trim(tmpRec(i, j, 1)) &"' and "&_
				 "whsno='"& trim(tmpRec(i, j, 7)) &"' and job='EV0' and yymm>='"& yymm &"' "
			conn.execute(sqlx)  
			
			sql="select * from empbhgt where yymm='"& yymm &"' and empid='"&  trim(tmpRec(i, j, 1)) &"'  " 
			Set rs = Server.CreateObject("ADODB.Recordset")
			RS.OPEN SQL, CONN, 3, 3  
			if not rs.eof then 
				sql="update empbhgt set BHXH5='"& trim(tmpRec(i, j, 21)) &"', BHYT1='"& trim(tmpRec(i, j, 22)) &"', "&_
					"BHTN1='"& trim(tmpRec(i, j, 32)) &"', "&_
					"BHTOT='"& trim(tmpRec(i, j, 28)) &"', GTAMT='"& trim(tmpRec(i, j, 23)) &"' , "&_
					"whsno='"& trim(tmpRec(i, j, 7)) &"', groupid='"& trim(tmpRec(i, j, 9)) &"' , "&_
					"BB='"& trim(tmpRec(i, j, 20)) &"' , bhdat='"& trim(tmpRec(i, j, 26)) &"' , "&_
					"GTDAT='"& trim(tmpRec(i, j, 27)) &"', CLOSEYN='N', muser='"& session("NETUSER")&"' , "&_
					"kh1='"&  trim(tmpRec(i, j, 29)) &"', chanjia='"&  trim(tmpRec(i, j, 30)) &"' , memo='"& trim(tmpRec(i, j, 31)) &"' "&_
					"where yymm='"& yymm &"' and empid='"& trim(tmpRec(i, j, 1))  &"' "  
					
					x= x+1  
				conn.execute(sql) 
				response.write sql&"<BR>" 	
			else
				sql="insert into  empbhgt (empid, whsno, groupid, bb, bhdat, gtdat, bhxh5, bhyt1, bhtn1, bhtot, gtamt, yymm, mdtm, muser, closeYN , kh1, chanjia , memo ) "&_
					"values ( "&_
					"'"& trim(tmpRec(i, j, 1))&"', '"& trim(tmpRec(i, j, 7))&"', '"& trim(tmpRec(i, j, 9))&"', '"& trim(tmpRec(i, j, 20))&"', "&_
					"'"& trim(tmpRec(i, j, 26))&"', '"& trim(tmpRec(i, j, 27))&"', '"& trim(tmpRec(i, j, 21))&"','"& trim(tmpRec(i, j, 22))&"', "&_
					"'"& trim(tmpRec(i, j, 32))&"', '"& trim(tmpRec(i, j, 28))&"', '"& trim(tmpRec(i, j, 23))&"',  '"& yymm &"', getdate(), '"& session("NETUSER")&"' , 'N', "&_
					"'"& trim(tmpRec(i, j, 29))&"','"& trim(tmpRec(i, j, 30))&"','"& trim(tmpRec(i, j, 31))&"' ) " 
					
					x= x+1   
				conn.execute(sql) 	
				response.write sql&"<BR>"
			end if  
		end if   
		
	next
next

'RESPONSE.END

if conn.Errors.Count = 0 or err.number=0  then
	conn.CommitTrans
	Set Session("empBHGTD") = Nothing
	Set conn = Nothing
%>	<SCRIPT LANGUAGE=VBSCRIPT>
		ALERT "資料處理成功!!"&"<%=X%>"&"筆"&chr(13)&"DATA CommitTrans Success !!"
		OPEN "empbhgt.asp" , "_self"
	</script>
<%
ELSE
	conn.RollbackTrans
%>	<SCRIPT LANGUAGE=VBSCRIPT>
		ALERT "資料處理失敗!!"&chr(13)&"DATA CommitTrans ERROR !!"
		OPEN "empbhgt.asp" , "_self"
	</script>
<%	response.end
END IF
%>
 
<%@LANGUAGE="VBSCRIPT" codepage=65001%>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<!-- #include file = "../GetSQLServerConnection.fun" -->
<!-- #include file="../ADOINC.inc" -->
<%
Response.Expires = 0
Response.Buffer = true

CurrentPage = request("CurrentPage")
TotalPage = request("TotalPage")
PageRec = request("PageRec")
gTotalPage = request("gTotalpage")

Set conn = GetSQLServerConnection()
tmpRec = Session("empBHGTD")
a1 = session("a1cols") 

Set CONN = GetSQLServerConnection()

'response.write trim(tmpRec(i, j, 21))
'response.end
conn.BeginTrans
x = 0
y = ""
YYMM=REQUEST("YYMM")
allcols = request("allcols")
MMDAYS  = REQUEST("MMDAYS")
EMPIDSTR=""
for i = 1 to TotalPage
	for j = 1 to PageRec
		'RESPONSE.WRITE TotalPage &"<br>"
		'RESPONSE.WRITE PageRec &"<br>" 		
			
		if trim(tmpRec(i, j, 1))<>"" then			 
		
			if  trim(tmpRec(i, j, 7)) ="DN" then 
				n_phu = trim(tmpRec(i, j, 36))
			else
				n_phu = 0 
			end if   
			
			sqlx="update bempj set job ='EV1' where empid='"& trim(tmpRec(i, j, 1)) &"' and "&_
				 "whsno='"& trim(tmpRec(i, j, 7)) &"' and job='EV0' and yymm>='"& yymm &"' "
			conn.execute(sqlx)   
			
			BHTOT = cdbl(trim(tmpRec(i, j, 21)))+cdbl(trim(tmpRec(i, j, 22)))+cdbl(trim(tmpRec(i, j, 32))) 
			
			sql="select * from empbhgt where yymm='"& yymm &"' and empid='"&  trim(tmpRec(i, j, 1)) &"'  " 
			Set rs = Server.CreateObject("ADODB.Recordset")
			RS.OPEN SQL, CONN, 3, 3  
			if not rs.eof then 
				sql="update empbhgt set BHXH5='"& trim(tmpRec(i, j, 21)) &"', BHYT1='"& trim(tmpRec(i, j, 22)) &"', "&_
					"BHTN1='"& trim(tmpRec(i, j, 32)) &"', "&_
					"BHTOT='"& BHTOT &"', GTAMT='"& trim(tmpRec(i, j, 23)) &"' , "&_
					"whsno='"& trim(tmpRec(i, j, 7)) &"', groupid='"& trim(tmpRec(i, j, 9)) &"' , "&_
					"BB='"& trim(tmpRec(i, j, 20)) &"' , bhdat='"& trim(tmpRec(i, j, 26)) &"' , "&_
					"GTDAT='"& trim(tmpRec(i, j, 27)) &"', CLOSEYN='N', muser='"& session("NETUSER")&"' , "&_
					"kh1='"&  trim(tmpRec(i, j, 29)) &"', chanjia='"&  trim(tmpRec(i, j, 30)) &"' , memo='"& trim(tmpRec(i, j, 31)) &"' ,  "&_
					"CV='"&  trim(tmpRec(i, j, 35)) &"', phu='"& n_phu &"' , BHP='"&  trim(tmpRec(i, j, 34)) &"', "&_ 
					"colsname='"& trim(tmpRec(i, j, 36+cdbl(allcols)+1)) &"' " 
				for z1 = 1 to 7 
					if z1<=cdbl(allcols) then 
						sql=sql & ", C"& z1 &"="&  trim(tmpRec(i, j, 36+cdbl(z1)))  
					else
						sql=sql & ", C"& z1 &"=0" 
					end if 
				next 
				sql=sql & " where yymm='"& yymm &"' and empid='"& trim(tmpRec(i, j, 1))  &"' "  
					
					x= x+1  
				conn.execute(sql) 
				response.write sql&"<BR>" 	
			else
				sql="insert into  empbhgt (empid, whsno, groupid, bb, bhdat, gtdat, bhxh5, bhyt1, bhtn1, bhtot, gtamt, yymm, "&_
					"mdtm, muser, closeYN , kh1, chanjia , memo , cv, phu, bhp , colsname   " 
				for z2 = 1 to 7 					
						sql=sql & ", C"& z2  					 	
				next 	
				sql=sql&") values ( "&_
					"'"& trim(tmpRec(i, j, 1))&"', '"& trim(tmpRec(i, j, 7))&"', '"& trim(tmpRec(i, j, 9))&"', '"& trim(tmpRec(i, j, 20))&"', "&_
					"'"& trim(tmpRec(i, j, 26))&"', '"& trim(tmpRec(i, j, 27))&"', '"& trim(tmpRec(i, j, 21))&"','"& trim(tmpRec(i, j, 22))&"', "&_
					"'"& trim(tmpRec(i, j, 32))&"', '"& BHTOT &"', '"& trim(tmpRec(i, j, 23))&"',  '"& yymm &"', getdate(), '"& session("NETUSER")&"' , 'N', "&_
					"'"& trim(tmpRec(i, j, 29))&"','"& trim(tmpRec(i, j, 30))&"','"& trim(tmpRec(i, j, 31))&"', "&_
					"'"& trim(tmpRec(i, j, 35))&"','"& n_phu &"','"& trim(tmpRec(i, j, 34))&"', '"& trim(tmpRec(i, j, 36+cdbl(allcols)+1)) &"' " 
				for z3 = 1 to 7 
					if z3 <=cdbl(allcols) then 
						sql=sql & "," & trim(tmpRec(i, j, 36+z3)) 
					else	
						sql=sql & ", 0 " 
					end if 	
				next 	
				sql=sql & " ) "	
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
		OPEN "yece0701.asp" , "_self"
	</script>
<%
ELSE
	conn.RollbackTrans
%>	<SCRIPT LANGUAGE=VBSCRIPT>
		ALERT "資料處理失敗!!"&chr(13)&"DATA CommitTrans ERROR !!"
		OPEN "yece0701.asp" , "_self"
	</script>
<%	response.end
END IF
%>
 
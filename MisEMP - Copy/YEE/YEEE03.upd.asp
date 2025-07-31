<%@LANGUAGE="VBSCRIPT" codepage=65001%>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<!-- #include file = "../GetSQLServerConnection.fun" -->
<!-- #include file="../ADOINC.inc" -->
<%
Response.Expires = 0
Response.Buffer = true

self="YEEE03"
CurrentPage = request("CurrentPage")
TotalPage = request("TotalPage")
PageRec = request("PageRec")
gTotalPage = request("gTotalpage")

Set conn = GetSQLServerConnection()
tmpRec = Session("YEEE03B")
 
Set CONN = GetSQLServerConnection()

if session("netuser")="" then 
	response.write "使用者帳號為空,請重新登入!!<br>"
	response.write "Ma so khong ton tai, Xin Hay Dang Nhap Lai He Thong!!!!!<br>"
	response.end
end if 	

conn.BeginTrans
x = 0
y = ""
YYMM=REQUEST("YYMM") 

for i = 1 to TotalPage
	for j = 1 to PageRec
		empid = tmpRec(i, j, 1) 	
		indat = tmpRec(i, j, 5) 	
		outdat = tmpRec(i, j, 6) 	
		job = tmpRec(i, j, 7) 		
		whsno = tmpRec(i, j, 8) 	
		groupid = tmpRec(i, j, 9)	
		country = tmpRec(i, j, 4) 
		all_JiaE_hr = tmpRec(i, j, 12)
		sys_TXH = tmpRec(i, j, 25)	'系統計算 年假時數		    
		sys_TXD	=	tmpRec(i, j, 26)	'系統計算 年假天數		    
		'hr_money = replace(trim(tmpRec(i, j, 32)),",","")	'時薪(年度平均)	
		hr_money = replace(trim(tmpRec(i, j, 20)),",","")	'時薪(年度平均)			
		
		TZ_txD = trim(request("njtz")(j))  '調整年假
		if TZ_txD="" then TZ_txD = 0 
		Atz_TXD = trim(request("txdays")(j))  '調整後實際年假天數
		Atz_TXH = trim(request("tx_hr")(j))  '調整後實際年假時數
		NowTxH = trim(request("nowtx")(j))  '調整後剩餘年假(時數) 
		NJ_amt = replace(trim(request("njAmt")(j)),",","")
		if NJ_amt="" then NJ_amt = 0  
		txmemo = trim(tmpRec(i, j, 36))
		
		sqlx="select * from EMPTXAMT where TYear='"& yymm &"'  and empid = '"& empid &"' "
		Set rs = Server.CreateObject("ADODB.Recordset")   
		rs.open sqlx , conn, 1, 3 
		if rs.eof then 
			response.write  j &":_1 = " & empid  & "<br>"
			sql="INSERT INTO [dbo].[EMPTXAMT]( [TYear], [whsno], [country], [empid], [indat], [outdat], [groupid], [job], [all_JiaE_hr], "&_
					"[sys_TXD], [sys_TXH], [TZ_txD], [Atz_TXD], [Atz_TXH], [NowTxH], [hr_money], [NJ_amt], memo, [mdtm], [muser], [keyindate], [keyinby]) "&_
					"values ( "&_
					"'"&yymm&"','"&whsno&"','"&country&"','"&empid&"','"&indat&"','"&outdat&"','"&groupid&"','"&job&"','"&all_JiaE_hr&"', "&_
					"'"&sys_TXD&"','"&sys_TXH&"','"&TZ_txD&"','"&Atz_TXD&"','"&Atz_TXH&"','"&NowTxH&"','"&hr_money&"','"& NJ_amt &"', "&_
					"N'"&txmemo&"',getdate(),'"& session("netuser") &"',getdate(),'"& session("netuser") &"' ) " 
			x = x + 1  		
			response.write sql &"<br>"
			conn.execute(Sql)
		else 
			response.write  j &":_2 = " & empid  & "<br>"
			sql="UPDATE [dbo].[EMPTXAMT] set "&_
					"whsno='"& whsno &"' , country='"& country &"' ,  indat='"& indat &"' , outdat='"& outdat &"' , "&_
					"groupid='"& groupid &"', job='"& job &"' , [all_JiaE_hr]='"& all_JiaE_hr &"' , [sys_TXD]='"& sys_TXD &"', [sys_TXH]='"&sys_TXH&"', "&_
					"[TZ_txD]='"&TZ_txD&"', [Atz_TXD]='"&Atz_TXD&"', [Atz_TXH]='"& Atz_TXH &"',[NowTxH]='"&NowTxH&"', [hr_money]='"&hr_money&"', "&_					
					"[NJ_amt]='"&NJ_amt&"', memo=N'"&txmemo&"', [mdtm]=getdate(), [muser]='"& session("netuser") &"' "&_
					"WHERE  TYear='"& yymm &"'  and empid = '"& empid &"' " 
			x = x + 1 		
			response.write sql &"<br>"
			conn.execute(Sql)
		end if 
		
	next
next
response.write err.number &"<BR>"
response.write conn.errors.count &"<BR>"


for g =0 to conn.errors.count-1
	response.write conn.errors.item(g)&"<br>"
	response.write Err.Description
next  

'RESPONSE.END

if err.number = 0 then
	conn.CommitTrans
	Set Session("YEEE03") = Nothing
	Set conn = Nothing 
%>	<SCRIPT LANGUAGE=VBSCRIPT>
		ALERT "資料處理成功!!"&"<%=X%>"&"筆"&chr(13)&"DATA CommitTrans Success !!"
		OPEN "<%=self%>.asp" , "_self"
	</script>
<% 
ELSE
	conn.RollbackTrans 
%>	<SCRIPT LANGUAGE=VBSCRIPT>
		ALERT "資料處理失敗!!"&chr(13)&"DATA CommitTrans ERROR !!"
		OPEN "empsalaryHW.asp" , "_self"
	</script>
<%  response.end
END IF
%>
 
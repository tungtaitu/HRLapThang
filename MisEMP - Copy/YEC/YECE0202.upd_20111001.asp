<%@LANGUAGE="VBSCRIPT" codepage=65001%>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<!-- #include file = "../GetSQLServerConnection.fun" -->
<!-- #include file="../ADOINC.inc" -->
<%
self="yece0202" 
CurrentPage = request("CurrentPage")
TotalPage = request("TotalPage")
PageRec = request("PageRec")
gTotalPage = request("gTotalpage")

Set conn = GetSQLServerConnection()
tmpRec = Session("YECE0202B")
cfg  = request("cfg") 

firstday  = request("calcdat")
endday = request("ccdt") 

response.write "TotalPage=" & TotalPage &"<BR>"
response.write "PageRec=" & PageRec &"<BR>" 
response.write "cfg=" & cfg 
'response.end 

conn.BeginTrans
x = 0
y = ""
YYMM=REQUEST("YYMM")

MMDAYS  = REQUEST("MMDAYS")
EMPIDSTR=""
for i = 1 to TotalPage
	for j = 1 to PageRec
		RESPONSE.WRITE TotalPage &"<br>"
		RESPONSE.WRITE PageRec &"<br>" 
		if trim(tmpRec(i, j, 1))<>"" then
			empid=trim(tmpRec(i, j, 1))
			empwhsno = trim(tmpRec(i, j, 7))
			indat = trim(tmpRec(i, j, 5))
			outdate = trim(tmpRec(i, j, 17))
			endJbdat = trim(tmpRec(i, j, 18))
			h1=trim(tmpRec(i, j, 21))
			if h1="" then h1=0 
			h2=trim(tmpRec(i, j, 22))
			if h2="" then h2=0 
			h3=trim(tmpRec(i, j, 23))			
			if h3="" then h3=0 
			ov_h1=trim(tmpRec(i, j, 24))
			if ov_h1="" then ov_h1=0 
			ov_h2=trim(tmpRec(i, j, 25))
			if ov_h2="" then ov_h2=0 
			ov_h3=trim(tmpRec(i, j, 26))
			if ov_h3="" then ov_h3=0 
			B3=trim(tmpRec(i, j, 27))
			if B3="" then B3=0 
			bb = trim(tmpRec(i, j, 28))
			cv = trim(tmpRec(i, j, 29))
			phu = trim(tmpRec(i, j, 30))
			hourM = trim(tmpRec(i, j, 31))
			jbm = trim(tmpRec(i, j, 32))
			ov_jbm = trim(tmpRec(i, j, 33))
			ov_b3 =trim(tmpRec(i, j, 36))
			if ov_b3="" then ov_b3=0 
			
			sql="select isnull(flag,'') N_flag, * from empJBtim where empid='"& empid &"' and yymm='"& YYMM &"' "
			Set rst = Server.CreateObject("ADODB.Recordset")
			rst.open sql, conn, 1, 3 
			if not rst.eof then 
				if rst("n_flag")<>"Y" then 
					sql="update empJBtim set h1='"& h1 &"' , h2='"& h2 &"', h3='"& h3 &"', b3='"& b3 &"' , "&_
						"ov_h1='"& ov_h1 &"' , ov_h2='"& ov_h2 &"', ov_h3='"& ov_h3 &"' , mdtm=getdate() , muser='"& session("netuser") &"',  "&_
						"bb='"& bb &"', cv='"& cv &"', phu='"& phu &"', jbm='"& jbm &"' , ov_jbm='"& ov_jbm &"', endjbdat='"& endJbdat &"' , "&_
						"hourM='"& hourM &"' , ov_b3='"& ov_b3 &"' where empid='"& empid &"' and yymm='"& YYMM &"' " 
					conn.execute(sql)	
					response.write "1=" & sql &"<BR>"
					x = x + 1 
				end if 
			else
				sql="insert into empJBtim ( empwhsno, empid,indat, outdate, yymm, endJbdat , "&_
					"h1, h2, h3, b3, ov_h1, ov_h2, ov_h3, mdtm, muser, bb, cv, phu, jbm, ov_jbm, hourM , ov_b3  ) values ( "&_
					"'"& empwhsno&"','"& empid&"','"& indat&"','"& outdate&"','"& yymm&"','"& endJbdat&"', "&_
					"'"& h1&"','"& h2&"','"& h3&"','"& B3&"','"& ov_h1&"','"&ov_h2&"','"& ov_h3&"', "&_
					"getdate(), '"& session("netuser") &"','"& bb &"','"& cv &"','"& phu &"','"& jbm &"','"& ov_jbm &"','"& hourM &"','"& ov_b3 &"' )  "
				conn.execute(Sql)	
				response.write "2=" & sql &"<br>"				
				x = x+ 1 
			end if 	
		'else
			'response.write "???"
		END IF
	next
next
'response.end 
response.write err.number &"<BR>"
response.write conn.errors.count &"<BR>"

for g =0 to conn.errors.count-1
	response.write conn.errors.item(g)&"<br>"
	response.write Err.Description
next  
response.clear
'RESPONSE.END

if err.number = 0 then
	conn.CommitTrans
	Set Session("YECE0202B") = Nothing
	Set conn = Nothing      
%>	<SCRIPT LANGUAGE=VBSCRIPT>
		ALERT "資料處理成功!!"& "<%=X%>" & "筆" & chr(13) &"DATA CommitTrans Success !!"
		OPEN "<%=self%>.asp" , "_parent"
	</script>
<%'  response.end  
ELSE
	Set Session("YECE0202B") = Nothing
	Set conn = Nothing     
	conn.RollbackTrans       
%>	<SCRIPT LANGUAGE=VBSCRIPT>
		ALERT "資料處理失敗!!"&chr(13)&"DATA CommitTrans ERROR !!"
		OPEN "<%=self%>.asp" , "_parent"
	</script>
<%
	response.end
END IF
%>
 
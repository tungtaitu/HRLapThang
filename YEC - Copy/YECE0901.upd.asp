<%@LANGUAGE="VBSCRIPT" codepage=65001%>
<!-- #include file = "../GetSQLServerConnection.fun" -->
<!-- #include file="../ADOINC.inc" -->
<%
Response.Expires = 0
Response.Buffer = true
self="YECE0901"
CurrentPage = request("CurrentPage")
TotalPage = request("TotalPage")
PageRec = request("PageRec")
gTotalPage = request("gTotalpage")

 
tmpRec = Session("YECE0901F")

Set CONN = GetSQLServerConnection()

conn.BeginTrans
x = 0
y = ""
YYMM=REQUEST("YYMM")

MMDAYS  = REQUEST("MMDAYS")
EMPIDSTR=""
for i = 1 to TotalPage
	for j = 1 to PageRec 
			empid=trim(tmpRec(i, j,1)) 
			
			if  (tmpRec(i, j,23))="" THEN
				XIANM  = 0    '現金
			ELSE
				XIANM = tmpRec(i,j,23)
			END IF

			if  (tmpRec(i, j,24)) ="" THEN
				ZHUANM = 0   '轉款
			ELSE
				ZHUANM = tmpRec(i,j,24)
			END IF
			sql="update empdsalary set flag='Y', zhuanM='"& ZHUANM &"' , XIANM='"& XIANM &"' where yymm='"& yymm &"' and empid='"& empid &"' " 
			conn.execute(Sql)
			
			sqlx="update empdsalary_bak set  flag='Y', zhuanM='"& ZHUANM &"' , XIANM='"& XIANM &"' where yymm='"& yymm &"' and empid='"& empid &"'" 
			conn.execute(Sqlx) 
			
			X=X+1
 			response.write sql&"<BR>"
 			response.write sqlx&"<BR>"
	next
next

'RESPONSE.END

if err.number = 0 then
	conn.CommitTrans
	Set Session("YECE0901F") = Nothing
	Set conn = Nothing
%>	<SCRIPT LANGUAGE=VBSCRIPT>
		ALERT "資料處理成功!!"&chr(13)&"DATA CommitTrans Success !!"
		OPEN "<%=self%>.asp" , "_self"
	</script>
<%
ELSE
	conn.RollbackTrans
%>	<SCRIPT LANGUAGE=VBSCRIPT>
		ALERT "資料處理失敗!!"&chr(13)&"DATA CommitTrans ERROR !!"
		OPEN "<%=self%>.asp" , "_self"
	</script>
<%	response.end
END IF
%>
 
<%@LANGUAGE="VBSCRIPT" codepage=65001%>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<!-- #include file = "../GetSQLServerConnection.fun" -->
<!-- #include file="../ADOINC.inc" -->
<%
Response.Expires = 0
Response.Buffer = true

self="YEDE02"
CurrentPage = request("CurrentPage")
TotalPage = request("TotalPage")
PageRec = request("PageRec")
gTotalPage = request("gTotalpage")

Set conn = GetSQLServerConnection()
tmpRec = Session("YEDE02B")
  
Set CONN = GetSQLServerConnection()

conn.BeginTrans
x = 0
y = ""
YYMM=REQUEST("YYMM")

MMDAYS  = REQUEST("MMDAYS")
EMPIDSTR=""
for i = 1 to TotalPage
	for j = 1 to PageRec
		'if trim(tmpRec(i, j, 0))="*" then
			if replace(trim(tmpRec(i, j, 1)),"/","") <>"" then 
			 	workdat=replace(trim(tmpRec(i, j, 1)),"/","")
			 	empid=trim(tmpRec(i, j, 4))
			 	dsts=trim(tmpRec(i, j, 26))
				T1=replace(trim(tmpRec(i, j, 2)),":","")&"00"
				if trim(tmpRec(i, j, 2))="" then
					T1="000000"
				end if 	
				T2=replace(trim(tmpRec(i, j, 3)),":","")&"00"
				if trim(tmpRec(i, j, 3))="" then
					T2="000000"
				end if 
				
				kzhour=trim(tmpRec(i, j, 10))
				if trim(tmpRec(i, j, 10))=""  or isnull(tmpRec(i, j, 10)) then kzhour=0  
				
				forget=trim(tmpRec(i, j, 9))
				if trim(tmpRec(i, j, 9))=""  or isnull(tmpRec(i, j, 9)) then forget=0 
				
				if trim(tmpRec(i, j, 7))="" or isnull(tmpRec(i, j, 7)) then 
					toth=0 
				else
					toth=trim(tmpRec(i, j, 7))
				end if 		
				
				latefor=trim(tmpRec(i, j, 12))
				if trim(tmpRec(i, j, 12))=""  or isnull(tmpRec(i, j, 12)) then latefor=0 
				
				b3=trim(tmpRec(i, j, 16))
				if trim(tmpRec(i, j, 16))=""  or isnull(tmpRec(i, j, 16)) then b3=0 
				
				JB=trim(tmpRec(i, j, 25))
				if trim(tmpRec(i, j, 25))=""  or isnull(tmpRec(i, j, 25)) then jb=0 
				
				if dsts="H1" then 
					H1=JB
					H2="0"
					H3="0"
				elseif dsts="H2" then 
					H1="0"
					H2=JB
					H3="0"					
				elseif dsts="H3" then 
					H1="0"
					H2="0"
					H3=JB
				end if 	 
				
				sqlx="select * from empwork where workdat='"& workdat &"' and empid='"& empid &"' "
				Set rds = Server.CreateObject("ADODB.Recordset") 		
				rds.open sqlx, conn, 1, 3 	 
				if rds.eof then 
					SQL = "INSERT INTO empwork (EMPID , workdat, timeup, timedown, toth, forget, kzhour , H1, H2, H3, B3 , "&_
				 		  "latefor , flag , yymm, mdtm, muser , userIP   ) values ( "&_
				 		  "'"& empid &"', '"& workdat &"', '"& T1 &"', '"& T2 &"' , '"& toth &"' ,"&_
				 		  "'"& forget  &"','"& kzhour &"', '"& H1 &"', '"& H2 &"', "&_
				 		  "'"& H3 &"', '"& B3 &"', '"& latefor &"', '*' , '"& left(workdat,6) &"', getdate(), "&_
				 		  "'"& session("netuser") &"', '"& session("vnLogIP") &"'   ) " 
				 	conn.execute(sql) 
				 	response.write sql&"<BR>"		  
				 	X = X + 1			
				else 			
					sql="update empwork set timeup='"&T1 &"', timedown='"& T2 &"', kzhour='"& kzhour &"', forget='"& forget &"' , "&_
						"toth='"& toth &"',latefor='"& latefor &"', H1='"& H1 &"', H2='"& H2 &"', H3='"& H3 &"', "&_
						"b3='"& b3 &"' , upd='*', mdtm=getdate(), muser='"& session("NetUSer") &"', yymm='"& left(workdat,6) &"'  "&_
						"where  empid='"& empid &"' and workdat='"& workdat &"' " 
					conn.execute(sql)
					response.write sql&"<BR>"	
					x = x+1 
				end if 
			end if 	
		'END IF
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
	Set Session("YECE03") = Nothing
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
 
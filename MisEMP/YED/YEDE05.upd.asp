<%@LANGUAGE="VBSCRIPT" codepage=65001%>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<!-- #include file = "../GetSQLServerConnection.fun" -->
<!-- #include file="../ADOINC.inc" -->
<%
Response.Expires = 0
Response.Buffer = true

self="YEDE05" 
PageRec = request("PageRec")

Set conn = GetSQLServerConnection()
 
conn.BeginTrans
x = 0
y = ""
YYMM=REQUEST("YYMM")

MMDAYS  = REQUEST("MMDAYS")
EMPIDSTR=""
for i = 1 to PageRec
  op = request("op")(i)
  lsempid = request("lsempid")(i)
  cardNo = request("cardno")(i)
  whsno="LA"
  empid= UCase(trim(request("empid")(i)))
  workdat = trim(request("workdat")(i))
  T1 = trim(request("T1")(i))
  T2 = trim(request("T2")(i))
  jb = trim(request("jb")(i))   
  if op="Y" then 
	DT1=left(workdat,4)&"/"&mid(workdat,5,2)&"/"&right(workdat,2) &" "&T1
	DT2=left(workdat,4)&"/"&mid(workdat,5,2)&"/"&right(workdat,2) &" "&T2
	
	if right(workdat,2)>="26" then 
		if right(left(workdat,6),2)+1 > "12" then 
			yymm=left(workdat,4)+1&"01" 
		else
			yymm=left(workdat,6)+1 
		end if 
	else
		yymm = left(workdat,6) 
	end if  
	if left(T2,2)<left(T1,2) then 
		toth = round( datediff( "N" ,cdate(dt1), CDATE(dt2) ) / 30,0) / 2 + 24 
	else
		toth = round( datediff( "N" ,cdate(dt1), CDATE(dt2) ) / 30,0) / 2
	end if 	
	response.write "toth=" & toth  &"<br>"
	response.write lsempid &"<br>"
	response.write cardNo &"<br>"
	response.write workdat &"<br>"
	response.write empid &"<br>"
	response.write yymm &"<br>"
	response.write dt1 &"<br>"
	response.write dt2 &"<br>"
	sql="insert into empforget ( whsno,empid,Lsempid, dat,timeup, timedown, toth, status,yymm,mdtm, muser,cardno)  values ( "&_
		"'"&whsno&"','"&empid&"','"&Lsempid&"','"& left(workdat,4)&"/"&mid(workdat,5,2)&"/"&right(workdat,2) &"','"&T1&"','"&T2&"','"& toth &"', "&_
		"'', '"&yymm&"', getdate(),'"& session("netuser") &"','"& cardNo &"' ) " 
	conn.execute(Sql)
	response.write sql &"<Br>"	
	
	nt1= replace(T1,":","")&"00" 
	nt2= replace(T2,":","")&"00"
	sqlx="select * from empwork where empid='"& empid &"' and workdat='"& workdat &"'  "
	response.write sqlx &"<Br>"
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.open sqlx, conn, 1,3 
	if rs.eof then 
		sql="insert into empwork (empwhsno, empid, workdat, timeup, timedown , totH, yymm, memo, flag, mdtm, muser ) values ( "&_
			"'"& whsno &"','"& empid &"','"& workdat &"','"& Nt1 &"','"& nt2 &"','"& toth &"', '"& yymm &"',  "&_
			"'"& lsempid &"'+'-'+'"& cardNo &"' , 'INS', getdate(),'"& session("netuser") &"' )  "
		conn.execute(Sql)	
	else
		if ( rs("timeup")="" and rs("timedown")="" ) or ( rs("timeup")="000000" and  rs("timedown")="000000"  )   then  		
			sql="update empwork set timeup='"& Nt1 &"' , timedown='"& nt2 &"' , mdtm=getdate(), muser='"& session("userid") &"', "&_ 
				"yymm='"& yymm &"' , memo='"& lsempid &"'+'-'+'"& cardNo &"' "&_
				"where empid='"& empid &"' and workdat='"& workdat &"'   "
			conn.execute(Sql)	
		else 	
			sql="update empwork set memo='"& lsempid &"'+'-'+'"& cardNo &"' "&_
				"where empid='"& empid &"' and workdat='"& workdat &"'  "
			conn.execute(Sql)		
		end if 
	end if 
	response.write sql&"<br>" 
	set rs=nothing  
	'------------------------------------------------------- 工號 ---------------------------------------出勤日期---------------------------------------------------------------------------------------------------------上下班時間 ----------------------------排班----------工時---------------------------no use(number) 	
	sql2="exec A_UpdWokTime '"& empid &"' , '"& left(workdat,4)&"/"&mid(workdat,5,2)&"/"&right(workdat,2) &"' , '"& T1 &"' ,'"& T2 &"' ,'', '"& toth &"' , '0', '0'  "
	conn.execute(sql2)
  end if 
  
next

'response.end  
response.write err.number &"<BR>"
response.write conn.errors.count &"<BR>"
for g =0 to conn.errors.count-1
	response.write conn.errors.item(g)&"<br>"
	response.write Err.Description
next   
'RESPONSE.END

if err.number = 0 then
	conn.CommitTrans
	'response.end  
	Set Session("YEDE0") = Nothing
	Set conn = Nothing 
%>	<SCRIPT LANGUAGE=VBSCRIPT>
		ALERT "資料處理成功!!"&"<%=X%>"&"筆"&chr(13)&"DATA CommitTrans Success !!"
		OPEN "<%=self%>.asp" , "_parent"
	</script>
<% 
ELSE
	conn.RollbackTrans 
%>	<SCRIPT LANGUAGE=VBSCRIPT>
		ALERT "資料處理失敗!!"&chr(13)&"DATA CommitTrans ERROR !!"
		OPEN "<%=self%>.asp" , "_parent"
	</script>
<%  response.end
END IF
%>
 
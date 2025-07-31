<%@LANGUAGE="VBSCRIPT" codepage=65001%>
<!-- #include file = "../../GetSQLServerConnection.fun" --> 
<!-- #include file="../../ADOINC.inc" -->   
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="refresh"> 
</head>
<%
Response.Expires = 0
Response.Buffer = true 

empid=request("empid")
HOLIDAY_TYPE  = request("HOLIDAY_TYPE")
HHDAT1 = request("HHDAT1")
HHDAT2 = request("HHDAT2")
HHTIM1 = request("HHTIM1")
HHTIM2 = request("HHTIM2")
toth = request("toth")
memo = request("memo") 
HDcnt = request("HDcnt")  



Set CONN = GetSQLServerConnection()  
conn.BeginTrans
sqlstr = "select * from empHoliday where empid='"& empid &"' and dateup='"& HHDAT1 &"' and timeup='"& HHTIM1 &"' "&_
		 "and dateDown='"& HHDAT2 &"' and timeDown='"& HHTIM2 &"' " 
'set rst=conn.execute(Sqlstr)  
Set rst = Server.CreateObject("ADODB.Recordset")   
rst.open sqlstr, conn, 3, 3
if rst.eof then 	
	if HHDAT2 > HHDAT1 then 
		days = fix(toth/8)+cdbl(HDcnt)  
		for x = 1 to days
			cdatestr = year(cdate(HHDAT1)+(x-1))&"/"&right("00"& month(cdate(HHDAT1)+(x-1)),2)&"/"&right("00"& day(cdate(HHDAT1)+(x-1)),2)
			sqlstra = "select * from ydbmcale where status='H1' and convert(char(10), dat, 111) = '"& cdatestr &"' "
			Set rs = Server.CreateObject("ADODB.Recordset")   
			rs.open sqlstra, conn, 3, 3			
			if not rs.eof then 
				'response.write cdatestr  &"<BR>" 
				sql="insert into empHoliday ( empid, jiaType, DateUP, TimeUP, DateDown, TimeDown, HHour, memo, Muser ) values ( "&_
					"'"& empid &"', '"& HOLIDAY_TYPE &"', '"& cdatestr &"', '08:00', '"& cdatestr &"', '17:00', "&_
					"'8', '"& memo &"', '"& session("NETUSER") &"' ) " 
				'response.write sql &"<BR>"
				conn.execute(sql) 
				
				if HOLIDAY_TYPE="G" then 
					f1_toth=8 
				else 
					f1_toth=0  	
				end if 	
				
				sql2="select * from empwork where empid='"& empid &"' and workdat='"& trim(replace(cdatestr,"/","")) &"' "
				Set rds = Server.CreateObject("ADODB.Recordset")   
				rds.open sql2, conn, 3, 3
				if rds.eof then 
					sql3="insert into empwork ( empid, workdat, timeup, timedown, toth, flag, yymm) values ( "&_
						 "'"& empid &"' , '"& trim(replace(cdatestr,"/","")) &"' , '080000', '170000', "& f1_toth &", 'JIA',  '"& left(trim(replace(cdatestr,"/","")),6) &"'  ) "  
					response.write sql3 &"<BR>"
					conn.execute(sql3)
					sql4="update empwork set JIA"&HOLIDAY_TYPE&" = JIA"&HOLIDAY_TYPE&" +'8' where empid='"& empid &"' and workdat='"& trim(replace(cdatestr,"/","")) &"' "
					response.write sql4 &"<BR>"
					conn.execute(sql4)
				else
					sql3="update empwork set kzhour=0, JIA"&HOLIDAY_TYPE&" =JIA"&HOLIDAY_TYPE&" + '8' where empid='"& empid &"' and workdat='"& trim(replace(cdatestr,"/","")) &"' "
					response.write sql3 &"<BR>"
					conn.execute(sql3)
				end if 
				set rds=nothing 
				
			end if 
			set rs=nothing   
		next 
	else
		sql="insert into empHoliday ( empid, jiaType, DateUP, TimeUP, DateDown, TimeDown, HHour, memo, Muser ) values ( "&_
			"'"& empid &"', '"& HOLIDAY_TYPE &"', '"& HHDAT1 &"', '"&  HHTIM1  &"', '"& HHDAT2 &"', '"& HHTIM2 &"', "&_
			"'"& toth &"', '"& memo &"', '"& session("NETUSER") &"' ) " 
		''response.write sql 	
		conn.execute(sql)
		
		if HOLIDAY_TYPE="G" then 
			f1_toth=8 
		else 
			f1_toth=0  	
		end if 	
					 
		sql2="select * from empwork where empid='"& empid &"' and workdat='"& trim(replace(HHDAT1,"/","")) &"' "
		Set rds = Server.CreateObject("ADODB.Recordset")   
		rds.open sql2, conn, 3, 3
		if rds.eof then 
			sql3="insert into empwork ( empid, workdat, timeup, timedown, flag, yymm ) values ( "&_
				 "'"& empid &"' , '"& trim(replace(HHDAT1,"/","")) &"' , '"& replace(HHTIM1,":","")&"00"  &"', '"& replace(HHTIM2,":","")&"00" &"', 'JIA',  '"& left(trim(replace(HHDAT1,"/","")),6) &"'  ) "  
			response.write sql3 &"<BR>" 
			conn.execute(sql3)
			
			sql4="update empwork set kzhour=0 , JIA"&HOLIDAY_TYPE&" = JIA"&HOLIDAY_TYPE&" +'"& toth &"' where empid='"& empid &"' and workdat='"& trim(replace(HHDAT1,"/","")) &"' "
			response.write sql4 &"<BR>"
			conn.execute(sql4)
		else
			sql3="update empwork set  flag='JIA', kzhour = isnull(kzhour,0) - "& toth &" ,  "&_
				 "JIA"&HOLIDAY_TYPE&" =isnull( JIA"&HOLIDAY_TYPE&" ,0) + "& toth &" "&_
				 "where empid='"& empid &"' and workdat='"& trim(replace(HHDAT1,"/","")) &"' "
				response.write sql3 &"<BR>"
				conn.execute(sql3) 

	end if	
else
	response.write "請假資料重複( Data CommitTrans Error) !!<BR>" 
	response.write "<a href=empholiday.asp>回主畫面重新申請</a>"
	response.end 
end if 	 
'set rst=nothing 
'RESPONSE.END  

if conn.Errors.Count = 0 then 
	conn.CommitTrans	
	'response.redirect "empfile.salary.asp?empidstr=" & empidstr 
	Set conn = Nothing
%>	<SCRIPT LANGUAGE=VBSCRIPT>
		ALERT "DATA CommitTrans  SUCCESS!!"&chr(13)&"資料處理成功!!"		 
		open "empholiday.asp", "_self"
	</script>		
<%
ELSE
	conn.RollbackTrans	
%>	<SCRIPT LANGUAGE=VBSCRIPT>
		ALERT "DATA CommitTrans ERROR !!"
		OPEN "empfile.salary.asp" , "_self" 
	</script>	
<%	response.end 
END IF  
%>
 </html>
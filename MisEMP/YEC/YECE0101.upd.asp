<%@LANGUAGE="VBSCRIPT" CODEPAGE=65001%>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<!-- #include file = "../GetSQLServerConnection.fun" --> 
<!-- #include file="../ADOINC.inc" -->  
<!--#include file="../include/checkpower.asp"--> 
<%
Response.Expires = 0
Response.Buffer = true  

CurrentPage = request("CurrentPage")
TotalPage = request("TotalPage")
PageRec = request("PageRec")
gTotalPage = request("gTotalpage")

tmpRec = Session("empsalary01") 

Set CONN = GetSQLServerConnection()  
calcYM = request("calcYM") 

cDatestr=CDate(LEFT(calcYM,4)&"/"&RIGHT(calcYM,2)&"/01") 
days = DAY(cDatestr+(32-DAY(cDatestr))-DAY(cDatestr+(32-DAY(cDatestr))))   '一個月有幾天   
'response.write "A="&request("bbcode")
'response.end
conn.BeginTrans
x = 0
y = ""  
YYMM=REQUEST("YYMM")
EMPIDSTR="" 
if session("netuser")<>""   then 
	for i = 1 to TotalPage 
		for j = 1 to PageRec  
			'RESPONSE.WRITE TotalPage &"<br>"
			'RESPONSE.WRITE PageRec &"<br>"
			if trim(tmpRec(i, j, 1))<>"" then 
				if trim(tmpRec(i, j, 4))="VN" or trim(tmpRec(i, j, 4))="CT" then 
					dm="VND"
				else
					dm="USD"
				end if 
				'IF trim(tmpRec(i, j, 0))="UPD" THEN 
				sqlx="select * from bemps where  empid='"& TRIM(tmpRec(i, j, 1)) &"' and yymm='"& calcYM &"'"
				Set rst = Server.CreateObject("ADODB.Recordset")   
				rst.open sqlx,conn, 1, 3 
				bb=tmpRec(i, j, 20)
				cv=tmpRec(i, j, 22)
				phu=tmpRec(i, j, 23)
				nn=tmpRec(i, j, 24)
				kt=tmpRec(i, j, 25)
				MT=tmpRec(i, j, 26)
				wp_bb = 0 'cdbl(tmpRec(i, j, 19)) 
				wp_CV = cdbl(tmpRec(i, j, 19)) 
				TTKH=tmpRec(i, j, 27)  
				btien=tmpRec(i, j, 44) 
				
				' if  trim(tmpRec(i, j, 4))="MA" or  trim(tmpRec(i, j, 4))="TW" then 
					' TTKH= 0  
					' if tmpRec(i, j, 27)<>"" and tmpRec(i, j, 27)>"0" then 
						 'wp_bb = 0 'cdbl(tmpRec(i, j, 20))    -chang by elin 20090428  不區分 , 全部放入海外津貼
						 'wp_CV = cdbl(tmpRec(i, j, 27))   '-cdbl(tmpRec(i, j, 20))					
					' else
						' wp_bb=0							
						' wp_cv=0
					' end if 	
				' else
					' TTKH=tmpRec(i, j, 27) 
				' end if 	
  
				if rst.eof then 						
					sqln="insert into bemps( yymm,whsno,country,empid,groupid,Job,DM,BB,CV,PHU,NN,KT,MT,TTKH,QC,MEMO, MUSER, lncode , wp , btien , tien3 ) values ( "&_
						 "'"& calcYM &"','"& trim(tmpRec(i, j, 7)) &"','"& trim(tmpRec(i, j, 4)) &"','"& TRIM(tmpRec(i, j, 1)) &"', "&_
						 "'"& trim(tmpRec(i, j, 9)) &"','"& tmpRec(i, j, 6) &"','"& dm &"','"& BB &"','"& tmpRec(i, j, 22) &"', "&_
						 "'"& tmpRec(i, j, 23) &"','"& tmpRec(i, j, 24) &"','"& tmpRec(i, j, 25) &"','"& tmpRec(i, j, 26) &"', "&_
						 "'"& ttkh &"','"& tmpRec(i, j, 32) &"',N'"& tmpRec(i, j, 34) &"','"& session("netuser") &"', "&_
						 "'"& tmpRec(i, j, 47) &"' ,'"& tmpRec(i, j, 19) &"' ,'"& tmpRec(i, j, 44) &"' ,'"& tmpRec(i, j, 46) &"' ) " 
					conn.execute(sqln) 					 
				else
					sqln="UPDATE bemps set whsno='"& trim(tmpRec(i, j, 7)) &"',country='"& trim(tmpRec(i, j, 4)) &"',  "&_
						 "groupid='"& trim(tmpRec(i, j, 9)) &"', job='"& tmpRec(i, j, 6) &"',dm='"& dm &"', bb='"& tmpRec(i, j, 20) &"', "&_
						 "cv='"& tmpRec(i, j, 22) &"', phu='"& tmpRec(i, j, 23) &"', nn='"& tmpRec(i, j, 24) &"', "&_
						 "KT='"& tmpRec(i, j, 25) &"', MT='"& tmpRec(i, j, 26) &"' , TTKH='"& ttkh &"', "&_
						 "QC='"& tmpRec(i, j, 32) &"', mdtm=getdate(), muser='"& session("netuser") &"' , lncode='"& tmpRec(i, j, 47) &"' ,  "&_
						 "wp='"& tmpRec(i, j, 19) &"' , btien='"& tmpRec(i, j, 44) &"' , tien3='"& tmpRec(i, j, 46) &"' "&_
						 "where empid='"& TRIM(tmpRec(i, j, 1)) &"'and yymm='"& calcYM &"' " 
					conn.execute(sqln)	 
				end if 
				set rst=nothing 
				'response.write sqln &"<BR>" 
				
		 		sqlx="select * from bempj where  empid='"& TRIM(tmpRec(i, j, 1)) &"' and yymm='"& calcYM &"'"
				Set rst = Server.CreateObject("ADODB.Recordset") 
				rst.open sqlx, conn, 1,3 
				if rst.eof then 
					sql2="insert into bempj (yymm, whsno, country, empid, job, memo, mdtm, muser ) values ( "&_
						 "'"& calcYM &"','"& trim(tmpRec(i, j, 7)) &"','"& trim(tmpRec(i, j, 4)) &"', "&_
						 "'"& TRIM(tmpRec(i, j, 1)) &"','"& tmpRec(i, j, 6) &"','YECE0101 Insert', getdate(), '"& session("netuser") &"' ) "
					conn.execute(sql2)	 
				else
					if rst("job")<>tmpRec(i, j, 6) then  
						sql2x="insert into bempjT select * from bempj where  empid='"& TRIM(tmpRec(i, j, 1)) &"' and yymm='"& calcYM &"' "
						conn.execute(sql2x)
					end if 	
					sql2="update bempj set job='"& tmpRec(i, j, 6) &"', whsno='"& trim(tmpRec(i, j, 7)) &"', "&_
						 "country='"& trim(tmpRec(i, j, 4)) &"' , memo='YECE0101 IPD', mdtm=getdate(), muser='"& session("netuser") &"' "&_
						 "where  empid='"& TRIM(tmpRec(i, j, 1)) &"' and yymm='"& calcYM &"' " 
					conn.execute(sql2)	 
				end if 
				set rst=nothing  
				
				if wp_bb >0 or wp_CV > 0 then 
					sqlx="select * from salarywp where yymm='"& calcYM &"' and empid='"& TRIM(tmpRec(i, j, 1)) &"' and dm='"& dm &"'   " 
					Set rds = Server.CreateObject("ADODB.Recordset") 		
					rds.open sqlx, conn, 1, 3 	
						if rds.eof and  op<>"DEL" then  
						 	sql="insert into salarywp (yymm,workdays,empid,country,whsno,empname,BB, cv, dm,totAMT, mdtm,muser,userIP,closeflag	) values (  "&_
						 		"'"& yymm &"', '"& days &"', '"& TRIM(tmpRec(i, j, 1))  &"','"& TRIM(tmpRec(i, j, 4))  &"','"& TRIM(tmpRec(i, j, 7))  &"',N'"& TRIM(tmpRec(i, j, 2))  &"', "&_
						 		"'"& wp_bb &"','"& wp_CV &"','"&dm&"', '"& cdbl(wp_bb)+cdbl(wp_cv)  &"', getdate(), '"& session("netuser") &"', "&_
						 		"'"& session("vnlogIP") &"','') " 
						 	conn.execute(sql)	 	
						 	'response.write sql &"<BR>"
					 	else
					 		sql="update salarywp set  whsno='"& TRIM(tmpRec(i, j, 7)) &"' , country='"& TRIM(tmpRec(i, j, 4)) &"',  "&_
					 			"empname=N'"& TRIM(tmpRec(i, j, 2))   &"' , bb='"& wp_bb &"', cv='"& wp_cv &"',    "&_
					 			"totAMT='"& cdbl(wp_bb)+cdbl(wp_cv) &"' ,    "&_
					 			"mdtm=getdate(), muser='"& session("netuser")  &"', userip='"& session("vnlogIP")  &"' "&_
					 			"where yymm='"& yymm &"'  and empid='"& TRIM(tmpRec(i, j, 1)) &"' and dm='"& dm &"'  "  
					 		conn.execute(sql)	
					 		response.write sql &"<BR>"
					 	end if 				
					
				end if 
				
			END IF 
		next
	next 	
end if 
'RESPONSE.END 

if conn.Errors.Count = 0 or err.number=0 then 
	conn.CommitTrans
	Set Session("empsalary01") = Nothing 	  	
	Set conn = Nothing
%>	<SCRIPT LANGUAGE=VBSCRIPT>
		ALERT "資料處理成功!!"&chr(13)&"DATA CommitTrans Success !!"
		OPEN "YECE0101.asp" , "_self" 
	</script>	
<% 
ELSE
	conn.RollbackTrans	
%>	<SCRIPT LANGUAGE=VBSCRIPT>
		ALERT "資料處理失敗!!"&chr(13)&"DATA CommitTrans ERROR !!"
		OPEN "YECE0101.asp" , "_self" 
	</script>	
<%	response.end 
END IF 
%>
 
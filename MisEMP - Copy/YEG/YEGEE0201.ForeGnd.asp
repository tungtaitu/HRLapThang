<%@ Language=VBScript codepage=65001%> 
<!-- #include file = "../GetSQLServerConnection.fun" --> 
<!-- #include file="../ADOINC.inc" --> 
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="refresh">
<link rel="stylesheet" href="../Include/style.css" type="text/css">
<link rel="stylesheet" href="../Include/style2.css" type="text/css">
<%
session.codepage="65001"
self="YEGEE0201"
dat1 = cdate(trim(request("dat1")))
dat2 = cdate(trim(request("dat2")) )
eid  = request("eid")


D1 = replace(dat1, "/", "" )
D2 = replace(dat2, "/", "" )   
df = datediff("d",dat1,dat2) 
eid= Ucase(trim(request("eid")))
 
aidstr = session("netuser") & minute(now())&second(now()) 

Set conn = GetSQLServerConnection()	  
Set cnykt = GetaccessConnection()	  
Set rs = Server.CreateObject("ADODB.Recordset")    
pagerec= 1 
sqln="select * from view_empfile where isnull(status,'')<>'D' " 
Set rds = Server.CreateObject("ADODB.Recordset")    	
rds.open sqln, conn, 1, 3 
if not rds.eof then 
	pagerec = rds.recordCount 
end if 
chkemp=""
while not rds.eof 	
	chkemp = chkemp & rds("empid")&","
	rds.movenext
wend 	
set rds=nothing 


'---step1  
sql_1="exec  A_tranEmpData  '"& eid &"' , '"& dat1 &"', '"& dat2 &"' " 
conn.execute(Sql_1)  

conn.BeginTrans   

'---step2 抓access資料 
'sql="select * from Report_Day  where emp_id like '%"& eid &"' and sign_date  between #"& dat1 &"# and #"& dat2 &"# order by  sign_date , emp_id "
sql="select b.card_id , a.* from   "&_
	"(select * from Report_Day where emp_id like '%"& eid &"' and sign_date  between #"& dat1 &"# and #"& dat2 &"#   ) a "&_
	"left join( select emp_id, card_id from employee ) b on b.emp_id = a.emp_id "&_ 
	"order by  a.sign_date , a.emp_id "
'response.write sql &"<BR>"
'response.end 
rs.open sql, cnykt, 3, 3 

response.write rs.recordcount &"<BR>" 	

while not rs.eof 
	if instr(1,chkemp,right(left(rs("emp_id"),6),5))>0 then 
		empid=right(left(rs("emp_id"),6),5)   
		if trim(rs("in1"))<>"" then 
			timeup=trim(rs("in1")) 'replace(rs("in1"),":","")&"00"
		else
			if rs("out1")<>"" then 
				timeup=trim(rs("out1")) 'replace(rs("out1"),":","")&"00"
			else
				timeup=""
			end if 	
		end if 	
		if trim(rs("out1"))<>"" then 		
			timedown= trim(rs("out1")) 'replace(rs("out1"),":","")&"00"
		else
			if rs("in1")<>"" then 
				timedown= trim(rs("in1")) 'replace(rs("in1"),":","")&"00"
			else
				timedown=""
			end if 	
		end if 	
		workdat=year(rs("sign_date"))&"/"&right("00"&month(rs("sign_date")),2)&"/"&right("00"&day(rs("sign_date")),2)
		shift_id= replace( trim(rs("shift_id")),"+","")
		fact_hrs=trim(rs("fact_hrs"))  '總工時
		ot_hrs=trim(rs("ot_hrs"))  '加班時數 
		
		sql_2="exec A_UpdWokTime '"& empid &"', '"& workdat &"','"& timeup &"' , '"& timedown &"' , '"& shift_id &"', '"& fact_hrs &"', '"& ot_hrs &"'"
		conn.execute(sql_2) 
		response.write sql_2 &"<BR>"
	else
		empid=rs("emp_id") &"---NG" &"<BR>"
		sqlstr="select * from TWorkTime where emp_id ='"& trim(rs("emp_id")) &"'  "&_
			   "and card_id = '"& trim(rs("card_id")) &"' and  workdat= convert(char(8), '"& rs("sign_date") &"',112) "
		'response.write sqlstr &"<BR>"
		'response.end 			   
		Set rst1 = Server.CreateObject("ADODB.Recordset")    
		rst1.open sqlstr, conn, 3, 3 
		if rst1.eof then 
			sqlix="insert into TWorkTime (emp_id,card_id, workdat, timeup, timedown ) values ( "&_
				  "'"& rs("emp_id") &"', '"& rs("card_id") &"' , convert(char(8),'"& rs("sign_date") &"' ,112) , "&_
				  "'"& rs("in1") &"' ,'"& rs("out1") &"'  ) " 
			conn.execute(sqlix) 	  
		else
			sqlix="update TWorkTime set timeup='"& rs("in1") &"' , timedown='"& rs("out1") &"' "&_
				  "where  emp_id ='"& trim(rs("emp_id")) &"'  "&_
				  "and card_id = '"& trim(rs("card_id")) &"' and  workdat= convert(char(8), '"& rs("sign_date") &"',112) "
			conn.execute(sqlix) 	  
		end if 
		
		response.write sqlix 
	end if 	
	
	'response.write  empid
rs.movenext
wend   
set cnykt=nothing  
response.write "OK"
'response.end 

if  conn.Errors.Count=0 or err.number=0  then 
	conn.CommitTrans 
%>	<script language=vbscript>
		alert "Data CommitTrans sunccess!!  OK!!"
		open "<%=self%>.asp" , "_parent"
	</script>
<%	response.end 
else
	conn.RollbackTrans	%>
	<script language=vbscript>
		alert "Data CommitTrans Fail!!  Error!!"
		open "<%=self%>.asp" , "_parent"
	</script>	
<%  response.end 
end if 	
%>
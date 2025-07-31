<%@LANGUAGE="VBSCRIPT" codepage=65001%>
<!-- #include file = "../GetSQLServerConnection.fun" -->
<!-- #include file="../ADOINC.inc" -->
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="refresh">
<%
Response.Expires = 0
Response.Buffer = true
self="YFYEMPJXA"
JXYM = REQUEST("JXYM")
SALARYYM = REQUEST("SALARYYM")
GROUPID =request("GID")
zuno =request("zID")
icount = request("icount")
PageRec = request("PageRec")
RecordInDB = request("RecordInDB")
SHIFTN=Trim(request("shiftn"))
'response.write JXYM
'response.end 
empid1 = REQUEST("empid1")
F_GROUPID = REQUEST("F_GROUPID")
F_SHIFT=REQUEST("F_SHIFT")
F_zuno = REQUEST("F_zuno")
F_WHSNO=request("F_whsno")


Set conn = GetSQLServerConnection() 
tmpRec = Session("YFYEMPJXM")
'ARRAYS = Session("EMPJX")

conn.BeginTrans

'sqlstr="delete VYFYMYJX where YYMM='"& SalaryYM &"' and JXYM='"& JXYM &"' "&_
'	   "and groupid='"& GROUPID &"' and shift='"& shiftN &"' and isnull(zuno,'') like '"& zuno &"%'  "
'conn.execute(sqlstr)
  


FOR I = 1 TO PageRec
	if trim(tmprec(1,I,1))<>"" then
		empid=tmprec(1,I,1)
		SHIFT=tmprec(1,I,3)
		UNITJX=tmprec(1,I,4)
		FL= tmprec(1,I,8)
		FLM= request("FLMONEY")(I)
		FQD=request("FQD")(I)
		khfen=request("khfen")(I)
		gnn=request("gnn")(I)
		SUKM=request("SSmoeny")(I)
		TOTJX = round(request("TOTJX")(I),0)
		RELJX = (cdbl(TOTJX)\1000) * 1000
		NRelJXM = request("realJXM")(I)
		js=request("workJs")(I)
		NowGroup=request("NowGroup")(I)
		NowShift=request("NowShift")(I)
		NowZuno=request("NowZuno")(I)
		JxGroup=request("jxGroup")(I)
		jxShift=request("jxShift")(I)
		jxZuno=request("jxZuno")(I) 
		jxyn = request("jxyn")(I) 
		newjs = request("newjs")(I) 
		hrjs = request("hrjs")(I) 
		
				
		F1_colsName=""
		F1_JXM=""	
		sumJX = 0 
		js=1
		
		for x = 1 to 5
			colsName="JX"&chr(x+64)
			JXM = REQUEST(colsName)(I)
			F1_colsName = F1_colsName & "," & colsName
			F1_JXM = F1_JXM & "," & JXM
			SumJX = cdbl(SumJX) + cdbl(JXM)
		next 
		'response.write F1_colsName &"<BR>"
		'response.write F1_JXM &"<BR>" 
		'response.end 
 
	 
			sqla="delete VYFYMYJX where yymm='"& SalaryYM  &"' and jxym='"& JXYM &"' and empid='"& EMPID &"' "
			'response.write sqla 
			conn.execute(sqla)
			
			sql="insert into VYFYMYJX(YYMM, JXYM, GROUPID, jxGroup, SHIFT, jxShift, zuno, jxzuno, EMPID, UNITJX  "
			sql=sql & F1_colsName & ",sumJX, FL, FLM, FQD, SUKM, TOTJXM, RELJXM, js,  muser , fensu, rp_cnt ,newjs , hrjx  ) values (  "
			sql=sql & "'"& SalaryYM &"', '"& JXYM &"', '"& NowGroup &"', '"& jxGroup &"', '"& nowSHIFT &"', '"& jxSHift &"' , "&_
				"'"& Nowzuno &"', '"& jxzuno &"', '"& EMPID &"', '"& UNITJX  &"' "
			sql=sql & F1_JXM &", '"& SumJX &"', '"& FL &"', '"& FLM &"', '"& FQD &"', '"& SUKM &"', '"& TOTJX &"',  "&_
			"'"& RELJX &"','"& js &"',   '"& session("NetUser") &"','"&khfen&"','"&gnn&"','"&newjs&"','"&hrjs&"'   ) "
			response.write "D= " & sql &"<BR>"
			conn.execute(sql) 
			
			' sql="update VYFYMYJX set FLM='"& FLM &"', fensu='"& khfen &"', rp_cnt='"& gnn &"' "&_
					' "where yymm='"& SalaryYM  &"' and jxym='"& JXYM &"' and empid='"& EMPID &"' "
			' response.write sql &"<BR>"
			'conn.execute(Sql) 
			
			sql2="update empJXYN set jxyn='"&jxyn&"' where empid='"&EMPID&"' and yymm='"& jxym &"' "
			conn.execute(sql2)
			response.write sql2&"<BR>" 
	 
	end if
NEXT
'response.end

if conn.Errors.Count = 0 or err.number=0 then
	'response.clear
	conn.CommitTrans
	Set conn = Nothing
	set Session("YFYEMPJXM")=nothing
	set Session("EMPJX")=nothing
%>	<SCRIPT LANGUAGE=VBSCRIPT>
		ALERT "資料處理成功!!"&chr(13)&"DATA CommitTrans Success !!"
		OPEN "YECE06.Fore.asp?F_WHSNO="&"<%=F_WHSNO%>"&"&f_groupid="&"<%=f_groupid%>" , "_self"
	</script>
<%
ELSE
	conn.RollbackTrans
%>	<SCRIPT LANGUAGE=VBSCRIPT>
		ALERT "資料處理失敗!!"&chr(13)&"DATA CommitTrans ERROR !!"
		OPEN "YECE06.asp" , "_self"
	</script>
<%	response.end
END IF
%>
 
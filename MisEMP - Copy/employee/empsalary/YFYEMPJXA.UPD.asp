<%@LANGUAGE="VBSCRIPT" codepage=65001%>
<!-- #include file = "../../GetSQLServerConnection.fun" -->
<!-- #include file="../../ADOINC.inc" -->
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
'response.write GROUPID
'response.end
Set conn = GetSQLServerConnection()
response.write conn
tmpRec = Session("YFYEMPJXM")
ARRAYS = Session("EMPJX")

conn.BeginTrans

sqlstr="delete VYFYMYJX where YYMM='"& SalaryYM &"' and JXYM='"& JXYM &"' "&_
	   "and groupid='"& GROUPID &"' and shift='"& shiftN &"' and isnull(zuno,'') like '"& zuno &"%'  "
conn.execute(sqlstr)
  


FOR I = 1 TO PageRec
	if trim(tmprec(1,I,1))<>"" then
		empid=tmprec(1,I,1)
		SHIFT=tmprec(1,I,3)
		UNITJX=tmprec(1,I,4)
		FL= tmprec(1,I,8)
		FLM= request("FLMONEY")(I)
		FQD=request("FQD")(I)
		SUKM=request("SSmoeny")(I)
		TOTJX = request("TOTJX")(I)
		RELJX = (cdbl(TOTJX)\1000) * 1000
		NRelJXM = request("realJXM")(I)
		js=request("workJs")(I)
		NowGroup=request("NowGroup")(I)
		NowShift=request("NowShift")(I)
		NowZuno=request("NowZuno")(I)
		JxGroup=request("jxGroup")(I)
		jxShift=request("jxShift")(I)
		jxZuno=request("jxZuno")(I)		
				
		F1_colsName=""
		F1_JXM=""
		sumJX = 0
		for x = 1 to cdbl(ICOUNT)
			colsName="JX" & ARRAYS(x,5)
			JXM = REQUEST(colsName)(I)
			F1_colsName = F1_colsName & "," & colsName
			F1_JXM = F1_JXM & "," & JXM
			SumJX = cdbl(SumJX) + cdbl(JXM)
		next
		'response.write F1_colsName &"<BR>"
		'response.write F1_JXM &"<BR>" 
		sqla="delete VYFYMYJX where yymm='"& SalaryYM  &"' and jxym='"& JXYM &"' and empid='"& EMPID &"' "
		'response.write sqla 
		conn.execute(sqla)
		
		sql="insert into VYFYMYJX(YYMM, JXYM, GROUPID, jxGroup, SHIFT, jxShift, zuno, jxzuno, EMPID, UNITJX  "
		sql=sql & F1_colsName & ",sumJX, FL, FLM, FQD, SUKM, TOTJXM, RELJXM, js, NrelJXM, muser  ) values (  "
		sql=sql & "'"& SalaryYM &"', '"& JXYM &"', '"& NowGroup &"', '"& jxGroup &"', '"& nowSHIFT &"', '"& jxSHift &"' , "&_
			"'"& Nowzuno &"', '"& jxzuno &"', '"& EMPID &"', '"& UNITJX  &"' "
		sql=sql & F1_JXM &", '"& SumJX &"', '"& FL &"', '"& FLM &"', '"& FQD &"', '"& SUKM &"', '"& TOTJX &"', '"& RELJX &"','"& js &"',  '"& NRelJXM &"', '"& session("NetUser") &"'  ) "
		response.write "D= " & sql &"<BR>"
		conn.execute(sql)
	end if
NEXT
'response.end

if conn.Errors.Count = 0 then
	'response.clear
	conn.CommitTrans
	Set conn = Nothing
	set Session("YFYEMPJXM")=nothing
	set Session("EMPJX")=nothing
%>	<SCRIPT LANGUAGE=VBSCRIPT>
		ALERT "資料處理成功!!"&chr(13)&"DATA CommitTrans Success !!"
		OPEN "YFYEMPJXA.asp" , "_self"
	</script>
<%
ELSE
	conn.RollbackTrans
%>	<SCRIPT LANGUAGE=VBSCRIPT>
		ALERT "資料處理失敗!!"&chr(13)&"DATA CommitTrans ERROR !!"
		OPEN "YFYEMPJXA.asp" , "_self"
	</script>
<%	response.end
END IF
%>
 
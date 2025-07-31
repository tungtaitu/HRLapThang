<%@LANGUAGE="VBSCRIPT"  codepage=950%>
<!-- #include file = "../GetSQLServerConnection.fun" -->
<!-- #include file="../ADOINC.inc" -->
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<%
Response.Expires = 0
Response.Buffer = true

modify=request("modify")

sgno = request("sgno")
pddate = request("pddate")
session("pddate")=pddate
sgym = request("sgym")
TOTcost = request("TOTcost")
cfGroup = request("cfGroup")
empid = request("empid")
cfdw = request("cfdw")
sgcost = request("sgcost")
sgmemo = request("sgmemo")
NNY = request("NNY")
NDY = request("NDY")
whsno=request("whsno")
autoid= request("autoid")
DM= request("DM")
shift = request("shift")
FsalaryYM = request("FsalaryYM")

whsno = request("whsno")

pddateYM = left(pddate,4)&mid(pddate,6,2)

if TOTcost="" then TOTcos=0
AID= month(date()) & day(date()) &  trim(minute(time)) &  trim(second(time))

'response.write aid 
'response.end 
IF autoid<>"" THEN
	AID=autoid
ELSE
	AID=AID
END IF


Set conn = GetSQLServerConnection()
conn.BeginTrans
sqlstr="select * from YFYMSUCO where autoid = '"& AID &"'    "
Set rDs = Server.CreateObject("ADODB.Recordset")
RDS.OPEN SQLSTR, CONN, 3, 3
if rds.eof then
	sql="insert into YFYMSUCO (autoid, sgno,pdDate,totCost,SgCost,SgYM,sgmemo,mdtm,muser,flag,whsno ) values ( "&_
		"'"& AID &"', '"& sgno &"', '"& pddate &"',  '"& totCost &"' ,  '"& sgcost &"', "&_
		"'"& sgYM &"',N'"& sgmemo &"' , getdate(), '"& SESSION("NETUSER") &"' , 'N','"& whsno&"'  ) "
	conn.execute(sql)
else
	sql="update YFYMSUCO set ToTcost='"& totCost &"' , sgmemo=N'"& sgmemo &"', sgno='"& sgno &"', pddate='"& pddate &"' ,  "&_
		"mdtm=getdate(), muser='"& SESSION("NETUSER")  &"'  where autoid ='"& AID &"' "
	conn.execute(sql)
end if
set rds=nothing
response.write sql&"<BR>"

IF autoid<>"" THEN
	AID=autoid
ELSE
	AID=AID
END IF

Daid = request("Daid")
if daid<>"" then
	sql="delete YFYDSUCO where aid='"& Daid &"' "
	conn.execute(sql)
	response.write sql&"<BR>"
end if

if cfgroup="E" or left(cfgroup,2)="A0" then
	sql="insert into YFYDSUCO (Autoid, sgno, sgym, cfgroup, cfdw,  SUKM ,DM , shift ) values ( "&_
		"'"& AID &"','"& sgno &"' ,'"&  sgYM &"' , '"& cfGroup &"', N'"& cfdw &"', '0' , 'VND', '"& shift &"'   ) "
	conn.execute(sql)
end if


for X = NNY to NDY
	for Y = 1 to 12
		CurrenRow= (cdbl(X*12)+Y) - cdbl(NNY*12)
		SSmoney = request("SSYM")(CurrenRow)

		'response.write X &"-" & Y & SSYM &"<BR>"
		if SSmoney<>"0" then
			SSYM=cstr(X)&cstr(right("00"&y,2))
			sql="insert into YFYDSUCO (Autoid, sgno, sgym, cfgroup, cfdw, empid, YM, SUKM , DM ) values ( "&_
				"'"& AID &"','"& sgno &"' ,'"&  sgYM &"' , '"& cfGroup &"', N'"& cfdw &"', '"& empid &"', '"& SSYM &"',  '"& SSmoney &"' , '"& DM &"'  ) "
			response.write  X &"-" & Y &"-"& sql&"<BR>"
			conn.execute(sql)
		end if
	next
next

'response.end

if conn.Errors.Count = 0 then
	conn.CommitTrans
	'Set conn = Nothing

%>	<SCRIPT LANGUAGE=VBSCRIPT>
		F="<%=modify%>"
		'alert F
		ALERT "資料處理成功!!"&chr(13)&"DATA CommitTrans Success !!"
		if F="E" then
			OPEN "vyfysuco.sch.asp?sgym="&"<%=sgym%>" &"&salaryYM=" &"<%=FsalaryYM%>" &"&sgno="&"<%=sgno%>" , "_self"
		else
			OPEN "vyfysuco.asp" , "Fore"
		end if
	</script>
<%
ELSE
	conn.RollbackTrans
%>	<SCRIPT LANGUAGE=VBSCRIPT>
		F="<%=modify%>"
		ALERT "資料處理失敗!!"&chr(13)&"DATA CommitTrans ERROR !!"
		if F="E" then
			OPEN "vyfysuco.sch.asp" , "_self"
		else
			OPEN "vyfysuco.asp" , "Fore"
		end if
	</script>
<%	response.end
END IF
%>
 
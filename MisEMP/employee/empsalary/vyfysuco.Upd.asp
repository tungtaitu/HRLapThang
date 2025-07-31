<%@LANGUAGE="VBSCRIPT"  codepage=950%>
<!-- #include file = "../../GetSQLServerConnection.fun" --> 
<!-- #include file="../../ADOINC.inc" -->  
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<%
Response.Expires = 0
Response.Buffer = true 

sgno = request("sgno")
pddate = request("pddate")
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

if TOTcost="" then TOTcos=0   
AID= trim(minute(time)) + trim(second(time)) 

IF autoid<>"" THEN 
	AID=autoid 
ELSE
	AID=AID 
END IF 	 


Set conn = GetSQLServerConnection()
conn.BeginTrans 
sqlstr="select * from YFYMSUCO where sgno='"& sgno &"' and convert(char(10),pddate,111) = '"& pddate &"'   "
Set rDs = Server.CreateObject("ADODB.Recordset")   
RDS.OPEN SQLSTR, CONN, 3, 3 
if rds.eof then 
	sql="insert into YFYMSUCO (autoid, sgno,pdDate,totCost,SgCost,SgYM,sgmemo,mdtm,muser,flag ) values ( "&_
		"'"& AID &"', '"& sgno &"', '"& pddate &"',  '"& totCost &"' ,  '"& sgcost &"', "&_
		"'"& sgYM &"','"& sgmemo &"' , getdate(), '"& SESSION("NETUSER") &"' , 'N'  ) "  
	conn.execute(sql) 
else
	sql="update YFYMSUCO set ToTcost='"& totCost &"' , sgmemo='"& sgmemo &"', "&_
		"mdtm=getdate(), muser='"& SESSION("NETUSER")  &"'  where  sgno='"& sgno &"' and convert(char(10),pddate,111) = '"& pddate &"' and autoid ='"& AID &"' "	
	conn.execute(sql)	
end if	
set rds=nothing 
response.write sql&"<BR>"

IF autoid<>"" THEN 
	AID=autoid 
ELSE
	AID=AID 
END IF 		
for X = NNY to NDY 
	for Y = 1 to 12 
		CurrenRow= (cdbl(X*12)+Y) - cdbl(NNY*12)  
		SSmoney = request("SSYM")(CurrenRow) 		
		'response.write X &"-" & Y & SSYM &"<BR>"  
		if SSmoney<>"0" then 
			SSYM=cstr(X)&cstr(right("00"&y,2))
			sql="insert into YFYDSUCO (Autoid, sgno, sgym, cfgroup, cfdw, empid, YM, SUKM , DM ) values ( "&_
				"'"& AID &"','"& sgno &"' ,'"&  sgYM &"' , '"& cfGroup &"', '"& cfdw &"', '"& empid &"', '"& SSYM &"',  '"& SSmoney &"' , '"& DM &"'  ) " 
			response.write  X &"-" & Y &"-"& sql&"<BR>"	
			conn.execute(sql)
		end if 
	next 
next 

'response.end 

if conn.Errors.Count = 0 then 
	conn.CommitTrans	
	Set conn = Nothing
%>	<SCRIPT LANGUAGE=VBSCRIPT>
		ALERT "資料處理成功!!"&chr(13)&"DATA CommitTrans Success !!"
		OPEN "vYFYSUCO.asp" , "Fore" 
	</script>	
<%  
ELSE
	conn.RollbackTrans	
%>	<SCRIPT LANGUAGE=VBSCRIPT>
		ALERT "資料處理失敗!!"&chr(13)&"DATA CommitTrans ERROR !!"
		OPEN "VYFYSUCO.asp" , "Fore" 
	</script>	
<%	response.end 
END IF  
%>
 
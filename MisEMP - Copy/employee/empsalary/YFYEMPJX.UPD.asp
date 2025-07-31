<%@LANGUAGE="VBSCRIPT"  codepage=950%>
<!-- #include file = "../../GetSQLServerConnection.fun" --> 
<!-- #include file="../../ADOINC.inc" -->  
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<%
Response.Expires = 0
Response.Buffer = true 

JXYM = REQUEST("JXYM")
SALARYYM = REQUEST("SALARYYM")
GROUPID = REQUEST("GROUPID")
SHIFT  = REQUEST("SHIFT") 
zuno  = REQUEST("zuno") 
sts=request("sts")
whsno=request("jxwhsno")


Set conn = GetSQLServerConnection()	  

conn.BeginTrans 

if sts="NALL" then 
	PageRec = request("PageRec")
	FOR I = 1 TO PageRec 
		groupid=REQUEST("groupid")(I)
		shift=REQUEST("shift")(I)
		zuno=REQUEST("zuno")(I)
		STT=REQUEST("STT")(I)
		DESCP=REQUEST("DESCP")(I)
		HXSL=REQUEST("HXSL")(I)
		PER=REQUEST("PER")(I)
		HESO=REQUEST("HESO")(I)
		
		IF TRIM(DESCP)<>"" AND TRIM(PER)<>"" AND TRIM(HESO)<>"" AND HXSL<>""  THEN   
			sqlstr="select* from YFYMJIXO where jxwhsno='"& whsno &"' and jxym='"& JXYM &"' and salaryYM='"& SALARYYM &"' and "&_ 
				   "groupid='"& groupid &"' and shift='"& shift &"' and zuno='"&zuno &"' and stt='"& STT &"' and per='"& PER &"' " 
			Set rds = Server.CreateObject("ADODB.Recordset")					    
			rds.open sqlstr, conn, 3, 3 
			if rds.eof then 
				SQL="INSERT INTO YFYMJIXO (jxwhsno,JXYM,SALARYYM,GROUPID,SHIFT, zuno, STT,DESCP,HXSL,HESO,PER) VALUES (  "&_
					"'"&whsno&"','"& JXYM &"' , '"& SALARYYM &"' ,'"& GROUPID &"' ,'"& SHIFT &"' ,'"&ZUNO&"','"& STT &"' , "&_
					"'"& DESCP &"' , '"& HXSL &"' ,'"& HESO &"' ,'"& PER &"' ) " 
				CONN.EXECUTE(SQL)	
				response.write "A=" & sql &"<BR>"
			else
				sql="update YFYMJIXO set descp='"& descp  &"' ,   "&_
					"HXSL='"& HXSL  &"' , HESO='"& HESO  &"' , "&_
					"per='"& per &"' "&_
					"where jxwhsno='"&whsno&"' and jxym='"& jxym &"' and salaryYM='"& SALARYYM &"' and groupid='"& groupid &"' and shift='"& shift &"' "&_
					"and Zuno='"&zuno&"' and stt='"& stt &"' and per='"& PER &"' "  
				conn.execute(sql)
				response.write "B=" & sql &"<BR>"	
			end if  			
		END IF
	NEXT 	
else  
	FOR I = 1 TO 5 
		STT=REQUEST("STT")(I)
		DESCP=REQUEST("DESCP")(I)
		HXSL=REQUEST("HXSL")(I)
		PER=REQUEST("PER")(I)
		HESO=REQUEST("HESO")(I)
		
		IF TRIM(DESCP)<>"" AND TRIM(PER)<>"" AND TRIM(HESO)<>"" AND HXSL<>""  THEN 
			SQL="INSERT INTO YFYMJIXO (jxwhsno,JXYM,SALARYYM,GROUPID,zuno, SHIFT,STT,DESCP,HXSL,HESO,PER) VALUES (  "&_
				"'"& whsno &"' ,'"& JXYM &"' , '"& SALARYYM &"' ,'"& GROUPID &"' ,'"& zuno &"','"& SHIFT &"' ,'"& STT &"' , "&_
				"'"& DESCP &"' , '"& HXSL &"' ,'"& HESO &"' ,'"& PER &"' ) " 
			CONN.EXECUTE(SQL)		
			RESPONSE.WRITE "B=" & SQL&"<br>"
		END IF 		
	NEXT  
end if 
'response.end 

if conn.Errors.Count = 0 then 
	conn.CommitTrans
	Set conn = Nothing
	if sts="NALL" then%>
	<SCRIPT LANGUAGE=VBSCRIPT>
		ALERT "資料處理成功!!"&chr(13)&"DATA CommitTrans Success !!"
		OPEN "YFYEMPJX.copydata.asp" , "_self" 
	</script>	
<%	
	else
%>	<SCRIPT LANGUAGE=VBSCRIPT>
		ALERT "資料處理成功!!"&chr(13)&"DATA CommitTrans Success !!"
		OPEN "YFYEMPJX.asp" , "_self" 
	</script>	
<%  end if
ELSE
	conn.RollbackTrans	
%>	<SCRIPT LANGUAGE=VBSCRIPT>
		ALERT "資料處理失敗!!"&chr(13)&"DATA CommitTrans ERROR !!"
		OPEN "YFYEMPJX.asp" , "_self" 
	</script>	
<%	response.end 
END IF  
%>
 
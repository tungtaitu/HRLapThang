<%@LANGUAGE="VBSCRIPT"  codepage=65001%>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<!-- #include file = "../GetSQLServerConnection.fun" --> 
<!-- #include file="../ADOINC.inc" -->  

<%
Response.Expires = 0
Response.Buffer = true 

JXYM = REQUEST("JXYM")
saym = REQUEST("saym")
z1=request("z1")
  
'response.write z1
'response.end 
Set conn = GetSQLServerConnection()	  

conn.BeginTrans 
	for x = 1 to z1 
		xid = trim(request("xid")(x))
		groupid = request("groupid")(x)
		zuno = request("zuno")(x)
		shift = request("shift")(x)
		sttA = request("A")(x)
		sttB = request("B")(x)
		sttC = request("C")(x)
		sttD = request("D")(x) 
		'response.write xid &"<BR>"
		'response.write groupid &"<BR>"
		'response.write zuno &"<BR>"
		'response.write shift &"<BR>"
		if xid="1" then 
			a_desc="盈餘目標"  
		elseif xid="2" then	
			a_desc="本月事故"
		elseif xid="3" then	
			a_desc="機故時間"
			b_desc="用油量"
		elseif xid="4" then	
			a_desc="堆高機當月維修費用"
			b_desc="原紙庫存"
			c_desc="殘捲數"
		elseif xid>="5" and xid<"7" then	
			a_desc="生產M2/產能M2"	
			b_desc="本月事故"
			c_desc="績效損耗"
			d_desc="產能效率"
		else	
			a_desc="生產M2/產能M2"	
			b_desc="本月事故"
			c_desc="產量/H"		
			d_desc="欠量率"
		end if 
		
		if sttA<>"" then 
			SQL="INSERT INTO YFYMJIXO (JXYM,SALARYYM,GROUPID,ZUNO,SHIFT,xid,STT,DESCP,HXSL ) VALUES (  "&_
				"'"& JXYM &"' , '"& SAYM &"' ,'"& GROUPID &"' ,'"& ZUNO &"','"& SHIFT &"' ,'"& XID &"' , "&_
				"'"& A &"' ,'"& a_desc &"' , '"& sttA &"'   ) " 
				'CONN.EXECUTE(SQL)	
			response.write sql &"<BR>"	
		elseif sttB<>"" then 
			SQL="INSERT INTO YFYMJIXO (JXYM,SALARYYM,GROUPID,ZUNO,SHIFT,xid,STT,DESCP,HXSL ) VALUES (  "&_
				"'"& JXYM &"' , '"& SAYM &"' ,'"& GROUPID &"' ,'"& ZUNO &"','"& SHIFT &"' ,'"& XID &"' , "&_
				"'"& B &"' ,'"& B_DESC &"' , '"& sttB &"'   ) " 
				'CONN.EXECUTE(SQL)	
			response.write sql &"<BR>"			
		elseif sttC<>"" then 
			SQL="INSERT INTO YFYMJIXO (JXYM,SALARYYM,GROUPID,ZUNO,SHIFT,xid,STT,DESCP,HXSL ) VALUES (  "&_
				"'"& JXYM &"' , '"& SAYM &"' ,'"& GROUPID &"' ,'"& ZUNO &"','"& SHIFT &"' ,'"& XID &"' , "&_
				"'"& C &"' ,'"& C_DESC &"' , '"& sttC &"'   ) " 
				'CONN.EXECUTE(SQL)	
			response.write sql &"<BR>"	
		elseif sttD<>"" then 
			SQL="INSERT INTO YFYMJIXO (JXYM,SALARYYM,GROUPID,ZUNO,SHIFT,xid,STT,DESCP,HXSL ) VALUES (  "&_
				"'"& JXYM &"' , '"& SAYM &"' ,'"& GROUPID &"' ,'"& ZUNO &"','"& SHIFT &"' ,'"& XID &"' , "&_
				"'"& D &"' ,'"& d_DESC &"' , '"& sttD &"'   ) " 
				'CONN.EXECUTE(SQL)	
			response.write sql &"<BR>"				
		END IF	
 
	next	
response.end 

if conn.Errors.Count = 0 or err.number=0 then 
	conn.CommitTrans
	Set conn = Nothing
%>	<SCRIPT LANGUAGE=VBSCRIPT>
		ALERT "資料處理成功!!"&chr(13)&"DATA CommitTrans Success !!"
		OPEN "YFYEMPJX.asp" , "_self" 
	</script>	
<%  
ELSE
	conn.RollbackTrans	
%>	<SCRIPT LANGUAGE=VBSCRIPT>
		ALERT "資料處理失敗!!"&chr(13)&"DATA CommitTrans ERROR !!"
		OPEN "YFYEMPJX.asp" , "_self" 
	</script>	
<%	response.end 
END IF  
%>
 
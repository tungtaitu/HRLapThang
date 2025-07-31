<%@LANGUAGE="VBSCRIPT" codepage=65001 %>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<!-- #include file = "../../GetSQLServerConnection.fun" --> 

<%
'Session("NETUSER")=""
'SESSION("NETUSERNAME")=""

'Response.Expires = 0
'Response.Buffer = true 

Set CONN = GetSQLServerConnection()   
 
DAT1 = REQUEST("DAT1")
DAT2 = REQUEST("DAT2")
whsno = trim(request("whsno"))
groupid = trim(request("groupid"))
country = trim(request("country"))  
QUERYX = trim(request("empid1"))  

'response.write DAT1 & DAT2 & whsno & groupid & country 
tmpRec = Session("EMPHOLIDAYB")   
 
pagerec=request("pagerec") 
totalpage=request("totalpage")
conn.BeginTrans 
x = 0 
For i = 1 to totalpage 
	for j = 1 to  pagerec     
		'response.write tmpRec(i, j, 0) &"<BR>"
	    if  tmpRec(i, j, 0) = "del"  then 
	    	sql="delete empholiday  where  autoid='"& tmpRec(i, j, 24)  &"'  "   
	    	conn.execute(sql)
	    	'response.write sql&"<BR>"	  	 
	    	X = X+1        
				SQLA="UPDATE  EMPWORK SET  "&_
						 "JIA"&TRIM(tmpRec(i, j, 22))&" = JIA"&TRIM(tmpRec(i, j, 22))&"-"& TRIM(tmpRec(i, j, 23))&"  "&_
						 "WHERE  EMPID='"& TRIM(tmpRec(i, j, 1)) &"'  "&_
						 "AND WORKDAT = '"& REPLACE( TRIM(tmpRec(i, j, 17)), "/", "" ) &"'  "
				CONN.EXECUTE(SQLA) 
	    	RESPONSE.WRITE SQLA &"<br>"			 
	    	' SQLSTR = "SELECT * FROM  EMPWORK WHERE EMPID='"& TRIM(tmpRec(i, j, 1)) &"'  "&_
	    			 ' "AND WORKDAT = '"& REPLACE( TRIM(tmpRec(i, j, 17)), "/", "" ) &"'  "&_ 
	    			 ' "AND JIA"&TRIM(tmpRec(i, j, 22))&"  ='"& TRIM(tmpRec(i, j, 23)) &"'  " &_
	    			 ' "AND FLAG='JIA' " 
	    	' RESPONSE.WRITE SQLSTR   &"<br>"  	
	    	' Set rs = Server.CreateObject("ADODB.Recordset")        	 
	    	' RS.OPEN SQLSTR, CONN, 3, 3  
	    	' IF RS.EOF THEN   
	    		' SQLA="UPDATE  EMPWORK SET  "&_
	    			 ' "JIA"&TRIM(tmpRec(i, j, 22))&" = JIA"&TRIM(tmpRec(i, j, 22))&"-"& TRIM(tmpRec(i, j, 23))&"  "&_
	    			 ' "WHERE  EMPID='"& TRIM(tmpRec(i, j, 1)) &"'  "&_
	    			 ' "AND WORKDAT = '"& REPLACE( TRIM(tmpRec(i, j, 17)), "/", "" ) &"'  "   	    			 
	    	' ELSE
	    		' SQLA="DELETE EMPWORK WHERE EMPID='"& TRIM(tmpRec(i, j, 1)) &"'  "&_
	    			 ' "AND WORKDAT = '"& REPLACE( TRIM(tmpRec(i, j, 17)), "/", "" ) &"'  "&_ 
	    			 ' "AND JIA"&TRIM(tmpRec(i, j, 22))&"  ='"& TRIM(tmpRec(i, j, 23)) &"'  " &_
	    			 ' "AND FLAG='JIA' "  
	    	' END IF 		
	 
	    end if 	
    next 
Next 

'response.end 
'response.redirect "empbasicB.fore.asp?yymm="& yymm 

if ( conn.Errors.Count = 0 or err.number=0 ) and x<>0  then 
	conn.CommitTrans
	Set Session("EMPHOLIDAYB") = Nothing  
	Set conn = Nothing
%>	<SCRIPT LANGUAGE=VBSCRIPT>		
		open "empholidayB.Fore.asp?DAT1="& "<%=DAT1%>" &"&DAT2="& "<%=DAT2%>" &"&whsno="& "<%=whsno%>" &"&groupid="&  "<%=groupid%>" &"&country="& "<%=country%>" &"&empid1="& "<%=QUERYX%>" , "Fore" 
	</script>	
<%ELSE
	conn.RollbackTrans	
%>	<SCRIPT LANGUAGE=VBSCRIPT>
		ALERT "DATA CommitTrans ERROR !!"
		open "empholidayB.Fore.asp?DAT1="& "<%=DAT1%>" &"&DAT2="& "<%=DAT2%>" &"&whsno="& "<%=whsno%>" &"&groupid="&  "<%=groupid%>" &"&country="& "<%=country%>" &"&empid1="& "<%=QUERYX%>" , "Fore" 
	</script>	
<%	response.end 
END IF  
%>
  
 
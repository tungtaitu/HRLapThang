<%@LANGUAGE="VBSCRIPT" %>
<!-- #include file = "../../GetSQLServerConnection.fun" --> 
<!-- #include file="../../ADOINC.inc" -->  
<%
'Session("NETUSER")=""
'SESSION("NETUSERNAME")=""

'Response.Expires = 0
'Response.Buffer = true 

Set CONN = GetSQLServerConnection()   

yymm= request("salarymm") 


tmpRec = Session("EMPBASIC")   

pagerec=request("pagerec") 
totalpage=request("totalpage") 
totalpage=1 
 conn.BeginTrans 
x = 0 

For i = 1 to totalpage 
	for j = 1 to  pagerec  
		
	    if  tmpRec(i, j, 0) = "del"  then 
	    	sql="delete empsalarybasic where  autoid='"& tmpRec(i, j, 5)  &"'  "   
	    	conn.execute(sql)
	    	'response.write sql&"<BR>"	 
	    	x= x + 1 
	    elseif tmpRec(i, j, 0) = "upd"  then 
	    	if tmpRec(i, j, 5)="" then   
	    		if trim(tmpRec(i, j, 1))<>"" and trim(tmpRec(i, j, 3))<>"" and trim(tmpRec(i, j, 4))<>""  then 
	    			sql="insert into empsalarybasic ( func, code, bonus, descp, JOB, COUNTRY, dm, bwhsno ) values ( "&_
	    				"'"&trim(tmpRec(i, j, 1))&"' , '"&trim(tmpRec(i, j, 3))&"', '"&trim(tmpRec(i, j, 4))&"' ,  "&_
	    				"'"&trim(tmpRec(i, j, 2))&"', '"&UCASE(trim(tmpRec(i, j, 6)))&"', "&_
	    				"'"&UCASE(trim(tmpRec(i, j, 9)))&"', '"&UCASE(trim(tmpRec(i, j, 8)))&"', "&_
	    				"'"&UCASE(trim(tmpRec(i, j, 11))) &"'    ) "  
	    			conn.execute(sql) 
	    			x= x + 1  
	    		end if 
	    	else
	    		sql="update empsalarybasic set code='"& trim(tmpRec(i, j, 3)) &"' , bonus='"& trim(tmpRec(i, j, 4)) &"', "&_
	    			"dm='"&trim(tmpRec(i, j, 8)) &"',  descp='"& trim(tmpRec(i, j, 2)) &"' , COUNTRY='"& UCASE(trim(tmpRec(i, j, 9))) &"', JOB='"& UCASE(trim(tmpRec(i, j, 6))) &"' ,"&_
	    			"yymm='"&trim(tmpRec(i, j, 10)) &"', bwhsno='"&tmpRec(i, j, 11)&"'  where autoid='"& tmpRec(i, j, 5) &"'  "  
	    		conn.execute(sql) 	 
	    		x= x + 1  	
	    	end if  
	    	'response.write sql&"<BR>"	 
	    end if 	
	    
    next 
Next 

'response.end 
'response.redirect "empbasicB.fore.asp?yymm="& yymm


if conn.Errors.Count = 0 and x<>0  then 
	conn.CommitTrans
	Set Session("EMPBASIC") = Nothing 	    
	response.redirect "empbasic.Fore.asp"
	Set conn = Nothing
	
ELSE
	conn.RollbackTrans	
%>	<SCRIPT LANGUAGE=VBSCRIPT>
		ALERT "DATA CommitTrans ERROR !!"
		OPEN "empbasic.asp" , "main" 
	</script>	
<%	response.end 
END IF  
%>
  
 
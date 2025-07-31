<%@LANGUAGE="VBSCRIPT" codepage=950%>
<!-- #include file = "../../GetSQLServerConnection.fun" --> 
<!-- #include file="../../ADOINC.inc" -->  

<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta HTTP-EQUIV="refresh" >
</head>
</html>
<%
Response.Expires = 0
Response.Buffer = true 
session.codepage="950"

CurrentPage = request("CurrentPage")
TotalPage = request("TotalPage")
PageRec = request("PageRec")
gTotalPage = request("gTotalpage") 
'response.write "CurrentPage=" & CurrentPage &"<BR>"
'response.write "TotalPage=" & TotalPage &"<BR>"
'response.write "PageRec=" & PageRec &"<BR>"
'response.write "gTotalpage=" & gTotalpage &"<BR>"
tmpRec = Session("empde0401") 
Set CONN = GetSQLServerConnection()  
conn.BeginTrans
x = 0
y = ""  

for i = 1 to gTotalpage 
	for j = 1 to PageRec		
		'response.write trim(tmpRec(i,j,1)) &"<BR>"
		'response.write trim(tmpRec(i,j,3)) &"<BR>"
		'response.write trim(tmpRec(i,j,4)) &"<BR>"
		'response.write trim(tmpRec(i,j,5)) &"<BR>"
		'response.write trim(tmpRec(i,j,6)) &"<BR>"
		if trim( tmpRec(i,j,1))<>"" and  trim( tmpRec(i,j,3))<>"" and trim( tmpRec(i,j,4))<>""  and trim( tmpRec(i,j,5))<>"" and trim( tmpRec(i,j,6))>"0"  then 
			yymm = year(trim( tmpRec(i,j,3)))&right("00"& month(trim( tmpRec(i,j,3))),2) 
		 	sql="insert into empholiday(empid,JiaType,DateUP,TimeUP,DateDown,TimeDown,HHour,memo,mdtm,muser  ) values ( "&_
		 		"'"& trim(tmpRec(i,j,1)) &"','G'  , '"& trim(tmpRec(i,j,3)) &"' , '"& trim(tmpRec(i,j,4)) &"', "&_
		 		"'"& trim(tmpRec(i,j,3)) &"' ,'"& trim(tmpRec(i,j,5)) &"', '"& trim(tmpRec(i,j,6)) &"' , '因公未刷卡',  "&_
		 		"getdate(), '"& session("netuser") &"')  " 
		 	x= x+1 	
		 	conn.execute(sql) 
			response.write sql &"<BR>"	
			wkdat = replace(trim(tmpRec(i,j,3)),"/","") 
			sqlstr = "update  empwork set kzhour = kzhour-"& trim(tmpRec(i,j,6)) &", JIAG='"& trim(tmpRec(i,j,6)) &"'  where  empid='"&  trim(tmpRec(i,j,1)) &"' and  "&_
					 "workdat='"&  wkdat  &"'  " 
			conn.execute(sqlstr)		 
		end if 		 
	next
next 	

'RESPONSE.END 

if conn.Errors.Count = 0 then 
	conn.CommitTrans
	'Set Session("empworkbC") = Nothing 	    
%>	<SCRIPT LANGUAGE=VBSCRIPT>
		ALERT "資料處理成功!!"&"<%=X%>"&"筆"&chr(13)&"DATA CommitTrans SUCCESS !!"
		OPEN "empde0401.asp" , "_self" 
		'window.close()
	</script>	
<%
ELSE
	conn.RollbackTrans	
%>	<SCRIPT LANGUAGE=VBSCRIPT>
		ALERT "資料處理失敗!!"&chr(13)&"DATA CommitTrans ERROR !!"
		OPEN "empde0401.asp" , "_self" 
	</script>	
<%	response.end 
END IF  
%>
 
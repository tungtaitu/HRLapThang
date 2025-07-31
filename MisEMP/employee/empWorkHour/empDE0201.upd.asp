<%@LANGUAGE="VBSCRIPT" codepage=65001%> 
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<!-- #include file = "../../GetSQLServerConnection.fun" -->  
<%
'Response.Expires = 0
'Response.Buffer = true 

CurrentPage = request("CurrentPage")
TotalPage = request("TotalPage")
PageRec = request("PageRec")
gTotalPage = request("gTotalpage") 
response.write "CurrentPage=" & CurrentPage &"<BR>"
response.write "TotalPage=" & TotalPage &"<BR>"
response.write "PageRec=" & PageRec &"<BR>"
response.write "gTotalpage=" & gTotalpage &"<BR>"
tmpRec = Session("empde0201") 
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
		'response.write trim(tmpRec(i,j,16)) &"<BR>"
		if trim( tmpRec(i,j,1))<>"" and  trim( tmpRec(i,j,3))<>"" and trim( tmpRec(i,j,4))<>""  and trim( tmpRec(i,j,5))<>"" and trim( tmpRec(i,j,6))>"0"  then 
			sqlstr="select * from empforget where empid='"& trim(tmpRec(i,j,1)) &"' and convert(char(10), dat,111)='"& trim(tmpRec(i,j,3)) &"' and isnull(status,'')<>'D'  "
			Set rs = Server.CreateObject("ADODB.Recordset")    
			rs.open sqlstr ,conn ,3, 3 
			if rs.eof then 
				yymm = year(trim( tmpRec(i,j,3)))&right("00"& month(trim( tmpRec(i,j,3))),2) 
			 	'sql="insert into empforget(whsno, empid, lsempid, dat, timeup, timedown, toth, yymm, mdtm, muser, cab3  ) values ( "&_
				sql="insert into empforget(whsno, empid, lsempid, dat, timeup, timedown, toth, yymm, mdtm, muser, cardno  ) values ( "&_
			 		"'"& trim(tmpRec(i,j,12)) &"','"& trim(tmpRec(i,j,1)) &"'  , '"& trim(tmpRec(i,j,2)) &"' , '"& trim(tmpRec(i,j,3)) &"', "&_
			 		"'"& trim(tmpRec(i,j,4)) &"' ,'"& trim(tmpRec(i,j,5)) &"', '"& trim(tmpRec(i,j,6)) &"' , '"& yymm &"' , "&_
			 		"getdate(), '"& session("netuser") &"' ,'"& trim(tmpRec(i,j,16)) &"' )  " 
			 	x= x+1 	
			 	conn.execute(sql) 
				response.write sql &"<BR>"	 
				
				if trim(tmpRec(i,j,6)) ="" then 
					toth=0 
				else 
					 if cdbl(trim(tmpRec(i,j,6)))>=8 then 
					 	toth=8 
					 else	
					 	toth = cdbl(trim(tmpRec(i,j,6)))
					 end  if 	
				end if 	
				wkdat = replace(trim(tmpRec(i,j,3)),"/","") 
				sql2="update empwork set kzhour=kzhour-'"& toth &"', toth =toth+'"& toth &"' , forget=forget+1 where empid='"& trim(tmpRec(i,j,1)) &"' and workdat='"& wkdat &"'  "
				conn.execute(Sql2)
				response.write sql2 &"<BR>" 
				sucmsg = sucmsg & trim(tmpRec(i,j,1))&"-"&tmpRec(i, j, 10)& tmpRec(i, j, 11)&" "&trim(tmpRec(i,j,3))& " 新增成功OK .............." &"<BR>" 
			else				
				sqlx="update empforget set toth=0 where empid='X' "
				conn.execute(sqlx)
				errmsg=	errmsg & trim(tmpRec(i,j,1))&"-"&tmpRec(i, j, 10)& tmpRec(i, j, 11)&" "&trim(tmpRec(i,j,3))& " 資料已存在!!新增失敗!!(Fail)" &"<BR>"
			end  if 	
		end if 		 
	next
next 	

'RESPONSE.END 
response.clear  
session.codepage=65001
%>

<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<%
if conn.Errors.Count = 0 then 
	conn.CommitTrans
	response.write sucmsg 
	response.write "------------------------------------------------------<BR>"
	response.write errmsg 
	response.write "<P><a href=empde0201.asp>回忘刷卡資料處理</a>"
	response.end 
	'Set Session("empworkbC") = Nothing 	    
%>	<SCRIPT LANGUAGE=VBSCRIPT>
		ALERT "資料處理成功!!"&chr(13)&"DATA CommitTrans SUCCESS !!"
		OPEN "empde0201.asp" , "_self" 
		'window.close()
	</script>	
<%
ELSE
	conn.RollbackTrans		
	response.write sucmsg  
	response.write "------------------------------------------------------<BR>"
	response.write errmsg 
	response.write "<P><a href=empde0201.asp>回忘刷卡資料處理</a>"
	'response.end 
%>	<SCRIPT LANGUAGE=VBSCRIPT>
		ALERT "資料處理失敗!!"&chr(13)&"DATA CommitTrans ERROR !!"
		OPEN "empde0201.asp" , "_self" 
	</script>	
<%	response.end 
END IF  
%>
 
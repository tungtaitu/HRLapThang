<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<meta HTTP-EQUIV="Content-Type" content="text/html; charset=utf-8">
<!-- #include file = "../GetSQLServerConnection.fun" --> 
<!-- #include file="../ADOINC.inc" --> 

<%
Response.Buffer = true
Response.Expires = 0
 
CurrentPage = request("CurrentPage")
TotalPage = request("TotalPage")
PageRec = request("PageRec")
gTotalPage = request("gTotalpage")


Set conn = GetSQLServerConnection()
tmpRec = Session("YSBAE1701")

'on error resume next
conn.BeginTrans
x = 0
y = ""
for i = 1 to gTotalPage 
	for j = 1 to PageRec 
		if tmpRec(i, j , 0) = "update" then		
			'if tmpRec(i, j, 1) <> "" then 			       
			if tmpRec(i, j, 1) <> "" and Ucase(trim(tmpRec(i,j, 1)))>=(year(date) & right(0 & month(date),2)) then 			       
					sql1 = "select * from VYFYEXRT where yyyymm = '" & tmpRec(i, j, 1) & "' and code = '" & tmpRec(i, j, 2) & "'"
					set rs1=conn.execute(sql1)
					if rs1.eof then
					   sql2 = "insert into VYFYEXRT(yyyymm,code,exrt) values ('" & tmpRec(i, j, 1) & "','" & tmpRec(i, j, 2) & "' ," & tmpRec(i, j, 3) & ") "
					else
					   sql2 = "update VYFYEXRT set yyyymm = '" & tmpRec(i, j, 1) & "' , code = '" & tmpRec(i, j, 2) & "' , exrt = " & tmpRec(i, j, 3) & " where yyyymm = '" & tmpRec(i, j, 1) & "' and code = '" & tmpRec(i, j, 2) & "'"
					end if
					rs1.close
					'Response.Write sql2
					'Response.End 
					conn.Execute ( sql2 )
					x = x + 1	
					if y = "" then
					   y = "年月 : " & tmpRec(i, j, 1) & ",幣別 : " & tmpRec(i, j, 2) & ",匯率 : " & tmpRec(i, j, 3)
					else
					   y = y + "<BR>" & "稅別 : " & tmpRec(i, j, 1) & ",銷售別 : " & tmpRec(i, j, 2) & ",匯率 : " & tmpRec(i, j, 3)
					end if				
			end if 			
		else 
			if tmpRec(i, j , 0) = "del" then 
				if tmpRec(i, j, 1) <> "" and Ucase(trim(tmpRec(i,j, 1)))>=(year(date) & right(0 & month(date),2)) then 
					sql2 = "delete VYFYEXRT where yyyymm = '" & tmpRec(i, j, 1) & "' and code = '" & tmpRec(i, j, 2) & "'"
					conn.Execute ( sql2 )
					x = x + 1
					if y = "" then
					   y = "年月 : " & tmpRec(i, j, 1) & ",幣別 : " & tmpRec(i, j, 2) & ",匯率 : " & tmpRec(i, j, 3)
					else
					   y = y + "<BR>" & "年月 : " & tmpRec(i, j, 1) & ",幣別 : " & tmpRec(i, j, 2) & ",匯率 : " & tmpRec(i, j, 3)
					end if				
				end if 				
			end if 
		end if 		
	next
next 
'response.write "sql="& sql2
'response.end 
if conn.Errors.Count = 0 or err.number=0  then 
	conn.CommitTrans
	Set Session("YSBAE1701") = Nothing 	    	
	Set conn = Nothing
%>	<SCRIPT LANGUAGE=javascript>
		alert("DATA CommitTrans Success !!");
		open("yfyexrt.asp?pgid=<%=request("pgid")%>" , "main" );
	</script>	
<%	
	
ELSE
	conn.RollbackTrans	
%>	<SCRIPT LANGUAGE=javascript>
		alert("DATA CommitTrans ERROR !!");
		open("yfyexrt.asp" , "main" );
	</script>	
<%	response.end 
END IF  
%>

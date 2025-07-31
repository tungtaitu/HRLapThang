<%@LANGUAGE="VBSCRIPT" CODEPAGE=65001%>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<!-- #include file = "../GetSQLServerConnection.fun" --> 
<!-- #include file="../ADOINC.inc" -->  
<%
 

Set CONN = GetSQLServerConnection()  
f_w1=request("f_w1")
f_ct=request("f_ct") 

endym=request("endym") 
dm=request("dm") 
 
conn.BeginTrans
x = 0
y = ""  

PageRec = request("PageRec")  '總列數
TableRec = request("TableRec")  '總欄位
frows= request("frows")  '以輸入系統欄位
flines= request("flines")  '輸入系統之列數

'response.write  PageRec &" "&TableRec&" "&flines&" "&frows
for x = 1 to PageRec
	line_Code = request("line_Code")(x)
	descp = trim(request("descp")(x))
	chkNum= replace(trim(request("chkNum")(x)),",","")
	'response.write chkNum &"<BR>" 
	if chkNum>0 then  
		B0 = replace(trim(request("B0")(x)),",","")
		if b0<>"" and B0>"0"  then 
			lncode=line_Code&"B0"
			response.write lncode &"<BR>"
			'sql="exec proc_sp_yecb02_chkCode '"& lncode&"' "
			'conn.execute(Sql)
			sqlx="select * from  empsalarybasic   where country='VN' and func='aa'  and bwhsno='"&f_w1&"'  and code='"&lncode&"'  "
			response.write sqlx&"<br>"
			Set rs = Server.CreateObject("ADODB.Recordset")
			rs.Open sqlx, conn, 3, 3
			
			if not rs.eof then 
				sql="update  empsalarybasic   set bonus='"&b0&"' , descp=N'"&descp&"', yymm='"&endym&"', dm='"&dm&"' , unit='*BB', "&_
						"mdtm=getdate(), muser='"&session("netuser")&"' where  country='VN' and func='aa'  and bwhsno='"&f_w1&"'  and code='"&lncode&"' "
				conn.execute(sql)
				response.write sql &"<BR>"
			else				
				sql="insert into empsalarybasic (bwhsno, func,code,country, bonus, descp, yymm, dm , mdtm, muser , unit)  values  (  "&_
						"'"&f_w1&"','AA','"&lncode&"','VN','"&b0&"',N'"&descp&"','"&endym&"','"&dm&"',getdate(),'"&session("netuser")&"','*BB' ) "
				conn.execute(sql)
				response.write sql &"<BR>"
			end if 
			
			set rs=nothing 
		end if 	
		for zz = 5 to TableRec
			bx = replace(trim(request("Bx")( ((x-1)*(TableRec-4))+(zz-4))),",","")
			'response.write line_Code& "B"&zz-4&"  "&bx&"<BR>" 		
			descp = trim(request("descp")(x))
			chkNum= replace(trim(request("chkNum")(x)),",","")
			if bx<>"" and bx>"0"  then
				lncode=line_Code&"B"&zz-4
				if zz-4>=10 then  lncode=line_Code&"B"& chr((zz-4)+55)
				 
				response.write lncode &"<BR>"
				'sql="exec proc_sp_yecb02_chkCode '"& lncode&"' "
				'conn.execute(Sql)
				sqlx="select * from  empsalarybasic   where country='VN' and func='aa'  and bwhsno='"&f_w1&"'  and code='"&lncode&"' "
				response.write sqlx&"<br>"
				Set rs = Server.CreateObject("ADODB.Recordset")
				rs.Open sqlx, conn, 3, 3
				
				if not rs.eof then 
					sql="update  empsalarybasic   set bonus='"&bx&"' , descp=N'"&descp&"', yymm='"&endym&"', dm='"&dm&"' , unit='*BB', "&_
							"mdtm=getdate(), muser='"&session("netuser")&"' where  country='VN' and func='aa'  and bwhsno='"&f_w1&"'  and code='"&lncode&"' "
					conn.execute(sql)
					response.write sql &"<BR>"
				else				
					sql="insert into empsalarybasic (bwhsno, func,code,country, bonus, descp, yymm, dm , mdtm, muser,unit )  values  (  "&_
							"'"&f_w1&"','AA','"&lncode&"','VN','"&bx&"',N'"&descp&"','"&endym&"','"&dm&"',getdate(),'"&session("netuser")&"' ,'*BB' ) "
					conn.execute(sql)
					response.write sql &"<BR>"
				end if 
			
			end if
		next 	 
	end if 
next 
response.clear 
'RESPONSE.END 
%>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<%
if conn.Errors.Count = 0 or err.number=0 then 
	conn.CommitTrans
	Set Session("empsalary01") = Nothing 	  	
	Set conn = Nothing
%>	<SCRIPT LANGUAGE=VBSCRIPT>
		ALERT "資料處理成功!!"&chr(13)&"DATA CommitTrans Success !!"
		OPEN "YECB02.fore_vn.asp?f_w1="&"<%=f_w1%>"&"&f_ct="&"<%=f_ct%>", "_self" 
	</script>	
<% 
ELSE
	conn.RollbackTrans	
%>	<SCRIPT LANGUAGE=VBSCRIPT>
		ALERT "資料處理失敗!!"&chr(13)&"DATA CommitTrans ERROR !!"
		OPEN "YECB02.fore_vn.asp" , "_self" 
	</script>	
<%	response.end 
END IF 
%>
 
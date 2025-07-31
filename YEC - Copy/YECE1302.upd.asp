<%@LANGUAGE="VBSCRIPT" codepage=65001%>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<!-- #include file = "../GetSQLServerConnection.fun" -->
<!-- #include file="../ADOINC.inc" -->
<%
Response.Expires = 0
Response.Buffer = true

self="YECE1302"
CurrentPage = request("CurrentPage")
TotalPage = request("TotalPage")
PageRec = request("PageRec")
gTotalPage = request("gTotalpage")

Set conn = GetSQLServerConnection()


 

WHSNO = request("f_WHSNO")  

tmpRec = Session("YECE1302B") 

' response.write "years=" & firstday &"<BR>"
' response.write "endday=" & endday &"<BR>" 
' response.write "cfg=" & cfg 
'response.end 
Set CONN = GetSQLServerConnection()

conn.BeginTrans
x = 0
y = ""
 
nam=request("f_years")  
ct=request("f_country")  
whsno=whsno 
f_empid=request("f_empid")  
f_groupid	=request("f_groupid")  
for i = 1 to pagerec	
	'if trim(ct)="" then country=Ucase(trim(request("country")(i))) 	
	years = trim(request("years")(i))
	country = trim(request("country")(i))
	w1= trim(request("whsno")(i)) 
	empid= trim(request("empid")(i))    	 	
	grande= trim(request("kj")(i))     '考績  	
	bonus= replace(trim(request("bonus")(i)) ,",","")    '自動計算之獎金( VN未進位)
	Tjamt= replace(trim(request("khac")(i)) ,",","")  '其他調整
	tax= replace(trim(request("tax")(i)) ,",","")     '稅 (VN only ) 
	r_bonus= replace(trim(request("r_bonus")(i)),",","")  ' 實領獎金( VN 已進位(5000 ) )
	bodays= trim(request("bodays")(i))   '發放天數  
	js 	= trim(request("hs")(i))    '係數
	memos= trim(request("memos")(i))     '備註	
 
	indat = tmpRec(1,i,5) 
	empnam_cn = tmpRec(1,i,6)  
	empnam_vn =tmpRec(1,i,7)  
	nz = trim(request("nianzi")(i))  
	groupid = tmpRec(1,i,9)  
	dm = tmpRec(1,i,11)  
	fs= tmpRec(1,i,12)  
	
	totamAMT = replace(trim(request("basicBZM")(i)),",","") 
	
	
	sqlx="select * from empnzjj  where yymm='"&years &"' and empid='"& empid &"' "
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.open sqlx, conn, 3, 1 
	response.write sqlx &"<BR>"
	if rs.eof then 
		response.write "r_bonus =" & r_bonus &"<BR>"
		if cdbl(r_bonus)>0 then 
			sql="INSERT INTO [YFYNET].[dbo].[EmpNZJJ] "&_
					"([yymm], [whsno], [country], [groupid], [empid], [empnam_cn], [empnam_vn], [indat], [nz],  "&_
					"[totamAMT], [Bonus], [KTAXM], [Tjamt], [RealAMT], [w1], [fs], [grande], [js], [days], [months], [DM], [memos],mdtm,muser ) values ( "&_
					"'"&years&"','"&whsno&"','"&country&"','"&groupid&"','"&empid&"','"&empnam_cn&"',N'"&empnam_vn&"','"&indat&"','"&nz&"', "&_
					"'"&totamAMT&"','"&Bonus&"','"&tax&"','"&Tjamt&"','"&r_bonus&"','"& w1 &"','"&fs&"','"&grande&"','"&js&"','"&bodays&"','"&nz&"','"&dm&"', "&_
					"N'"&memos&"', getdate(), '"&session("netuser")&"' ) "
			conn.execute(sql)		 			
			
		end if 
		x = x+1		
	else
		sql="update EmpNZJJ set grande='"&grande&"' , js='"&js&"' , w1='"&w1 &"', nz='"&nz&"', memos=N'"&memos&"', Tjamt='"&Tjamt&"',Bonus='"&Bonus&"',RealAMT='"&r_bonus&"', KTAXM='"& tax &"',  "&_
				"days='"&bodays&"',totamAMT='"&totamAMT&"' , mdtm=getdate(), muser='"&session("netuser")&"' "&_
				"where yymm='"&years &"' and empid='"& empid &"' "		 
		conn.execute(sql)	
		x = x+1		
	end if 
	response.write sql &"<BR>"
 
	rs.close : set rs=nothing 

	
	'sqlx="select * from empnzkh  where years='"&years &"' and empid='"& empid &"' "
	'response.write sqlx &"<BR>"
	'Set rs = Server.CreateObject("ADODB.Recordset")
	'rs.open sqlx, conn, 3, 3 
	'if rs.eof then 		
	'	sqlb="insert into EmpNZKH([years], [whsno], [country],[empid],   [fensu], [kj], [mdtm], [muser]) values ( "&_					
	'			 "'"&years&"','"&whsno&"','"&country&"','"&empid&"','"&fs&"','"&grande&"',N'"&memos&"', getdate(), '"&session("netuser")&"' ) "
	'	conn.execute(sqlb)	 
	'	response.write sqlb
	'else
	'	if ucase(trim(rs("kj")))<>ucase(trim(grande)) then 
	'		sqlb="insert into EmpNZKH([years], [whsno], [country],[empid],   [fensu], [kj], [mdtm], [muser]) values ( "&_					
	'				 "'"&years&"','"&whsno&"','"&country&"','"&empid&"','"&fs&"','"&grande&"', getdate(), '"&session("netuser")&"' ) "
	'		conn.execute(sqlb)	 
	'		response.write sqlb
	'	end if 		
	'end if  
	'set rs=nothing  
	
next
response.write err.number &"<BR>"
response.write conn.errors.count &"<BR>"

for g =0 to conn.errors.count-1
	response.write conn.errors.item(g)&"<br>"
	response.write Err.Description
next  

response.clear
'RESPONSE.END

%>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<%

if err.number = 0 then
	conn.CommitTrans
	Set Session("YECE1302B") = Nothing
	conn.close : Set conn = Nothing 
%>	<SCRIPT LANGUAGE=VBSCRIPT>
		ALERT "資料處理成功!!"&"<%=X%>"&"筆"&chr(13)&"DATA CommitTrans Success !!"
		OPEN "<%=self%>.fore.asp?f_years="&"<%=nam%>"&"&f_country="&"<%=ct%>"&"&f_whsno="&"<%=WHSNO%>"&"&f_empid="&"<%=f_empid%>"&"&f_groupid="&"<%=f_groupid%>" , "_self"
	</script>
<% 
ELSE	
	conn.RollbackTrans 
%>	<SCRIPT LANGUAGE=VBSCRIPT>
		ALERT "資料處理失敗!!"&chr(13)&"DATA CommitTrans ERROR !!"
		OPEN "<%=self%>.fore.asp" , "_self"
	</script>
<%  response.end
END IF
%>
 
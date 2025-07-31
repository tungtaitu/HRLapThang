<%@LANGUAGE="VBSCRIPT"  codepage=65001%>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<!-- #include file = "../GetSQLServerConnection.fun" --> 
<!-- #include file="../ADOINC.inc" -->  

<%
Response.Expires = 0
Response.Buffer = true 
self="YEIE0101"
JXYM = REQUEST("JXYM")
khym = REQUEST("khym") 
z1=REQUEST("pagerec")
tt=REQUEST("tt") 
khweek = REQUEST("khweek") 

F_whsno= REQUEST("F_whsno") 
F_country= REQUEST("F_country") 
F_groupid= REQUEST("F_groupid") 
F_zuno= REQUEST("F_zuno") 
F_shift= REQUEST("F_shift") 

days = REQUEST("days")
if khweek = "" then 
	tt = cdbl(days)\7 
else
	tt = 1 
end if 			
'response.write z1
'response.end 
Set conn = GetSQLServerConnection()

xx=0
y=0
conn.BeginTrans 
	for x = 1 to z1 
		empid = request("empid")(x)
		memo = trim(request("memo")(x))
		KHW = request("whsno")(x)
		KHG = request("groupid")(x)
		khZ = request("zuno")(x)
		khS = request("shift")(x) 
		hjsts = request("hjsts")(x)
		cqmemo = request("cqmemo")(x)
		'response.write "hjsts="  & hjsts &"<BR>"
		for y = 1 to tt
			'response.write x &"-"
			'response.write y &"-"
			'response.write (x-1)*tt+y &"<BR>"			
			monthweek = request("monthweek")(y)			
			fnA=request("fensuA")((X-1)*tt+y)
			fnB=request("fensuB")((X-1)*tt+y)
			fnC=request("fensuC")((X-1)*tt+y)
			fnD=request("fensuD")((X-1)*tt+y)
			aid=request("aid")((X-1)*tt+y)
			if FnA="" then fnA="0" 
			if FnB="" then fnB="0" 
			if FnC="" then fnC="0" 
			if FnD="" then fnD="0" 
			'response.write empid &"-" & monthweek & "-" & fna & "-" & fnB & "-" & fnC & "-" & fnD &"<BR>"
			if empid<>"" then 
				if monthweek<>""  then 
					sqln="select * from empKHB where empid='"& empid &"' and khym='"& khym &"' and khweek='"& monthweek &"' "
					Set rs = Server.CreateObject("ADODB.Recordset")	  
					rs.open sqln, conn, 1, 3  
					if rs.eof then 
						sql="insert into empKHB (empid, khym, khweek, fnA, fnB, fnC, fnD, mdtm, muser, memo, khw, khg, khz, khs ) values ( "&_
							"'"& empid &"','"& khym &"','"& monthweek &"','"& fnA &"','"& fnB &"','"&fnC&"','"&fnD&"', "&_
							"getdate(),'"&session("netuser")&"',N'"& memo &"','"& khw &"','"& khg &"','"& khz &"','"& khs &"' ) " 
						response.write sql &"<BR>"	
						conn.execute(Sql)
						xx = xx+ 1
					else
						sql="update empKHB set fnA='"&fnA&"' , fnB='"&fnB&"', fnC='"&fnC&"' , fnD='"&fnD&"' "&_
							",mdtm=getdate(), muser='"& session("netuser") &"', memo=N'"& memo &"'  "&_
							",khw='"& khw &"',khg='"& khg &"',khz='"& khz &"',khs='"& khs&"'  where "&_
							"empid='"& empid &"' and khym='"& khym &"' and khweek='"& monthweek &"' "
						conn.execute(Sql)	
						response.write sql &"<BR>"	
						xx = xx + 1
					end if 	
					set rs=nothing
				end if		
			end if  
			if request("ht")="Y" then 
				sqlt="select isnull(status,'') khsts, * from empkhb where aid='"& aid &"' "
				Set rst = Server.CreateObject("ADODB.Recordset")	  
				rst.open sqlt, conn, 1, 3   
				if not rst.eof then 
					if rst("khsts")="Y" then
						sql2="update empkhb set cqmemo=N'"& cqmemo &"' , hjmdtm=getdate(), hjmuser='"& session("netuser") &"'  "&_
							 "where aid='"& aid &"' " 					 
						conn.execute(sql2)
						xx = xx + 1
					else
						if hjsts="Y" then 
							sql2="update empkhb set cqmemo=N'"& cqmemo &"' , status='Y' , hjmdtm=getdate(), hjmuser='"& session("netuser") &"'  "&_
								 "where aid='"& aid &"' and isnull(status,'')<>'Y' " 
							conn.execute(Sql2)	 
							xx = xx + 1
						'else
						'	sql2="update empkhb set cqmemo=N'"& cqmemo &"' , hjmdtm=getdate(), hjmuser='"& session("netuser") &"'  "&_
						'		 "where aid='"& aid &"' " 					
						'	conn.execute(Sql2)	 
						'	xx = xx + 1
						end if 												
					end if 
				end if 
				set rst=nothing 
				'response.write "***sq2***=" & sql2 &"<BR>"
			end if 	
		next		
	next	
'response.write "xx=" & xx
'response.end 

if xx > "0" and ( conn.Errors.Count = 0 or err.number=0  ) then 
	conn.CommitTrans
	Set conn = Nothing
%>	<SCRIPT LANGUAGE=VBSCRIPT>
		ALERT "資料處理成功!!"&chr(13)&"DATA CommitTrans Success !!"
		OPEN "<%=self%>.Fore.asp?khym="&"<%=khym%>"&"&F_whsno="&"<%=F_whsno%>"&"&F_groupid="&"<%=F_groupid%>"&"&F_zuno="&"<%=F_zuno%>"&"&F_shift="&"<%=F_shift%>"&"&F_country="&"<%=F_country%>"&"&khweek="&"<%=khweek%>", "_self" 
	</script>	
<%  
ELSE
	conn.RollbackTrans	
	if xx = "0" then 
%>		<SCRIPT LANGUAGE=VBSCRIPT>
		ALERT "無任何處理資料!!"&chr(13)&"DATA CommitTrans ERROR !!"
		OPEN "<%=self%>.asp" , "_self" 
		</script>	
<% 		response.end 
	else
%>		<SCRIPT LANGUAGE=VBSCRIPT>
		ALERT "資料處理失敗!!"&chr(13)&"DATA CommitTrans ERROR !!"
		OPEN "<%=self%>.asp" , "_self" 
		</script>	
<%	response.end 
	end if 
END IF  
%>
 
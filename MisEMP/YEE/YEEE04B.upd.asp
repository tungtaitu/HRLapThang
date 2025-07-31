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
'tmpRec = Session("EMPHOLIDAYB")   
 
pagerec=request("pagerec") 
totalpage=request("totalpage")
zjdays = request("zjdays") 
'response.write cdbl(pagerec)+cdbl(zjdays)
'response.end 
conn.BeginTrans 
x = 0 
f_xid= ""
xidstr=""
for j = 1 to  cdbl(pagerec) 
		'response.write tmpRec(i, j, 0) &"<BR>"
		autoid = trim(request("autoid")(j))
		jb = Ucase(trim(request("jiatype")(j)))
		place = Ucase(trim(request("place")(j)))
		op = trim(request("op")(j)) 
		'HHDAT1= trim(request("HHDAT1")(j)) 
		'jiamemo= trim(request("jiamemo")(j)) 
		xid= trim(request("xid")(j)) 
		
		if op="D" then 
			sql="delete empholiday where autoid='"& autoid &"' "
		else
			'if autoid<>"" then 
				sql="update empholiday set place='"& place &"', jiatype='"& jb &"' , xid='"& xid &"' where autoid='"& autoid &"'  "	
			'else
			'	if jb<>"" and HHDAT1<>"" and place<>""  then 
			'		sql="insert into empHoliday ( empid, jiaType, DateUP, TimeUP, DateDown, TimeDown, HHour, memo, Muser, place, xid   ) values ( "&_
			'				"'"& QUERYX &"', '"& jb &"', '"& HHDAT1 &"', '08:00', '"& HHDAT1 &"', '17:00', "&_
			'				"'8', '"& jiamemo &"', '"& session("NETUSER") &"' , '"& place &"' , '"& xid &"'  ) "  												
			'	end if 
			'end if 	
		end if 	
		response.write sql &"<BR>" 
		conn.execute(sql)  
		
		if   xidstr <> xid then
			f_xid	= f_xid &  xid &"," 			
		end if 	
		xidstr = xid 
		response.write f_xid 	&"<BR>"
next  

for kk = 1 to 3 
	jb2=Ucase(trim(request("jb2")(kk)))
	n_dat1=trim(request("n_dat1")(kk))
	n_dat2=trim(request("n_dat2")(kk))
	n_place=Ucase(trim(request("n_place")(kk)))
	n_xid=Ucase(trim(request("n_xid")(kk)))
	jbmemo=trim(request("jbmemo")(kk))
	if jb2<>"" and n_dat1<>"" and n_dat2<>"" and n_place<>"" then 
		sql="insert into empHoliday ( empid, jiaType, DateUP, TimeUP, DateDown, TimeDown, HHour, memo, Muser, place, xid, mdtm    ) "&_
				"select '"& QUERYX &"',  '"& jb2&"', convert(char(10),dat,111) , '08:00', convert(char(10),dat,111) , '17:00', '8',  "&_
				"'"& jbmemo &"' , '"& session("netuser") &"' , '"&n_place&"','"&n_xid&"', getdate()  "&_ 				
				"from ydbmcale where convert(char(10),dat,111) between '"& n_dat1 &"' and '"&n_dat2&"' "
		response.write sql &"<BR>"		 
		conn.execute(sql)
		
		if n_dat2 > dat2 then  DAT2 = n_dat2 			
	end if 
next 
' response.write y &"<BR>"
'response.end   
' for zz = 0 to ubound(split(f_xid,","))	
	' n_xid = split(f_xid,",")(zz)	
	' if trim(n_xid)<>""     then 
		 ' sqlx="select jiatype, xid , min(convert(char(10),dateup,111) as mindat  , max(convert(char(10),dateup,111) as maxdat "&_
				  ' "from empHoliday where isnull(xid,'')='"& n_xid  &"' and group by jiatype, xid "
		 ' response.write " f .sql="& sqlx &"<BR>"
	' end if 
' next 	

'response.end 
'response.redirect "empbasicB.fore.asp?yymm="& yymm 



if ( conn.Errors.Count = 0 or err.number=0 )  then 
	conn.CommitTrans	
	Set conn = Nothing
%>	<SCRIPT LANGUAGE=VBSCRIPT>		
		open "yeee04.showdata.asp?DAT1="& "<%=DAT1%>" &"&DAT2="& "<%=DAT2%>" &"&whsno="& "<%=whsno%>" &"&groupid="&  "<%=groupid%>" &"&country="& "<%=country%>" &"&empid1="& "<%=QUERYX%>" , "_self" 
	</script>	
<%ELSE
	conn.RollbackTrans	
%>	<SCRIPT LANGUAGE=VBSCRIPT>
		ALERT "DATA CommitTrans ERROR !!"
		open "yeee04.showdata.asp?DAT1="& "<%=DAT1%>" &"&DAT2="& "<%=DAT2%>" &"&whsno="& "<%=whsno%>" &"&groupid="&  "<%=groupid%>" &"&country="& "<%=country%>" &"&empid1="& "<%=QUERYX%>" , "_self" 
	</script>	
<%	response.end 
END IF  
%>
  
 
<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<!-- #include file = "../GetSQLServerConnection.fun" -->
<!-- #include file="../ADOINC.inc" -->
<%
'Session("NETUSER")=""
'SESSION("NETUSERNAME")=""

'Response.Expires = 0
'Response.Buffer = true

if session("netuser")="" then 
	response.write "使用者帳號為空!!請重新登入!!"
	response.end 
end if 	 
Set CONN = GetSQLServerConnection()

empautoid = TRIM(REQUEST("empautoid"))
EMPID=TRIM(REQUEST("EMPID"))	'員工編號
INDAT=TRIM(REQUEST("INDAT"))	'到職日
WHSNO=TRIM(REQUEST("WHSNO"))	'廠別
UNITNO=TRIM(REQUEST("UNITNO"))	'處/所
GROUPID=TRIM(REQUEST("GROUPID"))	'組/部門
ZUNO=TRIM(REQUEST("ZUNO"))	'單位
EMPNAM_CN=TRIM(REQUEST("NAM_CN"))	'姓名(中)
EMPNAM_VN=TRIM(REQUEST("NAM_VN"))	'姓名(越)
COUNTRY=TRIM(REQUEST("COUNTRY"))	'國籍
BYY=TRIM(REQUEST("BYY"))
BMM=TRIM(REQUEST("BMM"))
BDD=TRIM(REQUEST("BDD"))
AGES=TRIM(REQUEST("AGES"))	'年齡
SEX=TRIM(REQUEST("SEXSTR"))	'性別
JOB=TRIM(REQUEST("JOB"))	'職等
PERSONID=TRIM(REQUEST("PERSONID"))	'身分証字號
taxcode=TRIM(REQUEST("taxCOde"))	'MST
BHDAT=TRIM(REQUEST("BHDAT"))	'簽約日(保險日)
PASSPORTNO=TRIM(REQUEST("PASSPORTNO"))	'護照號碼
VISANO=TRIM(REQUEST("VISANO"))	'簽證號碼
PDUEDATE=TRIM(REQUEST("PDUEDATE"))	'護照有效期
VDUEDATE=TRIM(REQUEST("VDUEDATE"))	'簽證有效期
PHONE=TRIM(REQUEST("PHONE"))	'聯絡電話
MOBILEPHONE=TRIM(REQUEST("MOBILEPHONE"))	'手機
HOMEADDR=replace(TRIM(REQUEST("HOMEADDR")),"'","")	'聯絡地址
EMAIL=TRIM(REQUEST("EMAIL"))	'EMAIL
OUTDAT=TRIM(REQUEST("OUTDAT"))	'離職日 
MEMO=TRIM(REQUEST("MEMO"))	'其他說明
GTDAT=TRIM(REQUEST("GTDAT"))	'加入工團(年月)
marryed = REQUEST("marryed") '婚姻狀況
SCHOOL = REQUEST("SCHOOL") '婚姻狀況
PHOTOS=TRIM(REQUEST("PHOTOS"))	'照片檔名 
grps  = request("grps")  '工時計算基準
masobh  = request("masobh")  '保險號碼

IF MEMO<>"" THEN
	MEMOSTR = REPLACE(MEMO, "'", "" )
	MEMOSTR = REPLACE (MEMOSTR, vbCrLf ,"<br>")
END IF
'Steven add 20200628. Save data [WiseEye].[dbo].[UserInfo]
'Begin-------------------------------------
UserFullName=EMPNAM_VN
UserSex=1
IF EMPNAM_VN="" THEN
	UserFullName=EMPNAM_CN
END IF
IF SEX="M" THEN
	UserSex=0
END IF
sql_InsWiseEye="declare @UserEnrollNumber int ; set @UserEnrollNumber=(select max(UserEnrollNumber)+1 FROM [WiseEye].[dbo].[UserInfo]); "&_
		"if not exists ( select * from [WiseEye].[dbo].[UserInfo] where UserFullCode='"&EMPID&"'  )  "&_ 
		"insert into [WiseEye].[dbo].[UserInfo] ( [UserFullCode],[UserFullName],[UserLastName] ,[UserEnrollNumber],[UserEnrollName],[UserHireDay] ,[UserEnabled],[UserIDD],[UserSex],[UserCardNo],[UserBirthPlace] ) "& _ 
		"values ( '"& EMPID &"' , N'"& UserFullName &"', N'"& UserFullName &"', @UserEnrollNumber,'"& EMPID &"' ,'"& INDAT &"',1,0,'"&UserSex&"','0000000000','0') "
sql_UpdWiseEye="Update [WiseEye].[dbo].[UserInfo] set UserFullName='"&UserFullName&"',UserLastName='"&UserFullName&"',UserSex='"&UserSex&"', UserHireDay='"&INDAT&"' where UserFullCode='"&EMPID&"'"
'END--------------------------------------

SHIFT = REQUEST("SHIFT") 
nowmonth  = request("nowmonth") 
studyjob  = request("studyjob")  
INDATym=left(indat,4)&mid(indat,6,2)   

BANKID = request("BANKID") 

NowDat=year(date())& "/"& right("00"&month(date()),2) & "/"& right("00"&day(date()),2) 
Nowym = year(date())& right("00"&month(date()),2)
'--------------------------------
totalpage = request("totalpage")
currentpage = request("currentpage")
RecordInDB = request("RecordInDB") 

if session("netuser")="" then 
	response.write "請重新登入!!"
	response.end
end if 	

conn.BeginTrans  

IF REQUEST("ACT")="EMPEDIT"  THEN
	'年假(特休),以每月1日為基礎,工作滿一個月年假一天,隔年3個月內需休完,年資滿5,10...年,年假13 ,14...天
	IF RIGHT(INDAT,2)<>"01" THEN
		IF MID(INDAT,6,2)+1>12 THEN
			CALCDAT=TRIM(CSTR(YEAR(INDAT)+1))&"/01/01"
		ELSE
			CALCDAT=TRIM(CSTR(YEAR(INDAT)))&"/"&TRIM(CSTR(MONTH(INDAT)+1))&"/01"
		END IF
	ELSE
		CALCDAT = INDAT
	END IF
	'特休天數 ----------------------- +	 每滿5年加一天年假
	TXD=DATEDIFF("M",CALCDAT, DATE()) + ( FIX(CINT(YEAR(DATE())-YEAR(INDAT))/CINT(5)))
	'RESPONSE.WRITE CALCDAT & "-"&TXD &"<br>"
	'RESPONSE.END 
	
	if TRIM(OUTDAT)<>"" THEN 
		SQL="UPDATE EMPFILE SET INDAT='"& INDAT &"',    "&_
			"EMPNAM_CN='"& EMPNAM_CN &"', EMPNAM_VN=N'"& EMPNAM_VN &"', "&_
			"COUNTRY='"& COUNTRY &"', BYY='"& BYY &"', BMM='"& BMM &"', BDD='"& BDD &"',  AGES='"& AGES &"', SEX='"& SEX &"',"&_
			"PERSONID='"& PERSONID &"', BHDAT='"& BHDAT &"', PASSPORTNO=N'"& PASSPORTNO &"', bankid='"& bankid &"', "&_
			"VISANO='"& VISANO &"', PDUEDATE='"& PDUEDATE &"', VDUEDATE='"& VDUEDATE &"', PHONE='"& PHONE &"', "&_
			"MOBILEPHONE='"& MOBILEPHONE &"', HOMEADDR=N'"& HOMEADDR &"', EMAIL='"& EMAIL &"', OUTDAT='"& OUTDAT &"', "&_
			"GTDAT='"& GTDAT &"', MEMO='"& MEMOSTR &"', PHOTOS='"& PHOTOS &"' ,mdtm=getdate(), muser='"& session("NETuser") &"', "&_
			"MARRYED='"& MARRYED &"' , SCHOOL='"& SCHOOL &"' , grps='"& grps &"', studyjob=N'"& studyjob &"' , taxcode='"&taxcode&"' "&_
			"WHERE AUTOID='"& empautoid &"' AND EMPID='"& EMPID &"' ;"
	ELSE
		SQL="UPDATE EMPFILE SET INDAT='"& INDAT &"',"&_
			"EMPNAM_CN='"& EMPNAM_CN &"', EMPNAM_VN=N'"& EMPNAM_VN &"', "&_
			"COUNTRY='"& COUNTRY &"', BYY='"& BYY &"', BMM='"& BMM &"', BDD='"& BDD &"',  AGES='"& AGES &"', SEX='"& SEX &"',"&_
			"PERSONID='"& PERSONID &"', BHDAT='"& BHDAT &"', PASSPORTNO=N'"& PASSPORTNO &"',bankid='"& bankid &"', "&_
			"VISANO='"& VISANO &"', PDUEDATE='"& PDUEDATE &"', VDUEDATE='"& VDUEDATE &"', PHONE='"& PHONE &"', "&_
			"MOBILEPHONE='"& MOBILEPHONE &"', HOMEADDR=N'"& HOMEADDR &"', EMAIL='"& EMAIL &"', OUTDAT=NULL,  "&_
			"GTDAT='"& GTDAT &"', MEMO='"& MEMOSTR &"', PHOTOS='"& PHOTOS &"' ,mdtm=getdate(), muser='"& session("NETuser") &"', "&_
			"MARRYED='"& MARRYED &"' , SCHOOL='"& SCHOOL &"' , grps='"& grps &"',  studyjob=N'"& studyjob &"' , taxcode='"&taxcode&"'  "&_
			"WHERE AUTOID='"& empautoid &"' AND EMPID='"& EMPID &"' ;"
	END IF	
	sql=sql&sql_UpdWiseEye
	conn.execute(sql)
	'RESPONSE.WRITE SQL &"<BR>"
	
	sqlb="select * from bempJ where yymm >='"& INDATym &"' and empid='"& empid &"' "
	Set rst1 = Server.CreateObject("ADODB.Recordset")
	rst1.open sqlb , conn, 3,3
	if rst1.eof then 
		sqlx="insert into BempJ ( yymm, empid, whsno, country, job, memo, mdtm, muser ) values ( "&_
			 " '"& nowmonth &"','"& empid  &"','"& whsno &"','"& country  &"','"& job &"', '', "&_
			 "getdate() ,'"& Session("NETUSER") &"' ) " 
	else
		sqlx="update BempJ set job='"& job &"' , mdtm=getdate(), muser='"& session("NETuser") &"' where empid='"& empid &"' and yymm>='"& nowmonth &"' " 
	end if
	conn.execute(sqlx)
	set rst1=nothing  
	
	sqlC="select  * from bempg where yymm >='"& INDATym &"' and empid='"& empid &"' order by yymm desc "  		 
	'response.write sqlc&"<BR>"
	Set rst2 = Server.CreateObject("ADODB.Recordset")
	rst2.open sqlC, conn, 3,3  
	if rst2.eof then  
		sqlz="insert into BempG ( empid, whsno, country, groupid, zuno, shift, memo, mdtm, muser, yymm  ) values ( "&_
			 "'"& empid &"','"& whsno &"','"& country &"','"& groupid &"','"& zuno &"','"& shift &"','', "&_
			 "getdate(), '"& session("NETuser") &"','"& INDATym &"' ) " 
		conn.execute(sqlz) 		 
	else	 
		sqlz="update BempG set whsno='"& whsno &"', groupid='"& groupid &"', zuno='"& zuno &"', mdtm=getdate() , "&_
			 "shift='"& shift &"', muser='"& session("NETuser") &"' "&_
			 "where empid='"& empid &"' and yymm >= '"& nowmonth  &"' "	
		conn.execute(sqlz)	 
	end if
 	
	response.write sqlz &"<BR>"
	RESPONSE.END 
	if  conn.errors.count=0  OR ERR.NUMBER=0 then 
		conn.CommitTrans
		conn.close
		set conn=nothing
		'RESPONSE.REDIRECT "empfile.EDIT.ASP?GROUPID="& GROUPID &"&EMPID="& EMPID
%>		<script language=vbscript>			
			'open "empfile.EDIT.asp?empid1="&"<%=EMPID%>" , "_self",  "top=10, left=10, width=550, scrollbars=yes"		
			window.close()
		</script>	
<%	else
		conn.RollbackTrans
		conn.close
		set conn=nothing
		%>
		<script language=vbscript>
			alert "資料處理失敗!!Dara commit Error!!"
			'open "empfile.EDIT.asp?empid1="&"<%=EMPID%>" , "Fore",  "top=10, left=10, width=550, scrollbars=yes"
			window.close()
		</script>
<%
	end if	
	'response.end 
ELSEIF REQUEST("ACT")="EMPDEL"  THEN
	SQL="UPDATE EMPFILE SET STATUS='D', MDTM=GETDATE(), MUSER='"&SESSION("NETUSER")&"' WHERE EMPID='"& EMPID &"' "
	conn.execute(sql) 	
	if  conn.errors.count=0  OR ERR.NUMBER=0 then 
		conn.CommitTrans	
%>		<script language=vbscript>
			'alert "資料處理失敗!!Dara commit Error!!"
			'open "empfile.EDIT.asp?empid1="&"<%=EMPID%>" , "Fore",  "top=10, left=10, width=550, scrollbars=yes"
			window.close()
		</script>  
<%	else
		conn.RollbackTrans %>
		<script language=vbscript>
			'alert "資料處理失敗!!Dara commit Error!!"
			'open "empfile.EDIT.asp?empid1="&"<%=EMPID%>" , "Fore",  "top=10, left=10, width=550, scrollbars=yes"
			window.close()
		</script> 		
<%		response.end 
	end if	
ELSEIF REQUEST("ACT")="EMPADDNEW" THEN
	
	sql="if not exists (select * from empfile where empid='"& EMPID &"' ) INSERT INTO EMPFILE (EMPID, INDAT, TX, EMPNAM_CN, EMPNAM_VN, COUNTRY, "&_
		"BYY, BMM, BDD, AGES, SEX, PERSONID, BHDAT, PASSPORTNO, PDUEDATE, VISANO, VDUEDATE, PHONE, "&_
		"MOBILEPHONE, HOMEADDR, EMAIL, GTDAT, MEMO, MDTM, MUSER,marryed, SCHOOL, taxcode ) VALUES ( "&_
		"'"& EMPID &"', '"& INDAT &"', '0', N'"& EMPNAM_CN &"', "&_
		"N'"& EMPNAM_VN &"', '"& COUNTRY &"', '"& BYY &"', '"& BMM &"', '"& BDD &"', '"& AGES &"', "&_
		"'"& SEX &"','"& PERSONID &"', '"& BHDAT &"' , N'"& PASSPORTNO &"' ,'"& PDUEDATE &"' , "&_
		"'"& VISANO &"', '"& VDUEDATE &"', '"& PHONE &"', '"& MOBILEPHONE &"' , N'"& HOMEADDR &"' ,'"& EMAIL &"' , "&_
 		"'"& GTDAT &"', '"& MEMO &"', GETDATE() , '"& SESSION("NETUSER") &"', '"& marryed &"', '"& SCHOOL &"','"& taxCode &"' )  ; "&_
		"if not exists ( select * from empfileB where empid='"&EMPID&"'  )  "&_ 
		"insert into empfileB ( empid,  sobh, mdtm, muser , b_whsno, b_groupid, b_zuno, b_shift, b_job ) "& _ 
		"values ( '"& EMPID &"' , '"& masobh &"', getdate(), '"& SESSION("NETUSER") &"','"& whsno &"' ,'"& groupid &"','"& zuno &"','"& shift &"'  , '"& job &"') ; "
	sql=sql&sql_InsWiseEye
		
	conn.execute(sql)
 	response.write sql &"<BR>"	
	'response.end
	
 	sqlb="select * from bempJ where yymm >='"& INDATym &"' and empid='"& empid &"' "
	Set rst1 = Server.CreateObject("ADODB.Recordset")
	rst1.open sqlb , conn, 3,3
	if rst1.eof then 		
		sqlx="insert into BempJ ( yymm, empid, whsno, country, job, memo, mdtm, muser ) values ( "&_
			 " '"& nowmonth &"','"& empid  &"','"& whsno &"','"& country  &"','"& job &"', '', "&_
			 "getdate() ,'"& Session("NETUSER") &"' ) " 
	else
		sqlx="update BempJ set job='"& job &"' , mdtm=getdate(), muser='"& session("NETuser") &"' where empid='"& empid &"' and yymm>='"& nowmonth &"' " 
	end if
	conn.execute(sqlx)
	'response.write sqlX &"<BR>"
	set rst1=nothing
	sqlz="if not exists ( select * from BempG where empid='"& empid &"' and  yymm ='"& INDATym &"' ) insert into BempG ( empid, whsno, country, groupid, zuno, shift, memo, mdtm, muser, yymm ) values ( "&_
		 "'"& empid &"','"& whsno &"','"& country &"','"& groupid &"','"& zuno &"','"& shift &"','', "&_
		 "getdate(), '"& session("NETuser") &"','"& INDATym &"' ) "  
	conn.execute(sqlz)	 
	'response.write sqlz &"<BR>"		 
	'RESPONSE.END 
	if  conn.errors.count=0  then 
		conn.CommitTrans
%>		<script language=vbscript>
			'open "empfile.EDIT.asp?empid=<%=EMPID%>" , "Fore",  "top=10, left=10, width=550, scrollbars=yes"
			alert ("員工新增成功(OK)!!")
			OPEN "YEBE0101.ASP" , "Fore"
		</script>
<%	else
		conn.RollbackTrans%>
		<script language=vbscript>
			alert "DATA ERROR !!"
			OPEN "YEBE0101.ASP" , "Fore"
		</script>
<%
	'RESPONSE.REDIRECT "empfile.fore.ASP"
	end if 
	response.end
END IF
%>
 
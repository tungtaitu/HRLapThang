<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<!-- #include file = "../GetSQLServerConnection.fun" -->
<!-- #include file="../ADOINC.inc" -->
<%

self="YEBE0102"

'宣告變數
'Dim MSU
'Dim intCount
'建立 aspSmartUpload 物件
'Set MSU=Server.CreateObject("aspSmartUpload.SmartUpload")

'filename=Request.QueryString("H1")
'savefilename=Server.MapPath("/yfyemp/yeb/pic")&"\"&filename   '存檔檔名

'執行上傳
'MSU.CodePage = "utf-8"
'MSU.Upload
'將檔案存放到指定位置，這裡的指定位置可以使用相對路徑或絕對路徑
'intCount = MSU.Save("/yfyemp/yeb/pic")
'MSU.AllowedFilesList = "jpg,gif"  '指定副檔名
'fileFlag= MSU.Files("FILE1").Filename
'Response.Write "fileFlag : "
'Response.Write fileFlag & "<br>"
'response.end
'if fileFlag<>"" then
'	MSU.Files.Item(1).SaveAs("/yfyemp/yeb/pic/temp.jpg")
'	Set fso = CreateObject("Scripting.FileSystemObject")
'	fso.CopyFile Server.MapPath("/yfyemp/yeb/pic/temp.jpg"), savefilename
'	fso.DeleteFile Server.MapPath("/yfyemp/yeb/pic/temp.jpg")
'	Set fso=Nothing
'	Response.Write(filename & " file(s) uploaded.")
'end if
'如表單使用 ENCTYPE="multipart/form-data" 的方式 ,需以下列方式取得資料

empid=trim(Request("empid"))   '工號
empautoid = TRIM(Request("empautoid"))
emptype = Request("emptype")     '類別 *
act = Request("act")     '執行類別 *

INDAT=TRIM(Request("INDAT"))	'到職日
WHSNO=TRIM(Request("WHSNO"))	'廠別
UNITNO=TRIM(Request("UNITNO"))	'處/所
GROUPID=TRIM(Request("GROUPID"))	'組/部門
ZUNO=TRIM(Request("ZUNO"))	'單位
EMPNAM_CN=TRIM(Request("NAM_CN"))	'姓名(中)
EMPNAM_VN=TRIM(Request("NAM_VN"))	'姓名(越)
COUNTRY=TRIM(Request("COUNTRY"))	'國籍
BYY=TRIM(Request("BYY"))
BMM=TRIM(Request("BMM"))
BDD=TRIM(Request("BDD"))
AGES=TRIM(Request("AGES"))	'年齡
SEX=TRIM(Request("SEXSTR"))	'性別
JOB=TRIM(Request("JOB"))	'職等
PERSONID=TRIM(Request("PERSONID"))	'身分証字號
taxcode=TRIM(REQUEST("taxCOde"))	'MST

BHDAT=TRIM(Request("BHDAT"))	'簽約日(保險日)
VISANO=TRIM(Request("VISANO"))	'簽證號碼
PASSPORTNO=TRIM(Request("PASSPORTNO"))	'護照號碼
PDUEDATE=TRIM(Request("PDUEDATE"))	'護照有效期
pissuedate=TRIM(Request("pissuedate"))	'護照簽發日 *
VDUEDATE=TRIM(Request("VDUEDATE"))	'簽證有效期
PHONE=TRIM(Request("PHONE"))	'聯絡電話
MOBILEPHONE=TRIM(Request("MOBILEPHONE"))	'手機
'HOMEADDR=TRIM(Request("HOMEADDR"))	'聯絡地址
HOMEADDR=replace(TRIM(REQUEST("HOMEADDR")),"'","")
EMAIL=TRIM(Request("EMAIL"))	'EMAIL
OUTDAT=TRIM(Request("OUTDAT"))	'離職日
MEMO=TRIM(Request("MEMO"))	'備註
IF MEMO<>"" THEN  '備註
	MEMOSTR = REPLACE(MEMO, "'", "" )
	MEMOSTR = REPLACE (MEMOSTR, vbCrLf ,"<br>")
END IF
GTDAT=TRIM(Request("GTDAT"))	'加入工團(年月)
marryed = trim(Request("marryed"))   '婚姻狀況
SCHOOL = trim(Request("SCHOOL"))  '教育程度
'PHOTOS=TRIM(REQUEST("PHOTOS"))	'照片檔名
if fileFlag<>"" then
	photos = empid & ".jpg"
else
	photos=""
end if
grps  = Request("grps")  '工時計算基準
BANKID = trim(Request("BANKID"))  '銀行帳號
SHIFT = Request("SHIFT")  '班別
studyjob  = trim(Request("studyjob"))  '職能學習

WKD_No = trim(Request("WKD_No"))  '工作証號碼 *B
WKD_DueDate = trim(Request("WKD_DueDate"))  '工作證到期日 *B
experience = trim(Request("experience"))  '經歷 *B
urgent_person =  Request("urgent_person")  '緊急聯絡人*B
releation = trim(Request("releation"))  '關係*B
urgent_phone = trim(Request("urgent_phone"))  '緊急聯繫電話*B
urgent_mobile = trim(Request("urgent_mobile"))  '國內聯繫手機*B
urgent_addr = trim(Request("urgent_addr")) '國內地址*B
bh_person = trim(Request("bh_person"))   '保險受益人*B
bh_personID = trim(Request("bh_personID"))  '受益人身分證號*B
masobh = trim(Request("masobh"))  '保險號*B
nowmonth  =  Request("nowmonth") 

whsno_acc = request("whsno_acc")

RESPONSE.WRITE "MMMM" & urgent_person  &"<br>"
RESPONSE.WRITE "act" & act  &"<br>"
INDATym=left(indat,4)&mid(indat,6,2)
NowDat=year(date())& "/"& right("00"&month(date()),2) & "/"& right("00"&day(date()),2)
nowym = year(date())&right("00"&month(date()),2)

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
	
sql_UpdWiseEye="Update [WiseEye].[dbo].[UserInfo] set UserFullName=N'"&UserFullName&"',UserLastName=N'"&UserFullName&"',UserSex='"&UserSex&"', UserHireDay='"&INDAT&"' where UserFullCode='"&EMPID&"'"
'END--------------------------------------
'--------------------------------
Set CONN = GetSQLServerConnection()

conn.BeginTrans
'response.write "ACT="&ACT
'response.end
IF ACT="EMPEDIT"  THEN 
	
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
	'RESPONSE.WRITE HOMEADDR & "-"&TXD &"<br>"
	'RESPONSE.END
	'-----EMPFILE -----------------------------------------------------------------------------------------------------------------------------------------------
	if TRIM(OUTDAT)<>"" THEN
		SQL="UPDATE EMPFILE SET emptype='"& emptype &"', INDAT='"& INDAT &"', "&_
			"EMPNAM_CN=N'"& EMPNAM_CN &"', EMPNAM_VN=N'"& EMPNAM_VN &"', "&_
			"COUNTRY='"& COUNTRY &"', BYY='"& BYY &"', BMM='"& BMM &"', BDD='"& BDD &"',  AGES='"& AGES &"', SEX='"& SEX &"',"&_
			"PERSONID='"& PERSONID &"', BHDAT='"& BHDAT &"', PASSPORTNO=N'"& PASSPORTNO &"', bankid='"& bankid &"', "&_
			"VISANO=N'"& VISANO &"', PISSUEDATE='"& pissuedate &"', PDUEDATE='"& PDUEDATE &"', VDUEDATE='"& VDUEDATE &"', PHONE='"& PHONE &"', "&_
			"MOBILEPHONE='"& MOBILEPHONE &"', HOMEADDR=N'"& HOMEADDR &"', EMAIL='"& EMAIL &"', OUTDAT='"& OUTDAT &"', "&_
			"GTDAT='"& GTDAT &"', MEMO=N'"& MEMOSTR &"',  mdtm=getdate(), muser='"& session("NETuser") &"', "&_
			"MARRYED='"& MARRYED &"' , SCHOOL='"& SCHOOL &"' ,studyjob=N'"& studyjob &"', grps='"& grps &"', taxcode='"&taxcode&"'  "&_
			"WHERE AUTOID='"& empautoid &"' AND EMPID='"& EMPID &"' ;"
	ELSE
		SQL="UPDATE EMPFILE SET emptype='"& emptype &"', INDAT='"& INDAT &"',"&_
			"EMPNAM_CN=N'"& EMPNAM_CN &"', EMPNAM_VN=N'"& EMPNAM_VN &"', "&_
			"COUNTRY='"& COUNTRY &"', BYY='"& BYY &"', BMM='"& BMM &"', BDD='"& BDD &"',  AGES='"& AGES &"', SEX='"& SEX &"',"&_
			"PERSONID='"& PERSONID &"', BHDAT='"& BHDAT &"', PASSPORTNO=N'"& PASSPORTNO &"',bankid='"& bankid &"', "&_
			"VISANO=N'"& VISANO &"',PISSUEDATE='"& pissuedate &"',  PDUEDATE='"& PDUEDATE &"', VDUEDATE='"& VDUEDATE &"', PHONE='"& PHONE &"', "&_
			"MOBILEPHONE='"& MOBILEPHONE &"', HOMEADDR=N'"& HOMEADDR &"', EMAIL='"& EMAIL &"', OUTDAT=NULL,  "&_
			"GTDAT='"& GTDAT &"', MEMO=N'"& MEMOSTR &"',  mdtm=getdate(), muser='"& session("NETuser") &"', "&_
			"MARRYED='"& MARRYED &"' , SCHOOL='"& SCHOOL &"' ,studyjob=N'"& studyjob &"', grps='"& grps &"', taxcode='"&taxcode&"'  "&_
			"WHERE AUTOID='"& empautoid &"' AND EMPID='"& EMPID &"' ;"  
	END IF
	SQL=SQL&sql_UpdWiseEye	
	conn.execute(SQL)
	
	RESPONSE.WRITE SQL &"<BR>"	
	'-----EMPFILEB -----------------------------------------------------------------------------------------------------------------------------------------------
	sqlbx="select * from empfileb where empid='"& empid &"'  "
	Set rsb = Server.CreateObject("ADODB.Recordset")
	rsb.open sqlbx, conn, 1, 3
	if rsb.eof then
	 	sqlB="insert into empfileB (empid,WKD_No,WKD_dueDate,experience,urgent_person,releation,"&_
	 		 "urgent_addr,urgent_tel,urgent_mobile,bh_person,bh_personID, mdtm, muser , sobh ,b_whsno,b_groupid, b_zuno , b_shift , b_job ) values ( "&_
	 		 "'"& empid &"','"& WKD_No &"','"& WKD_dueDate &"',N'"& experience &"','"& urgent_person &"','"& releation &"','"& urgent_addr &"', "&_
	 		 "'"& urgent_phone &"','"& urgent_mobile &"','"& bh_person &"','"& bh_personID &"',GETDATE() ,'"& SESSION("NETUSER") &"' "&_ 
			 " ,'"& masobh &"', '"& whsno &"' ,'"& groupid &"','"& zuno &"','"& shift &"','"& job &"' ) "
	else
		sqlB="update EMPFILEB SET WKD_No='"& WKD_No &"', WKD_dueDate='"& WKD_dueDate &"',experience=N'"&experience&"',  "&_
			 "urgent_person='"&urgent_person &"', releation='"& releation &"', "&_
			 "urgent_addr='"& urgent_addr&"',urgent_tel='"& urgent_phone &"', urgent_mobile='"& urgent_mobile &"', "&_
			 "bh_person='"& bh_person &"',bh_personID='"&bh_personID&"',MDTM=GETDATE(), MUSER='"& SESSION("NETUSER") &"' "&_
			 ",sobh='"& masobh &"',  b_whsno='"& whsno &"' ,b_groupid='"& groupid &"', b_zuno= '"& zuno &"', b_shift= '"& shift &"', b_job = '"& job &"'   "&_
			 "WHERE EMPID='"& EMPID &"' "
	end if
	CONN.EXECUTE(sqlB)
	SET  RSB=NOTHING

	response.write sqlB&"<BR>"
	'if nowym = INDATym then 
	old_whsno = request("old_whsno")  '原部門
	old_grp = request("old_grp")  '原部門
	old_zuno = request("old_zuno")  '原組別
	old_shift = request("old_shift")  '原班別
	old_job = request("old_job")  '原職務 
	
	'-----bempJ  職務----------------------------------------------------------------------------------------------------------------------------------------------- 
		sqlb="select * from bempJ where yymm ='"& nowym &"' and empid='"& empid &"' "
		Set rst1 = Server.CreateObject("ADODB.Recordset")
		rst1.open sqlb , conn, 3,3
		if rst1.eof then
			sqlx="insert into BempJ ( yymm, empid, whsno, country, job, memo, mdtm, muser ) values ( "&_
				 " '"& nowym &"','"& empid  &"','"& whsno &"','"& country  &"','"& job &"', '', "&_
				 "getdate() ,'"& Session("NETUSER") &"' ) "
			conn.execute(sqlx) 	
		end if	
		set rst1=nothing
		if old_job<> job then 	''職務是否更新    
			sqlx="update BempJ set job='"& job &"' , mdtm=getdate(), muser='"& session("NETuser") &"'  "&_
				 "where empid='"& empid &"' and yymm='"& nowym &"' "
			conn.execute(sqlx) 		
			response.write sqlx &"<BR>" 
		end if 
	'-----bempg  單位部門 ----------------------------------------------------------------------------------------------------------------------------------------------- 		  
	sqlC="select  * from bempg where yymm ='"& nowym &"' and empid='"& empid &"'   "
	'response.write sqlc&"<BR>"
	Set rst2 = Server.CreateObject("ADODB.Recordset")
	rst2.open sqlC, conn, 3,3
	if rst2.eof then
		sqlz="insert into BempG ( yymm,empid, whsno, country, groupid, zuno, shift, memo, mdtm, muser ) values ( "&_
			 "'"& nowym &"','"& empid &"','"& whsno &"','"& country &"','"& groupid &"','"& zuno &"','"& shift &"','', "&_
			 "getdate(), '"& session("NETuser") &"' ) "
		conn.execute(sqlz)
	end if 
	set rst2=nothing   
		
	if ( old_whsno<>whsno or old_grp<>groupid or old_zuno<>zuno or old_shift<>shift ) then  		
		sqlz="update BempG set whsno='"& whsno &"' ,mdtm=getdate() , muser='"& session("NETuser") &"' , groupid='"& groupid &"' , shift='"& shift &"',  "&_
			 "zuno='"& zuno &"' where empid='"& empid &"' and yymm ='"& nowym &"'  "
		conn.execute(sqlz)  		
	end if 
	response.write sqlz &"<BR>" 
	
	'立帳單位  'new add elin 201308 
	sql="if exists ( select * from empfile_acc where empid='"& empid &"' )   "&_
			"update empfile_acc set whsno_acc='"&whsno_acc&"' , worknum='"& bhdat &"' where empid='"& empid &"'  "&_
			"else "&_
			"insert into empfile_acc ( empid, whsno_acc, country , worknum ) values ( '"&empid&"','"&whsno_acc&"','"&country&"','"&bhdat&"' ) "
	conn.execute(Sql)		
	'RESPONSE.END
	if  conn.errors.count=0  OR ERR.NUMBER=0 then
		conn.CommitTrans
		conn.close 
		set conn=nothing
		'RESPONSE.REDIRECT "empfile.EDIT.ASP?GROUPID="& GROUPID &"&EMPID="& EMPID
%>		<script language=vbscript>
			'alert "ok"
			'open "empfile.EDIT.asp?empid1="&"<%=EMPID%>" , "_self",  "top=10, left=10, width=550, scrollbars=yes" 
			alert "資料處理成功"
			parent.close()
		</script>
<%
	else
		conn.RollbackTrans 
		conn.close 
		set conn=nothing		
%>
		<script language=vbscript>
			alert "資料處理失敗!!Dara commit Error!!"
			'open "empfile.EDIT.asp?empid1="&"<%=EMPID%>" , "Fore",  "top=10, left=10, width=550, scrollbars=yes"
			parent.close()
		</script>
<%
	end if
	response.end  '----- 修改end 
ELSEIF ACT="EMPDEL"  THEN
	SQL="UPDATE EMPFILE SET STATUS='D', MDTM=GETDATE(), MUSER='"&SESSION("NETUSER")&"' WHERE EMPID='"& EMPID &"' "
	conn.execute(sql)
	RESPONSE.REDIRECT "empfile.edit.ASP"
	

ELSEIF ACT="EMPADDNEW" THEN
	'-----------------------------------------------------------------------------------------------------------------------------------------------
	SQL="INSERT INTO EMPFILE ( Emptype, EMPID, INDAT, TX,  EMPNAM_CN, EMPNAM_VN, COUNTRY, "&_
		"BYY, BMM, BDD, AGES, SEX, PERSONID, BHDAT, PASSPORTNO, PISSUEDATE, PDUEDATE, VISANO, VDUEDATE, PHONE, "&_
		"MOBILEPHONE, HOMEADDR, EMAIL, GTDAT, MEMO, MDTM, MUSER, marryed, SCHOOL,taxcode   ) VALUES ( "&_
		"'"& Emptype &"','"& EMPID &"', '"& INDAT &"', '0', N'"& EMPNAM_CN &"', "&_
		"N'"& EMPNAM_VN &"', '"& COUNTRY &"', '"& BYY &"', '"& BMM &"', '"& BDD &"', '"& AGES &"', "&_
		"'"& SEX &"', '"& PERSONID &"', '"& BHDAT &"' , N'"& PASSPORTNO &"' ,'"& PISSUEDATE &"', '"& PDUEDATE &"' , "&_
		"N'"& VISANO &"', '"& VDUEDATE &"', '"& PHONE &"', '"& MOBILEPHONE &"' , N'"& HOMEADDR &"' ,'"& EMAIL &"' , "&_
 		"'"& GTDAT &"', '"& MEMO &"', GETDATE() , '"& SESSION("NETUSER") &"', '"& marryed &"', '"& SCHOOL &"','"& taxcode &"'    ) "
 	conn.execute(sql)
 	'-----------------------------------------------------------------------------------------------------------------------------------------------
 	sqln="insert into empfileB (empid,WKD_No,WKD_dueDate,experience,urgent_person,releation,"&_
 		 "urgent_addr,urgent_tel,urgent_mobile,bh_person,bh_personID, mdtm, muser ) values ( "&_
 		 "'"& empid &"','"& WKD_No &"','"& WKD_dueDate &"',N'"& experience &"','"& urgent_person &"','"& releation &"','"& urgent_addr &"', "&_
 		 "'"& urgent_phone &"','"& urgent_mobile &"','"& bh_person &"','"& bh_personID &"',GETDATE() ,'"& SESSION("NETUSER") &"' ) "
 	conn.execute(sqln)
 	'-----------------------------------------------------------------------------------------------------------------------------------------------
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
	response.write sqlX &"<BR>"
	set rst1=nothing
	'-----------------------------------------------------------------------------------------------------------------------------------------------
	sqlz="insert into BempG (yymm, empid, whsno, country, groupid, zuno, shift, memo, mdtm, muser ) values ( "&_
		 "'"& nowmonth &"','"& empid &"','"& whsno &"','"& country &"','"& groupid &"','"& zuno &"','"& shift &"','', "&_
		 "getdate(), '"& session("NETuser") &"' ) "
	conn.execute(sqlz)

	'RESPONSE.END
	if  conn.errors.count=0 or err.number=0  then
		conn.CommitTrans
		conn.close 
		set conn=nothing

%>		<script language=vbscript>
			alert "OK"
			'open "empfile.EDIT.asp?empid=<%=EMPID%>" , "Fore",  "top=10, left=10, width=550, scrollbars=yes"
			OPEN "<%=self%>.ASP" , "_self"
		</script>
<%	else
		conn.RollbackTrans
		conn.close 
		set conn=nothing		
%>
		<script language=vbscript>
			alert "DATA ERROR !!"
			OPEN "<%=self%>.ASP" , "_self"
		</script>
<%
	'RESPONSE.REDIRECT "empfile.fore.ASP"
	end if
	response.end
END IF
%>
 
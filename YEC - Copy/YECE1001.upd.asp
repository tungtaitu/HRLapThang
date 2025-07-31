<%@LANGUAGE="VBSCRIPT" codepage=65001%>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<!-- #include file = "../GetSQLServerConnection.fun" -->
<!-- #include file="../ADOINC.inc" -->
<%
Response.Expires = 0
Response.Buffer = true

CurrentPage = request("CurrentPage")
TotalPage = request("TotalPage")
PageRec = request("PageRec")
gTotalPage = request("gTotalpage")

Set conn = GetSQLServerConnection()
tmpRec = Session("YECE1001HW")

Set CONN = GetSQLServerConnection()

conn.BeginTrans
x = 0
y = ""
YYMM=REQUEST("YYMM")

MMDAYS  = REQUEST("MMDAYS")
EMPIDSTR="" 

session.codepage=65001
for i = 1 to TotalPage
	for j = 1 to PageRec
		'RESPONSE.WRITE TotalPage &"<br>"
		'RESPONSE.WRITE PageRec &"<br>"
		if trim(tmpRec(i, j, 1))<>"" then
			'IF trim(tmpRec(i, j, 0))="UPD" THEN
			'	SQL="UPDATE EMPFILE SET BB='"& tmpRec(i, j, 19) &"', CV='"& tmpRec(i, j, 22) &"',"&_
			'		"PHU='"& tmpRec(i, j, 23) &"', KT='"& tmpRec(i, j, 24) &"', "&_
			'		"JOB='"& tmpRec(i, j, 6) &"',  "&_
			'		"MDTM_S=GETDATE(), MUSER_S='"& SESSION("NETUSER") &"' "&_
			'		"WHERE EMPID='"& TRIM(tmpRec(i, j, 1)) &"' "
			'	conn.execute(Sql) 
			'	'RESPONSE.WRITE SQL &"<br>"
			'	'RESPONSE.WRITE EMPIDSTR &"<BR>"
			'	EMPIDSTR = EMPIDSTR & "'" & TRIM(tmpRec(i, j, 1)) &"',"
			'END IF
			
			IF tmpRec(i, j, 20)="" THEN
				BB = 0
			ELSE
				BB = CDBL(tmpRec(i, j, 20))
			END IF
			IF tmpRec(i, j, 22)="" THEN
				CV = 0
			ELSE
				CV = CDBL(tmpRec(i, j, 22))
			END IF
			IF tmpRec(i, j, 23)="" THEN
				PHU = 0
			ELSE
				PHU = CDBL(tmpRec(i, j, 23))
			END IF
			IF tmpRec(i, j, 24)="" THEN
				KT=0
			ELSE
				KT=	CDBL(tmpRec(i, j, 24))
			END IF
			IF tmpRec(i, j, 25)="" THEN
				TTKH=0
			ELSE
				TTKH=	CDBL(tmpRec(i, j, 25))
			END IF 
			IF tmpRec(i, j, 26)="" THEN
				MT=0
			ELSE
				MT=	CDBL(tmpRec(i, j, 26))
			END IF 			
			
			IF tmpRec(i, j, 28)="" THEN
				TNKH=0
			ELSE
				TNKH=	CDBL(tmpRec(i, j, 28))
			END IF
			IF tmpRec(i, j, 29)="" THEN
				JX=0
			ELSE
				JX=CDBL(tmpRec(i, j, 29))
			END IF

			IF tmpRec(i, j, 30)="" THEN
				BH=0
			ELSE
				BH=	CDBL(tmpRec(i, j, 30))
			END IF

			IF tmpRec(i, j, 31)="" THEN
				QITA=0
			ELSE
				QITA=	CDBL(tmpRec(i, j, 31))
			END IF
			IF tmpRec(i, j, 32)="" THEN
				KTAXM=0
			ELSE
				KTAXM=	CDBL(tmpRec(i, j, 32))
			END IF


			IF tmpRec(i, j, 36)="" THEN
				RELTOTMONEY=0
			ELSE
				RELTOTMONEY=ROUND(CDBL(tmpRec(i, j, 36)),2)
			END IF

			IF tmpRec(i, j, 38)="" THEN
				MONEY_H=0
			ELSE
				MONEY_H=CDBL(tmpRec(i, j, 38))
			END IF

			IF tmpRec(i, j, 41)="" THEN
				TOTM=0
			ELSE
				TOTM=	CDBL(tmpRec(i, j, 41))
			END IF

			IF tmpRec(i, j, 42)="" THEN   '夜班天數
				B3=0
			ELSE
				B3=CDBL(tmpRec(i, j, 42))
			END IF

			IF tmpRec(i, j, 43)="" THEN  '夜班津貼(1天2元美金)  200801改為5USD/天
				B3M=0
			ELSE
				B3M=CDBL(tmpRec(i, j, 43))
			END IF


			'零數 (以10,000為單位)
			if trim(tmpRec(i, j, 4))="VN" then
				LAOHN= fix( RELTOTMONEY/ 10000 ) * 10000
				SOLE =  RELTOTMONEY -   LAOHN
				if cdbl(SOLE) > 5000 then
					LAOHN = LAOHN + 10000
					SOLE = RELTOTMONEY -  LAOHN
				else
					LAOHN = LAOHN
					SOLE = SOLE
				end if
			else
				LAOHN = RELTOTMONEY
				SOLE = 0
			end if


			if  CDBL(tmpRec(i, j,34)) < CDBL(MMDAYS) THEN
				BZKM = tmpRec(i,j,39)   '不足月扣款
			ELSE
				BZKM = 0
			END IF


			if  (tmpRec(i, j,44))="" THEN
				ZHUANM = 0    '轉款
			ELSE
				ZHUANM = tmpRec(i,j,44)
			END IF

			if  (tmpRec(i, j,45)) ="" THEN
				XIANM = 0   '現金
			ELSE
				XIANM = tmpRec(i,j,45)
			END IF

			workhour = cdbl(trim(tmpRec(i, j, 34)))*8  

			MEMOSTR = trim(REPLACE(tmpRec(i,j,47), "'", "" ))
			MEMOSTR = trim(REPLACE (MEMOSTR, vbCrLf ,"<br>")) 
			
			if  (tmpRec(i, j,48)) ="" THEN
				DKM = 0   '暫扣款  (未滿半年離職應扣稅 25% , 半年後全數發回) 
			ELSE
				DKM = tmpRec(i,j,48)
			END IF 
			
			acc = tmpRec(i,j,49)  
			
			SQLSTR="SELECT * FROM  EMPDSALARY WHERE YYMM='"& YYMM  &"' AND EMPID='"& trim(tmpRec(i, j, 1)) &"'  "
			Set rs = Server.CreateObject("ADODB.Recordset")
			RS.OPEN SQLSTR, CONN, 3, 3

			IF RS.EOF THEN
				SQL1="INSERT INTO EMPDSALARY ( WHSNO,COUNTRY, EMPID, indat, outdat, GROUPID,JOB,BB,CV,PHU,NN,KT,MT,TTKH,QC,TNKH, "&_
					 "JX,TBTR,H1M, H2M, H3M, B3M, TOTM, H1, H2, H3, B3, BH, HS,GT,QITA, FL, JIAA, JIAB, KZHOUR, "&_
					 "KZM, JIAAM, JIABM, MONEY_H, REAL_TOTAL, LAONH,SOLE,YYMM,MDTM,MUSER, WORKDAYS , workshour,  JISHU, LZBZJ, BZKM, KTAXM , DM,  "&_
					 "ZHUANM, XIANM , memo , userip, DKM,acc  ) VALUES ( "&_
					 "'"& tmpRec(i, j, 7) &"' , '"& TRIM(tmpRec(i, j, 4)) &"' , '"& tmpRec(i, j, 1) &"', '"& Trim(tmpRec(i, j, 5)) &"', '"& Trim(tmpRec(i, j, 27)) &"',  "&_
					 "'"& tmpRec(i, j, 9) &"' , '"& tmpRec(i, j, 6) &"'  ,  "&_
					 "'"& BB &"', '"& CV &"', '"& PHU &"'  , '0'  , '"& KT &"'  , '"& MT &"'  , '"& TTKH &"'  , "&_
					 "'0', '"& TNKH &"', '"& JX &"', '0', '0', '0', '0', '"& B3M &"', '"& TOTM &"' , "&_
					 "'0', '0', '0', '"& B3 &"', '"& BH &"'  , '0'  ,  "&_
					 "'0'  , '"& QITA &"', '0', '0', '0', '0', "&_
					 "'0', '0', '0', '"& MONEY_H &"',  "&_
					 "'"& RELTOTMONEY &"', '"& LAOHN &"', "& SOLE &", '"& YYMM &"', GETDATE(), '"& SESSION("NETUSER")&"', "&_
					 "'"& trim(tmpRec(i, j, 34)) &"','"& workhour &"',  '0', '0' , '"& BZKM &"', '"& KTAXM &"' , 'USD' , "&_
					 "'"& ZHUANM &"', '"& XIANM &"', '"& memostr &"','"& session("vnlogIP") &"', '"& DKM &"' ,'"& acc &"'   ) "
				'RESPONSE.WRITE SQL1 &"<br>"
				X = X + 1
				conn.execute(SQL1)
			ELSE
				SQL2="update EMPDSALARY set whsno = '"& trim(tmpRec(i, j, 7)) &"', GROUPID='"& trim(tmpRec(i, j, 9)) &"', JOB='"& trim(tmpRec(i, j, 6))  &"', "&_
					 "indat='"& Trim(tmpRec(i, j, 5)) &"', outdat='"& Trim(tmpRec(i, j, 27)) &"', "&_ 
					 "COUNTRY='"& Trim(tmpRec(i, j, 4)) &"' , BB='"& BB &"' , CV='"& CV &"' , PHU='"& PHU &"', "&_
					 "KT='"& KT &"' , MT='"& MT &"', TTKH='"& TTKH &"',  JX='"& JX &"',   "&_
					 "TNKH='"& TNKH &"' ,  BH='"& BH &"' , "&_
					 "QITA='"& QITA &"' , MONEY_H ='"& MONEY_H &"' , "&_
					 "REAL_TOTAL='"& RELTOTMONEY  &"' , LAONH='"& LAOHN &"' ,SOLE='"& SOLE &"', "&_
					 "TOTM='"& TOTM &"' , B3='"& B3 &"' , B3M='"& B3M &"' ,  "&_
					 "WORKDAYS='"& trim(tmpRec(i, j, 34)) &"',  ZHUANM='"& ZHUANM &"', XIANM='"& XIANM &"',  "&_
					 "workshour = '"& workhour &"', BZKM='"& BZKM &"' , KTAXM='"& KTAXM &"', DM='USD' , memo='"& memostr &"' , "&_
					 "mdtm=getdate(), muser='"& session("NETUSER") &"' , userip='"& session("vnlogIP") &"' , DKM='"& DKM &"', acc='"& acc &"'  "&_
					 "where YYMM='"& YYMM  &"' AND EMPID='"& trim(tmpRec(i, j, 1)) &"'   "
				RESPONSE.WRITE SQL2 &"<br>"
				X = X + 1
				conn.execute(SQL2)
			END IF 
			
			SQLSTR="SELECT * FROM  EMPDSALARY_BAK WHERE YYMM='"& YYMM  &"' AND EMPID='"& trim(tmpRec(i, j, 1)) &"' AND WHSNO='"& trim(tmpRec(i, j, 7)) &"'  "
			Set rsT = Server.CreateObject("ADODB.Recordset")
			RST.OPEN SQLSTR, CONN, 3, 3 
			if not rst.eof then 
				sql3="update EMPDSALARY_BAK set ZHUANM='"& ZHUANM &"', XIANM='"& XIANM &"' , dkm='"& dkm &"', memo = '"& memostr &"' "&_
					 "where YYMM='"& YYMM  &"' AND EMPID='"& trim(tmpRec(i, j, 1)) &"'   " 
				conn.execute(sql3) 	 
				response.write sql3 &"<BR>"
			end if 
		END IF
	next
next

'RESPONSE.END

if err.number = 0 then
	conn.CommitTrans
	Set Session("YECE1001HW") = Nothing
	Set conn = Nothing
%>	<SCRIPT LANGUAGE=VBSCRIPT>
		ALERT "資料處理成功!!"&"<%=X%>"&"筆"&chr(13)&"DATA CommitTrans Success !!"
		OPEN "YECE1001.asp" , "_self"
	</script>
<%
ELSE
	conn.RollbackTrans
%>	<SCRIPT LANGUAGE=VBSCRIPT>
		ALERT "資料處理失敗!!"&chr(13)&"DATA CommitTrans ERROR !!"
		OPEN "YECE1001.asp" , "_self"
	</script>
<%	response.end
END IF
%>
 
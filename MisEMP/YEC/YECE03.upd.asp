<%@LANGUAGE="VBSCRIPT" codepage=65001%>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<!-- #include file = "../GetSQLServerConnection.fun" -->
<!-- #include file="../ADOINC.inc" -->
<%
Response.Expires = 0
Response.Buffer = true

self="YECE03"
CurrentPage = request("CurrentPage")
TotalPage = request("TotalPage")
PageRec = request("PageRec")
gTotalPage = request("gTotalpage")

Set conn = GetSQLServerConnection()
tmpRec = Session("YECE03")
cfg  = request("cfg") 

firstday  = request("calcdat")
endday = request("ccdt") 

response.write "firstday=" & firstday &"<BR>"
response.write "endday=" & endday &"<BR>" 
response.write "cfg=" & cfg 
'response.end 
Set CONN = GetSQLServerConnection()

conn.BeginTrans
x = 0
y = ""
YYMM=REQUEST("YYMM")

MMDAYS  = REQUEST("MMDAYS")
EMPIDSTR=""
for i = 1 to TotalPage
	for j = 1 to PageRec
		'RESPONSE.WRITE TotalPage &"<br>"
		'RESPONSE.WRITE PageRec &"<br>"
		if trim(tmpRec(i, j, 1))<>"" then
			IF trim(tmpRec(i, j, 0))="UPD" THEN
				'SQL="UPDATE EMPFILE SET BB='"& tmpRec(i, j, 19) &"', CV='"& tmpRec(i, j, 22) &"',"&_
				'	"PHU='"& tmpRec(i, j, 23) &"', NN='"& tmpRec(i, j, 24) &"', "&_
				'	"KT='"& tmpRec(i, j, 25) &"', MT='"& tmpRec(i, j, 26) &"' , "&_
				'	"TTKH='"& tmpRec(i, j, 27) &"',JOB='"& tmpRec(i, j, 6) &"', "&_
				'	"MDTM_S=GETDATE(), MUSER_S='"& SESSION("NETUSER") &"' "&_
				'	"WHERE EMPID='"& TRIM(tmpRec(i, j, 1)) &"' "

				'RESPONSE.WRITE SQL &"<br>"
				'RESPONSE.WRITE EMPIDSTR &"<BR>"
				'conn.execute(Sql)
				EMPIDSTR = EMPIDSTR & "'" & TRIM(tmpRec(i, j, 1)) &"',"
			END  IF
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
				NN=0
			ELSE
				NN=	CDBL(tmpRec(i, j, 24))
			END IF
			IF tmpRec(i, j, 25)="" THEN
				KT=0
			ELSE
				KT=	CDBL(tmpRec(i, j, 25))
			END IF
			IF tmpRec(i, j, 26)="" THEN
				MT=0
			ELSE
				MT=	CDBL(tmpRec(i, j, 26))
			END IF
			IF tmpRec(i, j, 27)="" THEN
				TTKH=0
			ELSE
				TTKH=	CDBL(tmpRec(i, j, 27))
			END IF
			IF tmpRec(i, j, 31)="" THEN
				QC=0
			ELSE
				QC=	CDBL(tmpRec(i, j, 31))
			END IF
			IF tmpRec(i, j, 32)="" THEN
				TNKH=0
			ELSE
				TNKH=CDBL(tmpRec(i, j, 32))
			END IF
			IF tmpRec(i, j, 33)="" THEN
				TBTR=0
			ELSE
				TBTR=CDBL(tmpRec(i, j, 33))
			END IF
			IF tmpRec(i, j, 34)="" THEN
				BH=0
			ELSE
				BH=	CDBL(tmpRec(i, j, 34))
			END IF
			IF tmpRec(i, j, 35)="" THEN
				HS=0
			ELSE
				HS=	CDBL(tmpRec(i, j, 35))
			END IF
			IF tmpRec(i, j, 36)="" THEN
				GT=0
			ELSE
				GT=	CDBL(tmpRec(i, j, 36))
			END IF
			IF tmpRec(i, j, 37)="" THEN
				QITA=0
			ELSE
				QITA=	CDBL(tmpRec(i, j, 37))
			END IF
			IF tmpRec(i, j, 38)="" THEN
				MONEY_H=0
			ELSE
				MONEY_H=CDBL(tmpRec(i, j, 38))
			END IF
			IF tmpRec(i, j, 39)="" THEN
				REAL_TOTAL=0
			ELSE
				REAL_TOTAL=	CDBL(tmpRec(i, j, 39))
			END IF
			IF tmpRec(i, j, 40)="" THEN
				H1=0
			ELSE
				H1=	CDBL(tmpRec(i, j, 40))
			END IF
			IF tmpRec(i, j, 40)="" THEN
				H1=0
			ELSE
				H1=	CDBL(tmpRec(i, j, 40))
			END IF
			IF tmpRec(i, j, 41)="" THEN
				H2=0
			ELSE
				H2=	CDBL(tmpRec(i, j, 41))
			END IF
			IF tmpRec(i, j, 42)="" THEN
				H3=0
			ELSE
				H3=CDBL(tmpRec(i, j, 42))
			END IF
			IF tmpRec(i, j, 43)="" THEN
				B3=0
			ELSE
				B3=CDBL(tmpRec(i, j, 43))
			END IF
			IF tmpRec(i, j, 44)="" THEN
				KZHOUR=0
			ELSE
				KZHOUR=CDBL(tmpRec(i, j, 44))
			END IF
			IF tmpRec(i, j, 45)="" THEN
				JIAA=0
			ELSE
				JIAA=CDBL(tmpRec(i, j, 45))
			END IF
			IF tmpRec(i, j, 46)="" THEN
				JIAB=0
			ELSE
				JIAB=CDBL(tmpRec(i, j, 46))
			END IF

			IF tmpRec(i, j, 48)="" THEN
				FL=0
			ELSE
				FL=CDBL(tmpRec(i, j, 48))
			END IF

			IF tmpRec(i, j, 47)="" THEN
				RELTOTMONEY=0
			ELSE
				RELTOTMONEY=ROUND(CDBL(tmpRec(i, j, 47)),0)
			END IF
			

			IF TRIM(tmpRec(i, j, 51))="" THEN
				H1M=0
			ELSE
				H1M=round(CDBL(tmpRec(i, j, 51)),0)
			END IF

			IF TRIM(tmpRec(i, j, 52))="" THEN
				H2M=0
			ELSE
				H2M=round(CDBL(tmpRec(i, j, 52)),0)
			END IF

			IF TRIM(tmpRec(i, j, 53))="" THEN
				H3M=0
			ELSE
				H3M=round(CDBL(tmpRec(i, j, 53)),0)
			END IF

			IF TRIM(tmpRec(i, j, 54))="" THEN
				B3M=0
			ELSE
				B3M=CDBL(tmpRec(i, j, 54))
			END IF

			IF TRIM(tmpRec(i, j, 55))="" THEN
				KZM=0
			ELSE
				KZM=CDBL(tmpRec(i, j, 55))
			END IF

			IF TRIM(tmpRec(i, j, 56))="" THEN
				JIAAM=0
			ELSE
				JIAAM=CDBL(tmpRec(i, j, 56))
			END IF

			IF TRIM(tmpRec(i, j, 57))="" THEN
				JIABM=0
			ELSE
				JIABM=CDBL(tmpRec(i, j, 57))
			END IF

			IF TRIM(tmpRec(i, j, 58))="" THEN  '績效獎金
				JX=0
			ELSE
				JX=CDBL(tmpRec(i, j, 58))
			END IF

			IF TRIM(tmpRec(i, j, 61))="" THEN  '基數
				JISHU=0
			ELSE
				JISHU=CDBL(tmpRec(i, j, 61))
			END IF

			IF TRIM(tmpRec(i, j, 62))="" THEN  '離職補助金
				LZBZJ=0
			ELSE
				LZBZJ=CDBL(tmpRec(i, j, 62))
			END IF

			IF TRIM(tmpRec(i, j, 64))="" THEN
				TOTM=0
			ELSE
				TOTM=CDBL(tmpRec(i, j, 64))
			END IF
			
			if tmpRec(i, j, 4) = "VN" then 
				KTAXM = 0 
			else	
				IF TRIM(tmpRec(i, j, 66))="" THEN   '稅金
					KTAXM=0
				ELSE
					KTAXM=CDBL(tmpRec(i, j, 66))
					'KTAXM=0
				END IF
			end if 	

			if  CDBL(tmpRec(i, j,59)) < CDBL(MMDAYS) THEN
				BZKM = tmpRec(i,j,65)   '不足月扣款
			ELSE
				BZKM = 0
			END IF 

			workhour = cdbl(trim(tmpRec(i, j, 59)))*8

					
			
			'零數 (以10,000為單位) , 如當月離職(以1,000為單位) 
			if trim(tmpRec(i, j, 30))<>"" and (  (trim(tmpRec(i, j, 30)))> (firstday)   and  (trim(tmpRec(i, j, 30))) <= (endday) )   then 
				if trim(tmpRec(i, j, 4))="VN" then 
					response.write trim(tmpRec(i, j,1 ))&"LZBZJ-------------------"&"<BR>"
					LAOHN= fix( (RELTOTMONEY+LZBZJ) / 1000 ) * 1000 
					SOLE =  RELTOTMONEY - LAOHN 					
				else
					LAOHN = RELTOTMONEY
					SOLE = 0 	
				end if 	
				
			else
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
			end if 
			
			memo = trim(tmpRec(i, j, 69)) 
			
			if trim(tmpRec(i, j, 4))="VN" then
				dm="VND"
				ZHUANM = 0
				XIANM= LAOHN
				dkm = 0
			else
				dm="USD"
				ZHUANM = trim(tmpRec(i, j, 67))
				XIANM = trim(tmpRec(i, j, 68))
				dkm = round(trim(tmpRec(i, j, 70)),0)
			end if   	

			SQLSTR="SELECT * FROM  EMPDSALARY WHERE YYMM='"& YYMM  &"' AND EMPID='"& trim(tmpRec(i, j, 1)) &"' AND WHSNO='"& trim(tmpRec(i, j, 7)) &"'  "
			Set rs = Server.CreateObject("ADODB.Recordset")
			RS.OPEN SQLSTR, CONN , 3, 3

			IF RS.EOF THEN
				rs.close 
				set rs=nothing
				SQL1="INSERT INTO EMPDSALARY ( WHSNO, Country, EMPID, indat, outdat , GROUPID,JOB,BB,CV,PHU,NN,KT,MT,TTKH,QC,TNKH, "&_
					 "JX,TBTR,H1M, H2M, H3M, B3M, TOTM, H1, H2, H3, B3, BH, HS,GT,QITA, FL, JIAA, JIAB, KZHOUR, "&_
					 "KZM, JIAAM, JIABM, MONEY_H, REAL_TOTAL, LAONH,SOLE,YYMM,MDTM,MUSER, WORKDAYS , workshour,  "&_
					 "JISHU, LZBZJ, BZKM, KTAXM, dm, ZHUANM, XIANM , userIP, dkm, memo   ) VALUES ( "&_
					 "'"& tmpRec(i, j, 7) &"' ,'"& Trim(tmpRec(i, j, 4)) &"' , '"& tmpRec(i, j, 1) &"','"& tmpRec(i, j, 5) &"', '"& tmpRec(i, j, 30) &"',  '"& tmpRec(i, j, 9) &"' , '"& tmpRec(i, j, 6) &"'  ,  "&_
					 "'"& BB &"', '"& CV &"', '"& PHU &"'  , '"& NN &"'  , '"& KT &"'  , '"& MT &"'  , '"& TTKH &"'  , "&_
					 "'"& QC &"', '"& TNKH &"', '"& JX &"', '"& TBTR &"', '"& H1M &"', '"& H2M &"', '"& H3M &"', '"& B3M &"', '"& TOTM &"' , "&_
					 "'"& H1 &"', '"& H2 &"', '"& H3 &"', '"& B3 &"', '"& BH &"'  , '0'  ,  "&_
					 "'"& GT &"'  , '"& QITA &"', '"& FL &"', '"& JIAA &"', '"& JIAB &"', '"& KZHOUR &"', "&_
					 "'"& KZM &"', '"& JIAAM &"', '"& JIABM &"', '"& MONEY_H &"',  "&_
					 "'"& RELTOTMONEY &"', '"& LAOHN &"', "& SOLE &", '"& YYMM &"', GETDATE(), '"& SESSION("NETUSER")&"', "&_
					 "'"& trim(tmpRec(i, j, 59)) &"','"& workhour &"',  '"& JISHU &"', '"& LZBZJ &"' , '"& BZKM &"', '"& KTAXM &"' , "&_
					 "'"& DM &"' ,'"& ZHUANM &"', '"& XIANM &"', '"& session("vnlogIP") &"' , '"& dkm &"', '"& memo &"'    ) "
				RESPONSE.WRITE SQL1 &"<br>"
				X = X + 1
				'response.end
				conn.execute(SQL1)
			ELSE
				SQL2="update EMPDSALARY set GROUPID='"& trim(tmpRec(i, j, 9)) &"', JOB='"& trim(tmpRec(i, j, 6))  &"', "&_
					 "indat='"& Trim(tmpRec(i, j, 5)) &"', outdat='"& Trim(tmpRec(i, j, 30)) &"', "&_
					 "COUNTRY='"& Trim(tmpRec(i, j, 4)) &"' , BB='"& BB &"' , CV='"& CV &"' , PHU='"& PHU &"', NN='"& NN &"', "&_
					 "KT='"& KT &"' , MT='"& MT &"' , TTKH='"& TTKH &"', QC='"& QC &"',  "&_
					 "TNKH='"& TNKH &"' ,  BH='"& BH &"' ,  GT='"& GT &"' , "&_
					 "QITA='"& QITA &"' , MONEY_H ='"& MONEY_H &"' , "&_
					 "REAL_TOTAL='"& RELTOTMONEY  &"' , LAONH='"& LAOHN &"' ,SOLE='"& SOLE &"', "&_
					 "H1M='"& H1M  &"', H2M='"& H2M &"', H3M='"& H3M &"', B3M='"& B3M &"', TOTM='"& TOTM &"' , "&_
					 "KZM='"& KZM  &"', JIAAM='"& JIAAM &"', JIABM='"& JIABM &"', "&_
					 "H1='"& H1  &"', H2='"& H2 &"', H3='"& H3 &"', B3='"& B3 &"',  "&_
					 "KZHOUR='"& KZHOUR  &"', FL='"& FL &"', JIAA='"& JIAA &"', JIAB='"& JIAB &"', "&_
					 "WORKDAYS='"& trim(tmpRec(i, j, 59)) &"', JX='"& JX &"', TBTR='"& TBTR &"',  DM='"& DM &"' , ZHUANM='"& ZHUANM &"', XIANM='"& XIANM &"' ,    "&_
					 "JISHU='"& JISHU &"' , LZBZJ='"& LZBZJ &"' , workshour = '"& workhour &"', BZKM='"& BZKM &"' , KTAXM='"& KTAXM &"', "&_
					 "mdtm=getdate(), muser='"& session("NETUSER") &"' , userIP='"& session("vnlogIP")&"', dkm = '"& dkm &"', memo='"& memo &"' "&_
					 "where YYMM='"& YYMM  &"' AND EMPID='"& trim(tmpRec(i, j, 1)) &"' AND WHSNO='"& trim(tmpRec(i, j, 7)) &"'  "
				RESPONSE.WRITE SQL2 &"<br>"
				X = X + 1
				conn.execute(SQL2)
			END IF 
			set rs=nothing
			SQLSTR="SELECT * FROM  EMPDSALARY_BAK WHERE YYMM='"& YYMM  &"' AND EMPID='"& trim(tmpRec(i, j, 1)) &"' AND WHSNO='"& trim(tmpRec(i, j, 7)) &"'  "
			Set rsT = Server.CreateObject("ADODB.Recordset")
			RST.OPEN SQLSTR, CONN, 1, 1 
			if not rst.eof then 
				sql3="update EMPDSALARY_BAK set ZHUANM='"& ZHUANM &"', XIANM='"& XIANM &"' "&_
					 "where YYMM='"& YYMM  &"' AND EMPID='"& trim(tmpRec(i, j, 1)) &"' AND WHSNO='"& trim(tmpRec(i, j, 7)) &"' " 
				conn.execute(sql3) 	 
				response.write sql3 &"<BR>"
			end if  	 
			set rst=nothing 		
		END IF
	next
next
response.write err.number &"<BR>"
response.write conn.errors.count &"<BR>"

for g =0 to conn.errors.count-1
	response.write conn.errors.item(g)&"<br>"
	response.write Err.Description
next  

'RESPONSE.END

if err.number = 0 then
	conn.CommitTrans
	Set Session("YECE03") = Nothing
	Set conn = Nothing 
%>	<SCRIPT LANGUAGE=VBSCRIPT>
		ALERT "資料處理成功!!"&"<%=X%>"&"筆"&chr(13)&"DATA CommitTrans Success !!"
		OPEN "<%=self%>.asp" , "_self"
	</script>
<% 
ELSE
	conn.RollbackTrans 
%>	<SCRIPT LANGUAGE=VBSCRIPT>
		ALERT "資料處理失敗!!"&chr(13)&"DATA CommitTrans ERROR !!"
		OPEN "empsalaryHW.asp" , "_self"
	</script>
<%  response.end
END IF
%>
 
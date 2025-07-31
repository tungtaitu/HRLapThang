<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<!-- #include file = "../GetSQLServerConnection.fun" -->
<!-- #include file="../ADOINC.inc" -->
<%
'on error resume next
session.codepage="65001"
SELF = "YECE02"
Set conn = GetSQLServerConnection()
Set rs = Server.CreateObject("ADODB.Recordset")
YYMM=REQUEST("YYMM")
whsno = trim(request("whsno"))
unitno = trim(request("unitno"))
groupid = trim(request("groupid"))
COUNTRY = trim(request("COUNTRY"))
job = trim(request("job"))
empid1 = trim(request("empid1"))
outemp = request("outemp")
lastym = left(yymm,4) &  right("00" & cstr(right(yymm,2)-1) ,2 )
if right(yymm,2)="01"  then
	lastym = left(yymm,4)-1 &"12"
end if
shift = request("shift")

PERAGE = REQUEST("PERAGE")
nowmonth = year(date())&right("00"&month(date()),2)  
calcdt = left(YYMM,4)&"/"& right(yymm,2)&"/01"
'下個月
if right(yymm,2)="12" then
	ccdt = cstr(left(YYMM,4)+1)&"/01/01"
else
	ccdt = left(YYMM,4)&"/"& right("00" & right(yymm,2)+1,2)  &"/01"
end if
'response.write ccdt

 '一個月有幾天
cDatestr=CDate(LEFT(YYMM,4)&"/"&RIGHT(YYMM,2)&"/01")
days = DAY(cDatestr+(32-DAY(cDatestr))-DAY(cDatestr+(32-DAY(cDatestr))))   '一個月有幾天
'本月最後一天
ENDdat = CDate(LEFT(YYMM,4)&"/"&RIGHT(YYMM,2)&"/"&DAYS)
ENDdat=year(ENDdat)&"/"&right("00"&month(Enddat),2)&"/"&right("00"&day(Enddat),2)

'RESPONSE.WRITE days &"<BR>"
'RESPONSE.WRITE ENDdat &"<BR>" 

'本月假日天數 (星期日) 
SQL=" SELECT * FROM YDBMCALE WHERE CONVERT(CHAR(6),DAT,112)  ='"& YYMM &"' AND  isnull(status,'')<>'H1' "
'response.write "1=" & sql &"<BR>"
Set rsTT = Server.CreateObject("ADODB.Recordset")
RSTT.OPEN SQL, CONN, 3, 3
IF NOT RSTT.EOF THEN
	HHCNT = CDBL(RSTT.RECORDCOUNT) 
	wcnt =  CDBL(RSTT.RECORDCOUNT)  
ELSE
	HHCNT = 0
END IF
SET RSTT=NOTHING 	

'RESPONSE.WRITE HHCNT &"<br>"
'RESPONSE.END
'本月應記薪天數
MMDAYS = CDBL(days)-CDBL(HHCNT) 
'RESPONSE.WRITE  "MMDAYS="& MMDAYS &"<BR>"
'RESPONSE.END 

'年假計算日期 
if right(ENDdat,5)<="03/31" and yymm <= nowmonth  then 
	dat_s = cstr(left(yymm,4)-1) & "/04/01"	
else
	dat_s = cstr(left(yymm,4)) & "/04/01"	
end if 

'response.write "xxx=" &  tx_enddat 
'response.end 

'----------------------------------------------------------------------------------------
sqlstr = "update empwork set kzhour=0 where yymm='"& YYMM &"'  and kzhour<0 "
conn.execute(sqlstr)

if right(yymm,2) mod 2 = 0 then 
	ccx=35
else
	ccx=36
end if 		

gTotalPage = 1
PageRec = 10    'number of records per page
TableRec = 80    'number of fields per record

njyy = request("njyy")   '年度年假未修代金

'---------------------------------------------------------------------------------

	sql="select isnull(m.empid,'') eid2, case when isnull(m.empid,'')='' or ( isnull(m.jx,0) < isnull(k.RELJXM,0) )then isnull(k.RELJXM,0) else isnull(m.jx,0)  end TOTJXM, "&_
		"isnull( L.sole , 0 ) TBTR , case when isnull(m.empid,'')='' then  case when isnull(nj.nj_amt,0)>0 then isnull(nj.nj_amt,0) else 0 end  else  isnull(m.TNKH,0) end TNKH ,  isnull(m.jx,0) jx,  "&_
		"case when a.country='VN' then case when isnull(bg.empid,'')='' then 0 else bg.bhtot end else 0 end as BH  ,  "&_
		"case when a.country='VN' then case when isnull(bg.empid,'')='' then 0 else bg.gtamt end else 0 end as GT,  "&_
		"ISNULL(M.QITA,0) QITA, xj.job as Ljob , isnull(tx.njh1,0) njh1 ,  isnull(tx.njh2,0) njh2 , isnull(NOverJ,0) NOverJ ,  "&_
		"isnull(n.forget,0) forget, isnull(n.h1,0) h1, isnull(n.h2,0) h2 , isnull(n.h3,0) h3, isnull(n.b3,0) b3 ,"&_
		"isnull(JA.jiaa,0) jiaa,isnull(JB.jiab,0) jiab,isnull(Je.jiaE,0) jiaE,isnull(n.kzhour,0) kzhour, isnull(n.latefor,0) latefor, "&_
		"case when isnull(m.empid,'')<>'' then m.bb else bs.bb end as N_bb,  isnull(m.memo,'')    as salarymemo, "&_		
		"isnull(bs.bb,0) as bs_bb, isnull(bs.cv,0) as  bs_cv, isnull(bs.phu,0) as bs_phu, isnull(bs.nn,0) as bs_nn, isnull(bs.kt,0) as bs_kt,  "&_
		"isnull(bs.mt,0) as bs_mt, isnull(bs.ttkh,0) as bs_ttkh, isnull(bs.qc,0) as bs_qc,  isnull(nj.NJ_amt,0) NJ_amt,  "&_
		"isnull(nt.person_qty,0) person_qty , isnull(nt.tot_mtax,0) notax_Amt, isnull(bx.sys_value,4000000) taxSet, isnull(tt.sukm,0) sukm , a.* from  "&_
		"( select * from  view_empfile where CONVERT(CHAR(10), indat, 111)< '"& ccdt &"' and ( isnull(outdate,'')='' or outdate>'"& calcdt &"' )  "&_
		"and whsno like '"& whsno &"%'  and groupid like '"& groupid &"%' and COUNTRY like '"& COUNTRY  &"%' and  empid like '%"& empid1 &"%'  "&_
		" ) a  "&_
		"left join ( select * from empdsalary where  yymm='"& lastym &"' ) l on L.empid = a.empid and L.whsno = a.whsno "&_
		"left join ( select * from empdsalary where yymm='"& yymm &"' ) m on m.empid = a.empid  "&_
		"left join ( select * from VYFYMYJX where yymm='"& yymm &"' ) k on k.empid = a.empid "&_
		"LEFT JOIN ( select empid empidN,  (sum(isnull(forget,0)))  forget  , (sum(isnull(h1,0))) h1, (sum(isnull(h2,0))) h2, (sum(isnull(h3,0))) h3, (sum(isnull(b3,0))) b3 ,  "&_
	 	"(sum(isnull(jiaa,0))) jiaa, (sum(isnull(jiab,0))) jiab, ( sum(isnull(toth,0))) toth , ( sum(isnull(kzhour,0))) kzhour , (sum(latefor)) latefor "&_
	 	"from empwork   where yymm='"& YYMM &"' GROUP BY EMPID )  N ON N.empidN = A.EMPID  "&_
	 	"left join  (  "&_
		"select jiaType as Ja , empid as EIDA, sum(hhour) as   jiaa   from  empholiday    where  convert(char(6), dateup, 112)='"& yymm &"'  and jiatype='A'  group  by empid, jiatype   "&_
		")  JA on JA.EIDA = a.empid   "&_
		"left join ( "&_
		"select jiaType as Jb , empid as EIDB, sum(hhour) as   jiaB   from  empholiday    where  convert(char(6), dateup, 112)='"& yymm &"'  and jiatype='B'  group  by empid, jiatype   "&_
		") JB on JB.eidb = a.empid   "&_
		"left join ( "&_
		"select jiaType as Je , empid as EIDE, sum(hhour) as   jiaE   from  empholiday where convert(char(10),dateup,111) between '"& dat_s &"' and '"& enddat &"'  and jiatype='E'  group  by empid, jiatype   "&_
		") JE on JE.eidE = a.empid   "&_
		"left join ( select* from bemps where yymm='"& yymm &"' ) bs on bs.empid = a.empid  "&_
		"left join ( select* from empbhgt where yymm='"& yymm &"' ) bg on bg.empid = a.empid  "&_
		"left join ( select * from bempJ where yymm='"& yymm &"' ) XJ on xj.empid = a.empid  "&_
		"left join ( select (toth+h1+h2) as NTj, ( (h1+h2)-"& cdbl(ccx)&" ) as NOverJ , h1 as Njh1, (h2) Njh2, * from empworkper where yymm='"& yymm &"' ) TX on TX.empid = a.empid "&_
		"left join (select * from empnotax ) nt on nt.empid = a.empid "&_ 
		"left join (select* from  BasicCode  where func='tax' and sys_type='"& COUNTRY &"'  ) bx on bx.sys_type = a.country "&_ 
		"left join (select * from emptxamt where NJ_amt > 0  and Tyear ='"& njyy &"' )  nj on nj.empid = a.empid "&_
		"left join ( select empid, sum(case when dm='USD' and isnull(country,'')='VN' then  sukm*rate else sukm  end ) as sukm  from VIEW_YFYDSUCO  "&_
		" where  isnull(empid,'')<>''  and ym='"& yymm &"'  group by empid  )  TT on tt.empid = a.empid "&_		
		"where a.empid<>''  "
		if outemp="D" then
			sql=sql&" and ( isnull(a.outdat,'')<>'' and  ( a.outdate>'"& calcdt &"' and  a.outdate<='"& ccdt &"' ))  "
		end if
		if shift="C" then
			sql=sql&" and isnull(a.shift,'') NOT IN ('ALL', 'A', 'B' ) "
		ELSEif shift<>"" then 
			sql=sql&" and isnull(a.shift,'')= '"& shift &"'    "
		end if
		sql=sql&"order by a.empid   "

		'response.write sql 
		'response.end  
		
'sql="exec sp_Calc_empsalary  '"& yymm &"', '"& whsno&"', '"&unitno&"', '"&groupid&"', '"&COUNTRY&"', '"&job&"', '"&QUERYX&"', '"&outemp&"' "  		

if request("TotalPage") = "" or request("TotalPage") = "0" then
	CurrentPage = 1
	rs.Open SQL, conn, 3, 3  
	IF NOT RS.EOF THEN
		rs.PageSize = PageRec
		RecordInDB = rs.RecordCount
		TotalPage = rs.PageCount
		gTotalPage = TotalPage
	END IF
'	response.write RecordInDB 
'	response.end 
	Redim tmpRec(gTotalPage, PageRec, TableRec)   'Array

	for i = 1 to TotalPage
	 for j = 1 to PageRec
		if not rs.EOF then
				tmpRec(i, j, 0) = "no"
				tmpRec(i, j, 1) = trim(rs("empid"))
				tmpRec(i, j, 2) = trim(rs("empnam_cn"))
				tmpRec(i, j, 3) = trim(rs("empnam_vn"))
				tmpRec(i, j, 4) = rs("country")
				tmpRec(i, j, 5) = rs("nindat")
				tmpRec(i, j, 6) = rs("job")
				tmpRec(i, j, 7) = rs("whsno")
				tmpRec(i, j, 8) = rs("unitno")
				tmpRec(i, j, 9)	=RS("groupid")
				tmpRec(i, j, 10)=RS("zuno")
				tmpRec(i, j, 11)=RS("wstr")
				tmpRec(i, j, 12)=RS("ustr")
				tmpRec(i, j, 13)=RS("gstr")
				tmpRec(i, j, 14)=RS("zstr")
				tmpRec(i, j, 15)=RS("jstr")
				tmpRec(i, j, 16)=RS("cstr")
				tmpRec(i, j, 17)=RS("autoid")
				IF RS("zuno")="XX" THEN
					tmpRec(i, j, 18)=""
				ELSE
					tmpRec(i, j, 18)=RS("zuno")
				END IF
				tmpRec(i, j, 19)=RS("BB")
				tmpRec(i, j, 20)=RS("bs_bb")  '基本薪資
				tmpRec(i, j, 21)=RS("job")
				tmpRec(i, j, 22)=RS("bs_cv")  '職務加給
				tmpRec(i, j, 23)=RS("bs_PHU")		'Y獎金
				tmpRec(i, j, 24)=RS("bs_NN")  '語言加給
				tmpRec(i, j, 25)=RS("bs_KT") '技術加給
				tmpRec(i, j, 26)=RS("bs_MT") '環境加給
				tmpRec(i, j, 27)=RS("bs_TTKH")  '其他加給
				tmpRec(i, j, 28)=RS("BHDAT") '買保險日期
				tmpRec(i, j, 29)=RS("GTDAT") '工團日期
				tmpRec(i, j, 30)= trim(RS("OUTDATE")) '離職日期

				tmpRec(i, j, 32)=ROUND(RS("TNKH"),0) '其他收入
				tmpRec(i, j, 33)=ROUND(RS("TBTR"),0) '上月補款

				tmpRec(i, j, 34)=RS("BH") '保險費(-)
				tmpRec(i, j, 36)=RS("GT") '入工團費
				'tmpRec(i, j, 37)=RS("QITA") '其他扣除額
				if ( rs("eid2")="" or (cdbl(rs("sukm"))> cdbl(RS("QITA")) and cdbl(RS("QITA"))=0 ) ) then 
					tmpRec(i, j, 37)=RS("sukm") '事故扣款
				else  			
					tmpRec(i, j, 37)=RS("QITA") '其他扣除額
				end if 	

				TOTY=  CDBL( ( CDBL(tmpRec(i, j, 20))+CDBL(tmpRec(i, j, 22))+CDBL(tmpRec(i, j, 23)) )  )  'BB+CV+PHU
				if rs("country")="VN" then
					tmpRec(i, j, 38) =  round( CDBL(TOTY)/26/8,0)   '時薪
				else
					tmpRec(i, j, 38) =  round( CDBL(TOTY)/30/8,3)   '時薪
				end if
				'response.write  TOTY &"<BR>"
				'response.write  	tmpRec(i, j, 38) &"<BR>"

				'TMONEY=CDBL(TOTY)+CDBL(tmpRec(i, j, 24))+CDBL(tmpRec(i, j, 25))+CDBL(tmpRec(i, j, 26))+CDBL(tmpRec(i, j, 27))+CDBL(tmpRec(i, j, 31))+CDBL(tmpRec(i, j, 32))+CDBL(tmpRec(i, j, 33))-CDBL(tmpRec(i, j, 34))-CDBL(tmpRec(i, j, 35))-CDBL(tmpRec(i, j, 36))-CDBL(tmpRec(i, j, 37))
				'tmpRec(i, j, 39) = CDBL(TMONEY)

				tmpRec(i, j, 40) = CDBL(RS("H1"))
				tmpRec(i, j, 41) = CDBL(RS("H2"))
				tmpRec(i, j, 42) = CDBL(RS("H3"))
				tmpRec(i, j, 43) = CDBL(RS("B3"))
				tmpRec(i, j, 44) = CDBL(RS("KZHOUR"))
				tmpRec(i, j, 45) = CDBL(RS("JIAA"))
				tmpRec(i, j, 46) = CDBL(RS("JIAB"))
				'40~43 加班費(+)
				'44~46 請假或曠職(-)
				if tmpRec(i, j, 4)="VN" then
					'response.write "aaa"
					H1_money = ROUND((tmpRec(i, j, 38)*1.5) * cdbl(tmpRec(i, j, 40))+0.01,0) '平日加班工資(+) 時薪*1.5
				else
					
					H1_money = ROUND((tmpRec(i, j, 38)*1.37) * cdbl(tmpRec(i, j, 40)),0) '平日加班工資(+) 時薪*1.37(泰國)
				end if
				
				if tmpRec(i, j, 4)="VN" then
					H2_money = ROUND((tmpRec(i, j, 38)*2) * cdbl(tmpRec(i, j, 41))+0.01,0) '休假加班工資(+)時薪*2
				else
					H2_money = ROUND((tmpRec(i, j, 38)*1)* cdbl(tmpRec(i, j, 41)),0)   '平日加班工資(+) 時薪*1(泰國)
				end if
				if tmpRec(i, j, 4)="VN" then
					H3_money = ROUND((tmpRec(i, j, 38)*3) * cdbl(tmpRec(i, j, 42))+0.01,0) '節日加班工資(+)時薪*3
				else
					H3_money = 0
				end if
				if tmpRec(i, j, 4)="VN" then
					b3_money = ROUND((tmpRec(i, j, 38)*0.3) * cdbl(tmpRec(i, j, 43))+0.01,0) '夜班加班工資(+)時薪*0.3
				else
					b3_money = 0
				end if
				kz_money = ROUND(tmpRec(i, j, 38) * tmpRec(i, j, 44),0)
				jiaa_money = ROUND(tmpRec(i, j, 38) * tmpRec(i, j, 45),0)
				jiab_money = ROUND(tmpRec(i, j, 38) * tmpRec(i, j, 46),0)

				'tmpRec(i, j, 47) = CDBL(TMONEY)+ CDBL(H1_money)+CDBL(H2_money)+CDBL(H3_money)+CDBL(b3_money)-CDBL(kz_money)-CDBL(jiaa_money)-CDBL(jiab_money)
				tmpRec(i, j, 48) = cdbl(rs("forget"))+cdbl(rs("latefor"))
				'--總加班工資
				tmpRec(i, j, 49) = CDBL(H1_MONEY) + CDBL(H2_money) + CDBL(H3_money) + CDBL(b3_money)
				'時假
				tmpRec(i, j, 50) = kz_money + jiaa_money + jiab_money

				tmpRec(i, j, 51) = H1_money
				tmpRec(i, j, 52) = H2_money
				tmpRec(i, j, 53) = H3_money
				tmpRec(i, j, 54) = B3_money
				tmpRec(i, j, 55) = kz_money
				tmpRec(i, j, 56) = jiaa_money
				tmpRec(i, j, 57) = jiab_money
				tmpRec(i, j, 58) = rs("TOTJXM")  '績效獎金  
				
				'response.write "h1_money=" & h1_money  &"<BR>"
				'response.write "h2_money=" & h2_money  &"<BR>"
				'response.write "h3_money=" & h3_money  &"<BR>"

				'response.write H1_money &"<BR>"
				'response.end 

				'員工工作天數(記薪天數)
				SQLX="SELECT EMPID, ISNULL(SUM(HHOUR),0) HHOUR  FROM EMPHOLIDAY WHERE EMPID='"& tmpRec(i, j, 1) &"' AND CONVERT(CHAR(6), DATEUP,112)='"& YYMM &"' "&_
		 				  "AND JIATYPE in ('A','B', 'F' ) GROUP BY EMPID "
		 		'response.write sqlx &"<BR>"
				Set rDs = Server.CreateObject("ADODB.Recordset")
		 		RDS.OPEN  SQLX, CONN, 3, 3
		 		IF NOT RDS.EOF THEN   '員工本月請事病假天數
		 			A4=FIX(CDBL(RDS("HHOUR"))/8)
		 		ELSE
		 			A4=0
		 		END IF
				'response.write "A4="& A4 &"<BR>"
				SET RDS=NOTHING
				'1.本月離職員工(不含1日) 從本月1日計算至離職日前一天
		 		IF trim(tmpRec(i, j, 30))="" THEN  '未離職
		 			MWORKDAYS = CDBL(days) - CDBL(HHCNT)
		 			tmpRec(i, j, 59) = MWORKDAYS 
		 			'response.write "xx" &"<BR>"
		 			'response.write tmpRec(i, j, 59) &"<BR>"
		 		ELSE
		 			IF  tmpRec(i, j, 30) >= ccdt THEN  '非本月離職
		 				MWORKDAYS = CDBL(days) - CDBL(HHCNT)
		 				tmpRec(i, j, 59) = MWORKDAYS
		 			ELSE
		 				'response.write tmpRec(i, j, 30) &"<BR>"
			 			A1=DATEDIFF("D",CDATE(calcdt),CDATE(tmpRec(i, j, 30)) )  '從1日到離職日天數
			 			'RESPONSE.WRITE A1 &"<br>"
			 			'到離職日前的假日有幾天
			 			SQLS1="SELECT * FROM YDBMCALE WHERE CONVERT(CHAR(6),DAT,112)='"& YYMM &"' AND  CONVERT(CHAR(10),DAT,111)>= '"& tmpRec(i, j, 5)&"'  AND  CONVERT(CHAR(10),DAT,111)< '"& tmpRec(i, j, 30)&"' AND STATUS IN ('H2', 'H3' ) "
			 			'RESPONSE.WRITE SQLS1 &"<br>"
			 			Set rDs = Server.CreateObject("ADODB.Recordset")
			 			RDS.OPEN  SQLS1, CONN, 3, 3
			 			IF NOT RDS.EOF THEN   '到離職日前的假日有幾天
			 				A2=RDS.RECORDCOUNT
			 			ELSE
			 				A2=0
			 			END IF
			 			SET RDS=NOTHING
			 			'RESPONSE.WRITE A2 &"<br>"
			 			SQLS2="SELECT EMPID, SUM(HHOUR) HHOUR  FROM EMPHOLIDAY WHERE EMPID='"& tmpRec(i, j, 1) &"' AND CONVERT(CHAR(6), DATEUP,112)='"& YYMM &"' "&_
			 				  "AND JIATYPE IN ('A','B') AND CONVERT(CHAR(10), DATEUP,111)<'"& tmpRec(i, j, 30) &"' GROUP BY EMPID "
			 			Set rDs = Server.CreateObject("ADODB.Recordset")
			 			RDS.OPEN  SQLS2, CONN, 3, 3
			 			'RESPONSE.WRITE SQLS2
			 			IF NOT RDS.EOF THEN   '離職前請事假OR病假有幾天
			 				A3=FIX(CDBL(RDS("HHOUR"))/8)
			 			ELSE
			 				A3=0
			 			END IF
			 			SET RDS=NOTHING
			 			'RESPONSE.WRITE "A1=" &A1&"<br>"
			 			'RESPONSE.WRITE "A2=" &A2&"<br>"
			 			'RESPONSE.WRITE "A3=" &A3&"<br>"
			 			'MWORKDAYS = CDBL(A1)-CDBL(A2)-CDBL(A3) '**********本月工作天數**********
			 			MWORKDAYS = CDBL(A1)  '-CDBL(A3) '**********本月工作天數,員工離職休假日應計**********
			 			comdays = CDBL(A1)-CDBL(A2)-CDBL(A3) '扣飯錢天數
			 			tmpRec(i, j, 59)  = MWORKDAYS    '**********本月工作天數**********
			 		END IF
		 		END IF
		 		'RESPONSE.WRITE  MWORKDAYS
		 		'2.本月新進員工 從到職日計算到本月底
		 		IF CDATE(tmpRec(i, j, 5))>CDATE(calcdt) THEN
		 			iF tmpRec(i, j, 30)="" THEN
			 			A1= DATEDIFF("D", CDATE(tmpRec(i, j, 5)), CDATE(ENDdat))
			 			'RESPONSE.WRITE "x=" & a1 &"<br>"

			 			SQLS1="SELECT * FROM YDBMCALE WHERE CONVERT(CHAR(6),DAT,112)='"& YYMM &"' AND  CONVERT(CHAR(10),DAT,111)>= '"& tmpRec(i, j, 5)&"' AND STATUS IN ('H2', 'H3' ) "
			 			'RESPONSE.WRITE SQLS1 &"<br>"
			 			Set rDs = Server.CreateObject("ADODB.Recordset")
			 			RDS.OPEN  SQLS1, CONN, 3, 3
			 			IF NOT RDS.EOF THEN   '到職後到月底的假日有幾天
			 				A2=RDS.RECORDCOUNT
			 			ELSE
			 				A2=0
			 			END IF
			 			SET RDS=NOTHING
			 			'RESPONSE.WRITE A2&"<br>"
			 			SQLS2="SELECT EMPID, SUM(HHOUR) HHOUR  FROM EMPHOLIDAY WHERE EMPID='"& tmpRec(i, j, 1) &"' AND CONVERT(CHAR(6), DATEUP,112)='"& YYMM &"' "&_
			 				  "AND JIATYPE IN ('A','B') AND CONVERT(CHAR(10), DATEUP,111)>='"& tmpRec(i, j, 5) &"' GROUP BY EMPID "
			 			Set rDs = Server.CreateObject("ADODB.Recordset")
			 			RDS.OPEN  SQLS2, CONN, 3, 3
			 			'RESPONSE.WRITE SQLS2
			 			IF NOT RDS.EOF THEN   '到職後本月請事假OR病假有幾天
			 				A3=FIX(CDBL(RDS("HHOUR"))/8)
			 			ELSE
			 				A3=0
			 			END IF
			 			SET RDS=NOTHING
			 			MWORKDAYS = cdbl(A1)-CDBL(A2)-CDBL(A3)+1
			 			comdays = MWORKDAYS
			 			tmpRec(i, j, 59) = MWORKDAYS
			 		ELSE
			 			if tmpRec(i, j, 30) >= ENDdat  then  
			 				calcenddat =ccdt  
			 			else
			 				calcenddat = tmpRec(i, j, 30) 
			 			end  if 	 
			 			'RESPONSE.WRITE "y=" & calcenddat  &"<br>"  
			 			A1= DATEDIFF("D", CDATE(tmpRec(i, j, 5)), CDATE(calcenddat))
			 			'RESPONSE.WRITE "y=" & a1 &"<br>" 
			 			SQLS1="SELECT * FROM YDBMCALE WHERE CONVERT(CHAR(6),DAT,112)='"& YYMM &"' AND  CONVERT(CHAR(10),DAT,111)>= '"& tmpRec(i, j, 5)&"'  AND  CONVERT(CHAR(10),DAT,111)< '"& calcenddat &"' AND STATUS IN ('H2', 'H3' ) "
			 			'RESPONSE.WRITE SQLS1 &"<br>"
			 			Set rDs = Server.CreateObject("ADODB.Recordset")
			 			RDS.OPEN  SQLS1, CONN, 3, 3
			 			IF NOT RDS.EOF THEN   '到職後到月底的假日有幾天
			 				A2=RDS.RECORDCOUNT
			 			ELSE
			 				A2=0
			 			END IF
			 			SET RDS=NOTHING
			 			'RESPONSE.WRITE "y=" & a2 &"<br>" 
			 			SQLS2="SELECT EMPID, SUM(HHOUR) HHOUR  FROM EMPHOLIDAY WHERE EMPID='"& tmpRec(i, j, 1) &"' AND CONVERT(CHAR(6), DATEUP,112)='"& YYMM &"' "&_
			 				  "AND JIATYPE IN ('A','B') AND CONVERT(CHAR(10), DATEUP,111)>='"& tmpRec(i, j, 5) &"' AND CONVERT(CHAR(10), DATEUP,111)<='"& tmpRec(i, j, 30) &"' GROUP BY EMPID "
			 			Set rDs = Server.CreateObject("ADODB.Recordset")
			 			RDS.OPEN  SQLS2, CONN, 3, 3
			 			'RESPONSE.WRITE SQLS2
			 			IF NOT RDS.EOF THEN   '到職後本月請事假OR病假有幾天
			 				A3=FIX(CDBL(RDS("HHOUR"))/8)
			 			ELSE
			 				A3=0
			 			END IF
			 			SET RDS=NOTHING
			 			MWORKDAYS = cdbl(A1)-CDBL(A2)-CDBL(A3)
			 			comdays = MWORKDAYS
			 			tmpRec(i, j, 59) = MWORKDAYS
			 		END IF
		 		ELSE
		 			'MWORKDAYS =  CDBL(days) - CDBL(HHCNT) - CDBL(A4)
		 			'tmpRec(i, j, 59) = tmpRec(i, j, 59)- CDBL(A4) 
		 			tmpRec(i, j, 59) = tmpRec(i, j, 59)
		 		END IF
		 		'response.write tmpRec(i, j, 59)  &"<BR>" 
				
				if cdbl(A4) >= cdbl(MMDAYS) then 
					tmpRec(i, j, 59) = 0 
				end if 


		 		'全勤
				IF CDATE(tmpRec(i, j, 5)) >CDATE(calcdt) THEN
					tmpRec(i, j, 31) = 0
				ELSEif tmpRec(i, j, 59)< ( CDBL(MMDAYS)  ) THEN
					tmpRec(i, j, 31) = 0
				else
					IF CDBL(RS("FORGET")+cdbl(rs("latefor")))>=3 and ( CDBL(RS("FORGET"))+cdbl(rs("latefor"))) < 6 THEN
						if cdbl(rs("jiaa")) + cdbl(rs("jiab")) <=8 then
							tmpRec(i, j, 31)=CDBL(RS("bs_QC"))/2  '全勤
						else
							tmpRec(i, j, 31)= 0
						end if
					ELSEIF 	( CDBL(RS("FORGET"))+cdbl(rs("latefor"))) >=6 THEN
						tmpRec(i, j, 31)= 0
					else
						if  cdbl(rs("jiaa")) + cdbl(rs("jiab"))+cdbl(rs("kzhour"))=0 then
							tmpRec(i, j, 31)=CDBL(RS("bs_QC"))
						elseif  cdbl(rs("jiaa")) + cdbl(rs("jiab"))+cdbl(rs("kzhour"))>=1 and  cdbl(rs("jiaa")) + cdbl(rs("jiab"))+cdbl(rs("kzhour")) <=8 then
							tmpRec(i, j, 31)=CDBL(RS("bs_QC"))/2
						elseif 	cdbl(rs("jiaa")) + cdbl(rs("jiab"))+cdbl(rs("kzhour"))>=9 then
							tmpRec(i, j, 31)= 0
						else
							tmpRec(i, j, 31)=CDBL(RS("bs_QC"))
						end if
					end if
				END IF

		 		 '伙食費
		 		if rs("whsno")="LA" then  '200807起取消伙食費
			 		IF RS("COUNTRY")="VN" THEN
			 			if CDBL(days) <26 then hsdays=CDBL(days) else hsdays=26
				 		IF tmpRec(i, j, 59)< ( CDBL(days) - CDBL(HHCNT) ) THEN
				 			tmpRec(i, j, 35) =0 '1000 * CDBL(comdays)
				 		ELSE
							tmpRec(i, j, 35)=0 '1000*cdbl(hsdays)
						END IF
					ELSE
						tmpRec(i, j, 35)=0
					END IF
				else
					tmpRec(i, j, 35)=0 
				end if 	

		 		'應領薪資合計
		 		'如為本月新進員工薪資OR本月離職工作未滿13天: 總薪資/26 * 工作天數
		 		'舊員工本月離職, 工作天數13天(含)以上 : (BB+CV+PHU / 26 )* 工作天數  + ( NN+KT+MT+TTKH+QC 全薪 )
		 		'如本月工作天數>26天,以實際工作天數計
		 		ALLM=CDBL(tmpRec(i, j, 20))+CDBL(tmpRec(i, j, 22))+CDBL(tmpRec(i, j, 23))+CDBL(tmpRec(i, j, 24))+CDBL(tmpRec(i, j, 25))+CDBL(tmpRec(i, j, 26))+CDBL(tmpRec(i, j, 27))+CDBL(tmpRec(i, j, 31))  '本薪BB+CV+PHU+NN+KT+MT+TTKH+QC
		 		OTRM= CDBL(tmpRec(i, j, 32))+CDBL(tmpRec(i, j, 33))+cdbl(tmpRec(i, j, 58))  '其他收入+上月補款+績效
		 		if trim(tmpRec(i, j, 30))<>"" and tmpRec(i, j, 30) < calcdt then
		 			tmpRec(i, j, 60) = 0
		 		else
			 		if MMDAYS <=26 then   			 		
				 		IF tmpRec(i, j, 59)< ( CDBL(days) - CDBL(HHCNT) ) THEN
				 			IF ( CDBL(tmpRec(i, j, 59)) < 3  and CDATE(tmpRec(i, j, 5))>=CDATE(calcdt) )   then
				 				tmpRec(i, j, 60) = 0
				 			else
					 			IF  CDBL(tmpRec(i, j, 59)) < 13  THEN
					 				tmpRec(i, j, 60) = ( ROUND( CDBL(ALLM) /26 ,0) * CDBL(tmpRec(i, j, 59))) + OTRM
					 				'ALLM = tmpRec(i, j, 60)  
					 				'response.write  tmpRec(i, j, 60) &"<BR>"
					 				'RESPONSE.WRITE "X1" &"<br>"
					 			ELSE
					 				tmpRec(i, j, 60) = ( ROUND(TOTY / 26 ,0) *  CDBL(tmpRec(i, j, 59))) + CDBL(tmpRec(i, j, 24))+CDBL(tmpRec(i, j, 25))+CDBL(tmpRec(i, j, 26))+CDBL(tmpRec(i, j, 27))+CDBL(tmpRec(i, j, 31))+ OTRM
					 				'ALLM = tmpRec(i, j, 60)   
					 				'RESPONSE.WRITE "X2" &"<br>" 
					 				'RESPONSE.WRITE TOTY &"<br>"
					 				'RESPONSE.WRITE ( ROUND(TOTY / 26 ,0) * CDBL(tmpRec(i, j, 59)))   &"<br>" 
					 				'RESPONSE.WRITE  CDBL(tmpRec(i, j, 24)) &"<br>" 
					 				'RESPONSE.WRITE  CDBL(tmpRec(i, j, 25)) &"<br>" 
					 				'RESPONSE.WRITE  CDBL(tmpRec(i, j, 26)) &"<br>" 
					 				'RESPONSE.WRITE  CDBL(tmpRec(i, j, 27)) &"<br>" 
					 				'RESPONSE.WRITE CDBL(tmpRec(i, j, 31)) &"<br>" 
					 				'RESPONSE.WRITE  OTRM&"<br>"  
					 			END IF
					 		end if
				 		ELSE
				 			tmpRec(i, j, 60) = ALLM + OTRM
				 			'RESPONSE.WRITE "X3" &"<br>"
				 		END IF  
				 	else  '如本月工作天數>26天
				 		IF tmpRec(i, j, 59)< ( CDBL(days) - CDBL(HHCNT) ) THEN
				 			IF ( CDBL(tmpRec(i, j, 59)) < 3  and CDATE(tmpRec(i, j, 5))>=CDATE(calcdt) )   then
				 				tmpRec(i, j, 60) = 0
				 			else
								IF  CDBL(tmpRec(i, j, 59)) < 13  THEN
									tmpRec(i, j, 60) = ( ROUND( CDBL(ALLM) /cdbl(MMDAYS)  ,0) * CDBL(tmpRec(i, j, 59))) + OTRM
					 				'ALLM = tmpRec(i, j, 60)  
					 				'response.write  tmpRec(i, j, 60) &"<BR>"
					 				'RESPONSE.WRITE "X1" &"<br>"
					 			ELSE
					 				tmpRec(i, j, 60) = ( ROUND(TOTY / cdbl(MMDAYS) ,0) *  CDBL(tmpRec(i, j, 59))) + CDBL(tmpRec(i, j, 24))+CDBL(tmpRec(i, j, 25))+CDBL(tmpRec(i, j, 26))+CDBL(tmpRec(i, j, 27))+CDBL(tmpRec(i, j, 31))+ OTRM
					 				'ALLM = tmpRec(i, j, 60)   
					 			'RESPONSE.WRITE "X2" &"<br>" 
					 				'RESPONSE.WRITE TOTY &"<br>"
					 				'RESPONSE.WRITE ( ROUND(TOTY / cdbl(MMDAYS) ,0) * CDBL(tmpRec(i, j, 59)))   &"<br>" 
					 				'RESPONSE.WRITE  CDBL(tmpRec(i, j, 24)) &"<br>" 
					 				'RESPONSE.WRITE  CDBL(tmpRec(i, j, 25)) &"<br>" 
					 				'RESPONSE.WRITE  CDBL(tmpRec(i, j, 26)) &"<br>" 
					 				'RESPONSE.WRITE  CDBL(tmpRec(i, j, 27)) &"<br>" 
					 				'RESPONSE.WRITE CDBL(tmpRec(i, j, 31)) &"<br>" 
					 				'RESPONSE.WRITE  OTRM&"<br>"  
					 			END IF
					 		end if
				 		ELSE
				 			tmpRec(i, j, 60) = ALLM + OTRM
				 			'RESPONSE.WRITE "X3" &"<br>"
				 		END IF   
				 	end if	
			 		
			 		'response.write "60-" & tmpRec(i, j, 60)  &"<BR>" 
		 		end if
		 		'應發工資
		 		if tmpRec(i, j, 60) > 0 then
		 			tmpRec(i, j, 39) = CDBL(tmpRec(i, j, 60))-CDBL(tmpRec(i, j, 34))-CDBL(tmpRec(i, j, 35))-CDBL(tmpRec(i, j, 36))-CDBL(tmpRec(i, j, 37))		 			
		 			'response.write "34-" & tmpRec(i, j, 34)  &"<BR>"
		 			'response.write "35-" & tmpRec(i, j, 35)  &"<BR>"
		 			'response.write "36-" & tmpRec(i, j, 36)  &"<BR>"
		 			'response.write "37-" & tmpRec(i, j, 37)  &"<BR>"
		 		else
		 			tmpRec(i, j, 39) = 0
		 		end if
		 		'實領薪資  
		 		'response.write "yy" &   tmpRec(i, j, 59) & ( CDBL(days) - CDBL(HHCNT) ) &"<BR>"
		 		if tmpRec(i, j, 60) > 0 then
		 			IF tmpRec(i, j, 59)< ( CDBL(days) - CDBL(HHCNT) ) THEN
		 				tmpRec(i, j, 47) = CDBL(tmpRec(i, j, 39))+ CDBL(H1_money)+CDBL(H2_money)+CDBL(H3_money)+CDBL(b3_money)-CDBL(kz_money)-CDBL(jiaa_money)-CDBL(jiab_money)
		 				'RESPONSE.WRITE  tmpRec(i, j, 39)  &"<br>"
		 				'RESPONSE.WRITE  kz_money  &"<br>" 
		 				'RESPONSE.WRITE  jiaa_money  &"<br>" 
		 				'RESPONSE.WRITE  jiab_money   &"<br>" 
		 				
		 				'RESPONSE.WRITE "A"&"<br>"
		 			else
		 				tmpRec(i, j, 47) = CDBL(tmpRec(i, j, 39))+ CDBL(H1_money)+CDBL(H2_money)+CDBL(H3_money)+CDBL(b3_money)-CDBL(kz_money)-CDBL(jiaa_money)-CDBL(jiab_money)
		 				'RESPONSE.WRITE "B" 
		 			end if
		 		else
		 			tmpRec(i, j, 47) = 0
		 		end if

				'離職補助金
				if tmpRec(i, j, 4)="VN" and rs("nindat")<"2009/01/01"  then
					inMonth = 0
					jishu = 0
					if  tmpRec(i, j, 30)="" then
						inMonth  = datediff("m", CDATE(tmpRec(i, j, 5)) , CDATE(calcdt) )
					else
					 	inMonth  = datediff("m", CDATE(tmpRec(i, j, 5)) , CDATE(tmpRec(i, j, 30)) )
					end if
					if inMonth >= 12 then
					 	if inMonth mod 12 = 0 then
					 		jishu=round(fix(inMonth/12 )*0.5,2) 
					 		'response.write  "1" &"<BR>"
					 	elseif inMonth mod 12 > 6 then
					 		jishu=round(fix(inMonth/12 )*0.5+0.25 ,2)
					 		'response.write  "2"&"<BR>"
					 	else
					 		jishu=round(fix(inMonth/12 )*0.5 ,2)
					 		'response.write  "3"&"<BR>"
					 	end if
					else
					 	jishu=0
					end if
					tmpRec(i, j, 61) = jishu   '基數
					IF jishu > 0 THEN
					 	tmpRec(i, j, 62) = ROUND( ( CDBL(tmpRec(i, j, 20)) + CDBL(tmpRec(i, j, 22))+CDBL(tmpRec(i, j, 23))+CDBL(tmpRec(i, j, 24))+CDBL(tmpRec(i, j, 25))+CDBL(tmpRec(i, j, 26)) )  * CDBL(jishu) ,0)
					ELSE
						tmpRec(i, j, 62) = 0
					END IF
				else
					tmpRec(i, j, 62) = 0
				end if 
				
				 'response.write   "年直:" & jishu  &"<BR>"
				 
				 '檢核總薪資
				 tmpRec(i, j, 63) = ( cdbl(ALLM) + cdbl(OTRM)  + cdbl(tmpRec(i, j, 49)) ) -  cdbl(tmpRec(i, j, 50) ) -CDBL(tmpRec(i, j, 34))-CDBL(tmpRec(i, j, 35))-CDBL(tmpRec(i, j, 36))-CDBL(tmpRec(i, j, 37))
				 tmpRec(i, j, 64) =   ( cdbl(ALLM) + cdbl(OTRM) + cdbl(tmpRec(i, j, 49))  )
				 tmpRec(i, j, 65) =  cdbl(tmpRec(i, j, 63)) - cdbl(tmpRec(i, j, 47))
				 				 
				 if len(trim(rs("studyjob")))>0 then 				 
				 	tmpRec(i, j, 66) =  "Blue"
				 else	
				 	tmpRec(i, j, 66) =  "black"
				 end if 	 
				 
				 if trim(rs("outdate"))<>""   then 
				 	if cdbl(tmpRec(i, j, 62))>"0" and len(trim(rs("studyjob")))=0  then 
				 		tmpRec(i, j, 66) = "#ff0099" 
				 	elseif cdbl(tmpRec(i, j, 62))>"0"  and len(trim(rs("studyjob")))>0	then 
				 		tmpRec(i, j, 66) = "#3333ff" 
				 	end if 
				 else 
				 	tmpRec(i, j, 66) =  "black"
				 end if 		
				 	
				 'response.write  tmpRec(i, j, 63) &"<BR>"
				 'response.write  tmpRec(i, j, 64) &"<BR>"
				'response.write  ALLM &"<BR>"
				'response.write  OTRM &"<BR>"				 
				'response.write  tmpRec(i, j, 66) &"<BR>"  
				if rs("NoverJ") > "0" then
					tmpRec(i, j, 67) = cdbl(RS("njh1"))- cdbl(rs("NoverJ"))
				else 
					tmpRec(i, j, 67) =  (rs("njh1"))
				end if	 
				tmpRec(i, j, 68)=(cdbl(tmpRec(i, j, 20))+cdbl(tmpRec(i, j, 22))+cdbl(tmpRec(i, j, 23))+cdbl(tmpRec(i, j, 24))+cdbl(tmpRec(i, j, 25))+cdbl(tmpRec(i, j, 26)))/204 
				
				'個人所得稅計算
				F_TAX = 0 
				real_TOTAMT =  cdbl(tmpRec(i, j, 47))  ' 實領金額  
				totB = ( cdbl(rs("taxSet"))+cdbl(rs("notax_amt")) ) 
				if real_TOTAMT > ( cdbl(rs("taxSet"))+cdbl(rs("notax_amt")) ) then 
					if left(yymm,4)>"2008" then 
						sql2="exec sp_calctax '"& real_TOTAMT &"','"& cdbl(totB) &"' "
						'response.write sql2 
						'response.end 
						set ors=conn.execute(sql2) 
						f_tax = ors("tax")
					else
						sql2="exec sp_calctax_2008 '"& real_TOTAMT &"' "
						set ors=conn.execute(sql2) 
						f_tax = ors("tax")
					end if  				
					set ors=nothing 
				else
					f_tax = 0 
				end if 	
				
				tmpRec(i, j, 69) = ROUND(F_tax,0)  				
				tmpRec(i, j, 47) = cdbl(tmpRec(i, j, 47)) -  cdbl(tmpRec(i, j, 69))
				tmpRec(i, j, 70) = rs("salarymemo")
				tmpRec(i, j, 71) = rs("studyJob")    			
				if trim(rs("outdate"))="" then 
					tmpRec(i, j, 72) = round(datediff("d",cdate(rs("indat")),date())/30,1)
				else
					tmpRec(i, j, 72) = round(datediff("d",cdate(rs("indat")),cdate(rs("outdate")))/30,1)
				end if 	 				
				'年度特休 (前ㄧ年  4/1 計算至當年 4/1 ) 
				if trim(rs("outdate"))="" then 
					'tx_enddat = ENDdat 
					tx_enddat = ccdt
				else
					if right(rs("outdate"),5)>= "03/31" and  year(rs("indat")) < year(date())  then 
						'tx_enddat = cstr(left(yymm,4))&"/04/01"
						tx_enddat = rs("outdate") 
					else
						tx_enddat = rs("outdate") 
					end if 	
				end if 
				'response.write "xxx=" &  tx_enddat 
				'response.end  
 
				if  right(yymm,2)<="03" and yymm<=nowmonth then 
					if cdate( rs("calcTxdat") ) <= cdate(cstr(left(yymm,4)-1)&"/03/31") then 
						cc_indat = cdate(cstr(left(yymm,4)-1)&"/04/01") 
					else
						cc_indat = rs("calcTxdat")
					end if	
 				else
 					if rs("nindat")>left(yymm,4)&"/03/31"  then 
 						cc_indat = rs("calcTxdat")
 					else 					
 						cc_indat = cstr(left(yymm,4))&"/04/01"				
 					end if	
 				end if 
				
				if cdbl(datediff("m",cdate(cc_indat),cdate(tx_enddat) )) < 0 then 
					tmpRec(i, j, 73) = 0
				else					
					tmpRec(i, j, 73) = cdbl(datediff("m",cdate(cc_indat),cdate(tx_enddat) ))*8
				end if	
				tmpRec(i, j, 74) = CDBL(RS("jiaE"))
			 		
				tmpRec(i, j, 75)  = TOTY + cdbl(tmpRec(i, j, 24))+cdbl(tmpRec(i, j, 25))+cdbl(tmpRec(i, j, 26))+cdbl(tmpRec(i, j, 27))
				
				tmpRec(i, j, 76) = rs("person_qty")  '扶養人數
				tmpRec(i, j, 77) = rs("notax_amt")  '免稅額
				tmpRec(i, j, 78) = rs("taxSet")  '扣稅基本額 
				
		 
			rs.MoveNext
		else
			exit for
		end if
	 next

	 if rs.EOF then
		rs.Close
		Set rs = nothing
		exit for
	 end if
	next
	Session("empfilesalary") = tmpRec
else
	ccdt = request("ccdt")
	calcdat = request("calcdat")
	enddat = request("enddat")
	YYMM = request("yymm")
	MMDAYS = request("MMDAYS")
	TotalPage = cint(request("TotalPage"))
	'StoreToSession()
	CurrentPage = cint(request("CurrentPage"))
	RecordInDB  = REQUEST("RecordInDB")
	tmpRec = Session("empfilesalary")

	Select case request("send")
	     Case "FIRST"
		      CurrentPage = 1
	     Case "BACK"
		      if cint(CurrentPage) <> 1 then
			     CurrentPage = CurrentPage - 1
		      end if
	     Case "NEXT"
		      if cint(CurrentPage) < cint(TotalPage) then
			     CurrentPage = CurrentPage + 1
			  else
			  	 CurrentPage = TotalPage
		      end if
	     Case "END"
		      CurrentPage = TotalPage
	     Case Else
		      CurrentPage = 1
	end Select
end if

'response.end

FUNCTION FDT(D)
	Response.Write YEAR(D)&"/"&RIGHT("00"&MONTH(D),2)&"/"&RIGHT("00"&DAY(D),2)

END FUNCTION
%>

<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="refresh">
<link rel="stylesheet" href="../Include/style.css" type="text/css">
<link rel="stylesheet" href="../Include/style2.css" type="text/css">
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>

'-----------------enter to next field
function enterto()
	if window.event.keyCode = 13 then window.event.keyCode =9
end function

function f()
	'<%=self%>.PHU(0).focus()
	'<%=self%>.PHU(0).SELECT()
end function

function chgdata()
	<%=self%>.action="empfile.salary.asp?totalpage=0"
	<%=self%>.submit
end function

</SCRIPT>
</head>
<body   topmargin="0" leftmargin="0"  marginwidth="0" marginheight="0"  onkeydown="enterto()" onload="f()" bgproperties="fixed"  >
<form name="<%=self%>" method="post" action="<%=SELF%>.salary.asp">
<INPUT TYPE=HIDDEN NAME="UID" VALUE=<%=SESSION("NETUSER")%>>
<INPUT TYPE=hidden NAME=TotalPage VALUE="<%=TotalPage%>">
<INPUT TYPE=hidden NAME=CurrentPage VALUE="<%=CurrentPage%>">
<INPUT TYPE=hidden NAME=RecordInDB VALUE="<%=RecordInDB%>">
<INPUT TYPE=hidden NAME=PageRec VALUE="<%=PageRec%>">
<INPUT TYPE=hidden NAME=gTotalPage VALUE="<%=gTotalPage%>">
<INPUT TYPE='hidden' NAME=YYMM VALUE="<%=YYMM%>">
<INPUT TYPE='hidden' NAME=calcdat VALUE="<%=calcdt%>">
<INPUT TYPE='hidden' NAME=enddat VALUE="<%=year(enddat)&"/"&right("00"&month(enddat),2)&"/"&right("00"&day(enddat),2)%>">
<INPUT TYPE='hidden' NAME=ccdt VALUE="<%=ccdt%>">
<INPUT TYPE='hidden' NAME=MMDAYS VALUE="<%=MMDAYS%>" title="本月工作天數">

<table width="600" border="0" cellspacing="0" cellpadding="0">
	<tr>
	<TD width=430>
	<img border="0" src="../image/icon.gif" align="absmiddle">
	人事薪資系統( 員工薪資管理 )　
	計薪年月：<%=YYMM%>	</TD>
	</tr>
</table>
<hr size=0	style='border: 1px dotted #999999;' align=left width=500>

 
<TABLE  CLASS="txt8" BORDER="0" cellspacing="1" cellpadding="1" BGCOLOR="LightGrey"   >
	<TR HEIGHT=25 BGCOLOR="#e4e4e4"   >
 		<TD ROWSPAN=2 >項次</TD>
 		<TD align=center>工號<br>So the</TD>
 		<TD COLSPAN=3  >員工姓名(中,英,越)&nbsp;Ho Ten</TD>
 		<TD  align=center>職等<br>Chuc vu</TD>
 		<td align=center>時薪</td>
 		<td align=center>到職日<br>NVX</td>
 		<td align=center>離職日期<br>NTV</td>
 		<TD align=center>工作天數<br>So Ngay<br>Lam viec</TD>
 		<TD align=center>上月補款<br>T.B.T.R</TD>
 		<TD align=center>總加班費<br>Phi Tang ca</TD> 		
 		<TD align=center>離職補助<br>Tro cap<br>Thoi viec</TD> 
		<TD align=center>人數<br>So nguoi</TD> 
 		<td align=center>(-)扣時假</td>
 		<TD align=center>(-)不足月</TD>
 		<TD align=center>(-)所得稅</TD> 						
 		<TD ALIGN=CENTER >實領工資</TD>
 		<TD bgcolor="#ffcc99" align=center><font color=blue>年假</font></TD>
 		<TD COLSPAN=4 ALIGN=CENTER bgcolor="#ccff99">加班(H)</TD>
 		<TD bgcolor="#ffcc99"></TD>
 		<TD bgcolor="#ffcc99"></TD> 		
 		<TD COLSPAN=2 ALIGN=CENTER bgcolor="#ffcccc">請假(H)</TD>
 	</TR>
 	<tr BGCOLOR="#e4e4e4"  HEIGHT=25 >
 		<TD align=center>基薪(BB)</TD>
 		<TD align=center>職加(CV)</TD>
 		<TD align=center>獎金(Y)</TD>
 		<td align=center>語言(NN)</td>
 		<td align=center>技術(KT)</td>
 		<td align=center>環境(MT)</td>
 		<td align=center>其加(TTKH)</td>
 		<td align=center>薪資合計</td>
 		<td align=center>全勤獎金</td>
 		<td align=center><font color=red>績效獎金</font></td>
 		<td align=center>其他收入</td>
 		<TD align=center>應領薪資</TD>
		<TD align=center>家境免稅<br>tien mien thue</TD> 
 		<td align=center>(-)其他</td>
 		<td align=center>(-)工團費</td>
 		<td align=center>(-)保險費</td>
 		<td align=center>(-)伙食費</td>
 		
 		<TD ALIGN=CENTER bgcolor="#ffcc99"><font color=Brown>尚有年假</font></TD>
 		<TD ALIGN=CENTER bgcolor="#ccff99">平日</TD>
 		<TD ALIGN=CENTER bgcolor="#ccff99">休息</TD>
 		<TD ALIGN=CENTER bgcolor="#ccff99">假日</TD>
 		<TD ALIGN=CENTER bgcolor="#ccff99">夜班</TD>
 		<TD ALIGN=CENTER bgcolor="#ffcc99"><font color=Brown>曠職</font></TD>
 		<TD ALIGN=CENTER bgcolor="#ffcc99"><font color=Brown>忘遲</font></TD> 		
 		<TD ALIGN=CENTER bgcolor="#ffcccc" >事假</TD>
 		<TD ALIGN=CENTER bgcolor="#ffcccc" >病假</TD>
 	</tr>
	<%for CurrentRow = 1 to PageRec
		IF CurrentRow MOD 2 = 0 THEN
			WKCOLOR="LavenderBlush"
			'wkcolor="#ffffff"
		ELSE
			WKCOLOR="#DFEFFF"
			'wkcolor="#ffffff"
		END IF
		'if tmpRec(CurrentPage, CurrentRow, 1) <> "" then
	%>
	<TR BGCOLOR=<%=WKCOLOR%> >
		<TD ROWSPAN=2 ALIGN=CENTER >
		<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then %><%=(CURRENTROW)+((CURRENTPAGE-1)*10)%><%END IF %>
		</TD>
 		<TD align=center >
 			<a href='vbscript:editmemo(<%=CURRENTROW-1%>)'>
 				<font color="<%=tmpRec(CurrentPage, CurrentRow, 66)%>"><u><%=tmpRec(CurrentPage, CurrentRow, 1)%></u></font>
 			</a>
 			<input type=hidden name=empid value="<%=tmpRec(CurrentPage, CurrentRow, 1)%>"  >
 			<input type=hidden name="empautoid" value="<%=tmpRec(CurrentPage, CurrentRow, 17)%>">
 		</TD>
 		<TD COLSPAN=3 >
 			<a href='vbscript:oktest(<%=tmpRec(CurrentPage, CurrentRow, 17)%>)'> 				
 				<font class=txt8 color="<%=tmpRec(CurrentPage, CurrentRow, 66)%>">
 					<%=left(tmpRec(CurrentPage, CurrentRow, 2)&tmpRec(CurrentPage, CurrentRow, 3),26)%>
 				</font>
 			</a>
 		</TD>
 		<td  align=center><!--職等-->
 			<font color="<%=tmpRec(CurrentPage, CurrentRow, 66)%>"><%=left(tmpRec(CurrentPage, CurrentRow, 15),5)%></font>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
	 			<input type=hidden name=F1_JOB  class="txt8"   value="<%=trim(tmpRec(CurrentPage, CurrentRow, 6))%>"> 				
			<%else%>
				<input type=hidden name=F1_JOB >
			<%end if %>  		 			
 		</td>
 		<TD >
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
 				<%if session("netuser")="LSARY" then %> 
 					<INPUT NAME=HHMOENYjd  VALUE="<%=round(tmpRec(CurrentPage, CurrentRow, 38),0)%>" CLASS='INPUTBOX8' SIZE=7 STYLE="TEXT-ALIGN:RIGHT;COLOR:#3366cc;BACKGROUND-COLOR:LIGHTYELLOW" >
 					<INPUT type=hidden NAME=HHMOENY  VALUE="<%=tmpRec(CurrentPage, CurrentRow, 38)%>" CLASS='INPUTBOX8' SIZE=7 STYLE="TEXT-ALIGN:RIGHT;COLOR:#3366cc;BACKGROUND-COLOR:LIGHTYELLOW" >
 				<%else%>
 					<INPUT NAME=HHMOENY  VALUE="<%=tmpRec(CurrentPage, CurrentRow, 38)%>" CLASS='INPUTBOX8' SIZE=7 STYLE="TEXT-ALIGN:RIGHT;COLOR:#3366cc;BACKGROUND-COLOR:LIGHTYELLOW" >
 				<%end if%>	
 			<%ELSE%>
 				<INPUT NAME=HHMOENY TYPE=HIDDEN>
 			<%END IF %>
 		</TD>
 		<TD  ALIGN=CENTER ><FONT CLASS=TXT8><%=RIGHT(tmpRec(CurrentPage, CurrentRow, 5),8)%></FONT></TD>
 		<TD  ALIGN=CENTER ><FONT CLASS=TXT8 color=red><b><%=RIGHT(tmpRec(CurrentPage, CurrentRow, 30),8)%></b></FONT></TD>
 		<TD >
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
	 			<INPUT NAME=WORKDAYS CLASS='INPUTBOX8' READONLY  SIZE=7 VALUE="<%=tmpRec(CurrentPage, CurrentRow, 59)%>(<%=tmpRec(CurrentPage, CurrentRow, 72)%>)" STYLE="TEXT-ALIGN:RIGHT;BACKGROUND-COLOR:Gainsboro"  title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 工作天數">
	 			
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=WORKDAYS >
	 		<%END IF%>
 		</TD>
 		<TD>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
	 			<INPUT NAME=TBTR CLASS='INPUTBOX8' SIZE=7 VALUE="<%=tmpRec(CurrentPage, CurrentRow, 33)%>" STYLE="TEXT-ALIGN:RIGHT;BACKGROUND-COLOR:LIGHTYELLOW"  READONLY  title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 上月補款" >
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=TBTR >
	 		<%END IF%>
 		</TD> 		
 		<TD >
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
	 			<INPUT NAME=TOTJB CLASS='INPUTBOX8' SIZE=7 VALUE="<%=tmpRec(CurrentPage, CurrentRow, 49)%>" STYLE="TEXT-ALIGN:RIGHT;BACKGROUND-COLOR:LIGHTYELLOW"  READONLY  title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 總加班費">
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=TOTJB >
	 		<%END IF%>
 		</TD>	 
 		<TD  ALIGN=RIGHT><!--離職補助-->
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %> 				
	 			<INPUT NAME=LZBZJ CLASS='INPUTBOX8' SIZE=8 VALUE="<%=formatnumber(tmpRec(CurrentPage, CurrentRow, 62),0)%>" STYLE="TEXT-ALIGN:RIGHT;BACKGROUND-COLOR:LIGHTYELLOW;color:MediumBlue" READONLY   title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 離職補助金">
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=LZBZJ >	 			
	 		<%END IF%> 
		</TD> 
		<TD  ALIGN=RIGHT><!--扶養人數-->
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %> 				
	 			<INPUT NAME=person_Qty CLASS='INPUTBOX8' SIZE=8 VALUE="<%=formatnumber(tmpRec(CurrentPage, CurrentRow, 76),0)%>" STYLE="TEXT-ALIGN:RIGHT;BACKGROUND-COLOR:LIGHTYELLOW;color:MediumBlue" READONLY   title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 離職補助金">
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=person_Qty >	 			
	 		<%END IF%> 
		</TD> 
 		<TD ><!--扣時假-->
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
	 			<INPUT NAME=TOTKJ CLASS='INPUTBOX8' SIZE=7 VALUE="<%=formatnumber(tmpRec(CurrentPage, CurrentRow, 50),0)%>" STYLE="TEXT-ALIGN:RIGHT;BACKGROUND-COLOR:LIGHTYELLOW" READONLY   title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 扣時假">
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=TOTKJ >
	 		<%END IF%>
 		</TD>
 		<TD>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %> 
	 			<INPUT NAME=BZKM CLASS='INPUTBOX8' SIZE=7 VALUE="<%=formatnumber(tmpRec(CurrentPage, CurrentRow, 65),0)%>" STYLE="TEXT-ALIGN:RIGHT;BACKGROUND-COLOR:LIGHTYELLOW" READONLY title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 不足月扣款"> 
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=BZKM   >
	 		<%END IF%>
 		</TD>
 		<TD >
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
 				<INPUT NAME=KTAXM CLASS='INPUTBOX8' SIZE=7 VALUE="<%=tmpRec(CurrentPage, CurrentRow, 69)%>" STYLE="TEXT-ALIGN:RIGHT;BACKGROUND-COLOR:LIGHTYELLOW;color:MediumBlue" READONLY   title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 所得稅">	 			
	 		<%ELSE%>	 			
	 			<INPUT TYPE=HIDDEN NAME=KTAXM >
	 		<%END IF%>
 		</TD> 
 		<TD>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
	 			<INPUT NAME=RELTOTMONEY CLASS='INPUTBOX8' VALUE="<%=formatnumber( tmpRec(CurrentPage, CurrentRow, 47),0)%>" SIZE=8 STYLE="TEXT-ALIGN:RIGHT;BACKGROUND-COLOR:LIGHTYELLOW;;color:#cc0000" READONLY title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 實領工資">
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=RELTOTMONEY  >
	 		<%END IF%>
 		</TD> 		
		<TD class=txt8 align=right><font color=blue><%=tmpRec(CurrentPage, CurrentRow, 73)%></font></TD> 		
 		<TD COLSPAN=4 align=center>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
 				<FONT CLASS=TXT8><u><div style="cursor:hand" onclick="view1(<%=currentrow-1%>)"><%=tmpRec(CurrentPage, CurrentRow, 1)%> 出勤紀錄</div></u></font>
 			<%END IF %>
 		</TD>
 		<TD ></TD>
 		<TD ></TD> 		
 		<TD COLSPAN=2 align=center>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
 				<FONT CLASS=TXT8><u><div style="cursor:hand" onclick="view2(<%=currentrow-1%>)">請假紀錄</div></u></font>
 			<%END IF %>  			
 		</TD>
	</TR>
	<TR BGCOLOR=<%=WKCOLOR%> ><!------ line 2 ------------------------->
 		<TD ALIGN=RIGHT >
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
		 		<input  type=hidden  value="<%=trim(tmpRec(CurrentPage, CurrentRow, 19))%>"  name=BBCODE  class="txt8"  > 
			<%else%>
				<input type=hidden name=BBCODE >
			<%end if %> 
			
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
	 			<INPUT NAME=BB CLASS='INPUTBOX8' SIZE=7 VALUE="<%=tmpRec(CurrentPage, CurrentRow, 20)%>" STYLE="TEXT-ALIGN:RIGHT;BACKGROUND-COLOR:LIGHTYELLOW"  READONLY title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 資本薪資">
	 		<%else%>
				<input type=hidden name=BB >
			<%end if %>						
			
 		</TD>
 		<TD ALIGN=RIGHT><!--職加--> 
 		 	<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
 		 		<INPUT NAME=CV CLASS='INPUTBOX8' SIZE=7 VALUE="<%=tmpRec(CurrentPage, CurrentRow, 22)%>" STYLE="TEXT-ALIGN:right;BACKGROUND-COLOR:LIGHTYELLOW"  READONLY  title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 職務加給" >
 		 		<input type=hidden name=CVCODE VALUE="<%=tmpRec(CurrentPage, CurrentRow, 21)%>" SIZE=3>
 		 	<%else%>
				<input type=hidden name=CV >
				<input type=hidden name=CVCODE >
			<%end if %>	 
 		</TD> 
 		<TD  ALIGN=RIGHT>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
	 			<INPUT NAME=PHU CLASS='readonly8' readonly  SIZE=7 VALUE="<%=tmpRec(CurrentPage, CurrentRow, 23)%>" STYLE="TEXT-ALIGN:RIGHT"  title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 補助獎金(Y)" >
	 		<%ELSE%>
				<INPUT TYPE=HIDDEN NAME=PHU	>
			<%END IF%>
 		</TD>
 		<TD  ALIGN=RIGHT>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
 				<INPUT NAME=NN CLASS='readonly8' readonly SIZE=7 VALUE="<%=tmpRec(CurrentPage, CurrentRow, 24)%>" STYLE="TEXT-ALIGN:RIGHT"   title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 語言加給" >
 			<%ELSE%>
				<INPUT TYPE=HIDDEN NAME=NN >
			<%END IF%>
 		</TD>
 		<TD  ALIGN=RIGHT>
	 		<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
	 			<INPUT NAME=KT CLASS='readonly8' readonly SIZE=7 VALUE="<%=tmpRec(CurrentPage, CurrentRow, 25)%>" STYLE="TEXT-ALIGN:RIGHT"   title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 技術加給" >
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=KT >
	 		<%END IF%>
 		</TD>
 		<TD  ALIGN=RIGHT>
			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
	 			<INPUT NAME=MT CLASS='readonly8' readonly SIZE=7 VALUE="<%=tmpRec(CurrentPage, CurrentRow, 26)%>" STYLE="TEXT-ALIGN:RIGHT"   title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 環境加給">
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=MT >
	 		<%END IF%>
 		</TD>
 		<TD  ALIGN=RIGHT><!--其加-->
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
	 			<INPUT NAME=TTKH CLASS='readonly8' readonly SIZE=7 VALUE="<%=tmpRec(CurrentPage, CurrentRow, 27)%>" STYLE="TEXT-ALIGN:RIGHT"  title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 其他加給">
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=TTKH >
	 		<%END IF%>
 		</TD>
 		<TD ALIGN=RIGHT >
			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
				<INPUT NAME=totbsalary CLASS='INPUTBOX8' SIZE=7 VALUE="<%=formatnumber(tmpRec(CurrentPage, CurrentRow, 75),0)%>" STYLE="TEXT-ALIGN:RIGHT;BACKGROUND-COLOR:LIGHTYELLOW" READONLY  title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 薪資合計">
			<%else%>	
				<INPUT TYPE=HIDDEN NAME=totbsalary >
			<%end if %>
 		</TD>	
 		<TD>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
	 			<INPUT NAME=QC CLASS='INPUTBOX8' SIZE=7 VALUE="<%=formatnumber(tmpRec(CurrentPage, CurrentRow, 31),0)%>" STYLE="TEXT-ALIGN:RIGHT;BACKGROUND-COLOR:LIGHTYELLOW" READONLY  title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 全勤">
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=QC >
	 		<%END IF%>
 		</TD> 
 		<TD>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
	 			<INPUT NAME=JX SIZE=7 VALUE="<%=tmpRec(CurrentPage, CurrentRow, 58)%>" STYLE="TEXT-ALIGN:RIGHT" onblur="DATACHG(<%=CURRENTROW-1%>)"  title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 績效獎金" CLASS='inpt8red'  >
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=JX >
	 		<%END IF%>
 		</TD>  		 
 		<TD>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
	 			<INPUT NAME=TNKH SIZE=7 VALUE="<%=tmpRec(CurrentPage, CurrentRow, 32)%>" STYLE="TEXT-ALIGN:RIGHT" onblur="DATACHG(<%=CURRENTROW-1%>)" title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 其他收入" CLASS='inpt8blue' >
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=TNKH >
	 		<%END IF%>
 		</TD>
 		<TD>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
	 			<INPUT NAME=TOTM CLASS='INPUTBOX8' SIZE=8 VALUE="<%=formatnumber( tmpRec(CurrentPage, CurrentRow, 64),0)%>" STYLE="TEXT-ALIGN:RIGHT;BACKGROUND-COLOR:LIGHTYELLOW;color:#cc0000"  READONLY  title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 應領薪資">
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=TOTM >
	 		<%END IF%>
 		</TD>
		<TD  ALIGN=RIGHT><!--免稅額-->
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %> 				
	 			<INPUT NAME=Notax_amt CLASS='INPUTBOX8' SIZE=8 VALUE="<%=formatnumber(tmpRec(CurrentPage, CurrentRow, 77),0)%>" STYLE="TEXT-ALIGN:RIGHT;BACKGROUND-COLOR:LIGHTYELLOW;color:MediumBlue" READONLY   title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 離職補助金">
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=Notax_amt >	 			
	 		<%END IF%> 
		</TD>		
 		<TD>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
	 			<INPUT NAME=QITA CLASS='INPUTBOX8' SIZE=7 VALUE="<%=tmpRec(CurrentPage, CurrentRow, 37)%>" STYLE="TEXT-ALIGN:RIGHT" onblur="DATACHG(<%=CURRENTROW-1%>)" title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 扣除其他" >
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=QITA >
	 		<%END IF%>
 		</TD>
 		<TD>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
	 			<INPUT NAME=GT CLASS='readonly8' SIZE=7 VALUE="<%=tmpRec(CurrentPage, CurrentRow, 36)%>" STYLE="TEXT-ALIGN:RIGHT"   title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 工團費"  readonly >
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=GT >
	 		<%END IF%>
 		</TD>
 		<TD>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
	 			<INPUT NAME=BH CLASS='readonly8' SIZE=7 VALUE="<%=tmpRec(CurrentPage, CurrentRow, 34)%>" STYLE="TEXT-ALIGN:RIGHT;COLOR:#3366cc"  title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 保險費"  readonly  >
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=BH >
	 		<%END IF%>
 		</TD>
 		<TD>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
	 			<INPUT NAME=HS CLASS='INPUTBOX8' SIZE=8 VALUE="<%=tmpRec(CurrentPage, CurrentRow, 35)%>" STYLE="TEXT-ALIGN:RIGHT" onblur="DATACHG(<%=CURRENTROW-1%>)" title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 伙食費" >
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=HS >
	 		<%END IF%>
 		</TD> 
 		<TD>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
	 			<INPUT NAME=YTX CLASS='INPUTBOX8' VALUE="<%=cdbl(tmpRec(CurrentPage, CurrentRow, 73))-cdbl(tmpRec(CurrentPage, CurrentRow, 74))%>"  SIZE=3 STYLE="TEXT-ALIGN:RIGHT;BACKGROUND-COLOR:LIGHTYELLOW;color:red" READONLY >
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=YTX >
	 		<%END IF%>
 		</TD> 	 
 		<TD>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
 				<%if session("netuser")="LSARY" then %>
 					<INPUT NAME=H1jd CLASS='INPUTBOX8' VALUE="<%=tmpRec(CurrentPage, CurrentRow, 67)%>"  SIZE=3 STYLE="TEXT-ALIGN:RIGHT;BACKGROUND-COLOR:LIGHTYELLOW;color:MediumBlue" READONLY >
 					<INPUT type=hidden NAME=H1 CLASS='INPUTBOX8' VALUE="<%=tmpRec(CurrentPage, CurrentRow, 40)%>"  SIZE=3 STYLE="TEXT-ALIGN:RIGHT;BACKGROUND-COLOR:LIGHTYELLOW;color:MediumBlue" READONLY >
 				<%else%>
	 				<INPUT NAME=H1 CLASS='INPUTBOX8' VALUE="<%=tmpRec(CurrentPage, CurrentRow, 40)%>"  SIZE=3 STYLE="TEXT-ALIGN:RIGHT;BACKGROUND-COLOR:LIGHTYELLOW;color:MediumBlue" READONLY >
	 			<%end if%>
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=H1 >
	 		<%END IF%>	 		
 		</TD>
 		<TD>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
	 			<INPUT NAME=H2 CLASS='INPUTBOX8' VALUE="<%=tmpRec(CurrentPage, CurrentRow, 41)%>"  SIZE=3 STYLE="TEXT-ALIGN:RIGHT;BACKGROUND-COLOR:LIGHTYELLOW;color:MediumBlue" READONLY >
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=H2 >
	 		<%END IF%>
 		</TD>
 		<TD>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
	 			<INPUT NAME=H3 CLASS='INPUTBOX8' VALUE="<%=tmpRec(CurrentPage, CurrentRow, 42)%>"  SIZE=3 STYLE="TEXT-ALIGN:RIGHT;BACKGROUND-COLOR:LIGHTYELLOW;color:MediumBlue" READONLY >
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=H3 >
	 		<%END IF%>
 		</TD>
 		<TD>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
	 			<INPUT NAME=B3 CLASS='INPUTBOX8'  VALUE="<%=tmpRec(CurrentPage, CurrentRow, 43)%>"  SIZE=3  STYLE="TEXT-ALIGN:RIGHT;BACKGROUND-COLOR:LIGHTYELLOW;color:MediumBlue" READONLY >
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=B3 >
	 		<%END IF%>
 		</TD>
 		<TD>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
	 			<INPUT NAME=KZHOUR CLASS='INPUTBOX8' VALUE="<%=tmpRec(CurrentPage, CurrentRow, 44)%>"  SIZE=3  STYLE="TEXT-ALIGN:RIGHT;BACKGROUND-COLOR:LIGHTYELLOW;color:Brown" READONLY >
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=KZHOUR >
	 		<%END IF%>
 		</TD>
 		<TD>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
	 			<INPUT NAME=Forget CLASS='INPUTBOX8' VALUE="<%=tmpRec(CurrentPage, CurrentRow, 48)%>"  SIZE=3  STYLE="TEXT-ALIGN:RIGHT;BACKGROUND-COLOR:LIGHTYELLOW;color:Brown" READONLY >
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=Forget >
	 		<%END IF%>
 		</TD> 		
 		<TD>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
	 			<INPUT NAME=JIAA CLASS='INPUTBOX8' VALUE="<%=tmpRec(CurrentPage, CurrentRow, 45)%>"  SIZE=3 STYLE="TEXT-ALIGN:RIGHT;BACKGROUND-COLOR:LIGHTYELLOW;color:ForestGreen" READONLY >
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=JIAA >
	 		<%END IF%>
 		</TD>
 		<TD>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
	 			<INPUT NAME=JIAB CLASS='INPUTBOX8' VALUE="<%=tmpRec(CurrentPage, CurrentRow, 46)%>"  SIZE=3 STYLE="TEXT-ALIGN:RIGHT;BACKGROUND-COLOR:LIGHTYELLOW;color:ForestGreen READONLY >
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=JIAB >	 			
	 		<%END IF%> 
	 		
 		</TD> 		
	</TR>
	<INPUT TYPE=hidden VALUE="0" NAME=XIANM>
	<INPUT TYPE=hidden VALUE="0" NAME=ZHUANM >
	<%next%>
</TABLE>
<input type=hidden name=empid>
<input type=hidden name=BBCODE>
<input type=hidden name=BB>
<input type=hidden name=F1_JOB>
<input type=hidden name=CV>
<input type=hidden name=CVCODE>
<input type=hidden name=PHU>
<input type=hidden name=NN>
<input type=hidden name=KT>
<input type=hidden name=MT>
<input type=hidden name=TTKH>
<INPUT TYPE=HIDDEN NAME=ZHUANM  >
<INPUT TYPE=HIDDEN NAME=XIANM  >
<INPUT TYPE=HIDDEN NAME=notax_amt  ><!--免稅額-->


<TABLE border=0 width=650 class=font9 >
<tr>
    <td align="CENTER" height=40 WIDTH=60%>
	<% If CurrentPage > 1 Then %>
		<input type="submit" name="send" value="FIRST" class=button>
		<input type="submit" name="send" value="BACK" class=button>
	<% Else %>
		<input type="submit" name="send" value="FIRST" disabled class=button>
		<input type="submit" name="send" value="BACK" disabled class=button>
	<% End If %>
	<% If cint(CurrentPage) < cint(TotalPage) Then %>
		<input type="submit" name="send" value="NEXT" class=button>
		<input type="submit" name="send" value="END" class=button>
	<% Else %>
		<input type="submit" name="send" value="NEXT" disabled class=button>
		<input type="submit" name="send" value="END" disabled class=button>
	<% End If %>
	<FONT CLASS=TXT8>&nbsp;&nbsp;PAGE:<%=CURRENTPAGE%> / <%=TOTALPAGE%> , COUNT:<%=RECORDINDB%></FONT>
	</TD>
	<TD WIDTH=40% ALIGN=RIGHT>
	<%if session("netuser")="LSARY" then %>
			<%if yymm>=nowmonth then %>
				<input type="BUTTON" name="send" value="(Y)確　認" class=button ONCLICK="GO()">
				<input type="BUTTON" name="send" value="(N)取　消" class=button onclick="clr()">
			<%end if%>
		<%else%>	
			<input type="button" name="send" value="(Y)確　定" class=button onclick="go()" >
			<input type="BUTTON" name="send" value="(N)取　消" class=button  onclick="clr()">
		<%end if%>	
	</TD>
</TR>

</TABLE>
</form>




</body>
</html>
<%
Sub StoreToSession()
	Dim CurrentRow
	tmpRec = Session("empfilesalary")
	for CurrentRow = 1 to PageRec
		'tmpRec(CurrentPage, CurrentRow, 0) = request("op")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 6) = request("F1_JOB")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 19) = request("BBCODE")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 20) = request("BB")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 21) = request("CVCODE")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 22) = request("CV")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 23) = request("PHU")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 24) = request("NN")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 25) = request("KT")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 26) = request("MT")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 27) = request("TTKH")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 58) = request("JX")(CurrentRow)

	next
	Session("empfilesalary") = tmpRec

End Sub
%>

<script language=vbscript>
function BACKMAIN()
	open "../main.asp" , "_self"
end function

function clr()
	open "<%=SELF%>.asp" , "_self"
end function

function go()
	<%=self%>.action="<%=SELF%>.upd.asp"
	<%=self%>.submit()
end function

function oktest(N)
	tp=<%=self%>.totalpage.value
	cp=<%=self%>.CurrentPage.value
	rc=<%=self%>.RecordInDB.value
	open "empfile.show.asp?empautoid="& N , "_blank" , "top=10, left=10, width=550, scrollbars=yes"
end function 


function editmemo(index)
	tp=<%=self%>.totalpage.value
	cp=<%=self%>.CurrentPage.value
	rc=<%=self%>.RecordInDB.value
	YYMM = <%=self%>.YYMM.value
	open "<%=self%>.memo.asp?index="& index &"&currentpage=" & cp &"&yymm=" & yymm  , "_blank" , "top=10, left=10, width=450, height=450, scrollbars=yes"
end function 

FUNCTION BBCODECHG(INDEX)  
	codestr=<%=self%>.bbcode(index).value
	daystr=<%=self%>.MMDAYS.value
	open "<%=SELF%>.back.asp?ftype=A&index="& index &"&CurrentPage="& <%=CurrentPage%> & _
		 "&days=" & daystr & "&code=" &	codestr , "Back"
	'DATACHG(INDEX)

	'PARENT.BEST.COLS="70%,30%"
END FUNCTION

FUNCTION JOBCHG(INDEX)
	codestr=<%=self%>.F1_JOB(index).value
	daystr=<%=self%>.MMDAYS.value
	open "<%=SELF%>.back.asp?ftype=B&index="& index &"&CurrentPage="& <%=CurrentPage%> & _
		 "&days=" &daystr & "&code=" &	codestr , "Back"
	'PARENT.BEST.COLS="70%,30%"
	'DATACHG(INDEX)
END FUNCTION

FUNCTION DATACHG(INDEX)
	if isnumeric(<%=SELF%>.PHU(INDEX).VALUE)=false then
		alert "請輸入數字!!"
		<%=self%>.phu(index).focus()
		<%=self%>.phu(index).value=0
		<%=self%>.phu(index).select()
		exit FUNCTION
	end if
	if isnumeric(<%=SELF%>.NN(INDEX).VALUE)=false then
		alert "請輸入數字!!"
		<%=self%>.NN(index).value=0
		<%=self%>.NN(index).focus()
		<%=self%>.NN(index).select()
		exit FUNCTION
	end if
	if isnumeric(<%=SELF%>.KT(INDEX).VALUE)=false then
		alert "請輸入數字!!"
		<%=self%>.KT(index).value=0
		<%=self%>.KT(index).focus()
		<%=self%>.KT(index).select()
		exit FUNCTION
	end if
	if isnumeric(<%=SELF%>.MT(INDEX).VALUE)=false then
		alert "請輸入數字!!"
		<%=self%>.MT(index).value=0
		<%=self%>.MT(index).focus()
		<%=self%>.MT(index).select()
		exit FUNCTION
	end if
	if isnumeric(<%=SELF%>.TTKH(INDEX).VALUE)=false then
		alert "請輸入數字!!"
		<%=self%>.TTKH(index).value=0
		<%=self%>.TTKH(index).focus()
		<%=self%>.TTKH(index).select()
		exit FUNCTION
	end if

	if isnumeric(<%=SELF%>.TNKH(INDEX).VALUE)=false then  '其他收入
		alert "請輸入數字!!"
		<%=self%>.TNKH(index).value=0
		<%=self%>.TNKH(index).focus()
		<%=self%>.TNKH(index).select()
		exit FUNCTION
	end if

	if isnumeric(<%=SELF%>.HS(INDEX).VALUE)=false then  '伙食費(-)
		alert "請輸入數字!!"
		<%=self%>.HS(index).value=0
		<%=self%>.HS(index).focus()
		<%=self%>.HS(index).select()
		exit FUNCTION
	end if
	if isnumeric(<%=SELF%>.QITA(INDEX).VALUE)=false then  '其他扣除額(-)
		alert "請輸入數字!!"
		<%=self%>.QITA(index).value=0
		<%=self%>.QITA(index).focus()
		<%=self%>.QITA(index).select()
		exit FUNCTION
	end if

	if isnumeric(<%=SELF%>.JX(INDEX).VALUE)=false then  '其他扣除額(-)
		alert "請輸入數字!!"
		<%=self%>.JX(index).value=0
		<%=self%>.JX(index).focus()
		<%=self%>.JX(index).select()
		exit FUNCTION
	end if
	TTM = ( cdbl(<%=self%>.bb(index).value) + cdbl(<%=self%>.CV(index).value) + cdbl(<%=self%>.PHU(index).value) )
	if TTM mod (26*8)<>0 then
		TTMH = FIX (CDBL(TTM)/26/8 ) +1   '時薪
	else
		TTMH = FIX (CDBL(TTM)/26/8 )    '時薪
	end if
	'alert  TTMH
	'<%=self%>.HHMOENY(index).value = TTMH

	CODESTR01 = <%=SELF%>.PHU(INDEX).VALUE
	CODESTR02 = <%=SELF%>.NN(INDEX).VALUE
	CODESTR03 = <%=SELF%>.KT(INDEX).VALUE
	CODESTR04 = <%=SELF%>.MT(INDEX).VALUE
	CODESTR05 = <%=SELF%>.TTKH(INDEX).VALUE
	CODESTR06 = <%=SELF%>.TNKH(INDEX).VALUE
	CODESTR07 = <%=SELF%>.HS(INDEX).VALUE
	CODESTR08 = <%=SELF%>.QITA(INDEX).VALUE
	CODESTR09 = <%=SELF%>.JX(INDEX).VALUE
	CODESTR10 = <%=SELF%>.BH(INDEX).VALUE
	CODESTR11 = <%=SELF%>.GT(INDEX).VALUE
	daystr=<%=self%>.MMDAYS.value
	yymmstr=<%=self%>.yymm.value
	'ALERT CODESTR02
	'ALERT CODESTR03

	open "<%=SELF%>.back.asp?ftype=CDATACHG&index="&index &"&CurrentPage="& <%=CurrentPage%> & _
		 "&CODESTR01="& CODESTR01 &_
		 "&CODESTR02="& CODESTR02 &_
		 "&CODESTR03="& CODESTR03 &_
		 "&CODESTR04="& CODESTR04 &_
		 "&CODESTR05="& CODESTR05 &_
		 "&CODESTR06="& CODESTR06 &_
		 "&CODESTR07="& CODESTR07 &_
		 "&CODESTR08="& CODESTR08 &_
		 "&CODESTR09="& CODESTR09 &_
		 "&CODESTR10="& CODESTR10 &_
		 "&CODESTR11="& CODESTR11 &_
		 "&yymm="& yymmstr &_
		 "&days=" & daystr , "Back"

	'PARENT.BEST.COLS="70%,30%"

END FUNCTION

function view1(index)
	yymmstr = <%=self%>.yymm.value
	'alert yymmstr
	empidstr = <%=self%>.empid(index).value
	idstr= <%=SELF%>.empautoid(INDEX).VALUE
	'OPEN "../zzz/getempWorkTime.asp?yymm=" & yymmstr &"&EMPID=" & empidstr , "_blank" , "top=10, left=10,  width=650, scrollbars=yes"
	OPEN "empworkB.fore.asp?yymm=" & yymmstr &"&EMPID=" & empidstr , "_blank" , "top=10, left=10,  width=650, scrollbars=yes"
end function 

function view2(index)	
	yymmstr = <%=self%>.yymm.value
	'alert yymmstr
	empidstr = <%=self%>.empid(index).value
	idstr= <%=SELF%>.empautoid(INDEX).VALUE
	OPEN "showholiday.asp?yymm=" & yymmstr &"&EMPID=" & empidstr , "_blank" , "top=10, left=10,  width=650, scrollbars=yes" 	
end function

</script>


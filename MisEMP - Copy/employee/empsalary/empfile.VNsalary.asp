<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<!-- #include file = "../../GetSQLServerConnection.fun" --> 
<!-- #include file="../../ADOINC.inc" -->
<%
'on error resume next   
session.codepage="65001"
SELF = "empfilesalary"
 
Set conn = GetSQLServerConnection()	  
Set rs = Server.CreateObject("ADODB.Recordset")   
YYMM=REQUEST("YYMM")
whsno = trim(request("whsno"))
unitno = trim(request("unitno"))
groupid = trim(request("groupid"))
COUNTRY = trim(request("COUNTRY"))
job = trim(request("job"))
QUERYX = trim(request("empid1"))  
outemp = request("outemp")
lastym = left(yymm,4) &  right("00" & cstr(right(yymm,2)-1) ,2 )
if right(yymm,2)="01"  then 
	lastym = left(yymm,4)-1 &"12" 
end if 	 

PERAGE = REQUEST("PERAGE")

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
      

'本月假日天數 (星期日)
SQL=" SELECT * FROM YDBMCALE WHERE CONVERT(CHAR(6),DAT,112)  ='"& YYMM &"' AND  DATEPART( DW,DAT ) ='1'  " 
Set rsTT = Server.CreateObject("ADODB.Recordset")   
RSTT.OPEN SQL, CONN, 3, 3 
IF NOT RSTT.EOF THEN 
	HHCNT = CDBL(RSTT.RECORDCOUNT)
ELSE
	HHCNT = 0 
END IF 
SET RSTT=NOTHING   

'RESPONSE.WRITE HHCNT &"<br>" 
'RESPONSE.END  
'本月應記薪天數 
MMDAYS = CDBL(days)-CDBL(HHCNT) 
'RESPONSE.WRITE  MMDAYS 
'RESPONSE.END 
'----------------------------------------------------------------------------------------

sqlstr = "update empwork set kzhour=0 where yymm='"& YYMM &"'  and kzhour<0 " 
conn.execute(sqlstr) 

gTotalPage = 1
PageRec = 10    'number of records per page
TableRec = 70    'number of fields per record  

sql="select isnull(k.TOTJXM,0) TOTJXM,  isnull( L.sole , 0 ) TBTR , isnull( m. TNKH,0) TNKH ,  isnull(m.jx,0) jx,  "&_	
	"case when a.country='VN' then case when ISNULL(M.EMPID,'')='' then CASE WHEN ISNULL(A.BHDAT,'')<>'' THEN   bb_bonus * 0.06  ELSE 0 END else case when m.bh<>0 then  bb_bonus * 0.06 else m.bh  end end else 0 end  as BH  ,  "&_
	"case when a.country='VN' then case when ISNULL(M.EMPID,'')='' then case when  isnull(gtdat,'')<>'' then 5000 else 0 end else m.gt end  else 0 end as GT,  "&_
	"ISNULL(M.QITA,0) QITA, o.exrt,  "&_
	"isnull(n.forget,0) forget, isnull(n.h1,0) h1, isnull(n.h2,0) h2 , isnull(n.h3,0) h3, isnull(n.b3,0) b3 ,"&_
	"isnull(JA.jiaa,0) jiaa,isnull(JB.jiab,0) jiab,isnull(n.kzhour,0) kzhour, isnull(n.latefor,0) latefor, a.* from  "&_
	"( select * from  view_employee ) a  "&_
	"left join ( select * from empdsalary where  yymm='"& lastym &"' ) l on L.empid = a.empid and L.whsno = a.whsno "&_
	"left join ( select * from empdsalary where yymm='"& yymm &"' ) m on m.empid = a.empid and m.whsno = a.whsno   "&_  
	"left join ( select * from VYFYMYJX where yymm='"& yymm &"' ) k on k.empid = a.empid and k.groupid = a.groupid  "&_  
	"LEFT JOIN ( select empid empidN,  (sum(isnull(forget,0)))  forget  , (sum(isnull(h1,0))) h1, (sum(isnull(h2,0))) h2, (sum(isnull(h3,0))) h3, (sum(isnull(b3,0))) b3 ,  "&_
 	"(sum(isnull(jiaa,0))) jiaa, (sum(isnull(jiab,0))) jiab, ( sum(isnull(toth,0))) toth , ( sum(isnull(kzhour,0))) kzhour , (sum(latefor)) latefor "&_
 	"from empwork   where yymm='"& YYMM &"' GROUP BY EMPID )  N ON N.empidN = A.EMPID  "&_	
 	"left join  (  "&_
	"select jiaType as Ja , empid as EIDA, sum(hhour) as   jiaa   from  empholiday    where  convert(char(6), dateup, 112)='"& yymm &"'  and jiatype='A'  group  by empid, jiatype   "&_
	")  JA on JA.EIDA = a.empid   "&_
	"left join ( "&_
	"select jiaType as Jb , empid as EIDB, sum(hhour) as   jiaB   from  empholiday    where  convert(char(6), dateup, 112)='"& yymm &"'  and jiatype='B'  group  by empid, jiatype   "&_
	") JB on JB.eidb = a.empid   "&_ 	
	"LEFT JOIN ( SELECT *FROM VYFYEXRT  WHERE  YYYYMM='"& yymm &"' ) O ON O.code = a.DM  "&_
	"where CONVERT(CHAR(10), indat, 111)< '"& ccdt &"' and ( isnull(a.outdat,'')='' or a.outdat>'"& calcdt &"' )  "&_
	"and a.whsno like '%"& whsno &"%' and a.unitno like '%"& unitno &"%' and a.groupid like '%"& groupid &"%'  "&_
	"and a.COUNTRY like '%"& COUNTRY  &"%' and A.job like '%"& job &"%' and a.empid like '%"& QUERYX &"%' " 
	if outemp="D" then  
		sql=sql&" and ( isnull(a.outdat,'')<>'' and  a.outdat>'"& calcdt &"' )  " 
	end if
	
sql=sql&"order by a.empid   "
	
 
'response.write sql 
'response.end 
if request("TotalPage") = "" or request("TotalPage") = "0" then 
	CurrentPage = 1	
	rs.Open SQL, conn, 3, 3 
	IF NOT RS.EOF THEN 
		rs.PageSize = PageRec 
		RecordInDB = rs.RecordCount 
		TotalPage = rs.PageCount  
		gTotalPage = TotalPage
	END IF 	 

	Redim tmpRec(gTotalPage, PageRec, TableRec)   'Array
	
	for i = 1 to TotalPage 
	 for j = 1 to PageRec
		if not rs.EOF then 
			for k=1 to TableRec-1
				tmpRec(i, j, 0) = "no"
				tmpRec(i, j, 1) = trim(rs("empid"))
				tmpRec(i, j, 2) = trim(rs("empnam_cn"))
				tmpRec(i, j, 3) = trim(rs("empnam_vn"))
				tmpRec(i, j, 4) = rs("country")
				tmpRec(i, j, 5) = rs("date1")
				tmpRec(i, j, 6) = rs("job")				
				tmpRec(i, j, 7) = rs("whsno")	 
				tmpRec(i, j, 8) = rs("unitno")	 
				tmpRec(i, j, 9)	=RS("groupid") 
				tmpRec(i, j, 10)=RS("zuno") 				
				tmpRec(i, j, 11)=RS("whsnodesc") 	
				tmpRec(i, j, 12)=RS("unitdesc") 	
				tmpRec(i, j, 13)=RS("groupdesc") 	
				tmpRec(i, j, 14)=RS("zunodesc") 	
				tmpRec(i, j, 15)=RS("jobdesc") 	
				tmpRec(i, j, 16)=RS("countrydesc") 	
				tmpRec(i, j, 17)=RS("autoid") 	
				IF RS("zuno")="XX" THEN 
					tmpRec(i, j, 18)=""
				ELSE
					tmpRec(i, j, 18)=RS("zuno")
				END IF 
				tmpRec(i, j, 19)=RS("BB")
				tmpRec(i, j, 20)=RS("BB_bonus")  '基本薪資
				tmpRec(i, j, 21)=RS("CVcode")
				tmpRec(i, j, 22)=RS("CV_bonus")  '職務加給
				tmpRec(i, j, 23)=RS("PHU")		'Y獎金 
				tmpRec(i, j, 24)=RS("NN")  '語言加給
				tmpRec(i, j, 25)=RS("KT") '技術加給
				tmpRec(i, j, 26)=RS("MT") '環境加給
				tmpRec(i, j, 27)=RS("TTKH")  '其他加給
				tmpRec(i, j, 28)=RS("BHDAT") '買保險日期
				tmpRec(i, j, 29)=RS("GTDAT") '工團日期
				tmpRec(i, j, 30)=RS("OUTDAT") '離職日期
								
				tmpRec(i, j, 32)=ROUND(RS("TNKH"),0) '其他收入
				tmpRec(i, j, 33)=ROUND(RS("TBTR"),0) '上月補款  
				
				tmpRec(i, j, 34)=RS("BH") '保險費(-)  				
				tmpRec(i, j, 36)=RS("GT") '入工團費
				tmpRec(i, j, 37)=RS("QITA") '其他扣除額 
				
				TOTY=  CDBL( ( CDBL(tmpRec(i, j, 20))+CDBL(tmpRec(i, j, 22))+CDBL(tmpRec(i, j, 23)) )  )  'BB+CV+PHU				
				if rs("country")="VN" then 
					tmpRec(i, j, 38) =  round( CDBL(TOTY)/26/8,0)   '時薪 
				else 
					tmpRec(i, j, 38) =  round( CDBL(TOTY)/30/8,3)   '時薪 
				end if 	 				
				
				'TMONEY=CDBL(TOTY)+CDBL(tmpRec(i, j, 24))+CDBL(tmpRec(i, j, 25))+CDBL(tmpRec(i, j, 26))+CDBL(tmpRec(i, j, 27))+CDBL(tmpRec(i, j, 31))+CDBL(tmpRec(i, j, 32))+CDBL(tmpRec(i, j, 33))-CDBL(tmpRec(i, j, 34))-CDBL(tmpRec(i, j, 35))-CDBL(tmpRec(i, j, 36))-CDBL(tmpRec(i, j, 37))
				'tmpRec(i, j, 39) = CDBL(TMONEY) 
				
				tmpRec(i, j, 40) = CDBL(RS("H1")) 
				tmpRec(i, j, 41) = CDBL(RS("H2")) 
				tmpRec(i, j, 42) = CDBL(RS("H3")) 
				if rs("country")="VN" then 
					tmpRec(i, j, 43) = CDBL(RS("B3")) 
				else
					tmpRec(i, j, 43) =  0 
				end if 	
				tmpRec(i, j, 44) = CDBL(RS("KZHOUR")) 
				tmpRec(i, j, 45) = CDBL(RS("JIAA")) 
				tmpRec(i, j, 46) = CDBL(RS("JIAB")) 
				'40~43 加班費(+) 
				'44~46 請假或曠職(-)
				if tmpRec(i, j, 4)="VN" then 
					H1_money = ROUND((tmpRec(i, j, 38)*1.5) * cdbl(tmpRec(i, j, 40)),0) '平日加班工資(+) 時薪*1.5
				elseif tmpRec(i, j, 4)="TA" then
					if YYMM>="200608" then  
						H1_money = ROUND((tmpRec(i, j, 38)*1.37) * cdbl(tmpRec(i, j, 40)),3) '平日加班工資(+) 時薪*1.37(泰國)
					else
						H1_money = ROUND((tmpRec(i, j, 38)*1) * cdbl(tmpRec(i, j, 40)),3) '平日加班工資(+) 時薪*1(泰國)
					end if 	
				else
					H1_money = ROUND((tmpRec(i, j, 38)) * cdbl(tmpRec(i, j, 40)),0) '平日加班工資(+)
				end if 	
				if tmpRec(i, j, 4)="VN" then 
					H2_money = ROUND((tmpRec(i, j, 38)*2) * cdbl(tmpRec(i, j, 41)),0) '休假加班工資(+)時薪*2 
				elseif tmpRec(i, j, 4)="TA" then 
					H2_money = ROUND((tmpRec(i, j, 38)*1)* cdbl(tmpRec(i, j, 41)),3)   '休假加班工資(+) 時薪*1(泰國) 
				else
					H2_money = ROUND((tmpRec(i, j, 38)*1)* cdbl(tmpRec(i, j, 41)),0)   '休假加班工資(+) 時薪
				end if 	
				if tmpRec(i, j, 4)="VN" then 
					H3_money = ROUND((tmpRec(i, j, 38)*3) * cdbl(tmpRec(i, j, 42)),0) '節日加班工資(+)時薪*3
				else
					H3_money = 0
				end if 
				if tmpRec(i, j, 4)="VN" then 	
					b3_money = ROUND((tmpRec(i, j, 38)*0.3) * cdbl(tmpRec(i, j, 43)),0) '夜班加班工資(+)時薪*0.3
				else
					b3_money = 0 
				end if 	
				kz_money = ROUND(tmpRec(i, j, 38) * tmpRec(i, j, 44),0)
				jiaa_money = ROUND(tmpRec(i, j, 38) * tmpRec(i, j, 45),0)
				jiab_money = ROUND(tmpRec(i, j, 38) * tmpRec(i, j, 46),0) 
				
				'tmpRec(i, j, 47) = CDBL(TMONEY)+ CDBL(H1_money)+CDBL(H2_money)+CDBL(H3_money)+CDBL(b3_money)-CDBL(kz_money)-CDBL(jiaa_money)-CDBL(jiab_money)
				tmpRec(i, j, 48) = cdbl(rs("forget"))+cdbl(rs("latefor"))
				'--總加班工資
				tmpRec(i, j, 49) = round( CDBL(H1_MONEY) + CDBL(H2_money) + CDBL(H3_money) + CDBL(b3_money) ,0)   
				'response.write 
				'時假    
				tmpRec(i, j, 50) = kz_money + jiaa_money + jiab_money 
				
				tmpRec(i, j, 51) = H1_money 
				tmpRec(i, j, 52) = H2_money 
				tmpRec(i, j, 53) = H3_money 
				tmpRec(i, j, 54) = B3_money 
				tmpRec(i, j, 55) = kz_money 
				tmpRec(i, j, 56) = jiaa_money 
				tmpRec(i, j, 57) = jiab_money  				
				if rs("country")="TA" then 
					tmpRec(i, j, 58) = 0   '績效獎金
				else				
					tmpRec(i, j, 58) = rs("TOTJXM")  '績效獎金
				end if	
				
				'response.write H1_money &"<BR>"
				
				'員工工作天數(記薪天數) 
				SQLX="SELECT EMPID, ISNULL(SUM(HHOUR),0) HHOUR  FROM EMPHOLIDAY WHERE EMPID='"& tmpRec(i, j, 1) &"' AND CONVERT(CHAR(6), DATEUP,112)='"& YYMM &"' "&_
		 				  "AND JIATYPE IN ('A','B') GROUP BY EMPID " 
				Set rDs = Server.CreateObject("ADODB.Recordset")   		
		 		RDS.OPEN  SQLX, CONN, 3, 3   
		 		IF NOT RDS.EOF THEN   '員工本月請事病假天數
		 			A4=FIX(CDBL(RDS("HHOUR"))/8)
		 		ELSE
		 			A4=0  
		 		END IF  
		 		
				SET RDS=NOTHING 
				'1.本月離職員工(不含1日) 從本月1日計算至離職日前一天
		 		IF tmpRec(i, j, 30)="" THEN  '未離職 
		 			MWORKDAYS = CDBL(days) - CDBL(HHCNT)  
		 			tmpRec(i, j, 59) = MWORKDAYS 
		 		ELSE
		 			IF  tmpRec(i, j, 30) >= ccdt THEN  '非本月離職  
		 				MWORKDAYS = CDBL(days) - CDBL(HHCNT)  
		 				tmpRec(i, j, 59) = MWORKDAYS  
		 			ELSE
			 			A1=DATEDIFF("D",CDATE(calcdt),CDATE(tmpRec(i, j, 30)) )  '從1日到離職日天數  	  
			 			'RESPONSE.WRITE A1 &"<br>"
			 			'到離職日前的假日有幾天 	
			 			SQLS1="SELECT * FROM YDBMCALE WHERE CONVERT(CHAR(6),DAT,112)='"& YYMM &"' AND  CONVERT(CHAR(10),DAT,111)< '"& tmpRec(i, j, 30)&"' AND DATEPART( DW,DAT ) ='1'   " 
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
			 			MWORKDAYS = CDBL(A1)-CDBL(A2)-CDBL(A3) '**********本月工作天數********** 
			 			tmpRec(i, j, 59)  = MWORKDAYS    '**********本月工作天數**********
			 		END IF 	
		 		END IF  		 		
		 		'RESPONSE.WRITE  MWORKDAYS   		
		 		'2.本月新進員工 從到職日計算到本月底  
		 		IF CDATE(tmpRec(i, j, 5))>CDATE(calcdt) THEN  
		 			iF tmpRec(i, j, 30)="" THEN 
			 			A1= DATEDIFF("D", CDATE(tmpRec(i, j, 5)), CDATE(ENDdat))  		 			
			 			'RESPONSE.WRITE a1 &"<br>"
			 			
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
			 			tmpRec(i, j, 59) = MWORKDAYS 
			 		ELSE
			 			A1= DATEDIFF("D", CDATE(tmpRec(i, j, 5)), CDATE(tmpRec(i, j, 30))) 
			 			'RESPONSE.WRITE a1 &"<br>"			 			
			 			SQLS1="SELECT * FROM YDBMCALE WHERE CONVERT(CHAR(6),DAT,112)='"& YYMM &"' AND  CONVERT(CHAR(10),DAT,111)>= '"& tmpRec(i, j, 5)&"'  AND  CONVERT(CHAR(10),DAT,111)< '"& tmpRec(i, j, 30)&"' AND STATUS IN ('H2', 'H3' ) " 
			 			'RESPONSE.WRITE SQLS1 &"<br>"
			 			Set rDs = Server.CreateObject("ADODB.Recordset")   		
			 			RDS.OPEN  SQLS1, CONN, 3, 3 
			 			IF NOT RDS.EOF THEN   '到職後到月底的假日有幾天
			 				A2=RDS.RECORDCOUNT     
			 			ELSE
			 				A2=0 
			 			END IF 	
			 			SET RDS=NOTHING  
			 			'RESPONSE.WRITE a2 &"<br>"
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
			 			tmpRec(i, j, 59) = MWORKDAYS 
			 		END IF	
		 		ELSE
		 			'MWORKDAYS =  CDBL(days) - CDBL(HHCNT) - CDBL(A4) 
		 			tmpRec(i, j, 59) = tmpRec(i, j, 59)  
		 		END IF 	   
		 		
		 		'全勤 
				IF CDATE(tmpRec(i, j, 5)) >CDATE(calcdt) THEN   
					tmpRec(i, j, 31) = 0 
				ELSEif tmpRec(i, j, 59)< ( CDBL(days) - CDBL(HHCNT) ) THEN
					tmpRec(i, j, 31) = 0
				else						
					IF CDBL(RS("FORGET")+cdbl(rs("latefor")))>=3 and ( CDBL(RS("FORGET"))+cdbl(rs("latefor"))) < 6 THEN 
						if cdbl(rs("jiaa")) + cdbl(rs("jiab")) <=8 then 
							tmpRec(i, j, 31)=CDBL(RS("QC"))/2  '全勤 
						else 
							tmpRec(i, j, 31)= 0  
						end if 	
					ELSEIF 	( CDBL(RS("FORGET"))+cdbl(rs("latefor"))) >=6 THEN 
						tmpRec(i, j, 31)= 0 
					else						
						if  cdbl(rs("jiaa")) + cdbl(rs("jiab"))+cdbl(rs("kzhour"))=0 then 
							tmpRec(i, j, 31)=CDBL(RS("QC")) 
						elseif  cdbl(rs("jiaa")) + cdbl(rs("jiab"))+cdbl(rs("kzhour"))>=1 and  cdbl(rs("jiaa")) + cdbl(rs("jiab"))+cdbl(rs("kzhour")) <=8 then 
							tmpRec(i, j, 31)=CDBL(RS("QC"))/2 
						elseif 	cdbl(rs("jiaa")) + cdbl(rs("jiab"))+cdbl(rs("kzhour"))>=9 then 
							tmpRec(i, j, 31)= 0  
						else
							tmpRec(i, j, 31)=CDBL(RS("QC"))  
						end if 						
					end if 	
				END IF 	 
		 		
		 		 '伙食費
		 		IF RS("COUNTRY")="VN" THEN 		 			
			 		IF tmpRec(i, j, 59)< ( CDBL(days) - CDBL(HHCNT) ) THEN 
			 			tmpRec(i, j, 35) = 1000 * CDBL(tmpRec(i, j, 59))
			 		ELSE
						tmpRec(i, j, 35)=26000
					END IF 						
				ELSE
					tmpRec(i, j, 35)= 0 
				END IF 	 
				
		 		'應領薪資合計  
		 		'如為本月新進員工薪資OR本月離職工作未滿13天: 總薪資/26 * 工作天數 
		 		'舊員工本月離職, 工作天數13天(含)以上 : (BB+CV+PHU / 26 )* 工作天數  + ( NN+KT+MT+TTKH+QC 全薪 ) 
		 		ALLM=CDBL(tmpRec(i, j, 20))+CDBL(tmpRec(i, j, 22))+CDBL(tmpRec(i, j, 23))+CDBL(tmpRec(i, j, 24))+CDBL(tmpRec(i, j, 25))+CDBL(tmpRec(i, j, 26))+CDBL(tmpRec(i, j, 27))+CDBL(tmpRec(i, j, 31))  '本薪BB+CV+PHU+NN+KT+MT+TTKH+QC
		 		OTRM= CDBL(tmpRec(i, j, 32))+CDBL(tmpRec(i, j, 33))+cdbl(tmpRec(i, j, 58))  '其他收入+上月補款+績效		 		
		 		
			 	if trim(tmpRec(i, j, 30))<>"" and tmpRec(i, j, 30) < calcdt then 
			 		tmpRec(i, j, 60) = 0 
			 	else	 
					IF tmpRec(i, j, 59)< ( CDBL(days) - CDBL(HHCNT) ) THEN 
				 		IF ( CDBL(tmpRec(i, j, 59)) < 3  and CDATE(tmpRec(i, j, 5))>=CDATE(calcdt) )   then 
				 			tmpRec(i, j, 60) = 0 
				 		else 	
							IF  CDBL(tmpRec(i, j, 59)) < 13  THEN    		 			
					 			tmpRec(i, j, 60) = ( ROUND( CDBL(ALLM) /26 ,0) * CDBL(tmpRec(i, j, 59))) + OTRM  
					 			'response.write  tmpRec(i, j, 60) &"<BR>"
					 		ELSE 			 				
					 			tmpRec(i, j, 60) = ( ROUND(TOTY / 26 ,0) *  CDBL(tmpRec(i, j, 59))) + CDBL(tmpRec(i, j, 24))+CDBL(tmpRec(i, j, 25))+CDBL(tmpRec(i, j, 26))+CDBL(tmpRec(i, j, 27))+CDBL(tmpRec(i, j, 31))+ OTRM 
					 		END IF 	
					 	end if 
				 	ELSE 
						tmpRec(i, j, 60) = ALLM + OTRM 
					END IF 	 
			 	end if  
			 	
		 		'應發工資
		 		if tmpRec(i, j, 60) > 0 then  
		 			tmpRec(i, j, 39) = CDBL(tmpRec(i, j, 60))-CDBL(tmpRec(i, j, 34))-CDBL(tmpRec(i, j, 35))-CDBL(tmpRec(i, j, 36))-CDBL(tmpRec(i, j, 37)) 		 			
		 			'response.write tmpRec(i, j, 60)  &"<BR>"
		 			'response.write tmpRec(i, j, 34)  &"<BR>"
		 			'response.write tmpRec(i, j, 35)  &"<BR>"
		 			'response.write tmpRec(i, j, 36)  &"<BR>"
		 			'response.write tmpRec(i, j, 37)  &"<BR>"
		 		else
		 			tmpRec(i, j, 39) = 0 
		 		end if 
		 		'實領薪資
		 		if tmpRec(i, j, 60) > 0 then 			 			
		 			IF tmpRec(i, j, 59)< ( CDBL(days) - CDBL(HHCNT) ) THEN 			 				
			 			tmpRec(i, j, 47) = CDBL(tmpRec(i, j, 39))+ CDBL(H1_money)+CDBL(H2_money)+CDBL(H3_money)+CDBL(b3_money)-CDBL(kz_money)-CDBL(jiaa_money)-CDBL(jiab_money)
			 		else
			 			tmpRec(i, j, 47) = CDBL(tmpRec(i, j, 39))+ CDBL(H1_money)+CDBL(H2_money)+CDBL(H3_money)+CDBL(b3_money)-CDBL(kz_money)-CDBL(jiaa_money)-CDBL(jiab_money)
		 			end if 	
		 			IF rs("country")="TA" then  '超過8百萬越幣需扣所得稅
		 				if cdbl(tmpRec(i, j, 47))*cdbl(rs("exrt"))>=8000000 then 
		 					tmpRec(i, j, 66) =round(  ( ( cdbl(tmpRec(i, j, 47))*cdbl(rs("exrt"))-8000000 ) / cdbl(rs("exrt")) )  * 0.1 ,0 ) 
		 					tmpRec(i, j, 47) = round( round(cdbl(tmpRec(i, j, 47)),0) - cdbl(tmpRec(i, j, 66)) , 0) 
		 				else
		 					tmpRec(i, j, 66) = tmpRec(i, j, 35) 
		 				end if 	
		 			end if 		 				
		 		else
		 			tmpRec(i, j, 47) = 0
		 		end if 	
		 		
				'離職補助金 
				if tmpRec(i, j, 4)="VN" then 
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
					 	elseif inMonth mod 12 > 6 then 
					 		jishu=round(fix(inMonth/12 )*0.5 +0.25 ,2) 
					 	else
					 		jishu=round(fix(inMonth/12 )*0.5 ,2)
					 	end if 
					else
					 	jishu=0 
					end if 
					tmpRec(i, j, 61) = jishu   '基數
					IF jishu > 0 THEN 
					 	tmpRec(i, j, 62) = ROUND( CDBL(tmpRec(i, j, 20)) * CDBL(jishu) ,0) + CDBL(tmpRec(i, j, 22))+CDBL(tmpRec(i, j, 23))+CDBL(tmpRec(i, j, 24))+CDBL(tmpRec(i, j, 25))+CDBL(tmpRec(i, j, 26))
					ELSE
						tmpRec(i, j, 62) = 0 
					END IF 	  
				else 
					tmpRec(i, j, 62) = 0  
				end if 		
				 '檢核總薪資 
				 tmpRec(i, j, 63) = ( cdbl(ALLM) + cdbl(OTRM)  + cdbl(tmpRec(i, j, 49)) ) -  cdbl(tmpRec(i, j, 50) ) -CDBL(tmpRec(i, j, 34))-CDBL(tmpRec(i, j, 35))-CDBL(tmpRec(i, j, 36))- CDBL(tmpRec(i, j, 37))-CDBL(tmpRec(i, j, 66))    
				 tmpRec(i, j, 64) =   ( cdbl(ALLM) + cdbl(OTRM) + cdbl(tmpRec(i, j, 49))  )    
				 tmpRec(i, j, 65) =  round(cdbl(tmpRec(i, j, 63)),0) - round(cdbl(tmpRec(i, j, 47)),0) 
				 'response.write  tmpRec(i, j, 63) &"<BR>"   
				 'response.write  tmpRec(i, j, 47) &"<BR>"   
			next
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


FUNCTION FDT(D)
	Response.Write YEAR(D)&"/"&RIGHT("00"&MONTH(D),2)&"/"&RIGHT("00"&DAY(D),2) 
	
END FUNCTION 
%>

<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="refresh">
<link rel="stylesheet" href="../../Include/style.css" type="text/css">
<link rel="stylesheet" href="../../Include/style2.css" type="text/css"> 
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
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
<form name="<%=self%>" method="post" action="empfile.salary.asp">
<INPUT TYPE=HIDDEN NAME="UID" VALUE=<%=SESSION("NETUSER")%>>
<INPUT TYPE=hidden NAME=TotalPage VALUE="<%=TotalPage%>">
<INPUT TYPE=hidden NAME=CurrentPage VALUE="<%=CurrentPage%>">
<INPUT TYPE=hidden NAME=RecordInDB VALUE="<%=RecordInDB%>">
<INPUT TYPE=hidden NAME=PageRec VALUE="<%=PageRec%>">
<INPUT TYPE=hidden NAME=gTotalPage VALUE="<%=gTotalPage%>"> 	
<INPUT TYPE=hidden NAME=YYMM VALUE="<%=YYMM%>"> 	
<INPUT TYPE=hidden NAME=MMDAYS VALUE="<%=MMDAYS%>">
<table width="600" border="0" cellspacing="0" cellpadding="0">
	<tr>
	<TD width=430>
		<img border="0" src="../../image/icon.gif" align="absmiddle">
		人事薪資系統( 員工薪資管理 )　
		計薪年月：<%=YYMM%>
	</TD>	
	</tr>
</table> 

<hr size=0	style='border: 1px dotted #999999;' align=left width=500>		

<TABLE  CLASS="FONT9" BORDER=0 cellspacing="0" cellpadding="1" > 	
	<TR HEIGHT=25 BGCOLOR="LightGrey"   >
 		<TD ROWSPAN=2 >項次</TD>
 		<TD align=center>工號</TD> 		
 		<TD COLSPAN=3  >員工姓名(中,英,越)</TD> 
 		<td  align=center>時薪</td>
 		<td align=center>到職日期</td>
 		<td align=center>離職日期</td>
 		<TD align=center>工作天數</TD>
 		<TD align=center>總加班費</TD> 		 		 
 		<td align=center>(-)扣時假</td>
 		<td align=center>(-)其他</td> 	 		
 		<TD align=center>不足月扣款</TD> 		
 		<TD COLSPAN=4 ALIGN=CENTER bgcolor="#ccff99">加班(H)</TD>
 		<TD bgcolor="#ffcc99"></TD>
 		<TD bgcolor="#ffcc99"></TD>	
 		<TD COLSPAN=2 ALIGN=CENTER bgcolor="#ffcccc">請假(H)</TD> 		
 	</TR>
 	<tr BGCOLOR="LightGrey"  HEIGHT=25 > 	
 		<TD align=center>薪資代碼</TD>
 		<TD align=center>基本薪資</TD>
 		<TD align=center>職專</TD> 			
 		<TD align=center>職專加給</TD>	
 		<TD align=center>獎金(Y)</TD> 		 		
 		<TD align=center>績效獎金</TD>
 		<td align=center>其他加給</td> 	 		
 		<td align=center>其他收入</td>  	
 		<TD align=center>應領薪資</TD>	
 		<td align=center>(-)保險費</td>
 		<td align=center>(-)所得稅</td>  		
 		<TD ALIGN=CENTER >實領工資</TD>
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
		ELSE
			WKCOLOR="#DFEFFF"
		END IF 	 
		'if tmpRec(CurrentPage, CurrentRow, 1) <> "" then 
	%>
	<TR BGCOLOR=<%=WKCOLOR%> > 		
		<TD ROWSPAN=2 ALIGN=CENTER >
		<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then %><%=(CURRENTROW)+((CURRENTPAGE-1)*10)%><%END IF %>
		</TD>
 		<TD  >
 			<a href='vbscript:oktest(<%=tmpRec(CurrentPage, CurrentRow, 17)%>)'>
 				<%=tmpRec(CurrentPage, CurrentRow, 1)%>
 			</a>
 			<input type=hidden name=empid value="<%=tmpRec(CurrentPage, CurrentRow, 1)%>">
 			<input type=hidden name="empautoid" value="<%=tmpRec(CurrentPage, CurrentRow, 17)%>">
 		</TD> 		
 		<TD COLSPAN=3>
 			<a href='vbscript:oktest(<%=tmpRec(CurrentPage, CurrentRow, 17)%>)'>
 				<%=tmpRec(CurrentPage, CurrentRow, 2)%>
 				<font class=txt8><%=tmpRec(CurrentPage, CurrentRow, 3)%></font>
 			</a>
 		</TD>
 		<TD >
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>	
 				<INPUT NAME=HHMOENY  VALUE="<%=tmpRec(CurrentPage, CurrentRow, 38)%>" CLASS='INPUTBOX8' SIZE=7 STYLE="TEXT-ALIGN:RIGHT;COLOR:#3366cc;BACKGROUND-COLOR:LIGHTYELLOW" >  
 			<%ELSE%>	
 				<INPUT NAME=HHMOENY TYPE=HIDDEN> 
 			<%END IF %>	 				
 		</TD>
 		<TD  ALIGN=CENTER ><FONT CLASS=TXT8><%=RIGHT(tmpRec(CurrentPage, CurrentRow, 5),8)%></FONT></TD>
 		<TD  ALIGN=CENTER ><FONT CLASS=TXT8><%=RIGHT(tmpRec(CurrentPage, CurrentRow, 30),8)%></FONT></TD>
 		<TD > 		 		
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>	
	 			<INPUT NAME=WORKDAYS CLASS='INPUTBOX8' READONLY  SIZE=7 VALUE="<%=tmpRec(CurrentPage, CurrentRow, 59)%>" STYLE="TEXT-ALIGN:RIGHT;BACKGROUND-COLOR:Gainsboro" onblur="DATACHG(<%=CURRENTROW-1%>)" title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 工作天數">
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=WORKDAYS >
	 		<%END IF%>	
 		</TD>   		
 		<TD > 			
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
	 			<INPUT NAME=TOTJB CLASS='INPUTBOX8' SIZE=7 VALUE="<%=tmpRec(CurrentPage, CurrentRow, 49)%>" STYLE="TEXT-ALIGN:RIGHT;BACKGROUND-COLOR:LIGHTYELLOW"  READONLY  title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 總加班費">
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=TOTJB >
	 		<%END IF%>	  		
 		</TD>
 		<!--TD ALIGN=CENTER ><FONT CLASS=TXT8 COLOR="SeaGreen"><%=RIGHT(tmpRec(CurrentPage, CurrentRow, 28),8)%></FONT></TD> 		
 		<TD ALIGN=CENTER ><FONT CLASS=TXT8><%IF (tmpRec(CurrentPage, CurrentRow, 29))<>"" THEN%>Y<%END IF%></FONT></TD-->
 		
 		<TD >
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
	 			<INPUT NAME=TOTKJ CLASS='INPUTBOX8' SIZE=7 VALUE="<%=tmpRec(CurrentPage, CurrentRow, 50)%>" STYLE="TEXT-ALIGN:RIGHT;BACKGROUND-COLOR:LIGHTYELLOW" READONLY   title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 扣時假">
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=TOTKJ >
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
	 			<INPUT NAME=BZKM CLASS='INPUTBOX8' SIZE=10 VALUE="<%=formatnumber(tmpRec(CurrentPage, CurrentRow, 65),0)%>" STYLE="TEXT-ALIGN:RIGHT;BACKGROUND-COLOR:LIGHTYELLOW" READONLY title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 不足月扣款">
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=BZKM   >
	 		<%END IF%>	  		
	 		
 		</TD> 
 		<TD COLSPAN=4 align=center>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
 				<FONT CLASS=TXT8><u><div style="cursor:hand" onclick="view1(<%=currentrow-1%>)"><%=tmpRec(CurrentPage, CurrentRow, 1)%> 出勤紀錄</div></u></font>
 			<%END IF %>	
 		</TD> 
 		<TD ></TD>
 		<TD ></TD>
 		<TD COLSPAN=2 align=center>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
 				<FONT CLASS=TXT8><u><div style="cursor:hand" >看請假紀錄</div></u></font>
 			<%END IF %>		
 		</TD>  		
	</TR>
	<TR BGCOLOR=<%=WKCOLOR%> >
 		<TD ALIGN=RIGHT > 			
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>			
		 		<select name=BBCODE  class="txt8" style="width:60" onchange="bbcodechg(<%=currentrow-1%>)">				
					<%SQL="SELECT * FROM empsalarybasic WHERE FUNC='AA'  ORDER BY CODE "
					SET RST = CONN.EXECUTE(SQL)
					WHILE NOT RST.EOF  
					%>
					<option value="<%=RST("CODE")%>" <%IF RST("CODE")=trim(tmpRec(CurrentPage, CurrentRow, 19)) THEN %> SELECTED <%END IF%> ><%=RST("CODE")%></option>				 
					<%
					RST.MOVENEXT
					WEND 
					%>
				</SELECT>
				<%SET RST=NOTHING %>
			<%else%>
				<input type=hidden name=BBCODE >	
			<%end if %>
 		</TD>
 		<TD ALIGN=RIGHT>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
	 			<INPUT NAME=BB CLASS='INPUTBOX8' SIZE=7 VALUE="<%=tmpRec(CurrentPage, CurrentRow, 20)%>" STYLE="TEXT-ALIGN:RIGHT;BACKGROUND-COLOR:LIGHTYELLOW"  READONLY title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 資本薪資">	 			
	 		<%else%>
				<input type=hidden name=BB >	
			<%end if %>	
 		</TD>
 		<TD ALIGN=RIGHT ><!--職等--> 			
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
	 			<select name=F1_JOB  class="txt8" style="width:60" ONCHANGE="JOBCHG(<%=CURRENTROW-1%>)">				
					<%SQL="SELECT * FROM BASICCODE WHERE FUNC='LEV'  ORDER BY SYS_TYPE "
					SET RST = CONN.EXECUTE(SQL)
					WHILE NOT RST.EOF  
					%>
					<option value="<%=RST("SYS_TYPE")%>" <%IF RST("SYS_TYPE")=trim(tmpRec(CurrentPage, CurrentRow, 6)) THEN %> SELECTED <%END IF%> ><%=RST("SYS_VALUE")%></option>				 
					<%
					RST.MOVENEXT
					WEND 
					%>
				</SELECT>
				<%SET RST=NOTHING %>
			<%else%>
				<input type=hidden name=F1_JOB >	
			<%end if %>
 		</TD>
 		 <TD  ALIGN=RIGHT>
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
	 			<INPUT NAME=PHU CLASS='INPUTBOX8' SIZE=7 VALUE="<%=tmpRec(CurrentPage, CurrentRow, 23)%>" STYLE="TEXT-ALIGN:RIGHT" onblur="DATACHG(<%=CURRENTROW-1%>)" title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 補助獎金(Y)" >
	 		<%ELSE%>	
				<INPUT TYPE=HIDDEN NAME=PHU	>
			<%END IF%>	
 		</TD> 		 		
 		<TD>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
	 			<INPUT NAME=JX CLASS='INPUTBOX8' SIZE=7 VALUE="<%=tmpRec(CurrentPage, CurrentRow, 58)%>" STYLE="TEXT-ALIGN:RIGHT" onblur="DATACHG(<%=CURRENTROW-1%>)"  title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 績效獎金">
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=JX >
	 		<%END IF%>	
 		</TD>
 		<TD  ALIGN=RIGHT>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>	
	 			<INPUT NAME=TTKH CLASS='INPUTBOX8' SIZE=7 VALUE="<%=tmpRec(CurrentPage, CurrentRow, 27)%>" STYLE="TEXT-ALIGN:RIGHT" onblur="DATACHG(<%=CURRENTROW-1%>)" title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 其他加給">
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=TTKH >
	 		<%END IF%>			
 		</TD> 		
 		
 		<TD>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
	 			<INPUT NAME=TNKH CLASS='INPUTBOX8' SIZE=7 VALUE="<%=tmpRec(CurrentPage, CurrentRow, 32)%>" STYLE="TEXT-ALIGN:RIGHT" onblur="DATACHG(<%=CURRENTROW-1%>)" title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 其他收入">
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=TNKH >
	 		<%END IF%>	
 		</TD>
 		<TD>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
	 			<INPUT NAME=TOTM CLASS='INPUTBOX8' SIZE=7 VALUE="<%=tmpRec(CurrentPage, CurrentRow, 64)%>" STYLE="TEXT-ALIGN:RIGHT;BACKGROUND-COLOR:LIGHTYELLOW;color:#cc0000"  READONLY  title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 應領薪資">
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=TOTM >
	 		<%END IF%>	
 		</TD>  
 		
 		<TD>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>	
	 			<INPUT NAME=BH CLASS='INPUTBOX8' SIZE=7 VALUE="<%=tmpRec(CurrentPage, CurrentRow, 34)%>" STYLE="TEXT-ALIGN:RIGHT;COLOR:#3366cc" onblur="DATACHG(<%=CURRENTROW-1%>)"  title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 保險費"  >
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=BH >
	 		<%END IF%>	
 		</TD>
 		<TD>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>	
	 			<INPUT NAME=HS CLASS='INPUTBOX8' SIZE=7 VALUE="<%=tmpRec(CurrentPage, CurrentRow, 66)%>" STYLE="TEXT-ALIGN:RIGHT" onblur="DATACHG(<%=CURRENTROW-1%>)" title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 所得稅" >
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=HS>
	 		<%END IF%>	
 		</TD>
 		<TD>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>	 					
	 			<INPUT NAME=RELTOTMONEY CLASS='INPUTBOX8' VALUE="<%=formatnumber( tmpRec(CurrentPage, CurrentRow, 47),0)%>" SIZE=10  STYLE="TEXT-ALIGN:RIGHT;BACKGROUND-COLOR:LIGHTYELLOW;;color:#cc0000" READONLY title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 實領工資">
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=RELTOTMONEY  >
	 		<%END IF%>	
 		</TD>
 		<TD>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>	 					
	 			<INPUT NAME=H1 CLASS='INPUTBOX8' VALUE="<%=tmpRec(CurrentPage, CurrentRow, 40)%>"  SIZE=4 STYLE="TEXT-ALIGN:RIGHT;BACKGROUND-COLOR:LIGHTYELLOW;color:MediumBlue" READONLY >
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=H1 >
	 		<%END IF%>	
 		</TD>
 		<TD>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>	 					
	 			<INPUT NAME=H2 CLASS='INPUTBOX8' VALUE="<%=tmpRec(CurrentPage, CurrentRow, 41)%>"  SIZE=4 STYLE="TEXT-ALIGN:RIGHT;BACKGROUND-COLOR:LIGHTYELLOW;color:MediumBlue" READONLY >
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=H2 >
	 		<%END IF%>	
 		</TD>
 		<TD>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>	 					
	 			<INPUT NAME=H3 CLASS='INPUTBOX8' VALUE="<%=tmpRec(CurrentPage, CurrentRow, 42)%>"  SIZE=4 STYLE="TEXT-ALIGN:RIGHT;BACKGROUND-COLOR:LIGHTYELLOW;color:MediumBlue" READONLY >
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=H3 >
	 		<%END IF%>	
 		</TD>
 		<TD>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>	 					
	 			<INPUT NAME=B3 CLASS='INPUTBOX8'  VALUE="<%=tmpRec(CurrentPage, CurrentRow, 43)%>"  SIZE=4  STYLE="TEXT-ALIGN:RIGHT;BACKGROUND-COLOR:LIGHTYELLOW;color:MediumBlue" READONLY >
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=B3 >
	 		<%END IF%>	
 		</TD> 		
 		<TD>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>	 					
	 			<INPUT NAME=KZHOUR CLASS='INPUTBOX8' VALUE="<%=tmpRec(CurrentPage, CurrentRow, 44)%>"  SIZE=4  STYLE="TEXT-ALIGN:RIGHT;BACKGROUND-COLOR:LIGHTYELLOW;color:Brown" READONLY >
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=KZHOUR >
	 		<%END IF%>	
 		</TD>
 		<TD>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>	 					
	 			<INPUT NAME=Forget CLASS='INPUTBOX8' VALUE="<%=tmpRec(CurrentPage, CurrentRow, 48)%>"  SIZE=4  STYLE="TEXT-ALIGN:RIGHT;BACKGROUND-COLOR:LIGHTYELLOW;color:Brown" READONLY >
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=Forget >
	 		<%END IF%>	
 		</TD> 		
 		<TD>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>	 					
	 			<INPUT NAME=JIAA CLASS='INPUTBOX8' VALUE="<%=tmpRec(CurrentPage, CurrentRow, 45)%>"  SIZE=4 STYLE="TEXT-ALIGN:RIGHT;BACKGROUND-COLOR:LIGHTYELLOW;color:ForestGreen" READONLY >
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=JIAA >
	 		<%END IF%>	
 		</TD>
 		<TD>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>	 					
	 			<INPUT NAME=JIAB CLASS='INPUTBOX8' VALUE="<%=tmpRec(CurrentPage, CurrentRow, 46)%>"  SIZE=4 STYLE="TEXT-ALIGN:RIGHT;BACKGROUND-COLOR:LIGHTYELLOW;color:ForestGreen" READONLY >
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=JIAB >
	 		<%END IF%>	
 		</TD>
 		
	</TR>
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


<TABLE border=0 width=500 class=font9 >
<tr>
    <td align="CENTER" height=40 WIDTH=75%>    
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
	<TD WIDTH=25% ALIGN=RIGHT>		
		<input type="BUTTON" name="send" value="確　認" class=button ONCLICK="GO()">
		<input type="BUTTON" name="send" value="取　消" class=button onclick="clr()">
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
	open "empsalary.fore.asp" , "_self"
end function 

function go()	
	<%=self%>.action="empsalary.upd.asp"  
	<%=self%>.submit()
end function 

function oktest(N)
	tp=<%=self%>.totalpage.value 
	cp=<%=self%>.CurrentPage.value 
	rc=<%=self%>.RecordInDB.value 
	open "empfile.show.asp?empautoid="& N , "_blank" , "top=10, left=10, width=550, scrollbars=yes" 
end function 

FUNCTION BBCODECHG(INDEX)
	codestr=<%=self%>.bbcode(index).value 
	daystr=<%=self%>.MMDAYS.value 	
	open "empsalary.back.asp?ftype=A&index="& index &"&CurrentPage="& <%=CurrentPage%> & _
		 "&days=" & daystr & "&code=" &	codestr , "Back" 		 
	'DATACHG(INDEX)	  
	 
	'PARENT.BEST.COLS="70%,30%"	 	
END FUNCTION 

FUNCTION JOBCHG(INDEX)
	codestr=<%=self%>.F1_JOB(index).value 
	daystr=<%=self%>.MMDAYS.value 
	open "empsalary.back.asp?ftype=B&index="& index &"&CurrentPage="& <%=CurrentPage%> & _
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
	
	if isnumeric(<%=SELF%>.JX(INDEX).VALUE)=false then  '績效(-)
		alert "請輸入數字!!"
		<%=self%>.JX(index).value=0 		
		<%=self%>.JX(index).focus()
		<%=self%>.JX(index).select()
		exit FUNCTION 
	end if  
	TTM = ( cdbl(<%=self%>.bb(index).value) + cdbl(<%=self%>.CV(index).value) + cdbl(<%=self%>.PHU(index).value) ) 
	TTMH = round (CDBL(TTM)/26/8,0 )    '時薪
	
	'alert  TTMH 
	'<%=self%>.HHMOENY(index).value = TTMH 
	
	CODESTR01 = <%=SELF%>.PHU(INDEX).VALUE
	CODESTR02 = 0
	CODESTR03 = 0
	CODESTR04 = 0
	CODESTR05 = <%=SELF%>.TTKH(INDEX).VALUE
	CODESTR06 = <%=SELF%>.TNKH(INDEX).VALUE
	CODESTR07 = <%=SELF%>.HS(INDEX).VALUE
	CODESTR08 = <%=SELF%>.QITA(INDEX).VALUE
	CODESTR09 = <%=SELF%>.JX(INDEX).VALUE 
	CODESTR10 = <%=SELF%>.BH(INDEX).VALUE 
	CODESTR11 = 0
	daystr=<%=self%>.MMDAYS.value  
	'ALERT CODESTR02
	'ALERT CODESTR03
	
	open "empsalary.back.asp?ftype=CDATACHG&index="&index &"&CurrentPage="& <%=CurrentPage%> & _
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
		 "&CODESTR11="& CODESTR11 &"&days=" & daystr , "Back"  
		 
	'PARENT.BEST.COLS="70%,30%"	 
	
END FUNCTION  

function view1(index) 	 
	yymmstr = <%=self%>.yymm.value 
	'alert yymmstr
	empidstr = <%=self%>.empid(index).value 
	idstr= <%=SELF%>.empautoid(INDEX).VALUE
	OPEN "../zzz/getempWorkTime.asp?yymm=" & yymmstr &"&EMPID=" & empidstr , "_blank" , "top=10, left=10,  scrollbars=yes" 
end function 
	
</script>


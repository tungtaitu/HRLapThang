<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<!-- #include file = "../GetSQLServerConnection.fun" --> 
<!-- #include file="../ADOINC.inc" -->
<%response.buffer=true%>
<%
'on error resume next   
session.codepage="65001"
SELF = "YECE03"
 
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
f_exrt = REQUEST("f_exrt")
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
if COUNTRY="VN" then 
	MMDAYS = CDBL(days)-CDBL(HHCNT) 
else
	MMDAYS = CDBL(days)
end if 	


'RESPONSE.WRITE  MMDAYS &"<BR>"
'RESPONSE.END 
'----------------------------------------------------------------------------------------

sqlstr = "update empwork set kzhour=0 where yymm='"& YYMM &"'  and kzhour<0 " 
conn.execute(sqlstr)  

recalc  = request("recalc")
if recalc="Y" then 
	sql="delete empdsalary where yymm='"& YYMM &"' and isnull(country,'')='"& COUNTRY &"' and isnull(whsno,'') like '%"& whsno &"'"
	conn.execute(Sql)
end if   

gTotalPage = 1
PageRec = 10    'number of records per page
TableRec = 71    'number of fields per record  

sql="select isnull(m.empid,'') as Nowemp, ROUND(isnull(k.TOTJXM,0)/ISNULL(O.EXRT,1),0)  TOTJXM,  isnull( L.sole , 0 ) TBTR , isnull( m. TNKH,0) TNKH ,  isnull(m.jx,0) jx,  "&_	
	"isnull(m.empid,'') eid, case when a.country='VN' then case when ISNULL(M.EMPID,'')='' then CASE WHEN ISNULL(A.BHDAT,'')<>'' THEN   a.bb * 0.06  ELSE 0 END else case when m.bh<>0 then  a.bb * 0.06 else m.bh  end end else 0 end  as BH  ,  "&_
	"case when a.country='VN' then case when ISNULL(M.EMPID,'')='' then case when  isnull(gtdat,'')<>'' then 5000 else 0 end else m.gt end  else 0 end as GT,  "&_
	"ISNULL(M.QITA,0) QITA, o.exrt Nexrt,  isnull(m.zhuanM,0) ZhuanM, isnull(m.xianM,0) XianM, isnull(m.memo,'') as salarymemo,  isnull(m.dkm,0) dkm, "&_
	"isnull(n.forget,0) forget, isnull(n.h1,0) h1, isnull(n.h2,0) h2 , isnull(n.h3,0) h3, isnull(n.b3,0) b3 ,"&_
	"isnull(JA.jiaa,0) jiaa,isnull(JB.jiab,0) jiab,isnull(n.kzhour,0) kzhour, isnull(n.latefor,0) latefor, a.* from  "&_
	"( select * from  view_empfile where country='TA' and CONVERT(CHAR(10), indat, 111)< '"& ccdt &"' and ( isnull(outdat,'')='' or outdat>'"& calcdt &"' )  ) a  "&_
	"left join ( select * from view_empgroup  where  yymm='"& yymm  &"' ) c  on c.empid = a.empid  "&_ 
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
	"where  "&_
	"isnull(c.lw,a.whsno) like '"& whsno &"%'   and isnull(c.lg,a.groupid ) like '"& groupid &"%'  "&_
	"and a.COUNTRY like '"& COUNTRY  &"%' and A.job like '"& job &"%' and a.empid like '%"& QUERYX &"%' " 
	if outemp="D" then  
		sql=sql&" and ( isnull(a.outdat,'')<>'' and  a.outdat>'"& calcdt &"' )  " 
	end if
	
sql=sql&"order by c.lw desc , a.empid   "
	
 
'response.write sql 
'response.end 
if request("TotalPage") = "" or request("TotalPage") = "0" then 
	CurrentPage = 1	
	rs.Open SQL, conn, 3, 3 
	IF NOT RS.EOF THEN 
		f_exrt = rs("nexrt")
		rs.PageSize = PageRec 
		RecordInDB = rs.RecordCount 
		TotalPage = rs.PageCount  
		gTotalPage = TotalPage
	END IF 	 

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
				tmpRec(i, j, 20)=RS("BB")  '基本薪資
				BB=cdbl(RS("BB"))
				tmpRec(i, j, 21)=RS("cv")
				tmpRec(i, j, 22)=RS("CV")  '職務加給				
				CV=cdbl(rs("CV"))				
				tmpRec(i, j, 23)=RS("PHU")		'Y獎金 
				PHU=cdbl(rs("PHU"))
				tmpRec(i, j, 24)=RS("NN")  '語言加給
				NN=cdbl(rs("NN"))
				tmpRec(i, j, 25)=RS("KT") '技術加給
				KT=cdbl(rs("KT"))
				tmpRec(i, j, 26)=RS("MT") '環境加給
				MT=cdbl(rs("MT"))
				tmpRec(i, j, 27)=RS("TTKH")  '其他加給
				TTKH=cdbl(rs("TTKH"))
				tmpRec(i, j, 28)=RS("BHDAT") '買保險日期
				tmpRec(i, j, 29)=RS("GTDAT") '工團日期
				tmpRec(i, j, 30)=trim(RS("OUTDATE")) '離職日期								
				tmpRec(i, j, 32)=ROUND(RS("TNKH"),0) '其他收入
				TNKH=cdbl(rs("TNKH")) 
				tmpRec(i, j, 33) = 0  '上月補款    	
				tmpRec(i, j, 34)=0 '保險費(-) 				
				tmpRec(i, j, 36)=RS("GT") '入工團費
				tmpRec(i, j, 37)=RS("QITA") '其他扣除額  	
				
				 
				
				TOTY= CDBL(tmpRec(i, j, 20))+CDBL(tmpRec(i, j, 22))+CDBL(tmpRec(i, j, 23))+cdbl(tmpRec(i, j, 25))     'BB+CV+PHU+KT
				

				tmpRec(i, j, 38) =  round(TOTY/30/8,3)   '時薪  				
				'tmpRec(i, j, 39) = cdbl(TOTY) + cdbl(tmpRec(i, j, 27))   		'no use 				 
				tmpRec(i, j, 40) = CDBL(RS("H1")) 
				tmpRec(i, j, 41) = CDBL(RS("H2")) 
				tmpRec(i, j, 42) = CDBL(RS("H3")) 
				tmpRec(i, j, 43) =  0   '夜班津貼 (外國人無) 
				
				tmpRec(i, j, 44) = CDBL(RS("KZHOUR")) 
				tmpRec(i, j, 45) = CDBL(RS("JIAA")) 
				tmpRec(i, j, 46) = CDBL(RS("JIAB")) 
				'40~43 加班費(+) 
				'44~46 請假或曠職(-)
				'泰國加班費 (平日=時薪*1.37 * 加班時數  , 假日=時薪*加班時數  )
				if YYMM>="200608" then    
					H1_money = ROUND((tmpRec(i, j, 38)*1.37) * cdbl(tmpRec(i, j, 40)),0) '平日加班工資(+) 時薪*1.37(泰國)
				else
					H1_money = ROUND((tmpRec(i, j, 38)*1) * cdbl(tmpRec(i, j, 40)),0) '假日加班工資(+) 時薪*1(泰國)
				end if 	 				 

				H2_money = ROUND((tmpRec(i, j, 38)*1)* cdbl(tmpRec(i, j, 41)),0)   '休假加班工資(+) 時薪*1(泰國)  				 
				H3_money = ROUND((tmpRec(i, j, 38)*1)* cdbl(tmpRec(i, j, 42)),0)   '休假加班工資(+) 時薪*1(泰國)  				  
				b3_money = 0   
				'曠職以扣平日加班方式計算算  				
				kz_money = ROUND(tmpRec(i, j, 38)*1.37 * tmpRec(i, j, 44),0) 
				
				if yymm>="200812" then 
					ALLM=BB 
				else
					ALLM=BB+CV+PHU+NN+KT+MT+TTKH+QC
				end if 	 				 
				'if cdbl(tmpRec(i, j, 45))<=24 then 
				jiaa_money = ROUND(CDBL(tmpRec(i, j, 38)) * tmpRec(i, j, 45),0)  '請事假  (  基本薪資/30/8 * 小時 )     
				'end if 	
				jiab_money = ROUND(CDBL(tmpRec(i, j, 38)) * tmpRec(i, j, 46),0)  '請病假 
				
				jiaAB = cdbl(jiaa_money)+cdbl(jiab_money)
				
				'response.write jiaa_money  
				'tmpRec(i, j, 47) = CDBL(TMONEY)+ CDBL(H1_money)+CDBL(H2_money)+CDBL(H3_money)+CDBL(b3_money)-CDBL(kz_money)-CDBL(jiaa_money)-CDBL(jiab_money)
				tmpRec(i, j, 48) = cdbl(rs("forget"))+cdbl(rs("latefor"))  '忘刷遲到早退
				'--總加班工資				
				tmpRec(i, j, 49)=round( CDBL(H1_MONEY) + CDBL(H2_money) + CDBL(H3_money) + CDBL(b3_money) ,0)    				
				'response.write 
				'時假   
				tmpRec(i, j, 50) = kz_money + jiaa_money + jiab_money  '(曠職扣款+事假扣款+病假扣款) 
				'tmpRec(i, j, 50) = kz_money 
				
				tmpRec(i, j, 51) = H1_money 
				tmpRec(i, j, 52) = H2_money 
				tmpRec(i, j, 53) = H3_money 
				tmpRec(i, j, 54) = B3_money 
				tmpRec(i, j, 55) = kz_money 
				tmpRec(i, j, 56) = jiaa_money 
				tmpRec(i, j, 57) = jiab_money  				
				'if rs("country")="TA" then 
				'	tmpRec(i, j, 58) = 0   '績效獎金
				'else				
				if rs("eid")="" then 
					tmpRec(i, j, 58) = rs("TOTJXM")  '績效獎金
				else	
					tmpRec(i, j, 58) = rs("jx")  '績效獎金
				end if 	
				'end if	
				
				'response.write H1_money &"<BR>"
				
				'員工工作天數(記薪天數,扣除返鄉休假)
				SQLX="SELECT EMPID, ISNULL(SUM(HHOUR),0) HHOUR , min(convert(char(10),dateup,111)) as minJiaDat  , max(convert(char(10),dateup,111)) as maxJiaDat FROM EMPHOLIDAY WHERE EMPID='"& tmpRec(i, j, 1) &"' AND CONVERT(CHAR(6), DATEUP,112)='"& YYMM &"' "&_
		 				  "AND (JIATYPE ='I'  or isnull(place,'')='W' ) GROUP BY EMPID "
		 		'response.write sqlx &"<BR>"
				Set rDs = Server.CreateObject("ADODB.Recordset")
		 		RDS.OPEN  SQLX, CONN, 3, 3
		 		IF NOT RDS.EOF THEN   '員工本月請事病假天數
					allJiaABDays = RDS("HHOUR")
					if cdbl(RDS("HHOUR"))>=24 then 
						A4 = datediff("d",rds("minJiaDat"), rds("maxJiaDat"))+1
						'response.write "???="& datediff("d",rds("minJiaDat"), rds("maxJiaDat")) &"<BR>"
					else
						A4=FIX(CDBL(RDS("HHOUR"))/8)
					end if 	
		 		ELSE
					allJiaABDays = 0 
		 			A4=0
		 		END IF 
				'response.write rs("empid") & A4  &"<BR>"
				set rds=nothing  
				 
				'1.本月離職員工(不含1日) 從本月1日計算至離職日前一天
		 		IF tmpRec(i, j, 30)="" THEN  '未離職 
		 			'MWORKDAYS = CDBL(days) - CDBL(HHCNT)  
		 			MWORKDAYS = CDBL(days) 
		 			tmpRec(i, j, 59) = MWORKDAYS -A4
		 		ELSE
		 			IF  tmpRec(i, j, 30) >= ccdt THEN  '非本月離職  
		 				MWORKDAYS = CDBL(days) 
		 				tmpRec(i, j, 59) = MWORKDAYS  -A4
		 			ELSE
			 			A1=DATEDIFF("D",CDATE(calcdt),CDATE(tmpRec(i, j, 30)) )  '從1日到離職日天數  	  
			 			MWORKDAYS = CDBL(A1)+1 '**********本月工作天數********** 
			 			tmpRec(i, j, 59)  = MWORKDAYS-A4    '**********本月工作天數**********
			 		END IF 	
		 		END IF  		 		
		 		'RESPONSE.WRITE  MWORKDAYS   		
		 		'2.本月新進員工 從到職日計算到本月底  
		 		IF CDATE(tmpRec(i, j, 5))>CDATE(calcdt) THEN  '本月到職本月仍在職
		 			iF tmpRec(i, j, 30)="" THEN 
			 			A1= DATEDIFF("D", CDATE(tmpRec(i, j, 5)), CDATE(ENDdat))  		 			
			 			'RESPONSE.WRITE a1 &"<br>"			 			
			 			MWORKDAYS = cdbl(A1)
			 			tmpRec(i, j, 59) = MWORKDAYS+1 '**********本月工作天數**********
			 		ELSE  
			 			A1= DATEDIFF("D", CDATE(tmpRec(i, j, 5)), CDATE(tmpRec(i, j, 30))) '本月到職本月離職
			 			'RESPONSE.WRITE a1 &"<br>"			 						 			 
			 			MWORKDAYS = cdbl(A1) 
			 			tmpRec(i, j, 59) = MWORKDAYS -A4 '**********本月工作天數**********
			 		END IF	
		 		ELSE		 			
		 			tmpRec(i, j, 59) = tmpRec(i, j, 59)   '**********本月工作天數**********
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
				tmpRec(i, j, 35)= 0 
		
				
		 		'應領薪資合計  
		 		'如為本月新進員工薪資OR本月離職工作未滿13天: 總薪資/26 * 工作天數 
		 		'舊員工本月離職, 工作天數13天(含)以上 : (BB+CV+PHU / 26 )* 工作天數  + ( NN+KT+MT+TTKH+QC 全薪 ) 
		 		'本薪BB+CV+PHU+NN+KT+MT+TTKH+QC
		 		'ALLM=CDBL(tmpRec(i, j, 20))+CDBL(tmpRec(i, j, 22))+CDBL(tmpRec(i, j, 23))+CDBL(tmpRec(i, j, 24))+CDBL(tmpRec(i, j, 25))+CDBL(tmpRec(i, j, 26))+CDBL(tmpRec(i, j, 27))+CDBL(tmpRec(i, j, 31))  
		 		ALLM=BB+CV+PHU+NN+KT+MT
		 		OTRM= CDBL(tmpRec(i, j, 32))+CDBL(tmpRec(i, j, 33))+cdbl(tmpRec(i, j, 58))+cdbl(TTKH)  '其他收入+上月補款+績效		 		
		 		
			 	if trim(tmpRec(i, j, 30))<>"" and tmpRec(i, j, 30) < calcdt then 
			 		tmpRec(i, j, 60) = 0 
			 	else	 
					IF tmpRec(i, j, 59)< ( CDBL(days) ) THEN 
				 		IF ( CDBL(tmpRec(i, j, 59)) < 3  and CDATE(tmpRec(i, j, 5))>=CDATE(calcdt) )   then 
				 			tmpRec(i, j, 60) = 0 
				 			'response.write "A" &"<BR>"
				 		else 	
							IF  CDBL(tmpRec(i, j, 59)) < 13  THEN    		 			
					 			tmpRec(i, j, 60) = round ( cdbl(ALLM)/cdbl(MMDAYS)* CDBL(tmpRec(i, j, 59)),0) + OTRM  					 			
					 			'response.write "B" &"<BR>"
					 		ELSE 			 				
					 			tmpRec(i, j, 60) = round ( cdbl(ALLM)/cdbl(MMDAYS)* CDBL(tmpRec(i, j, 59)),0) + OTRM 
					 			' response.write Allm  &"<BR>"
								' response.write ( (cdbl(ALLM)/cdbl(MMDAYS))* CDBL(tmpRec(i, j, 59)))  &"<BR>"
								' response.write tmpRec(i, j, 59) &"," & cdbl(MMDAYS)  &"<BR>"
								' response.write OTRM  &"<BR>"
								' response.write tmpRec(i, j, 60)  &"<BR>"
								' response.write "C" &"<BR>"
					 		END IF 	
					 	end if 
				 	ELSE 
						IF cdbl(tmpRec(i, j, 44)) >= 208  then 
							tmpRec(i, j, 60) = 0 
						else
							tmpRec(i, j, 60) = cdbl(ALLM) + cdbl(OTRM)
						end if 	
						'response.write "D" &"<BR>"
					END IF 	 
			 	end if  
			 	'response.write  tmpRec(i, j, 1) &"<BR>"
			 	'response.write  allm &"<BR>"
			  'response.write  tmpRec(i, j, 59) &"<BR>"
			  'response.write  "應發工資=" & tmpRec(i, j, 60) &"<BR>" 	
		 		'應發工資 tmpRec(i, j, 60) = 工作天數應領薪資 
		 		'tmpRec(i, j, 39) = 工作天數應領薪資-保險-其他   
				bh = cdbl(tmpRec(i, j, 34))
				QITA = cdbl(tmpRec(i, j, 37))
		 		if tmpRec(i, j, 60) > 0 then  
		 			tmpRec(i, j, 39) = CDBL(tmpRec(i, j, 60))-CDBL(tmpRec(i, j, 34))- CDBL(tmpRec(i, j, 37))
		 		else
		 			tmpRec(i, j, 39) = 0 
		 		end if 
		 		
				'實領薪資 				
		 		if tmpRec(i, j, 60) > 0 then 			 		
					if cdbl(allJiaABDays)<=24 then 
						tmpRec(i, j, 47) = CDBL(tmpRec(i, j, 39))+ CDBL(tmpRec(i, j, 49))-CDBL(tmpRec(i, j, 50))		 			
					else
						tmpRec(i, j, 47) = CDBL(tmpRec(i, j, 39))+ CDBL(tmpRec(i, j, 49))-CDBL(kz_money)		 			
					end if 	
		 		else
		 			tmpRec(i, j, 47) = 0
		 		end if 	

				real_TOTAMT = cdbl(tmpRec(i, j, 47)) *cdbl(rs("nexrt")) ' 實領金額
 
				totb = 4000000
				if left(yymm,4)>"2008" then 
					sql2="exec sp_calctax '"& real_TOTAMT &"','"& totb &"' "
					set ors=conn.execute(sql2) 
					F_tax = ors("tax")
				else
					sql2="exec sp_calctax_HW_2008 '"& real_TOTAMT &"' "
					set ors=conn.execute(sql2) 
					F_tax = ors("tax")
				end if  				
				set ors=nothing  				
				tmpRec(i, j, 66) = round(cdbl(F_tax) /cdbl(rs("nexrt")),0)
				KTAXM = round(cdbl(F_tax) /cdbl(rs("nexrt")),0)  				
		 		tmpRec(i, j, 47)  = cdbl(tmpRec(i, j, 47))- cdbl(KTAXM)
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
				 if cdbl(allJiaABDays)<=24 then 					
					tmpRec(i, j, 63) = ( cdbl(ALLM) + cdbl(OTRM)  + cdbl(tmpRec(i, j, 49)) ) - cdbl(tmpRec(i, j, 34) )- cdbl(tmpRec(i, j, 50) ) - CDBL(tmpRec(i, j, 37))- CDBL(tmpRec(i, j, 66)  )    
					'response.write rs("empid") & ", A1=" & tmpRec(i, j, 63) &  ", 47="& tmpRec(i, j, 47) &"<BR>"
				 else	 					
					tmpRec(i, j, 63) = ( cdbl(ALLM) + cdbl(OTRM)  + cdbl(tmpRec(i, j, 49)) ) - cdbl(tmpRec(i, j, 34) )- cdbl(kz_money) - CDBL(tmpRec(i, j, 37))- CDBL(tmpRec(i, j, 66)  )    
					'response.write rs("empid") & ", A2=" & tmpRec(i, j, 63)& ", 47="& tmpRec(i, j, 47) & "<BR>"
				 end if 
				 tmpRec(i, j, 64) =   ( cdbl(ALLM) + cdbl(OTRM) + cdbl(tmpRec(i, j, 49))  )    
				 tmpRec(i, j, 65) =  round(cdbl(tmpRec(i, j, 63)),0) - round(cdbl(tmpRec(i, j, 47)),0)   '不足月
				 
				'response.write  tmpRec(i, j, 63) &"<BR>"   
				'response.write  tmpRec(i, j, 64) &"<BR>"   
				'response.write  tmpRec(i, j, 47) &"<BR>"    				 
				'response.write rs("empid") & ", "& tmpRec(i, j, 65) &"<BR>"   				
				' RESPONSE.WRITE ALLM &"<br>"
				' RESPONSE.WRITE OTRM &"<br>" 
				  if rs("Nowemp")="" then 
				 	tmpRec(i, j, 67) = tmpRec(i, j, 47) 
				 	tmpRec(i, j, 68) = 0 
				 else
				 	tmpRec(i, j, 67) =tmpRec(i, j, 47) 
				 	tmpRec(i, j, 68) =0
				 end if   
				 tmpRec(i, j, 69) = rs("salarymemo") 
				 
				 if datediff("d",rs("nindat"), enddat)<180 then 
				 	tmpRec(i, j, 70) =  cdbl(tmpRec(i, j, 47)) * 0.25 
				 else
				 	tmpRec(i, j, 70) = 0 
				 end if 	
				 tmpRec(i, j, 71)  = rs("nexrt")
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
	Session("YECE03") = tmpRec	
else    
	TotalPage = cint(request("TotalPage"))
	'StoreToSession()
	CurrentPage = cint(request("CurrentPage"))
	RecordInDB  = REQUEST("RecordInDB")
	tmpRec = Session("YECE03")
	
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
<link rel="stylesheet" href="../Include/style.css" type="text/css">
<link rel="stylesheet" href="../Include/style2.css" type="text/css"> 
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
<form name="<%=self%>" method="post" action="<%=self%>.salary.asp">
<INPUT TYPE=HIDDEN NAME="UID" VALUE=<%=SESSION("NETUSER")%>>
<INPUT TYPE=hidden NAME=TotalPage VALUE="<%=TotalPage%>">
<INPUT TYPE=hidden NAME=CurrentPage VALUE="<%=CurrentPage%>">
<INPUT TYPE=hidden NAME=RecordInDB VALUE="<%=RecordInDB%>">
<INPUT TYPE=hidden NAME=PageRec VALUE="<%=PageRec%>">
<INPUT TYPE=hidden NAME=gTotalPage VALUE="<%=gTotalPage%>"> 	
<INPUT TYPE=hidden NAME=YYMM VALUE="<%=YYMM%>"> 	
<INPUT TYPE=hidden NAME=country VALUE="<%=country%>"> 	
<INPUT TYPE=hidden NAME=MMDAYS VALUE="<%=MMDAYS%>">
<INPUT TYPE=hidden NAME=cfg VALUE="TA">
<INPUT TYPE=hidden NAME="f_exrt" VALUE="<%=f_exrt%>">

<table width="600" border="0" cellspacing="0" cellpadding="0" class="txt">
	<tr>
	<TD width=500>
		<img border="0" src="../image/icon.gif" align="absmiddle">
		<%=session("pgname")%>　計薪年月：<%=YYMM%> 國籍：<%=COUNTRY%>, exrt=<%=f_exrt%></TD>	
	</tr>
</table> 
<hr size=0	style='border: 1px dotted #999999;' align=left width=500>		
<TABLE  CLASS="FONT9" BORDER=0 cellspacing="0" cellpadding="1" > 	
	<TR HEIGHT=25 BGCOLOR="LightGrey"   >
 		<TD ROWSPAN=2 >項次<br>STT</TD>
 		<TD align=center valign="top">工號<br>So the</TD> 		
 		<TD COLSPAN=3  valign="top">員工姓名(中,英,越)Ho Ten</TD>  		
 		<td align=center valign="top">到職日<br>NVX</td>
 		<td align=center valign="top">離職日<br>NTV</td>
 		<td  align=center valign="top">時薪</td>
 		<TD align=center valign="top">工作天數<br>So ngay<br>lam viec</TD>
 		<TD align=center valign="top">總加班費</TD> 		 		 
 		<td align=center valign="top">(-)扣時假</td>
 		<td align=center valign="top">(-)保險費<br>phi B.H</td>
 		<td align=center valign="top">暫扣款</td>
 		<TD align=center valign="top">扣不足月</TD> 		
  		
 		<TD COLSPAN=3 ALIGN=CENTER bgcolor="#ccff99">加班(H)</TD> 		 			
 		<TD COLSPAN=3 ALIGN=CENTER >備註Ghi chu</TD> 		
 	</TR>
 	<tr BGCOLOR="LightGrey"  HEIGHT=25 > 	 		
 		<TD align=center valign="top">基薪<br>CB</TD>
 		<TD align=center valign="top">職專<br>ma CV</TD> 			
 		<TD align=center valign="top">職加<br>CV</TD>	
 		<TD align=center valign="top">補助(Y)<br>phu cap</TD> 		 		
 		<TD align=center valign="top">技術<br>KT</TD>		
 		<td align=center valign="top">其加<br>Phu cap<br>khac</td> 	 		
		<td align=center valign="top">合計<br>total</td>
		<TD align=center valign="top">績效獎金<br>Tien thuong</TD>
 		<td align=center valign="top">其他收入<br>Thu nhap<br>khac</td>  	
 		<TD align=center valign="top">應領薪資</TD>	 		
 		<td align=center valign="top">(-)其他<br>khac</td> 
 		<td align=center valign="top">(-)所得稅<br>Thue</td>  		
 		<TD ALIGN=CENTER valign="top">實領工資</TD>
 		
 		<TD ALIGN=CENTER bgcolor="#ccff99">平日</TD>
 		<TD ALIGN=CENTER bgcolor="#ccff99">休息</TD>
 		<TD ALIGN=CENTER bgcolor="#ccff99">假日</TD> 		
 		<TD ALIGN=CENTER bgcolor="#ffcc99"><font color=Brown>曠職</font></TD> 		
 		<TD ALIGN=CENTER bgcolor="#ffcccc" >事假</TD>
 		<TD ALIGN=CENTER bgcolor="#ffcccc" >病假</TD> 		
 	</tr> 
	<% Response.Flush %>
	<%for CurrentRow = 1 to PageRec
		IF CurrentRow MOD 2 = 0 THEN 
			WKCOLOR="LavenderBlush"
		ELSE
			WKCOLOR="#DFEFFF"
		END IF 	 
		'if tmpRec(CurrentPage, CurrentRow, 1) <> "" then   
		if trim(tmpRec(CurrentPage, CurrentRow, 1))<>"" then
			bb = tmpRec(CurrentPage, CurrentRow, 20)  
			cv = tmpRec(CurrentPage, CurrentRow, 22)  
			phu = tmpRec(CurrentPage, CurrentRow, 23)  
			kt = tmpRec(CurrentPage, CurrentRow, 25)  
			ttkh  = tmpRec(CurrentPage, CurrentRow, 27)  
			TB_money = cdbl(BB)+cdbl(CV)+cdbl(Phu)+cdbl(kt)+cdbl(Ttkh) 
		else
			TB_money  = 0
		end if 	
	%>
	<TR BGCOLOR=<%=WKCOLOR%> > 		
		<TD ROWSPAN=2 ALIGN=CENTER  >
		<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then %><%=(CURRENTROW)+((CURRENTPAGE-1)*10)%><%END IF %>
		</TD>
 		<TD  valign="top" >
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
 		
 		<TD  ALIGN=CENTER ><FONT CLASS=TXT8><%=RIGHT(tmpRec(CurrentPage, CurrentRow, 5),8)%></FONT></TD>
 		<TD  ALIGN=CENTER ><FONT CLASS=TXT8><%=RIGHT(tmpRec(CurrentPage, CurrentRow, 30),8)%></FONT></TD>
 		<TD >
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>	
 				<INPUT NAME=HHMOENY  VALUE="<%=tmpRec(CurrentPage, CurrentRow, 38)%>" CLASS='INPUTBOX8' SIZE=7 STYLE="TEXT-ALIGN:RIGHT;COLOR:#3366cc;BACKGROUND-COLOR:LIGHTYELLOW" >  
 			<%ELSE%>	
 				<INPUT NAME=HHMOENY TYPE=HIDDEN> 
 			<%END IF %>	 				
 		</TD>
 		<TD > 		 		
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>	
	 			<INPUT NAME=WORKDAYS CLASS='INPUTBOX8' READONLY  SIZE=7 VALUE="<%=tmpRec(CurrentPage, CurrentRow, 59)%>" STYLE="TEXT-ALIGN:center;BACKGROUND-COLOR:Gainsboro" onblur="DATACHG(<%=CURRENTROW-1%>)" title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 工作天數">
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
	 			<INPUT NAME=BH CLASS='INPUTBOX8' SIZE=7 VALUE="<%=tmpRec(CurrentPage, CurrentRow, 34)%>" STYLE="TEXT-ALIGN:RIGHT;COLOR:#3366cc" onblur="DATACHG(<%=CURRENTROW-1%>)"  title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 保險費"  >
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=BH >
	 		<%END IF%>	
 		</TD>		  		
 		<TD>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>	 					
	 			<INPUT NAME=DKM CLASS='INPUTBOX8' SIZE=7 VALUE="<%=formatnumber(tmpRec(CurrentPage, CurrentRow, 70),0)%>" STYLE="TEXT-ALIGN:RIGHT;BACKGROUND-COLOR:LIGHTYELLOW" READONLY title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 暫扣款,代收代付">
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=DKM >
	 		<%END IF%>	 		
 		</TD>  	 
 		<TD>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>	 					
	 			<INPUT NAME=BZKM CLASS='INPUTBOX8' VALUE="<%=formatnumber(tmpRec(CurrentPage, CurrentRow, 65),0)%>" STYLE="TEXT-ALIGN:RIGHT;BACKGROUND-COLOR:LIGHTYELLOW<%if cdbl(tmpRec(CurrentPage, CurrentRow, 65))>0 then%>;color:blue;<%end if%>"  SIZE="7" READONLY title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 不足月扣款">
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=BZKM   > 	 			
	 		<%END IF%>	   	 		
	 		<INPUT type=hidden NAME=XIANM CLASS='INPUTBOX8' SIZE=10 VALUE="<%=formatnumber(tmpRec(CurrentPage, CurrentRow, 68),0)%>" STYLE="TEXT-ALIGN:RIGHT;BACKGROUND-COLOR:LIGHTYELLOW;color:blue"  readonly  >
 		</TD>  	 	 
 		<TD COLSPAN=3 align=center>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
 				<FONT CLASS=TXT8><u><div style="cursor:hand" onclick="view1(<%=currentrow-1%>)"><%=tmpRec(CurrentPage, CurrentRow, 1)%> 出勤紀錄</div></u></font>
 			<%END IF %>	
 		</TD> 
 		<TD COLSPAN=3 align=center>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
 				<input name=memo class=inputbox maxlength=255 value="<%=tmprec(currentpage, currentRow, 69)%>" onchange=memochg(<%=currentrow-1%>) >
 			<%else%>	
 				<input name=memo type=hidden >
 			<%END IF %>		
 		</TD>  		
	</TR>
	<TR BGCOLOR=<%=WKCOLOR%> ><!--Line 2 ----------------> 		
 		<TD ALIGN=RIGHT>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
	 			<INPUT NAME=BB CLASS='INPUTBOX8' SIZE=7 VALUE="<%=tmpRec(CurrentPage, CurrentRow, 20)%>" STYLE="TEXT-ALIGN:RIGHT;BACKGROUND-COLOR:LIGHTYELLOW"  READONLY title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 資本薪資">	 			
				<input type=hidden name=BBCODE value="<%=trim(tmpRec(CurrentPage, CurrentRow, 19))%>" >	
	 		<%else%>
				<input type=hidden name=BB >	
				<input type=hidden name=BBCODE   >	
			<%end if %>	
 		</TD>
 		<TD ALIGN=RIGHT ><!--職等--> 			
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
	 			<select name=F1_JOB_S  class="txt8" style="width:60" disabled  >				
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
				<input type=hidden name=F1_JOB_S >	
			<%end if %>
			<input type=hidden name=F1_JOB value="<%=trim(tmpRec(CurrentPage, CurrentRow, 6))%>">	
 		</TD>
 		 <TD  ALIGN=RIGHT>
 		 	<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
 		 		<INPUT NAME=CV CLASS='readonly8' READONLY SIZE=7 VALUE="<%=tmpRec(CurrentPage, CurrentRow, 22)%>" STYLE="TEXT-ALIGN:right;BACKGROUND-COLOR:LIGHTYELLOW"    title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 職務加給" >
 		 		<input type=hidden name=CVCODE VALUE="<%=tmpRec(CurrentPage, CurrentRow, 21)%>" SIZE=3>
 		 	<%else%>
				<input type=hidden name=CV >	
				<input type=hidden name=CVCODE >
			<%end if %>	
 		 </TD>
 		<TD  ALIGN=RIGHT>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
	 			<INPUT NAME=PHU CLASS='readonly8' READONLY SIZE=7 VALUE="<%=tmpRec(CurrentPage, CurrentRow, 23)%>" STYLE="TEXT-ALIGN:RIGHT"  title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 補助獎金(Y)" >
	 		<%ELSE%>	
				<INPUT TYPE=HIDDEN NAME=PHU	>
			<%END IF%>	
 		</TD> 		 		
		<TD ALIGN=RIGHT > 			
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>			
				<INPUT NAME=KT CLASS='readonly8' readonly SIZE=7 VALUE="<%=tmpRec(CurrentPage, CurrentRow, 25)%>" STYLE="TEXT-ALIGN:RIGHT;BACKGROUND-COLOR:LIGHTYELLOW"  READONLY title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 技術">	 							
			<%else%>
				<input type=hidden name="kt" value="0">		
			<%end if %>
 		</TD>	
 		<TD  ALIGN=RIGHT>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>	
	 			<INPUT NAME=TTKH CLASS='readonly8' readonly  SIZE=7 VALUE="<%=tmpRec(CurrentPage, CurrentRow, 27)%>" STYLE="TEXT-ALIGN:RIGHT"   title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 其他加給">
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=TTKH >
	 		<%END IF%>			
 		</TD> 		
		<td><!--基本薪資合計-->
			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
	 			<INPUT NAME=TB_money CLASS='readonly8' readonly SIZE=7 VALUE="<%=TB_money%>" STYLE="TEXT-ALIGN:RIGHT"  title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 薪資小計">
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=TB_money value="0"  >
	 		<%END IF%>			
		</td>
 		<TD>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
	 			<INPUT NAME=JX CLASS='inpt8blue' SIZE=7 VALUE="<%=tmpRec(CurrentPage, CurrentRow, 58)%>" STYLE="TEXT-ALIGN:RIGHT" onblur="DATACHG(<%=CURRENTROW-1%>)"  title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 績效獎金">
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=JX >
	 		<%END IF%>	
 		</TD>		
 		
 		<TD>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
	 			<INPUT NAME=TNKH CLASS='inpt8red' SIZE=7 VALUE="<%=tmpRec(CurrentPage, CurrentRow, 32)%>" STYLE="TEXT-ALIGN:RIGHT" onblur="DATACHG(<%=CURRENTROW-1%>)" title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 其他收入">
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
	 			<INPUT NAME=QITA CLASS='INPUTBOX8' SIZE=7 VALUE="<%=tmpRec(CurrentPage, CurrentRow, 37)%>" STYLE="TEXT-ALIGN:RIGHT" onblur="DATACHG(<%=CURRENTROW-1%>)" title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 扣除其他" >
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=QITA >
	 		<%END IF%>
 		</TD>
 		<TD>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>	
	 			<INPUT NAME=ktaxM CLASS='INPUTBOX8' SIZE=7 VALUE="<%=tmpRec(CurrentPage, CurrentRow, 66)%>" STYLE="TEXT-ALIGN:RIGHT;color:blue" onblur="DATACHG(<%=CURRENTROW-1%>)" title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 所得稅" >
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=katxm>
	 		<%END IF%>	
 		</TD>
 		<TD>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>	 					
	 			<INPUT NAME=RELTOTMONEY CLASS='INPUTBOX8' VALUE="<%=formatnumber( tmpRec(CurrentPage, CurrentRow, 47),0)%>" SIZE=7  STYLE="TEXT-ALIGN:RIGHT;BACKGROUND-COLOR:LIGHTYELLOW;;color:#cc0000" READONLY title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 實領工資">
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=RELTOTMONEY value="0" > 	 		
	 		<%END IF%>	
	 			<INPUT type=hidden NAME=ZHUANM CLASS='INPUTBOX8' SIZE=10 VALUE="<%=formatnumber(tmpRec(CurrentPage, CurrentRow, 67),0)%>" STYLE="TEXT-ALIGN:RIGHT"  onblur="zhuanmchg(<%=CURRENTROW-1%>)"  >
	 			<INPUT TYPE=HIDDEN NAME=exrt value="<%=tmpRec(CurrentPage, CurrentRow, 71)%>"   >
				<INPUT TYPE=HIDDEN NAME=XIANMBAK  VALUE="<%=tmpRec(CurrentPage, CurrentRow, 68)%>" >
				<INPUT TYPE=HIDDEN NAME=ZHUANMBAK  VALUE="<%=tmpRec(CurrentPage, CurrentRow, 67)%>" >
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
	 			<INPUT  type=hidden  NAME=B3 CLASS='INPUTBOX8'  VALUE="<%=tmpRec(CurrentPage, CurrentRow, 43)%>"  SIZE=4  STYLE="TEXT-ALIGN:RIGHT;BACKGROUND-COLOR:LIGHTYELLOW;color:MediumBlue" READONLY >
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
<input type=hidden name=BB value="0">
<input type=hidden name=F1_JOB>
<input type=hidden name=CV value="0">
<input type=hidden name=CVCODE >
<input type=hidden name=PHU value="0">
<input type=hidden name=NN value="0">
<input type=hidden name=KT value="0">
<input type=hidden name=MT value="0">
<input type=hidden name=TTKH value="0">
<INPUT TYPE=HIDDEN NAME=exrt  value="0">
<INPUT TYPE=HIDDEN NAME=XIANMBAK   >
<INPUT TYPE=HIDDEN NAME=ZHUANMBAK  >
<INPUT TYPE=HIDDEN NAME=XIANM   value="0" >
<INPUT TYPE=HIDDEN NAME=ZHUANM value="0" >  
<INPUT TYPE=HIDDEN NAME=ktaxm value="0">  
<INPUT TYPE=HIDDEN NAME=RELTOTMONEY value="0" > 	 		


<TABLE border=0 width=550 class=font9 >
<tr>
    <td align="CENTER" height=40 WIDTH=70%>    
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
	<TD WIDTH=30% ALIGN=RIGHT nowrap>		
		<input type="BUTTON" name="send" value="(Y)Confirm" class=button ONCLICK="GO()">
		<input type="BUTTON" name="send" value="(N)Cancel" class=button onclick="clr()">
	</TD>
</TR>
	
</TABLE> 
</form>
  



</body>
</html>
<%
Sub StoreToSession()
	Dim CurrentRow
	tmpRec = Session("YECE03")
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
		tmpRec(CurrentPage, CurrentRow, 66) = request("ktaxm")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 47) = request("RELTOTMONEY")(CurrentRow)
	next 
	Session("YECE03") = tmpRec
	
End Sub
%> 

<script language=vbscript>
function BACKMAIN() 	
	open "../main.asp" , "_self"
end function   

function clr()
	open "<%=self%>.asp" , "_self"
end function 

function go()	
	<%=self%>.action="<%=self%>.upd.asp"  
	<%=self%>.submit()
end function 

function oktest(N)
	tp=<%=self%>.totalpage.value 
	cp=<%=self%>.CurrentPage.value 
	rc=<%=self%>.RecordInDB.value 
	open "empfile.show.asp?empautoid="& N , "_blank" , "top=10, left=10, width=550, scrollbars=yes" 
end function  

function zhuanmchg(index)
	F_EXRT = CDBL(<%=SELF%>.EXRT(INDEX).VALUE)
	REL_TOTAMT = CDBL(<%=SELF%>.RELTOTMONEY(INDEX).VALUE)	
	IF ISNUMERIC(<%=SELF%>.ZHUANM(INDEX).VALUE)=FALSE THEN 
		ALERT "請輸入數值!!" 
		<%=SELF%>.ZHUANM(INDEX).VALUE = <%=SELF%>.ZHUANMBAK(INDEX).VALUE 
		<%=SELF%>.XIANM(INDEX).VALUE = <%=SELF%>.XIANMBAK(INDEX).VALUE 
		'<%=SELF%>.ZHUANM(INDEX).SELECTED()
        <%=SELF%>.ZHUANM(INDEX).FOCUS()
		EXIT FUNCTION 
	END  IF 
	F_ZHUANM = CDBL(<%=SELF%>.ZHUANM(INDEX).VALUE)
	F_XIANM = CDBL(<%=SELF%>.XIANM(INDEX).VALUE)
	if  cdbl(F_ZHUANM)=CDBL(REL_TOTAMT)  then 
		<%=SELF%>.XIANM(INDEX).VALUE = CDBL(REL_TOTAMT) - CDBL(F_ZHUANM)		
	end if 
	IF  CDBL(<%=SELF%>.ZHUANM(INDEX).VALUE)+cdbl(<%=SELF%>.XIANM(INDEX).VALUE)  > CDBL(REL_TOTAMT)  THEN 
		ALERT "轉款金額輸入錯誤!!(大於實領金額)"
		<%=SELF%>.ZHUANM(INDEX).VALUE = <%=SELF%>.ZHUANMBAK(INDEX).VALUE 
		<%=SELF%>.XIANM(INDEX).VALUE = <%=SELF%>.XIANMBAK(INDEX).VALUE 			
		'<%=SELF%>.ZHUANM(INDEX).SELECTED()
		<%=SELF%>.ZHUANM(INDEX).FOCUS()
		EXIT FUNCTION 
	END  IF 
	<%=SELF%>.XIANM(INDEX).VALUE = CDBL(REL_TOTAMT) - CDBL(F_ZHUANM)
    '<%=SELF%>.XIANM(INDEX).FOCUS()        
    CODESTR01 = F_ZHUANM 
    CODESTR02 = CDBL(REL_TOTAMT) - CDBL(F_ZHUANM)
    open "<%=self%>.back.asp?ftype=ZXCHG&index="& index &"&CurrentPage="& <%=CurrentPage%> & _
		 "&CODESTR01=" & CODESTR01 & "&CODESTR02=" &	CODESTR02 , "Back"
    'PARENT.BEST.COLS="70%,30%"
END FUNCTION  

FUNCTION BBCODECHG(INDEX)
	codestr=<%=self%>.bbcode(index).value 
	daystr=<%=self%>.MMDAYS.value 	
	open "<%=self%>.back.asp?ftype=A&index="& index &"&CurrentPage="& <%=CurrentPage%> & _
		 "&days=" & daystr & "&code=" &	codestr , "Back" 		 
	'DATACHG(INDEX)	  
	 
	'PARENT.BEST.COLS="70%,30%"	 	
END FUNCTION 

FUNCTION JOBCHG(INDEX)
	codestr=<%=self%>.F1_JOB(index).value 
	daystr=<%=self%>.MMDAYS.value 
	open "<%=self%>.back.asp?ftype=B&index="& index &"&CurrentPage="& <%=CurrentPage%> & _
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
	CODESTR07 = <%=SELF%>.ktaxm(INDEX).VALUE
	CODESTR08 = <%=SELF%>.QITA(INDEX).VALUE
	CODESTR09 = <%=SELF%>.JX(INDEX).VALUE 
	CODESTR10 = <%=SELF%>.BH(INDEX).VALUE 
	CODESTR11 = 0
	
	yymmstr=<%=self%>.yymm.value  
	
	daystr=<%=self%>.MMDAYS.value  
	'ALERT CODESTR02
	'ALERT CODESTR03
	exrt = <%=self%>.f_exrt.value 
	open "<%=self%>.back.asp?ftype=CDATACHG&index="&index &"&CurrentPage="& <%=CurrentPage%> & _
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
		 "&CODESTR11="& CODESTR11 &"&days=" & daystr &"&exrt="& exrt &"&yymm="& yymmstr   , "Back"  
		 
	'PARENT.BEST.COLS="70%,30%"	 
	
END FUNCTION   

FUNCTION memochg(INDEX)
	yymmstr=<%=self%>.yymm.value 
 	memostr = escape(<%=self%>.memo(index).value)
 	open "<%=SELF%>.back.asp?ftype=memochk&index="&index &"&CurrentPage="& <%=CurrentPage%> & _ 
 		 "&yymm="& yymmstr &_
 		 "&memo=" & memostr  , "Back" 
    'parent.best.cols="70%,30%" 		 
END FUNCTION 

function view1(index) 	 
	yymmstr = <%=self%>.yymm.value 
	'alert yymmstr
	empidstr = <%=self%>.empid(index).value 
	idstr= <%=SELF%>.empautoid(INDEX).VALUE
	
	wt = (window.screen.width )*0.8
	ht = window.screen.availHeight*0.7
	tp = (window.screen.width )*0.02
	lt = (window.screen.availHeight)*0.02	
	
	OPEN "empworkB.Fore.asp?yymm=" & yymmstr &"&EMPID=" & empidstr , "_blank" , "top="& tp &", left="& lt &", width="& wt &",height="& ht &", scrollbars=yes"  
end function 
	
</script>

<%response.end %>
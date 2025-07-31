<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<!-- #include file = "../GetSQLServerConnection.fun" -->
<!-- #include file="../ADOINC.inc" -->
<%response.buffer=true%>
<%
'on error resume next
session.codepage="65001"
SELF = "YECE12" 

if session("netuser")="" then 
response.write "請重新登入"
response.end 
end if 

Set conn = GetSQLServerConnection()
Set rs = Server.CreateObject("ADODB.Recordset")
YYMM=REQUEST("YYMM")
whsno = trim(request("whsno"))
unitno = trim(request("unitno"))
groupid = trim(request("groupid"))
country = trim(request("country"))
job = trim(request("job"))
QUERYX = trim(request("empid1"))
outemp = request("outemp")
lastym = left(yymm,4) &  right("00" & cstr(right(yymm,2)-1) ,2 )
if right(yymm,2)="01"  then
	lastym = left(yymm,4)-1 &"12"
end if
rate = request("rate")
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
SQL=" SELECT * FROM YDBMCALE WHERE CONVERT(CHAR(6),DAT,112)  ='"& YYMM &"' AND     status <>'h1'  "
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
'if trim(QUERYX)<>"" then COUNTRY="" 
'本月應記薪天數
if country="VN" then 
	MMDAYS = CDBL(days)- cdbl(HHCNT)
else
	MMDAYS = CDBL(days)  
end if 	 

'if YYMM="202202" then 
'RESPONSE.WRITE  MMDAYS
'RESPONSE.END

sqlx="select * from VYFYEXRT where  yyyymm='"& yymm &"' and code='USD'  "
set rdsx= conn.execute(sqlx)
if rdsx.eof then 
	response.write "本月匯率尚未建檔!!" 
	response.end 
else	
	rate = rdsx("exrt")		
end if 	
rdsx.close : set rdsx=nothing 
'---------------------------------------------------------------------------------------- 
 

gTotalPage = 50
PageRec = 10    'number of records per page
TableRec = 91    'number of fields per record

NJYY = REQUEST("NJYY")

'response.write sql
'response.end 
allcnt = 0 

'年假計算日期 
if right(ENDdat,5)<="03/31" and yymm <= nowmonth  then 
	dat_s = cstr(left(yymm,4)-1) & "/04/01"	
else
	dat_s = cstr(left(yymm,4)) & "/04/01"	
end if  
'sql="exec sp_Calc_empsalary  '"& yymm &"', '"& whsno&"', '"&unitno&"', '"&groupid&"', '"&COUNTRY&"', '"&job&"', '"&QUERYX&"', '"&outemp&"' "  

if request("TotalPage") = "" or request("TotalPage") = "0" then
	CurrentPage = 1 
	
	sqlx="exec proc_CalcSalary '"& yymm &"', '"& whsno&"', '"&unitno&"', '"&groupid&"', '"&COUNTRY&"', '"&job&"', '"&QUERYX&"', '"&outemp&"' ,'"& dat_s &"'  " 
	conn.execute(sqlx) 
'response.write sqlx &"<br>" 
'response.end
	
	if outemp="D" then 
		sql="select a.*, isnull(nj.nj_amt,0) nj_amt , isnull(lz.lzxisu,0) lzxisu from "&_
				"( select * from Tab_CalcSalary  where isnull(outdate,'')<>'' and outdate >'"& calcdt &"' ) a "&_ 
				"left join ( select * from emptxamt where NJ_amt > 0  and Tyear ='"& njyy &"' )  nj on nj.empid = a.empid "&_
				"left join ( select * from fn_emplzxisu ( '')  ) lz on lz.empid=a.empid "
		sql=sql&"order by a.empid,lw desc, "&_
				"case when a.country='TW' then '1' else case when a.country='MA' then '2' else case when a.country='CN' then '3' else left(a.country,1) end end end ,   "&_
				"nindat "
	else
		sql="select a.*, isnull(nj.nj_amt,0) nj_amt  , isnull(lz.lzxisu,0) lzxisu  from "&_ 
				"(select * from Tab_CalcSalary   ) a "&_
				"left join ( select * from emptxamt where NJ_amt > 0  and Tyear ='"& njyy &"' )  nj on nj.empid = a.empid "&_
				"left join ( select * from fn_emplzxisu ( '')  ) lz on lz.empid=a.empid "
		sql=sql&"order by  a.empid,lw desc, "&_
						"case when a.country='TW' then '1' else case when a.country='MA' then '2' else case when a.country='CN' then '3' else left(a.country,1) end end end ,   "&_
						"nindat "
	end if 	
	
	rs.Open SQL, conn, 3, 3
	'response.write SQL&"<br>"
	'response.end
	IF NOT RS.EOF THEN		 
		rs.PageSize = PageRec
		RecordInDB = rs.recordcount
		TotalPage = rs.PageCount 
		gTotalPage = TotalPage
		'rs.movefirst
	END IF

	Redim tmpRec(gTotalPage, PageRec, TableRec)   'Array

	for i = 1 to TotalPage
	 for j = 1 to PageRec
		if not rs.EOF then
				rate = rs("rate")
				tmpRec(i, j, 0) = "no"
				tmpRec(i, j, 1) = trim(rs("empid"))
				tmpRec(i, j, 2) = trim(rs("empnam_cn"))
				tmpRec(i, j, 3) = trim(rs("empnam_vn"))
				tmpRec(i, j, 4) = rs("country")
				tmpRec(i, j, 5) = rs("nindat")  '到職日期
				tmpRec(i, j, 6) = rs("lj")
				tmpRec(i, j, 7) = rs("lw")
				tmpRec(i, j, 8) = "" 'rs("lu")
				tmpRec(i, j, 9)	=RS("lg")
				tmpRec(i, j, 10)=RS("lz")
				tmpRec(i, j, 11)=RS("lwstr")
				tmpRec(i, j, 12)="" 'RS("lustr")
				tmpRec(i, j, 13)=RS("lgstr")
				tmpRec(i, j, 14)=RS("lzstr")
				tmpRec(i, j, 15)=RS("ljstr")
				tmpRec(i, j, 16)= "" 'RS("lcstr")
				tmpRec(i, j, 17)=RS("autoid")
				IF RS("lz")="XX" THEN
					tmpRec(i, j, 18)=""
				ELSE
					tmpRec(i, j, 18)=RS("lz")
				END IF 
				
				tmpRec(i, j, 19)=trim(RS("OUTDATE")) 
				tmpRec(i, j, 20)=cdbl(RS("BB"))  '基本薪資				
				tmpRec(i, j, 21)=cdbl(RS("CV"))  '職務加給
				tmpRec(i, j, 22)=cdbl(RS("PHU")) 		'Y獎金 (海外津貼)
				tmpRec(i, j, 23)=cdbl(RS("NN"))		'Y獎金 (海外津貼)
				tmpRec(i, j, 24)=cdbl(RS("KT")) '技術加給(固定項目)
				tmpRec(i, j, 25)=cdbl(rs("MT"))  'MT   環境加給 , CN=年資(匯率)津貼  								
				tmpRec(i, j, 26)=cdbl(RS("TTKH"))  '其加
				tmpRec(i, j, 27)=cdbl(RS("QC"))  '全勤 				 				
				tmpRec(i, j, 28)=cdbl(RS("TNKH")) '其他收入		 
				if cdbl(rs("nj_amt"))>0 and rs("country")="VN" then 
					tmpRec(i, j, 28) = cdbl(tmpRec(i, j, 28)) + cdbl(rs("nj_amt"))
				end if 
				
				'response.write rs("eid")&"<br>"
				'response.write rs("reljxm_VND") &"<br>"
				'response.write rs("jx") &"<br>"
				'response.write "c="& rs("country")&"<br>"
				'response.write "rr="& rs("reljxm_VND")&"<br>"
				'response.write "ku="& rs("JXM_USD")&"<br>"
				'績效獎金 					
				if rs("eid")<>""  and   cdbl(rs("reljxm_VND"))=0   then  '已存在薪資檔
					tmpRec(i, j, 29)=rs("jx")  
				else
					if rs("country")="VN" then 
						tmpRec(i, j, 29)=rs("reljxm_VND")  
					else
						tmpRec(i, j, 29)=rs("JXM_USD")  
					end if 
				end if 		 
				'response.write rs("eid")&","& rs("reljxm_VND") &","& tmpRec(i, j, 29) &"<br>"
				tmpRec(i, j, 30)=RS("gtamt")   '工團費 				
				tmpRec(i, j, 31)=RS("bhp")   '本月保險基薪
				tmpRec(i, j, 32)=RS("bhtot")   '保險費 ( only  VN)      
				if yymm>="201001" then 
					tmpRec(i, j, 33)=RS("kh1")   '保險卡未歸還扣款     (KH1=月份)*(保險基薪*3%)  				
				else	
					tmpRec(i, j, 33) =0  
				end if 	
				
				'扣除其他	change by elin 20101129  201011起薪資 除營業外,其餘單位事故扣款接扣在績效獎金,此處不再扣款	 				
				if rs("eid")="" then 									
					if rs("lg")="A051" then 
						tmpRec(i, j, 34)=rs("dmsgKM")  '事故扣款 (系統計算) 
					else	
						if yymm>"201010" then tmpRec(i, j, 34)= 0  else tmpRec(i, j, 34)=rs("dmsgKM")
					end if 	
				else					
					if rs("lg")="A051" and (cdbl(rs("dmsgKM")) > cdbl(rs("QITA"))) then 
						tmpRec(i, j, 34)=rs("dmsgKM") 
					else
						tmpRec(i, j, 34)=rs("QITA")   '自行輸入的其他扣款  (存在empdsalary) 
					end if 	
				end if 
				
				
				
				'RESPONSE.WRITE tmpRec(i, j, 25) &"<br>"
				'RESPONSE.WRITE tmpRec(i, j, 30) &"<br>"
				
				if rs("country")="VN" then F_dm="VND" else F_dm="USD" 
				
				'if rs("EID")="" then 										
								
				'end if 	
				'tmpRec(i, j, 32)=rs("KTAXM") '所得稅
				
				BB=CDBL(tmpRec(i, j, 20))
				CV=CDBL(tmpRec(i, j, 21))
				PHU=CDBL(tmpRec(i, j, 22))				
				NN=CDBL(tmpRec(i, j, 23))
				KT=CDBL(tmpRec(i, j, 24))
				MT=CDBL(tmpRec(i, j, 25)) '匯率津貼
				TTKH=CDBL(tmpRec(i, j, 26))		'其加		
				QC=CDBL(tmpRec(i, j, 27))				
				TNKH=CDBL(tmpRec(i, j, 28))  '其他收入
				JX=CDBL(tmpRec(i, j, 29))  '績效獎金
				bt=CDBL(rs("btien"))  '補薪
				tien3=CDBL(rs("tien3"))  'TN年資
				'response.write "TNKH=" & TNKH &"<BR>"
				
				GTAMT=cdbl(RS("gtamt"))  '工團 				
				QITA=round(CDBL(tmpRec(i, j, 34)),0)  '- 其他 (=事故加其他)
				TOTBHP=round(CDBL(tmpRec(i, j, 32)),0)  '- 員工自付保險費 				
				if yymm>="201001" then 
					BH_NGT = 0 
				else
					BH_NGT= round( cdbl(tmpRec(i, j, 31))*0.045*cdbl(tmpRec(i, j, 33)) , 0)  '- 保險卡未歸還  (only VN) 
				end if 	
				
				
				'TOTY = All_CB 
							
				'KTAXM=CDBL(tmpRec(i, j, 32))  '-所得稅 
				
				 
				
				'總薪資
				TOTY=CDBL(BB)+CDBL(CV)+CDBL(PHU)+CDBL(NN)+CDBL(KT)+CDBL(MT)+CDBL(TTKH)+cdbl(tien3)
				All_CB =  cdbl(rs("all_cb"))
				'LT qui dinh de tru tien nghi phep theo sl ngay nghi
				All_CB1 =  CDBL(BB)+CDBL(CV)+CDBL(KT)+CDBL(MT)
				'response.write cdbl(rs("All_CB"))   &"<BR>"
				Hr_salary = cdbl(rs("Hr_salary"))
				
				tmpRec(i, j, 33)=  cdbl(rs("All_CB"))    '全薪 
				
				

				
				'****外國員工請假扣款方式******* 
				'1.境內休假  = 時薪  * 請假時數  
				'2.境外休假 = (全薪)/本月天數)  * (請假天數-12)   
				jiaAB_hr_i = cdbl(rs("jiaAB_hr_i")) '事病假 境內
				jiaAB_hr_w = cdbl(rs("jiaAB_hr_w")) '事病假 境外
				jiaI_hr_i = cdbl(rs("jiaI_hr_i"))   '返鄉休假 境內
				jiaI_hr_w = cdbl(rs("jiaI_hr_w"))   '返鄉休假 境外
				jiaF_hr_i = cdbl(rs("jiaF_hr_i"))  : if jiaF_hr_i="" then jiaF_hr_i=0  '產假 境內  
				jiaF_hr_w = cdbl(rs("jiaF_hr_w"))  : if jiaF_hr_w="" then jiaF_hr_w=0  '產假 境內   '產假 境外
				jiaJ_hr = cdbl(rs("jiaJ_hr"))   '留職停薪
				
				tmpRec(i, j, 35)=  jiaAB_hr_i+jiaAB_hr_w  '事病假
				tmpRec(i, j, 36)=  jiaI_hr_i+jiaI_hr_w  '返鄉休假
				tmpRec(i, j, 37)=  jiaF_hr_i+jiaF_hr_w  '產假 
				tmpRec(i, j, 38)=  jiaJ_hr  '留職停薪   			
				
				'response.write "事病="& tmpRec(i, j, 35)  &"<BR>"
				'response.write "產="& jiaF_hr_i  &"<BR>"
				'set rds=nothing  

				''員工工作天數(記薪天數) ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
				'1. 非新進員工 本月離職員工(不含1日) 從本月1日計算至離職日前一天				
		 		IF tmpRec(i, j, 19)="" or tmpRec(i, j, 19) >= ccdt THEN  '未離職或非本月離職
					wk_days = CDBL(MMDAYS)    '工作天數					 
					tmpRec(i, j, 39) = CDBL(wk_days)		  '工作天數 	
					
					'response.write "<br> 289 wk_days="&wk_days
					'response.write "<br>ccdt="&ccdt
					'response.write "<br>A1="&CDBL(wk_days)
					'response.write "<br> line 292 wk_days="&tmpRec(i, j, 39)
		 		ELSE	 
					SQL=" SELECT * FROM YDBMCALE WHERE CONVERT(CHAR(6),DAT,112)  ='"& YYMM &"' and CONVERT(CHAR(10),DAT,111)< '"& tmpRec(i, j, 19) &"'  AND  status<>'h1'  "
					Set rsTT = Server.CreateObject("ADODB.Recordset")
					RSTT.OPEN SQL, CONN, 3, 3
					IF NOT RSTT.EOF THEN
						T_HHCNT = CDBL(RSTT.RECORDCOUNT)
					ELSE
						T_HHCNT = 0
					END IF
					SET RSTT=NOTHING	 
					
			 		A1=DATEDIFF("D",CDATE(calcdt),CDATE(tmpRec(i, j, 19)) )  '從1日到離職日天數
					wk_days = CDBL(A1)	  '工作天數 
					if rs("country")	="VN" then 
						tmpRec(i, j, 39) = CDBL(wk_days)-T_HHCNT		  '越籍應扣假日天數
						'response.write "<br> line 308 wk_days="&tmpRec(i, j, 39)
					else
						tmpRec(i, j, 39) = CDBL(wk_days)		  '外籍以整月算,不扣假日
						'response.write "<br> line 323 wk_days="&tmpRec(i, j, 39)
					end if 
		 		END IF
				
				'加回產假時數避免重複扣款 20201010 
				'tmpRec(i, j, 39) = tmpRec(i, j, 39) +(cdbl(tmpRec(i, j, 37)/8))
				
				'2.老員工請假時數(產假,留職停薪,TA返鄉休假皆不計天數) 
				if rs("country")="VN" then 
					if cdbl(tmpRec(i, j, 37))>=208 then  '產假   本月休假時數超過 208(26*8)  H  
						wk_days = 0 '工作天數 = 0 			
						tmpRec(i, j, 39) = CDBL(wk_days)		   
					elseif 	cdbl(tmpRec(i, j, 37))>0 then 						
						tmpRec(i, j, 39) = round(CDBL(wk_days)-(cdbl(tmpRec(i, j, 37)) / 8 ) ,0)
					end if  
					'response.write "<br> 322 產假 "&tmpRec(i, j, 37) 
					'response.write "<br> line 323 wk_days="&tmpRec(i, j, 39)					
				else  '外籍員工休產假(境外) 	
					if cdbl(jiaF_hr_w)>=208 then  '(境外) 產假本月休假時數超過 208(26*8)  H  
						wk_days = 0 '工作天數 = 0 			
						tmpRec(i, j, 39) = CDBL(wk_days)		  '工作天數			 					 							
					end if   					
				end if 	
				
				if cdbl(tmpRec(i, j, 38))>=208 then  ' 留職停薪本月時數超過 208(26*8)  H   
					wk_days = 0 '工作天數 = 0 			
					tmpRec(i, j, 39) = CDBL(wk_days)		  '工作天數			 					 	
				elseif cdbl(tmpRec(i, j, 38))>=8 then  	
					wk_days = wk_days - (cdbl(tmpRec(i, j, 38))/8.0)
					tmpRec(i, j, 39) = CDBL(wk_days) 			 	
				end if   	
'				response.write 	"xxx=" &cdbl(tmpRec(i, j, 38)) &"days="& wk_days   &"<BR>"
				
				if rs("country")="TA" then   '泰籍員工境外返鄉休假應扣天數
					wk_days = wk_days - (cdbl(jiaI_hr_w)/8.0)
					tmpRec(i, j, 39) = CDBL(wk_days) 			 	
				end if 
				
			  if cdbl(jiaAB_hr_w)>208   then '境外請事病假
					wk_days = 0 '工作天數 = 0 			
					tmpRec(i, j, 39) = CDBL(wk_days)
				elseif 	cdbl(jiaAB_hr_w)>8 then 
					wk_days = wk_days - (CDBL(jiaAB_hr_w)/8.0)
					tmpRec(i, j, 39) = CDBL(wk_days)
					
					'response.write "<br> line 349 wk_days="&wk_days
					'response.write "<br>ji="&cdbl(jiaAB_hr_w)/8.0
					
				end if
				'response.write 	"xxx=" &cdbl(tmpRec(i, j, 39)) &"days="& wk_days   &"<BR>"
		 		'2.本月新進員工 從到職日計算到本月底
		 		IF CDATE(tmpRec(i, j, 5))>CDATE(calcdt) THEN
					SQL=" SELECT * FROM YDBMCALE WHERE CONVERT(CHAR(6),DAT,112)  ='"& YYMM &"' and CONVERT(CHAR(10),DAT,111)>='"& tmpRec(i, j, 5) &"'  AND   status<>'H1'  "
					Set rsTT = Server.CreateObject("ADODB.Recordset")
					RSTT.OPEN SQL, CONN, 3, 3
					IF NOT RSTT.EOF THEN
						F_HHCNT = CDBL(RSTT.RECORDCOUNT)
					ELSE
						F_HHCNT = 0
					END IF
					SET RSTT=NOTHING 
					
		 			iF tmpRec(i, j, 19)="" or tmpRec(i, j, 19) >= ccdt THEN  '本月到職本月仍在職 (到職日算至月底)			 			
			 			wk_days = DATEDIFF("D", CDATE(tmpRec(i, j, 5)), CDATE(ENDdat))+1  
						if rs("country")="VN" then 
							tmpRec(i, j, 39) = cdbl(wk_days) - F_HHCNT 
						else
							tmpRec(i, j, 39) = cdbl(wk_days) 
						end if 	
						
						'response.write "X5"&"<BR>"
			 		ELSE '本月到職本月離職	
						SQL=" SELECT * FROM YDBMCALE WHERE CONVERT(CHAR(6),DAT,112)  ='"& YYMM &"' and "&_
								" ( CONVERT(CHAR(10),DAT,111)>= '"& tmpRec(i, j, 5) &"' and CONVERT(CHAR(10),DAT,111)< '"& tmpRec(i, j, 19) &"'  ) AND  status<>'H1'  "
						'response.write sql &"<br>"
						Set rsTT = Server.CreateObject("ADODB.Recordset")
						RSTT.OPEN SQL, CONN, 3, 1
						IF NOT RSTT.EOF THEN
							F_HHCNT = CDBL(RSTT.RECORDCOUNT)
						ELSE
							F_HHCNT = 0
						END IF
						rstt.close : SET RSTT=NOTHING					
						
			 			wk_days =  DATEDIFF("D", CDATE(tmpRec(i, j, 5)), CDATE(tmpRec(i, j, 19)))   						
			 			tmpRec(i, j, 39) = cdbl(wk_days)-F_HHCNT		

						'response.write "line 386 ---"  & tmpRec(i, j, 39) &","&  wk_days & ","& F_HHCNT &"<BR>"						
			 			 						
			 		END IF
		 		ELSE
		 			tmpRec(i, j, 39) = tmpRec(i, j, 39)   '**********本月工作天數**********
					'response.write "X6"
		 		END IF				
							
				'加回產假時數避免重複扣款(天數) 20201010 
				tmpRec(i, j, 39) = tmpRec(i, j, 39) +(cdbl(tmpRec(i, j, 37)/8))
				
				'response.write "1=="& tmpRec(i, j, 39) &"<br>"
		 		'應領薪資合計
		 		'如為本月新進員工薪資OR本月離職: 總薪資/30 * 工作天數 	 (不足月時應領薪資)  				 
'				 if  yymm="202202"   and isnull(rs("outdate"))=false  then tmpRec(i, j, 39)=cdbl(tmpRec(i, j, 39))+4
				'response.write "2=="& tmpRec(i, j, 39) & ",,"& tmpRec(i, j, 19) & "<br>"
				'扣款計算  
				allKM = 0  
				'response.write "lin 369 ngay nghi=" & round(cdbl(tmpRec(i, j, 39)) -cdbl(tmpRec(i, j, 35))/8 -cdbl(tmpRec(i, j, 61))/8,2) &"<BR>"
				'So ngay nghi
				
				ngaynghi=cdbl(MMDAYS)-round(cdbl(tmpRec(i, j, 39)) -cdbl(tmpRec(i, j, 35))/8 -cdbl(tmpRec(i, j, 61))/8,2)
				'response.write ngaynghi
				'response.write "MMDAYS="&cdbl(MMDAYS) &"<br>"
				'response.write "wk_days="&round(cdbl(tmpRec(i, j, 39)) -cdbl(tmpRec(i, j, 35))/8 -cdbl(tmpRec(i, j, 61))/8,2) &"<br>"
				'response.end
				if rs("country")="VN" then   '日薪
					'day_salary = round( cdbl(All_CB) / ( cdbl(MMDAYS) ) ,0) '---VN計薪天數    '日薪
					'LT Neu ngay cong < 13 ngay
					'if round(cdbl(tmpRec(i, j, 39)) -cdbl(tmpRec(i, j, 35))/8 -cdbl(tmpRec(i, j, 61))/8,2) < 13  then
					'	day_salary = round( cdbl(All_CB) / 26 ,0) 'tat ca
					'else 'nguoc lai
					'	day_salary = round( cdbl(All_CB1) / 26 ,0) 'chi tinh theo cac cot nay All_CB1 =  CDBL(BB)+CDBL(CV)+CDBL(KT)+CDBL(MT)
					'end if
					if round(cdbl(ngaynghi)) >= 13 then
						day_salary = round( cdbl(All_CB) / 26 ,0) 'tat ca
					else 'nguoc lai
						day_salary = round( cdbl(All_CB1) / 26 ,0) 'chi tinh theo cac cot nay All_CB1 =  CDBL(BB)+CDBL(CV)+CDBL(KT)+CDBL(MT)
					end if
				else
					day_salary = round( cdbl(All_CB) / cdbl(MMDAYS) ,2)   '日薪  
				end if				
				
				'response.write "day_salary=" & day_salary &"<BR>"
				'response.write "lin 370 allCB=" & all_cb &"<BR>"
				
				if rs("country")="VN" then 		' ***-----VN --------------------------------------------------------------------------------------------------------------------------------------------------
					if cdbl(tmpRec(i, j, 39))=0 or ( cdbl(tmpRec(i, j, 39)) <3 and CDATE(tmpRec(i, j, 5))>CDATE(calcdt)) then   'VN工作天數<3天, 不計薪 (LA) 
						allKM  = allKM + cdbl(All_CB) 
						'response.write allKM&"<BR>"
					else 
						if cdbl(tmpRec(i, j, 35))>0 then 
							if cdbl(tmpRec(i, j, 35))/8 >= cdbl(tmpRec(i, j, 39))   then 
								allKM  =  cdbl(All_CB) 								
							else
								'if cdbl(tmpRec(i, j, 35))/8>= 13  then  '事病假超過或等於本月工作天數的1/2  
								'	allKM = cdbl(allKM) + round( cdbl(day_salary)* ((cdbl(tmpRec(i, j, 35))/8)) ,0)   '事病假天數*日薪
								'	'response.write rs("empid")&"-lin 147 allKM=" & day_salary &","&  allKM &"<BR>"
								'else
								'	'allKM = allKM + round( round(Hr_salary,0) * round(cdbl(tmpRec(i, j, 35)),1) , 0)
									allKM = allKM + round( round(cdbl(day_salary),0)/8 * round(cdbl(tmpRec(i, j, 35)),1) , 0)
								'	'response.write rs("empid")&"-lin 420 allKM=" & allKM &"<BR>"
								'end if 	
							end if 	
							'response.write "lin 382 allKM=" & allKM &"<BR>"
						end if 
						'response.write "lin 382 產假=" & tmpRec(i, j, 37) &"<BR>"
						if cdbl(tmpRec(i, j, 37))>0 then '201401 產假	
								if cdbl(tmpRec(i, j, 37))/8 >= cdbl(MMDAYS) then 
									allKM = All_CB    
									jiafCo = All_CB
								elseif cdbl(tmpRec(i, j, 37))/8>= cdbl(MMDAYS)/2  then  '產假超過或等於本月工作天數的1/2  
									allKM = cdbl(allKM) + round( cdbl(day_salary)* ((cdbl(tmpRec(i, j, 37))/8)) ,0)   '天數*日薪									
									jiafCo = cdbl(day_salary)* ((cdbl(tmpRec(i, j, 37))/8))
									'response.write  "line431 -1 產假應扣 = "  & jiafCo
								else
									'allKM = cdbl(allKM) + round( round(Hr_salary,0) * round(cdbl(tmpRec(i, j, 37)),2) , 0) '時薪*時數 事病假									
									allKM = cdbl(allKM) + round( round(Hr_salary,0) * round(cdbl(tmpRec(i, j, 37)),1) , 0) '時薪*時數 產假 20200917 elin								
									jiafCo = round(Hr_salary,0) * round(cdbl(tmpRec(i, j, 37)),1)
									'response.write  "line431 -2 產假應扣 = "  & jiafCo
								end if 	
								cj_days =  cdbl(tmpRec(i, j, 37))/8								
								'response.write  "line431 -2 產假應扣 = "  &  MMDAYS&"---"& cj_days  &","& jiafCo
						else 		
								cj_days =  0 
						end if  
						cj_days =  round(cdbl(tmpRec(i, j, 37))/8 ,1)
						'response.write  "line431 -2 產假應扣 = "  & jiafCo
						'response.write "445 allKM ="& allKM &"<br>"
						'response.write "工作天數="& tmpRec(i, j, 39) &"<br>"
						
						'如為本月新進員工薪資OR本月離職工作未滿13天: 總薪資/26 * 工作天數
						'舊員工本月離職, 工作天數13天(含)以上 : (BB+CV+PHU / 26 )* 工作天數  + ( NN+KT+MT+TTKH+QC 全薪 )  						
						'response.write "39="&cdbl(tmpRec(i, j, 39))
						if cdbl(tmpRec(i, j, 39))  < 13 then    '避免重複扣款要加回產假天數 201408
							allKM = cdbl(allKM) + round( cdbl(day_salary)* ((cdbl(MMDAYS))-(cdbl(tmpRec(i, j, 39))) ) ,0)  
							'response.write " "&rs("empid")&"日新="& day_salary &"-line 444 lizhi , allKM = "& cdbl(allKM) + round( cdbl(day_salary)* ((cdbl(MMDAYS))-cdbl(tmpRec(i, j, 39)) ) ,0) 
						elseif 	cdbl(tmpRec(i, j, 39))  >= 13    then   '避免重複扣款要加回產假天數 201408
							'allKM = cdbl(allKM) + round( ( cdbl(BB+CV+PHU)/cdbl(MMDAYS))*(cdbl(MMDAYS)-( cdbl(tmpRec(i, j, 39)) ) ) ,0)
							'Steven 2023/05/05
							allKM = cdbl(allKM) + round( cdbl(day_salary)*(cdbl(MMDAYS)-( cdbl(tmpRec(i, j, 39)) ) ) ,2)
							'response.write " "&rs("empid")&"-461 XXX2= " & cdbl(BB+CV+PHU) &","& cdbl(MMDAYS)&","& tmpRec(i, j, 39)&"," & round( ( cdbl(BB+CV+PHU)/cdbl(MMDAYS))*(cdbl(MMDAYS)-cdbl(tmpRec(i, j, 39)) ) ,0)
						end if 
						'response.write "allKM="&allKM
					end if 	
				else  ' 外籍
					if cdbl(tmpRec(i, j, 39)) <= 0 then  
					
						allKM  = allKM + cdbl(All_CB) 
					end if 
					if cdbl(jiaAB_hr_i)>0 then  						
						allKM  = allKM + round( round(Hr_salary,2)* round(cdbl(jiaAB_hr_i),1) ,0 ) '時薪*時數 事病假境內							
						'response.write "X2="& allKM &","&jiaAB_hr_i&","&Hr_salary&"<BR>"
					end if 
					if cdbl(jiaAB_hr_w)>0 then  
						allKM  = allKM + round( round(cdbl(day_salary),2) * round(cdbl(jiaAB_hr_w)/8.0,1) ,0 ) '日薪*天數 事病假境外
						'response.write "X2="& allKM &","&jiaAB_hr_w&","&day_salary&"<BR>"
					end if   
					' if rs("country")="TA" then 
						' if cdbl(jiaI_hr_w)>0 then 
							' allKM  = allKM + round( round( (BB+CV+PHU+KT)/cdbl(MMDAYS),2) * round(cdbl(jiaI_hr_w)/8.0,1) ,0)  'TA返鄉休假應扣
						' end if 	
					' end if 
					if cdbl(tmpRec(i, j, 39)) > 0 and cdbl(tmpRec(i, j, 39))<> cdbl(MMDAYS) and cdbl(jiaAB_hr_w)=0 then   
						if rs("country")="TA" then 	
							 allKM  = allKM + round( round( (BB+CV+PHU+KT)/cdbl(MMDAYS),2) * round(cdbl(jiaI_hr_w)/8.0,1) ,0)  
						else
							allKM  = allKM +  round( cdbl(day_salary) * (cdbl(MMDAYS)-cdbl(tmpRec(i, j, 39))) ,0)
						end if 	
					end if 
				end if 		
				'response.write " "&rs("empid")&"- , lin 474 allKM=" & allKM &"<BR>"
				Basic_allKM = allKM  ' 不含曠職應扣款
				
				'Steven 2022/04/04
				'曠職應扣曠職應扣 tru tien ko phep				
				if 	cdbl(rs("kzhour"))>0 then
					if rs("country")="VN" then 
						'allKM = allKM +  round( cdbl(rs("kzhour"))*round(Hr_salary,0) , 0) 
						'Steven 2023/05/05
						allKM = allKM +  round( cdbl(rs("kzhour"))*round(day_salary/8,2) , 2)
					elseif rs("country")="TA" then 
						if  yymm<="202107" then 
							allKM = allKM +  round( cdbl(rs("kzhour"))*(round(Hr_salary,3)*1.37) , 0) 	'202108
						else
							allKM = allKM +  round( cdbl(rs("kzhour"))*(round(Hr_salary,3)) , 0)   '扣時薪
						end if 
					end if 
				end if  
				
				if cdbl(allKM) > cdbl(all_cb) then  allKM = all_CB
				
				'response.write rs("empid")&"-曠職=" & cdbl(rs("kzhour")) &"H <BR>"  
				'RESPONSE.WRITE  "Hr_salary="& (Hr_salary)  &"<br>"
				' if rs("country")="VN" then  				
					' response.write "曠職扣=" & round( cdbl(rs("kzhour"))*round(Hr_salary,0) , 0)   &"<BR>"
				' else
				'	response.write "曠職扣=" &  round( cdbl(rs("kzhour"))*((Hr_salary)*1.37) , 0)   &"<BR>"
				' end if 
								
				
				'留職停薪 扣款  --  工作天數已經扣了(lin 315) ,此處不需再扣  20100701 change by elin 				
				'if cdbl(jiaJ_hr)>8 then '留職停薪
				'	allKM = allKM + round( (cdbl(jiaJ_hr)/8.0)*cdbl(day_salary) , 0)  '日薪*天數
		   	'end if 	   			
				'Response.write "L541="& allkm &"<BR>"

				
				'加班費 
				H12M=0 '晚班平日加班*1.7 201305 new add 
				B3_2=rs("b3_2") : if B3_2="" then B3_2=0  '周日夜班加班  201401 add   
				'response.write rs("h1") &","& rs("h2")&","& rs("h3")&","& rs("b3") &",假日夜班:"& rs("b3_2") &	", hr_salqry="& Hr_salary &"---<br>"  
				if rs("country")="VN" then 
					if yymm<="201304" then 
						H1M = round( cdbl(rs("allh1"))*1.5*round(Hr_salary,0)+0.01, 0) 
						H12M=0
					else
						H1M = round( cdbl(rs("h1"))*1.5*round(Hr_salary,0)+0.01, 0) 
						H12M= round( cdbl(rs("h12"))*1.7*round(Hr_salary,0)+0.01, 0)
						H1M=cdbl(H1M)+cdbl(H12M)
					end if
					H2M = round( cdbl(rs("h2"))*2*round(Hr_salary,0)+0.01, 0) 
					H3M = round( cdbl(rs("h3"))*3*round(Hr_salary,0)+0.01, 0) 
					'周日夜班加班 ( 21:00-06:00 9H*時新*0.2) 201401 add  
					'Steven 2020/12/21. Quan yeu cau vao proc_CalcSalary tinh lai b3 va b3_2
					'B3M = round( cdbl(rs("B3"))*0.3*round(Hr_salary,0)+cdbl(rs("B3_2"))*0.2*round(Hr_salary,0)+0.01 ,0)  
					B3M =  round(cdbl(rs("B3"))*0.3*round(Hr_salary,0),0)
					B3_2M =round(cdbl(rs("B3_2"))*0.6*round(Hr_salary,0),0)
					B4M =  round(cdbl(rs("B4"))*1.5*round(Hr_salary,0),0)
					
					'B3M = round( cdbl(rs("B3"))*0.3*round(Hr_salary,0)+0.01 ,0)  
					'response.write  ",原加班費:"& round(cdbl(H1M)+cdbl(H2M)+round(cdbl(rs("B3"))*0.3*round(Hr_salary,0)+0.01,0),0) &	", hr_salqry="& Hr_salary &"---<br>"  
					'response.write  ",假日夜班:"& round(cdbl(rs("b3_2"))*0.2*round(Hr_salary,0),0) &	", hr_salqry="& Hr_salary &"---<br>"  
						
				elseif rs("country")="TA" then 
					H1M = round( cdbl(rs("h1"))*1.5*round(Hr_salary,3), 0)
					'add by Steven 2012/02/01
					'nguoi yeu cau: L0197  --201901改為 *1.5 elin
					if YYMM < "201201" then
						H2M = round( cdbl(rs("h2"))*1*round(Hr_salary,3), 0) 
						H3M = round( cdbl(rs("h3"))*1*round(Hr_salary,3), 0)
					elseif YYMM <= "201812" then
						H2M = round( cdbl(rs("h2"))*1.37*round(Hr_salary,3), 0) 
						H3M = round( cdbl(rs("h3"))*1.37*round(Hr_salary,3), 0)
					else 
						H2M = round( cdbl(rs("h2"))*1.5*round(Hr_salary,3), 0) 
						H3M = round( cdbl(rs("h3"))*1.5*round(Hr_salary,3), 0)	
					end if
					B3M = 0
					B4M = 0
				elseif rs("country")="CN" then
					H1M = 0
					H2M = 0
					H3M = 0
					B3M = cdbl(rs("B3")) *5 
					B4M = 0
				else
					H1M = 0
					H2M = 0
					H3M = 0
					B3M = 0 
					B4M = 0
				end if 
				
				'全勤
				KQCM = 0 
				IF CDATE(tmpRec(i, j, 5)) >CDATE(calcdt) THEN   '新進人員
					KQCM = cdbl(qc)
					'response.write "1" &"<br>"
				ELSEif tmpRec(i, j, 39)< ( CDBL(MMDAYS)) THEN  '工作天數不足
					KQCM = cdbl(qc)
					'response.write "2" &"<br>"
				end if  
				'response.write "567 = " & allKM  &"<BR>"
				IF CDBL(RS("FORGET")+cdbl(rs("latefor")))>=3 and ( CDBL(RS("FORGET"))+cdbl(rs("latefor"))) < 6 THEN						
					KQCM = cdbl(qc)*0.5 
				elseIF 	( CDBL(RS("FORGET"))+cdbl(rs("latefor"))) >=6 THEN	
					KQCM = cdbl(qc) 
				end if 	 
				jiaAB_hr_c = rs("jiaAB_hr_c")  '不扣全勤時數  201407 事病假-不扣全勤時數 
				'response.write "jiaAB_hr_c line599 : "& tmpRec(i, j, 35)&" ----  , "& jiaAB_hr_c & " --  , "  & (cdbl(tmpRec(i, j, 35)) - cdbl(jiaAB_hr_c) ) & "<BR>"
				
				if ( (cdbl(tmpRec(i, j, 35))- cdbl(jiaAB_hr_c) ) > 0 and  ( cdbl(tmpRec(i, j, 35)) - cdbl(jiaAB_hr_c)) <= 8 ) or ( cdbl(rs("kzhour"))>0 and cdbl(rs("kzhour"))<= 8)then
					KQCM = cdbl(KQCM) + ( cdbl(qc)*0.5 )				
					'response.write "576-3 KQCM = " & KQCM  &"<br>" 					
				end if 			
				if ( cdbl(tmpRec(i, j, 35))- cdbl(jiaAB_hr_c)) >8 or cdbl(rs("kzhour"))>8 then 
					KQCM = cdbl(qc) 
'					response.write  "608 = = "& KQCM  &"<br/>"
				end if 
				if cdbl(kqcm) >  cdbl(QC) then 
					tmpRec(i, j, 27) = 0 
					qc = 0
				else
					tmpRec(i, j, 27) = cdbl(QC)-cdbl(KQCM)   
					qc =  cdbl(QC)-cdbl(KQCM)    
				end if 	
				if cdbl(cj_days)>= CDBL(MMDAYS) then   '產假一個月 ,無全勤
					tmpRec(i, j, 27) = 0 	
					qc = 0  
				end if 	
				'response.write "jzhour=" & tmpRec(i, j, 35) &"<BR>"
				'response.write "Ko全勤獎金="	& KQCM &"<BR>"		
				'response.write "全勤獎金="	& qc &"<BR>"
				'應領薪資		 
				tbtr = rs("tbtr")  '上月補款 
				mtax_qty = rs("person_qty")  '免稅人數
				mtax_Amt = rs("tot_Mtax")  '免稅額度
				
				'if rs("country")=
				if cdbl(tmpRec(i, j, 39)) <= 0 and (  cdbl(allKM ) >  cdbl(all_cb)+cdbl(qc)+cdbl(jx)+cdbl(tnkh)+cdbl(allJBM)+cdbl(tbtr)  ) then 
					'response.write "598-1"&"<BR>"
					allKM=cdbl(all_cb)+cdbl(qc)+cdbl(jx)+cdbl(tnkh)+cdbl(allJBM)+cdbl(tbtr)   
				else
					'response.write "598-2"&"<BR>"
					allKM=allKM 
				end if 	 					
				
'				response.write "634--XXXX-all_cb  = " & tmpRec(i, j, 33)  &"<BR>"
				all_salary = cdbl(all_cb)+cdbl(TNKH)+cdbl(H1M)+cdbl(H2M)+cdbl(H3M)+cdbl(b3M)+cdbl(b4M)+cdbl(b3_2M)+cdbl(jx)+cdbl(QC)+cdbl(tbtr)-cdbl(allKM)-cdbl(QITA)-cdbl(gtamt)-cdbl(TOTBHP)-cdbl(BH_NGT)  'add -cdbl(BH_NGT)20100429 
				'response.write "mtax_Amt=" & cdbl(mtax_Amt) &"<br>"				
				'response.write "ln 605 allKM=" & allKM &"<br>"
				'response.write "all_salary=" & all_salary &"<br>"
				'response.write "all_salary(VN)=" & all_salary* cdbl(rs("rate")) &"<br>"
				if rs("country")="VN" then 
					real_TOTAMT = all_salary-cdbl(mtax_Amt)
				else
					real_TOTAMT =(all_salary * cdbl(rs("rate")) )-cdbl(mtax_Amt)   '201201新增 ... 外籍也可申請免稅 elin 
				end if 	
				
				'if real_TOTAMT < 0 then real_TOTAMT = 0 
				if real_TOTAMT>0 then
				 	'----個人所得稅計算
			 		'RESPONSE.WRITE  ( CDBL(KTAXM)+CDBL(BH)+CDBL(QITA) )&"<br>"   
					'response.write "應扣金額=" & CDBL(BH)+CDBL(QITA)+cdbl(jiaAB) &"<BR>"
					'response.write "real_TOTAMT(USD)="& (cdbl(tmpRec(i, j, 35))- (CDBL(BH)+CDBL(QITA)+cdbl(jiaAB)) )  &"<BR>"
					'real_TOTAMT =  (cdbl(tmpRec(i, j, 35))- (CDBL(BH)+CDBL(QITA)+cdbl(jiaAB)) )  *cdbl(rs("exrt")) ' 實領金額 
					'response.write "real_TOTAMT="& real_TOTAMT &"<BR>" .   
					'totB=4000000   '   
					totB=9000000   '   201306, 900萬以上扣稅
					totB=11000000   '   202006, 1100萬以上扣稅
					F_TAX = 0					
					'Add by Steven , 2011/08/31
					if left(yymm,4)<="2008" then
						sql2="exec sp_calctax_HW_2008 '"& real_TOTAMT &"' "
						set ors=conn.execute(sql2) 
						F_tax = ors("tax")
						taxper = "0"
					elseif yymm>="201108" and yymm<"201112" then
						sql2="exec sp_calctax_201108 '"& real_TOTAMT &"' , '"& totB &"','"& rs("empid") &"' "						
						set ors=conn.execute(sql2) 
						F_tax = ors("tax")
						taxper = ors("taxper")					
					else
						sql2="exec sp_calctax_2010 '"& real_TOTAMT &"' , '"& totB &"','"& rs("empid") &"' "
						'response.write sql2
						set ors=conn.execute(sql2) 
						F_tax = ors("tax")
						taxper = ors("taxper")
					end if
					
					'if left(yymm,4)>"2008" then 
					'	sql2="exec sp_calctax_2010 '"& real_TOTAMT &"' , '"& totB &"','"& rs("empid") &"' "
					'	set ors=conn.execute(sql2) 
					'	F_tax = ors("tax")
					'	taxper = ors("taxper")
					'else
					'	sql2="exec sp_calctax_HW_2008 '"& real_TOTAMT &"' "
					'	set ors=conn.execute(sql2) 
					'	F_tax = ors("tax")
					'	taxper = "0"
					'end if  	
					
					set ors=nothing  
					'response.write  rs("empid") & " "& f_tax &"---" & sql2 &"<br>"
					 
					if rs("country")="VN" then 
						tmpRec(i, j,40) = round(cdbl(F_tax),0)
						KTAXM = round(cdbl(F_tax),0) 
					else
						tmpRec(i, j,40) = round(cdbl(F_tax) /cdbl(rs("rate")),0)
						KTAXM = round(cdbl(F_tax) /cdbl(rs("rate")),0) 
					end if 	
				else
					taxper = 0 
					KTAXM = 0 
					real_TOTAMT = 0 
				end if  
				 
				'實領薪資 = 應領薪資+績效+其他加給+其他收入-所得稅-醫療險自付額-其他扣除 				 
				final_salary = cdbl(all_salary) - cdbl(KTAXM)  
				if final_salary<0 then final_salary = 0 
				
				tmpRec(i, j,41) = rs("tbtr") '上月補款
				tmpRec(i, j,42) = final_salary 
				
				'離職補助金
				if tmpRec(i, j, 4)="VN" and rs("nindat")<"2009/01/01"  then
					inMonth = 0
					jishu = 0
					if  tmpRec(i, j, 19)="" then
						inMonth  = datediff("m", CDATE(tmpRec(i, j, 5)) , CDATE("2008/12/31") )
					else
						if tmpRec(i, j, 19)>"2009/01/01" then 
							end_dat2 = "2008/12/31" 
						end if 	
					 	inMonth  = datediff("m", CDATE(tmpRec(i, j, 5)) , CDATE(end_dat2) )
					end if
					
					'離職補助金 計算至2008/12/31
					'if  tmpRec(i, j, 19)="" and  tmpRec(i, j, 19)>"2008/12/31" then
					'	inMonth  = datediff("m", CDATE(tmpRec(i, j, 5)) , "2008/12/31" )
					'else
					' 	inMonth  = datediff("m", CDATE(tmpRec(i, j, 5)) , "2008/12/31" )
					'end if
					  
					'if inMonth >= 12 then
					' 	if inMonth mod 12 = 0 then
					' 		jishu=round(fix(inMonth/12 )*0.5,2) 
					' 		'response.write  "1" &"<BR>"
					' 	elseif inMonth mod 12 >= 6 then
					' 		jishu=round(fix(inMonth/12 )*0.5+0.25 ,2)
					' 		'response.write  "2"&"<BR>"
					' 	else
					' 		jishu=round(fix(inMonth/12 )*0.5 ,2)
					' 		'response.write  "3"&"<BR>"
					' 	end if
					'elseif inMonth >= 6 then
					'	if inMonth mod 6 = 0 then
					' 		jishu=round(fix(inMonth/6 )*0.5,2) 
					' 		'response.write  "1" &"<BR>" 
					' 	elseif inMonth mod 6 >= 3 then
					'		jishu=round(fix(inMonth/12 )*0.5+0.25 ,2)
					'	else
					' 		jishu=round(fix(inMonth/6 )*0.5 ,2)
					' 		'response.write  "3"&jishu & "<BR>"
					' 	end if
					'else 
 					' 	jishu=round(fix(inMonth/6 )*0.5+0.25 ,2)
					'end if 
  
					if cdbl(rs("lzxisu")) <=0 then   tmpRec(i, j, 43)  = 0  else tmpRec(i, j, 43) = rs("lzxisu")  '離職補助金系數 					
					jishu =cdbl( tmpRec(i, j, 43)  )
					
					IF jishu > 0 THEN
					 	tmpRec(i, j, 44) = ROUND( ( CDBL(all_cb)  * CDBL(jishu) ),0) '離職補助金
					ELSE
						tmpRec(i, j, 44) = 0
					END IF
				else
					tmpRec(i, j, 44) = 0
				end if  

				if trim(rs("studyjob"))<>""  then 				 
				 	tmpRec(i, j, 45) =  "Blue"
				 else	
				 	tmpRec(i, j, 45) =  "black"
				 end if 	 
				 
				 if trim(rs("outdate"))<>""   then 
				 	if cdbl(tmpRec(i, j, 43))>"0" and len(trim(rs("studyjob")))=0  then 
				 		tmpRec(i, j, 45) = "#ff0099" 
				 	elseif cdbl(tmpRec(i, j, 43))>"0"  and len(trim(rs("studyjob")))>0	then 
				 		tmpRec(i, j, 45) = "#3333ff" 
				 	end if 
				 else 
				 	tmpRec(i, j, 45) =  "black"
				 end if 		 

				if trim(rs("outdate"))="" then   '年資
					tmpRec(i, j, 46) = round(datediff("d",cdate(rs("indat")),cdate(ENDdat))/30,1)
				else
					tmpRec(i, j, 46) = round(datediff("d",cdate(rs("indat")),cdate(rs("outdate")))/30,1)
				end if 	 						
				
				tmpRec(i, j, 47)=h1m
				tmpRec(i, j, 48)=h2m
				tmpRec(i, j, 49)=h3m
				tmpRec(i, j, 50)=b3m
				tmpRec(i, j, 51)=cdbl(h1m)+cdbl(h2m)+cdbl(h3m)+cdbl(B3m)+cdbl(B4m)+cdbl(B3_2m) '總加班費
				allJBM = cdbl(h1m)+cdbl(h2m)+cdbl(h3m)+cdbl(B3m)+cdbl(B4m)+cdbl(B3_2m)
				tmpRec(i, j, 52)=rs("person_qty")
				tmpRec(i, j, 53)=rs("tot_mtax")
				tmpRec(i, j, 54)=allKM  
				
				
				if rs("country")<>"VN" then 
					tmpRec(i, j, 55) = 0 
				else	
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
						tmpRec(i, j, 55) = 0
					else					
						tmpRec(i, j, 55) = cdbl(datediff("m",cdate(cc_indat),cdate(tx_enddat) ))*8
					end if	 				
				end if 
				tmpRec(i, j, 56)=cdbl(all_cb)+cdbl(qc)+cdbl(jx)+cdbl(tnkh)+cdbl(allJBM)+cdbl(tbtr)
				if yymm<="201304" then 
					tmpRec(i, j, 57)=rs("allh1")
				else
					tmpRec(i, j, 57)=rs("h1")
				end if
				tmpRec(i, j, 58)=rs("h2")
				tmpRec(i, j, 59)=rs("h3")
				tmpRec(i, j, 60)=rs("b3")
				tmpRec(i, j, 61)=rs("kzhour") 
				'response.write  "kz" & tmpRec(i, j, 61)
				tmpRec(i, j, 62)=cdbl(rs("forget"))+cdbl(rs("latefor")) 
				tmpRec(i, j, 63)=BH_NGT  '保險卡未歸還扣款
				tmpRec(i, j, 64)=cdbl(rs("Hr_salary"))  '時薪  
				if F_tax = "0" then 
					tmpRec(i, j, 65)="0%"  '稅率   
				else
					tmpRec(i, j, 65)=taxper&"%"  '稅率   
				end if 	
				tmpRec(i, j, 66)=rs("sy_memo")  '備註說明
				tmpRec(i, j, 67)=Basic_allKM    ' 不含曠職應扣款  
				
				'機票補助款 
				nums = 0 				
				if rs("country")="TA" then 
					Set rs2 = Server.CreateObject("ADODB.Recordset")
					c_dat1 = left(yymm,4)-1&"1101"   
					c_dat2 = left(yymm,4)&"0430"
					sqlx="select count(*) h_nums, convert(char(10),min(dateup),111) as minDat, convert(char(10),max(dateup),111) maxdat , empid "&_
							 "from  empholiday WITH(NOLOCK) where jiatype='I' and empid='"& rs("empid") &"'  "&_
							 "and convert(char(8),dateup,112) between '"& c_dat1 &"' and '"& c_dat2&"' group by empid "
					Set rs2 = Server.CreateObject("ADODB.Recordset")
					rs2.open sqlx, conn, 1, 3		  					
					
					c_dat1 = left(yymm,4)&"0501" 
					c_dat2 = left(yymm,4)&"1031"
					sqlx="select count(*) h_nums, convert(char(10),min(dateup),111) as minDat, convert(char(10),max(dateup),111) maxdat , empid "&_
							 "from  empholiday WITH(NOLOCK) where jiatype='I' and empid='"& rs("empid") &"'  "&_
							 "and convert(char(8),dateup,112) between '"& c_dat1 &"' and '"& c_dat2&"' group by empid "
					Set rs3 = Server.CreateObject("ADODB.Recordset")
					rs3.open sqlx, conn, 1, 3		  
 
					if right(yymm,2)="04" then 
						if rs2.eof then 
							nums=1
							tmpRec(i, j, 68)=" 本年度第"&nums&"次未休假補助(機票)款"
						end if 
					elseif right(yymm,2)="10" then 
						if rs2.eof and rs3.eof then 
							nums=2 
							tmpRec(i, j, 68)=" 本年度第"&nums&"次未休假補助(機票)款" 
						elseif rs3.eof then 	
							nums=1
							tmpRec(i, j, 68)=" 本年度第"&nums&"次未休假補助(機票)款" 
							tmpRec(i, j, 71)="上次休假:" &rs2("mindat") &"~" & rs2("maxdat")
						else 
							nums=0 
						end if  	 
					end if 
					'response.write sqlx 	&"<BR>" 
				end if 
				if trim(rs("eid"))="" then 
					tmpRec(i, j, 66) = tmpRec(i, j, 66) & tmpRec(i, j, 68) 
				end if 	
				tmpRec(i, j, 69) = nums  
				
				if rs("country")="VN" then 
					tmpRec(i, j, 70) = 0  
				else '暫扣款(外國人半年內應暫扣總薪資*25% , 半年後仍在職全數歸還, 半年內去(離)職不發還 )					
					if trim(rs("outdate"))="" then 
						f_enddat=enddat
					else
						f_enddat=trim(rs("outdate"))
					end if 
					
					if datediff("d",cdate(rs("nindat")),cdate(f_enddat))<180 then 
					 	tmpRec(i, j, 70) = round(  all_salary *0.25 ,0) 					 						 
					end if	

					if datediff("m",rs("nindat"),ENDdat)>=6 and datediff("m",rs("nindat"),ENDdat)<7  then  
						tmpRec(i, j, 68)="本月歸還暫扣款"
						if trim(rs("eid"))="" then 
							tmpRec(i, j, 66) = tmpRec(i, j, 66) & " ,本月歸還暫扣款" 
						end if 	
					end if 	  
				end if   
				
				if nums>=1 and rs("country")="TA" then   '機票補助
					tmpRec(i, j, 72) = 300 
				else
					tmpRec(i, j, 72) = 0 
				end if 	  
				
				tmpRec(i, j, 73)=rs("jiaE")
				
				'201011 起只有營業才於薪資扣事故扣款
				if rs("dmsgKM")>"0" and rs("lg")="A051" then 
					if len(tmpRec(i, j, 66))=0 then 
						tmpRec(i, j, 66) = tmpRec(i, j, 66) & ",獎金扣款("&rs("dmsgKM")&")"& F_dm
					end if					
				end if 	
				
				if cdbl(rs("nj_amt"))>0 then 
					if trim(rs("eid"))="" then 
						tmpRec(i, j, 66) = njYY & "年假未修代金(tien thuong phep nam "&njyy&"): "& rs("nj_amt") &"VND"& chr(13) & tmpRec(i, j, 66)
					end if 	
				end if 	
				'曠職扣款 
				if 	cdbl(rs("kzhour"))>0 then
					if rs("country")="VN" then 
						tmpRec(i, j, 74) = round( cdbl(rs("kzhour"))*round(Hr_salary,0) , 0) 
					elseif rs("country")="TA" then 
						tmpRec(i, j, 74) =  round( cdbl(rs("kzhour"))*(round(Hr_salary,3)*1.37) , 0) 					
					end if 
				end if 
				
				tmpRec(i, j, 75)=rs("dmsgKM")  
				tmpRec(i, j, 76)=rs("btien")  
				if rs("country")="VN" and yymm>="202203" then   'change by elin 年資
					tmpRec(i, j, 76)=rs("tien3")  
				end if 
				tmpRec(i, j, 77)=rs("tien2")   'empmoney.money1  特別獎金
				if yymm<="201304" then tmpRec(i, j, 78)=0  else tmpRec(i, j, 78)=rs("h12")
				if yymm<="201304" then tmpRec(i, j, 79)=0  else tmpRec(i, j, 79)=H12M
				tmpRec(i, j, 88)=rs("b3_2")
				tmpRec(i, j, 89)=B3_2M
				
				tmpRec(i, j, 90)=rs("b4")
				tmpRec(i, j, 91)=B4M
				'Steven add lai cho dung ngay cong
				tmpRec(i, j, 39) = round(cdbl(tmpRec(i, j, 39)) -cdbl(tmpRec(i, j, 35))/8 -cdbl(tmpRec(i, j, 61))/8,2)
'				response.write "958--XXXX-all_cb  = " & tmpRec(i, j, 33)  &"<BR>"
				' response.write "全新="& tmpRec(i, j, 33) &"<BR>"
				' response.write "時新="& tmpRec(i, j, 64) &"<BR>"
				' response.write "日薪=" & day_salary &"<br>"
				' RESPONSE.WRITE  "本月計新天數="& MMDAYS & ", 工作天數:" &  tmpRec(i, j, 39)  &"<br>"
				' response.write "qc=" & qc &"<br>"  
				' RESPONSE.WRITE  "JX="  & jx &"<BR>"
				 'RESPONSE.WRITE  "加班費="&rs("h1")&"-"&rs("h12") & h1m & "," & h2m &","& h3M &","&  b3m &","& cdbl(h1m+h2M+h3M)   &"<br>" 
				' RESPONSE.WRITE  "allKM="& tmpRec(i, j, 54)  &"<br>"
				' response.write "扣qita=" & qita &"<br>"  				
				' response.write "tbtr=" & tbtr &"<br>"  
				' response.write "工團=" & gtamt &"<br>"  
				' response.write "保險=" & totbhp &"<br>" 
				' response.write "免稅=" & mtax_qty &", "&  mtax_Amt &"<br>" 				
				' response.write "1 TAX =" & KTAXM  &"<BR>"
				' response.write "2 final =" & rs("empid") & "----------  "&final_salary  &"<BR>"   
 
 
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
	Session("yece12B") = tmpRec
else
	TotalPage = cint(request("TotalPage"))
	StoreToSession()
	CurrentPage = cint(request("CurrentPage"))
	RecordInDB  = REQUEST("RecordInDB")
	tmpRec = Session("yece12B")
	COUNTRY = REQUEST("COUNTRY")

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
 
</head>
<body   topmargin="0" leftmargin="0"  marginwidth="0" marginheight="0" bgproperties="fixed" onkeydown='enterto()'  >
<form name="<%=self%>" method="post" action="<%=SELF%>.salary.asp">
<INPUT TYPE=HIDDEN NAME="UID" VALUE=<%=SESSION("NETUSER")%>>
<INPUT TYPE=hidden NAME=TotalPage VALUE="<%=TotalPage%>">
<INPUT TYPE=hidden NAME=CurrentPage VALUE="<%=CurrentPage%>">
<INPUT TYPE=hidden NAME=RecordInDB VALUE="<%=RecordInDB%>">
<INPUT TYPE=hidden NAME=PageRec VALUE="<%=PageRec%>">
<INPUT TYPE=hidden NAME=gTotalPage VALUE="<%=gTotalPage%>">
<INPUT TYPE=hidden NAME=YYMM VALUE="<%=YYMM%>">
<INPUT TYPE=hidden NAME=country VALUE="<%=country%>">
<INPUT TYPE=hidden NAME=rate VALUE="<%=rate%>">
<INPUT TYPE='hidden' NAME=calcdat VALUE="<%=calcdt%>">
<INPUT TYPE='hidden' NAME=enddat VALUE="<%=year(enddat)&"/"&right("00"&month(enddat),2)&"/"&right("00"&day(enddat),2)%>">
<INPUT TYPE='hidden' NAME=ccdt VALUE="<%=ccdt%>">
<INPUT TYPE='hidden' NAME=MMDAYS VALUE="<%=MMDAYS%>" title="本月工作天數">

<table width="600" border="0" cellspacing="0" cellpadding="0">
	<tr>
	<TD width=500 class="txt">
		<img border="0" src="../image/icon.gif" align="absmiddle">
		薪資計算( 員工薪資管理 )　計薪年月：<%=YYMM%>(<%=MMDAYS%>) 國籍：<%=COUNTRY%> , rate = <%=rate%></TD>
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
		<td align=center>
		<%if country<>"VN" then %>家境免稅<br>tien mien thue<%end if%>
		</td>
 		<TD align=center>工作天數<br>So Ngay<br>Lam viec</TD>
 		<TD align=center>
		<%if country="VN" then %>
			上月補款<br>T.B.T.R
		<%else%>	
			暫扣款
		<%end if%>	
		</TD> 
 		<TD align=center>總加班費<br>Phi Tang ca</TD> 		
 		<%if country="VN" then%><TD align=center colspan=2>離職補助<br>Tro cap<br>Thoi viec</TD> <%end if%>
		<%if country="VN" then%><TD align=center>保險未還<br>PHÁT SINH(T)</TD> <%end if%>
 		<td align=center>(-)扣時假</td>
 		<%if country="VN" then%><TD align=center>(-)不足月</TD><%end if%>
 		<TD align=center>(-)所得稅</TD> 						
 		<TD ALIGN=CENTER >實領工資</TD>
 		<%if country="VN" then%><TD bgcolor="#ffcc99" align=center><font color=blue>年假</font></TD><%end if%>
		<TD COLSPAN=7 ALIGN=CENTER bgcolor="#ccff99">加班(H)</TD>
 		<TD bgcolor="#ffcc99"></TD>
 		<TD bgcolor="#ffcc99"></TD> 		
 		<TD COLSPAN=2 ALIGN=CENTER bgcolor="#ffcccc">請假(H)</TD>
 	</TR>
 	<tr BGCOLOR="#e4e4e4"  HEIGHT=25 >
 		<TD align=center>基薪(BB)</TD>
 		<TD align=center>職加(CV)</TD>
 		<TD align=center>補助(Y)</TD>
		<td align=center><%if country="VN" then%>年資(TN)<%else%>補薪<br>(BL)<%end if%></td> 		
		<td align=center>語言(NN)</td>
 		<td align=center>技術(KT)</td>
 		<td align=center><%if country="VN" then%>環境(MT)<%else%>津貼(年資)<br>MT<%end if%></td>
 		<td align=center>其加(TTKH)</td>		
 		<td align=center>薪資合計</td>
 		<%if country="VN" then%><td align=center>全勤獎金</td><%end if%>
 		<td align=center><font color=red>績效獎金</font></td>
 		<td align=center>其他收入</td>
 		<TD align=center <%if country="VN" then%>colspan=2<%end if%>>應領薪資</TD>
		<%if country="VN" then%><TD align=center>家境免稅<br>tien mien thue</TD> <%end if%>
 		<td align=center>(-)其他</td>
 		<%if country="VN" then%><td align=center>(-)工團費</td><%end if%>
 		<td align=center>(-)保險費</td>
 		<td align=center>稅率(%)</td>
 		
 		<%if country="VN" then%><TD ALIGN=CENTER bgcolor="#ffcc99"><font color=Brown>尚有年假</font></TD><%end if%>
 		<TD ALIGN=CENTER bgcolor="#ccff99">平日(D)<br>1.5</TD>
		<TD ALIGN=CENTER bgcolor="#ccff99">平日(N)<br>1.7</TD>
 		<TD ALIGN=CENTER bgcolor="#ccff99">休息<br>2.0</TD>
 		<TD ALIGN=CENTER bgcolor="#ccff99">假日<br>3.0</TD>
 		<TD ALIGN=CENTER bgcolor="#ccff99">津貼<br>0.5</TD>
		<TD ALIGN=CENTER bgcolor="#ccff99">夜班<br>0.3</TD>
		<TD ALIGN=CENTER bgcolor="#ccff99">夜班<br>0.6</TD>
 		<TD ALIGN=CENTER bgcolor="#ffcc99"><font color=Brown>曠職</font></TD>
 		<TD ALIGN=CENTER bgcolor="#ffcc99"><font color=Brown>忘遲</font></TD> 		
 		<TD ALIGN=CENTER bgcolor="#ffcccc" >事病</TD>		
 		<TD ALIGN=CENTER bgcolor="#ffcccc" ><%if country<>"VN" then%>返鄉<%else%>&nbsp;<%end if%></TD>
 	</tr> 
 
<%response.flush%>
<%for CurrentRow = 1 to PageRec
		IF CurrentRow MOD 2 = 0 THEN
			WKCOLOR="LavenderBlush"
			'wkcolor="#ffffff"
		ELSE
			WKCOLOR="#DFEFFF"
			'wkcolor="#ffffff"
		END IF 
		'if tmpRec(CurrentPage, CurrentRow, 4)="VN" then xs=0 else xs=2
		'if tmpRec(CurrentPage, CurrentRow, 1) <> "" then
	%>
	<TR BGCOLOR=<%=WKCOLOR%> >
		<TD ROWSPAN=2 ALIGN=CENTER >
		<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then %><%=(CURRENTROW)+((CURRENTPAGE-1)*10)%><%END IF %>
		</TD>
 		<TD align=center >
 			<a href='vbscript:editmemo(<%=CURRENTROW-1%>)'>
 				<font color="<%=tmpRec(CurrentPage, CurrentRow, 45)%>"><u><%=tmpRec(CurrentPage, CurrentRow, 1)%></u></font>
 			</a>
 			<input type=hidden name=empid value="<%=tmpRec(CurrentPage, CurrentRow, 1)%>"  >
 			<input type=hidden name="empautoid" value="<%=tmpRec(CurrentPage, CurrentRow, 17)%>">
			<input type=hidden name="emp_ct" value="<%=tmpRec(CurrentPage, CurrentRow, 4)%>"  >
 		</TD>
 		<TD COLSPAN=3 >
 			<a href='vbscript:oktest(<%=CurrentRow-1%>)'> 				
 				<font class=txt8 color="<%=tmpRec(CurrentPage, CurrentRow, 45)%>">
					<%if whsno=""  then %>
					<%=tmpRec(CurrentPage, CurrentRow, 7)%> 
 					<%end if%>
					<%=left(tmpRec(CurrentPage, CurrentRow, 2)&tmpRec(CurrentPage, CurrentRow, 3),26)%>					
 				</font>
 			</a>
 		</TD>
 		<td  align=center><!--職等-->
 			<font color="<%=tmpRec(CurrentPage, CurrentRow, 45)%>"><%=left(tmpRec(CurrentPage, CurrentRow, 15),7)%></font>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
	 			<input type=hidden name=F1_JOB  class="readonly8" readonly   value="<%=trim(tmpRec(CurrentPage, CurrentRow, 6))%>"> 				
			<%else%>
				<input type=hidden name=F1_JOB >
			<%end if %>  		 			
 		</td>
 		<TD >
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %> 				 				
 				<INPUT NAME=HHMOENY  VALUE="<%=tmpRec(CurrentPage, CurrentRow, 64)%>" CLASS="readonly8r" readonly SIZE=7 STYLE="COLOR:#3366cc" >
 			<%ELSE%>
 				<INPUT NAME=HHMOENY TYPE=HIDDEN>
 			<%END IF %>
 		</TD>
 		<TD  ALIGN=CENTER ><FONT CLASS=TXT8><%=RIGHT(tmpRec(CurrentPage, CurrentRow, 5),8)%></FONT></TD>
 		<TD  ALIGN=CENTER ><FONT CLASS=TXT8 color=red><b><%=RIGHT(tmpRec(CurrentPage, CurrentRow, 19),8)%></b></FONT></TD>
		<!-- 201201顯示外籍免稅額度-->
		<td align="right"><%if country<>"VN" and tmpRec(CurrentPage, CurrentRow, 53)>"0" then%><%=formatnumber(tmpRec(CurrentPage, CurrentRow, 53),0)%><%end if%></td>
 		<TD >
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
	 			<INPUT NAME=WORKDAYS CLASS="readonly8r" READONLY  SIZE=7 VALUE="<%=tmpRec(CurrentPage, CurrentRow, 39)%>(<%=tmpRec(CurrentPage, CurrentRow, 46)%>)"    title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 工作天數">
	 			
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=WORKDAYS >
	 		<%END IF%>
 		</TD>
 		<TD>
 		<%if country="VN" then%>
			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
	 			<INPUT NAME=TBTR CLASS="readonly8r" READONLY  SIZE=7 VALUE="<%=formatnumber(tmpRec(CurrentPage, CurrentRow, 41),0)%>"  READONLY  title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 上月補款" >
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=TBTR >
	 		<%END IF%>
			<INPUT TYPE=HIDDEN NAME=DKM  VALUE="<%=formatnumber(tmpRec(CurrentPage, CurrentRow, 70),0)%>" >
		<%else%>	
			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
	 			<INPUT NAME=DKM CLASS="readonly8r" READONLY  SIZE=7 VALUE="<%=formatnumber(tmpRec(CurrentPage, CurrentRow, 70),0)%>"  READONLY  title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 上月補款" >
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=DKM >
	 		<%END IF%>			
			<INPUT type="hidden" NAME=TBTR  SIZE=7 VALUE="<%=formatnumber(tmpRec(CurrentPage, CurrentRow, 41),0)%>"  READONLY  title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 上月補款" >
		<%end if %>
 		</TD> 		
 		<TD >
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
	 			<INPUT NAME=TOTJB CLASS="readonly8r" READONLY  SIZE="<%if country="VN" then%>7<%else%>8<%end if%>"  VALUE="<%=formatnumber(tmpRec(CurrentPage, CurrentRow, 51),0)%>"   title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 總加班費">
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=TOTJB >
	 		<%END IF%>			
 		</TD>	 
		<%if country="VN" then%>
		<TD  ALIGN=RIGHT ><!--離職補助係數-->
			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
			<INPUT NAME=LZxisu CLASS="readonly8r" READONLY  SIZE=2 VALUE="<%=(tmpRec(CurrentPage, CurrentRow, 43))%>" >
			<%else%>	
			<INPUT TYPE=HIDDEN NAME=LZxisu CLASS="readonly8r" READONLY  SIZE=3 VALUE="0" > 
			<% end if%>
		</td>
 		<TD  ALIGN=RIGHT ><!--離職補助-->
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %> 								
	 			<INPUT NAME=LZBZJ CLASS="readonly8r" READONLY  SIZE=9 VALUE="<%=formatnumber(tmpRec(CurrentPage, CurrentRow, 44),0)%>" STYLE="TEXT-ALIGN:RIGHT;BACKGROUND-COLOR:LIGHTYELLOW;color:MediumBlue" READONLY   title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 離職補助金">
	 		<%ELSE%>				
	 			<INPUT TYPE=HIDDEN NAME=LZBZJ >	 			
	 		<%END IF%> 
		</TD> 
		<%else%>
			<INPUT TYPE=HIDDEN NAME=LZxisu CLASS="readonly8r" READONLY  SIZE=3 VALUE="0" > 
			<INPUT type="hidden" NAME=LZBZJ CLASS='INPUTBOX8' SIZE=8 VALUE="<%=formatnumber(tmpRec(CurrentPage, CurrentRow, 44),0)%>" STYLE="TEXT-ALIGN:RIGHT;BACKGROUND-COLOR:LIGHTYELLOW;color:MediumBlue" READONLY   title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 離職補助金">
		<%end if%>
		<%if country="VN" then%>
		<TD  ALIGN=RIGHT><!--保險卡未歸還-->
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %> 				
	 			<INPUT NAME=BH_NGT CLASS="readonly8r" READONLY  SIZE=8 VALUE="<%=formatnumber(tmpRec(CurrentPage, CurrentRow, 63),0)%>" STYLE="TEXT-ALIGN:RIGHT;BACKGROUND-COLOR:LIGHTYELLOW;color:MediumBlue" READONLY   title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 離職補助金">
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=BH_NGT >	 			
	 		<%END IF%> 
		</TD> 
		<%else%>
			<INPUT type="hidden" NAME=BH_NGT CLASS='INPUTBOX8' SIZE=8 VALUE="<%=formatnumber(tmpRec(CurrentPage, CurrentRow, 63),0)%>" STYLE="TEXT-ALIGN:RIGHT;BACKGROUND-COLOR:LIGHTYELLOW;color:MediumBlue" READONLY   title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 離職補助金">
		<%end if%>
 		<TD ><!--扣時假-->
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
	 			<INPUT NAME=TOTKJ CLASS="readonly8r" READONLY  SIZE=7 VALUE="<%=formatnumber(tmpRec(CurrentPage, CurrentRow, 54),0)%>" STYLE="TEXT-ALIGN:RIGHT;BACKGROUND-COLOR:LIGHTYELLOW;color:blue" READONLY   title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 扣時假">
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=TOTKJ value="0">
	 		<%END IF%>
			<INPUT type="hidden" NAME="old_TOTKJ" SIZE=7 VALUE="<%=formatnumber(tmpRec(CurrentPage, CurrentRow, 54),0)%>"  >
			<!--不含曠職的扣款-->
			<INPUT type="hidden" NAME="B_TOTKJ"   VALUE="<%=formatnumber(tmpRec(CurrentPage, CurrentRow, 67),0)%>"   >
 		</TD>
 		<%if country="VN" then%>
		<TD>			<!--不足月  200910 no use 以全數計入扣時假-->
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %> 
	 			<INPUT NAME=BZKM CLASS='INPUTBOX8' SIZE=7 VALUE="0" STYLE="TEXT-ALIGN:RIGHT;BACKGROUND-COLOR:LIGHTYELLOW" READONLY title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 不足月扣款"> 
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=BZKM   >
	 		<%END IF%>
 		</TD>
		<%else%>
			<INPUT  type="hidden"  NAME=BZKM CLASS='INPUTBOX8' SIZE=7 VALUE="0" STYLE="TEXT-ALIGN:RIGHT;BACKGROUND-COLOR:LIGHTYELLOW" READONLY title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 不足月扣款"> 
		<%end if%>
 		<TD >
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
 				<INPUT NAME=KTAXM CLASS="readonly8r" READONLY  SIZE=7 VALUE="<%=formatnumber(tmpRec(CurrentPage, CurrentRow, 40),0)%>" STYLE="TEXT-ALIGN:RIGHT;BACKGROUND-COLOR:LIGHTYELLOW;color:MediumBlue" READONLY   title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 所得稅">	 			
	 		<%ELSE%>	 			
	 			<INPUT TYPE=HIDDEN NAME=KTAXM >
	 		<%END IF%>
 		</TD> 
 		<TD>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
	 			<INPUT NAME=RELTOTMONEY CLASS="readonly8r" READONLY  size="8" VALUE="<%=formatnumber( tmpRec(CurrentPage, CurrentRow, 42),0)%>" STYLE="TEXT-ALIGN:RIGHT;BACKGROUND-COLOR:LIGHTYELLOW;;color:red" READONLY title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 實領工資">
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=RELTOTMONEY  >
	 		<%END IF%>
 		</TD> 		
		<%if country="VN" then%><!--年假-->
		<TD class=txt8 align=right><font color=blue><%=tmpRec(CurrentPage, CurrentRow, 55)%></font></TD>
		<%end if%>
 		<TD COLSPAN=7 align=center>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
 				<FONT CLASS=TXT8><u><div style="cursor:hand" onclick="view1(<%=currentrow-1%>)">
				<%=tmpRec(CurrentPage, CurrentRow, 1)%> 出勤紀錄</div></u></font>
 			<%END IF %>
 		</TD>
 		<TD ></TD>
 		<TD ></TD> 		
 		<TD COLSPAN=2 align=center nowrap>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
 				<FONT CLASS=TXT8><u><div style="cursor:hand" onclick="view2(<%=currentrow-1%>)">請假紀錄</div></u></font>				
 			<%END IF %>  			
 		</TD>
	</TR>
	<TR BGCOLOR=<%=WKCOLOR%> ><!------ line 2 ------------------------->
 		<TD ALIGN=RIGHT > 
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
	 			<INPUT NAME=BB CLASS="readonly8r" readonly   SIZE=7 VALUE="<%=formatnumber(tmpRec(CurrentPage, CurrentRow, 20),0)%>" title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 資本薪資">
	 		<%else%>
				<input type=hidden name=BB >
			<%end if %>		 
 		</TD>
 		<TD ALIGN=RIGHT><!--職加--> 
 		 	<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
 		 		<INPUT NAME=CV CLASS="readonly8r" readonly  SIZE=7 VALUE="<%=formatnumber(tmpRec(CurrentPage, CurrentRow, 21),0)%>"    title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 職務加給" > 		 		
 		 	<%else%>
				<input type=hidden name=CV >				 
			<%end if %>	 
 		</TD> 
 		<TD  ALIGN=RIGHT>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
	 			<INPUT NAME=PHU CLASS="readonly8r"   readonly  SIZE=7 VALUE="<%=formatnumber(tmpRec(CurrentPage, CurrentRow, 22),0)%>"   title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 補助(Y)" >
	 		<%ELSE%>
				<INPUT TYPE=HIDDEN NAME=PHU	>
			<%END IF%>
 		</TD>
		<TD  ALIGN=RIGHT><!--補薪-->
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
	 			<INPUT NAME=BTIRN CLASS="readonly8r" READONLY  SIZE=7 VALUE="<%=formatnumber(tmpRec(CurrentPage, CurrentRow, 76),0)%>"    title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 76))%> 補薪">
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=BTIRN >
	 		<%END IF%>
 		</TD>		
 		<TD  ALIGN=RIGHT>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
 				<INPUT NAME=NN CLASS="readonly8r" READONLY SIZE=7 VALUE="<%=formatnumber(tmpRec(CurrentPage, CurrentRow, 23),0)%>"    title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 語言加給" >
 			<%ELSE%>
				<INPUT TYPE=HIDDEN NAME=NN >
			<%END IF%>
 		</TD>
 		<TD  ALIGN=RIGHT>
	 		<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
	 			<INPUT NAME=KT CLASS="readonly8r" READONLY  SIZE=7 VALUE="<%=formatnumber(tmpRec(CurrentPage, CurrentRow, 24),0)%>"   title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 技術加給" >
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=KT >
	 		<%END IF%>
 		</TD>
 		<TD  ALIGN=RIGHT>
			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
	 			<INPUT NAME=MT CLASS="readonly8r" READONLY  SIZE=7 VALUE="<%=formatnumber(tmpRec(CurrentPage, CurrentRow, 25),0)%>" title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 環境加給">
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=MT >
	 		<%END IF%>
 		</TD>
 		<TD  ALIGN=RIGHT><!--其加-->
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
	 			<INPUT NAME=TTKH CLASS="readonly8r" READONLY  SIZE=7 VALUE="<%=formatnumber(tmpRec(CurrentPage, CurrentRow, 26),0)%>"    title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 其他加給">
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=TTKH >
	 		<%END IF%>
 		</TD> 
 		<TD ALIGN=RIGHT >
			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
				<INPUT NAME=totbsalary CLASS="readonly8r" READONLY  SIZE=8 VALUE="<%=formatnumber(tmpRec(CurrentPage, CurrentRow, 33),0)%>"   title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 薪資合計">
			<%else%>	
				<INPUT TYPE=HIDDEN NAME=totbsalary >
			<%end if %>
 		</TD>	
		<%if country="VN" then%>
 		<TD>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
	 			<INPUT NAME=QC CLASS="readonly8r" READONLY  SIZE=7 VALUE="<%=formatnumber(tmpRec(CurrentPage, CurrentRow, 27),0)%>"   title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 全勤">
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=QC >
	 		<%END IF%>
 		</TD> 
		<%else%>
			<INPUT  type="hidden"  NAME=QC CLASS="readonly8r" READONLY  SIZE=7 VALUE="<%=formatnumber(tmpRec(CurrentPage, CurrentRow, 27),0)%>"   title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 全勤">
		<%end if%>
 		<TD>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
	 			<INPUT NAME=JX SIZE=8 VALUE="<%=formatnumber(tmpRec(CurrentPage, CurrentRow, 29),0)%>" STYLE="TEXT-ALIGN:RIGHT" onblur="DATACHG(<%=CURRENTROW-1%>,1)"  title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 績效獎金" CLASS='inpt8red'  >
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=JX >
	 		<%END IF%>
 		</TD>  		 
 		<TD>
			<%newtnkh =cdbl(tmpRec(CurrentPage, CurrentRow, 28))%>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
	 			<INPUT NAME=TNKH SIZE=8 VALUE="<%=formatnumber(newtnkh,0)%>" STYLE="TEXT-ALIGN:RIGHT" onblur="DATACHG(<%=CURRENTROW-1%>,2)" title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 其他收入 + 特別獎金 : <%=trim(tmpRec(CurrentPage, CurrentRow, 77))%>" CLASS='inpt8blue' >
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=TNKH >
	 		<%END IF%>
			<INPUT TYPE=HIDDEN NAME=money1 value="<%=(tmpRec(CurrentPage, CurrentRow, 77))%>" >
 		</TD>
 		<TD <%if country="VN" then%>colspan=2<%end if%>>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
	 			<INPUT NAME=TOTM CLASS="readonly8r" READONLY  SIZE=<%if country="VN" then%>15<%else%>8<%end if%> VALUE="<%=formatnumber( tmpRec(CurrentPage, CurrentRow, 56),0)%>" STYLE="color:#cc0000"    title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 應領薪資">
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=TOTM >
	 		<%END IF%>
 		</TD>
		<%if country="VN" then%>
		<TD  ALIGN=RIGHT><!--免稅額-->
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %> 		
				<INPUT type="hidden" NAME=person_Qty CLASS='INPUTBOX8' SIZE=8 VALUE="<%=formatnumber(tmpRec(CurrentPage, CurrentRow, 52),0)%>"   >			
	 			<INPUT NAME=Notax_amt CLASS="readonly8r" READONLY  SIZE=8 VALUE="<%=formatnumber(tmpRec(CurrentPage, CurrentRow, 53),0)%>"   title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 離職補助金">
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=Notax_amt >	 			
				<INPUT TYPE=HIDDEN NAME=person_Qty >	
	 		<%END IF%> 
		</TD>		
		<%else%>
			<INPUT type="hidden" NAME=person_Qty CLASS="readonly8r" READONLY  SIZE=8 VALUE="<%=formatnumber(tmpRec(CurrentPage, CurrentRow, 52),0)%>"     >			
	 		<INPUT type="hidden" NAME=Notax_amt CLASS="readonly8r" READONLY  SIZE=8 VALUE="<%=formatnumber(tmpRec(CurrentPage, CurrentRow, 53),0)%>"    title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 離職補助金">
		<%end if%>
 		<TD>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
	 			<INPUT NAME=QITA CLASS='INPUTBOX8' SIZE=8 VALUE="<%=formatnumber(tmpRec(CurrentPage, CurrentRow, 34),0)%>" STYLE="TEXT-ALIGN:RIGHT" onblur="DATACHG(<%=CURRENTROW-1%>,3)" title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 扣除其他" >
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=QITA >
	 		<%END IF%>
 		</TD>
		<%if country="VN" then%>
 		<TD>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
	 			<INPUT NAME=GT CLASS="readonly8r" READONLY  SIZE=7 VALUE="<%=formatnumber(tmpRec(CurrentPage, CurrentRow, 30),0)%>" STYLE="TEXT-ALIGN:RIGHT"   title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 工團費"  readonly >
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=GT >
	 		<%END IF%>
 		</TD>
		<%else%>
			<INPUT  type="hidden"  NAME=GT CLASS="readonly8r" READONLY SIZE=7 VALUE="<%=formatnumber(tmpRec(CurrentPage, CurrentRow, 30),0)%>" STYLE="TEXT-ALIGN:RIGHT"   title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 工團費"  readonly >
		<%end if%>
 		<TD>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
	 			<INPUT NAME=BH CLASS='readonly8r' READONLY SIZE=7 VALUE="<%=formatnumber(tmpRec(CurrentPage, CurrentRow, 32),0)%>" STYLE="COLOR:#3366cc"  title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 保險費"  readonly  >
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=BH >
	 		<%END IF%>
 		</TD>
 		<TD><!--稅額-->			
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>				
				<INPUT NAME="taxper" CLASS="readonly8r" READONLY  SIZE=8 VALUE="<%=tmpRec(CurrentPage, CurrentRow, 65)%>" STYLE="TEXT-ALIGN:RIGHT;COLOR:#3366cc"  title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 稅額"  readonly  >
	 			<INPUT type="hidden" NAME=HS CLASS="readonly8r" READONLY  SIZE=8 VALUE="0" STYLE="TEXT-ALIGN:RIGHT" onblur="DATACHG(<%=CURRENTROW-1%>,4)" title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 伙食費" >
	 		<%ELSE%>
	 			<INPUT TYPE="HIDDEN" NAME="HS" >
				<INPUT TYPE="hidden" NAME="taxper" > 
	 		<%END IF%>
			<INPUT TYPE="hidden" value='0' NAME="empZAmt"  >
			<INPUT TYPE="hidden" value='0' NAME="hsf"  >
			<INPUT TYPE="hidden" value='0' NAME="govBB"  >
			<INPUT TYPE="hidden" value='0' NAME="butax"  >
			<INPUT TYPE="hidden" value='0' NAME="govall"  >
			<INPUT TYPE="hidden" value='0' NAME="after_tax"  >
			<INPUT TYPE="hidden" value='0' NAME="cty_saveamt"  >
			<INPUT TYPE="hidden" value='0' NAME="cty_buamt"  >
 		</TD> 
		
		<%if country="VN" then%>
 		<TD> <!--剩餘年假-->
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
	 			<INPUT NAME=YTX CLASS="readonly8r" READONLY  VALUE="<%=cdbl(tmpRec(CurrentPage, CurrentRow, 55))-cdbl(tmpRec(CurrentPage, CurrentRow, 73))%>"  SIZE=3 STYLE="color:red"  >
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=YTX >
	 		<%END IF%>
 		</TD> 	 
		<%else%>
			<INPUT type="hidden" NAME=YTX CLASS="readonly8r" readonly VALUE="<%=cdbl(tmpRec(CurrentPage, CurrentRow, 73))-cdbl(tmpRec(CurrentPage, CurrentRow, 74))%>"  SIZE=3  >
		<%end if%>
 		<TD>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %> 				
	 			<INPUT NAME=H1 CLASS="readonly8r" <%if country="VN" then%>readonly<%else%> onblur="h1chg(<%=CurrentRow-1%>)" <%end if%>  VALUE="<%=tmpRec(CurrentPage, CurrentRow, 57)%>"  SIZE=3 STYLE="color:MediumBlue"   >
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=H1 >
	 		<%END IF%>	 		
 		</TD>
		<TD>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %> 				
	 			<INPUT NAME=H12 CLASS="readonly8r"  readonly   VALUE="<%=tmpRec(CurrentPage, CurrentRow, 78)%>"  SIZE=3 STYLE="color:MediumBlue"   >
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=H12 >
	 		<%END IF%>	 		
 		</TD>
 		<TD>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
	 			<INPUT NAME=H2 CLASS="readonly8r" <%if country="VN" then%>readonly<%else%> onblur="h2chg(<%=CurrentRow-1%>)" <%end if%>  VALUE="<%=tmpRec(CurrentPage, CurrentRow, 58)%>"  SIZE=3 STYLE="color:MediumBlue"  >
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=H2 >
	 		<%END IF%>
 		</TD>
 		<TD>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
	 			<INPUT NAME=H3 CLASS="readonly8r"  <%if country="VN" then%>readonly<%else%> onblur="h3chg(<%=CurrentRow-1%>)" <%end if%> VALUE="<%=tmpRec(CurrentPage, CurrentRow, 59)%>"  SIZE=3 STYLE="color:MediumBlue"  >
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=H3 >
	 		<%END IF%>
 		</TD>
 		<TD>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
	 			<INPUT NAME=B4 CLASS="readonly8r"  <%if country="VN" then%>readonly<%else%> onblur="b3chg(<%=CurrentRow-1%>)" <%end if%>  VALUE="<%=tmpRec(CurrentPage, CurrentRow, 90)%>"  SIZE=3  STYLE="color:MediumBlue"  >
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=B4 >
	 		<%END IF%>			
 		</TD>
		<TD>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
	 			<INPUT NAME=B3 CLASS="readonly8r"  <%if country="VN" then%>readonly<%else%> onblur="b3chg(<%=CurrentRow-1%>)" <%end if%>  VALUE="<%=tmpRec(CurrentPage, CurrentRow, 60)%>"  SIZE=3  STYLE="color:MediumBlue"  >
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=B3 >
	 		<%END IF%>
			<INPUT TYPE=HIDDEN NAME=H1M  value="<%=tmpRec(CurrentPage, CurrentRow, 47)%>">
			<INPUT TYPE=HIDDEN NAME=H2M  value="<%=tmpRec(CurrentPage, CurrentRow, 48)%>">
			<INPUT TYPE=HIDDEN NAME=H3M  value="<%=tmpRec(CurrentPage, CurrentRow, 49)%>">
			<INPUT TYPE=HIDDEN NAME=B3M  value="<%=tmpRec(CurrentPage, CurrentRow, 50)%>">
			<INPUT TYPE=HIDDEN NAME=B4M  value="<%=tmpRec(CurrentPage, CurrentRow, 91)%>">
			<INPUT TYPE=HIDDEN NAME=B3_2M  value="<%=tmpRec(CurrentPage, CurrentRow, 89)%>">
 		</TD>
		<TD>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
	 			<INPUT NAME=B3_2 CLASS="readonly8r"  readonly  VALUE="<%=tmpRec(CurrentPage, CurrentRow, 88)%>"  SIZE=3  STYLE="color:MediumBlue"  >
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=B3_2 >
	 		<%END IF%>			
 		</TD>
 		<TD>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
	 			<INPUT NAME=KZHOUR CLASS="readonly8r"   <%if country="VN" then%>readonly<%else%> onblur="kzchg(<%=CurrentRow-1%>)" <%end if%>  VALUE="<%=tmpRec(CurrentPage, CurrentRow, 61)%>"  SIZE=3  STYLE="color:Brown"  >
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=KZHOUR >
	 		<%END IF%>
 		</TD>
 		<TD>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
	 			<INPUT NAME=Forget CLASS="readonly8r" readonly  VALUE="<%=tmpRec(CurrentPage, CurrentRow, 62)%>"  SIZE=3  STYLE="color:Brown"  >
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=Forget >
	 		<%END IF%>
 		</TD> 		
 		<TD>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
	 			<INPUT NAME=JIAA CLASS="readonly8r" readonly  VALUE="<%=tmpRec(CurrentPage, CurrentRow, 35)%>"  SIZE=3 STYLE="color:ForestGreen"  >
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=JIAA >
	 		<%END IF%>			
		</td>
		<Td>
		<%if country="VN" then%> 
			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
				<INPUT  NAME=JIAB CLASS="readonly8r" readonly  VALUE="0"  SIZE=3 STYLE="color:ForestGreen"  >			
			<%else%>
				<INPUT type="hidden" NAME=JIAI CLASS="readonly8r" readonly  VALUE="<%=tmpRec(CurrentPage, CurrentRow, 36)%>"  SIZE=3 STYLE="color:ForestGreen"  >
				<INPUT TYPE=HIDDEN NAME=JIAB >
			<%end if%>	
 		<%else%>	
			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
			<INPUT NAME=JIAI CLASS="readonly8r" readonly  VALUE="<%=tmpRec(CurrentPage, CurrentRow, 36)%>"  SIZE=3 STYLE="color:ForestGreen"  >		
			<%else%>
			<INPUT type="hidden" NAME=JIAI CLASS="readonly8r" readonly  VALUE="<%=tmpRec(CurrentPage, CurrentRow, 36)%>"  SIZE=3 STYLE="color:ForestGreen"  >
			<INPUT TYPE=HIDDEN NAME=JIAB >
			<%end if%>
		<%end if%> 
		</td>
	</TR>
	<%if country<>"VN" then %>	
	<tr BGCOLOR=<%=WKCOLOR%>>
		<Td></td>
		<Td colspan=10>			
			<%if tmpRec(CurrentPage, CurrentRow, 69)>"0" then %>
				<font color="blue">
				&nbsp;應補助:<%=tmpRec(CurrentPage, CurrentRow, 1)%>&nbsp;
				<%=tmpRec(CurrentPage, CurrentRow, 68)%>	&nbsp;			
				<%=tmpRec(CurrentPage, CurrentRow, 71)%>
				</font>
			<%else%>				
				<%if  tmpRec(CurrentPage, CurrentRow, 68)<>"" then %><font color="blue"><%=tmpRec(CurrentPage, CurrentRow, 68)%></font><%end if%>
			<%end if%>
		</td>
	</tr>
	<%else%>		
	<%end if%>
	<%next%>
</TABLE>  
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
			<input type="BUTTON" name="send" value="(Y)Confirm" class=button ONCLICK="GO()">
			<input type="BUTTON" name="send" value="(N)Cancel" class=button onclick="clr()">
		<%end if%>
	<%else%>	
		<input type="button" name="send" value="(Y)Confirm" class=button onclick="go()" >
		<input type="BUTTON" name="send" value="(N)Cancel" class=button  onclick="clr()">
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
	tmpRec = Session("yece12B")
	for CurrentRow = 1 to PageRec
		'tmpRec(CurrentPage, CurrentRow, 0) = request("op")(CurrentRow)
		' tmpRec(CurrentPage, CurrentRow, 6) = request("F1_JOB")(CurrentRow)
		' tmpRec(CurrentPage, CurrentRow, 19) = request("BBCODE")(CurrentRow)
		' tmpRec(CurrentPage, CurrentRow, 20) = request("BB")(CurrentRow)
		' tmpRec(CurrentPage, CurrentRow, 21) = request("CVCODE")(CurrentRow)
		' tmpRec(CurrentPage, CurrentRow, 22) = request("CV")(CurrentRow)
		' tmpRec(CurrentPage, CurrentRow, 23) = request("PHU")(CurrentRow)
		' tmpRec(CurrentPage, CurrentRow, 24) = request("KT")(CurrentRow)
		' tmpRec(CurrentPage, CurrentRow, 25) = request("TTKH")(CurrentRow)
		' tmpRec(CurrentPage, CurrentRow, 26) = request("MT")(CurrentRow)
		' tmpRec(CurrentPage, CurrentRow, 28) = request("TNKH")(CurrentRow)
		' tmpRec(CurrentPage, CurrentRow, 29) = request("JX")(CurrentRow)
		' tmpRec(CurrentPage, CurrentRow, 30) = request("BH")(CurrentRow)
		' tmpRec(CurrentPage, CurrentRow, 31) = request("QITA")(CurrentRow)
		' tmpRec(CurrentPage, CurrentRow, 32) = request("KTAXM")(CurrentRow)
		' tmpRec(CurrentPage, CurrentRow, 42) = request("B3")(CurrentRow)
		' tmpRec(CurrentPage, CurrentRow, 43) = request("B3M")(CurrentRow)
		' tmpRec(CurrentPage, CurrentRow, 44) = request("ZHUANM")(CurrentRow)
		' tmpRec(CurrentPage, CurrentRow, 45) = request("XIANM")(CurrentRow)
	next
	Session("yece12B") = tmpRec

End Sub
%>
 
<script language=vbscript>
function BACKMAIN()
	open "../main.asp" , "_self"
end function

function clr()
	open "<%=self%>.fore.asp" , "_self"
end function

function go()	
	<%=self%>.action="<%=SELF%>.upd.asp"
	<%=self%>.submit()
end function

function oktest(index)
		
	yymmstr = <%=self%>.YYMM.value
	empidstr = <%=self%>.empid(index).value 	 
	
	wt = (window.screen.width )*0.8
	ht = window.screen.availHeight*0.7
	tp = (window.screen.width )*0.02
	lt = (window.screen.availHeight)*0.02	
	
	OPEN "showsalary.asp?yymm=" & yymmstr &"&EMPID1=" & empidstr , "_blank" , "top="& tp &", left="& lt &", width="& wt &",height="& ht &", scrollbars=yes"  
	
	
end function  

function editmemo(index)
	tp=<%=self%>.totalpage.value
	cp=<%=self%>.CurrentPage.value
	rc=<%=self%>.RecordInDB.value
	YYMM = <%=self%>.YYMM.value
	open "<%=self%>.memo.asp?index="& index &"&currentpage=" & cp &"&yymm=" & yymm  , "_blank" , "top=10, left=10, width=450,height=450, scrollbars=yes"
end function  

function dkmchg(index)	
	CODESTR01 = <%=self%>.dkm(index).value
	YYMM = <%=self%>.YYMM.value
	open "<%=SELF%>.back.asp?ftype=dkmCHG&index="& index &"&CurrentPage="& <%=CurrentPage%> & _
		 "&CODESTR01=" & CODESTR01  &"&yymm=" & yymm  , "Back"
    'PARENT.BEST.COLS="70%,30%"
end function  

function h1chg(index)	
	if <%=self%>.h1(index).value<>"" then 
		if isnumeric(<%=self%>.h1(index).value)=false then 			
			alert "請輸入數字!!xin danh lai so"
			<%=self%>.h1(index).value="0"
			<%=self%>.h1(index).select()
			exit function 
		else				
			clcjbm(index)
		end if 
	end if 	
end function 
function h2chg(index)
	if <%=self%>.h2(index).value<>"" then 
		if isnumeric(<%=self%>.h2(index).value)=false then 			
			alert "請輸入數字!!xin danh lai so"
			<%=self%>.h2(index).value="0"
			<%=self%>.h2(index).select()
			exit function 
		else	
			clcjbm(index)
		end if 
	end if 	
end function 
function h3chg(index)
	if <%=self%>.h3(index).value<>"" then 
		if isnumeric(<%=self%>.h3(index).value)=false then 			
			alert "請輸入數字!!xin danh lai so"
			<%=self%>.h3(index).value="0"
			<%=self%>.h3(index).select()
			exit function 
		else	
			clcjbm(index)
		end if 
	end if 	
end function 
function b3chg(index)
	if <%=self%>.b3(index).value<>"" then 
		if isnumeric(<%=self%>.b3(index).value)=false then 			
			alert "請輸入數字!!xin danh lai so"
			<%=self%>.b3(index).value="0"
			<%=self%>.b3(index).select()
			exit function 
		else	
			clcjbm(index)
		end if 
	end if 	
end function 
function kzchg(index)
	if <%=self%>.kzhour(index).value<>"" then 
		if isnumeric(<%=self%>.kzhour(index).value)=false then 			
			alert "請輸入數字!!xin danh lai so"
			<%=self%>.kzhour(index).value="0"
			<%=self%>.kzhour(index).select()
			exit function 
		else	
			clckz(index)
		end if 
	end if 	
end function   

 

function clcjbm(index)
	f_ct = <%=self%>.emp_ct(index).value  '國籍
	f_hrm = <%=self%>.HHMOENY(index).value  '時薪
	f_h1=<%=self%>.h1(index).value
	f_h2=<%=self%>.h2(index).value
	f_h3=<%=self%>.h3(index).value
	f_b3=<%=self%>.b3(index).value
	f_allJBM = 0 
	
	if f_ct = "TA" then 
		if f_h1<>"" then <%=self%>.h1m(index).value = round( cdbl(f_h1)*1.37*round(f_hrm,3) , 0 ) 
		if f_h2<>"" then <%=self%>.h2m(index).value = round( cdbl(f_h2)*1*round(f_hrm,3) , 0 ) 
		if f_h3<>"" then <%=self%>.h3m(index).value = round( cdbl(f_h3)*1*round(f_hrm,3) , 0 ) 
		
		if f_h1<>"" and f_h2<>"" and f_h3<>""   then 
			f_allJBM = round( cdbl(f_h1)*1.37*round(f_hrm,3) , 0 ) + round( cdbl(f_h2)*1*round(f_hrm,3) , 0 )+round( cdbl(f_h3)*1*round(f_hrm,3) , 0 ) 
		end if 	
	elseif f_ct="CN" then 
		if f_b3<>"" then <%=self%>.B3M(index).value = round(cdbl(f_b3)*5,0)
		if f_b3<>"" then 
			f_allJBM = cdbl(f_b3)*5
		end if 
	end if 
	
	<%=self%>.TOTJB(index).value=f_allJBM  	 	
	
	todatachg(index)
end function 

function clckz(index)
	f_ct = <%=self%>.emp_ct(index).value  '國籍
	f_hrm = <%=self%>.HHMOENY(index).value  '時薪
	f_kz=<%=self%>.kzhour(index).value 
	f_bkj = <%=self%>.b_totkj(index).value  
	f_kj = 0 
	if f_ct = "TA" then 
		if f_kz<>"" then 
			f_kj = round(cdbl(f_kz)*1.37*round(f_hrm,3) , 0 )  
		end if 		 
	end if  
	<%=self%>.totkj(index).value=cdbl(f_bkj)+cdbl(f_kj)
	todatachg(index)
end function 
 
 

FUNCTION DATACHG(INDEX,a) 
	if a = 1 then  
		if isnumeric(<%=SELF%>.JX(INDEX).VALUE)=false then  '績效(+)
			alert "請輸入數字!!"
			<%=self%>.JX(index).value=0
			<%=self%>.JX(index).focus()
			<%=self%>.JX(index).select()
			exit FUNCTION
		else	
			<%=SELF%>.jx(INDEX).VALUE = formatnumber(<%=SELF%>.jx(INDEX).VALUE,0)	
			todatachg(index)
		end if
	end if  
	if a=2 then 
		if isnumeric(<%=SELF%>.TNKH(INDEX).VALUE)=false then  '其他收入
			alert "請輸入數字!!"
			<%=self%>.TNKH(index).value=0
			<%=self%>.TNKH(index).focus()
			<%=self%>.TNKH(index).select()
			exit FUNCTION
		else	
			<%=SELF%>.TNKH(INDEX).VALUE = formatnumber(<%=SELF%>.TNKH(INDEX).VALUE,0)
			todatachg(index)
		end if
	end if 
	if a = 3 then 
		if isnumeric(<%=SELF%>.QITA(INDEX).VALUE)=false then  '其他扣除額(-)
			alert "請輸入數字!!"
			<%=self%>.QITA(index).value=0
			<%=self%>.QITA(index).focus()
			<%=self%>.QITA(index).select()
			exit FUNCTION
		else	
			<%=SELF%>.qita(INDEX).VALUE = formatnumber(<%=SELF%>.qita(INDEX).VALUE,0)	
			todatachg(index) 
		end if
	end if 
END function  

function todatachg(index)   
	CODESTR01 = replace(<%=SELF%>.TNKH(INDEX).VALUE,",","")
	CODESTR02 = replace(<%=SELF%>.JX(INDEX).VALUE,",","") 
	CODESTR03 = replace(<%=SELF%>.QITA(INDEX).VALUE,",","")   
	CODESTR04 = replace(<%=SELF%>.TOTJB(INDEX).VALUE,",","")     '51  總加班
	CODESTR05 = replace(<%=SELF%>.h1(INDEX).VALUE,",","")   '57
	CODESTR06 = replace(<%=SELF%>.h2(INDEX).VALUE,",","")   '58
	CODESTR07 = replace(<%=SELF%>.h3(INDEX).VALUE,",","")   '59
	CODESTR08 = replace(<%=SELF%>.b3(INDEX).VALUE,",","")   '60
	CODESTR09 = replace(<%=SELF%>.kzhour(INDEX).VALUE,",","")   '61 
	CODESTR10 = replace(<%=SELF%>.TOTKJ(INDEX).VALUE,",","")   '扣時假  54
	CODESTR11 = replace(<%=SELF%>.h1M(INDEX).VALUE,",","")   '47
	CODESTR12 = replace(<%=SELF%>.h2M(INDEX).VALUE,",","")   '48
	CODESTR13 = replace(<%=SELF%>.h3M(INDEX).VALUE,",","")   '49
	CODESTR14 = replace(<%=SELF%>.b3M(INDEX).VALUE,",","")   '50
	CODESTR15 = replace(<%=SELF%>.b4M(INDEX).VALUE,",","")   '91
	'alert  CODESTR10
	rate = <%=self%>.rate.value
	daystr=<%=self%>.MMDAYS.value
	yymmstr=<%=self%>.yymm.value
	'ALERT CODESTR06
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
		 "&CODESTR12="& CODESTR12 &_		
		 "&CODESTR13="& CODESTR13 &_		
		 "&CODESTR14="& CODESTR14 &_
		 "&CODESTR15="& CODESTR15 &_
		 "&rate="& rate &_
		 "&yymm="& yymmstr &_
		 "&days=" & daystr , "Back"

PARENT.BEST.COLS="100%,0%"

END FUNCTION


FUNCTION memochg(INDEX)
	yymmstr=<%=self%>.yymm.value 
 	memostr = escape(<%=self%>.memo(index).value)
 	open "<%=SELF%>.back.asp?ftype=memochk&index="&index &"&CurrentPage="& <%=CurrentPage%> & _ 
 		 "&yymm="& yymmstr &_
 		 "&memo=" & memostr  , "Back" 
   ' parent.best.cols="70%,30%" 		 
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
	
	OPEN "../zzz/getempWorkTime.asp?yymm=" & yymmstr &"&EMPID=" & empidstr , "_blank" , "top="& tp &", left="& lt &", width="& wt &",height="& ht &", scrollbars=yes"  
end function 

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
	
function view2(index)	
	yymmstr = <%=self%>.yymm.value
	'alert yymmstr
	empidstr = <%=self%>.empid(index).value
	idstr= <%=SELF%>.empautoid(INDEX).VALUE
	OPEN "showholiday.asp?yymm=" & yymmstr &"&EMPID=" & empidstr , "_blank" , "top=10, left=10,  width=650, scrollbars=yes" 	
end function

</script>
<%response.end%>
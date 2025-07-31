<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<!-- #include file = "../GetSQLServerConnection.fun" -->
<!-- #include file="../ADOINC.inc" -->
<%
'on error resume next
session.codepage="65001"
SELF = "YECE1001"

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
MMDAYS = CDBL(days) 
'RESPONSE.WRITE  MMDAYS
'RESPONSE.END
'---------------------------------------------------------------------------------------- 

recalc  = request("recalc")
if recalc="Y" then 
	sql="delete empdsalary where yymm='"& YYMM &"' and isnull(country,'')='"& COUNTRY &"' and isnull(whsno,'') like '%"& whsno &"'"
	conn.execute(Sql)
end if 

sqlstr = "update empwork set kzhour=0 where yymm='"& YYMM &"'  and kzhour<0 "
conn.execute(sqlstr)

gTotalPage = 1
PageRec = 10    'number of records per page
TableRec = 52    'number of fields per record

sql="select case when isnull(m.qita,0)<>0 and a.whsno in ('PZ','BC') then m.qita else round( isnull(z.sukm,0)/cast(o.exrt as decimal(9,0)),2) end  sukm  ,  "&_
	"case when isnull(m.empid,'')='' then ROUND( isnull(k.TOTJXM,0)/ISNULL(O.EXRT,1),0)  else  m.jx end  TOTJXM, isnull( m. TNKH,0) TNKH ,  isnull(m.jx,0) jx,  "&_
	"ISNULL(M.BH,0)  as BH  , ISNULL(M.KTAXM,0) AS KTAXM,  ISNULL(M.MT,0) MTS,  "&_
	"0 as GT,  ISNULL(M.EMPID,'') AS EID2, ISNULL(M.TTKH,0) AS DTTKH, "&_
	"ISNULL(M.QITA,0) QITA, o.exrt, CASE WHEN ISNULL(M.EMPID,'')='' THEN 'N' ELSE 'Y' END AS EMPSTS,  "&_
    "ISNULL(M.ZHUANM,0) ZHUANM, ISNULL(M.XIANM,0) XIANM, "&_
	"isnull(n.forget,0) forget, isnull(n.h1,0) h1, isnull(n.h2,0) h2 , isnull(n.h3,0) h3, isnull(m.b3,0) b3 , isnull(m.b3m,0) b3m, "&_
	"isnull(JA.jiaa,0) jiaa,isnull(JB.jiab,0) jiab,isnull(n.kzhour,0) kzhour, isnull(n.latefor,0) latefor, isnull(m.memo,'') as salarymemo, "&_
	"isnull(m.dkm,0) dkm, isnull(m.acc,'') acc, isnull(l.empid,'') leid, "&_
	"isnull(l.bb,0) lbb, isnull(l.cv,0) lcv, isnull(l.phu,0) lphu, isnull(l.nn,0) lnn, "&_
	"isnull(l.kt,0) lkt, isnull(l.mt,0) lmt, isnull(l.ttkh,0) lttkh, isnull(l.qc,0) lqc , a.* from  "&_
	"( select * from  view_empfile where empid<>'pelin' and country not in ('VN', 'CN', 'TA') )  a  "&_
	"left join ( select * from bemps where  yymm='"& yymm &"' ) l on L.empid = a.empid "&_
	"left join ( select * from empdsalary where yymm='"& yymm &"' ) m on m.empid = a.empid   "&_
	"left join ( select * from VYFYMYJX where yymm='"& yymm &"' ) k on k.empid = a.empid and k.groupid = a.groupid  "&_
	"LEFT JOIN ( select empid empidN,  (sum(isnull(forget,0))) forget  , (sum(isnull(h1,0))) h1, (sum(isnull(h2,0))) h2, (sum(isnull(h3,0))) h3, (sum(isnull(b3,0))) b3 ,  "&_
 	"(sum(isnull(jiaa,0))) jiaa, (sum(isnull(jiab,0))) jiab, ( sum(isnull(toth,0))) toth , ( sum(isnull(kzhour,0))) kzhour , (sum(latefor)) latefor "&_
 	"from empwork   where yymm='"& YYMM &"' GROUP BY EMPID )  N ON N.empidN = A.EMPID  "&_
 	"left join  (  "&_
	"select jiaType as Ja , empid as EIDA, sum(hhour) as jiaa from empholiday where convert(char(6), dateup, 112)='"& yymm &"'  and jiatype='A'  group  by empid, jiatype   "&_
	")  JA on JA.EIDA = a.empid   "&_
	"left join ( "&_
	"select jiaType as Jb , empid as EIDB, sum(hhour) as jiaB from empholiday where convert(char(6), dateup, 112)='"& yymm &"'  and jiatype='B'  group  by empid, jiatype   "&_
	") JB on JB.eidb = a.empid   "&_
	"LEFT JOIN ( SELECT * FROM VYFYEXRT  WHERE  YYYYMM='"& yymm &"' and code='USD' ) O ON O.yyyymm='"& yymm &"'  "&_
	"left join ( select   empid, sum(sukm*exrt) sukm   from yfydsuco a , vyfyexrt b   where a.dm = isnull(b.code,'VND')  and a.ym = b.yyyymm and ym='"& yymm &"'  group by   empid   ) z on z.empid = a.empid  "&_
	"where CONVERT(CHAR(10), a.indat, 111)< '"& ccdt &"' and ( isnull(a.outdat,'')='' or a.outdat>'"& calcdt &"' )  "&_
	"and a.whsno like '"& whsno &"%' and a.unitno like '%"& unitno &"%' and a.groupid like '"& groupid &"%'  "&_
	"and a.COUNTRY like '"& COUNTRY  &"%' and A.job like '"& job &"%' and a.empid like '"& QUERYX &"%' "
	if outemp="D" then
		sql=sql&" and ( isnull(a.outdate,'')<>'' and  a.outdate>'"& calcdt &"' )  "
	end if

sql=sql&"order by a.whsno, a.empid   "


'response.write sql&"<br>"
'response.end
if request("TotalPage") = "" or request("TotalPage") = "0" then
	CurrentPage = 1
	rs.Open SQL, conn, 3, 3
	IF NOT RS.EOF THEN
		F_exrt = rs("exrt")
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
				tmpRec(i, j, 5) = rs("nindat")  '到職日期
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
				tmpRec(i, j, 19)=RS("lBB")
				tmpRec(i, j, 20)=RS("lbb")  '基本薪資
				tmpRec(i, j, 21)=RS("lcv")
				tmpRec(i, j, 22)=RS("lCV")  '職務加給
				tmpRec(i, j, 23)=RS("lPHU")		'Y獎金 (海外津貼)
				tmpRec(i, j, 24)=RS("lKT") '技術加給(固定項目)
				IF 	RS("EID2")="" THEN
					tmpRec(i, j, 25)=cdbl(RS("lTTKH"))  '其他加給(保險公司補助) only中國
				ELSE
					tmpRec(i, j, 25)=cdbl(RS("DTTKH"))
				END IF 				
				IF 	RS("EID2")="" THEN 
					tmpRec(i, j, 26)=cdbl(rs("lMT"))  'MT -- 匯率津貼
				else
					tmpRec(i, j, 26)=cdbl(rs("MTS"))  'MT -- 匯率津貼(已存在薪資檔)
				end if  							
				tmpRec(i, j, 27)=trim(RS("OUTDATE")) '離職日期
				tmpRec(i, j, 28)=ROUND(RS("TNKH"),0) '其他收入
				tmpRec(i, j, 29)=rs("totjxm")  '績效獎金
				'IF 	RS("EID2")="" THEN '保險費(-)
				'	tmpRec(i, j, 30) = 4
				'ELSE
				'	tmpRec(i, j, 30)=RS("BH")
				'END IF  
				tmpRec(i, j, 30) = 0 
				'RESPONSE.WRITE tmpRec(i, j, 25) &"<br>"
				'RESPONSE.WRITE tmpRec(i, j, 30) &"<br>"				
				
				tmpRec(i, j, 31)=rs("QITA")   '0  'RS("sukm") '其他扣除額
				tmpRec(i, j, 32)=rs("KTAXM") '所得稅

				'總薪資
				TOTY=CDBL(tmpRec(i, j, 20))+CDBL(tmpRec(i, j, 22))+CDBL(tmpRec(i, j, 23))+CDBL(tmpRec(i, j, 24))+CDBL(tmpRec(i, j, 26))

				BB=CDBL(tmpRec(i, j, 20))
				CV=CDBL(tmpRec(i, j, 22))
				PHU=CDBL(tmpRec(i, j, 23))
				KT=CDBL(tmpRec(i, j, 24))
				TTKH=CDBL(tmpRec(i, j, 25))
				MT=CDBL(tmpRec(i, j, 26)) '匯率津貼
				TNKH=CDBL(tmpRec(i, j, 28))
				JX=CDBL(tmpRec(i, j, 29))
				BH=CDBL(tmpRec(i, j, 30))
				QITA=CDBL(tmpRec(i, j, 31))
				KTAXM=CDBL(tmpRec(i, j, 32))
				
				tmpRec(i, j, 33)=TOTY  
				
				'員工工作天數(記薪天數)
				SQLX="SELECT EMPID, ISNULL(SUM(HHOUR),0) HHOUR  FROM EMPHOLIDAY WHERE EMPID='"& tmpRec(i, j, 1) &"' AND CONVERT(CHAR(6), DATEUP,112)='"& YYMM &"' "&_
		 				  "AND JIATYPE in ('A','B') GROUP BY EMPID "
		 		'response.write sqlx &"<BR>"
				Set rDs = Server.CreateObject("ADODB.Recordset")
		 		RDS.OPEN  SQLX, CONN, 3, 3
		 		IF NOT RDS.EOF THEN   '員工本月請事病假天數
		 			A4=FIX(CDBL(RDS("HHOUR"))/8)
		 		ELSE
		 			A4=0
		 		END IF 
				'response.write "A4=" & A4 &"<BR>"
				'工作天數---------------------------------------------------------------------
				'1.本月離職員工(不含1日) 從本月1日計算至離職日前一天
		 		IF tmpRec(i, j, 27)="" THEN  '未離職
		 			MWORKDAYS = CDBL(days)
		 			tmpRec(i, j, 34) = MWORKDAYS
		 			tmpRec(i, j, 34) = MWORKDAYS - A4
		 		ELSE
		 			IF  tmpRec(i, j, 27) >= ccdt THEN  '非本月離職
		 				MWORKDAYS = CDBL(days)
		 				tmpRec(i, j, 34) = MWORKDAYS
		 				tmpRec(i, j, 34) = MWORKDAYS - A4 
		 			ELSE '本月離職
			 			A1=DATEDIFF("D",CDATE(calcdt),CDATE(tmpRec(i, j, 27)) )  '從1日到離職日天數
			 			MWORKDAYS = CDBL(A1)+1 '**********本月工作天數**********
			 			tmpRec(i, j, 34)  = MWORKDAYS    '**********本月工作天數**********
			 		END IF
		 		END IF
		 		'RESPONSE.WRITE  MWORKDAYS
		 		'2.本月新進員工 從到職日計算到本月底
		 		IF CDATE(tmpRec(i, j, 5))>CDATE(calcdt) THEN
		 			iF tmpRec(i, j, 27)="" THEN  '本月到職本月仍在職
			 			A1= DATEDIFF("D", CDATE(tmpRec(i, j, 5)), CDATE(ENDdat))+1
			 			MWORKDAYS = cdbl(A1)
			 			tmpRec(i, j, 34) = MWORKDAYS
			 		ELSE '本月到職本月離職
			 			A1= DATEDIFF("D", CDATE(tmpRec(i, j, 5)), CDATE(tmpRec(i, j, 27)))
			 			MWORKDAYS = cdbl(A1)
			 			tmpRec(i, j, 34) = MWORKDAYS '**********本月工作天數**********
			 		END IF
		 		ELSE
		 			tmpRec(i, j, 34) = tmpRec(i, j, 34)  '**********本月工作天數**********
		 		END IF


		 		'應領薪資合計
		 		'如為本月新進員工薪資OR本月離職: 總薪資/30 * 工作天數 	 (不足月時應領薪資)
			 	if trim(tmpRec(i, j, 27))<>"" and tmpRec(i, j, 27) < calcdt then
			 		tmpRec(i, j, 35) = 0
			 	else	 '總薪資(BB+CV+PHU+KT+MT)/(本月天數)*工作天數+績效獎金+夜班津貼
			 		IF tmpRec(i, j, 34)< CDBL(days) THEN 
						tmpRec(i, j, 35) = ( ROUND((TOTY/cdbl(days))* CDBL(tmpRec(i, j, 34)),0)) + CDBL(tmpRec(i, j, 29))+ cdbl(rs("B3M"))+CDBL(TTKH)+CDBL(TNKH)-CDBL(QITA)
					ELSE 	
						tmpRec(i, j, 35) = CDBL(TOTY)+ CDBL(tmpRec(i, j, 29))+ cdbl(rs("B3M")) +CDBL(TTKH)+CDBL(TNKH)-CDBL(QITA)
					END IF
			 	end if
			 	'RESPONSE.WRITE tmpRec(i, j, 35) &"<br>"

			 	'超過800萬越幣應繳稅10% (總薪資+績效獎金>800VND) 

				F_TAX = 0 
				real_TOTAMT = cdbl(tmpRec(i, j, 35)) *cdbl(rs("exrt")) ' 實領金額
				' if cdbl(real_TOTAMT)>8000000 then 
					' if  cdbl(real_TOTAMT) <=20000000 then 
						' F_tax = ( cdbl(real_TOTAMT) - 8000000 ) * 0.1 
					' elseif cdbl(real_TOTAMT) > 20000000 and cdbl(cdbl(real_TOTAMT)) <= 50000000 then 
						' F_tax = ( (20000000-8000000)* 0.1 )+((cdbl(real_TOTAMT) - 20000000)*0.2)
					' elseif cdbl(real_TOTAMT) > 50000000 and cdbl(cdbl(real_TOTAMT)) <= 80000000 then 	
						' F_tax = ((20000000-8000000)* 0.1 )+ ( (50000000-20000000)* 0.2 ) + ((cdbl(real_TOTAMT) - 50000000)*0.3)
					' elseif cdbl(real_TOTAMT) > 80000000 then 
					  	' F_tax = ((20000000-8000000)* 0.1 )+ ( (50000000-20000000)* 0.2 ) + ( (80000000-50000000)* 0.3 ) + ((cdbl(real_TOTAMT) - 80000000)*0.4)
					' end if 
				' else
					' F_tax = 0  
				' end if 		
				totb = 4000000 
				if left(yymm,4)>"2008" then 
					sql2="exec sp_calctax '"& real_TOTAMT &"' ,'"& totb &"'"
					set ors=conn.execute(sql2) 
					F_tax = ors("tax")
				else
					sql2="exec sp_calctax_HW_2008 '"& real_TOTAMT &"' "
					set ors=conn.execute(sql2) 
					F_tax = ors("tax")
				end if  				
				set ors=nothing  	 				
				tmpRec(i, j, 32) = round(cdbl(F_tax) /cdbl(rs("exrt")),0)
				KTAXM = round(cdbl(F_tax)/cdbl(rs("exrt")),0)  

			 	''實領薪資 = 應領薪資+績效+其他加給+其他收入-所得稅-醫療險自付額-其他扣除-所得稅
		 		if tmpRec(i, j, 35) > 0 then
		 			tmpRec(i, j, 36) = CDBL(tmpRec(i, j, 35))-CDBL(KTAXM)-CDBL(BH)
		 		else
		 			tmpRec(i, j, 36) = 0
		 		end if		 		
		 		
		 		'response.write "35=" & tmpRec(i, j, 35) &"<BR>"
		 		'response.write "36=" & tmpRec(i, j, 36) &"<BR>"
		 		'response.write "KTAXM=" & KTAXM &"<BR>"
		 		
				 tmpRec(i, j, 37) = cdbl(TOTY)+CDBL(JX)+CDBL(TTKH)+CDBL(TNKH)+ cdbl(rs("B3M"))-CDBL(BH)-CDBL(QITA)-CDBL(KTAXM)
				 tmpRec(i, j, 38) =  round( CDBL(tmpRec(i, j, 20))/30/8,3)   '時薪(本薪/240)
				 'if rs("empid")<>"A0021" then 
				 	tmpRec(i, j, 39) = cdbl(tmpRec(i, j, 37)) - cdbl(tmpRec(i, j, 36)) '不足月扣款
				 'else
				 '	tmpRec(i, j, 39) = 0 
				 'end if 	
				 
				 'response.write "37=" & tmpRec(i, j, 37) &"<BR>" 
				 tmpRec(i, j, 40) = RS("EXRT")
				 tmpRec(i, j, 41) = cdbl(TOTY)+CDBL(JX)+CDBL(TTKH)+CDBL(TNKH)+ cdbl(rs("B3M")) 
				 tmpRec(i, j, 42) = rs("b3")
				 tmpRec(i, j, 43) = rs("b3m")
                 IF RS("EMPSTS")="N" THEN 
				    tmpRec(i, j, 44) = round(fix(tmpRec(i, j, 36)),0)
				    tmpRec(i, j, 45) = 0   '(round( round(cdbl(tmpRec(i, j, 36)),2) - round(fix(tmpRec(i, j, 36)),0),2)* cdbl(rs("exrt"))\1000)*1000
                 ELSE
                    tmpRec(i, j, 44) = RS("ZHUANM")
                    tmpRec(i, j, 45) = RS("XIANM")
                 END IF 
				 tmpRec(i, j, 46) = rs("exrt")				 
				 'response.write tmpRec(i, j, 44) &"<BR>"
				 
				 tmpRec(i, j, 47) = rs("salarymemo")
				 
				 if datediff("d",rs("nindat"),ENDdat)<180  and rs("empsts")="N" then  
				 	tmpRec(i, j, 48) = 0  'round( tmpRec(i, j, 36)*0.25 ,0) 
				 else				 
				 	tmpRec(i, j, 48) = rs("dkm")
				 end if	 
				 if trim(rs("acc"))="" then 
				 	if rs("whsno")="DN" then 
				 		tmpRec(i, j, 49) = "DN"
				 	else
				 		tmpRec(i, j, 49) = "VN"
				 	end if 	
				 else
				 	tmpRec(i, j, 49) = rs("acc") 
				 end if 	
 				'response.write tmpRec(i, j, 49) &"<BR>"  
 				if trim(rs("leid"))="" then 
 					tmpRec(i, j, 50) = "未建立基本薪資,請至C.E/1新增"
 					tmpRec(i, j, 51) ="red"
 				else
 					tmpRec(i, j, 50)=""
 					tmpRec(i, j, 51) ="black"
 				end if 		
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
	Session("YECE1001HW") = tmpRec
else
	TotalPage = cint(request("TotalPage"))
	'StoreToSession()
	CurrentPage = cint(request("CurrentPage"))
	RecordInDB  = REQUEST("RecordInDB")
	tmpRec = Session("YECE1001HW")
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
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
	function enterto()
		if window.event.keyCode = 13 then window.event.keyCode =9
	end function

	function f()
		'<%=self%>.PHU(0).focus()
		'<%=self%>.PHU(0).SELECT()
	end function

	function chgdata()
		<%=self%>.action="<%=self%>.salary.asp?totalpage=0"
		<%=self%>.submit
	end function
</SCRIPT>
</head>
<body   topmargin="0" leftmargin="0"  marginwidth="0" marginheight="0"  bgproperties="fixed"  onkeydown=enterto() >
<form name="<%=self%>" method="post" action="<%=SELF%>.salary.asp">
<INPUT TYPE=HIDDEN NAME="UID" VALUE=<%=SESSION("NETUSER")%>>
<INPUT TYPE=hidden NAME=TotalPage VALUE="<%=TotalPage%>">
<INPUT TYPE=hidden NAME=CurrentPage VALUE="<%=CurrentPage%>">
<INPUT TYPE=hidden NAME=RecordInDB VALUE="<%=RecordInDB%>">
<INPUT TYPE=hidden NAME=PageRec VALUE="<%=PageRec%>">
<INPUT TYPE=hidden NAME=gTotalPage VALUE="<%=gTotalPage%>">
<INPUT TYPE=hidden NAME=YYMM VALUE="<%=YYMM%>">
<INPUT TYPE=hidden NAME=MMDAYS VALUE="<%=MMDAYS%>">
<INPUT TYPE=hidden NAME=COUNTRY VALUE="<%=COUNTRY%>">
<INPUT TYPE=hidden NAME=exrt VALUE="<%=f_exrt%>">
<table width="600" border="0" cellspacing="0" cellpadding="0" class=txt>
	<tr>
	<TD width=600>
		<img border="0" src="../image/icon.gif" align="absmiddle">
		<%=session("pgname")%>　計薪年月：<%=YYMM%> exrt: <%=f_exrt%> , 國籍：<%=COUNTRY%>　</TD>
	</tr>
</table>
<hr size=0	style='border: 1px dotted #999999;' align=left width=500>

<TABLE  CLASS="FONT9" BORDER=0 cellspacing="0" cellpadding="2" >
	<TR HEIGHT=25 BGCOLOR="LightGrey"   >
 		<TD ROWSPAN=2 >項次</TD>
 		<td align=left>廠別</td>
 		<TD align=left>工號</TD>
 		<TD COLSPAN=3  >員工姓名(中,英,越)</TD> 		
 		<td align=center>到職日期</td>  
 		<TD align=center>離職日期</TD>
 		<TD align=center>工作天數</TD> 		
 		<td align=center colspan=3></td>
 	</TR>
 	<tr BGCOLOR="LightGrey"  HEIGHT=25 >
 		<TD align=left>立帳單位</TD> 		
 		<TD align=left>基本薪資</TD> 		
 		<TD align=center>職務加給</TD> 	     	  
 		<TD align=center>績效獎金</TD> 
 		<TD align=center>其他收入</TD> 		
 		<td align=center>應領薪資</td>
 		<td align=center>(-)其他</td>
 		<td align=center>(-)所得稅</td>
 		<td ALIGN=CENTER>(-)不足月</td>
 		<TD align=center>暫扣款</TD> 	
 		<TD ALIGN=CENTER >實領工資</TD> 
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
		<TD  ALIGN=left ><FONT CLASS=TXT8><%=tmpRec(CurrentPage, CurrentRow, 11)%></FONT></TD>
 		<TD  >
 			<a href='vbscript:editmemo(<%=currentRow-1%>)'>
 				<%=tmpRec(CurrentPage, CurrentRow, 1)%>
 			</a>
 			<input type=hidden name=empid value="<%=tmpRec(CurrentPage, CurrentRow, 1)%>">
 			<input type=hidden name="empautoid" value="<%=tmpRec(CurrentPage, CurrentRow, 17)%>">
 		</TD>
 		<TD COLSPAN=3>
 			<a href='vbscript:oktest(<%=tmpRec(CurrentPage, CurrentRow, 17)%>)'>
 				<font color="<%=tmpRec(CurrentPage, CurrentRow, 51)%>">
 				<%=tmpRec(CurrentPage, CurrentRow, 2)%>
 				<font class=txt8><%=tmpRec(CurrentPage, CurrentRow, 3)%></font></font>
 			</a>
 		</TD>
 		<TD  ALIGN=CENTER ><FONT CLASS=TXT8><%=RIGHT(tmpRec(CurrentPage, CurrentRow, 5),8)%></FONT></TD><!--到職日--> 
 		<TD ><FONT CLASS=TXT8><%=RIGHT(tmpRec(CurrentPage, CurrentRow, 27),8)%></FONT> <!--離職日--> 
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
	 			<INPUT TYPE=HIDDEN NAME=WORKDAYS CLASS='INPUTBOX8' READONLY  SIZE=7 VALUE="<%=tmpRec(CurrentPage, CurrentRow, 34)%>" STYLE="TEXT-ALIGN:RIGHT;BACKGROUND-COLOR:Gainsboro" onblur="DATACHG(<%=CURRENTROW-1%>)" title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 工作天數">
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=WORKDAYS >
	 		<%END IF%> 
	 		<INPUT TYPE=HIDDEN NAME=HHMOENY  VALUE="<%=tmpRec(CurrentPage, CurrentRow, 38)%>" CLASS='INPUTBOX8' SIZE=7 STYLE="TEXT-ALIGN:RIGHT;COLOR:#3366cc;BACKGROUND-COLOR:LIGHTYELLOW" > 
	 		<INPUT TYPE=HIDDEN NAME=B3 SIZE=7 VALUE="<%=tmpRec(CurrentPage, CurrentRow, 42)%>" CLASS="INPUTBOX8" STYLE="TEXT-ALIGN:RIGHT"  >
			<INPUT TYPE=HIDDEN NAME=B3M SIZE=7 VALUE="<%=tmpRec(CurrentPage, CurrentRow, 43)%>" CLASS="INPUTBOX8" READONLY  STYLE="TEXT-ALIGN:RIGHT" >
			<INPUT TYPE=HIDDEN NAME=TOTKJ CLASS='INPUTBOX8' SIZE=7 VALUE="0" STYLE="TEXT-ALIGN:RIGHT;BACKGROUND-COLOR:LIGHTYELLOW" READONLY   title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 扣時假">
			<INPUT TYPE=HIDDEN NAME=NN SIZE=7 VALUE="0" CLASS="INPUTBOX8" READONLY  STYLE="TEXT-ALIGN:RIGHT;BACKGROUND-COLOR:LIGHTYELLOW;color:blue" >
 		</TD> 
 		<TD align=center class=txt8><%=tmpRec(CurrentPage, CurrentRow, 34)%>  </TD> 		
 		<td colspan=3 class=txt8><font color=red><%=tmpRec(CurrentPage, CurrentRow, 50)%></font></td> 
	</TR>
	<TR BGCOLOR=<%=WKCOLOR%> ><!!---Line 2 ------------------------->
		<TD>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
		 		<select name=ACC  class="txt8" style="width:80" onchange="datachg(<%=currentrow-1%>)">
					<%SQL="SELECT * FROM basiccode WHERE FUNC='ACC' ORDER BY sys_value desc  "
					SET RST = CONN.EXECUTE(SQL)
					WHILE NOT RST.EOF
						if trim(tmpRec(CurrentPage, CurrentRow, 49))<>"" then 
							wstr =trim(tmpRec(CurrentPage, CurrentRow, 49))
						else 
							if trim(tmpRec(CurrentPage, CurrentRow, 7)) = "DN" then 
								wstr="DN" 
							else
								wstr="VN"
							end if 	
						end if 		
					%>
					<option value="<%=RST("sys_type")%>" <%IF RST("sys_type")=wstr THEN %> SELECTED <%END IF%> ><%=RST("sys_type")%>-<%=RST("sys_value")%></option>
					<%
					RST.MOVENEXT
					WEND
					%>
				</SELECT>
				<%SET RST=NOTHING %>
			<%else%>
				<input type=hidden name=ACC >
			<%end if %>
 		</TD>
 		<TD ALIGN=RIGHT >
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
	 			<INPUT NAME=BB CLASS='INPUTBOX8' SIZE=7 VALUE="<%=tmpRec(CurrentPage, CurrentRow, 20)%>" STYLE="TEXT-ALIGN:RIGHT;BACKGROUND-COLOR:LIGHTYELLOW"  READONLY title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 資本薪資">
	 		<%else%>
				<input type=hidden name=BB >
			<%end if %>
			<input type=hidden name=BBCODE value="<%=trim(tmpRec(CurrentPage, CurrentRow, 19))%>" >
 		</TD>
 		<TD ALIGN=RIGHT ><!--職等-->
 		 	<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
 		 		<INPUT NAME=CV CLASS='INPUTBOX8' SIZE=7 VALUE="<%=tmpRec(CurrentPage, CurrentRow, 22)%>" STYLE="TEXT-ALIGN:right;BACKGROUND-COLOR:LIGHTYELLOW"  READONLY  title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 職務加給" >
 		 		<input type=hidden name=CVCODE VALUE="<%=tmpRec(CurrentPage, CurrentRow, 21)%>" SIZE=3>
 		 	<%else%>
				<input type=hidden name=CV >
				<input type=hidden name=CVCODE >
			<%end if %> 
			<input type=hidden name=F1_JOB value="<%=trim(tmpRec(CurrentPage, CurrentRow, 6))%>">			
			<INPUT TYPE=HIDDEN NAME=PHU CLASS='INPUTBOX8' SIZE=7 VALUE="<%=tmpRec(CurrentPage, CurrentRow, 23)%>" STYLE="TEXT-ALIGN:RIGHT" onblur="DATACHG(<%=CURRENTROW-1%>)" title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 補助獎金(Y)" > 
			<INPUT TYPE=HIDDEN NAME=KT CLASS='INPUTBOX8' SIZE=7 VALUE="<%=tmpRec(CurrentPage, CurrentRow, 24)%>" STYLE="TEXT-ALIGN:RIGHT" onblur="DATACHG(<%=CURRENTROW-1%>)"  title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 績效獎金">
			<INPUT TYPE=HIDDEN NAME=TTKH CLASS='INPUTBOX8' SIZE=7 VALUE="<%=tmpRec(CurrentPage, CurrentRow, 25)%>" STYLE="TEXT-ALIGN:RIGHT" onblur="DATACHG(<%=CURRENTROW-1%>)" title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 其他加給">
			<INPUT TYPE=HIDDEN NAME=MT CLASS='INPUTBOX8' SIZE=7 VALUE="<%=tmpRec(CurrentPage, CurrentRow, 26)%>" STYLE="TEXT-ALIGN:RIGHT;COLOR:#3366cc" onblur="DATACHG(<%=CURRENTROW-1%>)"  title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 保險費"  >
 		</TD> 
 		<TD>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
	 			<INPUT NAME=JX CLASS='INPUTBOX8' SIZE=7 VALUE="<%=tmpRec(CurrentPage, CurrentRow, 29)%>" STYLE="TEXT-ALIGN:RIGHT" onblur="DATACHG(<%=CURRENTROW-1%>)"  title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 績效獎金">
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=JX >
	 		<%END IF%>
 		</TD> 
 		<TD>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
	 			<INPUT NAME=TNKH CLASS='INPUTBOX8' SIZE=7 VALUE="<%=tmpRec(CurrentPage, CurrentRow, 28)%>" STYLE="TEXT-ALIGN:RIGHT" onblur="DATACHG(<%=CURRENTROW-1%>)" title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 其他收入">
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=TNKH >
	 		<%END IF%>
 		</TD> 
 		<TD>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
	 			<INPUT NAME=TOTM CLASS='INPUTBOX8' SIZE=7 VALUE="<%=tmpRec(CurrentPage, CurrentRow, 41)%>" STYLE="TEXT-ALIGN:RIGHT;BACKGROUND-COLOR:LIGHTYELLOW;color:#cc0000"  READONLY  title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 應領薪資">
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=TOTM >
	 		<%END IF%>
 		</TD>
 		<TD>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
	 			<INPUT NAME=QITA CLASS='INPUTBOX8' SIZE=7 VALUE="<%=tmpRec(CurrentPage, CurrentRow, 31)%>" STYLE="TEXT-ALIGN:RIGHT;COLOR:RED" onblur="DATACHG(<%=CURRENTROW-1%>)" title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 扣除其他" >
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=QITA >
	 		<%END IF%>
 		</TD>
 		<TD>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
	 			<INPUT NAME=KTAXM CLASS='INPUTBOX8' SIZE=7 VALUE="<%=tmpRec(CurrentPage, CurrentRow, 32)%>" onblur="DATACHG(<%=CURRENTROW-1%>)"  STYLE="TEXT-ALIGN:RIGHT;COLOR:#3366cc" title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 所得稅" >
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=KTAXM >
	 		<%END IF%>
 		</TD>
 		<TD >
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
	 			<INPUT NAME=BZKM CLASS='INPUTBOX8' SIZE=7 VALUE="<%=formatnumber(tmpRec(CurrentPage, CurrentRow, 39),0)%>" STYLE="TEXT-ALIGN:RIGHT;BACKGROUND-COLOR:LIGHTYELLOW" READONLY title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 不足月扣款">
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=BZKM   >
	 		<%END IF%>
 		</TD>
 		<TD > 		
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
	 			<INPUT NAME=DKM CLASS='INPUTBOX8' SIZE=7 VALUE="<%=formatnumber(tmpRec(CurrentPage, CurrentRow, 48),0)%>" STYLE="TEXT-ALIGN:RIGHT;BACKGROUND-COLOR:LIGHTYELLOW" <%if session("Netuser")<>"PELIN" then %>READONLY <%else%> onchange="DATACHG(<%=CURRENTROW-1%>)" <%end if %>title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 暫扣款,代收代付">
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=DKM   >
	 		<%END IF%>	
	 		<INPUT TYPE=HIDDEN NAME=BH CLASS='INPUTBOX8' SIZE=7 VALUE="<%=tmpRec(CurrentPage, CurrentRow, 30)%>" STYLE="TEXT-ALIGN:RIGHT;COLOR:#3366cc"   title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 保險費"  >
	 		<INPUT type=hidden NAME=XIANM CLASS='INPUTBOX8' SIZE=10 VALUE="<%=formatnumber(tmpRec(CurrentPage, CurrentRow, 45),0)%>" STYLE="TEXT-ALIGN:RIGHT;BACKGROUND-COLOR:LIGHTYELLOW;color:blue"  readonly  > 
 			<INPUT type=hidden NAME=ZHUANM CLASS='INPUTBOX8' SIZE=10 VALUE="<%=formatnumber(tmpRec(CurrentPage, CurrentRow, 44),0)%>" STYLE="TEXT-ALIGN:RIGHT"  onblur="zhuanmchg(<%=CURRENTROW-1%>)"  >
 		</TD>   
 		<TD>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
	 			<INPUT NAME=RELTOTMONEY CLASS='INPUTBOX8' VALUE="<%=( tmpRec(CurrentPage, CurrentRow, 36))%>" SIZE=7  STYLE="TEXT-ALIGN:RIGHT;BACKGROUND-COLOR:LIGHTYELLOW;;color:#cc0000" READONLY title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 實領工資">
	 		<%ELSE%>
	 			<INPUT TYPE=HIDDEN NAME=RELTOTMONEY  >
	 		<%END IF%>
	 		<INPUT type=hidden NAME=H1 CLASS='INPUTBOX8' VALUE="0"  SIZE=4  READONLY >
	 		<INPUT type=hidden NAME=H2 CLASS='INPUTBOX8' VALUE="0"  SIZE=4  READONLY >
	 		<INPUT type=hidden NAME=H3 CLASS='INPUTBOX8' VALUE="0"  SIZE=4  READONLY >
	 		<INPUT type=hidden NAME=KZHOUR CLASS='INPUTBOX8' VALUE="0"  SIZE=4   READONLY >
	 		<INPUT type=hidden NAME=Forget CLASS='INPUTBOX8' VALUE="0"  SIZE=4  READONLY >
	 		<INPUT type=hidden NAME=JIAA CLASS='INPUTBOX8' VALUE="0"  SIZE=4 READONLY >
	 		<INPUT type=hidden NAME=JIAB CLASS='INPUTBOX8' VALUE="0"  SIZE=4 READONLY >
	 		<INPUT TYPE=HIDDEN NAME=exrt value="<%=tmpRec(CurrentPage, CurrentRow, 46)%>"   >
			<INPUT TYPE=HIDDEN NAME=XIANMBAK  VALUE="<%=tmpRec(CurrentPage, CurrentRow, 45)%>" >
			<INPUT TYPE=HIDDEN NAME=ZHUANMBAK  VALUE="<%=tmpRec(CurrentPage, CurrentRow, 44)%>" >
 		</TD>
 
	</TR>
	<%next%>
</TABLE>

<input type=hidden name=empid>
<input type=hidden name="empautoid"  >
<INPUT NAME=HHMOENY TYPE=HIDDEN>
<INPUT TYPE=HIDDEN NAME=WORKDAYS VALUE="0">
<INPUT TYPE=HIDDEN NAME=TNKH VALUE="0">
<INPUT TYPE=HIDDEN NAME=TOTKJ VALUE="0">
<INPUT TYPE=HIDDEN NAME=QITA VALUE="0">
<INPUT TYPE=HIDDEN NAME=BZKM   >
<input type=hidden name=BBCODE VALUE="0">
<input type=hidden name=BB>
<input type=hidden name=F1_JOB >
<input type=hidden name=CV VALUE="0">
<input type=hidden name=CVCODE VALUE="0">
<input type=hidden name=PHU VALUE="0">
<input type=hidden name=KT VALUE="0">
<INPUT TYPE=HIDDEN NAME=JX VALUE="0">
<INPUT TYPE=HIDDEN NAME=MT VALUE="0">
<INPUT TYPE=HIDDEN NAME=NN VALUE="0">
<input type=hidden name=TTKH VALUE="0">
<input type=hidden name=TNKH VALUE="0">
<INPUT TYPE=HIDDEN NAME=TOTM VALUE="0">
<INPUT TYPE=HIDDEN NAME=BH VALUE="0">
<INPUT TYPE=HIDDEN NAME=KTAXM VALUE="0">
<INPUT TYPE=HIDDEN NAME=RELTOTMONEY  VALUE="0">
<INPUT TYPE=HIDDEN NAME=H1 VALUE="0">
<INPUT TYPE=HIDDEN NAME=H2 VALUE="0">
<INPUT TYPE=HIDDEN NAME=H3 VALUE="0">
<INPUT TYPE=HIDDEN NAME=B3 VALUE="0">
<INPUT TYPE=HIDDEN NAME=KZHOUR VALUE="0">
<INPUT TYPE=HIDDEN NAME=Forget VALUE="0">
<INPUT TYPE=HIDDEN NAME=JIAA VALUE="0">
<INPUT TYPE=HIDDEN NAME=JIAB VALUE="0"> 
<INPUT TYPE=HIDDEN NAME=exrt >
<INPUT TYPE=HIDDEN NAME=XIANMBAK   >
<INPUT TYPE=HIDDEN NAME=ZHUANMBAK  >
<INPUT TYPE=HIDDEN NAME=XIANM   >
<INPUT TYPE=HIDDEN NAME=ZHUANM >
 
			

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
	tmpRec = Session("YECE1001HW")
	for CurrentRow = 1 to PageRec
		'tmpRec(CurrentPage, CurrentRow, 0) = request("op")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 6) = request("F1_JOB")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 19) = request("BBCODE")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 20) = request("BB")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 21) = request("CVCODE")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 22) = request("CV")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 23) = request("PHU")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 24) = request("KT")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 25) = request("TTKH")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 26) = request("MT")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 28) = request("TNKH")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 29) = request("JX")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 30) = request("BH")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 31) = request("QITA")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 32) = request("KTAXM")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 42) = request("B3")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 43) = request("B3M")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 44) = request("ZHUANM")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 45) = request("XIANM")(CurrentRow)
	next
	Session("YECE1001HW") = tmpRec

End Sub
%>

<script language=vbscript>
function BACKMAIN()
	open "../main.asp" , "_self"
end function

function clr()
	open "<%=SELF%>.asp" , "_parent"
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
	open "<%=self%>.memo.asp?index="& index &"&currentpage=" & cp &"&yymm=" & yymm  , "_blank" , "top=10, left=10, width=450,height=450, scrollbars=yes"
end function  


function zhuanmchg(index)
	F_EXRT = CDBL(<%=SELF%>.EXRT(INDEX).VALUE)
	REL_TOTAMT = CDBL(<%=SELF%>.RELTOTMONEY(INDEX).VALUE)	
	yymmstr=<%=self%>.yymm.value 
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
	else 
		IF  CDBL(<%=SELF%>.ZHUANM(INDEX).VALUE)+cdbl(<%=SELF%>.XIANM(INDEX).VALUE)  > CDBL(REL_TOTAMT)  THEN 
			ALERT "轉款金額輸入錯誤!!(大於實領金額)"
			<%=SELF%>.ZHUANM(INDEX).VALUE = <%=SELF%>.ZHUANMBAK(INDEX).VALUE 
			<%=SELF%>.XIANM(INDEX).VALUE = <%=SELF%>.XIANMBAK(INDEX).VALUE 			
			'<%=SELF%>.ZHUANM(INDEX).SELECTED()
			<%=SELF%>.ZHUANM(INDEX).FOCUS()
			EXIT FUNCTION 
		END  IF  
	end if 	
	<%=SELF%>.XIANM(INDEX).VALUE = CDBL(REL_TOTAMT) - CDBL(F_ZHUANM)
    '<%=SELF%>.XIANM(INDEX).FOCUS()        
    CODESTR01 = <%=SELF%>.ZHUANM(INDEX).VALUE  
    CODESTR02 = <%=SELF%>.XIANM(INDEX).VALUE
    open "<%=SELF%>.back.asp?ftype=ZXCHG&index="& index &"&CurrentPage="& <%=CurrentPage%> & _
		 "&CODESTR01=" & CODESTR01 & "&CODESTR02=" &CODESTR02 & "&yymm="& yymmstr , "Back"
    'PARENT.BEST.COLS="70%,30%"
END FUNCTION   



FUNCTION BBCODECHG(INDEX)
	codestr=<%=self%>.bbcode(index).value
	daystr=<%=self%>.MMDAYS.value
	yymmstr=<%=self%>.yymm.value 
	open "<%=SELF%>.back.asp?ftype=A&index="& index &"&CurrentPage="& <%=CurrentPage%> & _
		 "&days=" & daystr & "&code=" &	codestr & "&yymm="& yymmstr  , "Back"
	'DATACHG(INDEX)

	'PARENT.BEST.COLS="70%,30%"
END FUNCTION

FUNCTION JOBCHG(INDEX)
	codestr=<%=self%>.F1_JOB(index).value
	daystr=<%=self%>.MMDAYS.value
	yymmstr=<%=self%>.yymm.value 
	open "<%=SELF%>.back.asp?ftype=B&index="& index &"&CurrentPage="& <%=CurrentPage%> & _
		 "&days=" &daystr & "&code=" &	codestr & "&yymm="& yymmstr  , "Back"
	'PARENT.BEST.COLS="70%,30%"
	'DATACHG(INDEX)
END FUNCTION

FUNCTION B3CHG(INDEX)
	IF ISNUMERIC(<%=SELF%>.B3(INDEX).VALUE) THEN
		 
		<%=SELF%>.B3M(INDEX).VALUE=CDBL(<%=SELF%>.B3(INDEX).VALUE)*5   '200801起夜班津貼原先[2USD]改為[5USD]  
		DATACHG(INDEX)
	ELSE
		ALERT "請輸入正確天數!!"
		<%=SELF%>.B3(INDEX).VALUE="0"
		<%=SELF%>.B3M(INDEX).VALUE="0"
		<%=SELF%>.B3(INDEX).FOCUS()
	END IF
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

	if isnumeric(<%=SELF%>.KTAXM(INDEX).VALUE)=false then  '稅金(-)
		alert "請輸入數字!!"
		<%=self%>.KTAXM(index).value=0
		<%=self%>.KTAXM(index).focus()
		<%=self%>.KTAXM(index).select()
		exit FUNCTION
	end if
	if isnumeric(<%=SELF%>.QITA(INDEX).VALUE)=false then  '其他扣除額(-)
		alert "請輸入數字!!"
		<%=self%>.QITA(index).value=0
		<%=self%>.QITA(index).focus()
		<%=self%>.QITA(index).select()
		exit FUNCTION
	end if

	if isnumeric(<%=SELF%>.JX(INDEX).VALUE)=false then  '績效(+)
		alert "請輸入數字!!"
		<%=self%>.JX(index).value=0
		<%=self%>.JX(index).focus()
		<%=self%>.JX(index).select()
		exit FUNCTION
	end if
	TTM = ( cdbl(<%=self%>.bb(index).value) + cdbl(<%=self%>.CV(index).value) + cdbl(<%=self%>.PHU(index).value) + CDBL(<%=self%>.KT(index).value))
	TTMH = round (cdbl(<%=self%>.bb(index).value)/30/8,3)    '時薪

	'alert  TTMH
	'<%=self%>.HHMOENY(index).value = TTMH

	CODESTR01 = <%=SELF%>.BB(INDEX).VALUE
	CODESTR02 = <%=SELF%>.CV(INDEX).VALUE
	CODESTR03 = <%=SELF%>.PHU(INDEX).VALUE
	CODESTR04 = <%=SELF%>.KT(INDEX).VALUE
	CODESTR05 = <%=SELF%>.TTKH(INDEX).VALUE
	CODESTR06 = <%=SELF%>.TNKH(INDEX).VALUE
	CODESTR07 = <%=SELF%>.JX(INDEX).VALUE
	CODESTR08 = <%=SELF%>.BH(INDEX).VALUE
	CODESTR09 = <%=SELF%>.QITA(INDEX).VALUE
	CODESTR10 = <%=SELF%>.KTAXM(INDEX).VALUE
	CODESTR11 = <%=SELF%>.B3(INDEX).VALUE
	CODESTR12 = <%=SELF%>.B3M(INDEX).VALUE
	CODESTR13 = <%=SELF%>.exrt(INDEX).VALUE
	CODESTR14 = <%=SELF%>.MT(INDEX).VALUE
	CODESTR15 = <%=SELF%>.acc(INDEX).VALUE
	CODESTR16 = <%=SELF%>.dkm(INDEX).VALUE
	
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
		 "&CODESTR16="& CODESTR16 &_
		 "&yymm="& yymmstr &_
		 "&days=" & daystr , "Back"

	'PARENT.BEST.COLS="70%,30%"

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
	OPEN "../zzz/getempWorkTime.asp?yymm=" & yymmstr &"&EMPID=" & empidstr , "_blank" , "top=10, left=10,  scrollbars=yes"
end function

</script>
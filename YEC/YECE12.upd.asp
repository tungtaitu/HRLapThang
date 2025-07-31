<%@LANGUAGE="VBSCRIPT" codepage=65001%>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<!-- #include file = "../GetSQLServerConnection.fun" -->
<!-- #include file="../ADOINC.inc" -->
<%

CurrentPage = request("CurrentPage")
TotalPage = request("TotalPage")
PageRec = request("PageRec")
gTotalPage = request("gTotalpage")

Set conn = GetSQLServerConnection()
tmpRec = Session("yece12B")
cfg  = request("cfg") 

firstday  = request("calcdat")
endday = request("ccdt") 

response.write "firstday=" & firstday &"<BR>"
response.write "endday=" & endday &"<BR>" 
response.write "cfg=" & cfg  

country = request("country")
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
			IF TRIM(tmpRec(i, j, 20))="" THEN BB = 0 ELSE BB = replace(CDBL(tmpRec(i, j, 20)),",","")
			IF TRIM(tmpRec(i, j, 21))="" THEN CV = 0 ELSE CV = replace(CDBL(tmpRec(i, j, 21)),",","")			
			IF TRIM(tmpRec(i, j, 22))="" THEN PHU=0 ELSE PHU = replace(CDBL(tmpRec(i, j, 22)),",","")			
			IF TRIM(tmpRec(i, j, 23))="" THEN NN=0 ELSE NN=	replace(CDBL(tmpRec(i, j, 23)),",","")			
			IF TRIM(tmpRec(i, j, 24))="" THEN KT=0 ELSE KT=	replace(CDBL(tmpRec(i, j, 24)),",","")			
			IF TRIM(tmpRec(i, j, 25))="" THEN MT=0 ELSE MT=	replace(CDBL(tmpRec(i, j, 25)),",","")			
			IF TRIM(tmpRec(i, j, 26))="" THEN TTKH=0 ELSE TTKH=	replace(CDBL(tmpRec(i, j, 26))	,",","")		
			IF TRIM(tmpRec(i, j, 27))="" THEN QC=0	ELSE QC=	replace(CDBL(tmpRec(i, j, 27)),",","")
			IF TRIM(tmpRec(i, j, 28))="" THEN TNKH=0 ELSE TNKH=replace(CDBL(tmpRec(i, j, 28)),",","")			
			IF TRIM(tmpRec(i, j, 29))="" THEN JX=0 ELSE JX=replace(CDBL(tmpRec(i, j, 29))	,",","")	  
			IF TRIM(tmpRec(i, j, 30))="" THEN GT=0 ELSE	GT=	replace(CDBL(tmpRec(i, j, 30)),",","")
			IF TRIM(tmpRec(i, j, 32))="" THEN BH=0 ELSE BH=	replace(CDBL(tmpRec(i, j, 32)),",","")
			IF TRIM(tmpRec(i, j, 34))="" THEN QITA=0 ELSE QITA=	replace(CDBL(tmpRec(i, j, 34))	,",","")					
			IF TRIM(tmpRec(i, j, 41))="" THEN TBTR=0 ELSE TBTR=replace(CDBL(tmpRec(i, j, 41)),",","")
					
			'IF tmpRec(i, j, 35)="" THEN HS=0 ELSE HS=	CDBL(tmpRec(i, j, 35))			' 2009起不扣伙食費 
			HS = 0 
			
			IF tmpRec(i, j, 64)="" THEN MONEY_H=0	ELSE MONEY_H=replace(CDBL(tmpRec(i, j, 64)),",","")	'時薪
			IF tmpRec(i, j, 42)="" THEN	REAL_TOTAL=0 ELSE REAL_TOTAL=	replace(CDBL(tmpRec(i, j, 42)),",","")			 '實領薪資 
			IF tmpRec(i, j, 42)="" THEN	RELTOTMONEY=0	ELSE RELTOTMONEY= replace(ROUND(CDBL(tmpRec(i, j, 42)),0),",","")  '實領

			'請假應扣款 (請假應扣款全部計入 empdsalary.BZKM ) 	
			if tmpRec(i, j, 54)="" then allKM=0 else allKM=replace(cdbl(tmpRec(i, j, 54)),",","") 				
			IF tmpRec(i, j, 57)="" THEN	H1=0 ELSE H1=	replace(CDBL(tmpRec(i, j, 57)),",","")						
			IF tmpRec(i, j, 58)="" THEN H2=0 ELSE	H2=	replace(CDBL(tmpRec(i, j, 58)),",","")			
			IF tmpRec(i, j, 59)="" THEN	H3=0 ELSE	H3=replace(CDBL(tmpRec(i, j, 59))	,",","")		
			IF tmpRec(i, j, 60)="" THEN	B3=0 ELSE B3=replace(CDBL(tmpRec(i, j, 60)),",","")		 
			IF tmpRec(i, j, 90)="" THEN	B4=0 ELSE B4=replace(CDBL(tmpRec(i, j, 90)),",","")
			
			
			IF tmpRec(i, j, 61)="" THEN	KZHOUR=0 ELSE	KZHOUR=replace(CDBL(tmpRec(i, j, 61)),",","")
			IF tmpRec(i, j, 62)="" THEN	FL=0 ELSE FL=replace(CDBL(tmpRec(i, j, 62)),",","") 
			
			IF tmpRec(i, j, 35)="" THEN	JIAA=0 ELSE	JIAA=replace(CDBL(tmpRec(i, j, 35)),",","")   '200910事病假合併計算計入 jiaA 			
			'IF tmpRec(i, j, 46)="" THEN	JIAB=0 ELSE	JIAB=CDBL(tmpRec(i, j, 46))
			jiaB=0 			
			'稅金
			IF TRIM(tmpRec(i, j, 40))="" THEN KTAXM=0	ELSE KTAXM=replace(CDBL(tmpRec(i, j, 40)),",","")  '所得稅
			IF YYMM<"200802" THEN 				
				RELTOTMONEY = RELTOTMONEY + KTAXM 
				KTAXM = 0
			END  IF 
			
			
			'response.write tmpRec(i, j, 1) & "需繳稅=" & KTAXM &"<BR>"  						
			'200910 請假扣款全數計入allkm 
			 if tmpRec(i, j, 54)="" then allKM=0 else allKM=replace(cdbl(tmpRec(i, j, 54)),",","")   
			 IF TRIM(tmpRec(i, j, 47))="" THEN	H1M=0	ELSE	H1M=CDBL(tmpRec(i, j, 47))
			 IF TRIM(tmpRec(i, j, 48))="" THEN	H2M=0	ELSE	H2M=CDBL(tmpRec(i, j, 48))
			 IF TRIM(tmpRec(i, j, 49))="" THEN	H3M=0	ELSE	H3M=CDBL(tmpRec(i, j, 49))
			 IF TRIM(tmpRec(i, j, 50))="" THEN	B3M=0	ELSE	B3M=CDBL(tmpRec(i, j, 50))
			 IF TRIM(tmpRec(i, j, 91))="" THEN	B4M=0	ELSE	B4M=CDBL(tmpRec(i, j, 91))
			 IF tmpRec(i, j, 78)="" THEN H12=0 ELSE H12=	replace(CDBL(tmpRec(i, j, 78)),",","")  '平(夜)加班
			 IF TRIM(tmpRec(i, j, 79))="" THEN	H12M=0	ELSE	H12M=CDBL(tmpRec(i, j, 79))  '平(夜)加班費
			' IF TRIM(tmpRec(i, j, 74))="" THEN	KZM=0	ELSE	KZM=CDBL(tmpRec(i, j, 74))	 
			'IF TRIM(tmpRec(i, j, 35))="" THEN	JIAAM=0	ELSE JIAAM=CDBL(tmpRec(i, j, 35))	 '200910事病假合併計算計入 jiaA  	 
			'IF TRIM(tmpRec(i, j, 57))="" THEN	JIABM=0	ELSE JIABM=CDBL(tmpRec(i, j, 57))		
			
			kzm=0
			jiaAM=0 
			jiaBM=0 
			
			IF TRIM(tmpRec(i, j, 43))="" THEN JISHU=0	ELSE JISHU=replace(CDBL(tmpRec(i, j, 43)),",","")		  '離職補助金基數 
			IF TRIM(tmpRec(i, j, 44))="" THEN LZBZJ=0	ELSE LZBZJ=replace(CDBL(tmpRec(i, j, 44)),",","")    '離職補助金 ( 自200901起沒有 ) 
			IF TRIM(tmpRec(i, j, 56))="" THEN TOTM=0  ELSE	TOTM=replace(CDBL(tmpRec(i, j, 56)),",","")   '總薪資(+項)
			IF TRIM(tmpRec(i, j, 75))="" THEN sgkm=0  ELSE	sgkm=replace(CDBL(tmpRec(i, j, 75)),",","")   '事故扣款
			IF TRIM(tmpRec(i, j, 76))="" THEN btien=0  ELSE	btien=replace(CDBL(tmpRec(i, j, 76)),",","")   '補薪			
			IF TRIM(tmpRec(i, j, 77))="" THEN money1=0  ELSE	money1=replace(CDBL(tmpRec(i, j, 77)),",","")   '特別獎金
			TNKH = cdbl(TNKH)   '202007 elin			 
			
			if tmpRec(i, j, 80)="" then butax=0 else butax=replace(CDBL(tmpRec(i, j, 80)),",","")   '伙食費(外籍)
			if tmpRec(i, j, 81)="" then hsf=0 else hsf=replace(CDBL(tmpRec(i, j, 81)),",","")   '伙食費(外籍)
			if tmpRec(i, j, 82)="" then zhuanM=0 else zhuanM=replace(CDBL(tmpRec(i, j, 82)),",","")  'emp實領			
			if tmpRec(i, j, 83)="" then tien3=0 else tien3=replace(CDBL(tmpRec(i, j, 83)),",","")  'tien3	  =buamt 補住
			if tmpRec(i, j, 84)="" then vnbbtax=0 else vnbbtax=replace(CDBL(tmpRec(i, j, 84)),",","")  'vnbbtax
			if tmpRec(i, j, 86)="" then govbb=0 else govbb=replace(CDBL(tmpRec(i, j, 86)),",","")  'govBB
			if tmpRec(i, j, 87)="" then govOadd=0 else govOadd=0 'replace(CDBL(tmpRec(i, j, 86)),",","") 
			IF tmpRec(i, j, 88)="" THEN	B5=0 ELSE B5=replace(CDBL(tmpRec(i, j, 88)),",","")
			IF TRIM(tmpRec(i, j, 89))="" THEN	B5M=0	ELSE	B5M=CDBL(tmpRec(i, j, 89))
			
			if tmpRec(i, j, 4) ="VN" or tmpRec(i, j, 4) ="CT" then 
				vnbb=0 
				response.write "<br>===>1"&"<br>"
			elseif  tmpRec(i, j, 4)="TW" or tmpRec(i, j, 4)="MA"   then 
				if tmpRec(i, j, 85)="B" then 
					response.write "===>2"&"<br>"
					vnbb =cdbl(BB)+cdbl(ttkh)+cdbl(cv)
				else
					response.write "===>3"&"<br>"
					vnbb =cdbl(BB)+cdbl(ttkh)
				end if 
			elseif tmpRec(i, j, 4)="CN"  then 
				vnbb =cdbl(BB)+cdbl(ttkh)+cdbl(phu)+cdbl(cv)-cdbl(qita)
				response.write "===>4"&"<br>"
			elseif tmpRec(i, j, 4)="TA"  then 
				vnbb =cdbl(BB)+cdbl(phu)+cdbl(cv)+cdbl(nn)+cdbl(kt)+cdbl(mt)+cdbl(ttkh)-cdbl(qita)				
				response.write "<br>===>5"&"<br>"
			else 	
				response.write "<br>===>6"&"<br>"
			end if 
			
			response.write "xxxxvnbb="& tmpRec(i, j, 4) &"---"& vnbb &"<br>" 
			if vnbb="" then vnbb=0 
			'govOadd			
			'if tmpRec(i, j, 89)="" then butax=0 else butax=replace(CDBL(tmpRec(i, j, 89)),",","")  'emp實領
			
			'tmpRec(i, j, 83)=rs("buamt")  				
			'tmpRec(i, j, 89)=rs("tien3") '公司補 (留存用) , TA適用 + TW +MA 有獎金 
				
			'TNKH = cdbl(TNKH)-cdbl(money1) 
			
			
			'
			
			'response.write  tmpRec(i, j, 1) &","& TNKH &"<br>" 
			'response.end 
			'if  CDBL(tmpRec(i, j,59)) < CDBL(MMDAYS) THEN BZKM = tmpRec(i,j,65) else 	BZKM = 0		 
			

			workhour = cdbl(trim(tmpRec(i, j, 39)))*8  '總工時
			f_wkdays = trim(tmpRec(i, j, 39)) 
		
			'零數 (以10,000為單位) , 如當月離職(以1,000為單位) 
			'當月離職
			if trim(tmpRec(i, j, 19))<>"" and ( (trim(tmpRec(i, j, 19)))> (firstday)   and  (trim(tmpRec(i, j, 19))) <= (endday) )   then   
				if trim(tmpRec(i, j, 4))="VN" then 
					'response.write trim(tmpRec(i, j,1 ))&"LZBZJ-------------------"&"<BR>" 
					'LAOHN= fix( (RELTOTMONEY+LZBZJ) / 1000 ) * 1000 
					'SOLE =  RELTOTMONEY- LAOHN 		 
					'200901不計離職補助金   
					if cdbl(RELTOTMONEY)<=0 then 
						LAOHN = 0 
						SOLE =  cdbl(TBTR) + cdbl(RELTOTMONEY)
					else
						LAOHN=  RELTOTMONEY ' LAOHN'fix( (RELTOTMONEY) / 1000 ) * 1000  '--20180731不用零數
						SOLE =  RELTOTMONEY - LAOHN 					
					end if 	
				else 
					LAOHN = RELTOTMONEY
					SOLE = 0 	
				end if 	
				
			else
				if trim(tmpRec(i, j, 4))="VN" or trim(tmpRec(i, j, 4))="CT" then
					if cdbl(RELTOTMONEY)<=0 then 
						LAOHN = 0 
						SOLE =  cdbl(TBTR) + cdbl(RELTOTMONEY)
					else  					
						LAOHN=  cdbl(RELTOTMONEY)  'fix( RELTOTMONEY/ 10000 ) * 10000  '--20180731不用零數
						SOLE =  RELTOTMONEY - LAOHN
						'if cdbl(SOLE) > 5000 then
						'	LAOHN = LAOHN + 10000
						'	SOLE = RELTOTMONEY -  LAOHN
						'else
						'	LAOHN = LAOHN
						'	SOLE = SOLE
						'end if
					end if 	
				else
					LAOHN = RELTOTMONEY
					SOLE = 0
				end if 
			end if 
						
			if trim(tmpRec(i, j, 4))="VN" or trim(tmpRec(i, j, 4))="CT" then
				dm="VND"
				ZHUANM = LAOHN
				XIANM=  0 
			else
				dm="USD"
				'ZHUANM = LAOHN
				XIANM = 0
			end if    		 
			
			 
			MEMOSTR = trim(REPLACE(tmpRec(i,j,66), "'", "" ))
			MEMOSTR = trim(REPLACE (MEMOSTR, vbCrLf ,"<br>")) 

			SQLSTR="SELECT * FROM  EMPDSALARY WHERE YYMM='"& YYMM  &"' AND EMPID='"& trim(tmpRec(i, j, 1)) &"'  "
			Set rs = Server.CreateObject("ADODB.Recordset")
			RS.OPEN SQLSTR, CONN , 3, 3

			IF RS.EOF THEN
				rs.close 
				set rs=nothing
				SQL1="INSERT INTO EMPDSALARY ( WHSNO, Country, EMPID, indat, outdat , GROUPID,JOB,BB,CV,PHU,NN,KT,MT,TTKH,QC,TNKH, "&_
					 "JX,TBTR,H1M, H2M, H3M, B3M,B4M,TOTM, H1, H2, H3, B3,B4,BH, HS,GT,QITA, FL, JIAA, JIAB, KZHOUR, "&_
					 "KZM, JIAAM, JIABM, MONEY_H, REAL_TOTAL, LAONH,SOLE,YYMM,MDTM,MUSER, WORKDAYS , workshour,  "&_
					 "JISHU, LZBZJ, BZKM, KTAXM, dm, ZHUANM, XIANM , userIP , memo, sgkm , btien , h12, h12m  , tien2,  tien3,  vnbb , vnbbtax, hsf, govBB, govOadd, butax,B5,B5M  ) VALUES ( "&_
					 "'"& tmpRec(i, j, 7) &"' ,'"& Trim(tmpRec(i, j, 4)) &"' , '"& tmpRec(i, j, 1) &"','"& tmpRec(i, j, 5) &"', '"& tmpRec(i, j, 19) &"',  '"& tmpRec(i, j, 9) &"' , '"& tmpRec(i, j, 6) &"'  ,  "&_
					 "'"& BB &"', '"& CV &"', '"& PHU &"'  , '"& NN &"'  , '"& KT &"'  , '"& MT &"'  , '"& TTKH &"'  , "&_
					 "'"& QC &"', '"& TNKH &"', '"& JX &"', '"& TBTR &"', '"& H1M &"', '"& H2M &"', '"& H3M &"', '"& B3M &"','"& B4M &"', '"& TOTM &"' , "&_
					 "'"& H1 &"', '"& H2 &"', '"& H3 &"', '"& B3 &"','"& B4 &"', '"& BH &"'  , '"& HS &"'  ,  "&_
					 "'"& GT &"'  , '"& QITA &"', '"& FL &"', '"& JIAA &"', '"& JIAB &"', '"& KZHOUR &"', "&_
					 "'"& KZM &"', '"& JIAAM &"', '"& JIABM &"', '"& MONEY_H &"',  "&_
					 "'"& RELTOTMONEY &"', '"& LAOHN &"', "& SOLE &", '"& YYMM &"', GETDATE(), '"& SESSION("NETUSER")&"', "&_
					 "'"& f_wkdays &"','"& workhour &"',  '"& JISHU &"', '"& LZBZJ &"' , '"& allkm &"', '"& KTAXM &"' , "&_
					 "'"& DM &"' ,'"& ZHUANM &"', '"& XIANM &"', '"& session("vnlogIP") &"' ,'"& MEMOSTR &"','"& sgkm &"' ,'"& btien &"' , '"& h12 &"','"& h12m &"' ,'"&money1&"' ,'"&tien3&"','"&vnbb&"' ,'"&vnbbtax&"' ,'"&hsf&"' , '"&govbb&"', '"&govOadd&"' , '"&butax&"', '"&B5&"', '"&B5M&"' ) "
				RESPONSE.WRITE SQL1 &"<br>"
				X = X + 1
				'response.end
				conn.execute(SQL1)
			ELSE
				SQL2="update EMPDSALARY set GROUPID='"& trim(tmpRec(i, j, 9)) &"', JOB='"& trim(tmpRec(i, j, 6))  &"', "&_
					 "indat='"& Trim(tmpRec(i, j, 5)) &"', outdat='"& Trim(tmpRec(i, j, 19)) &"', "&_
					 "COUNTRY='"& Trim(tmpRec(i, j, 4)) &"' , BB='"& BB &"' , CV='"& CV &"' , PHU='"& PHU &"', NN='"& NN &"', "&_
					 "KT='"& KT &"' , MT='"& MT &"' , TTKH='"& TTKH &"', QC='"& QC &"',  "&_
					 "TNKH='"& TNKH &"' ,  BH='"& BH &"' , HS='"& HS &"' , GT='"& GT &"' , "&_
					 "QITA='"& QITA &"' , MONEY_H ='"& MONEY_H &"' , "&_
					 "REAL_TOTAL='"& RELTOTMONEY  &"' , LAONH='"& LAOHN &"' ,SOLE='"& SOLE &"', "&_
					 "H1M='"& H1M  &"', H2M='"& H2M &"', H3M='"& H3M &"', B3M='"& B3M &"',B4M='"& B4M &"',B5M='"& B5M &"', TOTM='"& TOTM &"' , "&_
					 "KZM='"& KZM  &"', JIAAM='"& JIAAM &"', JIABM='"& JIABM &"', "&_
					 "H1='"& H1  &"', H2='"& H2 &"', H3='"& H3 &"', B3='"& B3 &"',B4='"& B4 &"',B5='"& B5 &"',  "&_
					 "KZHOUR='"& KZHOUR  &"', FL='"& FL &"', JIAA='"& JIAA &"', JIAB='"& JIAB &"', "&_
					 "WORKDAYS='"& f_wkdays &"', JX='"& JX &"', TBTR='"& TBTR &"',  DM='"& DM &"' , ZHUANM='"& ZHUANM &"', XIANM='"& XIANM &"' ,    "&_
					 "JISHU='"& JISHU &"' , LZBZJ='"& LZBZJ &"' , workshour = '"& workhour &"', BZKM='"& allkm &"' , KTAXM='"& KTAXM &"', "&_
					 "sgkm='"& sgkm &"' ,btien='"& btien &"' , h12='"& h12 &"' , h12M='"& h12m &"'   "&_
					 ",tien2='"& money1&"' "&_
					 ",tien3='"&tien3 &"' "&_
					 ",hsf='"& hsf &"' "&_
					 ",vnbb='"& vnbb&"' "&_
					 ",vnbbtax='"& vnbbtax&"' "&_
					 ",govBB='"& govBB &"' "&_
					 ",govOadd='"& govOadd & "' "&_
					 ",butax='"& butax &"' "&_
					 ",mdtm=getdate(), muser='"& session("NETUSER") &"' , userIP='"& session("vnlogIP")&"',memo='"& MEMOSTR &"' "&_
					 ", WHSNO='"& trim(tmpRec(i, j, 7)) &"' "&_
					 "where YYMM='"& YYMM  &"' AND EMPID='"& trim(tmpRec(i, j, 1)) &"'  "
				RESPONSE.WRITE SQL2 &"<br>"
				X = X + 1
				conn.execute(SQL2)
			END IF 
			set rs=nothing
			'response.end 
			SQLSTR="SELECT * FROM  EMPDSALARY_BAK WHERE YYMM='"& YYMM  &"' AND EMPID='"& trim(tmpRec(i, j, 1)) &"' AND WHSNO='"& trim(tmpRec(i, j, 7)) &"'  "
			Set rsT = Server.CreateObject("ADODB.Recordset")
			RST.OPEN SQLSTR, CONN, 1, 1 
			if not rst.eof then 
				sql3="update EMPDSALARY_BAK set ZHUANM='"& ZHUANM &"', XIANM='"& XIANM &"', memo='"& MEMOSTR &"'  "&_
					 "where YYMM='"& YYMM  &"' AND EMPID='"& trim(tmpRec(i, j, 1)) &"' AND WHSNO='"& trim(tmpRec(i, j, 7)) &"' " 
				conn.execute(sql3) 	 
				response.write sql3 &"<BR>"
			end if  	 
			set rst=nothing 		
		END IF
		
		if trim(tmpRec(i, j, 4)) ="VN" or trim(tmpRec(i, j, 4)) ="CT" then 
		sqla="update  VYFYMYJX set RELJXM='"&jx&"' , mdtm=getdate(), muser='"& session("netuser") &"' where yymm='"& YYMM  &"'   and empid='"& trim(tmpRec(i, j, 1)) &"' "
		conn.execute(sqla)
		end if 
		
	next
next
response.write err.number &"<BR>"
response.write conn.errors.count &"<BR>"

for g =0 to conn.errors.count-1
	response.write conn.errors.item(g)&"<br>"
	response.write Err.Description
next  
'response.clear
'RESPONSE.END 

if err.number = 0 then
	conn.CommitTrans
	Set Session("yece12B") = Nothing
	Set conn = Nothing 
	
%><SCRIPT LANGUAGE=VBSCRIPT>
			ALERT "資料處理成功!!"&"<%=X%>"&"筆"&chr(13)&"DATA CommitTrans Success !!"
			OPEN "YECE12.Fore.asp?CT="&"<%=country%>" , "_self"
	</script>
<%
ELSE
	conn.RollbackTrans 
%>	<SCRIPT LANGUAGE=VBSCRIPT>
		ALERT "資料處理失敗!!"&chr(13)&"DATA CommitTrans ERROR !!"
		OPEN "YECE12.Fore.asp?CT="&"<%=country%>" , "_self"
	</script>
<%
	response.end
END IF
%>
 
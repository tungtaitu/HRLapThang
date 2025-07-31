<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<!-- #include file = "../../GetSQLServerConnection.fun" -->
<!-- #include file="../../ADOINC.inc" -->
<%
SELF = "empsalaryCN"
session.codepage=65001
ftype = request("ftype")
code = request("code")
index=request("index")
CurrentPage = request("CurrentPage")

yymm=request("yymm")
 '一個月有幾天
cDatestr=CDate(LEFT(YYMM,4)&"/"&RIGHT(YYMM,2)&"/01")
Cdays = DAY(cDatestr+(32-DAY(cDatestr))-DAY(cDatestr+(32-DAY(cDatestr))))   '一個月有幾天
response.write days
ENDdat = CDate(LEFT(YYMM,4)&"/"&RIGHT(YYMM,2)&"/"&Cdays)

CODESTR01 = REQUEST("CODESTR01")
CODESTR02 = REQUEST("CODESTR02")
CODESTR03 = REQUEST("CODESTR03")
CODESTR04 = REQUEST("CODESTR04")
CODESTR05 = REQUEST("CODESTR05")
CODESTR06 = REQUEST("CODESTR06")
CODESTR07 = REQUEST("CODESTR07")
CODESTR08 = REQUEST("CODESTR08")
CODESTR09 = REQUEST("CODESTR09")
CODESTR10 = REQUEST("CODESTR10")
CODESTR11 = REQUEST("CODESTR11")
CODESTR12 = REQUEST("CODESTR12")
CODESTR13 = REQUEST("CODESTR13")
CODESTR14 = REQUEST("CODESTR14")
workdays = REQUEST("days") 
memostr = replace(trim(request("memo")),vbCrLf,"<BR>")
response.write  "CODESTR13=" & CODESTR13 &"<BR>"

tmpRec = Session("empfilesalaryCN")
response.write "index=" & index &"<BR>"
response.write "ftype=" & ftype &"<BR>"
RESPONSE.WRITE "RELWORKDAY=" & tmpRec(CurrentPage,index + 1,34) &"<BR>"
RESPONSE.WRITE "EXRT=" & tmpRec(CurrentPage,index + 1,40) &"<BR>"

Set conn = GetSQLServerConnection()
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="refresh">
</head>
<%
select case ftype
	case "A"
		sql="select * from empsalaryBasic where func='AA' and code='"& code &"' and country='"& tmpRec(CurrentPage,index + 1,4) &"'  "
		response.write sql &"<br>"
		'response.end
		Set rst= Server.CreateObject("ADODB.Recordset")
  		rst.Open SQL, conn, 3,3
  		if not rst.eof then
  			tmpRec(CurrentPage,index + 1,0) = "UPD"
  			tmpRec(CurrentPage,index + 1,19) = code
  			tmpRec(CurrentPage,index + 1,20) = rst("bonus")
	  		TTMH = round( (tmpRec(CurrentPage,index + 1,20)/30/8), 3 )
  			tmpRec(CurrentPage,index + 1,38)=TTMH   '時薪

  			'+(加項)
	  		BB = tmpRec(CurrentPage,index + 1,20)
	  		CV = tmpRec(CurrentPage,index + 1,22)
	  		PHU = tmpRec(CurrentPage,index + 1,23)
	  		KT = tmpRec(CurrentPage,index + 1,24)
	  		TTKH = tmpRec(CurrentPage,index + 1,25)
	  		MT = tmpRec(CurrentPage,index + 1,26)
	  		TNKH = tmpRec(CurrentPage,index + 1,28)
	  		JX = tmpRec(CurrentPage,index + 1,29)  '績效獎金
	  		B3M =  tmpRec(CurrentPage,index + 1,43)
	  		'-(減項)
	  		BH = tmpRec(CurrentPage,index + 1,30)
	  		QITA = tmpRec(CurrentPage,index + 1,31)
	  		KTAXM = tmpRec(CurrentPage,index + 1,32)

	  		IF (tmpRec(CurrentPage,index + 1,34)) < cdbl(workdays) then
				F1_MONEY = round( (( CDBL(BB)+CDBL(CV)+CDBL(PHU)+CDBL(KT)+CDBL(MT)) /cdbl(Cdays))*cdbl(tmpRec(CurrentPage,index + 1,34)),0)+cdbl(JX)+CDBL(B3M)+CDBL(TTKH)+CDBL(TNKH)
			else
				F1_MONEY = CDBL(BB)+CDBL(CV)+CDBL(PHU)+CDBL(KT)+cdbl(JX)+cdbl(MT)+CDBL(B3M)+CDBL(TTKH)+CDBL(TNKH)
		  	end if
		  	RESPONSE.WRITE "F1_MONEY=" & F1_MONEY &"<BR>"  
				
			F2_MONEY =CDBL(BH)+CDBL(QITA)    '應扣款
			RESPONSE.WRITE "F2_MONEY=" & F2_MONEY &"<BR>"
			TMONEY = CDBL(F1_MONEY) - CDBL(F2_MONEY) 			
			RESPONSE.WRITE "TMONEY=" & TMONEY &"<BR>"

			exrt = tmpRec(CurrentPage,index + 1,46)
			'超過800萬越幣應繳稅10% (總薪資+績效獎金>800萬VND) 
			'稅額計算 ( 累加 ) 
			'1.8,000,000~20,000,000   tax:10%
			'2.20,000,000~50,000,000  tax:20%
			'3.50,000,000~80,000,000  tax:30%
			'4.>80,000,000            tax:40% 
			F_TAX = 0 
			real_TOTAMT = (CDBL(TMONEY))* CDBL(tmpRec(CurrentPage,index + 1,40))  ' 實領金額
			' if cdbl(real_TOTAMT)>8000000 then 
				' if  cdbl(real_TOTAMT) <=20000000 then 
					' F_tax = ( cdbl(real_TOTAMT) - 8000000 ) * 0.1 
					' response.write "TAX1"&"<BR>"
				' elseif cdbl(real_TOTAMT) > 20000000 and cdbl(cdbl(real_TOTAMT)) <= 50000000 then 
					' F_tax = ( (20000000-8000000)* 0.1 )+((cdbl(real_TOTAMT) - 20000000)*0.2)
					' response.write "TAX2"&"<BR>"
				' elseif cdbl(real_TOTAMT) > 50000000 and cdbl(cdbl(real_TOTAMT)) <= 80000000 then 	
					' F_tax = ((20000000-8000000)* 0.1 )+ ( (50000000-20000000)* 0.2 ) + ((cdbl(real_TOTAMT) - 50000000)*0.3)
					' response.write "TAX3"&"<BR>"
				' elseif cdbl(real_TOTAMT) > 80000000 then 
				  	' F_tax = ((20000000-8000000)* 0.1 )+ ( (50000000-20000000)* 0.2 ) + ( (80000000-50000000)* 0.3 ) + ((cdbl(real_TOTAMT) - 80000000)*0.4)
				  	' response.write "TAX4"&"<BR>"
				' end if 
			' else
				' F_tax = 0  
			' end if 		
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
				
			tmpRec(CurrentPage,index + 1,32) = round(cdbl(F_tax) /cdbl(exrt),0)
			KTAXM = round(F_tax / cdbl(exrt) ,0)
			response.write F_tax  			
			
			
			tmpRec(CurrentPage,index + 1,36) = TMONEY '實領金額  
			
			'自200801起不補貼稅額 (稅額自付) 
			tmpRec(CurrentPage,index + 1,37)  = CDBL(BB)+CDBL(CV)+CDBL(PHU)+CDBL(KT)+ CDBL(MT)+cdbl(JX)+CDBL(TTKH)+CDBL(TNKH)+CDBL(B3M)-CDBL(F2_MONEY)-cdbl(ktaxm) '應領
			RESPONSE.WRITE "RELMONEY=" & tmpRec(CurrentPage,index + 1,37) &"<BR>"
			'不足月扣款
			tmpRec(CurrentPage,index + 1,39) = CDBL(tmpRec(CurrentPage,index + 1,37))-CDBL(tmpRec(CurrentPage,index + 1,36))
	  		tmpRec(CurrentPage,index + 1,41) = CDBL(BB)+CDBL(CV)+CDBL(PHU)+CDBL(KT)+cdbl(JX)+CDBL(TTKH)+CDBL(TNKH)+CDBL(MT)
	  		SUMTOT = CDBL(BB)+CDBL(CV)+CDBL(PHU)+CDBL(KT)+cdbl(JX)+CDBL(TTKH)+CDBL(TNKH)+CDBL(MT) 
		  	RESPONSE.WRITE "SUMTOT=" & SUMTOT &"<BR>"
		  	tmpRec(CurrentPage,index + 1,44) = round(fix(tmpRec(CurrentPage,index + 1, 36)),0)
		  	tmpRec(CurrentPage,index + 1,45) = (round( round(cdbl(tmpRec(CurrentPage, index + 1, 36)),2) - round(fix(tmpRec(CurrentPage, index + 1, 36)),0),2)* cdbl(codestr13)\1000)*1000
%>			<script language=vbs>
				Parent.Fore.<%=self%>.bb(<%=index%>).value=<%=rst("bonus")%>
				'Parent.Fore.<%=self%>.TNKH(<%=index%>).value=<%=CDBL(KTAXM)%>
				Parent.Fore.<%=self%>.HHMOENY(<%=index%>).value=<%=TTMH%>
				Parent.Fore.<%=self%>.BZKM(<%=index%>).value=<%=tmpRec(CurrentPage,index + 1,39)%>
				Parent.Fore.<%=self%>.TOTM(<%=index%>).value=<%=SUMTOT%>
				Parent.Fore.<%=self%>.RELTOTMONEY(<%=index%>).value=<%=(TMONEY)%>
				Parent.Fore.<%=self%>.ZHUANM(<%=index%>).value=<%=tmpRec(CurrentPage,index + 1,44)%>
				Parent.Fore.<%=self%>.XIANM(<%=index%>).value=<%=tmpRec(CurrentPage,index + 1,45)%>
				Parent.Fore.<%=self%>.PHU(<%=index%>).select()
			</script>
<% 		end if
		set rs=nothing
	case "B"
		sql="select * from empsalaryBasic where func='BB' AND JOB='"& code &"' AND COUNTRY='cn'  " 
		response.write sql
		'response.end 
		Set rst= Server.CreateObject("ADODB.Recordset")
  		rst.Open SQL, conn, 3,3
  		if not rst.eof then
  			tmpRec(CurrentPage,index + 1,0) = "UPD"
  			tmpRec(CurrentPage,index + 1,6) = code
  			tmpRec(CurrentPage,index + 1,21) = rst("CODE")
  			'tmpRec(CurrentPage,index + 1,22) = rst("bonus")
  			TTM = cdbl(rst("bonus"))+cdbl(tmpRec(CurrentPage, index+1, 20))+cdbl(tmpRec(CurrentPage, index+1, 23))
  			if tmpRec(CurrentPage,index + 1,4)="VN" then
	  			TTMH = round( (TTM/26/8), 0 )
	  		else
	  			TTMH = round( (TTM/30/8), 3 )
	  		end if
  			tmpRec(CurrentPage,index + 1,38)=TTMH

%>			<script language=vbs>
				Parent.Fore.<%=self%>.CV(<%=index%>).value=<%=rst("bonus")%>
				Parent.Fore.<%=self%>.CVCODE(<%=index%>).value="<%=rst("CODE")%>"
				Parent.Fore.<%=self%>.HHMOENY(<%=index%>).value=<%=TTMH%>
			</script>
<% 		end if
		set rs=nothing
	case "CDATACHG"
		RESPONSE.WRITE "XXX-----------------"&"<BR>"
		tmpRec(CurrentPage,index + 1,0) = "UPD"
  		tmpRec(CurrentPage,index + 1,20) = CODESTR01
  		tmpRec(CurrentPage,index + 1,22) = CODESTR02
  		tmpRec(CurrentPage,index + 1,23) = CODESTR03
  		tmpRec(CurrentPage,index + 1,24) = CODESTR04
  		tmpRec(CurrentPage,index + 1,25) = CODESTR05
  		tmpRec(CurrentPage,index + 1,26) = CODESTR14
  		tmpRec(CurrentPage,index + 1,28) = CODESTR06
  		tmpRec(CurrentPage,index + 1,29) = CODESTR07
  		tmpRec(CurrentPage,index + 1,30) = CODESTR08
  		tmpRec(CurrentPage,index + 1,31) = CODESTR09
  		tmpRec(CurrentPage,index + 1,32) = CODESTR10
  		tmpRec(CurrentPage,index + 1,42) = CODESTR11
  		tmpRec(CurrentPage,index + 1,43) = CODESTR12

  		TTM = cdbl(tmpRec(CurrentPage, index+1, 20))
  		if tmpRec(CurrentPage,index + 1,4)="VN" then
	  		TTMH = round( (TTM/26/8), 0 )
	  	else
	  		TTMH = round( (TTM/30/8), 3 )
	  	end if
  		tmpRec(CurrentPage,index + 1,38)=TTMH

  		'+(加項)
  		BB = tmpRec(CurrentPage,index + 1,20)
  		CV = tmpRec(CurrentPage,index + 1,22)
  		PHU = tmpRec(CurrentPage,index + 1,23)
  		KT = tmpRec(CurrentPage,index + 1,24)
  		TTKH = tmpRec(CurrentPage,index + 1,25)
  		MT  = tmpRec(CurrentPage,index + 1,26)
  		TNKH = tmpRec(CurrentPage,index + 1,28)
  		JX = tmpRec(CurrentPage,index + 1,29)  '績效獎金
			B3=tmpRec(CurrentPage,index + 1,42)
			B3M=tmpRec(CurrentPage,index + 1,43)

  		response.write   "BB=" & BB &"<BR>"
  		response.write   "CV=" & CV &"<BR>"
  		response.write   "PHU=" & PHU &"<BR>"
  		response.write   "KT=" & KT &"<BR>"
  		response.write   "MT=" & MT &"<BR>"
  		response.write   "TTKH=" & TTKH &"<BR>"
  		response.write   "TNKH=" & TNKH &"<BR>"
  		response.write   "JX=" & JX &"<BR>"
  		response.write   "B3M=" & B3M &"<BR>"
  		'-(減項)
  		BH = tmpRec(CurrentPage,index + 1,30)
  		QITA = tmpRec(CurrentPage,index + 1,31)
  		KTAXM = tmpRec(CurrentPage,index + 1,32)

  		response.write   "BH=" & BH &"<BR>"
  		response.write   "QITA=" & QITA &"<BR>"
  		response.write   "old_KTAXM=" & KTAXM &"<BR>"

  		IF (tmpRec(CurrentPage,index + 1,34)) < cdbl(workdays) then
			'F1_MONEY = round( ( CDBL(BB)+CDBL(CV)+CDBL(PHU)+CDBL(KT))/cdbl(days),0) *cdbl(tmpRec(CurrentPage,index + 1,34))+cdbl(JX)
			F1_MONEY = round( (( CDBL(BB)+CDBL(CV)+CDBL(PHU)+CDBL(KT)+CDBL(MT)) /cdbl(Cdays))*cdbl(tmpRec(CurrentPage,index + 1,34)),0)+cdbl(JX)+CDBL(B3M)+CDBL(TTKH)+CDBL(TNKH)
		else
			F1_MONEY = CDBL(BB)+CDBL(CV)+CDBL(PHU)+CDBL(KT)+CDBL(MT)+cdbl(JX)+CDBL(B3M)+CDBL(TTKH)+CDBL(TNKH)
	  	end if
	  	RESPONSE.WRITE "F1_MONEY=" & F1_MONEY &"<BR>"

			'超過800萬越幣應繳稅10% (總薪資+績效獎金>800萬VND) 
			'稅額計算 ( 累加 ) 
			'1.8,000,000~20,000,000   tax:10%
			'2.20,000,000~50,000,000  tax:20%
			'3.50,000,000~80,000,000  tax:30%
			'4.>80,000,000            tax:40%
			F_TAX = 0 
			real_TOTAMT = (CDBL(F1_MONEY)-CDBL(QITA))* CDBL(tmpRec(CurrentPage,index + 1,40))  ' 實領金額
			' if cdbl(real_TOTAMT)>8000000 then 
				' if  cdbl(real_TOTAMT) <=20000000 then 
					' F_tax = ( cdbl(real_TOTAMT) - 8000000 ) * 0.1 
					' response.write "TAX1"&"<BR>"
				' elseif cdbl(real_TOTAMT) > 20000000 and cdbl(cdbl(real_TOTAMT)) <= 50000000 then 
					' F_tax = ( (20000000-8000000)* 0.1 )+((cdbl(real_TOTAMT) - 20000000)*0.2)
					' response.write "TAX2"&"<BR>"
				' elseif cdbl(real_TOTAMT) > 50000000 and cdbl(cdbl(real_TOTAMT)) <= 80000000 then 	
					' F_tax = ((20000000-8000000)* 0.1 )+ ( (50000000-20000000)* 0.2 ) + ((cdbl(real_TOTAMT) - 50000000)*0.3)
					' response.write "TAX3"&"<BR>"
				' elseif cdbl(real_TOTAMT) > 80000000 then 
				  	' F_tax = ((20000000-8000000)* 0.1 )+ ( (50000000-20000000)* 0.2 ) + ( (80000000-50000000)* 0.3 ) + ((cdbl(real_TOTAMT) - 80000000)*0.4)
				  	' response.write "TAX4"&"<BR>"
				' end if 
			' else
				' F_tax = 0  
			' end if 	
			totb = 4000000
				if left(yymm,4)>"2008" then 
					sql2="exec sp_calctax '"& real_TOTAMT &"' ,'"& totb &"' "
					set ors=conn.execute(sql2) 
					F_tax = cdbl(ors("tax"))
				else
					sql2="exec sp_calctax_HW_2008 '"& real_TOTAMT &"' "
					set ors=conn.execute(sql2) 
					F_tax = cdbl(ors("tax"))
				end if  				
				set ors=nothing  			
			response.write F_tax &"<BR>"
			tmpRec(CurrentPage,index + 1,32) = round(cdbl(F_tax) / CDBL(tmpRec(CurrentPage,index + 1,40)),0)
			KTAXM = round(cdbl(F_tax)/CDBL(tmpRec(CurrentPage,index + 1,40)),0)


		F2_MONEY =CDBL(BH)+CDBL(KTAXM)+CDBL(QITA)    '應扣款
		RESPONSE.WRITE "F2_MONEY=" & F2_MONEY &"<BR>"
		TMONEY = CDBL(F1_MONEY) - CDBL(F2_MONEY)
		RESPONSE.WRITE "TMONEY=" & TMONEY &"<BR>"
		tmpRec(CurrentPage,index + 1,36) = TMONEY '實領金額
		
		tmpRec(CurrentPage,index + 1,37)  = CDBL(BB)+CDBL(CV)+CDBL(PHU)+CDBL(KT)+CDBL(MT)+cdbl(JX)+CDBL(TTKH)+CDBL(TNKH)+CDBL(B3M)-CDBL(F2_MONEY) '應領
		RESPONSE.WRITE "RELMONEY=" & tmpRec(CurrentPage,index + 1,37) &"<BR>"
		'不足月扣款
		tmpRec(CurrentPage,index + 1,39) = CDBL(tmpRec(CurrentPage,index + 1,37))-CDBL(tmpRec(CurrentPage,index + 1,36))
  		tmpRec(CurrentPage,index + 1,41) = CDBL(BB)+CDBL(CV)+CDBL(PHU)+CDBL(KT)+CDBL(MT)+cdbl(JX)+CDBL(TTKH)+CDBL(TNKH)+CDBL(B3M) 
  		SUMTOT = CDBL(BB)+CDBL(CV)+CDBL(PHU)+CDBL(KT)+CDBL(MT)+cdbl(JX)+CDBL(TTKH)+CDBL(TNKH)+CDBL(B3M) 
	  	RESPONSE.WRITE "SUMTOT=" & SUMTOT &"<BR>"
		tmpRec(CurrentPage,index + 1,44) = round(fix(tmpRec(CurrentPage,index + 1, 36)),0)
		tmpRec(CurrentPage,index + 1,45) = (round( round(cdbl(tmpRec(CurrentPage, index + 1, 36)),2) - round(fix(tmpRec(CurrentPage, index + 1, 36)),0),2)* cdbl(codestr13)\1000)*1000
		
		if datediff("d",tmpRec(CurrentPage, index + 1, 5),ENDdat)< 180 then 
			dkm = round(TMONEY * 0.25 ,0)
			tmpRec(CurrentPage,index + 1,48) = round(TMONEY * 0.25 ,0)
		else
			dkm = 0
			tmpRec(CurrentPage,index + 1,48) = 0 
		end if 
%>		<script language=vbs>
			Parent.Fore.<%=self%>.HHMOENY(<%=index%>).value=<%=(TTMH)%>
			Parent.Fore.<%=self%>.TNKH(<%=index%>).value=<%=CDBL(TNKH)%>
			Parent.Fore.<%=self%>.KTAXM(<%=index%>).value=<%=(KTAXM)%>
			'Parent.Fore.<%=self%>.NN(<%=index%>).value=<%=(KTAXM)%>
			Parent.Fore.<%=self%>.BZKM(<%=index%>).value=<%=tmpRec(CurrentPage,index + 1,39)%>
			Parent.Fore.<%=self%>.TOTM(<%=index%>).value=<%=SUMTOT%>
			Parent.Fore.<%=self%>.RELTOTMONEY(<%=index%>).value=<%=(TMONEY)%>
			Parent.Fore.<%=self%>.DKM(<%=index%>).value=<%=(DKM)%>
			Parent.Fore.<%=self%>.ZHUANM(<%=index%>).value=<%=tmpRec(CurrentPage,index + 1,44)%>
			Parent.Fore.<%=self%>.XIANM(<%=index%>).value=<%=tmpRec(CurrentPage,index + 1,45)%>
		'	alert <%=TTMH%>
		</script>
<%
    CASE "ZXCHG"
        tmpRec(CurrentPage,index + 1,0)  = "upd"
        tmpRec(CurrentPage,index + 1,44) = CODESTR01  
        tmpRec(CurrentPage,index + 1,45) = CODESTR02 
        response.write tmpRec(CurrentPage,index + 1,44) &"<BR>"
        response.write tmpRec(CurrentPage,index + 1,45) &"<BR>"  
	CASE "memochk"        
		tmpRec(CurrentPage,index + 1,47) = memostr  
		response.write "memochk"        
		response.write "index= " & index + 1  &"<BR>"
		response.write tmpRec(CurrentPage,index + 1,47) &"<BR>"  
	CASE "dkmCHG"        
		tmpRec(CurrentPage,index + 1,48) = CODESTR01  
		response.write "memochk"        
		response.write "index= " & index + 1  &"<BR>"
		response.write tmpRec(CurrentPage,index + 1,48) &"<BR>"  		
end  select
Session("empfilesalaryCN") = tmpRec
%>
</html>

<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<!-- #include file = "../GetSQLServerConnection.fun" --> 
<!-- #include file="../ADOINC.inc" --> 
<%
SELF = "YECE02" 

ftype = request("ftype") 
code = request("code") 
index=request("index")  
CurrentPage = request("CurrentPage") 

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
workdays = REQUEST("days")  
response.write  "workdays=" & workdays &"<BR>" 

yymm=request("yymm")
calcdt = left(YYMM,4)&"/"& right(yymm,2)&"/01" 

tmpRec = Session("empfilesalary") 
response.write "index=" & index &"<BR>"
response.write "ftype=" & ftype &"<BR>"
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
		sql="select * from empsalaryBasic where func='AA' and code='"& code &"'  "
		response.write sql
		'response.end  
		Set rst= Server.CreateObject("ADODB.Recordset")
  		rst.Open SQL, conn, 3,3      	
  		if not rst.eof then  
  			tmpRec(CurrentPage,index + 1,0) = "UPD"
  			tmpRec(CurrentPage,index + 1,19) = code
  			tmpRec(CurrentPage,index + 1,20) = rst("bonus")
  			tmpRec(CurrentPage,index + 1,34) = cdbl(rst("bonus"))*0.06
  			TTM = cdbl(rst("bonus"))+cdbl(tmpRec(CurrentPage, index+1, 22))+cdbl(tmpRec(CurrentPage, index+1, 23)) 
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
	  		NN = tmpRec(CurrentPage,index + 1,24)
	  		KT = tmpRec(CurrentPage,index + 1,25)
	  		MT = tmpRec(CurrentPage,index + 1,26)
	  		TTKH = tmpRec(CurrentPage,index + 1,27)
	  		QC = tmpRec(CurrentPage,index + 1,31)
	  		TNKH = tmpRec(CurrentPage,index + 1,32)
	  		TBTR = tmpRec(CurrentPage,index + 1,33) 
	  		JX = tmpRec(CurrentPage,index + 1,58)  '績效獎金
	  		'-(減項)
	  		BH = tmpRec(CurrentPage,index + 1,34)	  		
	  		HS = tmpRec(CurrentPage,index + 1,35)
	  		GT = tmpRec(CurrentPage,index + 1,36)
	  		QITA = tmpRec(CurrentPage,index + 1,37) 
	  		  
	  		
	  		if cdbl(workdays)<=26 then  
		  		if cdbl(tmpRec(CurrentPage,index + 1,59)) < 3 and tmpRec(CurrentPage,index + 1,5) >=  calcdt   then 
		  			F1_MONEY = 0 
		  		else 
			  		if 	cdbl(tmpRec(CurrentPage,index + 1,59)) < cdbl(workdays) then 
			  			if cdbl(tmpRec(CurrentPage,index + 1,59)) <13 then 
				  			F1_MONEY = round( ( CDBL(BB)+CDBL(CV)+CDBL(PHU)+CDBL(NN)+CDBL(KT)+CDBL(MT)+CDBL(TTKH) )/26 ,0) * cdbl(tmpRec(CurrentPage,index + 1,59)) + CDBL(TNKH)+cdbl(JX)+CDBL(TBTR) 
				  		else
				  			F1_MONEY = round( ( CDBL(BB)+CDBL(CV)+CDBL(PHU)) /26   ,0) *cdbl(tmpRec(CurrentPage,index + 1,59)) +CDBL(NN)+CDBL(KT)+CDBL(MT)+CDBL(TTKH) + CDBL(TNKH)+cdbl(JX)+CDBL(TBTR)  
				  		end if 	
				  		response.write "4"&"<BR>"
				  	else
				  		F1_MONEY = CDBL(BB)+CDBL(CV)+CDBL(PHU)+CDBL(NN)+CDBL(KT)+CDBL(MT)+CDBL(TTKH)+CDBL(QC)+CDBL(TNKH)+cdbl(JX)+CDBL(TBTR)
				  		response.write "3"&"<BR>"
				  	end if 
				end  if   	
			  	response.write "1"&"<BR>"
			 else
			 	if cdbl(tmpRec(CurrentPage,index + 1,59)) < 3 then 
			 		F1_MONEY = 0 
			 	else	
				 	if 	cdbl(tmpRec(CurrentPage,index + 1,59)) < cdbl(workdays) then 
			  			if cdbl(tmpRec(CurrentPage,index + 1,59)) <13 then 
				  			F1_MONEY = round( ( CDBL(BB)+CDBL(CV)+CDBL(PHU)+CDBL(NN)+CDBL(KT)+CDBL(MT)+CDBL(TTKH) )/ cdbl(workdays) ,0) * cdbl(tmpRec(CurrentPage,index + 1,59)) + CDBL(TNKH)+cdbl(JX)+CDBL(TBTR) 
				  		else
				  			F1_MONEY = round( ( CDBL(BB)+CDBL(CV)+CDBL(PHU))/cdbl(workdays)  ,0) *cdbl(tmpRec(CurrentPage,index + 1,59)) +CDBL(NN)+CDBL(KT)+CDBL(MT)+CDBL(TTKH) + CDBL(TNKH)+cdbl(JX)+CDBL(TBTR)  
				  		end if 	
				  	else
				  		F1_MONEY = CDBL(BB)+CDBL(CV)+CDBL(PHU)+CDBL(NN)+CDBL(KT)+CDBL(MT)+CDBL(TTKH)+CDBL(QC)+CDBL(TNKH)+cdbl(JX)+CDBL(TBTR)
				  	end if  
				  	response.write "2"&"<BR>"
				end if 	
			 end if  	
	  		response.write  "F1_MONEY=" & F1_MONEY &"<BR>"	 
	  		F2_MONEY =CDBL(BH)+CDBL(HS)+CDBL(GT)+CDBL(QITA)    '應扣款 
	  		response.write "F2_MONEY=" & F2_MONEY &"<BR>"		
	  		'tmpRec(CurrentPage,index + 1,50) : 曠職事假病假扣款
	  		'tmpRec(CurrentPage,index + 1,49) 總加班費  
	  		
		  	if CDBL(F1_MONEY) = 0 then 
		  		TMONEY = 0 
		  	else
		  		TMONEY = CDBL(F1_MONEY) - CDBL(F2_MONEY)
		  	end if 
	  		
	  		response.write 	 " TMONEY=" & TMONEY &"<BR>"		
	  		tmpRec(CurrentPage,index + 1,39) = TMONEY '應發金額     (不含加班扣減時假)
	  		relTOTM = TMONEY + cdbl(tmpRec(CurrentPage,index + 1,49)) - cdbl(tmpRec(CurrentPage,index + 1,50))  
	  		
  			
  		'個人所得稅計算
	 		F_TAX = 0 
			real_TOTAMT =  relTOTM   ' 實領金額
			'基本額 + 免稅額
			B1 = tmpRec(CurrentPage,index + 1,78) 
			B2 = tmpRec(CurrentPage,index + 1,77) 
			Tot_B = cdbl(B1)+cdbl(B2)			
 	 			
			if real_TOTAMT > cdbl(Tot_B) then 
				if left(yymm,4)>"2008" then 
					sql2="exec sp_calctax '"& real_TOTAMT &"' , '"& cdbl(tot_b) &"' "
					set ors=conn.execute(sql2) 
					f_tax = ors("tax")
				else
					sql2="exec sp_calctax_2008 '"& real_TOTAMT &"' "
					set ors=conn.execute(sql2) 
					f_tax = ors("tax")
				end if  				
				set ors=nothing 			
				tmpRec(CurrentPage, index + 1, 47)= cdbl( relTOTM )- CDBL(ROUND(F_tax,0))		  			  	  		
	  	else
				f_tax = 0 
				tmpRec(CurrentPage, index + 1, 47)= cdbl( relTOTM )- CDBL(ROUND(F_tax,0))		  			  	  		
			end if 	 
   
			tmpRec(CurrentPage,index + 1,64)  = CDBL(BB)+CDBL(CV)+CDBL(PHU)+CDBL(NN)+CDBL(KT)+CDBL(MT)+CDBL(TTKH)+CDBL(QC)+CDBL(TNKH)+cdbl(JX)+CDBL(TBTR)+cdbl(tmpRec(CurrentPage,index + 1,49)) 
  		tmpRec(CurrentPage,index + 1,65) = cdbl(tmpRec(CurrentPage,index + 1,64))-cdbl(F2_MONEY)-cdbl(tmpRec(CurrentPage,index + 1,50))-  CDBL(ROUND(F_tax,0))- cdbl(tmpRec(CurrentPage,index + 1,47))  
 			tmpRec(CurrentPage,index + 1,67) = tmpRec(CurrentPage,index + 1,47) 
 			tmpRec(CurrentPage,index + 1,68) = 0
 			tmpRec(CurrentPage, index + 1, 69) = ROUND(F_tax,0) 				 
	  		
%>			<script language=vbs>								
				Parent.Fore.<%=self%>.bb(<%=index%>).value=<%=rst("bonus")%>
				IF Parent.Fore.<%=self%>.BH(<%=index%>).VALUE>0 THEN 
					Parent.Fore.<%=self%>.BH(<%=index%>).value=<%=cdbl(rst("bonus"))*0.06%>
				END IF 	
				Parent.Fore.<%=self%>.HHMOENY(<%=index%>).value=<%=TTMH%>
				Parent.Fore.<%=self%>.BZKM(<%=index%>).value="<%=formatnumber(tmpRec(CurrentPage,index + 1,65),0)%>"
				Parent.Fore.<%=self%>.TOTM(<%=index%>).value="<%=formatnumber(tmpRec(CurrentPage,index + 1,64),0)%>"
				Parent.Fore.<%=self%>.RELTOTMONEY(<%=index%>).value="<%=formatnumber(relTOTM,0)%>"
				'Parent.Fore.<%=self%>.ZHUANM(<%=index%>).value=<%=tmpRec(CurrentPage,index + 1,67)%>
				'Parent.Fore.<%=self%>.XIANM(<%=index%>).value=<%=tmpRec(CurrentPage,index + 1,68)%>
				Parent.Fore.<%=self%>.ktaxm(<%=index%>).value=<%=tmpRec(CurrentPage,index + 1,69)%>
			</script>
<% 		end if  
		set rs=nothing 		
	case "B"		
		sql="select * from empsalaryBasic where func='BB' and JOB='"& code &"' "
		response.write sql
		'response.end  
		Set rst= Server.CreateObject("ADODB.Recordset")
  		rst.Open SQL, conn, 3,3      	
  		if not rst.eof then  
  			tmpRec(CurrentPage,index + 1,0) = "UPD"
  			tmpRec(CurrentPage,index + 1,6) = code
  			tmpRec(CurrentPage,index + 1,21) = rst("CODE")
  			tmpRec(CurrentPage,index + 1,22) = rst("bonus")
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
		sql2="select * from empsalaryBasic where func='CC' and JOB='"& code &"'  "
		response.write sql2 &"<BR>"
		'response.end  
		Set rst= Server.CreateObject("ADODB.Recordset")
  		rst.Open SQL2, conn, 3,3      	
  		if not rst.eof then  
  			tmpRec(CurrentPage,index + 1,0) = "UPD"
  			tmpRec(CurrentPage,index + 1,31) = rst("bonus")		 
  			
  			'+(加項)
	  		BB = tmpRec(CurrentPage,index + 1,20)
	  		CV = tmpRec(CurrentPage,index + 1,22)
	  		PHU = tmpRec(CurrentPage,index + 1,23)
	  		NN = tmpRec(CurrentPage,index + 1,24)
	  		KT = tmpRec(CurrentPage,index + 1,25)
	  		MT = tmpRec(CurrentPage,index + 1,26)
	  		TTKH = tmpRec(CurrentPage,index + 1,27)
	  		'QC = tmpRec(CurrentPage,index + 1,31)
	  		TNKH = tmpRec(CurrentPage,index + 1,32)
	  		TBTR = tmpRec(CurrentPage,index + 1,33)
	  		JX = tmpRec(CurrentPage,index + 1,58)  '績效獎金
	  		'-(減項)
	  		BH = tmpRec(CurrentPage,index + 1,34)
	  		HS = tmpRec(CurrentPage,index + 1,35)
	  		GT = tmpRec(CurrentPage,index + 1,36)
	  		QITA = tmpRec(CurrentPage,index + 1,37)  
	  		
  			IF 	cdbl(tmpRec(CurrentPage,index + 1,59)) < cdbl(workdays) THEN 
  				QC_MONEY=0
  			ELSE	
		  		IF CDBL( tmpRec(CurrentPage,index + 1,48))>= 6 THEN 
		  			QC_MONEY = 0 
		  		ELSEIF	CDBL( tmpRec(CurrentPage,index + 1,48))>= 3 THEN 
		  			QC_MONEY =  rst("bonus") / 2 
		  		ELSE 
		  			QC_MONEY = rst("bonus")
		  		END IF 
		  	END IF 	  
		  	QC = QC_MONEY 
		  	tmpRec(CurrentPage,index + 1,31) = QC_MONEY 
		  	
		  	if cdbl(workdays)<=26 then  
		  		if cdbl(tmpRec(CurrentPage,index + 1,59)) < 3 and tmpRec(CurrentPage,index + 1,5) >= calcdt  then 
		  			F1_MONEY = 0 
		  		else
			  		if 	cdbl(tmpRec(CurrentPage,index + 1,59)) < cdbl(workdays) then 
			  			if cdbl(tmpRec(CurrentPage,index + 1,59)) <13 then 
				  			F1_MONEY = round( ( CDBL(BB)+CDBL(CV)+CDBL(PHU)+CDBL(NN)+CDBL(KT)+CDBL(MT)+CDBL(TTKH) )/26 ,0) * cdbl(tmpRec(CurrentPage,index + 1,59)) + CDBL(TNKH)+cdbl(JX)+CDBL(TBTR) 
				  		else
				  			F1_MONEY = round( ( CDBL(BB)+CDBL(CV)+CDBL(PHU)) /26   ,0) *cdbl(tmpRec(CurrentPage,index + 1,59)) +CDBL(NN)+CDBL(KT)+CDBL(MT)+CDBL(TTKH) + CDBL(TNKH)+cdbl(JX)+CDBL(TBTR)  
				  		end if 	
				  	else
				  		F1_MONEY = CDBL(BB)+CDBL(CV)+CDBL(PHU)+CDBL(NN)+CDBL(KT)+CDBL(MT)+CDBL(TTKH)+CDBL(QC)+CDBL(TNKH)+cdbl(JX)+CDBL(TBTR)
				  	end if 
				 end if  	
			 else
				if cdbl(tmpRec(CurrentPage,index + 1,59)) < 3 then  
					F1_MONEY = 0 
				else
				 	if 	cdbl(tmpRec(CurrentPage,index + 1,59)) < cdbl(workdays) then 
			  			if cdbl(tmpRec(CurrentPage,index + 1,59)) <13 then 
				  			F1_MONEY = round( ( CDBL(BB)+CDBL(CV)+CDBL(PHU)+CDBL(NN)+CDBL(KT)+CDBL(MT)+CDBL(TTKH) )/ cdbl(workdays) ,0) * cdbl(tmpRec(CurrentPage,index + 1,59)) + CDBL(TNKH)+cdbl(JX)+CDBL(TBTR) 
				  		else
				  			F1_MONEY = round( ( CDBL(BB)+CDBL(CV)+CDBL(PHU))/cdbl(workdays)  ,0) *cdbl(tmpRec(CurrentPage,index + 1,59)) +CDBL(NN)+CDBL(KT)+CDBL(MT)+CDBL(TTKH) + CDBL(TNKH)+cdbl(JX)+CDBL(TBTR)  
				  		end if 	
				  	else
				  		F1_MONEY = CDBL(BB)+CDBL(CV)+CDBL(PHU)+CDBL(NN)+CDBL(KT)+CDBL(MT)+CDBL(TTKH)+CDBL(QC)+CDBL(TNKH)+cdbl(JX)+CDBL(TBTR)
				  	end if  
				end if 
			 end if  
	  		
	  		F2_MONEY =CDBL(BH)+CDBL(HS)+CDBL(GT)+CDBL(QITA)    '應扣款 
	  		
	  		'tmpRec(CurrentPage,index + 1,50) : 曠職事假病假扣款
	  		'tmpRec(CurrentPage,index + 1,49) 總加班費  
	  	  	response.write  F1_MONEY &"<BR>"	  		
	  		response.write  F2_MONEY &"<BR>"		
		  	if CDBL(F1_MONEY) = 0 then 
		  		TMONEY = 0 
		  	else
		  		TMONEY = CDBL(F1_MONEY) - CDBL(F2_MONEY)
		  	end if 
	  		response.write  "TMONEY=" & TMONEY &"<BR>"
	  		tmpRec(CurrentPage,index + 1,39) = TMONEY '應發金額     (不含加班扣減時假)
	  		relTOTM = TMONEY + cdbl(tmpRec(CurrentPage,index + 1,49)) - cdbl(tmpRec(CurrentPage,index + 1,50))  
  			
  		'個人所得稅計算
	  	F_TAX = 0 
			real_TOTAMT =  relTOTM   ' 實領金額
			'基本額 + 免稅額
			B1 = tmpRec(CurrentPage,index + 1,78) 
			B2 = tmpRec(CurrentPage,index + 1,77) 
			Tot_B = cdbl(B1)+cdbl(B2)			 
			if real_TOTAMT > cdbl(Tot_B) then 			
				if left(yymm,4)>"2008" then 
					sql2="exec sp_calctax '"& real_TOTAMT &"', '"& cdbl(tot_b) &"' "
					set ors=conn.execute(sql2) 
					f_tax = ors("tax")
				else
					sql2="exec sp_calctax_2008 '"& real_TOTAMT &"' "
					set ors=conn.execute(sql2) 
					f_tax = ors("tax")
				end if  				
				set ors=nothing 								
	  	else
				f_tax =0 
			end if 
			
			tmpRec(CurrentPage,index + 1,47)  = relTOTM - CDBL(ROUND(F_tax,0))  '實領金額(含加班扣減時假-所得稅)       			  	  		
	  	tmpRec(CurrentPage,index + 1,64)  = CDBL(BB)+CDBL(CV)+CDBL(PHU)+CDBL(NN)+CDBL(KT)+CDBL(MT)+CDBL(TTKH)+CDBL(QC)+CDBL(TNKH)+cdbl(JX)+CDBL(TBTR)+cdbl(tmpRec(CurrentPage,index + 1,49)) 
	  	tmpRec(CurrentPage,index + 1,65) = cdbl(tmpRec(CurrentPage,index + 1,64))-cdbl(F2_MONEY)-cdbl(tmpRec(CurrentPage,index + 1,50))-CDBL(ROUND(F_tax,0)) - cdbl(tmpRec(CurrentPage,index + 1,47)) 
  		tmpRec(CurrentPage,index + 1,67) = tmpRec(CurrentPage,index + 1,47) 
  		tmpRec(CurrentPage,index + 1,68) = 0   		
  		tmpRec(CurrentPage, index + 1, 69) = ROUND(F_tax,0)   
  			
%>			<script language=vbs>												
				Parent.Fore.<%=self%>.QC(<%=index%>).value=<%=QC_MONEY%>				 
				'Parent.Fore.<%=self%>.TOTMONEY(<%=index%>).value=<%=(TMONEY)%> 
				Parent.Fore.<%=self%>.RELTOTMONEY(<%=index%>).value="<%=formatnumber(relTOTM,0)%>"
				'Parent.Fore.<%=self%>.TOTM(<%=index%>).value=<%=F1_MONEY%> 
				Parent.Fore.<%=self%>.BZKM(<%=index%>).value="<%=formatnumber(tmpRec(CurrentPage,index + 1,65),0)%>"
				Parent.Fore.<%=self%>.TOTM(<%=index%>).value="<%=formatnumber(tmpRec(CurrentPage,index + 1,64),0)%>"
				'Parent.Fore.<%=self%>.ZHUANM(<%=index%>).value=<%=tmpRec(CurrentPage,index + 1,67)%>
				'Parent.Fore.<%=self%>.XIANM(<%=index%>).value=<%=tmpRec(CurrentPage,index + 1,68)%>
				Parent.Fore.<%=self%>.ktaxm(<%=index%>).value=<%=tmpRec(CurrentPage,index + 1,69)%>
			</script>
<% 		end if   
		set rs=nothing   
		
	case "CDATACHG"		
		tmpRec(CurrentPage,index + 1,0) = "UPD"		   		
  		tmpRec(CurrentPage,index + 1,23) = CODESTR01
  		tmpRec(CurrentPage,index + 1,24) = CODESTR02
  		tmpRec(CurrentPage,index + 1,25) = CODESTR03
  		tmpRec(CurrentPage,index + 1,26) = CODESTR04
  		tmpRec(CurrentPage,index + 1,27) = CODESTR05
  		tmpRec(CurrentPage,index + 1,32) = CODESTR06
  		tmpRec(CurrentPage,index + 1,35) = CODESTR07
  		tmpRec(CurrentPage,index + 1,37) = CODESTR08
  		tmpRec(CurrentPage,index + 1,58) = CODESTR09
  		tmpRec(CurrentPage,index + 1,34) = CODESTR10
  		tmpRec(CurrentPage,index + 1,36) = CODESTR11
  		TTM = cdbl(tmpRec(CurrentPage, index+1, 20))+cdbl(tmpRec(CurrentPage, index+1, 22))+cdbl(tmpRec(CurrentPage, index+1, 23)) 
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
  		NN = tmpRec(CurrentPage,index + 1,24)
  		KT = tmpRec(CurrentPage,index + 1,25)
  		MT = tmpRec(CurrentPage,index + 1,26)
  		TTKH = tmpRec(CurrentPage,index + 1,27)
  		QC = tmpRec(CurrentPage,index + 1,31)
  		TNKH = tmpRec(CurrentPage,index + 1,32)
  		TBTR = tmpRec(CurrentPage,index + 1,33)
  		JX = tmpRec(CurrentPage,index + 1,58)  '績效獎金 
  		
  		response.write   "BB=" & BB &"<BR>"
  		response.write   "CV=" & CV &"<BR>"
  		response.write   "PHU=" & PHU &"<BR>"
  		response.write   "NN=" & NN &"<BR>"
  		response.write   "KT=" & KT &"<BR>"
  		response.write   "MT=" & MT &"<BR>"
  		response.write   "TTKH=" & TTKH &"<BR>"
  		'-(減項)
  		BH = tmpRec(CurrentPage,index + 1,34)
  		HS = tmpRec(CurrentPage,index + 1,35)
  		GT = tmpRec(CurrentPage,index + 1,36)
  		QITA = tmpRec(CurrentPage,index + 1,37) 
  		
  		if cdbl(workdays)<=26 then 
		  	if cdbl(tmpRec(CurrentPage,index + 1,59)) < 3 and tmpRec(CurrentPage,index + 1,5) >= calcdt  then 
		  		F1_MONEY = 0 
		  	else  		
		  		if 	cdbl(tmpRec(CurrentPage,index + 1,59)) < cdbl(workdays) then 
		  			if cdbl(tmpRec(CurrentPage,index + 1,59)) <13 then 
			  			F1_MONEY = round( ( CDBL(BB)+CDBL(CV)+CDBL(PHU)+CDBL(NN)+CDBL(KT)+CDBL(MT)+CDBL(TTKH)+CDBL(QC) )/26 ,0) * cdbl(tmpRec(CurrentPage,index + 1,59)) + CDBL(TNKH)+cdbl(JX)+CDBL(TBTR) 
			  		else
			  			F1_MONEY = round( ( CDBL(BB)+CDBL(CV)+CDBL(PHU) )/26 ,0) *cdbl(tmpRec(CurrentPage,index + 1,59)) +CDBL(NN)+CDBL(KT)+CDBL(MT)+CDBL(TTKH) + CDBL(TNKH)+cdbl(JX)+CDBL(TBTR)+CDBL(QC)
			  		end if 	
			  	else
			  		F1_MONEY = CDBL(BB)+CDBL(CV)+CDBL(PHU)+CDBL(NN)+CDBL(KT)+CDBL(MT)+CDBL(TTKH)+CDBL(QC)+CDBL(TNKH)+cdbl(JX)+CDBL(TBTR)
			  	end if 
			end if   	
		else
		  	if cdbl(tmpRec(CurrentPage,index + 1,59)) < 3 then 
		  		F1_MONEY = 0 
		  	else  
				if 	cdbl(tmpRec(CurrentPage,index + 1,59)) < cdbl(workdays) then 
		  			if cdbl(tmpRec(CurrentPage,index + 1,59)) <13 then 
			  			F1_MONEY = round( ( CDBL(BB)+CDBL(CV)+CDBL(PHU)+CDBL(NN)+CDBL(KT)+CDBL(MT)+CDBL(TTKH)+CDBL(QC) )/cdbl(workdays) ,0) * cdbl(tmpRec(CurrentPage,index + 1,59)) + CDBL(TNKH)+cdbl(JX)+CDBL(TBTR) 
			  		else
			  			F1_MONEY = round( ( CDBL(BB)+CDBL(CV)+CDBL(PHU) )/cdbl(workdays) ,0) *cdbl(tmpRec(CurrentPage,index + 1,59)) +CDBL(NN)+CDBL(KT)+CDBL(MT)+CDBL(TTKH) + CDBL(TNKH)+cdbl(JX)+CDBL(TBTR)+CDBL(QC)
			  		end if 	
			  	else
			  		F1_MONEY = CDBL(BB)+CDBL(CV)+CDBL(PHU)+CDBL(NN)+CDBL(KT)+CDBL(MT)+CDBL(TTKH)+CDBL(QC)+CDBL(TNKH)+cdbl(JX)+CDBL(TBTR)
			  	end if 
			end if 			  	
		end if   	
	  		
	  	RESPONSE.WRITE 	"F1_MONEY=" & F1_MONEY &"<br>"
	  	tmpRec(CurrentPage,index + 1,60) = CDBL(F1_MONEY)
	  	
	  	F2_MONEY =CDBL(BH)+CDBL(HS)+CDBL(GT)+CDBL(QITA)    '應扣款  	  	
	  	RESPONSE.WRITE 	"F2_MONEY=" & F2_MONEY &"<br>"	
	  	
	  	'tmpRec(CurrentPage,index + 1,50) : 曠職事假病假扣款
	  	'tmpRec(CurrentPage,index + 1,49) 總加班費  
	  	if CDBL(F1_MONEY) = 0 then 
	  		TMONEY = 0 
	  	else
	  		TMONEY = CDBL(F1_MONEY) - CDBL(F2_MONEY)
	  	end if 	
	  	tmpRec(CurrentPage,index + 1,39) = TMONEY '應發金額     (不含加班扣減時假)
	  	RESPONSE.WRITE "TMONEY=" & TMONEY &"<br>"
	  	relTOTM = TMONEY + cdbl(tmpRec(CurrentPage,index + 1,49)) - cdbl(tmpRec(CurrentPage,index + 1,50))  
  		
  		'個人所得稅計算
	  	F_TAX = 0 
			real_TOTAMT =  relTOTM   ' 實領金額
			'基本額 + 免稅額
			B1 = tmpRec(CurrentPage,index + 1,78) 
			B2 = tmpRec(CurrentPage,index + 1,77) 
			Tot_B = cdbl(B1)+cdbl(B2)			 
			if real_TOTAMT > cdbl(Tot_B) then 			
				if left(yymm,4)>"2008" then 
					sql2="exec sp_calctax '"& real_TOTAMT &"' , '"& cdbl(tot_b) &"'"
					set ors=conn.execute(sql2) 
					f_tax = ors("tax")
				else
					sql2="exec sp_calctax_2008 '"& real_TOTAMT &"' "
					set ors=conn.execute(sql2) 
					f_tax = ors("tax")
				end if  				
				set ors=nothing 								
	  	else
				f_tax =0 
			end if 
			
			tmpRec(CurrentPage,index + 1,47)  = relTOTM - CDBL(ROUND(F_tax,0))  '實領金額(含加班扣減時假-所得稅)    
			  
  		tmpRec(CurrentPage,index + 1,64)  = CDBL(BB)+CDBL(CV)+CDBL(PHU)+CDBL(NN)+CDBL(KT)+CDBL(MT)+CDBL(TTKH)+CDBL(QC)+CDBL(TNKH)+cdbl(JX)+CDBL(TBTR)+cdbl(tmpRec(CurrentPage,index + 1,49)) 
	  	tmpRec(CurrentPage,index + 1,65) = cdbl(tmpRec(CurrentPage,index + 1,64))-cdbl(F2_MONEY)-cdbl(tmpRec(CurrentPage,index + 1,50))- CDBL(ROUND(F_tax,0)) - cdbl(tmpRec(CurrentPage,index + 1,47)) 	
  		tmpRec(CurrentPage,index + 1,67) = tmpRec(CurrentPage,index + 1,47) 
  		tmpRec(CurrentPage,index + 1,68) = 0
  		tmpRec(CurrentPage,index + 1,69) = ROUND(F_TAX,0)
  		
%>		<script language=vbs>														
			Parent.Fore.<%=self%>.HHMOENY(<%=index%>).value=<%=(TTMH)%>
			'Parent.Fore.<%=self%>.TOTMONEY(<%=index%>).value=<%=(TMONEY)%>			
			'Parent.Fore.<%=self%>.TOTM(<%=index%>).value=<%=(F1_MONEY)%> 
			Parent.Fore.<%=self%>.BZKM(<%=index%>).value="<%=formatnumber(tmpRec(CurrentPage,index + 1,65),0)%>" 
			Parent.Fore.<%=self%>.TOTM(<%=index%>).value="<%=formatnumber(tmpRec(CurrentPage,index + 1,64),0)%>"
			Parent.Fore.<%=self%>.RELTOTMONEY(<%=index%>).value="<%=formatnumber(tmpRec(CurrentPage,index + 1,47) ,0)%>"
			'Parent.Fore.<%=self%>.ZHUANM(<%=index%>).value=<%=tmpRec(CurrentPage,index + 1,67)%>
			'Parent.Fore.<%=self%>.XIANM(<%=index%>).value=<%=tmpRec(CurrentPage,index + 1,68)%>
			Parent.Fore.<%=self%>.KTAXM(<%=index%>).value=<%=tmpRec(CurrentPage,index + 1,69)%>
		'	alert <%=TTMH%>
		</script>
<%
	CASE "ZXCHG"
        tmpRec(CurrentPage,index + 1,0) = "UPD"		   		
        tmpRec(CurrentPage,index + 1,67) = CODESTR01  
        tmpRec(CurrentPage,index + 1,68) = CODESTR02 
        response.write tmpRec(CurrentPage,index + 1,67) &"<BR>"
        response.write tmpRec(CurrentPage,index + 1,68) &"<BR>"
end  select   		
Session("empfilesalary") = tmpRec
%>
</html>

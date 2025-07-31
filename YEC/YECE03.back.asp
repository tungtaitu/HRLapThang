<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<!-- #include file = "../GetSQLServerConnection.fun" --> 
<!-- #include file="../ADOINC.inc" --> 
<%
SELF = "YECE03" 

ftype = request("ftype") 
code = request("code") 
index=request("index")  
CurrentPage = request("CurrentPage") 
yymm = request("yymm")
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
memostr = request("memo")
response.write  "workdays=" & workdays &"<BR>"

tmpRec = Session("YECE03") 
response.write "index=" & index &"<BR>"
response.write "ftype=" & ftype &"<BR>" 
exrt = request("exrt")  

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
		sql="select * from empsalaryBasic where func='AA' and code='"& code &"' and country='"& tmpRec(CurrentPage,index + 1,4) &"' "
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
	  		
	  		
	  		
	  		'if cdbl(workdays)<=26 then  
		  	'	if 	cdbl(tmpRec(CurrentPage,index + 1,59)) < cdbl(workdays) then 
		  	'		if cdbl(tmpRec(CurrentPage,index + 1,59)) <13 then 
			'  			F1_MONEY = round( ( CDBL(BB)+CDBL(CV)+CDBL(PHU)+CDBL(NN)+CDBL(KT)+CDBL(MT)+CDBL(TTKH) )/26 ,0) * cdbl(tmpRec(CurrentPage,index + 1,59)) + CDBL(TNKH)+cdbl(JX)+CDBL(TBTR) 
			'  		else
			'  			F1_MONEY = round( ( CDBL(BB)+CDBL(CV)+CDBL(PHU)) /26   ,0) *cdbl(tmpRec(CurrentPage,index + 1,59)) +CDBL(NN)+CDBL(KT)+CDBL(MT)+CDBL(TTKH) + CDBL(TNKH)+cdbl(JX)+CDBL(TBTR)  
			'  		end if 	
			'  	else
			'  		F1_MONEY = CDBL(BB)+CDBL(CV)+CDBL(PHU)+CDBL(NN)+CDBL(KT)+CDBL(MT)+CDBL(TTKH)+CDBL(QC)+CDBL(TNKH)+cdbl(JX)+CDBL(TBTR)
			'  	end if 
			' else
			 	if 	cdbl(tmpRec(CurrentPage,index + 1,59)) < cdbl(workdays) then 
		  			if cdbl(tmpRec(CurrentPage,index + 1,59)) <13 then 
			  			F1_MONEY = round( ( CDBL(BB)+CDBL(CV)+CDBL(PHU)+CDBL(NN)+CDBL(KT)+CDBL(MT)+CDBL(TTKH) )/ cdbl(workdays) ,0) * cdbl(tmpRec(CurrentPage,index + 1,59)) + CDBL(TNKH)+cdbl(JX)+CDBL(TBTR) 
			  		else
			  			F1_MONEY = round( ( CDBL(BB)+CDBL(CV)+CDBL(PHU))/cdbl(workdays)  ,0) *cdbl(tmpRec(CurrentPage,index + 1,59)) +CDBL(NN)+CDBL(KT)+CDBL(MT)+CDBL(TTKH) + CDBL(TNKH)+cdbl(JX)+CDBL(TBTR)  
			  		end if 	
			  	else
			  		F1_MONEY = CDBL(BB)+CDBL(CV)+CDBL(PHU)+CDBL(NN)+CDBL(KT)+CDBL(MT)+CDBL(TTKH)+CDBL(QC)+CDBL(TNKH)+cdbl(JX)+CDBL(TBTR)
			  	end if  
			' end if  	
	  		
	  		F2_MONEY =CDBL(BH)+CDBL(HS)+CDBL(GT)+CDBL(QITA)    '應扣款 
	  		
	  		'tmpRec(CurrentPage,index + 1,50) : 曠職事假病假扣款
	  		'tmpRec(CurrentPage,index + 1,49) 總加班費  
	  		
	  		TMONEY = F1_TMONEY - F2_MONEY  	  		
	  		tmpRec(CurrentPage,index + 1,39) = TMONEY '應發金額     (不含加班扣減時假)
	  		relTOTM = TMONEY + cdbl(tmpRec(CurrentPage,index + 1,49)) - cdbl(tmpRec(CurrentPage,index + 1,50))  
  			tmpRec(CurrentPage,index + 1,47)  = relTOTM  '實領金額(含加班扣減時假)
	  		tmpRec(CurrentPage,index + 1,64)  = CDBL(BB)+CDBL(CV)+CDBL(PHU)+CDBL(NN)+CDBL(KT)+CDBL(MT)+CDBL(TTKH)+CDBL(QC)+CDBL(TNKH)+cdbl(JX)+CDBL(TBTR)+cdbl(tmpRec(CurrentPage,index + 1,49)) 
	  		tmpRec(CurrentPage,index + 1,65) = cdbl(tmpRec(CurrentPage,index + 1,64))-cdbl(F2_MONEY)-cdbl(tmpRec(CurrentPage,index + 1,50))-cdbl(tmpRec(CurrentPage,index + 1,47)) 
  			tmpRec(CurrentPage,index + 1,67) = tmpRec(CurrentPage,index + 1,47) 
  			tmpRec(CurrentPage,index + 1,68) = 0
	  		
%>			<script language=vbs>								
				Parent.Fore.<%=self%>.bb(<%=index%>).value=<%=rst("bonus")%>
				IF Parent.Fore.<%=self%>.BH(<%=index%>).VALUE>0 THEN 
					Parent.Fore.<%=self%>.BH(<%=index%>).value=<%=cdbl(rst("bonus"))*0.06%>
				END IF 	
				Parent.Fore.<%=self%>.HHMOENY(<%=index%>).value=<%=TTMH%>
				Parent.Fore.<%=self%>.BZKM(<%=index%>).value=<%=tmpRec(CurrentPage,index + 1,65)%> 
				Parent.Fore.<%=self%>.TOTM(<%=index%>).value=<%=tmpRec(CurrentPage,index + 1,64)%>  
				Parent.Fore.<%=self%>.RELTOTMONEY(<%=index%>).value=<%=(relTOTM)%> 
				Parent.Fore.<%=self%>.ZHUANM(<%=index%>).value=<%=tmpRec(CurrentPage,index + 1,67)%>
				Parent.Fore.<%=self%>.XIANM(<%=index%>).value=<%=tmpRec(CurrentPage,index + 1,68)%>
			</script>
<% 		end if  
		set rs=nothing 		
	case "B"		
		sql="select * from empsalaryBasic where func='BB' and JOB='"& code 
		response.write sql
		'response.end  &"'  "
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
		response.write sql2
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
		  	
		  	'if cdbl(workdays)<=26 then  
		  	'	if 	cdbl(tmpRec(CurrentPage,index + 1,59)) < cdbl(workdays) then 
		  	'		if cdbl(tmpRec(CurrentPage,index + 1,59)) <13 then 
			'  			F1_MONEY = round( ( CDBL(BB)+CDBL(CV)+CDBL(PHU)+CDBL(NN)+CDBL(KT)+CDBL(MT)+CDBL(TTKH) )/26 ,0) * cdbl(tmpRec(CurrentPage,index + 1,59)) + CDBL(TNKH)+cdbl(JX)+CDBL(TBTR) 
			'  		else
			'  			F1_MONEY = round( ( CDBL(BB)+CDBL(CV)+CDBL(PHU)) /26   ,0) *cdbl(tmpRec(CurrentPage,index + 1,59)) +CDBL(NN)+CDBL(KT)+CDBL(MT)+CDBL(TTKH) + CDBL(TNKH)+cdbl(JX)+CDBL(TBTR)  
			'  		end if 	
			'  	else
			'  		F1_MONEY = CDBL(BB)+CDBL(CV)+CDBL(PHU)+CDBL(NN)+CDBL(KT)+CDBL(MT)+CDBL(TTKH)+CDBL(QC)+CDBL(TNKH)+cdbl(JX)+CDBL(TBTR)
			'  	end if 
			' else
			 	if 	cdbl(tmpRec(CurrentPage,index + 1,59)) < cdbl(workdays) then 
		  			if cdbl(tmpRec(CurrentPage,index + 1,59)) <13 then 
			  			F1_MONEY = round( ( CDBL(BB)+CDBL(CV)+CDBL(PHU)+CDBL(NN)+CDBL(KT)+CDBL(MT)+CDBL(TTKH) )/ cdbl(workdays) ,0) * cdbl(tmpRec(CurrentPage,index + 1,59)) + CDBL(TNKH)+cdbl(JX)+CDBL(TBTR) 
			  		else
			  			F1_MONEY = round( ( CDBL(BB)+CDBL(CV)+CDBL(PHU))/cdbl(workdays)  ,0) *cdbl(tmpRec(CurrentPage,index + 1,59)) +CDBL(NN)+CDBL(KT)+CDBL(MT)+CDBL(TTKH) + CDBL(TNKH)+cdbl(JX)+CDBL(TBTR)  
			  		end if 	
			  	else
			  		F1_MONEY = CDBL(BB)+CDBL(CV)+CDBL(PHU)+CDBL(NN)+CDBL(KT)+CDBL(MT)+CDBL(TTKH)+CDBL(QC)+CDBL(TNKH)+cdbl(JX)+CDBL(TBTR)
			  	end if  
			' end if  
	  		
	  		F2_MONEY =CDBL(BH)+CDBL(HS)+CDBL(GT)+CDBL(QITA)    '應扣款 
	  		
	  		'tmpRec(CurrentPage,index + 1,50) : 曠職事假病假扣款
	  		'tmpRec(CurrentPage,index + 1,49) 總加班費  
	  		
	  		TMONEY = F1_TMONEY - F2_MONEY  	  		
	  		tmpRec(CurrentPage,index + 1,39) = TMONEY '應發金額     (不含加班扣減時假)
	  		relTOTM = TMONEY + cdbl(tmpRec(CurrentPage,index + 1,49)) - cdbl(tmpRec(CurrentPage,index + 1,50))  
  			tmpRec(CurrentPage,index + 1,47)  = relTOTM  '實領金額(含加班扣減時假) 
  			tmpRec(CurrentPage,index + 1,64)  = CDBL(BB)+CDBL(CV)+CDBL(PHU)+CDBL(NN)+CDBL(KT)+CDBL(MT)+CDBL(TTKH)+CDBL(QC)+CDBL(TNKH)+cdbl(JX)+CDBL(TBTR)+cdbl(tmpRec(CurrentPage,index + 1,49)) 
	  		tmpRec(CurrentPage,index + 1,65) = cdbl(tmpRec(CurrentPage,index + 1,64))-cdbl(F2_MONEY)-cdbl(tmpRec(CurrentPage,index + 1,50))-cdbl(tmpRec(CurrentPage,index + 1,47))
  			tmpRec(CurrentPage,index + 1,67) = tmpRec(CurrentPage,index + 1,47) 
  			tmpRec(CurrentPage,index + 1,68) = 0
%>			<script language=vbs>												
				Parent.Fore.<%=self%>.QC(<%=index%>).value=<%=QC_MONEY%>				 
				'Parent.Fore.<%=self%>.TOTMONEY(<%=index%>).value=<%=TMONEY%> 
				Parent.Fore.<%=self%>.RELTOTMONEY(<%=index%>).value=<%=relTOTM%>
				'Parent.Fore.<%=self%>.TOTM(<%=index%>).value=<%=F1_MONEY%> 
				Parent.Fore.<%=self%>.BZKM(<%=index%>).value=<%=tmpRec(CurrentPage,index + 1,65)%> 
				Parent.Fore.<%=self%>.TOTM(<%=index%>).value=<%=tmpRec(CurrentPage,index + 1,64)%> 
				Parent.Fore.<%=self%>.ZHUANM(<%=index%>).value=<%=tmpRec(CurrentPage,index + 1,67)%>
				Parent.Fore.<%=self%>.XIANM(<%=index%>).value=<%=tmpRec(CurrentPage,index + 1,68)%>
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
  		allJBM = tmpRec(CurrentPage,index + 1,49)   '全部加班費
  		response.write   "BB=" & BB &"<BR>"
  		response.write   "CV=" & CV &"<BR>"
  		response.write   "PHU=" & PHU &"<BR>"
  		response.write   "NN=" & NN &"<BR>"
  		response.write   "KT=" & KT &"<BR>"
  		response.write   "MT=" & MT &"<BR>"
  		response.write   "TTKH=" & TTKH &"<BR>"
			response.write   "jx=" & jx &"<BR>"
			response.write   "allJBM=" & allJBM &"<BR>"
  		'-(減項)
  		BH = tmpRec(CurrentPage,index + 1,34)  		
  		GT = tmpRec(CurrentPage,index + 1,36)
  		QITA = tmpRec(CurrentPage,index + 1,37)  
			khhM = cdbl(tmpRec(CurrentPage,index + 1,50))   '扣時假 
  		all_KM = cdbl(bh)+cdbl(GT)+cdbl(qita)+cdbl(khhm)
  		'if cdbl(workdays)<=26 then 
	  	'	if 	cdbl(tmpRec(CurrentPage,index + 1,59)) < cdbl(workdays) then 
	  	'		if cdbl(tmpRec(CurrentPage,index + 1,59)) <13 then 
		'  			F1_MONEY = round( ( CDBL(BB)+CDBL(CV)+CDBL(PHU)+CDBL(NN)+CDBL(KT)+CDBL(MT)+CDBL(TTKH)+CDBL(QC) )/26 ,0) * cdbl(tmpRec(CurrentPage,index + 1,59)) + CDBL(TNKH)+cdbl(JX)+CDBL(TBTR) 
		'  		else
		'  			F1_MONEY = round( ( CDBL(BB)+CDBL(CV)+CDBL(PHU) )/26 ,0) *cdbl(tmpRec(CurrentPage,index + 1,59)) +CDBL(NN)+CDBL(KT)+CDBL(MT)+CDBL(TTKH) + CDBL(TNKH)+cdbl(JX)+CDBL(TBTR)+CDBL(QC)
		'  		end if 	
		'  	else
		'  		F1_MONEY = CDBL(BB)+CDBL(CV)+CDBL(PHU)+CDBL(NN)+CDBL(KT)+CDBL(MT)+CDBL(TTKH)+CDBL(QC)+CDBL(TNKH)+cdbl(JX)+CDBL(TBTR)
		'  	end if 
		'else
		  response.write "days = "& tmpRec(CurrentPage,index + 1,59) &"<BR>"
			if 	cdbl(tmpRec(CurrentPage,index + 1,59)) < cdbl(workdays) then 
	  			if cdbl(tmpRec(CurrentPage,index + 1,59)) <13 then 
						response.write "X1"
		  			F1_MONEY = round( ( CDBL(BB)+CDBL(CV)+CDBL(PHU)+CDBL(NN)+CDBL(KT)+CDBL(MT) )/cdbl(workdays)  * cdbl(tmpRec(CurrentPage,index + 1,59)),0) +CDBL(TTKH)+CDBL(TNKH)+cdbl(JX)+CDBL(TBTR)+cdbl(allJBM) 
		  		else
						response.write "X2"
		  			F1_MONEY = round( ( CDBL(BB) +CDBL(CV)+CDBL(PHU)+CDBL(NN)+CDBL(KT)+CDBL(MT) )/cdbl(workdays) * cdbl(tmpRec(CurrentPage,index + 1,59)),0) +CDBL(TTKH)+CDBL(TNKH)+cdbl(JX)+CDBL(TBTR)+cdbl(allJBM) 
		  		end if 	
		  	else
					response.write "X3"
		  		F1_MONEY = CDBL(BB)+CDBL(CV)+CDBL(PHU)+CDBL(NN)+CDBL(KT)+CDBL(MT)+CDBL(TTKH)+CDBL(QC)+CDBL(TNKH)+cdbl(JX)+CDBL(TBTR)+cdbl(allJBM)
		  	end if 
		'end if   	
	  		
	  	RESPONSE.WRITE 	"F1_MONEY=" & F1_MONEY &"<br>"
	  	tmpRec(CurrentPage,index + 1,60) = CDBL(F1_MONEY)
	  	
	  	F2_MONEY =CDBL(BH)+CDBL(HS)+CDBL(GT)+CDBL(QITA)+cdbl(khhm)    '應扣款  	  	
	  	RESPONSE.WRITE 	"F2_MONEY=" & F2_MONEY &"<br>"	
	  	
	  	'tmpRec(CurrentPage,index + 1,50) : 曠職事假病假扣款
	  	'tmpRec(CurrentPage,index + 1,49) 總加班費  
	  	
	  	TMONEY = CDBL(F1_MONEY) - CDBL(F2_MONEY)
	  	tmpRec(CurrentPage,index + 1,39) = TMONEY '應領金額  
	  	RESPONSE.WRITE "TMONEY=" & TMONEY &"<br>"
	  	
  		
  		tmpRec(CurrentPage,index + 1,64)  =  CDBL(BB)+CDBL(CV)+CDBL(PHU)+CDBL(NN)+CDBL(KT)+CDBL(MT)+CDBL(TTKH)+CDBL(QC)+CDBL(TNKH)+cdbl(JX)+CDBL(TBTR)+cdbl(tmpRec(CurrentPage,index + 1,49)) 
  		'tmpRec(CurrentPage,index + 1,64)  =  CDBL(F1_MONEY) +  cdbl(tmpRec(CurrentPage,index + 1,49))  
			response.write  "totM (64)=" & tmpRec(CurrentPage,index + 1,64) &"<BR>"
			
			'relTOTM = TMONEY  
			exrt = tmpRec(CurrentPage,index + 1,71)
			response.write "exrt=" & exrt 
			'response.end 
			real_TOTAMT = (CDBL(TMONEY))* CDBL(exrt)  ' 實領金額 
			response.write "應領金額(VND)=" & real_TOTAMT &"<BR>"
 
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
			response.write "F_tax="& F_tax &"<br>" 
			tmpRec(CurrentPage,index + 1,66) = round(cdbl(F_tax) / cdbl(exrt),0)
			KTAXM = round(cdbl(F_tax)/cdbl(exrt),0)
			response.write "KTAXM="& KTAXM &"<br>"  
			relTOTM  = cdbl(TMONEY) - cdbl(KTAXM)
			
	  	tmpRec(CurrentPage,index + 1,47)  = cdbl(relTOTM)    '實領金額(含加班扣減時假)   	 
			tmpRec(CurrentPage,index + 1,65) = cdbl(tmpRec(CurrentPage,index + 1,64)) - cdbl(tmpRec(CurrentPage,index + 1,47))-(cdbl(KTAXM)+cdbl(qita)+cdbl(bh)+cdbl(khhM))
  		tmpRec(CurrentPage,index + 1,67) = tmpRec(CurrentPage,index + 1,47) 
  		tmpRec(CurrentPage,index + 1,68) = 0
%>		<script language=vbs>														
			Parent.Fore.<%=self%>.HHMOENY(<%=index%>).value=<%=(TTMH)%>
			'Parent.Fore.<%=self%>.TOTMONEY(<%=index%>).value=<%=(TMONEY)%>			
			'Parent.Fore.<%=self%>.TOTM(<%=index%>).value=<%=(F1_MONEY)%> 
			Parent.Fore.<%=self%>.BZKM(<%=index%>).value=<%=tmpRec(CurrentPage,index + 1,65)%> 
			Parent.Fore.<%=self%>.TOTM(<%=index%>).value=<%=tmpRec(CurrentPage,index + 1,64)%>   
			Parent.Fore.<%=self%>.RELTOTMONEY(<%=index%>).value=<%=(relTOTM)%>
			Parent.Fore.<%=self%>.ZHUANM(<%=index%>).value=<%=tmpRec(CurrentPage,index + 1,67)%>
			Parent.Fore.<%=self%>.XIANM(<%=index%>).value=<%=tmpRec(CurrentPage,index + 1,68)%>
			Parent.Fore.<%=self%>.ktaxM(<%=index%>).value=<%=KTAXM%>
		'	alert <%=TTMH%>
		</script>
<%
	CASE "ZXCHG"
        tmpRec(CurrentPage,index + 1,0) = "UPD"		   		
        tmpRec(CurrentPage,index + 1,67) = CODESTR01  
        tmpRec(CurrentPage,index + 1,68) = CODESTR02 
        response.write tmpRec(CurrentPage,index + 1,67) &"<BR>"
        response.write tmpRec(CurrentPage,index + 1,68) &"<BR>" 

	CASE "memochk"        
		tmpRec(CurrentPage,index + 1,69) = memostr  
		response.write "memochk"        
		response.write "index= " & index + 1  &"<BR>"
		response.write tmpRec(CurrentPage,index + 1,69) &"<BR>"  
		        
end  select   		
response.write "32=" & tmpRec(CurrentPage,index + 1,32)  &"<br>"
response.write "47=" & tmpRec(CurrentPage,index + 1,47) &"<br>"
Session("YECE03") = tmpRec
%>
</html>

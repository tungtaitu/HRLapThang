<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<!-- #include file = "../GetSQLServerConnection.fun" --> 
<!-- #include file="../ADOINC.inc" --> 
<%
SELF = "YECE12" 

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
CODESTR12 = REQUEST("CODESTR12") 
CODESTR13 = REQUEST("CODESTR13") 
CODESTR14 = REQUEST("CODESTR14") 
workdays = REQUEST("days") 
actdays = REQUEST("actdays")  

response.write  "workdays=" & workdays &"<BR>"  
yymm=request("yymm") 
rate = request("rate")  

calcdt = left(YYMM,4)&"/"& right(yymm,2)&"/01" 

tmpRec = Session("yece12B") 
response.write "index=" & index &"<BR>"
response.write "ftype=" & ftype &"<BR>"
response.write "CurrentPage=" & CurrentPage &"<BR>"

Set conn = GetSQLServerConnection()	  
'response.end 
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
</head>
<%
select case ftype    
	case "CDATACHG"		
		tmpRec(CurrentPage,index + 1,0) = "UPD"		   		
  		tmpRec(CurrentPage,index + 1,28) = CODESTR01  '其他收入
  		tmpRec(CurrentPage,index + 1,29) = CODESTR02	'jx
  		tmpRec(CurrentPage,index + 1,34) = CODESTR03	'- 其他  
			tmpRec(CurrentPage,index + 1,51)  = CODESTR04	' 總加班費
			tmpRec(CurrentPage,index + 1,57) = CODESTR05	'h1
			tmpRec(CurrentPage,index + 1,58) = CODESTR06	'
			tmpRec(CurrentPage,index + 1,59) = CODESTR07	'
			tmpRec(CurrentPage,index + 1,60) = CODESTR08	'B3 
			tmpRec(CurrentPage,index + 1,61) = CODESTR09	'-曠職
			tmpRec(CurrentPage,index + 1,54) = CODESTR10	'扣時假
			tmpRec(CurrentPage,index + 1,47) = CODESTR11	'扣時假
			tmpRec(CurrentPage,index + 1,48) = CODESTR12	'扣時假
			tmpRec(CurrentPage,index + 1,49) = CODESTR13	'扣時假
			tmpRec(CurrentPage,index + 1,50) = CODESTR14	'扣時假
			
  		 
			all_cb = tmpRec(CurrentPage,index + 1,33)			
			QC = tmpRec(CurrentPage,index + 1,27) 
			TBTR = tmpRec(CurrentPage,index + 1,41) 
			all_JBM = tmpRec(CurrentPage,index + 1,51) 
  		'+(加項)
 
  		
  		JX = tmpRec(CurrentPage,index + 1,29)  '績效獎金 
  		tnkh =  tmpRec(CurrentPage,index + 1,28)  
 
  		'-(減項)
  		GT = tmpRec(CurrentPage,index + 1,30)
			BH = tmpRec(CurrentPage,index + 1,32)
  		QITA = tmpRec(CurrentPage,index + 1,34) 
			HS =  0 
			allKM =tmpRec(CurrentPage,index + 1,54) 
		Notax_amt=tmpRec(CurrentPage,index + 1,53)	
  		'response.write  "allKM="& f_allKM &"<BR>"
			 
			'relTOTM = cdbl(all_cb)+cdbl(QC)+cdbl(tbtr)+cdbl(all_JBM)+cdbl(JX )+cdbl(tnkh)-cdbl(allKM)-cdbl(GT)-cdbl(bh)-cdbl(qita)-cdbl(hs)
			'relTOTM_dm = cdbl(all_cb)+cdbl(QC)+cdbl(tbtr)+cdbl(all_JBM)+cdbl(JX )+cdbl(tnkh)-cdbl(allKM)-cdbl(GT)-cdbl(bh)-cdbl(qita)-cdbl(hs)
			
			relTOTM = cdbl(all_cb)+cdbl(QC)+cdbl(tbtr)+cdbl(all_JBM)+cdbl(JX )+cdbl(tnkh)-cdbl(GT)-cdbl(bh)-cdbl(qita)-cdbl(hs)-cdbl(Notax_amt)
			relTOTM_dm = cdbl(all_cb)+cdbl(QC)+cdbl(tbtr)+cdbl(all_JBM)+cdbl(JX )+cdbl(tnkh)-cdbl(GT)-cdbl(bh)-cdbl(qita)-cdbl(hs)
			
			response.write  "ALL+ = " & all_cb &"+"&QC &"+"&TBTR &"+"&all_JBM &"+" &JX &"+" &tnkh &"="&  cdbl(all_cb)+cdbl(QC)+cdbl(tbtr)+cdbl(all_JBM)+cdbl(JX )+cdbl(tnkh)  &"<BR>"
			response.write  "ALL- = " & gt &"+"& bh &"+"& qita &"+"& hs &"+" & allKM  &"="&  cdbl(allKM)+cdbl(GT)+cdbl(bh)+cdbl(qita)+cdbl(hs) &"<BR>"
			
			if tmpRec(CurrentPage,index + 1,4) <>"VN" then 
				relTOTM = relTOTM * cdbl(rate)
			end if 
  		'個人所得稅計算
	  	F_TAX = 0 
			real_TOTAMT =  relTOTM   ' 實領金額
			'基本額 + 免稅額
			'B1 = tmpRec(CurrentPage,index + 1,78)
			if left(yymm,4)>"2008" then 
				B1="4000000"  
				B1="9000000"  '201306  , 900晚以上扣稅
			else
				B1="5000000"
			end if 
			B2 = tmpRec(CurrentPage,index + 1,53)  ' 免稅額
			
			Tot_B = cdbl(B1)+cdbl(B2)			 
			if real_TOTAMT > 0 then					
				
			sql2="exec sp_calctax_2020 '"& real_TOTAMT &"' ,'"&tmpRec(CurrentPage,index + 1,1)&"' "
			'response.write sql2
			set ors=conn.execute(sql2) 
			F_tax = ors("tax")
			taxper = ors("taxper")
				
				
				set ors=nothing 								
	  	else
				f_tax =0 
				taxper = 0 
			end if 
			
			if tmpRec(CurrentPage,index + 1,4) <>"VN" then  
				f_tax = round( cdbl(f_tax) / cdbl(rate) ,0)
			end if 			
			
			response.write "f_tax="   & f_tax &"<br>"
			response.write "relTOTM="& relTOTM &"<BR>"
			response.write "x final salary="& cdbl(relTOTM_dm) - cdbl(f_tax) &"<BR>"
			tmpRec(CurrentPage,index + 1,42)  = relTOTM_dm - CDBL(ROUND(F_tax,0))  '實領金額(含加班扣減時假-所得稅)    
			  
  		tmpRec(CurrentPage,index + 1,56)  = cdbl(all_cb)+cdbl(QC)+cdbl(tbtr)+cdbl(all_JBM)+cdbl(JX )+cdbl(tnkh)
  		tmpRec(CurrentPage,index + 1,40) = ROUND(F_TAX,0)
			if f_tax = "0" then 
				tmpRec(CurrentPage,index + 1,65) ="0%"  '稅率   
			else
				tmpRec(CurrentPage,index + 1,65) =taxper&"%"  '稅率   
			end if 
  		
%>		<script language=vbs>																	
			Parent.Fore.<%=self%>.TOTM(<%=index%>).value="<%=formatnumber(tmpRec(CurrentPage,index + 1,56),0)%>"
			Parent.Fore.<%=self%>.RELTOTMONEY(<%=index%>).value="<%=formatnumber(tmpRec(CurrentPage,index + 1,42) ,0)%>"			
			Parent.Fore.<%=self%>.KTAXM(<%=index%>).value="<%=formatnumber(tmpRec(CurrentPage,index + 1,40),0)%>"		
			Parent.Fore.<%=self%>.taxper(<%=index%>).value="<%=tmpRec(CurrentPage,index + 1,65)%>"		
		</script>
<%
end  select   		
response.write "54=" & tmpRec(CurrentPage,index + 1,54) &"<BR>"
Session("yece12B") = tmpRec
%>
</html>

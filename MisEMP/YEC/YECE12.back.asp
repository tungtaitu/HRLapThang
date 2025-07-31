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

CODESTR01 = REQUEST("CODESTR01")  : if CODESTR01="" then CODESTR01=0
CODESTR02 = REQUEST("CODESTR02")  : if CODESTR02="" then CODESTR02=0
CODESTR03 = REQUEST("CODESTR03")  : if CODESTR03="" then CODESTR03=0
CODESTR04 = REQUEST("CODESTR04")  : if CODESTR04="" then CODESTR04=0
CODESTR05 = REQUEST("CODESTR05")  : if CODESTR05="" then CODESTR05=0
CODESTR06 = REQUEST("CODESTR06")  : if CODESTR06="" then CODESTR06=0
CODESTR07 = REQUEST("CODESTR07")  : if CODESTR07="" then CODESTR07=0
CODESTR08 = REQUEST("CODESTR08")  : if CODESTR08="" then CODESTR08=0
CODESTR09 = REQUEST("CODESTR09")  : if CODESTR09="" then CODESTR09=0
CODESTR10 = REQUEST("CODESTR10")  : if CODESTR10="" then CODESTR10=0
CODESTR11 = REQUEST("CODESTR11")  : if CODESTR11="" then CODESTR11=0
CODESTR12 = REQUEST("CODESTR12")  : if CODESTR12="" then CODESTR12=0
CODESTR13 = REQUEST("CODESTR13")  : if CODESTR13="" then CODESTR13=0
CODESTR14 = REQUEST("CODESTR14")  : if CODESTR14="" then CODESTR14=0
CODESTR15 = REQUEST("CODESTR15")  : if CODESTR15="" then CODESTR15=0
CODESTR16 = REQUEST("CODESTR16")  : if CODESTR16="" then CODESTR16=0
CODESTR17 = REQUEST("CODESTR17")  : if CODESTR17="" then CODESTR17=0




workdays = REQUEST("days")  
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
			tmpRec(CurrentPage,index + 1,51) = CODESTR04	' 總加班費
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
			if tmpRec(CurrentPage,index + 1,4) <>"VN" and tmpRec(CurrentPage,index + 1,4) <>"CT" then  
				tmpRec(CurrentPage,index + 1,81) = CODESTR15	'hsf
				tmpRec(CurrentPage,index + 1,83) = CODESTR16	'buamt=tien3 			
				tmpRec(CurrentPage,index + 1,26) = CODESTR17	'ttkh
			end if 
			
			BB=CDBL(tmpRec(CurrentPage, index + 1, 20))
			CV=CDBL(tmpRec(CurrentPage, index + 1, 21))
			PHU=CDBL(tmpRec(CurrentPage, index + 1, 22))				
			NN=CDBL(tmpRec(CurrentPage, index + 1, 23))
			KT=CDBL(tmpRec(CurrentPage, index + 1, 24))
			MT=CDBL(tmpRec(CurrentPage, index + 1, 25)) '匯率津貼	
			TTKH = cdbl(tmpRec(CurrentPage,index + 1,26))  '其家		
			btien = cdbl(tmpRec(CurrentPage,index + 1,76))		
			QC=CDBL(tmpRec(CurrentPage, index + 1, 27))				
			
			if tmpRec(CurrentPage,index + 1,4) <>"VN" and tmpRec(CurrentPage,index + 1,4) <>"CT" then  
				tmpRec(CurrentPage,index + 1,33) = bb+cv+phu+nn+kt+mt+ttkh+biten
			end if 
			empid=tmpRec(CurrentPage,index + 1,1)  
			QC = tmpRec(CurrentPage,index + 1,27) 
			TBTR = tmpRec(CurrentPage,index + 1,41) 
			
			all_JBM = tmpRec(CurrentPage,index + 1,51)  
			
			bb = tmpRec(CurrentPage,index + 1,20)  
			CV = tmpRec(CurrentPage,index + 1,21)
			PHU = tmpRec(CurrentPage,index + 1,22)
			
			TNKH = tmpRec(CurrentPage,index + 1,28)  '其他收入
			JX = tmpRec(CurrentPage,index + 1,29)
			hsf = tmpRec(CurrentPage,index + 1,81) 
			if tmpRec(CurrentPage,index + 1,4) ="VN" or tmpRec(CurrentPage,index + 1,4) ="CT" then  
				all_cb = tmpRec(CurrentPage,index + 1,33)  
			else
				all_cb = cdbl(tmpRec(CurrentPage,index + 1,33)) 
			end if			
			
			govbb =  tmpRec(CurrentPage,index + 1,77)  'for tw, ma 
			
			
  		'+(加項)
 
  		
  		JX = tmpRec(CurrentPage,index + 1,29)  '績效獎金 
  		tnkh =  tmpRec(CurrentPage,index + 1,28)  
 
  		'-(減項)
  		GT = tmpRec(CurrentPage,index + 1,30)
			BH = tmpRec(CurrentPage,index + 1,32)
			QITA = tmpRec(CurrentPage,index + 1,34) 
			HS =  0 
			allKM =tmpRec(CurrentPage,index + 1,54) 			
			v_buamt=CODESTR16   
			
			response.write  "all_cb="& all_cb &"<BR>"
			response.write  "allKM="& f_allKM &"<BR>"
			 
			relTOTM =    cdbl(all_cb)+cdbl(QC)+cdbl(tbtr)+cdbl(all_JBM)+cdbl(JX) +cdbl(tnkh)-cdbl(allKM)-cdbl(GT)-cdbl(bh)-cdbl(qita)-cdbl(hs) 
			relTOTM_dm = cdbl(all_cb)+cdbl(QC)+cdbl(tbtr)+cdbl(all_JBM)+cdbl(JX)+cdbl(tnkh)-cdbl(allKM)-cdbl(GT)-cdbl(bh)-cdbl(qita)-cdbl(hs) 
  		
			response.write  "ALL+ = " & all_cb &"+"&QC &"+"&TBTR &"+"&all_JBM &"+" &JX &"+" &tnkh &"="&  cdbl(all_cb)+cdbl(QC)+cdbl(tbtr)+cdbl(all_JBM)+cdbl(JX )+cdbl(tnkh)  &"<BR>"
			response.write  "ALL- = " & gt &"+"& bh &"+"& qita &"+"& hs &"+" & allKM  &"="&  cdbl(allKM)+cdbl(GT)+cdbl(bh)+cdbl(qita)+cdbl(hs) &"<BR>"
			if tmpRec(CurrentPage,index + 1,4) <>"VN" and tmpRec(CurrentPage,index + 1,4) <>"CT" then  
				relTOTM = relTOTM+cdbl(v_buamt)
			end if 
			if tmpRec(CurrentPage,index + 1,4) <>"VN" and tmpRec(CurrentPage,index + 1,4) <>"CT" then 
				relTOTM = relTOTM * cdbl(rate)
			end if 
			'個人所得稅計算
			F_TAX = 0 
			real_TOTAMT =  relTOTM   ' 實領金額
			'基本額 + 免稅額
			'B1 = tmpRec(CurrentPage,index + 1,78)
			if left(yymm,4)>"2008" then 
				B1="4000000"  
				B1="9000000"  '201306  , 900萬以上扣稅
				B1="11000000"    '202006  , 1100萬以上扣稅
			else
				B1="5000000"
			end if 
			B2 = tmpRec(CurrentPage,index + 1,53)  ' 免稅額
			
			Tot_B = cdbl(B1)+cdbl(B2)
			
			if tmpRec(CurrentPage, index+1, 4)="TW"  and tmpRec(CurrentPage,index + 1,85)="B"  then 
				BBAMT=CDBL(BB)+CDBL(cv)+cdbl(ttkh)
				wpamt=cdbl(jx)+cdbl(TNKH)-cdbl(qita)
				empallamt =cdbl(all_cb)+cdbl(QC)+cdbl(tbtr)+cdbl(all_JBM)+cdbl(JX)+cdbl(tnkh)-cdbl(allKM)-cdbl(GT)-cdbl(bh)-cdbl(qita)-cdbl(hs)
				if cdbl(jx)+cdbl(TNKH)>0 then 
					buamt=cdbl(v_buamt)+cdbl(hsf)
				else
					buamt=cdbl(v_buamt)+cdbl(hsf)
				end if 
			elseif  tmpRec(CurrentPage, index+1, 4)="TW"  or tmpRec(CurrentPage, index+1, 4)="MA"   then 
				BBAMT=CDBL(BB)+cdbl(ttkh)
				wpamt=CDBL(cv)+cdbl(jx)+cdbl(TNKH)-cdbl(qita)
				empallamt =cdbl(all_cb)+cdbl(QC)+cdbl(tbtr)+cdbl(all_JBM)+cdbl(JX )+cdbl(tnkh)-cdbl(allKM)-cdbl(GT)-cdbl(bh)-cdbl(qita)-cdbl(hs)
				if cdbl(jx)+cdbl(TNKH)>0 then 
					buamt=cdbl(v_buamt)+cdbl(hsf)
				else
					buamt=cdbl(v_buamt)+cdbl(hsf) 'cdbl(hsf)
				end if 
			elseif  tmpRec(CurrentPage, index+1, 4)="CN"  then 
				'if cdbl(real_TOTAMT) > cdbl(Tot_B) then   						
					BBAMT= cdbl(bb)+cdbl(cv)+cdbl(phu)+cdbl(ttkh)-cdbl(allKM)-cdbl(GT)-cdbl(bh)-cdbl(qita) 
					response.write  "BBAMT=" & BBAMT &"<br>"
					wpamt=cdbl(jx)+cdbl(TNKH)
					govbb = BBAMT 
					buamt =  v_buamt+cdbl(hsf) 'if tmpRec(CurrentPage, index+1, 4)="TA" then  buamt=500  else  buamt = 0 
					empallamt = cdbl(all_cb)+cdbl(QC)+cdbl(tbtr)+cdbl(all_JBM)+cdbl(JX )+cdbl(tnkh)-cdbl(allKM)-cdbl(GT)-cdbl(bh)-cdbl(qita)-cdbl(hs) 							
				'end if 
			elseif   tmpRec(CurrentPage, index+1, 4)="TA"  then 
				BBAMT=cdbl(all_cb)+cdbl(QC)+cdbl(tbtr)+cdbl(all_JBM)+cdbl(JX)-cdbl(allKM)-cdbl(GT)-cdbl(bh)-cdbl(qita)   
				response.write  "BBAMT=" & BBAMT &"<br>"
				wpamt=cdbl(TNKH)
				govbb = BBAMT
				buamt =  v_buamt+cdbl(hsf)  'if tmpRec(CurrentPage, index+1, 4)="TA" then  buamt=500  else  buamt = 0 
				empallamt = cdbl(all_cb)+cdbl(QC)+cdbl(tbtr)+cdbl(all_JBM)+cdbl(JX )+cdbl(tnkh)-cdbl(allKM)-cdbl(GT)-cdbl(bh)-cdbl(qita)-cdbl(hs)
			end if 
			response.write "empallamt >>>>> " & empallamt &"<br>"
			if real_TOTAMT > cdbl(Tot_B) then  
				if left(yymm,4)<="2008" then 
					sql2="exec sp_calctax_HW_2008 '"& real_TOTAMT &"' "
					set ors=conn.execute(sql2) 
					f_tax = ors("tax")
					taxper = 0	
				elseif yymm>="201108" and yymm<"201112" then
					sql2="exec sp_calctax_201108 '"& real_TOTAMT &"' , '"& cdbl(tot_b) &"' ,'"&tmpRec(CurrentPage,index + 1,1)&"'  "
					response.write sql2&"<BR>"
					set ors=conn.execute(sql2) 
					f_tax = ors("tax")
					taxper = ors("taxper")
				else
					sql2="exec sp_calctax_2010 '"& real_TOTAMT &"' , '"& cdbl(tot_b) &"' ,'"&tmpRec(CurrentPage,index + 1,1)&"' "
					response.write sql2&"<BR>"
					set ors=conn.execute(sql2) 
					f_tax = ors("tax")
					taxper = ors("taxper")
				end if   
				set ors=nothing  

' [dbo].[aSp_2020calc_buTaxamt]    
' @yymm as varchar(6)
',@eid as varchar(20)  
',@empbbamt as varchar(20)  --自行負擔
',@empwpamt as varchar(20)  --(薪資-自行負擔)+獎金+其他-扣除 
',@bu_amt as varchar(20)  --沖帳金額
',@empallamt as varchar(20) --全部薪資
',@gov_BB as varchar(20)  --報稅薪資
',@typs as varchar(5)  -- 傳1表示系統取值 , 空白表示輸入值  
',@Tot_B as varchar(20) --免稅額 
',@emp_jx as varchar(20)
',@emp_tnkh  as varchar(20)
',@emp_hsf as varchar(20)
',@emp_qita as varchar(20)
',@emp_bzkm as varchar(20)				
				
				if tmpRec(CurrentPage, index+1, 4)<>"VN" and tmpRec(CurrentPage, index+1, 4)<>"CT"  then 
					'扣稅 TW+MA  = BB , ( 派任 + CN + TA ) BB+CV+PHU+..  全部扣稅 20200915	 
					if empid="A0079" then   govbb = cdbl(govbb) + cdbl(jx)+ cdbl(tnkh)
					sqlb="exec [aSp_2020calc_buTaxamt]  '"&yymm&"' ,'"&empid&"' ,'"&BBAMT&"','"&wpamt&"' ,'"&buamt&"','"&empallamt&"','"&govbb&"','' ,'"&tot_b&"' ,'"&jx&"','"&tnkh&"','"&hsf&"','"&qita&"','"&allKM&"'  " 
					'response.write sqlb&"<BR>"
					'response.end 
					set nrs=conn.execute(sqlb) 					
					
					emp_dftax = nrs("df_tax") '自行負擔稅額
					emp_aftertax = nrs("empaftertax") '員工稅後金額
					emp_newamt = nrs("new_amt")    '員工報稅金額 					
					emp_newtax  = nrs("new_tax") '員工稅額
					emp_buamt = nrs("new_buamt") '員工補稅金額
					emp_realAmt = nrs("relamt") '員工實領(入賬)
					
					gov_dftax = nrs("gov_dftax") 'gov負擔稅額
					gov_aftertax = nrs("govaftertax") 'gov稅後金額
					gov_newamt = nrs("govnew_amt") 	'GOV報稅金額 
					gov_newtax  = nrs("govtax")   	'GOV稅額	 
					govtaxper	= nrs("govtaxper")   	'GOV稅率	 				
					gov_buamt = cdbl(gov_newtax)-cdbl(gov_dftax) 	'GOV補稅金額
					'gov_realAmt = nrs("relamt") 	'GOV實領(入賬)
					f_tax = gov_newtax
					response.write  empid &" ==> govBB > " & govbb&"<BR>"
					response.write  " ==> BBAMT >" & BBAMT &"<BR>"
					response.write  " ==> wpamt >" & wpamt &"<BR>"
					response.write  " ==> empallamt >" & empallamt&"<BR>"
					response.write  " ==> emp_newamt >" & emp_newamt &"<BR>"
					response.write  " ==> emp_dftax >" & emp_dftax &"<BR>"
					response.write  " ==> emp_buamt >" & emp_buamt&"<BR>"
					response.write  " ==> emp_aftertax >" & emp_aftertax&"<BR>"
					response.write  " ==> emp_realAmt >" & emp_realAmt&"<BR>"					
					response.write  " ==> gov_bb >" & govbb&"<BR>"
					response.write  " ==> gov_dftax >" & gov_dftax&"<BR>"
					response.write  " ==> gov_newamt >" & gov_newamt&"<BR>"
					response.write  " ==> gov_newtax >" & gov_newtax &"<BR>"
					response.write  " ==> gov_aftertax >" & gov_aftertax&"<BR>"
					
					Session("yece12B") = tmpRec
				end if 
			else
				f_tax =0 
				taxper = 0 
				empallamt = empallamt  
				emp_aftertax = empallamt 'cdbl(all_cb)+cdbl(QC)+cdbl(tbtr)+cdbl(all_JBM)+cdbl(JX)-cdbl(allKM)-cdbl(GT)-cdbl(bh)-cdbl(qita) 
				gov_aftertax =emp_aftertax 
				gov_newamt = empallamt
				govbb =empallamt 
				govtaxper = 0 
				gov_newtax = 0 
				emp_dftax = 0 
				
			end if 
			
			if tmpRec(CurrentPage,index + 1,4) <>"VN" and tmpRec(CurrentPage,index + 1,4) <>"CT" then  
				f_tax = round( cdbl(f_tax) / cdbl(rate) ,0)
			end if 			
			butax = cdbl(gov_newamt)-cdbl(govbb)-cdbl(v_buamt)-cdbl(hsf)  'cdbl(gov_newamt)-cdbl(empallamt)
			before_tax =cdbl(gov_newamt)-cdbl(empallamt)
			after_tax = cdbl(gov_aftertax)-cdbl(emp_aftertax)
			cty_save = cdbl(after_tax)-cdbl(hsf)
			response.write "差額=" & butax &"== >("& gov_newamt &"-"& empallamt &")<br>"
			response.write "稅後差額=" & after_tax &"== >("& gov_aftertax &"-"& emp_aftertax &")<br>"
			response.write "留存=" & cdbl(after_tax)-cdbl(hsf) &"== >("& after_tax &"-"& hsf &")<br>"
			response.write "補稅=" & cdbl(gov_newamt)-cdbl(govbb)-cdbl(v_buamt)-cdbl(hsf) &"== >("& after_tax &"-"& hsf &")<br>"
			response.write "f_tax="   & f_tax &"<br>"
			response.write "relTOTM="& relTOTM &"<BR>"
			response.write "x final salary="& cdbl(relTOTM_dm) - cdbl(f_tax) &"<BR>"
			
			
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
			
  		'response.write  "allKM="& f_allKM &"<BR>"
			 
			relTOTM = cdbl(all_cb)+cdbl(QC)+cdbl(tbtr)+cdbl(all_JBM)+cdbl(JX )+cdbl(tnkh)-cdbl(allKM)-cdbl(GT)-cdbl(bh)-cdbl(qita)-cdbl(hs) 
			relTOTM_dm = cdbl(all_cb)+cdbl(QC)+cdbl(tbtr)+cdbl(all_JBM)+cdbl(JX )+cdbl(tnkh)-cdbl(allKM)-cdbl(GT)-cdbl(bh)-cdbl(qita)-cdbl(hs) 
			
			
  		'tmpRec(CurrentPage,index + 1,56)  = cdbl(all_cb)+cdbl(QC)+cdbl(tbtr)+cdbl(all_JBM)+cdbl(JX )+cdbl(tnkh)
		if tmpRec(CurrentPage, index+1, 4)="VN" or tmpRec(CurrentPage, index+1, 4)="CT"   then  
			tmpRec(CurrentPage,index + 1,56)  = cdbl(all_cb)+cdbl(QC)+cdbl(tbtr)+cdbl(all_JBM)+cdbl(JX)+cdbl(tnkh)
		else 
			tmpRec(CurrentPage,index + 1,56)  = (cdbl(all_cb)+cdbl(QC)+cdbl(tbtr)+cdbl(all_JBM)+cdbl(JX )+cdbl(tnkh))-(cdbl(allKM)+cdbl(GT)+cdbl(bh)+cdbl(qita)+cdbl(hs))  
		end if 
		if tmpRec(CurrentPage, index+1, 4)<>"VN" and tmpRec(CurrentPage, index+1, 4)<>"CT"   then  
			tmpRec(CurrentPage,index + 1,40) = ROUND(gov_newtax,0)
		ELSE
			tmpRec(CurrentPage,index + 1,40) = ROUND(F_TAX,0)
		end if 	 
		
			if f_tax = "0" then 
				tmpRec(CurrentPage,index + 1,65) ="0%"  '稅率   
			else
				tmpRec(CurrentPage,index + 1,65) =taxper&"%"  '稅率   
			end if 
		if tmpRec(CurrentPage, index+1, 4)="VN" or tmpRec(CurrentPage, index+1, 4)="CT" then 
			tmpRec(CurrentPage,index + 1,42) = relTOTM_dm - CDBL(ROUND(F_tax,0))  '實領金額(含加班扣減時假-所得稅)    	
			'tmpRec(CurrentPage,index + 1,65) = govtaxper&"%" 
		else 
			tmpRec(CurrentPage,index + 1,42) = gov_aftertax  'relTOTM_dm - CDBL(ROUND(F_tax,0))  '實領金額(含加班扣減時假-所得稅)    	
			tmpRec(CurrentPage,index + 1,65) = govtaxper&"%"
		end if  
		tmpRec(CurrentPage,index + 1,84) = emp_dftax    '原扣稅金額
		tmpRec(CurrentPage,index + 1,82) = emp_aftertax    'EMP 實領 zhuanM 
		tmpRec(CurrentPage,index + 1,86) = gov_newamt    '報稅 (govbb)
		tmpRec(CurrentPage,index + 1,87) = cdbl(gov_newamt)-cdbl(govbb)    '報稅其他
		'tmpRec(CurrentPage,index + 1,89) = butax    '報稅其他
		tmpRec(CurrentPage,index + 1,80) = butax
		tmpRec(CurrentPage,index + 1,89) = (before_tax) '稅前
		tmpRec(CurrentPage,index + 1,90) = (after_tax) '稅後
		if cdbl(tmpRec(CurrentPage,index + 1,42))<0 then tmpRec(CurrentPage,index + 1,42) = 0 
		govall= gov_newamt 
  		
%>		<script language=vbs>																	
			Parent.Fore.<%=self%>.TOTM(<%=index%>).value="<%=formatnumber(tmpRec(CurrentPage,index + 1,56),0)%>"
			Parent.Fore.<%=self%>.RELTOTMONEY(<%=index%>).value="<%=formatnumber(tmpRec(CurrentPage,index + 1,42) ,0)%>"			
			Parent.Fore.<%=self%>.KTAXM(<%=index%>).value="<%=formatnumber(tmpRec(CurrentPage,index + 1,40),0)%>"		
			Parent.Fore.<%=self%>.taxper(<%=index%>).value="<%=tmpRec(CurrentPage,index + 1,65)%>"		
			Parent.Fore.<%=self%>.empZAmt(<%=index%>).value="<%=formatnumber(tmpRec(CurrentPage,index + 1,82) ,0)%>"			
			Parent.Fore.<%=self%>.govALL(<%=index%>).value="<%=formatnumber(govall ,0)%>"			
			Parent.Fore.<%=self%>.butax(<%=index%>).value="<%=formatnumber(butax ,0)%>"			
			Parent.Fore.<%=self%>.after_tax(<%=index%>).value="<%=formatnumber(after_tax ,0)%>"			
			Parent.Fore.<%=self%>.cty_saveamt(<%=index%>).value="<%=formatnumber(cty_save ,0)%>"						
			
		</script>
<%
end  select   		
response.write "54=" & tmpRec(CurrentPage,index + 1,54) &"<BR>"
Session("yece12B") = tmpRec
%>
</html>

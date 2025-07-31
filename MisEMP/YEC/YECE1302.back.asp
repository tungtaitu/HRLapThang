<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<!-- #include file = "../GetSQLServerConnection.fun" --> 
<!-- #include file="../ADOINC.inc" --> 
<%
SELF = "YECE1302" 

func = request("func") 
code = request("code") 
index=request("index")  
CurrentPage = request("CurrentPage") 

CODE01 = REQUEST("CODE01")
CODE02 = REQUEST("CODE02")
CODE03 = REQUEST("CODE03")
CODE04 = REQUEST("CODE04")
CODE05 = REQUEST("CODE05")
CODE06 = REQUEST("CODE06")
CODE07 = REQUEST("CODE07")
CODE08 = REQUEST("CODE08")
CODE09 = REQUEST("CODE09") 
CODE10 = REQUEST("CODE10") 
CODE11 = REQUEST("CODE11") 
CODE12 = REQUEST("CODE12") 
CODE13 = REQUEST("CODE13") 
CODE14 = REQUEST("CODE14") 
workdays = REQUEST("days")  
response.write  "workdays=" & workdays &"<BR>"  
yymm=request("yymm") 
rate = request("rate")  

calcdt = left(YYMM,4)&"/"& right(yymm,2)&"/01" 

'tmpRec = Session("YECE1302B") 
response.write "index=" & index &"<BR>"
response.write "func=" & func &"<BR>"
response.write "CurrentPage=" & CurrentPage &"<BR>"

years = request("years")
country = request("country")
whsno = request("whsno")
nz=request("nz")
rhs=request("hs")


Set conn = GetSQLServerConnection()	   

 tmpRec  = Session("YECE1302B") 
'response.end 
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
</head>
<%
select case func    
	case "chkkj"		
		'tmpRec(CurrentPage,index + 1,0) = "UPD"
		sql="select * from EMPNZJJ_set where years='"& trim(years)&"' and country='"&trim(country)&"' and whsno='"& trim(whsno) &"' "&_
				"and ( kj='"& trim(CODE01) &"'  or grade='"& trim(code01) &"' )  "
		response.write sql  &"<BR>"

		set rs=conn.execute(sql)
		if rs.eof then    		
%>		<script language=vbs>																	
				alert "考績輸入錯誤!!(或尚未設定)"
				parent.Fore.<%=self%>.kj(<%=index%>).value = parent.Fore.<%=self%>.old_kj(<%=index%>).value
				parent.Fore.<%=self%>.kj(<%=index%>).select()
			</script>
<%		response.end 
			set rs=nothing 
		else 					
			khac = request("khac")  
			basicNZM = request("basicBZM")  
			if khac="" then khac=0 	
			response.write "調整=" & khac &"<BR>"
			'basicNZM = tmpRec(1, index+1, 23)
			response.write "basicNZM="& basicNZM &"<BR>"  
			nz = request("nz")
			hs = request("hs")
			if nz="" then nz="0"
			response.write "年資="& nz &"<BR>"  
			response.write "xisu 係數= "& hs  &"<BR>"  
			df_bonus = request("df_bonus") 
			
			if years<="2008"   then hs=round(rs("hs"),2) end if 
			
			if years<="2008" then 
				if cdbl(nz)>=12 then 
					bonus = cdbl(basicNZM) * cdbl(hs)
				else
					bonus = cdbl(basicNZM)*cdbl(hs)*round(nz,2)/12 
					response.write "BX1"&"<BR>"
				end if 	 
				response.write "bonus=" & bonus &"<BR>"
			else  			
				if cdbl(nz)>=12 then 
					bonus = cdbl(basicNZM) *cdbl(hs)
					response.write "BX1a =  "& cdbl(basicNZM) &","& cdbl(hs) & " <BR>"
				else
					bonus = cdbl(basicNZM)*cdbl(hs)*round(nz,1)/12 
					response.write "BX2"&"<BR>"
				end if 	 
				response.write "bonus=" & bonus &"<BR>"
			end if 	
			
			bonus = cdbl(df_bonus) 
			f_bonus = cdbl(bonus) + cdbl(khac)
			
			response.write "調整後=" & f_bonus &"<BR>"
			
			if country="VN" then  
				if years<="2008" then 
						sql2="exec sp_calctax_2008 '"& f_bonus  &"' "
						set ors=conn.execute(sql2) 
						F_tax = ors("tax")					
				elseif years>="2013" then 
						sql2="exec  sp_calctax_2010  '"& f_bonus  &"' , 0,'' "
						set ors=conn.execute(sql2) 
						F_tax = ors("tax")
						taxper = ors("taxper")
				else
						sql2="exec  sp_calctax  '"& f_bonus  &"' , '4000000' "
						set ors=conn.execute(sql2) 
						F_tax = ors("tax")
						taxper = ors("taxper")
				end if 
				set ors=nothing 
			else   '外國人不扣稅 
				f_tax = 0 
			end if 	
			
			response.write "f_yax=" & F_tax &"<BR>"			
			f2_bonus = cdbl(f_bonus) - cdbl(f_tax) 	 			
			
			if country="VN" then  			
				'if cdbl(f2_bonus) mod 500 <> 0  then 				
				'	rel_bonus = fix(f2_bonus/500+1) * 500 
				'else	
				'	rel_bonus = f2_bonus
				'end if 	 
				'rel_bonus = f2_bonus
				rel_bonus = ceil( (f2_bonus/1.0)) *1
			else
				rel_bonus = f2_bonus
			end if   
			
			response.write  "f2_bonus(+-調整-稅)="& f2_bonus & "  ,f2_bonus_int="& fix(f2_bonus/1) &"<BR>"			
			response.write "ceil="& ceil( (f2_bonus/1.0))&"<BR>"	
			response.write "ceil_result="& ceil( (f2_bonus/1.0)) *1 &"<BR>"	
%>
			<script language="vbscript">		
				'alert ( parent.Fore.document.getElementById("intrst1").innerText )
				i_basicbzm = replace(parent.Fore.document.getElementsByName("basicBZM")(<%=index%>).value ,",","")
				i_hs = replace(parent.Fore.document.getElementsByName("hs")(<%=index%>).value ,",","")
				i_bonus = replace(parent.Fore.document.getElementsByName("bonus")(<%=index%>).value ,",","")
				cols= "intrst"&"<%=(index+1)%>"
				parent.Fore.document.getElementById(cols).innerText = "<%=formatnumber(f2_bonus,0)%>"
				'parent.Fore.<%=self%>.bodays(<%=index%>).value = "<%=rs("days")%>"
				'parent.Fore.<%=self%>.hs(<%=index%>).value = "<%=hs%>"
				'parent.Fore.<%=self%>.bonus(<%=index%>).value = "<%=formatnumber(bonus,0)%>"
				'parent.Fore.<%=self%>.khac(<%=index%>).value = "<%=formatnumber(khac,0)%>"
				parent.Fore.<%=self%>.tax(<%=index%>).value = "<%=formatnumber(F_tax,0)%>"
				parent.Fore.<%=self%>.r_bonus(<%=index%>).value = "<%=formatnumber(rel_bonus,0)%>"
				'parent.Fore.<%=self%>.grade(<%=index%>).value = "<%=rs("grade")%>"
			</script>		
<%	end if 			
		rs.close : set rs=nothing  
		response.end 
end  select    		


function floor(x)
    dim temp

    temp = Round(x)

    if temp > x then
        temp = temp - 1
    end if

    floor = temp
end function

' Returns the smallest integer greater than or equal to the specified number.
function ceil(x)
    dim temp

    temp = Round(x)

    if temp < x then
        temp = temp + 1
    end if

    ceil = temp
end function 
%>
</html>

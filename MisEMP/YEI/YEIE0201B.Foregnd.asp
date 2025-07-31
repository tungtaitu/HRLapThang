<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<!-- #include file = "../GetSQLServerConnection.fun" --> 
<%  
SELF = "YEIE0201B"

Set conn = GetSQLServerConnection()
Set rs = Server.CreateObject("ADODB.Recordset")
Set rds = Server.CreateObject("ADODB.Recordset")

nowmonth = year(date())&right("00"&month(date()),2)
if month(date())="01" then
	calcmonth = year(date()-1)&"12"
else
	calcmonth =  year(date())&right("00"&month(date())-1,2)
end if 
 
'一個月有幾天 
'cDatestr=CDate(LEFT(khym,4)&"/"&RIGHT(khym,2)&"/01") 
'days = DAY(cDatestr+(32-DAY(cDatestr))-DAY(cDatestr+(32-DAY(cDatestr))))   '一個月有幾天  
'本月最後一天 
'ENDdat = LEFT(khym,4)&"/"&RIGHT(khym,2)&"/"&DAYS   


khyears=request("khyears")
khud=Ucase(Trim(request("khud")))
F_whsno = request("F_whsno")
F_groupid = request("F_groupid")
F_zuno = request("F_zuno") 
F_shift=request("F_shift")
F_empid =request("f_empid")
F_country=request("F_country") 
mylevel = request("mylevel") 
mylevel="Z"
mycqid =  request("mycqid") 
if mycqid="" then mycqid=session("netuser")

f_khBid=khyears&khud  
 
if  khud="U" then 
	d1_str = khyears&"0101"
	d2_str= khyears&"0630" 
elseif khud="D" then  
	d1_str = khyears&"0701"
	d2_str= khyears&"1231" 
end if 	 

'if khweek="" then khweek=(days\7)    


gTotalPage = 1
PageRec = 0    'number of records per page
TableRec = 30    'number of fields per record    
 
'response.end 		


' @empid as varchar(10), @khbid as varchar(10) , @empkhid as varchar(15) , @cqid as varchar(10)  
khyear = request("khyear")
empid = request("empid")
khbid=request("khbid") 
empkhid=request("empkhid")  
index = request("index")

tmprec = session("yeie0201B")

sql="select * from fn_yeie0201_empkh ( '"&empid&"','"&khbid&"','"&empkhid&"','"&mycqid&"') order by sttno " 
rs.Open SQL, conn, 1, 3    
'response.write sql  
IF NOT RS.EOF THEN
		PageRec = rs.RecordCount
		rs.PageSize = PageRec
		RecordInDB = rs.RecordCount
		TotalPage = rs.PageCount
		gTotalPage = TotalPage
	END IF
	Redim tmpRecb(gTotalPage, PageRec, TableRec)   'Array

	for i = 1 to TotalPage
		for j = 1 to PageRec
			if not rs.EOF then
				tmpRecB(i, j, 0) = rs("khbid")
				tmpRecB(i, j, 1) = rs("sttno")
				tmpRecB(i, j, 2) = rs("grade")
				tmpRecB(i, j, 3) = rs("fensu")
				tmpRecB(i, j, 4) = rs("khstr_cn")
				tmpRecB(i, j, 5) = rs("khstr_vn")
				tmpRecB(i, j, 6) = rs("aid")
				tmpRecB(i, j, 7) = rs("cqgrade") 
				tmpRecB(i, j, 8) = rs("zsts") 
				tmpRecB(i, j, 9) = rs("jsts") 
				tmpRecB(i, j, 10) = rs("hsts") 
				tmpRecB(i, j, 11) = rs("cqid") 
				tmpRecB(i, j, 12) = rs("cq_level") 
				tmpRecB(i, j, 13) = rs("cqfensu") 
				tmpRecB(i, j, 14) = rs("zcqmemos") 
				tmpRecB(i, j, 15) = rs("jcqmemos") 
				tmpRecB(i, j, 16) = rs("hcqmemos") 
				tmpRecB(i, j, 17) = rs("z_fs") 
				tmpRecB(i, j, 18) = rs("j_fs") 
				tmpRecB(i, j, 19) = rs("h_fs") 
				tmpRecB(i, j, 20) = rs("z_kj") 
				tmpRecB(i, j, 21) = rs("j_kj") 
				tmpRecB(i, j, 22) = rs("h_kj") 
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
 	
  
empname = tmprec(1,index+1,2) &" "&tmprec(1,index+1,3) 
emp_indate = tmprec(1,index+1,4)
emp_group = tmprec(1,index+1,6)&" "&tmprec(1,index+1,7)
emp_job = tmprec(1,index+1,24)
kzhour =tmprec(1,index+1,25)
flz =tmprec(1,index+1,26)
jiaA =tmprec(1,index+1,27)
jiaB =tmprec(1,index+1,28)
'response.write tmprec(1,index+1,2) 
'response.end 

FUNCTION FDT(D)
	Response.Write YEAR(D)&"/"&RIGHT("00"&MONTH(D),2)&"/"&RIGHT("00"&DAY(D),2)
END FUNCTION 	
%>

<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta HTTP-EQUIV="refresh" >
<link rel="stylesheet" href="../Include/style.css" type="text/css">
<link rel="stylesheet" href="../Include/style2.css" type="text/css">
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
 
'-----------------enter to next field
function enterto()
	if window.event.keyCode = 13 then window.event.keyCode =9
end function

function f()
 
end function
 

function datachg() 
	<%=self%>.totalpage.value="0"
	<%=self%>.sortvalue.value=""
	<%=self%>.action = "<%=self%>.ForeGnd.asp"
	<%=self%>.target="_self"
	<%=self%>.submit()
end function   

function datachg2() 
	<%=self%>.totalpage.value="0"	
	<%=self%>.act.value="A"
	<%=self%>.action = "<%=self%>.ForeGnd.asp"
	<%=self%>.target="_self"
	<%=self%>.submit()
end function 

function sortby(a)
	if a=1 then 
		<%=self%>.sortvalue.value="a.khz, a.empid"
	elseif a=2 then	
		<%=self%>.sortvalue.value="a.empid"
	elseif a=3 then	
		<%=self%>.sortvalue.value="b.nindat, a.empid"
	elseif a=4 then	
		<%=self%>.sortvalue.value="a.monthfen desc, a.empid"
	elseif a=5 then	
		<%=self%>.sortvalue.value="len(a.khs) desc, a.khs, a.khz, a.empid"
	else
		<%=self%>.sortvalue.value="b.country , a.khw, a.khg, a.empid"			
	end if 	 
	<%=self%>.totalpage.value="0"
	<%=self%>.action = "<%=self%>.ForeGnd.asp"
	<%=self%>.target="_self"
	<%=self%>.submit()
	'alert a 
end function 

 
</SCRIPT>
</head>
<body  topmargin="0" leftmargin="5"  marginwidth="0" marginheight="0"  onkeydown="enterto()" onload="f()">
<form  name="<%=self%>" method="post" action="<%=self%>.upd.asp"   >
<INPUT TYPE=HIDDEN NAME="UID" VALUE=<%=SESSION("NETUSER")%>>
<INPUT TYPE=hidden NAME=TotalPage VALUE="<%=TotalPage%>">
<INPUT TYPE=hidden NAME=CurrentPage VALUE="<%=CurrentPage%>">
<INPUT TYPE=hidden NAME=RecordInDB VALUE="<%=RecordInDB%>">
<INPUT TYPE=hidden NAME=PageRec VALUE="<%=PageRec%>">
<INPUT TYPE=hidden NAME=gTotalPage VALUE="<%=gTotalPage%>"> 	 

<INPUT TYPE=hidden NAME="khyear" VALUE="<%=khyear%>"> 	 
<INPUT TYPE=hidden NAME="empid" VALUE="<%=empid%>"> 	 
<INPUT TYPE=hidden NAME="khbid" VALUE="<%=khbid%>"> 	 
<INPUT TYPE=hidden NAME="empkhid" VALUE="<%=empkhid%>">
<INPUT TYPE=hidden NAME="pindex" VALUE="<%=index%>">
 

<table BORDER=0 cellspacing="1" cellpadding="1"  class=txt12 width=600>
	<TR  >		
		<Td align="center"> 永豐餘越南公司 ( <%=session("mywhsno")%> )</td>
	</tr>
	<TR   >		
	<Td  align="center"> <%=khyear%> 年度 <%=replace(replace(right(khyear,1),"U","上"),"D","下")%> 半年員工考核表 </td>
	</tr> 
	</TR>	 
</table>
<hr size=0	style='border: 1px dotted #999999;' align=left width=600> 
<TABLE WIDTH=550 CLASS=txt BORDER=0  cellspacing="1" cellpadding="1">
	<TR >
		<TD width=60 nowrap align=right >工號<br><font class=txt8>So The</font></TD>
		<TD ><%=EMPID%></TD>
		<TD width=60 nowrap align=right >姓名<br><font class=txt8>Ho Ten</font></TD>
		<TD ><%=empname%></TD>		
		<td     align=center valign=top  nowrap  rowspan="5" >
			<img src="../yeb/pic/<%=EMPID%>.jpg"  border=1 width=77 height=100  >
		</td>
	</TR>
	<tr>
		<td align="right">到職日<br>Nvx</td>
		<td ><%=emp_indate%></td> 
		<td align="right">單位<br>Bo phan</td>
		<td  ><%=emp_group%></td>
	</tr>
	<tr>
		<td align="right">職務<br>Chuc vu</td>
		<td colspan=3><%=emp_job%></td>
	</tr>
	<tr><td colspan=4></td></tr>
	<tr><td colspan=4></td></tr>	
</table>	
 
<hr size=0	style='border: 1px dotted #999999;' align=left width=600> 

<table width="630" border="1" cellspacing="1" cellpadding="1" class="txt8">
<tr bgcolor="#FFFFCB">
		<td width=40 nowrap align="center">項次</td>
		<td colspan=4>評核主項目</td>				
		<td width=35 nowrap  align="center">評分</td>
</tr>
<tr>
		<td width=40 nowrap></td>
		<td width=30 nowrap bgcolor="#CFF3CB">項次</td>
		<td width=70 nowrap bgcolor="#CFF3CB">評語</td>
		<td width=45 nowrap bgcolor="#CFF3CB">分數</td>
		<td bgcolor="#CFF3CB" colspan=2 >評核項目(中)</td>		 
</tr>
<%for x = 1 to PageRec 
	m_title = tmprecB(1,x,1) 
	if len(m_title)=3 then 	
		M_stt = mid("一二三四五六七八九十", right(m_title,1) , 1)  
	else
		m_stt=""
	end if 	

grade=tmprecB(1,x,2) 
fensu=tmprecB(1,x,3) 
khstr_cn=tmprecB(1,x,4) 
khstr_vn=tmprecB(1,x,5)  	
cqfensu=tmprecB(1,x,13)  	
'if cqfensu="" or cqfensu="0" then 

totfs = totfs+cdbl(cqfensu)
'end if  
if mylevel="Z" then 
	cq_kj=tmprecB(1,x,20)  	
	cqmemos=	tmprecB(1,x,14)  	
elseif	mylevel="J" then 
	cq_kj=tmprecB(1,x,21)  		
	cqmemos=tmprecB(1,x,15)  	
elseif	mylevel="H" then 
	cq_kj=tmprecB(1,x,22)  			
	cqmemos=tmprecB(1,x,16)
else
	cq_kj=""	
	cqmemos=""
end if 	



%>
<%if len(m_title)=3 then %>
<tr bgcolor="#FFFFCB">
	<Td align="center" valign="top"><%=m_stt%></td>
	<td colspan=4 style="cursor:hand" onclick="view_d(<%=x-1%>)"><%=khstr_cn%><Font color="blue">( <%=fensu%>&nbsp;分Điểm)</font><br><%=khstr_vn%>
		<input type="hidden"  name="M_khstr_CN" class="inputbox8" size=50 >
		<input type="hidden" name="M_khstr_VN" class="inputbox8" size=50 > 
	</td>	
	<td  width=45  align="right" >
		<input name="cqfensu" class="inputbox8" size=5 style="text-align:center;" value="<%=cqfensu%>" onblur="cqfensuchg(<%=x-1%>)">
	</td>  
</tr>
<%end if %>
<%if len(m_title)=5 then 
grade=tmprecB(1,x,2) 
fensu=tmprecB(1,x,3) 
khstr_cn=tmprecB(1,x,4) 
khstr_vn=tmprecB(1,x,5) 
%>

<tr id="<%=tmprecB(1,x,1)%>" style="display:'none'">
	<Td  width=40 nowrap ></td>
	<Td align="center" bgcolor="#CFF3CB" width=30 nowrap >
		<%=right(m_title,2)%>
	</td>
	<td bgcolor="#CFF3CB"  ><%=grade%></td>	
	<td bgcolor="#CFF3CB" align="center"><%=fensu%></td>	
	<td bgcolor="#CFF3CB"  ><%=khstr_cn%><br><%=khstr_vn%>
	<input type="hidden" name="cqfensu" class="inputbox8" size=5 style="text-align:center;" value="0"> 
	</td>	 
</tr>
<%end if %>

<input name="sttno" value="<%=tmprecB(1,x,1)%>" type="hidden" >
<input name="fs_set" value="<%=tmprecB(1,x,3)%>" type="hidden" >
<%next%> 
<input name="cqfensu" value="0" type="hidden" > 

<table width="600" border="0" cellspacing="1" cellpadding="1" class="txt8" bgcolor="#ccccccc" > 
	<tr bgcolor="#DBDBEF"  >
		<td  rowspan=3 align="right">考核總分<br>Tổng điểm xét duyệt</td>
		<td  rowspan=3>
			<input name="totFs" class="readonly8" style="height:40;text-align:center;" size=5 value="<%=totfs%>" readonly >
		</td>
		<Td align="center" width=72 nowrap style="cursor:hand;" onclick="kjchg(1)">A(優)</td>
		<Td align="center" width=72 nowrap style="cursor:hand;" onclick="kjchg(2)">B(良)</td>
		<Td align="center" width=72 nowrap style="cursor:hand;" onclick="kjchg(3)">C(甲)</td>
		<Td align="center" width=72 nowrap style="cursor:hand;" onclick="kjchg(4)">D(乙)</td>
		<Td align="center" width=72 nowrap style="cursor:hand;" onclick="kjchg(5)">E(丙)</td>
	</tr>
	<tr bgcolor="#DBDBEF">
		<Td align="center"><input name="kj1" class="inputbox" size=5 style="cursor:hand;text-align:center;" maxlength=1 readonly onclick="kjchg(1)" <%if cq_kj="A" then%>value="V"<%end if%>></td>
		<Td align="center"><input name="kj1" class="inputbox" size=5 style="cursor:hand;text-align:center;" maxlength=1 readonly onclick="kjchg(2)" <%if cq_kj="B" then%>value="V"<%end if%>></td>
		<Td align="center"><input name="kj1" class="inputbox" size=5 style="cursor:hand;text-align:center;" maxlength=1 readonly onclick="kjchg(3)" <%if cq_kj="C" then%>value="V"<%end if%>></td>
		<Td align="center"><input name="kj1" class="inputbox" size=5 style="cursor:hand;text-align:center;" maxlength=1 readonly onclick="kjchg(4)" <%if cq_kj="D" then%>value="V"<%end if%>></td>
		<Td align="center"><input name="kj1" class="inputbox" size=5 style="cursor:hand;text-align:center;" maxlength=1 readonly onclick="kjchg(5)" <%if cq_kj="E" then%>value="V"<%end if%>></td>
	</tr>	
	<tr bgcolor="#DBDBEF">
		<Td align="center"> >=90</td>
		<Td align="center"> >=85  </td>
		<Td align="center"> >=80 </td>
		<Td align="center"> >=70 </td>
		<Td align="center"> < 70 </td>
	</tr>
</table>
<table width="600" border="0" cellspacing="1" cellpadding="1" class="txt8" bgcolor="#E4e4e4" > 
	<tr bgcolor="#ffffff" height="22">
		<td rowspan=2 align="center">考勤(扣減點數)<br>Chuyên cần  Khấu trừ số điểm</td>
		<td align="center" width=72 nowrap>病假</td>
		<td align="center" width=72 nowrap>事假</td>
		<td align="center" width=72 nowrap>曠職</td>
		<td align="center" width=72 nowrap>遲到</td>
		<td align="center" width=72 nowrap>合計</td>
	</tr>
	<tr bgcolor="#ffffff" height="22">
		<Td align="center"><%=jiaB%></td>
		<Td align="center"><%=jiaA%></td>
		<Td align="center"><%=kzhour%></td>
		<Td align="center"><%=flz%></td>
		<Td align="center">&nbsp;</td>
	</tr>
	</tr>
</table>
<table width="600" border="0" cellspacing="1" cellpadding="1" class="txt8" bgcolor="#E4e4e4" > 
	<tr bgcolor="#ffffff" height="22">
		<td align="center" >備註說明<br>Gui chu<BR>考核<>C(甲等)需說明</td> 	
		<Td align="center"><textarea class="textarea" rows=3 cols=75 name="cqmemos"><%=cqmemos%></textarea></td>		
		<input name="end_Kj"type="hidden" value="<%=cq_kj%>">
	</tr>
	</tr>
</table>
<br>  
<TABLE WIDTH=600>
	<tr ALIGN=center>
	<TD >
		<input type=button  name=send value="(OK)完成考核" class=button onclick="go()"> 
		
		<input type=button  name=send value="(M)回主畫面" class=button onclick=backM()>
		
	</TD>
	</TR>
</TABLE>


</form>


</body>
</html>
<script language=vbscript> 

function khemp(index)
	empidstr = <%=self%>.empid(index).value	
	khyearud = <%=self%>.khyearsUd.value 
	
	open "<%=self%>B.Foregnd.asp?index="&index &"&empid="& empidstr &"&khyear="& khyearud , "Back" 
	parent.best.cols="50%,50%"
end function  

function view_d(index)
	s_name =  <%=self%>.sttno(index).value  	
	strtemp="文件中所有的標籤名稱TR的："
  set objall= document.all.tags("tr")
  strtempno="文件中標籤名稱 TR 的總數：" & objall.length 	
	' alert(strtempno)  
	d_no =5 '細項共5個   
	'alert index 
	for  inti = 1 to d_no 
		if index = 0 then 
			zz = inti + 9   
			if objall.item(zz).style.display="" then 
				objall.item(zz).style.display="none" 
			else
				objall.item(zz).style.display="" 
			end if  			
		else
			zz = inti + index+9  
			if objall.item(zz).style.display="" then 
				objall.item(zz).style.display="none" 
			else
				objall.item(zz).style.display="" 
			end if 		
			'alert zz 
		end  if 
	next  
end function  

function kjchg(a)
	for xx =1 to 5 
		if  xx <>a then 
			<%=self%>.kj1(xx-1).value="" 
		else
			<%=self%>.kj1(xx-1).value="V" 
		end if 
	next 
	<%=self%>.end_Kj.value=mid("ABCDE",a,1)
end function 

function cqfensuchg(index)
	if <%=self%>.cqfensu(index).value<>""  then 
		if isnumeric(<%=self%>.cqfensu(index).value)=false  then 
			alert "請輸入數字xin danh lai [so] !!" 
			<%=self%>.cqfensu(index).value="0"
			<%=self%>.cqfensu(index).select()
			exit function 
		elseif instr(	<%=self%>.cqfensu(index).value,".")>0 or cdbl(<%=self%>.cqfensu(index).value)<0 then 
			alert "請輸入數字xin danh lai [so] !!" 
			<%=self%>.cqfensu(index).value="0"
			<%=self%>.cqfensu(index).select()
			exit function 
		elseif 	cdbl(<%=self%>.cqfensu(index).value)>cdbl(<%=self%>.fs_set(index).value) then 
			alert "本項總分  不可 (khong duoc ) >  ["& <%=self%>.fs_set(index).value &"] 分"			 
			<%=self%>.cqfensu(index).value="0"
			<%=self%>.cqfensu(index).select()
			exit function   
		end if 
	end if 	
	calcTotfs()
end function  

function calcTotfs()
	for  i = 1 to <%=self%>.pagerec.value 
		totfs = cdbl(totfs) + <%=self%>.cqfensu(i-1).value  		
	next 
	<%=self%>.totFs.value=totfs
	if cdbl(totfs)>=90 then 
		<%=self%>.kj1(0).value="V"
	elseif	cdbl(totfs)>=85 then 
		<%=self%>.kj1(1).value="V"
	elseif	cdbl(totfs)>=80 then 
		<%=self%>.kj1(2).value="V"
	elseif	cdbl(totfs)>=70 then 
		<%=self%>.kj1(3).value="V"
	elseif	cdbl(totfs)<70 then 
		<%=self%>.kj1(4).value="V"			
	end if 	 

	for j =1 to 5 
		if cdbl(totfs)>=90 then 
			a = 1 			
		elseif	cdbl(totfs)>=85 then   
			a=2 	
		elseif	cdbl(totfs)>=80 then   
			a=3
		elseif	cdbl(totfs)>=70 then   
			a=4
		elseif	cdbl(totfs)<70then   
			a=5	
		end if 	
	next  
	
	for xx =1 to 5 
		if  xx <>a then 
			<%=self%>.kj1(xx-1).value="" 
		else
			<%=self%>.kj1(xx-1).value="V" 
		end if 
	next 
	<%=self%>.end_Kj.value=mid("ABCDE",a,1)  

end function 
 
function  backM()	
	open "<%=self%>.asp", "_self"
	parent.best.cols="100%,0%"
	
end function  

function go()
	if trim(<%=self%>.totFs.value)="" or trim(<%=self%>.totFs.value)="0" then 
		alert "請評分!!,khong co Tổng điểm xét duyệt!!"
		exit function 
	end if 

	if trim(<%=self%>.end_Kj.value)="" then 
		alert "請評考績!!,khong co Xét bậc !!"
		exit function 
	end if  
	
	if trim(<%=self%>.totFs.value)<>"" and trim(<%=self%>.end_Kj.value)<>"" then 
		if  <%=self%>.end_Kj.value<>"C" and <%=self%>.cqmemos.value=""  then 
			alert "考績非 C(甲)等,請於備註說明, xin danh lai ly do(ghi chu), Xét bậc khong phi [C] !!"
			<%=self%>.cqmemos.focus()
			exit function 
		else	
			parent.Fore.YEIE0201.khyn(<%=index%>).value="V"
			parent.Fore.YEIE0201.funsu(<%=index%>).value=<%=self%>.totFs.value
			parent.Fore.YEIE0201.grade(<%=index%>).value=<%=self%>.end_Kj.value
			parent.best.cols="100%,0%"
			<%=self%>.action = "<%=self%>.insDB.asp"
			<%=self%>.target = "Back"
			<%=self%>.submit()
		end if 		
	end if 
	
end function 
  
</script>


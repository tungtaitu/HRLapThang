<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<!-- #include file = "../GetSQLServerConnection.fun" -->
<!-- #include file="../ADOINC.inc" -->

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="refresh">
<%
self="YECE06"

Set conn = GetSQLServerConnection()

nowmonth = year(date())&right("00"&month(date()),2)
if month(date())="01" then
	calcmonth = year(date()-1)&"12"
else
	calcmonth = year(date())&right("00"&month(date())-1,2)
end if

if day(date())<=11 then
	if month(date())="01" then
		calcmonth = year(date()-1)&"12"
	else
		calcmonth = year(date())&right("00"&month(date())-1,2)
	end if
else
	calcmonth = nowmonth
end if


JXYM=REQUEST("JXYM")
salaryYM=REQUEST("salaryYM")
country=REQUEST("country")
JOBID=REQUEST("JOBID")
empid1 = REQUEST("empid1")
F_GROUPID = REQUEST("F_GROUPID")
F_SHIFT=REQUEST("F_SHIFT")
F_zuno = REQUEST("F_zuno")
F_whsno=request("F_whsno")

'response.write groupid &"<BR>"

Edays=left(salaryYM,4)&"/"&right(salaryYM,2)&"/01"  
e_edat = left(salaryYM,4)&"/"&right(salaryYM,2)&"/25"


TotalPage = 10
PageRec = 10  'number of records per page
TableRec = 50  'number of fields per record 

salaryYMDate = left(salaryYM,4)&"/"&right(salaryYM,2)&"/01"
days = DAY(cDatestr+(32-DAY(salaryYMDate))-DAY(cDatestr+(32-DAY(salaryYMDate))))   '月有幾天
enddays=left(salaryYM,4)&"/"&right(salaryYM,2)&"/"&right("00"&cstr(days),2)

JXYMdays=left(JXYM,4)&"/"&right(JXYM,2)&"/01"
JXdays = DAY(cDatestr+(32-DAY(JXYMdays))-DAY(cDatestr+(32-DAY(JXYMdays))))   '月有幾天
JXenddays=left(JXYM,4)&"/"&right(JXYM,2)&"/"&right("00"&cstr(JXdays),2)

' response.write "JXYMdays=" & JXYMdays &"<BR>"
' response.write "JXenddays=" & JXenddays &"<BR>"
'response.write "Edays=" & Edays &"<BR>" 

'response.end  

sqln="SELECT * FROM VYFYEXRT WHERE YYYYMM='"& jxym &"' and code='USD' "
set rsx=conn.execute(Sqln)
if not rsx.eof then 
	rate=rsx("exrt")
end if 	
set rsx=nothing 

sortby = request("sortby")
if request("sortby")="" then sortby="len(ls) desc, ls, lz, empid"  

sql="select * from fn_yece06 ('"& JXYM &"','"& salaryYM &"','"& country &"','"& F_whsno &"','"& F_GROUPID &"','"& F_zuno &"','"& F_SHIFT &"','"& empid1 &"','"& JXYMdays &"','"& JXenddays &"' )  "
sql = sql & "order by "& sortby	 

'response.write sql 
'response.end  
if request("totalpage")="" or request("totalpage")="0" then 
currentpage=1 
Set rs = Server.CreateObject("ADODB.Recordset")
RS.OPEN SQL,CONN, 1, 3  

if not rs.eof then 
	pagerec = rs.RecordCount
	rs.PageSize = pagerec
	RecordInDB = rs.RecordCount
	TotalPage = rs.PageCount
	lgstr = rs("lgstr")
end if 	
Redim tmpRec(TotalPage,PageRec, TableRec)  'Array
for i = 1 to TotalPage
	for j = 1 to PageRec
		if not rs.EOF then
			 
			tmpRec(i, j, 0)="no"
			tmpRec(i, j, 1)=trim(rs("EMPID"))
			tmpRec(i, j, 2)=trim(rs("LG"))  '績效年月時之單位
			tmpRec(i, j, 3)=trim(rs("LS"))  '績效年月時之班別
			tmpRec(i, j, 4)=rs("Jxbonus")
			tmpRec(i, j, 5)=rs("EMPNAM_CN")
			tmpRec(i, j, 6)=rs("EMPNAM_VN")
			tmpRec(i, j, 7)=rs("nindat")
			if JXYM>"200702" then 
				if  rs("country")="CN" then 
					tmpRec(i, j, 8) = 0
				else
					tmpRec(i, j, 8)= cdbl(RS("KZHOUR"))
				end if 	
			else
				if  rs("country")="CN" then 
					tmpRec(i, j, 8) = 0
				else
					tmpRec(i, j, 8)= rs("latefor")+rs("forget")
				end if 	
			end if	
			'tmpRec(i, j, 9)= 0 'rs("SUmoney") 
			if JXYM<="201010" then   '自計算201011月績效獎金開始..如11月有事故扣款..扣在績效獎金,不扣在薪資  change by elin 20101128
				tmpRec(i, j, 9) =  0
			else	
				tmpRec(i, j, 9) = rs("dmsgKM")   
			end if	
			
			tmpRec(i, j, 10)= rs("oldjob")
			tmpRec(i, j, 11)=left(rs("oldjobdesc"),4)

			F1_TOTJX = 0
			'for z = 1 to icount
			'	'共有幾項績效項目
			'	tmpRec(i, j,11+z)= round( (cdbl(rs("Jxbonus"))* ( cdbl(arrays(z,9))/100 ) ) * cdbl(arrays(z,8)) , 0) 
			'	F1_TOTJX = F1_TOTJX + cdbl(tmpRec(i, j, 11+z))
			'	'response.write tmpRec(i, j, 1) &"-" & 11+Z &"-" & tmpRec(i, j, 11+z) &"-" & arrays(z,9) &"-" &arrays(z,8) &"<BR>"
			'next
			
			'mark by  2009/09/28
			' tmpRec(i, j, 12) =  round(cdbl(rs("hesoA"))*(cdbl(rs("perA"))/100),0)
			' tmpRec(i, j, 13) =  round(cdbl(rs("hesoB"))*(cdbl(rs("perB"))/100),0)
			' tmpRec(i, j, 14) =  round(cdbl(rs("hesoC"))*(cdbl(rs("perC"))/100),0)
			' tmpRec(i, j, 15) =  round(cdbl(rs("hesoD"))*(cdbl(rs("perD"))/100),0)
			' tmpRec(i, j, 16) =  round(cdbl(rs("hesoE"))*(cdbl(rs("perE"))/100),0) 
			
			tmpRec(i, j, 12) =  round(cdbl(rs("STT_A_thuong")),0)
			tmpRec(i, j, 13) =  round(cdbl(rs("STT_B_thuong")),0)
			tmpRec(i, j, 14) =  round(cdbl(rs("STT_C_thuong")),0)
			tmpRec(i, j, 15) =  round(cdbl(rs("STT_D_thuong")),0)
			tmpRec(i, j, 16) =  round(cdbl(rs("STT_E_thuong")),0)			
			
			'應領獎金
			F1_TOTJX = cdbl(tmpRec(i, j, 12))+cdbl(tmpRec(i, j, 13))+cdbl(tmpRec(i, j, 14))+cdbl(tmpRec(i, j, 15))+cdbl(tmpRec(i, j, 16))			
			F1_TOTJX = cdbl(rs("N_totjx"))   			 
			F1_TOTJX = round(cdbl(rs("N_totjx"))*cdbl(rs("newjs")) ,0)
			'add elin 201202  
			f1_tothrs = rs("tothrs")   '部門應出勤工時
			f1_empwkhrs = cdbl(rs("toth"))+cdbl(rs("jiaEhr"))+cdbl(rs("jiaGhr"))  '員工實際出勤工時+年假+公假 
			IF RS("COUNTRY")="VN" THEN  							
				if f1_tothrs = 0 then 
					heso3 = 1 
				else
					heso3 = cdbl(f1_empwkhrs)/cdbl(f1_tothrs) 
				end if 	
			else
				heso3=1.0 
			end if 	
'response.write 	F1_TOTJX &"<BR>"		
'response.write 	heso3 &"<BR>" 
			F1_TOTJX = round(cdbl(F1_TOTJX)*cdbl(round(heso3,2)),0)
			'response.write 	F1_TOTJX &"<BR>" 
			'曠職應扣款
			if tmpRec(i, j, 8)< 8 then  
				tmpRec(i, j, 17)= 0
			elseif cdbl(tmpRec(i, j, 8))>=8 and cdbl(tmpRec(i, j, 8))<16 then
				tmpRec(i, j,17) = round(cdbl(F1_TOTJX)*0.5,0) '曠職ㄧ天以上...扣績效獎金一半
			elseif cdbl(tmpRec(i, j, 8))>=16 then
				tmpRec(i, j, 17)= cdbl(F1_TOTJX)  ''曠職2天以上...無績效獎金
			else	
				tmpRec(i, j, 17) = 0   '應扣款
			end if 			 
			'response.write "4ykm=" & tmpRec(i, j, 17) &"<BR>" 
			
			tmpRec(i, j, 18) =  round( cdbl(rs("N_totjx"))*cdbl(rs("newjs"))*cdbl(round(heso3,2)) ,0) 'F1_TOTJX
			
			
			'tmpRec(i, j, 14+icount) = round(cdbl(F1_TOTJX)-cdbl(tmpRec(i, j,12+icount))-cdbl(tmpRec(i, j, 9)),0)
			tmpRec(i, j, 20 ) = rs("outdate")
			IF RS("COUNTRY")="CN" THEN   'no use
				tmpRec(i, j, 21) = 1 
			ELSE 
				tmpRec(i, j, 21) = 1   
			END IF 	
 
			IF left(replace(rs("outdate"),"/",""),6)>=JXYM and trim(rs("outdate"))< e_edat  then 
				tmpRec(i, j, 19) = 0
			else
				if jxym>"200702" then 
					if cdbl(tmpRec(i, j, 21))>=1 then 
						tmpRec(i, j, 19) =  cdbl(F1_TOTJX) '( round(cdbl(F1_TOTJX)-cdbl(tmpRec(i, j,17))-cdbl(tmpRec(i, j, 9)),0)) 
					else	
						tmpRec(i, j, 19)  =  cdbl(F1_TOTJX) ' round(  cdbl(tmpRec(i, j, 21)) *round(cdbl(F1_TOTJX)-cdbl(tmpRec(i, j,17))-cdbl(tmpRec(i, j, 9)),0) ,0)
					end if 	
				else  '<="200702" 
					tmpRec(i, j, 19) = round(cdbl(F1_TOTJX)-cdbl(tmpRec(i, j,17))-cdbl(tmpRec(i, j, 9)),0) 
				end if 
			end if 			
			
			if rs("jxyn")="" or rs("jxyn")="N" then 
				tmpRec(i, j, 19) = 0 
			end if 
			
			tmpRec(i, j, 23 ) = rs("groupid")  '目前單位
			if rs("shift")="" then 
				tmpRec(i, j, 24 ) = rs("LS")  '計績效時班別
			else
				tmpRec(i, j, 24) = rs("shift")  '目前班別
			end if	
			
			tmpRec(i, j, 25 ) = rs("zuno") '目前組別
			tmpRec(i, j, 26 ) =rs("LZ")  '績效年月的組別
			tmpRec(i, j, 27 ) =rs("monfen")  '考核分數		
			

			if  rs("country")="VN" then 
				if cdbl(tmpRec(i, j, 27))<70 then 					
					tmpRec(i, j, 17) = tmpRec(i, j, 17) + round(cdbl(F1_TOTJX) *0.5 ,0)				
				else
					tmpRec(i, j, 17) = tmpRec(i, j, 17)
				end if 	
			else
				'tmpRec(i, j, 19) = tmpRec(i, j, 19)
				tmpRec(i, j, 17) = tmpRec(i, j, 17)
			end if 			
			
			'response.write "1ykm=" & tmpRec(i, j, 17) &"<BR>"
						
			tmpRec(i, j, 28) =rs("LZstr")  '績效年月的組別
			tmpRec(i, j, 29) = rs("DM")
			tmpRec(i, j, 30) = cdbl(rs("exrt"))
			'2009新增 
			tmpRec(i, j, 31) = cdbl(rs("jiaAhr"))
			
			'201802新增病假扣款  
			tmpRec(i, j, 44) = cdbl(rs("jiabhr")) 
			
			'Response.write  rs("empid") &"-" & tmpRec(i, j, 44)&"<BR>"
			
			'事假扣款
			if rs("country")="VN" then 
				if  cdbl(tmpRec(i, j, 31))>=24 and cdbl(tmpRec(i, j, 31))<40 then   '事假3天以上未滿5天 績效減半				
					tmpRec(i, j,17) = cdbl(tmpRec(i, j,17)) + round(cdbl(F1_TOTJX)*0.5,0) 
				elseif cdbl(tmpRec(i, j, 31)) >= 40	 then   '事假5天含以上 績效為 0 				
					tmpRec(i, j,17) = round( cdbl(tmpRec(i, j,17)) +  round(cdbl(F1_TOTJX),0) , 0) 
				else	
					tmpRec(i, j, 17) = round( tmpRec(i, j, 17) , 0)  
				end if 	  
			else 
				tmpRec(i, j,17) = cdbl(tmpRec(i, j,17)) + round( (cdbl(F1_TOTJX)/30)*round(cdbl(tmpRec(i, j, 31))/8,1),0) ' 按天數扣款
			end if  
			'201802新增病假扣款
			if rs("country")="VN" then 
				if  cdbl(tmpRec(i, j, 44))>=24 and cdbl(tmpRec(i, j, 44))<40 then   '病假3天以上未滿5天 績效減半				
					tmpRec(i, j,17) = cdbl(tmpRec(i, j,17)) + round(cdbl(F1_TOTJX)*0.5,0) 
				elseif cdbl(tmpRec(i, j, 44)) >= 40	 then   '病假5天含以上 績效為 0 				
					tmpRec(i, j,17) = round( cdbl(tmpRec(i, j,17)) +  round(cdbl(F1_TOTJX),0) , 0) 
				else	
					tmpRec(i, j, 17) = round( tmpRec(i, j, 17) , 0)  
				end if  
			end if 
			
			'response.write "3ykm=" & tmpRec(i, j, 17) &"<BR>" 
			'勸導單
			if rs("rp_cnt")>="1" then  '勸導書1張扣獎金一半
				tmpRec(i, j,17) = cdbl(tmpRec(i, j,17)) + round(cdbl(F1_TOTJX)*0.5,0) 
			elseif	rs("rp_cnt")>="2" then  '2張或以上無獎金
				tmpRec(i, j,17) = round(cdbl(F1_TOTJX),0) 
			end  if 	 
			
			
			'response.write "2ykm=" & tmpRec(i, j, 17) &"<BR>" 
			' if cdbl(tmpRec(i, j,17)) > cdbl(F1_TOTJX) then 
				' tmpRec(i, j,17) =  round( cdbl(F1_TOTJX) , 0) 
			' end if 	
			' cdbl(tmpRec(i, j,9))    = 事故扣款 
			if tmpRec(i, j, 23 )<>"A051" then   			
				if cdbl(tmpRec(i, j, 19))-cdbl(tmpRec(i, j,17))-cdbl(tmpRec(i, j,9))<0 then 
					tmpRec(i, j, 19) = 0 
				else	
					tmpRec(i, j, 19) = cdbl(tmpRec(i, j, 19))-cdbl(tmpRec(i, j,17))-cdbl(tmpRec(i, j,9))
				end if
			end if 	
			tmpRec(i, j, 32) = rs("jxyn")  '是否計績效
			tmpRec(i, j, 33) = rs("rp_cnt")  '勸導書張數
			tmpRec(i, j, 34) = rs("file_totjxm")  
			tmpRec(i, j, 35) = rs("newJs")     '' 工作系數 
			tmpRec(i, j, 36) = rs("country")      
			tmpRec(i, j, 37) = rs("tothrs")      '部門應出勤工時
			tmpRec(i, j, 38) = rs("toth")   '員工實際出勤工時    
			tmpRec(i, j, 39) = rs("jiaEhr") '年假
			tmpRec(i, j, 40) = rs("jiaGhr")	'公假
			tmpRec(i, j, 41) = cdbl(rs("toth"))+cdbl(rs("jiaGhr"))+cdbl(rs("jiaEhr"))
			tmpRec(i, j, 42) = round(heso3,2)
			tmpRec(i, j, 43) = rs("lgstr")
			rs.MoveNext
		else
			exit for
		end if
	next
NEXT
Session("YFYEMPJXM") = tmpRec 
end if  
%>
<link rel="stylesheet" href="../Include/style.css" type="text/css">
<link rel="stylesheet" href="../Include/style2.css" type="text/css">
</head>
<body  leftmargin="5" marginwidth="0" marginheight="0" onkeydown="enterto()" >
<form name="<%=self%>" method="post" action="<%=self%>.upd.asp">
<INPUT TYPE="hidden" NAME=TotalPage VALUE="<%=TotalPage%>">
<INPUT TYPE="hidden"  NAME=CurrentPage VALUE="<%=CurrentPage%>">
<INPUT TYPE="hidden" NAME=PageRec VALUE="<%=PageRec%>">
<INPUT TYPE="hidden" NAME=RecordInDB VALUE="<%=RecordInDB%>">
<INPUT TYPE="hidden" NAME="ICOUNT"   VALUE="<%=ICOUNT%>">
<INPUT TYPE="hidden" NAME="F_whsno"   VALUE="<%=F_whsno%>">
<INPUT TYPE="hidden" NAME="empid1"   VALUE="<%=empid1%>">
 
	<table border=0 style="height:70px"><tr><td>&nbsp;</td></tr></table>
	
	<table width="100%" BORDER=0 align=center cellpadding=0 cellspacing=0 >
		<tr>						
			<td>
				<table class="txt" cellpadding=3 cellspacing=3>
					<tr class="txt">
						<td ALIGN=right nowrap>績效年月<br><font class="txt8">Tien Thuong</font></td>
						<td><INPUT type="text" style="width:100px" NAME="JXYM" VALUE="<%=JXYM%>"  READONLY ></td>
						<td ALIGN=RIGHT>計薪年月<br><font class="txt8">Tien Luong</font></td> 
						<td><INPUT type="text" style="width:100px" NAME="SALARYYM" VALUE="<%=SALARYYM%>"    READONLY></td>
						<td ALIGN=RIGHT>部門<br><font class="txt8">Bo phan</font></td>
						<td colspan=5>
							<INPUT type="hidden" NAME="F_groupid" VALUE="<%=F_GROUPID%>"  READONLY size=5 >
							<INPUT type="text" style="width:120px" NAME="GSTR" VALUE="<%=lgstr%>"  READONLY >				
						</td>	
					</tr>
					<tr>
						<td ALIGN=RIGHT>班別<br><font class="txt8">Ca</font></td>
						<td>
							<select name=F_shift  onchange="dchg()" style="width:100px">
							<option value="">--</option>
							<%sqlt="select * from basicCode where func='shift' order by len(sys_type) desc, sys_type"
							set rds=conn.execute(sqlt)
							while not rds.eof
							%>
							<option value="<%=rds("sys_type")%>" <%if F_shift=rds("sys_type") then%>selected<%end if%>><%=rds("sys_type")%></option>
							<%rds.movenext
							wend
							set rds=nothing
							%>
							</select>
						</td>	
						<td ALIGN=RIGHT>單位<br><font class="txt8">To</font></td>
						<td colspan=3>
							<select name=F_zuno   onchange="dchg()" style="width:100px">				
							<option value="">--</option>
							<%sqlt="select * from basicCode where func='zuno' and left(sys_type,4) like '"& f_groupid&"'  order by sys_type"
							set rds=conn.execute(sqlt)
							while not rds.eof
							%>
							<option value="<%=rds("sys_type")%>" <%if f_zuno=rds("sys_type") then%>selected<%end if%>><%=rds("sys_type")%><%=rds("sys_value")%></option>
							<%rds.movenext
							wend
							set rds=nothing
							%>
							</select>
						</td>			
					 
						<td ALIGN=RIGHT>排序<BR><font class="txt8">Sap xep</font></td>
						<td  >
							<select name=sortby   onchange="dchg()" style="width:100px">
								<option value="len(ls) desc, ls, lz, country,empid" <%if sortby="len(ls) desc, ls, lz, country,empid" then%>selected<%end if%>>(1)依班別(Ca)</option>
								<option value="country, empid"  <%if sortby="country, empid" then%>selected<%end if%>>(2)依工號(so the)</option>
								<option value="lz,len(ls) desc, ls, country,empid" <%if sortby="lz,len(ls) desc, ls, country,empid" then%>selected<%end if%>>(3)依單位(to)</option>					
							</select>
						</td>	 
					</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td>
				<table  class="txt" cellpadding=3 cellspacing=3>
					<tr>
						<td style="width:20px">&nbsp;</td>
						<td class="txt8" nowrap>
							<font color="#cc0000">*(màu đỏ biểu thị không tính tiền thưởng)紅色表示不計績效</font> &nbsp;rate:=<%=rate%>
								&nbsp;&nbsp;更新所有績效係數=<input type="text" name="njs" value="" onblur="njschg()" style="width:100px;text-align:center">
						</td>
					</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td>
				<table id="myTableGrid" width="98%">
					<TR BGCOLOR="#FEF7CF" CLASS=TXT8>
						<TD width=30  ALIGN=CENTER nowrap>算績效<br>tính t.t</TD>
						<TD width=50  ALIGN=CENTER nowrap>工號<br>So the</TD>
						<TD width=100 ALIGN=CENTER nowrap>姓名<br>Ho Ten</TD>
						<TD width=70 ALIGN=CENTER nowrap>到職日<br>NVX<br>NTV</TD>
						<TD width=60 ALIGN=CENTER  nowrap>部門/班<BR>Bo phan/Ca</TD>
						<TD width=60 ALIGN=CENTER  nowrap>職等<br>Chuc vu</TD>			
						<TD width=40 ALIGN=CENTER nowrap>職務<BR>係數<br>He so</TD>
						<td  width=90 align=center>A</td>
						<td  width=90 align=center>B</td>
						<td  width=90 align=center>C</td>
						<td  width=90 align=center>D</td>
						<td  width=90 align=center>E</td>
						<TD width=40 ALIGN=CENTER nowrap>績效<BR>係數<br>He so2</TD>
						<TD width=40 ALIGN=CENTER nowrap>工時<BR>係數<br>He so3</TD>
						<TD width=40 ALIGN=CENTER nowrap>實際與<br>應出勤<BR>工時</TD>
						<td  width=100 align=center>小計<br>Total</td>
						<td width=30 align=center nowrap>曠職(H)<BR>kC(H)</td>
						<td width=30 align=center nowrap>事/病假(H)<BR>VR(H)</td>
						<td width=40 nowrap align=center>勸導書<br>Giay<br>nhac<br>nho</td>
						<td width=40 nowrap align=center>考核<br>Diem<br>nang<br>suat</td>
						<td width=60 align=center>應扣款<br>TOT TRU</td>
						<td width=40 nowrap align=center>反規定<br>Pham<br>quy<br>dinh</td>			
						<td width=60 nowrap align=center><font color="blue">事故扣款<br>TRU su co</font></td>			
						<td width=60 nowrap align=center>合計<br>Total</td>
						<td width=40 nowrap align=center>幣別<br>Loai Tien</td>
						<td width=40 nowrap align=center>合計<BR>Total (USD)</td>
						 
					</TR>
						<%for CurrentRow = 1 to PageRec
							IF CurrentRow MOD 2 = 0 THEN
								WKCOLOR="#FFFFFF"
							ELSE
								WKCOLOR="#FFFFFF"
							END IF
							if tmpRec(CurrentPage,  CurrentRow, 32)<>"Y" then
								ft_color="#cc0000"
							else
								ft_color="#000000"
							end if 	 
							if cdbl(tmpRec(CurrentPage,  CurrentRow, 34))<>cdbl(tmpRec(CurrentPage,  CurrentRow, 19)) then
								ft_clr="#cc0000"
							else
								ft_clr="blue"
							end if 
							if tmpRec(CurrentPage,  CurrentRow,  1)<>"" then
						%>
								<TR id=id1 bgColor="<%=wkcolor%>"  valign="top" align="center"  CLASS=TXT8>
									<td>
										<input name="fn2" type="checkbox"  <%if tmpRec(CurrentPage,CurrentRow, 32)="Y" then%>checked<%end if%> onclick="fn2chg(<%=CurrentRow-1%>)">
										<input name="jxyn" type="hidden" value="<%=tmpRec(CurrentPage,CurrentRow, 32)%>">
									</td>
									<TD nowrap align=center style="cursor:pointer" onclick="vbscript:oepnEmpWKT(<%=currentRow-1%>)"  >
										<font color="<%=ft_color%>"><%=tmpRec(CurrentPage,CurrentRow, 1)%></font>
										<input type="hidden" name="empid" value="<%=tmpRec(CurrentPage,CurrentRow, 1)%>">
										<input type="hidden" name="NowGroup" value="<%=tmpRec(CurrentPage,CurrentRow, 23)%>">
										<input type="hidden" name="NowShift" value="<%=tmpRec(CurrentPage,CurrentRow, 24)%>">
										<input type="hidden" name="NowZuno" value="<%=tmpRec(CurrentPage,CurrentRow, 25)%>">
										<input type="hidden" name="jxgroup" value="<%=tmpRec(CurrentPage,CurrentRow, 2)%>">
										<input type="hidden" name="jxShift" value="<%=tmpRec(CurrentPage,CurrentRow, 3)%>">
										<input type="hidden" name="jxZuno" value="<%=tmpRec(CurrentPage,CurrentRow, 26)%>">
										<input type="hidden" name="ct" value="<%=tmpRec(CurrentPage,CurrentRow, 36)%>"><!--country-->
									</TD>
									<TD nowrap style="cursor:pointer" onclick="vbscript:oepnEmpWKT(<%=currentRow-1%>)">
										<font color="<%=ft_color%>"><%=tmpRec(CurrentPage,CurrentRow, 5)%><BR><%=left(tmpRec(CurrentPage, CurrentRow, 6),15)%></font>
									</TD>
									<TD nowrap align=center valign="top"><%=tmpRec(CurrentPage,CurrentRow,7)%><br><font color=red><%=tmpRec(CurrentPage,CurrentRow,20)%></font></TD>
									<TD nowrap valign="top"><font color="<%=ft_color%>">
										<%=tmpRec(CurrentPage,CurrentRow,24)%><br>
										<%=(tmpRec(CurrentPage,CurrentRow,28))%>
										</font>
									</TD>
									<TD nowrap valign="top" align="left">
										<%=tmpRec(CurrentPage,CurrentRow,10)%>-
										<%=tmpRec(CurrentPage,CurrentRow,11)%>
									</TD>
									<TD nowrap align=center valign="top">
										<%=tmpRec(CurrentPage, CurrentRow, 4)%>_n
										<input name="jobjs"  type="hidden"   value="<%=tmpRec(CurrentPage,CurrentRow,4)%>"   >
									</TD> 
									<Td valign="top">
										<input type="text" style="width:100%" name="JXA"  class="inputbox8r" id="jxA" value="<%=tmpRec(CurrentPage,CurrentRow,12)%>"  onblur="chkdata('jxa',<%=currentRow-1%>)" >
										<input type="hidden" size=3 name="BJXA" id="BJXA"  value="<%=tmpRec(CurrentPage, CurrentRow,12)%>" >
									</td>	
									<Td valign="top">
										<input type="text" style="width:100%" name="JXB"  class="inputbox8r" id="jxB" value="<%=tmpRec(CurrentPage,CurrentRow,13)%>"  onblur="chkdata('jxb',<%=currentRow-1%>)" >
										<input type="hidden" size=3 name="BJXB" id="BJXB"  value="<%=tmpRec(CurrentPage, CurrentRow,13)%>" >
									</td>	
									<Td valign="top">
										<input type="text" style="width:100%" name="JXC"  class="inputbox8r" id="jxC" value="<%=tmpRec(CurrentPage,CurrentRow,14)%>"  onblur="chkdata('jxc',<%=currentRow-1%>)" >
										<input type="hidden" size=3 name="BJXC" id="BJXC"  value="<%=tmpRec(CurrentPage, CurrentRow,14)%>" >
									</td>	
									<Td valign="top">
										<input type="text" style="width:100%" name="JXD"  class="inputbox8r" id="jxD" value="<%=tmpRec(CurrentPage,CurrentRow,15)%>"  onblur="chkdata('jxd',<%=currentRow-1%>)" >
										<input type="hidden" size=3 name="BJXD" id="BJXD"  value="<%=tmpRec(CurrentPage, CurrentRow,15)%>" >
									</td>	
									<Td valign="top">
										<input type="text" style="width:100%" name="JXE"  class="inputbox8r" id="jxE" value="<%=tmpRec(CurrentPage,CurrentRow,16)%>"  onblur="chkdata('jxe',<%=currentRow-1%>)" >
										<input type="hidden" size=3 name="BJXE" id="BJXE"  value="<%=tmpRec(CurrentPage, CurrentRow,16)%>" >
									</td>	
									<TD nowrap align=center valign="top">
										<input type="text" style="width:100%" name="newJs" class="inputbox8r" id="newJs" value="<%=tmpRec(CurrentPage, CurrentRow,35)%>"  onblur="chkdata('newJs',<%=currentRow-1%>)"  > 
										<input type="hidden" size=3 name="BnewJs" id="BnewJs"  value="<%=tmpRec(CurrentPage, CurrentRow,35)%>" >
									</TD>
									<TD nowrap align=right valign="top">
										<input type="text" style="width:100%" name="hrJs" class="inputbox8r" id="hrjs" value="<%=tmpRec(CurrentPage, CurrentRow,42)%>"  onblur="chkdata('hrjs',<%=currentRow-1%>)"  >							
									</TD>
									<!--實際與英出勤工時-->
									<td align="right">
										<%=tmpRec(CurrentPage, CurrentRow,41)%><br>
										<%=tmpRec(CurrentPage, CurrentRow,37)%>
										<input type="hidden" name="relhr" value="<%=tmpRec(CurrentPage, CurrentRow,41)%>" >
										<input type="hidden" name="nedhr" value="<%=tmpRec(CurrentPage, CurrentRow,37)%>" >
									</td>
									<TD nowrap align=center valign="top">
										<input type="text" name="sumjx" class="readonly8s" value="<%=tmpRec(CurrentPage, CurrentRow,18)%>" readonly style='width:100%;text-align:right'> 
										<br><div style="right"></div>
									</TD>
									<TD nowrap align=center valign="top">
										<!--曠職H-->
										<%=tmpRec(CurrentPage, CurrentRow,8)%>
										<input type="hidden" name="FL" class="readonly8s" size=1 value="<%=tmpRec(CurrentPage, CurrentRow,8)%>" readonly style='text-align:right;<%if cdbl(tmpRec(CurrentPage, CurrentRow,8))<>0 then%>color:red<%end if%>'> 
									</TD>
									<TD nowrap align=center valign="top">
										<%=tmpRec(CurrentPage, CurrentRow,31)%><br><%=tmpRec(CurrentPage, CurrentRow,44)%>
										<input type="hidden" name="jiaAhr" class="readonly8s" size=2 value="<%=tmpRec(CurrentPage, CurrentRow,31)%>" > 
										<input type="hidden" name="jiaBhr" class="readonly8s" size=2 value="<%=tmpRec(CurrentPage, CurrentRow,44)%>" > 
									</TD>												
									<TD align=center valign="top">
										<!--勸導書-->
										<%=tmpRec(CurrentPage,  CurrentRow,33)%>
										<input type="hidden" name="gnn"  class="readonly8s" readonly size="3" value="<%=tmpRec(CurrentPage,  CurrentRow,33)%>" style='text-align:right;<%if cdbl(tmpRec(CurrentPage, CurrentRow,33))<>0 then%>color:red<%end if%>'  >							
									</TD> 
									<TD align=center valign="top">
										<%if tmpRec(CurrentPage,CurrentRow,27)< 70 then %>
											<input type="text" name="khfen"  class="readonly8s" readonly value="<%=tmpRec(CurrentPage,CurrentRow,27)%>"  style='width:100%;text-align:right;color:red' > 
										<%else%>
											<input type="hidden" name="khfen"   size="2" value="<%=tmpRec(CurrentPage,CurrentRow,27)%>" > 
										<%end if %>	
									</TD>						
									<TD nowrap align=center valign="top">	
										<!--應扣款-->
										<input type="text" name="FLmoney"  class="readonly8s" value="<%=tmpRec(CurrentPage,CurrentRow,17)%>" readonly style='width:100%;text-align:right' > 		
									</TD>
									<TD align=center valign="top">
										<input type="text" style="width:100%" name="FQD"  class="inputbox8r" value="0" onblur="FQDchg(<%=currentRow-1%>)" > 
									</TD>
									<TD nowrap align=center valign="top">	
										<!--事故扣款-->							
										<input type="text" style="width:100%"  name="SSmoeny"  class="inputbox8r" value="<%=tmpRec(CurrentPage,  CurrentRow,  9)%>" onchange="SUKMCHG(<%=currentRow-1%>)"  > 
										<input TYPE="HIDDEN" name="BSSmoeny"  class="readonly8s" value="<%=tmpRec(CurrentPage,  CurrentRow,  9)%>" >
									</TD>	
																				
									<TD  align=right valign="top">
										<!--績效獎金-->								
										<input type="text" style="width:100%" name="TOTJX" class="inputbox8r" value="<%=round(tmpRec(CurrentPage,CurrentRow,19),0)%>" style='color:<%=ft_clr%>'>
										<br><font color="#cccccc"><%=round(tmpRec(CurrentPage,CurrentRow,34),0)%></font>
										<input type="hidden" name="workJs" class="inputbox8r" size=5 value="<%=tmpRec(CurrentPage,CurrentRow,21)%>" >
										<input type="hidden" name="BTOTJX"  value="<%=tmpRec(CurrentPage,CurrentRow, 19)%>" >
										<input type="hidden" name="realJXM" class="inputbox8r" size=10 value="<%=tmpRec(CurrentPage,CurrentRow,22)%>" >
									</TD>		
									<TD align=center valign="top">
										<input type="text" name="dm"  class="readonly8" value="<%=tmpRec(CurrentPage,CurrentRow,29)%>" style='width:100%;text-align:right' ONCHANGE=SUKMCHG(<%=currentRow-1%>)>
									</TD> 		
									
									<TD align=center valign="top">
										<%
										if trim(tmpRec(CurrentPage,CurrentRow,29))="USD" then 
											usd_jxm=round(cdbl(tmpRec(CurrentPage,CurrentRow,19))/cdbl(tmpRec(CurrentPage,CurrentRow,30)),0)
										else
											usd_jxm=0
										end if 	
										%>
										<input type="text" name="jxmUSD"  class="readonly8s" value="<%=usd_jxm%>" style='width:100%;text-align:right'  > 
									</TD> 				
								</TR>
							<%else%>
								<input type="hidden" name="empid"  >
								<%for  zz  =  1  to  5  
									zzstr=chr(64+zz)
								 %>
									<input type="hidden" name="JX"&<%=zzstr%>>
									<input type="hidden" name="BJX"&<%=zzstr%>>
								<%next%> 
								<input type="hidden" name="FL" >
								<input type="hidden" name="FLmoney" >
								<input type="hidden" name="FQD" >
								<input type="hidden" name="gnn" >
								<input type="hidden" name="SSmoeny" >
								<input type="hidden" name="BSSmoeny">
								<input type="hidden" name="TOTJX" >
								<input type="hidden" name="BTOTJX" >	
								<input type="hidden" name="workJs" >	
								<input type="hidden" name="realJXM" >	
								<input type="hidden" name="khfen"  > <!--考核分數-->
								<input type="hidden" name="fn2"  >
								<input type="hidden" name="jxyn"  >
								<input type="hidden" name="newjs"  >
								<input type="hidden" name="jobjs"  >
								<input type="hidden" name="ct" value="">
								<input type="hidden" name="hrjs"  >
								<input type="hidden" name="relhr"  >
								<input type="hidden" name="nedhr"  >
						<%end if%>
						<%next%> 
				</TABLE>
			</td>
		</tr>
		<tr>
			<td align="center">
				<table  class="txt" cellpadding=3 cellspacing=3>
					<tr class=txt9>
						<td align=left>
						<% If CurrentPage > 1 Then %>
							<input type="submit"name="send"  value="FIRST" class="btn btn-sm btn-outline-secondary">
							<input type="submit"  name="send" value="BACK" class="btn btn-sm btn-outline-secondary">
						<% Else  %>
							<input  type="submit" name="send" value="FIRST"  disabled  class="btn btn-sm btn-outline-secondary">
							<input  type="submit"  name="send" value="BACK" disabled class="btn btn-sm btn-outline-secondary">
						<% End If %>
						<%  If  cint(CurrentPage)  <  cint(TotalPage)  Then %>
							<input type="submit" name="send" value="NEXT" class="btn btn-sm btn-outline-secondary">
							<input type="submit" name="send" value="END" class="btn btn-sm btn-outline-secondary">
						<% Else %>
							<input type="submit" name="send" value="NEXT" disabled class="btn btn-sm btn-outline-secondary">
							<input type="submit" name="send" value="END"  disabled  class="btn btn-sm btn-outline-secondary">
						<%  End  If  %>
						</TD>
						<td align=center>共<%=RecordInDB%>筆, 第<%=CurrentPage%>頁/共<%=TotalPage%>頁</td>
						<td  align=riht>
							<input type=button  name=btm  class="btn btn-sm btn-danger" value="(Y)Confirm" onclick="go()">
							<input Type=reset name=btM class="btn btn-sm btn-outline-secondary" value="(N)Cancel" onclick="goclr()">
							<%if session("netuser")<>"LSARY" then %>
							<input Type=button name=btM class="btn btn-sm btn-outline-secondary" value="Export To Excel" onclick="goexcel()">	
							<%end if%>
						</td>
					</tr>
				</table>
				<input type="hidden" name="empid" >
				<%for yy = 1to ICOUNT %>
				<input type="hidden" name="JX"&<%=arrays(yy,5)%>>
				<input type="hidden" name="BJX"&<%=arrays(yy,5)%>>
				<%next%>
				<input type="hidden" name="SUMJX" >
				<input type="hidden" name="FL" >
				<input type="hidden" name="FLmoney" >
				<input type="hidden" name="FQD" >
				<input type="hidden" name="gnn" >
				<input type="hidden" name="SSmoeny" >
				<input type="hidden" name="BSSmoeny" >
				<input type="hidden" name="TOTJX" >
				<input type="hidden" name="BTOTJX" >
				<input type="hidden" name="workJs" >		
				<input type="hidden" name="realJXM" >	
				<input type="hidden" name="khfen"  >
				<input type="hidden" name="fn2"  >
				<input type="hidden" name="jxyn"  >
			</td>
		</tr>
	</table>
			
</form>
</body>
</html>
<!-- #include file="../Include/func.inc" -->
<!-- #include file="../Include/SIDEINFO.inc" -->
<script language=vbscript >

function goexcel()
	'open "<%=self%>.toexcel.asp" , "Back" 
	parent.best.cols="100%,0%"
	<%=self%>.action="<%=self%>.toexcel.asp"
	<%=self%>.target="Back"
	<%=self%>.submit()	
end function  

function oepnEmpWKT(index)  
	empidstr  =<%=self%>.empid(index).value 
	yymmstr  = <%=self%>.JXym.value 	 
	open "../ZZZ/getEmpWorkTime.asp?yymm="& yymmstr & "&empid=" & empidstr , "_blank","top=10 , left=10, scrollbars=yes"
end function 

function dchg()
	<%=self%>.totalpage.value="0"
	<%=self%>.action="<%=self%>.foregnd.asp"
	<%=self%>.submit()
end function  

function fn2chg(index)
	if <%=self%>.fn2(index).checked=false then 
		<%=self%>.jxyn(index).value="N" 
		<%=self%>.totjx(index).value = 0 
	else
		<%=self%>.jxyn(index).value="Y"
		<%=self%>.totjx(index).value = <%=self%>.BTOTJX(index).value  
	end if 
end function 

function njschg()
	if <%=self%>.njs.value<>"" then 		
		if isnumeric(<%=self%>.njs.value)=false then 
			alert "請輸入數值!!"
			<%=self%>.njs.value=""
			<%=self%>.njs.focus()
			exit function 
		else
			for x =1 to <%=self%>.PageRec.value 
				<%=self%>.newjs(x-1).value = <%=self%>.njs.value 
				clcjx(x-1)
			next
		end if
	end if 		
end function 

function chkdata(sid,index) 
	'修正檢核及計算方式   20110429 elin 
	codestr = Document.Forms("<%=self%>").Elements(sid)(index).value
	bid = "B"&sid
	if isnumeric(codestr)=false then 
		alert "請輸入數字!!xin danh lai so"
		if sid<>"FQD" then 
			Document.Forms("<%=self%>").Elements(sid)(index).value=Document.Forms("<%=self%>").Elements(bid)(index).value	
		end if 	
		Document.Forms("<%=self%>").Elements(sid)(index).select()
		'exit function 
	end if	
	clcjx(index) 
end function 

function clcjx(index)

	'曠職ㄧ天以上...扣績效獎金一半, 曠職2天以上...無績效獎金 
	'事假3天以上未滿5天 績效減半, 事假5天含以上 績效為 0  (VN)				 	
	'勸導書1張扣獎金一半,2張含以上績效 獎金 = 0 
	
	f_qnn = <%=self%>.gnn(index).value  '勸導書
	f_FL = <%=self%>.fL(index).value   '曠職
	f_khfen =<%=self%>.khfen(index).value '考核分數
	f_jiaA =<%=self%>.jiaAhr(index).value '事假
			
	jobJs=<%=self%>.jobjs(index).value
	a1=<%=self%>.jxa(index).value
	a2=<%=self%>.jxb(index).value
	a3=<%=self%>.jxc(index).value
	a4=<%=self%>.jxd(index).value
	a5=<%=self%>.jxe(index).value
	njs=<%=self%>.newJS(index).value
	hrjs=<%=self%>.hrjs(index).value
	if njs="" then njs=1  
	'alert (cdbl(a1)+cdbl(a2)+cdbl(a3)+cdbl(a4)+cdbl(a5))
	N_sumjx = (cdbl(a1)+cdbl(a2)+cdbl(a3)+cdbl(a4)+cdbl(a5))*cdbl(jobJs)*cdbl(njs)*cdbl(hrjs) 
	
	tot_ykm = 0 
	if  f_qnn<>"" and f_qnn="1" then tot_ykm = tot_ykm+cdbl(N_sumjx)*0.5  
	if  f_qnn<>"" and f_qnn>="2" then tot_ykm = tot_ykm+cdbl(N_sumjx)
	if  f_FL<>"" then 		
		if cdbl(f_FL)>=16  then 
			tot_ykm = tot_ykm+cdbl(N_sumjx) 
			'alert tot_ykm
		elseif cdbl(f_FL)>=8 and  cdbl(f_FL)<16 then 
			tot_ykm = tot_ykm+cdbl(N_sumjx)*0.5  
			'alert tot_ykm
		end if 
	end if 		
	f_ct=<%=self%>.ct(index).value 
	if f_ct="VN" then 
		if f_khfen<>"" and f_khfen<"70" then tot_ykm = tot_ykm+cdbl(N_sumjx)*0.5    	
		if  f_jiaA<>"" and ( cdbl(f_jiaA)>=24 and  cdbl(f_jiaA)<40 )  then tot_ykm = tot_ykm+cdbl(N_sumjx)*0.5
		if  f_jiaA<>"" and cdbl(f_jiaA)>=40 then tot_ykm = tot_ykm+cdbl(N_sumjx)
	else
		tot_ykm = tot_ykm + (cdbl(N_sumjx)/30) * cdbl(f_jiaA/8)   '外國人依請假天數扣績效獎金
	end if 	
	
	<%=self%>.FLmoney(index).value = round(tot_ykm,0)
	<%=self%>.sumjx(index).value=round(n_sumjx,0)
	
	ykm = <%=self%>.FLmoney(index).value  '應扣款 
	n_fqd=<%=self%>.fqd(index).value  '反規定
	n_SSmoeny=<%=self%>.SSmoeny(index).value '事故扣款 
	n_TOTJX= cdbl(N_sumjx) - cdbl(ykm) - ( cdbl(n_fqd)*10000 ) - cdbl(n_SSmoeny) 
	if cdbl(n_TOTJX)<0 then n_TOTJX = 0 
	<%=self%>.totjx(index).value = round( n_TOTJX , 0) 
end function 

function goclr()
	open "<%=self%>.asp", "_self"
end function 

function  datachg(index)		 	'修正檢核及計算方式   20110429 elin  nouse  , 改用 chkdata 
	thiscols  =document.activeElement.name 
	Bcols="B"&document.activeElement.name
	if isnumeric(document.all(thiscols)(index).value)=false then
		alert  "請輸入數值"
		document.all(thiscols)(index).focus()
		document.all(thiscols)(index).value=document.all(Bcols)(index).value
		document.all(thiscols)(index).select()
	 	'exit function 		 
	else	
		alert cdbl(document.all(thiscols)(index).value) 
		NewJXM=cdbl(document.all(Bcols)(index).value) -  cdbl(document.all(thiscols)(index).value)		
		alert NewJXM 
		'NewTOTJX=cdbl(<%=self%>.BTOTJX(index).value)-cdbl(<%=self%>.TOTJX(index).value)		<%=self%>.sumjx(index).value=cdbl(<%=self%>.sumjx(index).value) - (cdbl(NewJXM)) 
		<%=self%>.TOTJX(index).value=cdbl(<%=self%>.TOTJX(index).value) - (cdbl(NewJXM)) - (+ewTOTJX)
	end if
end function

function SUKMCHG(INDEX)
	IF isnumeric(<%=SELF%>.SSMOENY(INDEX).VALUE)=false THEN
		ALERT "請輸入數值!!"
		<%=SELF%>.SSMOENY(INDEX).VALUE=<%=SELF%>.bSSMOENY(INDEX).VALUE
		<%=SELF%>.SSMOENY(INDEX).FOCUS()		
	ELSE
		if cdbl(<%=SELF%>.SSMOENY(INDEX).VALUE)<0 then 
			alert "不可 < 0 "
			<%=SELF%>.SSMOENY(INDEX).VALUE=<%=SELF%>.bSSMOENY(INDEX).VALUE
			<%=SELF%>.SSMOENY(INDEX).FOCUS()		
		end if 
		'<%=self%>.TOTJX(index).value=cdbl(<%=self%>.TOTJX(index).value) + (cdbl(NEWSUKM)) + (CDBL(NewTOTJX))
		if isnumeric(<%=self%>.fqd(index).value)=true then 
			fdqCnt = cdbl(<%=self%>.fqd(index).value) 
		end if  
		ykm =  cdbl(<%=SELF%>.FLmoney(INDEX).VALUE) 
		sukm = cdbl(<%=SELF%>.SSMOENY(INDEX).VALUE)  
		sumjx = cdbl(<%=SELF%>.sumjx(INDEX).VALUE)  		
		<%=self%>.TOTJX(index).value =  cdbl(sumjx)-cdbl(sukm)-cdbl(ykm)-cdbl(fdqCnt*10000)
		
	END IF
END function

function  fqdchg(index)  
	'200803反規定ㄧ次扣10000VND 
	'200901 當月勸導書1張績效減半 , 勸導書2張績效 = 0  , 
	if <%=self%>.fqd(index).value<>"" then 
		if isnumeric(<%=self%>.fqd(index).value)=false then 
			alert "請輸入數字!!"
			<%=self%>.fqd(index).value="0"
			<%=self%>.fqd(index).select()
			exit function 
		else
			fdqCnt =cdbl(<%=self%>.fqd(index).value)
			ykm =  cdbl(<%=SELF%>.FLmoney(INDEX).VALUE) 
			sukm = cdbl(<%=SELF%>.SSMOENY(INDEX).VALUE)  
			sumjx = cdbl(<%=SELF%>.sumjx(INDEX).VALUE)  
			<%=self%>.TOTJX(index).value =  cdbl(sumjx)-cdbl(sukm)-cdbl(ykm)-cdbl(fdqCnt*10000)
		end if 
	end if	 
	
	'old 適用至2008/02
	'if  cdbl(<%=self%>.fqd(index).value)=1  then
	'	<%=self%>.TOTJX(index).value=cdbl(<%=self%>.TOTJX(index).value)/2
	'elseif cdbl(<%=self%>.fqd(index).value)=2 then
	'	<%=self%>.TOTJX(index).value=0
	'else
 	'	<%=self%>.TOTJX(index).value = <%=self%>.TOTJX(index).value 
	'end if
	
end function

function TT(a)  
	alert <%=self%>.JXSTT(a).value
end function

FUNCTION GO()  	
	<%=self%>.action ="<%=self%>.upd.asp"
	<%=SELF%>.SUBMIT
END FUNCTION

</script> 
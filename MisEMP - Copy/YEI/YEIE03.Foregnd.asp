<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<!-- #include file = "../GetSQLServerConnection.fun" -->
<!-- #include file="../ADOINC.inc" -->
<!--#include file="../include/sideinfo.inc"-->
<%
session.codepage="65001"
SELF = "YEIE03"

Set conn = GetSQLServerConnection()
Set rs = Server.CreateObject("ADODB.Recordset")
Set rds = Server.CreateObject("ADODB.Recordset")

nowmonth = year(date())&right("00"&month(date()),2)
if month(date())="01" then
	calcmonth = year(date()-1)&"12"
else
	calcmonth =  year(date())&right("00"&month(date())-1,2)
end if 

F_whsno = request("F_whsno")
F_groupid = request("F_groupid")
F_zuno = request("F_zuno") 
F_shift=request("F_shift")
F_empid =request("empid")
F_country=request("F_country")
fclass = request("fclass")  
sortvalue = request("sortvalue") 
if sortvalue ="" then sortvalue="b.country , a.khw, a.khg, a.empid"

FUNCTION FDT(D)
	Response.Write YEAR(D)&"/"&RIGHT("00"&MONTH(D),2)&"/"&RIGHT("00"&DAY(D),2)
END FUNCTION 
khym = request("khym")
if request("khym")="" then 
	khym=nowmonth
end if  

 
 '一個月有幾天 
cDatestr=CDate(LEFT(khym,4)&"/"&RIGHT(khym,2)&"/01") 
days = DAY(cDatestr+(32-DAY(cDatestr))-DAY(cDatestr+(32-DAY(cDatestr))))   '一個月有幾天  
'本月最後一天 
ENDdat = LEFT(khym,4)&"/"&RIGHT(khym,2)&"/"&DAYS   

'if khweek="" then khweek=(days\7)  

gTotalPage = 1
PageRec = 10    'number of records per page
TableRec = 20    'number of fields per record   


sql="select  "&_
		"(ov_h1m+ov_h2m+ov_h3m+ov_b3m ) as totov ,     "&_
		"round( (ov_h1m+ov_h2m+ov_h3m+ov_b3m )/ (money_h*2.5),3) ff1  "&_
		", *  "&_
		"from (  "&_
		"			select g.empid,   b.empnam_cn, b.empnam_vn,  "&_
		"			g.yymm, g.money_h,  "&_
		"			isnull(g.h1m,0) h1m,  isnull(g.h2m,0) h2m,  isnull(g.h3m,0) h3m,  isnull(g.b3m,0) b3m  ,  "&_
		"			round(isnull(jb.h1,g.h1)*isnull(g.money_h,0)*1.5,0)  as N_h1m ,  "&_
		"			isnull(g.h1m,0)- round(isnull(jb.h1,g.h1)*isnull(g.money_h,0)*1.5,0)   as ov_h1m, "&_
		"			round(isnull(jb.h2,g.h2)*isnull(g.money_h,0)*2,0)  as N_h2m ,  "&_
		"			isnull(g.h2m,0) - round(isnull(jb.h2,g.h2)*isnull(g.money_h,0)*2,0)  as ov_h2m,  "&_
		"			round(isnull(jb.h3,g.h3)*isnull(g.money_h,0)*3,0)  as N_h3m ,  "&_
		"			isnull(g.h3m,0)- round(isnull(jb.h3,g.h3)*isnull(g.money_h,0)*3,0)  as ov_h3m, "&_
		"			round((isnull(g.b3,0)-isnull(jb.ov_b3,0))*isnull(g.money_h,0)*0.3,0) as N_b3m ,  "&_
		"			isnull(g.b3m,0)-round((isnull(g.b3,0)-isnull(jb.ov_b3,0))*isnull(g.money_h,0)*0.3,0) as ov_b3m  "&_
		"			, lw,lg,lz,lgstr, lzstr, lwstr , ls, b.nindat , b.outdate "&_
		"			 from    "&_
		"			( select * from empdsalary_bak  where country='"&F_country&"' and whsno='"&f_whsno&"'  and yymm ='"& khym &"'   )  g "&_
		"			left join  ( select * from view_empfile  ) b on b.empid = g.empid "&_
		"			left join ( select* from view_empgroup   ) i on i.empid = g.empid  and i.yymm = g.yymm "&_
		"			left join ( select * from empJBtim   ) jb on jb.yymm = g.yymm and jb.empid = g.empid "&_
		"			where i.lg like '"&f_groupid&"%' "&_
		") z 	"&_
		"order by  lg, empid " 	 
'response.write sql 	
'response.end 			
if request("TotalPage") = "" or request("TotalPage") = "0" then 
	CurrentPage = 1	
	rds.Open SQL, conn, 1, 3 
	IF NOT RdS.EOF THEN 	
		PageRec = rds.RecordCount 
		rds.PageSize = PageRec 
		RecordInDB = rds.RecordCount 
		TotalPage = rds.PageCount  
		gTotalPage = TotalPage
	END IF 	 
	Redim tmpRec(gTotalPage, PageRec, TableRec)   'Array
	
	for i = 1 to TotalPage 
		for j = 1 to PageRec
			if not rds.EOF then 			
				tmpRec(i, j, 0) = "no"
				tmpRec(i, j, 1) = rds("EMPID")
				tmpRec(i, j, 2) = rds("lw")
				tmpRec(i, j, 3) = rds("lg")
				tmpRec(i, j, 4) = rds("lz")
				tmpRec(i, j, 5) = rds("ls")
				tmpRec(i, j, 6) = rds("empnam_cn")
				tmpRec(i, j, 7) = rds("empnam_vn")
				tmpRec(i, j, 8) = rds("nindat")
				tmpRec(i, j, 9) = rds("outdate")
				tmpRec(i, j, 10) = rds("lgstr")
				tmpRec(i, j, 11) = rds("lzstr")
				tmpRec(i, j, 12) = rds("money_h")
				tmpRec(i, j, 13) = right("00.000"&replace(formatnumber(rds("ff1"),3),",",""),6)
				tmpRec(i, j, 14) = rds("totov")		 
				if 	isnull(rds("totov")) or rds("totov")="" then tmpRec(i, j, 14)= 0 
				'response.write tmpRec(i, j, 13) &"-"& tmpRec(i, j, 14) & "-" & len(replace(tmpRec(i, j, 13),",","")) &"<br>"				
				rds.MoveNext 
			else 
				exit for 
			end if 
	 	next 
	
	 	if rds.EOF then 
			rds.Close 
			Set rds = nothing
			exit for 
	 	end if 
	next
end if	
%>

<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
<link rel="stylesheet" href="../Include/style.css" type="text/css">
<link rel="stylesheet" href="../Include/style2.css" type="text/css"> 
</head>
<body  topmargin="0" leftmargin="5"  marginwidth="0" marginheight="0"  onkeydown="enterto()" onload="f()">
<form  name="<%=self%>" method="post" action="<%=self%>.upd.asp"   >
<INPUT TYPE=HIDDEN NAME="UID" VALUE=<%=SESSION("NETUSER")%>>
<INPUT TYPE=hidden NAME=TotalPage VALUE="<%=TotalPage%>">
<INPUT TYPE=hidden NAME=CurrentPage VALUE="<%=CurrentPage%>">
<INPUT TYPE=hidden NAME=RecordInDB VALUE="<%=RecordInDB%>">
<INPUT TYPE=hidden NAME=PageRec VALUE="<%=PageRec%>">
<INPUT TYPE=hidden NAME=gTotalPage VALUE="<%=gTotalPage%>"> 	
<INPUT TYPE=hidden NAME=nowmonth VALUE="<%=nowmonth%>">
<INPUT TYPE=hidden NAME=days VALUE="<%=days%>">
<INPUT TYPE=hidden NAME=sortvalue VALUE="<%=sortvalue%>">
<input name=act value="<%=act%>" type=hidden >

	<table border=0 style="height:70px"><tr><td>&nbsp;</td></tr></table>
	
	<table width="100%" BORDER=0 align=center cellpadding=0 cellspacing=0 >
		<tr>
			<td>
				<table class="txt" cellpadding=3 cellspacing=3>
					<TR>		
						<TD align=right>考核年月</TD>
						<td>
							<input type="text" style="width:100px"  name=khym value="<%=khym%>" maxlength=6  onchange=datachg2()>
							<input name=khweek type=hidden>				
							<input type="hidden" name="fclass" value="A" > 
						</td>
						<td align=right>國籍<BR>Quoc tich</td>
						<td>
							<select name=F_country  onchange="datachg()" style="width:120px">										
							<%							
								SQL="SELECT * FROM BASICCODE WHERE FUNC='country'  and sys_type='VN' ORDER BY SYS_TYPE desc "
							
							SET RST = CONN.EXECUTE(SQL)
							WHILE NOT RST.EOF
							%>
							<option value="<%=RST("SYS_TYPE")%>" <%if RST("SYS_TYPE")=F_country then%>selected<%end if%>><%=RST("SYS_VALUE")%></option>
							<%
							RST.MOVENEXT
							WEND
							%>
							</SELECT>
							<%SET RST=NOTHING %>	
						</td>
					</tr>
					<tr  height=22>	
						<TD align=right>廠別<br>Xuong</TD>
						<td>
							<select name=F_whsno style="width:120px">					
									<%
									if session("rights")=0 then 
										SQL="SELECT * FROM BASICCODE WHERE FUNC='WHSNO' ORDER BY SYS_TYPE "%>
										<option value="" selected >全部(Toan bo) </option>
									<%		
									else
										SQL="SELECT * FROM BASICCODE WHERE FUNC='WHSNO' and sys_type='"& session("NETWHSNO") &"' ORDER BY SYS_TYPE "
									end if	
									SET RST = CONN.EXECUTE(SQL)
									WHILE NOT RST.EOF
									%>
									<option value="<%=RST("SYS_TYPE")%>" <%if RST("SYS_TYPE")=F_whsno then%>selected<%end if%>><%=RST("SYS_VALUE")%></option>
									<%
									RST.MOVENEXT
									WEND
									%>
								</SELECT>
							<%SET RST=NOTHING %>	
						</td>									
						<TD align=right>部門<br>Bo Phan</TD>
						<td>
							<table border=0 class="txt">
								<tr>
									<td>
										<select name=F_groupid   onchange="datachg()" style="width:120px">			
											<option value="">全部(Toan bo) </option>
											<% 
											SQL="SELECT * FROM BASICCODE WHERE FUNC='GROUPID' and sys_type <>'AAA' ORDER BY SYS_TYPE " 
											SET RST = CONN.EXECUTE(SQL)
											RESPONSE.WRITE SQL 
											WHILE NOT RST.EOF  
											%>
												<option value="<%=RST("SYS_TYPE")%>" <%if RST("SYS_TYPE")=F_groupid then%>selected<%end if%>><%=RST("SYS_TYPE")%><%=RST("SYS_VALUE")%></option>				 
											<%
												RST.MOVENEXT
												WEND 
											%>
										</SELECT>
										<%SET RST=NOTHING %>
									</td>
									<td>
										<select name=F_zuno   onchange="datachg()" style="width:120px">				
											<option value=""></option>
											<%
												SQL="SELECT * FROM BASICCODE WHERE FUNC='zuno' and sys_type <>'XX' and  left(sys_type,4)= '"& F_groupid &"' ORDER BY SYS_TYPE "
												SET RST = CONN.EXECUTE(SQL)
												RESPONSE.WRITE SQL 
												WHILE NOT RST.EOF  
											%>
												<option value="<%=RST("SYS_TYPE")%>" <%if RST("SYS_TYPE")=F_zuno then%>selected<%end if%>><%=right(RST("SYS_TYPE"),1)%>-<%=RST("SYS_VALUE")%></option>				 
											<%
												RST.MOVENEXT
												WEND 
											%>
										</SELECT>
										<%SET RST=NOTHING %>
									</td>
								</tr>
							</table>
						</td>
						<td align=right>班別<BR>Ca</td>
						<td>
							<select name=F_shift   onchange="datachg()"  style="width:120px">
								<option value=""></option>
								<option value="ALL" <%if F_shift="ALL" then%>selected<%end if%>>日</option>
								<option value="A" <%if F_shift="A" then%>selected<%end if%>>A班</option>
								<option value="B" <%if F_shift="B" then%>selected<%end if%>>B班</option>
							</select>										
						</td>
						<td><input type=button name=send value="(S)查詢" class="btn btn-sm btn-outline-secondary" onclick="datachg()"></td>
					</TR>	 
				</table>
			</td>
		</tr>
		<tr>
			<td align="center">
				<table id="myTableGrid" width="98%">
					<tr class="header">
						<Td nowrap align=center  >STT</td>
						<Td  nowrap align=center  >部門</td>
						<Td  nowrap align=center   style='cursor:pointer;' onclick=sortby(1) title='依單位排序' >單位<br><img src="../picture/soryby.gif"></td>
						<Td   nowrap align=center  style='cursor:pointer;' onclick=sortby(5) title='依班別排序'>班別<br><img src="../picture/soryby.gif"></td>
						<Td   nowrap align=center  style='cursor:pointer;' onclick=sortby(2) title='依工號排序'>工號<br><img src="../picture/soryby.gif"></td>
						<Td   nowrap align=center  >姓名</td>									
						<Td   nowrap align=center  >時薪</td>
						<td align=center>A</td>
						<td align=center>B</td>
						<td align=center>C</td>
						<td align=center>D</td>
						<td align=center>E</td>
						<td  align=center style='cursor:pointer;' onclick=sortby(4)  nowrap title='依總分排序'>總分<br><img src="../picture/soryby.gif"></td>
						<td  align=center style='cursor:pointer;' onclick=sortby(4)  nowrap title='依總分排序'>考核獎金<br><img src="../picture/soryby.gif"></td>
						
					</tr> 	
					<%
					totm  = 0 
					for CurrentRow = 1 to PageRec
						IF CurrentRow MOD 2 = 0 THEN 
							WKCOLOR="LavenderBlush"
						ELSE
							WKCOLOR="#DFEFFF"
						END IF 	 		
						
						FA=left(tmpRec(CurrentPage, CurrentRow, 13),1)
						FB=mid(tmpRec(CurrentPage, CurrentRow, 13),2,1)
						FC=mid(tmpRec(CurrentPage, CurrentRow, 13),4,1)
						FD=mid(tmpRec(CurrentPage, CurrentRow, 13),5,1)
						FE=mid(tmpRec(CurrentPage, CurrentRow, 13),6,1)
						if fa="" then fa=0
						if fb="" then fb=0
						if fc="" then fc=0 
						if fd="" then fd=0 
						if fe="" then fe=0 
						totf =cdbl(fa)*10 +cdbl(fb)*1+cdbl(fc)*0.1+cdbl(fd)*0.01+cdbl(fe)*0.001 		
					  
						totm = totm + cdbl(tmpRec(CurrentPage, CurrentRow, 14))
						if cdbl(tmpRec(CurrentPage, CurrentRow, 14))< 0 then 
							f_TOTM  = 0 
						else
							f_totm = cdbl(tmpRec(CurrentPage, CurrentRow, 14))
						end if 
					%>	
					<TR BGCOLOR="<%=WKCOLOR%>" class="txt9"> 	 
						<Td align=center><%=CurrentRow%></td>
						<Td align=center><%=tmpRec(CurrentPage, CurrentRow, 10)%></td>
						<td align=left><%=tmpRec(CurrentPage, CurrentRow, 11)%></td>
						<Td align=center><%=tmpRec(CurrentPage, CurrentRow, 5)%></td>
						<Td align=center >
							<%=tmpRec(CurrentPage, CurrentRow, 1)%>
							<input type=hidden name=F_empid  value="<%=tmpRec(CurrentPage, CurrentRow, 1)%>">
						</td>
						<Td  style="cursor:pointer" nowrap onclick="vbscript:oepnEmpWKT(<%=currentRow-1%>)" title="**點選可看出勤紀錄**">
							<%=tmpRec(CurrentPage, CurrentRow, 6)%><br><%=left(tmpRec(CurrentPage, CurrentRow, 7),15)%>
						</td>									
						<td align="right">
								<%=formatnumber(tmpRec(CurrentPage, CurrentRow, 12),0)%>
								<input name=moneyh class=inputbox8 size=4 type=hidden value="<%=formatnumber(tmpRec(CurrentPage, CurrentRow, 12),0)%>">
						</td><!--時薪-->
						<td align="right"><input type="text" name="FA" class=inputbox8r  value="<%=fa%>" onblur="fenchg(1,<%=CurrentRow-1%>)" style="width:100%"></td>
						<td align="right"><input type="text" name="FB" class=inputbox8r  value="<%=fb%>" onblur="fenchg(2,<%=CurrentRow-1%>)" style="width:100%"></td>
						<td align="right"><input type="text" name="FC" class=inputbox8r  value="<%=fc%>" onblur="fenchg(3,<%=CurrentRow-1%>)" style="width:100%"></td>
						<td align="right"><input type="text" name="FD" class=inputbox8r  value="<%=fd%>" onblur="fenchg(4,<%=CurrentRow-1%>)" style="width:100%"></td>
						<td align="right"><input type="text" name="FE" class=inputbox8r  value="<%=fe%>" onblur="fenchg(5,<%=CurrentRow-1%>)" style="width:100%"></td>
						<td align="right"><input type="text" name="TOTf" class=inputbox8r  value="<%=formatnumber(totf,3)%>" readonly style="width:100%"></td>
						<td align="right"><input type="text" name="TOTM" class=inputbox8r value="<%=formatnumber(f_totm,0)%>" readonly  style="width:100%"></td>		
					</tr>				
					<%next%> 
					<tr bgcolor="lightyellow">
						<td colspan=13></td>
						<td align="right" class="txt8"><%=formatnumber(totm,0)%></td>
					</tr>
					<input type=hidden name=empid  value="">
					<input name=moneyh class=inputbox8 size=4 type=hidden value=0>
					<input name=fa class=inputbox8 size=4 type=hidden value=0>
					<input name=fb class=inputbox8 size=4 type=hidden value=0>
					<input name=fc class=inputbox8 size=4 type=hidden value=0>
					<input name=fd class=inputbox8 size=4 type=hidden value=0>
					<input name=fe class=inputbox8 size=4 type=hidden value="0">
					<input name=TOTf class=inputbox8 size=4 type=hidden value="0">
					<input name=TOTM class=inputbox8 size=4 type=hidden value="0">
				</table>
			</td>
		</tr>
		<tr>
			<td align="center">
				<table class="table-borderless table-sm text-secondary txt">
					<tr ALIGN=center>
						<TD >
							<input type=button  name=send value="(M)回主畫面" class="btn btn-sm btn-outline-secondary" onclick=backM()>
							<input type=button  name=send value="下載到Excel" class="btn btn-sm btn-outline-secondary" onclick=goexcel()>
						</TD>
					</TR>
				</TABLE>
			</td>
		</tr>
	</table>
			
</form>


</body>
</html>
<script language=vbscript> 
function oepnEmpWKT(index)
	empidstr = <%=self%>.F_empid(index).value
	yymmstr = <%=self%>.khym.value
	khweekstr = "" '<%=self%>.khweek.value
	open "../yed/yedq01.Foregnd.asp?fr=A&yymm="& yymmstr & "&empid=" & empidstr &"&khweek=" & khweekstr , "_blank", "top=10 , left=10, height=500, width=700,scrollbars=yes"
end function


function fenchg(a,index)
	f_fa = <%=self%>.fa(index).value 
	f_fb = <%=self%>.fb(index).value 
	f_fc = <%=self%>.fc(index).value 
	f_fd = <%=self%>.fd(index).value 
	f_fe = <%=self%>.fe(index).value 
	totfen = eval(f_fa*10+f_fb*1+f_fc*0.1+f_fd*0.01+f_fe*0.001)
	<%=self%>.TOTf(index).value= formatnumber(totfen,3)
	f_moneyh=cdbl(<%=self%>.moneyh(index).value)
	f_totm=eval(totfen*f_moneyh*2.5)
	<%=self%>.TOTM(index).value= formatnumber(f_totm,0)
end function 

function fnAcng(index) 	
	maxfensu = 25
	if trim(<%=self%>.fensuA(index).value)<>"" then 
		if isnumeric(<%=self%>.fensuA(index).value)=false then  
			alert "請輸入數字!!hay nhap so vao!!"
			<%=self%>.fensuA(index).value="0"
			<%=self%>.fensuA(index).select()
			exit function  
		else 
			if cdbl(<%=self%>.fensuA(index).value)> cdbl(maxfensu) then 
				alert "分數超過 [25] 分"
				<%=self%>.fensuA(index).value="0"
				<%=self%>.fensuA(index).select()
				exit function  
			end if 	 
			calctfn(index)
		end if	
	end if 		
end function  

function goexcel()
	<%=self%>.action = "<%=self%>.toexcel.asp"
	<%=self%>.target="Back"
	'parent.best.cols="50%,50%"
	<%=self%>.submit()
	
end function 
 
function fnBcng(index) 	  
	maxfensu = 25
	if trim(<%=self%>.fensuB(index).value)<>"" then 
		if isnumeric(<%=self%>.fensuB(index).value)=false then  
			alert "請輸入數字!!hay nhap so vao!!"
			<%=self%>.fensuB(index).value=""
			<%=self%>.fensuB(index).focus()
			exit function 	
		else
			if cdbl(<%=self%>.fensuB(index).value)> cdbl(maxfensu) then 
				alert "分數超過 [25] 分"
				<%=self%>.fensuB(index).value="0"
				<%=self%>.fensuB(index).select()
				exit function  
			end if 						
			calctfn(index)
		end if	
	end if 	
end function 
 
function fnCcng(index) 	  
	maxfensu = 25
	if trim(<%=self%>.fensuC(index).value)<>"" then 
		if isnumeric(<%=self%>.fensuC(index).value)=false then  
			alert "請輸入數字!!hay nhap so vao!!"
			<%=self%>.fensuC(index).value=""
			<%=self%>.fensuC(index).focus()
			exit function 	
		else
			if cdbl(<%=self%>.fensuC(index).value)> cdbl(maxfensu) then 
				alert "分數超過 [25] 分"
				<%=self%>.fensuC(index).value="0"
				<%=self%>.fensuC(index).select()
				exit function  
			end if 			
			calctfn(index)
		end if	
	end if 	
end function   

function fnDcng(index) 
	maxfensu = 25	  
	if trim(<%=self%>.fensuD(index).value)<>"" then 
		if isnumeric(<%=self%>.fensuD(index).value)=false then  
			alert "請輸入數字!!hay nhap so vao!!"
			<%=self%>.fensuD(index).value=""
			<%=self%>.fensuD(index).focus()
			exit function 	
		else
			if cdbl(<%=self%>.fensuD(index).value)> cdbl(maxfensu) then 
				alert "分數超過 [25] 分"
				<%=self%>.fensuD(index).value="0"
				<%=self%>.fensuD(index).select()
				exit function  
			end if 	
			calctfn(index)			
		end if	
	end if 	
end function  

function calctfn(index)	
	'alert index
	if <%=self%>.khweek.value="" then 
		A1 = round(cdbl(<%=self%>.days.value)\7,0)
	else
		A1=1 
	end if 
'	c_tfn = 0 	
	for A2 = 1 to  A1 		
		'alert (index\A1)*A1+A2-1
		C_fna = (<%=self%>.fensuA((index\A1)*A1+A2-1).value)
		C_fnb = (<%=self%>.fensuB((index\A1)*A1+A2-1).value)
		C_fnc = (<%=self%>.fensuC((index\A1)*A1+A2-1).value)
		C_fnd = (<%=self%>.fensuD((index\A1)*A1+A2-1).value)
		c_tfn = c_tfn + cdbl(c_fnA)+ cdbl(c_fnB)+ cdbl(c_fnC)+ cdbl(c_fnD)
	next  
	<%=self%>.tfn(index\A1).value = c_tfn
end function 
 
function  backM()	
	open "<%=self%>.asp", "_self" 	
end function 


function enterto()
	if window.event.keyCode = 13 then window.event.keyCode =9
end function

function f()
	if <%=self%>.act.value="A" then 
		<%=self%>.F_whsno.focus()
	else
		<%=self%>.khym.focus()
		<%=self%>.khym.select()
	end if 	
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

  
</script>


<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<!-- #include file = "../GetSQLServerConnection.fun" -->
<!-- #include file="../ADOINC.inc" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="refresh">
<link rel="stylesheet" href="../Include/style.css" type="text/css">
<link rel="stylesheet" href="../Include/style2.css" type="text/css">

</head>
<%

self="YECE1302"  
nowmonth = year(date())&right("00"&month(date()),2)  
if month(date())="1" then  
	calcmonth = year(date())-1&"12"  	
else
	calcmonth =  year(date())&right("00"&month(date())-1,2)   
end if 	   

if day(date())<=11 then 
	if month(date())="1" then  
		calcmonth = year(date())-1&"12" 
	else
		calcmonth =  year(date())&right("00"&month(date())-1,2)   
	end if 	 	
else
	calcmonth = nowmonth 
end if 	

Set conn = GetSQLServerConnection()
Set rs = Server.CreateObject("ADODB.Recordset")


f_years = request("f_years")
f_country = request("f_country")
f_whsno = request("f_whsno")
f_groupid = request("f_groupid")
f_empid = request("f_empid") 
f_kj = request("f_kj") 
f_indat = request("f_indat") 


gTotalPage = 1
PageRec = 0    'number of records per page
TableRec = 35    'number of fields per record 


'基本扣稅稅額設定   (  Store Procude 設定   2009 = 4,000,000VND , 2008 = 5,000,000 VND以上 , VN扣稅，外國人不扣 )
if f_years<="2008"  then 
	B_taxAmt = 5000000 
else
	B_taxAmt = 4000000 	
end if 	

enddat = f_years&"1231"


if flag="S" then   	
	sql="select a.* , isnull(b.hr,0) JAw , c.cv , sys_value job    from "&_
			"( select * from Temp_empNZJJ where years='"&f_years&"' and whsno='"&f_whsno&"' and grouid like'"&f_groupid &"%' and empid like'"&f_empid&"%'  "&_
			"and ( country ='"& f_country &"' or case when country in('TW','MA') then 'TM' else country end  = '"& f_country &"' "&_
			"or case when country ='VN' then country else 'HW' end = '"& f_country &"' ) "&_
			") a "&_
			"left join (select empid, sum(hhour) hr from empholiday where convert(char(4),dateup,112)='"&f_years&"' and jiatype='A' and isnull(place,'')<>'W'  "&_
			"group by convert(char(4),dateup,112) , empid ) b on b.empid= a.empid "&_ 
			"left join (select empid, outdat, indat, cv=(select top 1 job from bempj where empid=aa.empid oder by yymm  desc  )  from empfile ) c on c.empid= a.empid "&_	
			"left join ( select * from basiccode  where func='lev' ) zz on zz.sys_type=c.cv   "&_			
			"where isnull(outdat,'')=''  and convert(char(4),c.indat,112)<='"&f_years&"' "&_
			
			"order by a.whsno, a.years, a.country, case when a.country='VN' then a.groupid else a.country end ,  a.indate,a.empid "	
else
	if f_years<>"" and f_whsno<>"" then 
		sqlx="exec SP_calcNzjj '"& f_years &"','"& f_whsno &"' " 
		conn.execute(sqlx)
		
		sql="select a.* , convert(char(10),c.outdat,111) outdate,  isnull(b.hr,0) JAw ,  "&_
			"case when round(x.nz,2)<>round(a.nznew,2) then x.nz else a.nznew end  as  f_nz , "&_
			"case when round(x.totamamt,0)<>round(a.basicNZM,0) then x.totamamt else basicNZM end as F_bas , c.cv , sys_value job  from "&_
				"(select empid, outdat, indat , cv=(select top 1 job from bempj where empid=aa.empid oder by yymm  desc  ) from empfile aa where  ( isnull(outdat,'')='' or  convert(char(8),outdat,112)>='"& enddat &"')   "&_
				"and convert(char(4),indat,112)<='"&f_years&"'   "&_
				"and ( country = '"& f_country &"' or case when country in('TW','MA') then 'TM' else country end  = '"& f_country &"' "&_
				"or case when country ='VN' then country else 'HW' end = '"& f_country &"' ) "&_
				"and empid like'"&f_empid&"%'  "&_
				") c  "&_
				"left join  ( select * from Temp_empNZJJ where years='"&f_years&"' and whsno='"&f_whsno&"' and groupid like'"&f_groupid &"%'  "&_				
				") a  on c.empid= a.empid "&_
				"left join (select empid, sum(hhour) hr from empholiday where convert(char(4),dateup,112)='"&f_years&"' and jiatype='A' and isnull(place,'')<>'W'  "&_
				"group by convert(char(4),dateup,112) , empid ) b on b.empid= a.empid "&_	
				"left join ( select * from empnzjj where yymm='"&f_years&"' and whsno='"&f_whsno&"' and groupid like'"&f_groupid &"%' ) x on x.empid = c.empid  "&_
				"left join ( select * from basiccode  where func='lev' ) zz on zz.sys_type=c.cv   "&_
				"where isnull(a.whsno,'')='"&f_whsno&"' and convert(char(4),c.indat,112)<='"&f_years&"' "				
		if f_kj<>"" then 
			if f_kj="X" then 
				sql=sql&"and  isnull(kj,'') ='' " 
			else
				sql=sql&"and  isnull(kj,'') like '"& f_kj &"%' " 
			end if 	
		end if 
		if f_indat<>"" then 
			sql=sql&" and convert(char(10),c.indat,111)>='"& f_indat&"' "
		end if 		

		if f_indat<>"" then 
			sql=sql&"order by a.whsno, a.years, a.country,  a.indate, a.empid "			
		else
			sql=sql&"order by a.whsno, a.years, a.country, case when a.country='VN' then a.groupid else a.country end , a.indate, a.empid "			
		end if	
	else	
		sql="select * from Temp_empNZJJ where years='x' "
	end if
end if 
'response.write sqlx&"<br>"
'response.write sql&"<br>"
'response.end 
pagerec = 400 
gTotalPage = 1 
TableRec = 50 
if f_years<>""   then  
	sql = "exec sp_api_calcnzjj '"&f_years&"','"&f_whsno&"','"&f_country&"','"&f_groupid&"','"&f_empid&"','"&f_indat&"','"&f_kj&"' " 
else	
		sql="select * from Temp_empNZJJ where years='x' "
end if
'response.write sql&"<br>" 
'response.end 
p=0 
if request("TotalPage") = "" or request("TotalPage") = "0" then 
	CurrentPage = 1	
	rs.Open SQL, conn, 3, 3 
	 
	Redim tmpRec(gTotalPage, PageRec, TableRec)   'Array 
	
	for i = 1 to gTotalPage 
		for j = 1 to PageRec
			if not rs.EOF then 		
				p = p + 1 
				tmpRec(i, j, 0) = "no"
				tmpRec(i, j, 1) = trim(rs("whsno"))
				tmpRec(i, j, 2) = trim(rs("years"))
				tmpRec(i, j, 3) = trim(rs("country"))
				tmpRec(i, j, 4) = rs("empid")
				tmpRec(i, j, 5) = rs("indate")
				tmpRec(i, j, 6) = rs("empnam_cn")
				tmpRec(i, j, 7) = rs("empnam_vn")				
				tmpRec(i, j, 8) = round(rs("nz"),2) 
				'if round(rs("nz"),2) >=12 then tmpRec(i, j, 8) = round(round(rs("nz"),2)/12,1)
				tmpRec(i, j, 9) = rs("groupid")
				tmpRec(i, j, 10) = rs("gstr")
				tmpRec(i, j, 11) = rs("dm")
				tmpRec(i, j, 12) = rs("fensu")
				tmpRec(i, j, 13) = rs("kj")
				tmpRec(i, j, 14) = rs("hs")
				tmpRec(i, j, 15) = rs("days")
				tmpRec(i, j, 16) = rs("bonus") 
				if cdbl(rs("bonus")) = 0 then tmpRec(i, j, 16)=rs("f_bonus") 
				tmpRec(i, j, 17) = rs("clc_relamt")  '實發
				if cdbl(rs("realamt")) = 0 then tmpRec(i, j, 17)=rs("clc_relamt")  'ceiling 500 (扣稅後
				tmpRec(i, j, 18) = rs("tjamt")   '調整
				tmpRec(i, j, 19) = 0   '稅金 
				
				'totB=9000000   '   201306, 900萬以上扣稅 
				'Set oconn = GetSQLServerConnection()
				'if f_years<="2008" then 
				'	sql2="exec sp_calctax_2008 '"& tmpRec(i, j, 16) &"' "
				'	set ors=oconn.execute(sql2) 
				'	F_tax = ors("tax")					
				'elseif f_years>="2013" then 
				'	sql2="exec  sp_calctax_2010  '"& tmpRec(i, j, 16) &"' , 0 ,'' "
				'	set ors=oconn.execute(sql2) 
				'	F_tax = ors("tax")
				'	taxper = ors("taxper")
				'else
				'	sql2="exec  sp_calctax  '"& tmpRec(i, j, 16) &"' , '"& B_taxAmt &"' "
				'	set ors=oconn.execute(sql2) 
				'	F_tax = ors("tax")
				'	taxper = ors("taxper")
				'end if 
				'ors.close :set ors=nothing  
				'oconn.close : set oconn=nothing 
			
				'tmpRec(i, j, 19) = F_tax 
				tmpRec(i, j, 19) = rs("tax")
				
				tmpRec(i, j, 20) = rs("memos")   '調整說明
				tmpRec(i, j, 21) = cdbl(tmpRec(i, j, 16))+cdbl(tmpRec(i, j, 18))-cdbl(tmpRec(i, j, 19))
				'201801  進位改為 500→ 5000 同excel 估算 '201901 2018起不算零數
				if rs("country")="VN" then 				
					if (cdbl(tmpRec(i, j, 16))+cdbl(tmpRec(i, j, 18))-cdbl(tmpRec(i, j, 19)))  mod 1 = 0 then 
						tmpRec(i, j, 22) = (cdbl(tmpRec(i, j, 16))+cdbl(tmpRec(i, j, 18))-cdbl(tmpRec(i, j, 19))) 
					else
						tmpRec(i, j, 22) = ( (cdbl(tmpRec(i, j, 16))+cdbl(tmpRec(i, j, 18))-cdbl(tmpRec(i, j, 19)))/1 )*1
					end if 	
					'tmpRec(i, j, 22) = (cdbl(tmpRec(i, j, 16))+cdbl(tmpRec(i, j, 18))-cdbl(tmpRec(i, j, 19)))  
				else
					tmpRec(i, j, 22) = (cdbl(tmpRec(i, j, 16))+cdbl(tmpRec(i, j, 18))-cdbl(tmpRec(i, j, 19))) 
				end if 
				tmpRec(i, j, 22) =rs("rstamt") '扣稅後
				'if rs("realamt")="0" then  
				'	tmpRec(i, j, 17) = tmpRec(i, j, 22)
				'end if 
				
				tmpRec(i, j, 23) = rs("F_basic") 
				'CN境內事假 扣年獎 5% ( 天 )
				tmpRec(i, j, 24) = rs("jaw") '境內事假
				if  rs("whsno")="" then 
					tmpRec(i, j, 25) = f_whsno
				else
					tmpRec(i, j, 25) = rs("whsno") 
				end if 	
				
				tmpRec(i, j, 26) = rs("outdate")
				tmpRec(i, j, 27) = rs("grade")
				tmpRec(i, j, 28) = 0 'rs("tax")
				tmpRec(i, j, 29) =rs("nz")
				tmpRec(i, j, 30) =rs("realamt")
				tmpRec(i, j, 31) =rs("jobcode")
				tmpRec(i, j, 32) =rs("jobstr") 
				tmpRec(i, j, 33) =rs("days_cv")
				tmpRec(i, j, 34) =rs("hs_cv")
				tmpRec(i, j, 35) =rs("days_bs")
				tmpRec(i, j, 36) =rs("hs_bs")
				if rs("jobstr") >="EV5"   then 
					tmpRec(i, j, 37) = trs("days_cv")
				else
					tmpRec(i, j, 37) = rs("days_bs")
				end if  
				if rs("jobstr") >="EV5"  then 
					tmpRec(i, j, 38) = trs("hs_cv")
				else
					tmpRec(i, j, 38) = rs("hs_bs")
				end if 
				tmpRec(i, j, 39) = rs("days_add")
				rs.MoveNext 
			else 
				exit for 		 
			end if 
			'response.write tmpRec(i, j, 0) &","&tmpRec(i, j, 2)
		next 	
		 if rs.EOF then 
			rs.Close  
			Set rs = nothing
			exit for 
		 end if 					
	next 	
end if  
Session("YECE1302B") = tmpRec 

w1=session("mywhsno") 
'if f_whsno="" then f_whsno=session("mywhsno")
%>

<body  topmargin="5" leftmargin="5"  marginwidth="0" marginheight="0"  onkeydown="enterto()" onload="f()"  >
<form name="<%=self%>" method="post" >
<INPUT TYPE=HIDDEN NAME="UID" VALUE=<%=SESSION("NETUSER")%>>
<INPUT TYPE=hidden NAME=TotalPage VALUE="<%=TotalPage%>">
<INPUT TYPE=hidden NAME=CurrentPage VALUE="<%=CurrentPage%>">
<INPUT TYPE=hidden NAME=RecordInDB VALUE="<%=RecordInDB%>">
<INPUT TYPE=hidden NAME=PageRec VALUE="<%=P%>">
<INPUT TYPE=hidden NAME=gTotalPage VALUE="<%=gTotalPage%>"> 	
<INPUT TYPE=hidden NAME="flag" VALUE="S"> 	
<table width="460" border="0" cellspacing="0" cellpadding="0">
	<tr><TD>
	<img border="0" src="../image/icon.gif" align="absmiddle">
	<%=SESSION("PGNAME")%></TD></tr>
</table>
<hr size=0	style='border: 1px dotted #999999;' align=left width=500>
<table width=500  ><tr><td >
	<table width=600  border=0 cellspacing="1" cellpadding="1"  class="txt8" >
		<tr>	
			<TD nowrap  align=right height=30 >年度<br>(Nam)</TD>			
			<TD   > 
				<input name="f_years" class="inputbox" size=6  maxlength=4 value="<%=f_years%>"  >
			</td>			
			<TD nowrap   align=right   >廠別<br>(Xuong)</TD>
			<TD  > 
				<select name="f_WHSNO"  class="txt8"  >
					<option value="" <%if trim(f_whsno)="" then %>selected<%end if%>>----</option>
					<%SQL="SELECT * FROM BASICCODE WHERE FUNC='WHSNO' ORDER BY SYS_TYPE "
					SET RST = CONN.EXECUTE(SQL)
					WHILE NOT RST.EOF  
					%>
					<option value="<%=RST("SYS_TYPE")%>" <%if f_whsno=rst("sys_type") then%>selected<%end if%> ><%=RST("SYS_TYPE")%>-<%=RST("SYS_VALUE")%></option>				 
					<%
					RST.MOVENEXT
					WEND 
					%>
				</SELECT>
				<%SET RST=NOTHING %>
			</TD> 		

			<TD nowrap  align=right height=30 >國籍<br>(Quoc tich)</TD>			
			<TD   colspan=4> 
				<%
				if session("rights")<=0  then 				
					sql="select *from basiccode where func='country' order by sys_type" 
				else	
					sql="select *from basiccode where func='country' and sys_type in ('VN','TA') order by sys_type" 
				end if 	
				set rst=conn.execute(sql)				
				%>
				<select name="f_country" class="txt8"   > 
					<option value="">----</option>
					<%while not rst.eof%>
						<option value="<%=rst("sys_type")%>"  <%if f_country=rst("sys_type") then%>selected<%end if%> ><%=rst("sys_type")%>-<%=rst("sys_value")%></option>
					<%rst.movenext
					wend
					set rst=nothing 
					%>	
					<%if session("rights")<="0" then %>
						<option value="HW" <%if f_country="HW" then%>selected<%end if%> >海外</option>						
					<%end if%>
				</select> 
			</td>			 			
		</TR>			
		<tr>
			<td align="right">工號<br>So the</td>
			<td><input name="f_empid"  class="inputbox8" size=10 value="<%=Ucase(trim(f_empid))%>"></td>
			<td align="right">部門<br>Bo phan</td>
			<td>
				<%sql="select *from basiccode where func='groupid' order by sys_type" 
				set rst=conn.execute(sql)				
				%>
				<select name="f_groupid" class="txt8"   > 
					<option value="">----</option>
					<%while not rst.eof%>
						<option value="<%=rst("sys_type")%>"  <%if f_groupid=rst("sys_type") then%>selected<%end if%> ><%=rst("sys_type")%>-<%=rst("sys_value")%></option>
					<%rst.movenext
					wend
					set rst=nothing 
					%>	
				</select> 			
			</td>
			<Td align="right">到職日>=</td>
			<td><input name="f_indat"  class="inputbox8" size=11 value="<%=Ucase(trim(f_indat))%>"></td>
			<Td align="right">考績</td>
			<td>
			<select name="f_kj" class="txt8"   > 
				<option value="">ALL</option>
				<option value="優" <%if f_kj="優" then%>selected<%end if%>>優</option>
				<option value="良" <%if f_kj="良" then%>selected<%end if%>>良</option>
				<option value="甲" <%if f_kj="甲" then%>selected<%end if%>>甲</option>
				<option value="乙" <%if f_kj="乙" then%>selected<%end if%>>乙</option>
				<option value="丙" <%if f_kj="丙" then%>selected<%end if%>>丙</option>
				<option value="N" <%if f_kj="N" then%>selected<%end if%>>未考核</option>
				<option value="X" <%if f_kj="X" then%>selected<%end if%>>無考核資料</option>
			</select>
			</td>
			<Td align="right">				
				<input type="button" name="btn" class="button" value="(S)查詢" onclick="gos()">
			</td>
		</tr>	
	</table>	
	
	<table   align=center border=0 cellspacing="1" cellpadding="1"  class="txt8" >
		<tr height=22  bgcolor="#e4e4e4">
			<Td align="center" width=30 nowrap >STT</td>									
			<Td align="center" width=40 nowrap   >年度<br>Nam</td>
			<Td align="center" width=40 nowrap   >不計<br>KO tinh</td>
			<Td align="center" width=40 nowrap   >廠別<br>Xuong</td>
			<Td align="center" width=30 nowrap   >國籍<br>Quoc tich</td>
			<Td align="center" width=40 nowrap  >工號<br>So the</td>
			<Td align="center" width=100 nowrap  >姓名<br>Ho ten</td>
			<Td align="center" width=70 nowrap  >到職日<br>NVX</td>
			<Td align="center" width=60 nowrap  >單位<br>bo phan</td> 		
			<Td align="center" width=60 nowrap  >職務<br>CV</td> 		 			
			<Td align="center" width=40 nowrap  >年資<br>so<br>thang<br>lam<br>viec</td>
			<Td align="center" width=40 nowrap  >境內<br>事假<br>(H)</td> 
			<Td align="center" width=30 nowrap  >考績<br></td>
			<Td align="center" width=60 nowrap  >年終獎金<br>1(SYS)</td>
			<Td align="center" width=60 nowrap  >其他調整<br>2(+-)</td>
			<Td align="center" width=60 nowrap  >-稅金<br>3.TAX</td>	
			<Td align="center" width=60 nowrap  >實領獎金<br>4</td>								
			<Td align="center" width=60 nowrap >年獎基準</td>						
			<Td align="center" width=40 nowrap  >天數</td>
			<Td align="center" width=40 nowrap  >C職務<br>天數</td>
			<Td align="center" width=40 nowrap  >係數</td>
						
		</tr>		
		<%for x = 1 to pagerec 
		if x mod 2 = 0 then 
			wkclr="#FFE7E7"
		else			
			wkclr="#DBE7FB"
		end if 	 
		if  tmprec(currentpage,x,4)<>"" then 
		%>
			<Tr bgcolor="<%=wkclr%>">
				<Td align="center" width=30 valign="top" rowspan=2><%=x%>
				
				</td>				 
				<Td align="center"  valign="top" rowspan=2><%=tmprec(currentpage,x,2)%></td>	
				<Td align="center"  valign="top" rowspan=2>
				<input type="checkbox" name="fnnot" onclick="fnnotchg(<%=x-1%>)">
				</td>	
				<Td align="center"  valign="top" rowspan=2>					
					<select name="whsno" class="txt8" onchange="whsnochg(<%=x-1%>)">
						<%sql2="select *from basicCode where func='whsno' and sys_type<>'All' order by sys_type" 
						  set rs2=conn.execute(Sql2)
							while  not rs2.eof 
						%>
						<option value="<%=rs2("sys_type")%>" <% if rs2("sys_type")=tmprec(currentpage,x,25) then%>selected<%end if%>><%=rs2("sys_type")%></option>
						<%rs2.movenext
						wend
						set rs2=nothing 
						%>
					</select> 
				</td>	
				<Td align="center"  valign="top" rowspan=2><%=tmprec(currentpage,x,3)%></td>	
				<Td align="center"  valign="top" rowspan=2><%=tmprec(currentpage,x,4)%></td>	
				<Td  valign="top" title="點選可看出勤與獎懲紀錄" style="cursor:hand" rowspan=2>
					<a onclick="vbscript:empdata(<%=x-1%>)"><%=tmprec(currentpage,x,6)%><br>
					<%=left(tmprec(currentpage,x,7),18)%></a>
					<input type="hidden" name="years" value="<%=tmprec(currentpage,x,2)%>" >
					<input type="hidden" name="country" value="<%=tmprec(currentpage,x,3)%>" >					
					<input type="hidden" name="empid" value="<%=tmprec(currentpage,x,4)%>" >
				</td>	
				<Td align="center" valign="top" rowspan=2><%=tmprec(currentpage,x,5)%><br><font color="red"><%=tmprec(currentpage,x,26)%></font></td>	
				<Td  valign="top" rowspan=2><%=tmprec(currentpage,x,10)%></td>	
				<Td  valign="top"  rowspan=2 ><%=tmprec(currentpage,x,31)%><br> 
				 <%=tmprec(currentpage,x,32)%></td>	
				
				<Td align="right"  valign="top" rowspan=2> 
				<input   name="nianZi" value="<%=tmprec(currentpage,x,8)%>" class="readonly8" size=5 style="text-align:right"  readonly ><br>
				<%=round(tmprec(currentpage,x,29),2)%>
				</td>					
				<Td align="center"  valign="top" rowspan=2><%=tmprec(currentpage,x,24)%></td><!--境內事假-->
				<Td  valign="top" rowspan=2>
					<input name="kj"  value="<%=tmprec(currentpage,x,13)%>" class="inputbox8"   readonly size=2 style="text-align:center"  >
					<input name="old_kj"  type="hidden" value="<%=tmprec(currentpage,x,13)%>"  >
					<input name="grade"  type="hidden" value="<%=tmprec(currentpage,x,27)%>"  >
				</td>				
				<Td valign="top" rowspan=2>
					<input name="bonus"  value="<%=formatnumber(tmprec(currentpage,x,16),0)%>" class="readonly8r"  readonly size=9   >	
					<input type='hidden' name="bonus_df"  value="<%=formatnumber(tmprec(currentpage,x,16),0)%>" class="readonly8r"  readonly size=9   >						
				</td>
				<Td  valign="top" rowspan=2><input name="khac"  value="<%=formatnumber(tmprec(currentpage,x,18),0)%>" class="inputbox8r"   size=7  onchange="khacchg(<%=x-1%>)"  ></td>
				<Td  valign="top" rowspan=2 align="right">
					<input name="tax"  value="<%=formatnumber(tmprec(currentpage,x,19),0)%>" class="readonly8r"  size=7  readonly >
					
				</td>
				<Td align="right" valign="top" rowspan=2>					
					<input name="r_bonus"  value="<%=formatnumber(tmprec(currentpage,x,17),0)%>" class="readonly8r"   size=10  readonly style="color:<%if round( cdbl(tmprec(currentpage,x,17)),0)<>round(cdbl(tmprec(currentpage,x,30)),0)then%>red;<%end if%>"  >
					<input type='hidden' name="rbonus_df"  value="<%=formatnumber(tmprec(currentpage,x,17),0)%>" class="readonly8r"  readonly size=9   >
					<span style="color:#999999"><%=formatnumber(tmprec(currentpage,x,30),0)%></span>
					<br><span id="intrst<%=x%>"><%=formatnumber(tmprec(currentpage,x,22),0)%></span>
				</td> 				
				<Td align="right" valign="top" >					
					<input  class="readonly8" name="basicBZM" value="<%=formatnumber(tmprec(currentpage,x,23),0)%>" size=10   style="text-align:right"   >
				</td>	
				<% if session("rights")<=0 then  types="text" else types="hidden" 
					types="text"
				%>	
					<Td align="center" valign="top">					
						<input type="<%=types%>" name="bodays"  value="<%=tmprec(currentpage,x,15)%>" class="readonly8" readonly   size=4 style="text-align:right">
					</td>	
					<Td align="center" valign="top"> 
						<input type="<%=types%>" name="daysadd"  value="<%=tmprec(currentpage,x,39)%>" class="readonly8" readonly   size=4 style="text-align:right">
					</td>
					<Td align="center" valign="top"> 
						<input type="<%=types%>" name="hs"  value="<%=tmprec(currentpage,x,14)%>" class="readonly8" readonly   size=4 style="text-align:right">
					</td>	
							
			</tr>			
			<tr bgcolor="<%=wkclr%>">
				<td colspan=4><input name="memos" class="inputbox8" size=40 value="<%=tmprec(currentpage,x,20)%>">
				</td>
			</tr>
			<%end if%>			
		<%next%>
		
		<input type="hidden" name="func" value="" >				
		<input type="hidden" name="years" value="" >
		<input type="hidden" name="country" value="" >
		<input type="hidden" name="whsno" value="" >
		<input type="hidden" name="empid" value="" >
		<input type="hidden" name="op" value="" >				
		<input type="hidden" name="nianzi" value="" >				
		<input type="hidden" name="kj" value="" >
		<input type="hidden" name="grade" value="" >
		<input type="hidden" name="old_kj" value="" >
		<input type="hidden" name="bonus" value="0" >
		<input type="hidden" name="khac" value="0" >
		<input type="hidden" name="tax" value="0" >
		<input type="hidden" name="r_bonus" value="0" >
		<input type="hidden" name="bodays" value="0" >
		<input type="hidden" name="hs" value="0" >
		<input type="hidden" name="basicBZM" value="0" >
		<input type="hidden" name="memos" value="" >
		<input type="hidden" name="bonus_df" value="0" > 
		<input type="hidden" name="daysadd" value="0" >
		<input type="hidden" name="fnnot" value="" >
		<input type="hidden" name="rbonus_df" value="0" > 
	</table>
	<br>
	<Table width=550>
		<tr> 
			<TD  ALIGN=center nowrap>		
			<input type="BUTTON" name="send" value="(Y)Confirm" class=button ONCLICK="GO()">
			<input type="BUTTON" name="send" value="(N)Cancel" class=button onclick="clr()">&nbsp;
			<input type="BUTTON" name="send" value="To Excel" class="button" onclick="goexcel()">&nbsp;
			<%if session("rights")="0" then%>
			<input type="BUTTON" name="send" value="(S)查詢年獎天數K.Tra" class=button onclick="gok()">
			<%end if %>
			</TD>
		</tr>
	</table>	
<hr size=0	style='border: 1px dotted #999999;' align=left >	

</td></tr></table>
<%set conn=nothing	%>
</body>
</html>
 
<!-- #include file="../Include/func.inc" -->
<script language=vbs> 

function fnnotchg(index)
	if <%=self%>.fnnot(index).checked=true  then   
		<%=self%>.bonus(index).value="0" 
		<%=self%>.r_bonus(index).value="0" 
	else
		<%=self%>.bonus(index).value=<%=self%>.bonus_df(index).value
		<%=self%>.r_bonus(index).value=<%=self%>.rbonus_df(index).value
	end if 
end function  

function f() 
		<%=self%>.f_years.focus()  
end function

function kjchg(index)
	years= trim(<%=self%>.years(index).value)
	country= trim(<%=self%>.country(index).value)
	whsno= trim(<%=self%>.whsno(index).value)
	khac = (trim(<%=self%>.khac(index).value))		
	basicBZM =replace((trim(<%=self%>.basicBZM(index).value)),",","")
	code01 = escape(ucase(trim(<%=self%>.kj(index).value))) 
	nz = ucase(trim(<%=self%>.nianzi(index).value)) 
	if trim(<%=self%>.kj(index).value)<>""   then 
		if   trim(<%=self%>.kj(index).value)<>trim(<%=self%>.old_kj(index).value) then 
			open "<%=self%>.back.asp?func=chkkj&index="& index &"&code01="&  code01 &"&khac="& khac &"&years="& years &"&whsno="& whsno &"&country="& country  &"&basicBZM="& basicBZM &"&nz="& nz , "Back"
			parent.best.cols="50%,50%"
		end if 	
	end if   
end function    

function goexcel() 
	parent.best.cols="100%,0%"
	<%=self%>.action="<%=self%>.toexcel.asp"
	<%=self%>.target="Back"
	<%=self%>.submit()
end function 

function khacchg(index)
	years= trim(<%=self%>.years(index).value)
	country= trim(<%=self%>.country(index).value)
	whsno= trim(<%=self%>.whsno(index).value)
	khac = (trim(<%=self%>.khac(index).value))		
	basicBZM =replace((trim(<%=self%>.basicBZM(index).value)),",","")
	code01 = escape(ucase(trim(<%=self%>.kj(index).value)))
	nz = ucase(trim(<%=self%>.nianzi(index).value)) 	
	if khac<>"" and code01<>""  then 
		open "<%=self%>.back.asp?func=chkkj&index="& index &"&CODE01="&  CODE01 &"&khac="& khac &"&years="& years &"&whsno="& whsno &"&country="& country &"&basicBZM="& basicBZM &"&nz="& nz , "Back"
		parent.best.cols="100%,0%"
	end if  
end function    
 
 function whsnochg(index)
 	years= trim(<%=self%>.years(index).value)
	country= trim(<%=self%>.country(index).value)
	whsno= trim(<%=self%>.whsno(index).value)
	khac = (trim(<%=self%>.khac(index).value))		
	basicBZM =replace((trim(<%=self%>.basicBZM(index).value)),",","")
	code01 = escape(ucase(trim(<%=self%>.kj(index).value)))
	
	nz = ucase(trim(<%=self%>.nianzi(index).value)) 
	if trim(<%=self%>.kj(index).value)<>""   then 		
		open "<%=self%>.back.asp?func=chkkj&index="& index &"&CODE01="&  CODE01 &"&khac="& khac &"&years="& years &"&whsno="& whsno &"&country="& country &"&basicBZM="& basicBZM &"&nz="& nz , "Back"
		'parent.best.cols="50%,50%"		
	end if  		
 end function 



function gos()
	' pg = <%=self%>.pagerec.value
	' if <%=self%>.years.value<>"" then 
		' for x = 1 to 6 
			' if trim(<%=self%>.nam(x-1).value)="" then 
				' <%=self%>.nam(x-1).value = trim(<%=self%>.years.value)
			' end if 	
		' next 
	' end if 
	<%=self%>.totalpage.value="0"
	<%=self%>.action="<%=self%>.fore.asp"
	<%=self%>.target="Fore"
	<%=self%>.submit()
end function 

function gok()
	wt = (window.screen.width )*0.6
	ht = window.screen.availHeight*0.6
	tp = (window.screen.width )*0.02
	lt = (window.screen.availHeight)*0.02 
	years = <%=self%>.f_years.value 
	
	open  "yece1301.show.asp?years="&years , "_blank", "top="& tp &", left="& lt &", width="& wt &",height="& ht &", scrollbars=yes"		 
	
end function  

function del(index)
	if <%=self%>.func(index).checked=true then 
		<%=self%>.op(index).value="D"
	else
		<%=self%>.op(index).value=""
	end if 
end function 

function Gocopy()
	open "<%=self%>B.fore.asp" , "_self"
end function 

function clr()
	open "<%=SELF%>.asp" , "_self"
end function 

function dayschg(index)
	if <%=self%>.days(index).value<>"" then 
		if isnumeric(<%=self%>.days(index).value)=false then 
			alert "請輸入數字,xin danh lai [so] !!"
			<%=self%>.days(index).value=""
			<%=self%>.days(index).focus()
			exit function 
		else
			<%=self%>.hs(index).value = round(cdbl(<%=self%>.days(index).value)/30+0.001,2)
		end if 
	end if 	
end function 

function tot_Mtaxchg(index,a)
	if a=1 then 
		if <%=self%>.person_qty(index).value<>"" then 
			if isnumeric(<%=self%>.person_qty(index).value)=false then 
				alert "請輸入數字,xin danh lai [so] !!"
				<%=self%>.person_qty(index).value=""
				<%=self%>.person_qty(index).focus()
				exit function 
			end if 
		end if 	
	elseif a=2 then 
		if <%=self%>.ut_mtax(index).value<>"" then 
			if isnumeric(<%=self%>.ut_mtax(index).value)=false then 
				alert "請輸入數字,xin danh lai [so] !!"
				<%=self%>.ut_mtax(index).value=""
				<%=self%>.ut_mtax(index).focus()
				exit function 
			else	
				<%=self%>.ut_mtax(index).value=formatnumber(<%=self%>.ut_mtax(index).value,0)
			end if 
		end if 	
	end if 
	
	if trim(<%=self%>.ut_mtax(index).value)<>"" and trim(<%=self%>.person_qty(index).value)<>"" then 
		<%=self%>.tot_Mtax(index).value=formatnumber( cdbl(<%=self%>.person_qty(index).value)*cdbl(<%=self%>.ut_mtax(index).value) , 0)
	end if 
end function 

function empidchg(index)
	code1=UCase(trim(<%=self%>.empid(index).value))
	if <%=self%>.empid(index).value<>"" then 		
		open "<%=self%>.back.asp?func=chkemp&code1="& code1 &"&index="& index  , "Back"
		'parent.best.cols="50%,50%"
	end if 
end function 
function strchg(a)
	if a=1 then
		<%=self%>.empid1.value = Ucase(<%=self%>.empid1.value)
	elseif a=2 then
		<%=self%>.empid2.value = Ucase(<%=self%>.empid2.value)
	end if
end function

function go()  
	<%=self%>.action="<%=SELF%>.upd.asp"
	<%=self%>.target="Fore"
 	<%=self%>.submit() 
end function


'*******檢查日期*********************************************
FUNCTION date_change(a)

if a=1 then
	INcardat = Trim(<%=self%>.indat1.value)
elseif a=2 then
	INcardat = Trim(<%=self%>.indat2.value)
end if

IF INcardat<>"" THEN
	ANS=validDate(INcardat)
	IF ANS <> "" THEN
		if a=1 then
			Document.<%=self%>.indat1.value=ANS
		elseif a=2 then
			Document.<%=self%>.indat2.value=ANS
		end if
	ELSE
		ALERT "EZ0067:輸入日期不合法 !!"
		if a=1 then
			Document.<%=self%>.indat1.value=""
			Document.<%=self%>.indat1.focus()
		elseif a=2 then
			Document.<%=self%>.indat2.value=""
			Document.<%=self%>.indat2.focus()
		end if
		EXIT FUNCTION
	END IF

ELSE
	'alert "EZ0015:日期該欄位必須輸入資料 !!"
	EXIT FUNCTION
END IF 
END FUNCTION 

function empdata(index)
	wt = (window.screen.width )*0.8
	ht = window.screen.availHeight*0.7
	tp = (window.screen.width )*0.02
	lt = (window.screen.availHeight)*0.02	
	country=<%=self%>.country(index).value 
	empid=<%=self%>.empid(index).value 
	if country="VN" then 
		open "YEBQ01B.editVN.asp?empid="& empid   , "_balnk"  , "top="& tp &", left="& lt &", width="& wt &",height="& ht &", scrollbars=yes"  
		'open "YEBQ01B.editVN.asp?empid="& empid   , "_self"  
	else
		open "YEBQ01B.editHW.asp?empid="& empid   , "_balnk"  , "top="& tp &", left="& lt &", width="& wt &",height="& ht &", scrollbars=yes"  
	end if 	
end function 
</script> 
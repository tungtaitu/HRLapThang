<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<!-- #include file = "../GetSQLServerConnection.fun" -->
<!-- #include file="../ADOINC.inc" -->
<!-- #include file="../Include/func.inc" -->
<!--#include file="../include/sideinfo.inc"-->
<%
Set conn = GetSQLServerConnection()
Set rs = Server.CreateObject("ADODB.Recordset")
self="YECE0202"  
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

yymm=request("YYMM" )  

calcdt = left(YYMM,4)&"/"& right(yymm,2)&"/01"
'下個月
if right(yymm,2)="12" then
	ccdt = cstr(left(YYMM,4)+1)&"/01/01"
else
	ccdt = left(YYMM,4)&"/"& right("00" & right(yymm,2)+1,2)  &"/01"
end if
'response.write ccdt

 '一個月有幾天
cDatestr=CDate(LEFT(YYMM,4)&"/"&RIGHT(YYMM,2)&"/01")
days = DAY(cDatestr+(32-DAY(cDatestr))-DAY(cDatestr+(32-DAY(cDatestr))))   '一個月有幾天
'本月最後一天
ENDdat = CDate(LEFT(YYMM,4)&"/"&RIGHT(YYMM,2)&"/"&DAYS)
ENDdat=year(ENDdat)&"/"&right("00"&month(Enddat),2)&"/"&right("00"&day(Enddat),2) 


if instr(session("vnlogip"),"168")>0 then 
	whsno="LA"
elseif instr(session("vnlogip"),"169")>0 then 
	whsno="DN"
elseif instr(session("vnlogip"),"47")>0 then 
	whsno="BC" 
else 
	whsno="LA"
end if 	 
eid = request("eid")		
GROUPID = request("GROUPID")	

gTotalPage = 1
PageRec = 15    'number of records per page
TableRec = 40    'number of fields per record  

if GROUPID <>"" then 
	sqlx="delete AEMpJB_Edat where yymm='"&yymm&"' and "&_
			 "empid in ( select empid from bempg where yymm='"&yymm&"' and groupid='"&GROUPID&"' ) " 
elseif eid<>"" then 
	sqlx="delete AEMpJB_Edat where yymm='"&yymm&"' and  empid='"&eid&"'" 
else
	sqlx="delete AEMpJB_Edat where yymm='"&yymm&"'  "
end if 	
conn.execute(sqlx)
'response.write sqlx 
			 
	sql="select isnull(convert(char(10),b.endJBdat,111),'') endJBdat,  isnull(b.yymm,'') as eyymm, isnull(n.h1,0) as totjbh, "&_
		"isnull(n.forget,0) forget, isnull(n.h1,0) h1, isnull(n.h2,0) h2 , isnull(n.h3,0) h3, isnull(n.b3,0) b3 ,"&_
		"isnull(n.nh1,0) nh1, isnull(n.nh2,0) nh2 , isnull(n.nh3,0) nh3, isnull(n.nb3,0) nb3 ,"&_
		"isnull(n.ovh1,0) ovh1, isnull(n.ovh2,0) ovh2 , isnull(n.ovh3,0) ovh3, isnull(n.ovb3,0) ovb3 ,"&_
		"isnull(n.kzhour,0) kzhour, isnull(n.latefor,0) latefor, "&_		
		"isnull(bs.bb,0) as bs_bb, isnull(bs.cv,0) as  bs_cv, isnull(bs.phu,0) as bs_phu, "&_
		"isnull(r_bb,0) r_bb , isnull(r_cv,0) r_cv, isnull(r_phu,0) r_phu, isnull(money_h,0) money_h  , "&_
		" a.* from  "&_
		"( select * from  view_empfile where  CONVERT(CHAR(10), indat, 111)< '"& ccdt &"' and ( isnull(outdate,'')='' or outdate>'"& calcdt &"' )  "&_
		" and    COUNTRY='VN' and empid like '"& eid &"%' "&_ 
		" ) a  "&_ 
		"left join ( select * from view_empgroup where yymm='"& yymm &"' ) ax on  ax.empid = a.empid  "&_
		"left join (select *from AEMpJB_Edat where yymm = '"& yymm &"' ) b on b.empid = a.empid "&_
		"left join ( select* from bemps where yymm='"& yymm &"' ) bs on bs.empid = a.empid  "&_
		"left join ( select empid, whsno, yymm,  bb R_bb, cv r_cv , phu  r_phu , money_h  from empdsalary  where yymm='"& yymm &"' ) d on d.empid = a.empid  "&_
		"LEFT JOIN ( select empid empidN,  (sum(isnull(forget,0)))  forget  , (sum(isnull(h1,0))) h1, (sum(isnull(h2,0))) h2, (sum(isnull(h3,0))) h3, (sum(isnull(b3,0))) b3 ,  "&_		
		"(sum(isnull(nh1,0))) nh1, (sum(isnull(nh2,0))) nh2, (sum(isnull(nh3,0))) nh3, (sum(isnull(nb3,0))) nb3 , "&_
		"(sum(isnull(ovh1,0))) ovh1, (sum(isnull(ovh2,0))) ovh2, (sum(isnull(ovh3,0))) ovh3, (sum(isnull(ovb3,0))) ovb3 , "&_
	 	"(sum(isnull(jiaa,0))) jiaa, (sum(isnull(jiab,0))) jiab, ( sum(isnull(toth,0))) toth , ( sum(isnull(kzhour,0))) kzhour , (sum(latefor)) latefor "&_
	 	"from empwork   where yymm='"& YYMM &"' GROUP BY EMPID )  N ON N.empidN = A.EMPID  "&_	 
		"where datediff(d, a.indat, isnull(a.outdat,'"& ENDdat &"')) >=1  and isnull(ax.lw,'')='"& whsno &"' and ax.lg like'"& groupid&"%' " 		
		sql=sql&"order by a.groupid, a.shift,a.empid   "  

		'response.write sql &"<br>"
		'response.end  

if request("TotalPage") = "" or request("TotalPage") = "0" then	
	'sql2="exec SP_calcEMPJBEdat '"& yymm &"' , '"& eid &"' , '"& GROUPID &"','"&session("netuser")&"' " 
	'response.write sql2 
	'conn.execute(Sql2) 	
	
	CurrentPage = 1
	rs.Open SQL, conn, 3, 3  
	IF NOT RS.EOF THEN
		rs.PageSize = PageRec
		RecordInDB = rs.RecordCount
		TotalPage = rs.PageCount
		gTotalPage = TotalPage
	END IF
'	response.write RecordInDB 
'	response.end 
	Redim tmpRec(gTotalPage, PageRec, TableRec)   'Array

	for i = 1 to TotalPage
	 for j = 1 to PageRec
		if not rs.EOF then
				tmpRec(i, j, 0) = "no"
				tmpRec(i, j, 1) = trim(rs("empid"))
				tmpRec(i, j, 2) = trim(rs("empnam_cn"))
				tmpRec(i, j, 3) = trim(rs("empnam_vn"))
				tmpRec(i, j, 4) = rs("country")
				tmpRec(i, j, 5) = rs("nindat")
				tmpRec(i, j, 6) = rs("job")
				tmpRec(i, j, 7) = rs("whsno")
				tmpRec(i, j, 8) = rs("unitno")
				tmpRec(i, j, 9)	=RS("groupid")
				tmpRec(i, j, 10)=RS("zuno")
				tmpRec(i, j, 11)=RS("wstr")
				tmpRec(i, j, 12)=RS("ustr")
				tmpRec(i, j, 13)=RS("gstr")
				tmpRec(i, j, 14)=RS("zstr")
				tmpRec(i, j, 15)=RS("jstr")
				tmpRec(i, j, 16)=RS("cstr")
				tmpRec(i, j, 17)=RS("outdate")
				
				tmpRec(i, j, 18)=RS("endJBdat")
				tmpRec(i, j, 19)=RS("eyymm")
				tmpRec(i, j, 20)=RS("totJBH") 				
				
				h1 = rs("h1")
				h2 = rs("h2")
				h3 = rs("h3")
				b3 = rs("b3")
				nh1 = rs("nh1")
				nh2 = rs("nh2")
				nh3 = rs("nh3")
				nb3 = rs("nb3")
				ovh1 = rs("ovh1")
				ovh2 = rs("ovh2")
				ovh3 = rs("ovh3")
				ovb3 = rs("ovb3")
				
				tmpRec(i, j, 21)=nh1
				tmpRec(i, j, 22)=nh2
				tmpRec(i, j, 23)=nh3
				tmpRec(i, j, 24)=ovh1
				tmpRec(i, j, 25)=ovh2
				tmpRec(i, j, 26)=ovh3
				tmpRec(i, j, 27)=rs("B3")
				
				tmpRec(i, j, 28)=RS("bs_bb")
				tmpRec(i, j, 29)=RS("bs_cv")
				tmpRec(i, j, 30)=RS("bs_phu")
				tmpRec(i, j, 31)=round( (cdbl(RS("bs_bb"))+cdbl(RS("bs_cv"))+cdbl(RS("bs_phu")))/26/8, 0)

				tmpRec(i, j, 32)= round(cdbl(tmpRec(i, j, 31))*1.5*nh1,0) + round(cdbl(tmpRec(i, j, 31))*2*nh2,0) + round(cdbl(tmpRec(i, j, 31))*3*nh3,0) 
				tmpRec(i, j, 33)= round(cdbl(tmpRec(i, j, 31))*1.5*ovh1,0) + round(cdbl(tmpRec(i, j, 31))*2*ovh2,0) + round(cdbl(tmpRec(i, j, 31))*3*ovh3,0) 
				tmpRec(i, j, 34)=cdbl(tmpRec(i, j, 32)) + cdbl(tmpRec(i, j, 33)) 
				' sqlx="select a.empid, sum(b3) b3, "&_
						 ' "sum( case when isnull(b.endjbdat,'')='' or a.workdat <= convert(char(8),b.endjbdat,112) then b3  "&_
						 ' "else case when b3>0 and b3>h1 then b3-h1 else 0 end end ) as b31 "&_
						 ' "from (select * from empwork where yymm='"&yymm&"' and empid='"& trim(tmprec(i,j,1)) &"'  ) a  "&_
						 ' "left join (select * from AEMpJB_Edat where yymm='"&yymm&"' ) b on b.empid= a.empid "&_
						 ' "group by a.empid  "  				
				' set rsx=conn.execute(Sqlx)	
                ' if rsx.eof then  				
					' tmpRec(i, j, 35) = rsx("b31")  '合理的夜班 	
					' tmpRec(i, j, 36)=cdbl(rs("b3"))-cdbl(tmpRec(i, j, 35))  'over夜班					
				' else	
					' tmpRec(i, j, 35) = 0	
					' tmpRec(i, j, 36)=cdbl(rs("b3"))-cdbl(tmpRec(i, j, 35))  'over夜班				
				' end if 
				' tmpRec(i, j, 36)=cdbl(rs("b3"))-cdbl(tmpRec(i, j, 35))  'over夜班
				' set rsx=nothing 
				tmpRec(i, j, 35) = nb3
				tmpRec(i, j, 36) = ovb3
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
	Session("YECE0202B") = tmpRec
else
	ccdt = request("ccdt")
	calcdat = request("calcdat")
	enddat = request("enddat")
	YYMM = request("yymm")
	MMDAYS = request("MMDAYS")
	TotalPage = cint(request("TotalPage"))
	'StoreToSession()
	CurrentPage = cint(request("CurrentPage"))
	RecordInDB  = REQUEST("RecordInDB")
	tmpRec = Session("YECE0202B")

	Select case request("send")
	     Case "FIRST"
		      CurrentPage = 1
	     Case "BACK"
		      if cint(CurrentPage) <> 1 then
			     CurrentPage = CurrentPage - 1
		      end if
	     Case "NEXT"
		      if cint(CurrentPage) < cint(TotalPage) then
			     CurrentPage = CurrentPage + 1
			  else
			  	 CurrentPage = TotalPage
		      end if
	     Case "END"
		      CurrentPage = TotalPage
	     Case Else
		      CurrentPage = 1
	end Select
end if

'response.end

FUNCTION FDT(D)
	Response.Write YEAR(D)&"/"&RIGHT("00"&MONTH(D),2)&"/"&RIGHT("00"&DAY(D),2)

END FUNCTION 

%>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="refresh">
<link rel="stylesheet" href="../Include/style.css" type="text/css">
<link rel="stylesheet" href="../Include/style2.css" type="text/css">
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--
function f()
	<%=self%>.YYMM.focus()
	<%=self%>.YYMM.SELECT()
end function

function god()
	<%=self%>.totalpage.value="" 
	<%=self%>.submit()
end function  


-->
</SCRIPT>
</head>
<body  topmargin="5" leftmargin="5"  marginwidth="0" marginheight="0"  onkeydown="enterto()" onload='f()' >
<form name="<%=self%>" method="post"  action="<%=self%>.foregnd.asp">
<INPUT TYPE=HIDDEN NAME="UID" VALUE=<%=SESSION("NETUSER")%>>
<INPUT TYPE=hidden NAME=TotalPage VALUE="<%=TotalPage%>">
<INPUT TYPE=hidden NAME=CurrentPage VALUE="<%=CurrentPage%>">
<INPUT TYPE=hidden NAME=RecordInDB VALUE="<%=RecordInDB%>">
<INPUT TYPE=hidden NAME=PageRec VALUE="<%=PageRec%>">
<INPUT TYPE=hidden NAME=gTotalPage VALUE="<%=gTotalPage%>">
<INPUT TYPE='hidden' NAME=calcdat VALUE="<%=calcdt%>">
<INPUT TYPE='hidden' NAME=enddat VALUE="<%=year(enddat)&"/"&right("00"&month(enddat),2)&"/"&right("00"&day(enddat),2)%>">
<INPUT TYPE='hidden' NAME=ccdt VALUE="<%=ccdt%>"> 

	<table border=0 style="height:70px"><tr><td>&nbsp;</td></tr></table>
	<table width="100%" border=0 >
		<tr>
			<td>
				<table width="100%" BORDER=0 align=center cellpadding=0 cellspacing=0 >
					<tr>
						<td>
							<table class="txt" cellpadding=3 cellspacing=3>
								<tr height=30 >
									<TD nowrap align=right>計薪年月：</TD>
									<TD ><INPUT type="text" style="width:100px" NAME=YYMM  VALUE="<%=yymm%>" onchange="god()"></TD>
									<td align=right >部門:</td>			
									<TD nowrap >
										<select name=GROUPID   style="width:120px"  onchange="gos()">
										<option value="" selected >--All--</option>
										<%
										SQL="SELECT * FROM BASICCODE WHERE FUNC='GROUPID' and sys_type <>'AAA' ORDER BY SYS_TYPE "
										'SQL="SELECT * FROM BASICCODE WHERE FUNC='GROUPID'  ORDER BY SYS_TYPE "
										SET RST = CONN.EXECUTE(SQL)
										'RESPONSE.WRITE SQL 
										WHILE NOT RST.EOF  
										%>
										<option value="<%=RST("SYS_TYPE")%>" <%if rst("sys_type")=groupid then%>selected<%end if%> ><%=RST("SYS_TYPE")%><%=RST("SYS_VALUE")%></option>				 
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
					</tr>
					<tr>
						<td align="center">
							<table id="myTableGrid" width="98%"> 
								<Tr bgcolor="#e4e4e4" height=22 class="txt9">
									<td rowspan=2 width=30 nowrap align=center>STT</td>
									<td rowspan=2 width=60 nowrap align=center>單位</td>
									<td rowspan=2 width=50 nowrap align=center>工號</td>
									<td rowspan=2 width=110 nowrap align=center>姓名</td>
									<td rowspan=2 width=70 nowrap align=center>到職日<br>離職日</td>				
									<td rowspan=2 width=70 nowrap align=center>加班<br>截止日</td>
									<td rowspan=2 width=40 nowrap align=center>累計(H)</td>
									<td rowspan=2 width=40 nowrap align=center>累計B3(H)</td>
									<td rowspan=2 width=50 nowrap align=center> 時薪<BR>(VND)</td>
									<Td colspan=5 align=center align=center>正常加班</td>
									<Td colspan=5 align=center bgcolor="#CCFF99">Over JB</td>
									<td rowspan=2 width=60 nowrap align=center>總計</td>
								</tr>
								<tr bgcolor="#e4e4e4" height=22 class="txt9">
									<td width=30 nowrap align=center>h1(H)</td>
									<td width=30 nowrap align=center>h2(H)</td>
									<td width=30 nowrap align=center>h3(H)</td>
									<td width=30 nowrap align=center>B3(H)</td>
									<td width=50 nowrap align=center>合計</td>
									<td bgcolor="#CCFF99" width=30 nowrap align=center>h1(H)</td>
									<td bgcolor="#CCFF99" width=30 nowrap align=center>h2(H)</td>
									<td bgcolor="#CCFF99"width=30 nowrap align=center>h3(H)</td>
									<td bgcolor="#CCFF99"width=30 nowrap align=center>B3(H)</td>
									<td width=50 nowrap bgcolor="#CCFF99"  align=center>合計</td>
								</tr>
							<%for x = 1 to PageRec
								IF x MOD 2 = 0 THEN
									WKCOLOR="LavenderBlush"
									'wkcolor="#ffffff"
								ELSE
									WKCOLOR="#DFEFFF"
									'wkcolor="#ffffff"
								END IF
								if tmpRec(CurrentPage, x, 1) <> "" then
							%>
								<tr bgcolor="<%=wkcolor%>" class="txt9">
									<td><%=((currentpage-1)*pagerec)+x%></td>
									<td nowrap><a href="vbscript:showworTim(<%=x-1%>)"><%=tmprec(currentpage,x,13)%></a></td>
									<td nowrap><a href="vbscript:showworTim(<%=x-1%>)"><%=tmprec(currentpage,x,1)%></a></td>
									<td nowrap><a href="vbscript:showworTim(<%=x-1%>)"><%=tmprec(currentpage,x,2)%><br><%=tmprec(currentpage,x,3)%></a></td>
									<input type="hidden" name="empid" value="<%=tmprec(currentpage,x,1)%>" > 			
									<td><%=tmprec(currentpage,x,5)%><br><%=tmprec(currentpage,x,17)%></td>
									<td><%=tmprec(currentpage,x,18)%></td>
									<td align=right><%=tmprec(currentpage,x,20)%></td>			
									<td align=right><%=tmprec(currentpage,x,27)%></td>	 <!--累計夜班-->
									<td align=right><%=formatnumber(tmprec(currentpage,x,31),0)%></td>			<!--時薪-->

									<Td align=right><%=tmprec(currentpage,x,21)%></td>	
									<Td align=right><%=tmprec(currentpage,x,22)%></td>	
									<Td align=right><%=tmprec(currentpage,x,23)%></td>
									<Td align=right><%=tmprec(currentpage,x,35)%></td>
									<Td align=right><%=formatnumber(tmprec(currentpage,x,32),0)%></td>	
									<Td align=right bgcolor="#CCFF99"><%=tmprec(currentpage,x,24)%></td>	
									<Td align=right bgcolor="#CCFF99"><%=tmprec(currentpage,x,25)%></td>	
									<Td align=right bgcolor="#CCFF99"><%=tmprec(currentpage,x,26)%></td>				
									<Td align=right bgcolor="#CCFF99"><%=tmprec(currentpage,x,36)%></td>	
									<Td align=right bgcolor="#CCFF99"><%=formatnumber(tmprec(currentpage,x,33),0)%></td>				
									<Td align=right><%=formatnumber(tmprec(currentpage,x,34),0)%></td>	
								</tr>
								<%end if%>
							<%next%>	
							<input type="hidden" name="empid" value="" >
							</table>
						</td>
					</tr>
					<tr>
						<td align="center">
							<table class="txt">
								<tr class=font9>
									<td align="CENTER" height=40 WIDTH=60%>
									<% If CurrentPage > 1 Then %>
										<input type="submit" name="send" value="FIRST" class="btn btn-sm btn-outline-secondary">
										<input type="submit" name="send" value="BACK" class="btn btn-sm btn-outline-secondary">
									<% Else %>
										<input type="submit" name="send" value="FIRST" disabled class="btn btn-sm btn-outline-secondary">
										<input type="submit" name="send" value="BACK" disabled class="btn btn-sm btn-outline-secondary">
									<% End If %>
									<% If cint(CurrentPage) < cint(TotalPage) Then %>
										<input type="submit" name="send" value="NEXT" class="btn btn-sm btn-outline-secondary">
										<input type="submit" name="send" value="END" class="btn btn-sm btn-outline-secondary">
									<% Else %>
										<input type="submit" name="send" value="NEXT" disabled class="btn btn-sm btn-outline-secondary">
										<input type="submit" name="send" value="END" disabled class="btn btn-sm btn-outline-secondary">
									<% End If %>
									<FONT CLASS=TXT8>&nbsp;&nbsp;PAGE:<%=CURRENTPAGE%> / <%=TOTALPAGE%> , COUNT:<%=RECORDINDB%></FONT>
									</TD>
									<TD WIDTH=40% ALIGN=RIGHT>
										<%if session("netuser")="LSARY" then %>
											<%if yymm>=nowmonth then %>
												<input type="BUTTON" name="send" value="(Y)確　認" class="btn btn-sm btn-outline-secondary" ONCLICK="GO()">
												<input type="BUTTON" name="send" value="(N)取　消" class="btn btn-sm btn-outline-secondary" onclick="clr()">
											<%end if%>
										<%else%>	
											<input type="button" name="send" value="(Y)確　定" class="btn btn-sm btn-danger" onclick="go()" >
											<input type="BUTTON" name="send" value="(N)取　消" class="btn btn-sm btn-outline-secondary"  onclick="clr()">
										<%end if%>	
									</TD>
								</TR>
							</TABLE>
						</td>
					</tr>
				</table>
			</td>
		</tr>
	</table>
</form>
</body>
</html>


<script language=vbs> 



function showworTim(index)
	yymm_str = <%=self%>.yymm.value
	empid_str = <%=self%>.empid(index).value
	
	wt = (window.screen.width )*0.8
	ht = window.screen.availHeight*0.7
	tp = (window.screen.width )*0.02
	lt = (window.screen.availHeight)*0.02	
	
	OPEN "../zzz/getempWorkTime.asp?yymm=" & yymm_str &"&EMPID=" & empid_str , "_blank" , "top="& tp &", left="& lt &", width="& wt &",height="& ht &", scrollbars=yes"   	
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
 	<%=self%>.submit() 
end function

function gos()
	<%=self%>.totalpage.value=""
	<%=self%>.action ="<%=self%>.foregnd.asp"
	<%=self%>.submit() 
end function 

 
</script> 
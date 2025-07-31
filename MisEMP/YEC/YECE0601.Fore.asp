<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%> 
<!-- #include file = "../GetSQLServerConnection.fun" --> 
<!-- #include file="../ADOINC.inc" -->
<!-- #include file="../Include/func.inc" -->
<!-- #include file="../Include/SIDEINFO.inc" -->
<%

self="YECE0601"   

nowmonth = year(date())&right("00"&month(date()),2)
if month(date())="01" then  
	calcmonth = year(date()-1)&"12"  	
else
	calcmonth =  year(date())&right("00"&month(date())-1,2)   
end if 	   

if day(date())<=11 then 
	if month(date())="01" then  
		calcmonth = year(date()-1)&"12" 
	else
		calcmonth =  year(date())&right("00"&month(date())-1,2)   
	end if 	 	
else
	calcmonth = nowmonth 
end if  

if right(calcmonth,2)="01" then 
	sgym = left(calcmonth,4)-1 & "12" 
else
	sgym = left(calcmonth,4)&right("00"&right(calcmonth,2)-1,2)
end if 	 




YYMM = request("YYMM")
if YYMM="" then YYMM=calcmonth 
JXYM = request("JXYM") 
if JXYM="" then JXYM=sgym  

m_f_day =  left(JXYM,4)&"/"&right(JXYM,2)&"/01"

Set conn = GetSQLServerConnection()	  
Set rs = Server.CreateObject("ADODB.Recordset") 

F_WHSNO = request("F_WHSNO")
F_groupid=request("F_groupid") 
F_shift=request("F_shift") 
country=request("country") 
empid1=request("empid1") 
ktrajx = request("ktrajx")

gTotalPage = 1
PageRec = 10    'number of records per page
TableRec = 50    'number of fields per record 

Redim tmpRec(gTotalPage, PageRec, TableRec)   'Array
if request("sflag")="" then 
	sql="select * from empfile where empid='X'"
else
	sql="select  isnull(c.jxyn,'') jxyn, isnull(c.memo,'') jxynmemo, isnull(c.aid,'') c_aid,  a.empid, a.empnam_cn, a.empnam_vn, a.nindat, a.outdate,  a.bhdat, a.country, "&_
			"d.lj, d.ljstr,  b.* from   "&_
			"( select * from view_empfile where ( isnull(outdat,'')='' or outdate>'"& m_f_day &"')  ) a "&_ 
			"left join ( select * from view_empgroup  where yymm='"& jxym &"')  b on b.empid = a.empid   "&_
			"left join ( select * from view_empjob  where yymm='"& jxym &"')  d on d.empid = a.empid   "&_
			"left join (select * from Empjxyn where yymm='"& jxym &"'  ) c on  c.empid = a.empid "&_
			"where isnull(b.lw,'') like '"& f_whsno &"%' and isnull(b.lg,'') like '"& f_groupid &"%' "&_
			"and isnull(b.ls,'') like '"& f_shift &"%' and a.empid like '%"& empid1 &"%' and a.country like '"& country &"%' "
	if ktrajx="" then 
		sql=sql&" and a.empid<>'' " 
	elseif ktrajx="Y" then 
		sql=sql&" and isnull(c.jxyn,'')='Y' " 
	else	
		sql=sql&" and isnull(c.jxyn,'N')='N' " 
	end if 
		sql=sql&"order by a.empid "
	'response.write sql 		
	if request("TotalPage") = "" or request("TotalPage") = "0" then
		CurrentPage = 1
		rs.Open SQL, conn, 3, 3  
		IF NOT RS.EOF THEN
			PageRec = rs.RecordCount
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
				tmpRec(i, j, 0) = trim(rs("jxyn"))
				tmpRec(i, j, 1) = trim(rs("empid"))
				if trim(rs("jxyn"))="" then tmpRec(i, j, 0)="Y"
				tmpRec(i, j, 2) = trim(rs("empnam_cn"))
				tmpRec(i, j, 3) = trim(rs("empnam_vn"))
				tmpRec(i, j, 4) = trim(rs("country"))
				tmpRec(i, j, 5) = trim(rs("nindat"))
				tmpRec(i, j, 6) = trim(rs("bhdat"))
				tmpRec(i, j, 7) = trim(rs("lw"))
				tmpRec(i, j, 8) = trim(rs("lg"))
				tmpRec(i, j, 9) = trim(rs("lz"))
				tmpRec(i, j, 10) = trim(rs("lj"))
				tmpRec(i, j, 11) = trim(rs("lgstr"))
				tmpRec(i, j, 12) = trim(rs("lzstr"))
				tmpRec(i, j, 13) = trim(rs("outdate"))
				tmpRec(i, j, 14) = trim(rs("jxynmemo"))
				tmpRec(i, j, 15) = trim(rs("c_aid"))
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
		Session("yece0601B") = tmpRec
	else
		
	end if 
	
	
end if 

%>

<html> 
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="refresh">
<link rel="stylesheet" href="../Include/style.css" type="text/css">
<link rel="stylesheet" href="../Include/style2.css" type="text/css">
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--
function f()
	'<%=self%>.salaryYM.focus()		
end function 

function dchg()
	<%=self%>.action = "<%=self%>.Fore.asp"
	<%=self%>.submit()
end  function   

function newPage()
	open "yece0601.asp"  , "_blank" , "top=10, left=10"
	
end function 
   
-->
</SCRIPT>   
</head> 
<body leftmargin="5"  marginwidth="0" marginheight="0"  onkeydown="enterto()" onload='f()' >
<form name="<%=self%>" method="post" action="<%=self%>.Fore.asp">
<input type="hidden" name="sflag" value="T">
<INPUT TYPE=HIDDEN NAME="UID" VALUE=<%=SESSION("NETUSER")%>>
<INPUT TYPE=hidden NAME=TotalPage VALUE="<%=TotalPage%>">
<INPUT TYPE=hidden NAME=CurrentPage VALUE="<%=CurrentPage%>">
<INPUT TYPE=hidden NAME=RecordInDB VALUE="<%=RecordInDB%>">
<INPUT TYPE=hidden NAME=PageRec VALUE="<%=PageRec%>">
<INPUT TYPE=hidden NAME=gTotalPage VALUE="<%=gTotalPage%>">
 
	<table border=0 style="height:70px"><tr><td>&nbsp;</td></tr></table>
	<table width="100%" border=0 >
		<tr>
			<td>
				<table width="80%" BORDER=0 align=center cellpadding=0 cellspacing=0 >
					<tr>
						<td>
							<table class="table-borderless table-sm bg-white text-secondary">		
								<tr height=30  class="txt8">
									<TD nowrap align=right>績效年月<br>YYMM</TD>
									<TD ><INPUT NAME=JXYM  CLASS=INPUTBOX VALUE="<%=jxym%>" SIZE=10 maxlength=6></TD>	
									<TD nowrap align=right height=30 >廠別<br>Xuong</TD>
									<TD > 
										<select name=F_WHSNO  class=txt8 style="width:70"  onchange="GETDATA()"> 
											<%
											if session("rights")="0" then 
												SQL="SELECT * FROM BASICCODE WHERE FUNC='WHSNO' ORDER BY SYS_TYPE "
											%><option value=""></option>
											<%	
											else		
												SQL="SELECT * FROM BASICCODE WHERE FUNC='WHSNO' and sys_type='"& session("netwhsno") &"' ORDER BY SYS_TYPE "
											end if	
											SET RST = CONN.EXECUTE(SQL)
											WHILE NOT RST.EOF  
											%>
											<option value="<%=RST("SYS_TYPE")%>" <%if f_whsno=RST("SYS_TYPE") then %>selected<%end if%>><%=RST("SYS_VALUE")%></option>				 
											<%
											RST.MOVENEXT
											WEND 
											%>
										</SELECT>
										<%SET RST=NOTHING %>
									</TD>  
									<TD nowrap align=right height=30 >國籍<br>Quoc Tich</TD>
									<TD >
										<select name=country  class="txt8" style='width:80' onchange="GETDATA()" >
											<option value=""></option>
											<%SQL="SELECT * FROM BASICCODE WHERE FUNC='country'  ORDER BY SYS_type desc  "
											SET RST = CONN.EXECUTE(SQL)
											WHILE NOT RST.EOF  
											%>
											<option value="<%=RST("SYS_TYPE")%>" <%if country=RST("SYS_TYPE") then %>selected<%end if%>><%=RST("SYS_VALUE")%></option>				 
											<%
											RST.MOVENEXT
											WEND 
											%>
										</SELECT>
										<%SET RST=NOTHING %>
									</TD>	 
									<td>統計</td>
									<td>
										<select name="ktrajx" class="txt8" onchange="GETDATA()" >
											<option value="" <%if ktrajx="" then %>selected<%end if%>>All</option>
											<option value="Y" <%if ktrajx="Y" then %>selected<%end if%>>Y-計績效</option>
											<option value="N" <%if ktrajx="N" then %>selected<%end if%>>N-不計績效</option>
										</select>
									</td>
								</tr>
								<tr class="txt8">	
									<TD align=right>部門<br>Bo phan</TD>
										<TD>
										<SELECT NAME=F_GROUPID CLASS="txt8"  onchange="GETDATA()" >
											<option value=""></option>
											<%SQL="SELECT* FROM BASICCODE WHERE FUNC='GROUPID'  and sys_type<>'AAA'   "&_	 					  
												  "order by   sys_type    "
											  SET RST=CONN.EXECUTE(SQL)
											  WHILE NOT RST.EOF 
											%><OPTION VALUE="<%=RST("SYS_TYPE")%>" <%if rst("sys_type")=f_groupid then%>selected<%end if%>><%=RST("SYS_TYPE")%><%=RST("SYS_VALUE")%></OPTION>
											<%RST.MOVENEXT%>
											<%WEND%>
											<%set rst=nothing %>
										</SELECT>	 			
									</TD>  
									<td nowrap align=right >班別<br>Ca</td>
									<td  >
										<select name="F_shift" class=txt8 onchange="GETDATA()"> 			 		
											<option value=""></option>
											<%SQL="SELECT* FROM BASICCODE WHERE FUNC='shift'      "&_	 					  
												  "order by   len(sys_type) desc, sys_type    "
											  SET RST=CONN.EXECUTE(SQL)
											  WHILE NOT RST.EOF 
											%><OPTION VALUE="<%=RST("SYS_TYPE")%>" <%if rst("sys_type")=f_shift then%>selected<%end if%>><%=RST("SYS_VALUE")%></OPTION>
											<%RST.MOVENEXT%>
											<%WEND%>
											<%set rst=nothing %>
										</select>	
									</td>
							 
									<td nowrap align=right >工號<br>so the</td>
									<td  >
										<input name=empid1 class=inputbox size=10 value="<%=empid1%>"> 			 	
									</td>					
									<td align=center colspan=2>				
										<input type=reset  name=btm class=button value="(S)查詢K.Tra" ONCLICK="GETDATA()">				
									</td>
								</tr>	
							</table>
						</td>
					</tr>
					<tr>
						<td>
							<table class="table table-bordered table-sm bg-white text-secondary"> 		
								<tr bgcolor="#ffcc99" height=22><td colspan=10>　　『<b>v</b>』:表示計算績效(Y)</td></tr>
								<tr bgcolor="#ffffff" class="txt8">
									<td width=30 align="center" nowrap>STT</td>
									<td width=30 align="center" nowrap>計算<br>績效</td>
									<td width=50 align="center" nowrap>工號</td>
									<td width=120 align="center" nowrap>姓名</td>
									<td width=30 align="center" nowrap>國籍</td>
									<td width=70 align="center" nowrap>到職日</td>
									<td width=70 align="center" nowrap>簽合同</td>
									<td width=100 align="center" nowrap>單位</td>
									<td width=70 align="center" nowrap>離職日</td>
									<td>備註</td>
								</tr>		
								<%for x = 1 to pagerec%>
								<tr bgcolor="#ffffff" class="txt8">	
									<td align="center" ><%=x%></td>
									<td align="center" >
										<%if tmpRec(CurrentPage, x, 1)<>"" then %>
											<%if tmpRec(CurrentPage, x, 0)="Y" then %> 
												<input type="checkbox"  name="func" checked onclick="chkynchg(<%=x-1%>)">
											<%else%>
												<input type="checkbox"  name="func" onclick="chkynchg(<%=x-1%>)">
											<%end if%>
										<%else%>	
											<input type="hidden"  name="func" >
										<%end if%>
										<input type="hidden" value="<%=tmpRec(CurrentPage, x, 0)%>" name="jxYN">
										<input type="hidden" value="<%=tmpRec(CurrentPage, x, 15)%>" name="c_aid">
									</td>			
									<td align="center" ><%=tmpRec(CurrentPage, x, 1)%></td>
									<td><%=tmpRec(CurrentPage, x, 2)%><br><%=tmpRec(CurrentPage, x, 3)%></td> 
									<td align="center" ><%=tmpRec(CurrentPage, x, 4)%></td>
									<td align="center" ><%=tmpRec(CurrentPage, x, 5)%></td>
									<td align="center" ><%=tmpRec(CurrentPage, x, 6)%></td>
									<td><%=tmpRec(CurrentPage, x, 8)%><%=tmpRec(CurrentPage, x, 11)%><br><%=tmpRec(CurrentPage, x, 12)%></td> 
									<td><%=tmpRec(CurrentPage, x, 13)%></td>
									<td>
										<%if tmpRec(CurrentPage, x, 1)<>"" then %>
											<input name="jxynMemo" size=20 value="<%=tmpRec(CurrentPage, x, 14)%>" class="inputbox8">
										<%else%>	
											<input name="jxynMemo" size=20 type="hidden">
										<%end if%>	
									</td>
								</tr>
								<%next%>
							</table>
						</td>
					</tr>
					<tr>
						<td>
							<input type="hidden" name="func" value="">
							<input type="hidden" name="jxYN" value="">
							<input type="hidden" name="c_aid" value="">
							<input type="hidden" name="jxynMemo" value="">
							
							<table class="table-borderless table-sm bg-white text-secondary">
								<tr><td align="center">
								<%if session("mode")="W" then %>
									<input type="button"  name="btn" value="(Y)Confirm" class="btn btn-sm btn-danger" onclick="go()">
								<%end if%>
								<input type="button"  name="btn" value="(X)Close" class="btn btn-sm btn-outline-secondary" onclick="parent.close()"></td></tr>
							</table>
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

function chkynchg(index)
	if <%=self%>.func(index).checked=true then 
		<%=self%>.jxyn(index).value="Y"
	else
		<%=self%>.jxyn(index).value="N"
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
 	<%=self%>.action="<%=SELF%>.updateDB.asp"
	<%=self%>.target="_self"
 	<%=self%>.submit
end function   

FUNCTION GETDATA()
	<%=self%>.totalpage.value=""
	<%=self%>.action="<%=self%>.Fore.asp"
 	<%=self%>.submit()
END FUNCTION  
	

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


'_________________DATE CHECK___________________________________________________________________

function validDate(d)
	if len(d) = 8 and isnumeric(d) then
		d = left(d,4) & "/" & mid(d, 5, 2) & "/" & right(d,2)
		if isdate(d) then
			validDate = formatDate(d)
		else
			validDate = ""
		end if
	elseif len(d) >= 8 and isdate(d) then
			validDate = formatDate(d)
		else
			validDate = ""
		end if
end function

function formatDate(d)
		formatDate = Year(d) & "/" & _
		Right("00" & Month(d), 2) & "/" & _
		Right("00" & Day(d), 2)
end function
'________________________________________________________________________________________  

 
</script> 
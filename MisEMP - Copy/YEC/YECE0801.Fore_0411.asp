<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%> 
<!-- #include file = "../GetSQLServerConnection.fun" --> 
<!-- #include file="../ADOINC.inc" -->
<!-- #include file="../Include/func.inc" -->
<%
Set conn = GetSQLServerConnection()	  
self="YECE0801" 
yymm = request("YYMM") 
 
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



if yymm="" then yymm=nowmonth

	'一個月有幾天  
	cDatestr=CDate(LEFT(yymm,4)&"/"&RIGHT(yymm,2)&"/01") 
	days = DAY(cDatestr+(32-DAY(cDatestr))-DAY(cDatestr+(32-DAY(cDatestr))))   '一個月有幾天   
	'response.write days  
	
	ENDdat = (LEFT(YYMM,4)&"/"&RIGHT(YYMM,2)&"/"&DAYS) 
	calcdt = left(YYMM,4)&"/"& right(yymm,2)&"/01" 

Set conn = GetSQLServerConnection()
Set rs = Server.CreateObject("ADODB.Recordset")

gTotalPage = 1
PageRec = 20    'number of records per page
TableRec = 35    'number of fields per record 

if right(yymm,2)="01" then  
	ymstr = left(yymm,4)-1 & "12" 
else
	ymstr = left(yymm,4)& right("00"&right(yymm,2)-1,2)
end if 	

sqln="select * from salaryWP where yymm='"& ymstr &"'" 
SET RDS=CONN.EXECUTE(SQLN) 
IF RDS.EOF THEN 
	LASTData = 0 
else
	LastData = 1 
end if 
set rds=nothing 	

sqlstr="select yymm, isnull(closeflag,'') Nowflag from salaryWP where  yymm='"& yymm &"' group by yymm , isnull(closeflag,'')  "
set rds2=conn.execute(sqlstr)
if not rds2.eof then 
	closeYN=rds2("Nowflag")
end if 
set rds2=nothing		
 
transFlag = request("transFlag")  
if request("transFlag") = "Y" then  
	sqlstr="select isnull(vnTmat,0) vnTmat, isnull(c.nindat,'') nindat, isnull(c.outdate,'') outdate,  "&_
			 "isnull(d.ktaxm,0) tax, isnull(d.real_total,0) realTotal , isnull(e.jiaAh,0) jiaAh, z.* , isnull(f.exrt,0) as rate , "&_
			 "isnull(b.tien3 ,0) as wpbtien from ( "&_
		   "select * from ( "&_
		   "select   a.* from "&_
		   "(select isnull(closeflag,'') closeYN, * from salaryWP  where  yymm='"& ymstr &"' ) a  "&_
		   "left join (select *  from salaryWP where  yymm='"& yymm &"' ) b on b.empid = a.empid "&_
		   "where isnull(b.empid,'')='' "&_
		   "union all "&_
		   "select isnull(closeflag,'') closeYN, * from salaryWP where  yymm='"& yymm &"' "&_
		   " ) z  ) z "&_
		   "left join ( select bb+cv+phu+nn+kt+mt+ttkh+qc as vnTmat, * from  bemps where yymm='"& yymm &"' ) b on b.empid= z.empid "&_
			 "left join (select * from view_empfile ) c on c.empid = z.empid "&_
			 "left join (Select * from empdsalary where yymm='"& yymm &"' ) d on d.empid = z.empid  "&_
			 "left join (select sum(hhour) jiaAh, empid  from empholiday where jiatype='A' and convert(char(6), dateup,112)='"& yymm &"' group by empid ) e on e.empid = z.empid "&_
			 "left join ( select * from VYFYEXRT where    code='USD' )  f on f.yyyymm=z.yymm "&_
		   "order by len(z.whsno) desc, z.whsno, z.empid "   
else  
	sqlstr="select isnull(vnTmat,0) vnTmat, isnull(c.nindat,'') nindat, isnull(c.outdate,'') outdate,  "&_
			 "isnull(d.ktaxm,0) tax, isnull(d.real_total,0) realTotal , isnull(e.jiaAh,0) jiaAh,  a.* , isnull(f.exrt,0) as rate "&_
			 ",isnull(b.tien3 ,0) as wpbtien  from "&_
		   "(select isnull(closeflag,'') closeYN, * from salaryWP "&_
		   "where  yymm='"& yymm &"' ) a "&_
		   "left join ( select bb+cv+phu+nn+kt+mt+ttkh+qc as vnTmat, * from  bemps where yymm='"& yymm &"' ) b on b.empid= a.empid "&_
		   "left join (select * from view_empfile ) c on c.empid = a.empid "&_
			 "left join (Select* from empdsalary where yymm='"& yymm &"' ) d on d.empid = a.empid  "&_
			 "left join (select sum(hhour) jiaAh, empid  from empholiday where jiatype='A' and convert(char(6), dateup,112)='"& yymm &"' group by empid ) e on e.empid = a.empid "&_
			 "left join ( select * from VYFYEXRT where    code='USD' )  f on f.yyyymm=a.yymm "&_
			 "order by len(a.whsno) desc, a.whsno, a.empid  "    
end if 	
'response.write sqlstr 
'response.end 

clcrate="20900"
if request("TotalPage") = "" or request("TotalPage") = "0" then
	CurrentPage = 1
	rs.Open sqlstr, conn, 3, 1
	IF NOT RS.EOF THEN
		PageRec = RS.RECORDCOUNT + 5 
		rs.PageSize = PageRec
		RecordInDB = rs.RecordCount
		TotalPage = rs.PageCount
		gTotalPage = TotalPage

	END IF 
	Redim tmpRec(gTotalPage, PageRec, rs.fields.count+10)   'Array  	
	for i = 1 to TotalPage
	 for j = 1 to PageRec
		if not rs.EOF then
		
			ratex=rs("rate") 
			'ratex="20900" 
			
			tmpRec(i, j, 0) = "no"
			tmpRec(i, j, 1) = trim(rs("yymm"))
			tmpRec(i, j, 2) = trim(rs("empid"))
			tmpRec(i, j, 3) = trim(rs("country"))
			tmpRec(i, j, 4) = rs("whsno")
			tmpRec(i, j, 5) = trim(replace(rs("empname"),"v",""))
			tmpRec(i, j, 6) = rs("rzM")
			if rs("rzM")=""  or isnull(rs("rzM")) then 
				tmpRec(i, j, 6) = 0 
			end if 		
			tmpRec(i, j, 7) = rs("rzdays")
			if rs("rzdays")=""  or isnull(rs("rzdays")) then 
				tmpRec(i, j, 7) = 0 
			end if 
			tmpRec(i, j, 8) = rs("JrM")
			if rs("JrM")=""  or isnull(rs("JrM")) then 
				tmpRec(i, j, 8) = 0 
			end if 
			tmpRec(i, j, 9)	=RS("jrdays")
			if rs("jrdays")=""  or isnull(rs("jrdays")) then 
				tmpRec(i, j, 9) = 0 
			end if 
			tmpRec(i, j, 10)=RS("tnkh")
			tmpRec(i, j, 11)=RS("qita")
			tmpRec(i, j, 12)=RS("dm")
			tmpRec(i, j, 13)=RS("totAMT")
			if rs("totAMT")=""  or isnull(rs("totAMT")) then 
				tmpRec(i, j, 13) = 0 
			end if 
			
			tmpRec(i, j, 13) = cdbl(rs("bb"))+cdbl(rs("cv"))+cdbl(rs("tnkh"))-cdbl(rs("qita"))-cdbl(rs("zgm"))
			tmpRec(i, j, 14)=RS("zkM")
				if rs("zkM")=""  or isnull(rs("zkM")) then 
				tmpRec(i, j, 14) = 0 
			end if 
			tmpRec(i, j, 15)=RS("memo") 
			tmpRec(i, j, 16)=RS("bb")  '職務加給
			tmpRec(i, j, 17)=RS("cv")   '海外津貼 
			tmpRec(i, j, 18)=RS("salaryType") 
			tmpRec(i, j, 19)=RS("closeflag")
			tmpRec(i, j, 20)=RS("aid")
			tmpRec(i, j, 21)=RS("zgm")
			tmpRec(i, j, 22)=RS("vnTmat")
			tmpRec(i, j, 23)=cdbl(RS("totAMT"))+cdbl(RS("vnTmat"))
			
			tmpRec(i, j, 24)=trim(RS("nindat"))
			tmpRec(i, j, 27)=trim(RS("outdate"))
			'response.write cdate(tmpRec(i, j, 24)) & ","& ENDdat &"<BR>" 
			
			if trim(RS("nindat"))<>"" then 
				IF (tmpRec(i, j, 24))>(calcdt) THEN
					iF tmpRec(i, j, 27)="" THEN  '本月到職本月仍在職
						A1= DATEDIFF("D", CDATE(tmpRec(i, j, 24)), CDATE(ENDdat))+1
						MWORKDAYS = cdbl(A1)
						tmpRec(i, j, 25) = MWORKDAYS
						if trim(tmpRec(i, j, 15))="" then 
							tmpRec(i, j, 15) = tmpRec(i, j, 24)&"到職 ,"& tmpRec(i, j, 15)  
						end if 	
					ELSE '本月到職本月離職
						A1= DATEDIFF("D", CDATE(tmpRec(i, j, 24)), CDATE(tmpRec(i, j, 27)))
						MWORKDAYS = cdbl(A1)
						tmpRec(i, j, 25) = MWORKDAYS '**********本月工作天數**********
						
					END IF
				ELSE '舊員工
					iF tmpRec(i, j, 27)="" THEN  '仍在職						
						tmpRec(i, j, 25) = days
					ELSEiF tmpRec(i, j, 27)<=ENDdat  then  '本月離職
						A1= DATEDIFF("D", calcdt, CDATE(tmpRec(i, j, 27)))+1
						MWORKDAYS = cdbl(A1)
						tmpRec(i, j, 25) = MWORKDAYS '**********本月工作天數********** 
						if trim(tmpRec(i, j, 15))="" then 
							tmpRec(i, j, 15) = tmpRec(i, j, 27)&"離職 "
						end if 	
					else  '非本月離職	
						tmpRec(i, j, 25) = days
					END IF
					'tmpRec(i, j, 25) = days  '**********本月工作天數**********
				END IF
			else
				tmpRec(i, j, 25)=0
			end if 
		'用海外津貼扣稅 同VN算法 ,統一稅率15% elin 20150401
			'RS("cv") 
			
			totB="9000000"
			sql2="exec sp_calctax_2010 '"& cdbl(RS("cv"))*cdbl(clcrate) &"' , '"& totB &"','"& rs("empid") &"' "							
			
			set ors=conn.execute(sql2) 
			kother = ors("kother") 
			'response.write sql2 &" , kother="& kother & "usd="& cdbl(kother)/cdbl(clcrate) &"<br>" 
			ors.close : set ors=nothing 
			r = (cdbl(RS("cv"))*cdbl(clcrate)-(cdbl(totB)))
			rr=r*0.15	
			if cdbl(clcrate) = 0 then 
				taxamt=0 
			else
				taxamt= round( (((cdbl(RS("cv"))*cdbl(clcrate)-(cdbl(totB)))*(15/100.00)+ (cdbl(kother)*-1)) )/cdbl(clcrate) ,0)
			end if
			'taxper = ors("taxper")
			if rs("empid")="A0006" or rs("empid")="A0045" then taxamt = 0 
			
			'response.write rs("empid") &"-"& cdbl(RS("cv"))*cdbl(rs("rate")) &"," & r  &","& rr&","& (kother)&"==="&  taxamt &"<BR>"
			if taxamt <=0 then taxamt=0
			tmpRec(i, j, 31)	 = taxamt  
			
			'response.write rs("empid") &"-taxamt:"&  taxamt&"<br>"
			'response.write rs("empid") &"-wptax:"&  rs("wptax")&"<br>"			
			'response.write rs("empid") &"-tmprec 31:"&  rs("wptax")&"<br>" 
			'response.write rs("empid") &"-cdbl(tmpRec(i, j, 13))="&  cdbl(tmpRec(i, j, 13))&"<br>"
			tmpRec(i, j, 13) = cdbl(tmpRec(i, j, 13))-cdbl(taxamt)  
			'response.write rs("empid") &"-cdbl(tmpRec(i, j, 13))="&  cdbl(tmpRec(i, j, 13))&"<br>"			
			
		if cdbl(tmpRec(i, j, 25))< cdbl(days) and cdbl(tmpRec(i, j, 25)) > 0  then 
			if cdbl(tmpRec(i, j, 11))=0 then 
				tmpRec(i, j, 11)= round( cdbl(RS("totAMT")) - ( cdbl(RS("totAMT"))/cdbl(days)*cdbl(tmpRec(i, j, 25)) ) ,0)
				tmpRec(i, j, 13)= cdbl(RS("totAMT")) - cdbl(tmpRec(i, j, 11)) 				
			end if 	
		end if 
		tmpRec(i, j, 23)=cdbl(tmpRec(i, j, 13))+cdbl(RS("vnTmat"))
			tmpRec(i, j, 28)=rs("tax")
			tmpRec(i, j, 29)=rs("realTotal")
			tmpRec(i, j, 30)=cdbl(tmpRec(i, j, 13)) +cdbl(tmpRec(i, j, 29))
			if cdbl(rs("jiaAh")) >0  then 
					if trim(tmpRec(i, j, 15))="" then 
						tmpRec(i, j, 15) = tmpRec(i, j, 15) &",事假 "& round(cdbl(rs("jiaAh"))/8.0,1) &" 天"
					end if 	
			end if 
			
			tmpRec(i, j, 13) = cdbl(tmpRec(i, j, 13))+cdbl(rs("wpbtien"))
			tmpRec(i, j, 23) = cdbl(tmpRec(i, j, 23))+cdbl(rs("wpbtien"))
			tmpRec(i, j, 30) = cdbl(tmpRec(i, j, 30))+cdbl(rs("wpbtien"))
			tmpRec(i, j, 35)=rs("wpbtien")
			tmpRec(i, j, 36)=cdbl(tmpRec(i, j, 23))+cdbl(taxamt)
			
		
			
			rs.MoveNext
		else
			'tmpRec(i, j, 6) = 0 
			'tmpRec(i, j, 7) = 0 
			'tmpRec(i, j, 8) = 0 
			'tmpRec(i, j, 9) = 0 
			'tmpRec(i, j, 10) = 0 
			'tmpRec(i, j, 11) = 0 
			'tmpRec(i, j, 13) = 0 
			'tmpRec(i, j, 14) = 0 
			'tmpRec(i, j, 16) = 0 
			'tmpRec(i, j, 17) = 0 
			exit for
		end if 
		
		
	 next

	 if rs.EOF then
		rs.Close
		Set rs = nothing
		exit for
	 end if
	next
	'Session("YECE0801") = tmpRec
else
	TotalPage = cint(request("TotalPage"))
	'StoreToSession()
	CurrentPage = cint(request("CurrentPage"))
	RecordInDB  = REQUEST("RecordInDB")
	'tmpRec = Session("YECE0801")

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


FUNCTION FDT(D)
	Response.Write YEAR(D)&"/"&RIGHT("00"&MONTH(D),2)&"/"&RIGHT("00"&DAY(D),2)

END FUNCTION 


%>

<html> 
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="refresh">
<link rel="stylesheet" href="../Include/style.css" type="text/css">
<link rel="stylesheet" href="../Include/style2.css" type="text/css">
</head> 
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
 
function f()
	<%=self%>.YYMM.focus()	
	<%=self%>.YYMM.SELECT()
end function  
function yymmchg()
	<%=self%>.transFlag.value=""
	<%=self%>.totalpage.value="0"
	if isnumeric(<%=self%>.yymm.value)=false then 
		alert "輸入錯誤!!"
		exit function 
	end if 	
	<%=self%>.action="<%=self%>.Fore.ASP"
	<%=self%>.submit() 	
end function    

function transDatachg()	
	<%=self%>.transFlag.value="Y"
	<%=self%>.totalpage.value="0"
	<%=self%>.action="<%=self%>.Fore.ASP"
	<%=self%>.submit() 	
end  function 
 
</SCRIPT>   
<body  topmargin="5" leftmargin="5"  marginwidth="0" marginheight="0"  onkeydown="enterto()" onload='f()' >
<form name="<%=self%>" method="post" action="<%=self%>.Fore.ASP">
<INPUT TYPE=hidden NAME=TotalPage VALUE="<%=TotalPage%>">
<INPUT TYPE=hidden NAME=CurrentPage VALUE="<%=CurrentPage%>">
<INPUT TYPE=hidden NAME=RecordInDB VALUE="<%=RecordInDB%>">
<INPUT TYPE=hidden NAME=PageRec VALUE="<%=PageRec%>">
<INPUT TYPE=hidden NAME=gTotalPage VALUE="<%=gTotalPage%>"> 
<INPUT TYPE=hidden NAME=flag VALUE=""> 
<INPUT TYPE=hidden NAME=closeyn VALUE="<%=closeyn%>"> 

<input type="hidden" name=lastData value="<%=LastData%>">
<input type="hidden" name=transFlag value="Y" > 

<table width="460" border="0" cellspacing="0" cellpadding="0">
	<tr><TD>
	<img border="0" src="../image/icon.gif" align="absmiddle">
	<%=session("pgname")%></TD></tr>
</table>
<hr size=0	style='border: 1px dotted #999999;' align=left width=500>	
 
<table width=600  ><tr><td >
	<table width=700 border=0 cellspacing="0" cellpadding="0"  > 
		<tr height=30 >
			<TD   align=left width=90 >計薪年月：</TD>
			<TD   >
				<INPUT NAME=YYMM  CLASS=INPUTBOX VALUE="<%=yymm%>" SIZE=8   >
				&nbsp;
				<INPUT type="button" name="btn" value="(Y)confirm"  CLASS="button" onclick="yymmchg()" >
			</TD>	
			<TD  nowrap="nowrap" align=left   >工作天數：<%=days%>  <INPUT type=hidden NAME=workdays  CLASS=INPUTBOX VALUE="<%=days%>" size=3 >
			Rate :<%=ratex%>
			</TD> 
			<TD    >
				<%if closeyn="Y" then%><font color=red>本月已關帳,不可修改</font>
				<%else%>
					<%if lastData >0  then %>
						<input type="button" name=TransData class=button  value="轉入上月(<%=ymstr%>)資料" onclick="transDatachg()" >
					<%else%>&nbsp;
						<input type="hidden" name=TransData class=button value="轉入資料">
					<%end if%>
				<%end if%>
				 				
			</TD>	 
		</TR>
	</table>
	
	<TABLE   CLASS="txt8" BORDER=0 cellspacing="1" cellpadding="1" >
		<TR HEIGHT=25 BGCOLOR="LightGrey"   >
			<td>STT</td>
	 		<TD align=center   nowrap>刪除</TD>
	 		<TD align=center   nowrap>代號</TD>	 		
	 		<TD align=center   nowrap>國籍</TD> 		 
	 		<TD align=center  nowrap>廠別</TD>		 		
	 		<TD    >員工姓名<BR>(中,英,越)</TD> 
			<TD align=center  nowrap>幣別</TD>
			<TD align=center  nowrap>工作<br>天數</TD>		 		
	 		<!--TD align=center   nowrap>職務<br>加給</TD-->
	 		<TD align=center   nowrap>海外<BR>津貼</TD>	 		
			<TD align=center   nowrap>WP<BR>其加</TD>	 		
	 		<!--td align=center nowrap>日支<BR>USD</td> 			
	 		<TD align=center   nowrap>天數</TD>
	 		<TD align=center  nowrap>假日<BR>USD</TD>
	 		<TD align=center  nowrap>天數</TD-->	 		
	 		<td align=center  nowrap>其他<br>收入</td>  		 
	 		<td align=center  nowrap>列帳<br>暫估</td>  		 			
	 		<td align=center  nowrap>其他<br>扣除</td>  	 			 	 		
			<td align=center  nowrap>暫扣款</td>  	
			<td align=center  nowrap>合計</td>	
			<td align=center  nowrap>WP<br/>稅額</td>	 				
			<td align=center  nowrap>境內<br>薪資</td>
			<td align=center  nowrap><br>應領<br>薪資</td>
			<td align=center  nowrap width="50">薪資<br>未含稅</td>
			<!--td align=center  nowrap>境內<br>實領</td-->	 		
			<td align=center  nowrap>實領<br>金額</td>
	 		
	 		<td align=center  nowrap>備註</td>
 		</TR>
 		<%
 		TOTusd  = 0 
 		TOTvnd  = 0 
 		totbb = 0 
 		totcv = 0 
 		tottnkh = 0 
 		totqita = 0 
 		totamt = 0 
 		totzkm = 0 
 		
 		for CurrentRow = 1 to PageRec
			IF CurrentRow MOD 2 = 0 THEN
				WKCOLOR="LavenderBlush"
			ELSE
				WKCOLOR=""
			END IF 
			
			IF 	tmpRec(CurrentPage, CurrentRow, 13)<>"" AND  tmpRec(CurrentPage, CurrentRow, 12)="USD" THEN 
				TOTusd = TOTusd + CDBL(tmpRec(CurrentPage, CurrentRow, 13)) 
			END IF 	
			IF 	tmpRec(CurrentPage, CurrentRow, 13)<>"" AND  tmpRec(CurrentPage, CurrentRow, 12)="VND"  THEN 
				TOTvnd = TOTvnd + CDBL(tmpRec(CurrentPage, CurrentRow, 13)) 
			END IF  			

 		
 		%>
 		<tr bgcolor="<%=WKCOLOR%>">
 			<td align=center>
 				<%=currentRow%>
 			</td>
 			<td align=center>
 				<%if closeYN="" and ( tmpRec(CurrentPage, CurrentRow, 13)<>"" and tmpRec(CurrentPage, CurrentRow, 13)>"0" ) then %>
 					<INPUT type=checkbox  NAME=fun onclick="delchg(<%=currentrow-1%>)">
 					<INPUT type=hidden  NAME=op value="" >
 				<%else%>	
 					<INPUT type=hidden type=checkbox  NAME=fun >
 					<INPUT type=hidden  NAME=op value="" >
 				<%end if%>
 				<INPUT type=hidden  NAME=aid value="<%=tmpRec(CurrentPage, CurrentRow, 20)%>" >
 			</td>
 			<td>
 				<INPUT  NAME=empid  size=6 value="<%=tmpRec(CurrentPage, CurrentRow, 2)%>" class="INPUTBOX8" STYLE="background-color: #EDEFFF" ondblclick="getempid(<%=CurrentRow-1%>)" onchange="empidchg(<%=CurrentRow-1%>)"> 
 			</td>
 			<td>
 				<INPUT  NAME=country  size=3 value="<%=tmpRec(CurrentPage, CurrentRow, 3)%>" class="readonly8"   readonly > 
 			</td>
 			<td>
 				<INPUT  NAME=whsno  size=3 value="<%=tmpRec(CurrentPage, CurrentRow, 4)%>" class="readonly8"   readonly > 
 			</td>
 			<td>
				<%msg="到職日:"& tmpRec(CurrentPage, CurrentRow, 24) & chr(13)&"離職日:" & tmpRec(CurrentPage, CurrentRow, 27) %>
 				<INPUT  NAME=empname  size=12 value="<%=tmpRec(CurrentPage, CurrentRow, 5)%>" class="INPUTBOX8"  title="<%=msg%>"   > 
 			</td>  	 	
			<td>
 				<select name=dm class="INPUTBOX8"  >
 					<option value="" <%if tmpRec(CurrentPage, CurrentRow, 12)="" then%>selected<%end if%>>---</option>
 					<option value="USD" <%if tmpRec(CurrentPage, CurrentRow, 12)="USD" then%>selected<%end if%>>USD</option>
 					<option value="VND" <%if tmpRec(CurrentPage, CurrentRow, 12)="VND" then%>selected<%end if%>>VND</option> 					
 				</select>	
 			</td> 		
			<td>
 				<INPUT  NAME=wkdays  size=2 value="<%=tmpRec(CurrentPage, CurrentRow, 25)%>" class="readonly8"   readonly > 
 			</td>  	 			
 			<td>
				<INPUT  type="hidden" NAME=BB  size=6 value="<%=tmpRec(CurrentPage, CurrentRow, 16)%>" class="INPUTBOX8" STYLE="TEXT-ALIGN:RIGHT" onblur="bbchg(<%=currentrow-1%>)"> 
				<INPUT  NAME=CV  size=5 value="<%=tmpRec(CurrentPage, CurrentRow, 17)%>" class="INPUTBOX8" STYLE="TEXT-ALIGN:RIGHT" onblur="cvchg(<%=currentrow-1%>)"> 				
				<!--日薪USD,天數-->
 				<INPUT  type="hidden" NAME=rzM  size=3 value="<%=tmpRec(CurrentPage, CurrentRow, 6)%>" class="INPUTBOX8" STYLE="TEXT-ALIGN:RIGHT" onblur="rzmchg(<%=currentrow-1%>)"> 
				<INPUT  type="hidden" NAME=rzdays  size=3 value="<%=tmpRec(CurrentPage, CurrentRow, 7)%>" class="INPUTBOX8" STYLE="TEXT-ALIGN:RIGHT" onblur="rzMdayschg(<%=currentrow-1%>)"> 
				<!--假日日薪USD,天數-->
				<INPUT  type="hidden" NAME=JRM  size=3 value="<%=tmpRec(CurrentPage, CurrentRow, 8)%>" class="INPUTBOX8" STYLE="TEXT-ALIGN:RIGHT" onblur="JRMchg(<%=currentrow-1%>)"> 
				<INPUT  type="hidden" NAME=jrdays  size=3 value="<%=tmpRec(CurrentPage, CurrentRow, 9)%>" class="INPUTBOX8" STYLE="TEXT-ALIGN:RIGHT" onblur="jrMdayschg(<%=currentrow-1%>)"> 
 			</td> 
			<td>
 				<INPUT  NAME=wpbtien  size=5 value="<%=tmpRec(CurrentPage, CurrentRow, 35)%>" class="READONLY8" STYLE="TEXT-ALIGN:RIGHT" readonly  >
 			</td>
 			<td>
 				<INPUT  NAME=tnkh  size=5 value="<%=tmpRec(CurrentPage, CurrentRow, 10)%>" class="INPUTBOX8" STYLE="TEXT-ALIGN:RIGHT" onblur="tnkhchg(<%=currentrow-1%>)" >
 			</td>
 			<td>
 				<INPUT  NAME=zgm  size=5 value="<%=tmpRec(CurrentPage, CurrentRow, 21)%>" class="INPUTBOX8" STYLE="TEXT-ALIGN:RIGHT" onblur="zgmchg(<%=currentrow-1%>)" >
 			</td>			
 			<td>
 				<INPUT  NAME=qita  size=5 value="<%=tmpRec(CurrentPage, CurrentRow, 11)%>" class="INPUTBOX8" STYLE="TEXT-ALIGN:RIGHT" onblur="qitachg(<%=currentrow-1%>)" > 
 			</td>
			 
			<td>
 				<INPUT  NAME=zkm  size=5 value="<%=tmpRec(CurrentPage, CurrentRow, 14)%>" class="INPUTBOX8" STYLE="TEXT-ALIGN:RIGHT" > 
 			</td>
			<td>
 				<INPUT  NAME=totAMT  size=7 value="<%=tmpRec(CurrentPage, CurrentRow, 13)%>" readonly class="READONLY8" STYLE="TEXT-ALIGN:RIGHT" > 
 			</td>
			<td>
				<INPUT  NAME="kwptax"  size=5 value="<%=tmpRec(CurrentPage, CurrentRow, 31)%>" readonly class="READONLY8" STYLE="TEXT-ALIGN:RIGHT" > 
			</td>			
 			<td>
 				<INPUT  NAME=vnAMT  size=6 value="<%=tmpRec(CurrentPage, CurrentRow, 22)%>" readonly class="READONLY8" STYLE="TEXT-ALIGN:RIGHT" > 
 			</td>
 			<td>
 				<INPUT  NAME=AllAmt  size=7 value="<%=tmpRec(CurrentPage, CurrentRow, 23)%>" readonly class="READONLY8" STYLE="TEXT-ALIGN:RIGHT" > 
 			</td>
 			<td align="right">
				<%=tmpRec(CurrentPage, CurrentRow, 36)%>
 				<INPUT  type="hidden" NAME=vnTAX  size=6 value="<%=tmpRec(CurrentPage, CurrentRow, 28)%>" readonly class="READONLY8" STYLE="TEXT-ALIGN:RIGHT" > 
				<INPUT  type="hidden" NAME=vnRealAMT  size=6 value="<%=tmpRec(CurrentPage, CurrentRow, 29)%>" readonly class="READONLY8" STYLE="TEXT-ALIGN:RIGHT" > 
 			</td>

			<td>
 				<INPUT  NAME=WPVNAMT  size=6 value="<%=tmpRec(CurrentPage, CurrentRow, 30)%>" readonly class="READONLY8" STYLE="TEXT-ALIGN:RIGHT" > 
 			</td>
 			<td>
 				<INPUT  NAME=memo  size=50 value="<%=tmpRec(CurrentPage, CurrentRow, 15)%>" class="INPUTBOX8" STYLE="TEXT-ALIGN:left" > 
 			</td>	
 		</tr>	 			 			
 		<%next%>
 		<tr>
 			<td></td>
 			<td></td>
 			<td></td>
 			<td></td>
 			<td></td>
 			<td></td>
 			<td></td><!--bb-->
 			<td></td><!--CV-->
 			<td></td>
 			<td></td>
 			<td></td>
 			<td></td>
 			<td></td><!--tnkh-->
 			<td></td><!--qita-->
 			<td></td>
 			<td></td>
 			<td></td> 
 			<td></td> 
 			
 		</tr>
 		<tr bgcolor=#e4e4e4  height=25>
 			<td colspan=18 align=center >
 				Total USD : <INPUT NAME=TOTUSD CLASS=READONLY8 READONLY VALUE="<%=formatnumber(TOTUSD,0)%>" style='text-align:right'> 　　
 				Total VND : <INPUT NAME=TOTVND CLASS=READONLY8 READONLY VALUE="<%=formatnumber(TOTVND,0)%>" style='text-align:right'> 　
 			</td>
 		</tr>
 	</table>
	<br>
	<table width=600  class=font9>
		<tr >
		    <td align="left" height=40 width=60%>
		    PAGE:<%=CURRENTPAGE%> / <%=TOTALPAGE%> , COUNT:<%=RECORDINDB%>
			<% If CurrentPage > 1 Then %>
				<input type="submit" name="send" value="FIRST" class=button>
				<input type="submit" name="send" value="BACK" class=button>
			<% Else %>
				<input type="submit" name="send" value="FIRST" disabled class=button>
				<input type="submit" name="send" value="BACK" disabled class=button>
			<% End If %>
			<% If cint(CurrentPage) < cint(TotalPage) Then %>
				<input type="submit" name="send" value="NEXT" class=button>
				<input type="submit" name="send" value="END" class=button>
			<% Else %>
				<input type="submit" name="send" value="NEXT" disabled class=button>
				<input type="submit" name="send" value="END" disabled class=button>
			<% End If %>
			</td>		
			<td align=right>
				<input type=button  name=btm class=button value="<%if closeyn<>"Y" then %>確   認<%else%>已關帳<%end if%>" onclick="go()" onkeydown="go()" <%if closeyn="Y" then%>disabled<%end if%> >
				<input type=reset  name=btm class=button value="取   消">
				
				<input type=button  name=btm class=button value="關帳(CLOSE)" <%if closeyn="Y" then%>disabled<%end if%> onclick="goclose()" onkeydown="goclose()" > 
			</td>
		</tr>	
	</table>	

</td></tr></table> 

</body>
</html>


<script language=vbs>  
function delchg(index) 
	if <%=self%>.fun(index).checked=true then 
		<%=self%>.op(index).value="DEL"
	else
		<%=self%>.op(index).value=""
	end if 
end function 

function getempid(index)
	open "getempdata.asp?formName="&"<%=self%>" &"&index=" & index , "Back" 
	parent.best.cols="60%,40%"
end function 

function goclose()
	<%=self%>.flag.value="Y"
	<%=self%>.action = "<%=self%>.upd.asp"
	<%=self%>.submit()
end function 

function empidchg(index) 
	if <%=self%>.empid(index).value<>"" then 
		empid_str = Ucase(trim(<%=self%>.empid(index).value)) 
		open "<%=self%>.back.asp?func=getemp&index=" & index &"&CODESTR01=" & empid_str  , "Back" 
		
	else
		<%=self%>.empid(index).value=""	
		<%=self%>.whsno(index).value=""
		<%=self%>.country(index).value=""
		<%=self%>.empname(index).value="" 
	end if  
	parent.best.cols="100%,0%"
end  function   
      
function bbchg(index)
	if <%=self%>.bb(index).value<>"" then 
		if isnumeric(<%=self%>.bb(index).value)=false then 
			alert "請輸入數字!!"
			<%=self%>.bb(index).value=""
			<%=self%>.bb(index).focus()
			exit function 
		else
			'<%=self%>.cv(index).focus()
			calctotM(index) 			
		end if
	end if 			
end function     

function cvchg(index)
	if <%=self%>.cv(index).value<>"" then 
		if isnumeric(<%=self%>.cv(index).value)=false then 
			alert "請輸入數字!!"
			<%=self%>.cv(index).value=""
			<%=self%>.cv(index).focus()
			exit function 
		else
			'<%=self%>.rzm(index).focus()
			calctotM(index) 			
		end if
	end if 			
end function  

function cvchg2(index)
	if <%=self%>.wpbtien(index).value<>"" then 
		if isnumeric(<%=self%>.wpbtien(index).value)=false then 
			alert "請輸入數字!!"
			<%=self%>.wpbtien(index).value=""
			<%=self%>.wpbtien(index).focus()
			exit function 
		else
			'<%=self%>.rzm(index).focus()
			calctotM(index) 			
		end if
	end if 			
end function   
 

function rzmchg(index)
	if <%=self%>.rzm(index).value<>"" then 
		if isnumeric(<%=self%>.rzm(index).value)=false then 
			alert "請輸入數字!!"
			<%=self%>.rzm(index).value=""
			<%=self%>.rzm(index).focus()
			exit function 
		else
			'<%=self%>.rzdays(index).focus()
			calctotM(index) 			
		end if
	end if 			
end function

function rzmdayschg(index)
	if <%=self%>.rzdays(index).value<>"" then 
		if isnumeric(<%=self%>.rzdays(index).value)=false then 
			alert "請輸入數字!!"
			<%=self%>.rzdays(index).value=""
			<%=self%>.rzdays(index).focus()
			exit function 
		else
			calctotM(index) 			
		end if
	end if 			
end function    
function jrMchg(index)
	if <%=self%>.jrM(index).value<>"" then 
		if isnumeric(<%=self%>.jrM(index).value)=false then 
			alert "請輸入數字!!"
			<%=self%>.jrM(index).value=""
			<%=self%>.jrM(index).focus()
			exit function 
		else
			calctotM(index) 			
		end if
	end if 			
end function   

function jrMdayschg(index)
	if <%=self%>.jrdays(index).value<>"" then 
		if isnumeric(<%=self%>.jrdays(index).value)=false then 
			alert "請輸入數字!!"
			<%=self%>.jrdays(index).value=""
			<%=self%>.jrdays(index).focus()
			exit function 
		else
			calctotM(index) 			
		end if
	end if 		 	
end function    

function tnkhchg(index)
	if <%=self%>.tnkh(index).value<>"" then 
		if isnumeric(<%=self%>.tnkh(index).value)=false then 
			alert "請輸入數字!!"
			<%=self%>.tnkh(index).value=""
			<%=self%>.tnkh(index).focus()
			exit function 
		else
			'<%=self%>.qita(index).focus()
			calctotM(index) 			
		end if
	end if 			
end function  
function zgmchg(index)
	if <%=self%>.zgm(index).value<>"" then 
		if isnumeric(<%=self%>.zgm(index).value)=false then 
			alert "請輸入數字!!"
			<%=self%>.zgm(index).value=""
			<%=self%>.zgm(index).focus()
			exit function 
		else
			'<%=self%>.zgm(index).focus()
			calctotM(index) 			
		end if
	end if 			
end function 

function qitachg(index)
	if <%=self%>.qita(index).value<>"" then 
		if isnumeric(<%=self%>.qita(index).value)=false then 
			alert "請輸入數字!!"
			<%=self%>.qita(index).value=""
			<%=self%>.qita(index).focus()
			exit function 
		else
			'<%=self%>.dm(index).focus()
			calctotM(index) 			
		end if
	end if 			
end function  

function calctotM(index)
	if trim(<%=self%>.bb(index).value)<>"" and isnumeric(<%=self%>.bb(index).value)=true then 
		F_BB = cdbl(<%=self%>.bb(index).value) 
	else
		F_BB = 0	
	end if 	
	if trim(<%=self%>.cv(index).value)<>"" and isnumeric(<%=self%>.cv(index).value)=true then 
		F_CV = cdbl(<%=self%>.cv(index).value) 
	else
		F_CV = 0	
	end if 		
	if trim(<%=self%>.rzM(index).value)<>"" and isnumeric(<%=self%>.rzM(index).value)=true then 
		F_rzM = cdbl(<%=self%>.rzM(index).value) 
	else
		F_rzM = 0	
	end if 	 
	if trim(<%=self%>.rzdays(index).value)<>"" and isnumeric(<%=self%>.rzdays(index).value)=true then 
		F_rzMdays = cdbl(<%=self%>.rzdays(index).value) 
	else
		F_rzMdays = 0
	end if 	
	if trim(<%=self%>.jrM(index).value)<>"" and isnumeric(<%=self%>.jrM(index).value)=true then 
		F_jrM = cdbl(<%=self%>.jrM(index).value) 
	else
		F_jrM = 0	
	end if 	
	if trim(<%=self%>.jrdays(index).value)<>"" and isnumeric(<%=self%>.jrdays(index).value)=true then 
		F_jrMdays = cdbl(<%=self%>.jrdays(index).value) 
	else
		F_jrMdays = 0
	end if 			 
	if trim(<%=self%>.tnkh(index).value)<>"" and isnumeric(<%=self%>.tnkh(index).value)=true then 
		F_tnkh = cdbl(<%=self%>.tnkh(index).value) 
	else 	
		F_tnkh = 0
	end if 		 
	if trim(<%=self%>.zgm(index).value)<>"" and isnumeric(<%=self%>.zgm(index).value)=true then 
		F_zgm = cdbl(<%=self%>.zgm(index).value) 
	else 	
		F_zgm = 0
	end if 	
	if trim(<%=self%>.qita(index).value)<>"" and isnumeric(<%=self%>.qita(index).value)=true then 
		F_qita = cdbl(<%=self%>.qita(index).value) 
	else
		F_qita = 0	
	end if 			
 
	if trim(<%=self%>.wpbtien(index).value)<>"" and isnumeric(<%=self%>.wpbtien(index).value)=true then 
		F_wpbtien = cdbl(<%=self%>.wpbtien(index).value) 
	else
		F_wpbtien = 0	
	end if 	
	
	f_kwptax = <%=self%>.kwptax(index).value  
	if f_kwptax="" then f_kwptax = 0 
	
	
	
	
	'if <%=self%>.salaryType(index).value="1" then 
	<%=self%>.totAmt(index).value = F_BB + F_CV + (F_rzM*F_rzMdays ) + ( F_jrM * F_jrMdays ) + F_tnkh + F_ZGM +cdbl(F_wpbtien) - cdbl(f_kwptax)
	'end if 	
	vnamt = (<%=self%>.vnamt(index).value)
	if (<%=self%>.vnamt(index).value)="" then vnamt = 0  
	
	
	
	
	<%=self%>.allamt(index).value = cdbl(vnamt)+cdbl(<%=self%>.totAmt(index).value)
	
	'<%=self%>.WPVNAMT(index).value = cdbl(<%=self%>.allamt(index).value) -cdbl(<%=self%>.vnTAX(index).value)
	c_totamt_usd = 0
	c_totamt_vnd = 0 
	for zz = 1 to 20		
		if <%=self%>.totamt(zz-1).value<>"" and <%=self%>.totamt(zz-1).value>"0" then  
			if <%=self%>.dm(zz-1).value="USD" then 
				c_totamt_usd = c_totamt_usd + cdbl(<%=self%>.totamt(zz-1).value)
			elseif <%=self%>.dm(zz-1).value="VND" then 
				c_totamt_vnd = c_totamt_vnd + cdbl(<%=self%>.totamt(zz-1).value)
			end if	
		end if 
	next 
	<%=self%>.totusd.value=formatnumber(c_totamt_usd,0)
	<%=self%>.totvnd.value=formatnumber(c_totamt_vnd,0)
	
end function 


function chksts()
	if <%=self%>.chk1.checked=true then 
		<%=self%>.recalc.value="Y"
	else
		<%=self%>.recalc.value="N"
	end if 
end function 
function dataclick(a)
	if a = 1 then 		
		open "empbasic/empbasic.asp" , "_self"
	elseif a = 2 then 		
		open "empfile/empfile.asp" , "_self"
	elseif a = 3 then 		
		open "empworkHour/empwork.asp" , "_self"	
	elseif a = 4 then 		
		open "holiday/empholiday.asp" , "_self"	
	elseif a = 5 then 		
		open "AcceptCaTime/main.asp" , "_self"				
	elseif a = 6 then 		
		open "../report/main.asp" , "_self"		
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
</script> 
<%@language=vbscript codepage=65001%>
<!-- #include file = "../../getsqlserverconnection.fun" -->
<!-- #include file="../../adoinc.inc" -->
<%
session.codepage="65001"
self = "empworkb"

set conn = getsqlserverconnection()
set rs = server.createobject("adodb.recordset")
set rst = server.createobject("adodb.recordset")

gtotalpage = 1
pagerec = 10    'number of records per page
tablerec = 30    'number of fields per record


yymm = request("yymm")
'response.write yymm
if yymm="" then
	yymm = year(date())&right("00"&month(date()),2)
	'yymm="200601"
	cdatestr=date()
	days = day(cdatestr+(32-day(cdatestr))-day(cdatestr+(32-day(cdatestr))))   '一個月有幾天
else
	cdatestr=cdate(left(yymm,4)&"/"&right(yymm,2)&"/01")
	days = day(cdatestr+(32-day(cdatestr))-day(cdatestr+(32-day(cdatestr))))   '一個月有幾天
end if

nowmonth = year(date())&right("00"&month(date()),2) 
if month(date())="01" then
	calcmonth = year(date()-1)&"12"
else
	calcmonth =  year(date())&right("00"&month(date())-1,2)
end if
empid = trim(request("empid"))
empautoid = trim(request("empautoid"))

ftotalpage = request("ftotalpage")
fcurrentpage = request("fcurrentpage")
frecordindb = request("frecordindb")
'response.end 
 

'-------------------------------------------------------------------------------------- 
'response.end 

gtotalpage = 1
'pagerec = 31    'number of records per page
if yymm="" then
	pagerec = 31
else
	pagerec = days
end if
tablerec = 40    'number of fields per record

'出缺勤紀錄 --------------------------------------------------------------------------------------

'sqlstra= " sp_calcworktime '"& empid &"', '"& yymm &"' "
'response.write sqlstra 
'response.end 
'response.write 	request("totalpage")


viewid = session("netuser")  
'viewid = "lsary"  
If Request.ServerVariables("REQUEST_METHOD") = "POST"   and request("btn") ="Y" then
	'response.write request("days")
	empid=request("empid")
	yymm=request("yymm")
	for i = 1 to request("days")
		flag= request("flag")(i) 		
		workdat=replace(trim(request("workdat")(i)),"/","")
		t1=REPLACE(request("timeup")(i),":","")&"00"
		t2=REPLACE(request("timedown")(i),":","")&"00"
		tothr=request("tothour")(i)
		tmpt1 	=  request("tmpt1")(i)  '修息時間1
		tmpt2 	=  request("tmpt2")(i)  '歇息時間2
		vh1 	=  request("h1")(i)   
		vh2 	=  request("h2")(i)  
		vh3 	=  request("h3")(i)  
		vb3 	=  request("b3")(i)  
		'response.write flag
		
			'response.write  yymm &"<BR>"
			'response.write  empid &"<BR>"
			'response.write  workdat &"<BR>"
			'response.write  t1 &"<BR>"
			'response.write  t2 &"<BR>"
			'response.write  tothr_day &"<BR>"	
			'if request("timeup")(i)="" then t1="000000"
			'if request("timedown")(i)="" then t2="000000"
			if trim(request("timeup")(i))="" and trim(request("timedown")(i))="" then 
			sql="delete empworkjd where yymm='"&yymm&"' and workdat='"&workdat&"' and empid='"&empid&"' "
			else 
			sql="if not exists ( select * from empworkjd where yymm='"&yymm&"' and workdat='"&workdat&"' and empid='"&empid&"' )  "&_  
					"insert into empworkjd ( yymm, workdat, empid, timeup, timedown, toth ,  mdtm, muser ,h1,h2,h3,b3) values ( "&_
					"'"&yymm&"','"&workdat&"','"&empid&"','"&t1&"','"&t2&"','"&tothr&"',getdate(),'"&session("netuser")&"','"&vh1&"','"&vh2&"','"&vh3&"','"&vb3&"' )   "&_ 
					"else update empworkjd set timeup='"&t1&"' , timedown='"&t2&"' , toth='"&tothr&"'  , mdtm=getdate(), muser='"&session("netuser")&"' "&_
					",h1='"&vh1&"',h2='"&vh2&"',h3='"&vh3&"',b3='"&vb3&"' where yymm='"&yymm&"' and workdat='"&workdat&"' and empid='"&empid&"' " 
			
			end if 
			conn.execute(sql)
		if flag<>"" then 	
			if ( tmpt1<>"" and len(tmpt1)>1 ) or ( tmpt2<>"" and len(tmpt2)>1 ) then 
				sqlx="if not exists (select * from empwork_xx where emp_id='"&empid&"' and work_dat='"& workdat &"' ) "&_
						 "insert into  empwork_xx ( emp_id,work_dat, tmpt1, tmpt2, mdtm,muser ) values ( '"&empid&"','"& workdat &"'  "&_
						 ",'"&tmpt1&"','"&tmpt2&"',getdate(),'"& session("netuser")&"')   "&_
						 "else "&_
						 "update  empwork_xx set tmpt1='"&tmpt1&"' , tmpt2='"&tmpt2&"', mdtm=getdate(), muser='"& session("netuser") &"' "&_
						 "where  emp_id='"&empid&"' and work_dat='"& workdat &"'  "  
				conn.execute(Sqlx)		 
				'response.write sqlx &"<BR>"		 
			elseif tmpt1="" and   tmpt2="" then 
				sql="delete empwork_xx where emp_id='"&empid&"' and work_dat='"& workdat &"' " 
				conn.execute(Sql)
				'response.write sql &"<BR>"		 
			end if  
			
			
		end if 
	next 
	'response.end

end if 

uid=session("netuser")
uid="LSARY"

sql="select * from fn_empwork_jd ( '"& empid &"' ,'"&yymm&"') order by dat "
'if request("totalpage") = "" or request("totalpage") = "0" then 
'response.write sql &"<BR>" 
'response.end 
	currentpage = 1
	'response.write sql 
	'response.end
	rs.open sql, conn, 1, 3
	if not rs.eof then
		empworkid = rs("empworkid")
		gstr = rs("gstr")
		groupid=rs("groupid")
		zuno=rs("zuno")
		zstr=rs("zstr")
		job = rs("job")
		jstr =rs("jstr")
		empnam_cn = rs("empnam_cn")
		empnam_vn = rs("empnam_vn")
		tx = rs("tx")
		nindat = rs("nindat")
		outdate= rs("outdate")
		shiht=rs("shift")
		grps=rs("grps")
		country=rs("country")
		rs.pagesize = pagerec
		recordindb = days 'rs.recordcount
		totalpage = 1 'rs.pagecount
		gtotalpage = totalpage 
		'response.write rs("shift")
		rshift =  rs("shift")
	end if

	redim tmprec(gtotalpage, pagerec, tablerec)   'array

	for i = 1 to totalpage
	 for j = 1 to pagerec
		if not rs.eof then
			tmprec(i, j, 0) = "no"
			tmprec(i, j, 1) = trim(rs("dat")) 
			' if ucase(session("netuser"))="lsary"  then  
			' response.write  session("netuser") &","& tmprec(i, j, 1)&","&","&rs("status")&","&rs("timeup")&","&rs("timedown")&"<br>"				
			' end if  
			'response.write  tmprec(i, j, 1) &"<br>" 			 
			if uid="LSARY"  then  
				if rs("status")="H1" then 
					if cdbl(rs("newtoth"))>8 then 	
						tmpRec(i, j, 2) = left(RS("timeup") ,2)&":"&mid(RS("timeup"),3,2) 
						if  mid(RS("timeup"),3,2) <"30"  then 
							tmpRec(i, j, 2) = left(RS("timeup") ,2)&":3"&mid(RS("timeup"),4,1)
						end if 
						tmpRec(i, j, 3) = left(RS("timedown") ,2)&":"&mid(RS("timedown"),3,2)
						'if ( tmpRec(i, j, 2) >="06:00" and tmpRec(i, j, 2)<="07:29"  )  then tmpRec(i, j, 2)="08:00"							
						if right(tmpRec(i, j, 2),2)<>"00" and mid(tmpRec(i, j, 2),4,1)<>"0"  then 
							clct1=right("00"&left(tmpRec(i, j, 2),2)+1,2)&":00"
							'response.write "1"	& right(tmpRec(i, j, 2),2) &"," & mid(tmpRec(i, j, 2),4,1)				
						else 
							clct1=tmpRec(i, j, 2)
							'response.write "2"	
						end if 
						
						if tmpRec(i, j, 2)<>"" then 
							'response.write "xxx= "& rs("newtoth") &","& rs("dat")& ","&  tmpRec(i, j, 2) &","& tmpRec(i, j, 3)& ","&-1*(rs("newtoth")*60)  &"<BR>"							
							newT2  = dateadd("n", (rs("newtoth")*60), trim(rs("dat")) &" "&clct1 ) 							
							
							if tmpRec(i, j, 3)<>tmpRec(i, j, 2)  then 
								if mid(rs("timedown"),3,2)>30 then clcmin = right("00"&mid(rs("timedown"),3,2)-30,2) else clcmin= mid(rs("timedown"),3,2)
								tmpRec(i, j, 3) = right("00"&hour(newT2),2)&":"& clcmin ' mid(rs("timedown"),3,2) 'right("00"&Minute(newT2),2)
								'response.write  tmpRec(i, j, 3)  &"<BR>" 
								if rs("groupid")="A061" then tmpRec(i, j, 3) =  right("00"&hour(newT2)+1,2)&":"& clcmin 
							end if 	
						end if
						
					else '工時未超過8小時
						if left( RS("timeup"),4) >="0600" and left(RS("timeup"),4)<="0729" and cdbl(rs("newtoth"))=8 then 
							tmpRec(i, j, 2) ="08:00"							
							clct1=tmpRec(i, j, 2)
							'response.write "3"	
							newT2  = dateadd("n", (rs("newtoth")*60), clct1 ) 							
							if mid(rs("timedown"),3,2)>"30" then clcmin = right("00"&mid(rs("timedown"),3,2)-30,2) else clcmin= mid(rs("timedown"),3,2)
								'response.wrie  newT2 &"<BR>"								
							tmpRec(i, j, 3) = right("00"&hour(newT2),2)&":"& clcmin 'mid(RS("timedown"),3,2)'right("00"&Minute(newT2),2)
							
							'tmpRec(i, j, 3) ="16:00" 
							'tmpRec(i, j, 3) ="16:00"		 
						else		'工時>8小時					
							if RS("timeup")<>"" and RS("timeup")<>"000000" then 
								tmpRec(i, j, 2) = left(RS("timeup") ,2)&":"&mid(RS("timeup"),3,2)
								if  mid(RS("timeup"),3,2) <"30"  then 
									tmpRec(i, j, 2) = left(RS("timeup") ,2)&":3"&mid(RS("timeup"),4,1)
								end if 
							else	
								tmpRec(i, j, 2) = ""
							end if	
							if ( tmpRec(i, j, 2) >="07:30" and tmpRec(i, j, 2)<="08:30"  ) then 
								'response.write rs("dat")&tmpRec(i, j, 3)&","& "xxx1"&"<BR>"
								if tmpRec(i, j, 3)<"16:00" then tmpRec(i, j, 2)="08:00"													
							end if  
							
							if RS("timedown")<>"" and RS("timedown")<>"000000" then 
								'tmpRec(i, j, 3) = left(RS("timedown") ,2)&":"&mid(RS("timedown"),3,2)								
								if right(tmpRec(i, j, 2),2)<>"00" and mid(tmpRec(i, j, 2),4,1)<>"0"  then 
									clct1=right("00"&left(tmpRec(i, j, 2),2)+1,2)&":00"  
									'response.write "4"	
								else
									clct1=tmpRec(i, j, 2)
									'response.write "5"	
								end if 									
								'response.write "," & rs("workdat") &","& clct1&"<BR>"	
								'response.end 
								newT2  = dateadd("n", (rs("newtoth")*60), clct1 ) 							
								if mid(rs("timedown"),3,2)>"30" then clcmin = right("00"&mid(rs("timedown"),3,2)-30,2) else clcmin= mid(rs("timedown"),3,2)
								'response.wrie  newT2 &"<BR>"								
								tmpRec(i, j, 3) = right("00"&hour(newT2),2)&":"& clcmin 'mid(RS("timedown"),3,2)'right("00"&Minute(newT2),2)
								if rs("groupid")="A061" then tmpRec(i, j, 3) =  right("00"&hour(newT2)+1,2)&":"& clcmin 
								
							else	
								tmpRec(i, j, 3) = ""
							end if
						end if 	
					end if 	
				else  '假日加班不計(lsary)
					tmpRec(i, j, 2) = ""
					tmpRec(i, j, 3) = ""
				end if 
			else 
				if RS("timeup")<>"" and RS("timeup")<>"000000" then 
					tmpRec(i, j, 2) = left(RS("timeup") ,2)&":"&mid(RS("timeup"),3,2)
				else	
					tmpRec(i, j, 2) = ""
				end if	
				if RS("timedown")<>"" and RS("timedown")<>"000000" then 
					tmpRec(i, j, 3) = left(RS("timedown") ,2)&":"&mid(RS("timedown"),3,2)
				else	
					tmpRec(i, j, 3) = "" 
				end if	
			end if 
			
			'tmpRec(i, j, 3) = RS("timedown") 
			'response.write tmpRec(i, j, 3) &"<BR>"
			if uid="LSARY"  then  
				if ( rs("endjbdat")="" or rs("dat")<=rs("endjbdat") )  and rs("status")="H1" then 
					tmpRec(i, j, 4) = cdbl(RS("toth"))
				else
					if rs("status")="H1" then   'eidt 20090620 假日加班不計
						tmpRec(i, j, 4) = cdbl(RS("toth"))-(cdbl(rs("H1"))+cdbl(rs("H2"))+cdbl(rs("H3")) )  
					else
						tmpRec(i, j, 4) = 0 
					end if 		
					if tmpRec(i, j, 4)<=0 then 
						tmpRec(i, j, 4) = 0 
						tmpRec(i, j, 2) = ""
						tmpRec(i, j, 3) = "" 						
					else
						if tmpRec(i, j, 2)<>"" then 
							response.write "xxx=" & tmpRec(i, j, 2) &","& tmpRec(i, j, 3)  &"<BR>"							
							newT2  = dateadd("n", 1*(rs("h1")*60), tmpRec(i, j, 2) ) 							
							response.wrie  newT2 &"<BR>"
							if tmpRec(i, j, 3)<>tmpRec(i, j, 2)  then 
								tmpRec(i, j, 3) = right("00"&hour(newT2),2)&":"&right("00"&Minute(newT2),2)
							end if 	
						end if 	
					end if 
				end if  
				''假日加班不計
				if rs("status")="H1" then  
					'tmpRec(i, j, 3) = rs("nt1")
					tmpRec(i, j, 4) = cdbl(RS("newtoth")) 					
				else
					tmpRec(i, j, 4) = 0 
					tmpRec(i, j, 2) = ""
					tmpRec(i, j, 3) = "" 
				end if
				'response.write "Line:189 = "& tmpRec(i, j, 3)&"<BR>"	 				
			else
				tmpRec(i, j, 4) = cdbl(RS("toth"))
			end if
 
			
			
			if uid="LSARY"  then 
				if rs("endjbdat")="" or rs("dat")<=rs("endjbdat") then 
					tmpRec(i, j, 5) = rs("H1")	
					if rs("status")="H1" then 
						tmpRec(i, j, 6) = rs("H2")			 
						tmpRec(i, j, 7) = rs("H3")		 
					else
						tmpRec(i, j, 6) = 0
						tmpRec(i, j, 7) = 0
					end if 	
					'response.write "b3" &cdbl(rs("B3")) &","& cdbl(rs("H1")) &"<BR>"
					tmpRec(i, j, 8) = cdbl(rs("B3"))-cdbl(rs("H1"))							 'rs("B3")
				else
					tmpRec(i, j, 5) = 0
					tmpRec(i, j, 6) = 0
					tmpRec(i, j, 7) = 0
					'從21:00開始算夜班   ( change by elin 2009/09/15)  
					if left(tmpRec(i, j, 3),2)>="21" or left(tmpRec(i, j, 3),1)="0" then 
						if cdbl(rs("B3")) > 0 then 	
							'response.write "b3,h1" &cdbl(rs("B3")) &","& cdbl(rs("H1")) &"<BR>"
							tmpRec(i, j, 8) = cdbl(rs("B3"))  '-cdbl(rs("H1"))							
						else
							tmpRec(i, j, 8) = 0 	
						end if 	
					else	
						tmpRec(i, j, 8) = 0 
					end if	
				end if 
				tmpRec(i, j, 5) = rs("nH1")		 
				tmpRec(i, j, 6) = rs("nH2")			 
				tmpRec(i, j, 7) = rs("nH3")		 
				tmpRec(i, j, 8) = rs("nB3")				
				'response.write  j &"="& tmpRec(i, j, 8) &"<BR>"
			else
				tmpRec(i, j, 5) = rs("H1")		 
				tmpRec(i, j, 6) = rs("H2")			 
				tmpRec(i, j, 7) = rs("H3")		 
				tmpRec(i, j, 8) = rs("B3")
			end if 
			
			if rs("xb3")<>"0" then tmpRec(i, j, 8)=rs("xb3")
			'tmpRec(i, j, 8) = rs("B3")
				 
			tmpRec(i, j, 9) = rs("jiaa_h")			
			tmpRec(i, j, 10) = rs("jiab_h")			
			tmpRec(i, j, 11) = rs("jiac_h")			
			tmpRec(i, j, 12) = rs("jiad_h")			
			tmpRec(i, j, 13) = rs("jiae_h")			
			tmpRec(i, j, 14) = rs("jiaf_h")
			tmpRec(i, j, 15)= mid("日一二三四五六",weekday(tmpRec(i, j, 1)) , 1 ) 
			
			tmpRec(i, j, 21) = rs("jiag_h") 			
			tmprec(i,j,23)=rs("jiah_h")
			
			totjia = cdbl(tmpRec(i, j, 9))+cdbl(tmpRec(i, j, 10))+cdbl(tmpRec(i, j, 11))+cdbl(tmpRec(i, j, 12))+cdbl(tmpRec(i, j, 13))+cdbl(tmpRec(i, j, 14))+cdbl(tmpRec(i, j, 21))+cdbl(tmpRec(i, j, 23))
			'response.write totjia &"<BR>"
			if rs("TIMEUP")<>"000000" or  ( cdbl(rs("fgh"))+cdbl(totjia))>=8  then
				tmpRec(i, j, 16) = "readonly  "
			else
				tmpRec(i, j, 16) = "inputbox"
			end if
			if rs("TIMEDOWN")<>"000000"  or  ( cdbl(rs("fgh"))+cdbl(totjia))>=8  then
				tmpRec(i, j, 17) = "readonly "
			else
				tmpRec(i, j, 17) = "inputbox"
			end if

			tmpRec(i, j, 18)=RS("STATUS")
			
			'所有假加總			
			'totjia = cdbl(tmpRec(i, j, 9))+cdbl(tmpRec(i, j, 10))+cdbl(tmpRec(i, j, 11))+cdbl(tmpRec(i, j, 12))+cdbl(tmpRec(i, j, 13))+cdbl(tmpRec(i, j, 14))+cdbl(tmpRec(i, j, 21)) 
			tmpRec(i, j, 22) = totjia  			
			
			if tmpRec(i, j, 22)>=8 then   '請假超過或等於8小時工時為0
				tmpRec(i, j, 4) = 0
			else
				tmpRec(i, j, 4) = cdbl(tmpRec(i, j, 4))
			end if

			
			'曠職
			tmpRec(i, j, 19) = (RS("kzhour"))  '8 - cdbl(tmpRec(i, j, 4)) - tmpRec(i, j, 22)

			'忘刷
			tmpRec(i, j,20) = rs("fgcnt") 
			
			tmprec(i,j,24)=rs("lsempid")
			tmprec(i,j,25)=rs("latefor") 
			if rs("status")<>"H1" then   ''' 節假日不計 (忘刷或遲到次數) elin 20110501 
				tmprec(i,j,24)= 0
				tmprec(i,j,25)= 0
			end if 
			
			tmprec(i,j,26)= rs("tmpt1")
			tmprec(i,j,27)= rs("tmpt2")
			tmprec(i,j,28)= rs("st")
			
			if rs("st")="*" then 
				tmpRec(i, j, 2)=left(RS("timeup") ,2)&":"&mid(RS("timeup"),3,2)
				tmpRec(i, j, 3) = left(RS("timedown") ,2)&":"&mid(RS("timedown"),3,2)
				tmpRec(i, j, 4) = cdbl(RS("toth"))
			end if
			
			'response.write "Ln290:="& tmpRec(i, j, 3) &"<BR>"
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
'	session("empworkbc") = tmprec 
	
	'response.end 	
 



'--------------------------------------------------------------------------------------
function fdt(d)
if d <> "" then
	response.write year(d)&"/"&right("00"&month(d),2)&"/"&right("00"&day(d),2)
end if
end function
'--------------------------------------------------------------------------------------
sql="select * from basiccode where func='closep' and sys_type='"& yymm &"' "
set rds=conn.execute(sql)
if rds.eof then
	pcntfg = 1 '可異動
	msgstr=""
else
	pcntfg = 0 '不可異動該月出勤紀錄
	msgstr="已結算，不可異動"
end if
set rds=nothing
if pcntfg = "0" then
	inputsts="readonly"
else
	inputsts="inputbox"
end if
'---------------------------------------------------------------------------------
%>

<html>

<head>
<meta http-equiv="content-type" content="text/html; charset=utf-8">
<meta http-equiv="pragma" content="no-cache">
<meta http-equiv="refresh" >
<link rel="stylesheet" href="../../include/style.css" type="text/css">
<link rel="stylesheet" href="../../include/style2.css" type="text/css">
<script src="../../include/enter2tab.js"></script> 
</head>


<body  topmargin="5" leftmargin="5"  marginwidth="0" marginheight="0"  bgcolor="#e4e4e4"    >
<form name="<%=self%>" id="form1" method="post" action = "<%=self%>.upd.asp" >
<input type=hidden name="pcntfg" value=<%=pcntfg%>>
<input type=hidden name="uid" value=<%=session("netuser")%>>
<input type=hidden name="empautoid" value=<%=empautoid%>>
<input type=hidden name=totalpage value="<%=totalpage%>">
<input type=hidden name=currentpage value="<%=currentpage%>">
<input type=hidden name=recordindb value="<%=recordindb%>">
<input type=hidden name=ftotalpage value="<%=ftotalpage%>">
<input type=hidden name=fcurrentpage value="<%=fcurrentpage%>">
<input type=hidden name=frecordindb value="<%=frecordindb%>">
<input type=hidden name=pagerec value="<%=pagerec%>">
<!-- table width="460" border="0" cellspacing="0" cellpadding="0">
	<tr>
		<td align=center >員工差勤作業 </td>
	</tr>
</table>
<hr size=0	style='border: 1px dotted #999999;' align=left width=500 -->
<table width=500   class=font9>
	<tr>
		<td >查詢年月:</td>
		<td colspan=3>
			<select name=yymm class=font9  onchange="dchg()">
				<%for z = 1 to 24
					if   z mod 12 = 0  then 
						if z\12 = 1  then 
							yy =year(date())-((z\12))
						else
							yy =year(date())
						end if 	
						zz = 12 
					elseif z > 12 and z mod 12 <> 0  then 
						yy = year(date())  
						zz = z mod 12
					else
						zz = z 
						yy = year(date()-365)						 
					end if 	
				  yymmvalue = yy&right("00"&zz,2)
				%>
					<option value="<%=yymmvalue%>" <%if yymmvalue=yymm then %>selected<%end if%>><%=yymmvalue%></option>
				<%next%>
			</select> (模擬)
			<input type=hiddent class=readonly readonly  name=days value="<%=days%>" size=5>
			　<font color=red><%=msgstr%></font>
		</td>
	</tr>
	<tr height=30>
		<td width=60>員工編號:</td>
		<td>
			<input name=empid value="<%=empid%>" size=7 class="readonly" readonly style="height:22">
			<input name=empnam value="<%=empnam_cn&" "&empnam_vn%>" size=30 class="readonly" readonly style="height:22">
		</td>
		<td align=right>單位:</td>
		<td>
			<input name=groupidstr value="<%=gstr%>" size=7 class="readonly" readonly  style="height:22">
			<input name=zunostr value="<%=zstr%>" size=5 class="readonly" readonly style="height:22" >
			<input type=hidden name=groupid value="<%=groupid%>" size=5 >
			<input type=hidden name=zuno value="<%=zuno%>" size=5 >
		</td>
		

	</tr>
</table>
<table width=500 class=font9 >
	<tr>
		<td width=60>到職日期:</td>
		<td><input name="indat" value="<%=nindat%>" size=11 class="readonly" readonly  style="height:22"></td>

		<td>職等:</td>
		<td><input name=job value="<%=jstr%>" size=12 class="readonly" readonly  style="height:22"></td>
		<td align=right>特休(天/小時):</td>
		<td>
			<input name=tx value="<%=tx%>" size=5 class="readonly" readonly  style="height:22">
			<input name=txh value="<%=cdbl(tx)*8%>" size=5 class="readonly" readonly  style="height:22">
		</td>
	</tr>
	<tr>
		<td width=60>離職日期:</td>
		<td  ><input name="outdat" value="<%=outdate%>" size=11 class="readonly" readonly  style="height:22"></td>
		<td align=right>班別:</td>
		<td>
			<input name=shift value="<%=rshift%>" size=5 class="readonly" readonly  style="height:22">
			<input name=grps value="<%=grps%>" size=5 class="readonly" readonly  style="height:22">
		</td>
		<td align=right></td>
		<td>			
		</td>
	</tr>
</table>
<hr size=0	style='border: 1px dotted #999999;' align=left width=500>
<table width=610 class=font9 >
	<tr bgcolor=#cccccc>
		<td rowspan=2 align=center>日期</td>
		<td rowspan=2 align=center>上班</td>
		<td rowspan=2 align=center>休息<br>1</td>
		<td rowspan=2 align=center>休息<br>2</td>
		<td rowspan=2 align=center>下班</td>
		<td rowspan=2 align=center>工時</td>
		<td rowspan=2 align=center>曠職</td>
		<td rowspan=2 align=center>忘<br>刷<br>卡</td>
		<td rowspan=2 align=center>遲到</td>
		<td colspan=4 align=center>加班(單位：小時)</td>
		<td colspan=7 align=center>休假(單位：小時)</td>
	</tr>
	<tr bgcolor=#cccccc>
		<td align=center>一般(1.5)</td>
		<td align=center>休息(2)</td>
		<td align=center>假日(3)</td>
		<td align=center>夜班(0.3)</td>
		<td align=center>公假</td>
		<td align=center>年假</td>
		<td align=center>事假</td>
		<td align=center>病假</td>
		<td align=center>婚假</td>
		<td align=center>喪假</td>
		<td align=center>產假</td>
	</tr>
	<%
	sum_tothour = 0
	sum_kzhour = 0
	sum_forget = 0
	sum_h1 = 0
	sum_h2 = 0
	sum_h3 = 0
	sum_b3 = 0
	um_jiaa = 0
	sum_jiab = 0
	sum_jiac = 0
	sum_jiad = 0
	sum_jiae = 0
	sum_jiaf = 0
	sum_jiag = 0
	sum_latefor = 0

	for currentrow = 1 to pagerec
	'response.write  pagerec &"<br>"
		if tmprec(currentpage, currentrow, 18)<>"H1" then
			wkcolor = "#cccccc"
		else
			if currentrow mod 2 = 0 then
				wkcolor="lavenderblush"
			else
				wkcolor=""
			end if
		end if
		
		if  tmprec(currentpage, currentrow, 28)="*" then  
		wkcolor="yellow"
		end if 
		
		'if tmprec(currentpage, currentrow, 1) <> "" then 
		'response.write currentrow &"="&  tmprec(currentpage, currentrow, 8) &"<br/>"
	%>
	<%
		sum_tothour = sum_tothour + cdbl(tmprec(currentpage, currentrow, 4))
		sum_latefor  = sum_latefor + cdbl(tmprec(currentpage, currentrow, 31))
		sum_kzhour  = sum_kzhour + cdbl(tmprec(currentpage, currentrow, 19))
		sum_forget  = sum_forget + cdbl(tmprec(currentpage, currentrow, 20))
		sum_h1 = sum_h1 + cdbl(tmprec(currentpage, currentrow, 5))
		sum_h2 = sum_h2 + cdbl(tmprec(currentpage, currentrow, 6))
		sum_h3 = sum_h3 + cdbl(tmprec(currentpage, currentrow, 7))
		sum_b3 = sum_b3 + cdbl(tmprec(currentpage, currentrow, 8))
		sum_jiaa = sum_jiaa + cdbl(tmprec(currentpage, currentrow, 9))
		sum_jiab = sum_jiab	+ cdbl(tmprec(currentpage, currentrow, 10))
		sum_jiac = sum_jiac + cdbl(tmprec(currentpage, currentrow, 11))
		sum_jiad = sum_jiad + cdbl(tmprec(currentpage, currentrow, 12))
		sum_jiae = sum_jiae + cdbl(tmprec(currentpage, currentrow, 13))
		sum_jiaf = sum_jiaf + cdbl(tmprec(currentpage, currentrow, 14))
		sum_jiag = sum_jiag + cdbl(tmprec(currentpage, currentrow, 21))
		
		
		'response.write currentrow &"="&  tmprec(currentpage, currentrow, 8) & "," & sum_b3 &"<br/>" 
	%>
	<tr bgcolor=<%=wkcolor%>>
		<td align=center nowrap class=txt8 >
		<%if tmprec(currentpage, currentrow, 1)>=indat and ( trim(outdat)="" or ( trim(outdat)<>"" and tmprec(currentpage, currentrow, 1)<=trim(outdat)) ) then%>			
			<input name=func type=checkbox  onclick="if (this.checked){document.forms[0].flag[<%=currentrow-1%>].value = 'Y' ;} else {document.forms[0].flag[<%=currentrow-1%>].value = '' ;}"  >			
		<%else%>	
			<input name=func type=hidden >			
		<%end if%>	
		<input type="hidden" name="flag" value="" size=1 class=inputbox8 readonly >
		<a href="javascript:" onclick="showworktime(<%=currentrow-1%>)" ><font color=blue><%=tmprec(currentpage, currentrow, 1)&"("&tmprec(currentpage, currentrow, 15)&")"%></font></a>
		<input type=hidden name="workdatim" value="<%=tmprec(currentpage, currentrow, 1)&"("&tmprec(currentpage, currentrow, 15)&")"%>" class=readonly readonly  size=15 style="text-align:center;color:<%if weekday(tmprec(currentpage, currentrow, 1))=1 then %>royalblue<%else%>black<%end if%>">		
		<input type=hidden name="workdat" value="<%=tmprec(currentpage, currentrow, 1)%>"  >
		<input type=hidden name="status" size="2" value="<%=tmprec(currentpage, currentrow, 18)%>" class=inputbox8  >
		<input type=hidden name="lsempid" value="<%=tmprec(currentpage, currentrow, 24)%>"  >
		</td>
		<td align=center><input name="timeup" value="<%=tmprec(currentpage, currentrow, 2)%>" class=<%=tmprec(currentpage, currentrow, 16)%> size=6 style="text-align:center" onchange="timchg(this.name,<%=currentrow-1%>,this.value)" maxlength=5   title='<%=tmprec(currentpage, currentrow, 24)%>' ></td>
		<td align=center> <!--休息時間-->
		<input name="tmpt1" value="<%=tmprec(currentpage, currentrow, 26)%>" class="readonly" size=6 style="text-align:center"  maxlength=4   title='<%=tmprec(currentpage, currentrow, 24)%>'   >
		</td>
		<td align=center> <!--休息時間-->
		<input name="tmpt2" value="<%=tmprec(currentpage, currentrow, 27)%>" class="readonly" size=6 style="text-align:center"   maxlength=4  title='<%=tmprec(currentpage, currentrow, 24)%>'  >
		</td>
		<td align=center>
			<input name="timedown" value="<%=tmprec(currentpage, currentrow, 3)%>" class=<%=tmprec(currentpage, currentrow, 17)%> size=6 style="text-align:center" onchange="timchg(this.name,<%=currentrow-1%>,this.value)"  maxlength=5    title='<%=tmprec(currentpage, currentrow, 24)%>'>			
		</td>
		<td align=center><input name="tothour" value="<%=tmprec(currentpage, currentrow, 4)%>" class="readonly"   size=3 style="text-align:right" onchange="tothr()"  ></td>
		<td align=center><input name="kzhour" value="<%=tmprec(currentpage, currentrow, 19)%>" class="readonly"    size=3 style="text-align:right;color:<%if tmprec(currentpage, currentrow, 19)<>"0" then %>red<%else%>black<%end if%>"   ></td>
		<td align=center><input name="forget" value="<%=tmprec(currentpage, currentrow, 20)%>" class="readonly" size=3 style="text-align:right;color:<%if tmprec(currentpage, currentrow, 20)<>"0" then %>red<%else%>black<%end if%>"    ></td>
		<td align=center><input name="latefor" value="<%=tmprec(currentpage, currentrow, 25)%>" class=<%=inputsts%> size=3 style="text-align:right;color:<%if tmprec(currentpage, currentrow, 25)<>"0" then %>red<%else%>black<%end if%>"    ></td>
		<td align=center bgcolor="#fbe5ce"><input name="h1" value="<%=tmprec(currentpage, currentrow, 5)%>" class=<%=inputsts%>  size=3 style="text-align:right;color:<%if tmprec(currentpage, currentrow, 5)<>"0" then %>red<%else%>black<%end if%>" ></td>
		<td align=center bgcolor="#d5fbdf"><input name="h2" value="<%=tmprec(currentpage, currentrow, 6)%>" class=<%=inputsts%>  size=3 style="text-align:right;color:<%if tmprec(currentpage, currentrow, 6)<>"0" then %>red<%else%>black<%end if%>" ></td>
		<td align=center bgcolor="#f4dcfb"><input name="h3" value="<%=tmprec(currentpage, currentrow, 7)%>" class=<%=inputsts%>  size=3 style="text-align:right;color:<%if tmprec(currentpage, currentrow, 7)<>"0" then %>red<%else%>black<%end if%>" ></td>
		<td align=center bgcolor="#e8b5a1"><input name="b3" value="<%=tmprec(currentpage, currentrow, 8)%>" class=<%=inputsts%>  size=3 style="text-align:right;color:<%if tmprec(currentpage, currentrow, 8)<>"0" then %>red<%else%>black<%end if%>" ></td>
		<td align=center><input name="jiag" value="<%=tmprec(currentpage, currentrow, 21)%>" class="readonly" readonly  size=3 style="text-align:right;color:<%if tmprec(currentpage, currentrow, 21)<>"0" then %>red<%else%>black<%end if%>" ></td>
		<td align=center><input name="jiae" value="<%=tmprec(currentpage, currentrow, 13)%>" class="readonly" readonly  size=3 style="text-align:right;color:<%if tmprec(currentpage, currentrow, 13)<>"0" then %>red<%else%>black<%end if%>" ></td>
		<td align=center><input name="jiaa" value="<%=tmprec(currentpage, currentrow, 9)%>" class="readonly" readonly  size=3 style="text-align:right;color:<%if tmprec(currentpage, currentrow, 9)<>"0" then %>red<%else%>black<%end if%>" ></td>
		<td align=center><input name="jiab" value="<%=tmprec(currentpage, currentrow, 10)%>" class="readonly" readonly  size=3 style="text-align:right;color:<%if tmprec(currentpage, currentrow, 10)<>"0" then %>red<%else%>black<%end if%>" ></td>
		<td align=center><input name="jiac" value="<%=tmprec(currentpage, currentrow, 11)%>" class="readonly" readonly  size=3 style="text-align:right;color:<%if tmprec(currentpage, currentrow, 11)<>"0" then %>red<%else%>black<%end if%>" ></td>
		<td align=center><input name="jiad" value="<%=tmprec(currentpage, currentrow, 12)%>" class="readonly" readonly  size=3 style="text-align:right;color:<%if tmprec(currentpage, currentrow, 12)<>"0" then %>red<%else%>black<%end if%>" ></td>
		<td align=center><input name="jiaf" value="<%=tmprec(currentpage, currentrow, 14)%>" class="readonly" readonly  size=3 style="text-align:right;color:<%if tmprec(currentpage, currentrow, 14)<>"0" then %>red<%else%>black<%end if%>"  ></td>
	</tr>
	
	<%next%>
	<tr bgcolor="lavender" >
		<td align=right colspan=5 height=22>總計</td>
		<td align=right ><input name="sum_tothour" value="<%=sum_tothour%>" class=readonly readonly  size=3 style="text-align:right;color:#002ca5"></td>
		<td align=right ><input name="sum_kzhour" value="<%=sum_kzhour%>" class=readonly   size=3 style="text-align:right;color:#002ca5"></td>
		<td align=right ><input name="sum_forget" value="<%=sum_forget%>" class=readonly readonly  size=3 style="text-align:right;color:#002ca5"></td>
		<td align=right ><input name="sum_latefor" value="<%=sum_latefor%>" class=readonly readonly  size=3 style="text-align:right;color:#002ca5"></td>
		<td align=right ><input name="sum_h1" value="<%=sum_h1%>" class=readonly readonly  size=3 style="text-align:right;color:#002ca5"></td>
		<td align=right ><input name="sum_h2" value="<%=sum_h2%>" class=readonly readonly  size=3 style="text-align:right;color:#002ca5"></td>
		<td align=right ><input name="sum_h3" value="<%=sum_h3%>" class=readonly readonly  size=3 style="text-align:right;color:#002ca5"></td>
		<td align=right ><input name="sum_b3" value="<%=sum_b3%>" class=readonly readonly  size=3 style="text-align:right;color:#002ca5"></td>
		<td align=right ><input name="sum_jiag" value="<%=sum_jiag%>" class=readonly readonly  size=3 style="text-align:right;color:#800000"></td>
		<td align=right ><input name="sum_jiae" value="<%=sum_jiae%>" class=readonly readonly  size=3 style="text-align:right;color:#800000"></td>
		<td align=right ><input name="sum_jiaa" value="<%=sum_jiaa%>" class=readonly readonly  size=3 style="text-align:right;color:#800000"></td>
		<td align=right ><input name="sum_jiab" value="<%=sum_jiab%>" class=readonly readonly  size=3 style="text-align:right;color:#800000"></td>
		<td align=right ><input name="sum_jiac" value="<%=sum_jiac%>" class=readonly readonly  size=3 style="text-align:right;color:#800000"></td>
		<td align=right ><input name="sum_jiad" value="<%=sum_jiad%>" class=readonly readonly  size=3 style="text-align:right;color:#800000"></td>
		<td align=right ><input name="sum_jiaf" value="<%=sum_jiaf%>" class=readonly readonly  size=3 style="text-align:right;color:#800000"></td>
	</tr>
</table>

<table border=0 width=600 class=font9 >
<tr>
  <td align="center" height=40  >
	<input type=button name=send value="關閉此視窗(close)"  class=button onclick="vbscript:window.close()">　　
	</td>
	<td align=right>	 		
			<input type="button" name="send" value="確　定" class=button onclick="go()" <%=types%> >
			<input type="reset" name="send" value="取　消" class=button>			
			
		 		
	</td>
	<td width=120  nowrap align='center'>
	  <input type="button" name="send" value="(P)出勤明細表" class=button onclick="goprt()">			
	</td>
</tr> 
</table>

</form>
<script type='text/javascript'>
	function gom(){
	window.open ("empwork_jd.asp?empid="+"<%=empid%>"+"&yymm="+"<%=yymm%>","_blank","top=50,left=60,width=800,height=500,scrollbars=yes,resizable=yes") 
	}
	function dchg(){ 
		var yymm=document.forms[0].yymm.value ;
		window.location.href="empwork_jd.asp?yymm=" +yymm +"&empid="+"<%=empid%>"	
	}
	
	function showworktime(index){
	//alert (index);
	var m = document.forms[0];
	var empidstr = m.empid.value ;
	var workdatstr = m.workdat[index].value  ;
	window.open ("showworktime.asp?empid=" + empidstr +"&workdat=" + workdatstr  , "_blank"   , "top=100, left=100, width=500, height=400, scrollbars=yes");  
	}
	
	function goprt(){
		var m = document.forms[0];
		var yymm="<%=yymm%>" ;
		var eid="<%=empworkid%>" ;
		var country="<%=country%>" ;
		var grpid="<%=groupid%>" ;
		//alert  (yymm);
		//alert  (eid);
		//alert  (grpid);
		window.open ("http://172.22.168.33/yfyemprpt/yef/yefp03.getrpt.asp?netuser=LSARY&yymm="+yymm+"&whsno=LA&groupid="+grpid+"&empid1="+eid+"&empid2="+eid , "_new" , "top=50,left=100,width=900,height=650,resizable=yes,scrollbars=yes") ;
	}
</script> 
</body>
</html>
<script type='text/javascript'>

	function timchg(sid,index,ival){
		var m = document.forms[0]; 
		
		//alert (index);
		
		if (ival !=""){			
			if (sid=="timeup") {  
				if ( left(ival,2) >="24" || right(ival,2)>"59" ){ alert ("時間輸入錯誤!!");  
					document.getElementsByName(sid)[index].value="";
					window.setTimeout( function(){document.getElementsByName(sid)[index].focus(); }, 0);					
					return;
				} 
					
			}else if (sid=="timedown") {  
				if ( left(ival,2) >="24" || right(ival,2)>"59"  ){ alert ("時間輸入錯誤!!");  
					document.getElementsByName(sid)[index].value="";
					window.setTimeout( function(){document.getElementsByName(sid)[index].focus(); }, 0);
					return;
				}
			}
			
			document.getElementsByName(sid)[index].value=left(trim(ival),2)+":"+right(trim(ival),2);
			clchour(index);
		}
		m.flag[index].value="*"; 		
	}

	function clchour(index){
		var m = document.forms[0]; 
		var NDDT=""; var NDDD="" ;
		//alert (NDDT);
		//alert ( trim(m.timeup[index].value)) ;
		//alert ( trim(m.timedown[index].value)) ;
		if ( trim(m.timeup[index].value) !="" && trim(m.timedown[index].value) !="" ){			
			if ( trim(m.timeup[index].value) == trim(m.timedown[index].value) ) {
				NDDT= m.timeup(index).value ;  			
				NDDD= m.timedown(index).value ;  			
			}
			else {
			//	alert ("bb");
				if ( m.timeup[index].value >="05:00" && m.timeup[index].value <="06:03" ){  NDDT  = "06:00" ;} 
				else if ( m.timeup[index].value >="06:04" && m.timeup[index].value <="07:03" )  { NDDT  = "07:00" ;}
				else if ( m.timeup[index].value >="07:00" && m.timeup[index].value <="08:03" )  { NDDT  = "08:00" ;}
				else if ( m.timeup[index].value >="12:00" && m.timeup[index].value <="13:03" )  { NDDT  = "13:00" ;}
				else if ( m.timeup[index].value >="15:00" && m.timeup[index].value <="16:03" )  { NDDT  = "16:00" ;}
				else if ( m.timeup[index].value >="16:04" && m.timeup[index].value <="17:03" )  { NDDT  = "17:00" ;}
				else if ( m.timeup[index].value >="19:00" && m.timeup[index].value <="20:03" )  { NDDT  = "20:00" ;}
				else if (right (m.timeup[index].value,2) > "15" ){ NDDT  = left(m.timeup[index].value,2)+":30" ; }
				else { NDDT = m.timeup[index].value ;}
				//alert (NDDT); 
				if ( right(m.timedown[index].value ,2)<"30" ) { NDDD=left(m.timedown[index].value ,2)+":00"  ;}
				else { 	NDDD=left(m.timedown[index].value ,2)+":30" ; } 
				
			}
			//遲到
			if (m.status[index].value=="H1"  && right (NDDT,2) > "03" && right(NDDT,2)<="15"  ){ m.latefor[index].value="1"; m.latefor[index].style.color="red";}			
			
			var dd1=""; var dd2="" ;
			dd1 = m.workdat[index].value+" "+NDDT	;
			dd2 = m.workdat[index].value+" "+NDDD		;
			//alert (dd1) ;
			//alert (dd2);
			var totjia = 0 ; var toth = 0 ; 
			totjia =  eval(m.jiaa[index].value*1)+eval(m.jiab[index].value*1)+eval(m.jiac[index].value*1)+eval(m.jiad[index].value*1)+eval(m.jiae[index].value*1)+eval(m.jiaf[index].value*1)+eval(m.jiag[index].value*1) ;
			//alert (totjia)
			if ( dd1.length==16 && dd2.length ==16 ){
				if ( NDDD < NDDT ) { toth= Math.ceil(DateDiff("n",new Date(dd1),new Date(dd2))/30)*0.5+24 ; }
				else { toth = Math.ceil(DateDiff("n",new Date(dd1),new Date(dd2))/30)*0.5 ;}			
				m.tothour[index].value=toth;
			} 
			//曠職 
			var maxt = 8 ;
			if ( m.groupid.value=="A061" || m.zuno.value=="A0591" || left(NDDT,2)>"12" || (NDDT=NDDD) ) maxt = 8 ; else maxt = 9 
			if ( eval(toth*1+totjia*1) < 8 && m.status[index].value=="H1"  && m.workdat[index].value  >= m.indat.value  ){
				//alert (trim(m.outdat.value)); 				
				if ( trim(m.outdat.value) ==""  ) {
					m.kzhour[index].value = eval(maxt-(toth*1+totjia*1));
					m.kzhour[index].style.color="red" ;
				}	
				else if  ( trim(m.outdat.value) <= trim(m.workdat[index].value) ){ 
					m.kzhour[index].value = 0 ;
					m.kzhour[index].style.color="black";
				}
			}	
			else {
					m.kzhour[index].value = 0;
					m.kzhour[index].style.color="black";
			}
			
		}
		tothr();
	}
	
	function tothr(){
		var m = document.forms[0];
		var e = m.days.value ;
		//alert (e);
		//(月)總工時		
		var tothr=0		;		var f_kzhour = 0		;		var f_forget = 0		;		var f_latefor = 0		;
		var f_h1 = 0		;		var f_h2 = 0		;		var f_h3 = 0		;		var f_b3 = 0		; 		
		var arrq = new Array();
		for ( j=0; j<=7; j++ ){ arrq[j]=0;  }
		
		for ( i = 1 ; i<=e; i++){
		 	if (m.tothour[i-1].value=="") arrq[0]=0 ; else arrq[0]=eval(m.tothour[i-1].value*1) ; 		
			if (m.kzhour[i-1].value=="") arrq[1]=0 ; else arrq[1]=eval(m.kzhour[i-1].value*1) ; 
			if (m.forget[i-1].value=="") arrq[2]=0 ; else arrq[2]=eval(m.forget[i-1].value*1) ; 
			if (m.latefor[i-1].value=="") arrq[3]=0 ; else arrq[3]=eval(m.latefor[i-1].value*1) ; 
			if (m.h1[i-1].value=="") arrq[4]=0 ; else arrq[4]=eval(m.h1[i-1].value*1) ; 
			if (m.h2[i-1].value=="") arrq[5]=0 ; else arrq[5]=eval(m.h2[i-1].value*1) ; 
			if (m.h3[i-1].value=="") arrq[6]=0 ; else arrq[6]=eval(m.h3[i-1].value*1) ; 
			if (m.b3[i-1].value=="") arrq[7]=0 ; else arrq[7]=eval(m.b3[i-1].value*1) ; 
			tothr += eval(arrq[0]*1);
			f_kzhour += eval(arrq[1]*1);
			f_forget += eval(arrq[2]*1);
			f_latefor += eval(arrq[3]*1);
			f_h1 +=eval(arrq[4]*1);
			f_h2 +=eval(arrq[5]*1);
			f_h3 +=eval(arrq[6]*1);
			f_b3 +=eval(arrq[7]*1);
		}
		//alert (tothr);
		m.sum_tothour.value = tothr;
		m.sum_kzhour.value = f_kzhour;
		m.sum_forget.value = f_forget;
		m.sum_latefor.value = f_latefor;
		m.sum_h1.value = f_h1;
		m.sum_h2.value = f_h2;
		m.sum_h3.value = f_h3;
		m.sum_b3.value = f_b3;
	}
	
	function DateDiff(interval,date1,date2){
	 var long = date2.getTime() - date1.getTime(); //相差毫秒
	 switch(interval.toLowerCase()){
		case "y": return parseInt(date2.getFullYear() - date1.getFullYear());
		case "m": return parseInt((date2.getFullYear() - date1.getFullYear())*12 + (date2.getMonth()-date1.getMonth()));
		case "d": return parseInt(long/1000/60/60/24);
		case "w": return parseInt(long/1000/60/60/24/7);
		case "h": return parseInt(long/1000/60/60);
		case "n": return parseInt(long/1000/60);
		case "s": return parseInt(long/1000);
		case "l": return parseInt(long);
	 }		
	} 
	
	function left(mainStr,lngLen) { 
	 if (lngLen>0) {return mainStr.substring(0,lngLen)} 
	 else{return null} 
	 }  

	function right(mainStr,lngLen) { 
	
	 if (mainStr.length-lngLen>=0 && mainStr.length>=0 && mainStr.length-lngLen<=mainStr.length) { 
	 return mainStr.substring(mainStr.length-lngLen,mainStr.length)} 
	 else{return null} 
	 } 
	function mid(mainStr,starnum,endnum){ 
	 if (mainStr.length>=0){ 
	 return mainStr.substr(starnum,endnum) 
	 }else{return null} 
	 
	}	
	function trim(str) {
  var start = -1,
  end = str.length;
  while (str.charCodeAt(--end) < 33);
  while (str.charCodeAt(++start) < 33);
  return str.slice(start, end + 1);
	}
	
	function go() {
		document.forms[0].action="empwork_jd.asp?btn=Y" ;
		document.forms[0].submit();
	}
</script> 



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

sql="select x.status, convert(char(10),x.dat,111) as dat,  b.empnam_cn, b.empnam_vn, a.empid, isnull(a.workdat,convert(char(8),x.dat,112)) workdat , "&_ 
		"isnull(a.timeup,'') timeup, isnull(timedown,'') timedown, isnull(a.toth,0) toth, isnull(a.h1,0) h1, isnull(a.h2,0) h2, isnull(a.h3,0) h3 , isnull(a.b3,0) b3, "&_
		"isnull(nt1,'') nt1 ,isnull(a.newtoth,0) newtoth, isnull(a.nh1,0) nh1, isnull(a.nh2,0) nh2, isnull(a.nh3,0) nh3 , isnull(a.nb3,0) nb3, "&_
		"isnull(a.kzhour,0) kzhour, isnull(latefor,0) latefor, "&_
		"b.empid as empworkid, isnull(b.tx,0) tx,  b.groupid, b.gstr, b.zuno, b.zstr, b.job, b.jstr, b.country, b.nindat, b.outdate  , b.shift, "&_
		"isnull(ja.hhour,0) jiaa_h , isnull(jb.hhour,0) jiab_h , isnull(jc.hhour,0) jiac_h , isnull(jd.hhour,0) jiad_h ,  "&_
		"isnull(je.hhour,0) jiae_h , isnull(jf.hhour,0) jiaf_h , isnull(jg.hhour,0) jiag_h , isnull(jh.hhour,0) jiah_h , "&_
		"isnull(fg.fgcnt,0) fgcnt, isnull(fg.fgh,0) fgh , isnull(fgt1,'') fgt1, isnull(fgt2,'') fgt2, isnull(lsempid,'') lsempid, isnull(c.endjbdat,'') endjbdat  "&_
		",isnull(xi.tmpt1,'')  tmpt1, isnull(xi.tmpt2,'')  tmpt2 from "&_
		"(select convert(char(6), dat, 112) as yymm , * from ydbmcale where convert(char(6), dat, 112)='"& yymm &"' ) x  "&_		
		"left join (select * from empwork where empid='"& empid &"' and yymm='"& yymm &"' ) a on a.yymm = x.yymm and a.workdat = convert(char(8),x.dat,112) "&_
		"left join (select * from view_empfile ) b on b.empid = a.empid "&_ 
		"left join (select convert(char(8),dateup,112) jiadat ,  empid, sum(hhour) hhour   from empholiday  where jiatype='a' and empid='"& empid &"' group by empid, convert(char(8),dateup,112)   ) ja on ja.jiadat = convert(char(8),x.dat,112) "&_
		"left join (select convert(char(8),dateup,112) jiadat ,  empid, sum(hhour) hhour   from empholiday  where jiatype='b' and empid='"& empid &"' group by empid, convert(char(8),dateup,112)  ) jb on jb.jiadat = convert(char(8),x.dat,112) "&_
		"left join (select convert(char(8),dateup,112) jiadat ,  empid, sum(hhour) hhour   from empholiday  where jiatype='c' and empid='"& empid &"' group by empid, convert(char(8),dateup,112)  ) jc on jc.jiadat = convert(char(8),x.dat,112) "&_
		"left join (select convert(char(8),dateup,112) jiadat ,  empid, sum(hhour) hhour   from empholiday  where jiatype='d' and empid='"& empid &"' group by empid, convert(char(8),dateup,112)  ) jd on jd.jiadat = convert(char(8),x.dat,112) "&_
		"left join (select convert(char(8),dateup,112) jiadat ,  empid, sum(hhour) hhour   from empholiday  where jiatype='e' and empid='"& empid &"' group by empid, convert(char(8),dateup,112)  ) je on je.jiadat = convert(char(8),x.dat,112) "&_
		"left join (select convert(char(8),dateup,112) jiadat ,  empid, sum(hhour) hhour   from empholiday  where jiatype='f' and empid='"& empid &"' group by empid, convert(char(8),dateup,112)  ) jf on jf.jiadat = convert(char(8),x.dat,112) "&_
		"left join (select convert(char(8),dateup,112) jiadat ,  empid, sum(hhour) hhour   from empholiday  where jiatype='g' and empid='"& empid &"' group by empid, convert(char(8),dateup,112)  ) jg on jg.jiadat = convert(char(8),x.dat,112) "&_
		"left join (select convert(char(8),dateup,112) jiadat ,  empid, sum(hhour) hhour   from empholiday  where jiatype='h' and empid='"& empid &"' group by empid, convert(char(8),dateup,112)  ) jh on jh.jiadat = convert(char(8),x.dat,112) "&_  
		"left join (select empid,ltrim(rtrim(isnull(lsempid,''))) lsempid, convert(char(8), dat, 112) fgdat , min(timeup) fgt1, max(timedown) fgt2,  "&_
		"sum(toth) as fgh,  sum( case when ltrim(rtrim(lsempid)) ='' then 1 else 0 end  ) fgcnt  from   empforget   where isnull(status,'')<>'d' "&_
		"group by  empid, convert(char(8), dat, 112),ltrim(rtrim(isnull(lsempid,'')))  ) fg on fg.empid = a.empid and fg.fgdat = convert(char(8), x.dat, 112) "&_
		"left join ( select  * from empjbtim ) c on c.empid = a.empid and c.yymm = a.yymm  "&_
		"left join ( select  * from empwork_xx ) xi on xi.emp_id = a.empid and xi.work_dat = a.workdat "&_
		"order by x.dat " 
		'response.write sql &"<br>"
'		response.end  
sql="select * from fn_empwork ( '"& empid &"' ,'"&yymm&"') order by dat "
if request("totalpage") = "" or request("totalpage") = "0" then
	currentpage = 1
	'response.write sqlstra
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
			if ucase(session("netuser"))="LSARY"  then  
				if rs("status")="h1" then 
					if cdbl(rs("newtoth"))>8 then 	
						tmprec(i, j, 2) = left(rs("timeup") ,2)&":"&mid(rs("timeup"),3,2) 
						if  mid(rs("timeup"),3,2) <"30"  then 
							tmprec(i, j, 2) = left(rs("timeup") ,2)&":3"&mid(rs("timeup"),4,1)
						end if 
						tmprec(i, j, 3) = left(rs("timedown") ,2)&":"&mid(rs("timedown"),3,2)
						'if ( tmprec(i, j, 2) >="06:00" and tmprec(i, j, 2)<="07:29"  )  then tmprec(i, j, 2)="08:00"							
						if right(tmprec(i, j, 2),2)<>"00" and mid(tmprec(i, j, 2),4,1)<>"0"  then 
							clct1=right("00"&left(tmprec(i, j, 2),2)+1,2)&":00"
							'response.write "1"	& right(tmprec(i, j, 2),2) &"," & mid(tmprec(i, j, 2),4,1)				
						else 
							clct1=tmprec(i, j, 2)
							'response.write "2"	
						end if 
						
						if tmprec(i, j, 2)<>"" then 
							'response.write "xxx= "& rs("newtoth") &","& rs("dat")& ","&  tmprec(i, j, 2) &","& tmprec(i, j, 3)& ","&-1*(rs("newtoth")*60)  &"<br>"							
							newt2  = dateadd("n", (rs("newtoth")*60), trim(rs("dat")) &" "&clct1 ) 							
							
							if tmprec(i, j, 3)<>tmprec(i, j, 2)  then 
								if mid(rs("timedown"),3,2)>30 then clcmin = right("00"&mid(rs("timedown"),3,2)-30,2) else clcmin= mid(rs("timedown"),3,2)
								tmprec(i, j, 3) = right("00"&hour(newt2),2)&":"& clcmin ' mid(rs("timedown"),3,2) 'right("00"&minute(newt2),2)
								'response.write  tmprec(i, j, 3)  &"<br>" 
								if rs("groupid")="a061" then tmprec(i, j, 3) =  right("00"&hour(newt2)+1,2)&":"& clcmin 
							end if 	
						end if
						
					else '工時未超過8小時
						if left( rs("timeup"),4) >="0600" and left(rs("timeup"),4)<="0729" and cdbl(rs("newtoth"))=8 then 
							tmprec(i, j, 2) ="08:00"							
							clct1=tmprec(i, j, 2)
							'response.write "3"	
							newt2  = dateadd("n", (rs("newtoth")*60), clct1 ) 							
							if mid(rs("timedown"),3,2)>"30" then clcmin = right("00"&mid(rs("timedown"),3,2)-30,2) else clcmin= mid(rs("timedown"),3,2)
								'response.wrie  newt2 &"<br>"								
							tmprec(i, j, 3) = right("00"&hour(newt2),2)&":"& clcmin 'mid(rs("timedown"),3,2)'right("00"&minute(newt2),2)
							
							'tmprec(i, j, 3) ="16:00" 
							'tmprec(i, j, 3) ="16:00"		 
						else							
							if rs("timeup")<>"" and rs("timeup")<>"000000" then 
								tmprec(i, j, 2) = left(rs("timeup") ,2)&":"&mid(rs("timeup"),3,2)
								if  mid(rs("timeup"),3,2) <"30"  then 
									tmprec(i, j, 2) = left(rs("timeup") ,2)&":3"&mid(rs("timeup"),4,1)
								end if 
							else	
								tmprec(i, j, 2) = ""
							end if	
							if ( tmprec(i, j, 2) >="07:30" and tmprec(i, j, 2)<="08:30"  ) then 
								'response.write rs("dat")&tmprec(i, j, 3)&","& "xxx1"&"<br>"
								if tmprec(i, j, 3)<"16:00" then tmprec(i, j, 2)="08:00"													
							end if  							
							if rs("timedown")<>"" and rs("timedown")<>"000000" then 
								'tmprec(i, j, 3) = left(rs("timedown") ,2)&":"&mid(rs("timedown"),3,2)								
								if right(tmprec(i, j, 2),2)<>"00" and mid(tmprec(i, j, 2),4,1)<>"0"  then 
									clct1=right("00"&left(tmprec(i, j, 2),2)+1,2)&":00"  
									'response.write "4"	
								else
									clct1=tmprec(i, j, 2)
									'response.write "5"	
								end if 									
								'response.write clct1&"<br>"	
								newt2  = dateadd("n", (rs("newtoth")*60), clct1 ) 							
								if mid(rs("timedown"),3,2)>"30" then clcmin = right("00"&mid(rs("timedown"),3,2)-30,2) else clcmin= mid(rs("timedown"),3,2)
								'response.wrie  newt2 &"<br>"								
								tmprec(i, j, 3) = right("00"&hour(newt2),2)&":"& clcmin 'mid(rs("timedown"),3,2)'right("00"&minute(newt2),2)
								if rs("groupid")="a061" then tmprec(i, j, 3) =  right("00"&hour(newt2)+1,2)&":"& clcmin 
								
							else	
								tmprec(i, j, 3) = ""
							end if
						end if 	
					end if 	
				else  '假日加班不計(lsary)
					tmprec(i, j, 2) = ""
					tmprec(i, j, 3) = ""
				end if 
			else  'real login 
				if rs("timeup")<>"" and rs("timeup")<>"000000" then 
					tmprec(i, j, 2) = left(rs("timeup") ,2)&":"&mid(rs("timeup"),3,2)
				else	
					tmprec(i, j, 2) = ""
				end if	
				if rs("timedown")<>"" and rs("timedown")<>"000000" then 
					tmprec(i, j, 3) = left(rs("timedown") ,2)&":"&mid(rs("timedown"),3,2)
				else	
					tmprec(i, j, 3) = "" 
				end if	
			end if 	
			'tmprec(i, j, 3) = rs("timedown") 
			'response.write tmprec(i, j, 3) &"<br>"
			if ucase(session("netuser"))="LSARY"  then  
				if ( rs("endjbdat")="" or rs("dat")<=rs("endjbdat") )  and rs("status")="h1" then 
					tmprec(i, j, 4) = cdbl(rs("toth"))
				else
					if rs("status")="h1" then   'eidt 20090620 假日加班不計
						tmprec(i, j, 4) = cdbl(rs("toth"))-(cdbl(rs("h1"))+cdbl(rs("h2"))+cdbl(rs("h3")) )  
					else
						tmprec(i, j, 4) = 0 
					end if 		
					if tmprec(i, j, 4)<=0 then 
						tmprec(i, j, 4) = 0 
						tmprec(i, j, 2) = ""
						tmprec(i, j, 3) = "" 						
					else
						if tmprec(i, j, 2)<>"" then 
							response.write "xxx=" & tmprec(i, j, 2) &","& tmprec(i, j, 3)  &"<br>"							
							newt2  = dateadd("n", 1*(rs("h1")*60), tmprec(i, j, 2) ) 							
							response.wrie  newt2 &"<br>"
							if tmprec(i, j, 3)<>tmprec(i, j, 2)  then 
								tmprec(i, j, 3) = right("00"&hour(newt2),2)&":"&right("00"&minute(newt2),2)
							end if 	
						end if 	
					end if 
				end if  
				''假日加班不計
				if rs("status")="h1" then  
					'tmprec(i, j, 3) = rs("nt1")
					tmprec(i, j, 4) = cdbl(rs("newtoth")) 					
				else
					tmprec(i, j, 4) = 0 
					tmprec(i, j, 2) = ""
					tmprec(i, j, 3) = "" 
				end if
				'response.write "line:189 = "& tmprec(i, j, 3)&"<br>"	 				
			else
				tmprec(i, j, 4) = cdbl(rs("toth"))
			end if 	
			
			
			if viewid="LSARY"  then 
				if rs("endjbdat")="" or rs("dat")<=rs("endjbdat") then 
					tmprec(i, j, 5) = rs("h1")	
					if rs("status")="h1" then 
						tmprec(i, j, 6) = rs("h2")			 
						tmprec(i, j, 7) = rs("h3")		 
					else
						tmprec(i, j, 6) = 0
						tmprec(i, j, 7) = 0
					end if 	
					'response.write "b3" &cdbl(rs("b3")) &","& cdbl(rs("h1")) &"<br>"
					tmprec(i, j, 8) = cdbl(rs("b3"))-cdbl(rs("h1"))							 'rs("b3")
				else
					tmprec(i, j, 5) = 0
					tmprec(i, j, 6) = 0
					tmprec(i, j, 7) = 0
					'從21:00開始算夜班   ( change by elin 2009/09/15)  
					if left(tmprec(i, j, 3),2)>="21" or left(tmprec(i, j, 3),1)="0" then 
						if cdbl(rs("b3")) > 0 then 	
							'response.write "b3,h1" &cdbl(rs("b3")) &","& cdbl(rs("h1")) &"<br>"
							tmprec(i, j, 8) = cdbl(rs("b3"))-cdbl(rs("h1"))							
						else
							tmprec(i, j, 8) = 0 	
						end if 	
					else	
						tmprec(i, j, 8) = 0 
					end if	
				end if 
				tmprec(i, j, 5) = rs("nh1")		 
				tmprec(i, j, 6) = rs("nh2")			 
				tmprec(i, j, 7) = rs("nh3")		 
				tmprec(i, j, 8) = rs("nb3")				
			else
				tmprec(i, j, 5) = rs("h1")		 
				tmprec(i, j, 6) = rs("h2")			 
				tmprec(i, j, 7) = rs("h3")		 
				tmprec(i, j, 8) = rs("b3")
			end if 
				 
			tmprec(i, j, 9) = rs("jiaa_h")			
			tmprec(i, j, 10) = rs("jiab_h")			
			tmprec(i, j, 11) = rs("jiac_h")			
			tmprec(i, j, 12) = rs("jiad_h")			
			tmprec(i, j, 13) = rs("jiae_h")			
			tmprec(i, j, 14) = rs("jiaf_h")
			tmprec(i, j, 15)= mid("日一二三四五六",weekday(tmprec(i, j, 1)) , 1 ) 
			
			tmprec(i, j, 21) = rs("jiag_h") 			
			tmprec(i,j,23)=rs("jiah_h")
			
			totjia = cdbl(tmprec(i, j, 9))+cdbl(tmprec(i, j, 10))+cdbl(tmprec(i, j, 11))+cdbl(tmprec(i, j, 12))+cdbl(tmprec(i, j, 13))+cdbl(tmprec(i, j, 14))+cdbl(tmprec(i, j, 21))+cdbl(tmprec(i, j, 23))
			'response.write totjia &"<br>"
			if rs("timeup")<>"000000" or  ( cdbl(rs("fgh"))+cdbl(totjia))>=8  then
				tmprec(i, j, 16) = "readonly  "
			else
				tmprec(i, j, 16) = "inputbox"
			end if
			if rs("timedown")<>"000000"  or  ( cdbl(rs("fgh"))+cdbl(totjia))>=8  then
				tmprec(i, j, 17) = "readonly "
			else
				tmprec(i, j, 17) = "inputbox"
			end if

			tmprec(i, j, 18)=rs("status")
			
			'所有假加總			
			'totjia = cdbl(tmprec(i, j, 9))+cdbl(tmprec(i, j, 10))+cdbl(tmprec(i, j, 11))+cdbl(tmprec(i, j, 12))+cdbl(tmprec(i, j, 13))+cdbl(tmprec(i, j, 14))+cdbl(tmprec(i, j, 21)) 
			tmprec(i, j, 22) = totjia  			
			
			if tmprec(i, j, 22)>=8 then   '請假超過或等於8小時工時為0
				tmprec(i, j, 4) = 0
			else
				tmprec(i, j, 4) = cdbl(tmprec(i, j, 4))
			end if

			'曠職
			tmprec(i, j, 19) = (rs("kzhour"))  '8 - cdbl(tmprec(i, j, 4)) - tmprec(i, j, 22)

			'忘刷
			tmprec(i, j,20) = rs("fgcnt") 
			
			tmprec(i,j,24)=rs("lsempid")
			tmprec(i,j,25)=rs("latefor") 
			if rs("status")<>"h1" then   ''' 節假日不計 (忘刷或遲到次數) elin 20110501 
				tmprec(i,j,24)= 0
				tmprec(i,j,25)= 0
			end if 
			
			tmprec(i,j,26)= rs("tmpt1")
			tmprec(i,j,27)= rs("tmpt2")
			
			'response.write "ln290:="& tmprec(i, j, 3) &"<br>"
			rs.movenext
		else
			exit for
		end if
	 next

	 if rs.eof then
		rs.close
		set rs = nothing
		exit for
	 end if
	next
	session("empworkbc") = tmprec 
	
	'response.end 	
else
	totalpage = cint(request("totalpage"))
	'storetosession()
	currentpage = cint(request("currentpage"))
	recordindb  = request("recordindb")
	tmprec = session("empworkbc")

	select case request("send")
	     case "first"
		      currentpage = 1
	     case "back"
		      if cint(currentpage) <> 1 then
			     currentpage = currentpage - 1
		      end if
	     case "next"
		      if cint(currentpage) <= cint(totalpage) then
			     currentpage = currentpage + 1
		      end if
	     case "end"
		      currentpage = totalpage
	     case else
		      currentpage = 1
	end select
end if



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


<body  topmargin="5" leftmargin="5"  marginwidth="0" marginheight="0"     >
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
			</select>
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
		if tmprec(currentpage, currentrow, 18)<>"h1" then
			wkcolor = "#cccccc"
		else
			if currentrow mod 2 = 0 then
				wkcolor="lavenderblush"
			else
				wkcolor=""
			end if
		end if
		
		'if tmprec(currentpage, currentrow, 1) <> "" then
	%>
	<tr bgcolor=<%=wkcolor%>>
		<td align=center nowrap class=txt8 >
		<%if tmprec(currentpage, currentrow, 1)>=indat and ( trim(outdat)="" or ( trim(outdat)<>"" and tmprec(currentpage, currentrow, 1)<=trim(outdat)) ) then%>			
			<input name=func type=checkbox  onclick='funcchg(<%=currentrow-1%>)'   >
			<input type=hidden name="flag" value="" size=1 class=inputbox8 readonly >
		<%else%>	
			<input name=func type=hidden >
			<input type=hidden name="flag" value="" size=1 class=inputbox8  readonly >
		<%end if%>	
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
		<td align=center><input name="tothour" value="<%=tmprec(currentpage, currentrow, 4)%>" class="readonly"   size=3 style="text-align:right"  ></td>
		<td align=center><input name="kzhour" value="<%=tmprec(currentpage, currentrow, 19)%>" class="readonly"    size=3 style="text-align:right;color:<%if tmprec(currentpage, currentrow, 19)<>"0" then %>red<%else%>black<%end if%>"   ></td>
		<td align=center><input name="forget" value="<%=tmprec(currentpage, currentrow, 20)%>" class="readonly" size=3 style="text-align:right;color:<%if tmprec(currentpage, currentrow, 20)<>"0" then %>red<%else%>black<%end if%>"    ></td>
		<td align=center><input name="latefor" value="<%=tmprec(currentpage, currentrow, 25)%>" class=<%=inputsts%> size=3 style="text-align:right;color:<%if tmprec(currentpage, currentrow, 25)<>"0" then %>red<%else%>black<%end if%>"    ></td>
		<td align=center bgcolor="#fbe5ce"><input name="h1" value="<%=tmprec(currentpage, currentrow, 5)%>" class=<%=inputsts%>  size=3 style="text-align:right;color:<%if tmprec(currentpage, currentrow, 5)<>"0" then %>red<%else%>black<%end if%>" onblur="tothn(<%=currentrow-1%>)" ></td>
		<td align=center bgcolor="#d5fbdf"><input name="h2" value="<%=tmprec(currentpage, currentrow, 6)%>" class=<%=inputsts%>  size=3 style="text-align:right;color:<%if tmprec(currentpage, currentrow, 6)<>"0" then %>red<%else%>black<%end if%>" onblur="tothn(<%=currentrow-1%>)"></td>
		<td align=center bgcolor="#f4dcfb"><input name="h3" value="<%=tmprec(currentpage, currentrow, 7)%>" class=<%=inputsts%>  size=3 style="text-align:right;color:<%if tmprec(currentpage, currentrow, 7)<>"0" then %>red<%else%>black<%end if%>" onblur="tothn(<%=currentrow-1%>)"></td>
		<td align=center bgcolor="#e8b5a1"><input name="b3" value="<%=tmprec(currentpage, currentrow, 8)%>" class=<%=inputsts%>  size=3 style="text-align:right;color:<%if tmprec(currentpage, currentrow, 8)<>"0" then %>red<%else%>black<%end if%>" onblur="tothn(<%=currentrow-1%>)"></td>
		<td align=center><input name="jiag" value="<%=tmprec(currentpage, currentrow, 21)%>" class="readonly" readonly  size=3 style="text-align:right;color:<%if tmprec(currentpage, currentrow, 21)<>"0" then %>red<%else%>black<%end if%>" ></td>
		<td align=center><input name="jiae" value="<%=tmprec(currentpage, currentrow, 13)%>" class="readonly" readonly  size=3 style="text-align:right;color:<%if tmprec(currentpage, currentrow, 13)<>"0" then %>red<%else%>black<%end if%>" ></td>
		<td align=center><input name="jiaa" value="<%=tmprec(currentpage, currentrow, 9)%>" class="readonly" readonly  size=3 style="text-align:right;color:<%if tmprec(currentpage, currentrow, 9)<>"0" then %>red<%else%>black<%end if%>" ></td>
		<td align=center><input name="jiab" value="<%=tmprec(currentpage, currentrow, 10)%>" class="readonly" readonly  size=3 style="text-align:right;color:<%if tmprec(currentpage, currentrow, 10)<>"0" then %>red<%else%>black<%end if%>" ></td>
		<td align=center><input name="jiac" value="<%=tmprec(currentpage, currentrow, 11)%>" class="readonly" readonly  size=3 style="text-align:right;color:<%if tmprec(currentpage, currentrow, 11)<>"0" then %>red<%else%>black<%end if%>" ></td>
		<td align=center><input name="jiad" value="<%=tmprec(currentpage, currentrow, 12)%>" class="readonly" readonly  size=3 style="text-align:right;color:<%if tmprec(currentpage, currentrow, 12)<>"0" then %>red<%else%>black<%end if%>" ></td>
		<td align=center><input name="jiaf" value="<%=tmprec(currentpage, currentrow, 14)%>" class="readonly" readonly  size=3 style="text-align:right;color:<%if tmprec(currentpage, currentrow, 14)<>"0" then %>red<%else%>black<%end if%>"  ></td>
	</tr>
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
	%>
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
		<%if session("netuser")="lsary" then %>
			<%if yymm>=nowmonth then %>
				<input type="button" name="send" value="確　定" class=button onclick="go()" >
				<input type="reset" name="send" value="取　消" class=button>			
			<%end if%>
		<%else%>	
			<%
			if yymm>=nowmonth then 
					types=""	
				else
					if day(date())>2 and session("rights")>"0" then 
						types="disabled"
					else
						if session("netuser")="l0197" or session("rights")="0" then 
							types=""
						else	
							types="disabled"
						end if	
					end if 						
				end if 
			%>			
			<input type="button" name="send" value="確　定" class=button onclick="go()"  >
			<input type="reset" name="send" value="取　消" class=button>
			
		<%end if%> 		
	</td>
	<td width=120  nowrap align='center'><input type="reset" name="send" value="(m)模 擬" onclick="gom()" class="button"></td>
</tr> 
</table>

</form>
<script type='text/javascript'>
	function gom(){
		location.href="empwork_jd.asp?empid="+"<%=empid%>"+"&yymm="+"<%=yymm%>" ;
		//window.open ("empwork_jd.asp?empid="+"<%=empid%>"+"&yymm="+"<%=yymm%>","_blank","top=50,left=60,width=800,height=500,scrollbars=yes,resizable=yes") 
	}
	function dchg(){ 
		var yymm=document.forms[0].yymm.value ;
		window.location.href="empwork_a_new.asp?yymm=" +yymm +"&empid="+"<%=empid%>"	
	}
	
	function showworktime(index){
	//alert (index);
	var m = document.forms[0];
	var empidstr = m.empid.value ;
	var workdatstr = m.workdat[index].value  ;
	window.open ("showworktime.asp?empid=" + empidstr +"&workdat=" + workdatstr  , "_blank"   , "top=100, left=100, width=500, height=400, scrollbars=yes");  
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
			m.flag[index].value="*";
			document.getElementsByName(sid)[index].value=left(trim(ival),2)+":"+right(trim(ival),2);
			clchour(index);
		}
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

			//加班時數			
			if (m.status[index].value=="H1" ) {
				if ( m.groupid.value=="A061" || m.zuno.value=="A0591" ) {
					if ( eval(toth*1) > 8) {
						m.h1[index].value=eval(toth*1-8);
						m.h1[index].style.color="red";
					}	
					else{
						m.h1[index].value=0 ;
						m.h1[index].style.color="black";
					}	
				}else {  //not a061 and a0591 
					//alert (m.groupid.value) ; 
					//alert (m.zuno.value ); 
					if ( eval(toth*1)  >= 9 ) {
						if ( left(NDDT,2)>="19" &&  left(NDDT,2)<="20" ) {
							m.h1[index].value=eval(toth*1-8)  ;
							m.h1[index].style.color="red";
						} else {
							m.h1[index].value=eval(toth*1-9)  ;
							m.h1[index].style.color="red";
						}	
					}	else {
						m.h1[index].value=0 ;
						m.h1[index].style.color="black";
					}
				}
				m.h2[index].value=0 ;
				m.h3[index].value=0 ;
			} else {  //節假日加班
				if (m.status[index].value=="H2" ) { 
					if ( m.groupid.value=="A061" || m.zuno.value=="A0591" ) 
						m.h2[index].value=eval(toth*1)  ; 						
					else 
						if ( eval(toth*1)>8 ) m.h2[index].value=eval(toth*1-1)  ;  else m.h2[index].value=eval(toth*1)  ; 
						
					m.h2[index].style.color="red";
					m.h1[index].value=0 ;
					m.h3[index].value=0 ;
				}	
				else if (m.status[index].value=="H3" )  {
					if ( m.groupid.value=="A061" || m.zuno.value=="A0591" ) 
						m.h3[index].value=eval(toth*1)  ; 
					else 
						if ( eval(toth*1)>8 ) m.h3[index].value=eval(toth*1-1)  ; else m.h3[index].value=eval(toth*1)  ;
					
					m.h3[index].style.color="red";					
					m.h2[index].value=0 ;
					m.h1[index].value=0 ;
				}
			}

			//	夜班			
			var endb3 =""; endb3 = dd2; var b3str="";
			if ( left(m.timedown[index].value,2)>="21" || left(m.timedown[index].value,1)=="0"  || (left(m.timedown[index].value,2)>="00" && left(m.timedown[index].value,2)<="06" ) )			
			{  
				if ( (left(m.timeup[index].value,2)>="21" || left(m.timeup[index].value,1)>="0" ) &&  left(m.timeup[index].value,2)<="06" )				
				{ b3str = dd1;	} else { b3str = m.workdat[index].value+" 21:00";}				
				if ( m.timeup[index].value > m.timedown[index].value ) 
				{ m.b3[index].value = Math.ceil(DateDiff("n",new Date(b3str),new Date(dd2))/30)*0.5+24 ; } 
				else if ( m.timeup[index].value == m.timedown[index].value ){
					m.b3[index].value = 0 
				}
				else {
					m.b3[index].value = Math.ceil(DateDiff("n",new Date(b3str),new Date(dd2))/30)*0.5 ;
				}
			//	alert  ( m.timedown[index].value ) ;
			} 
			//alert ( b3str );
			//DateDiff("n",new Date(b3str),new Date(dd2)) );
			
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
</script> 



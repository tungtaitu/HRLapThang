<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<!-- #include file = "../../GetSQLServerConnection.fun" -->
<!-- #include file="../../ADOINC.inc" -->
<%
SESSION.CODEPAGE="65001"
SELF = "empworkb"

Set conn = GetSQLServerConnection()
Set rs = Server.CreateObject("ADODB.Recordset")
Set RST = Server.CreateObject("ADODB.Recordset")

gTotalPage = 1
PageRec = 10    'number of records per page
TableRec = 30    'number of fields per record


YYMM = REQUEST("YYMM")
'response.write yymm
IF YYMM="" THEN
	YYMM = year(date())&right("00"&month(date()),2)
	'YYMM="200601"
	cDatestr=date()
	days = DAY(cDatestr+(32-DAY(cDatestr))-DAY(cDatestr+(32-DAY(cDatestr))))   '一個月有幾天
ELSE
	cDatestr=CDate(LEFT(YYMM,4)&"/"&RIGHT(YYMM,2)&"/01")
	days = DAY(cDatestr+(32-DAY(cDatestr))-DAY(cDatestr+(32-DAY(cDatestr))))   '一個月有幾天
END IF

nowmonth = year(date())&right("00"&month(date()),2) 
if month(date())="01" then
	calcmonth = year(date()-1)&"12"
else
	calcmonth =  year(date())&right("00"&month(date())-1,2)
end if
EMPID = TRIM(REQUEST("EMPID"))
empautoid = TRIM(REQUEST("empautoid"))

Ftotalpage = request("Ftotalpage")
Fcurrentpage = request("Fcurrentpage")
FRecordInDB = request("FRecordInDB")
'RESPONSE.END 
 

'-------------------------------------------------------------------------------------- 
'response.end 

gTotalPage = 1
'PageRec = 31    'number of records per page
if yymm="" then
	PageRec = 31
else
	PageRec = days
end if
TableRec = 40    'number of fields per record

'出缺勤紀錄 --------------------------------------------------------------------------------------

'SQLSTRA= " SP_CALCWORKTIME '"& EMPID &"', '"& YYMM &"' "
'response.write sqlstra 
'response.end 
'response.write 	request("TotalPage")

'response.write session("rights")
viewid = session("netuser")  
'viewid = "LSARY"  

sql="select x.status, convert(char(10),x.dat,111) as dat,  b.empnam_cn, b.empnam_vn, a.empid, isnull(a.workdat,convert(char(8),x.dat,112)) workdat , "&_ 
		"isnull(a.timeup,'') timeup, isnull(timedown,'') timedown, isnull(a.toth,0) toth, isnull(a.h1,0) h1, isnull(a.h2,0) h2, isnull(a.h3,0) h3 , isnull(a.b3,0) b3, "&_
		"isnull(nt1,'') nt1 ,isnull(a.newtoth,0) newtoth, isnull(a.nh1,0) nh1, isnull(a.nh2,0) nh2, isnull(a.nh3,0) nh3 , isnull(a.nb3,0) nb3, "&_
		"isnull(a.kzhour,0) kzhour, isnull(latefor,0) latefor, "&_
		"b.empid as empworkid, isnull(b.tx,0) tx,  b.groupid, b.gstr, b.zuno, b.zstr, b.job, b.jstr, b.country, b.nindat, b.outdate  , b.shift, "&_
		"isnull(ja.hhour,0) jiaA_h , isnull(jb.hhour,0) jiaB_h , isnull(jC.hhour,0) jiaC_h , isnull(jd.hhour,0) jiaD_h ,  "&_
		"isnull(jE.hhour,0) jiaE_h , isnull(jf.hhour,0) jiaF_h , isnull(jg.hhour,0) jiaG_h , isnull(jh.hhour,0) jiaH_h , "&_
		"isnull(fg.fgcnt,0) fgcnt, isnull(fg.fgh,0) fgh , isnull(fgt1,'') fgt1, isnull(fgt2,'') fgt2, isnull(lsempid,'') lsempid, isnull(c.endjbdat,'') endjbdat  "&_
		",isnull(xi.tmpt1,'')  tmpt1, isnull(xi.tmpt2,'')  tmpt2 from "&_
		"(select convert(char(6), dat, 112) as yymm , * from ydbmcale where convert(char(6), dat, 112)='"& yymm &"' ) x  "&_		
		"left join (select * from empwork where empid='"& empid &"' and yymm='"& yymm &"' ) a on a.yymm = x.yymm and a.workdat = convert(char(8),x.dat,112) "&_
		"left join (select * from view_empfile ) b on b.empid = a.empid "&_ 
		"left join (select convert(char(8),dateup,112) jiaDat ,  empid, sum(hhour) hhour   from empholiday  where jiatype='A' and empid='"& empid &"' group by empid, convert(char(8),dateup,112)   ) ja on ja.jiadat = convert(char(8),x.dat,112) "&_
		"left join (select convert(char(8),dateup,112) jiaDat ,  empid, sum(hhour) hhour   from empholiday  where jiatype='B' and empid='"& empid &"' group by empid, convert(char(8),dateup,112)  ) jb on jb.jiadat = convert(char(8),x.dat,112) "&_
		"left join (select convert(char(8),dateup,112) jiaDat ,  empid, sum(hhour) hhour   from empholiday  where jiatype='C' and empid='"& empid &"' group by empid, convert(char(8),dateup,112)  ) jc on jc.jiadat = convert(char(8),x.dat,112) "&_
		"left join (select convert(char(8),dateup,112) jiaDat ,  empid, sum(hhour) hhour   from empholiday  where jiatype='D' and empid='"& empid &"' group by empid, convert(char(8),dateup,112)  ) jd on jd.jiadat = convert(char(8),x.dat,112) "&_
		"left join (select convert(char(8),dateup,112) jiaDat ,  empid, sum(hhour) hhour   from empholiday  where jiatype='E' and empid='"& empid &"' group by empid, convert(char(8),dateup,112)  ) je on je.jiadat = convert(char(8),x.dat,112) "&_
		"left join (select convert(char(8),dateup,112) jiaDat ,  empid, sum(hhour) hhour   from empholiday  where jiatype='F' and empid='"& empid &"' group by empid, convert(char(8),dateup,112)  ) jf on jf.jiadat = convert(char(8),x.dat,112) "&_
		"left join (select convert(char(8),dateup,112) jiaDat ,  empid, sum(hhour) hhour   from empholiday  where jiatype='G' and empid='"& empid &"' group by empid, convert(char(8),dateup,112)  ) jg on jg.jiadat = convert(char(8),x.dat,112) "&_
		"left join (select convert(char(8),dateup,112) jiaDat ,  empid, sum(hhour) hhour   from empholiday  where jiatype='H' and empid='"& empid &"' group by empid, convert(char(8),dateup,112)  ) jh on jh.jiadat = convert(char(8),x.dat,112) "&_  
		"left join (select empid,ltrim(rtrim(isnull(lsempid,''))) lsempid, convert(char(8), dat, 112) fgdat , min(timeup) fgT1, max(timedown) fgT2,  "&_
		"sum(toth) as fgH,  sum( case when ltrim(rtrim(lsempid)) ='' then 1 else 0 end  ) fgcnt  from   empforget   where isnull(status,'')<>'D' "&_
		"group by  empid, convert(char(8), dat, 112),ltrim(rtrim(isnull(lsempid,'')))  ) FG on fg.empid = a.empid and fg.fgdat = convert(char(8), x.dat, 112) "&_
		"left join ( select  * from empjbTim ) c on c.empid = a.empid and c.yymm = a.yymm  "&_
		"left join ( select  * from empwork_xx ) xi on xi.emp_id = a.empid and xi.work_dat = a.workdat "&_
		"order by x.dat " 
		'response.write sql &"<br>"
'		response.end  
sql="select * from fn_empwork ( '"& empid &"' ,'"&yymm&"') order by dat "
if request("TotalPage") = "" or request("TotalPage") = "0" then
	CurrentPage = 1
	'RESPONSE.WRITE sql
	'RESPONSE.END
	rs.Open sql, conn, 1, 3
	IF NOT RS.EOF THEN
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
		shiht =  rs("shift")
		rs.PageSize = PageRec
		RecordInDB = days 'rs.RecordCount
		TotalPage = 1 'rs.PageCount
		gTotalPage = TotalPage
	END IF

	Redim tmpRec(gTotalPage, PageRec, TableRec)   'Array
	
	for i = 1 to TotalPage
	 for j = 1 to PageRec
		if not rs.EOF then
			tmpRec(i, j, 0) = "no"
			tmpRec(i, j, 1) = trim(rs("dat"))
				
			' if Ucase(session("netuser"))="LSARY"  then  
			' response.write  session("netuser") &","& tmpRec(i, j, 1)&","&","&rs("status")&","&RS("timeup")&","&RS("timedown")&"<BR>"				
			' end if  
			if Ucase(session("netuser"))="LSARY"  then  
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
						else							
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
								'response.write clct1&"<BR>"	
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
			if ucase(session("netuser"))="LSARY"  then  
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
			
			if viewid="LSARY"  then 
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
					if cdbl(rs("x_b3"))<>0 then  tmpRec(i, j, 8) = rs("x_b3")
				else
					tmpRec(i, j, 5) = 0
					tmpRec(i, j, 6) = 0
					tmpRec(i, j, 7) = 0
					'從21:00開始算夜班   ( change by elin 2009/09/15)  
					if left(tmpRec(i, j, 3),2)>="21" or left(tmpRec(i, j, 3),1)="0" then 
						if cdbl(rs("B3")) > 0 then 	
							'response.write "b3,h1" &cdbl(rs("B3")) &","& cdbl(rs("H1")) &"<BR>"
							tmpRec(i, j, 8) = cdbl(rs("B3"))-cdbl(rs("H1"))							
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
			else
				tmpRec(i, j, 5) = rs("H1")		 
				tmpRec(i, j, 6) = rs("H2")			 
				tmpRec(i, j, 7) = rs("H3")		 
				tmpRec(i, j, 8) = rs("B3")
			end if 
				 
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
			
			if ucase(session("netuser"))="LSARY"  then   
				if rs("st")="*" then 
					tmpRec(i, j, 2)=left(RS("x_times") ,2)&":"&mid(RS("x_times"),3,2)
					tmpRec(i, j, 3) = left(RS("x_timew") ,2)&":"&mid(RS("x_timew"),3,2)
					tmpRec(i, j, 4) = cdbl(RS("x_toth"))				
				end if 
			end if 
			
			tmprec(i,j,28)= rs("fgt1")
			tmprec(i,j,29)= rs("fgt2")
			tmprec(i,j,30)= rs("fgh")
			'tmprec(i,j,31)= rs("fgempid")
			'tmprec(i,j,32)= rs("cab3")
			'response.write "Ln290:="& tmpRec(i, j, 3) &"<BR>"
			'response.write "a="&ucase(session("netuser"))
			'RESPONSE.END
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
	Session("empworkbC") = tmpRec 
	
	'response.end 	
else
	TotalPage = cint(request("TotalPage"))
	'StoreToSession()
	CurrentPage = cint(request("CurrentPage"))
	RecordInDB  = REQUEST("RecordInDB")
	tmpRec = Session("empworkbC")

	Select case request("send")
	     Case "FIRST"
		      CurrentPage = 1
	     Case "BACK"
		      if cint(CurrentPage) <> 1 then
			     CurrentPage = CurrentPage - 1
		      end if
	     Case "NEXT"
		      if cint(CurrentPage) <= cint(TotalPage) then
			     CurrentPage = CurrentPage + 1
		      end if
	     Case "END"
		      CurrentPage = TotalPage
	     Case Else
		      CurrentPage = 1
	end Select
end if



'--------------------------------------------------------------------------------------
FUNCTION FDT(D)
IF D <> "" THEN
	Response.Write YEAR(D)&"/"&RIGHT("00"&MONTH(D),2)&"/"&RIGHT("00"&DAY(D),2)
END IF
END FUNCTION
'--------------------------------------------------------------------------------------
SQL="SELECT * FROM BASICCODE WHERE FUNC='CLOSEP' AND SYS_TYPE='"& YYMM &"' "
SET RDS=CONN.EXECUTE(SQL)
IF RDS.EOF THEN
	PCNTFG = 1 '可異動
	MSGSTR=""
ELSE
	PCNTFG = 0 '不可異動該月出勤紀錄
	MSGSTR="已結算，不可異動"
END IF
SET RDS=NOTHING
IF PCNTFG = "0" THEN
	INPUTSTS="READONLY"
ELSE
	INPUTSTS="INPUTBOX"
END IF
'---------------------------------------------------------------------------------
%>

<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta HTTP-EQUIV="refresh" >
<link rel="stylesheet" href="../../Include/style.css" type="text/css">
<link rel="stylesheet" href="../../Include/style2.css" type="text/css">
<SCRIPT   LANGUAGE=vbscript>
 
function enterto()
	if window.event.keyCode = 13 then window.event.keyCode =9
end function

function f()
	<%=self%>.TIMEUP(0).SELECT()
end function

function colschg(index)
	thiscols = document.activeElement.name
	if window.event.keyCode = 38 then
		IF INDEX<>0 THEN
			document.all(thiscols)(index-1).SELECT()
		END IF
	end if
	if window.event.keyCode = 40 then
		document.all(thiscols)(index+1).SELECT()
	end if 
end function 

</SCRIPT>
</head>
<body  topmargin="5" leftmargin="5"  marginwidth="0" marginheight="0"  onkeydown="enterto()"   >
<form name="<%=self%>"  method="post" action = "<%=self%>.upd.asp" >
<INPUT TYPE=HIDDEN NAME="PCNTFG" VALUE=<%=PCNTFG%>>
<INPUT TYPE=HIDDEN NAME="UID" VALUE=<%=SESSION("NETUSER")%>>
<INPUT TYPE=HIDDEN NAME="empautoid" VALUE=<%=empautoid%>>
<INPUT TYPE=hidden NAME=TotalPage VALUE="<%=TotalPage%>">
<INPUT TYPE=hidden NAME=CurrentPage VALUE="<%=CurrentPage%>">
<INPUT TYPE=hidden NAME=RecordInDB VALUE="<%=RecordInDB%>">
<INPUT TYPE=hidden NAME=FTotalPage VALUE="<%=FTotalPage%>">
<INPUT TYPE=hidden NAME=FCurrentPage VALUE="<%=FCurrentPage%>">
<INPUT TYPE=hidden NAME=FRecordInDB VALUE="<%=FRecordInDB%>">
<INPUT TYPE=hidden NAME=PageRec VALUE="<%=PageRec%>">
 
<table   class=font9>
	<TR>
		<td nowrap>查詢年月<br>Ngày tìm kiếm:</td>
		<td COLSPAN=3>
			<select name=yymm class=font9  onchange="dchg()">
				<%for z = 1 to 24
					if   z mod 12 = 0  then 
						if Z\12 = 1  then 
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
			<input type=hiddenT class=readonly readonly  name=days value="<%=days%>" size=5>
			　<FONT COLOR=RED><%=MSGSTR%></FONT>
		</td>
	</TR>
	<tr height=30>
		<td  nowrap >員工編號<br>Mã số nhân viên:</td>
		<td>
			<input name=empid value="<%=EMPID%>" size=7 class="readonly" readonly style="height:22">
			<input name=empnam value="<%=empnam_cn&" "&empnam_vn%>" size=30 class="readonly" readonly style="height:22">
		</td>
		<td align=right>單位<br>Đơn vị:</td>
		<td>
			<input name=groupidstr value="<%=gstr%>" size=7 class="readonly" readonly  style="height:22">
			<input name=zunostr value="<%=zstr%>" size=5 class="readonly" readonly style="height:22" >
			<input TYPE=HIDDEN name=groupid value="<%=groupid%>" size=5 >
			<input TYPE=HIDDEN name=zuno value="<%=zuno%>" size=5 >
		</td>
		

	</tr>
</table>
<table  class=font9 >
	<tr>
		<td nowrap>到職日期<br>Thời gian Tx:</td>
		<td><input name=indat value="<%=nindat%>" size=11 class="readonly" readonly  style="height:22"></td>

		<td>職等:<br>Cấp</td>
		<td><input name=job value="<%=jstr%>" size=12 class="readonly" readonly  style="height:22"></td>
		<td align=right>特休(天/小時)<br>Nghỉ đặc biệt (Ngày/giờ):</td>
		<td>
			<input name=TX value="<%=tx%>" size=5 class="readonly" readonly  style="height:22">
			<input name=TXH value="<%=cdbl(tx)*8%>" size=5 class="readonly" readonly  style="height:22">
		</td>
	</tr>
	<tr>
		<td width=60>離職日期<br> Ngày nghỉ việc:</td>
		<td  ><input name=outdat value="<%=outdate%>" size=11 class="readonly" readonly  style="height:22"></td>
		<td align=right>班別<br> Ca:</td>
		<td>
			<input name=shift value="<%=shift%>" size=5 class="readonly" readonly  style="height:22">
			<input name=grps value="<%=grps%>" size=5 class="readonly" readonly  style="height:22">
		</td>
		<td align=right></td>
		<td>			
		</td>
	</tr>
</table>
<hr size=0	style='border: 1px dotted #999999;' align=left width=500>
<TABLE WIDTH=610 CLASS=FONT9 >
	<TR BGCOLOR=#CCCCCC>
		<TD ROWSPAN=2 ALIGN=CENTER>日期<br>Ngày</TD>
		<TD ROWSPAN=2 ALIGN=CENTER>上班<br>lên ca</TD>
		<TD ROWSPAN=2 ALIGN=CENTER>休息<br>Nghỉ ngơi<br>1</TD>
		<TD ROWSPAN=2 ALIGN=CENTER>休息<br>Nghỉ ngơi<br>2</TD>
		<TD ROWSPAN=2 ALIGN=CENTER>下班<br>Xuống ca</TD>
		<TD ROWSPAN=2 ALIGN=CENTER>工時<br>Công giờ</TD>
		<TD ROWSPAN=2 ALIGN=CENTER>曠職<br>Vắng</TD>
		<TD ROWSPAN=2 ALIGN=CENTER>忘<br>刷<br>卡<br>Quên Bấm Thẻ</TD>
		<TD ROWSPAN=2 ALIGN=CENTER>遲到<br>Đi trễ</TD>
		<TD COLSPAN=4 ALIGN=CENTER>加班(單位：小時)<br>Tăng ca (Đơn vị :giờ)</TD>
		<TD COLSPAN=7 ALIGN=CENTER>休假(單位：小時)<br>Nghỉ phép (Đơn vị :giờ)</TD>
		<TD COLSPAN=4 ALIGN=CENTER>忘刷<br>Quên bấm thẻ</TD>
	</TR>
	<TR BGCOLOR=#CCCCCC>
		<TD ALIGN=CENTER>一般(1.5)<br>Thường</TD>
		<TD ALIGN=CENTER>休息(2)<br>Nghỉ</TD>
		<TD ALIGN=CENTER>假日(3)<br>Lễ</TD>
		<TD ALIGN=CENTER>夜班(0.3)<br>Ca đêm</TD>
		<TD ALIGN=CENTER>公假<br>Phép công</TD>
		<TD ALIGN=CENTER>年假<br>Phép năm</TD>
		<TD ALIGN=CENTER>事假<br>Việc riêng</TD>
		<TD ALIGN=CENTER>病假<br>Phép bệnh</TD>
		<TD ALIGN=CENTER>婚假<br>Phép cưới</TD>
		<TD ALIGN=CENTER>喪假<br>Phép tang</TD>
		<TD ALIGN=CENTER>產假<br>Phép thai sản</TD>
		<TD ALIGN=CENTER>FGT1</TD>
		<TD ALIGN=CENTER>FGT2</TD>
		<TD ALIGN=CENTER>FGHr</TD>
		<TD ALIGN=CENTER>CA(dem)</TD>
		
	</TR>
	<%
	sum_TOTHOUR = 0
	sum_KZhour = 0
	sum_Forget = 0
	sum_H1 = 0
	sum_H2 = 0
	sum_H3 = 0
	sum_B3 = 0
	um_JIAA = 0
	sum_JIAB = 0
	sum_JIAC = 0
	sum_JIAD = 0
	sum_JIAE = 0
	sum_JIAF = 0
	sum_JIAG = 0
	sum_LATEFOR = 0

	for CurrentRow = 1 to PageRec
	'response.write  PageRec &"<BR>"
		IF tmpRec(CurrentPage, CurrentRow, 18)<>"H1" THEN
			WKCOLOR = "#cccccc"
		ELSE
			IF CurrentRow MOD 2 = 0 THEN
				WKCOLOR="LavenderBlush"
			ELSE
				WKCOLOR=""
			END IF
		END IF
		
		'if tmpRec(CurrentPage, CurrentRow, 1) <> "" then
	%>
	<TR BGCOLOR=<%=WKCOLOR%>>
		<TD ALIGN=CENTER NOWRAP class=txt8 >
		<%if tmpRec(CurrentPage, CurrentRow, 1)>=INDAT and ( trim(outdat)="" or ( trim(outdat)<>"" and tmpRec(CurrentPage, CurrentRow, 1)<=trim(outdat)) ) then%>			
			<input name=func type=checkbox style="display:nonex"  onclick='funcchg(<%=CurrentRow-1%>)'   >
			<input type=hidden name=flag value="" size=1 class=inputbox8 readonly >
		<%else%>	
			<input name=func type=hidden >
			<input type=hidden name=flag value="" size=1 class=inputbox8  readonly >
		<%end if%>	
		<a href="vbscript:showWorkTime(<%=currentrow-1%>)" ><font color=blue><%=tmpRec(CurrentPage, CurrentRow, 1)&"("&tmpRec(CurrentPage, CurrentRow, 15)&")"%></font></a>
		<INPUT type=hidden NAME=WORKDATIM VALUE="<%=tmpRec(CurrentPage, CurrentRow, 1)&"("&tmpRec(CurrentPage, CurrentRow, 15)&")"%>" CLASS=READONLY READONLY  SIZE=15 STYLE="TEXT-ALIGN:CENTER;color:<%if weekday(tmpRec(CurrentPage, CurrentRow, 1))=1 then %>RoyalBlue<%else%>black<%end if%>">		
		<INPUT TYPE=HIDDEN NAME=WORKDAT VALUE="<%=tmpRec(CurrentPage, CurrentRow, 1)%>"  >
		<INPUT TYPE=HIDDEN NAME=STATUS VALUE="<%=tmpRec(CurrentPage, CurrentRow, 18)%>"  >
		<INPUT TYPE=HIDDEN NAME=lsempid VALUE="<%=tmpRec(CurrentPage, CurrentRow, 24)%>"  >
		</TD>
		<TD ALIGN=CENTER><INPUT NAME=TIMEUP VALUE="<%=tmpRec(CurrentPage, CurrentRow, 2)%>" CLASS=<%=tmpRec(CurrentPage, CurrentRow, 16)%> SIZE=6 STYLE="TEXT-ALIGN:CENTER" ONBLUR="TIMEUP_chg(<%=CurrentRow-1%>)" maxlength=5  onkeydown="colschg(<%=CurrentRow-1%>)"  title='<%=tmpRec(CurrentPage, CurrentRow, 24)%>'></TD>
		<TD ALIGN=CENTER> <!--休息時間-->
		<INPUT NAME=tmpt1 VALUE="<%=tmpRec(CurrentPage, CurrentRow, 26)%>" CLASS="readonly" SIZE=6 STYLE="TEXT-ALIGN:CENTER"  maxlength=4   title='<%=tmpRec(CurrentPage, CurrentRow, 24)%>'   >
		</td>
		<TD ALIGN=CENTER> <!--休息時間-->
		<INPUT NAME=tmpt2 VALUE="<%=tmpRec(CurrentPage, CurrentRow, 27)%>" CLASS="readonly" SIZE=6 STYLE="TEXT-ALIGN:CENTER"   maxlength=4  title='<%=tmpRec(CurrentPage, CurrentRow, 24)%>'  >
		</td>
		<TD ALIGN=CENTER>
			<INPUT NAME=TIMEDOWN VALUE="<%=tmpRec(CurrentPage, CurrentRow, 3)%>" CLASS=<%=tmpRec(CurrentPage, CurrentRow, 17)%> SIZE=6 STYLE="TEXT-ALIGN:CENTER" ONBLUR="TIMEDOWN_chg(<%=CurrentRow-1%>)" maxlength=5  onkeydown="colschg(<%=CurrentRow-1%>)" title='<%=tmpRec(CurrentPage, CurrentRow, 24)%>'>			
		</TD>
		<TD ALIGN=CENTER><INPUT NAME=TOTHOUR VALUE="<%=tmpRec(CurrentPage, CurrentRow, 4)%>" CLASS="readonly"   SIZE=3 STYLE="TEXT-ALIGN:RIGHT"  onkeydown="colschg(<%=CurrentRow-1%>)" ></TD>
		<TD ALIGN=CENTER><INPUT NAME=KZhour VALUE="<%=tmpRec(CurrentPage, CurrentRow, 19)%>" CLASS="readonly"    SIZE=3 STYLE="TEXT-ALIGN:RIGHT;color:<%if tmpRec(CurrentPage, CurrentRow, 19)<>"0" then %>red<%else%>black<%end if%>" onkeydown="colschg(<%=CurrentRow-1%>)"  ></TD>
		<TD ALIGN=CENTER><INPUT NAME=Forget VALUE="<%=tmpRec(CurrentPage, CurrentRow, 20)%>" CLASS="readonly" SIZE=3 STYLE="TEXT-ALIGN:RIGHT;color:<%if tmpRec(CurrentPage, CurrentRow, 20)<>"0" then %>red<%else%>black<%end if%>" onkeydown="colschg(<%=CurrentRow-1%>)"  ></TD>
		<TD ALIGN=CENTER><INPUT NAME=LATEFOR VALUE="<%=tmpRec(CurrentPage, CurrentRow, 25)%>" CLASS=<%=INPUTSTS%> SIZE=3 STYLE="TEXT-ALIGN:RIGHT;color:<%if tmpRec(CurrentPage, CurrentRow, 25)<>"0" then %>red<%else%>black<%end if%>" onkeydown="colschg(<%=CurrentRow-1%>)" ></TD>
		<TD ALIGN=CENTER bgcolor="#FBE5CE"><INPUT NAME=H1 VALUE="<%=tmpRec(CurrentPage, CurrentRow, 5)%>" CLASS=<%=INPUTSTS%>  SIZE=3 STYLE="TEXT-ALIGN:RIGHT;color:<%if tmpRec(CurrentPage, CurrentRow, 5)<>"0" then %>red<%else%>black<%end if%>" onkeydown="colschg(<%=CurrentRow-1%>)" onblur="tothn(<%=CurrentRow-1%>)" ></TD>
		<TD ALIGN=CENTER bgcolor="#D5FBDF"><INPUT NAME=H2 VALUE="<%=tmpRec(CurrentPage, CurrentRow, 6)%>" CLASS=<%=INPUTSTS%>  SIZE=3 STYLE="TEXT-ALIGN:RIGHT;color:<%if tmpRec(CurrentPage, CurrentRow, 6)<>"0" then %>red<%else%>black<%end if%>" onkeydown="colschg(<%=CurrentRow-1%>)" onblur="tothn(<%=CurrentRow-1%>)"></TD>
		<TD ALIGN=CENTER bgcolor="#F4DCFB"><INPUT NAME=H3 VALUE="<%=tmpRec(CurrentPage, CurrentRow, 7)%>" CLASS=<%=INPUTSTS%>  SIZE=3 STYLE="TEXT-ALIGN:RIGHT;color:<%if tmpRec(CurrentPage, CurrentRow, 7)<>"0" then %>red<%else%>black<%end if%>" onkeydown="colschg(<%=CurrentRow-1%>)" onblur="tothn(<%=CurrentRow-1%>)"></TD>
		<TD ALIGN=CENTER bgcolor="#E8B5A1"><INPUT NAME=B3 VALUE="<%=tmpRec(CurrentPage, CurrentRow, 8)%>" CLASS=<%=INPUTSTS%>  SIZE=3 STYLE="TEXT-ALIGN:RIGHT;color:<%if tmpRec(CurrentPage, CurrentRow, 8)<>"0" then %>red<%else%>black<%end if%>" onkeydown="colschg(<%=CurrentRow-1%>)" onblur="tothn(<%=CurrentRow-1%>)"></TD>
		<TD ALIGN=CENTER><INPUT NAME=JIAG VALUE="<%=tmpRec(CurrentPage, CurrentRow, 21)%>" CLASS="readonly" readonly  SIZE=3 STYLE="TEXT-ALIGN:RIGHT;color:<%if tmpRec(CurrentPage, CurrentRow, 21)<>"0" then %>red<%else%>black<%end if%>" ></TD>
		<TD ALIGN=CENTER><INPUT NAME=JIAE VALUE="<%=tmpRec(CurrentPage, CurrentRow, 13)%>" CLASS="readonly" readonly  SIZE=3 STYLE="TEXT-ALIGN:RIGHT;color:<%if tmpRec(CurrentPage, CurrentRow, 13)<>"0" then %>red<%else%>black<%end if%>" ></TD>
		<TD ALIGN=CENTER><INPUT NAME=JIAA VALUE="<%=tmpRec(CurrentPage, CurrentRow, 9)%>" CLASS="readonly" readonly  SIZE=3 STYLE="TEXT-ALIGN:RIGHT;color:<%if tmpRec(CurrentPage, CurrentRow, 9)<>"0" then %>red<%else%>black<%end if%>" ></TD>
		<TD ALIGN=CENTER><INPUT NAME=JIAB VALUE="<%=tmpRec(CurrentPage, CurrentRow, 10)%>" CLASS="readonly" readonly  SIZE=3 STYLE="TEXT-ALIGN:RIGHT;color:<%if tmpRec(CurrentPage, CurrentRow, 10)<>"0" then %>red<%else%>black<%end if%>" ></TD>
		<TD ALIGN=CENTER><INPUT NAME=JIAC VALUE="<%=tmpRec(CurrentPage, CurrentRow, 11)%>" CLASS="readonly" readonly  SIZE=3 STYLE="TEXT-ALIGN:RIGHT;color:<%if tmpRec(CurrentPage, CurrentRow, 11)<>"0" then %>red<%else%>black<%end if%>" ></TD>
		<TD ALIGN=CENTER><INPUT NAME=JIAD VALUE="<%=tmpRec(CurrentPage, CurrentRow, 12)%>" CLASS="readonly" readonly  SIZE=3 STYLE="TEXT-ALIGN:RIGHT;color:<%if tmpRec(CurrentPage, CurrentRow, 12)<>"0" then %>red<%else%>black<%end if%>" ></TD>
		<TD ALIGN=CENTER><INPUT NAME=JIAF VALUE="<%=tmpRec(CurrentPage, CurrentRow, 14)%>" CLASS="readonly" readonly  SIZE=3 STYLE="TEXT-ALIGN:RIGHT;color:<%if tmpRec(CurrentPage, CurrentRow, 14)<>"0" then %>red<%else%>black<%end if%>"  ></TD>
		<TD ALIGN=CENTER bgcolor="#ccff66"><INPUT NAME=fgt1 VALUE="<%=tmpRec(CurrentPage, CurrentRow, 28)%>" CLASS="readonly" readonly  SIZE=3 STYLE="TEXT-ALIGN:RIGHT;color:<%if tmpRec(CurrentPage, CurrentRow, 30)<>"0" then %>red<%else%>black<%end if%>"  ></TD>
		<TD ALIGN=CENTER bgcolor="#ccff66"><INPUT NAME=fgt2 VALUE="<%=tmpRec(CurrentPage, CurrentRow, 29)%>" CLASS="readonly" readonly  SIZE=3 STYLE="TEXT-ALIGN:RIGHT;color:<%if tmpRec(CurrentPage, CurrentRow, 30)<>"0" then %>red<%else%>black<%end if%>"  ></TD>
		<TD ALIGN=CENTER bgcolor="#ccff66"><INPUT NAME=fghr VALUE="<%=tmpRec(CurrentPage, CurrentRow, 30)%>" CLASS="readonly" readonly  SIZE=3 STYLE="TEXT-ALIGN:RIGHT;color:<%if tmpRec(CurrentPage, CurrentRow, 30)<>"0" then %>red<%else%>black<%end if%>"  ></TD>
		<TD ALIGN=CENTER bgcolor="#ccff66"><INPUT NAME=cab3 VALUE="<%=tmpRec(CurrentPage, CurrentRow, 32)%>" CLASS="readonly" readonly  SIZE=3 STYLE="TEXT-ALIGN:center;color:<%if tmpRec(CurrentPage, CurrentRow, 30)<>"0" then %>red<%else%>black<%end if%>"  ></TD>
	</TR>
	<%
		sum_TOTHOUR = sum_TOTHOUR + cdbl(tmpRec(CurrentPage, CurrentRow, 4))
		sum_LATEFOR  = sum_LATEFOR + cdbl(tmpRec(CurrentPage, CurrentRow, 25))
		sum_KZhour  = sum_KZhour + cdbl(tmpRec(CurrentPage, CurrentRow, 19))
		sum_Forget  = sum_Forget + cdbl(tmpRec(CurrentPage, CurrentRow, 20))
		sum_H1 = sum_H1 + cdbl(tmpRec(CurrentPage, CurrentRow, 5))
		sum_H2 = sum_H2 + cdbl(tmpRec(CurrentPage, CurrentRow, 6))
		sum_H3 = sum_H3 + cdbl(tmpRec(CurrentPage, CurrentRow, 7))
		sum_B3 = sum_B3 + cdbl(tmpRec(CurrentPage, CurrentRow, 8))
		sum_JIAA = sum_JIAA + cdbl(tmpRec(CurrentPage, CurrentRow, 9))
		sum_JIAB = sum_JIAB	+ cdbl(tmpRec(CurrentPage, CurrentRow, 10))
		sum_JIAC = sum_JIAC + cdbl(tmpRec(CurrentPage, CurrentRow, 11))
		sum_JIAD = sum_JIAD + cdbl(tmpRec(CurrentPage, CurrentRow, 12))
		sum_JIAE = sum_JIAE + cdbl(tmpRec(CurrentPage, CurrentRow, 13))
		sum_JIAF = sum_JIAF + cdbl(tmpRec(CurrentPage, CurrentRow, 14))
		sum_JIAG = sum_JIAG + cdbl(tmpRec(CurrentPage, CurrentRow, 21))
	%>
	<%next%>
	<tr BGCOLOR="Lavender" >
		<td align=right colspan=5 HEIGHT=22>總計 Thống Kê</td>
		<td align=right ><INPUT NAME="sum_TOTHOUR" VALUE="<%=sum_TOTHOUR%>" CLASS=READONLY READONLY  SIZE=3 STYLE="TEXT-ALIGN:RIGHT;color:#002CA5"></td>
		<td align=right ><INPUT NAME="sum_KZhour" VALUE="<%=sum_KZhour%>" CLASS=READONLY   SIZE=3 STYLE="TEXT-ALIGN:RIGHT;color:#002CA5"></td>
		<td align=right ><INPUT NAME="sum_Forget" VALUE="<%=sum_Forget%>" CLASS=READONLY READONLY  SIZE=3 STYLE="TEXT-ALIGN:RIGHT;color:#002CA5"></td>
		<td align=right ><INPUT NAME="sum_LATEFOR" VALUE="<%=sum_LATEFOR%>" CLASS=READONLY READONLY  SIZE=3 STYLE="TEXT-ALIGN:RIGHT;color:#002CA5"></td>
		<td align=right ><INPUT NAME="sum_H1" VALUE="<%=sum_H1%>" CLASS=READONLY READONLY  SIZE=3 STYLE="TEXT-ALIGN:RIGHT;color:#002CA5"></td>
		<td align=right ><INPUT NAME="sum_H2" VALUE="<%=sum_H2%>" CLASS=READONLY READONLY  SIZE=3 STYLE="TEXT-ALIGN:RIGHT;color:#002CA5"></td>
		<td align=right ><INPUT NAME="sum_H3" VALUE="<%=sum_H3%>" CLASS=READONLY READONLY  SIZE=3 STYLE="TEXT-ALIGN:RIGHT;color:#002CA5"></td>
		<td align=right ><INPUT NAME="sum_B3" VALUE="<%=sum_B3%>" CLASS=READONLY READONLY  SIZE=3 STYLE="TEXT-ALIGN:RIGHT;color:#002CA5"></td>
		<td align=right ><INPUT NAME="sum_JIAG" VALUE="<%=sum_JIAG%>" CLASS=READONLY READONLY  SIZE=3 STYLE="TEXT-ALIGN:RIGHT;color:#800000"></td>
		<td align=right ><INPUT NAME="sum_JIAE" VALUE="<%=sum_JIAE%>" CLASS=READONLY READONLY  SIZE=3 STYLE="TEXT-ALIGN:RIGHT;color:#800000"></td>
		<td align=right ><INPUT NAME="sum_JIAA" VALUE="<%=sum_JIAA%>" CLASS=READONLY READONLY  SIZE=3 STYLE="TEXT-ALIGN:RIGHT;color:#800000"></td>
		<td align=right ><INPUT NAME="sum_JIAB" VALUE="<%=sum_JIAB%>" CLASS=READONLY READONLY  SIZE=3 STYLE="TEXT-ALIGN:RIGHT;color:#800000"></td>
		<td align=right ><INPUT NAME="sum_JIAC" VALUE="<%=sum_JIAC%>" CLASS=READONLY READONLY  SIZE=3 STYLE="TEXT-ALIGN:RIGHT;color:#800000"></td>
		<td align=right ><INPUT NAME="sum_JIAD" VALUE="<%=sum_JIAD%>" CLASS=READONLY READONLY  SIZE=3 STYLE="TEXT-ALIGN:RIGHT;color:#800000"></td>
		<td align=right ><INPUT NAME="sum_JIAF" VALUE="<%=sum_JIAF%>" CLASS=READONLY READONLY  SIZE=3 STYLE="TEXT-ALIGN:RIGHT;color:#800000"></td>
	</tr>
</TABLE>

<TABLE border=0 width=600 class=font9 >
<tr>
  <td align="CENTER" height=40  >
	<input type=BUTTON name=send value="關閉此視窗(CLOSE)"  class=button ONCLICK="vbscript:window.close()">　　
	</td>
	<td align=right>
		<%if session("netuser")="LSARY" then %>
			<%if yymm>=nowmonth then %>
				<input type="button" name="send" value="確　定 Xác Nhận " class=button onclick="go()" style="display:none" >
				<input type="RESET" name="send" value="取　消 Hủy" class=button style="display:none">			
			<%end if%>
		<%else%>	
			<%
				if yymm>=nowmonth then 
					types=""	
				else
					if session("netuser")="SEN" or session("rights")="0" then  
						types=""
					else
					 types="disabled"
					end if 					
				end if 
			%>			
			<input type="button" name="btn" value="確　定 Xác Nhận" class=button onclick="go()" >
			<input type="RESET" name="send" value="取　消 Hủy" class=button>
			
		<%end if%> 		
	</td>
	<%if session("netuser")<>"LSARY" then %>
	<td width=120  nowrap align='center'><input type="reset" name="send" value="(m)模 擬" onclick="gom()" class="button"></td>
	<%end if%>
</TR> 
</TABLE>

</form>
<script type='text/javascript'>
	function gom(){
	window.open ("empwork_jd.asp?empid="+"<%=empid%>"+"&yymm="+"<%=yymm%>","_blank","top=50,left=60,width=800,height=500,status=yes,scrollbars=yes,resizable=yes") 
	}
</script>

</body>
</html>

<script language=vbscript >
 

function funcchg(index)
	if <%=self%>.func(index).checked =true then 
		<%=self%>.flag(index).value="Y"
	else
		<%=self%>.flag(index).value=""
	end if 
end function 

function dchg()
	'<%=SELF%>.ACTION="empworkB.FORE.ASP"
	'<%=SELF%>.SUBMIT()
	<%=self%>.TotalPage.value=""
	ymstr = <%=self%>.yymm.value
	tp=<%=self%>.totalpage.value
	cp=<%=self%>.CurrentPage.value
	rc=<%=self%>.RecordInDB.value
	n = <%=self%>.empid.value
	
	open "empworkBN.fore.asp?empid="& N &"&YYMM="&ymstr &"&Ftotalpage=" & tp &"&Fcurrentpage=" & cp &"&FRecordInDB=" & rc , "_self"
	'alert <%=self%>.yymm.value
END function

function TIMEUP_chg(index)
	IF TRIM(<%=SELF%>.TIMEUP(INDEX).VALUE)<>"" THEN
		IF ( LEFT(<%=SELF%>.TIMEUP(INDEX).VALUE,2)>="24" OR RIGHT(<%=SELF%>.TIMEUP(INDEX).VALUE,2)>="60" ) OR LEN(<%=SELF%>.TIMEUP(INDEX).VALUE)>5  THEN
			ALERT "時間格式輸入錯誤!!"
			<%=SELF%>.TIMEUP(INDEX).VALUE=""
			<%=SELF%>.TIMEUP(INDEX).FOCUS()
			CALL CALCHOUR(INDEX)
			EXIT FUNCTION
		ELSE
			<%=SELF%>.TIMEUP(INDEX).VALUE=LEFT(<%=SELF%>.TIMEUP(INDEX).VALUE,2)&":"&RIGHT(<%=SELF%>.TIMEUP(INDEX).VALUE,2)
			CALL CALCHOUR(INDEX)
		END IF
	ELSE
		<%=SELF%>.TOTHOUR(INDEX).VALUE="0"
		<%=SELF%>.KZhour(INDEX).VALUE="0"
		<%=SELF%>.Forget(INDEX).VALUE="0"
		<%=SELF%>.LATEFOR(INDEX).VALUE="0"
		<%=SELF%>.H1(INDEX).VALUE="0"
		<%=SELF%>.H2(INDEX).VALUE="0"
		<%=SELF%>.H3(INDEX).VALUE="0"
		<%=SELF%>.B3(INDEX).VALUE="0"
	END IF

End function

function TIMEDOWN_chg(index)
	IF TRIM(<%=SELF%>.TIMEDOWN(INDEX).VALUE)<>"" THEN
		IF ( LEFT(<%=SELF%>.TIMEDOWN(INDEX).VALUE,2)>="24" OR RIGHT(<%=SELF%>.TIMEDOWN(INDEX).VALUE,2)>="60" ) OR LEN(<%=SELF%>.TIMEDOWN(INDEX).VALUE)>5  THEN
			ALERT "時間格式輸入錯誤!!"
			<%=SELF%>.TIMEDOWN(INDEX).VALUE=""
			<%=SELF%>.TIMEDOWN(INDEX).FOCUS()
			CALL CALCHOUR(INDEX)
			EXIT FUNCTION
		ELSE
			<%=SELF%>.TIMEDOWN(INDEX).VALUE=LEFT(<%=SELF%>.TIMEDOWN(INDEX).VALUE,2)&":"&RIGHT(<%=SELF%>.TIMEDOWN(INDEX).VALUE,2)
			CALL CALCHOUR(INDEX)
		END IF
	ELSE
		<%=SELF%>.TOTHOUR(INDEX).VALUE="0"
		<%=SELF%>.KZhour(INDEX).VALUE="0"
		<%=SELF%>.Forget(INDEX).VALUE="0"
		<%=SELF%>.LATEFOR(INDEX).VALUE="0"
		<%=SELF%>.H1(INDEX).VALUE="0"
		<%=SELF%>.H2(INDEX).VALUE="0"
		<%=SELF%>.H3(INDEX).VALUE="0"
		<%=SELF%>.B3(INDEX).VALUE="0"
	END IF

End function



FUNCTION CALCHOUR(INDEX)
	IF TRIM(<%=SELF%>.TIMEUP(INDEX).VALUE)<>"" AND TRIM(<%=SELF%>.TIMEDOWN(INDEX).VALUE)<>""   THEN
		IF TRIM(<%=SELF%>.TIMEUP(INDEX).VALUE)<>TRIM(<%=SELF%>.TIMEDOWN(INDEX).VALUE) THEN
			if <%=SELF%>.TIMEUP(INDEX).VALUE >="05:00" and <%=SELF%>.TIMEUP(INDEX).VALUE<="06:03" then
				NDDT="06:00"
			elseif <%=SELF%>.TIMEUP(INDEX).VALUE >="06:04" and <%=SELF%>.TIMEUP(INDEX).VALUE<="07:03" then
				NDDT="07:00"
			elseif 	<%=SELF%>.TIMEUP(INDEX).VALUE >="07:00" and <%=SELF%>.TIMEUP(INDEX).VALUE<="08:03" then
				NDDT="08:00"
			elseif <%=SELF%>.TIMEUP(INDEX).VALUE >="12:00" and <%=SELF%>.TIMEUP(INDEX).VALUE<="13:03" then
				NDDT="13:00"
			elseif 	<%=SELF%>.TIMEUP(INDEX).VALUE >="15:00" and <%=SELF%>.TIMEUP(INDEX).VALUE<="16:03" then
				NDDT="16:00"
			elseif 	<%=SELF%>.TIMEUP(INDEX).VALUE >="16:04" and <%=SELF%>.TIMEUP(INDEX).VALUE<="17:03" then
				NDDT="17:00"
			elseif 	<%=SELF%>.TIMEUP(INDEX).VALUE >="19:00" and <%=SELF%>.TIMEUP(INDEX).VALUE<="20:03" then
				NDDT="20:00"
			elseIF RIGHT(<%=SELF%>.TIMEUP(INDEX).VALUE,2)>"15" and ( TRIM(<%=SELF%>.TIMEUP(INDEX).VALUE)<>TRIM(<%=SELF%>.TIMEDOWN(INDEX).VALUE) )  THEN
				NDDT=LEFT(<%=SELF%>.TIMEUP(INDEX).VALUE,2)&":30"
			else
				NDDT=<%=SELF%>.TIMEUP(INDEX).VALUE
			end if
		ELSE
			NDDT= <%=SELF%>.TIMEUP(INDEX).VALUE
		END if

		'遲到 		
		if right(NDDT,2) >"03" and right(NDDT,2) <= "15"  then 
			<%=SELF%>.lateFor(INDEX).VALUE="1"
		end if 
		
		IF TRIM(<%=SELF%>.TIMEUP(INDEX).VALUE)<>TRIM(<%=SELF%>.TIMEDOWN(INDEX).VALUE) THEN
			if right(<%=SELF%>.TIMEDOWN(INDEX).VALUE ,2)<"30"  THEN
				NDDD=LEFT(<%=SELF%>.TIMEDOWN(INDEX).VALUE ,2)&":00"
			ELSE
				NDDD=LEFT(<%=SELF%>.TIMEDOWN(INDEX).VALUE ,2)&":30"
			END IF
		ELSE
			NDDD = <%=SELF%>.TIMEDOWN(INDEX).VALUE
		END IF

		DD1 = <%=SELF%>.WORKDAT(INDEX).VALUE&" "&NDDT
		DD2 = <%=SELF%>.WORKDAT(INDEX).VALUE&" "&NDDD
		'alert dd1
		'alert dd2
		
		'所有請假時數
		TOTJIA = CDBL(<%=SELF%>.JIAA(INDEX).VALUE)+ CDBL(<%=SELF%>.JIAB(INDEX).VALUE)+ CDBL(<%=SELF%>.JIAC(INDEX).VALUE)+ CDBL(<%=SELF%>.JIAD(INDEX).VALUE)+ CDBL(<%=SELF%>.JIAE(INDEX).VALUE)+ CDBL(<%=SELF%>.JIAF(INDEX).VALUE)+ CDBL(<%=SELF%>.JIAG(INDEX).VALUE) 
		
		'(日)工時		
		if <%=SELF%>.Forget(INDEX).VALUE="0" and TOTJIA<"8"  then  
			IF LEFT(<%=SELF%>.TIMEDOWN(INDEX).VALUE,2)< LEFT(<%=SELF%>.TIMEUP(INDEX).VALUE,2) THEN
				TOTH=ROUND( DATEDIFF("N", DD1, DD2 ) /30 ,0 ) /2 + 24
			ELSE
				TOTH=ROUND( DATEDIFF("N", DD1, DD2 ) /30 ,0 ) /2
			END IF
			if TOTH <= 0 then TOTH = 0  
			<%=SELF%>.TOTHOUR(INDEX).VALUE=TOTH 
		else
			TOTH =<%=SELF%>.TOTHOUR(INDEX).VALUE 
			'alert  <%=SELF%>.TOTHOUR(INDEX).VALUE  
		end if 	
		
		

		'alert <%=SELF%>.Forget(INDEX).VALUE 
		'alert toth 
		'忘刷卡
		'if <%=SELF%>.STATUS(INDEX).VALUE="H1" then
		'	IF ( <%=SELF%>.TIMEUP(INDEX).VALUE = <%=SELF%>.TIMEDOWN(INDEX).VALUE ) AND (  TRIM(<%=SELF%>.TIMEUP(INDEX).VALUE)<>"" AND TRIM(<%=SELF%>.TIMEDOWN(INDEX).VALUE)<>"" ) AND (  TRIM(<%=SELF%>.TIMEUP(INDEX).VALUE)<>"00:00" AND TRIM(<%=SELF%>.TIMEDOWN(INDEX).VALUE)<>"00:00" )   THEN
		'		<%=SELF%>.Forget(INDEX).VALUE="1"
		'		<%=self%>.Forget(INDEX).style.color="red"
		'	ELSEIF  TOTH<="0"  AND (  TRIM(<%=SELF%>.TIMEUP(INDEX).VALUE)<>"00:00" AND TRIM(<%=SELF%>.TIMEDOWN(INDEX).VALUE)<>"00:00" )  THEN
		'		<%=SELF%>.Forget(INDEX).VALUE="1"
		'		<%=self%>.Forget(INDEX).style.color="red"
		'	ELSE
		'		<%=SELF%>.Forget(INDEX).VALUE="0"
		'		<%=self%>.Forget(INDEX).style.color="black"
		'	END IF
		'else
		'	<%=SELF%>.Forget(INDEX).VALUE="0"
		'	<%=self%>.Forget(INDEX).style.color="black"
		'end if
		'曠職
		
		IF  TOTH+TOTJIA < 8 AND <%=SELF%>.STATUS(INDEX).VALUE="H1" and <%=self%>.WORKDATIM(index).value>=<%=self%>.indat.value THEN		
			IF Trim(<%=self%>.outdat.value)=""   then
				<%=SELF%>.KZhour(INDEX).VALUE = 8-(TOTH+TOTJIA)
				<%=self%>.KZhour(INDEX).style.color="red"
			ELSEIF  Trim(<%=self%>.outdat.value) <= trim(<%=self%>.workdat(index).value)  THEN
				<%=SELF%>.KZhour(INDEX).VALUE = 0
				<%=self%>.KZhour(INDEX).style.color="black"
			ELSE
				<%=SELF%>.KZhour(INDEX).VALUE = 0
				<%=self%>.KZhour(INDEX).style.color="black"
			END IF
		ELSE
			<%=SELF%>.KZhour(INDEX).VALUE = 0
			<%=self%>.KZhour(INDEX).style.color="black"
		END IF

		'加班時數
		IF <%=SELF%>.STATUS(INDEX).VALUE="H1" THEN
			if <%=self%>.groupid.value="A061" or   <%=self%>.zuno.value="A0591" then
				if totH > 8 then
					<%=SELF%>.H1(INDEX).VALUE=cdbl(totH) - 8
					<%=self%>.H1(INDEX).style.color="red"
				else
					<%=SELF%>.H1(INDEX).VALUE=0
					<%=self%>.H1(INDEX).style.color="black"
				end if
			else
				if totH >= 9 then
					if   <%=self%>.zuno.value="A0656" 	 then
						<%=SELF%>.H1(INDEX).VALUE=cdbl(totH) - 9
						<%=self%>.H1(INDEX).style.color="red"
					elseif left(NDDT,2)>="19" and left(NDDT,2)<="20" then
						<%=SELF%>.H1(INDEX).VALUE=cdbl(totH) - 8
						<%=self%>.H1(INDEX).style.color="red"
					else
						<%=SELF%>.H1(INDEX).VALUE=cdbl(totH) - 9
						<%=self%>.H1(INDEX).style.color="red"
					end if
				else
					<%=SELF%>.H1(INDEX).VALUE=0
					<%=self%>.H1(INDEX).style.color="black"
				end if
			end if
		ELSE
			<%=SELF%>.H1(INDEX).VALUE=0
			<%=self%>.H1(INDEX).style.color="black"
		END IF

		IF <%=SELF%>.STATUS(INDEX).VALUE="H2" THEN
			<%=SELF%>.H2(INDEX).VALUE=totH
			<%=self%>.H2(INDEX).style.color="red"
		ELSE
			<%=SELF%>.H2(INDEX).VALUE=0
			<%=self%>.H2(INDEX).style.color="black"
		END IF

		IF <%=SELF%>.STATUS(INDEX).VALUE="H3" THEN
			<%=SELF%>.H3(INDEX).VALUE=cdbl(totH)
			<%=self%>.H3(INDEX).style.color="red"
		ELSE
			<%=SELF%>.H3(INDEX).VALUE=0
			<%=self%>.H3(INDEX).style.color="black"
		END IF


		'值夜班時數
		if left(<%=SELF%>.TIMEDOWN(INDEX).VALUE,2)>"06"  then 
			'endB3 = <%=SELF%>.WORKDAT(INDEX).VALUE&" 06:00"
			endB3 = DD2 
		else
			endB3 = DD2 
		end if 	
		'自21:00開始算夜班 (change by elin 2009/09/15) 
		IF ( LEFT(<%=SELF%>.TIMEDOWN(INDEX).VALUE,2)>="21" OR left(<%=SELF%>.TIMEDOWN(INDEX).VALUE,1)="0"  or (left(<%=SELF%>.TIMEDOWN(INDEX).VALUE,2)>="00" and left(<%=SELF%>.TIMEDOWN(INDEX).VALUE,2)<="06" ) )   THEN
			if LEFT(<%=SELF%>.TIMEUP(INDEX).VALUE,2)>="21"  or ( left(<%=SELF%>.TIMEUP(INDEX).VALUE,1)>="0"  and left(<%=SELF%>.TIMEUP(INDEX).VALUE,2)<="06" ) then
				B3STR = <%=SELF%>.WORKDAT(INDEX).VALUE&" "&<%=SELF%>.TIMEUP(INDEX).VALUE
			else
				B3STR = <%=SELF%>.WORKDAT(INDEX).VALUE&" 21:00"
			end if
			
			IF left(<%=SELF%>.TIMEDOWN(INDEX).VALUE,1)="0" and ( <%=SELF%>.TIMEDOWN(INDEX).VALUE < <%=SELF%>.TIMEUP(INDEX).VALUE ) THEN
				<%=SELF%>.B3(INDEX).VALUE = round(DATEDIFF("N", B3STR, endB3)/30 ,0)/2 +24
				<%=self%>.B3(INDEX).style.color="red"
			ELSEif   <%=SELF%>.TIMEDOWN(INDEX).VALUE = <%=SELF%>.TIMEUP(INDEX).VALUE then
				<%=self%>.B3(INDEX).VALUE=0
				<%=self%>.B3(INDEX).style.color="BLACK"
			else
				<%=SELF%>.B3(INDEX).VALUE = round(DATEDIFF("N", B3STR, endB3)/30 ,0)/2
				<%=self%>.B3(INDEX).style.color="red"
			END IF
		ELSE
			<%=self%>.B3(INDEX).VALUE=0
			<%=self%>.B3(INDEX).style.color="BLACK"
		END IF


		'(月)總工時
		TOTHR=0
		F_KZhour = 0
		F_Forget = 0
		F_LATEFOR = 0
		F_H1 = 0
		F_H2 = 0
		F_H3 = 0
		F_B3 = 0

		FOR Z = 1 TO (<%=SELF%>.DAYS.VALUE)
			TOTHR=TOTHR+CDBL(<%=SELF%>.TOTHOUR(Z-1).VALUE)
			F_KZhour=F_KZhour+CDBL(<%=SELF%>.KZhour(Z-1).VALUE)
			F_Forget=F_Forget+CDBL(<%=SELF%>.Forget(Z-1).VALUE)
			F_LATEFOR=F_LATEFOR+CDBL(<%=SELF%>.LATEFOR(Z-1).VALUE)
			F_H1=F_H1+CDBL(<%=SELF%>.H1(Z-1).VALUE)
			F_H2=F_H2+CDBL(<%=SELF%>.H2(Z-1).VALUE)
			F_H3=F_H3+CDBL(<%=SELF%>.H3(Z-1).VALUE)
			F_B3=F_B3+CDBL(<%=SELF%>.B3(Z-1).VALUE)
		NEXT
		<%=SELF%>.SUM_TOTHOUR.VALUE = TOTHR
		<%=SELF%>.sum_KZhour.VALUE = F_KZhour
		<%=SELF%>.sum_Forget.VALUE = F_Forget
		<%=SELF%>.sum_LATEFOR.VALUE = F_LATEFOR
		<%=SELF%>.sum_H1.VALUE = F_H1
		<%=SELF%>.sum_H2.VALUE = F_H2
		<%=SELF%>.sum_H3.VALUE = F_H3
		<%=SELF%>.sum_B3.VALUE = F_B3
	ELSE
		<%=SELF%>.SUM_TOTHOUR.VALUE=CDBL(<%=SELF%>.SUM_TOTHOUR.VALUE)-CDBL(<%=SELF%>.TOTHOUR(INDEX).VALUE)
		<%=SELF%>.TOTHOUR(INDEX).VALUE = 0
		<%=SELF%>.H1(INDEX).VALUE = 0
		<%=SELF%>.H2(INDEX).VALUE = 0
		<%=SELF%>.H3(INDEX).VALUE = 0
		<%=SELF%>.B3(INDEX).VALUE = 0
	END IF
	'<%=self%>.flag(index).value="*"
END FUNCTION


function tothn(index)
	FOR Y = 1 TO (<%=SELF%>.DAYS.VALUE)
			TOTHR=TOTHR+CDBL(<%=SELF%>.TOTHOUR(Y-1).VALUE)
			F_KZhour=F_KZhour+CDBL(<%=SELF%>.KZhour(Y-1).VALUE)
			F_Forget=F_Forget+CDBL(<%=SELF%>.Forget(Y-1).VALUE)
			F_LATEFOR=F_LATEFOR+CDBL(<%=SELF%>.LATEFOR(Y-1).VALUE)
			F_H1=F_H1+CDBL(<%=SELF%>.H1(Y-1).VALUE)
			F_H2=F_H2+CDBL(<%=SELF%>.H2(Y-1).VALUE)
			F_H3=F_H3+CDBL(<%=SELF%>.H3(Y-1).VALUE)
			F_B3=F_B3+CDBL(<%=SELF%>.B3(Y-1).VALUE)
	NEXT
	<%=SELF%>.SUM_TOTHOUR.VALUE = TOTHR
	<%=SELF%>.sum_KZhour.VALUE = F_KZhour
	<%=SELF%>.sum_Forget.VALUE = F_Forget
	<%=SELF%>.sum_LATEFOR.VALUE = F_LATEFOR
	<%=SELF%>.sum_H1.VALUE = F_H1
	<%=SELF%>.sum_H2.VALUE = F_H2
	<%=SELF%>.sum_H3.VALUE = F_H3
	<%=SELF%>.sum_B3.VALUE = F_B3
end function


'*******檢查日期*********************************************
FUNCTION date_change(a)

if a=1 then
	INcardat = Trim(<%=self%>.indat.value)
elseif a=2 then
	INcardat = Trim(<%=self%>.BHDAT.value)
elseif a=3 then
	INcardat = Trim(<%=self%>.pduedate.value)
elseif a=4 then
	INcardat = Trim(<%=self%>.vduedate.value)
elseif a=5 then
	INcardat = Trim(<%=self%>.outdat.value)
end if

IF INcardat<>"" THEN
	ANS=validDate(INcardat)
	IF ANS <> "" THEN
		if a=1 then
			Document.<%=self%>.indat.value=ANS
		elseif a=2 then
			Document.<%=self%>.BHDAT.value=ANS
		elseif a=3 then
			Document.<%=self%>.pduedate.value=ANS
		elseif a=4 then
			Document.<%=self%>.vduedate.value=ANS
		elseif a=5 then
			Document.<%=self%>.outdat.value=ANS
		end if
	ELSE
		ALERT "EZ0067:輸入日期不合法 !!"
		if a=1 then
			Document.<%=self%>.indat.value=""
			Document.<%=self%>.indat.focus()
		elseif a=2 then
			Document.<%=self%>.BHDAT.value=""
			Document.<%=self%>.BHDAT.focus()
		elseif a=3 then
			Document.<%=self%>.pduedate.value=""
			Document.<%=self%>.pduedate.focus()
		elseif a=4 then
			Document.<%=self%>.vduedate.value=""
			Document.<%=self%>.vduedate.focus()
		elseif a=5 then
			Document.<%=self%>.outdat.value=""
			Document.<%=self%>.outdat.focus()
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


function go()
	'alert ("ok") 	
	<%=self%>.action="<%=self%>.upd.asp"
	<%=self%>.submit()
end function 


function showWorkTime(index) 
	empidstr = <%=self%>.empid.value 	
	workdatstr = <%=self%>.workdat(index).value 
	
	open "showWorkTime.asp?empid=" & empidstr &"&workdat=" & workdatstr  , "_blank"   , "top=100, left=100, width=500, height=400, scrollbars=yes  " 
end function 
</script>



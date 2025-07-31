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
'response.write  session("netuser") &"<BR>"

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

if right(yymm,2) mod 2 = 0 then 
	ccx=35
else
	ccx=36
end if 		 


'--------------------------------------------------------------------------------------
SQL="SELECT convert(char(10), indat, 111) as Nindate, b.sys_value as groupstr, c.sys_value as zunostr, "&_
	"d.sys_value as jobstr , e.sys_value as grpsstr, f.toth1, f.toth2 , g.tj, g.OverJ, g.jh1, g.jh2 , "&_
	"isnull(h.NOverJ,0) NoverJ , isnull(H.njh1,0) njh1 , isnull(h.njh2,0) njh2 , a.* from  "&_
	"( SELECT * FROM  EMPFILE WHERE ISNULL(STATUS,'')<>'D' AND autoid='"& empautoid &"' ) a "&_
	"left join ( select * from basicCode where func='groupid' ) b on b.sys_type = a.groupid "&_
	"left join ( select * from basicCode where func='zuno' ) c on c.sys_type = a.zuno "&_
	"left join ( select * from basicCode where func='lev' ) d on d.sys_type = a.job "&_
	"left join ( select * from basicCode where func='grps' ) e on e.sys_type = isnull(a.grps,'') "&_
	"left join ( select empid , suM(h1) as totH1, sum(h2) as toth2 from empwork where left(workdat,4)='"& left(yymm,4) &"' group by empid  ) f on f.empid = a.empid "&_
	"left join ( select empid, sum(toth+h1+h2) as Tj,sum( (h1+h2)-"& cdbl(ccx)&" ) as OverJ , sum(h1) jh1, sum(h2) jh2  from empworkper "&_
	"where left(yymm,4)='"& left(yymm,4) &"' group by empid  ) g on g.empid = a.empid "&_
	"left join ( select empid, (toth+h1+h2) as NTj, ( (h1+h2)-"& cdbl(ccx)&" ) as NOverJ , h1 as Njh1, (h2) Njh2  from empworkper "&_
	"where yymm='"& yymm &"') h on h.empid = a.empid "
	'RESPONSE.WRItE SQL
	'RESPONSE.END
	RST.OPEN SQL , CONN, 3, 3
IF NOT RST.EOF THEN
	empautoid = TRIM(RST("AUTOID"))
	EMPID=TRIM(RST("EMPID"))	'員工編號
	INDAT=TRIM(RST("Nindate"))	'到職日
	TX=TRIM(RST("TX"))	'特休
	WHSNO=TRIM(RST("WHSNO"))	'廠別
	UNITNO=TRIM(RST("UNITNO"))	'處/所
	GROUPID=TRIM(RST("GROUPID"))	'組/部門
	ZUNO=TRIM(RST("ZUNO"))	'單位
	JOB=TRIM(RST("JOB"))	'職等
	EMPNAM_CN=TRIM(RST("EMPNAM_CN"))	'姓名(中)
	EMPNAM_VN=TRIM(RST("EMPNAM_VN"))	'姓名(越)
	COUNTRY=TRIM(RST("COUNTRY"))	'國籍
	GROUPSTR = TRIM(RST("GROUPSTR"))  '組/部門
	ZUNOSTR = TRIM(RST("ZUNOSTR"))  '單位
	JOBSTR = TRIM(RST("JOBSTR"))  '職等
	outdat = TRIM(RST("outdat"))  '離職日
	'shift = TRIM(RST("shift"))  '班別
	if TRIM(RST("shift")) ="A" then 
		shift="A班"	
	elseif TRIM(RST("shift")) ="B" then 
		shift="B班"
	elseif TRIM(RST("shift")) ="ALL" then 
		shift="常日班"	
	else
		shift=""
	end if 		
	grps = TRIM(RST("grpsstr"))   
	if session("netuser")="LSARY" then 		 
		toth1 = cdbl(TRIM(RST("jh1"))) 	  		
		toth2 = cdbl(TRIM(RST("jh2")))  
		if rst("NoverJ") > "0" then
			Njh1 = cdbl(RSt("njh1"))- cdbl(rst("NoverJ"))
			Njh2 =  (rst("njh2"))
		else 
			Njh1 =  (rst("njh1"))
			Njh2 =  (rst("njh2"))
		end if	 
		overJ = rst("NoverJ") 
	else
		toth1 = cdbl(TRIM(RST("totH1")))
		toth2 = cdbl(TRIM(RST("toth2")))
	end if 	 
	
END IF
SET RST=NOTHING

'response.write "overJ=" & overJ &"<BR>"


gTotalPage = 1
'PageRec = 31    'number of records per page
if yymm="" then
	PageRec = 31
else
	PageRec = days
end if
TableRec = 40    'number of fields per record  

  
'response.end 
	  

'出缺勤紀錄 --------------------------------------------------------------------------------------

SQLSTRA= " SP_CALCWORKTIME_N '"& EMPID &"', '"& YYMM &"'  "

'response.write 	request("TotalPage")
zz =   0 
if request("TotalPage") = "" or request("TotalPage") = "0" then
	CurrentPage = 1
	'RESPONSE.WRITE SQLSTRA
	'RESPONSE.END
	rs.Open SQLSTRA, conn, 3, 3
	IF NOT RS.EOF THEN
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
			tmpRec(i, j, 1) = trim(rs("DAT"))
			IF RS("TIMEUP")="000000"  AND RS("STATUS")<>"H1"  THEN
				tmpRec(i, j, 2) =""
			ELSE
				tmpRec(i, j, 2) = RS("T1")
			END IF
			IF RS("TIMEDOWN") ="000000" AND RS("STATUS")<>"H1"  THEN
				tmpRec(i, j, 3) = ""
			ELSE
				tmpRec(i, j, 3) = RS("T2")
			END IF  			
			
			'response.write   "456-" & RS("TIMDIF") & "-" & instr(1,cstr(RS("TIMDIF")),".") &"<BR>" 			
			if  instr(1,cstr(RS("TIMDIF")),".") > 0 then  
				'response.write  tmpRec(i, j, 1) &"-" & right(TIMDIF,1) & "-" & cdbl(RS("TIMDIF")) &"<BR>"				
				if right(RS("TIMDIF"),1)="5" then 
					tmpRec(i, j, 4) = cdbl(RS("TIMDIF"))
				else	
					tmpRec(i, j, 4)= round(RS("TIMDIF"),0)
				end if 	
			else 
				tmpRec(i, j, 4) = cdbl(RS("TIMDIF"))
			end if 	 
			'response.write tmpRec(i, j, 4) &"<BR>"
		
			'遲到
			IF ISNULL(RS("flag")) or trim(RS("flag"))="AUTO"  THEN
	 			tmpRec(i, j, 31) = rs("calclateFor")
	 		else
	 			tmpRec(i, j, 31) = rs("lateFor")
	 		end if

			'RESPONSE.WRITE NDAT1  &"<br>"
			'RESPONSE.WRITE tmpRec(i, j, 30)   &"<br>"
			'RESPONSE.WRITE NDAT2  &"<br>"
			'RESPONSE.WRITE NEWTOTH  &"<br>"

			
			 
			'response.write tmpRec(i, j, 1) &"-"&tmpRec(i, j, 5) &"<BR>"

			IF ISNULL(RS("flag")) or trim(RS("flag"))="AUTO"  THEN
				if trim(RS("STATUS"))="H2" then
					tmpRec(i, j, 6) = rs("h2Times")
				else
					tmpRec(i, j, 6)=0
				end if
			ELSE
				'response.write   tmpRec(i, j, 1) &  "3" &"<BR>"
				tmpRec(i, j, 6) = rs("H2")
			END IF
			IF ISNULL(RS("flag")) or trim(RS("flag"))="AUTO"  THEN
				if RS("STATUS")="H3" then
					tmpRec(i, j, 7) = rs("h3times")
				else
					tmpRec(i, j, 7)=0
				end if
			ELSE
				tmpRec(i, j, 7) = rs("H3")
			END IF
			IF ISNULL(RS("flag")) or trim(RS("flag"))="AUTO"  THEN
				tmpRec(i, j, 8) =  rs("b3times")	
			ELSE
				tmpRec(i, j, 8) = rs("B3")
			END IF
			IF ISNULL(RS("JIAA")) THEN
				tmpRec(i, j, 9) = 0
			ELSE
				tmpRec(i, j, 9) = rs("hhoura")
			END IF
			IF ISNULL(RS("JIAB")) THEN
				tmpRec(i, j, 10) = 0
			ELSE
				tmpRec(i, j, 10) = rs("hhourb")
			END IF
			IF ISNULL(RS("JIAC")) THEN
				tmpRec(i, j, 11) = 0
			ELSE
				tmpRec(i, j, 11) = rs("hhourc")
			END IF
			IF ISNULL(RS("JIAD")) THEN
				tmpRec(i, j, 12) = 0
			ELSE
				tmpRec(i, j, 12) = rs("hhourd")
			END IF
			IF ISNULL(RS("JIAE")) THEN
				tmpRec(i, j, 13) = 0
			ELSE
				tmpRec(i, j, 13) = rs("hhoure")
			END IF
			IF ISNULL(RS("JIAF")) THEN
				tmpRec(i, j, 14) = 0
			ELSE
				tmpRec(i, j, 14) = rs("hhourf")
			END IF
			IF ISNULL(RS("JIAG")) THEN
				tmpRec(i, j, 21) = 0
			ELSE
				tmpRec(i, j, 21) = rs("hhourg")
			END IF
			
			tmpRec(i, j, 15)= mid("日一二三四五六",weekday(tmpRec(i, j, 1)) , 1 )
			
			totjia = cdbl(tmpRec(i, j, 9))+cdbl(tmpRec(i, j, 10))+cdbl(tmpRec(i, j, 11))+cdbl(tmpRec(i, j, 12))+cdbl(tmpRec(i, j, 13))+cdbl(tmpRec(i, j, 14))+cdbl(tmpRec(i, j, 21)) 
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
			'tmpRec(i, j, 22) =0 'cdbl(tmpRec(i, j, 9)+cdbl(tmpRec(i, j, 10))+cdbl(tmpRec(i, j, 11))+cdbl(tmpRec(i, j, 12))+cdbl(tmpRec(i, j, 13))+cdbl(tmpRec(i, j, 14))+cdbl(tmpRec(i, j, 21)))
			tmpRec(i, j, 22) = cdbl(tmpRec(i, j, 9))+cdbl(tmpRec(i, j, 10))+cdbl(tmpRec(i, j, 11))+cdbl(tmpRec(i, j, 12))+cdbl(tmpRec(i, j, 13))+cdbl(tmpRec(i, j, 14))+cdbl(tmpRec(i, j, 21))
			'response.write  "22=" & tmpRec(i, j, 22) &"<BR>"
			
			if tmpRec(i, j, 22)>=8 then   '請假超過或等於8小時工時為0
				tmpRec(i, j, 4) = 0
			else
				tmpRec(i, j, 4) = cdbl(tmpRec(i, j, 4))
			end if 
			
			if  overJ<>"" and overJ > "0" then    
				if RS("STATUS")="H1" and  tmpRec(i, j, 4) > 0 then 
					'response.write "1---"&"<BR>"
					if zz <=cdbl(Njh1) then 
						'response.write "y2=" & zz &"<BR>"
						'response.write "njh1-zz=" & cdbl(Njh1)-cdbl(zz) &"<BR>"
						'response.write "instr=" & instr(1,zz,".") &"<BR>" 
						if cdbl(Njh1)>26 then
							if day(rs("DAT")) mod 3 = 0 then 
								if cdbl(Njh1)-cdbl(zz) > 1.5 or ( cdbl(Njh1)-cdbl(zz)<=1.5  and instr(1,njh1,".")=0 )  then 
									if zz+2 > cdbl(Njh1) then 
										tmpRec(i, j, 5) = 0  
										zz = zz + 0 
									else
										tmpRec(i, j, 5) = 2 
										zz = zz + 2
									end if 	
								'	response.write "xx1=" & zz &"<BR>"						 
								else 
									tmpRec(i, j, 5) = cdbl(Njh1)-zz 
									zz = cdbl(Njh1) 
							'		response.write "x=" &  zz &"<BR>"
								end if 								
							else 
								if cdbl(Njh1)-cdbl(zz) > 1.5 or ( cdbl(Njh1)-cdbl(zz)<=1.5  and instr(1,njh1,".")=0 )  then 
									if zz+1 > cdbl(Njh1) then 
										tmpRec(i, j, 5) = 0  
										zz = zz + 0
									else
										tmpRec(i, j, 5) = 1  
										zz = zz + 1 
									end if 	
								'	response.write "xx2=" & zz &"<BR>"						 
								else 
									tmpRec(i, j, 5) = cdbl(Njh1)-zz 
									zz = cdbl(Njh1) 
								'	response.write "x=" &  zz &"<BR>"
								end if 
							end if								
						else						
							if cdbl(Njh1)-cdbl(zz) > 1.5 or ( cdbl(Njh1)-cdbl(zz)<=1.5  and instr(1,njh1,".")=0 )  then 
								if zz+1 > cdbl(Njh1) then 
									tmpRec(i, j, 5) = 0  
									zz = zz + 0
								else
									tmpRec(i, j, 5) = 1  
									zz = zz + 1 
								end if 	
								'response.write "xx=" & zz &"<BR>"						 
							else 
								tmpRec(i, j, 5) = cdbl(Njh1)-zz 
								zz = cdbl(Njh1) 
								'response.write "x=" &  zz &"<BR>"
							end if 
						end if 	
					else
						tmpRec(i, j, 5)= 0
					end if 	  
				else
					tmpRec(i, j, 5)= 0 
				end if	 
								
			else
				if  round(rs("h1Times"),0) >4 then 
					tmpRec(i, j, 5)=4					
				else
					'tmpRec(i, j, 5)= rs("h1Times")					
					tmpRec(i, j, 5) = rs("H1")
					
					'response.write "xxx111<BR>"
				end if 	
			end if 	 			 
			
'工時
				'response.write rs("grps") 
				if rs("grps")="A" or (rs("grps")="B" and left(tmpRec(i, j, 2),2)>="20" )  then 
					cch=8 
				else
					cch=9
				end if	
				if tmpRec(i, j, 2)<>""  and tmpRec(i, j, 2)<>"00:00" then  
					if rs("status")="H1" then 
						clchh = round(cdbl(left(tmpRec(i, j, 2),2))+cdbl(cch)+cdbl(tmpRec(i, j, 5))+1,0)
						'response.write clchh &"<BR>"
						if  clchh  >= 24 then 
							tmpRec(i, j, 3) = right("00"&(clchh-24),2)& right(tmpRec(i, j, 3),3) 
							
						else				
							tmpRec(i, j, 3) = right("00"&clchh,2)& right(tmpRec(i, j, 3),3)
						end if 
						if tmpRec(i, j, 4) > 0 then 
							tmpRec(i, j, 4) =  cch + cdbl(tmpRec(i, j, 5)) 
						end if 	  
						 
			 			if clchh>22 then 
			 				if clchh>24 then 		 					
			 					tmpRec(i, j, 8) = clchh - 24 +2
			 				else
			 					tmpRec(i, j, 8) = clchh - 22  
			 				end if 
			 				if right(tmpRec(i, j, 3),2)>="30" then  
			 					tmpRec(i, j, 8) =  tmpRec(i, j, 8)+ 0.5 
			 				end if 
			 			else
			 				tmpRec(i, j, 8) =  0 
			 			end if	
					end if 		
				end if 	 
			
			

			'曠職
			if  cdate(trim(rs("DAT"))) < cdate(date()) and cdate(trim(rs("DAT")))>= cdate(trim(INDAT)) and  tmpRec(i, j, 18)="H1"  then												
				if ( trim(rs("outdat"))<>""  and trim(rs("DAT"))>= trim(rs("outdat")) )  then 
					tmpRec(i, j, 19) = 0  					
					'response.write "c" &"<BR>"
					'response.write trim(rs("DAT"))  &"<BR>"
					'response.write trim(rs("outdat")) &"<BR>"
					'response.write "c" &"<BR>"
				else
					tmpRec(i, j, 19) = 8 - cdbl(tmpRec(i, j, 4)) - tmpRec(i, j, 22)
					if tmpRec(i, j, 19) < 0 then tmpRec(i, j, 19) = 0 
					if tmpRec(i, j, 19) > 0.5 and tmpRec(i, j, 19) < 1  then tmpRec(i, j, 19) = 0.5 
					'if tmpRec(i, j, 19) < 0.5 then tmpRec(i, j, 19) = 1 
					'response.write "B" &"<BR>"
				end  if 		
			else
				tmpRec(i, j, 19) = 0  	
				'response.write "A" &"<BR>"
			end if
			'response.write tmpRec(i, j, 19) &"<BR>"

			'忘刷
			tmpRec(i, j, 20) = rs("fgcnt")			
			tmprec(i,j,23)=""
			
			sqlT1="delete EMPWORKJD where empid='"& EMPID &"' and workdat='"&  replace(tmpRec(i, j, 1),"/","") &"'"  
			conn.execute(sqlT1) 
			sqlT="insert into  EMPWORKJD ( empid, workdat, timeup, timedown, toth, forget, latefor, kzhour, h1, h2, h3, b3,  yymm, mdtm ) values ( "&_
				 "'"& EMPID &"', '"& replace(tmpRec(i, j, 1),"/","") &"', '"& tmpRec(i, j, 2) &"', '"& tmpRec(i, j, 3) &"', '"& tmpRec(i, j, 4) &"', "&_
				 "'"& tmpRec(i, j, 20) &"', '"& tmpRec(i, j, 31) &"', '"& tmpRec(i, j, 19)&"', '"& tmpRec(i, j, 5) &"', '"& tmpRec(i, j, 6) &"', "&_
				 "'"& tmpRec(i, j, 7) &"', '"& tmpRec(i, j, 8) &"', '"& left(replace(tmpRec(i, j, 1),"/",""),6) &"', getdate() ) " 
			conn.execute(sqlT)	 

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
'response.end 


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

'response.write "x1=" & Njh1 &"<BR>"
'response.write "x2=" & Njh2 &"<BR>"
%>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta HTTP-EQUIV="refresh" >
<link rel="stylesheet" href="../../Include/style.css" type="text/css">
<link rel="stylesheet" href="../../Include/style2.css" type="text/css">
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--
'-----------------enter to next field
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

-->
</SCRIPT>
</head>
<body  topmargin="5" leftmargin="5"  marginwidth="0" marginheight="0"  onkeydown="enterto()"  ONLOAD="F()" >
<form name="<%=self%>"  method="post" action="<%=self%>.upd.asp" >
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
<INPUT TYPE=hidden NAME=njh1 VALUE="<%=njh1%>">
<INPUT TYPE=hidden NAME=njh2 VALUE="<%=njh2%>">
<!-- table width="460" border="0" cellspacing="0" cellpadding="0">
	<tr>
		<TD align=center >員工差勤作業 </TD>
	</tr>
</table>
<hr size=0	style='border: 1px dotted #999999;' align=left width=500 -->
<table width=500   class=font9>
	<TR>
		<td >查詢年月:</td>
		<td COLSPAN=3>
			<select name=yymm class=font9  onchange="dchg()">
				<%for z = 1 to 12
				  yymmvalue = year(date())&right("00"&z,2)
				%>
					<option value="<%=yymmvalue%>" <%if yymmvalue=yymm then %>selected<%end if%>><%=yymmvalue%></option>
				<%next%>
			</select>
			<input type=hiddenT class=readonly readonly  name=days value="<%=days%>" size=5>
			　<FONT COLOR=RED><%=MSGSTR%></FONT>
		</td>
	</TR>
	<tr height=30>
		<td width=60>員工編號:</td>
		<td>
			<input name=empid value="<%=EMPID%>" size=7 class="readonly" readonly style="height:22">
			<input name=empnam value="<%=empnam_cn&" "&empnam_vn%>" size=30 class="readonly" readonly style="height:22">
		</td>
		<td align=right>單位:</td>
		<td>
			<input name=groupidstr value="<%=GROUPSTR%>" size=7 class="readonly" readonly  style="height:22">
			<input name=zunostr value="<%=zunoSTR%>" size=5 class="readonly" readonly style="height:22" >
			<input TYPE=HIDDEN name=groupid value="<%=groupid%>" size=5 >
			<input TYPE=HIDDEN name=zuno value="<%=zuno%>" size=5 >
		</td>
		

	</tr>
</table>
<table width=500 class=font9 >
	<tr>
		<td width=60>到職日期:</td>
		<td><input name=indat value="<%=indat%>" size=11 class="readonly" readonly  style="height:22"></td>

		<td>職等:</td>
		<td><input name=job value="<%=jobSTR%>" size=12 class="readonly" readonly  style="height:22"></td>
		<td align=right>特休(天/小時):</td>
		<td>
			<input name=TX value="<%=tx%>" size=5 class="readonly" readonly  style="height:22">
			<input name=TXH value="<%=tx*8%>" size=5 class="readonly" readonly  style="height:22">
		</td>
	</tr>
	<tr>
		<td width=60>離職日期:</td>
		<td  ><input name=outdat value="<%=outdat%>" size=11 class="readonly" readonly  style="height:22"></td>
		<td align=right>班別:</td>
		<td>
			<input name=shift value="<%=shift%>" size=5 class="readonly" readonly  style="height:22">
			<input name=grps value="<%=grps%>" size=5 class="readonly" readonly  style="height:22">
		</td>
		<td align=right>累計加班(H):</td>
		<td>
			<input name=totJiaH value="<%if fix(toth1+toth2)>300 then %><%=300-cdbl(right(empid,2))%><%else%><%=fix(toth1+toth2)%><%end if%>" size=10 class="readonly" readonly  style="height:22;text-align:right">		
		</td>
	</tr>
</table>
<hr size=0	style='border: 1px dotted #999999;' align=left width=500>
<TABLE WIDTH=610 CLASS=FONT9 >
	<TR BGCOLOR=#CCCCCC>
		<TD ROWSPAN=2 ALIGN=CENTER>日期</TD>
		<TD ROWSPAN=2 ALIGN=CENTER>上班</TD>
		<TD ROWSPAN=2 ALIGN=CENTER>下班</TD>
		<TD ROWSPAN=2 ALIGN=CENTER>工時</TD>
		<TD ROWSPAN=2 ALIGN=CENTER>曠職</TD>
		<TD ROWSPAN=2 ALIGN=CENTER>忘<br>刷<br>卡</TD>
		<TD ROWSPAN=2 ALIGN=CENTER>遲到</TD>
		<TD COLSPAN=4 ALIGN=CENTER>加班(單位：小時)</TD>
		<TD COLSPAN=7 ALIGN=CENTER>休假(單位：小時)</TD>
	</TR>
	<TR BGCOLOR=#CCCCCC>
		<TD ALIGN=CENTER>一般(1.5)</TD>
		<TD ALIGN=CENTER>休息(2)</TD>
		<TD ALIGN=CENTER>假日(3)</TD>
		<TD ALIGN=CENTER>夜班(0.3)</TD>
		<TD ALIGN=CENTER>公假</TD>
		<TD ALIGN=CENTER>年假</TD>
		<TD ALIGN=CENTER>事假</TD>
		<TD ALIGN=CENTER>病假</TD>
		<TD ALIGN=CENTER>婚假</TD>
		<TD ALIGN=CENTER>喪假</TD>
		<TD ALIGN=CENTER>產假</TD>
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
		<TD ALIGN=CENTER NOWRAP >
		<a href="vbscript:showWorkTime(<%=currentrow-1%>)" ><font color=blue><%=tmpRec(CurrentPage, CurrentRow, 1)&"("&tmpRec(CurrentPage, CurrentRow, 15)&")"%></font></a>
		<INPUT type=hidden NAME=WORKDATIM VALUE="<%=tmpRec(CurrentPage, CurrentRow, 1)&"("&tmpRec(CurrentPage, CurrentRow, 15)&")"%>" CLASS=READONLY READONLY  SIZE=15 STYLE="TEXT-ALIGN:CENTER;color:<%if weekday(tmpRec(CurrentPage, CurrentRow, 1))=1 then %>RoyalBlue<%else%>black<%end if%>">
		<input type=hidden name=flag value="" size=1>
		<INPUT TYPE=HIDDEN NAME=WORKDAT VALUE="<%=tmpRec(CurrentPage, CurrentRow, 1)%>" onfocus="this.select()">
		<INPUT TYPE=HIDDEN NAME=STATUS VALUE="<%=tmpRec(CurrentPage, CurrentRow, 18)%>" onfocus="this.select()">
		</TD>
		<TD ALIGN=CENTER><INPUT NAME=TIMEUP VALUE="<%=tmpRec(CurrentPage, CurrentRow, 2)%>" CLASS=<%=tmpRec(CurrentPage, CurrentRow, 16)%> SIZE=6 STYLE="TEXT-ALIGN:CENTER"  maxlength=5 <%if yymm>=nowmonth then %>  ONBLUR="TIMEUP_chg(<%=CurrentRow-1%>)" onkeydown="colschg(<%=CurrentRow-1%>)" <%end if %>></TD>
		<TD ALIGN=CENTER><INPUT NAME=TIMEDOWN VALUE="<%=tmpRec(CurrentPage, CurrentRow, 3)%>" CLASS=<%=tmpRec(CurrentPage, CurrentRow, 17)%> SIZE=6 STYLE="TEXT-ALIGN:CENTER" ONBLUR="TIMEDOWN_chg(<%=CurrentRow-1%>)" maxlength=5  onkeydown="colschg(<%=CurrentRow-1%>)" ></TD>
		<TD ALIGN=CENTER><INPUT NAME=TOTHOUR VALUE="<%=tmpRec(CurrentPage, CurrentRow, 4)%>" CLASS="readonly"   SIZE=3 STYLE="TEXT-ALIGN:RIGHT"  onkeydown="colschg(<%=CurrentRow-1%>)" ></TD>
		<TD ALIGN=CENTER><INPUT NAME=KZhour VALUE="<%=tmpRec(CurrentPage, CurrentRow, 19)%>" CLASS="readonly" readonly  SIZE=3 STYLE="TEXT-ALIGN:RIGHT;color:<%if tmpRec(CurrentPage, CurrentRow, 19)<>"0" then %>red<%else%>black<%end if%>" onkeydown="colschg(<%=CurrentRow-1%>)"  ></TD>
		<TD ALIGN=CENTER><INPUT NAME=Forget VALUE="<%=tmpRec(CurrentPage, CurrentRow, 20)%>" CLASS="readonly" SIZE=3 STYLE="TEXT-ALIGN:RIGHT;color:<%if tmpRec(CurrentPage, CurrentRow, 20)<>"0" then %>red<%else%>black<%end if%>" onkeydown="colschg(<%=CurrentRow-1%>)"  ></TD>
		<TD ALIGN=CENTER><INPUT NAME=LATEFOR VALUE="<%=tmpRec(CurrentPage, CurrentRow, 31)%>" CLASS=<%=INPUTSTS%> SIZE=3 STYLE="TEXT-ALIGN:RIGHT;color:<%if tmpRec(CurrentPage, CurrentRow, 31)<>"0" then %>red<%else%>black<%end if%>" onkeydown="colschg(<%=CurrentRow-1%>)" ></TD>
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
	</TR>
	<%
		sum_TOTHOUR = sum_TOTHOUR + cdbl(tmpRec(CurrentPage, CurrentRow, 4))
		sum_LATEFOR  = sum_LATEFOR + cdbl(tmpRec(CurrentPage, CurrentRow, 31))
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
		<td align=right colspan=3 HEIGHT=22>總計</td>
		<td align=right ><INPUT NAME="sum_TOTHOUR" VALUE="<%=sum_TOTHOUR%>" CLASS=READONLY READONLY  SIZE=3 STYLE="TEXT-ALIGN:RIGHT;color:#002CA5"></td>
		<td align=right ><INPUT NAME="sum_KZhour" VALUE="<%=sum_KZhour%>" CLASS=READONLY   SIZE=3 STYLE="TEXT-ALIGN:RIGHT;color:#002CA5"></td>
		<td align=right ><INPUT NAME="sum_Forget" VALUE="<%=sum_Forget%>" CLASS=READONLY READONLY  SIZE=3 STYLE="TEXT-ALIGN:RIGHT;color:#002CA5"></td>
		<td align=right ><INPUT NAME="sum_LATEFOR" VALUE="<%=sum_LATEFOR%>" CLASS=READONLY READONLY  SIZE=3 STYLE="TEXT-ALIGN:RIGHT;color:#002CA5"></td>
		<td align=right ><INPUT NAME="sum_H1" VALUE="<%=Njh1%>" CLASS=READONLY READONLY  SIZE=3 STYLE="TEXT-ALIGN:RIGHT;color:#002CA5"></td>
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
				<input type="button" name="send" value="確　定" class=button onclick="go()" >
				<input type="RESET" name="send" value="取　消" class=button>			
			<%end if%>
		<%else%>	
			<input type="button" name="send" value="確　定" class=button onclick="go()" >
			<input type="RESET" name="send" value="取　消" class=button>
		<%end if%>
		
	</td>
</TR>
</TABLE>

</form>


</body>
</html>

<script language=vbscript >
function BACKMAIN()
	open "../main.asp" , "_self"
end function

FUNCTION dchg()
	
	
	<%=SELF%>.TOTALPAGE.VALUE=""
	<%=SELF%>.ACTION="empworkBT.FORE.ASP"
	<%=SELF%>.SUBMIT() 
	'ymstr = <%=self%>.yymm.value
	'tp=<%=self%>.totalpage.value
	'cp=<%=self%>.CurrentPage.value
	'rc=<%=self%>.RecordInDB.value
	'n = <%=self%>.empautoid.value
	'open "empworkB.fore.asp?totalpage=0&empautoid="& N &"&YYMM="&ymstr &"&Ftotalpage=" & tp &"&Fcurrentpage=" & cp &"&FRecordInDB=" & rc , "_self"
	'alert <%=self%>.yymm.value
END FUNCTION

function hback()
	'alert <%=currentpage%>
	open "EMPWORK.FORE.ASP?send=NEXT&totalpage=" & <%=Ftotalpage%> &"&currentpage=" & <%=Fcurrentpage-1%> &"&RecordInDB=" & <%=FRecordInDB%>  , "_self"
	'OPEN "EMPWORK.FORE.ASP", "_self"
end function

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
		
		IF  TOTH+TOTJIA < 8   AND  <%=SELF%>.STATUS(INDEX).VALUE="H1" and  <%=self%>.WORKDATIM(index).value>=<%=self%>.indat.value    THEN		
			IF Trim(<%=self%>.outdat.value)=""   then
				<%=SELF%>.KZhour(INDEX).VALUE = 8-(TOTH+TOTJIA)
				<%=self%>.KZhour(INDEX).style.color="red"
			ELSEIF  Trim(<%=self%>.outdat.value) <= <%=self%>.workdat(index).value  THEN
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
		IF ( LEFT(<%=SELF%>.TIMEDOWN(INDEX).VALUE,2)>="22" OR left(<%=SELF%>.TIMEDOWN(INDEX).VALUE,1)="0"  or (left(<%=SELF%>.TIMEDOWN(INDEX).VALUE,2)>="00" and left(<%=SELF%>.TIMEDOWN(INDEX).VALUE,2)<="06" ) )   THEN
			if LEFT(<%=SELF%>.TIMEUP(INDEX).VALUE,2)>="22"  or ( left(<%=SELF%>.TIMEUP(INDEX).VALUE,1)>="0"  and left(<%=SELF%>.TIMEUP(INDEX).VALUE,2)<="06" ) then
				B3STR = <%=SELF%>.WORKDAT(INDEX).VALUE&" "&<%=SELF%>.TIMEUP(INDEX).VALUE
			else
				B3STR = <%=SELF%>.WORKDAT(INDEX).VALUE&" 22:00"
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
'________________________________________________________________________________________

FUNCTION CHKVALUE(N)
IF N=1 THEN
	IF TRIM(<%=SELF%>.BYY.VALUE)<>"" THEN
		IF ISNUMERIC(<%=SELF%>.BYY.VALUE)=FALSE OR INSTR(1,<%=SELF%>.BYY.VALUE,"-")>0 THEN
			ALERT "輸入錯誤!!請輸入正確數字"
			<%=SELF%>.BYY.VALUE=""
			<%=SELF%>.BYY.FOCUS()
			EXIT FUNCTION
		END IF
	END IF
ELSEIF N=2 THEN
	IF TRIM(<%=SELF%>.BMM.VALUE)<>"" THEN
		IF ISNUMERIC(<%=SELF%>.BMM.VALUE)=FALSE OR INSTR(1,<%=SELF%>.BMM.VALUE,"-")>0 THEN
			ALERT "輸入錯誤!!請輸入正確數字"
			<%=SELF%>.BMM.VALUE=""
			<%=SELF%>.BMM.FOCUS()
			EXIT FUNCTION
		END IF
	END IF
ELSEIF N=3 THEN
	IF TRIM(<%=SELF%>.BDD.VALUE)<>"" THEN
		IF ISNUMERIC(<%=SELF%>.BDD.VALUE)=FALSE OR INSTR(1,<%=SELF%>.BDD.VALUE,"-")>0 THEN
			ALERT "輸入錯誤!!請輸入正確數字"
			<%=SELF%>.BDD.VALUE=""
			<%=SELF%>.BDD.FOCUS()
			EXIT FUNCTION
		END IF
	END IF
ELSEIF N=4 THEN
	IF TRIM(<%=SELF%>.AGES.VALUE)<>"" THEN
		IF ISNUMERIC(<%=SELF%>.AGES.VALUE)=FALSE OR INSTR(1,<%=SELF%>.AGES.VALUE,"-")>0 THEN
			ALERT "輸入錯誤!!請輸入正確數字"
			<%=SELF%>.AGES.VALUE=""
			<%=SELF%>.AGES.FOCUS()
			EXIT FUNCTION
		END IF
	END IF
ELSEIF N=5 THEN
	IF TRIM(<%=SELF%>.GTDAT.VALUE)<>"" THEN
		IF ISNUMERIC(<%=SELF%>.GTDAT.VALUE)=FALSE OR INSTR(1,<%=SELF%>.GTDAT.VALUE,"-")>0 THEN
			ALERT "輸入錯誤!!請輸入正確數字"
			<%=SELF%>.GTDAT.VALUE=""
			<%=SELF%>.GTDAT.FOCUS()
			EXIT FUNCTION
		END IF
	END IF
END IF

END FUNCTION


function go()
	'alert "ok" 
	if <%=self%>.UID.value="LSARY" then 
		if ( cdbl(<%=self%>.sum_h1.value)+cdbl(<%=self%>.sum_h2.value) )+ cdbl(<%=self%>.totjiaH.value) > 300  then 
			alert "加班時數超過(>)300小時,不可再加班!!"
			exit function
		else
			<%=self%>.action="<%=self%>.upd.asp"
			<%=self%>.submit()		
		end if	
	else
		<%=self%>.action="<%=self%>.upd.asp"
		<%=self%>.submit()
	end if	
end function 


function showWorkTime(index) 
	empidstr = <%=self%>.empid.value 	
	workdatstr = <%=self%>.workdat(index).value 
	
	open "showWorkTime.asp?empid=" & empidstr &"&workdat=" & workdatstr  , "_blank"   , "top=100, left=100, width=500, height=400, scrollbars=yes  " 
end function 

</script>



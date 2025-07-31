<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<!-- #include file = "../GetSQLServerConnection.fun" --> 
<!-- #include file="../Include/SIDEINFO.inc" -->

<%
'on error resume next   
session.codepage="65001"
SELF = "YECE0101"
 
Set conn = GetSQLServerConnection()	  
Set rs = Server.CreateObject("ADODB.Recordset")   
YYMM=REQUEST("YYMM")
F_whsno = trim(request("F_whsno"))
unitno = trim(request("unitno"))
groupid = trim(request("groupid"))
F_country = trim(request("F_country"))  
code03=replace(F_country," ","")
'response.write code03
'response.end 
job = trim(request("job"))
QUERYX = trim(request("empid1"))  
outemp = request("outemp")
lastym = left(yymm,4) &  right("00" & cstr(right(yymm,2)-1) ,2 )
nowmonth = left(year(date()),4) &  right("00" & cstr(month(date())) ,2 ) 
nzs=request("nzs")
'response.write code03 
'response.end 
calcdt = left(YYMM,4)&"/"& right(yymm,2)&"/01"   
'下個月
if right(yymm,2)="12" then 
	ccdt = cstr(left(YYMM,4)+1)&"/01/01" 
else
	ccdt = left(YYMM,4)&"/"& right("00" & right(yymm,2)+1,2)  &"/01"  
end if 	 
'response.write ccdt  
 

'上ㄧ個月
if right(yymm,2)="01"  then 
	lastym = left(yymm,4)-1 &"12" 
else
	lastym=left(yymm,4)&right("00"&right(yymm,2)-1,2)
end if 	
 
 '一個月有幾天 
cDatestr=CDate(LEFT(YYMM,4)&"/"&RIGHT(YYMM,2)&"/01") 
days = DAY(cDatestr+(32-DAY(cDatestr))-DAY(cDatestr+(32-DAY(cDatestr))))   '一個月有幾天  
'本月最後一天 
ENDdat = CDate(LEFT(YYMM,4)&"/"&RIGHT(YYMM,2)&"/"&DAYS) 


      

'本月假日天數 (星期日)
SQL=" SELECT * FROM YDBMCALE WHERE CONVERT(CHAR(6),DAT,112)  ='"& YYMM &"' AND  DATEPART( DW,DAT ) ='1'  " 
Set rsTT = Server.CreateObject("ADODB.Recordset")   
RSTT.OPEN SQL, CONN, 3, 3 
IF NOT RSTT.EOF THEN 
	HHCNT = CDBL(RSTT.RECORDCOUNT)
ELSE
	HHCNT = 0 
END IF 
SET RSTT=NOTHING   

'RESPONSE.WRITE HHCNT &"<br>" 
'RESPONSE.END  
'本月應記薪天數 
MMDAYS = CDBL(days)-CDBL(HHCNT) 
'RESPONSE.WRITE  MMDAYS 
'RESPONSE.END 
'----------------------------------------------------------------------------------------


gTotalPage = 1
PageRec = 10    'number of records per page
TableRec = 50    'number of fields per record  

  
'response.write sql&"<BR>"
'response.end 



if request("TotalPage") = "" or request("TotalPage") = "0" then 
	CurrentPage = 1	
	sqlx="exec sp_yece0101 '"& yymm &"','"& code03&"','"&f_whsno&"','"&groupid&"','"&QUERYX&"','"&outemp&"','"&nzs&"'  "
	'response.write "<BR><BR><BR>"&sqlx  &"<BR>"
	conn.execute(sqlx)
	
	sql="select distinct * from Tmp_yece01 where empid<>'' and DATEDIFF(day,indat,isnull(outdat,'9999/01/30')) > 3 and empid like'%"&QUERYX&"%' "
	if outemp="D" then  
		sql=sql&" and ( isnull(outdat,'')<>'' and  outdat>'"& calcdt &"' )  " 
	elseif len(outemp)=6 then
		sql=sql&" and  convert(char(6),indat,112)='"& outemp &"'    " 
	end if  
	if nzs<>"" then 
		sql=sql&" and  datediff( m, indat, '"& calcdt &"'  ) " & nzs 
	end if 
	sql=sql&" order by empid,whsno  ,  country ,  indat" 
	
	'response.write "<BR><BR><BR>"&sql&"<BR>"	
	'response.end 
	
	rs.Open SQL, conn, 3,1 
	
	IF NOT RS.EOF THEN 
		if F_whsno="DN" then 
			pagerec = rs.RecordCount  
		end if 	
		rs.PageSize = PageRec 
		RecordInDB = rs.RecordCount  
		TotalPage = rs.PageCount  
		gTotalPage = TotalPage
		TableRec = rs.fields.count
	END IF 	 

	Redim tmpRec(gTotalPage, PageRec, TableRec)   'Array   	
	for i = 1 to TotalPage 
	 for j = 1 to PageRec
		if not rs.EOF then 
				tmpRec(i, j, 0) = "no"
				tmpRec(i, j, 1) = trim(rs("empid"))				
				tmpRec(i, j, 2) = trim(rs("empnam_cn"))
				tmpRec(i, j, 3) = trim(rs("empnam_vn"))&"("&rs("school")&")"
				tmpRec(i, j, 4) = rs("country")
				tmpRec(i, j, 5) = rs("nindat") 				
				tmpRec(i, j, 6) = rs("lj")	
				'response.write rs("lj") &"  "&rs("nz") &"<BR>"
				if (rs("country")="VN" or rs("country")="CT") and ( trim(RS("BHDAT")) <>"" and RS("BHDAT") <=  calcdt )  then 
					if rs("lj")="EV0" then 
						tmpRec(i, j, 6)="EV1"
					end if 
				 elseif rs("country")="CN"   then 					
					if cdbl(rs("nz"))>3 then   'CN過試用期(3個月後) 轉為副班長(管理員)
						if left(rs("lj"),3)<"EV3" then 
							tmpRec(i, j, 6)="EV3"
						end if 
					end if 
				end if 			
								
				tmpRec(i, j, 7) = rs("lw")	  
				tmpRec(i, j, 8) = "" 'rs("unitno") 
				tmpRec(i, j, 9)	=RS("lg")  				
				tmpRec(i, j, 10)=RS("lz") 				
				tmpRec(i, j, 11)="" 'RS("lwstr") 	
				tmpRec(i, j, 12)=""  'RS("ustr") 	
				tmpRec(i, j, 13)=RS("lgstr") 	
				tmpRec(i, j, 14)=RS("lzstr") 	
				tmpRec(i, j, 15)=RS("ljstr") 	
				tmpRec(i, j, 16)="" 'RS("cstr") 	
				tmpRec(i, j, 17)=RS("autoid") 	
				IF RS("lz")="XX" THEN 
					tmpRec(i, j, 18)=""
				ELSE
					tmpRec(i, j, 18)=RS("lz")
				END IF 
				tmpRec(i, j, 19)=RS("b_wp") 			'wp 薪資
				tmpRec(i, j, 20)=cdbl(RS("B_BB"))  '基本薪資				 				
				tmpRec(i, j, 21)="" 'rs("bbcode")
				if rs("country")="VN" or rs("country")="TA" or rs("country")="CT"  then 
					tmpRec(i, j, 22)=cdbl(RS("b_cv"))  '職務加給
				else
					tmpRec(i, j, 22)=cdbl(rs("B_CV"))
				end if 	
				tmpRec(i, j, 23)=cdbl(RS("B_PHU"))		'Y獎金 (陸幹為其他加給)
				tmpRec(i, j, 24)=cdbl(RS("B_NN"))  '語言加給
				tmpRec(i, j, 25)=cdbl(RS("B_KT")) '技術加給
				tmpRec(i, j, 26)=cdbl(RS("B_MT")) '環境加給(陸幹為年資加給)
				tmpRec(i, j, 27)=cdbl(RS("B_TTKH"))  '其他加給(陸幹為補助醫療)
				tmpRec(i, j, 28)=RS("BHDAT") '買保險日期
				tmpRec(i, j, 29)=RS("GTDAT") '工團日期
				tmpRec(i, j, 30)=RS("OUTDATE") '離職日期 		 
				TOTY=  CDBL( ( CDBL(tmpRec(i, j, 20))+CDBL(tmpRec(i, j, 22))+CDBL(tmpRec(i, j, 23)) )  )  'BB+CV+PHU
				if rs("country")="VN" or rs("country")="CT" then 					
					if TOTY mod (26*8)<>0 then 
		  				TTMH = fix(TOTY/26/8)+1 		  				
		  			else
		  				TTMH = fix(TOTY/26/8) 
		  			end if 
		  			tmpRec(i, j, 31) = TTMH 
				else
					tmpRec(i, j, 31) = round(tmpRec(i, j, 20)/30,3)
				end if 	
				
				'if cdbl(rs("qcb")) > cdbl(RS("B_QC")) then 
				tmpRec(i, j, 32)=cdbl(rs("b_qc"))
								 	
				tmpRec(i, j, 33)= cdbl(TOTY)+cdbl(rs("b_btien"))+cdbl(tmpRec(i, j, 19))+cdbl(tmpRec(i, j, 24))+cdbl(tmpRec(i, j, 25))+cdbl(tmpRec(i, j, 26))+cdbl(tmpRec(i, j, 27))+cdbl(tmpRec(i, j, 32))
				if F_whsno="DN" and F_country="VN" then 
					tmpRec(i, j, 34)=rs("b_memo")  & " "&rs("empflag")
				else	
					tmpRec(i, j, 34)=rs("b_memo")  	
				end if 	
				if rs("eid")="" then 
					tmpRec(i, j, 35) = "red"
				else
					tmpRec(i, j, 35) = "black"
				end if 
				tmpRec(i, j, 37) = rs("nz")
				' tmpRec(i, j, 38) = (rs("nz")\6)*20 
				'陸幹年資加給(半年20USD)  				
				'response.write cdbl(rs("nz")) & "---" & round( cdbl(rs("JbHour"))/8/30 ,0) &"<BR>"
				'response.write cdbl(rs("nz")) - round( cdbl(rs("JbHour"))/8/30 ,0) &"<BR>" 
				
				if cdbl(rs("JbHour")) > 0 then '扣除留職停薪與產假的年資 
					f_nz = cdbl(rs("nz")) - round( cdbl(rs("JbHour"))/8.0/30.0 ,0)
				else
					f_nz  = rs("nz")
				end if 	
				tmpRec(i, j, 38) = cdbl(RS("B_MT")) ' (cdbl(f_nz)\6)*20    '計算至201002為止,201003起取消(CN)年資加給  modify by elin 20010331
				
				
				if rs("lw")="DN" and rs("country")="VN" then 
					if trim(RS("BHDAT")) ="" and tmpRec(i, j, 37) < 2 then 
						tmpRec(i, j, 39)="B0" 
					elseif 	tmpRec(i, j, 37) >=2 and tmpRec(i, j, 37)<=12 then 
						tmpRec(i, j, 39)="B1" 
					elseif 	tmpRec(i, j, 37)>12 and tmpRec(i, j, 37)<=24 then 
						tmpRec(i, j, 39)="B2" 	
					elseif 	tmpRec(i, j, 37)>24 and tmpRec(i, j, 37)<=36 then 
						tmpRec(i, j, 39)="B3" 		
					elseif 	tmpRec(i, j, 37)>36 and tmpRec(i, j, 37)<=48 then 
						tmpRec(i, j, 39)="B4" 			
					elseif 	tmpRec(i, j, 37)>48 and tmpRec(i, j, 37)<=60 then 
						tmpRec(i, j, 39)="B5" 				
					elseif 	tmpRec(i, j, 37)>60 and tmpRec(i, j, 37)<=72 then 
						tmpRec(i, j, 39)="B6" 					
					elseif 	tmpRec(i, j, 37)>72 and tmpRec(i, j, 37)<=84 then 
						tmpRec(i, j, 39)="B7" 						
					elseif 	tmpRec(i, j, 37)>84 and tmpRec(i, j, 37)<=96 then 
						tmpRec(i, j, 39)="B8" 							
					elseif 	tmpRec(i, j, 37)>96 and tmpRec(i, j, 37)<=108 then 
						tmpRec(i, j, 39)="B9" 		
					elseif 	tmpRec(i, j, 37)>108 and tmpRec(i, j, 37)<=120 then 
						tmpRec(i, j, 39)="BA" 			
					end if 	
				else
					if trim(RS("BHDAT")) ="" then 
						tmpRec(i, j, 39)="B0" 
					elseif 	tmpRec(i, j, 37)<12 then 
						tmpRec(i, j, 39)="B1" 
					elseif 	tmpRec(i, j, 37)>=12 and tmpRec(i, j, 37)<24 then 
						tmpRec(i, j, 39)="B2" 	
					elseif 	tmpRec(i, j, 37)>=24 and tmpRec(i, j, 37)<36 then 
						tmpRec(i, j, 39)="B3" 		
					elseif 	tmpRec(i, j, 37)>=36 and tmpRec(i, j, 37)<48 then 
						tmpRec(i, j, 39)="B4" 			
					elseif 	tmpRec(i, j, 37)>=48 and tmpRec(i, j, 37)<60 then 
						tmpRec(i, j, 39)="B5" 				
					elseif 	tmpRec(i, j, 37)>=60 and tmpRec(i, j, 37)<72 then 
						tmpRec(i, j, 39)="B6" 					
					elseif 	tmpRec(i, j, 37)>=72 and tmpRec(i, j, 37)<84 then 
						tmpRec(i, j, 39)="B7" 						
					elseif 	tmpRec(i, j, 37)>=84 and tmpRec(i, j, 37)<96 then 
						tmpRec(i, j, 39)="B8" 							
					elseif 	tmpRec(i, j, 37)>=96 and tmpRec(i, j, 37)<108 then 
						tmpRec(i, j, 39)="B9" 		
					elseif 	tmpRec(i, j, 37)>=108 and tmpRec(i, j, 37)<120 then 
						tmpRec(i, j, 39)="BA" 			
					end if 	
				end if 	
				if rs("lw")="DN" and rs("country")="VN" then 
					if trim(rs("lests"))="v" then 
						tmpRec(i, j, 39) = "Z5"&tmpRec(i, j, 39) 
					elseif rs("empflag")="T" then 
						tmpRec(i, j, 39) = "Z1"&tmpRec(i, j, 39) 
					elseif rs("empflag")="CN" then 
						tmpRec(i, j, 39) = "Z2"&tmpRec(i, j, 39) 
					elseif rs("empflag")="NV" then 
						tmpRec(i, j, 39) = "Z3"&tmpRec(i, j, 39) 	
					end if 
				else
					if trim(rs("lests"))="v" then 
						tmpRec(i, j, 39) = "Z3"&tmpRec(i, j, 39) 
					else
						tmpRec(i, j, 39) = "Z1"&tmpRec(i, j, 39) 
					end if 
				end if 	
				if rs("country")="VN" or rs("country")="CT" then 
					tmpRec(i, j, 40) = 0 'rs("CB")
					tmpRec(i, j, 41) = 0 'cdbl(rs("CB"))-cdbl(rs("BB"))-cdbl(rs("sys_cv"))-cdbl(rs("b_phu"))
				else
					tmpRec(i, j, 40) = 0 
					tmpRec(i, j, 41) =  0
				end if 
				tmpRec(i, j, 42) = rs("lests") 
				tmpRec(i, j, 43) = rs("code")'rs("lncode") 
				tmpRec(i, j, 44) = rs("B_btien")  				
				'tmpRec(i, j, 45) = rs("whsno_acc") no exist
				tmpRec(i, j, 45) = rs("whsno")
				if  rs("country")<>"VN" then  
					tmpRec(i, j, 46) = 0'rs("wpbtien")  
				else	
					tmpRec(i, j, 46) = 0 
				end if 				
				tmpRec(i, j, 47) = rs("code") 
				tmpRec(i, j, 33)= cdbl(tmpRec(i, j, 33))+cdbl(tmpRec(i, j, 46))
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
	Session("empsalary01") = tmpRec	
else    
	TotalPage = cint(request("TotalPage"))
	'StoreToSession()
	CurrentPage = cint(request("CurrentPage"))
	RecordInDB  = REQUEST("RecordInDB")
	tmpRec = Session("empsalary01")
	
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
<script src='../js/enter2tab.js'></script>
 
</head>   
<body leftmargin="0"  marginwidth="0" marginheight="0"    bgproperties="fixed"  >
<form name="<%=self%>" method="post" action="YECE0101.foregnd.asp">
<INPUT TYPE=HIDDEN NAME="UID" VALUE=<%=SESSION("NETUSER")%>>
<INPUT TYPE=hidden NAME=TotalPage VALUE="<%=TotalPage%>">
<INPUT TYPE=hidden NAME=CurrentPage VALUE="<%=CurrentPage%>">
<INPUT TYPE=hidden NAME=RecordInDB VALUE="<%=RecordInDB%>">
<INPUT TYPE=hidden NAME=PageRec VALUE="<%=PageRec%>">
<INPUT TYPE=hidden NAME=gTotalPage VALUE="<%=gTotalPage%>"> 	
<INPUT TYPE=hidden NAME=YYMM VALUE="<%=YYMM%>"> 	
<INPUT TYPE=hidden NAME=MMDAYS VALUE="<%=MMDAYS%>">
<INPUT TYPE=hidden NAME=F_country VALUE="<%=F_country%>">
<INPUT TYPE=hidden NAME=F_whsno VALUE="<%=F_whsno%>">
 
	<table border=0 style="height:30px"><tr><td>&nbsp;</td></tr></table>

	<table width="100%" BORDER=0 align=center cellpadding=0 cellspacing=0 >
		<tr>		
			<TD>
				<table border=0  cellpadding=3 cellspacing=3 >
					<tr>
						<td align="right">計薪年月</td>
						<td><input type="text" style="width:100px" name=calcYM value="<%=YYMM%>" maxlength=6></td> 
			　　		<td><%if F_country="CN" then%><a href="salary_CN_ver200801.pdf" target="_blank"><font color=blue>**查看薪資結構表**</font>(*.pdf 需安裝 Acrobat)</a><%end if%></td>
					</tr>
				</table>									
			</TD>									 
		</tr>
		<tr>
			<td align="center">
				<table id="myTableGrid" width="98%">					
					<TR HEIGHT=25 BGCOLOR="LightGrey" class="txt">
						<TD ROWSPAN=2 width=30 align=center>項次<br>STT</TD>
						<TD width=30 align=center>廠別</TD>
						<TD align=center>工號<br>Số thẻ</TD> 		
						<TD COLSPAN=3  >員工姓名<br>Họ tên</TD>  		
						<td align=center><%if F_country="CN" then %><%ELSE%>理論代碼<br>Mã bậc lương<%END IF%></td>
						<td align=center>到職日期<BR>Ngày vào xưởng</td>
						<td align=center>離職日期<BR>Ngày thôi việc</td>
						<TD align=center><%if F_country="CN" then %><%ELSE%>保險日期<br>Ngày bảo hiểm<%END IF%></TD>
						<%if instr(f_country,"VN")>0 then mycol=4 else mycol=3 %>
						<td align=center colspan="<%=mycol%>">備註<br>Ghi chú</td>
						<%if F_whsno="DN" then %>
							<td align=center>CB</td>			
						<%end if%>	
					</TR>
					<tr BGCOLOR="LightGrey"  HEIGHT=25  class="txt"> 
						<TD width=30 align=center>立帳</TD>
						<TD align=center title="Mã lương">薪資代碼<br>Mã lương</TD>
						<TD align=center title="BB">基薪<br>Cơ bản</TD>
						<TD align=center title="Mã chức vụ">職專<br>Mã chức vụ</TD> 			
						<TD align=center title="CV">職務加給<br>PC chức vụ</TD>	
						<TD align=center title="PHU">電話津貼<br>PC điện thoại</TD>
						<%if instr(f_country,"VN")>0 then %>
							<td align=center title="BB"><font color="blue">補薪<br>Bù lương</font></td>
						<%end if%>
						<td align=center title="NN">燃油津貼<br>PC xăng xe</td>
						<td align=center title="KT">技術加給<BR>Kỷ thuật</td>
						<td align=center title="MT">環境<br>Môi trường</td>						
						<td align=center title="TTKH">住房支持<br>Hỗ trợ nhà ở</td>						
						<td align=center title="QC">全勤獎金<br>Chuyên cần</td>
						<td align=center>薪資合計<br>Tổng Cộng</td>
						<%if F_whsno="DN" then %>		
							<td align=center>DIFF</td>
						<%end if%>	
					</tr> 
					<% Response.Flush %>
					<%for CurrentRow = 1 to PageRec
						IF CurrentRow MOD 2 = 0 THEN 
							WKCOLOR="LavenderBlush"
						ELSE
							WKCOLOR="#DFEFFF"
						END IF 	 
						'if tmpRec(CurrentPage, CurrentRow, 1) <> "" then 
					%>
					<TR BGCOLOR=<%=WKCOLOR%> class="txt"> 		
						<TD ROWSPAN=2 ALIGN=CENTER >
						<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then %><%=(CURRENTROW)+((CURRENTPAGE-1)*10)%><%END IF %>
						</TD>
						<TD  align="center"><%=tmpRec(CurrentPage, CurrentRow, 7)%></td>
						<TD  >
							<a href='vbscript:oktest(<%=tmpRec(CurrentPage, CurrentRow, 17)%>)'>
								<font color="<%=tmpRec(CurrentPage, CurrentRow, 35)%>"><%=tmpRec(CurrentPage, CurrentRow, 1)%></font>
							</a>
							<input type=hidden name=empid value="<%=tmpRec(CurrentPage, CurrentRow, 1)%>">
							<input type=hidden name="empautoid" value="<%=tmpRec(CurrentPage, CurrentRow, 17)%>">
							<input type=hidden name="COUNTRY" value="<%=tmpRec(CurrentPage, CurrentRow, 4)%>"> 
							<input type=hidden name="whsno" value="<%=tmpRec(CurrentPage, CurrentRow, 7)%>">	
						</TD> 		
						<TD COLSPAN=3>
							<a href='vbscript:oktest(<%=tmpRec(CurrentPage, CurrentRow, 17)%>)'>
								<font color="<%=tmpRec(CurrentPage, CurrentRow, 35)%>"><%=tmpRec(CurrentPage, CurrentRow, 42)%>&nbsp;<%=tmpRec(CurrentPage, CurrentRow, 2)%>
								<font class=txt8><%=tmpRec(CurrentPage, CurrentRow, 3)%></font></font>
							</a>
						</TD>
						<TD align="right" class="txt8">
							<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>	
								<INPUT TYPE="text"  NAME="lncode"  CLASS='INPUTBOX8' READONLY  VALUE="<%=tmpRec(CurrentPage, CurrentRow, 47)%>" style="width:40%;align:right;BACKGROUND-COLOR:Gainsboro"  title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 實薪">
								
								<INPUT TYPE=HIDDEN  NAME=HHMOENY  CLASS='INPUTBOX8' READONLY  SIZE=10 VALUE="<%=tmpRec(CurrentPage, CurrentRow, 31)%>" STYLE="TEXT-ALIGN:RIGHT;BACKGROUND-COLOR:Gainsboro" onblur="DATACHG(<%=CURRENTROW-1%>)" title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 實薪">
							<%ELSE%>
								<INPUT TYPE=HIDDEN NAME=HHMOENY >
								<INPUT TYPE=HIDDEN NAME="lncode" >
							<%END IF%>	 		
						</TD> 
						
						<TD  ALIGN=CENTER >
							<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>	
								<INPUT type="text" style="width:100%;TEXT-ALIGN:CENTER;BACKGROUND-COLOR:Gainsboro" NAME=INDAT CLASS='INPUTBOX8' READONLY  VALUE="<%=(right(tmpRec(CurrentPage, CurrentRow, 5),8))%>(<%=tmpRec(CurrentPage, CurrentRow, 37)%>)" onblur="DATACHG(<%=CURRENTROW-1%>)" title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 到職日">
							<%ELSE%>
								<INPUT TYPE=HIDDEN NAME=INDAT >
							<%END IF%> 		
						</TD>	
						<TD  ALIGN=CENTER >
							<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>	
								<INPUT type="text" style="width:100%;TEXT-ALIGN:CENTER;BACKGROUND-COLOR:Gainsboro" NAME=OUTDAT CLASS='INPUTBOX8' READONLY  VALUE="<%=(tmpRec(CurrentPage, CurrentRow, 30))%>"  onblur="DATACHG(<%=CURRENTROW-1%>)" title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 離職日">
							<%ELSE%>
								<INPUT TYPE=HIDDEN NAME=OUTDAT >
							<%END IF%> 	
						</TD>
						<TD  ALIGN=right > 			
							<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>	
								<%if tmpRec(CurrentPage, CurrentRow, 4)="CN" then%>
									<%if cdbl(tmpRec(CurrentPage, CurrentRow, 26))<>cdbl(tmpRec(CurrentPage, CurrentRow, 38)) then%>
										<font class=txt8 color=red ><%=tmpRec(CurrentPage, CurrentRow, 38)%></font>
									<%end if%>
								<%end if%>
								<INPUT style="width:100%;TEXT-ALIGN:CENTER;BACKGROUND-COLOR:Gainsboro" <%if tmpRec(CurrentPage, CurrentRow, 4)<>"VN" then%>TYPE=HIDDEN <% ELSE %>type="text"<%end if%> NAME=BHDAT CLASS='INPUTBOX8' READONLY VALUE="<%=tmpRec(CurrentPage, CurrentRow, 28)%>"  onblur="DATACHG(<%=CURRENTROW-1%>)" title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> ">
							<%ELSE%>
								<INPUT TYPE=HIDDEN NAME=BHDAT >
							<%END IF%> 	
						</TD>  		
						<TD colspan=3>
							<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
								<INPUT type="text" style="width:100%" NAME=memo class="inputbox" value="<%=tmpRec(CurrentPage, CurrentRow, 34)%>" onchange="DATACHG(<%=CURRENTROW-1%>)"> 				
							<%else%>	
								<INPUT TYPE=HIDDEN NAME=memo >
							<%end if%>
						</TD>		
						<td>	
						<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>	
							<%if F_whsno="DN" then %>
								<INPUT NAME=CB class="inputbox8r" value="<%=tmpRec(CurrentPage, CurrentRow, 40)%>" readonly  STYLE="TEXT-ALIGN:RIGHT;BACKGROUND-COLOR:LIGHTYELLOW"  > 				
							<%else %>
								<INPUT NAME=CB size=10 type="hidden">
							<%end if %>
						<%else %>
							<INPUT NAME=CB size=10 type="hidden">
						<%end if %>	
						</td>									
					</TR>
					<TR BGCOLOR=<%=WKCOLOR%> class="txt">
						<TD  align="center"><%=tmpRec(CurrentPage, CurrentRow, 45)%></td>
						<TD ALIGN=RIGHT > 			 			
							<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  
									if tmpRec(CurrentPage, CurrentRow, 4)="VN" or tmpRec(CurrentPage, CurrentRow, 4)="CT" then 
										SQL="SELECT ''  as sys_value, * FROM empsalarybasic WHERE FUNC='AA' and country='VN' and bwhsno='"& tmpRec(CurrentPage, CurrentRow,7) &"'  AND Bonus>0 ORDER BY right(code,1), CODE " 
									else
										SQL="select a.*, b.sys_value from "&_
												"( SELECT * FROM empsalarybasic WHERE FUNC='AA'  and bwhsno='"& tmpRec(CurrentPage, CurrentRow,7) &"' and country='"& tmpRec(CurrentPage, CurrentRow,4) &"'  ) a "&_
												"LEFT JOIN ( SELECT * FROM  BAsicCODE WHERE  FUNC='LEV' ) B ON B.SYs_type = a.job  "&_
												"order by right(code,1),  a.code "
									end if	
									
									'RESPONSE.WRITE SQL 
							%>	
								<select name=BBCODE  class="txt8" style="width:100%" onchange="bbcodechg(<%=currentrow-1%>)">					 			
									<option value="0" <%IF cdbl(trim(tmpRec(CurrentPage, CurrentRow, 20)))=0 THEN %> SELECTED <%END IF%> >0</option>
									<%					
									
									SET RST = CONN.EXECUTE(SQL)
									WHILE NOT RST.EOF  
										if tmpRec(CurrentPage, CurrentRow,4) ="VN" or tmpRec(CurrentPage, CurrentRow,4) ="CT" or tmpRec(CurrentPage, CurrentRow,4)="TW" or tmpRec(CurrentPage, CurrentRow,4)="MA" then 
											showCode =rst("code") 
										else 
											showCode =rst("code")  							
										end if	
									%>
									<option value="<%=RST("CODE")%>" <%IF trim(RST("CODE"))=trim(tmpRec(CurrentPage, CurrentRow, 43)) THEN %> SELECTED <%END IF%> ><%=showcode%>-<%=RST("bonus")%></option>
									<%
									RST.MOVENEXT
									WEND 
									%>
								</SELECT>
								<%SET RST=NOTHING %>
							<%else%>
								<input type=hidden name=BBCODE >	
							<%end if %>
							 
						</TD>
						<TD ALIGN=RIGHT>
							<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
								<INPUT type="text" NAME=BB CLASS='INPUTBOX8' VALUE="<%=tmpRec(CurrentPage, CurrentRow, 20)%>" STYLE="width:100%;TEXT-ALIGN:RIGHT;BACKGROUND-COLOR:LIGHTYELLOW"  READONLY title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 資本薪資">	 			
							<%else%>
								<input type=hidden name=BB >	
							<%end if %>	
						</TD>
						<TD ALIGN=RIGHT ><!--職等--> 			
							<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
								<select name=F1_JOB  class="txt8" style="width:100%" ONCHANGE="JOBCHG(<%=CURRENTROW-1%>)">			
									<option value="" <%IF  trim(tmpRec(CurrentPage, CurrentRow, 6))="" THEN %> SELECTED <%END IF%>>---</option>	
									<%SQL="SELECT * FROM BASICCODE WHERE FUNC='LEV'  ORDER BY SYS_TYPE "
									SET RST = CONN.EXECUTE(SQL)
									WHILE NOT RST.EOF  
									%>
									<option value="<%=RST("SYS_TYPE")%>" <%IF RST("SYS_TYPE")=trim(tmpRec(CurrentPage, CurrentRow, 6)) THEN %> SELECTED <%END IF%> ><%=RST("SYS_TYPE")%>-<%=RST("SYS_VALUE")%></option>				 
									<%
									RST.MOVENEXT
									WEND 
									%>
								</SELECT>
								<%SET RST=NOTHING %>
							<%else%>
								<input type=hidden name=F1_JOB >	
							<%end if %>
						</TD>
						 <TD  ALIGN=RIGHT>
							<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
								<INPUT type="text" NAME=CV CLASS='INPUTBOX8' VALUE="<%=tmpRec(CurrentPage, CurrentRow, 22)%>"  STYLE="width:100%;TEXT-ALIGN:RIGHT" title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 職務加給"    >
								<input type=hidden name=CVCODE VALUE="<%=tmpRec(CurrentPage, CurrentRow, 21)%>" SIZE=3>
							<%else%>
								<input type=hidden name=CV >	
								<input type=hidden name=CVCODE >
							<%end if %>	
						 </TD>
						<TD  ALIGN=RIGHT>
							<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
								<INPUT type="text" NAME=PHU CLASS='INPUTBOX8' VALUE="<%=tmpRec(CurrentPage, CurrentRow, 23)%>" STYLE="width:100%;TEXT-ALIGN:RIGHT" onblur="DATACHG(<%=CURRENTROW-1%>)" title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 補助獎金(Y)" >
							<%ELSE%>	
								<INPUT TYPE=HIDDEN NAME=PHU	>
							<%END IF%>	
						</TD> 		
						<%if instr(f_country,"VN")>0 then %> 	
							<Td>
							<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
								<INPUT type="text" NAME=btien CLASS='inpt8blue' VALUE="<%=tmpRec(CurrentPage, CurrentRow, 44)%>" STYLE="width:100%;TEXT-ALIGN:RIGHT" onblur="DATACHG(<%=CURRENTROW-1%>)" title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 補薪" >
							<%ELSE%>	
								<INPUT TYPE=HIDDEN NAME=btien >
							<%END IF%>	
							</td>
						<%else%>
							<INPUT TYPE=HIDDEN NAME=btien value=0 >
						<%end if%>
						<TD  ALIGN=RIGHT>
							<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
								<INPUT type="text" NAME=NN CLASS='INPUTBOX8' VALUE="<%=tmpRec(CurrentPage, CurrentRow, 24)%>" STYLE="width:100%;TEXT-ALIGN:RIGHT" onblur="DATACHG(<%=CURRENTROW-1%>)" title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 語言加給" >
							<%ELSE%>	
								<INPUT TYPE=HIDDEN NAME=NN >
							<%END IF%>		
						</TD>
						<TD  ALIGN=RIGHT>
							<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
								<INPUT type="text" NAME=KT CLASS='INPUTBOX8' VALUE="<%=tmpRec(CurrentPage, CurrentRow, 25)%>" STYLE="width:100%;TEXT-ALIGN:RIGHT" onblur="DATACHG(<%=CURRENTROW-1%>)" title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 技術加給" >
							<%ELSE%>
								<INPUT TYPE=HIDDEN NAME=KT >
							<%END IF%>			
						</TD>
						<TD  ALIGN=RIGHT>
							<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %> 
								<INPUT type="text" NAME=MT CLASS='INPUTBOX8' VALUE="<%=tmpRec(CurrentPage, CurrentRow, 26)%>" onblur="DATACHG(<%=CURRENTROW-1%>)" title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 環境加給"  STYLE="width:100%;TEXT-ALIGN:RIGHT;" >
							<%ELSE%>
								<INPUT TYPE=HIDDEN NAME=MT >
							<%END IF%>			
						</TD>
						
							<%if tmpRec(CurrentPage, CurrentRow, 4)="TW" or tmpRec(CurrentPage, CurrentRow, 4)="MA" then 
									fcolor="X"
							  else
								fcolor=""	
							  end if	
							%>
							<%if tmpRec(CurrentPage, CurrentRow, 4)="TW" or tmpRec(CurrentPage, CurrentRow, 4)="MA"  then %>
							<td align="right">
								<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>			 	
									<INPUT type="text" NAME="wp" CLASS='INPUTBOX8' VALUE="<%=tmpRec(CurrentPage, CurrentRow, 19)%>" STYLE="width:100%;TEXT-ALIGN:RIGHT;<%if fcolor<>"" then%>border: 1px solid #ff6633 ; <%end if%>" onblur="DATACHG(<%=CURRENTROW-1%>)" title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> Willpower">					
								<%ELSE%>
									<INPUT TYPE=HIDDEN NAME="wp" >
								<%END IF%>			
								<INPUT TYPE=HIDDEN NAME="TTKH"  value="0">				
							</td>	
							<td align="right">
								<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>			 	
									<INPUT type="text" NAME="wpbtien" CLASS='INPUTBOX8' VALUE="<%=tmpRec(CurrentPage, CurrentRow, 46)%>" STYLE="width:100%;TEXT-ALIGN:RIGHT;<%if fcolor<>"" then%>border: 1px solid #ff6633 ; <%end if%>" onblur="DATACHG(<%=CURRENTROW-1%>)" title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> Willpower">					
								<%ELSE%>
									<INPUT TYPE=HIDDEN NAME="wpbtien" value="0">
								<%END IF%>							
							</td>	
							<%else%>
							<td align="right">
								<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>	
									<INPUT type="text" NAME=TTKH CLASS='INPUTBOX8' VALUE="<%=tmpRec(CurrentPage, CurrentRow, 27)%>" STYLE="width:100%;TEXT-ALIGN:RIGHT;<%if fcolor<>"" then%>border: 1px solid #ff6633 ; <%end if%>" onblur="DATACHG(<%=CURRENTROW-1%>)" title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 其他加給">
								<%ELSE%>
									<INPUT TYPE=HIDDEN NAME=TTKH >
								<%END IF%>			
								<INPUT TYPE=HIDDEN NAME="wp"  value="<%=tmpRec(CurrentPage, CurrentRow, 19)%>">
								<INPUT TYPE=HIDDEN NAME="wpbtien"  value="0">
							</td>	
							<%end if%>	
						
						<TD>
							<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>	
								<INPUT type="text" NAME=QC CLASS='INPUTBOX8' VALUE="<%=tmpRec(CurrentPage, CurrentRow, 32)%>" STYLE="width:100%;TEXT-ALIGN:RIGHT;BACKGROUND-COLOR:LIGHTYELLOW" READONLY  title="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))&" "&trim(tmpRec(CurrentPage, CurrentRow, 3))%> 全勤">
							<%ELSE%>
								<INPUT TYPE=HIDDEN NAME=QC >
							<%END IF%>	
						</TD>
						<TD>
							<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
								<INPUT type="text" NAME=totamt CLASS='INPUTBOX8' VALUE="<%=formatnumber(tmpRec(CurrentPage, CurrentRow, 33),0)%>" STYLE="width:100%;TEXT-ALIGN:RIGHT; color:darkred;BACKGROUND-COLOR:LIGHTYELLOW"   >
							<%ELSE%>
								<INPUT TYPE=HIDDEN NAME=totamt >
							<%END IF%>	
						</td>
						<Td>		
						<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>	
							<%if F_whsno="DN" then %>
								<INPUT type="hidden" NAME=CBdiff class="inputbox8r" value="<%=tmpRec(CurrentPage, CurrentRow, 41)%>"  readonly STYLE="width:100%;TEXT-ALIGN:RIGHT;BACKGROUND-COLOR:LIGHTYELLOW;color:red"  > 				
							<%else %>
								<INPUT NAME=CBdiff size=10 type="hidden">
							<%end if %>
						<%else%>	
							<INPUT NAME=CBdiff size=10 type="hidden">
						<%end if %>
						</TD>  		
					</TR>
					<%if tmpRec(CurrentPage, CurrentRow, 4) <>"VN" and ( tmpRec(CurrentPage, CurrentRow, 6)="EV0" and tmpRec(CurrentPage, CurrentRow, 37)>3 ) then   %>
					<tr  BGCOLOR=<%=WKCOLOR%>>
						<td ></td>
						<td colspan=11><font color="Red">已過試用期，請調整職務</font></td>
					</tr>
					<%end if %>
					<%next%>
				</TABLE>
			</td>
		</tr>
		<tr>
			<td align="center">
				<input type=hidden name="empid">
				<input type=hidden name="empautoid"> 
				<INPUT TYPE=HIDDEN NAME=INDAT > 
				<INPUT TYPE=HIDDEN NAME=INDAT > 
				<INPUT TYPE=HIDDEN NAME=OUTDAT >
				<INPUT TYPE=HIDDEN NAME=BHDAT > 
				<INPUT TYPE=HIDDEN NAME=GTDAT > 
				<INPUT TYPE=HIDDEN NAME=HHMOENY > 
				<INPUT TYPE=HIDDEN NAME=lncode >
				<input type=hidden name="BBCODE">
				<input type=hidden name="BB">
				<input type=hidden name="F1_JOB">
				<input type=hidden name="CV">
				<input type=hidden name="CVCODE">
				<input type=hidden name="PHU">
				<input type=hidden name="NN">
				<input type=hidden name="KT">
				<input type=hidden name="MT">
				<input type=hidden name="TTKH">
				<input type=hidden name="wp">
				<INPUT TYPE=HIDDEN NAME=QC > 
				<INPUT TYPE=HIDDEN NAME=totamt > 
				<INPUT TYPE=HIDDEN NAME=memo > 
				<INPUT TYPE=HIDDEN NAME=cb > 
				<INPUT TYPE=HIDDEN NAME=cbdiff > 

				<table class="txt">
					<tr>
						<td align="CENTER" height=40 WIDTH=75%>    
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
						<TD WIDTH=25% ALIGN=RIGHT>		
							<input type="BUTTON" name="send" value="確　認" class="btn btn-sm btn-danger" ONCLICK="GO()">
							<input type="BUTTON" name="send" value="取　消" class="btn btn-sm btn-outline-secondary" onclick="clr()">
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
function BACKMAIN() 	
	open "../main.asp" , "_self"
end function   

function clr()
	open "<%=self%>.fore.asp" , "_self"
end function 

function go()	
	<%=self%>.action="<%=self%>.upd.asp"  
	<%=self%>.submit()
end function 

function oktest(N)
	tp=<%=self%>.totalpage.value 
	cp=<%=self%>.CurrentPage.value 
	rc=<%=self%>.RecordInDB.value 
	open "empfile.show.asp?empautoid="& N , "_blank" , "top=10, left=10, width=550, scrollbars=yes" 
end function 

FUNCTION BBCODECHG(INDEX)
	
	whsno = <%=self%>.whsno(index).value 
	'whsno = <%=self%>.whsno.value
	'alert whsno
	codestr=<%=self%>.bbcode(index).value 
	daystr=<%=self%>.MMDAYS.value
	<%=self%>.lncode(index).value=codestr
	'alert whsno
	'alert codestr
	'alert daystr	
	open "<%=self%>.back.asp?ftype=A&index="& index &"&CurrentPage="& <%=CurrentPage%> & _
		 "&days=" & daystr & "&code=" &	codestr &"&whsno="& whsno&"&lncode="&codestr , "Back" 		 
	'DATACHG(INDEX)	  
	'PARENT.BEST.COLS="70%,30%"	 	
END FUNCTION 

FUNCTION JOBCHG(INDEX)	
	whsno = <%=self%>.whsno(index).value 
	COUNTRYstr = <%=self%>.COUNTRY(index).value 	
	codestr=<%=self%>.F1_JOB(index).value 
	daystr=<%=self%>.MMDAYS.value 	
'	if COUNTRYstr="VN" then 
		open "<%=self%>.back.asp?ftype=B&index="& index &"&CurrentPage="& <%=CurrentPage%> & _
			 "&days=" &daystr & "&whsno="& whsno & "&code=" &	codestr , "Back"
'	end if 		 
	PARENT.BEST.COLS="100%,0%"	 	 
	'DATACHG(INDEX)	  
END FUNCTION 

FUNCTION DATACHG(INDEX) 	 	
	whsno = <%=self%>.whsno(index).value 
	if isnumeric(<%=SELF%>.PHU(INDEX).VALUE)=false then 
		alert "請輸入數字!!"
		<%=self%>.phu(index).focus()
		<%=self%>.phu(index).value=0
		<%=self%>.phu(index).select()
		exit FUNCTION 
	end if 	
	if isnumeric(<%=SELF%>.NN(INDEX).VALUE)=false then 
		alert "請輸入數字!!"
		<%=self%>.NN(index).value=0 		
		<%=self%>.NN(index).focus()
		<%=self%>.NN(index).select()
		exit FUNCTION 
	end if 	
	if isnumeric(<%=SELF%>.KT(INDEX).VALUE)=false then 
		alert "請輸入數字!!"
		<%=self%>.KT(index).value=0 		
		<%=self%>.KT(index).focus()
		<%=self%>.KT(index).select()
		exit FUNCTION 
	end if 	
	if isnumeric(<%=SELF%>.MT(INDEX).VALUE)=false then 
		alert "請輸入數字!!"
		<%=self%>.MT(index).value=0 		
		<%=self%>.MT(index).focus()
		<%=self%>.MT(index).select()
		exit FUNCTION 
	end if 	
	if isnumeric(<%=SELF%>.TTKH(INDEX).VALUE)=false then 
		alert "請輸入數字!!"
		<%=self%>.TTKH(index).value=0 		
		<%=self%>.TTKH(index).focus()
		<%=self%>.TTKH(index).select()
		exit FUNCTION 
	end if 	 
	if isnumeric(<%=SELF%>.wp(INDEX).VALUE)=false then 
		alert "請輸入數字!!"
		<%=self%>.wp(index).value=0 		
		<%=self%>.wp(index).focus()
		<%=self%>.wp(index).select()
		exit FUNCTION 
	end if 	 	
 
	 
	TTM = ( cdbl(<%=self%>.bb(index).value) + cdbl(<%=self%>.CV(index).value) + cdbl(<%=self%>.PHU(index).value) ) 
	if TTM mod (26*8)<>0 then 
		TTMH = FIX (CDBL(TTM)/26/8 ) +1   '時薪
	else
		TTMH = FIX (CDBL(TTM)/26/8 )    '時薪
	end if 
	'alert  TTMH 
	'<%=self%>.HHMOENY(index).value = TTMH 
	
	CODESTR01 = <%=SELF%>.PHU(INDEX).VALUE
	CODESTR02 = <%=SELF%>.NN(INDEX).VALUE
	CODESTR03 = <%=SELF%>.KT(INDEX).VALUE
	CODESTR04 = <%=SELF%>.MT(INDEX).VALUE
	CODESTR05 = <%=SELF%>.TTKH(INDEX).VALUE
	CODESTR06 = <%=SELF%>.QC(INDEX).VALUE
	CODESTR07 = <%=SELF%>.BB(INDEX).VALUE
	CODESTR08 = <%=SELF%>.CV(INDEX).VALUE		
	CODESTR09 =  (escape(trim(<%=SELF%>.memo(INDEX).VALUE)))
	CODESTR10 =  (escape(trim(<%=SELF%>.wp(INDEX).VALUE)))
	CODESTR11 =  (escape(trim(<%=SELF%>.btien(INDEX).VALUE)))
	CODESTR10wp =  (escape(trim(<%=SELF%>.wpbtien(INDEX).VALUE)))
	'daystr=<%=self%>.MMDAYS.value  
	'ALERT CODESTR09
	'ALERT CODESTR03
	
	open "<%=self%>.back.asp?ftype=CDATACHG&index="&index &"&CurrentPage="& <%=CurrentPage%> & _
		 "&CODESTR01="& CODESTR01 &_
		 "&CODESTR02="& CODESTR02 &_
		 "&CODESTR03="& CODESTR03 &_
		 "&CODESTR04="& CODESTR04 &_
		 "&CODESTR05="& CODESTR05 &_
		 "&CODESTR06="& CODESTR06 &_
		 "&CODESTR07="& CODESTR07 &_	
		 "&CODESTR08="& CODESTR08 &_	
		 "&CODESTR09="& CODESTR09 &_
		 "&CODESTR10="& CODESTR10 &_
		 "&CODESTR11="& CODESTR11 &_
		 "&CODESTR10wp="& CODESTR10wp &_
		 "&whsno="& whsno  , "Back"  
		 
	'PARENT.BEST.COLS="80%,20%"	 
	
END FUNCTION  

function view1(index) 
	 
	yymmstr = <%=self%>.yymm.value 
	'alert yymmstr
	empidstr = <%=self%>.empid(index).value 
	idstr=  <%=self%>.empautoid(index).value 
	open "empworkb.fore.asp?yymm=" & yymmstr &"&EMPID=" & empidstr &"&empautoid=" & idstr , "_blank" , "top=10, left=10, scrollbars=yes" 
end function 
	
</script>
<%response.end%>

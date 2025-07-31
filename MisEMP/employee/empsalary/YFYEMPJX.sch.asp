<%@LANGUAGE=VBSCRIPT CODEPAGE=950%> 
<!-- #include file = "../../GetSQLServerConnection.fun" --> 
<!-- #include file="../../ADOINC.inc" -->
<!-- #include file="../../Include/func.inc" -->
<%

self="JXsch"  

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

NNY=year(date())
NDY=year(date())+1   

jxym= request("JXYM")
S1=request("S1")
G1=request("G1")
W1=request("W1")
Set conn = GetSQLServerConnection()	  
Set rs = Server.CreateObject("ADODB.Recordset")   
gTotalPage = 1
TotalPage = 1
PageRec = 1    'number of records per page
TableRec = 15    'number of fields per record      

sql="select b.sys_value,  c.sys_value as zstr,a.* from  "&_
	"(SELECT* FROM YFYMJIXO where jxwhsno like '"&W1&"%' and JXYM='"& JXYM  &"' and groupid like '"& G1 &"%' and  isnull(shift,'') like '%"& S1 &"' ) a  "&_
	"left join ( select * from basicCode where func='groupid' ) b on b.sys_type=a.groupid "&_
	"left join ( select * from basicCode where func='zuno' ) c on c.sys_type=isnull(a.zuno,'') "&_
	"order by case when groupid='A065' then 'A000' else groupid end  , len(shift) desc , shift,  a.autoid , stt "   
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
	
	Redim tmpRec(TotalPage, PageRec, TableRec)   'Array
	
	for i = 1 to TotalPage
		for j = 1 to PageRec 
	 		if not rs.EOF then
			 	tmpRec(i, j, 0) = "no"
				tmpRec(i, j, 1) = trim(rs("autoid"))
				tmpRec(i, j, 2) = trim(rs("jxym"))
				tmpRec(i, j, 3) = trim(rs("groupid"))
				tmpRec(i, j, 4) = rs("shift")
				tmpRec(i, j, 5) = rs("stt")  
				tmpRec(i, j, 6) = rs("descp")
				tmpRec(i, j, 7) = rs("HXSL")
				tmpRec(i, j, 8) = rs("HESO")
				tmpRec(i, j, 9)	=RS("PER")
				tmpRec(i, j, 10)=RS("sys_value")
				tmpRec(i, j, 11)=RS("zuno")
				tmpRec(i, j, 12)=RS("zstr")
				tmpRec(i, j, 13)=RS("jxwhsno")
				rs.movenext 
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
	Session("empjxedit") = tmpRec 
else 
	TotalPage = cint(request("TotalPage"))
	'StoreToSession()
	CurrentPage = cint(request("CurrentPage"))
	RecordInDB  = REQUEST("RecordInDB")
	tmpRec = Session("empjxedit")
	COUNTRY = REQUEST("COUNTRY")

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
if request("w1")="" then 
	if instr(session("vnlogIP"),"168")>0 then 
			w1="LA" 
	elseif instr(session("vnlogIP"),"169")>0 then 	
			w1="DN" 
	elseif instr(session("vnlogIP"),"47")>0 then 	
			w1="BC" 
	end if 	
end if 
%>

<html> 
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="refresh">
<link rel="stylesheet" href="../../Include/style.css" type="text/css">
<link rel="stylesheet" href="../../Include/style2.css" type="text/css">
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!-- 
function f()
	<%=self%>.JXYM.focus()		
end function     

function SCHDATA()
	<%=self%>.TotalPage.value="0" 	
	<%=self%>.action="yfyempjx.sch.asp"
	<%=self%>.submit
end function 
-->
</SCRIPT>   
</head> 
<body  topmargin="5" leftmargin="5"  marginwidth="0" marginheight="0"  onkeydown="enterto()" onload='f()' >
<form name="<%=self%>" method="post" action="yfyempjx.foreEdit.asp">
<input type=hidden name="NNY" value="<%=NNY%>">
<input type=hidden name="NDY" value="<%=NDY%>"> 
<INPUT TYPE=hidden NAME=TotalPage VALUE="<%=TotalPage%>">
<INPUT TYPE=hidden NAME=CurrentPage VALUE="<%=CurrentPage%>">
<INPUT TYPE=hidden NAME=PageRec VALUE="<%=PageRec%>"> 

<table width="460" border="0" cellspacing="0" cellpadding="0">
	<tr><TD>
	<img border="0" src="../../image/icon.gif" align="absmiddle">
	績效獎金計算</TD></tr>
</table>
<hr size=0	style='border: 1px dotted #999999;' align=left width=500>
<table width=500 class=txt9>
 	<tr>
 		<td><a href="yfyempjx.Fore.asp">績效新增作業</a></td> 
 		<td><a href="yfyempjx.ForeEDIT.asp">績效修改作業</A></td>
 		<td><a href="yfyempjx.sch.asp">績效查詢作業</a></td>
 	</tr>
 	<tr><td colspan=3><hr size=0	style='border: 1px dotted #999999;' align=left ></td></tr>
</table>		
<table width=550 border=0 ><tr><td >
		<TABLE WIDTH=500 CLASS=TXT9>
	 		<TR>
	 			<TD width=50>廠別:</TD>
				<TD ALIGN=CENTER>
	 				<select name=w1 class="txt8" style="width:60" onchange="SCHDATA()">
	 				<option value="">---</option>
	 				<%sqlx="select *from basicCode where func='whsno' order by sys_type " 
	 				  set rst=conn.execute(Sqlx)
	 				  while not rst.eof 
	 				%>	
	 				<option value="<%=rst("sys_type")%>" <%if W1=rst("sys_type") then%>selected<%end if%>><%=rst("sys_type")%>-<%=rst("sys_value")%></option>
	 				<%rst.movenext
	 				wend
	 				rst.close
	 				set rst=nothing
	 				%>
	 				</select>
	 			</TD>
				<TD width=70>績效年月:</TD>
	 			<TD width=80 ><INPUT NAME=JXYM VALUE="<%=JXYM%>" SIZE=8 CLASS=INPUTBOX></TD> 	 			
	 			<TD width=50>單位:</TD>
	 			<TD width=200 >	
	 				<select name=g1 class=inputbox onchange="SCHDATA()" style="width:80" >
	 					<option value="">全部</option>
	 					<%sql="select * from basiccode where func='groupid' and  left(sys_type,3)='A06' or  sys_type in ('A051', 'A059', 'A033') order by case when sys_type='A065' then 'a000' else sys_type end " 
	 					  set rst=conn.execute(sql) 
	 					  while not rst.eof 
	 					%>
	 					<option value="<%=rst("sys_type")%>" <%if G1=rst("sys_type") then%>selected<%end if%>><%=rst("sys_type")%><%=rst("sys_value")%></option>
	 					<%rst.movenext
	 					wend
	 					%>
	 				<select>
	 				<select name=s1 class=inputbox onchange="SCHDATA()">
	 					<option value="">全部</option>
	 					<option value="ALL" <%if S1="ALL" then%>selected<%end if%>>常日班</option>
	 					<option value="A" <%if S1="A" then%>selected<%end if%>>A</option>
	 					<option value="B" <%if S1="B" then%>selected<%end if%> >B</option>
	 				</select>
	 			</TD>
	 			<td>
	 				<input type=button name=btm class=button value="查詢"  onclick="SCHDATA()" onkeydown="SCHDATA()">
	 			</td> 	 			
	 		</TR>
	 	</TABLE>		
	 	<hr size=0	style='border: 1px dotted #999999;' align=left width=500> 
	 	<TABLE WIDTH=610 CLASS=TXT9 BGCOLOR="#CCCCCC" BORDER=0 border="1" cellspacing="1" CLASS=TXT9 ALIGN=CENTER>
	 		<TR BGCOLOR="#FFF278">
				<TD WIDTH=50 HEIGHT=22 ALIGN=CENTER>STT</TD>
				<TD WIDTH=50 HEIGHT=22 ALIGN=CENTER>廠別</TD>
				<TD WIDTH=60 HEIGHT=22 ALIGN=CENTER>單位</TD>
				<TD WIDTH=50 HEIGHT=22 ALIGN=CENTER>班別</TD>
	 			<TD WIDTH=30 HEIGHT=22 ALIGN=CENTER>STT</TD>
	 			<TD WIDTH=140 ALIGN=CENTER>說明</TD>
	 			<TD WIDTH=100 ALIGN=CENTER>實績</TD>
	 			<TD WIDTH=60 ALIGN=CENTER>係數</TD>
	 			<TD WIDTH=60 ALIGN=CENTER>比例</TD>	 			
	 		</TR>
	 		<%for CurrentRow = 1 to PageRec
			IF CurrentRow MOD 2 = 0 THEN
				WKCOLOR="LavenderBlush"
			ELSE
				WKCOLOR="#DFEFFF"
			END IF
			
			'if tmpRec(CurrentPage, CurrentRow, 1) <> "" then
			%>
			<%if CurrentRow>1 and ( tmpRec(CurrentPage, CurrentRow, 10)<>tmpRec(CurrentPage, CurrentRow-1, 10) or  tmpRec(CurrentPage, CurrentRow, 4)<>tmpRec(CurrentPage, CurrentRow-1, 4) )then %>
	 		<tr  BGCOLOR="#FFFFFF" >
	 			<td colspan=8 height=15>	 			
	 				<hr style="border-style: dotted; border-width: 3px; padding: 0" noshade color="#000000" align="left" size="1">
	 			</td>
	 		</tr>
	 		<%end if%>
	 		<TR BGCOLOR="#FFFFFF" >
	 			<TD ALIGN=CENTER><%=CurrentRow%></TD>
	 			<TD ALIGN=CENTER>
	 				<INPUT NAME=whsno  CLASS="readonly2" READONLY  SIZE=4  value="<%=tmpRec(CurrentPage, CurrentRow, 13)%>" style="text-align:center">
	 			</TD>
				<TD ALIGN=CENTER>
	 				<INPUT NAME=groupid CLASS="readonly2" READONLY SIZE=8 value="<%=tmpRec(CurrentPage, CurrentRow, 10)%>">
	 			</TD>
	 			<TD ALIGN=CENTER>
	 				<INPUT NAME=shift  CLASS="readonly2" READONLY  SIZE=5  value="<%=tmpRec(CurrentPage, CurrentRow, 4)%>" style="text-align:center">
	 			</TD>
	 			<TD HEIGHT=22 ALIGN=CENTER >
	 				<INPUT NAME="STT"  SIZE=3 CLASS="readonly2" READONLY  value="<%=tmpRec(CurrentPage, CurrentRow, 5)%>" style="text-align:center">
	 				<INPUT  type=hidden  NAME="autoid"  value="<%=tmpRec(CurrentPage, CurrentRow, 1)%>" >
	 			</TD>
	 			<TD ALIGN=CENTER>
	 				<INPUT NAME=DESCP CLASS="readonly2" READONLY SIZE=22  value="<%=tmpRec(CurrentPage, CurrentRow, 6)%>" >
	 			</TD>
	 			<TD ALIGN=CENTER>
	 				<INPUT NAME=HXSL  CLASS="readonly2" READONLY SIZE=15  value="<%=tmpRec(CurrentPage, CurrentRow, 7)%>"  style="text-align:right">
	 			</TD>
	 			<TD ALIGN=CENTER>
	 				<INPUT NAME="HESO" SIZE=6 CLASS="readonly2"  READONLY value="<%=tmpRec(CurrentPage, CurrentRow, 8)%>"  style="text-align:right">
	 			</TD>
	 			<TD ALIGN=CENTER>
	 				<INPUT NAME="PER"  SIZE=6 CLASS="readonly2" READONLY  value="<%=tmpRec(CurrentPage, CurrentRow, 9)%>"  style="text-align:right">
	 			</TD>	 			
	 		</TR>
	 		
	 		<%	 			 		
	 		NEXT  	 		
	 		%>
	 	</TABLE>	<br>	 	
</td></tr></table> 

</body>
</html>


<script language=vbs>  

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

FUNCTION GO()
	IF <%=SELF%>.JXYM.VALUE="" THEN 
		ALERT "請輸入績效年月!!"
		<%=SELF%>.JXYM.FOCUS()
		EXIT FUNCTION 
	ELSEIF <%=SELF%>.SALARYYM.VALUE="" THEN 	
		ALERT "請輸入計薪年月!!"
		<%=SELF%>.SALARYYM.FOCUS()
		EXIT FUNCTION 
	ELSEIF <%=SELF%>.GROUPID.VALUE="" THEN 
		ALERT "請輸入單位!!"
		<%=SELF%>.GROUPID.FOCUS()
		EXIT FUNCTION 
	ELSEIF <%=SELF%>.SHIFT.VALUE="" THEN 
		ALERT "請輸入班別!!"
		<%=SELF%>.SHIFT.FOCUS()
		EXIT FUNCTION
	ELSE
		<%=SELF%>.ACTION="<%=SELF%>.UPD.ASP"
		<%=SELF%>.SUBMIT()		
	END IF 
END FUNCTION  

FUNCTION COPYDATA()
	IF <%=SELF%>.JXYM.VALUE="" THEN 
		ALERT "請輸入欲複製的績效年月"
		<%=SELF%>.JXYM.FOCUS()
		EXIT FUNCTION 
	END IF 
	<%=SELF%>.ACTION="<%=SELF%>.NEW.ASP"
	<%=SELF%>.SUBMIT()		
END FUNCTION 
</script> 
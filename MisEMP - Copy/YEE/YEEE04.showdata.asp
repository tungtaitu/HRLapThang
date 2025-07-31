<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<!-- #include file = "../GetSQLServerConnection.fun" --> 
<!-- #include file="../ADOINC.inc" --> 

<%
'on error resume next   
session.codepage="65001"
SELF = "yeee04B"
 
Set conn = GetSQLServerConnection()	  
Set rs = Server.CreateObject("ADODB.Recordset")   
DAT1 = REQUEST("DAT1")
DAT2 = REQUEST("DAT2")
whsno = trim(request("whsno"))
groupid = trim(request("groupid"))
country = trim(request("country"))  
QUERYX = trim(request("empid1"))  

unitno = trim(request("unitno"))
zuno = trim(request("zuno"))
job = trim(request("job")) 
jb = trim(request("jb")) 
ym1= trim(request("ym1")) 
ym2= trim(request("ym2")) 

gTotalPage = 1
PageRec = 16    'number of records per page
TableRec = 30    'number of fields per record  

if dat1="" and dat2="" and whsno="" and groupid="" and country="" and QUERYX="" then 
	sql="select * from empfile where empid='XX' "
else
	SQL="SELECT  A.JIATYPE,  CONVERT(CHAR(10), A.DATEUP, 111) DATEUP , A.TIMEUP, convert(char(10) , A.DATEDOWN , 111) datedown, "
	SQL=SQL&"A.TIMEDOWN , A.HHOUR, A.MEMO AS JIAMEMO  , a.autoid as jiaid,  B.*  , isnull(c.sys_value,'') as jia_str  , "
	SQL=SQL&"isnull(a.place,'') place, isnull(a.xid,'') xid  , a.autoid FROM   "
	SQL=SQL&"( SELECT * FROM EMPHOLIDAY   ) A  "
	SQL=SQL&"LEFT JOIN ( SELECT empid, nindat, empnam_cn, empnam_vn, gstr, wstr, zstr, jstr, whsno, groupid, job, zuno, country , outdate   FROM view_empfile ) B ON B.EMPID = A.EMPID  	 "
	SQL=SQL&"LEFT JOIN ( SELECT * FROM basicCode where func='JB'  ) c  on c.sys_type = a.JIATYPE  "
	SQL=SQL&"WHERE 1=1  " 	
	SQL=SQL&"and country like  '"& country &"%'  "
	SQL=SQL&"AND whsno like '"& whsno &"%'  and groupid like '"& groupid &"%'  " 
	SQL=SQL&"and zuno like '"& zuno &"%' and a.jiatype like '"& jb &"%' and b.empid like '%"& QUERYX &"%'  "
	
	IF DAT1<>"" and DAT2<>"" then  
	 	sql=sql& "and CONVERT(CHAR(10), A.DATEUP, 111) BETWEEN '"& DAT1 &"' AND '"& DAT2 &"' " 
	END IF 
	IF ym1<>"" and ym2<>"" then  
	 	sql=sql& "and CONVERT(CHAR(6), A.DATEUP, 112) BETWEEN '"& ym1 &"' AND '"& ym2 &"' " 
	END IF 
	SQL=SQL&"order by b.empid, A.DATEUP , a.jiaType "  
end if 	
'response.write sql 
'RESPONSE.END  
if request("TotalPage") = "" or request("TotalPage") = "0" then 
	CurrentPage = 1	
	rs.Open SQL, conn, 3, 3 
	IF NOT RS.EOF THEN 
		PageRec = rs.RecordCount 
		rs.PageSize = PageRec 
		RecordInDB = rs.RecordCount 
		TotalPage = rs.PageCount  
		gTotalPage = TotalPage 
		
		empname = rs("empnam_cn") &" "&rs("empnam_vn")
	END IF 	 

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
			tmpRec(i, j, 8) = "" 'rs("unitno")	 
			tmpRec(i, j, 9)	=RS("groupid") 
			tmpRec(i, j, 10)=RS("zuno") 				
			tmpRec(i, j, 11)=RS("wstr") 	
			tmpRec(i, j, 12)="" 'RS("ustr") 	
			tmpRec(i, j, 13)=RS("gstr") 	
			tmpRec(i, j, 14)=RS("zstr") 	
			tmpRec(i, j, 15)=RS("jstr") 	
			tmpRec(i, j, 16)="" 'RS("cstr")
			tmpRec(i, j, 17)=RS("DATEUP")
			tmpRec(i, j, 18)=RS("TIMEUP")
			tmpRec(i, j, 19)=RS("DATEDOWN")
			tmpRec(i, j, 20)=RS("TIMEDOWN")
			tmpRec(i, j, 21)=RS("JIAMEMO")
			tmpRec(i, j, 22)=RS("JIATYPE") 
			tmpRec(i, j, 23)=RS("hhour") 
			tmpRec(i, j, 24)=RS("jiaID")  
			tmpRec(i, j, 25)=RS("jia_str")  
			tmpRec(i, j, 26)=RS("DATEUP") &" "&mid("日一二三四五六",weekday(cdate(rs("DATEUP"))) , 1 )  
			tmpRec(i, j, 27)=RS("DATEDOWN") &" "&mid("日一二三四五六",weekday(cdate(rs("DATEDOWN"))) , 1 )  
			tmpRec(i, j, 28)=RS("place")
			tmpRec(i, j, 29)=RS("xid")
			tmpRec(i, j, 30)=RS("autoid")
			 
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
	Session("yeee04B2") = tmpRec	
else    
	TotalPage = cint(request("TotalPage"))
	'StoreToSession()
	CurrentPage = cint(request("CurrentPage"))
	RecordInDB  = REQUEST("RecordInDB")
	'tmpRec = Session("yeee04B2")
	
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


FUNCTION FDT(D)
	Response.Write YEAR(D)&"/"&RIGHT("00"&MONTH(D),2)&"/"&RIGHT("00"&DAY(D),2) 
	
END FUNCTION 

nowmonth = year(date())&right("00"&month(date()),2)  
if month(date())="01" then  
	calcmonth = year(date()-1)&"12" 
else
	calcmonth =  year(date())&right("00"&month(date())-1,2)   
end if 	 

zjdays = request("zjdays")
if zjdays="" then zjdays = 0 
%>

<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="refresh">
<link rel="stylesheet" href="../../Include/style.css" type="text/css">
<link rel="stylesheet" href="../../Include/style2.css" type="text/css"> 
 
</head>   
<body  topmargin="5" leftmargin="5"  marginwidth="0" marginheight="0"  onkeydown="enterto()" onload="f()">
<form name="<%=self%>" method="post" action="yeee04.showdata.asp">
<INPUT TYPE=HIDDEN NAME="UID" VALUE=<%=SESSION("NETUSER")%>>
<INPUT TYPE=hidden NAME=TotalPage VALUE="<%=TotalPage%>">
<INPUT TYPE=hidden NAME=CurrentPage VALUE="<%=CurrentPage%>">
<INPUT TYPE=hidden NAME=PageRec VALUE="<%=PageRec%>">
<INPUT TYPE=hidden NAME=RecordInDB VALUE="<%=RecordInDB%>">
<INPUT TYPE=hidden NAME=gTotalPage VALUE="<%=gTotalPage%>"> 	 

<INPUT NAME=ym1 VALUE="<%=ym1%>" TYPE="HIDDEN" >
<INPUT NAME=ym2 VALUE="<%=ym2%>" TYPE="HIDDEN"  >
<INPUT NAME=jb VALUE="<%=jb%>" TYPE="HIDDEN"  >

<table width="460" border="0" cellspacing="0" cellpadding="0">
	<tr><TD>
	<img border="0" src="../image/icon.gif" align="absmiddle">
	員工請假資料查詢</TD></tr>
</table> 
<hr size=0	style='border: 1px dotted #999999;' align=left width=500>		
<TABLE WIDTH=500 CLASS=txt8 BORDER=0  cellspacing="2" cellpadding="2" >   
	<TR height=25 >
		<TD nowrap align=right width=60>查詢日期<br>Ngay</TD>
		<TD nowrap > 
			<INPUT NAME=DAT1 VALUE="<%=DAT1%>"  class=inputbox size=11  onblur="date_change(1)" >
			<INPUT NAME=DAT2 VALUE="<%=DAT2%>"  class=inputbox size=11  onblur="date_change(2)" >
		</TD>  
		<TD nowrap align=right width=60>員工編號<br>So the</TD>
		<TD >
			<input name=empid1 class=inputbox size=8 maxlength=5   VALUE="<%=QUERYX%>" readonly  > 
			<br><%=empname%>
		</TD> 	 
	</TR>	
	 
</TABLE>

<hr size=0	style='border: 1px dotted #999999;' align=left width=500  >	 	
<!-------------------------------------------------------------------->  
<table  class="txt8"  cellspacing="1" cellpadding="1"  >
	<tr BGCOLOR="LightGrey" height=22>		 
		<TD width=20 nowrap align=center >STT</TD>
 		<TD width=30 nowrap align=center >刪<br>xoa di</TD>
 		<td width=30 nowrap  align=center >地點</td>		
		<TD align=center width=100 nowrap >假別<br>loai phep</TD>
 		<TD width=120 align=center nowrap >日期(起)<br>Ngay(tu)</TD>		
		<TD width=120 align=center nowrap >日期(迄)<br>Ngay(Den)</TD>		
		<td width=50 nowrap align=center  >時數<br>So gio</td>
		<td width=100 nowrap  align=center >事由<br>Ly do</td>		
		<td width=60 nowrap  align=center >請假編號<br>so nhap phep</td>		
		
	</tr>
	 
	<%for CurrentRow = 1 to PageRec
		IF CurrentRow MOD 2 = 0 THEN 
			WKCOLOR="LavenderBlush"
		ELSE
			WKCOLOR=""
		END IF 	 
		'if tmpRec(CurrentPage, CurrentRow, 1) <> "" then 
	%>
	<TR BGCOLOR="<%=WKCOLOR%>" height="20"> 			
		<Td><%=currentrow%></td>
 		<TD align="center">
			<%IF tmpRec(CurrentPage, CurrentRow, 1)<>"" THEN %>		
				<%IF tmpRec(CurrentPage, CurrentRow, 0)="del" THEN  %>
					<INPUT type=checkbox name=func value=del onclick="del(<%=CurrentRow - 1%>)" checked >
				<%ELSE%>	
					<INPUT type=checkbox name=func value=del onclick="del(<%=CurrentRow - 1%>)"   >
				<%END IF%>	
				<INPUT TYPE=HIDDEN NAME=OP >
			<%ELSE%>	
				<INPUT TYPE=HIDDEN NAME=FUNC  >
				<INPUT TYPE=HIDDEN NAME=OP   >
			<%END IF %> 			
			<INPUT TYPE=HIDDEN NAME=EMPID VALUE="<%=tmpRec(CurrentPage, CurrentRow, 1)%>">
			<INPUT TYPE=HIDDEN NAME="autoid" VALUE="<%=tmpRec(CurrentPage, CurrentRow, 30)%>">
			
 		</TD>
		<Td> 	
			<%IF tmpRec(CurrentPage, CurrentRow, 1)<>"" THEN %>	
				<INPUT  NAME="place" size=2 class="inputbox8" value="<%=tmpRec(CurrentPage, CurrentRow, 28)%>" style="text-align:center" maxlength=1 title="W:境外,I:境內" >
			<%else%>
				<INPUT  type="hidden" NAME="place" size=2 class="inputbox8" value="" style="text-align:center"  >
			<%end if%>
		</td><!--境內或境外-->		
 		<TD><%IF tmpRec(CurrentPage, CurrentRow, 1)<>"" THEN %>	
				<INPUT  NAME="jiatype" value="<%=tmpRec(CurrentPage, CurrentRow, 22)%>" class="inputbox8" size=1 style="text-align:center" maxlength=1><%=tmpRec(CurrentPage, CurrentRow, 25)%>
				<%else%>
				<INPUT  NAME="jiatype" value="" class="inputbox8" size=1 style="text-align:center" type="hidden" >
				<%end if%>
	 			<INPUT TYPE=HIDDEN NAME=HOLIDAY_TYPE value="<%=tmpRec(CurrentPage, CurrentRow, 22)%>" >
	 			<INPUT TYPE=HIDDEN NAME=HOLIDASTR value="<%=tmpRec(CurrentPage, CurrentRow, 22)%>&nbsp;<%=tmpRec(CurrentPage, CurrentRow, 25)%>" class=readonly  readonly size=12  > 	 			 
				
 		</TD>
 		<TD align=center><%=tmpRec(CurrentPage, CurrentRow, 26)%>&nbsp;<%=tmpRec(CurrentPage, CurrentRow, 18)%> 			
 				<input TYPE=HIDDEN name=HHDAT1 size=14 class=readonly readonly   value="<%=tmpRec(CurrentPage, CurrentRow, 26)%>" >
				<input TYPE=HIDDEN name=HHTIM1 size=5 class=readonly readonly   value="<%=tmpRec(CurrentPage, CurrentRow, 18)%>" style="text-align:center" > 			
				<input TYPE=HIDDEN name=HHDAT2 size=14 class=readonly readonly   value="<%=tmpRec(CurrentPage, CurrentRow, 27)%>" >
				<input TYPE=HIDDEN name=HHTIM2 size=5 class=readonly readonly   value="<%=tmpRec(CurrentPage, CurrentRow, 20)%>" style="text-align:center">
 		</TD> 		
 		<TD align=center><%=tmpRec(CurrentPage, CurrentRow, 27)%>&nbsp;<%=tmpRec(CurrentPage, CurrentRow, 20)%> 			
 		</TD>  		
 		<TD align=center><%=tmpRec(CurrentPage, CurrentRow, 23)%>
 			<%IF tmpRec(CurrentPage, CurrentRow, 1)<>"" THEN %>	
 				<input TYPE=HIDDEN name=toth size=4 class=readonly readonly   value="<%=tmpRec(CurrentPage, CurrentRow, 23)%>" style="text-align:right">
 			<%ELSE%>	
				<INPUT TYPE=HIDDEN NAME=toth  >				
			<%END IF %>
 		</TD> 
 		<TD align="left"><%=tmpRec(CurrentPage, CurrentRow, 21)%>
 			<%IF tmpRec(CurrentPage, CurrentRow, 1)<>"" THEN %>	
 				<input TYPE=HIDDEN name=JIAMEMO size=15 class=readonly readonly   value="<%=tmpRec(CurrentPage, CurrentRow, 21)%>" >
 			<%ELSE%>	
				<INPUT TYPE=HIDDEN NAME=JIAMEMO  >				
			<%END IF %>
 		</TD> 		
		<Td>
			<%IF tmpRec(CurrentPage, CurrentRow, 1)<>"" THEN %>					
 				<input  name="xid" size=8 class="inputbox8"   value="<%=tmpRec(CurrentPage, CurrentRow, 29)%>" >
 			<%ELSE%>	
				<INPUT  NAME="xid" size=8  class="inputbox8" >				
			<%END IF %>
		</td><!--xid-->

		
	</TR>
	<%next%>   		
</table>	
<hr size=0 width=750 align="left">	 
<table  class="txt8"  cellspacing="1" cellpadding="1"  >
	<tr height=22>		 
		<td colspan=6>新增請假資料(以天為單位) </td>
	</tr>	
	<tr BGCOLOR="LightGrey" height=22>		 
		<TD width=20 nowrap align=center >STT</TD>
 		<td width=30 nowrap  align=center >地點</td>
		<TD width=30 align=center  nowrap >假別</TD>
		<TD width=60 nowrap align=center >日期(起)</TD> 		
		<TD width=60 nowrap align=center >日期(迄)</TD> 		 		
		<td width=60 nowrap  align=center >請假編號<br>so nhap phep</td>				
		<TD width=120 align=center nowrap >事由</TD>		
	</tr> 
	<%
		for x = 1 to 3
	%>
	<tr>
		<td><%=X%></td>
		<td>
				<INPUT NAME="n_place" VALUE="" size=2 style="text-align:center" class="inputbox8"  >					
		</td>
		<td>
				<INPUT NAME="jb2" VALUE="" size=2 style="text-align:center" class="inputbox8" >					
		</td>
		<td>
			<input name="n_dat1" size=11 class="inputbox8"  value=""  onblur="datchg(<%=x-1%>,1)">
		</td>
		<td>
			<input name="n_dat2" size=11 class="inputbox8"  value=""  onblur="datchg(<%=x-1%>,2)">
		</td>		
		<td>
			<input name="n_xid" size=6 class="inputbox8"  value="" maxlength=6 >
		</td>
		<td>
			<input name="jbmemo" size=20 class="inputbox8"  value="" >
		</td>		
	</tr>
	<%next%>	
</table>	

<INPUT TYPE=HIDDEN NAME=EMPID VALUE="">
<INPUT TYPE=HIDDEN NAME="autoid" VALUE="">	
<INPUT  type="hidden" NAME="place" size=2 class="inputbox8" value=""  > 
<INPUT  type="hidden" NAME="jiatype" size=2 class="inputbox8" value=""  > 
 <INPUT TYPE=HIDDEN NAME="xid" VALUE=""> 
 <INPUT TYPE=HIDDEN NAME="op" VALUE=""> 
 <INPUT TYPE=HIDDEN NAME="func" VALUE=""> 
 <INPUT TYPE=HIDDEN NAME="HHDAT1" VALUE=""> 
 <INPUT TYPE=HIDDEN NAME="jiamemo" VALUE=""> 
 
<TABLE border=0 width=600 class=font9 >
<tr>
    <td align="CENTER" height=40 width=70%>
    
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
	PAGE:<%=CURRENTPAGE%> / <%=TOTALPAGE%> , COUNT:<%=RECORDINDB%>
	</td>
	<td align=right>		
		<%if session("rights")<="1" then %>
			<input type="BUTTON" name="send" value="(Y)Confirm" class=button ONCLICK="go()">	
		<%end if%>	
		<input type="BUTTON" name="send" value="(X)Close" class=button ONCLICK="window.close()">	
	</td>		
</TR>
</TABLE>  
<input type=hidden name=func >
<input type=hidden name=op >
<input type=hidden name=empid >

</form>

</body>
</html> 
<!-- #include file="../Include/func.inc" -->
<script language=vbscript> 
function del(index) 
	if <%=self%>.func(index).checked=true then 
		<%=self%>.op(index).value="D" 		
		'open "<%=self%>.back.asp?func=del&index="& index &"&CurrentPage="& <%=CurrentPage%> , "Back"
	else
		<%=self%>.op(index).value=""  
		'open "<%=self%>.back.asp?func=no&index="& index &"&CurrentPage="& <%=CurrentPage%> , "Back"
	end if 	 	
	'parent.best.cols="70%,30%"
end function 

function BACKMAIN()	
	open "../main.asp" , "_self"
end function   

function oktest(N)	
	tp=<%=self%>.totalpage.value 
	cp=<%=self%>.CurrentPage.value 
	rc=<%=self%>.RecordInDB.value 
	'open "empworkB.fore.asp?empautoid="& N &"&yymm="&"<%=calcmonth%>", "_self" 
	open "empworkB.fore.asp?empautoid="& N &"&YYMM="&"<%=calcmonth%>" &"&Ftotalpage=" & tp &"&Fcurrentpage=" & cp &"&FRecordInDB=" & rc , "_self" 
end function   

FUNCTION CLR()
	OPEN "<%=SELF%>.ASP" , "_self"
END FUNCTION 

function strchg(a)
	if a=1 then 
		<%=self%>.empid1.value = Ucase(<%=self%>.empid1.value) 
		'IF TRIM(<%=self%>.empid.value)<>"" THEN 
			<%=SELF%>.totalpage.VALUE=0
			<%=SELF%>.ACTION="<%=SELF%>.FORE.ASP"
			<%=SELF%>.SUBMIT()
		'END IF 
	elseif a=2 then 	
		<%=self%>.empid2.value = Ucase(<%=self%>.empid2.value)
	end if 	
end function   

function go()
	<%=self%>.action="<%=self%>.upd.asp" 
	<%=self%>.submit()
end function  


'*******檢查日期*********************************************
function date_change(a)	

if a=1 then
	INcardat = Trim(<%=self%>.dat1.value)  		
elseif a=2 then
	INcardat = Trim(<%=self%>.dat2.value)
end if		
		    
IF INcardat<>"" THEN
	ANS=validDate(INcardat)
	IF ANS <> "" THEN
		if a=1 then
			Document.<%=self%>.dat1.value=ANS			
		elseif a=2 then
			Document.<%=self%>.dat2.value=ANS
			datachg()	
		end if		
	ELSE
		ALERT "EZ0067:輸入日期不合法 !!" 
		if a=1 then
			Document.<%=self%>.dat1.value=""
			Document.<%=self%>.dat1.focus()
		elseif a=2 then
			Document.<%=self%>.dat2.value=""
			Document.<%=self%>.dat2.focus()
		end if		
		EXIT function 
	END IF
		 
ELSE
	'alert "EZ0015:日期該欄位必須輸入資料 !!" 		
	EXIT function
END IF     
END function   

function datchg(index,a)
	'alert index 
	if a=1 then 
		INcardat = Trim(<%=self%>.n_dat1(index).value)  		
	elseif a=2 then 	
		INcardat = Trim(<%=self%>.n_dat2(index).value)  		
	end if 	
	if INcardat<>"" then 
		ANS=validDate(INcardat)
		IF ANS <> "" THEN		
			if a=1 then 
				<%=self%>.n_dat1(index).value=ANS					
			elseif a=2 then 
				<%=self%>.n_dat2(index).value=ANS					
			end if 
		ELSE
			ALERT "EZ0067:輸入日期不合法 !!" 		
			if a=1 then 
				<%=self%>.n_dat1(index).value=""
				<%=self%>.n_dat1(index).focus()		
			elseif a=2 then 
				<%=self%>.n_dat2(index).value=""
				<%=self%>.n_dat2(index).focus()		
			end if 
			exit function
		END IF
	end if 	
end function 

function enterto()
	if window.event.keyCode = 13 then window.event.keyCode =9 		
end function  

function f()
	'<%=self%>.QUERYX.focus()	
end function   

function datachg()
	<%=SELF%>.totalpage.VALUE=0
	<%=self%>.action="yeee04.showdata.asp"
	<%=self%>.submit
end function 
	
</script>


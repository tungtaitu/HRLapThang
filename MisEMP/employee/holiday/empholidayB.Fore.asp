<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<!-- #include file = "../../GetSQLServerConnection.fun" --> 
<!-- #include file="../../ADOINC.inc" -->
<%
'on error resume next   
session.codepage="65001"
SELF = "EMPHOLIDAYB"
 
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


gTotalPage = 1
PageRec = 16    'number of records per page
TableRec = 30    'number of fields per record  

if dat1="" and dat2="" and whsno="" and groupid="" and country="" and QUERYX="" then 
	sql="select * from empfile where empid='XX' "
else
	SQL="SELECT  A.JIATYPE,  CONVERT(CHAR(10), A.DATEUP, 111) DATEUP , A.TIMEUP, convert(char(10) , A.DATEDOWN , 111) datedown, "
	SQL=SQL&"A.TIMEDOWN , A.HHOUR, A.MEMO AS JIAMEMO  , a.autoid as jiaid,  B.*  , isnull(c.sys_value,'') as jia_str  FROM   "
	SQL=SQL&"( SELECT * FROM EMPHOLIDAY   ) A  "
	SQL=SQL&"LEFT JOIN ( SELECT * FROM view_empfile ) B ON B.EMPID = A.EMPID  	 "
	SQL=SQL&"LEFT JOIN ( SELECT * FROM basicCode where func='JB'  ) c  on c.sys_type = a.JIATYPE  "
	SQL=SQL&"WHERE 1=1  " 	
	SQL=SQL&"and country like  '"& country &"%'  "
	SQL=SQL&"AND whsno like '"& whsno &"%' and unitno like '"& unitno &"%'  and groupid like '"& groupid &"%'  " 
	SQL=SQL&"and zuno like '"& zuno &"%' and job like '"& job &"%' and b.empid like '%"& QUERYX &"%'  "
	IF DAT1<>"" and DAT2<>"" then  
	 	sql=sql& "and CONVERT(CHAR(10), A.DATEUP, 111) BETWEEN '"& DAT1 &"' AND '"& DAT2 &"' " 
	END IF 
	SQL=SQL&"order by b.empid, A.DATEUP , a.jiaType "  
end if 	

if request("TotalPage") = "" or request("TotalPage") = "0" then 
	CurrentPage = 1	
	rs.Open SQL, conn, 3, 3 
	IF NOT RS.EOF THEN 
		rs.PageSize = PageRec 
		RecordInDB = rs.RecordCount 
		TotalPage = rs.PageCount  
		gTotalPage = TotalPage
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
			tmpRec(i, j, 8) = rs("unitno")	 
			tmpRec(i, j, 9)	=RS("groupid") 
			tmpRec(i, j, 10)=RS("zuno") 				
			tmpRec(i, j, 11)=RS("wstr") 	
			tmpRec(i, j, 12)=RS("ustr") 	
			tmpRec(i, j, 13)=RS("gstr") 	
			tmpRec(i, j, 14)=RS("zstr") 	
			tmpRec(i, j, 15)=RS("jstr") 	
			tmpRec(i, j, 16)=RS("cstr")
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
	Session("EMPHOLIDAYB") = tmpRec	
else    
	TotalPage = cint(request("TotalPage"))
	'StoreToSession()
	CurrentPage = cint(request("CurrentPage"))
	RecordInDB  = REQUEST("RecordInDB")
	tmpRec = Session("EMPHOLIDAYB")
	
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

%>

<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="refresh">
<link rel="stylesheet" href="../../Include/style.css" type="text/css">
<link rel="stylesheet" href="../../Include/style2.css" type="text/css"> 
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--
'-----------------enter to next field
function enterto()
	if window.event.keyCode = 13 then window.event.keyCode =9 		
end function  

function f()
	'<%=self%>.QUERYX.focus()	
end function   

function datachg()
	<%=SELF%>.totalpage.VALUE=0
	<%=self%>.action="<%=SELF%>.fore.asp?totalpage=0"
	<%=self%>.submit
end function 

-->
</SCRIPT>  
</head>   
<body  topmargin="5" leftmargin="5"  marginwidth="0" marginheight="0"  onkeydown="enterto()" onload="f()">
<form name="<%=self%>" method="post" action="<%=SELF%>.fore.asp">
<INPUT TYPE=HIDDEN NAME="UID" VALUE=<%=SESSION("NETUSER")%>>
<INPUT TYPE=hidden NAME=TotalPage VALUE="<%=TotalPage%>">
<INPUT TYPE=hidden NAME=CurrentPage VALUE="<%=CurrentPage%>">
<INPUT TYPE=hidden NAME=PageRec VALUE="<%=PageRec%>">
<INPUT TYPE=hidden NAME=RecordInDB VALUE="<%=RecordInDB%>">
<INPUT TYPE=hidden NAME=gTotalPage VALUE="<%=gTotalPage%>"> 	 
<INPUT NAME=DAT1 VALUE="<%=DAT1%>" TYPE=HIDDEN >
<INPUT NAME=DAT2 VALUE="<%=DAT2%>" TYPE=HIDDEN  >
<table width="460" border="0" cellspacing="0" cellpadding="0">
	<tr><TD>
	<img border="0" src="../../image/icon.gif" align="absmiddle">
	員工請假作業維護</TD></tr>
</table> 
<hr size=0	style='border: 1px dotted #999999;' align=left width=500>		
<TABLE WIDTH=400 CLASS=FONT9 BORDER=0>   
	<TR height=25 >
		<TD nowrap align=right>廠別</TD>
		<TD > 
			<select name=WHSNO  class=font9 onchange="datachg()">
				<option value=""></option>
				<%SQL="SELECT * FROM BASICCODE WHERE FUNC='WHSNO' ORDER BY SYS_TYPE "
				SET RST = CONN.EXECUTE(SQL)
				WHILE NOT RST.EOF  
				%>
				<option value="<%=RST("SYS_TYPE")%>" <%IF  RST("SYS_TYPE")=whsno THEN %> SELECTED <%END IF%> ><%=RST("SYS_VALUE")%></option>				 
				<%
				RST.MOVENEXT
				WEND 
				%>
			</SELECT>
			<%SET RST=NOTHING %>
		</TD> 
		<TD   nowrap align=right>國籍</TD>
		<TD > 
			<select name=country  class=font9 onchange="datachg()">
				<option value=""></option>
				<%SQL="SELECT * FROM BASICCODE WHERE FUNC='country' and sys_type<>'AAA' ORDER BY SYS_TYPE "
				SET RST = CONN.EXECUTE(SQL)
				WHILE NOT RST.EOF  
				%>
				<option value="<%=RST("SYS_TYPE")%>" <%IF  RST("SYS_TYPE")=country THEN %> SELECTED <%END IF%>><%=RST("SYS_VALUE")%></option>				 
				<%
				RST.MOVENEXT
				WEND 
				%>
			</SELECT>
			<%SET RST=NOTHING %>
		</TD>	
		<TD nowrap align=right >組/部門</TD>
		<TD >
			<select name=GROUPID  class=font9  onchange="datachg()">
				<option value=""></option>
				<%SQL="SELECT * FROM BASICCODE WHERE FUNC='GROUPID' and sys_type <>'AAA' ORDER BY SYS_TYPE "
				SET RST = CONN.EXECUTE(SQL)
				WHILE NOT RST.EOF  
				%>
				<option value="<%=RST("SYS_TYPE")%>" <%IF  RST("SYS_TYPE")=GROUPID THEN %> SELECTED <%END IF%> ><%=RST("SYS_VALUE")%></option>				 
				<%
				RST.MOVENEXT
				WEND 
				%>
			</SELECT>
			<%SET RST=NOTHING %>
		</TD> 
		<TD nowrap align=right >員工編號</TD>
		<TD >
			<input name=empid1 class=inputbox size=8 maxlength=5 ONBLUR=strchg(1) VALUE="<%=QUERYX%>"> 
		</TD> 	 		
	</TR>	
	 
</TABLE>

<hr size=0	style='border: 1px dotted #999999;' align=left width=500  >	 	
<!-------------------------------------------------------------------->  
<table width=580 class=font9 cellpadding="0" >
	<tr BGCOLOR="LightGrey" height=22>
		<TD width=30 nowrap align=center >刪除</TD> 
		<TD width=50 nowrap align=center >工號</TD> 		
 		<TD width=190 nowrap align=center >姓名</TD>
 		<TD align=center  >假別</TD>
 		<TD width=80 align=center nowrap >日期(起)</TD>
		<TD align=center  >時間(起)</TD>
		<TD width=80 align=center nowrap >日期(迄)</TD>
		<td align=center  >時間(迄)</td>
		<td align=center  >時數</td>
		<td align=center >事由</td>		
	</tr>
	 
	<%for CurrentRow = 1 to PageRec
		IF CurrentRow MOD 2 = 0 THEN 
			WKCOLOR="LavenderBlush"
		ELSE
			WKCOLOR=""
		END IF 	 
		'if tmpRec(CurrentPage, CurrentRow, 1) <> "" then 
	%>
	<TR BGCOLOR=<%=WKCOLOR%> > 	
		<TD align=center>
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
		</TD>
		<TD align=center>
			<%=tmpRec(CurrentPage, CurrentRow, 1)%>
			<INPUT TYPE=HIDDEN NAME=EMPID VALUE="<%=tmpRec(CurrentPage, CurrentRow, 1)%>">
		</TD> 		
 		<TD>
 			<%=tmpRec(CurrentPage, CurrentRow, 2)%>&nbsp;
 			<font class=txt8><%=tmpRec(CurrentPage, CurrentRow, 3)%></font>
 		</TD>
 		<TD>
 			<%IF tmpRec(CurrentPage, CurrentRow, 1)<>"" THEN 
 			%>
	 			<INPUT TYPE=HIDDEN NAME=HOLIDAY_TYPE value="<%=tmpRec(CurrentPage, CurrentRow, 22)%>" >
	 			<INPUT NAME=HOLIDASTR value="<%=tmpRec(CurrentPage, CurrentRow, 22)%>&nbsp;<%=tmpRec(CurrentPage, CurrentRow, 25)%>" class=readonly  readonly size=12  > 	 			 
			<%ELSE%>	
				<INPUT TYPE=HIDDEN NAME=HOLIDAY_TYPE  >	
				<INPUT TYPE=HIDDEN NAME=HOLIDASTR >			
			<%END IF %>
 		</TD>
 		<TD align=center>
 			<%IF tmpRec(CurrentPage, CurrentRow, 1)<>"" THEN %>	
 				<input name=HHDAT1 size=14 class=readonly readonly   value="<%=tmpRec(CurrentPage, CurrentRow, 26)%>" > 				
 			<%ELSE%>	
				<INPUT TYPE=HIDDEN NAME=HHDAT1  >								
			<%END IF %>
 		</TD>
 		<TD align=center>
 			<%IF tmpRec(CurrentPage, CurrentRow, 1)<>"" THEN %>	
 				<input name=HHTIM1 size=5 class=readonly readonly   value="<%=tmpRec(CurrentPage, CurrentRow, 18)%>" style="text-align:center" >
 			<%ELSE%>	
				<INPUT TYPE=HIDDEN NAME=HHTIM1  >				
			<%END IF %>
 		</TD>
 		<TD align=center>
 			<%IF tmpRec(CurrentPage, CurrentRow, 1)<>"" THEN %>	 				
 				<input name=HHDAT2 size=14 class=readonly readonly   value="<%=tmpRec(CurrentPage, CurrentRow, 27)%>" >
 			<%ELSE%>					
				<INPUT TYPE=HIDDEN NAME=HHDAT2  >			
			<%END IF %>
 		</TD> 
 		<TD align=center>
 			<%IF tmpRec(CurrentPage, CurrentRow, 1)<>"" THEN %>	
 				<input name=HHTIM2 size=5 class=readonly readonly   value="<%=tmpRec(CurrentPage, CurrentRow, 20)%>" style="text-align:center">
 			<%ELSE%>	
				<INPUT TYPE=HIDDEN NAME=HHTIM2  >				
			<%END IF %>
 		</TD>
 		<TD align=center>
 			<%IF tmpRec(CurrentPage, CurrentRow, 1)<>"" THEN %>	
 				<input name=toth size=4 class=readonly readonly   value="<%=tmpRec(CurrentPage, CurrentRow, 23)%>" style="text-align:right">
 			<%ELSE%>	
				<INPUT TYPE=HIDDEN NAME=toth  >				
			<%END IF %>
 		</TD> 
 		<TD align=center>
 			<%IF tmpRec(CurrentPage, CurrentRow, 1)<>"" THEN %>	
 				<input name=JIAMEMO size=15 class=readonly readonly   value="<%=tmpRec(CurrentPage, CurrentRow, 21)%>" >
 			<%ELSE%>	
				<INPUT TYPE=HIDDEN NAME=JIAMEMO  >				
			<%END IF %>
 		</TD> 		
	</TR>
	<%next%> 	 
	
</table>	
 
 
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
		<input type="BUTTON" name="send" value="確　定" class=button onclick="GO()" >
		<input type="BUTTON" name="send" value="取　消" class=button ONCLICK="CLR()">	
	</td>		
</TR>
</TABLE>  
<input type=hidden name=func >
<input type=hidden name=op >
<input type=hidden name=empid >

</form>

</body>
</html>
<%
Sub StoreToSession()
	Dim CurrentRow
	tmpRec = Session("EMPHOLIDAYB")
	for CurrentRow = 1 to PageRec
		tmpRec(CurrentPage, CurrentRow, 0) = request("op")(CurrentRow)	
	next 
	Session("EMPHOLIDAYB") = tmpRec
	
End Sub
%>  
<script language=vbscript> 
function del(index) 
	if <%=self%>.func(index).checked=true then 
		<%=self%>.op(index).value="del" 		
		open "<%=self%>.back.asp?func=del&index="& index &"&CurrentPage="& <%=CurrentPage%> , "Back"
	else
		<%=self%>.op(index).value="no"  
		open "<%=self%>.back.asp?func=no&index="& index &"&CurrentPage="& <%=CurrentPage%> , "Back"
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
			<%=SELF%>.ACTION="<%=SELF%>.FORE.ASP?TOTALPAE=0"
			<%=SELF%>.SUBMIT()
		'END IF 
	elseif a=2 then 	
		<%=self%>.empid2.value = Ucase(<%=self%>.empid2.value)
	end if 	
end function   

function go()
	<%=self%>.action="<%=self%>.updateDB.asp" 
	<%=self%>.submit()
end function 
	
</script>


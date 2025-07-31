<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<!-- #include file = "../../GetSQLServerConnection.fun" -->
<!-- #include file="../../ADOINC.inc" -->
<%
'on error resume next
session.codepage="65001"
SELF = "empBHGT"

Set conn = GetSQLServerConnection()
Set rs = Server.CreateObject("ADODB.Recordset")
YYMM=REQUEST("YYMM")
whsno = trim(request("whsno"))
unitno = trim(request("unitno"))
groupid = trim(request("groupid"))
COUNTRY = trim(request("COUNTRY"))
job = trim(request("job"))
QUERYX = trim(request("empid1"))
outemp = request("outemp")
lastym = left(yymm,4) &  right("00" & cstr(right(yymm,2)-1) ,2 )
if right(yymm,2)="01"  then
	lastym = left(yymm,4)-1 &"12"
end if
shift = request("shift")

PERAGE = REQUEST("PERAGE")

calcdt = left(YYMM,4)&"/"& right(yymm,2)&"/01"
'下個月
if right(yymm,2)="12" then
	ccdt = cstr(left(YYMM,4)+1)&"/01/01"
else
	ccdt = left(YYMM,4)&"/"& right("00" & right(yymm,2)+1,2)  &"/01"
end if
'response.write ccdt

 '一個月有幾天
cDatestr=CDate(LEFT(YYMM,4)&"/"&RIGHT(YYMM,2)&"/01")
days = DAY(cDatestr+(32-DAY(cDatestr))-DAY(cDatestr+(32-DAY(cDatestr))))   '一個月有幾天
'本月最後一天
ENDdat = CDate(LEFT(YYMM,4)&"/"&RIGHT(YYMM,2)&"/"&DAYS)
'RESPONSE.WRITE days

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

'RECALC = REQUEST("recalc")
'IF recalc="Y"  THEN 
'	sql="delete empbhgt where  yymm='"& yymm &"' "
'	conn.execute(Sql) 
'end if 	


gTotalPage = 1
PageRec = 15    'number of records per page
TableRec = 35    'number of fields per record

sql="select isnull(c.bb,0) BB, BHXH =  case when isnull(b.empid,'')='' then  CASE WHEN CONVERT(CHAR(10), A.BHDAT, 111)<'"& ccdt &"' AND ISNULL(A.BHDAT,'')<>'' THEN isnull(c.bb,0)*0.05 ELSE 0 END  else  isnull(b.bhxh5,isnull(c.bb,0)*0.05)  end  ,  "&_
	"BHYT =  case when isnull(b.empid,'')='' then  CASE WHEN CONVERT(CHAR(10), A.BHDAT, 111)<'"& ccdt &"' AND ISNULL(A.BHDAT,'')<>''THEN  isnull(c.bb,0)*0.01 ELSE 0 END  else  isnull(b.BHYT1,isnull(c.bb,0)*0.01)  end  ,   "&_
	"BHTN =  case when isnull(b.empid,'')='' then  CASE WHEN CONVERT(CHAR(10), A.BHDAT, 111)<'"& ccdt &"' AND ISNULL(A.BHDAT,'')<>''THEN  isnull(c.bb,0)*0.01 ELSE 0 END  else   isnull(b.BHTn1,isnull(c.bb,0)*0.01)   end  ,   "&_
	"GTAMT = case when isnull(b.empid,'')='' then case when ( isnull(a.gtdat,'')<>''AND isnull(a.gtdat,'')<='"& yymm &"' ) then 5000 else 0 end else isnull(b.gtamt,0) end , "&_
	"flag=case when isnull(b.empid,'')='' then 'Y' else 'N' end , isnull(b.kh1,0) kh1, isnull(b.chanjia,0) chanjia, isnull(b.memo,'') bhmemo,a.* from  "&_
	"( select * from  view_empfile  where  country='VN'   ) a  "&_
	"left join ( select * from empbhgt  where yymm='"& YYMM &"'  ) b on b.empid = a.empid "&_   
	"left join ( select * from bemps where yymm='"& yymm &"'  ) c on c.empid = a.empid "&_ 
	"where CONVERT(CHAR(10), indat, 111)< '"& ccdt &"' and ( isnull(a.outdat,'')='' or a.outdat>'"& calcdt &"' )  "&_
	"and a.whsno like '%"& whsno &"%' and a.unitno like '%"& unitno &"%' and a.groupid like '%"& groupid &"%'  "&_
	"and a.COUNTRY like '%"& COUNTRY  &"%' and A.job like '%"& job &"%' and a.empid like '%"& QUERYX &"%' "
	if outemp="D" then
		sql=sql&" and ( isnull(a.outdat,'')<>'')  "
	elseif 	outemp="N" then
		sql=sql&" and ( isnull(a.outdat,'')='' )  "
	end if
	if shift="C" then
		sql=sql&" and isnull(a.shift,'') NOT IN ('N', 'A', 'B' ) "
	ELSE
		sql=sql&" and isnull(a.shift,'') like '%"& shift &"'    "
	end if
	sql=sql&"order by a.empid   "
 
'response.write sql
'response.end
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
			tmpRec(i, j, 17)=RS("autoid")
			tmpRec(i, j, 18)=RS("outdate")
			tmpRec(i, j, 19)=RS("code")			
			tmpRec(i, j, 20)=RS("bb")
			tmpRec(i, j, 21)=RS("BHXH")
			tmpRec(i, j, 22)=RS("BHYT")
			tmpRec(i, j, 23)=RS("GTAMT")			
			tmpRec(i, j, 24)=RS("flag")	
			if rs("flag")="N" then 
				tmpRec(i, j, 25)="Blue"
			else
				tmpRec(i, j, 25)="Black"
			end if 	
			tmpRec(i, j, 26)=RS("BHDAT")
			tmpRec(i, j, 27)=RS("GTDAT")
			'tmpRec(i, j, 28)=CDBL(tmpRec(i, j, 21))+CDBL(tmpRec(i, j, 22))
			tmpRec(i, j, 29)=rs("KH1")
			tmpRec(i, j, 30)=rs("chanjia")
			tmpRec(i, j, 31)=rs("bhmemo")
			tmpRec(i, j, 32)=rs("BHTN")  '失業保險 since 200901 
			tmpRec(i, j, 28)=CDBL(tmpRec(i, j, 21))+CDBL(tmpRec(i, j, 22))+CDBL(tmpRec(i, j, 32))
			'response.write   tmpRec(i, j, 20) &"<Br>"
			'response.write  
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
	Session("empBHGTD") = tmpRec
else
	TotalPage = cint(request("TotalPage"))
	'StoreToSession()
	CurrentPage = cint(request("CurrentPage"))
	RecordInDB  = REQUEST("RecordInDB")
	tmpRec = Session("empBHGTD")

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

'response.end

FUNCTION FDT(D)
	Response.Write YEAR(D)&"/"&RIGHT("00"&MONTH(D),2)&"/"&RIGHT("00"&DAY(D),2)

END FUNCTION
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
	'<%=self%>.PHU(0).focus()
	'<%=self%>.PHU(0).SELECT()
end function

function chgdata()
	<%=self%>.action="empfile.salary.asp?totalpage=0"
	<%=self%>.submit
end function
-->
</SCRIPT>
</head>
<body   topmargin="0" leftmargin="0"  marginwidth="0" marginheight="0"  onkeydown="enterto()" onload="f()" bgproperties="fixed"  >
<form name="<%=self%>" method="post" action="empBHGT.ForeGnd.asp">
<INPUT TYPE=HIDDEN NAME="UID" VALUE=<%=SESSION("NETUSER")%>>
<INPUT TYPE=hidden NAME=TotalPage VALUE="<%=TotalPage%>">
<INPUT TYPE=hidden NAME=CurrentPage VALUE="<%=CurrentPage%>">
<INPUT TYPE=hidden NAME=RecordInDB VALUE="<%=RecordInDB%>">
<INPUT TYPE=hidden NAME=PageRec VALUE="<%=PageRec%>">
<INPUT TYPE=hidden NAME=gTotalPage VALUE="<%=gTotalPage%>">
<INPUT TYPE=hidden NAME=YYMM VALUE="<%=YYMM%>">

<table width="600" border="0" cellspacing="0" cellpadding="0">
	<tr>
	<TD width=430>
	<img border="0" src="../../image/icon.gif" align="absmiddle">
	員工保險與工團費
	計薪年月：<%=YYMM%></TD>
	</tr>
</table>
<hr size=0	style='border: 1px dotted #999999;' align=left width=500>

<TABLE  CLASS="FONT9" BORDER="0" cellspacing="1" cellpadding="1" BGCOLOR="LightGrey" WIDTH=865 >
	<TR HEIGHT=25 BGCOLOR="#C0C0C0"  class=txt8 >
 		<TD WIDTH=40 ALIGN=CENTER>項次<BR>STT</TD>
 		<TD WIDTH=60 align=center>工號<BR>So The</TD>
 		<TD WIDTH=120 NOWRAP >員工姓名(中,英,越)<BR>Ho Ten</TD> 		
 		<td WIDTH=55 align=center nowrap>到職日期<BR>NVX(yy/mm/dd)</td>
 		<td WIDTH=55 align=center nowrap>離職日期<BR>NTV(yy/mm/dd)</td> 		
 		<TD WIDTH=90 align=center nowrap>基本薪資<br>Bậc Lương</TD>
 		<td WIDTH=55 align=center>保險日期<BR>NBH(yy/mm/dd)</td>
 		<td WIDTH=50 nowrap align=center>PHÁT<Br>SINH<Br>Thang</td>
 		<td WIDTH=50 nowrap align=center>THAI<Br>SẢN<Br>Thang</td>
 		<TD WIDTH=60 align=center>BHXH</TD>
 		<TD WIDTH=60 align=center>BHYT</TD>
		<TD WIDTH=60 align=center>BHTN</TD>
 		<TD WIDTH=60 align=center>BHTOT</TD>
 		<td WIDTH=50 align=center>入工團</td>
 		<TD WIDTH=60 align=center>工團費</TD> 		
 	</tr>
	<%for CurrentRow = 1 to PageRec
		IF CurrentRow MOD 2 = 0 THEN
			WKCOLOR="LavenderBlush"
			'wkcolor="#ffffff"
		ELSE
			'WKCOLOR="#DFEFFF"
			'WKCOLOR="#D7EBFF"
			WKCOLOR="#E8F3FF"
			
			'wkcolor="#ffffff"
		END IF
		'if tmpRec(CurrentPage, CurrentRow, 1) <> "" then
	%>
	<TR BGCOLOR=<%=WKCOLOR%> >
		<TD ALIGN=CENTER ><FONT COLOR="<%=tmpRec(CurrentPage, CurrentRow, 25)%>">
		<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then %><%=(CURRENTROW)+((CURRENTPAGE-1)*15)%><%END IF %></FONT>
		</TD>
 		<TD ALIGN=CENTER>  		
 			<%=tmpRec(CurrentPage, CurrentRow, 1)%>	
 			<!--a href="vbscript:view1(<%=tmpRec(CurrentPage, CurrentRow, 17)%>)">
 				<FONT COLOR="<%=tmpRec(CurrentPage, CurrentRow, 25)%>"><%=tmpRec(CurrentPage, CurrentRow, 1)%></FONT>
 			</a-->
 			<input type=hidden name=empid value="<%=tmpRec(CurrentPage, CurrentRow, 1)%>">
 			<input type=hidden name="empautoid" value="<%=tmpRec(CurrentPage, CurrentRow, 17)%>">
 		</TD>
 		<TD  >
 			<a href='vbscript:editmemo(<%=CurrentRow-1%>)'>
 				<FONT COLOR="<%=tmpRec(CurrentPage, CurrentRow, 25)%>"><%=tmpRec(CurrentPage, CurrentRow, 2)%></FONT><br>
 				<font COLOR="<%=tmpRec(CurrentPage, CurrentRow, 25)%>" class=txt8><%=tmpRec(CurrentPage, CurrentRow, 3)%></font>
 			</a>
 		</TD> 		
 		<TD  ALIGN=CENTER nowrap width=55><FONT CLASS=TXT8><%=RIGHT(tmpRec(CurrentPage, CurrentRow, 5),8)%></FONT></TD>
 		<TD  ALIGN=CENTER nowrap width=55><FONT CLASS=TXT8><%=RIGHT(tmpRec(CurrentPage, CurrentRow, 18),8)%></FONT></TD> 		

 		<TD ALIGN=RIGHT>
			<%if tmpRec(CurrentPage, CurrentRow, 1) <> "" then%>
 			<select name='bb' class=txt8 onchange='bbchg(<%=currentrow-1%>)' style='width:95'> 
					<option value="0" <%if cdbl(n_bb)=0 then%>selected<%end if%>>0</option>
 				<% 
 				  n_bb=tmpRec(CurrentPage, CurrentRow, 20)
 				  if tmpRec(CurrentPage, CurrentRow, 20)="" then n_bb=0 
 				  sqlt="select * from  empsalarybasic  where  country='vn' and bwhsno='la' "&_
					   "and func='aa' and (isnull(yymm,'')='' or yymm>=convert(char(6),getdate(),112)  ) "
				  set rds=conn.execute(sqlt)
				  while not rds.eof  				  
 				%>
 				<option value="<%=rds("bonus")%>" <%if cdbl(rds("bonus"))=cdbl(n_bb) then%>selected<%end if%>><%=rds("code")%>-<%=rds("bonus")%></option>
 				<%rds.movenext
 				wend
 				set rds=nothing 
 				%>
 			</select> 
			<%else%>
			<input type=hidden size=5 name="bb"  value="<%=(trim(tmpRec(CurrentPage, CurrentRow, 20)))%>">			
			<%end if%>
			<input type=hidden size=5 name="BBCODE"  value="<%=(trim(tmpRec(CurrentPage, CurrentRow, 19)))%>">			
			
 		</TD>
 		<TD  ALIGN=CENTER nowrap width=55><FONT CLASS=TXT8><%=RIGHT(tmpRec(CurrentPage, CurrentRow, 26),8)%></FONT></TD>
 		<TD ALIGN=CENTER >
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
				<input name=KH1 class=INPUTBOX8r size=5  value="<%=(trim(tmpRec(CurrentPage, CurrentRow, 29)))%>" ONCHANGE="DATACHG(<%=currentrow-1%>)">
			<%else%>	
				<input name=KH1 class=INPUTBOX8r size=5 type=hidden>
			<%end if %>
 		</TD>
 		<TD ALIGN=CENTER >
			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>	
				<input name=chanjia class=INPUTBOX8r size=5  value="<%=(trim(tmpRec(CurrentPage, CurrentRow, 30)))%>" ONCHANGE="DATACHG(<%=currentrow-1%>)" >
			<%else%>	
				<input name=chanjia class=INPUTBOX8r size=5 type=hidden >
			<%end if%>
 		</TD> 		
 		<TD ALIGN=RIGHT>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
 				<INPUT NAME="BHXH" CLASS='INPUTBOX8' SIZE=6 VALUE="<%=tmpRec(CurrentPage, CurrentRow, 21)%>"  STYLE="TEXT-ALIGN:RIGHT"  ONCHANGE="DATACHG1(<%=currentrow-1%>)" > 
 			<%else%>
 				<INPUT NAME="BHXH" TYPE='HIDDEN'>-
 			<%end if%> 		
 		</TD>
 		<TD ALIGN=RIGHT>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
 				<INPUT NAME="BHYT" CLASS='INPUTBOX8' SIZE=6 VALUE="<%=tmpRec(CurrentPage, CurrentRow, 22)%>"  STYLE="TEXT-ALIGN:RIGHT"  ONCHANGE="DATACHG(<%=currentrow-1%>)" > 
 			<%else%>
 				<INPUT NAME="BHYT" TYPE='HIDDEN'>-
 			<%end if%>
 		</TD>
 		<TD ALIGN=RIGHT>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
 				<INPUT NAME="BHTN" CLASS='INPUTBOX8' SIZE=6 VALUE="<%=tmpRec(CurrentPage, CurrentRow, 32)%>"  STYLE="TEXT-ALIGN:RIGHT"  ONCHANGE="DATACHG(<%=currentrow-1%>)" > 
 			<%else%>
 				<INPUT NAME="BHTN" TYPE='HIDDEN'>-
 			<%end if%>
 		</TD>		
 		<TD ALIGN=RIGHT>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
 				<INPUT NAME="BHTOT" CLASS='INPUTBOX8' READONLY  SIZE=7 VALUE="<%=tmpRec(CurrentPage, CurrentRow, 28)%>"  STYLE="TEXT-ALIGN:RIGHT;BACKGROUND-COLOR:LIGHTYELLOW"  > 
 			<%else%>
 				<INPUT NAME="BHTOT" TYPE='HIDDEN'>-
 			<%end if%>
 		</TD> 
 		<TD  ALIGN=CENTER ><FONT CLASS=TXT8><%=tmpRec(CurrentPage, CurrentRow, 27)%></FONT></TD>
 		<TD ALIGN=RIGHT>
 			<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then  %>
 				<INPUT NAME="GTAMT" CLASS='INPUTBOX8' SIZE=6 VALUE="<%=tmpRec(CurrentPage, CurrentRow, 23)%>"  STYLE="TEXT-ALIGN:RIGHT"  ONBLUR="DATACHG(<%=currentrow-1%>)"  > 
 			<%else%>
 				<INPUT NAME="GTAMT" TYPE='HIDDEN'>-
 			<%end if%>
 		</TD>
	</TR>
	<%next%>
</TABLE>
<input type=hidden name=empid>
<input type=hidden name=BBCODE>
<input type=hidden name=BB>
<input type=hidden name=kh1>
<input type=hidden name=chanjia>
<INPUT NAME="BHXH" TYPE='HIDDEN'>
<INPUT NAME="BHYT" TYPE='HIDDEN'>
<INPUT NAME="BHTN" TYPE='HIDDEN'>
<INPUT NAME="GTAMT" TYPE='HIDDEN'>
<INPUT NAME="BHTOT" TYPE='HIDDEN'>

<TABLE border=0 width=500 class=font9 >
<tr>
    <td align="CENTER" height=40 WIDTH=75%>
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
	<FONT CLASS=TXT8>&nbsp;&nbsp;PAGE:<%=CURRENTPAGE%> / <%=TOTALPAGE%> , COUNT:<%=RECORDINDB%></FONT>
	</TD>
	<TD WIDTH=25% ALIGN=RIGHT>
		<input type="BUTTON" name="send" value="確　認" class=button ONCLICK="GO()">
		<input type="BUTTON" name="send" value="取　消" class=button onclick="clr()">
	</TD>
</TR>

</TABLE>
</form>




</body>
</html>
<%
Sub StoreToSession()
	Dim CurrentRow
	tmpRec = Session("empBHGTD")
	for CurrentRow = 1 to PageRec
		'tmpRec(CurrentPage, CurrentRow, 0) = request("op")(CurrentRow)		
		tmpRec(CurrentPage, CurrentRow, 19) = request("BBCODE")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 20) = request("BB")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 21) = request("BHXH")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 22) = request("BHYT")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 23) = request("GTAMT")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 28) = request("BHTOT")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 29) = request("kh1")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 30) = request("chanjia")(CurrentRow)
	next
	Session("empBHGTD") = tmpRec

End Sub
%>

<script language=vbscript>
function BACKMAIN()
	open "../main.asp" , "_self"
end function

function clr()
	open "EMPBHGT.fore.asp" , "_self"
end function

function go()
	<%=self%>.action="EMPBHGT.upd.asp"
	<%=self%>.submit()
end function

function oktest(N)
	tp=<%=self%>.totalpage.value
	cp=<%=self%>.CurrentPage.value
	rc=<%=self%>.RecordInDB.value
	open "empfile.show.asp?empautoid="& N , "_blank" , "top=10, left=10, width=550, scrollbars=yes"
end function

FUNCTION BBCODECHG(INDEX)
	codestr=<%=self%>.bbcode(index).value	
	open "empbhgt.back.asp?ftype=A&index="& index &"&CurrentPage="& <%=CurrentPage%> & _
		 "&code=" &	codestr , "Back"	
	'PARENT.BEST.COLS="70%,30%"
END FUNCTION


FUNCTION bbchg(INDEX)
	codestr=<%=self%>.bb(index).value	
	<%=self%>.BHXH(index).value = cdbl(codestr)*0.05
	<%=self%>.BHYT(index).value = cdbl(codestr)*0.01
	<%=self%>.BHTOT(index).value = cdbl(<%=self%>.BHXH(index).value)+cdbl(<%=self%>.BHYT(index).value)
	'open "empbhgt.back.asp?ftype=A&index="& index &"&CurrentPage="& <%=CurrentPage%> & _
	'	 "&code=" &	codestr , "Back"	
	'PARENT.BEST.COLS="70%,30%"
END FUNCTION



FUNCTION DATACHG(INDEX)
	if isnumeric(<%=SELF%>.BHXH(INDEX).VALUE)=false then
		alert "請輸入數字can so!!"
		<%=self%>.BHXH(index).focus()
		<%=self%>.BHXH(index).value=0
		<%=self%>.BHXH(index).select()
		exit FUNCTION
	end if
	if isnumeric(<%=SELF%>.BHYT(INDEX).VALUE)=false then
		alert "請輸入數字can so!!"
		<%=self%>.BHYT(index).focus()
		<%=self%>.BHYT(index).value=0		
		<%=self%>.BHYT(index).select()
		exit FUNCTION
	end if
	if isnumeric(<%=SELF%>.BHTN(INDEX).VALUE)=false then
		alert "請輸入數字can so!!"
		<%=self%>.BHTN(index).focus()
		<%=self%>.BHTN(index).value=0		
		<%=self%>.BHTN(index).select()
		exit FUNCTION
	end if	
	if isnumeric(<%=SELF%>.GTAMT(INDEX).VALUE)=false then
		alert "請輸入數字can so!!"
		<%=self%>.GTAMT(index).focus()
		<%=self%>.GTAMT(index).value=0		
		<%=self%>.GTAMT(index).select()
		exit FUNCTION
	end if  	 
	
	if isnumeric(<%=SELF%>.kh1(INDEX).VALUE)=false then
		alert "請輸入數字can so!!"
		<%=self%>.kh1(index).focus()
		<%=self%>.kh1(index).value=0		
		<%=self%>.kh1(index).select()
		exit FUNCTION
	end if  	 

	if isnumeric(<%=SELF%>.chanjia(INDEX).VALUE)=false then
		alert "請輸入數字can so!!"
		<%=self%>.chanjia(index).focus()
		<%=self%>.chanjia(index).value=0		
		<%=self%>.chanjia(index).select()
		exit FUNCTION
	end if  	 	
	
	<%=SELF%>.BHTOT(INDEX).VALUE=CDBL(<%=self%>.BHXH(index).value)+CDBL(<%=self%>.BHYT(index).value)+CDBL(<%=self%>.BHTN(index).value)

	CODESTR01 = <%=SELF%>.BB(INDEX).VALUE
	CODESTR02 = <%=SELF%>.BBCODE(INDEX).VALUE
	CODESTR03 = <%=SELF%>.BHXH(INDEX).VALUE
	CODESTR04 = <%=SELF%>.BHYT(INDEX).VALUE
	CODESTR05 = <%=SELF%>.GTAMT(INDEX).VALUE   
	CODESTR06 = <%=SELF%>.kh1(INDEX).VALUE   
	CODESTR07 = <%=SELF%>.chanjia(INDEX).VALUE   
	CODESTR08 = <%=SELF%>.BHTN(INDEX).VALUE   
	

	open "empbhgt.back.asp?ftype=CDATACHG&index="&index &"&CurrentPage="& <%=CurrentPage%> & _
		 "&CODESTR01="& CODESTR01 &_
		 "&CODESTR02="& CODESTR02 &_
		 "&CODESTR03="& CODESTR03 &_
		 "&CODESTR04="& CODESTR04 &_
		 "&CODESTR05="& CODESTR05 &_
		 "&CODESTR06="& CODESTR06 &_
		 "&CODESTR07="& CODESTR07 &_
		 "&CODESTR08="& CODESTR08  , "Back"

	'PARENT.BEST.COLS="70%,30%"

END FUNCTION  


FUNCTION DATACHG1(INDEX)
	if isnumeric(<%=SELF%>.BHXH(INDEX).VALUE)=false then
		alert "請輸入數字!!"
		<%=self%>.BHXH(index).focus()
		<%=self%>.BHXH(index).value=0
		<%=self%>.BHXH(index).select()
		exit FUNCTION
	end if  

	CODESTR01 = <%=SELF%>.BB(INDEX).VALUE
	CODESTR02 = <%=SELF%>.BBCODE(INDEX).VALUE
	CODESTR03 = <%=SELF%>.BHXH(INDEX).VALUE
	CODESTR04 = <%=SELF%>.BHYT(INDEX).VALUE
	CODESTR05 = <%=SELF%>.GTAMT(INDEX).VALUE   
 
	IF <%=SELF%>.BHXH(INDEX).VALUE="0" THEN 
		<%=SELF%>.BHYT(INDEX).VALUE="0"
		CODESTR04 = "0"
	END IF 
	<%=SELF%>.BHTOT(INDEX).VALUE=CDBL(<%=self%>.BHXH(index).value)+CDBL(<%=self%>.BHYT(index).value)
	
	open "empbhgt.back.asp?ftype=CDATACHG1&index="&index &"&CurrentPage="& <%=CurrentPage%> & _
		 "&CODESTR01="& CODESTR01 &_
		 "&CODESTR02="& CODESTR02 &_
		 "&CODESTR03="& CODESTR03 &_
		 "&CODESTR04="& CODESTR04 &_
		 "&CODESTR05="& CODESTR05  , "Back"	
	'PARENT.BEST.COLS="70%,30%"

END FUNCTION  

function view1(index)
	yymmstr = <%=self%>.yymm.value
	'alert yymmstr
	empidstr = <%=self%>.empid(index).value
	idstr= <%=SELF%>.empautoid(INDEX).VALUE
	OPEN "../zzz/getempWorkTime.asp?yymm=" & yymmstr &"&EMPID=" & empidstr , "_blank" , "top=10, left=10,  scrollbars=yes"
end function


function editmemo(index)
	tp=<%=self%>.totalpage.value
	cp=<%=self%>.CurrentPage.value
	rc=<%=self%>.RecordInDB.value
	YYMM = <%=self%>.YYMM.value
	open "<%=self%>.memo.asp?index="& index &"&currentpage=" & cp &"&yymm=" & yymm  , "_blank" , "top=10, left=10, width=450,height=450, scrollbars=yes"
end function  


</script>


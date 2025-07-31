<%@Language=VBScript codepage=65001%>
<!-------- #include file ="../GetSQLServerConnection.fun" --------->
<!-----#include file="../ADOINC.inc"------>
<!--#include file="../include/sideinfo.inc"-->
<%
session.codepage=65001
Dim gTotalPage, PageRec, TableRec
Dim CurrentRow, CurrentPage, TotalPage, RecordInDB
Dim tmpRec, i, j, k, SELF, conn, rs, Source
Dim WK_COLOR, StartToAdd

SELF = "YEAAE0101"
gTotalPage = 10
PageRec = 20    'number of records per page
TableRec = 7    'number of fields per record

Set CONN = GetSQLServerConnection()

A1 = Trim(request("A1"))
if request("TotalPage") = "" or request("TotalPage") = "0" then 
	CurrentPage = 1	
	Source = "select * from BASICCODE where func like '%"& Trim(request("a1")) &"' order by FUNC , SYS_TYPE "
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open Source, conn, 3		
	if not rs.eof then 
		RecordInDB = rs.RecordCount 
		PageRec = rs.RecordCount  + 10 
		rs.PageSize = PageRec  
		TotalPage = rs.PageCount 	
		gTotalPage =TotalPage 
	end if	
	Redim tmpRec(gTotalPage, PageRec, TableRec)   'Array	
	for i = 1 to gTotalPage 
		for j = 1 to PageRec
			if not rs.EOF then 				
				tmpRec(i, j, 0) = "no"
				tmpRec(i, j, 1) = rs("FUNC")				  	
				tmpRec(i, j, 2) = rs("SYS_TYPE")
				tmpRec(i, j, 3) = rs("SYS_VALUE")
				tmpRec(i, j, 4) = rs("AUTOID")
				if Trim(request("a1"))<>"" then 
					tmpRec(i, j, 5) = ""
				else
					tmpRec(i, j, 5) = ""
				end if	
				
			  	rs.MoveNext 
		   else 
				tmpRec(i, j, 1) = Trim(request("a1"))
				tmpRec(i, j, 5) = ""
			  	'exit for 
		   end if 			
		next		
		if rs.EOF then 
			rs.Close 
		  	Set rs = nothing
		  	exit for 
		end if 
	next 
	Session("ADMIN01") = tmpRec	
else
	TotalPage = cint(request("TotalPage"))
	gTotalPage = cint(request("gTotalPage"))
	'RESPONSE.WRITE "..."
	StoreToSession()
	RecordInDB  = REQUEST("RecordInDB")
	CurrentPage = cint(request("CurrentPage"))
	tmpRec = Session("ADMIN01")
	
	Select case request("send") 
	     Case "FIRST"
		      CurrentPage = 1			
	     Case "BACK"
		      if cint(CurrentPage) <> 1 then 
			     CurrentPage = CurrentPage - 1				
		      end if
	     Case "NEXT"
		      if cint(CurrentPage) < cint(gTotalPage) then 
			     CurrentPage = CurrentPage + 1 
		      end if			
	     Case "END"
		      CurrentPage = gTotalPage 			
	     Case Else 
		      CurrentPage = 1	
	end Select 
end if 
%>
<HTML>
<HEAD>
<meta HTTP-EQUIV="Content-Type" content="text/html; charset=utf-8">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
<link rel="stylesheet" href="../Include/style.css" type="text/css">
<link rel="stylesheet" href="../Include/style2.css" type="text/css">
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--
'-----------------enter to next field
function enterto()
	if window.event.keyCode = 13 then window.event.keyCode =9
end function

-->
</SCRIPT>
</HEAD>
<body background="bg_blue.gif"  topmargin=5 onload=f()  onkeydown="enterto()" >
<FORM action="<%=SELF%>.FORE.asp" method="POST" name="<%=SELF%>">
<table width="460" border="0" cellspacing="0" cellpadding="0">
	<tr><TD width=100%>
	<img border="0" src="../image/icon.gif" align="absmiddle">
	代碼檔維護 </TD></tr>
</table>
<hr size=0	style='border: 1px dotted #999999;' align=left width=500>
<INPUT TYPE=HIDDEN NAME="UID" VALUE=<%=SESSION("NETUSER")%>>
<INPUT TYPE=hidden NAME=TotalPage VALUE="<%=TotalPage%>">
<INPUT TYPE=hidden NAME=CurrentPage VALUE="<%=CurrentPage%>">
<INPUT TYPE=hidden NAME=PageRec VALUE="<%=PageRec%>">
<INPUT TYPE=hidden NAME=RecordInDB VALUE="<%=RecordInDB%>">
<INPUT TYPE=hidden NAME=gTotalPage VALUE="<%=gTotalPage%>"> 
<table width=600 ><tr><td align=center>	
		<table width=480 >
		<tr>
			<td>類別：
				<select name="A1" class=inputbox onchange=a1chg()>
					<option value="" <%if a1="" then %>selected<%end if%>>全部顯示</option>
					<%SQL="SELECT FUNC  FROM BASICCODE GROUP BY  FUNC ORDER BY FUNC "
					SET RDS=CONN.execute(SQL)
					WHILE NOT RDS.EOF
					%>	
						<OPTION VALUE="<%=RDS("FUNC")%>" <%if a1=RDS("FUNC") then %>selected<%end if%> ><%=RDS("FUNC")%></OPTION>
					<%
					RDS.MOVENEXT
					WEND
					'SET RDS=NOTHING
					%>	
				</select>
			</td>
		</td>
		</table>
		<TABLE WIDTH=480   > 
	 	<TR BGCOLOR="Gainsboro" class=txt>
	 		<TD WIDTH=50 height=25>刪除</TD>
	 		<TD WIDTH=50>大類</TD>
	 		<TD WIDTH=50 >代碼</TD>
	 		<TD WIDTH=280 >說明</TD>
	 	</TR>
	 	 <%
		for CurrentRow = 1 to PageRec
		
		j = 1
		if j=1 then 
			wk_color = ""
			j = 0
		else 
			wk_color = "#E4E4E4"
			j = 1
		end if 
      %> 
	 	<TR>
	 		<TD>
	 			<%if tmpRec(CurrentPage, CurrentRow, 0)="del" then %>
	 				<INPUT TYPE=CHECKBOX NAME=FUNC checked onclick="del(<%=CurrentRow-1%>)">
	 				<INPUT TYPE=HIDDEN NAME=OP VALUE="del" >
	 			<%else%>
	 				<INPUT TYPE=CHECKBOX NAME=FUNC onclick="del(<%=CurrentRow-1%>)">
	 				<INPUT TYPE=HIDDEN NAME=OP VALUE="UPD" >
	 			<%end if%>	
	 			<INPUT TYPE=HIDDEN NAME=autoid VALUE="<%= tmpRec(CurrentPage, CurrentRow, 4)%>" >
	 		</TD>
	 		<TD><INPUT NAME=FuncTYPE VALUE="<%= tmpRec(CurrentPage, CurrentRow, 1)%>" SIZE=10 CLASS=INPT onchange="datachg(<%=currentrow-1%>)" <%=tmpRec(CurrentPage, CurrentRow, 5)%>></TD>
	 		<TD><INPUT NAME=SYSTYPE VALUE="<%= tmpRec(CurrentPage, CurrentRow, 2)%>" SIZE=10 CLASS=INPT onchange="datachg(<%=currentrow-1%>)" ></TD>
	 		<TD><INPUT NAME=SYSVALUE VALUE="<%= tmpRec(CurrentPage, CurrentRow, 3)%>" SIZE=40 CLASS=INPT onchange="datachg(<%=currentrow-1%>)" > </TD>
	 	</TR>
	 	<%next%> 
	 </TABLE>	
	 <TABLE border=0 width=460  >
		<tr>
		    <td align="left" CLASS=FONT9 >
		    頁次:<%=CURRENTPAGE%> / <%=GTOTALPAGE%> , COUNT:<%=RECORDINDB%><BR>
			<% If CurrentPage > 1 Then %>
			<input type="submit" name="send" value="FIRST" class=button>
			<input type="submit" name="send" value="BACK" class=button>
			<% Else %>
			<input type="submit" name="send" value="FIRST" disabled class=button>
			<input type="submit" name="send" value="BACK" disabled class=button>
			<% End If %>
		
			<% If cint(CurrentPage) < cint(GTotalPage) Then %>
			<input type="submit" name="send" value="NEXT" class=button>
			<input type="submit" name="send" value="END" class=button>
			<% Else %>      
			<input type="submit" name="send" value="NEXT" disabled class=button>
			<input type="submit" name="send" value="END" disabled class=button>	
			<% End If %>
			</td>
			<td align=right width=150 CLASS=FONT9>
				<BR>
				<input type="button" name="send" value="Confirm"  class=button onclick=go()>		
				<input type="button" name="send" value="Reset"  class=button>		
			</td>			
		</TR>
	</TABLE>		
</td></tr></table> 
 </FORM>

</BODY>
</HTML>
<%
Sub StoreToSession()
	Dim CurrentRow
	tmpRec = Session("ADMIN01")
	for CurrentRow = 1 to PageRec
		tmpRec(CurrentPage, CurrentRow, 0) = request("op")(CurrentRow)		
		tmpRec(CurrentPage, CurrentRow, 1) = request("FuncTYPE")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 2) = request("SYSTYPE")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 3) = request("SYSVALUE")(CurrentRow)
	next 
	Session("ADMIN01") = tmpRec
End Sub
%> 
<script ID=clientEventHandlersVBS language="vbscript">
function Go()
	<%=self%>.action="<%=self%>.updateDB.asp"
	<%=self%>.submit
end function

function f()
	'<%=self%>.muser.focus()
end function

function chg1()
	'<%=self%>.muser.value=Ucase(<%=self%>.muser.value)
end function

function a1chg()
	<%=self%>.totalpage.value="0" 
	<%=self%>.action="<%=self%>.fore.asp"
	<%=self%>.submit()
end function   

function datachg(index)	
	<%=self%>.op(index).value="upd"
	CurrentPage = <%=self%>.CurrentPage.value
    codestr01 = (trim(<%=SELF%>.autoid(index).value))
    codestr02 = Ucase(trim(<%=SELF%>.FuncTYPE(index).value))
    codestr03 = Ucase(trim(<%=SELF%>.SYSTYPE(index).value))
    codestr04 = Ucase(trim(<%=SELF%>.SYSVALUE(index).value))
	open "<%=SELF%>.Back.asp?codestr01=" & codestr01 & _            
			 "&codestr02=" & codestr02 &_
			 "&codestr03=" & codestr03 &_
			 "&codestr04=" & codestr04 &_
		     "&CurrentPage="& CurrentPage &_
		     "&index=" & index & "&func=upd", "Back"	     	     
		   'parent.best.cols="50%, 50%" 
	 
end function 

</script>



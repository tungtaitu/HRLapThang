<%@LANGUAGE=VBSCRIPT CODEPAGE=950%>
<!-- #include file = "../../GetSQLServerConnection.fun" --> 
<!-- #include file="../../ADOINC.inc" --> 
<%
Dim gTotalPage, PageRec, TableRec
Dim CurrentRow, CurrentPage, TotalPage, RecordInDB
Dim tmpRec, i, j, k, SELF, conn, rs, Source
Dim WK_COLOR, StartToAdd

SELF = "YFYEXRT"
gTotalPage = 10
PageRec = 10    'number of records per page
TableRec = 3    'number of fields per record
Query = request("Query")
if trim(Query) = "" then
   if request("TotalPage") = "" or request("TotalPage") = "0" then 
      CurrentPage = 1
	  Set conn = GetSQLServerConnection()	
	  Source = "select * from VYFYEXRT order by yyyymm desc ,code"
	 
	  Set rs = Server.CreateObject("ADODB.Recordset")
	  rs.Open Source, conn, 3, 3
		if not rs.eof then 
	  rs.PageSize = PageRec 
	  RecordInDB = rs.RecordCount 
	  TotalPage = rs.PageCount 
		gtotalpage= TotalPage +1 
		end if 
	

	
	  Redim tmpRec(gTotalPage, PageRec, TableRec)   'Array
	
	  for i = 1 to TotalPage 
	 	 for j = 1 to PageRec
			if not rs.EOF then 				
				tmpRec(i, j, 0) = "no"
				tmpRec(i, j, 1) = trim(rs("yyyymm"))
				tmpRec(i, j, 2) = trim(rs("code"))
				tmpRec(i, j, 3) = rs("exrt")				
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
	  Session("YSBAE1701") = tmpRec	
   else    
	  TotalPage = cint(request("TotalPage"))
		gTotalPage = cint(request("gTotalPage"))
	  StoreToSession()
	  CurrentPage = cint(request("CurrentPage"))
	  tmpRec = Session("YSBAE1701")
	
	  Select case request("send") 
		     Case "第一頁"
			      CurrentPage = 1			
		     Case "上一頁"
			      if cint(CurrentPage) <> 1 then 
				     CurrentPage = CurrentPage - 1				
			      end if
		     Case "下一頁"
 			      if cint(CurrentPage) <= cint(gTotalPage) then 
				     CurrentPage = CurrentPage + 1 
			      end if			
		     Case "最末頁"
			      CurrentPage = TotalPage 			
		     Case Else 
			      CurrentPage = 1	
	  end Select 
   end if
else
   CurrentPage = 1
   Set conn = GetSQLServerConnection()	
   Source = "Select * from VYFYEXRT Where Code like '" & trim(Query) & "%' " & _
            "or yyyymm like '" & trim(Query) & "%' " & _
            "order by yyyymm desc ,Code "

   Set rs = Server.CreateObject("ADODB.Recordset")
   rs.Open Source, conn, 3, 3
	
		if not rs.eof then 
	  rs.PageSize = PageRec 
	  RecordInDB = rs.RecordCount 
	  TotalPage = rs.PageCount 
		gtotalpage= TotalPage +1 
		end if 
	 
	
   Redim tmpRec(gTotalPage, PageRec, TableRec)   'Array
	
   for i = 1 to TotalPage 
 	  for j = 1 to PageRec
		 if not rs.EOF then 			
			tmpRec(i, j, 0) = "no"
			tmpRec(i, j, 1) = trim(rs("yyyymm"))
			tmpRec(i, j, 2) = trim(rs("code"))
			tmpRec(i, j, 3) = rs("exrt")			
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
   Session("YSBAE1701") = tmpRec	   
end if   
	  Set conn = nothing 
%>
<html>
<head>
<meta HTTP-EQUIV="Content-Type" content="text/html; charset=BIG5">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="refresh">
<link rel="stylesheet" href="../../Include/style.css" type="text/css">
<link rel="stylesheet" href="../../Include/style2.css" type="text/css"> 

<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--
function m(index)
   <%=SELF%>.send(index).style.backgroundcolor="lightyellow"
   <%=SELF%>.send(index).style.color="red"
end function

function n(index)
   <%=SELF%>.send(index).style.backgroundcolor="khaki"
   <%=SELF%>.send(index).style.color="black"
end function

function Q_Data()
   <%=self%>.TotalPage.value = ""
   <%=self%>.submit
end function
-->
</SCRIPT>
<title>稅別</title></head>
<body background="bg_blue.gif"  topmargin=5>
<form method="post" action="<%=SELF%>.Fore.asp" name="<%=SELF%>">
<INPUT TYPE=HIDDEN NAME="UID" VALUE=<%=SESSION("NETUSER")%>>
<INPUT TYPE=hidden NAME=TotalPage VALUE="<%=TotalPage%>">
<INPUT TYPE=hidden NAME=CurrentPage VALUE="<%=CurrentPage%>">
<INPUT TYPE=hidden NAME=PageRec VALUE="<%=PageRec%>">
<INPUT TYPE=hidden NAME=RecordInDB VALUE="<%=RecordInDB%>">
<INPUT TYPE=hidden NAME=gTotalPage VALUE="<%=gTotalPage%>">
<table width="460" border="0" cellspacing="0" cellpadding="0">
	<tr><TD>
        <img border="0" src="../../image/icon.gif" align="absmiddle"> 
        匯率檔 </TD>	 
	</tr>
</table>
<hr size=0	style='border: 1px dotted #999999;' align=left width=500>		  
<table width=550><tr><td align=center> 
	<TABLE WIDTH=350 BORDER=0  cellpadding=0 cellspacing=3> 
      <tr bgcolor=#CCCCCC> 
        <td width=30 align=center class=txt9 height=20>刪除</td>
        <td width=125 align=center class=txt9>年月</td>
        <td width=125 align=center class=txt9>幣別</td>
        <td width=125 align=center class=txt9>匯率</td>
      </tr>
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
      <TR bgcolor="<%=wk_color%>">
        <td width="30" align=center valign="top"> 
			<%if tmpRec(CurrentPage, CurrentRow, 0) = "del" then%>
				<input type=checkbox name=func value=del onclick="del(<%=CurrentRow - 1%>)" checked>
				<input type=hidden name=op value=del>
			<%else%>
				<input type=checkbox name=func value=no onclick="del(<%=CurrentRow - 1%>)" <%=mode%>>
				<input type=hidden name=op value=no>
			<%end if%>				
        </td>
        <TD width="125" align=center valign="top"> 
			<%if tmpRec(CurrentPage, CurrentRow, 1) = "" then%>
				 <input size=8 class=inputbox maxlength=6 name=yyyymm value="<%=Ucase(trim(tmpRec(CurrentPage, CurrentRow, 1)))%>" >
				 </TD>
				 <TD width="125" align=center valign="top"> 
				 <INPUT size=5 class=inputbox maxlength=3 name=code value="<%=Trim(Ucase(tmpRec(CurrentPage, CurrentRow, 2)))%>" onchange="code_change(<%=CurrentRow - 1%>)">
				 </TD>
				 <TD width="125" align=center valign="top"> 
				 <INPUT size=15 class=inputbox maxlength=15 name=exrt value="<%=trim(Ucase(tmpRec(CurrentPage, CurrentRow, 3)))%>" onchange="exrt_change(<%=CurrentRow - 1%>)">
				 </TD>
			<%elseif Ucase(trim(tmpRec(CurrentPage, CurrentRow, 1)))>=(year(date) & right(0 & month(date),2)) then %>
				<input size=8 maxlength=6 class=inputbox name=yyyymm Readonly style="BACKGROUND-COLOR: lightyellow" value="<%=Ucase(trim(tmpRec(CurrentPage, CurrentRow, 1)))%>">				 
				 </TD>
				 <TD width="125" align=center valign="top"> 
				<input size=5 maxlength=3 class=inputbox name=code Readonly style="BACKGROUND-COLOR: lightyellow" value="<%=Ucase(trim(tmpRec(CurrentPage, CurrentRow, 2)))%>">				 
				 </TD>
				 <TD width="125" align=center valign="top"> 
				<input size=15 maxlength=15 class=inputbox name=exrt value="<%=Trim(tmpRec(CurrentPage, CurrentRow, 3))%>" onchange="exrt_change(<%=CurrentRow - 1%>)">				 
				 </TD>
			<%else%>
				<input size=8 maxlength=6 class=inputbox name=yyyymm Readonly style="BACKGROUND-COLOR: lightyellow" value="<%=Ucase(trim(tmpRec(CurrentPage, CurrentRow, 1)))%>">				 
				 </TD>
				 <TD width="125" align=center valign="top"> 
				<input size=5 maxlength=3 class=inputbox name=code Readonly style="BACKGROUND-COLOR: lightyellow" value="<%=Ucase(trim(tmpRec(CurrentPage, CurrentRow, 2)))%>">				 
				 </TD>
				 <TD width="125" align=center valign="top"> 
				<input size=15 maxlength=15 class=inputbox name=exrt Readonly style="BACKGROUND-COLOR: lightyellow" value="<%=tmpRec(CurrentPage, CurrentRow, 3)%>">				 
				 </TD>
			<%end if%>
      </TR>
      <%next%> 
      <tr>
      	<td colspan=4 class=txt9 align=center>
      	第<INPUT size=6 readonly class=readonlyN name=ta style='BORDER-BOTTOM-STYLE: none; BORDER-LEFT: medium none;BORDER-RIGHT-STYLE: none; BORDER-TOP-STYLE: none' value="<%=CurrentPage%>">頁,
      	共<INPUT size=6 readonly class=readonlyN name=ta style='BORDER-BOTTOM-STYLE: none; BORDER-LEFT: medium none;BORDER-RIGHT-STYLE: none; BORDER-TOP-STYLE: none' value="<%=Totalpage%>">頁
      	</td>
      </tr>
    </TABLE>
</td></tr></table>
<br>
<TABLE border=0 width=550 >
<tr>
    <td align="left"  >
<% If CurrentPage > 1 Then %>
	<input type="submit" name="send" value="第一頁" class=button>
	<input type="submit" name="send" value="上一頁" class=button>
<% Else %>
	<input type="submit" name="send" value="第一頁" disabled class=button>
	<input type="submit" name="send" value="上一頁" disabled class=button>
<% End If %>

<% If cint(CurrentPage) <= cint(TotalPage) Then %>
	<input type="submit" name="send" value="下一頁" class=button>
	<input type="submit" name="send" value="最末頁" class=button>
<% Else %>      
	<input type="submit" name="send" value="下一頁" disabled class=button>
	<input type="submit" name="send" value="最末頁" disabled class=button>	
<% End If %>
	</td>
    <td  >
    
	<INPUT TYPE="button" name=send VALUE="確    認" class=button onClick="Go()">	
	<INPUT TYPE="button" name=send VALUE="取    消" class=button onClick="Clear()">
	</TD>
</TR>
</TABLE>

</form>
</body>
</html>
<%
Sub StoreToSession()
	Dim CurrentRow
	tmpRec = Session("YSBAE1701")
	for CurrentRow = 1 to PageRec
		tmpRec(CurrentPage, CurrentRow, 0) = request("op")(CurrentRow)		
		tmpRec(CurrentPage, CurrentRow, 1) = request("yyyymm")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 2) = request("code")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 3) = request("exrt")(CurrentRow)
	next 
	Session("YSBAE1701") = tmpRec
End Sub
%>

<script language="vbscript">
<!--

function Go()
   <%=SELF%>.action = "<%=SELF%>.UpdateDB.asp"
   <%=SELF%>.submit
end function

function Clear()
	open "<%=SELF%>.asp", "Fore"
end function

function del(index)
	TaxCode_str = Ucase(trim(<%=SELF%>.code(index).value))
	Description_str = Ucase(trim(<%=SELF%>.yyyymm(index).value))
	Percentage_str = Ucase(trim(<%=SELF%>.exrt(index).value))
	 if <%=SELF%>.func(index).checked then 
		<%=SELF%>.op(index).value = "del"		
		open "<%=SELF%>.Back.asp?code=" & TaxCode_str & _            
		     "&CurrentPage=" & <%=CurrentPage%> & _
		     "&index=" & index & "&func=del", "Back"
	 else
	    <%=SELF%>.op(index).value = "update"		
		open "<%=SELF%>.Back.asp?code=" & TaxCode_str & _            
         "&CurrentPage=" & <%=CurrentPage%> & _
         "&yyyymm=" & Description_str & _
         "&exrt=" & Percentage_str & _
         "&index=" & index & "&func=TaxCode_change", "Back"
	 end if		
end function

function code_change(index)
    TaxCode_str = Ucase(trim(<%=SELF%>.code(index).value))
	<%=SELF%>.code(index).value = TaxCode_str
	Data_change(index)
end function

function exrt_change(index)
    Percentage_str = Ucase(trim(<%=SELF%>.exrt(index).value))
    <%=SELF%>.exrt(index).value = Percentage_str
    if Percentage_str <> "" then 
		if not isnumeric(Percentage_str) then 
			alert "匯率資料輸入錯誤!!"
			<%=SELF%>.exrt(index).value = "0"
			<%=SELF%>.exrt(index).focus()		
			Percentage_str = "0"
		end if 
	end if
	Data_change(index)
end function

    
function Data_change(index)
	<%=SELF%>.op(index).value = "update"
	TaxCode_str = Ucase(trim(<%=SELF%>.code(index).value))
	Description_str = Ucase(trim(<%=SELF%>.yyyymm(index).value))
	Percentage_str = Ucase(trim(<%=SELF%>.exrt(index).value))
	open "<%=SELF%>.Back.asp?code=" & TaxCode_str & _
         "&CurrentPage=" & <%=CurrentPage%> & _
         "&yyyymm=" & Description_str & _
         "&exrt=" & Percentage_str & _
         "&index=" & index & "&func=TaxCode_change", "Back"
end function

//-->
</script>



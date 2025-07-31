<%@LANGUAGE=VBSCRIPT CODEPAGE=950%> 
<!-- #include file = "../GetSQLServerConnection.fun" --> 
<!-- #include file="../ADOINC.inc" -->
<%
'Session("NETUSER")=""
'SESSION("NETUSERNAME")=""

Dim gTotalPage, PageRec, TableRec
Dim CurrentRow, CurrentPage, TotalPage, RecordInDB
Dim tmpRec, i, j, k, SELF, conn, rs, Source
Dim WK_COLOR, StartToAdd

SELF = "ADMIN01"
gTotalPage = 10
PageRec = 20    'number of records per page
TableRec = 7    'number of fields per record

A1 = Trim(request("A1"))
if request("TotalPage") = "" or request("TotalPage") = "0" then 
	CurrentPage = 1
	Set conn = GetSQLServerConnection()
	Source = "select * from BASICCODE where func like '%"& Trim(request("a1")) &"%' order by FUNC , SYS_TYPE "
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open Source, conn, 3
	rs.PageSize = PageRec 
	RecordInDB = rs.RecordCount 
	TotalPage = rs.PageCount 	
	Set conn = nothing 	
	Redim tmpRec(gTotalPage, PageRec, TableRec)   'Array	
	for i = 1 to gTotalPage 
	   for j = 1 to PageRec
		   if not rs.EOF then 
			  for k=1 to TableRec-1
				  tmpRec(i, j, 0) = "no"
				  tmpRec(i, j, 1) = rs("FUNC")
				  tmpRec(i, j, 2) = rs("SYS_TYPE")
				  tmpRec(i, j, 3) = rs("SYS_VALUE")
				  tmpRec(i, j, 4) = rs("AUTOID")
			  next
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
	Session("ADMIN01") = tmpRec	
else
	TotalPage = cint(request("TotalPage"))
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
<html>
<head>
<meta HTTP-EQUIV="Content-Type" content="text/html; charset=big5">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
<link rel="stylesheet" href="../Include/style.css" type="text/css">
<link rel="stylesheet" href="../Include/style2.css" type="text/css">
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}

function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}

function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}

function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}

function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}

function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}

function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}

function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}

function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}

function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}

function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}

function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}

function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}

function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}

function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}

function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}

function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}

function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}

function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}

function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}

function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}

function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}

function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}

function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}

function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}

function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}

function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}

function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}

function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}

function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}

function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}

function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}
//-->
</script>
<title>VYFYNET</title></head>  
<body leftmargin="0"  topmargin="0"  marginwidth="0" marginheight="0" onLoad="MM_preloadImages('../picture/BTN02.gif','../picture/BTNA2.GIF','../picture/BTNb2.gif','../picture/BTNC2.GIF','../picture/BTND2.gif','../picture/BTNe2.gif','../picture/BTNF2.gif','../picture/BTNG2.GIF')" onkeydown="enterto()" >
<FORM NAME="<%=SELF%>" METHOD=POST ACTION="ADMIN01.ASP"> 
<INPUT TYPE=HIDDEN NAME="UID" VALUE=<%=SESSION("NETUSER")%>>
<INPUT TYPE=hidden NAME=TotalPage VALUE="<%=TotalPage%>">
<INPUT TYPE=hidden NAME=CurrentPage VALUE="<%=CurrentPage%>">
<INPUT TYPE=hidden NAME=PageRec VALUE="<%=PageRec%>">
<INPUT TYPE=hidden NAME=RecordInDB VALUE="<%=RecordInDB%>">
<INPUT TYPE=hidden NAME=gTotalPage VALUE="<%=gTotalPage%>"> 
<table width="760" border="0" cellspacing="0" cellpadding="0" align=center>
	<tr> 
      <td><img src="../picture/LOGONEW.gif" width="760" height="76"></td>
    </tr>
</table>
<table width="760" border="0" cellspacing="0" cellpadding="0" align=center >
   	<tr>      
		<td><a href="vbscript:chgpage(1)" onMouseOver="MM_swapImage('A1','','../picture/BTN02.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="../picture/BTN01.gif" name="A1" width="95" height="24" border="0" id="A1"></a></td>
		<td><a href="vbscript:chgpage(2)" onMouseOver="MM_swapImage('A2','','../picture/BTNA2.GIF',1)" onMouseOut="MM_swapImgRestore()"><img src="../picture/BTNA1.GIF" name="A2" width="95" height="24" border="0" id="A2"></a></td>
		<td><a href="vbscript:chgpage(3)" onMouseOver="MM_swapImage('A3','','../picture/BTNB2.GIF',1)" onMouseOut="MM_swapImgRestore()"><img src="../picture/BTNB1.GIF" name="A3" width="95" height="24" border="0" id="A3"></a></td>		
		<td><a href="vbscript:chgpage(4)" onMouseOver="MM_swapImage('A4','','../picture/BTNC2.GIF',1)" onMouseOut="MM_swapImgRestore()"><img src="../picture/BTNC1.gif" name="A4" width="95" height="24" border="0" id="A4"></a></td>
		<td><a href="vbscript:chgpage(5)" onMouseOver="MM_swapImage('A5','','../picture/BTND2.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="../picture/BTND1.gif" name="A5" width="95" height="24" border="0" id="A5"></a></td>
		<td><a href="vbscript:chgpage(6)" onMouseOver="MM_swapImage('A6','','../picture/BTNe2.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="../picture/BTNe1.gif" name="A6" width="95" height="24" border="0" id="A6"></a></td>
		<td><a href="vbscript:chgpage(7)" onMouseOver="MM_swapImage('A7','','../picture/BTNf2.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="../picture/BTNf1.gif" name="A7" width="95" height="24" border="0" id="A7"></a></td>		
		<td><a href="vbscript:chgpage(8)" onMouseOver="MM_swapImage('A8','','../picture/BTNG2.GIF',1)" onMouseOut="MM_swapImgRestore()"><img src="../picture/BTNG1.gif" name="A8" width="95" height="24" border="0" id="A8"></a></td>		
   	</tr>
</table> 
<table width="760" border="0" cellspacing="0" cellpadding="0" align=center background="../picture/banner03.gif"  >
    <tr>	      
    	<td height="39" align="left" valign="middle" >
    	&nbsp;<img src="../picture/icon06.GIF"  width="22" height="17" border="0"  align="absmiddle" > 
    	<font color="#004080"><B>系統管理</b></font>&nbsp;&nbsp;&nbsp;   
    	│<a href="admin01.asp">1.代碼檔維護</a>│
    	│<a href="newuser.asp">2.新增使用者</a>│
    	│<a href="user.edit.asp">3.使用者維護</a>│  	
    	</td>
    	<td width=100 class=txt align=right valign="middle" ><font color=#FFFFFF>│</font><a href="vbscript:chgpage(9)"><font color=#FFFFFF><u>系統管理</u></font></a><font color=#FFFFFF>│</font></td>
    </tr>
</table>
<BR>
<table width=760 align=center><tr><td align=center>	
		<table width=480 >
		<tr>
			<td>類別：
				<select name=A1 class=inputbox onchange=a1chg()>
					<option value="" <%if a1="" then %>selected<%end if%>>全部顯示</option>
					<option value="whsno" <%if a1="whsno" then %>selected<%end if%>>廠別</option>
					<option value="unitno" <%if a1="unitno" then %>selected<%end if%>>處(部)</option>
					<option value="groupid" <%if a1="groupid" then %>selected<%end if%>>單位(組)</option>
					<option value="jobid" <%if a1="jobid" then %>selected<%end if%>>職等</option>
					<option value="country" <%if a1="country" then %>selected<%end if%>>國籍</option>
					<option value="ACCG" <%if a1="country" then %>selected<%end if%>>受文單位</option>
				</select>
			</td>
		</td>
		</table>
		<TABLE WIDTH=480   > 
	 	<TR BGCOLOR="Gainsboro">
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
	 		<TD><INPUT NAME=FuncTYPE VALUE="<%= tmpRec(CurrentPage, CurrentRow, 1)%>" SIZE=10 CLASS=INPT onchange="datachg(<%=currentrow-1%>)" ></TD>
	 		<TD><INPUT NAME=SYSTYPE VALUE="<%= tmpRec(CurrentPage, CurrentRow, 2)%>" SIZE=10 CLASS=INPT onchange="datachg(<%=currentrow-1%>)" ></TD>
	 		<TD><INPUT NAME=SYSVALUE VALUE="<%= tmpRec(CurrentPage, CurrentRow, 3)%>" SIZE=40 CLASS=INPT onchange="datachg(<%=currentrow-1%>)" s> </TD>
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
<script language=VBScript>
function del(index)	
     codestr01 = (trim(<%=SELF%>.autoid(index).value))
     codestr02 = Ucase(trim(<%=SELF%>.FuncTYPE(index).value))
     codestr03 = Ucase(trim(<%=SELF%>.SYSTYPE(index).value))
     codestr04 = Ucase(trim(<%=SELF%>.SYSVALUE(index).value))
	 if <%=SELF%>.func(index).checked = true then 
		<%=SELF%>.op(index).value = "del"		
		open "<%=SELF%>.Back.asp?codestr01=" & codestr01 & _            
		     "&CurrentPage=" & <%=CurrentPage%> & _
		     "&index=" & index & "&func=del", "Back"  
	'	 parent.best.cols="70%, 30%"    
		     
	 else
	    <%=SELF%>.op(index).value = "upd"		
		open "<%=SELF%>.Back.asp?codestr01=" & codestr01 & _            
			 "&codestr02=" & codestr02 &_
			 "&codestr03=" & codestr03 &_
			 "&codestr04=" & codestr04 &_
		     "&CurrentPage=" & <%=CurrentPage%> & _
		     "&index=" & index & "&func=upd", "Back"	     	     
		'   parent.best.cols="70%, 30%" 
	 end if		
end function 

function datachg(index)	
    codestr01 = (trim(<%=SELF%>.autoid(index).value))
    codestr02 = Ucase(trim(<%=SELF%>.FuncTYPE(index).value))
    codestr03 = Ucase(trim(<%=SELF%>.SYSTYPE(index).value))
    codestr04 = Ucase(trim(<%=SELF%>.SYSVALUE(index).value))
	open "<%=SELF%>.Back.asp?codestr01=" & codestr01 & _            
			 "&codestr02=" & codestr02 &_
			 "&codestr03=" & codestr03 &_
			 "&codestr04=" & codestr04 &_
		     "&CurrentPage=" & <%=CurrentPage%> & _
		     "&index=" & index & "&func=upd", "Back"	     	     
		 '  parent.best.cols="70%, 30%" 
	 
end function 

'-----------------enter to next field
function enterto()
	if window.event.keyCode = 13 then window.event.keyCode =9 		
end function

FUNCTION SYSNAME_CHG()
	<%=SELF%>.ACTION="<%=SELF%>.ASP?TOTALPAGE=0"
	<%=SELF%>.SUBMIT()
END FUNCTION  

FUNCTION openused()	
	OPEN "used.htm" , "_blank"  , "top=10, left=10, width=600, height=600"
END FUNCTION 

FUNCTION OPEND(N)
	OPEN "YFYFE01.SHOW.ASP?AUTOID="& N , "_BALNK" , "TOP=15, LEFT=15 , WIDTH=500, HEIGHT=550, SCROLLBARS=YES"
END  FUNCTION  

function chgpage(a)
	if a="1" then 
		if Trim(<%=self%>.uid.value)="" then  
			open "../login.asp", "_self"
		else
			open "../system.asp", "_self"
		end if
	elseif a="2" then  	
		open "../VYFYSELF/main.asp", "_self"		
	elseif a="3" then  	
		if Trim(<%=self%>.uid.value)="" then  
			alert "您必須先登入,才可使用電子公文!!"
		else
			open "../YFYBOARD/main.asp", "_self"
		end if
	elseif a="4" then  	
		if Trim(<%=self%>.uid.value)="" then  
			alert "您必須先登入,才可使用檔案中心!!"
		else
			open "../YFYEFILE/main.asp", "_self"
		end if	
	elseif a="5" then  	 
		open "../YFYEDOWN/main.asp", "_self"
	elseif a="6" then  	
		open "../YFYCONTS/main.asp", "_self"		
	elseif a="7" then  	
		open "../VYFYFIX/main.asp", "_self"			
	elseif a="8" then  	
		open "../YFYOTHER/main.asp", "_self"		
	elseif a="9" then  	
		if Trim(<%=self%>.uid.value)="" or Trim(<%=self%>.uid.value)<>"PELIN" then
			alert "系統管理者,才可使用此功能"
		else				
			open "../VYFYUSER/main.index.asp", "_parent"	 		
		end if 		
	end if 	
end function 

function a1chg()
	<%=self%>.totalpage.value="0" 
	<%=self%>.action="admin01.asp"
	<%=self%>.submit()
end function   

function go()
	<%=SELF%>.action = "<%=SELF%>.UpdateDB.asp"
   	<%=SELF%>.submit
end function 
</script> 

 

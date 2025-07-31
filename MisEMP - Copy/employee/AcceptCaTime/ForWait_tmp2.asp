<%@language=vbscript codepage=950%>
<%
Response.Buffer = true
Response.Expires = 0
func = request("func") 

'Response.Write func
'Response.End 


sdate = request("DAT1")
edate = request("DAT2")

tit = request("tit")
'Response.Write tit &"<P>"
self2=Request("self2")

'Response.Write self2
'Response.End  
'response.write now()

%>
<HTML>
<head>
<meta HTTP-EQUIV="Content-Type" content="text/html; charset=BIG5">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="refresh">
<link rel="stylesheet" href="../../Include/style.css" type="text/css">
<link rel="stylesheet" href="../../Include/style2.css" type="text/css">  

<script language="JavaScript">
<!--  
var timerID = null
var timerRunning = false

function stopclock(){
    if(timerRunning)
        clearTimeout(timerID)
    timerRunning = false
}

function startclock(){
    stopclock()
    showtime()
}

function showtime(){
    var now = new Date()
    var hours = now.getHours()
    var minutes = now.getMinutes()
    var seconds = now.getSeconds()
    var timeValue = "" + ((hours > 12) ? hours - 12 : hours)
    timeValue  += ((minutes < 10) ? ":0" : ":") + minutes
    timeValue  += ((seconds < 10) ? ":0" : ":") + seconds
    timeValue  += (hours >= 12) ? "PM" : "AM"
    document.clock.face.value = timeValue 
    timerID = setTimeout("showtime()",1000)
    timerRunning = true
}
//-->
</script>  

</head>
<body  leftmargin=5 rightmargin=0 topmargin=5 onload="startclock()"  > 
<table width=580 border=0><tr>
<td align=center height=60><h3><H3><%=tit%></H3></h3></font></td>
</tr></table>
<br>
<table width=630 border=0 class=txt>
<tr>
	<td align=right width=300>資料處理中，請稍後&nbsp;</td>
	<td align=left><img src=WDH01SLA.gif  align=absmiddle><img src=WDH01SLA.gif  align=absmiddle></td>	
</tr>
<tr>
	<td align=center colspan=2>請勿按下重新整理!!</td>	
</tr> 
<tr>
	<td align=center colspan=2> </td>	
</tr> 
</table>  
<TABLE WIDTH=630>
	<TR>
	<TD ALIGN=CENTER>
	<form name="clock" onsubmit="0">
	<input type="text" name="face" size="15" style="TEXT-ALIGN:CENTER;background-color: #FFFFFF; border-width: 0; font-family: Verdana;">
	</form> 
	</TD>
	</TR> 
</TABLE> 

<form name=form1>
<input type=hidden name=sdate value=<%=sdate%>>
<input type=hidden name=edate value=<%=edate%>>
</form>
</BODY>
</HTML>
<script language=vbscript>
sdate=form1.sdate.value
edate=form1.edate.value
 '  form1.action = "YCBBE0101.UpdateDB.asp" 
 open "acceptedcatime2.asp?DAT1=" & sdate & "&DAT2=" & edate  , "Back"
 'parent.best.cols="70%,30%"
 '  form1.submit



</script>

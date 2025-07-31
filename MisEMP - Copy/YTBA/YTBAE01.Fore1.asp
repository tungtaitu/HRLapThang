<%@ Language=VBScript CODEPAGE=65001%>
<!---------  #include file="../GetSQLServerConnection.fun"  -------->
<!--#include file="../include/checkpower.asp"-->
<!--#include file="../include/sideinfo.inc"-->

<%  
SELF = "YTBAE01"
Response.Buffer =true  
newpage="YES"    
SQL = "SELECT tblid,description from SCode_big where tblid <> '' order by tblid  " 	
Set conn = GetSQLServerConnection()
Set RS = Server.CreateObject("ADODB.Recordset")
RS.Open SQL, conn, 3, 3
IF RS.EOF THEN
%>
	<SCRIPT LANGUAGE=vbscript>
      alert "無代碼大類別存在"
	</SCRIPT>
<% END IF  %>	
<html>
<head>   
<meta HTTP-EQUIV="Content-Type" content="text/html; charset=UTF-8">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">	
<link rel="stylesheet" href="../../ydb/Include/style.css" type="text/css">
<link rel="stylesheet" href="../../ydb/Include/style2.css" type="text/css">  
</head>
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
<body background="bg_blue.gif"  topmargin=50>
<FORM method="POST" name="<%=self%>" action="<%=self%>.ForGnd.asp" >
<br><br>
<table width=600><tr><td align=center>
<TABLE CELLPADDING=5 CELLSPACING=5 BORDER=0 >
<TR>
<INPUT TYPE="HIDDEN"  NAME="newpage" VALUE="<%=newpage%>">
<TD align=right >代碼大類別(<U>T</U>)</TD>
<TD>  
    <select name="DB_TBLID"  class=txt >
    <%DO WHILE NOT RS.EOF  %>
	<option value=<%=RS(0)%> ><%=RS(0)%>-<% =RS(1) %>
	  <%  RS.MoveNext 
			  
	    LOOP 
	    RS.CLOSE%>
	</select>
</TD>  
</tr>
</TABLE>
<br>

<table width="100%">
<tr>
 <td align="CENTER">
 <input TYPE="submit" name="send" VALUE="輸     入" class="button"  >
 <input type="reset" name="send" value="取     消" class="button"   onclick="NEXTONE()">
 </td>
</tr>
</table>

</td></tr></table>
</FORM>

</BODY>
</HTML>


</script>
<script language="vbscript">
<!-- 
	function NEXTONE()
    	open "<%=self%>.asp", "content"
    end function
//-->
</script>

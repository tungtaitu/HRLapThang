<%@Language=VBScript Codepage=65001%>
<META HTTP-EQUIV="Content-Type" content="text/html; charset=utf-8">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="refresh">
<!---------  #include file="../../GetSQLServerConnection.fun"  -------->
<%
self="YTBDE01"
Set conn = GetSQLServerConnection()
Set rs = Server.CreateObject("ADODB.RECORDSET")

func = request("func")
code = request("code1")
index = request("index")
cols01 = request("cols01")
cols02 = request("cols02")
cols03 = request("cols03")

Select Case func
	   Case "del"
	        Sql = "update ytbdxmdp set status='D' , mdtm=getdate(), muser='"& session("userid") &"' where cmid='"& trim(Code) &"'   "
	        response.write sql
	        conn.execute(Sql)
			%>
			    <Script Language="vbscript">
			       parent.Fore.<%=self%>.totalpage.value=""
			       parent.Fore.<%=self%>.action = "<%=self%>.Fore.asp"
			       parent.Fore.<%=self%>.submit()
			    </Script>
			<%
			    Response.End
	   Case "chkempid"
			sql="select * from [yfynet].dbo.empfile  where ISNULL(OUTDAT,'')='' AND empid = '"& code &"'"
			response.write sql
	        rs.Open Sql,Conn,1,3
	        if not rs.eof  then
			%>
			    <Script Language="vbscript">
			       parent.fore.<%=self%>.empid.value = "<%=UCase(code)%>"
			       parent.fore.<%=self%>.ndpeople.value = "<%=rs("empnam_cn")%>"
			    </Script>
			<%
			    Response.End
			else
			%>
			    <Script Language="vbscript">
			       alert "No Data Complete!!"
			       parent.fore.<%=self%>.empid.value = ""
			       parent.fore.<%=self%>.ndpeople.value = ""
			       parent.fore.<%=self%>.empid.focus()
			    </Script>
			<%
			    Response.End
			end if
			set rs=nothing
End Select
%>
 
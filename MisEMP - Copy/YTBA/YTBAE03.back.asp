<%@Language=VBScript Codepage=65001%>
<META HTTP-EQUIV="Content-Type" content="text/html; charset=utf-8">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="refresh">
<!---------  #include file="../../GetSQLServerConnection.fun"  -------->
<%
self="YTBAE03"
Set conn = GetSQLServerConnection()
Set rs = Server.CreateObject("ADODB.RECORDSET")

func = request("func")
code = request("code1")
index = request("index")
cols01 = request("cols01")
cols02 = request("cols02")
cols03 = request("cols03")

response.write index &"<BR>"
response.write cols01 &"<BR>"
response.write cols02 &"<BR>"
response.write cols03 &"<BR>"

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
			       parent.Fore.<%=self%>.<%=cols01%>(<%=index%>).value = "<%=UCase(code)%>"
			       parent.Fore.<%=self%>.<%=cols02%>(<%=index%>).value = "<%=rs("empnam_cn")%>"
			       parent.Fore.<%=self%>.<%=cols03%>(<%=index%>).value = "<%=rs("email")%>"
						 parent.Fore.<%=self%>.<%=cols03%>(<%=index%>).focus()
			    </Script>
			<%
			    Response.End
			else
			%>
			    <Script Language="vbscript">
			       alert "No Data Complete!!"
			       parent.Fore.<%=self%>.<%=cols01%>(<%=index%>).value = ""
			       parent.Fore.<%=self%>.<%=cols02%>(<%=index%>).value = ""
			       parent.Fore.<%=self%>.<%=cols03%>(<%=index%>).value = ""
			       parent.Fore.<%=self%>.<%=cols01%>(<%=index%>).focus()
			    </Script>
			<%
			    Response.End
			end if
			set rs=nothing
End Select
%>
 
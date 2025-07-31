<%@ Language=VBScript codepage=65001%>
<!--#include file="../../ADOINC.inc"-->
<!--#include file="../../GetSQLServerConnection.fun"  -->
<head>
<meta HTTP-EQUIV="Content-Type" content="text/html; charset=utf-8">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">	
</head>
<%
self="empbe04"
func=request("func")
code1=request("code1")
index=request("index")
 

Set conn = GetSQLServerConnection()	 

select case func
    case "chkempid"
        sql="select convert(char(10), indat, 111) indate, * from  view_empfile where  empid='"& code1 &"' and isnull(status,'')<>'D' "
        RESPONSE.WRITE SQL     
        Set rs = Server.CreateObject("ADODB.Recordset")
        rs.Open Sql, conn, 3, 3 
        if rs.eof then 
%>          <script language=vbscript>
                alert "工號輸入錯誤!!"
                parent.Fore.<%=self%>.empid(<%=index%>).value=""
                parent.Fore.<%=self%>.empname(<%=index%>).value=""  
				parent.Fore.<%=self%>.country(<%=index%>).value=""  				
				parent.Fore.<%=self%>.indate(<%=index%>).value=""  				
                parent.Fore.<%=self%>.empid(<%=index%>).focus()
            </script>
<%          response.end
        else        	
%>          <script language=vbscript>
                parent.Fore.<%=self%>.empid(<%=index%>).value="<%=rs("empid")%>"
                parent.Fore.<%=self%>.empname(<%=index%>).value="<%=rs("empnam_cn")%>"
                parent.Fore.<%=self%>.country(<%=index%>).value="<%=rs("country")%>"                
               parent.Fore.<%=self%>.indate(<%=index%>).value="<%=rs("indate")%>"
                parent.Fore.<%=self%>.dat1(<%=index%>).focus()
            </script>    
<%      end if
end select   
%>                

<%@Language=VBScript CODEPAGE=65001%>
<meta HTTP-EQUIV="Content-Type" content="text/html; charset=UTF-8">
<!-------- #include file = "../../GetSQLServerConnection.fun" --------->

<%

whsno = request("whsno")
wbloai = request("wbloai") 
wbid = request("empid") 
if trim(request("indat"))="" then 
	indat="null" 
else
	indat = "'"&request("indat")&"'"
end if 
nam_cn = request("nam_cn") 
nam_vn = request("nam_vn") 
job = Ucase(trim(request("job")))
country = request("country") 
yy = request("byy") 
mm = request("bmm") 
dd = request("bdd") 
age = request("ages") 
sex = request("sexstr") 
personID = request("personID") 
cardno = request("cardno") 

if wbloai="01" then 
	fac = request("xhid") 
else
	fac = request("fac") 
end if 	
lorry=request("lorry")
soxe=request("soxe") 

phone = request("phone") 
mobile = request("mobile") 
addr=request("homeaddr")
wbmemo = request("memo") 

if trim(request("outdat"))="" then 
	outdat="null" 
else
	outdat = "'"&request("outdat")&"'"
end if  
flag = request("flag")  

outmemo = request("outmemo") 
filename = request("filename")
wbphotoid = request("wbphotoid")

Set conn = GetSQLServerConnection()



'on error resume next 
if session("netuser")="" then 
	response.write "使用者帳號為空,請重新登入!!<BR>"
	response.write "UserID is Null Please Login again !!!<BR>"
	response.write "Vao mang trong rong , hoac doi lau , hay nhan nut nhap mang tu dau !!! "
	response.end 
end if 
 	
 	
conn.BeginTrans 

response.write gTotalPage &"<BR>" 
response.write PageRec &"<BR>"
'response.end 

x = 0
y = ""
sql="select * from wbempfile where wbid='"& wbid &"' "
Set rs = Server.CreateObject("ADODB.Recordset") 
rs.open sql, conn, 1, 3 
if not rs.eof then 	
	sql="update wbempfile set wbname_cn=N'"&nam_cn&"',wbname_vn=N'"&nam_vn&"',indat="&indat&", "&_
		"job='"&job&"',country='"&country&"',yy='"&yy&"',mm='"&mm&"',dd='"&dd&"',age='"&age&"', "&_
		"sex='"&sex&"',personID='"&personID&"',cardno='"&cardno&"',fac='"&fac&"',lorry='"&lorry&"', "&_
		"soxe='"&soxe&"',phone='"&phone&"',mobile='"&mobile&"',wbmemo=N'"&wbmemo&"',outdat="&outdat&", "&_
		"outmemo=N'"&outmemo&"',flag='"&flag&"',addr=N'"&addr&"',mdtm=getdate(), muser='"&session("netuser")&"' "&_
		"filename='"& filename &"' "&_
		"where wbid='"&wbid&"'"
else
	sql="insert into wbempfile (loai,wbwhsno,wbid,wbname_cn,wbname_vn,indat,yy,mm,dd,age,sex,personid,phone,mobile, "&_
		"fac,lorry,soxe,job,addr,wbmemo,mdtm,muser, filename ) values ( "&_
		"'"&wbloai&"','"&whsno&"','"&wbid&"',N'"&nam_cn&"',N'"&nam_vn&"',"&indat&",'"&yy&"','"&mm&"','"&dd&"', "&_
		"'"&age&"','"&sex&"','"&personid&"','"&phine&"','"&mobile&"',N'"&fac&"','"&lorry&"','"&soxe&"', "&_
		"'"&job&"',N'"&addr&"',N'"&wbmemo&"',getdate(),'"&session("netuser")&"','"& filename&"' )  "
end if 
response.write sql 
conn.execute(sql)  

Nfname = Server.MapPath("wbphotos")&"\"& filename 
'Nfname = Server.MapPath("pic2")&"\"& file.FileName '測試
		
'response.write Nfname &"<BR>" 		
'response.end 
		
Set rsx= Server.CreateObject("ADODB.Recordset") 
sql="select * from wbempfile where wbid='"& wbid &"' and loai='"&wbloai&"' "		
response.write sql &"<BR>"
		
const adCmdText=1
const adOpenDynamic=2
const adLockOptimistic=3
const adOpenKeyset=1
		
Set mstream = Server.CreateObject("ADODB.Stream")
mstream.Type = 1
mstream.Open
		
mstream.LoadFromFile Nfname  
'mstream.LoadFromFile file.name  '測試
'response.end 	
		
rsx.Open SQL,conn,adOpenKeyset,adLockOptimistic,adCmdText
if not rsx.eof then
	rsx.Fields("photos").Value = mstream.Read
	rsx.Update
else
	rsx.addnew  		        
	rsx.Fields("photos").Value = mstream.Read        
	rsx.Update        
end if 
rsx.close   


sqlx="update wbphotos set wbwhsno='"& whsno &"' , wbempid='"& wbid &"' where aid='"& wbphotoid &"' "
conn.execute(Sqlx)
response.write sqlx 


'RESPONSE.WRITE Y 
'response.end  

 if conn.Errors.Count = 0 or err.number=0  then 
	conn.CommitTrans
	conn.close
	Set conn = Nothing
	Session("Title") = ""
	Session("Name") = "IMessage"
	Session("NO") = "Data Complete Success (OK) Count:" & x 
	Session("MessageCode") = "Success"
	Session("KeyValue") =  Y
	Session("SubmitValue") = replace(session("pgname"),"<BR>",chr(13))
	'Session("Action") = "YSBAE1403.Fore.asp"
	'Response.Redirect "YEBE0104.Fore.asp?whsno="&whsno&"&wbloai="&wbloai&"&soxe="&soxe
	Set conn = Nothing %>
	<script language=vbs>
		alert "資料處理成功(OK)!!"
		parent.close()
	</script>
<% else
	conn.RollbackTrans     
    Set cmd = Nothing
		conn.close
	Set conn = Nothing 
%>	<script language=vbs>
		alert "資料處理失敗(Fial)!!"
		open "YEBE0104.fore.asp"
	</script>
	
<% end if  %>

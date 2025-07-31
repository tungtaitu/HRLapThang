<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<!-- #include file = "../GetSQLServerConnection.fun" -->
<%
SELF = "YECE0700" 
Set conn = GetSQLServerConnection() 
setstr ="" 
for i = 1 to 7 
	if ucase(trim(request("fcode")(i)))<>"" then 
		setstr = setstr & 		ucase(trim(request("fcode")(i))) &"," 
	end if 	
next 
'newstr = replace(left(setstr,len(setstr)-1),",","+")
newstr = replace(setstr ,",","+")
w = request("whsno")
ym= request("ym")   

if request("flag")="C"  then 
	ym = request("CopyYm")
end if	

sqlx="select * from  empbh_set where w='"& w &"' and ym='"& ym &"' "  
set rds=conn.execute(Sqlx)
if rds.eof then 
	sql="insert into empbh_set (w, ym ,setstr ) values ('"&w&"','"& ym &"', '"& newstr &"' ) "
else
	sql="update empbh_set set setstr='"& newstr &"' where w='"& w &"' and ym='"&ym &"' "   
end if 
set rds=nothing 
conn.execute(sql) 
response.write sqlx&"<br>"
response.write sql &"<br>"  

emp_bhxh = request("emp_bhxh")
emp_bhyt = request("emp_bhyt")
emp_bhtn = request("emp_bhtn")
cty_bhxh = request("cty_bhxh")
Cty_bhyt = request("Cty_bhyt")
cty_bhtn = request("cty_bhtn")
emp_gtamt = request("emp_gtamt") : if emp_gtamt="" then emp_gtamt= 0 
gt_per = request("gt_per") : if gt_per="" then gt_per= 0 
ct1=request("ct1")
if  emp_bhxh<>"" then a=cdbl(emp_bhxh) else a=0
if  emp_bhyt<>"" then b=cdbl(emp_bhyt) else b=0
if  emp_bhtn<>"" then c=cdbl(emp_bhtn) else c=0
if  cty_bhxh<>"" then d=cdbl(cty_bhxh) else d=0
if  Cty_bhyt<>"" then e=cdbl(Cty_bhyt) else e=0
if  cty_bhtn<>"" then f=cdbl(cty_bhtn) else f=0  


sqls="select * from  empbh_per where  yymm='"& ym &"' and country='VN'"   
set rs2=conn.execute(sqls)
if rs2.eof then 
	sql="insert into  empbh_per (yymm,emp_bhxh, emp_bhyt, emp_bhtn, cty_bhxh, cty_bhyt, cty_bhtn, muser,country,emp_gtant ,gt_per ) values ( "&_
			"'"&ym&"','"&a&"','"&b&"','"&c&"','"&d&"','"&e&"','"&f&"','"&session("netuser")&"','"&ct1&"' ,'"&emp_gtamt&"','"&gt_per&"')"
	conn.execute(sql)		
else
	sql="update  empbh_per set emp_bhxh='"&a&"', emp_bhyt='"&b&"', emp_bhtn='"&c&"', "&_
			" cty_bhxh='"&d&"', cty_bhyt='"&e&"', cty_bhtn='"&f&"',country='"&ct1&"', "&_
			" mdtm=getdate(), muser='"&session("netuser") &"' , emp_gtant='"&emp_gtamt&"' , "&_
			" gt_per='"&gt_per&"' where yymm='"& ym &"' and country='VN' "
	conn.execute(Sql)		
end if 

'-*------ 外籍保險費率
ct2=request("ct2")
hwemp_bhxh = request("hwemp_bhxh")
hwemp_bhyt = request("hwemp_bhyt")
hwemp_bhtn = request("hwemp_bhtn")
hwcty_bhxh = request("hwcty_bhxh")
hwCty_bhyt = request("hwCty_bhyt")
hwcty_bhtn = request("hwcty_bhtn")
if  hwemp_bhxh<>"" then a=cdbl(hwemp_bhxh) else a=0
if  hwemp_bhyt<>"" then b=cdbl(hwemp_bhyt) else b=0
if  hwemp_bhtn<>"" then c=cdbl(hwemp_bhtn) else c=0
if  hwcty_bhxh<>"" then d=cdbl(hwcty_bhxh) else d=0
if  hwCty_bhyt<>"" then e=cdbl(hwCty_bhyt) else e=0
if  hwcty_bhtn<>"" then f=cdbl(hwcty_bhtn) else f=0 

sqls="select * from  empbh_per where  yymm='"& ym &"' and country='HW'"   
set rs2=conn.execute(sqls)
if rs2.eof then 
	sql="insert into  empbh_per (yymm,emp_bhxh, emp_bhyt, emp_bhtn, cty_bhxh, cty_bhyt, cty_bhtn, muser,country ) values ( "&_
			"'"&ym&"','"&a&"','"&b&"','"&c&"','"&d&"','"&e&"','"&f&"','"&session("netuser")&"','"&ct2&"' )"
	conn.execute(sql)		
else
	sql="update  empbh_per set emp_bhxh='"&a&"', emp_bhyt='"&b&"', emp_bhtn='"&c&"', "&_
			" cty_bhxh='"&d&"', cty_bhyt='"&e&"', cty_bhtn='"&f&"', country='"&ct2&"' , "&_
			" mdtm=getdate(), muser='"&session("netuser") &"' "&_
			" where yymm='"& ym &"' and country='HW' "
	conn.execute(Sql)		
end if 
'response.end 
%>
<script language=vbs>
	open "<%=self%>.fore.asp?whsno="&"<%=w%>" &"&ym="&"<%=ym%>" , "_self"
</script>
 
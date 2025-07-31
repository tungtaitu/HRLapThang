<script language="javascript" >
	
	function dblconfirm(str)
	{
		
		with (window) 
		{			
			var cnf			
			cnf=confirm("確定要〔"+ str.trim() +"〕嗎?");			
			if(cnf==true){
				dblconfirm=true;
			}
			else {
				dblconfirm=false;
			}			
		}
		return dblconfirm;
	}
</script>


<script language="vbs" >
'-----------------檢查不合法字元function
function chkchar(str)
	with window.<%=formname%>
		dim errstr,i,errarray
		errstr="%,:,"",'"
		errarray=split (errstr,",")
		for i=0 to ubound(errarray,1)
			if instr(1,str,errarray(i))<>0 then
				alert "欄位值含不合法字元......（" & errarray(i) & "）"
				chkchar=false
				exit function
			end if
		next
	end with
	chkchar=true
end function

'-----------------各類欄位型態判斷function
function chktype(val,index)
	select case index
		case 1 '-----字串不得為空白
			if val="" then
				alert ("該欄位不得為空白!!!")
				chktype=false
				exit function
			end if
		case 2 '-----欄位只能輸入數字
			if isnumeric(val)=false then
				alert ("該欄位請輸入數字!!!")
				chktype=false
				exit function
			end if
		case 3 '-----日期欄位檢查
			if isdate(val)=false then
				alert ("該欄位請輸入正確日期格式!!!ex:2003/10/10")
				chktype=false
				exit function
			end if
		case 4 '-----欄位只能輸入正數字
			if isnumeric(val)=false then
				alert ("該欄位請輸入數字!!!")
				chktype=false
				exit function
			elseif val<0 then
				alert ("該欄位請輸入正數字!!!")
				chktype=false
				exit function
			end if
	end select
	chktype=true
end function

'------------------改變物件背景色function
function diffcolor(name,index)
	with window.<%=formname%>
		set obj=eval("."& name)
		select case index
			case 1 '----淡灰色 & readonly
				obj.style.color="#000000"
				obj.style.backgroundcolor="#c0c0c0"
				obj.readonly=true
			case 2 '----亮黃色 & noreadonly
				obj.style.backgroundcolor="#ffffc0"
				obj.readonly=false
			case 3 '----藍字白底 & readonly
				obj.style.color="#0000cc"
				obj.style.backgroundcolor="#ffffff"
				obj.readonly=true
			case 4 '----黑字白底 & noreadonly
				obj.style.color="#000000"
				obj.readonly=false
			case 5 '----灰底黑字 & readonly
				obj.style.backgroundcolor="#eeeeee"
				obj.readonly=true	
			case 6 '----orange & noreadonly
				obj.style.backgroundcolor="#ccccff"
				obj.readonly=false
			case else
				exit function
		end select	
	end with
end function


'-----------------計算function
function getcompute(num1,num2,index)
	select case cstr(index)
		case "1" '---------無條件進入
			if num2<>0 then	
				if num1 mod num2 > 0 then 
					getcompute=num1 \ num2 + 1
				else
					getcompute=num1 \ num2 
				end if
			else
				getcompute=0
			end if
	end select
end function

'-----------------dblcheck function
function dblconfirm(str)
	with window
		dim cnf
		cnf=.confirm ("確定要〔"& trim(str) &"〕嗎?")
		if cnf=true then
			dblconfirm=true
		else
			dblconfirm=false
		end if
	end with
end function

'-----------------enter to next field
function enterto()
	with window
		if .event.keyCode=13 then .event.keyCode =9
	end with
end function

'-----------------將所有欄位值轉大寫
function UcaseChar()
	dim Elm
	for each Elm in <%=formname%>
		select case Elm.name
			case "ArrSql_tmp","strsql"
			case else
				Elm.value=replace(Elm.value,"'","′")
				Elm.value=replace(Elm.value,"""","”")
				Elm.value=ucase(Elm.value)
		end select
	next
end function 

'-----------------公司統編CHECK FUNCTION
'-----------------sBinID(需CHECK的統編)
'-----------------nErrCode(錯誤回傳碼)
'-----------------sErrDesp(錯誤回傳說明)

Function CheckCompanyId(sBinID, nErrCode, sErrDesp)
   Dim intA,intB,intSum,i,intMod,sChkID
 
   CheckCompanyId = True   
   sChkID = Trim(CStr(sBinID))
   nErrCode = 0
   sErrDesp = ""
   ReDim intA(8)
   ReDim intB(8)
   
   If Len(sChkID) <> 8 Then
      nErrCode = 1
      sErrDesp = "公司統一編號長度不正確" & Space(5)
      CheckCompanyId = False
      Exit Function
   End If
  
   If IsNumeric(sChkID) = False Then 
         nErrCode = 1
         sErrDesp = "公司統一編號中有非數字" & Space(5)
         CheckCompanyId = False
         Exit Function
   End If
   
   intA(1) = Clng(Mid(sChkID, 1, 1)) * 1   '第一位數*1
   intA(2) = Clng(Mid(sChkID, 2, 1)) * 2   '第二位數*2
   intA(3) = Clng(Mid(sChkID, 3, 1)) * 1   '第三位數*1
   intA(4) = Clng(Mid(sChkID, 4, 1)) * 2   '第四位數*2
   intA(5) = Clng(Mid(sChkID, 5, 1)) * 1   '第五位數*1
   intA(6) = Clng(Mid(sChkID, 6, 1)) * 2   '第六位數*2
   intA(7) = Clng(Mid(sChkID, 7, 1)) * 4   '第七位數*4
   intA(8) = Clng(Mid(sChkID, 8, 1)) * 1   '第八位數*1
   
   '二四六七位數可能會大於10,故需取其整數與餘數
   intB(1) = Clng(Int(intA(2) / 10))
   intB(2) = Clng(intA(2) Mod 10)
   intB(3) = Clng(Int(intA(4) / 10))
   intB(4) = Clng(intA(4) Mod 10)
   intB(5) = Clng(Int(intA(6) / 10))
   intB(6) = Clng(intA(6) Mod 10)
   intB(7) = Clng(Int(intA(7) / 10))
   intB(8) = Clng(intA(7) Mod 10)
   
   intSum = intA(1) + intA(3) + intA(5) + intA(8)
   For i = 1 To 8
      intSum = intSum + intB(i)
   Next
   intMod = intSum Mod 10
   If Clng(Mid(sChkID, 7, 1)) = 7 Then  ' 判斷第7位數是否為7
      If intMod = 0 Then ' 判斷餘數是否為0
          CheckCompanyId = True
          Exit Function
      Else
         intSum = intSum + 1
         intMod = intSum Mod 10
         If intMod = 0 Then
            CheckCompanyId = True
            Exit Function
         Else
            nErrCode = 2
            sErrDesp = "公司統一編號錯誤!!!" & Space(5)
            CheckCompanyId = False
            Exit Function
         End If
      End If
   Else
      If intMod = 0 Then
         CheckCompanyId = True
         Exit Function
      Else
         nErrCode = 2
         sErrDesp = "公司統一編號錯誤!!!" & Space(5)
         CheckCompanyId = False
         Exit Function
      End If
   End If
End Function

'-------------------日期檢核修正function
'-------------------傳入值為check field name
'-------------------連帶function為ValidDate
FUNCTION date_change(name)	
	with window.<%=formname%> 
		dim obj
		set obj=eval("."& name)
		INcardat = Trim(obj.value)  				    
		IF INcardat<>"" THEN
			ANS=validDate(INcardat)
			IF ANS <> "" THEN		
				obj.value=ANS
			ELSE
				ALERT "EZ0067:輸入日期不合法 !!" 		
					obj.value=""		
					obj.focus()
				EXIT FUNCTION
			END IF
		ELSE
			EXIT FUNCTION
		END IF    
	end with
END FUNCTION

'_________________DATE CHECK___________________________________________________________________

function validDate(d)
	if len(d) = 8 and isnumeric(d) then
		d = left(d,4) & "/" & mid(d, 5, 2) & "/" & right(d,2)
		if isdate(d) then
			validDate = formatDate(d)
		else
			validDate = ""
		end if
	elseif len(d) >= 8 and isdate(d) then
			validDate = formatDate(d)
		else
			validDate = ""
		end if
end function

function formatDate(d)
		formatDate = Year(d) & "/" & _
		Right("0" & Month(d), 2) & "/" & _
		Right("0" & Day(d), 2)
end function
</script>
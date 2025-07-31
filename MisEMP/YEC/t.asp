<html> 
<body> 
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="table1"> 
  <TR>
  	<TD COLSPAN=4>2008 SALARY </TD>
  </TR>
  <tr> 
    <td width="25%">學號</td> 
    <td width="25%">姓名</td> 
    <td width="25%">科目</td> 
    <td width="25%">成績</td> 
  </tr> 
  <tr> 
    <td width="25%">0001</td> 
    <td width="25%">王小明</td> 
    <td width="25%">國語</td> 
    <td width="25%">90</td> 
  </tr> 
  <tr> 
    <td width="25%">0002</td> 
    <td width="25%">李大名</td> 
    <td width="25%">國語</td> 
    <td width="25%">80</td> 
  </tr> 
  <tr> 
    <td width="25%">0003</td> 
    <td width="25%">趙中明</td> 
    <td width="25%">國語</td> 
    <td width="25%">70</td> 
  </tr> 
</table> 
<form name="f1"> 
  <input type="button" value="匯出至excel" name="B1" onClick="saveToExcel('table1');"> 
</form> 
</body> 
</html> 
<script language="JavaScript"> 
function saveToExcel(str) { 
   try { 
      var xls = new ActiveXObject("Excel.Application"); 
      xls.Visible = true; 
   } 
   catch(e) { 
      alert("開啟失敗，請確定你的電腦已經安裝excel，且瀏覽器必須允許ActiveX控件執行"); 
      return; 
   } 
   var objTable = document.getElementById(str); 
   var xlBook = xls.Workbooks.Add; 
   var xlsheet = xlBook.Worksheets(1); 
   for (var i=0;i<objTable.rows.length;i++) 
      for (var j=0;j<objTable.rows[i].cells.length;j++) 
         xlsheet.Cells(i+1,j+1).value = objTable.rows[i].cells[j].innerHTML; 
} 
</script>

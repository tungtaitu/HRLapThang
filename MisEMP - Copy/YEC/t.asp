<html> 
<body> 
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="table1"> 
  <TR>
  	<TD COLSPAN=4>2008 SALARY </TD>
  </TR>
  <tr> 
    <td width="25%">�Ǹ�</td> 
    <td width="25%">�m�W</td> 
    <td width="25%">���</td> 
    <td width="25%">���Z</td> 
  </tr> 
  <tr> 
    <td width="25%">0001</td> 
    <td width="25%">���p��</td> 
    <td width="25%">��y</td> 
    <td width="25%">90</td> 
  </tr> 
  <tr> 
    <td width="25%">0002</td> 
    <td width="25%">���j�W</td> 
    <td width="25%">��y</td> 
    <td width="25%">80</td> 
  </tr> 
  <tr> 
    <td width="25%">0003</td> 
    <td width="25%">������</td> 
    <td width="25%">��y</td> 
    <td width="25%">70</td> 
  </tr> 
</table> 
<form name="f1"> 
  <input type="button" value="�ץX��excel" name="B1" onClick="saveToExcel('table1');"> 
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
      alert("�}�ҥ��ѡA�нT�w�A���q���w�g�w��excel�A�B�s�����������\ActiveX�������"); 
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

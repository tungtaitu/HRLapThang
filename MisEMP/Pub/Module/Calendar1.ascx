<%@ Control Language="c#" Inherits="Pub_CommControl_Calendar1" CodeFile="Calendar1.ascx.cs" %>
<script language="javascript">
function ShowCanlendar(objThis,strFormat)
{
	var strLOC = document.location.toString().toLowerCase();
	var aStrName;
	var strDate = "";
	aStrName = strLOC.split("/");
	var tmpPath = "http://" + aStrName[2] + "/" + aStrName[3] + "/"+ aStrName[4] + "/";
	
	var sPath = tmpPath + "Pub/Html/Calendar.htm";
			      
	strFeatures = "dialogWidth=290px;dialogHeight=210px;center=yes;help=no;status=no;resizable=no";
	strDate = showModalDialog(sPath,strFormat,strFeatures);
	if (strDate != null)
	{
		objThis.previousSibling.previousSibling.value = strDate;			
		ReLoadLabsentHr();
	}
}
function ValidData1(objthis)
{
	var strValue = objthis.value;
	var yy = 0;
	var mm = 0
	var dd = 0
	try
	{
		if (isNaN(strValue))
		{
			ErrMsg(strValue);
			return;
		}
		else
		{
			if (strValue.length != 8)
			{
				ErrMsg(strValue);
				return;
			}
			else
			{
				yy = parseInt(strValue.substring(0,4),10);
				mm = parseInt(strValue.substring(4,6),10);
				dd = parseInt(strValue.substring(6,8),10);
				if (mm < 1 || mm > 12)
				{
					ErrMsg(strValue);
					return;
				}
				if (dd < 1 || dd > 31)
				{
					ErrMsg(strValue);
					return;
				}
			}
			
		}
	}
	catch(e)
	{
		alert("error");
	}
	
}

function ErrMsg(err)
{
	alert("您所輸入日期格式:" + err + " 不符，請改用如:20040810之日期格式輸入");
}
</script>
<%--<asp:TextBox id="UC_Calendar" onchange="ValidData1(this);" runat="server" Width="78px" MaxLength="10" SkinID="txt01"></asp:TextBox>
--%>
<asp:TextBox id="UC_Calendar"  runat="server" Width="78px" MaxLength="10" SkinID="txt01"></asp:TextBox>
<input type="button" id="btn1ByCalendar" value="..." onclick="ShowCanlendar(this,'yyyy/MM/dd');" class="btn01ss_0"  onmouseover="BtnMouseOver(this,'btn01ss_1');" onmouseout="BtnMouseOut(this,'btn01ss_0');" >

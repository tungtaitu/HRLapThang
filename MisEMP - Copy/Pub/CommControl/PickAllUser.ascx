<%@ Control Language="C#" AutoEventWireup="true" CodeFile="PickAllUser.ascx.cs" Inherits="Pub_CommControl_PickAllUser" %>
<!-- <LINK href="../../App_Themes/salary01/Button.css" type="text/css" rel="stylesheet"> -->
<script src="../../../Pub/Js/common.js" type="text/javascript" ></script>
<script type="text/javascript" language="javascript"> 

function Notifier(obj) {
	var strLOC = document.location.toString().toLowerCase();
	var aStrName;
	var strUser = "";
	var aStrUser = "";
	aStrName = strLOC.split("/");
	var sPath = "../../../Pub/CommControl/PikcAllEmps.aspx";
	//var tmpPath = "http://" + aStrName[2] + "/" + aStrName[3] + "/"+ aStrName[4] + "/";
	//var sPath = tmpPath + "Pub/CommControl/PikcAllEmps.aspx";
	

    strFeatures = "dialogWidth=750px;dialogHeight=500px;center=yes;help=no;status=no;resizable=no";
    
//	window.open(sPath,this,strFeatures);	//var UserID = window.open(sPath,'PickUser',strFeatures);
    window.open(sPath,'PickUser',strFeatures);
}
function SetValue(UserInfo)
{
    if (UserInfo == undefined)//直接關閉視窗
	{
	   obj.previousSibling.previousSibling.value  = "";
	}
	else
	{
	  var strSplit = UserInfo.split("$");
	  var obj=document.getElementById("btChoice");
	
	  obj.previousSibling.value  = strSplit[1];//ShowUserNm
	  obj.nextSibling.value= strSplit[0];//hideUserID
	  obj.nextSibling.nextSibling.value =  strSplit[1];//hideUserNm
	}	
}
</script>
<asp:TextBox ID="txt_UserInfo" runat="server"  ReadOnly="True" TextMode="MultiLine"
    Width="50%"></asp:TextBox><input id="btChoice" class="btn01ss_0" onclick=" Notifier(this);"
        type="button" value="..."  onmouseover="BtnMouseOver(this,'btn01ss_1');" onmouseout="BtnMouseOut(this,'btn01ss_0');" /><input id="txtReturnID" runat="server" name="txtReturnID"
            style="width: 56px" type="hidden" /><input id="hidReturnNm" runat="server" name="hidReturnNm"
                style="width: 56px" type="hidden" />

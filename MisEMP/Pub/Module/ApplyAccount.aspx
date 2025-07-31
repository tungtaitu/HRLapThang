<%@ Page language="c#" Inherits="Pub_Module_ApplyAccount"  CodeFile="ApplyAccount.aspx.cs" %>
<%@ Register TagPrefix="uc1" TagName="ValidNumber" Src="ValidNumber.ascx" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" >
<HTML>
	<HEAD runat="server">
		<title></title>
		<meta name="GENERATOR" Content="Microsoft Visual Studio 7.0">
		<meta name="CODE_LANGUAGE" Content="C#">
		<meta name="vs_defaultClientScript" content="JavaScript">
		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
		<LINK href="../Css/PccStyles.css" rel="stylesheet">
		<script language="javascript">
			function ApplicationChange()
			{
				if (ApplyAccount.ddlApplcation.value == "0")
					ApplyAccount.btnApply.disabled = true;
				else
					ApplyAccount.btnApply.disabled = false;
			}
			
			function ValidFact(objval,objargs)
			{
				if (objargs.Value == "0")
					objargs.IsValid = false;
					
				//ApplyAccount.SelectedDept.value = ApplyAccount.ddldept_id.value;
					
			}
			
			function ChangeFact(objThis)
			{
				var i;
				var a1;
				a1 = ApplyAccount.DeptData.value.split(";");
				
				ApplyAccount.ddldept_id.innerHTML = ""
				CreateSelOption("","0",ApplyAccount.ddldept_id);
				
				if (ApplyAccount.ddlfact_id.value != "0")
				{
					for (i = 0 ; i < a1.length ; i++)
					{
						if (ApplyAccount.ddlfact_id.value.split("-")[1] == a1[i].split(",")[0])
						{
							CreateSelOption(a1[i].split(",")[2],a1[i].split(",")[1],ApplyAccount.ddldept_id);
						}
					}
				}
				
			}
			
			function CreateSelOption(OptionData,OptionValue,SelectID)
			{
				var SelectOption = document.createElement('OPTION');
				SelectOption.value = OptionData;
				if(OptionData == "")
				{
					SelectOption.text = "--請選擇--";
					SelectOption.value = "0";
				}
				else
				{
					SelectOption.text = OptionData;
					SelectOption.value = OptionValue;
				}   
				SelectID.add(SelectOption);	
				
			}
			
			function ValidFunction(objval,objargs)
			{
				if (objargs.Value == "0")
					objargs.IsValid = false;		
			}
		</script>
	</HEAD>
	<body style="BACKGROUND-COLOR:#ffffff">
		<form  id="ApplyAccount" method="post" runat="server">
			<input id="DeptData" type="hidden" name="DeptData" runat="server"> <input id="SelectedDept" type="hidden" name="SelectedDept" runat="server">
			<table cellpadding="0" cellspacing="0" border="0" width="100%" height="100%" style="BACKGROUND-POSITION: 50% top; BACKGROUND-IMAGE: url(../../images/number_2.jpg); BACKGROUND-REPEAT: repeat-x">
				<tr>
					<td valign="top" style="PADDING-RIGHT: 0px; BACKGROUND-POSITION: left top; PADDING-LEFT: 300px; BACKGROUND-IMAGE: url(../../images/number_1.jpg); PADDING-BOTTOM: 0px; PADDING-TOP: 80px; BACKGROUND-REPEAT: no-repeat">
						<table cellpadding="0" cellspacing="0" border="0">
							<tr>
								<td colspan="2" class="font06" style="PADDING-BOTTOM:30px">歡迎來到寶成網路資訊系統，使用者帳號申請中心<br>
									為了您使用的權益，請您要使用正確資料喔！</td>
							</tr>
							<tr>
								<td colspan="2"><asp:label id="lblMsg" runat="server" ForeColor="Red"></asp:label></td>
							</tr>
							<tr style="LINE-HEIGHT:2">
								<td class="font05"><asp:label id="lblApplication" runat="server" Width="101px" Height="19px">應用程式：</asp:label></td>
								<td><asp:dropdownlist id="ddlApplcation" runat="server" Width="156px" CssClass="inputTxt02" onselectedindexchanged="ddlApplcation_SelectedIndexChanged"></asp:dropdownlist></td>
							</tr>
							<tr style="LINE-HEIGHT:2">
								<td class="font05"><asp:label id="lbluser_desc" runat="server" Width="101px">姓名：</asp:label></td>
								<td><asp:textbox id="txtuser_desc" runat="server" Width="156px" MaxLength="20" CssClass="inputTxt02"></asp:textbox><asp:requiredfieldvalidator id="rfvuser_desc" runat="server" ControlToValidate="txtuser_desc" ErrorMessage="*"></asp:requiredfieldvalidator></td>
							</tr>
							<tr style="LINE-HEIGHT:2">
								<td class="font05"><asp:label id="lbluser_nm" runat="server" Width="101px">電子郵件帳號：</asp:label></td>
								<td><asp:textbox id="txtuser_nm" runat="server" Width="288px" MaxLength="50" CssClass="inputTxt02"></asp:textbox><asp:requiredfieldvalidator id="rsvuser_nm" runat="server" ControlToValidate="txtuser_nm" ErrorMessage="*" Display="Dynamic"></asp:requiredfieldvalidator>
									<asp:RegularExpressionValidator id="RegularExpressionValidator1" runat="server" ErrorMessage="Email Format Error!" Display="Dynamic" ValidationExpression="\w+([-+.]\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*" ControlToValidate="txtuser_nm"></asp:RegularExpressionValidator></td>
							</tr>
							<tr style="LINE-HEIGHT:2">
								<td class="font05"><asp:label id="lblusr_pas" runat="server" Width="101px">密碼：</asp:label></td>
								<td><asp:textbox id="txtusr_pas" runat="server" Width="156px" MaxLength="20" TextMode="Password" CssClass="inputTxt02"></asp:textbox><asp:requiredfieldvalidator id="rsvusr_pas" runat="server" ControlToValidate="txtusr_pas" ErrorMessage="*"></asp:requiredfieldvalidator></td>
							</tr>
							<tr style="LINE-HEIGHT:2">
								<td class="font05"><asp:label id="lblReusr_pas" runat="server" Width="101px">密碼確認：</asp:label></td>
								<td><asp:textbox id="txtReusr_pas" runat="server" Width="156px" MaxLength="20" TextMode="Password" CssClass="inputTxt02"></asp:textbox><asp:requiredfieldvalidator id="rsvReusr_pas" runat="server" ControlToValidate="txtReusr_pas" ErrorMessage="*"></asp:requiredfieldvalidator><asp:comparevalidator id="CompareValidator1" runat="server" ControlToValidate="txtReusr_pas" ErrorMessage="Compare Error" ControlToCompare="txtusr_pas"></asp:comparevalidator></td>
							</tr>
							<tr style="LINE-HEIGHT:2">
								<td class="font05"><asp:label id="lblfact_id" runat="server" Width="101px" Height="19px">廠別：</asp:label></td>
								<td>
									<asp:dropdownlist id="ddlfact_id" runat="server" Width="156px" CssClass="inputTxt02"></asp:dropdownlist>
									<asp:CustomValidator id="CustomValidator1" runat="server" ErrorMessage="*" ControlToValidate="ddlfact_id" ClientValidationFunction="ValidFunction"></asp:CustomValidator>
								</td>
							</tr>
							<tr style="LINE-HEIGHT:2">
								<td class="font05"><asp:label id="lblemp_no" runat="server" Width="101px">工號：</asp:label></td>
								<td><asp:textbox id="txtemp_no" runat="server" Width="156" MaxLength="8" CssClass="inputTxt02"></asp:textbox></td>
							</tr>
							<tr style="LINE-HEIGHT:2">
								<td class="font05"><asp:label id="lblext" runat="server" Width="101px">分機：</asp:label></td>
								<td><asp:textbox id="txtext" runat="server" Width="156" MaxLength="10" CssClass="inputTxt02"></asp:textbox></td>
							</tr>
							<tr style="LINE-HEIGHT:2">
								<td class="font05"><asp:label id="lblCheck" runat="server" Width="101">請輸入下列數字：</asp:label></td>
								<td>
									<uc1:ValidNumber id="ValidNumber1" runat="server" CssClass="inputTxt02"></uc1:ValidNumber></td>
							</tr>
							<tr>
								<td colspan="2" align="middle" style="PADDING-RIGHT:0px; PADDING-LEFT:0px; PADDING-BOTTOM:0px; PADDING-TOP:20px">
									<!--<input type="button" class="inputBtn02" value="提出申請">&nbsp;
								<input type="button" class="inputBtn02" value="重新登錄">-->
									<asp:button id="btnApply" runat="server" Text="提出申請" CssClass="inputBtn02" onclick="btnApply_Click"></asp:button>&nbsp;
									<asp:button id="btnReLogin" runat="server" Text="重新登入" CausesValidation="False" CssClass="inputBtn02" onclick="btnReLogin_Click"></asp:button>&nbsp;
								</td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
		</form>
	</body>
</HTML>

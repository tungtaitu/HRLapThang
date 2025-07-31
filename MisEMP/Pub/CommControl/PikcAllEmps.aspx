<%@ Page Language="C#" AutoEventWireup="true" CodeFile="PikcAllEmps.aspx.cs" Inherits="SalaryReport_AuthGroupApplication_PikcAllEmps"  Theme="pwfBody"  Culture="auto"  UICulture="auto"%>
<%@ Register Src="NoAuthFEC.ascx" TagName="NoAuthFEC" TagPrefix="uc4" %>
<%@ Register Src="../../Pub/BaseControls/ddlComp.ascx" TagName="ddlComp" TagPrefix="uc1" %>
<%@ Register Src="../../Pub/BaseControls/ddlFact.ascx" TagName="ddlFact" TagPrefix="uc2" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title></title>
   <link href="../../App_Themes/pwfBody/Default.css" type="text/css" rel="stylesheet"/>  
    <script src="../../Pub/Js/common.js" type="text/javascript" ></script>
      
   <script language="javascript" type="text/javascript">
   function GetChoiceUserIno(UserInfo)
    {
        if(window.opener)
        {
            window.opener.SetValue(UserInfo);  
        }
    
	    //window.returnValue =  UserInfo ;//取得返回的userID
	    //alert(window.returnValue)
        window.close();
    }
   
 
   </script>
   <base target ="_self" /> 
</head>
<body  class="area_main" style="text-align:left;" >
    <form id="PikcAllEmps" runat="server">
    <div>
        <asp:ScriptManager ID="ScriptManager1" AsyncPostBackTimeOut="900" runat="server">
        </asp:ScriptManager>
        <table width="100%">
            <tr>
                <td >
                    <table width="100%">
                        <tr>
                            <td class="area_editBtn">
                                &nbsp;<asp:Button ID="btQuery" runat="server" Text="查詢" OnClick="btQuery_Click" CssClass="button"  onmouseover="BtnMouseOver(this,'button_1');" onmouseout="BtnMouseOut(this,'button');" /></td>
                        </tr>
                        <tr>
                        </tr>
                        <tr>
                            <td class="area_qry">
                                 <table width="100%">
                                     <tr>
                                        <td colspan="2" >
                                            <uc4:NoAuthFEC ID="NoAuthFEC1" runat="server"  />
                                            
                                        </td>
                                    </tr>
                                    <tr>
                                        <td width="10%" align="right" ><nobr>
                                                        <asp:Label ID="lblemp_nm" runat="server" SkinID="itm01" Text="姓名："></asp:Label></nobr></td>
                                        <td>
                                            <asp:TextBox ID="txtNm" runat="server"></asp:TextBox></td>
                                    </tr>
                                 </table>
                            </td>
                        </tr>
                        
                    </table>
                </td>
            </tr>
            <tr>
                <td>
                    <asp:Button ID="btChoice" runat="server"   
                        Text="選取" OnClick="btChoice_Click" CssClass="button"  onmouseover="BtnMouseOver(this,'button_1');" onmouseout="BtnMouseOut(this,'button');" />
                                <input id="hidCheckAllCount" runat="server" name="hidCheckAllCount" style="width: 33px"
                                    type="hidden" /></td>
            </tr>
            <tr>
                <td>
                    <asp:GridView ID="Gw_PickUser" runat="server" AutoGenerateColumns="False"
                        Width="100%" DataKeyNames="user_id,emp_name" OnPageIndexChanging="Gw_PickUser_PageIndexChanging" OnRowDataBound="Gw_PickUser_RowDataBound" SkinID="GridView02" AllowPaging="True" EmptyDataText="<%$ resources: Strings, notdata %>">
                        <Columns>
                            <asp:TemplateField HeaderText="勾選">
                                <ItemStyle HorizontalAlign="Center" />
                                <HeaderTemplate>
                                    <asp:CheckBox ID="chkAll" runat="server" AutoPostBack="True" OnCheckedChanged="chkAll_CheckedChanged"
                                        Text="<%$ resources: Strings, chk_all  %>" />
                                </HeaderTemplate>
                                <ItemTemplate>
                                    <asp:CheckBox ID="chkUserID" runat="server" meta:resourcekey="chkUserIDResource1"
                                        OnCheckedChanged="chkUserID_CheckedChanged" /><asp:HiddenField ID="HidID" runat="server"
                                            Value='<%# Eval("user_id") %>' /><asp:HiddenField ID="HidNM" runat="server"
                                            Value='<%# Eval("emp_name") %>' />
                                </ItemTemplate>
                            </asp:TemplateField>
                            <asp:BoundField DataField="emp_name" HeaderText="<%$ resources: Strings, emp_name %>" />
                            <asp:BoundField DataField="fact_cname" HeaderText="<%$ resources: Strings, fact %>" />
                            <asp:BoundField DataField="dept_cname" HeaderText="<%$ resources: Strings, dept %>" />
                            <asp:BoundField DataField="group_cname" HeaderText="<%$ resources: Strings, group %>" />
                           
                        </Columns>
                        <HeaderStyle HorizontalAlign="Center" />
                    </asp:GridView>
                    <asp:HiddenField ID="hidSelEmployee" runat="server" />
                    <asp:HiddenField ID="hidSelAgent" runat="server" />
                </td>
            </tr>
        </table>
    
    </div>
    </form>
</body>
</html>

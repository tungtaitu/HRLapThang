<%@ Page Language="C#" AutoEventWireup="true" CodeFile="WorkFlowSequence.aspx.cs" Inherits="Pub_RemoteMethods_WorkFlowSequence" Theme="pwfBody"   Culture="auto"  UICulture="auto"%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
        <asp:Literal ID="litSpliter1" Text="<#BreakChar#>" runat="server"></asp:Literal>
        <asp:Literal ID="litSpliter2" Text="<#BreakChar#>" runat="server"></asp:Literal>
        <asp:GridView ID="gvPowerDetail" runat="server" AutoGenerateColumns="False" DataKeyNames="user_id,authorityId" Width="100%" OnRowDataBound="gvPowerDetail_RowDataBound" SkinID="GridView02" EmptyDataText="<%$ resources: Strings, notauth %>">
            <EmptyDataRowStyle BackColor="White" BorderColor="White" />
            <Columns>
                <asp:TemplateField HeaderText="<%$ resources: Strings, number %>">
                    <EditItemTemplate>
                        <asp:TextBox ID="TextBox1" runat="server"></asp:TextBox>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="Label1" runat="server"></asp:Label>
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="<%$ resources: Strings, comp %>">
                    <EditItemTemplate>
                        <asp:TextBox ID="TextBox2" runat="server"></asp:TextBox>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="Label2" runat="server" Text='<%# Bind("comp_cname") %>'></asp:Label>
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="<%$ resources: Strings, fact %>">
                    <EditItemTemplate>
                        <asp:TextBox ID="TextBox3" runat="server"></asp:TextBox>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="Label3" runat="server" Text='<%# Bind("fact_cname") %>'></asp:Label>
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="<%$ resources: Strings, dept %>">
                    <EditItemTemplate>
                        <asp:TextBox ID="TextBox4" runat="server"></asp:TextBox>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="Label4" runat="server" Text='<%# Bind("dept_cname") %>'></asp:Label>
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="<%$ resources: Strings, group %>">
                    <EditItemTemplate>
                        <asp:TextBox ID="TextBox5" runat="server"></asp:TextBox>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="Label5" runat="server" Text='<%# Bind("group_cname") %>'></asp:Label>
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="<%$ resources: Strings, manage %>">
                    <EditItemTemplate>
                        <asp:TextBox ID="TextBox6" runat="server"></asp:TextBox>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <img src="../../../App_Themes/pwfBody/images/fnBtn_del_0.gif" onclick="delUserAuth(this);" style="cursor:pointer" alt="刪除(Delete)" />
                        <input type="hidden" runat="server" id="hidRecID" value='<%# Eval("id") %>' />
                    </ItemTemplate>
                </asp:TemplateField>
            </Columns>
            <RowStyle HorizontalAlign="Center" />
            
        </asp:GridView>
    </form>
</body>
</html>

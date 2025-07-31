<%@ Page Language="C#" AutoEventWireup="true" CodeFile="WorkFlow.aspx.cs" Inherits="Pub_RemoteMethods_WorkFlow" Theme="pwfBody" Culture="auto"  UICulture="auto"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <asp:Literal ID="litSpliter1" runat="server" Text="<#BreakChar#>"></asp:Literal><asp:Literal
            ID="litSpliter2" runat="server" Text="<#BreakChar#>"></asp:Literal><asp:GridView
                ID="gv_flowuser" runat="server" AutoGenerateColumns="False" Width="100%" SkinID="GridView02" OnRowDataBound="gv_flowuser_RowDataBound">
                <Columns>
                    <asp:BoundField HeaderText="<%$ resources: Strings, order %>" />
                    <asp:BoundField DataField="comp_cname" HeaderText="<%$ resources: Strings, comp %>" />
                    <asp:BoundField DataField="fact_cname" HeaderText="<%$ resources: Strings, fact %>" />
                    <asp:BoundField DataField="dept_cname" HeaderText="<%$ resources: Strings, dept %>" />
                    <asp:BoundField DataField="group_cname" HeaderText="<%$ resources: Strings, group %>" />
                    <asp:BoundField DataField="posit_nm" HeaderText="<%$ resources: Strings, posit %>" />
                    <asp:BoundField DataField="emp_name" HeaderText="<%$ resources: Strings, emp_name %>" />
                </Columns>
                <RowStyle HorizontalAlign="Center" />
            </asp:GridView>
    
    </div>
    </form>
</body>
</html>

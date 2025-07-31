<%@ Page Language="C#" AutoEventWireup="true" CodeFile="GroupUser.aspx.cs" Inherits="Pub_RemoteMethods_GroupUser" Theme="pwfBody"  Culture="auto"  UICulture="auto"%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        &nbsp;<asp:Literal ID="litSpliter1" runat="server" Text="<#BreakChar#>"></asp:Literal><asp:Literal
            ID="litSpliter2" runat="server" Text="<#BreakChar#>"></asp:Literal><asp:GridView
                ID="gv_GroupUser" runat="server" EmptyDataText="<%$ resources: Strings, notdata %>" AutoGenerateColumns="False" SkinID="GridView02" Width="100%">
                <Columns>
                    <asp:BoundField DataField="comp_cname" HeaderText="<%$ resources: Strings, comp %>" />
                    <asp:BoundField DataField="fact_cname" HeaderText="<%$ resources: Strings, fact %>" />
                    <asp:BoundField DataField="dept_cname" HeaderText="<%$ resources: Strings, dept %>" />
                    <asp:BoundField DataField="group_cname" HeaderText="<%$ resources: Strings, group %>" />
                    <asp:BoundField DataField="emp_no" HeaderText="<%$ resources: Strings, emp_no %>" />
                    <asp:BoundField DataField="emp_name" HeaderText="<%$ resources: Strings, emp_name %>" />
                </Columns>
              
            </asp:GridView>
    
    </div>
    </form>
</body>
</html>

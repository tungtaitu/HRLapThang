<%@ Page Language="C#" AutoEventWireup="true" CodeFile="Detail_Over.aspx.cs" Inherits="Pub_RemoteMethods_Detail_Over" Theme="pwfBody" Culture="auto"  UICulture="auto"%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN">
<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
     <div>
        <asp:Literal ID="litSpliter1" runat="server" Text="<#BreakChar#>"></asp:Literal><asp:Literal
            ID="litSpliter2" runat="server" Text="<#BreakChar#>"></asp:Literal>
            <asp:GridView ID="get_Detail" runat="server" AutoGenerateColumns="False" OnRowDataBound="get_Detail_RowDataBound" Width="100%" SkinID="GridView02"  EmptyDataText="<%$ resources:Strings, not_overtime_user %>">
             <Columns>
               <asp:BoundField HeaderText="<%$ resources: Strings, number %>" />
               <asp:BoundField DataField="emp_no" HeaderText="<%$ resources: Strings, emp_no %>" />
               <asp:BoundField DataField="emp_name" HeaderText="<%$ resources: Strings, emp_name %>" />
               <asp:BoundField DataField="date" HeaderText="<%$ resources: Strings, over_date %>" />
               <asp:BoundField DataField="time_over" HeaderText="<%$ resources: Strings, overtime_date %>" />
               <asp:BoundField DataField="reason" HeaderText="<%$ resources: Strings, Reson_over %>" />               
           </Columns> 
            <RowStyle HorizontalAlign="Center" />
        </asp:GridView>
        </div>
    </form>
</body>
</html>

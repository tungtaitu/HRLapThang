<%@ Page Language="C#" AutoEventWireup="true" CodeFile="Detail_Labsent.aspx.cs" Inherits="Pub_RemoteMethods_Detail_Labsent" Theme="pwfBody"  Culture="auto"  UICulture="auto" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <asp:Literal ID="litSpliter1" runat="server" Text="<#BreakChar#>"></asp:Literal><asp:Literal
            ID="litSpliter2" runat="server" Text="<#BreakChar#>"></asp:Literal><asp:GridView ID="get_LabsentDetail" runat="server" AutoGenerateColumns="False" OnRowDataBound="get_LabsentDetail_RowDataBound" Width="100%" SkinID="GridView02" EmptyDataText="<%$ resources:Strings, not_labsent_user %>">
             <Columns>
               <asp:BoundField HeaderText="<%$ resources:Strings, number %>" />
               <asp:BoundField DataField="emp_no" HeaderText="<%$ resources:Strings, emp_no %>" />
               <asp:BoundField DataField="emp_name" HeaderText="<%$ resources:Strings, emp_name %>" />
               <asp:BoundField DataField="dept_cname" HeaderText="<%$ resources:Strings, dept %>" />
               <asp:BoundField DataField="group_cname" HeaderText="<%$ resources:Strings, group %>" />
                 <asp:BoundField DataField="type_nm" HeaderText="<%$ resources:Strings, LabsentType %>" >
                     <ItemStyle HorizontalAlign="Center" />
                 </asp:BoundField>
               <asp:TemplateField HeaderText="<%$ resources:Strings, LabsentDate %>">                                                   
                   <ItemTemplate>
                  <nobr>
                      <asp:Label ID="lblStart_dateD" runat="server" Text='<%# Eval("sdate") %>'></asp:Label>~<asp:Label ID="lblEnd_dateD" runat="server" Text='<%# Eval("edate") %>'></asp:Label>
                       </nobr><br> ，<asp:Label ID="total" runat="server" Text="<%$ resources: Strings, total %>"></asp:Label>
                       <asp:Label ID="lblHr" runat="server" Text='<%# Eval("lhours") %>'></asp:Label>
                       <asp:Label ID="hr" runat="server" Text="<%$ resources: Strings, hr %>"></asp:Label>。
                   </ItemTemplate>
                   <ItemStyle HorizontalAlign="Left" />
               </asp:TemplateField>
                 <asp:BoundField DataField="lreason" HeaderText="<%$ resources:Strings, LabsentReason %>" />
                 <asp:BoundField DataField="agent_usernm" HeaderText="<%$ resources:Strings, agent_user %>" />
           </Columns> 
        </asp:GridView>
        </div>
    </form>
</body>
</html>

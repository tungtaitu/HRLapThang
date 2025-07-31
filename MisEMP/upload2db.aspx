<%@ Page Language="C#" AutoEventWireup="true" CodeFile="upload2DB.aspx.cs" Inherits="uplod2DB" debug="true" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>Import Excel Data into Database</title>
</head>
<body style="margin-left:60px;margin-top:10px">
	<form id="form1" runat="server">
	<div style="font-family:verdana;color:Green;font-size:12pt;font-weight:bold" >
		Step1 : Import Excel		
    
    <asp:Panel ID="Panel1" runat="server">
       <asp:FileUpload ID="FileUpload1" runat="server" />
       <asp:Button ID="btnUpload" runat="server" Text="Upload" OnClick="btnUpload_Click" />
        <asp:Label ID="Label4" runat="server" Text="Has Header ?" Visible="false" ></asp:Label>
        <asp:RadioButtonList ID="RadioButtonList1" runat="server" Visible="false" >
            <asp:ListItem Text = "Yes" Value = "Yes"  ></asp:ListItem>
            <asp:ListItem Text = "No" Value = "No" Selected = "True"></asp:ListItem>
        </asp:RadioButtonList>
				 
        <asp:GridView ID="GridView1" runat="server"     >
        </asp:GridView>
        <br/><br/>
        <asp:Label ID="lblMessage" runat="server" Text="" style="font-weight:bold"></asp:Label>
    </asp:Panel>
	</div>
    <asp:Panel ID="Panel2" runat="server" Visible = "false" >
        <asp:Label ID="Label5" runat="server" Text="File Name:" />
        <asp:Label ID="lblFileName" runat="server" Text="" style="color:blue"/>
				<asp:Label ID="lblrelt" runat="server" Text="" style="color:blue"/>
        <br /><br />
        <asp:Label ID="Label2" runat="server" Text="Select Sheet" />
        <asp:DropDownList ID="ddlSheets" runat="server" AppendDataBoundItems = "true">
        </asp:DropDownList>
		&nbsp; <asp:Button ID="btnSave" runat="server" Text="Save" OnClick="btnSave_Click" />
        <asp:Button ID="btnCancel" runat="server" Text="Cancel" OnClick="btnCancel_Click" />        
        <br />
        <asp:Label ID="Label3" runat="server" Text="Enter Source Table Name"  Visible="false"/>
        <asp:TextBox ID="txtTable" runat="server"  Visible="false"></asp:TextBox>
        <br />
        <asp:Label ID="Label1" runat="server" Text="Has Header Row?" Visible="false" ></asp:Label>
        <br />
        <asp:RadioButtonList ID="rbHDR" runat="server" Visible="false">
            <asp:ListItem Text = "Yes" Value = "Yes"   ></asp:ListItem>
            <asp:ListItem Text = "No" Value = "No" Selected = "True"></asp:ListItem>
        </asp:RadioButtonList>
        <br />
        
     </asp:Panel>
	 </div>	  
    </form>
</body>
</html>

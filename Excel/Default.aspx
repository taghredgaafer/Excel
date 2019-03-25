<%@ Page Title="Home Page" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="Default.aspx.cs" Inherits="Excel._Default" %>

<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">


    <asp:Panel ID="Panel1" runat="server">
        <div class="jumbotron">
            <asp:FileUpload ID="FileUpload1" runat="server" />
            <asp:Button ID="btnUpload" runat="server" Text="Upload" OnClick="btnUpload_Click" />
            <br />
            <asp:Label ID="lblMessage" runat="server" Text="" />


        </div>
    </asp:Panel>

    <asp:Panel ID="Panel2" runat="server" Visible="false">
        <asp:Label ID="Label1" runat="server" Text="File Name"></asp:Label>
        <asp:Label ID="lblFileName" runat="server" Text=""></asp:Label>

        <asp:Label ID="Label2" runat="server" Text="Select sheet"></asp:Label>
        <asp:DropDownList ID="ddlSheet" runat="server" AppendDataBoundItems="true"></asp:DropDownList>
        <br />

        <asp:Label ID="Label3" runat="server" Text="Enter Source Table Name" />

        <asp:TextBox ID="txtTable" runat="server"> </asp:TextBox>

        <br />

        <asp:Label ID="Label4" runat="server" Text="Has Header Row?" />

        <br />

        <asp:RadioButtonList ID="rbHDR" runat="server">

            <asp:ListItem Text="Yes" Value="Yes" Selected="True"> </asp:ListItem>

            <asp:ListItem Text="No" Value="No"></asp:ListItem>

        </asp:RadioButtonList>

        <br />

        <asp:Button ID="btnSave" runat="server" Text="Save"
            OnClick="btnSave_Click" />

        <asp:Button ID="btnCancel" runat="server" Text="Cancel"
            OnClick="btnCancel_Click" />
    </asp:Panel>
</asp:Content>


<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Page1.aspx.cs" Inherits="ReadWordDoc.Page1" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <asp:FileUpload ID="WordFileToRead" runat="server" Width="500px" /> 
        <asp:Button ID="btnUpload" runat="server" Text="Read File" OnClick="btnUpload_Click" /> 
    </div>
    </form>
</body>
</html>

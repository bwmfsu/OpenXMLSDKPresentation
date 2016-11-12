<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Default.aspx.cs" Inherits="GenerateDocxDemo._Default"
    ValidateRequest="false" ViewStateEncryptionMode="Always" %>

<%@ Register Assembly="AjaxControlToolkit" Namespace="AjaxControlToolkit" TagPrefix="asp" %>
<%@ Register TagPrefix="FTB" Namespace="FreeTextBoxControls" Assembly="FreeTextBox" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
    <asp:ToolkitScriptManager ID="toolkitScriptManager" runat="server">
    </asp:ToolkitScriptManager>
    <asp:TabContainer ID="tabContainerMain" runat="server" Height="300">
        <asp:TabPanel runat="server" HeaderText="Open XML SDK Demo">
            <ContentTemplate>
                <table>
                    <tr>
                        <td>
                            <FTB:FreeTextBox ID="freeTextoBox" runat="server" Height="200" />
                        </td>
                        <td valign="top" style="padding-left: 10px;">
                            Choose document to upload:
                            <asp:FileUpload ID="fileUploadDocx" runat="server" />
                            <br />
                            <br />
                            <asp:Button ID="btnMixedContent" runat="server" Text="Generate Document" />
                        </td>
                    </tr>
                </table>
            </ContentTemplate>
        </asp:TabPanel>
        <asp:TabPanel runat="server" HeaderText="Extra">
            <ContentTemplate>
                <table>
                    <tr>
                        <td>
                            Number of people:
                            <td>
                                <asp:TextBox ID="txtNumPeople" runat="server" />
                            </td>
                        </td>
                        <td>
                            <asp:RangeValidator ID="cmprNumPeople" runat="server" Type="Integer" ControlToValidate="txtNumPeople"
                                MaximumValue="300" MinimumValue="1" ErrorMessage="Must be a number" />
                        </td>
                    </tr>
                    <tr>
                        <td>
                            Number of giveaways:
                        </td>
                        <td>
                            <asp:TextBox ID="txtNumGiveAways" runat="server" />
                        </td>
                        <td>
                            <asp:RangeValidator ID="cmprNumGiveaways" runat="server" Type="Integer" ControlToValidate="txtNumGiveAways"
                                MaximumValue="300" MinimumValue="1" ErrorMessage="Must be a number" />
                        </td>
                    </tr>
                    </table>
                    <br />
                    <asp:Button ID="btnExtra" runat="server" Text="Generate Document" />
            </ContentTemplate>
        </asp:TabPanel>
    </asp:TabContainer>
    </form>
</body>
</html>

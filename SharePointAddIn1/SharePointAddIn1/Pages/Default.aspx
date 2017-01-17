<%-- The following 4 lines are ASP.NET directives needed when using SharePoint components --%>

<%@ Page Inherits="Microsoft.SharePoint.WebPartPages.WebPartPage, Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" MasterPageFile="~masterurl/default.master" Language="C#" %>

<%@ Register TagPrefix="Utilities" Namespace="Microsoft.SharePoint.Utilities" Assembly="Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Register TagPrefix="WebPartPages" Namespace="Microsoft.SharePoint.WebPartPages" Assembly="Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Register TagPrefix="SharePoint" Namespace="Microsoft.SharePoint.WebControls" Assembly="Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>

<%-- The markup and script in the following Content element will be placed in the <head> of the page --%>
<asp:Content ContentPlaceHolderID="PlaceHolderAdditionalPageHead" runat="server">
    <script type="text/javascript" src="../Scripts/jquery-1.9.1.min.js"></script>
    <SharePoint:ScriptLink name="sp.js" runat="server" OnDemand="true" LoadAfterUI="true" Localizable="false" />
    <meta name="WebPartPageExpansion" content="full" />

    <!-- Add your CSS styles to the following file -->
    <link rel="Stylesheet" type="text/css" href="../Content/App.css" />

    <!-- Add your JavaScript to the following file -->
    <script type="text/javascript" src="../Scripts/App.js"></script>
</asp:Content>

<%-- The markup in the following Content element will be placed in the TitleArea of the page --%>
<asp:Content ContentPlaceHolderID="PlaceHolderPageTitleInTitleArea" runat="server">
    Page Title
</asp:Content>

<%-- The markup and script in the following Content element will be placed in the <body> of the page --%>
<asp:Content ContentPlaceHolderID="PlaceHolderMain" runat="server">

    <div>
        <div>
            <p>The purpose of this add-in is to showcase an issue with the session/cookie handling in the latest Office 365 update.</p>
            <ol>
                <li>Upload a file to the host-web. <em>It's easiest to use the Documents library and creating an empty Word file.</em></li>
                <li>Click &quot;Get files&quot; to perform a REST-query and print the top 10 files from the host-web.</li>
                <li>Click on a file in the results below, the file opens in a new tab. <em>Notice the quick redirect via login.microsoftonline.com.</em> You may close the tab once loaded.</li>
                <li>Click &quot;Get files&quot; again, you will receive an error. <em>Your session/cookie is invalid, the server replies with an error.</em></li>
                <li>Click &quot;Re-open add-in&quot; to open the add-in in a new tab. <em>Notice the quick redirect via login.microsoftonline.com.</em> You may close the tab once loaded.</li>
                <li>Click &quot;Get Files&quot; and it works again. <em>The previous step made our browser update its session/cookie for this particular sub-domain, making the original tab work again.</em></li>
            </ol>
        </div>
        <div>
            <button id="test_query" type="button">Get files</button>
            <button id="test_reopen" type="button">Re-open add-in</button>
        </div>
        <hr />
        <div id="test_message">
        </div>
    </div>

</asp:Content>

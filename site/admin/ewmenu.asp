<!--#include file="ewcfg9.asp"-->
<!--#include file="aspfn9.asp"-->
<!--#include file="userfn9.asp"-->
<%

' Get Menu Text
Function GetMenuText(Id, Text)
	GetMenuText = Language.MenuPhrase(Id, "MenuText")
	If GetMenuText = "" Then GetMenuText = Text
End Function
%>
<!-- Begin Main Menu -->
<div class="aspmaker">
<%

' Initialize language object
Set Language = New cLanguage
Call Language.LoadPhrases()

' Generate all menu items
Dim RootMenu
Set RootMenu = new cMenu
RootMenu.Id = "RootMenu"
RootMenu.IsRoot = True
RootMenu.AddMenuItem &HFFFFFFFF, Language.Phrase("Logout"), "logout.asp", -1, "", "", IsLoggedIn(), False
RootMenu.AddMenuItem &HFFFFFFFF, Language.Phrase("Login"), "login.asp", -1, "", "", (Not IsLoggedIn() And Right(Request.ServerVariables("URL"), Len("login.asp")) <> "login.asp"), False
Call RootMenu.Render(False)
Set RootMenu = Nothing
%>
</div>
<!-- End Main Menu -->

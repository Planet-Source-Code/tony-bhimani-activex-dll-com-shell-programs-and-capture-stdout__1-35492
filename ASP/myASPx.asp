<%

Dim myASPx, strOut

' create an object of our DLL
Set myASPx = Server.CreateObject("myASPx.Exec")

myASPx.AppName = "help.exe"   ' display Dos help
myASPx.Params = ""            ' pass no parameters
myASPx.Showwindow = 0         ' hide the window so it doesn't pop up

strOut = myASPx.DosExec()

%>

<HTML>
<HEAD>
<TITLE>Capture Dos Help</TITLE>
</HEAD>
<BODY>
Shell of Dos Help and display the output.
<br><br>
Dos Help says:
<br><br>
<PRE>

<%= strOut %>

</PRE>
</BODY>
</HTML>
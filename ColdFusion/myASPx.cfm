<HTML>
<HEAD>
<TITLE>Ping Example</TITLE>
</HEAD>
<BODY>
<BR>
Calling myASPx.dll from ColdFusion

<cfobject type="COM" action="create" class="myASPx.Exec" name="ASPx">
<cfset ASPx.AppName = "ping.exe">
<cfset ASPx.Params = "-n 5 www.yahoo.com">
<cfset ASPx.Showwindow = 0>
<cfset strOut = ASPx.DosExec()>

<PRE>

<cfoutput>
Ping results for www.yahoo.com:
#strOut#
</cfoutput>

</PRE>

</BODY>
</HTML>
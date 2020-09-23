Project: myASPx.dll (COM DLL for use by apps or a
         Web Server. Can be used by ASP or ColdFusion)

Description: Very similar to a Free server add-on called
             ASPExec which executes DOS or Windows apps
             and returns the STDOUT. This can be used
             for many purposes. e.g. CGI, ping, tracert,
             shelling system tasks, etc... and displaying
             the output to a webpage.

Author: Tony Bhimani
Email: TonyBhimani@hotmail.com
ICQ: 21634206
Homepage: http://www.xenocafe.com
Demo scripts: http://demos.xenocafe.com

Standard I/O module (MGetCmdOutput.bas) was written by
Mattias Sjögren. Refer to the file for more information.


CONTENTS
==============================================

Folders:

ActiveX DLL - Contains the project and source code to the DLL
ASP         - Contains the ASP example file for using the DLL (this
              would go on your Web Server that has ASP installed.)
ColdFusion  - Contains the ColdFusion example file that for using the 
              DLL (this would go on your Web Server that has ColdFusion
              installed.)
VB Example  - Contains the source code for the VB example project that 
              uses the DLL


INTRODUCTION
==============================================

This project was written/compiled with Visual Basic 6. Just open up the
VB project and compile. It should not give any errors. It is an ActiveX
DLL which can be used to shell programs and retrieve any STDOUT produced
by the program. I have used it to execute tools such as ping, traceroute,
and whois. Another example could be an EXE you wrote to generate a password. 
Use this to shell it and retrieve the password and display it on a webpage.
It is similar to a freeware program called ASPExec. I wrote this just to
see if I could produce the same goals. =)


myASPx Methods and Properties
==============================================

myASPx has one class called Exec. From this class there are three properties
and two methods.

Properties include:

AppName    - Sets or returns the application path and/or name to be shelled.
Params     - Sets or returns the parameters to be passed to the shelled EXE.
ShowWindow - Sets or returns the state of the shelled EXE. 0 = Hide, 1 = Show normal
             If this DLL is to be used as a Web Server add-on, I would highly
             suggest shelling the program as hidden. That way it will not popup
             on the server and interrupt anything. Also, this is EXTREMELY 
             IMPORTANT to remember. Once the app is shelled, control will not
             be returned to the DLL till the called app is terminated. So if
             the process fails or locks up and the program was set to be hidden,
             terminating it will not be easy. Just keep that in mind.

Methods include:

DosExec    - Shells a Dos program (uses command.com in Windows 9x or cmd.exe for 
             Windows NT) and returns and STDOUT as a string.
WinExec    - Shells a Windows program and returns any STDOUT produced (console window)


INSTALLATION
==============================================

Once the DLL has been compiled you can move it to any location. I suggest
your Windows\system directory. You then have to register the component on 
your system if you plan to use it with a Web Server and call it through
ASP or ColdFusion.

To register use the regsvr32.exe program.

Example:

regsvr32.exe x:\path\dllname

or in other words

regsvr32.exe c:\winnt\system32\myASPx.dll

Once registered the class is available for use. 

There is a redirection issue on Windows 9x MS-DOS Applications. So in the 
ActiveX DLL folder there is a .cpp file that creates a 32-bit console window
that shells. This is to be used instead of 16-bit command.com on Win9x
systems. To compile the .cpp, create a new project called "RedirStub" in 
Visual C++ (make sure its a console application), tell it to create an Empty 
project and add the .cpp file. It should compile with no problems. Put the 
RedirStub.exe in the same location as you put the myASPx.dll otherwise using 
this DLL on Win95 will fail.

For more information on the Windows 9x issue, see the MS Knowledge Base article 
Q150956 or visit this link

http://support.microsoft.com/default.aspx?scid=kb;EN-US;q150956


USING THE DLL IN AN EXE
==============================================

You can also use this DLL in a program. Just set a reference to it through
the Project, References menu in VB. A sample application has been provided
on how to implement it. Using the DLL through C or C++ shouldn't be too
difficult. I haven't tried it yet, but I plan to at a later date.


USING THE DLL IN ASP OR COLDFUSION
==============================================

First step is to register the DLL with Windows. See the installation section.
I have provided two examples. First example is how to call the DLL through
ASP (Active Server Pages.) Very basic and easy to use. Think of it the same 
way you create any other object like ADODB for databases. Use 
Server.CreateObject and the rest is a piece of cake. The ColdFusion sample
works exactly the same way. Now I admit I have not used ColdFusion for more 
than a week. There may already be a method of shelling EXE's and capturing 
STDOUT. But as far as I know, there isn't.


CONCLUSION
==============================================

I admit that this might not be the most efficient or best coding in the world.
I am always open to suggestions and you can contact me at any time if you 
require assistance. Remember that this code is free and it is USE AT YOUR OWN
RISK. I cannot and will not be held liable for any data loss, loss of profits,
system crashes, etc... It works perfectly on my system and it should on yours 
also. If you use this code/DLL in any commercial app, please consult me first. 
That is all I ask. =)


Kind regards,
Tony
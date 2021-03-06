# IbmiXmlserviceStd
There is a **SSH-enabled version of the library commercially available named: IbmiXmlserviceStdSsh**. You can visit http://www.mobigogo.net or reach out via email to info@mobigogo.net for more info.
Documentation can be reviewed here: http://mobigogo.net/files/docs/IbmiXmlserviceStdSsh 

IBM i C# and VB.Net XMLSERVICE Data Access Service Wrapper for .Net and .Net Core

This class is used to interface with existing IBM i database, program calls, CL commands, service programs and 
data queues via the XMLSERVICE service program. It also needs an Apache instance hooked up for XMLCGI calls
since this class uses the HTTP interface. 
 
Return data is returned in a usable .Net DataTable format or you can process the raw XML responses yourself.
 
This class should work with V5R4 and above of the OS400/IBM i operating system running XMLSERVICE.
 
Requirement: The XMLSERVICE library must also exist and be compiled on the system, including the XMLCGI program. 
Apache instance must also be set up and running and configured for XMLCGI calls.
 
Note: For appropriate security you should configure your Apache instance for SSL and for user Authentication. This
way all traffic is secured and there is an extra User/Password layer. To enable HTTP authentication use the 
UseHttpCredentials parameter and set it to True on the SetHttpUserInfo method.

You can always refer to the Yips site for more install info on XMLSERVICE: 
http://yips.idevcloud.com/wiki/index.php/xmlservice/xmlserviceinstall 

Class has been tested with XML Tookit library (XMLSERVICE) V1.9.7. 

There is a Nuget package available to install the DLL in to a new .Net 4.6.1+ or .Net Core 2.0.0+ application.

IBM i Multi-platform Xmlservice .Net Wrapper for Dotnet Standard Nuget Site Link
https://www.nuget.org/packages/IbmiXmlserviceStd

XMLSERVICE is now typically packaged on the IBMi in library: **QXMLSERV** as part of the operating system, but if you want to play with the code, here's the current Github location as of 1/23/2019

https://github.com/IBM/xmlservice

Note: 5/12/2019 - When working with the source, make sure to work with the C# code in: ***IbmiXmlserviceStdCs***. The VB code in ***IbmiXmlserviceStd*** will not be maintained going forward.






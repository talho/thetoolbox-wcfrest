using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.EnterpriseServices;

// General Information about an assembly is controlled through the following 
// set of attributes. Change these attribute values to modify the information
// associated with an assembly.
[assembly: AssemblyTitle("ExchSvc")]
[assembly: AssemblyDescription("A .NET RESTful API for creating and managing Microsoft Exchange 2007 mailboxes")]
[assembly: AssemblyConfiguration("")]
[assembly: AssemblyCompany("TALHO")]
[assembly: AssemblyProduct("ExchSvc")]
[assembly: AssemblyCopyright("Copyright ©  2010")]
[assembly: AssemblyTrademark("")]
[assembly: AssemblyCulture("")]

// Setting ComVisible to false makes the types in this assembly not visible 
// to COM components.  If you need to access a type in this assembly from 
// COM, set the ComVisible attribute to true on that type.
[assembly: ComVisible(true)]

// The following GUID is for the ID of the typelib if this project is exposed to COM
[assembly: Guid("920134ef-ec22-4c8c-b12d-21eed4c1d1e9")]

// Version information for an assembly consists of the following four values:
//
//      Major Version
//      Minor Version 
//      Build Number
//      Revision
//
// You can specify all the values or you can default the Revision and Build Numbers 
// by using the '*' as shown below:
// [assembly: AssemblyVersion("1.0.*")]
[assembly: AssemblyVersion("1.0.0.0")]
[assembly: AssemblyFileVersion("1.0.0.0")]

[assembly: ApplicationActivation(ActivationOption.Server)]
[assembly: ApplicationName("PowerShellComponent")]
[assembly: Description("A Powershell Component for managing Microsoft Exchange 2007 accounts via a COM service")]
[assembly: ApplicationAccessControl(
           false,
           AccessChecksLevel = AccessChecksLevelOption.Application,
           Authentication = AuthenticationOption.None,
           ImpersonationLevel = ImpersonationLevelOption.Identify)]
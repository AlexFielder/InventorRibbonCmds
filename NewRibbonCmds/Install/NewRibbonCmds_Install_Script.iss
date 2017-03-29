; Script generated by the Inno Setup Script Wizard.
; SEE THE DOCUMENTATION FOR DETAILS ON CREATING INNO SETUP SCRIPT FILES!

#define MyAppName "CH2M HILL Ribbon Commands"
#define MyAppVerName "Ribbon Commands 0.0.1.1"
#define MyAppPublisher "CH2M HILL LIMITED"
#define MyAppURL "http://www.ch2m.com/"
#include "scripts\products.iss"

#include "scripts\products\winversion.iss"
#include "scripts\products\fileversion.iss"

//#include "scripts\products\iis.iss"

#include "scripts\products\kb835732.iss"
//#include "scripts\products\kb886903.iss"
//#include "scripts\products\kb928366.iss"

//#include "scripts\products\msi20.iss"
//#include "scripts\products\msi31.iss"
//#include "scripts\products\ie6.iss"

//#include "scripts\products\dotnetfx11.iss"
//#include "scripts\products\dotnetfx11lp.iss"
//#include "scripts\products\dotnetfx11sp1.iss"

//#include "scripts\products\dotnetfx20.iss"
//#include "scripts\products\dotnetfx20lp.iss"
//#include "scripts\products\dotnetfx20sp1.iss"
//#include "scripts\products\dotnetfx20sp1lp.iss"
//#include "scripts\products\dotnetfx20sp2.iss"
//#include "scripts\products\dotnetfx20sp2lp.iss"

//#include "scripts\products\dotnetfx35.iss"
//#include "scripts\products\dotnetfx35lp.iss"
//#include "scripts\products\dotnetfx35sp1.iss"
//#include "scripts\products\dotnetfx35sp1lp.iss"

//#include "scripts\products\mdac28.iss"
//#include "scripts\products\jet4sp8.iss"
//#include "scripts\products\sql2005express.iss"
#include "dotnetfx40.iss"

[Setup]
; NOTE: The value of AppId uniquely identifies this application.
; Do not use the same AppId value in installers for other applications.
; (To generate a new GUID, click Tools | Generate GUID inside the IDE.)
AppId={{971EF5C6-56A6-4A5D-98B3-E7D320682F37}
AppName={#MyAppName}
AppVerName={#MyAppVerName}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}
DefaultDirName={pf}\Ribbon Commands
DefaultGroupName=CH2M HILL\RIBBON COMMANDS
DisableProgramGroupPage=true
OutputDir=C:\Users\afielder\Documents\Visual Studio 2010\Projects\NewRibbonCmds\NewRibbonCmds\Install
OutputBaseFilename=setup
SetupIconFile=C:\Users\afielder\Documents\Visual Studio 2010\Projects\NewRibbonCmds\NewRibbonCmds\Resources\CH2M_Globe.ico
Compression=lzma
SolidCompression=true
ArchitecturesInstallIn64BitMode=x64
VersionInfoCompany=CH2M HILL LIMITED
VersionInfoDescription=Inventor Ribbon Commands
VersionInfoProductName=Inventor Ribbon Commands
VersionInfoProductVersion=0.0.1.1
AppCopyright=CH2MHILL

[Languages]
Name: en; MessagesFile: compiler:Default.isl

[Files]
Source: C:\Users\afielder\Documents\Visual Studio 2010\Projects\NewRibbonCmds\NewRibbonCmds\bin\NewRibbonCmds.dll; DestDir: {app}; Flags: ignoreversion
; NOTE: Don't use "Flags: ignoreversion" on any shared system files

[Icons]
Name: {group}\{cm:UninstallProgram,{#MyAppName}}; Filename: {uninstallexe}

[Run]
Filename: {dotnet40}\regasm.exe; Parameters: "/codebase ""{app}\NewRibbonCmds.dll"""; WorkingDir: {app}; Description: Registers the commands with Inventor, allows the tool to load.; StatusMsg: Registering with Inventor; Languages: 
[UninstallRun]
Filename: {dotnet40}\regasm.exe; Parameters: "/unregister ""{app}\NewRibbonCmds.dll"""; WorkingDir: {app}; Flags: runhidden
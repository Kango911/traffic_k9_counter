; install script for Traffic K9 Counter

[Setup]
AppName=Traffic K9 Counter
AppVersion=2.0.0
AppPublisher=Kango911
AppPublisherURL=https://github.com/Kango911/traffic_k9_counter
AppSupportURL=https://github.com/Kango911/traffic_k9_counter/issues
DefaultDirName={pf}\TrafficK9Counter
DefaultGroupName=Traffic K9 Counter
UninstallDisplayIcon={app}\TrafficK9Counter.exe
Compression=lzma2
SolidCompression=yes
OutputDir=installer_output
OutputBaseFilename=TrafficK9Counter_Setup_v2.0.0
SetupIconFile=icon.ico
WizardStyle=modern
PrivilegesRequired=admin

[Languages]
Name: "russian"; MessagesFile: "compiler:Languages\Russian.isl"
Name: "english"; MessagesFile: "compiler:Default.isl"

[Files]
Source: "dist\TrafficK9Counter.exe"; DestDir: "{app}"
Source: "icon.ico"; DestDir: "{app}"; Flags: dontcopy

[Icons]
Name: "{group}\Traffic K9 Counter"; Filename: "{app}\TrafficK9Counter.exe"
Name: "{group}\Uninstall Traffic K9 Counter"; Filename: "{uninstallexe}"
Name: "{autodesktop}\Traffic K9 Counter"; Filename: "{app}\TrafficK9Counter.exe"

[Run]
Filename: "{app}\TrafficK9Counter.exe"; Description: "{cm:LaunchProgram,Traffic K9 Counter}"; Flags: postinstall nowait skipifsilent

[UninstallDelete]
Type: files; Name: "{app}\vehicle_types_auto.json"
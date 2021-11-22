; Script generated by the Inno Setup Script Wizard.
; SEE THE DOCUMENTATION FOR DETAILS ON CREATING INNO SETUP SCRIPT FILES!

[Setup]
; NOTE: The value of AppId uniquely identifies this application.
; Do not use the same AppId value in installers for other applications.
; (To generate a new GUID, click Tools | Generate GUID inside the IDE.)
AppId={{01C7C60E-5580-4B10-8196-6484F41C60F8}
AppName=ECOLOG
AppVersion=6.0.2
AppPublisher=Ecoinformatics Studio
AppPublisherURL=http://ecolog.sourceforge.net/
AppSupportURL=http://ecolog.sourceforge.net/
AppUpdatesURL=http://ecolog.sourceforge.net/
AppCopyright=Copyright � 1990-2021 Mauro J. Cavalcanti
AppContact=maurobio@gmail.com
DefaultDirName={pf}\ECOLOG
DisableDirPage=no
DefaultGroupName=ECOLOG
DisableProgramGroupPage=yes
LicenseFile=gpl-2.0.txt
OutputBaseFilename=ECOLOG-6.0.2-setup
Compression=lzma
SolidCompression=yes
PrivilegesRequired=none
MinVersion=0,5.01sp3
WizardImageFile=C:\Users\mauro\My Documents\Projetos\ECOLOG\source\ECOLOG-5.0-src\calendula-bio.bmp
WizardSmallImageFile=C:\Users\mauro\My Documents\Projetos\ECOLOG\source\ECOLOG-5.0-src\sunflower.bmp
DisableWelcomePage=False

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"
Name: "brazilianportuguese"; MessagesFile: "compiler:Languages\BrazilianPortuguese.isl"
Name: "french"; MessagesFile: "compiler:Languages\French.isl"
Name: "portuguese"; MessagesFile: "compiler:Languages\Portuguese.isl"
Name: "spanish"; MessagesFile: "compiler:Languages\Spanish.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
Source: "ECOLOG.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "shapelib129.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "languages\ECOLOG.en.mo"; DestDir: "{app}\languages"; Flags: ignoreversion
Source: "languages\ECOLOG.en.po"; DestDir: "{app}\languages"; Flags: ignoreversion
Source: "languages\ECOLOG.po"; DestDir: "{app}\languages"; Flags: ignoreversion
Source: "languages\ECOLOG.pt_br.mo"; DestDir: "{app}\languages"; Flags: ignoreversion
Source: "languages\ECOLOG.pt_br.po"; DestDir: "{app}\languages"; Flags: ignoreversion
Source: "languages\ECOLOG.es.mo"; DestDir: "{app}\languages"; Flags: ignoreversion
Source: "languages\ECOLOG.es.po"; DestDir: "{app}\languages"; Flags: ignoreversion
Source: "C:\Users\mauro\My Documents\Projetos\ECOLOG\source\ECOLOG-5.0-src\samples\Exemplo de Planilha simples.json"; DestDir: "{app}\samples\"; Flags: ignoreversion
Source: "C:\Users\mauro\My Documents\Projetos\ECOLOG\source\ECOLOG-5.0-src\samples\Exemplo de Planilha simples.xls"; DestDir: "{app}\samples\"; Flags: ignoreversion
Source: "C:\Users\mauro\My Documents\Projetos\ECOLOG\source\ECOLOG-5.0-src\samples\Flora de Santa Genebra.xls"; DestDir: "{app}\samples\"; Flags: ignoreversion
Source: "C:\Users\mauro\My Documents\Projetos\ECOLOG\source\ECOLOG-5.0-src\samples\Flora de Santa Genebra.json"; DestDir: "{app}\samples\"; Flags: ignoreversion
Source: "C:\Users\mauro\My Documents\Projetos\ECOLOG\source\ECOLOG-5.0-src\samples\Flora do Itacolomi.ods"; DestDir: "{app}\samples\"; Flags: ignoreversion
Source: "C:\Users\mauro\My Documents\Projetos\ECOLOG\source\ECOLOG-5.0-src\samples\Flora do Itacolomi.json"; DestDir: "{app}\samples\"; Flags: ignoreversion
Source: "C:\Users\mauro\My Documents\Projetos\ECOLOG\source\ECOLOG-5.0-src\samples\Planilha de Variaveis Ambientais.xlsx"; DestDir: "{app}\samples\"; Flags: ignoreversion
Source: "C:\Users\mauro\My Documents\Projetos\ECOLOG\source\ECOLOG-5.0-src\scissor.png"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\mauro\My Documents\Projetos\ECOLOG\source\ECOLOG-5.0-src\gpl-2.0.txt"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\..\manual\Manual do ECOLOG.pdf"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\ECOLOG"; Filename: "{app}\ECOLOG.exe"; IconFilename: "{app}\ECOLOG.exe"; IconIndex: 0
Name: "{group}\{cm:ProgramOnTheWeb,ECOLOG}"; Filename: "http://ecolog.sourceforge.net"
Name: "{group}\{cm:UninstallProgram,ECOLOG}"; Filename: "{uninstallexe}"
Name: "{group}\Manual do ECOLOG"; Filename: "{app}\Manual do ECOLOG.pdf"
Name: "{commondesktop}\ECOLOG"; Filename: "{app}\ECOLOG.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\ECOLOG.exe"; Description: "{cm:LaunchProgram,ECOLOG}"; Flags: nowait postinstall skipifsilent

[Dirs]
Name: "{app}\samples\"
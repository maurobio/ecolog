unit about;

{$mode objfpc}{$H+}

interface

uses
    {$IFDEF Windows} Win32Proc, {$ENDIF} FileInfo, SysUtils, Classes, Graphics,
    Forms, Controls, StdCtrls, Buttons, ExtCtrls, LCLTranslator, LCLVersion,
    LCLIntf;

const
  VERSION = 'Dahlia';

type

    { TAboutBox }

    TAboutBox = class (TForm)
      Compiler: TLabel;
      IDE: TLabel;
      Comments: TLabel;
      Donation: TImage;
      Information: TLabel;
      Platform: TLabel;
      Website: TLabel;
      VersionNumber: TLabel;
      OperatingSystem: TLabel;
      Panel1: TPanel;
      ProgramIcon: TImage;
      ProductName: TLabel;
      Copyright: TLabel;
      OKButton: TButton;
      procedure FormShow(Sender: TObject);
      procedure DonationClick(Sender: TObject);
      procedure WebsiteClick(Sender: TObject);
    private
         { Private declarations }
    public
        { Public declarations }
    end;

    Str255 = string[255];

var
   AboutBox: TAboutBox;

function OSVersion: Str255;

implementation

uses main;

{$R *.lfm}

resourcestring
   strInformation = 'Informação do Sistema';
   strComments = 'Sistema de Banco de Dados para Levantamentos Ecológicos de Campo.';
   strWebsite = 'Página do ECOLOG na Internet';
   strCompiler = 'Compilador:';
   strPlatform = 'Plataforma:';

function OSVersion: Str255;
begin
  {$IFDEF LCLcarbon}
  OSVersion := 'Mac OS X 10.';
  {$ELSE}
  {$IFDEF Linux}
  OSVersion := 'Linux Kernel ';
  {$IFDEF CPU32}
  OsVersion := OSVersion + ' (32-bits)';
  {$ENDIF}
  {$IFDEF CPU64}
  OsVersion := OSVersion + ' (64-bits)';
  {$ENDIF}
  {$ELSE}
  {$IFDEF UNIX}
  OSVersion := 'Unix ';
  {$ELSE}
  {$IFDEF WINDOWS}
  if WindowsVersion = wv95 then OSVersion := 'Windows 95 '
   else if WindowsVersion = wvNT4 then OSVersion := 'Windows NT v.4 '
   else if WindowsVersion = wv98 then OSVersion := 'Windows 98 '
   else if WindowsVersion = wvMe then OSVersion := 'Windows ME '
   else if WindowsVersion = wv2000 then OSVersion := 'Windows 2000 '
   else if WindowsVersion = wvXP then OSVersion := 'Windows XP '
   else if WindowsVersion = wvServer2003 then OSVersion := 'Windows Server 2003 '
   else if WindowsVersion = wvVista then OSVersion := 'Windows Vista '
   else if WindowsVersion = wv7 then OSVersion := 'Windows 7 '
   else if WindowsVersion = wv8 then OSVersion := 'Windows 8 '
   else if WindowsVersion = wv10 then OSVersion := 'Windows 10 '
   else OSVersion:= 'Windows ';
  {$ENDIF}
  {$ENDIF}
  {$ENDIF}
  {$ENDIF}
end;

{ TAboutBox }

procedure TAboutBox.FormShow(Sender: TObject);
var
  Quad: TVersionQuad;
begin
  if GetProgramVersion(Quad) then
    VersionNumber.Caption := 'v.' + VersionQuadToStr(Quad) + ' (' + DateToStr(Now) + ') -- "' + VERSION + '"';
  Information.Caption := strInformation;
  Comments.Caption := strComments;
  Compiler.Caption := strCompiler + ' ' + Format('Free Pascal v.%s', [{$I %FPCVERSION%}]);
  IDE.Caption := 'IDE: ' + Format('Lazarus v.%s', [lcl_version]);
  Platform.Caption := strPlatform + ' ' + {$I %FPCTARGETOS%};
  OperatingSystem.Caption := 'OS: ' + OSVersion;
  Website.Caption := strWebsite;
end;

procedure TAboutBox.DonationClick(Sender: TObject);
var
  sUrl: string;
begin
  case sLang of
    'en', 'es':
      sUrl := 'https://www.paypal.com/donate?hosted_button_id=SM442YC8BTMKY';
    'pt_br':
      sUrl := 'https://www.paypal.com/donate?hosted_button_id=U4WEUX9S4CST2';
  end;
  OpenURL(sUrl);
end;

procedure TAboutBox.WebsiteClick(Sender: TObject);
begin
  OpenURL('http://ecolog.sourceforge.net');
end;

end.

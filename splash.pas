unit splash;

{$mode objfpc}{$H+}

interface

uses
  SysUtils, LCLIntf, LCLType, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, StdCtrls, Buttons;

type

  { TSplashForm }

  TSplashForm = class (TForm)
    Image: TImage;
    Timer: TTimer;
    procedure FormShow(Sender: TObject);
    procedure TimerTimer(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    Completed: Boolean;
  end;

var
  SplashForm: TSplashForm;

implementation

{$R *.lfm}

{ TSplashForm }

procedure TSplashForm.FormShow(Sender: TObject);
begin
  OnShow := nil;
  Completed := False;
  Timer.Interval := 2000; // 2s minimum time to show splash screen
  Timer.Enabled := True;
end;

procedure TSplashForm.TimerTimer(Sender: TObject);
begin
  Timer.Enabled := False;
  Completed := True;
end;

end.

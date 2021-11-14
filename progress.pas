unit progress;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, StdCtrls, ComCtrls;

type

  { TProgressDlg }

  TProgressDlg = class(TForm)
    CancelButton: TButton;
    LabelText: TLabel;
    ProgressBar: TProgressBar;
  private

  public

  end;

var
  ProgressDlg: TProgressDlg;

implementation

{$R *.lfm}

{ TProgressDlg }

end.


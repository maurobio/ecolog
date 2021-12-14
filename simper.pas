unit simper;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, StdCtrls, Spin;

type

  { TSIMPERDlg }

  TSIMPERDlg = class(TForm)
    OKButton: TButton;
    CancelButton: TButton;
    ComboBoxTransform: TComboBox;
    ComboBoxGroup: TComboBox;
    LabelTransform: TLabel;
    LabelGroup: TLabel;
    LabelNPerm: TLabel;
    SpinEditNPerm: TSpinEdit;
    procedure ComboBoxGroupChange(Sender: TObject);
    procedure FormCreate(Sender: TObject);
  private

  public
    Factors: string;
    procedure CreateSIMPER(fname: string; groupvar: string; transf, nperm: integer);
    procedure SIMPER(fname: string; groupvar: string; transf, nperm: integer;
      n, m: integer);
  end;

var
  SIMPERDlg: TSIMPERDlg;

implementation

{$R *.lfm}

uses group, report, useful;

resourcestring
  strSIMPER = 'PORCENTAGEM DE SIMILARIDADE';
  strGroup = 'GRUPOS: ';
  strTransform = 'Transformação: ';
  strTransformNone = 'Sem Transformação';
  strTransformCommonLog = 'Logaritmo comum (base 10)';
  strTransformNaturalLog = 'Logaritmo natural (base e)';
  strTransformSqrt = 'Raiz quadrada';
  strTransformArcsin = 'Arcosseno';
  strNPerm = 'Número de Permutações: ';
  strSamples = 'Amostras';
  strSpecies = 'Espécies';

{ TSIMPERDlg }

procedure TSIMPERDlg.FormCreate(Sender: TObject);
begin
  with ComboBoxTransform.Items do
  begin
    Add(strTransformNone);
    Add(strTransformCommonLog);
    Add(strTransformNaturalLog);
    Add(strTransformSqrt);
    Add(strTransformArcsin);
  end;
  ComboBoxTransform.ItemIndex := 0;
  ComboBoxGroup.ItemIndex := 0;
end;

procedure TSIMPERDlg.ComboBoxGroupChange(Sender: TObject);
begin
  GroupDlg.GroupVariable := ComboBoxGroup.Items[ComboBoxGroup.ItemIndex];
  if GroupDlg.ShowModal = mrOk then
    Factors := GroupDlg.Factors;
end;

procedure TSIMPERDlg.CreateSIMPER(fname: string; groupvar: string;
  transf, nperm: integer);
var
  stransf: string;
  outfile: TextFile;
begin
  case transf of
    0: stransf := '';
    1: stransf := 'log10';
    2: stransf := 'log';
    3: stransf := 'sqrt';
    4: stransf := 'asin';
  end;

  AssignFile(outfile, 'simper.R');
  Rewrite(outfile);
  WriteLn(outfile, 'options(warn=-1)');
  WriteLn(outfile, 'options(digits=4)');
  WriteLn(outfile, 'suppressPackageStartupMessages(library(vegan))');
  WriteLn(outfile, 'library(vegan, quietly=TRUE)');
  WriteLn(outfile, 'df.data <- read.csv("rdata.csv", row.names=1)');
  WriteLn(outfile, 'df.data <- t(df.data)');
  if Length(stransf) > 0 then
    WriteLn(outfile, 'df.data <- ' + stransf + '(df.data + 1)');
  WriteLn(outfile, 'groups <- as.factor(c(', groupvar, '))');
  WriteLn(outfile, 'sim <- simper(df.data, groups, permutations=', nperm, ')');
  WriteLn(outfile, 'sink("simper.txt")');
  WriteLn(outfile, 'print(summary(sim))');
  WriteLn(outfile, 'sink()');
  WriteLn(outfile, 'options(warn=0)');
  CloseFile(outfile);
end;

procedure TSIMPERDlg.SIMPER(fname: string; groupvar: string;
  transf, nperm: integer; n, m: integer);
var
  line: string;
  infile, outfile: TextFile;
begin
  AssignFile(outfile, fname);
  Rewrite(outfile);
  Header(outfile, strSIMPER);
  WriteLn(outfile, '<br>');
  WriteLn(outfile, IntToStr(n) + ' ' + LowerCase(strSamples) + ' x ' +
    IntToStr(m) + ' ' + LowerCase(strSpecies) + '<br><br>');

  if transf = 0 then
    WriteLn(outfile, strTransform + strTransformNone + '<br><br>')
  else if transf = 1 then
    WriteLn(outfile, strTransform + strTransformCommonLog + '<br><br>')
  else if transf = 2 then
    WriteLn(outfile, strTransform + strTransformNaturalLog + '<br><br>')
  else if transf = 3 then
    WriteLn(outfile, strTransform + strTransformSqrt + '<br><br>')
  else if transf = 4 then
    WriteLn(outfile, strTransform + strTransformArcsin + '<br><br>');

  Write(outfile, strNPerm + IntToStr(nperm) + '<br><br>');

  WriteLn(outfile, '<pre>');
  AssignFile(infile, 'simper.txt');
  Reset(infile);
  while not EOF(infile) do
  begin
    ReadLn(infile, line);
    WriteLn(outfile, line);
  end;
  CloseFile(infile);
  WriteLn(outfile, '</pre>');

  WriteLn(outfile, '</body>');
  WriteLn(outfile, '</html>');
  CloseFile(outfile);
  if FileExists('simper.txt') then
    DeleteFile('simper.txt');
end;

end.

unit cca;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, StdCtrls, StrUtils;

type

  { TCCADlg }

  TCCADlg = class(TForm)
    CheckBoxScale: TCheckBox;
    OKButton: TButton;
    CancelButton: TButton;
    ComboBoxTransform: TComboBox;
    LabelTransform: TLabel;
    procedure FormCreate(Sender: TObject);
  private

  public
    procedure CreateCCA(fname: string; transf, nvars: integer;
      scale: boolean; selected: string);
    procedure CCA(fname: string; transf: integer; n, m: integer; selected: string);
  end;

var
  CCADlg: TCCADlg;

implementation

{$R *.lfm}

uses report, useful;

resourcestring
  strCCa = 'ANÁLISE DE CORRESPONDÊNCIAS CANÔNICA';
  strScatterplot = 'DIAGRAMA DE DISPERSÃO';
  strAxis = 'EIXO ';
  strTransform = 'Transformação: ';
  strTransformNone = 'Sem Transformação';
  strTransformCommonLog = 'Logaritmo comum (base 10)';
  strTransformNaturalLog = 'Logaritmo natural (base e)';
  strTransformSqrt = 'Raiz quadrada';
  strTransformArcsin = 'Arcosseno';
  strScaleSamples = 'Amostras';
  strScaleSpecies = 'Espécies';
  strInertia = 'Inércia';
  strChiSquared = 'Qui-Quadrado';
  strVariables = 'Variáveis: ';

{ TCCADlg }

procedure TCCADlg.FormCreate(Sender: TObject);
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
  CheckBoxScale.Checked := False;
end;

procedure TCCADlg.CreateCCA(fname: string; transf, nvars: integer;
  scale: boolean; selected: string);
var
  model, stransf, figf: string;
  outfile: TextFile;
begin
  case transf of
    0: stransf := '';
    1: stransf := 'log10';
    2: stransf := 'log';
    3: stransf := 'sqrt';
    4: stransf := 'asin';
  end;

  if selected.CountChar(',') + 1 <> nvars then
    model := StringReplace(selected, ',', '+', [rfReplaceAll])
  else
    model := '';

  AssignFile(outfile, 'cca.R');
  Rewrite(outfile);
  WriteLn(outfile, 'options(warn=-1)');
  WriteLn(outfile, 'options(digits=4)');
  WriteLn(outfile, 'suppressPackageStartupMessages(library(vegan))');
  WriteLn(outfile, 'library(vegan, quietly=TRUE)');
  WriteLn(outfile, 'df.data <- read.csv("rdata1.csv", row.names=1)');
  WriteLn(outfile, 'df.data <- t(df.data)');
  WriteLn(outfile, 'df.vars <- read.csv2("rdata2.csv", row.names=1)');
  if Length(stransf) > 0 then
    WriteLn(outfile, 'df.data <- ' + stransf + '(df.data + 1)');
  WriteLn(outfile, 'attach(df.vars)');
  if Length(model) = 0 then
    WriteLn(outfile, 'cc <- cca(df.data, data=df.vars, scale=' +
      IfThen(scale, 'TRUE', 'FALSE') + ')')
  else
    WriteLn(outfile, 'cc <- cca(df.data~' + model + ', data=df.vars, scale=' +
      IfThen(scale, 'TRUE', 'FALSE') + ')');
  WriteLn(outfile, 'sink("cca.txt")');
  WriteLn(outfile, 'results <- summary(cc)');
  WriteLn(outfile, 'print(results)');
  WriteLn(outfile, 'sink()');
  WriteLn(outfile, 'ppi <- 100');
  figf := GetFileNameWithoutExt(fname) + '.png';
  WriteLn(outfile, 'png("' + figf + '", width=6*ppi, height=6*ppi, res=ppi)');
  WriteLn(outfile, 'par(mar=c(4,4,4,4))');
  WriteLn(outfile,
    'plot(cc, main=iconv("' + strScatterplot +
    '", from=''UTF-8'', to=''latin1''), xlab="' + strAxis + '1", ylab="' + strAxis +
    '2", col="darkgreen")');
  WriteLn(outfile, 'invisible(dev.off())');
  WriteLn(outfile, 'options(warn=0)');
  CloseFile(outfile);
end;

procedure TCCADlg.CCA(fname: string; transf: integer; n, m: integer; selected: string);
var
  line, figf: string;
  infile, outfile: TextFile;
begin
  AssignFile(outfile, fname);
  Rewrite(outfile);
  Header(outfile, strCCA);
  WriteLn(outfile, '<br>');
  WriteLn(outfile, IntToStr(n) + ' ' + LowerCase(strScaleSamples) +
    ' x ' + IntToStr(m) + ' ' + LowerCase(strScaleSpecies) + '<br><br>');

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

  if selected <> '' then
    WriteLn(outfile, strVariables + selected + '<br><br>');

  WriteLn(outfile, strInertia + ' = ' + strChiSquared + '<br><br>');

  WriteLn(outfile, '<pre>');
  AssignFile(infile, 'cca.txt');
  Reset(infile);
  while not EOF(infile) do
  begin
    ReadLn(infile, line);
    WriteLn(outfile, line);
  end;
  CloseFile(infile);
  WriteLn(outfile, '</pre>');

  figf := GetFileNameWithoutExt(fname) + '.png';
  WriteLn(outfile, '<p align="left"><img src="' + figf + '"></p>');

  WriteLn(outfile, '</body>');
  WriteLn(outfile, '</html>');
  CloseFile(outfile);
  if FileExists('cca.txt') then
    DeleteFile('cca.txt');
end;

end.

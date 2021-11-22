unit rda;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, StdCtrls, StrUtils;

type

  { TRDADlg }

  TRDADlg = class(TForm)
    ComboBoxInertia: TComboBox;
    ComboBoxScale: TComboBox;
    LabelInertia: TLabel;
    LabelScale: TLabel;
    OKButton: TButton;
    CancelButton: TButton;
    ComboBoxTransform: TComboBox;
    LabelTransform: TLabel;
    procedure FormCreate(Sender: TObject);
  private

  public
    procedure CreateRDA(fname: string; transf, scaling, inertia, nvars: integer;
      selected: string);
    procedure RDA(fname: string; transf, inertia: integer; n, m: integer;
      selected: string);
  end;

var
  RDADlg: TRDADlg;

implementation

{$R *.lfm}

uses report, useful;

resourcestring
  strRDA = 'ANÁLISE DE REDUNDÂNCIAS';
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
  strScaleSymmetric = 'Simétrica';
  strInertia = 'Inércia';
  strVariance = 'Variância';
  strCorrelation = 'Correlação';
  strVariables = 'Variáveis: ';

{ TRDADlg }

procedure TRDADlg.FormCreate(Sender: TObject);
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
  with ComboBoxScale.Items do
  begin
    Add(strScaleSamples);
    Add(strScaleSpecies);
    Add(strScaleSymmetric);
  end;
  ComboBoxScale.ItemIndex := 0;
  with ComboBoxInertia.Items do
  begin
    Add(strVariance);
    Add(strCorrelation);
  end;
  ComboBoxInertia.ItemIndex := 0;
end;

procedure TRDADlg.CreateRDA(fname: string; transf, scaling, inertia, nvars: integer;
  selected: string);
var
  model, stransf, figf: string;
  scale: boolean;
  outfile: TextFile;
begin
  case transf of
    0: stransf := '';
    1: stransf := 'log10';
    2: stransf := 'log';
    3: stransf := 'sqrt';
    4: stransf := 'asin';
  end;

  if inertia = 0 then
    scale := False
  else
    scale := True;

  if selected.CountChar(',') + 1 <> nvars then
    model := StringReplace(selected, ',', '+', [rfReplaceAll])
  else
    model := '';

  AssignFile(outfile, 'rda.R');
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
    WriteLn(outfile, 'ra <- rda(df.data, data=df.vars, scale=' +
      IfThen(scale, 'TRUE', 'FALSE') + ')')
  else
    WriteLn(outfile, 'ra <- rda(df.data~' + model + ', data=df.vars, scale=' +
      IfThen(scale, 'TRUE', 'FALSE') + ')');
  WriteLn(outfile, 'sink("rda.txt")');
  WriteLn(outfile, 'results <- summary(ra)');
  WriteLn(outfile, 'print(results)');
  WriteLn(outfile, 'sink()');
  WriteLn(outfile, 'ppi <- 100');
  figf := GetFileNameWithoutExt(fname) + '.png';
  WriteLn(outfile, 'png("' + figf + '", width=6*ppi, height=6*ppi, res=ppi)');
  WriteLn(outfile, 'par(mar=c(4,4,4,4))');
  WriteLn(outfile,
    'plot(ra, main=iconv("' + strScatterplot +
    '", from=''UTF-8'', to=''latin1''), xlab="' + strAxis + '1", ylab="' + strAxis +
    '2", scaling=' + IntToStr(scaling + 1) + ', col="darkgreen")');
  WriteLn(outfile, 'invisible(dev.off())');
  WriteLn(outfile, 'options(warn=0)');
  CloseFile(outfile);
end;

procedure TRDADlg.RDA(fname: string; transf, inertia: integer;
  n, m: integer; selected: string);
var
  line, figf: string;
  infile, outfile: TextFile;
begin
  AssignFile(outfile, fname);
  Rewrite(outfile);
  Header(outfile, strRDA);
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

  WriteLn(outfile, strInertia + ' = ' + IfThen(inertia = 0, strVariance,
    strCorrelation) + '<br><br>');

  WriteLn(outfile, '<pre>');
  AssignFile(infile, 'rda.txt');
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
  if FileExists('rda.txt') then
    DeleteFile('rda.txt');
end;

end.

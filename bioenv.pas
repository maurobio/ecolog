unit bioenv;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, StdCtrls, Spin;

type

  { TBIOENVDlg }

  TBIOENVDlg = class(TForm)
    ComboBoxMetric: TComboBox;
    LabelMetric: TLabel;
    OKButton: TButton;
    CancelButton: TButton;
    ComboBoxTransform: TComboBox;
    ComboBoxCoef: TComboBox;
    ComboBoxMethod: TComboBox;
    LabelTransform: TLabel;
    LabelCoef: TLabel;
    LabelMaxPar: TLabel;
    LabelMethod: TLabel;
    SpinEditMaxPar: TSpinEdit;
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
  private

  public
    NParms: integer;
    procedure CreateBIOENV(fname: string;
      nvars, transf, coef, upto, method, metric: integer; selected: string);
    procedure BIOENV(fname: string; transf, coef, upto, method, metric: integer;
      n, m: integer);
  end;

var
  BIOENVDlg: TBIOENVDlg;

implementation

{$R *.lfm}

uses report, useful;

resourcestring
  strBIOENV = 'COMBINAÇÃO DE VARIÁVEIS';
  strMethod = 'Método: ';
  strMetric = 'Métrica: ';
  strTransform = 'Transformação: ';
  strTransformNone = 'Sem Transformação';
  strTransformCommonLog = 'Logaritmo comum (base 10)';
  strTransformNaturalLog = 'Logaritmo natural (base e)';
  strTransformSqrt = 'Raiz quadrada';
  strTransformArcsin = 'Arcosseno';
  strCoef = 'Coeficiente: ';
  strCoefBray = 'Bray-Curtis';
  strCoefCanberra = 'Canberra';
  strCoefManhattan = 'Manhattan';
  strCoefEuclidean = 'Euclidiana';
  strCoefAvgEuclidean = 'Euclidiana normalizada';
  strCoefSqrEuclidean = 'Euclidiana quadrada';
  strCoefMorisita = 'Morisita-Horn';
  strMethodPearson = 'Pearson';
  strMethodKendall = 'Kendall';
  strMethodSpearman = 'Spearman';
  strMetricEuclidean = 'Euclidiana';
  strMetricMahalanobis = 'Mahalanobis';
  strMetricManhattan = 'Manhattan';
  strMetricGower = 'Gower';
  strMaxPar = 'Número Máximo de Parâmetros: ';
  strSamples = 'Amostras';
  strSpecies = 'Espécies';

{ TBIOENVDlg }

procedure TBIOENVDlg.FormCreate(Sender: TObject);
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
  with ComboBoxCoef.Items do
  begin
    Add(strCoefBray);
    Add(strCoefCanberra);
    Add(strCoefManhattan);
    Add(strCoefEuclidean);
    Add(strCoefAvgEuclidean);
    Add(strCoefSqrEuclidean);
    Add(strCoefMorisita);
  end;
  ComboBoxCoef.ItemIndex := 0;
  with ComboBoxMethod.Items do
  begin
    Add(strMethodPearson);
    Add(strMethodKendall);
    Add(strMethodSpearman);
  end;
  ComboBoxMethod.ItemIndex := 2;
  with ComboBoxMetric.Items do
  begin
    Add(strMetricEuclidean);
    Add(strMetricMahalanobis);
    Add(strMetricManhattan);
    Add(strMetricGower);
  end;
  ComboBoxMetric.ItemIndex := 0;
end;

procedure TBIOENVDlg.FormShow(Sender: TObject);
begin
  SpinEditMaxPar.Value := NParms; // ncol(env)
  SpinEditMaxPar.MaxValue := NParms;
end;

procedure TBIOENVDlg.CreateBIOENV(fname: string;
  nvars, transf, coef, upto, method, metric: integer; selected: string);
var
  model, stransf, scoef, smeth, smetr: string;
  outfile: TextFile;
begin
  case transf of
    0: stransf := '';
    1: stransf := 'log10';
    2: stransf := 'log';
    3: stransf := 'sqrt';
    4: stransf := 'asin';
  end;

  case coef of
    0: scoef := 'bray';
    1: scoef := 'canberra';
    2: scoef := 'manhattan';
    3: scoef := 'euclidean';
    4: scoef := 'avg.euclidean';
    5: scoef := 'sqr.euclidean';
    6: scoef := 'horn';
  end;

  case method of
    0: smeth := 'pearson';
    1: smeth := 'kendall';
    2: smeth := 'spearman';
  end;

  case metric of
    0: smetr := 'euclidean';
    1: smetr := 'mahalanobis';
    2: smetr := 'manhattan';
    3: smetr := 'gower';
  end;

  if selected.CountChar(',') + 1 <> nvars then
    model := StringReplace(selected, ',', '+', [rfReplaceAll])
  else
    model := '';

  AssignFile(outfile, 'bioenv.R');
  Rewrite(outfile);
  WriteLn(outfile, 'options(warn=-1)');
  WriteLn(outfile, 'options(digits=4)');
  WriteLn(outfile, 'suppressPackageStartupMessages(library(vegan))');
  WriteLn(outfile, 'library(vegan, quietly=TRUE)');
  WriteLn(outfile, 'df.data <- read.csv("rdata1.csv", row.names=1)');
  WriteLn(outfile, 'df.data <- t(df.data)');
  WriteLn(outfile, 'df.vars <- read.csv("rdata2.csv", row.names=1)');
  if Length(stransf) > 0 then
    WriteLn(outfile, 'df.data <- ' + stransf + '(df.data + 1)');
  WriteLn(outfile, 'attach(df.vars)');
  if Length(model) = 0 then
    WriteLn(outfile, 'sol <- bioenv(df.data, df.vars, method="' +
      smeth + '", index="' + scoef + '", upto=' + IntToStr(upto) +
      ', metric="' + smetr + '")')
  else
    WriteLn(outfile, 'sol <- bioenv(df.data~' + model + ', df.vars, method="' +
      smeth + '", index="' + scoef + '", upto=' + IntToStr(upto) +
      ', metric="' + smetr + '")');
  WriteLn(outfile, 'sink("bioenv.txt")');
  WriteLn(outfile, 'print(summary(sol))');
  WriteLn(outfile, 'sink()');
  WriteLn(outfile, 'options(warn=0)');
  CloseFile(outfile);
end;

procedure TBIOENVDlg.BIOENV(fname: string; transf, coef, upto, method, metric: integer;
  n, m: integer);
var
  line: string;
  infile, outfile: TextFile;
begin
  AssignFile(outfile, fname);
  Rewrite(outfile);
  Header(outfile, strBIOENV);
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

  if coef = 0 then
    WriteLn(outfile, strCoef + strCoefBray + '<br><br>')
  else if coef = 1 then
    WriteLn(outfile, strCoef + strCoefCanberra + '<br><br>')
  else if coef = 2 then
    WriteLn(outfile, strCoef + strCoefManhattan + '<br><br>')
  else if coef = 3 then
    WriteLn(outfile, strCoef + strCoefEuclidean + '<br><br>')
  else if coef = 4 then
    WriteLn(outfile, strCoef + strCoefMorisita + '<br><br>');

  if method = 0 then
    WriteLn(outfile, strMethod + strMethodPearson + '<br><br>')
  else if method = 1 then
    WriteLn(outfile, strMethod + strMethodKendall + '<br><br>')
  else if method = 2 then
    WriteLn(outfile, strMethod + strMethodSpearman + '<br><br>');

  if metric = 0 then
    WriteLn(outfile, strMetric + strMetricEuclidean + '<br><br>')
  else if metric = 1 then
    WriteLn(outfile, strMetric + strMetricMahalanobis + '<br><br>')
  else if metric = 2 then
    WriteLn(outfile, strMetric + strMetricManhattan + '<br><br>')
  else if metric = 3 then
    WriteLn(outfile, strMetric + strMetricGower + '<br><br>');

  Write(outfile, strMaxPar + IntToStr(upto) + '<br><br>');

  WriteLn(outfile, '<pre>');
  AssignFile(infile, 'bioenv.txt');
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
  if FileExists('bioenv.txt') then
    DeleteFile('bioenv.txt');
end;

end.

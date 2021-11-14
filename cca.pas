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
  strEigenval = 'AUTOVALOR';
  strEigenvals = 'AUTOVALORES';
  strPercentVar = '%VARIÂNCIA';
  strCumVar = '%CUMULATIVA';
  strAxis = 'EIXO ';
  strSample = 'AMOSTRA';
  strVariable = 'VARIÁVEL';
  strSampleScores = 'ESCORES DAS AMOSTRAS';
  strSpeciesScores = 'ESCORES DAS VARIÁVEIS';
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
  WriteLn(outfile, 'res <- summary(cc)');
  WriteLn(outfile, 'loadings <- res$cont$importance');
  WriteLn(outfile,
    'write.table(loadings, "loadings.csv", sep=" ", row.names=FALSE, col.names=FALSE)');
  WriteLn(outfile,
    'write.table(res$species, "rows.csv", sep=" ", row.names=FALSE, col.names=FALSE)');
  WriteLn(outfile,
    'write.table(res$sites, "cols.csv", sep=" ", row.names=FALSE, col.names=FALSE)');
  WriteLn(outfile, 'ppi <- 100');
  figf := GetFileNameWithoutExt(fname) + '.png';
  WriteLn(outfile, 'png("' + figf + '", width=6*ppi, height=6*ppi, res=ppi)');
  WriteLn(outfile, 'par(mar=c(4,4,4,4))');
  WriteLn(outfile,
    'plot(cc, main="", xlab="' + strAxis + '1", ylab="' + strAxis +
    '2", col="darkgreen")');
  WriteLn(outfile, 'invisible(dev.off())');
  WriteLn(outfile, 'options(warn=0)');
  CloseFile(outfile);
end;

procedure TCCADlg.CCA(fname: string; transf: integer; n, m: integer; selected: string);
const
  nvect = 2;
var
  eig_val, sumvariance, cumvariance: array of double;
  row_scores, col_scores: array of array of double;
  x1, x2: double;
  i, j, k: integer;
  figf: string;
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

  WriteLn(outfile, strEigenvals + '<br>');
  WriteLn(outfile, '<table border=1 cellspacing=1 cellpadding=1 width="100%">');
  WriteLn(outfile, '<tr><th>i</th>');
  WriteLn(outfile, '<th>' + strEigenval + '</th>');
  WriteLn(outfile, '<th>' + strPercentVar + '</th>');
  WriteLn(outfile, '<th>' + strCumVar + '</th>');
  WriteLn(outfile, '</tr>');

  AssignFile(infile, 'loadings.csv');
  Reset(infile);
  SetLength(eig_val, 2);
  SetLength(sumvariance, 2);
  SetLength(cumvariance, 2);
  while not EOF(infile) do
  begin
    ReadLn(infile, x1, x2);
    eig_val[0] := x1;
    eig_val[1] := x2;
    ReadLn(infile, x1, x2);
    sumvariance[0] := x1;
    sumvariance[1] := x2;
    ReadLn(infile, x1, x2);
    cumvariance[0] := x1;
    cumvariance[1] := x2;
  end;
  CloseFile(infile);

  for k := 0 to k - 1 do
  begin
    if eig_val[k] < 0.0001 then
      break;
    WriteLn(outfile, '<tr><td align="Center">' + IntToStr(k + 1) + '</td>');
    WriteLn(outfile, '<td align="Center">' + FloatToStrF(eig_val[k],
      ffFixed, 5, 3) + '</td>');
    WriteLn(outfile, '<td align="Center">' + FloatToStrF(sumvariance[k],
      ffFixed, 5, 2) + '</td>');
    WriteLn(outfile, '<td align="Center">' + FloatToStrF(cumvariance[k],
      ffFixed, 5, 2) + '</td>');
    WriteLn(outfile, '</tr>');
  end;
  WriteLn(outfile, '</table><br><br>');

  AssignFile(infile, 'cols.csv');
  Reset(infile);
  k := 0;
  SetLength(col_scores, 1, nvect);
  while not EOF(infile) do
  begin
    ReadLn(infile, x1, x2);
    col_scores[k, 0] := x1;
    col_scores[k, 1] := x2;
    Inc(k);
    SetLength(col_scores, Length(col_scores) + 1, nvect);
  end;
  CloseFile(infile);

  WriteLn(outfile, strSampleScores + '<br>');
  WriteLn(outfile, '<table border=1 cellspacing=1 cellpadding=1 width="100%">');
  WriteLn(outfile, '<tr><th>' + strSample + '</th>');
  for i := 1 to nvect do
    WriteLn(outfile, '<th>' + strAxis + IntToStr(i) + '</th>');
  WriteLn(outfile, '</tr>');
  for i := 1 to k do
  begin
    WriteLn(outfile, '<tr>');
    WriteLn(outfile, '<td align="Center">' + IntToStr(i) + '</td>');
    for j := 1 to nvect do
      WriteLn(outfile, '<td align="Center">' +
        FloatToStrF(col_scores[i - 1][j - 1], ffFixed, 5, 3) + '</td>');
    WriteLn(outfile, '</tr>');
  end;
  WriteLn(outfile, '</table><br><br>');
  WriteLn(outfile);

  AssignFile(infile, 'rows.csv');
  Reset(infile);
  k := 0;
  SetLength(row_scores, 1, nvect);
  while not EOF(infile) do
  begin
    ReadLn(infile, x1, x2);
    row_scores[k, 0] := x1;
    row_scores[k, 1] := x2;
    Inc(k);
    SetLength(row_scores, Length(row_scores) + 1, nvect);
  end;
  CloseFile(infile);

  WriteLn(outfile, strSpeciesScores + '<br>');
  WriteLn(outfile,
    '<table border=1 cellspacing=1 cellpadding=1 width="100%">');
  WriteLn(outfile, '<tr><th>' + strVariable + '</th>');
  for i := 1 to nvect do
    WriteLn(outfile, '<th>' + strAxis + IntToStr(i) + '</th>');
  WriteLn(outfile, '</tr>');
  for i := 1 to k do
  begin
    WriteLn(outfile, '<tr>');
    WriteLn(outfile, '<td align="Center">' + IntToStr(i) + '</td>');
    for j := 1 to nvect do
      WriteLn(outfile, '<td align="Center">' +
        FloatToStrF(row_scores[i - 1][j - 1], ffFixed, 5, 3) + '</td>');
    WriteLn(outfile, '</tr>');
  end;
  WriteLn(outfile, '</table><br><br>');
  WriteLn(outfile);

  figf := GetFileNameWithoutExt(fname) + '.png';
  WriteLn(outfile, '<p align="center"><img src="' + figf + '"></p>');

  WriteLn(outfile, '</body>');
  WriteLn(outfile, '</html>');
  CloseFile(outfile);
  if FileExists('loadings.csv') then
    DeleteFile('loadings.csv');
  if FileExists('rows.csv') then
    DeleteFile('rows.csv');
  if FileExists('cols.csv') then
    DeleteFile('cols.csv');
end;

end.





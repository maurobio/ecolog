unit ca;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, StdCtrls;

type

  { TCADlg }

  TCADlg = class(TForm)
    OKButton: TButton;
    CancelButton: TButton;
    ComboBoxTransform: TComboBox;
    ComboBoxScale: TComboBox;
    LabelTransform: TLabel;
    LabelScale: TLabel;
    procedure FormCreate(Sender: TObject);
  private

  public
    procedure CreateCA(fname: string; transf, scale: integer);
    procedure CA(fname: string; transf: integer; n, m: integer);
  end;

var
  CADlg: TCADlg;

implementation

{$R *.lfm}

uses report, useful;

resourcestring
  strCA = 'ANÁLISE DE CORRESPONDÊNCIAS';
  strEigenvals = 'AUTOVALORES';
  strEigenval = 'AUTOVALOR';
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

{ TCADlg }

procedure TCADlg.FormCreate(Sender: TObject);
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
end;

procedure TCADlg.CreateCA(fname: string; transf, scale: integer);
var
  stransf, figf: string;
  outfile: TextFile;
begin
  case transf of
    0: stransf := '';
    1: stransf := 'log10';
    2: stransf := 'log';
    3: stransf := 'sqrt';
    4: stransf := 'asin';
  end;

  AssignFile(outfile, 'ca.R');
  Rewrite(outfile);
  WriteLn(outfile, 'options(warn=-1)');
  WriteLn(outfile, 'suppressPackageStartupMessages(library(vegan))');
  WriteLn(outfile, 'suppressPackageStartupMessages(library(ca))');
  WriteLn(outfile, 'library(vegan, quietly=TRUE)');
  WriteLn(outfile, 'library(ca, quietly=TRUE)');
  WriteLn(outfile, 'df.data <- read.csv("rdata.csv", row.names=1)');
  WriteLn(outfile, 'df.data <- t(df.data)');
  if Length(stransf) > 0 then
    WriteLn(outfile, 'df.data <- ' + stransf + '(df.data + 1)');
  WriteLn(outfile, 'corresp <- ca(df.data, nd=2)');
  WriteLn(outfile, 'pct <- round(100*(corresp$sv^2)/sum(corresp$sv^2), 1)');
  WriteLn(outfile,
    'write.table(data.frame(corresp$sv, pct, cumsum(pct)), "ca.csv", sep=" ", row.names=FALSE, col.names=FALSE)');
  WriteLn(outfile,
    'write.table(corresp$rowcoord, "rows.csv", sep=" ", row.names=FALSE, col.names=FALSE)');
  WriteLn(outfile,
    'write.table(corresp$colcoord, "cols.csv", sep=" ", row.names=FALSE, col.names=FALSE)');
  WriteLn(outfile, 'ppi <- 100');
  figf := GetFileNameWithoutExt(fname) + '.png';
  WriteLn(outfile, 'png("' + figf + '", width=6*ppi, height=6*ppi, res=ppi)');
  WriteLn(outfile, 'par(mar=c(4,4,4,4))');
  if scale = 0 then
    WriteLn(outfile,
      'plot(corresp, main="", xlab=paste("' + strAxis +
      '1 (", pct[1], "%)", sep=""), ylab=paste("' + strAxis +
      '2 (", pct[2], "%)", sep=""), map="rowprincipal", col="darkgreen", arrows=c(TRUE, FALSE))')
  else if scale = 1 then
    WriteLn(outfile,
      'plot(corresp, main="", xlab=paste("' + strAxis +
      '1 (", pct[1], "%)", sep=""), ylab=paste("' + strAxis +
      '2 (", pct[2], "%)", sep=""), map="colprincipal", col="darkgreen", arrows=c(FALSE, TRUE))')
  else if scale = 2 then
    WriteLn(outfile,
      'plot(corresp, main="", xlab=paste("' + strAxis +
      '1 (", pct[1], "%)", sep=""), ylab=paste("' + strAxis +
      '2 (", pct[2], "%)", sep=""), map="symmetric", col="darkgreen", arrows=c(TRUE, TRUE))');
  WriteLn(outfile, 'invisible(dev.off())');
  WriteLn(outfile, 'options(warn=0)');
  CloseFile(outfile);
end;

procedure TCADlg.CA(fname: string; transf: integer; n, m: integer);
const
  nvect = 2;
var
  eig_val, sumvariance, cumvariance: array of double;
  row_scores, col_scores: array of array of double;
  i, j, k: integer;
  x1, x2, x3: double;
  figf: string;
  infile, outfile: TextFile;
begin
  AssignFile(outfile, fname);
  Rewrite(outfile);
  Header(outfile, strCA);
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

  WriteLn(outfile, strEigenvals + '<br>');
  WriteLn(outfile, '<table border=1 cellspacing=1 cellpadding=1 width="100%">');
  WriteLn(outfile, '<tr><th>i</th>');
  WriteLn(outfile, '<th>' + strEigenval + '</th>');
  WriteLn(outfile, '<th>' + strPercentVar + '</th>');
  WriteLn(outfile, '<th>' + strCumVar + '</th>');
  WriteLn(outfile, '</tr>');

  AssignFile(infile, 'ca.csv');
  Reset(infile);
  k := 0;
  SetLength(eig_val, 1);
  SetLength(sumvariance, 1);
  SetLength(cumvariance, 1);
  while not EOF(infile) do
  begin
    ReadLn(infile, x1, x2, x3);
    eig_val[k] := x1;
    sumvariance[k] := x2;
    cumvariance[k] := x3;
    Inc(k);
    SetLength(eig_val, Length(eig_val) + 1);
    SetLength(sumvariance, Length(sumvariance) + 1);
    SetLength(cumvariance, Length(cumvariance) + 1);
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

  WriteLn(outfile, strSampleScores + '<br>');
  WriteLn(outfile, '<table border=1 cellspacing=1 cellpadding=1 width="100%">');
  WriteLn(outfile, '<tr><th>' + strSample + '</th>');
  for i := 1 to nvect do
    WriteLn(outfile, '<th>' + strAxis + IntToStr(i) + '</th>');
  WriteLn(outfile, '</tr>');
  for i := 1 to n do
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

  WriteLn(outfile, strSpeciesScores + '<br>');
  WriteLn(outfile,
    '<table border=1 cellspacing=1 cellpadding=1 width="100%">');
  WriteLn(outfile, '<tr><th>' + strVariable + '</th>');
  for i := 1 to nvect do
    WriteLn(outfile, '<th>' + strAxis + IntToStr(i) + '</th>');
  WriteLn(outfile, '</tr>');
  for i := 1 to m do
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

  figf := GetFileNameWithoutExt(fname) + '.png';
  WriteLn(outfile, '<p align="center"><img src="' + figf + '"></p>');

  WriteLn(outfile, '</body>');
  WriteLn(outfile, '</html>');
  CloseFile(outfile);
  if FileExists('ca.csv') then
    DeleteFile('ca.csv');
  if FileExists('rows.csv') then
    DeleteFile('rows.csv');
  if FileExists('cols.csv') then
    DeleteFile('cols.csv');
end;

end.

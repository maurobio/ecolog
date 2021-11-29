unit kmeans;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, StdCtrls, Spin;

type

  { TKMeansDlg }

  TKMeansDlg = class(TForm)
    ComboBoxCoef: TComboBox;
    ComboBoxTransform: TComboBox;
    LabelCoef: TLabel;
    LabelTransform: TLabel;
    OKButton: TButton;
    CancelButton: TButton;
    LabelClusters: TLabel;
    LabelIterMax: TLabel;
    LabelNStart: TLabel;
    SpinEditClusters: TSpinEdit;
    SpinEditIterMax: TSpinEdit;
    SpinEditNStart: TSpinEdit;
    procedure FormCreate(Sender: TObject);
  private

  public
    procedure CreateKMeans(fname: string;
      transf, coef, centers, iter_max, nstart: integer);
    procedure KMeans(fname: string; transf, coef, centers, iter_max, nstart: integer;
      n, m: integer);
  end;

var
  KMeansDlg: TKMeansDlg;

implementation

{$R *.lfm}

uses report, useful;

resourcestring
  strKMeans = 'K-MÉDIAS';
  strCompactness = 'Compactness = ';
  strSamples = 'Amostras';
  strSpecies = 'Espécies';
  strDistance = 'DISTÂNCIA';
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
  strConfig = 'Configuração Inicial: ';
  strClusters = 'Número de Agrupamentos: ';
  strIter = 'Número de Iterações: ';
  strNStart = 'Número Inicial de Partições: ';

procedure TKMeansDlg.CreateKMeans(fname: string;
  transf, coef, centers, iter_max, nstart: integer);
var
  stransf, scoef, figf: string;
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

  AssignFile(outfile, 'kmeans.R');
  Rewrite(outfile);
  WriteLn(outfile, 'options(warn=-1)');
  WriteLn(outfile, 'options(digits=4)');
  WriteLn(outfile, 'suppressPackageStartupMessages(library(vegan))');
  WriteLn(outfile, 'library(vegan, quietly=TRUE)');
  WriteLn(outfile, 'df.data <- read.csv("rdata.csv", row.names=1)');
  WriteLn(outfile, 'df.data <- t(df.data)');
  if Length(stransf) > 0 then
    WriteLn(outfile, 'df.data <- ' + stransf + '(df.data + 1)');
  if coef in [0..3] then
    WriteLn(outfile, 'df.dist <- vegdist(df.data, method="' + scoef +
      '", binary=FALSE, diag=TRUE)')
  else if coef = 4 then
    WriteLn(outfile, 'df.dist <- dist(df.data, diag=TRUE)/nrow(df.data)')
  else if coef = 5 then
    WriteLn(outfile, 'df.dist <- dist(df.data, diag=TRUE)^2')
  else if coef = 6 then
    WriteLn(outfile, 'df.dist <- designdist(df.data, "1-J/sqrt(A*B)")');
  WriteLn(outfile, 'cl <- kmeans(df.dist, centers=', centers,
    ', iter.max=', iter_max, ', nstart=', nstart, ')');
  WriteLn(outfile, 'sink("kmeans.txt")');
  WriteLn(outfile, 'print(cl)');
  WriteLn(outfile, 'sink()');
  WriteLn(outfile, 'ppi <- 100');
  figf := GetFileNameWithoutExt(fname) + '.png';
  WriteLn(outfile, 'png("' + figf + '", width=6*ppi, height=6*ppi, res=ppi)');
  WriteLn(outfile, 'par(mar=c(4,4,4,4))');
  WriteLn(outfile, 'x <- cl$centers[1,]');
  WriteLn(outfile, 'y <- cl$centers[2,]');
  WriteLn(outfile, 'plot(x, y, xlab=iconv("' + strDistance +
    '", from=''UTF-8'', to=''latin1''), ylab=iconv("' + strDistance +
    '", from=''UTF-8'', to=''latin1''), main=paste("' + strCompactness +
    '", round(cl$betweenss/cl$totss*100.0, digits=1), "%", sep=" "), pch=19, col="blue")');
  WriteLn(outfile, 'text(x, y, labels=rownames(df.data), cex=0.5, pos=3)');
  WriteLn(outfile, 'invisible(dev.off())');
  WriteLn(outfile, 'options(warn=0)');
  CloseFile(outfile);
end;

procedure TKMeansDlg.KMeans(fname: string;
  transf, coef, centers, iter_max, nstart: integer; n, m: integer);
var
  line, figf: string;
  infile, outfile: TextFile;
begin
  AssignFile(outfile, fname);
  Rewrite(outfile);
  Header(outfile, strKMeans);
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

  WriteLn(outfile, strConfig + '<br>');
  WriteLn(outfile, '&nbsp;&nbsp;&nbsp;' + strClusters + IntToStr(centers) + '<br>');
  WriteLn(outfile, '&nbsp;&nbsp;&nbsp;' + strIter + IntToStr(iter_max) + '<br>');
  WriteLn(outfile, '&nbsp;&nbsp;&nbsp;' + strNStart + IntToStr(nstart) + '<br><br>');

  WriteLn(outfile, '<pre>');
  AssignFile(infile, 'kmeans.txt');
  Reset(infile);
  while not EOF(infile) do
  begin
    ReadLn(infile, line);
    if line.StartsWith('Available') then
      break;
    WriteLn(outfile, line);
  end;
  CloseFile(infile);
  WriteLn(outfile, '</pre>');

  figf := GetFileNameWithoutExt(fname) + '.png';
  WriteLn(outfile, '<p align="left"><img src="' + figf + '"></p>');

  WriteLn(outfile, '</body>');
  WriteLn(outfile, '</html>');
  CloseFile(outfile);
  if FileExists('kmeans.txt') then
    DeleteFile('kmeans.txt');
end;

{ TKMeansDlg }

procedure TKMeansDlg.FormCreate(Sender: TObject);
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
end;

end.

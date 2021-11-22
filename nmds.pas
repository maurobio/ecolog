unit nmds;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, StdCtrls, Spin;

type

  { TNMDSDlg }

  TNMDSDlg = class(TForm)
    OKButton: TButton;
    CancelButton: TButton;
    ComboBoxTransform: TComboBox;
    ComboBoxCoef: TComboBox;
    ComboBoxConfig: TComboBox;
    LabelTransform: TLabel;
    LabelCoef: TLabel;
    LabelIter: TLabel;
    LabelConfig: TLabel;
    SpinEditIter: TSpinEdit;
    procedure FormCreate(Sender: TObject);
  private

  public
    procedure CreateNMDS(fname: string; transf, coef, iter, config: integer);
    procedure NMDS(fname: string; transf, coef, iter, config: integer; n, m: integer);
  end;

var
  NMDSDlg: TNMDSDlg;

implementation

{$R *.lfm}

uses report, useful;

resourcestring
  strNMDS = 'ESCALONAMENTO MULTIDIMENSIONAL NÃO-MÉTRICO';
  strDim = 'DIMENSÃO ';
  strStress = 'Stress = ';
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
  strConfigPCO = 'Coordenadas Principais';
  strConfigRandom = 'Aleatória';
  strIter = 'Número de Iterações: ';
  strSamples = 'Amostras';
  strSpecies = 'Espécies';

{ TNMDSDlg }

procedure TNMDSDlg.FormCreate(Sender: TObject);
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
  with ComboBoxConfig.Items do
  begin
    Add(strConfigPCO);
    Add(strConfigRandom);
  end;
  ComboBoxConfig.ItemIndex := 0;
end;

procedure TNMDSDlg.CreateNMDS(fname: string; transf, coef, iter, config: integer);
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

  AssignFile(outfile, 'nmds.R');
  Rewrite(outfile);
  WriteLn(outfile, 'options(warn=-1)');
  WriteLn(outfile, 'options(digits=4)');
  WriteLn(outfile, 'suppressPackageStartupMessages(library(vegan))');
  WriteLn(outfile, 'suppressPackageStartupMessages(library(MASS))');
  WriteLn(outfile, 'library(vegan, quietly=TRUE)');
  WriteLn(outfile, 'library(MASS, quietly=TRUE)');
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
  WriteLn(outfile, 'sink("nmds.txt")');
  if config = 0 then
    WriteLn(outfile, 'nmds <- isoMDS(df.dist, k=2, maxit=' + IntToStr(iter) +
      ', trace=TRUE, tol=1e-7)')
  else if config = 1 then
    WriteLn(outfile,
      'nmds <- isoMDS(df.dist, initMDS(df.dist), k=2, maxit=' +
      IntToStr(iter) + ', trace=TRUE, tol=1e-7)');
  WriteLn(outfile, 'cat("\nstress =", nmds$stress)');
  WriteLn(outfile, 'sink()');
  WriteLn(outfile, 'ppi <- 100');
  figf := GetFileNameWithoutExt(fname) + '.png';
  WriteLn(outfile, 'png("' + figf + '", width=6*ppi, height=6*ppi, res=ppi)');
  WriteLn(outfile, 'par(mar=c(4,4,4,4))');
  WriteLn(outfile, 'x <- nmds$points[,1]');
  WriteLn(outfile, 'y <- nmds$points[,2]');
  WriteLn(outfile, 'plot(x, y, xlab=iconv("' + strDim +
    '1", from=''UTF-8'', to=''latin1''), ylab=iconv("' + strDim +
    '2", from=''UTF-8'', to=''latin1''), main=paste("' + strStress +
    '", round(nmds$stress, 7), sep=" "), pch=19, col="blue")');
  WriteLn(outfile, 'text(x, y, labels=rownames(nmds$points), pos=3, cex=0.7)');
  WriteLn(outfile, 'invisible(dev.off())');
  WriteLn(outfile, 'options(warn=0)');
  CloseFile(outfile);
end;

procedure TNMDSDlg.NMDS(fname: string; transf, coef, iter, config: integer;
  n, m: integer);
var
  line, figf: string;
  infile, outfile: TextFile;
begin
  AssignFile(outfile, fname);
  Rewrite(outfile);
  Header(outfile, strNMDS);
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

  Write(outfile, strConfig);
  if config = 0 then
    WriteLn(outfile, strConfigPCO + '<br><br>')
  else if config = 1 then
    WriteLn(outfile, strConfigRandom + '<br><br>');
  Write(outfile, strIter + IntToStr(iter) + '<br><br>');

  WriteLn(outfile, '<pre>');
  AssignFile(infile, 'nmds.txt');
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
  if FileExists('nmds.txt') then
    DeleteFile('nmds.txt');
end;

end.

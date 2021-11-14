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
  //strScatterplot = 'DIAGRAMA DE DISPERSÃO';
  strDim = 'DIMENSÃO ';
  strSample = 'AMOSTRA';
  strCoords = 'COORDENADAS';
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
  if config = 0 then
    WriteLn(outfile, 'nmds <- isoMDS(df.dist, k=2, maxit=' + IntToStr(iter) +
      ', trace=FALSE, tol=1e-7)')
  else if config = 1 then
    WriteLn(outfile,
      'nmds <- isoMDS(df.dist, initMDS(df.dist), k=2, maxit=' +
      IntToStr(iter) + ', trace=FALSE, tol=1e-7)');
  WriteLn(outfile,
    'write.table(data.frame(nmds), "nmds.csv", sep=" ", row.names=FALSE, col.names=FALSE)');
  WriteLn(outfile, 'cat(as.character(nmds$stress), file="stress.txt")');
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
const
  ndim = 2; //3;
var
  pts: array of array of double;
  i, j, k: integer;
  x1, x2, {x3,} stress: double;
  figf: string;
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

  AssignFile(infile, 'nmds.csv');
  Reset(infile);
  k := 0;
  SetLength(pts, 1, ndim);
  while not EOF(infile) do
  begin
    //ReadLn(infile, x1, x2, x3);
    ReadLn(infile, x1, x2);
    pts[k, 0] := x1;
    pts[k, 1] := x2;
    //pts[k, 2] := x3;
    Inc(k);
    SetLength(pts, Length(pts) + 1, ndim);
  end;
  CloseFile(infile);

  AssignFile(infile, 'stress.txt');
  Reset(infile);
  ReadLn(infile, stress);
  CloseFile(infile);
  if stress < 0.0001 then
    stress := 0.00;

  WriteLn(outfile, strStress + FloatToStrF(stress, ffFixed, 9, 7) + '<br><br>');
  WriteLn(outfile, strCoords + '<br>');
  WriteLn(outfile, '<table border=1 cellspacing=1 cellpadding=1 width="100%">');
  WriteLn(outfile, '<tr><th>' + strSample + '</th>');
  for i := 1 to ndim do
    WriteLn(outfile, '<th>' + strDim + IntToStr(i) + '</th>');
  WriteLn(outfile, '</tr>');
  for i := 0 to k - 1 do
  begin
    WriteLn(outfile, '<tr>');
    WriteLn(outfile, '<td align="Center">' + IntToStr(i + 1) + '</td>');
    for j := 0 to ndim - 1 do
      WriteLn(outfile, '<td align="Center">' +
        FloatToStrF(pts[i, j], ffFixed, 5, 3) + '</td>');
    WriteLn(outfile, '</tr>');
  end;
  WriteLn(outfile, '</table><br><br>');

  figf := GetFileNameWithoutExt(fname) + '.png';
  WriteLn(outfile, '<p align="center"><img src="' + figf + '"></p>');

  WriteLn(outfile, '</body>');
  WriteLn(outfile, '</html>');
  CloseFile(outfile);
  if FileExists('nmds.csv') then
    DeleteFile('nmds.csv');
  if FileExists('stress.txt') then
    DeleteFile('stress.txt');
end;

end.

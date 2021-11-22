unit cluster;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, StdCtrls;

type

  { TClusterDlg }

  TClusterDlg = class(TForm)
    ComboBoxCoef: TComboBox;
    ComboBoxTransform: TComboBox;
    LabelCoef: TLabel;
    LabelTransform: TLabel;
    OKButton: TButton;
    CancelButton: TButton;
    ComboBoxMethod: TComboBox;
    LabelMethod: TLabel;
    procedure FormCreate(Sender: TObject);
  private

  public
    procedure CreateCluster(fname: string; transf, coef, method: integer);
    procedure Cluster(fname: string; transf, coef, method: integer; n, m: integer);
  end;

var
  ClusterDlg: TClusterDlg;

implementation

{$R *.lfm}

uses report, useful;

resourcestring
  strClusterAnalysis = 'ANÁLISE DE AGRUPAMENTOS';
  strDendrogram = 'Dendrograma';
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
  strCorrelation = 'Correlação';
  strCoefJaccard = 'Jaccard';
  strCoefDice = 'Dice-Sorenson';
  strCoefKulczynski = 'Kulczynski';
  strCoefOchiai = 'Ochiai';
  strMethod = 'Método: ';
  strSLM = 'Ligação única (SLM)';
  strCLM = 'Ligação completa (CLM)';
  strUPGMA = 'Ligação média (UPGMA)';
  strWPGMA = 'Ligação ponderada (WPGMA)';
  strCentroid = 'Centroide (UPGMC)';
  strMedian = 'Mediana (WPGMC)';
  strWard = 'Método de Ward';
  strDistance = 'Distância';
  strSamples = 'Amostras';
  strSpecies = 'Espécies';
//strCorrelation = 'Correlação';
//strSimilarity = 'Similaridade';

{ TClusterDlg }

procedure TClusterDlg.CreateCluster(fname: string; transf, coef, method: integer);
var
  smethod, stransf, scoef, figf, acronym: string;
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
    7: scoef := 'cor';
    8: scoef := 'jaccard';
    9: scoef := 'sorensen';
   10: scoef := 'kulczynski';
   11: scoef := 'ochiai';
  end;

  case method of
    0: smethod := 'single';
    1: smethod := 'complete';
    2: smethod := 'average';
    3: smethod := 'mcquitty';
    4: smethod := 'centroid';
    5: smethod := 'median';
    6: smethod := 'ward.D2';
  end;

  case method of
    0: acronym := 'SLM';
    1: acronym := 'CLM';
    2: acronym := 'UPGMA';
    3: acronym := 'WPGMA';
    4: acronym := 'UPGMC';
    5: acronym := 'WPGMC';
    6: acronym := 'Ward';
  end;

  AssignFile(outfile, 'cluster.R');
  Rewrite(outfile);
  WriteLn(outfile, 'options(warn=-1)');
  WriteLn(outfile, 'suppressPackageStartupMessages(library(vegan))');
  WriteLn(outfile, 'library(vegan, quietly=TRUE)');
  WriteLn(outfile, 'df.data <- read.csv("rdata.csv", row.names=1)');
  WriteLn(outfile, 'df.data <- t(df.data)');
  if Length(stransf) > 0 then
    WriteLn(outfile, 'df.data <- ' + stransf + '(df.data + 1)');
  if coef in [0..3, 6] then
    WriteLn(outfile, 'df.dist <- vegdist(df.data, method="' + scoef +
      '", binary=FALSE, diag=TRUE)')
  else if coef in [8, 10] then
    WriteLn(outfile, 'df.dist <- vegdist(df.data, method="' + scoef +
      '", binary=TRUE, diag=TRUE)')
  else if coef = 9 then
    WriteLn(outfile, 'df.dist <- designdist(df.data, "(A+B-2*J)/(A+B)")')
  else if coef = 11 then
    WriteLn(outfile, 'df.dist <- designdist(df.data, "1-J/sqrt(A*B)")')
  else if coef = 7 then
    WriteLn(outfile, 'df.dist <- as.dist(cor(t(df.data)))')
  else if coef = 4 then
    WriteLn(outfile, 'df.dist <- dist(df.data, diag=TRUE)/nrow(df.data)')
  else if coef = 5 then
    WriteLn(outfile, 'df.dist <- dist(df.data, diag=TRUE)^2');
  WriteLn(outfile, 'hc <- hclust(df.dist, method="' + smethod + '")');
  WriteLn(outfile, 'dc <- cophenetic(hc)');
  WriteLn(outfile, 'r <- cor(df.dist, dc)');
  WriteLn(outfile, 'ppi <- 100');
  figf := GetFileNameWithoutExt(fname) + '.png';
  WriteLn(outfile, 'png("' + figf + '", width=6*ppi, height=6*ppi, res=ppi)');
  WriteLn(outfile, 'par(mar=c(4,4,4,4))');
  WriteLn(outfile,
    'plot(as.dendrogram(hc), xlab="' + strDistance + '", ylab="' +
    strSamples + '", main=paste("' + strDendrogram +
    ' (' + acronym + ', r = ", format(r, digits=4), ")"), horiz=TRUE, edgePar=list(col="blue", lwd=3))');
  WriteLn(outfile, 'invisible(dev.off())');
  WriteLn(outfile, 'options(warn=0)');
  CloseFile(outfile);
end;

procedure TClusterDlg.Cluster(fname: string; transf, coef, method: integer;
  n, m: integer);
var
  figf: string;
  outfile: TextFile;
begin
  AssignFile(outfile, fname);
  Rewrite(outfile);
  Header(outfile, strClusterAnalysis);
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
    WriteLn(outfile, strCoef + strCoefAvgEuclidean + '<br><br>')
  else if coef = 5 then
    WriteLn(outfile, strCoef + strCoefSqrEuclidean + '<br><br>')
  else if coef = 6 then
    WriteLn(outfile, strCoef + strCoefMorisita + '<br><br>')
  else if coef = 7 then
    WriteLn(outfile, strCoef + strCorrelation + '<br><br>')
  else if coef = 8 then
    WriteLn(outfile, strCoef + strCoefJaccard + '<br><br>')
  else if coef = 9 then
    WriteLn(outfile, strCoef + strCoefDice + '<br><br>')
  else if coef = 10 then
    WriteLn(outfile, strCoef + strCoefKulczynski + '<br><br>')
  else if coef =11 then
    WriteLn(outfile, strCoef + strCoefOchiai + '<br><br>');

  if method = 0 then
    WriteLn(outfile, strMethod + strSLM + '<br><br>')
  else if method = 1 then
    WriteLn(outfile, strMethod + strCLM + '<br><br>')
  else if method = 2 then
    WriteLn(outfile, strMethod + strUPGMA + '<br><br>')
  else if method = 3 then
    WriteLn(outfile, strMethod + strWPGMA + '<br><br>')
  else if method = 4 then
    WriteLn(outfile, strMethod + strCentroid + '<br><br>')
  else if method = 5 then
    WriteLn(outfile, strMethod + strMedian + '<br><br>')
  else if method = 6 then
    WriteLn(outfile, strMethod + strMethod + strWard + '<br><br>');

  figf := GetFileNameWithoutExt(fname) + '.png';
  WriteLn(outfile, '<p align="left"><img src="' + figf + '"></p>');

  WriteLn(outfile, '</body>');
  WriteLn(outfile, '</html>');
  CloseFile(outfile);
end;

procedure TClusterDlg.FormCreate(Sender: TObject);
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
    Add(strCorrelation);
    Add(strCoefJaccard);
    Add(strCoefDice);
    Add(strCoefKulczynski);
    Add(strCoefOchiai);
  end;
  ComboBoxCoef.ItemIndex := 0;
  with ComboBoxMethod.Items do
  begin
    Add(strSLM);
    Add(strCLM);
    Add(strUPGMA);
    Add(strWPGMA);
    Add(strCentroid);
    Add(strMedian);
    Add(strWard);
  end;
  ComboBoxMethod.ItemIndex := 2;
end;

end.

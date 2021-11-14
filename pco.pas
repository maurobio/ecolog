unit pco;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, StdCtrls;

type

  { TPCODlg }

  TPCODlg = class(TForm)
    OKButton: TButton;
    CancelButton: TButton;
    ComboBoxTransform: TComboBox;
    ComboBoxCoef: TComboBox;
    LabelTransform: TLabel;
    LabelCoef: TLabel;
    procedure FormCreate(Sender: TObject);
  private

  public
    procedure CreatePCOA(fname: string; transf, coef: integer);
    procedure PCOA(fname: string; transf, coef: integer; n, m: integer);
  end;

var
  PCODlg: TPCODlg;

implementation

{$R *.lfm}

uses report, useful;

resourcestring
  strPCO = 'ANÁLISE DE COORDENADAS PRINCIPAIS';
  strEigenvals = 'AUTOVALORES';
  strEigenval = 'AUTOVALOR';
  strEigenvecs = 'AUTOVETORES';
  strEigenvec = 'AUTOVETOR';
  strPercentVar = '%VARIÂNCIA';
  strCumVar = '%CUMULATIVA';
  strVariable = 'VARIÁVEL';
  strScatterplot = 'DIAGRAMA DE DISPERSÃO';
  strAxis = 'EIXO ';
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
  strSamples = 'Amostras';
  strSpecies = 'Espécies';

{ TPCODlg }

procedure TPCODlg.FormCreate(Sender: TObject);
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

procedure TPCODlg.CreatePCOA(fname: string; transf, coef: integer);
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

  AssignFile(outfile, 'pco.R');
  Rewrite(outfile);
  WriteLn(outfile, 'options(warn=-1)');
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
  WriteLn(outfile, 'pcovec <- cmdscale(df.dist, k=2, eig=FALSE)');
  WriteLn(outfile,
    'write.table(pcovec, "vectors.csv", sep=" ", row.names=FALSE, col.names=FALSE)');
  WriteLn(outfile, 'pco <- cmdscale(df.dist, k=2, eig=TRUE)');
  WriteLn(outfile, 'pcovar <- pco$eig / sum(pco$eig[pco$eig > 0])');
  WriteLn(outfile,
    'write.table(data.frame(pco$eig, pcovar*100, cumsum(pcovar)*100), "pco.csv", sep=" ", row.names=FALSE, col.names=FALSE)');
  WriteLn(outfile, 'ppi <- 100');
  figf := GetFileNameWithoutExt(fname) + '.png';
  WriteLn(outfile, 'png("' + figf + '", width=6*ppi, height=6*ppi, res=ppi)');
  WriteLn(outfile, 'par(mar=c(4,4,4,4))');
  WriteLn(outfile,
    'plot(pco$points, main=iconv("' + strScatterplot +
    '", from=''UTF-8'', to=''latin1''), xlab=paste("' + strAxis +
    '1 (", round(pcovar[1], 3)*100, "%)", sep=""), ylab=paste("' +
    strAxis + '2 (", round(pcovar[2], 3)*100, "%)", sep=""), pch=19, col="blue")');
  WriteLn(outfile, 'text(pco$points[,1:2], labels=rownames(pco$points), pos=3, cex=0.7)');
  WriteLn(outfile, 'invisible(dev.off())');
  WriteLn(outfile, 'options(warn=0)');
  CloseFile(outfile);
end;

procedure TPCODlg.PCOA(fname: string; transf, coef: integer; n, m: integer);
const
  nvec = 3;
var
  eig_val, sumvariance, cumvariance: array of double;
  eig_vec: array of array of double;
  i, j, k: integer;
  x1, x2, x3: double;
  figf: string;
  infile, outfile: TextFile;
begin
  AssignFile(outfile, fname);
  Rewrite(outfile);
  Header(outfile, strPCO);
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

  WriteLn(outfile, strEigenvals + '<br>');
  WriteLn(outfile, '<table border=1 cellspacing=1 cellpadding=1 width="100%">');
  WriteLn(outfile, '<tr><th>i</th>');
  WriteLn(outfile, '<th>' + strEigenval + '</th>');
  WriteLn(outfile, '<th>' + strPercentVar + '</th>');
  WriteLn(outfile, '<th>' + strCumVar + '</th>');
  WriteLn(outfile, '</tr>');

  AssignFile(infile, 'pco.csv');
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

  AssignFile(infile, 'vectors.csv');
  Reset(infile);
  k := 0;
  SetLength(eig_vec, 1, nvec);
  while not EOF(infile) do
  begin
    ReadLn(infile, x1, x2, x3);
    eig_vec[k, 0] := x1;
    eig_vec[k, 1] := x2;
    eig_vec[k, 2] := x3;
    Inc(k);
    SetLength(eig_vec, Length(eig_vec) + 1, nvec);
  end;
  CloseFile(infile);

  WriteLn(outfile, strEigenvecs + '<br>');
  WriteLn(outfile, '<table border=1 cellspacing=1 cellpadding=1 width="100%">');
  WriteLn(outfile, '<tr><th>' + strVariable + '</th>');
  for i := 1 to nvec do
    WriteLn(outfile, '<th>' + strAxis + IntToStr(i) + '</th>');
  WriteLn(outfile, '</tr>');
  for i := 0 to k - 1 do
  begin
    WriteLn(outfile, '<tr>');
    WriteLn(outfile, '<td align="Center">' + IntToStr(i + 1) + '</td>');
    for j := 0 to nvec - 1 do
      WriteLn(outfile, '<td align="Center">' +
        FloatToStrF(eig_vec[i, j], ffFixed, 5, 3) + '</td>');
    WriteLn(outfile, '</tr>');
  end;
  WriteLn(outfile, '</table>');

  figf := GetFileNameWithoutExt(fname) + '.png';
  WriteLn(outfile, '<p align="center"><img src="' + figf + '"></p>');

  WriteLn(outfile, '</body>');
  WriteLn(outfile, '</html>');
  CloseFile(outfile);
  if FileExists('pco.csv') then
    DeleteFile('pco.csv');
  if FileExists('vectors.csv') then
    DeleteFile('vectors.csv');
end;

end.

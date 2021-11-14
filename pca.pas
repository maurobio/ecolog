unit pca;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, StdCtrls, StrUtils;

type

  { TPCADlg }

  TPCADlg = class(TForm)
    OKButton: TButton;
    CancelButton: TButton;
    CheckBoxCenter: TCheckBox;
    CheckBoxStand: TCheckBox;
    ComboBoxTransform: TComboBox;
    ComboBoxCoef: TComboBox;
    LabelTransform: TLabel;
    LabelCoef: TLabel;
    procedure ComboBoxCoefChange(Sender: TObject);
    procedure FormCreate(Sender: TObject);
  private

  public
    procedure CreatePCA(fname: string; transf, coef: integer; center, scale: boolean);
    procedure PCA(fname: string; transf, coef: integer; center, scale: boolean;
      n, m: integer);
  end;

var
  PCADlg: TPCADlg;

implementation

{$R *.lfm}

uses report, useful;

resourcestring
  strPCA = 'ANÁLISE DE COMPONENTES PRINCIPAIS';
  strEigenvals = 'AUTOVALORES';
  strEigenval = 'AUTOVALOR';
  strEigenvecs = 'AUTOVETORES';
  strEigenvec = 'AUTOVETOR';
  strPercentVar = '%VARIÂNCIA';
  strCumVar = '%CUMULATIVA';
  strVariable = 'VARIÁVEL';
  strVariance = 'VARIÂNCIA';
  strScatterplot = 'DIAGRAMA DE DISPERSÃO';
  strScreeplot = 'GRÁFICO DE DECLIVIDADE';
  strBiplot = 'BIPLOT';
  strAxis = 'EIXO ';
  strTransform = 'Transformação: ';
  strTransformNone = 'Sem Transformação';
  strTransformCommonLog = 'Logaritmo comum (base 10)';
  strTransformNaturalLog = 'Logaritmo natural (base e)';
  strTransformSqrt = 'Raiz quadrada';
  strTransformArcsin = 'Arcosseno';
  strCoef = 'Coeficiente: ';
  strCovariance = 'Covariância';
  strCorrelation = 'Correlação';
  strSamples = 'Amostras';
  strSpecies = 'Espécies';
  strCenter = 'Dados Centrados';
  strStand = 'Dados Estandardizados';

{ TPCADlg }

procedure TPCADlg.FormCreate(Sender: TObject);
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
    Add(strCovariance);
    Add(strCorrelation);
  end;
  ComboBoxCoef.ItemIndex := 1;
  CheckBoxCenter.Checked := True;
  CheckBoxStand.Checked := True;
end;

procedure TPCADlg.ComboBoxCoefChange(Sender: TObject);
begin
  if ComboBoxCoef.ItemIndex = 0 then
    CheckBoxStand.Checked := False
  else if ComboBoxCoef.ItemIndex = 1 then
    CheckBoxStand.Checked := True;
end;

procedure TPCADlg.CreatePCA(fname: string; transf, coef: integer;
  center, scale: boolean);
var
  stransf, scoef, figf1, figf2, figf3: string;
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
    0: scoef := 'cor';
    1: scoef := 'varcov';
  end;

  AssignFile(outfile, 'pca.R');
  Rewrite(outfile);
  WriteLn(outfile, 'options(warn=-1)');
  WriteLn(outfile, 'df.data <- read.csv("rdata.csv", row.names=1)');
  WriteLn(outfile, 'df.data <- t(df.data)');
  if Length(stransf) > 0 then
    WriteLn(outfile, 'df.data <- ' + stransf + '(df.data + 1)');
  WriteLn(outfile, 'pca <- prcomp(df.data, scale=' + IfThen(scale, 'TRUE', 'FALSE') +
    ', ' + 'center=' + IfThen(center, 'TRUE', 'FALSE') + ')');
  WriteLn(outfile, 'pcavar <- round((pca$sdev^2) / sum((pca$sdev^2)), 3)*100');
  WriteLn(outfile,
    'write.table(data.frame(pca$sdev^2, pcavar, cumsum(pcavar)), "pca.csv", sep=" ", row.names=FALSE, col.names=FALSE)');
  WriteLn(outfile, 'loadings <- pca$rotation');
  WriteLn(outfile,
    'write.table(loadings, "loadings.csv", sep=" ", row.names=FALSE, col.names=FALSE)');
  WriteLn(outfile, 'scores <- pca$x');
  WriteLn(outfile, 'ppi <- 100');
  figf1 := GetFileNameWithoutExt(fname) + '1.png';
  WriteLn(outfile, 'png("' + figf1 + '", width=6*ppi, height=6*ppi, res=ppi)');
  WriteLn(outfile, 'par(mar=c(4,4,4,4))');
  WriteLn(outfile,
    'plot(scores, main=iconv("' + strScatterplot +
    '", from=''UTF-8'', to=''latin1''), xlab=paste("' + strAxis +
    '1 (", round(pcavar[1], 3), "%)", sep=""), ylab=paste("' + strAxis +
    '2 (", round(pcavar[2], 3), "%)", sep=""), pch=19, col="blue")');
  WriteLn(outfile, 'text(scores[,1], scores[,2], labels=rownames(df.data), pos=3, cex=0.7)');
  WriteLn(outfile, 'invisible(dev.off())');
  figf2 := GetFileNameWithoutExt(fname) + '2.png';
  WriteLn(outfile, 'png("' + figf2 + '", width=6*ppi, height=6*ppi, res=ppi)');
  WriteLn(outfile, 'plot(pca$sdev^2, xlab="' +
    AnsiProperCase(strEigenval, StdWordDelims) + '", ylab=iconv("' +
    AnsiProperCase(strVariance, StdWordDelims) +
    '", from=''UTF-8'', to=''latin1''), main=iconv("' + strScreeplot +
    '", from=''UTF-8'', to=''latin1''), pch=19, col="black")');
  WriteLn(outfile, 'lines(pca$sdev^2)');
  WriteLn(outfile, 'invisible(dev.off())');
  figf3 := GetFileNameWithoutExt(fname) + '3.png';
  WriteLn(outfile, 'png("' + figf3 + '", width=6*ppi, height=6*ppi, res=ppi)');
  WriteLn(outfile, 'biplot(pca, scale=0, cex=.7, xlab="' + strAxis +
    '1", ylab="' + strAxis + '2", main="' + strBiplot + '")');
  WriteLn(outfile, 'invisible(dev.off())');
  WriteLn(outfile, 'options(warn=0)');
  CloseFile(outfile);
end;

procedure TPCADlg.PCA(fname: string; transf, coef: integer;
  center, scale: boolean; n, m: integer);
const
  nvec = 3;
var
  eig_val, sumvariance, cumvariance: array of double;
  eig_vec: array of array of double;
  i, j, k: integer;
  x1, x2, x3: double;
  figf1, figf2, figf3: string;
  infile, outfile: TextFile;
begin
  AssignFile(outfile, fname);
  Rewrite(outfile);
  Header(outfile, strPCA);
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
    WriteLn(outfile, strCoef + strCovariance + '<br><br>')
  else if coef = 1 then
    WriteLn(outfile, strCoef + strCorrelation + '<br><br>');

  if center then
    WriteLn(outfile, strCenter + '<br><br>');
  if scale then
    WriteLn(outfile, strStand + '<br><br>');

  WriteLn(outfile, strEigenvals + '<br>');
  WriteLn(outfile, '<table border=1 cellspacing=1 cellpadding=1 width="100%">');
  WriteLn(outfile, '<tr><th>i</th>');
  WriteLn(outfile, '<th>' + strEigenval + '</th>');
  WriteLn(outfile, '<th>' + strPercentVar + '</th>');
  WriteLn(outfile, '<th>' + strCumVar + '</th>');
  WriteLn(outfile, '</tr>');

  AssignFile(infile, 'pca.csv');
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

  AssignFile(infile, 'loadings.csv');
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

  figf1 := GetFileNameWithoutExt(fname) + '1.png';
  WriteLn(outfile, '<p align="center"><img src="' + figf1 + '"></p>');
  figf2 := GetFileNameWithoutExt(fname) + '2.png';
  WriteLn(outfile, '<p align="center"><img src="' + figf2 + '"></p>');
  figf3 := GetFileNameWithoutExt(fname) + '3.png';
  WriteLn(outfile, '<p align="center"><img src="' + figf3 + '"></p>');

  WriteLn(outfile, '</body>');
  WriteLn(outfile, '</html>');
  CloseFile(outfile);
  if FileExists('pca.csv') then
    DeleteFile('pca.csv');
  if FileExists('loadings.csv') then
    DeleteFile('loadings.csv');
end;

end.

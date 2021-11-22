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
  strEigenval = 'AUTOVALOR';
  strVariance = 'VARIÂNCIA';
  strScatterplot = 'DIAGRAMA DE DISPERSÃO';
  strScreeplot = 'GRÁFICO DE DECLIVIDADE';
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
  WriteLn(outfile, 'options(digits=4)');
  WriteLn(outfile, 'df.data <- read.csv("rdata.csv", row.names=1)');
  WriteLn(outfile, 'df.data <- t(df.data)');
  if Length(stransf) > 0 then
    WriteLn(outfile, 'df.data <- ' + stransf + '(df.data + 1)');
  WriteLn(outfile, 'pca <- prcomp(df.data, scale=' + IfThen(scale, 'TRUE', 'FALSE') +
    ', ' + 'center=' + IfThen(center, 'TRUE', 'FALSE') + ')');
  WriteLn(outfile, 'sink("pca.txt")');
  WriteLn(outfile, 'summary(pca)');
  WriteLn(outfile, 'sink()');
  WriteLn(outfile, 'pcavar <- round((pca$sdev^2) / sum((pca$sdev^2)), 3)*100');
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
  WriteLn(outfile, 'options(warn=0)');
  CloseFile(outfile);
end;

procedure TPCADlg.PCA(fname: string; transf, coef: integer;
  center, scale: boolean; n, m: integer);
var
  line, figf1, figf2: string;
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

  WriteLn(outfile, '<pre>');
  AssignFile(infile, 'pca.txt');
  Reset(infile);
  while not EOF(infile) do
  begin
    ReadLn(infile, line);
    WriteLn(outfile, line);
  end;
  CloseFile(infile);
  WriteLn(outfile, '</pre>');

  figf1 := GetFileNameWithoutExt(fname) + '1.png';
  WriteLn(outfile, '<p align="left"><img src="' + figf1 + '"></p>');
  figf2 := GetFileNameWithoutExt(fname) + '2.png';
  WriteLn(outfile, '<p align="left"><img src="' + figf2 + '"></p>');

  WriteLn(outfile, '</body>');
  WriteLn(outfile, '</html>');
  CloseFile(outfile);
  if FileExists('pca.txt') then
    DeleteFile('pca.txt');
end;

end.

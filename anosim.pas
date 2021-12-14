unit anosim;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, StdCtrls, Spin;

type

  { TANOSIMDlg }

  TANOSIMDlg = class(TForm)
    OKButton: TButton;
    CancelButton: TButton;
    ComboBoxTransform: TComboBox;
    ComboBoxCoef: TComboBox;
    ComboBoxGroup: TComboBox;
    LabelTransform: TLabel;
    LabelCoef: TLabel;
    LabelNPerm: TLabel;
    LabelGroup: TLabel;
    SpinEditNPerm: TSpinEdit;
    procedure ComboBoxGroupChange(Sender: TObject);
    procedure FormCreate(Sender: TObject);
  private

  public
    Factors: string;
    procedure CreateANOSIM(fname: string; groupvar: string;
      transf, coef, nperm: integer);
    procedure ANOSIM(fname: string; groupvar: string; transf, coef, nperm: integer;
      n, m: integer);
  end;

var
  ANOSIMDlg: TANOSIMDlg;

implementation

{$R *.lfm}

uses group, report, useful;

resourcestring
  strANOSIM = 'ANÁLISE DE SIMILARIDADE';
  strGroup = 'GRUPOS: ';
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
  strNPerm = 'Número de Permutações: ';
  strSamples = 'Amostras';
  strSpecies = 'Espécies';

{ TANOSIMDlg }

procedure TANOSIMDlg.FormCreate(Sender: TObject);
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
  ComboBoxGroup.ItemIndex := 0;
end;

procedure TANOSIMDlg.ComboBoxGroupChange(Sender: TObject);
begin
  GroupDlg.GroupVariable := ComboBoxGroup.Items[ComboBoxGroup.ItemIndex];
  if GroupDlg.ShowModal = mrOk then
    Factors := GroupDlg.Factors;
end;

procedure TANOSIMDlg.CreateANOSIM(fname: string; groupvar: string;
  transf, coef, nperm: integer);
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

  AssignFile(outfile, 'anosim.R');
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
  WriteLn(outfile, 'groups <- as.factor(c(', groupvar, '))');
  WriteLn(outfile, 'ano <- anosim(df.data, groups, permutations=', nperm, ')');
  WriteLn(outfile, 'sink("anosim.txt")');
  WriteLn(outfile, 'print(summary(ano))');
  WriteLn(outfile, 'sink()');
  WriteLn(outfile, 'ppi <- 100');
  figf := GetFileNameWithoutExt(fname) + '.png';
  WriteLn(outfile, 'png("' + figf + '", width=6*ppi, height=6*ppi, res=ppi)');
  WriteLn(outfile, 'par(mar=c(4,4,4,4))');
  WriteLn(outfile, 'if (ano$permutations) {');
  WriteLn(outfile, '  pval <- format.pval(ano$signif)');
  WriteLn(outfile, '} else {');
  WriteLn(outfile, '  pval <- "not assessed"}');
  WriteLn(outfile,
    'boxplot(ano$dis.rank~ano$class.vec, data=df.data, main=paste("R = ", round(ano$statistic, 3), ", ", "P = ", pval), frame=FALSE, border="steelblue")');
  WriteLn(outfile, 'invisible(dev.off())');
  WriteLn(outfile, 'options(warn=0)');
  CloseFile(outfile);
end;

procedure TANOSIMDlg.ANOSIM(fname: string; groupvar: string;
  transf, coef, nperm: integer; n, m: integer);
var
  line, figf: string;
  infile, outfile: TextFile;
begin
  AssignFile(outfile, fname);
  Rewrite(outfile);
  Header(outfile, strANOSIM);
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

  Write(outfile, strNPerm + IntToStr(nperm) + '<br><br>');

  WriteLn(outfile, '<pre>');
  AssignFile(infile, 'anosim.txt');
  Reset(infile);
  while not EOF(infile) do
  begin
    ReadLn(infile, line);
    if not line.StartsWith('NULL') then
      WriteLn(outfile, line);
  end;
  CloseFile(infile);
  WriteLn(outfile, '</pre>');

  figf := GetFileNameWithoutExt(fname) + '.png';
  WriteLn(outfile, '<p align="left"><img src="' + figf + '"></p>');

  WriteLn(outfile, '</body>');
  WriteLn(outfile, '</html>');
  CloseFile(outfile);
  if FileExists('anosim.txt') then
    DeleteFile('anosim.txt');
end;

end.

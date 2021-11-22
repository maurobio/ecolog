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
  strBiplot = 'BIPLOT';
  strAxis = 'EIXO ';
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
  WriteLn(outfile, 'options(digits=4)');
  WriteLn(outfile, 'suppressPackageStartupMessages(library(vegan))');
  WriteLn(outfile, 'suppressPackageStartupMessages(library(ca))');
  WriteLn(outfile, 'library(vegan, quietly=TRUE)');
  WriteLn(outfile, 'library(ca, quietly=TRUE)');
  WriteLn(outfile, 'df.data <- read.csv("rdata.csv", row.names=1)');
  WriteLn(outfile, 'df.data <- t(df.data)');
  if Length(stransf) > 0 then
    WriteLn(outfile, 'df.data <- ' + stransf + '(df.data + 1)');
  WriteLn(outfile, 'corresp <- ca(df.data, nd=2)');
  WriteLn(outfile, 'pct <- corresp$sv^2/sum(corresp$sv^2)');
  WriteLn(outfile, 'results <- data.frame(corresp$sv, pct, cumsum(pct))');
  WriteLn(outfile, 'results <- t(results)');
  WriteLn(outfile,
    'rownames(results) <- c("Standard deviation", "Proportion of Variance", "Cumulative Proportion")');
  WriteLn(outfile, 'k <- length(corresp$sv)');
  WriteLn(outfile, 'nms <- c()');
  WriteLn(outfile, 'for (i in 1:k) {');
  WriteLn(outfile, 'nms <- c(nms, i)}');
  WriteLn(outfile, 'colnames(results) <- nms');
  WriteLn(outfile, 'sink("ca.txt")');
  WriteLn(outfile, 'cat("Importance of components\n\n")');
  WriteLn(outfile, 'print(results)');
  WriteLn(outfile, 'sink()');
  WriteLn(outfile, 'ppi <- 100');
  figf := GetFileNameWithoutExt(fname) + '.png';
  WriteLn(outfile, 'png("' + figf + '", width=6*ppi, height=6*ppi, res=ppi)');
  WriteLn(outfile, 'par(mar=c(4,4,4,4))');
  if scale = 0 then
    WriteLn(outfile,
      'plot(corresp, main=iconv("' + strBiplot +
    '", from=''UTF-8'', to=''latin1''), xlab=paste("' + strAxis +
      '1 (", round(pct[1], 3)*100, "%)", sep=""), ylab=paste("' + strAxis +
      '2 (", round(pct[2], 3)*100, "%)", sep=""), map="rowprincipal", col="darkgreen", arrows=c(TRUE, FALSE))')
  else if scale = 1 then
    WriteLn(outfile,
      'plot(corresp, main=iconv("' + strBiplot +
    '", from=''UTF-8'', to=''latin1''), xlab=paste("' + strAxis +
      '1 (", round(pct[1], 3)*100, "%)", sep=""), ylab=paste("' + strAxis +
      '2 (", round(pct[2], 3)*100, "%)", sep=""), map="colprincipal", col="darkgreen", arrows=c(FALSE, TRUE))')
  else if scale = 2 then
    WriteLn(outfile,
      'plot(corresp, main=iconv("' + strBiplot +
    '", from=''UTF-8'', to=''latin1''), xlab=paste("' + strAxis +
      '1 (", round(pct[1], 3)*100, "%)", sep=""), ylab=paste("' + strAxis +
      '2 (", round(pct[2], 3)*100, "%)", sep=""), map="symmetric", col="darkgreen", arrows=c(TRUE, TRUE))');
  WriteLn(outfile, 'invisible(dev.off())');
  WriteLn(outfile, 'options(warn=0)');
  CloseFile(outfile);
end;

procedure TCADlg.CA(fname: string; transf: integer; n, m: integer);
var
  line, figf: string;
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

  WriteLn(outfile, '<pre>');
  AssignFile(infile, 'ca.txt');
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
  if FileExists('ca.txt') then
    DeleteFile('ca.txt');
end;

end.

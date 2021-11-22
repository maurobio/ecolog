unit dca;

interface

uses
  Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, ExtCtrls, Spin;

type

  { TDCADlg }

  TDCADlg = class(TForm)
    ComboBoxScale: TComboBox;
    LabelScale: TLabel;
    OKButton: TButton;
    CancelButton: TButton;
    CheckBoxDetrended: TCheckBox;
    CheckBoxDownweight: TCheckBox;
    SpinEditNRescaling: TSpinEdit;
    SpinEditNSegs: TSpinEdit;
    TransformMemo: TMemo;
    CheckBoxRescaling: TCheckBox;
    LabelRescalingN: TLabel;
    LabelNSegs: TLabel;
    procedure FormCreate(Sender: TObject);
  private

  public
    procedure CreateDCA(fname: string; ira, iweigh, mk, scale, iresc, axis: integer);
    procedure DCA(fname: string; ira, iweigh, mk, scale, iresc: integer; n, m: integer);
  end;

var
  DCADlg: TDCADlg;

implementation

{$R *.lfm}

uses report, useful;

resourcestring
  strDCA = 'ANÁLISE DE CORRESPONDÊNCIAS CORRIGIDA';
  strScatterplot = 'DIAGRAMA DE DISPERSÃO';
  strEigenvals = 'AUTOVALORES';
  strDetrended = 'AUTOVALORES CORRIGIDOS';
  strAxis = 'EIXO ';
  strSamples = 'Amostras';
  strSpecies = 'Espécies';
  strScaleSamples = 'Amostras';
  strScaleSpecies = 'Espécies';
  strScaleSymmetric = 'Simétrica';

{ TDCADlg }

procedure TDCADlg.FormCreate(Sender: TObject);
begin
  with ComboBoxScale.Items do
  begin
    Add(strScaleSamples);
    Add(strScaleSpecies);
    Add(strScaleSymmetric);
  end;
  ComboBoxScale.ItemIndex := 0;
end;

procedure TDCADlg.CreateDCA(fname: string; ira, iweigh, mk, scale, iresc, axis: integer);
var
  figf: string;
  outfile: TextFile;
begin
  AssignFile(outfile, 'dca.R');
  Rewrite(outfile);
  WriteLn(outfile, 'options(warn=-1)');
  WriteLn(outfile, 'options(digits=4)');
  WriteLn(outfile, 'suppressPackageStartupMessages(library(vegan))');
  WriteLn(outfile, 'library(vegan, quietly=TRUE)');
  WriteLn(outfile, 'df.data <- read.csv("rdata.csv", row.names=1)');
  WriteLn(outfile, 'df.data <- t(df.data)');
  if scale = 0 then
    iresc := 0;
  WriteLn(outfile, 'dca <- decorana(df.data, iweigh=', iweigh,
    ', iresc=', iresc, ', ira=', ira, ', mk=', mk, ')');
  WriteLn(outfile, 'pct <- dca$evals.decorana^2/sum(dca$evals.decorana^2)');
  WriteLn(outfile, 'sink("dca.txt")');
  WriteLn(outfile, 'print(dca)');
  WriteLn(outfile, 'sink()');
  WriteLn(outfile, 'ppi <- 100');
  figf := GetFileNameWithoutExt(fname) + '.png';
  WriteLn(outfile, 'png("' + figf + '", width=6*ppi, height=6*ppi, res=ppi)');
  WriteLn(outfile, 'par(mar=c(4,4,4,4))');
  if axis = 0 then
    WriteLn(outfile,
      'plot(dca, main=iconv("' + strScatterplot +
    '", from=''UTF-8'', to=''latin1''), xlab=paste("' + strAxis +
      '1 (", round(pct[1], 3)*100, "%)", sep=""), ylab=paste("' + strAxis +
      '2 (", round(pct[2], 3)*100, "%)", sep=""), type="none", display="sites")')
  else if axis = 1 then
    WriteLn(outfile,
      'plot(dca, main=iconv("' + strScatterplot +
    '", from=''UTF-8'', to=''latin1''), xlab=paste("' + strAxis +
      '1 (", round(pct[1], 3)*100, "%)", sep=""), ylab=paste("' + strAxis +
      '2 (", round(pct[2], 3)*100, "%)", sep=""), type="none", display="species")')
  else if axis = 2 then
    WriteLn(outfile,
      'plot(dca, main=iconv("' + strScatterplot +
    '", from=''UTF-8'', to=''latin1''), xlab=paste("' + strAxis +
      '1 (", round(pct[1], 3)*100, "%)", sep=""), ylab=paste("' + strAxis +
      '2 (", round(pct[2], 3)*100, "%)", sep=""), type="none", display="both")');
  WriteLn(outfile, 'points(dca, pch=19, cex=1, col="blue")');
  WriteLn(outfile, 'text(dca, labels=rownames(df.data), cex=1, pos=3)');
  WriteLn(outfile, 'invisible(dev.off())');
  WriteLn(outfile, 'options(warn=0)');
  CloseFile(outfile);
end;

procedure TDCADlg.DCA(fname: string; ira, iweigh, mk, scale, iresc: integer;
  n, m: integer);
var
  line, figf: string;
  infile, outfile: TextFile;
begin
  AssignFile(outfile, fname);
  Rewrite(outfile);
  Header(outfile, strDCA);
  WriteLn(outfile, '<br>');
  WriteLn(outfile, IntToStr(n) + ' ' + LowerCase(strSamples) + ' x ' +
    IntToStr(m) + ' ' + LowerCase(strSpecies) + '<br><br>');

  WriteLn(outfile, '<pre>');
  AssignFile(infile, 'dca.txt');
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
  if FileExists('dca.txt') then
    DeleteFile('dca.txt');
end;

end.

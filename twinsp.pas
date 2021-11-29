unit twinsp;

interface

uses
  Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Grids, ExtCtrls, ComCtrls, Spin;

type

  { TTWSPDlg }

  TTWSPDlg = class(TForm)
    LabelNCutLevels: TLabel;
    OKButton: TButton;
    CancelButton: TButton;
    PseudospeciesStringGrid: TStringGrid;
    PseudoSpeciesRadioGroup: TRadioGroup;
    LabelMinGroupSize: TLabel;
    LabelMaxLevelOfDivs: TLabel;
    SpinEditNCutLevels: TSpinEdit;
    SpinEditMinimumGroupSize: TSpinEdit;
    SpinEditMaximumLevelDivision: TSpinEdit;
    procedure FormCreate(Sender: TObject);
    procedure OKButtonClick(Sender: TObject);
    procedure PseudoSpeciesRadioGroupSelectionChanged(Sender: TObject);
    procedure SpinEditNCutLevelsChange(Sender: TObject);
  private

  public
    CT: string;
    procedure CreateTWSP(fname: string; cut_levels: string;
      min_group_size, levels: integer);
    procedure TWSP(fname: string; cut_levels: string;
      min_group_size, levels: integer; n, m: integer);
  end;

var
  TWSPDlg: TTWSPDlg;

implementation

{$R *.lfm}

uses report, useful;

resourcestring
  strTWSP = 'ANÁLISE DE ESPÉCIES INDICADORAS';
  strSamples = 'Amostras';
  strSpecies = 'Espécies';

procedure DeleteRow(Grid: TStringGrid; ARow: integer);
var
  i: integer;
begin
  for i := ARow to Grid.RowCount - 2 do
    Grid.Rows[i].Assign(Grid.Rows[i + 1]);
  Grid.RowCount := Grid.RowCount - 1;
end;

{ TTWSPDlg }

procedure TTWSPDlg.PseudoSpeciesRadioGroupSelectionChanged(Sender: TObject);
var
  i, r: integer;
begin
  i := PseudoSpeciesRadioGroup.ItemIndex;
  case i of
    0:
    begin
      PseudoSpeciesStringGrid.InsertRowWithValues(1, ['1', '0']);
      PseudoSpeciesStringGrid.InsertRowWithValues(2, ['2', '2']);
      PseudoSpeciesStringGrid.InsertRowWithValues(3, ['3', '5']);
      PseudoSpeciesStringGrid.InsertRowWithValues(4, ['4', '10']);
      PseudoSpeciesStringGrid.InsertRowWithValues(5, ['5', '20']);
      PseudoSpeciesStringGrid.Enabled := False;
      LabelNCutLevels.Enabled := False;
      SpinEditNCutLevels.Enabled := False;
    end;
    1:
    begin
      PseudoSpeciesStringGrid.InsertRowWithValues(1, ['1', '0']);
      for i := 2 to 5 do
        PseudoSpeciesStringGrid.InsertRowWithValues(i, ['', '']);
      PseudoSpeciesStringGrid.Enabled := False;
      LabelNCutLevels.Enabled := False;
      SpinEditNCutLevels.Enabled := False;
    end;
    2:
    begin
      LabelNCutLevels.Enabled := True;
      SpinEditNCutLevels.Enabled := True;
      PseudoSpeciesStringGrid.Enabled := True;
      PseudoSpeciesStringGrid.RowCount := SpinEditNCutLevels.Value + 1;
      for r := 1 to PseudoSpeciesStringGrid.RowCount - 1 do
        DeleteRow(PseudoSpeciesStringGrid, r);
      for r := 1 to SpinEditNCutLevels.Value do
        PseudoSpeciesStringGrid.InsertRowWithValues(r, [IntToStr(r), '0']);
    end
  end;
end;

procedure TTWSPDlg.SpinEditNCutLevelsChange(Sender: TObject);
var
  r: integer;
begin
  PseudoSpeciesStringGrid.RowCount := SpinEditNCutLevels.Value + 1;
  for r := 1 to PseudoSpeciesStringGrid.RowCount - 1 do
    DeleteRow(PseudoSpeciesStringGrid, r);
  for r := 1 to SpinEditNCutLevels.Value do
    PseudoSpeciesStringGrid.InsertRowWithValues(r, [IntToStr(r), '0']);
end;

procedure TTWSPDlg.FormCreate(Sender: TObject);
begin
  PseudoSpeciesRadioGroup.ItemIndex := 0;
  PseudoSpeciesRadioGroupSelectionChanged(Sender);
  //PseudoSpeciesStringGrid.InsertRowWithValues(0, [strPseudospecies, strCutLevel, strWeight, strIndicatorPotential])
end;

procedure TTWSPDlg.OKButtonClick(Sender: TObject);
var
  r: integer;
  l: TStringList;
begin
  CT := '';
  l := TStringList.Create;
  for r := 1 to PseudoSpeciesStringGrid.RowCount - 1 do
  begin
    if PseudoSpeciesStringGrid.Cells[1, r] <> '' then
      l.Add(PseudoSpeciesStringGrid.Cells[1, r]);
  end;
  CT := l.DelimitedText;
  l.Free;
  CreateTWSP('', CT, SpinEditMinimumGroupSize.Value, SpinEditMaximumLevelDivision.Value);
end;

procedure TTWSPDlg.CreateTWSP(fname: string; cut_levels: string;
  min_group_size, levels: integer);
var
  outfile: TextFile;
begin
  AssignFile(outfile, 'twsp.R');
  Rewrite(outfile);
  WriteLn(outfile, 'options(warn=-1)');
  WriteLn(outfile, 'options(digits=4)');
  WriteLn(outfile, 'suppressPackageStartupMessages(library(twinspanR))');
  WriteLn(outfile, 'library(twinspanR, quietly=TRUE)');
  WriteLn(outfile, 'df.data <- read.csv("rdata.csv", row.names=1)');
  WriteLn(outfile, 'df.data <- t(df.data)');
  WriteLn(outfile, 'twsp <- twinspan(df.data, cut.levels=c(', cut_levels, ')',
    ', min.group.size=', min_group_size, ', levels=', levels, ')');
  WriteLn(outfile, 'sink("twsp.txt")');
  WriteLn(outfile, 'summary(twsp)');
  WriteLn(outfile, 'sink()');
  WriteLn(outfile,
    'write.table(file="table.txt", capture.output(print(twsp, what="table")), quote=FALSE, row.names=FALSE)');
  WriteLn(outfile, 'options(warn=0)');
  CloseFile(outfile);
end;

procedure TTWSPDlg.TWSP(fname: string; cut_levels: string;
  min_group_size, levels: integer; n, m: integer);
var
  line: string;
  infile, outfile: TextFile;
begin
  AssignFile(outfile, fname);
  Rewrite(outfile);
  Header(outfile, strTWSP);
  WriteLn(outfile, '<br>');
  WriteLn(outfile, IntToStr(n) + ' ' + LowerCase(strSamples) + ' x ' +
    IntToStr(m) + ' ' + LowerCase(strSpecies) + '<br><br>');

  WriteLn(outfile, '<pre>');
  AssignFile(infile, 'twsp.txt');
  Reset(infile);
  while not EOF(infile) do
  begin
    ReadLn(infile, line);
    WriteLn(outfile, line);
  end;
  CloseFile(infile);
  WriteLn(outfile);
  AssignFile(infile, 'table.txt');
  Reset(infile);
  while not EOF(infile) do
  begin
    ReadLn(infile, line);
    if not line.StartsWith('x') then
      WriteLn(outfile, line);
  end;
  CloseFile(infile);
  WriteLn(outfile, '</pre>');

  WriteLn(outfile, '</body>');
  WriteLn(outfile, '</html>');
  CloseFile(outfile);
  if FileExists('twsp.txt') then
    DeleteFile('twsp.txt');
  if FileExists('table.txt') then
    DeleteFile('table.txt');
end;

end.

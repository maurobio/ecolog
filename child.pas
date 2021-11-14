unit child;

{$mode objfpc}{$H+}

interface

uses
  LCLIntf, LCLType, Classes, SysUtils, DB, fpcsvexport, Forms, Controls,
  Graphics, Dialogs, ExtCtrls, Grids, DBGrids, fpspreadsheetctrls, fpspreadsheetgrid,
  fpsallformats, fpstypes, fpspreadsheet, fpsdataset, HtmlView, ZMConnection,
  ZMQueryDataSet;

type

  { TFrameTemplate }

  { TMDIChild }

  TMDIChild = class(TFrame)
    CSVExporter1: TCSVExporter;
    CSVExporter2: TCSVExporter;
    DataGrid: TDBGrid;
    DataSource1: TDataSource;
    Dataset1: TsWorksheetDataset;
    DataSource2: TDataSource;
    HtmlViewer: THtmlViewer;
    QueryConnection1: TZMConnection;
    QueryDataSet1: TZMQueryDataSet;
    Dataset2: TsWorksheetDataset;
    QueryConnection2: TZMConnection;
    QueryDataSet2: TZMQueryDataSet;
    procedure DataGridMouseMove(Sender: TObject; Shift: TShiftState; X, Y: integer);
    procedure DataGridTitleClick(Column: TColumn);
  private

  public

  end;

var
  MDIChild: TMDIChild;

implementation

uses main;

{$R *.lfm}

{ TMDIChild }

procedure TMDIChild.DataGridMouseMove(Sender: TObject; Shift: TShiftState;
  X, Y: integer);
var
  pt: TGridCoord;
begin
  pt := DataGrid.MouseCoord(x, y);
  if pt.y = 0 then
    DataGrid.Cursor := crHandPoint
  else
    DataGrid.Cursor := crDefault;
end;

procedure TMDIChild.DataGridTitleClick(Column: TColumn);
var
  i: integer;
begin
  with DataGrid do
    for i := 0 to Columns.Count - 1 do
      Columns[i].Title.Font.Style := Columns[i].Title.Font.Style - [fsBold];
  Column.Title.Font.Style := Column.Title.Font.Style + [fsBold];
  Dataset1.SortOnFields(Column.Field.FieldName);
end;

end.

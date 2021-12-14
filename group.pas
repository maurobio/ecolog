unit group;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, ValEdit, Grids,
  StdCtrls, StrUtils, DB;

type

  { TGroupDlg }

  TGroupDlg = class(TForm)
    OKButton: TButton;
    CancelButton: TButton;
    GroupEditor: TValueListEditor;
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure OKButtonClick(Sender: TObject);
  private

  public
    Factors: string;
    GroupVariable: string;
    GroupValues: TStringList;
  end;

var
  GroupDlg: TGroupDlg;

implementation

{$R *.lfm}

uses main, child;

resourcestring
  strWarning = 'Aviso';
  strGroup = 'Grupo';
  strNoGroups = 'Número insuficiente de grupos';

function FindString(const SearchKey: string; SearchList: TStrings): integer;
var
  Found: boolean;
  I: integer;
  SearchStr: string;
begin
  Found := False;
  I := 0;
  while (I < SearchList.Count) and (not Found) do
  begin
    SearchStr := SearchList.Strings[I];
    if (Pos(SearchKey, SearchStr) <> 0) then
    begin
      Result := I;
      Found := True;
      Exit;
    end;
    Inc(I);
  end;
  if not Found then
    Result := -1;
end;

{ TGroupDlg }

procedure TGroupDlg.FormShow(Sender: TObject);
var
  bm: TBookmark;
begin
  GroupEditor.Strings.Clear;
  with GroupEditor.TitleCaptions do
  begin
    Add(GroupVariable);
    Add(strGroup);
  end;
  with MainForm.MultiDoc.ActiveObject as TMDIChild do
  begin
    bm := Dataset1.GetBookmark;
    Dataset1.DisableControls;
    Dataset1.First;
    while not Dataset1.EOF do
    begin
      if FindString(Dataset1.FieldByName(GroupVariable).AsString, GroupValues) < 0 then
        GroupValues.Add(Dataset1.FieldByName(GroupVariable).AsString + '=0');
      Dataset1.Next;
    end;
    Dataset1.EnableControls;
    Dataset1.GotoBookmark(bm);
  end;
  GroupEditor.Strings.AddStrings(GroupValues);
  GroupEditor.KeyOptions := [keyEdit];
end;

procedure TGroupDlg.FormCreate(Sender: TObject);
begin
  GroupValues := TStringList.Create;
  Factors := '';
end;

procedure TGroupDlg.FormDestroy(Sender: TObject);
begin
  GroupValues.Free;
end;

procedure TGroupDlg.OKButtonClick(Sender: TObject);
var
  gp: string;
  i: integer;
begin
  gp := '';
  for i := 0 to GroupEditor.Strings.Count - 1 do
  begin
    if StrToInt(GroupEditor.Strings.ValueFromIndex[i]) > 0 then
      gp := gp + GroupEditor.Strings.ValueFromIndex[i] + ',';
  end;
  gp := Copy(gp, 1, RPos(',', gp) - 1);
  if Length(gp) > 0 then
    Factors := gp
  else
    MessageDlg(strWarning, strNoGroups, mtWarning, [mbOK], 0);
end;

end.


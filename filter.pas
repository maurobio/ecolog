unit filter;

{$mode objfpc}{$H+}

interface

uses
  LCLIntf, LCLType, Classes, Graphics, Forms, Controls, Buttons, StdCtrls,
  SysUtils, Dialogs, Menus, ExtCtrls, DB, sqldb, MaskEdit, CheckLst;

const
  //operL: array of string = ('AND', 'OR');
  operM: array of string = ('>', '>=', '<', '<=', '=', '<>', '*');

type

  { TFilterDlg }

  TFilterDlg = class(TForm)
    Bevel1: TBevel;
    Label1: TLabel;
    Label2: TLabel;
    comboFields: TComboBox;
    editFilter: TComboBox;
    Label3: TLabel;
    comboConditional: TComboBox;
    OKBtn: TButton;
    CancelBtn: TButton;
    ExcludeBtn: TSpeedButton;
    IncludeBtn: TSpeedButton;
    Label5: TLabel;
    Bevel2: TBevel;
    ClearBtn: TSpeedButton;
    listFilters: TCheckListBox;
    MSOfficeCaption: TLabel;
    rbAND: TRadioButton;
    rbOR: TRadioButton;
    procedure FormClose(Sender: TObject; var CloseAction: TCloseAction);
    procedure OKBtnClick(Sender: TObject);
    procedure IncludeBtnClick(Sender: TObject);
    procedure ClearBtnClick(Sender: TObject);
    procedure ExcludeBtnClick(Sender: TObject);
    procedure editFilterKeyDown(Sender: TObject; var Key: word; Shift: TShiftState);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure editFilterClick(Sender: TObject);
    procedure comboFieldsChange(Sender: TObject);
  private
    { Private declarations }
    Conditions, Queries: TStringList;
    procedure SetComboField;
    procedure SetQuery;
    procedure EnableControls;
  public
    { Public declarations }
    FFilter: string;
  published
  end;

var
  FilterDlg: TFilterDlg;

implementation

uses main, child;

resourcestring
    strGreater = 'Maior';
    strGreaterEqual = 'Maior ou Igual';
    strLesser = 'Menor';
    strLesserEqual = 'Menor ou Igual';
    strEqual = 'Igual';
    strDifferent = 'Diferente';
    strLike = 'Contendo';

{$R *.lfm}

procedure TFilterDlg.editFilterKeyDown(Sender: TObject; var Key: word;
  Shift: TShiftState);
begin
  if Key = VK_RETURN then
    IncludeBtnClick(Sender);
end;

procedure TFilterDlg.FormCreate(Sender: TObject);
begin
  Conditions := TStringList.Create;
  Queries := TStringList.Create;
  with ComboConditional.Items do
  begin
    Add(strGreater);
    Add(strGreaterEqual);
    Add(strLesser);
    Add(strLesserEqual);
    Add(strEqual);
    Add(strDifferent);
    Add(strLike);
  end;
end;

procedure TFilterDlg.FormDestroy(Sender: TObject);
begin
  Conditions.Free;
  Queries.Free;
end;

procedure TFilterDlg.FormShow(Sender: TObject);
begin
  SetComboField;
  comboConditional.ItemIndex := 4;
  editFilter.Text := '';
  listFilters.Clear;
  EnableControls;
end;

procedure TFilterDlg.IncludeBtnClick(Sender: TObject);
begin
  if editFilter.Text = '' then
    Exit;
  //if (operM[comboConditional.ItemIndex] = 'LIKE') or
  //  (operM[comboConditional.ItemIndex] = 'NOT LIKE') then
  //  Queries.Add(comboFields.Text + ' ' + operM[comboConditional.ItemIndex] + ' ' +
  //    ' %' + QuotedStr('%' + editFilter.Text + '%'))
  if (operM[comboConditional.ItemIndex] = '*') then
    Queries.Add(comboFields.Text + ' ' + operM[4] + ' ' +
      QuotedStr('*' + editFilter.Text + '*'))
  else
    Queries.Add(comboFields.Text + ' ' + operM[comboConditional.ItemIndex] +
      ' ' + QuotedStr(editFilter.Text));
  listFilters.Items.Add(comboFields.Text + ' ' + comboConditional.Text +
    ' ' + QuotedStr(editFilter.Text));
  listFilters.Checked[listFilters.Count - 1] := True;
  EnableControls;
end;

procedure TFilterDlg.ClearBtnClick(Sender: TObject);
begin
  listFilters.Clear;
  Queries.Clear;
  EnableControls;
end;

procedure TFilterDlg.ExcludeBtnClick(Sender: TObject);
begin
  with listFilters do
  begin
    if Items.Count > 0 then
    begin
      if ItemIndex > -1 then
      begin
        Items.Delete(ItemIndex);
        Queries.Delete(ItemIndex);
      end;
      if Items.Count > 0 then
        ItemIndex := 0;
    end;
    Refresh;
  end;
  EnableControls;
end;

procedure TFilterDlg.OKBtnClick(Sender: TObject);
var
  i: integer;
begin
  Conditions.Clear;
  if listFilters.Items.Count > 0 then
  begin
    for i := 0 to listFilters.Items.Count - 1 do
    begin
      if listFilters.Checked[i] then
      begin
        if i <> (listFilters.Items.Count - 1) then
        begin
          if rbAND.Checked then
            //Conditions.Add(listFilters.Items.Strings[i] + ' AND ')
            Conditions.Add(Queries[i] + ' AND ')
          else if rbOR.Checked then
            //Conditions.Add(listFilters.Items.Strings[i] + ' OR ');
            Conditions.Add(Queries[i] + ' OR ');
        end
        else
          //Conditions.Add(listFilters.Items.Strings[i]);
          Conditions.Add(Queries[i]);
      end;
    end;
  end;
  SetQuery;
  editFilter.Clear;
end;

procedure TFilterDlg.FormClose(Sender: TObject; var CloseAction: TCloseAction);
begin
  Queries.Clear;
end;

procedure TFilterDlg.SetComboField;
var
  col: cardinal;
begin
  comboFields.Clear;
  with MainForm.MultiDoc.ActiveObject as TMDIChild do
  begin
    for col := 0 to Dataset1.FieldCount - 1 do
      comboFields.Items.Add(Dataset1.Fields[col].FieldName);
  end;
  comboFields.ItemIndex := 0;
end;

procedure TFilterDlg.SetQuery;
var
  i: integer;
begin
  FFilter := '';
  if Conditions.Count > 0 then
  begin
    for i := 0 to Conditions.Count - 1 do
      FFilter := FFilter + Conditions[i] + ' ';
  end;
end;

procedure TFilterDlg.editFilterClick(Sender: TObject);
var
  listPos: integer;
  lookupList: TStringList;
begin
  lookupList := TStringList.Create;
  lookupList.Sorted := True;
  lookupList.Duplicates := dupIgnore;
  listPos := comboFields.Items.IndexOf(comboFields.Text);
  with MainForm.MultiDoc.ActiveObject as TMDIChild do
  begin
    Dataset1.First;
    while not Dataset1.EOF do
    begin
      lookupList.Add(Dataset1.Fields[listPos].AsString);
      Dataset1.Next;
    end;
    Dataset1.First;
  end;
  editFilter.Items.Clear;
  editFilter.Items.AddStrings(lookupList);
  lookupList.Free;
end;

procedure TFilterDlg.comboFieldsChange(Sender: TObject);
begin
  editFilter.Text := '';
  editFilter.Enabled := (comboFields.ItemIndex <> 7);
  comboConditional.Enabled := (comboFields.ItemIndex <> 7);
end;

procedure TFilterDlg.EnableControls;
begin
  rbAND.Enabled := listFilters.Items.Count > 0;
  rbOR.Enabled := listFilters.Items.Count > 0;
  if rbAND.Enabled then
  begin
    rbAND.Checked := rbAND.Enabled;
    rbOR.Checked := not rbAND.Enabled;
  end
  else
  begin
    rbAND.Checked := rbAND.Enabled;
    rbOR.Checked := rbAND.Enabled;
  end;
  ExcludeBtn.Enabled := listFilters.Items.Count > 0;
  ClearBtn.Enabled := listFilters.Items.Count > 0;
end;

end.

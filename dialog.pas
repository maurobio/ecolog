unit dialog;

{$mode objfpc}{$H+}

interface

uses
  SysUtils, LCLIntf, LCLType, Classes, Graphics, Controls,
  Forms, Dialogs, Menus, StdCtrls, ExtCtrls, Types;

function InputCombo(const ACaption, APrompt: string; const AList: TStrings): string;
function CheckListDialog(const ACaption, APrompt: string; const AList: TStrings;
  Checked: boolean = False): string;

implementation

uses CheckLst;

resourcestring
  strOK = 'OK';
  strCancel = 'Cancelar';

function GetCharSize(Canvas: TCanvas): TPoint;
var
  I: integer;
  Buffer: array[0..51] of char;
begin
  for I := 0 to 25 do
    Buffer[I] := Chr(I + Ord('A'));
  for I := 0 to 25 do
    Buffer[I + 26] := Chr(I + Ord('a'));
  GetTextExtentPoint(Canvas.Handle, Buffer, 52, TSize(Result));
  Result.X := Result.X div 52;
end;

function GetSelectedItems(CheckListBox: TCheckListBox): string;
var
  i: integer;
begin
  Result := '';
  for i := 0 to CheckListBox.Items.Count - 1 do
    if CheckListBox.State[i] = cbChecked then
      Result := Result + QuotedStr(CheckListBox.Items[i]) + ',';
  Result := Copy(Result, 0, Length(Result) - 1);
end;

function InputCombo(const ACaption, APrompt: string; const AList: TStrings): string;
var
  Form: TForm;
  Prompt: TLabel;
  Combo: TComboBox;
  DialogUnits: TPoint;
  ButtonTop, ButtonWidth, ButtonHeight: integer;
begin
  Result := '';
  Form := TForm.Create(Application);
  with Form do
    try
      Canvas.Font := Font;
      DialogUnits := GetCharSize(Canvas);
      BorderStyle := bsDialog;
      Caption := ACaption;
      ClientWidth := MulDiv(180, DialogUnits.X, 4);
      Position := poScreenCenter;
      Prompt := TLabel.Create(Form);
      with Prompt do
      begin
        Parent := Form;
        Caption := APrompt;
        Left := MulDiv(8, DialogUnits.X, 4);
        Top := MulDiv(8, DialogUnits.Y, 8);
        Constraints.MaxWidth := MulDiv(164, DialogUnits.X, 4);
        WordWrap := True;
      end;
      Combo := TComboBox.Create(Form);
      with Combo do
      begin
        Parent := Form;
        Style := csDropDownList;
        Items.Assign(AList);
        ItemIndex := 0;
        Left := Prompt.Left;
        Top := Prompt.Top + Prompt.Height + 5;
        Width := MulDiv(164, DialogUnits.X, 4);
      end;
      ButtonTop := Combo.Top + Combo.Height + 15;
      ButtonWidth := MulDiv(50, DialogUnits.X, 4);
      ButtonHeight := MulDiv(14, DialogUnits.Y, 8);
      with TButton.Create(Form) do
      begin
        Parent := Form;
        Caption := strOK; //'OK';
        ModalResult := mrOk;
        default := True;
        SetBounds(MulDiv(38, DialogUnits.X, 4), ButtonTop, ButtonWidth,
          ButtonHeight);
      end;
      with TButton.Create(Form) do
      begin
        Parent := Form;
        Caption := strCancel; //'Cancelar';
        ModalResult := mrCancel;
        Cancel := True;
        SetBounds(MulDiv(92, DialogUnits.X, 4), Combo.Top + Combo.Height + 15,
          ButtonWidth, ButtonHeight);
        Form.ClientHeight := Top + Height + 13;
      end;
      if ShowModal = mrOk then
        Result := Combo.Text;
    finally
      Form.Free;
    end;
end;

function CheckListDialog(const ACaption, APrompt: string; const AList: TStrings;
  Checked: boolean = False): string;
var
  Form: TForm;
  Prompt: TLabel;
  ListBox: TCheckListBox;
  DialogUnits: TPoint;
  ButtonTop, ButtonWidth, ButtonHeight, I: integer;
begin
  Result := '';
  Form := TForm.Create(Application);
  with Form do
    try
      Canvas.Font := Font;
      DialogUnits := GetCharSize(Canvas);
      BorderStyle := bsDialog;
      Caption := ACaption;
      ClientWidth := MulDiv(180, DialogUnits.X, 4);
      Position := poScreenCenter;
      Prompt := TLabel.Create(Form);
      with Prompt do
      begin
        Parent := Form;
        Caption := APrompt;
        Left := MulDiv(8, DialogUnits.X, 4);
        Top := MulDiv(8, DialogUnits.Y, 8);
        Constraints.MaxWidth := MulDiv(164, DialogUnits.X, 4);
        WordWrap := True;
      end;
      ListBox := TCheckListBox.Create(Form);
      with ListBox do
      begin
        Parent := Form;
        Items.Assign(AList);
        ItemIndex := 0;
        Left := Prompt.Left;
        Top := Prompt.Top + Prompt.Height + 5;
        Width := MulDiv(164, DialogUnits.X, 4);
      end;
      for I := 0 to ListBox.Count - 1 do
        if Checked then
          ListBox.Checked[I] := True
        else
          ListBox.Checked[I] := False;
      ButtonTop := ListBox.Top + ListBox.Height + 15;
      ButtonWidth := MulDiv(50, DialogUnits.X, 4);
      ButtonHeight := MulDiv(14, DialogUnits.Y, 8);
      with TButton.Create(Form) do
      begin
        Parent := Form;
        Caption := strOK; //'OK';
        ModalResult := mrOk;
        default := True;
        SetBounds(MulDiv(38, DialogUnits.X, 4), ButtonTop, ButtonWidth,
          ButtonHeight);
      end;
      with TButton.Create(Form) do
      begin
        Parent := Form;
        Caption := strCancel; //'Cancelar';
        ModalResult := mrCancel;
        Cancel := True;
        SetBounds(MulDiv(92, DialogUnits.X, 4), ListBox.Top + ListBox.Height + 15,
          ButtonWidth, ButtonHeight);
        Form.ClientHeight := Top + Height + 13;
      end;
      if ShowModal = mrOk then
        Result := GetSelectedItems(ListBox);
    finally
      Form.Free;
    end;
end;

end.

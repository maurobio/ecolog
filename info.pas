unit info;

{$mode objfpc}{$H+}

interface

uses
  LCLIntf, LCLType, Classes, Graphics, Forms, Controls, Buttons, StdCtrls,
  ExtCtrls, ComCtrls, SysUtils, Dialogs, Spin;

type

  { TProjectDlg }

  TProjectDlg = class(TForm)
    MethodComboBox: TComboBox;
    TitleEdit: TEdit;
    RoleEdit: TEdit;
    InstitutionEdit: TEdit;
    Address1Edit: TEdit;
    Address2Edit: TEdit;
    CityEdit: TEdit;
    UFEdit: TEdit;
    ZIPEdit: TEdit;
    PhoneEdit: TEdit;
    EmailEdit: TEdit;
    WebsiteEdit: TEdit;
    AuthorEdit: TEdit;
    FundingEdit: TEdit;
    SizeEdit: TEdit;
    CountryEdit: TEdit;
    StateEdit: TEdit;
    ProvinceEdit: TEdit;
    LatitudeEdit: TFloatSpinEdit;
    LongitudeEdit: TFloatSpinEdit;
    ElevationEdit: TSpinEdit;
    TitleLabel: TLabel;
    LatitudeLabel: TLabel;
    LongitudeLabel: TLabel;
    ElevationLabel: TLabel;
    RoleLabel: TLabel;
    InstitutionLabel: TLabel;
    Address1Label: TLabel;
    Address2Label: TLabel;
    CityLabel: TLabel;
    UFLabel: TLabel;
    ZIPLabel: TLabel;
    AuthorLabel: TLabel;
    PhoneLabel: TLabel;
    EmailLabel: TLabel;
    WebsiteLabel: TLabel;
    FundingLabel: TLabel;
    DescriptionLabel: TLabel;
    MethodLabel: TLabel;
    SizeLabel: TLabel;
    CountryLabel: TLabel;
    StateLabel: TLabel;
    ProvinceLabel: TLabel;
    LocalityLabel: TLabel;
    DescriptionMemo: TMemo;
    LocalityMemo: TMemo;
    OKButton: TButton;
    CancelButton: TButton;
    PageControl: TPageControl;
    TabSheetContact: TTabSheet;
    TabSheetDescription: TTabSheet;
    TabSheetSite: TTabSheet;
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure OKButtonClick(Sender: TObject);
  private
  public
  end;

var
  ProjectDlg: TProjectDlg;

implementation

uses main;

{$R *.lfm}

resourcestring
  strRandom = 'Coleta Aleatória';
  strStation = 'Estação';
  strPlot = 'Parcela';
  strPoint = 'Ponto';
  strQuadrat = 'Quadrante';
  strTransect = 'Transecção';

{ TProjectDlg }

procedure TProjectDlg.FormShow(Sender: TObject);
begin
  TitleEdit.Text := Metadata.Force('title').AsString;
  AuthorEdit.Text := Metadata.Force('author').AsString;
  DescriptionMemo.Text := Metadata.Force('description').AsString;
  MethodComboBox.Text := Metadata.Force('method').AsString;
  SizeEdit.Text := Metadata.Force('size').AsString;
  CountryEdit.Text := Metadata.Force('country').AsString;
  StateEdit.Text := Metadata.Force('state').AsString;
  ProvinceEdit.Text := Metadata.Force('province').AsString;
  LocalityMemo.Text := Metadata.Force('locality').AsString;
  LatitudeEdit.Text := Metadata.Force('latitude').AsString;
  LongitudeEdit.Text := Metadata.Force('longitude').AsString;
  ElevationEdit.Text := Metadata.Force('elevation').AsString;
  RoleEdit.Text := Metadata.Force('role').AsString;
  InstitutionEdit.Text := Metadata.Force('institution').AsString;
  Address1Edit.Text := Metadata.Force('address1').AsString;
  Address2Edit.Text := Metadata.Force('address2').AsString;
  CityEdit.Text := Metadata.Force('city').AsString;
  UFEdit.Text := Metadata.Force('uf').AsString;
  ZIPEdit.Text := Metadata.Force('zip').AsString;
  PhoneEdit.Text := Metadata.Force('phone').AsString;
  EmailEdit.Text := Metadata.Force('email').AsString;
  WebsiteEdit.Text := Metadata.Force('website').AsString;
  FundingEdit.Text := Metadata.Force('funding').AsString;
  PageControl.ActivePage := TabSheetDescription;
end;

procedure TProjectDlg.OKButtonClick(Sender: TObject);
begin
  Metadata.Find('title').AsString := TitleEdit.Text;
  Metadata.Find('author').AsString := AuthorEdit.Text;
  Metadata.Find('description').AsString := DescriptionMemo.Text;
  Metadata.Find('method').AsString := MethodComboBox.Text;
  Metadata.Find('size').AsString := SizeEdit.Text;
  Metadata.Find('country').AsString := CountryEdit.Text;
  Metadata.Find('state').AsString := StateEdit.Text;
  Metadata.Find('province').AsString := ProvinceEdit.Text;
  Metadata.Find('locality').AsString := LocalityMemo.Text;
  Metadata.Find('latitude').AsString := LatitudeEdit.Text;
  Metadata.Find('longitude').AsString := LongitudeEdit.Text;
  Metadata.Find('elevation').AsString := ElevationEdit.Text;
  Metadata.Find('role').AsString := RoleEdit.Text;
  Metadata.Find('institution').AsString := InstitutionEdit.Text;
  Metadata.Find('address1').AsString := Address1Edit.Text;
  Metadata.Find('address2').AsString := Address2Edit.Text;
  Metadata.Find('city').AsString := CityEdit.Text;
  Metadata.Find('uf').AsString := UFEdit.Text;
  Metadata.Find('zip').AsString := ZIPEdit.Text;
  Metadata.Find('phone').AsString := PhoneEdit.Text;
  Metadata.Find('email').AsString := EmailEdit.Text;
  Metadata.Find('website').AsString := WebsiteEdit.Text;
  Metadata.Find('funding').AsString := FundingEdit.Text;
end;

procedure TProjectDlg.FormCreate(Sender: TObject);
begin
   with MethodComboBox.Items do
   begin
     Add(strRandom);
     Add(strStation);
     Add(strPlot);
     Add(strPoint);
     Add(strQuadrat);
     Add(strTransect);
   end;
end;

{ TProjectDlg }

end.

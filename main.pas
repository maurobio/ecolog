{==============================================================================}
{ ECOLOG - Sistema Gerenciador de Banco de Dados para Levantamentos Ecológicos }
{        ECOLOG - Database Management System for Ecological Surveys            }
{      Copyright (c) 1990-2021 Mauro J. Cavalcanti. All rights reserved.       }
{                                                                              }
{   Este programa é software livre; você pode redistribuí-lo e/ou              }
{   modificá-lo sob os termos da Licença Pública Geral GNU, conforme           }
{   publicada pela Free Software Foundation; tanto a versão 2 da               }
{   Licença como (a seu critério) qualquer versão mais nova.                   }
{                                                                              }
{   Este programa é distribuído na expectativa de ser útil, mas SEM            }
{   QUALQUER GARANTIA; sem mesmo a garantia implícita de                       }
{   COMERCIALIZAÇÃO ou de ADEQUAÇÃO A QUALQUER PROPÓSITO EM                    }
{   PARTICULAR. Consulte a Licença Pública Geral GNU para obter mais           }
{   detalhes.                                                                  }
{                                                                              }
{   Você deve ter recebido uma cópia da Licença Pública Geral GNU              }
{   junto com este programa; se não, veja <http://ww w.gnu.org/licenses/>.     }
{                                                                              }
{   This program is free software: you can redistribute it and/or              }
{   modify it under the terms of the GNU General Public License as published   }
{   by the Free Software Foundation, either version 2 of the License, or       }
{   version 3 of the License, or (at your option) any later version.           }
{                                                                              }
{   This program is distributed in the hope that it will be useful,            }
{   but WITHOUT ANY WARRANTY; without even the implied warranty of             }
{   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See                   }
{   the GNU General Public License for more details.                           }
{                                                                              }
{   You should have received a copy of the GNU General Public License          }
{   along with this program. If not, see <http://www.gnu.org/licenses/>.       }
{                                                                              }
{   Requerimentos / Requirements:                                              }
{     Free Pascal Compiler 3.0+ (www.freepascal.org)                           }
{     Lazarus IDE 2.0+ (www.lazarus.freepascal.org)                            }
{     FPSpreadsheetDataset 1.13+ (github.com/wp-xyz/FPSpreadsheetDataset)      }
{     FuzzyWuzzy (github.com/DavidMoraisFerreira/FuzzyWuzzy.pas                }
{     HtmlViewer 11.9+ (wiki.freepascal.org/THtmlPort)                         }
{     JSON Tools (github.com/sysrpl/JsonTools)                                 }
{     MultiDoc 0.3+ (wiki.lazarus.freepascal.org/MultiDoc)                     }
{     NaturalSort (wiki.freepascal.org/NaturalSort)                            }
{                                                                              }
{   Histórico de Revisões / Revision History:                                  }
{     Versão 6.0.1 "Amaranthus", 14/11/2021                                    }
{       -- Utilização direta de planilhas eletrônicas nos formatos do          }
{          MS-Excel (xls, xlsx), OO-Calc (ods) e texto delimitado (csv, tsv).  }
{       -- Opções para exportação de dados em diferentes formatos de programas }
{          externos (Cornell, Fitopac, MVSP, BRAHMS).                          }
{       -- Integração com o sistema estatístico e gráfico R para análises      }
{          multivariadas de dados ecológicos (índices de similaridade, análise }
{          de agrupamentos e métodos de ordenação).                            }
{     Versão 6.0.2 "Begonia", 22/11/2021                                       }
{       -- Inclusão da Análise de Correspondências Corrigida (DECORANA).       }
{       -- Inclusãão das saídas do R nos relatórios analíticos do sistema.     }
{==============================================================================}
unit main;

{$mode objfpc}{$H+}

interface

uses
  LCLIntf, LCLType, LCLTranslator, Classes, SysUtils, DB, fpcsvexport, Forms,
  Controls, Graphics, Dialogs, Menus, ComCtrls, MultiFrame, Types, StrUtils,
  Process, Variants, IniFiles, JsonTools, ZMQueryDataSet, FPSpreadsheet,
  fpspreadsheetctrls, fpspreadsheetgrid, fpstypes, fpsdataset;

type

  { TMainForm }

  TMainForm = class(TForm)
    AddButton: TToolButton;
    CatalogButton: TToolButton;
    ChecklistButton: TToolButton;
    ClusterButton: TToolButton;
    D1: TToolButton;
    D2: TToolButton;
    D3: TToolButton;
    DiversityButton: TToolButton;
    FilterButton: TToolButton;
    FindButton: TToolButton;
    GeocodeButton: TToolButton;
    ImageList: TImageList;
    LabelsButton: TToolButton;
    MainMenu: TMainMenu;
    FileMenu: TMenuItem;
    LanguageMenu: TMenuItem;
    LanguageMenuEnglishItem: TMenuItem;
    LanguageMenuPortugueseItem: TMenuItem;
    LanguageMenuSpanishItem: TMenuItem;
    OrdinationDCAItem: TMenuItem;
    N7: TMenuItem;
    SearchFilterResetItem: TMenuItem;
    SearchMenu: TMenuItem;
    SearchSortItem: TMenuItem;
    SearchFindItem: TMenuItem;
    FileNewItem: TMenuItem;
    FileAddItem: TMenuItem;
    FileReloadItem: TMenuItem;
    FileCloseItem: TMenuItem;
    FileRemoveItem: TMenuItem;
    FileExportItem: TMenuItem;
    AnalysisClusterItem: TMenuItem;
    AnalysisOrdinationItem: TMenuItem;
    PCAItem: TMenuItem;
    PCOItem: TMenuItem;
    NMDSItem: TMenuItem;
    CAItem: TMenuItem;
    RAItem: TMenuItem;
    CCAItem: TMenuItem;
    NamesButton: TToolButton;
    NewButton: TToolButton;
    OpenButton: TToolButton;
    OrdinationButton: TToolButton;
    OrdinationMenu: TPopupMenu;
    ReloadButton: TToolButton;
    RemoveButton: TToolButton;
    ReportGeocodeItem: TMenuItem;
    ReportChecklistItem: TMenuItem;
    ReportStatisticsItem: TMenuItem;
    OrdinationPCAItem: TMenuItem;
    OrdinationPCOItem: TMenuItem;
    OrdinationNMDSItem: TMenuItem;
    OrdinationCAItem: TMenuItem;
    OrdinationRDAItem: TMenuItem;
    OrdinationCCAItem: TMenuItem;
    SortButton: TToolButton;
    StatisticsButton: TToolButton;
    ToolBar: TToolBar;
    WindowPreviousItem: TMenuItem;
    WindowNextItem: TMenuItem;
    MenuItem3: TMenuItem;
    WindowCloseItem: TMenuItem;
    N6: TMenuItem;
    N5: TMenuItem;
    N4: TMenuItem;
    WindowCascadeItem: TMenuItem;
    WindowMenu: TMenuItem;
    WindowTileHorizontalItem: TMenuItem;
    WindowTileVerticalItem: TMenuItem;
    MultiDoc: TMultiFrame;
    ReportLabelsItem: TMenuItem;
    N1: TMenuItem;
    SearchFilterItem: TMenuItem;
    SaveDialog: TSaveDialog;
    ReportNamesItem: TMenuItem;
    HelpMenu: TMenuItem;
    HelpAboutItem: TMenuItem;
    MenuItem11: TMenuItem;
    AnalysisDiversityItem: TMenuItem;
    FileOpenItem: TMenuItem;
    FileChangeItem: TMenuItem;
    FileExit: TMenuItem;
    N0: TMenuItem;
    ReportMenu: TMenuItem;
    AnalysisMenu: TMenuItem;
    ReportCatalogItem: TMenuItem;
    OpenDialog: TOpenDialog;
    StatusBar: TStatusBar;
    procedure AnalysisClusterItemClick(Sender: TObject);
    procedure AnalysisDiversityItemClick(Sender: TObject);
    procedure FileAddItemClick(Sender: TObject);
    procedure FileChangeItemClick(Sender: TObject);
    procedure FileCloseItemClick(Sender: TObject);
    procedure FileReloadItemClick(Sender: TObject);
    procedure FileRemoveItemClick(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: boolean);
    procedure LanguageMenuEnglishItemClick(Sender: TObject);
    procedure LanguageMenuPortugueseItemClick(Sender: TObject);
    procedure LanguageMenuSpanishItemClick(Sender: TObject);
    procedure OrdinationDCAItemClick(Sender: TObject);
    procedure ReportCatalogItemClick(Sender: TObject);
    procedure ReportChecklistItemClick(Sender: TObject);
    procedure ReportGeocodeItemClick(Sender: TObject);
    procedure ReportLabelsItemClick(Sender: TObject);
    procedure ReportNamesItemClick(Sender: TObject);
    procedure ReportStatisticsItemClick(Sender: TObject);
    procedure SearchFilterItemClick(Sender: TObject);
    procedure SearchFindItemClick(Sender: TObject);
    procedure FileExportItemClick(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure MultiDocDeleteChild(Sender: TObject);
    procedure OrdinationCAItemClick(Sender: TObject);
    procedure OrdinationCCAItemClick(Sender: TObject);
    procedure OrdinationNMDSItemClick(Sender: TObject);
    procedure OrdinationPCAItemClick(Sender: TObject);
    procedure OrdinationPCOItemClick(Sender: TObject);
    procedure OrdinationRDAItemClick(Sender: TObject);
    procedure SearchFilterResetItemClick(Sender: TObject);
    procedure SearchSortItemClick(Sender: TObject);
    procedure WindowCascadeItemClick(Sender: TObject);
    procedure FileNewItemClick(Sender: TObject);
    procedure WindowCloseItemClick(Sender: TObject);
    procedure WindowNextItemClick(Sender: TObject);
    procedure WindowPreviousItemClick(Sender: TObject);
    procedure WindowTileHorizontalItemClick(Sender: TObject);
    procedure WindowTileVerticalItemClick(Sender: TObject);
    procedure HelpAboutItemClick(Sender: TObject);
    procedure FileExitClick(Sender: TObject);
    procedure FileOpenItemClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
  private
    function CheckData(SrcFileName: string): boolean;
    function LocateR: boolean;
    function FindR: string;
    procedure CreateMDIChild(const AName: string);
    procedure UpdateMenuItems(Sender: TObject);
    procedure UpdateTitleBar(Filename: string);
    procedure RestoreSettings;
    procedure SaveSettings;
    procedure GridToArray(const filename: string; separator: char; var n, m: integer);
  public
    ChildCount: integer;
    ProjectFile: string;
  end;

var
  MainForm: TMainForm;
  RPath, sLang: string;
  Metadata: TJsonNode;

implementation

{$R *.lfm}

uses
  ChildFrame,
  child,
  about,
  info,
  filter,
  diversity,
  cluster,
  pca,
  pco,
  nmds,
  ca,
  dca,
  rda,
  cca,
  export,
  report,
  useful,
  dialog;

resourcestring
  { Main }
  strOpenFile = 'Abrir arquivo existente';
  strSaveFile = 'Salvar arquivo como';
  strJsonFilter = 'Projetos (*.json)|*.json';
  strHtmlFilter = 'Relatórios (*.htm *.html)|*.htm;*html';
  strCSVFilter = 'CSV (separado por vírgulas) (*.csv)|*.csv';
  strSheetFilter =
    'Planilhas (*.csv *.ods *.tsv *.xls *.xlsx)|*.csv;*.ods;*.tsv;*.xls;*.xlsx';
  strEmlFilter = 'Ecological Metadata Language (*.eml)|*.eml';
  strKmlFilter = 'Keyhole Markup Language (*.kml)|*.kml';
  strSHPFilter = 'Shapefiles (*.shp)|*.shp';
  strDbfFilter = 'dBase (*.dbf)|*.dbf';
  strFitopac1Filter = 'Arquivos Fitopac 1 (*.dad)|*.dad';
  strFitopac2Filter = 'Arquivos Fitopac 2 (*.fpd)|*.fpd';
  strCEPFilter = 'Cornell Ecology Programs (*.dta)|*.dta';
  strMVSPFilter = 'Multivariate Statistical Package (*.mvs)|*.mvs';
  strFPMFilter = 'Matriz Fitopac (*.fpm)|*.fpm';
  strInformation = 'Informação';
  strConfirmation = 'Confirmação';
  strError = 'Erro';
  strWarning = 'Aviso';
  strQuit = 'Deseja encerrar o programa?';
  strSearchCaption = 'Pesquisar';
  strSearchPrompt = 'Digite o texto a ser pesquisado:';
  //strTextFound = '%d ocorrência(s) encontrada(s)';
  strTextNotFound = 'Texto não encontrado';
  strDataNotFound = 'O projeto não contém arquivos de dados';
  strVariablesNotFound = 'O projeto não contém uma planilha de variáveis';
  strNoDescriptors = 'Não há descritores na planilha';
  strNoVariables = 'Não há variáveis na planilha';
  strNoSequences = 'Não há sequências na planilha';
  strNoCoordinates = 'Não há coordenadas geográficas na planilha';
  strFilter = 'Deseja filtrar os registros?';
  strDataSheet = 'Planilha de Dados';
  strVarSheet = 'Planilha de Variáveis';
  strPromptSheet = 'Planilha';
  strAddSheet = 'Adicionar';
  strDelSheet = 'Remover';
  strConnection = 'Não foi detectada uma conexão com a Internet';
  strGraph = 'Deseja construir gráfico?';
  strRDirectory = 'Selecione diretório do R';
  strRNotFound = 'R não instalado ou não encontrado.';
  strNotExecute = 'Erro na execução do programa %s.';
  strPlot = 'parcela';
  strQuadrat = 'quadrante';
  strInvalidMethod = 'Tipo de levantamento deve ser parcela ou quadrante';
  strAdd = 'Deseja adicionar planilhas ao projeto?';
  strColError = 'Número de oolunas da planilha deve ser maior ou igual a 4';
  strInvalidChars = 'Caracter inválido no nome da coluna %d (%s)';
  strNoFields = 'Campos insuficientes para geração do relatório';
  strWindowsOnly = 'Esta opção está disponível apenas para MS-Windows.';

  { Export }
  strExport = 'Exportar';
  strFormat = 'Formato:';
  strEML = 'Metadados em EML';
  strFITOPAC1 = 'Arquivos do Fitopac 1';
  strFITOPAC2 = 'Arquivo do Fitopac 2';
  strKML = 'Coordenadas em KML';
  strSHP = 'Shapefile ESRI';
  strBRAHMS = 'Arquivo RDE/BRAHMS';
  strCEP = 'Matriz CEP';
  strMVSP = 'Matriz MVSP';
  strFPM = 'Matriz Fitopac';
  strCSV = 'Matriz CSV';
  strExportMetadata = 'Metadados gravados no arquivo: %s';
  strExportData = 'Dados gravados no arquivo: %s para %d amostras(s) e %d espécie(s)';
  strExportRecords = '%d registros gravados no arquivo: %s';
  strExportFitopac1 =
    'Dados gravados no arquivo: %s para %d família(s), %d espécie(s) e %d indivíduo(s)';
  strExportFitopac2 =
    'Dados gravados no arquivo: %s para %d família(s), %d espécie(s), %d indivíduo(s) e %d amostras(s)';

  { Report }
  strReport = 'Relatório Estatístico';
  strStatistics = 'Estatísticas:';
  strFamilies = 'Famílias';
  strGenera = 'Gêneros';
  strSpecies = 'Espécies';
  strSites = 'Locais';
  strDescriptors = 'Descritores';
  strSequences = 'Sequências';
  strVariables = 'Variáveis';
  strStandard = 'Padrão';
  strRoman = 'Romano';
  strFull = 'Extenso';
  strDateFormat = 'Formato da Data';
  strLabelFormat = 'Formato:';

{ TMainForm }
function TMainForm.CheckData(SrcFileName: string): boolean;
var
  workbook: TsWorkbook;
  worksheet: TsWorksheet;
  i, j, n: integer;
  s: string;
  invalidCh: boolean;
begin
  workbook := TsWorkBook.Create;
  try
    workbook.ReadFromFile(SrcFileName);
    worksheet := workbook.GetFirstWorksheet;
    n := worksheet.GetLastColIndex(True);
    if n < 4 then
    begin
      MessageDlg(strError, strColError, mtError, [mbOK], 0);
      Result := False;
    end
    else
    begin
      Result := True;
      for j := 0 to n do
      begin
        s := worksheet.ReadAsText(0, j);
        if s <> '' then
        begin
          invalidCh := False;
          for i := 1 to length(s) do
            if s[i] > #127 then
              invalidCh := True;
          if invalidCh then
          begin
            MessageDlg(strError, Format(strInvalidChars, [j, s]), mtError, [mbOK], 0);
            Result := False;
          end;
        end;
      end;
    end;
  finally
    workbook.Free;
  end;
end;

function TMainForm.LocateR: boolean;
begin
  if not FileExists(RPath) then
    RPath := '';
  if RPath = '' then
  begin
    if Copy(OSVersion, 1, Pos(' ', OSVersion) - 1) <> 'Windows' then
      RPath := '/usr/bin/Rscript'
    else
    begin
      Screen.Cursor := crHourGlass;
      RPath := FindR;
      if RPath <> '' then
        RPath := Concat(RPath, 'Rscript.exe');
      Screen.Cursor := crDefault;
    end;
    if RPath = '' then
    begin
      MessageDlg(strError, Format(strRNotFound, ['R']), mtError, [mbOK], 0);
      Result := False;
    end
    else
      SaveSettings;
  end
  else
    Result := True;
end;

function TMainForm.FindR: string;
const
  FN = 'bin\RScript.exe';
  P = 'C:\Program Files\R';

var
  I: integer;
  LSearchResult: TStringList;

  procedure FileSearch(const dirName: string);
  var
    searchResult: TSearchRec;
  begin
    if FindFirst(dirName + '\*', faAnyFile, searchResult) = 0 then
    begin
      try
        repeat
          if (searchResult.Attr and faDirectory) = 0 then
          begin
            if SameText(ExtractFileExt(searchResult.Name), '.exe') then
              LSearchResult.Append(IncludeTrailingBackSlash(dirName) +
                searchResult.Name);
          end
          else if (searchResult.Name <> '.') and (searchResult.Name <> '..') then
            FileSearch(IncludeTrailingBackSlash(dirName) + searchResult.Name);
        until FindNext(searchResult) <> 0;
      finally
        FindClose(searchResult);
      end;
    end;
  end;

begin
  Result := '';
  LSearchResult := TStringList.Create;
  FileSearch(P);
  for I := 0 to LSearchResult.Count - 1 do
  begin
    if Pos(LowerCase(FN), LowerCase(LSearchResult[I])) > 0 then
      Result := ExtractFilePath(LSearchResult[I]);
  end;
  LSearchResult.Free;
end;

procedure TMainForm.CreateMDIChild(const AName: string);
var
  NewChild: TChildFrame;
  ChildFrame: TMDIChild;
begin
  NewChild := MultiDoc.NewChild('NewChild');
  ChildFrame := TMDIChild.Create(NewChild);
  NewChild.DockedObject := ChildFrame;
  NewChild.Caption := AName;
  Inc(ChildCount);
  MultiDoc.Cascade;
end;

procedure TMainForm.UpdateMenuItems(Sender: TObject);
begin
  if ChildCount = 0 then
    MainForm.Caption := Application.Title;
  FileChangeItem.Enabled := ChildCount > 0;
  FileReloadItem.Enabled := ChildCount > 0;
  FileCloseItem.Enabled := ChildCount > 0;
  FileAddItem.Enabled := ChildCount > 0;
  FileRemoveItem.Enabled := ChildCount > 0;
  FileExportItem.Enabled := ChildCount > 0;
  SearchSortItem.Enabled := ChildCount > 0;
  SearchFindItem.Enabled := ChildCount > 0;
  SearchFilterItem.Enabled := ChildCount > 0;
  with MultiDoc.ActiveObject as TMDIChild do
    SearchFilterResetItem.Enabled := (ChildCount > 0) and (Dataset1.Filtered);
  ReportCatalogItem.Enabled := ChildCount > 0;
  ReportLabelsItem.Enabled := ChildCount > 0;
  ReportChecklistItem.Enabled := ChildCount > 0;
  ReportStatisticsItem.Enabled := ChildCount > 0;
  ReportNamesItem.Enabled := ChildCount > 0;
  ReportGeocodeItem.Enabled := ChildCount > 0;
  AnalysisDiversityItem.Enabled := ChildCount > 0;
  AnalysisClusterItem.Enabled := ChildCount > 0;
  AnalysisOrdinationItem.Enabled := ChildCount > 0;
  WindowPreviousItem.Enabled := ChildCount > 0;
  WindowNextItem.Enabled := ChildCount > 0;
  WindowCascadeItem.Enabled := ChildCount > 0;
  WindowTileHorizontalItem.Enabled := ChildCount > 0;
  WindowTileVerticalItem.Enabled := ChildCount > 0;
  WindowCloseItem.Enabled := ChildCount > 0;
  N7.Visible := ChildCount > 0;
end;

procedure TMainForm.UpdateTitleBar(Filename: string);
begin
  if Length(Filename) > 0 then
    MainForm.Caption := Application.Title + ' - ' + Filename
  else
    MainForm.Caption := Application.Title;
end;

procedure TMainForm.RestoreSettings;
var
  sPath: string;
  IniFile: TIniFile;
begin
  sPath := GetAppConfigDir(False);
  IniFile := TIniFile.Create(sPath + 'ecolog.ini');
  sLang := IniFile.ReadString('Options', 'Language', 'pt_br');
  SetDefaultLang(sLang, 'languages', True);
  case sLang of
    'en':
    begin
      LanguageMenuEnglishItem.Checked := True;
      LanguageMenuPortugueseItem.Checked := False;
      LanguageMenuSpanishItem.Checked := False;
    end;
    'es':
    begin
      LanguageMenuSpanishItem.Checked := True;
      LanguageMenuEnglishItem.Checked := False;
      LanguageMenuPortugueseItem.Checked := False;
    end;
    'pt_br':
    begin
      LanguageMenuPortugueseItem.Checked := True;
      LanguageMenuEnglishItem.Checked := False;
      LanguageMenuSpanishItem.Checked := False;
    end;
  end;
  RPath := IniFile.ReadString('Options', 'RPath', '');
  Left := IniFile.ReadInteger('MainWindow', 'Left', Left);
  Top := IniFile.ReadInteger('MainWindow', 'Top', Top);
  Width := IniFile.ReadInteger('MainWindow', 'Width', Width);
  Height := IniFile.ReadInteger('MainWindow', 'Height', Height);
  WindowState := TWindowState(IniFile.ReadInteger('MainWindow', 'State',
    integer(WindowState)));
  IniFile.Free;
end;

procedure TMainForm.SaveSettings;
var
  sPath: string;
  IniFile: TIniFile;
begin
  sPath := GetAppConfigDir(False);
  IniFile := TIniFile.Create(sPath + 'ecolog.ini');
  IniFile.WriteString('Options', 'Language', sLang {GetDefaultLang});
  IniFile.WriteString('Options', 'RPath', RPath);
  IniFile.WriteInteger('MainWindow', 'Left', Left);
  IniFile.WriteInteger('MainWindow', 'Top', Top);
  IniFile.WriteInteger('MainWindow', 'Width', Width);
  IniFile.WriteInteger('MainWindow', 'Height', Height);
  IniFile.WriteInteger('MainWindow', 'State', integer(WindowState));
  IniFile.Free;
end;

procedure TMainForm.GridToArray(const filename: string; separator: char;
  var n, m: integer);
type
  TIntMatrix = array of array of integer;
var
  DatasetName1: string;
  lSPECIES, lSAMPLE: TStringList;
  fSPECIES, fSAMPLE: TField;
  bm: TBookmark;
  sums: TIntMatrix = nil;
  idxSPECIES, idxSAMPLE, r, c: integer;
  outfile: TextFile;
begin
  with MultiDoc.ActiveObject as TMDIChild do
  begin
    if not Dataset1.Active then
    begin
      DatasetName1 := Metadata.Find('dataset1').AsString;
      Dataset1.FileName := DatasetName1;
      Dataset1.AutoFieldDefs := True;
      Dataset1.Open;
    end;
    bm := Dataset1.GetBookmark;
    Dataset1.DisableControls;
    lSPECIES := TStringList.Create;
    lSAMPLE := TStringList.Create;
    try
      lSPECIES.Sorted := True;
      lSPECIES.Duplicates := dupIgnore;
      lSAMPLE.Sorted := True;
      lSAMPLE.Duplicates := dupIgnore;
      fSAMPLE := Dataset1.Fields[0];
      fSPECIES := Dataset1.Fields[3];

      Dataset1.First;
      while not Dataset1.EOF do
      begin
        lSPECIES.Add(fSPECIES.AsString);
        lSAMPLE.Add(fSAMPLE.AsString);
        Dataset1.Next;
      end;

      SetLength(sums, lSPECIES.Count, lSAMPLE.Count);

      Dataset1.First;
      while not Dataset1.EOF do
      begin
        lSPECIES.Find(fSPECIES.AsString, idxSPECIES);
        lSAMPLE.Find(fSAMPLE.AsString, idxSAMPLE);
        Inc(sums[idxSPECIES, idxSAMPLE]);
        Dataset1.Next;
      end;

      AssignFile(outfile, filename);
      Rewrite(outfile);
      Write(outfile, 'SPECIES,');
      for c := 1 to lSAMPLE.Count do
        if c < lSAMPLE.Count then
          Write(outfile, lSAMPLE[c - 1], separator)
        else
          Write(outfile, lSAMPLE[c - 1]);
      WriteLn(outfile);
      for r := 1 to lSPECIES.Count do
      begin
        Write(outfile, lSPECIES[r - 1], separator);
        for c := 1 to lSAMPLE.Count do
          if c < lSAMPLE.Count then
            Write(outfile, sums[r - 1, c - 1], separator)
          else
            Write(outfile, sums[r - 1, c - 1]);
        WriteLn(outfile);
      end;
      CloseFile(outfile);

      n := lSAMPLE.Count;
      m := lSPECIES.Count;

    finally
      lSPECIES.Free;
      lSAMPLE.Free;
      Dataset1.EnableControls;
      Dataset1.GotoBookmark(bm);
    end;
  end;
end;

procedure TMainForm.WindowTileHorizontalItemClick(Sender: TObject);
begin
  MultiDoc.TileHorizontal;
end;

procedure TMainForm.WindowTileVerticalItemClick(Sender: TObject);
begin
  MultiDoc.TileVertical;
end;

procedure TMainForm.FileNewItemClick(Sender: TObject);
begin
  SaveDialog.Title := strSaveFile;
  SaveDialog.DefaultExt := '.json';
  SaveDialog.Filter := strJsonFilter;
  saveDialog.FileName := '';
  if ProjectDlg.ShowModal = mrOk then
  begin
    if SaveDialog.Execute then
    begin
      ProjectFile := SaveDialog.Filename;
      if MessageDlg(strConfirmation, strAdd, mtConfirmation, [mbYes, mbNo], 0) =
        mrYes then
      begin
        FileAddItemClick(Sender);
        UpdateMenuItems(Self);
        UpdateTitleBar(ProjectFile);
      end;
    end;
  end;
end;

procedure TMainForm.WindowCloseItemClick(Sender: TObject);
begin
  MultiDoc.ActiveChild.Close;
end;

procedure TMainForm.WindowNextItemClick(Sender: TObject);
begin
  MultiDoc.NexChild;
end;

procedure TMainForm.WindowPreviousItemClick(Sender: TObject);
begin
  MultiDoc.NexChild;
end;

procedure TMainForm.WindowCascadeItemClick(Sender: TObject);
begin
  MultiDoc.Cascade;
end;

procedure TMainForm.MultiDocDeleteChild(Sender: TObject);
begin
  Dec(ChildCount);
  UpdateMenuItems(Self);
end;

procedure TMainForm.OrdinationCAItemClick(Sender: TObject);
var
  n, m: integer;
  s: ansistring;
begin
  if not LocateR then
    Exit;
  SaveDialog.Title := strSaveFile;
  SaveDialog.DefaultExt := '.htm';
  SaveDialog.Filter := strHtmlFilter;
  SaveDialog.Filename := '';
  if SaveDialog.Execute then
  begin
    MainForm.MultiDoc.SetActiveChild(0);
    with CADlg do
    begin
      if ShowModal = mrOk then
      begin
        GridToArray('rdata.csv', ',', n, m);
        CreateCA(SaveDialog.Filename, ComboBoxTransform.ItemIndex,
          ComboBoxScale.ItemIndex);
        Screen.Cursor := crHourGlass;
        if RunCommand(RPath, ['--vanilla', 'ca.R'], s, [poNoConsole]) then
        begin
          CA(SaveDialog.Filename, ComboBoxTransform.ItemIndex, n, m);
          CreateMDIChild(ExtractFileName(SaveDialog.Filename));
          with MultiDoc.ActiveObject as TMDIChild do
          begin
            DataGrid.Visible := False;
            HtmlViewer.Visible := True;
            HtmlViewer.LoadFromFile(SaveDialog.Filename);
          end;
          UpdateTitleBar(ProjectFile);
        end
        else
        begin
          Screen.Cursor := crDefault;
          MessageDlg(strError, Format(strNotExecute, ['R']), mtError, [mbOK], 0);
        end;
        Screen.Cursor := crDefault;
      end;
    end;
  end;
end;

procedure TMainForm.OrdinationCCAItemClick(Sender: TObject);
var
  n, m, i, nvars: integer;
  s: ansistring;
  DatasetName2, selected, SqlStr: string;
  SAMPLE: string;
  VarList: TStringList;
  f: TField;
begin
  if Length(Metadata.Find('dataset2').AsString) = 0 then
  begin
    MessageDlg(strWarning, strVariablesNotFound, mtWarning, [mbOK], 0);
    Exit;
  end;
  if not LocateR then
    Exit;
  SaveDialog.Title := strSaveFile;
  SaveDialog.DefaultExt := '.htm';
  SaveDialog.Filter := strHtmlFilter;
  SaveDialog.Filename := '';
  if SaveDialog.Execute then
  begin
    VarList := TStringList.Create;
    MainForm.MultiDoc.SetActiveChild(1);
    with CCADlg do
    begin
      if ShowModal = mrOk then
      begin
        with MainForm.MultiDoc.ActiveObject as TMDIChild do
        begin
          if not Dataset2.Active then
          begin
            DatasetName2 := Metadata.Find('dataset2').AsString;
            Dataset2.FileName := DatasetName2;
            Dataset2.AutoFieldDefs := True;
            Dataset2.Open;
          end;
          for i := 1 to Dataset2.FieldCount do
            VarList.Add(Dataset2.Fields[i - 1].FieldName);
          VarList.Delete(0);
          nvars := VarList.Count;
          selected := CheckListDialog(strVariables, '', VarList, True);
          selected := DelChars(selected, '''');
          VarList.Free;
          if selected = '' then
            Exit;
          CSVExporter2.FileName := 'rdata2.csv';
          CSVExporter2.Execute;
          QueryConnection1.DatabasePath := GetCurrentDir;
          QueryConnection1.Connect;
          QueryDataSet1.TableName := 'rdata2';
          QueryDataset1.LoadFromTable;
          SAMPLE := Dataset2.Fields[0].FieldName;
          SqlStr := 'SELECT ' + SAMPLE + ', ';
          for f in Dataset2.Fields do
            if f is TFloatField or f is TIntegerField then
              SqlStr += 'AVG(' + f.FieldName + '), ';
          Delete(SqlStr, Length(SqlStr) - 1, 1);
          SqlStr += 'FROM rdata2 GROUP BY ' + SAMPLE + ' ORDER BY ' + SAMPLE;
          QueryDataset1.SQL.Text := SqlStr;
          QueryDataset1.QueryExecute;
          QueryDataset1.SaveToTable;
        end;
        GridToArray('rdata1.csv', ',', n, m);
        CreateCCA(SaveDialog.Filename, ComboBoxTransform.ItemIndex,
          nvars, CheckBoxScale.Checked, selected);
        Screen.Cursor := crHourGlass;
        if RunCommand(RPath, ['--vanilla', 'cca.R'], s, [poNoConsole]) then
        begin
          CCA(SaveDialog.Filename, ComboBoxTransform.ItemIndex, n, m, selected);
          CreateMDIChild(ExtractFileName(SaveDialog.Filename));
          with MultiDoc.ActiveObject as TMDIChild do
          begin
            DataGrid.Visible := False;
            HtmlViewer.Visible := True;
            HtmlViewer.LoadFromFile(SaveDialog.Filename);
          end;
          UpdateTitleBar(ProjectFile);
        end
        else
        begin
          Screen.Cursor := crDefault;
          MessageDlg(strError, Format(strNotExecute, ['R']), mtError, [mbOK], 0);
        end;
        Screen.Cursor := crDefault;
      end;
    end;
  end;
end;

procedure TMainForm.OrdinationNMDSItemClick(Sender: TObject);
var
  n, m: integer;
  s: ansistring;
begin
  if not LocateR then
    Exit;
  SaveDialog.Title := strSaveFile;
  SaveDialog.DefaultExt := '.htm';
  SaveDialog.Filter := strHtmlFilter;
  SaveDialog.Filename := '';
  if SaveDialog.Execute then
  begin
    MainForm.MultiDoc.SetActiveChild(0);
    with NMDSDlg do
    begin
      if ShowModal = mrOk then
      begin
        GridToArray('rdata.csv', ',', n, m);
        CreateNMDS(SaveDialog.Filename, ComboBoxTransform.ItemIndex,
          ComboBoxCoef.ItemIndex, SpinEditIter.Value, ComboBoxConfig.ItemIndex);
        Screen.Cursor := crHourGlass;
        if RunCommand(RPath, ['--vanilla', 'nmds.R'], s, [poNoConsole]) then
        begin
          NMDS(SaveDialog.Filename, ComboBoxTransform.ItemIndex,
            ComboBoxCoef.ItemIndex, SpinEditIter.Value, ComboBoxConfig.ItemIndex, n, m);
          CreateMDIChild(ExtractFileName(SaveDialog.Filename));
          with MultiDoc.ActiveObject as TMDIChild do
          begin
            DataGrid.Visible := False;
            HtmlViewer.Visible := True;
            HtmlViewer.LoadFromFile(SaveDialog.Filename);
          end;
          UpdateTitleBar(ProjectFile);
        end
        else
        begin
          Screen.Cursor := crDefault;
          MessageDlg(strError, Format(strNotExecute, ['R']), mtError, [mbOK], 0);
        end;
        Screen.Cursor := crDefault;
      end;
    end;
  end;
end;

procedure TMainForm.OrdinationPCAItemClick(Sender: TObject);
var
  n, m: integer;
  s: ansistring;
begin
  if not LocateR then
    Exit;
  SaveDialog.Title := strSaveFile;
  SaveDialog.DefaultExt := '.htm';
  SaveDialog.Filter := strHtmlFilter;
  SaveDialog.Filename := '';
  if SaveDialog.Execute then
  begin
    MainForm.MultiDoc.SetActiveChild(0);
    with PCADlg do
    begin
      if ShowModal = mrOk then
      begin
        GridToArray('rdata.csv', ',', n, m);
        CreatePCA(SaveDialog.Filename, ComboBoxTransform.ItemIndex,
          ComboBoxCoef.ItemIndex, CheckBoxCenter.Checked, CheckBoxStand.Checked);
        Screen.Cursor := crHourGlass;
        if RunCommand(RPath, ['--vanilla', 'pca.R'], s, [poNoConsole]) then
        begin
          PCA(SaveDialog.Filename, ComboBoxTransform.ItemIndex,
            ComboBoxCoef.ItemIndex, CheckBoxCenter.Checked, CheckBoxStand.Checked, n, m);
          CreateMDIChild(ExtractFileName(SaveDialog.Filename));
          with MultiDoc.ActiveObject as TMDIChild do
          begin
            DataGrid.Visible := False;
            HtmlViewer.Visible := True;
            HtmlViewer.LoadFromFile(SaveDialog.Filename);
          end;
          UpdateTitleBar(ProjectFile);
        end
        else
        begin
          Screen.Cursor := crDefault;
          MessageDlg(strError, Format(strNotExecute, ['R']), mtError, [mbOK], 0);
        end;
        Screen.Cursor := crDefault;
      end;
    end;
  end;
end;

procedure TMainForm.OrdinationPCOItemClick(Sender: TObject);
var
  n, m: integer;
  s: ansistring;
begin
  if not LocateR then
    Exit;
  SaveDialog.Title := strSaveFile;
  SaveDialog.DefaultExt := '.htm';
  SaveDialog.Filter := strHtmlFilter;
  SaveDialog.Filename := '';
  if SaveDialog.Execute then
  begin
    MainForm.MultiDoc.SetActiveChild(0);
    with PCODlg do
    begin
      if ShowModal = mrOk then
      begin
        GridToArray('rdata.csv', ',', n, m);
        CreatePCOA(SaveDialog.Filename, ComboBoxTransform.ItemIndex,
          ComboBoxCoef.ItemIndex);
        Screen.Cursor := crHourGlass;
        if RunCommand(RPath, ['--vanilla', 'pco.R'], s, [poNoConsole]) then
        begin
          PCOA(SaveDialog.Filename, ComboBoxTransform.ItemIndex,
            ComboBoxCoef.ItemIndex, n, m);
          CreateMDIChild(ExtractFileName(SaveDialog.Filename));
          with MultiDoc.ActiveObject as TMDIChild do
          begin
            DataGrid.Visible := False;
            HtmlViewer.Visible := True;
            HtmlViewer.LoadFromFile(SaveDialog.Filename);
          end;
          UpdateTitleBar(ProjectFile);
        end
        else
        begin
          Screen.Cursor := crDefault;
          MessageDlg(strError, Format(strNotExecute, ['R']), mtError, [mbOK], 0);
        end;
        Screen.Cursor := crDefault;
      end;
    end;
  end;
end;

procedure TMainForm.OrdinationRDAItemClick(Sender: TObject);
var
  n, m, i, nvars: integer;
  s: ansistring;
  DatasetName2, selected, SQLStr: string;
  SAMPLE: string;
  VarList: TStringList;
  f: TField;
begin
  if Length(Metadata.Find('dataset2').AsString) = 0 then
  begin
    MessageDlg(strWarning, strVariablesNotFound, mtWarning, [mbOK], 0);
    Exit;
  end;
  if not LocateR then
    Exit;
  SaveDialog.Title := strSaveFile;
  SaveDialog.DefaultExt := '.htm';
  SaveDialog.Filter := strHtmlFilter;
  SaveDialog.Filename := '';
  if SaveDialog.Execute then
  begin
    VarList := TStringList.Create;
    MainForm.MultiDoc.SetActiveChild(1);
    with RDADlg do
    begin
      if ShowModal = mrOk then
      begin
        with MainForm.MultiDoc.ActiveObject as TMDIChild do
        begin
          if not Dataset2.Active then
          begin
            DatasetName2 := Metadata.Find('dataset2').AsString;
            Dataset2.FileName := DatasetName2;
            Dataset2.AutoFieldDefs := True;
            Dataset2.Open;
          end;
          for i := 1 to Dataset2.FieldCount do
            VarList.Add(Dataset2.Fields[i - 1].FieldName);
          VarList.Delete(0);
          nvars := VarList.Count;
          selected := CheckListDialog(strVariables, '', VarList, True);
          selected := DelChars(selected, '''');
          VarList.Free;
          if selected = '' then
            Exit;
          CSVExporter2.FileName := 'rdata2.csv';
          CSVExporter2.Execute;
          QueryConnection1.DatabasePath := GetCurrentDir;
          QueryConnection1.Connect;
          QueryDataSet1.TableName := 'rdata2';
          QueryDataset1.LoadFromTable;
          SAMPLE := Dataset2.Fields[0].FieldName;
          SqlStr := 'SELECT ' + SAMPLE + ', ';
          for f in Dataset2.Fields do
            if f is TFloatField or f is TIntegerField then
              SqlStr += 'AVG(' + f.FieldName + '), ';
          Delete(SqlStr, Length(SqlStr) - 1, 1);
          SqlStr += 'FROM rdata2 GROUP BY ' + SAMPLE + ' ORDER BY ' + SAMPLE;
          QueryDataset1.SQL.Text := SqlStr;
          QueryDataset1.QueryExecute;
          QueryDataset1.SaveToTable;
        end;
        GridToArray('rdata1.csv', ',', n, m);
        CreateRDA(SaveDialog.Filename, ComboBoxTransform.ItemIndex,
          ComboBoxScale.ItemIndex, ComboBoxInertia.ItemIndex, nvars, selected);
        Screen.Cursor := crHourGlass;
        if RunCommand(RPath, ['--vanilla', 'rda.R'], s, [poNoConsole]) then
        begin
          RDA(SaveDialog.Filename, ComboBoxTransform.ItemIndex,
            ComboBoxInertia.ItemIndex, n, m, selected);
          CreateMDIChild(ExtractFileName(SaveDialog.Filename));
          with MultiDoc.ActiveObject as TMDIChild do
          begin
            DataGrid.Visible := False;
            HtmlViewer.Visible := True;
            HtmlViewer.LoadFromFile(SaveDialog.Filename);
          end;
          UpdateTitleBar(ProjectFile);
        end
        else
        begin
          Screen.Cursor := crDefault;
          MessageDlg(strError, Format(strNotExecute, ['R']), mtError, [mbOK], 0);
        end;
        Screen.Cursor := crDefault;
      end;
    end;
  end;
end;

procedure TMainForm.SearchFilterResetItemClick(Sender: TObject);
begin
  with MultiDoc.ActiveObject as TMDIChild do
  begin
    Dataset1.Filter := '';
    Dataset1.Filtered := False;
    Dataset1.First;
  end;
  UpdateMenuItems(Self);
end;

procedure TMainForm.SearchSortItemClick(Sender: TObject);
var
  Selected: string;
begin
  with MainForm.MultiDoc.ActiveObject as TMDIChild do
  begin
    Selected := Dataset1.Fields[DataGrid.SelectedIndex].FieldName;
    Dataset1.SortOnFields(Selected);
  end;
end;

procedure TMainForm.AnalysisClusterItemClick(Sender: TObject);
var
  n, m: integer;
  s: ansistring;
begin
  if not LocateR then
    Exit;
  SaveDialog.Title := strSaveFile;
  SaveDialog.DefaultExt := '.htm';
  SaveDialog.Filter := strHtmlFilter;
  SaveDialog.Filename := '';
  if SaveDialog.Execute then
  begin
    MainForm.MultiDoc.SetActiveChild(0);
    with ClusterDlg do
    begin
      if ShowModal = mrOk then
      begin
        GridToArray('rdata.csv', ',', n, m);
        CreateCluster(SaveDialog.Filename, ComboBoxTransform.ItemIndex,
          ComboBoxCoef.ItemIndex,
          ComboBoxMethod.ItemIndex);
        Screen.Cursor := crHourGlass;
        if RunCommand(RPath, ['--vanilla', 'cluster.R'], s, [poNoConsole]) then
        begin
          Cluster(SaveDialog.Filename, ComboBoxTransform.ItemIndex,
            ComboBoxCoef.ItemIndex,
            ComboBoxMethod.ItemIndex, n, m);
          CreateMDIChild(ExtractFileName(SaveDialog.Filename));
          with MultiDoc.ActiveObject as TMDIChild do
          begin
            DataGrid.Visible := False;
            HtmlViewer.Visible := True;
            HtmlViewer.LoadFromFile(SaveDialog.Filename);
          end;
          UpdateTitleBar(ProjectFile);
        end
        else
        begin
          Screen.Cursor := crDefault;
          MessageDlg(strError, Format(strNotExecute, ['R']), mtError, [mbOK], 0);
        end;
        Screen.Cursor := crDefault;
      end;
    end;
  end;
end;

procedure TMainForm.AnalysisDiversityItemClick(Sender: TObject);
var
  n, m: integer;
  s: ansistring;
begin
  if not LocateR then
    Exit;
  SaveDialog.Title := strSaveFile;
  SaveDialog.DefaultExt := '.htm';
  SaveDialog.Filter := strHtmlFilter;
  SaveDialog.Filename := '';
  if SaveDialog.Execute then
  begin
    MainForm.MultiDoc.SetActiveChild(0);
    GridToArray('rdata.csv', ',', n, m);
    CreateDivers(SaveDialog.Filename);
    Screen.Cursor := crHourGlass;
    if RunCommand(RPath, ['--vanilla', 'diversity.R'], s, [poNoConsole]) then
    begin
      Divers(SaveDialog.Filename, n, m);
      CreateMDIChild(ExtractFileName(SaveDialog.Filename));
      with MultiDoc.ActiveObject as TMDIChild do
      begin
        DataGrid.Visible := False;
        HtmlViewer.Visible := True;
        HtmlViewer.LoadFromFile(SaveDialog.Filename);
      end;
      UpdateTitleBar(ProjectFile);
    end
    else
    begin
      Screen.Cursor := crDefault;
      MessageDlg(strError, Format(strNotExecute, ['R']), mtError, [mbOK], 0);
    end;
    Screen.Cursor := crDefault;
  end;
end;

procedure TMainForm.FileAddItemClick(Sender: TObject);
var
  options: TStringList;
  vResult: string;
  index: integer;
  f: TField;
begin
  options := TStringList.Create;
  try
    options.Add(strDataSheet);
    options.Add(strVarSheet);
    vResult := InputCombo(strAddSheet, strPromptSheet, options);
    index := options.IndexOf(vResult);
  finally
    options.Free;
  end;
  if index <> -1 then
  begin
    OpenDialog.Title := strOpenFile;
    OpenDialog.DefaultExt := '.xlsx';
    OpenDialog.Filter := strSheetFilter;
    OpenDialog.FileName := '';
    if OpenDialog.Execute then
    begin
      if index = 0 then
      begin
        if not CheckData(OpenDialog.FileName) then
          Exit;
        Metadata.Force('dataset1').AsString := ExtractFileName(OpenDialog.FileName);
        CreateMDIChild(Metadata.Find('dataset1').AsString);
        with MultiDoc.ActiveObject as TMDIChild do
        begin
          DataGrid.DataSource := DataSource1;
          Dataset1.FileName := OpenDialog.FileName;
          Dataset1.AutoFieldDefs := True;
          Dataset1.Open;
          for f in Dataset1.Fields do
            if f is TFloatField then
              TFloatField(f).DisplayFormat := '0.###';
          DataGrid.Visible := True;
          HtmlViewer.Visible := False;
        end;
      end
      else
      begin
        Metadata.Force('dataset2').AsString := ExtractFileName(OpenDialog.FileName);
        CreateMDIChild(Metadata.Find('dataset2').AsString);
        with MultiDoc.ActiveObject as TMDIChild do
        begin
          DataGrid.DataSource := DataSource2;
          Dataset2.FileName := OpenDialog.FileName;
          Dataset2.AutoFieldDefs := True;
          Dataset2.Open;
          for f in Dataset2.Fields do
            if f is TFloatField then
              TFloatField(f).DisplayFormat := '0.###';
          DataGrid.Visible := True;
          HtmlViewer.Visible := False;
        end;
      end;
      Metadata.SaveToFile(ProjectFile);
    end;
  end;
end;

procedure TMainForm.FileChangeItemClick(Sender: TObject);
begin
  SaveDialog.Title := strSaveFile;
  SaveDialog.DefaultExt := '.json';
  SaveDialog.Filter := strJsonFilter;
  SaveDialog.FileName := ProjectFile;
  if ProjectDlg.ShowModal = mrOk then
  begin
    if SaveDialog.Execute then
      Metadata.SaveToFile(SaveDialog.Filename);
  end;
end;

procedure TMainForm.FileCloseItemClick(Sender: TObject);
var
  i: integer;
begin
  ProjectFile := '';
  Metadata.Clear;
  try
    for i := MultiDoc.ChildCount - 1 downto 0 do
      MultiDoc.Childs[i].Close;
    UpdateMenuItems(Self);
  except
  end;
end;

procedure TMainForm.FileReloadItemClick(Sender: TObject);
var
  i: integer;
  FName1, FName2: string;
  f: TField;
begin
  for i := MultiDoc.ChildCount - 1 downto 0 do
    MultiDoc.Childs[i].Close;
  FName1 := Metadata.Find('dataset1').AsString;
  CreateMDIChild(FName1);
  with MultiDoc.ActiveObject as TMDIChild do
  begin
    DataGrid.DataSource := DataSource1;
    Dataset1.FileName := FName1;
    Dataset1.AutoFieldDefs := True;
    Dataset1.Open;
    for f in Dataset1.Fields do
      if f is TFloatField then
        TFloatField(f).DisplayFormat := '0.###';
    DataGrid.Visible := True;
    HtmlViewer.Visible := False;
  end;
  FName2 := Metadata.Find('dataset2').AsString;
  if Length(FName2) > 0 then
  begin
    CreateMDIChild(FName2);
    with MultiDoc.ActiveObject as TMDIChild do
    begin
      DataGrid.DataSource := DataSource2;
      Dataset2.FileName := FName2;
      Dataset2.AutoFieldDefs := True;
      Dataset2.Open;
      for f in Dataset2.Fields do
        if f is TFloatField then
          TFloatField(f).DisplayFormat := '0.###';
      DataGrid.Visible := True;
      HtmlViewer.Visible := False;
    end;
  end;
  MultiDoc.Cascade;
  UpdateMenuItems(Self);
  MultiDoc.SetActiveChild(0);
end;

procedure TMainForm.FileRemoveItemClick(Sender: TObject);
var
  options: TStringList;
  vResult: string;
  index, i: integer;
begin
  options := TStringList.Create;
  try
    options.Add(strDataSheet);
    options.Add(strVarSheet);
    vResult := InputCombo(strDelSheet, strPromptSheet, options);
    index := options.IndexOf(vResult);
  finally
    options.Free;
  end;
  if index <> -1 then
  begin
    if index = 0 then
    begin
      for i := 0 to MultiDoc.ChildCount - 1 do
        if MultiDoc.Childs[i].Caption = Metadata.Find('dataset1').AsString then
          MultiDoc.Childs[i].Close;
      Metadata.Find('dataset1').AsString := '';
    end
    else
    begin
      for i := 0 to MultiDoc.ChildCount - 1 do
        if MultiDoc.Childs[i].Caption = Metadata.Find('dataset2').AsString then
          MultiDoc.Childs[i].Close;
      Metadata.Find('dataset2').AsString := '';
    end;
    Metadata.SaveToFile(ProjectFile);
  end;
end;

procedure TMainForm.FormCloseQuery(Sender: TObject; var CanClose: boolean);
begin
  if MessageDlg(strConfirmation, strQuit, mtConfirmation, [mbYes, mbNo], 0) =
    mrYes then
    CanClose := True
  else
    CanClose := False;
end;

procedure TMainForm.LanguageMenuEnglishItemClick(Sender: TObject);
var
  sPath: string;
  IniFile: TIniFile;
begin
  LanguageMenuEnglishItem.Checked := True;
  LanguageMenuPortugueseItem.Checked := False;
  LanguageMenuSpanishItem.Checked := False;
  sPath := GetAppConfigDir(False);
  IniFile := TIniFile.Create(sPath + 'ecolog.ini');
  SetDefaultLang('en', 'language', True);
  IniFile.WriteString('Options', 'Language', 'en');
  IniFile.Free;
end;

procedure TMainForm.LanguageMenuPortugueseItemClick(Sender: TObject);
var
  sPath: string;
  IniFile: TIniFile;
begin
  LanguageMenuPortugueseItem.Checked := True;
  LanguageMenuEnglishItem.Checked := False;
  LanguageMenuSpanishItem.Checked := False;
  sPath := GetAppConfigDir(False);
  IniFile := TIniFile.Create(sPath + 'ecolog.ini');
  SetDefaultLang('pt_br', 'language', True);
  IniFile.WriteString('Options', 'Language', 'pt_br');
  IniFile.Free;
end;

procedure TMainForm.LanguageMenuSpanishItemClick(Sender: TObject);
var
  sPath: string;
  IniFile: TIniFile;
begin
  LanguageMenuSpanishItem.Checked := True;
  LanguageMenuEnglishItem.Checked := False;
  LanguageMenuPortugueseItem.Checked := False;
  sPath := GetAppConfigDir(False);
  IniFile := TIniFile.Create(sPath + 'ecolog.ini');
  SetDefaultLang('es', 'language', True);
  IniFile.WriteString('Options', 'Language', 'es');
  IniFile.Free;
end;

procedure TMainForm.OrdinationDCAItemClick(Sender: TObject);
var
  n, m: integer;
  s: ansistring;
begin
  if not LocateR then
    Exit;
  SaveDialog.Title := strSaveFile;
  SaveDialog.DefaultExt := '.htm';
  SaveDialog.Filter := strHtmlFilter;
  SaveDialog.Filename := '';
  if SaveDialog.Execute then
  begin
    MainForm.MultiDoc.SetActiveChild(0);
    with DCADlg do
    begin
      if ShowModal = mrOk then
      begin
        GridToArray('rdata.csv', ',', n, m);
        CreateDCA(SaveDialog.Filename,
          StrToInt(IfThen(CheckBoxDetrended.Checked, '0', '1')),
          StrToInt(IfThen(CheckBoxDownweight.Checked, '1', '0')),
          SpinEditNSegs.Value,
          StrToInt(IfThen(CheckBoxRescaling.Checked, '0', '1')),
          SpinEditNRescaling.Value, ComboBoxScale.ItemIndex);
        Screen.Cursor := crHourGlass;
        if RunCommand(RPath, ['--vanilla', 'dca.R'], s, [poNoConsole]) then
        begin
          DCA(SaveDialog.Filename,
            StrToInt(IfThen(CheckBoxDetrended.Checked, '0', '1')),
            StrToInt(IfThen(CheckBoxDownweight.Checked, '1', '0')),
            SpinEditNSegs.Value,
            StrToInt(IfThen(CheckBoxRescaling.Checked, '0', '1')),
            SpinEditNRescaling.Value, n, m);
          CreateMDIChild(ExtractFileName(SaveDialog.Filename));
          with MultiDoc.ActiveObject as TMDIChild do
          begin
            DataGrid.Visible := False;
            HtmlViewer.Visible := True;
            HtmlViewer.LoadFromFile(SaveDialog.Filename);
          end;
          UpdateTitleBar(ProjectFile);
        end
        else
        begin
          Screen.Cursor := crDefault;
          MessageDlg(strError, Format(strNotExecute, ['R']), mtError, [mbOK], 0);
        end;
        Screen.Cursor := crDefault;
      end;
    end;
  end;
end;

procedure TMainForm.ReportCatalogItemClick(Sender: TObject);
var
  i: integer;
  {found: boolean;}
begin
  {found := True;
  with MainForm.MultiDoc.ActiveObject as TMDIChild do
  begin
    for i := 0 to Dataset1.FieldCount - 1 do
    begin
      if not AnsiMatchStr(Copy(UpperCase(Dataset1.Fields[i].FieldName), 1, 3),
        ['COL', 'NUM', 'DAT', 'LOC', 'LAT', 'LON', 'ALT']) then
        found := False;
    end;
  end;
  if not found then
  begin
    MessageDlg(strWarning, strNoFields, mtWarning, [mbOK], 0);
    Exit;
  end;}
  SaveDialog.Title := strSaveFile;
  SaveDialog.DefaultExt := '.htm';
  SaveDialog.Filter := strHtmlFilter;
  SaveDialog.Filename := '';
  if SaveDialog.Execute then
  begin
    MainForm.MultiDoc.SetActiveChild(0);
    with MainForm.MultiDoc.ActiveObject as TMDIChild do
    begin
      if not Dataset1.Filtered then
      begin
        if MessageDlg(strConfirmation, strFilter, mtConfirmation, [mbYes, mbNo], 0) =
          mrYes then
        begin
          if FilterDlg.ShowModal = mrOk then
          begin
            Dataset1.Filtered := False;
            Dataset1.Filter := FilterDlg.FFilter;
            Dataset1.Filtered := True;
            Dataset1.First;
          end;
        end;
      end;
    end;
    Catalog(SaveDialog.Filename);
    CreateMDIChild(ExtractFileName(SaveDialog.Filename));
    with MultiDoc.ActiveObject as TMDIChild do
    begin
      DataGrid.Visible := False;
      HtmlViewer.Visible := True;
      HtmlViewer.LoadFromFile(SaveDialog.Filename);
    end;
    UpdateTitleBar(ProjectFile);
  end;
end;

procedure TMainForm.ReportChecklistItemClick(Sender: TObject);
begin
  SaveDialog.Title := strSaveFile;
  SaveDialog.DefaultExt := '.htm';
  SaveDialog.Filter := strHtmlFilter;
  SaveDialog.Filename := '';
  if SaveDialog.Execute then
  begin
    MainForm.MultiDoc.SetActiveChild(0);
    with MainForm.MultiDoc.ActiveObject as TMDIChild do
    begin
      if not Dataset1.Filtered then
      begin
        if MessageDlg(strConfirmation, strFilter, mtConfirmation, [mbYes, mbNo], 0) =
          mrYes then
        begin
          if FilterDlg.ShowModal = mrOk then
          begin
            Dataset1.Filtered := False;
            Dataset1.Filter := FilterDlg.FFilter;
            Dataset1.Filtered := True;
            Dataset1.First;
          end;
        end;
      end;
    end;
    Geral(SaveDialog.Filename);
    CreateMDIChild(ExtractFileName(SaveDialog.Filename));
    with MultiDoc.ActiveObject as TMDIChild do
    begin
      DataGrid.Visible := False;
      HtmlViewer.Visible := True;
      HtmlViewer.LoadFromFile(SaveDialog.Filename);
    end;
    UpdateTitleBar(ProjectFile);
  end;
end;

procedure TMainForm.ReportGeocodeItemClick(Sender: TObject);
begin
  if not IsOnline then
  begin
    MessageDlg(strError, strConnection, mtError, [mbOK], 0);
    Exit;
  end;
  SaveDialog.Title := strSaveFile;
  SaveDialog.DefaultExt := '.htm';
  SaveDialog.Filter := strHtmlFilter;
  SaveDialog.Filename := '';
  if SaveDialog.Execute then
  begin
    MainForm.MultiDoc.SetActiveChild(0);
    with MainForm.MultiDoc.ActiveObject as TMDIChild do
    begin
      if not Dataset1.Filtered then
      begin
        if MessageDlg(strConfirmation, strFilter, mtConfirmation, [mbYes, mbNo], 0) =
          mrYes then
        begin
          if FilterDlg.ShowModal = mrOk then
          begin
            Dataset1.Filtered := False;
            Dataset1.Filter := FilterDlg.FFilter;
            Dataset1.Filtered := True;
            Dataset1.First;
          end;
        end;
      end;
    end;
    Georef(SaveDialog.Filename);
    CreateMDIChild(ExtractFileName(SaveDialog.Filename));
    with MultiDoc.ActiveObject as TMDIChild do
    begin
      DataGrid.Visible := False;
      HtmlViewer.Visible := True;
      HtmlViewer.LoadFromFile(SaveDialog.Filename);
    end;
    UpdateTitleBar(ProjectFile);
  end;
end;

procedure TMainForm.ReportLabelsItemClick(Sender: TObject);
var
  options: TStringList;
  vResult: string;
  format, i: integer;
  {found: boolean;}
begin
  {found := True;
  with MainForm.MultiDoc.ActiveObject as TMDIChild do
  begin
    for i := 0 to Dataset1.FieldCount - 1 do
    begin
      if not AnsiMatchStr(Copy(UpperCase(Dataset1.Fields[i].FieldName), 1, 3),
        ['COL', 'NUM', 'DAT', 'LOC', 'LAT', 'LON', 'ALT']) then
        found := False;
    end;
  end;
  if not found then
  begin
    MessageDlg(strWarning, strNoFields, mtWarning, [mbOK], 0);
    Exit;
  end;}
  options := TStringList.Create;
  try
    options.Add(strStandard);
    options.Add(strRoman);
    options.Add(strFull);
    vResult := InputCombo(strDateFormat, strLabelFormat, options);
    format := options.IndexOf(vResult);
  finally
    options.Free;
  end;
  SaveDialog.Title := strSaveFile;
  SaveDialog.DefaultExt := '.htm';
  SaveDialog.Filter := strHtmlFilter;
  SaveDialog.Filename := '';
  if SaveDialog.Execute then
  begin
    MainForm.MultiDoc.SetActiveChild(0);
    with MainForm.MultiDoc.ActiveObject as TMDIChild do
    begin
      if not Dataset1.Filtered then
      begin
        if MessageDlg(strConfirmation, strFilter, mtConfirmation, [mbYes, mbNo], 0) =
          mrYes then
        begin
          if FilterDlg.ShowModal = mrOk then
          begin
            Dataset1.Filtered := False;
            Dataset1.Filter := FilterDlg.FFilter;
            Dataset1.Filtered := True;
            Dataset1.First;
          end;
        end;
      end;
    end;
    Labels(SaveDialog.Filename, format);
    CreateMDIChild(ExtractFileName(SaveDialog.Filename));
    with MultiDoc.ActiveObject as TMDIChild do
    begin
      DataGrid.Visible := False;
      HtmlViewer.Visible := True;
      HtmlViewer.LoadFromFile(SaveDialog.Filename);
    end;
    UpdateTitleBar(ProjectFile);
  end;
end;

procedure TMainForm.ReportNamesItemClick(Sender: TObject);
begin
  if not IsOnline then
  begin
    MessageDlg(strError, strConnection, mtError, [mbOK], 0);
    Exit;
  end;
  SaveDialog.Title := strSaveFile;
  SaveDialog.DefaultExt := '.htm';
  SaveDialog.Filter := strHtmlFilter;
  SaveDialog.Filename := '';
  if SaveDialog.Execute then
  begin
    MainForm.MultiDoc.SetActiveChild(0);
    with MainForm.MultiDoc.ActiveObject as TMDIChild do
    begin
      if not Dataset1.Filtered then
      begin
        if MessageDlg(strConfirmation, strFilter, mtConfirmation, [mbYes, mbNo], 0) =
          mrYes then
        begin
          if FilterDlg.ShowModal = mrOk then
          begin
            Dataset1.Filtered := False;
            Dataset1.Filter := FilterDlg.FFilter;
            Dataset1.Filtered := True;
            Dataset1.First;
          end;
        end;
      end;
    end;
    Check(SaveDialog.Filename);
    CreateMDIChild(ExtractFileName(SaveDialog.Filename));
    with MultiDoc.ActiveObject as TMDIChild do
    begin
      DataGrid.Visible := False;
      HtmlViewer.Visible := True;
      HtmlViewer.LoadFromFile(SaveDialog.Filename);
    end;
    UpdateTitleBar(ProjectFile);
  end;
end;

procedure TMainForm.ReportStatisticsItemClick(Sender: TObject);
var
  options: TStringList;
  vResult: string;
  index, i: integer;
  found, graph_it: boolean;
begin
  options := TStringList.Create;
  try
    options.Add(strFamilies);
    options.Add(strGenera);
    options.Add(strSpecies);
    options.Add(strSites);
    options.Add(strDescriptors);
    options.Add(strSequences);
    options.Add(strVariables);
    vResult := InputCombo(strReport, strStatistics, options);
    index := options.IndexOf(vResult) + 1;
  finally
    options.Free;
  end;
  if Length(vResult) = 0 then
    Exit;
  case index of
    1, 2, 3, 4:
    begin
      if MessageDlg(strConfirmation, strGraph, mtConfirmation, [mbYes, mbNo], 0) =
        mrYes then
        graph_it := True
      else
        graph_it := False;
      if graph_it then
      begin
        if not LocateR then
          Exit;
      end;
    end;
    5:
    begin
      with MainForm.MultiDoc.ActiveObject as TMDIChild do
      begin
        if DataGrid.Columns.Count <= 12 then
        begin
          MessageDlg(strWarning, strNoDescriptors, mtWarning, [mbOK], 0);
          Exit;
        end;
      end;
    end;
    6:
    begin
      found := False;
      with MainForm.MultiDoc.ActiveObject as TMDIChild do
      begin
        for i := 0 to Dataset1.FieldCount - 1 do
        begin
          if Pos(Copy(UpperCase('SEQ'), 1, 3), Dataset1.Fields[i].FieldName) > 0 then
            found := True;
        end;
      end;
      if not found then
      begin
        MessageDlg(strWarning, strNoSequences, mtWarning, [mbOK], 0);
        Exit;
      end;
    end;
    7:
    begin
      if Length(Metadata.Find('dataset2').AsString) = 0 then
      begin
        MessageDlg(strWarning, strVariablesNotFound, mtWarning, [mbOK], 0);
        Exit;
      end;
      with MainForm.MultiDoc.ActiveObject as TMDIChild do
      begin
        if DataGrid.Columns.Count <= 4 then
        begin
          MessageDlg(strWarning, strNoVariables, mtWarning, [mbOK], 0);
          Exit;
        end;
      end;
    end;
  end;
  SaveDialog.Title := strSaveFile;
  SaveDialog.DefaultExt := '.htm';
  SaveDialog.Filter := strHtmlFilter;
  SaveDialog.Filename := '';
  if SaveDialog.Execute then
  begin
    if index < 7 then
      MainForm.MultiDoc.SetActiveChild(0)
    else
      MainForm.MultiDoc.SetActiveChild(1);
    Stats(SaveDialog.Filename, index, graph_it);
    CreateMDIChild(ExtractFileName(SaveDialog.Filename));
    with MultiDoc.ActiveObject as TMDIChild do
    begin
      DataGrid.Visible := False;
      HtmlViewer.Visible := True;
      HtmlViewer.LoadFromFile(SaveDialog.Filename);
    end;
    UpdateTitleBar(ProjectFile);
  end;
end;

procedure TMainForm.SearchFilterItemClick(Sender: TObject);
begin
  MainForm.MultiDoc.SetActiveChild(0);
  if FilterDlg.ShowModal = mrOk then
  begin
    with MainForm.MultiDoc.ActiveObject as TMDIChild do
    begin
      Dataset1.Filtered := False;
      Dataset1.Filter := FilterDlg.FFilter;
      Dataset1.Filtered := True;
      Dataset1.First;
    end;
    UpdateMenuItems(Self);
  end;
end;

procedure TMainForm.SearchFindItemClick(Sender: TObject);
var
  SearchStr{, KeyField}: string;
  FDataset: TsWorksheetDataset;
  Found: boolean;
  i: integer;
begin
  SearchStr := InputBox(strSearchCaption, strSearchPrompt, '');
  if not IsEmptyStr(SearchStr, [' ']) then
  begin
    with MainForm.MultiDoc.ActiveObject as TMDIChild do
    begin
      if CompareStr(MultiDoc.ActiveChild.Caption,
        Metadata.Find('dataset1').AsString) = 0 then
        FDataset := Dataset1
      else
      if CompareStr(MultiDoc.ActiveChild.Caption,
        Metadata.Find('dataset2').AsString) = 0 then
        FDataset := Dataset2;
      FDataset.First;
      Found := False;
      while (not FDataset.EOF) and (not Found) do
      begin
        for i := 0 to FDataset.FieldCount - 1 do
          if (FDataset.Fields[i].Value <> null) and
            (Pos(Uppercase(SearchStr), Uppercase(FDataset.Fields[i].Value)) > 0) then
          begin
            Found := True;
            DataGrid.SelectedField := FDataset.Fields[i];
            Break;
          end;
        if Found then
          break;
        FDataset.Next;
      end;
      //KeyField := FDataset.Fields[DataGrid.SelectedIndex].FieldName;
      //if not FDataset.Locate(KeyField, SearchStr, []) then
      if not Found then
        MessageDlg(strInformation, strTextNotFound, mtInformation, [mbOK], 0);
    end;
  end;
end;

procedure TMainForm.FileExportItemClick(Sender: TObject);
var
  options: TStringList;
  vResult: string;
  index, col, kind, levels, recs, nfam, nspp, nind, nsam, nrows, ncols: integer;
  found: boolean;
begin
  options := TStringList.Create;
  try
    options.Add(strEML);
    options.Add(strFITOPAC1);
    options.Add(strFITOPAC2);
    options.Add(strKML);
    options.Add(strSHP);
    options.Add(strBRAHMS);
    options.Add(strCEP);
    options.Add(strMVSP);
    options.Add(strFPM);
    options.Add(strCSV);
    vResult := InputCombo(strExport, strFormat, options);
    index := options.IndexOf(vResult);
  finally
    options.Free;
  end;
  SaveDialog.FileName := '';
  case index of
    0:
    begin
      SaveDialog.Title := strSaveFile;
      SaveDialog.DefaultExt := '.eml';
      SaveDialog.Filter := strEmlFilter;
      if SaveDialog.Execute then
      begin
        MainForm.MultiDoc.SetActiveChild(0);
        levels := Export.toEML(SaveDialog.FileName);
        MessageDlg(strInformation, Format(strExportMetadata,
          [ExtractFileName(SaveDialog.FileName)]), mtInformation, [mbOK], 0);
      end;
    end;
    1:
    begin
      if CompareText(Metadata.Find('method').AsString, LowerCase(strPlot)) = 0 then
        kind := 1
      else if CompareText(Metadata.Find('method').AsString,
        LowerCase(strQuadrat)) = 0 then
        kind := 2
      else
      begin
        MessageDlg(strWarning, strInvalidMethod, mtWarning, [mbOK], 0);
        Exit;
      end;
      SaveDialog.Title := strSaveFile;
      SaveDialog.DefaultExt := '.dad';
      SaveDialog.Filter := strFitopac1Filter;
      if SaveDialog.Execute then
      begin
        MainForm.MultiDoc.SetActiveChild(0);
        Export.toFitopac1(SaveDialog.FileName, kind, nfam, nspp, nind);
        MessageDlg(strInformation, Format(strExportFitopac1,
          [ExtractFileName(SaveDialog.FileName), nfam, nspp, nind]),
          mtInformation, [mbOK], 0);
      end;
    end;
    2:
    begin
      SaveDialog.Title := strSaveFile;
      SaveDialog.DefaultExt := '.fpd';
      SaveDialog.Filter := strFitopac2Filter;
      if SaveDialog.Execute then
      begin
        MainForm.MultiDoc.SetActiveChild(0);
        Export.toFitopac2(SaveDialog.FileName, nfam, nspp, nind, nsam);
        MessageDlg(strInformation, Format(strExportFitopac2,
          [ExtractFileName(SaveDialog.FileName), nfam, nspp, nind, nsam]),
          mtInformation, [mbOK], 0);
      end;
    end;
    3:
    begin
      found := False;
      with MultiDoc.ActiveObject as TMDIChild do
      begin
        for col := 0 to Dataset1.FieldCount - 1 do
        begin
          if (AnsiContainsText(Dataset1.Fields[col].FieldName, 'LAT')) or
            (AnsiContainsText(Dataset1.Fields[col].FieldName, 'LON')) then
            found := True;
        end;
      end;
      if not found then
      begin
        MessageDlg(strWarning, strNoCoordinates, mtWarning, [mbOK], 0);
        Exit;
      end;
      SaveDialog.Title := strSaveFile;
      SaveDialog.DefaultExt := '.kml';
      SaveDialog.Filter := strKmlFilter;
      if SaveDialog.Execute then
      begin
        MainForm.MultiDoc.SetActiveChild(0);
        recs := Export.toKML(SaveDialog.FileName);
        MessageDlg(strInformation, Format(strExportRecords,
          [recs, ExtractFileName(SaveDialog.FileName)]),
          mtInformation, [mbOK], 0);
      end;
    end;
    4:
    begin
      if Copy(OSVersion, 1, Pos(' ', OSVersion) - 1) <> 'Windows' then
      begin
        MessageDlg(strWindowsOnly, mtInformation, [mbOK], 0);
        Exit;
      end;
      found := False;
      with MultiDoc.ActiveObject as TMDIChild do
      begin
        for col := 0 to Dataset1.FieldCount - 1 do
        begin
          if (AnsiContainsText(Dataset1.Fields[col].FieldName, 'LAT')) or
            (AnsiContainsText(Dataset1.Fields[col].FieldName, 'LON')) then
            found := True;
        end;
      end;
      if not found then
      begin
        MessageDlg(strWarning, strNoCoordinates, mtWarning, [mbOK], 0);
        Exit;
      end;
      SaveDialog.Title := strSaveFile;
      SaveDialog.DefaultExt := '.shp';
      SaveDialog.Filter := strShpFilter;
      if SaveDialog.Execute then
      begin
        MainForm.MultiDoc.SetActiveChild(0);
        {$IFDEF WINDOWS}
        recs := Export.toSHP(SaveDialog.FileName);
        {$ENDIF}
        {$IFDEF LINUX}
        recs := 0;
        {$ENDIF}
        MessageDlg(strInformation, Format(strExportRecords,
          [recs, ExtractFileName(SaveDialog.FileName)]),
          mtInformation, [mbOK], 0);
      end;
    end;
    5:
    begin
      SaveDialog.Title := strSaveFile;
      SaveDialog.DefaultExt := '.dbf';
      SaveDialog.Filter := strDbfFilter;
      if SaveDialog.Execute then
      begin
        MainForm.MultiDoc.SetActiveChild(0);
        recs := Export.toRDE(SaveDialog.FileName);
        MessageDlg(strInformation, Format(strExportRecords, [recs]) +
          LineEnding + ExtractFileName(SaveDialog.FileName), mtInformation, [mbOK], 0);
      end;
    end;
    6:
    begin
      SaveDialog.Title := strSaveFile;
      SaveDialog.DefaultExt := '.dta';
      SaveDialog.Filter := strCEPFilter;
      if SaveDialog.Execute then
      begin
        MainForm.MultiDoc.SetActiveChild(0);
        Export.toCEP(SaveDialog.FileName, nrows, ncols);
        MessageDlg(strInformation,
          Format(strExportData, [ExtractFileName(SaveDialog.FileName), nrows, ncols]),
          mtInformation, [mbOK], 0);
      end;
    end;
    7:
    begin
      SaveDialog.Title := strSaveFile;
      SaveDialog.DefaultExt := '.mvs';
      SaveDialog.Filter := strMVSPFilter;
      if SaveDialog.Execute then
      begin
        MainForm.MultiDoc.SetActiveChild(0);
        Export.toMVSP(SaveDialog.FileName, nrows, ncols);
        MessageDlg(strInformation, Format(strExportData,
          [ExtractFileName(SaveDialog.FileName), nrows, ncols]),
          mtInformation, [mbOK], 0);
      end;
    end;
    8:
    begin
      SaveDialog.Title := strSaveFile;
      SaveDialog.DefaultExt := '.fpm';
      SaveDialog.Filter := strFPMFilter;
      if SaveDialog.Execute then
      begin
        MainForm.MultiDoc.SetActiveChild(0);
        Export.toFPM(SaveDialog.FileName, nrows, ncols);
        MessageDlg(strInformation, Format(strExportData,
          [ExtractFileName(SaveDialog.FileName), nrows, ncols]),
          mtInformation, [mbOK], 0);
      end;
    end;
    9:
    begin
      SaveDialog.Title := strSaveFile;
      SaveDialog.DefaultExt := '.csv';
      SaveDialog.Filter := strCSVFilter;
      if SaveDialog.Execute then
      begin
        MainForm.MultiDoc.SetActiveChild(0);
        Export.toCSV(SaveDialog.FileName, nrows, ncols);
        MessageDlg(strInformation, Format(strExportData,
          [ExtractFileName(SaveDialog.FileName), nrows, ncols]),
          mtInformation, [mbOK], 0);
      end;
    end;
  end;
end;

procedure TMainForm.FormDestroy(Sender: TObject);
begin
  SaveSettings;
  Metadata.Free;
end;

procedure TMainForm.FileOpenItemClick(Sender: TObject);
var
  DatasetName1, DatasetName2: string;
  f: TField;
begin
  if MultiDoc.ChildCount > 0 then
    FileCloseItemClick(Sender);
  OpenDialog.Title := strOpenFile;
  OpenDialog.DefaultExt := '.json';
  OpenDialog.Filter := strJsonFilter;
  OpenDialog.FileName := '';
  if OpenDialog.Execute then
  begin
    ProjectFile := ExtractFilename(Copy(OpenDialog.Filename, 1,
      RPos(ExtractFileExt(OpenDialog.Filename), OpenDialog.Filename) - 1)) + '.json';
    ChDir(ExtractFileDir(OpenDialog.Filename));
    Metadata.LoadFromFile(OpenDialog.Filename);
    DatasetName1 := Metadata.Force('dataset1').AsString;
    DatasetName2 := Metadata.Force('dataset2').AsString;
    if FileExists(DatasetName1) then
    begin
      if not CheckData(DatasetName1) then
        Exit;
      CreateMDIChild(DatasetName1);
      with MultiDoc.ActiveObject as TMDIChild do
      begin
        DataGrid.DataSource := DataSource1;
        Dataset1.FileName := DatasetName1;
        Dataset1.AutoFieldDefs := True;
        Dataset1.Open;
        for f in Dataset1.Fields do
          if f is TFloatField then
            TFloatField(f).DisplayFormat := '0.###';
        DataGrid.Visible := True;
        HtmlViewer.Visible := False;
      end;
    end
    else
    begin
      MessageDlg(strError, strDataNotFound, mtError, [mbOK], 0);
      Exit;
    end;
    if FileExists(DatasetName2) then
    begin
      CreateMDIChild(DatasetName2);
      with MultiDoc.ActiveObject as TMDIChild do
      begin
        DataGrid.DataSource := DataSource2;
        Dataset2.FileName := DatasetName2;
        Dataset2.AutoFieldDefs := True;
        Dataset2.Open;
        for f in Dataset2.Fields do
          if f is TFloatField then
            TFloatField(f).DisplayFormat := '0.###';
        DataGrid.Visible := True;
        HtmlViewer.Visible := False;
      end;
    end;
    UpdateTitleBar(ProjectFile);
    UpdateMenuItems(Self);
    MultiDoc.SetActiveChild(0);
  end;
end;

procedure TMainForm.FormCreate(Sender: TObject);
begin
  DefaultFormatSettings.DecimalSeparator := '.';
  Metadata := TJsonNode.Create;
  RestoreSettings;
  Screen.OnActiveFormChange := @UpdateMenuItems;
end;

procedure TMainForm.FileExitClick(Sender: TObject);
begin
  Close;
end;

procedure TMainForm.HelpAboutItemClick(Sender: TObject);
begin
  AboutBox.ShowModal;
end;

end.

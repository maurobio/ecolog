unit report;

{$mode objfpc}{$H+}

interface

uses
  SysUtils, LCLIntf, LCLType, Classes, Graphics, Controls, Forms, Dialogs,
  Menus, StdCtrls, ExtCtrls, Math, StrUtils, DB, fphttpclient, opensslsockets,
  FuzzyWuzzy, JsonTools, DOM, XMLRead, XPath, Process;

{const
  SAMPLE = 0;
  SPECIMEN = 1;
  FAMILY = 2;
  SPECIES = 3;
  COLLECTOR = 4;
  NUMBER = 5;
  DATE = 6;
  OBS = 7;
  LOCALITY = 8;
  LATITUDE = 9;
  LONGITUDE = 10;
  ELEVATION = 11;}

procedure Header(var outf: TextFile; const title: string);
procedure Catalog(fname: string);
procedure Geral(fname: string);
procedure Labels(fname: string; formato: integer);
procedure Stats(fname: string; option: integer; graph_it: boolean);
procedure Check(fname: string);
procedure Georef(fname: string);

implementation

uses main, child, useful, eval, progress, diversity, DNA;

resourcestring
  strReportGeneral = 'RELATÓRIO GERAL DE ESPÉCIES';
  strReportCatalog = 'CATÁLOGO DE COLETAS';
  strReportStatistics = 'RELATÓRIO ESTATÍSTICO';
  strReportCheck = 'RELATÓRIO DE VERIFICAÇÃO NOMENCLATURAL';
  strReportGeoref = 'RELATÓRIO DE GEOCODIFICAÇÃO';
  strFamilyStatistics = 'Estatística de FAMÍLIAS';
  strGenusStatistics = 'Estatística de GÊNEROS';
  strSpeciesStatistics = 'Estatística de ESPÉCIES';
  strSampleStatistics = 'Estatística de LOCAIS';
  strDescriptorStatistics = 'Estatística de DESCRITORES';
  strVariableStatistics = 'Estatística de VARIÁVEIS AMBIENTAIS';
  strSequenceStatistics = 'Estatística de SEQUÊNCIAS';
  strProjectName = 'Projeto: ';
  strSurveyType = 'Tipo de Levantamento: ';
  strReportDate = 'Data - ';
  strReportTime = 'Hora - ';
  strFilter = 'FILTRO: ';
  strElevation = 'Alt. ';
  strDepth = 'Prof. ';
  strDate = 'Data: ';
  strProvince = '  Mun. ';
  strSpecimen = 'INDIVÍDUO';
  strSpecies = 'ESPÉCIE';
  strLocality = 'LOCALIDADE';
  strLatitude = 'LATITUDE';
  strLongitude = 'LONGITUDE';
  strTotalRecords = 'TOTAL DE REGISTROS ARMAZENADOS: ';
  strRetrievedRecords = 'NÚMERO DE REGISTROS RECUPERADOS: ';
  strPercent = '% DE RECUPERAÇÃO/TOTAL: ';
  strSuspectedNames = 'NOMES SUSPEITOS';
  strCheckedNames = 'NOMES VERIFICADOS';
  strTotalNames = 'TOTAL DE NOMES = ';
  strValidNames = 'TOTAL DE NOMES VÁLIDOS = ';
  strSynonyms = 'TOTAL DE SINÔNIMOS = ';
  strTotalSuspectedNames = 'TOTAL DE NOMES SUSPEITOS = ';
  strTaxonomicQualityIndex = 'ÍNDICE DE QUALIDADE TAXONÔMICA = ';
  strCoded = 'TOTAL DE REGISTROS GEOCODIFICADOS = ';
  strUncoded = 'TOTAL DE REGISTROS NÃO-GEOCODIFICADOS = ';
  strFamily = 'FAMÍLIA';
  strGenus = 'GÊNERO';
  strSample = 'LOCAL';
  strNoSpp = 'NO.SPP.';
  strPercentSpp = '%SPP';
  strNoInd = 'NO.IND.';
  strDom = 'DOM.%';
  strFreq = 'FREQ';
  strFR = 'F.R.%';
  strCodon = 'CÓDON';
  strCategory = 'CATEGORIA';
  strTotalFamilies = 'TOTAL DE FAMÍLIAS = ';
  strTotalGenera = 'TOTAL DE GÊNEROS = ';
  strTotalSpecies = 'TOTAL DE ESPÉCIES = ';
  strTotalSpecimens = 'TOTAL DE INDIVÍDUOS = ';
  strTotalSamples = 'TOTAL DE LOCAIS = ';
  strTotalDescriptors = 'TOTAL DE DESCRITORES = ';
  strTotalVariables = 'TOTAL DE VARIÁVEIS = ';
  strTotalSequences = 'TOTAL DE SEQUÊNCIAS = ';
  strTotalCodons = 'TOTAL DE CÓDONS = ';
  strFamiliesCount = ' família(s)';
  strGeneraCount = ' gênero(s)';
  strSpeciesCount = ' espécie(s)';
  strSpecimensCount = ' indivíduo(s)';
  strSamplesCount = ' locais';
  strDescriptorsCount = ' descritor(es)';
  strVariablesCount = ' variáveis';
  strSequencesCount = ' sequência(s)';
  strCodonsCount = ' códon(s)';
  strConstant = 'Constante';
  strAccessory = 'Acessória';
  strAccidental = 'Acidental';
  strDiversityAnalysis = 'Análise de DIVERSIDADE';
  strSpeciesRichness = 'RIQUEZA DE ESPÉCIES';
  strMargalefIndex = ' Índice de Margalef  (D1)      = ';
  strMenhinickIndex = ' Índice de Menhinick (D2)      = ';
  strDiversity = 'DIVERSIDADE';
  strSimpsonIndex = ' Índice de Simpson (C)         = ';
  strShannonIndex = ' Índice de Shannon-Weaver (H'') = ';
  strHurlbertIndex = ' Índice de Hurlbert (PIE)      = ';
  strBergerIndex = ' Índice de Berger-Parker (d)   = ';
  strMcIntoshIndex = ' Índice de McIntosh (M)        = ';
  strEvenness = 'EQUITABILIDADE';
  strEvennessIndex = ' Equitabilidade (J)            = ';
  strNumber = 'NÚMERO';
  strDescriptor = 'DESCRITOR';
  strVariable = 'VARIÁVEL';
  strType = 'TIPO';
  strUnit = 'UNIDADE';
  strMean = 'MÉDIA';
  strStdDev = 'DESVIO-PADRÃO';
  strSD = 'DESV.PAD.';
  strCV = 'COEF.VAR. (.%)';
  strMinimum = 'MÍNIMO';
  strMin = 'MÍN.';
  strMaximum = 'MÁXIMO';
  strMax = 'MÁX.';
  strCases = 'CASOS';
  strObs = 'N';
  strNumeric = 'Numérico';
  strText = 'Texto';
  strUnknown = 'Desconhecido';
  strNucleotideComposition = 'Composição de NUCLEOTÍDEOS';
  strCodonFrequency = 'Frequência de CÓDONS';
  strFamilyTitle = 'FREQUÊNCIA DE ESPÉCIES POR FAMÍLIA';
  strGenusTitle = 'FREQUÊNCIA DE ESPÉCIES POR GÊNERO';
  strSpeciesTitle = 'FREQUÊNCIA DE INDIVÍDUOS POR ESPÉCIE';
  strXLabel = 'FREQUÊNCIA RELATIVA (%)';
  strPlotTitle = 'CURVA DO COLETOR';
  strPlotXLabel = 'AMOSTRAS';
  strPlotYLabel = 'NÚMERO CUMULATIVO DE ESPÉCIES';
  strNotExecute = 'Erro na execução do programa %s.';
  strError = 'Erro';
  strCheck = 'Verificar Nomes';
  strCheckNames = 'Verificando nomes...';
  strSynonym = ' -- sinônimo de ';
  strValidName = ' -- nome válido';
  strNoRecord = ' -- sem registro';

procedure ProcessArray(ADataSet: TDataSet; FieldNumber: integer;
  var a: array of double; var len: integer);
var
  AField: TField;
  AValue: string;
  res: variant;
  cnt: integer;
begin
  len := 0;
  AField := ADataSet.Fields[FieldNumber];
  ADataSet.First;
  while not ADataSet.EOF do
  begin
    AValue := AField.AsString;
    if Pos('+', AValue) > 0 then
    begin
      res := eval.eval(AValue);
      cnt := AValue.CountChar('+') + 1;
      a[len] := (res / cnt);
    end
    else
      a[len] := StrToFloatDef(AValue, 0.0);
    Inc(len);
    ADataSet.Next;
  end;
end;

function Percent(Val1, Val2: integer): double;
begin
  if (Val2 = 0) then
    Result := 0.0
  else
    Result := (double(Val1) / double(Val2)) * 100.00;
end;

procedure Header(var outf: TextFile; const title: string);
var
  project, method: string;
begin
  project := Metadata.Find('title').AsString;
  method := Metadata.Find('method').AsString;
  WriteLn(outf, '<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 3.2//EN"">');
  WriteLn(outf, '<html>');
  WriteLn(outf, '<head>');
  WriteLn(outf, '<title>', project + ' - ' + title, '</title>');
  WriteLn(outf, '</head>');
  WriteLn(outf, '<body>');
  WriteLn(outf, '<pre>');
  WriteLn(outf, '* * * * * * * * * * * * * * * * * * * *');
  WriteLn(outf, '* * * * *     E C O L O G     * * * * *');
  WriteLn(outf, '* * * * * * * * * * * * * * * * * * * *');
  WriteLn(outf, '</pre>');
  WriteLn(outf, strReportDate, DateToStr(Now), '<br>');
  WriteLn(outf, strReportTime, TimeToStr(Now), '<br><br>');
  WriteLn(outf);
  WriteLn(outf, strProjectName, project, '<br><br>');
  WriteLn(outf, strSurveyType, method, '<br><br>');
  WriteLn(outf, title, '<br>');
  with MainForm.MultiDoc.ActiveObject as TMDIChild do
  begin
    if Length(Dataset1.Filter) > 0 then
      if Dataset1.Filtered then
        WriteLn(outf, '<br>', strFilter, Dataset1.Filter, '<br><br>');
  end;
end;

procedure Footer(var outf: TextFile; conta, total: integer);
begin
  if (conta > 0) or (total > 0) then
  begin
    WriteLn(outf, '<br>', '--', '<br>');
    WriteLn(outf, strTotalRecords, IntToStr(total), '<br>');
    WriteLn(outf, strRetrievedRecords, IntToStr(conta), '<br>');
    WriteLn(outf, strPercent, FloatToStrF(Percent(conta, total), ffFixed, 5, 2) + ' %');
  end;
  WriteLn(outf, '</body>');
  WriteLn(outf, '</html>');
end;

procedure Scissor(var outf: TextFile);
begin
  WriteLn(outf, '<p><img src="scissor.png" align="Bottom" alt=" ">',
    StringOfChar('-', 160), '</p>');
end;

procedure Catalog(fname: string);
var
  total, counter: integer;
  grpCheck, grpEval, subCheck, subEval: string;
  genus, nomen, spp, author1, ssp, infraname, author2: string;
  mplace, mloc, line: string;
  alt: integer;
  SPECIMEN, FAMILY, SPECIES, COLLECTOR, NUMBER, DATE, LOCALITY,
  LATITUDE, LONGITUDE, ELEVATION: string;
  outfile: TextFile;
begin
  with MainForm.MultiDoc.ActiveObject as TMDIChild do
  begin
    CSVExporter1.FileName := 'temp.csv';
    CSVExporter1.Execute;
    QueryConnection1.DatabasePath := GetCurrentDir;
    QueryConnection1.Connect;
    QueryDataSet1.TableName := 'temp';
    QueryDataset1.LoadFromTable;
    SPECIMEN := QueryDataset1.Fields[1].FieldName;
    FAMILY := QueryDataset1.Fields[2].FieldName;
    SPECIES := QueryDataset1.Fields[3].FieldName;
    COLLECTOR := QueryDataset1.Fields[4].FieldName;
    NUMBER := QueryDataset1.Fields[5].FieldName;
    DATE := QueryDataset1.Fields[6].FieldName;
    LOCALITY := QueryDataset1.Fields[8].FieldName;
    LATITUDE := QueryDataset1.Fields[9].FieldName;
    LONGITUDE := QueryDataset1.Fields[10].FieldName;
    ELEVATION := QueryDataset1.Fields[11].FieldName;
    QueryDataset1.SQL.Text := 'SELECT ' + SPECIMEN + ', ' + FAMILY +
      ', ' + SPECIES + ', ' + COLLECTOR + ', ' + NUMBER + ', ' +
      DATE + ', ' + LOCALITY + ', ' + ELEVATION + ' FROM temp ORDER BY ' +
      FAMILY + ', ' + SPECIES + ', ' + LOCALITY + ', ' + COLLECTOR + ', ' + DATE;
    QueryDataset1.QueryExecute;
    QueryDataset1.SaveToTable;
    total := Dataset1.RecordCount;
  end;

  mplace := Metadata.Find('province').AsString;
  mloc := Metadata.Find('locality').AsString;

  counter := 0;
  grpCheck := '';
  grpEval := '';
  subCheck := '';
  subEval := '';

  AssignFile(outfile, fname);
  Rewrite(outfile);
  Header(outfile, strReportCatalog);

  WriteLn(outfile, '<dl>');
  with MainForm.MultiDoc.ActiveObject as TMDIChild do
  begin
    QueryDataset1.First;
    while not QueryDataset1.EOF do
    begin
      grpCheck := QueryDataset1.Fields[1].AsString;
      if grpEval <> grpCheck then
      begin
        grpEval := grpCheck;
        WriteLn(outfile, '<br>', UpperCase(grpCheck), '<br>');
      end;

      subCheck := QueryDataset1.Fields[2].AsString;
      if subEval <> subCheck then
      begin
        subEval := subCheck;
        ParseName(QueryDataset1.Fields[2].AsString, genus, nomen,
          spp, author1, ssp, infraname, author2);
        WriteLn(outfile, '<br><dt><i>' + genus + ' ' + nomen +
          ' ' + spp + '</i> ' + author1 + ' ' + infraname + ' <i>' +
          infraname + '</i> ' + author2 + '</dt>');
      end;

      line := '<dd>';
      line += strProvince + Trim(mplace) + ': ' + Trim(mloc) + ', ' +
        QueryDataset1.Fields[6].AsString + '. ';
      alt := StrToIntDef(QueryDataset1.Fields[7].AsString, 0);
      if alt >= 0 then
        line += strElevation + IntToStr(alt)
      else if alt < 0 then
        line += strDepth + IntToStr(alt * -1);
      if alt <> 0 then
        line += ' ' + GetUnit(QueryDataset1.Fields[7].AsString) + '. ';
      line += '<u>' + QueryDataset1.Fields[3].AsString + ' ' +
        QueryDataset1.Fields[4].AsString + '</u>. ' +
        QueryDataset1.Fields[5].AsString;
      WriteLn(outfile, line, '</dd>');
      QueryDataset1.Next;
      Inc(counter);
    end;
    QueryConnection1.Disconnect;
  end;
  WriteLn(outfile, '</dl>');

  Footer(outfile, counter, total);
  CloseFile(outfile);
  if FileExists('temp.csv') then
    DeleteFile('temp.csv');
end;

procedure Geral(fname: string);
var
  total, counter: integer;
  grpCheck, grpEval: string;
  genus, nomen, spp, author1, ssp, infraname, author2: string;
  FAMILY, SPECIES: string;
  outfile: TextFile;
begin
  with MainForm.MultiDoc.ActiveObject as TMDIChild do
  begin
    CSVExporter1.FileName := 'temp.csv';
    CSVExporter1.Execute;
    QueryConnection1.DatabasePath := GetCurrentDir;
    QueryConnection1.Connect;
    QueryDataSet1.TableName := 'temp';
    QueryDataset1.LoadFromTable;
    FAMILY := QueryDataset1.Fields[2].FieldName;
    SPECIES := QueryDataset1.Fields[3].FieldName;
    QueryDataset1.SQL.Text := 'SELECT ' + FAMILY + ', ' + SPECIES +
      ' FROM temp GROUP BY ' + FAMILY + ', ' + SPECIES + ' ORDER BY ' +
      FAMILY + ', ' + SPECIES;
    QueryDataset1.QueryExecute;
    QueryDataset1.SaveToTable;
    total := Dataset1.RecordCount;
  end;

  counter := 0;
  grpCheck := '';
  grpEval := '';

  AssignFile(outfile, fname);
  Rewrite(outfile);
  Header(outfile, strReportGeneral);

  with MainForm.MultiDoc.ActiveObject as TMDIChild do
  begin
    QueryDataset1.First;
    while not QueryDataset1.EOF do
    begin
      grpCheck := Trim(QueryDataset1.Fields[0].AsString);
      if grpEval <> grpCheck then
      begin
        grpEval := grpCheck;
        WriteLn(outfile, '<p>', UpperCase(grpCheck), '</p>');
      end;
      ParseName(QueryDataset1.Fields[1].AsString, genus, nomen,
        spp, author1, ssp, infraname, author2);
      WriteLn(outfile, '&nbsp;&nbsp;&nbsp;&nbsp;<i>' + genus + ' ' +
        nomen + ' ' + spp + '</i> ' + author1 + ' ' + infraname +
        ' <i>' + infraname + '</i> ' + author2 + '<br>');
      QueryDataset1.Next;
      Inc(counter);
    end;
    QueryConnection1.Disconnect;
  end;

  Footer(outfile, counter, total);
  CloseFile(outfile);
  if FileExists('temp.csv') then
    DeleteFile('temp.csv');
end;

procedure Labels(fname: string; formato: integer);
var
  mfamily, mgenus, mcf, mspecies, mauthor1, msubsp, minfra, mauthor2: string;
  mcoletor, mdatacol, mobs: string;
  l_institute, l_header, l_footer, mcountry, mstate, mprov, mloc, mproc: string;
  vlatdeg, vlondeg, vlatmin, vlonmin, vlatsec, vlonsec: integer;
  vlath, vlonh: char;
  alt: integer;
  SPECIMEN, FAMILY, SPECIES, COLLECTOR, NUMBER, DATE, OBS, LOCALITY,
  LATITUDE, LONGITUDE, ELEVATION: string;
  outfile: TextFile;
begin
  with MainForm.MultiDoc.ActiveObject as TMDIChild do
  begin
    CSVExporter1.FileName := 'temp.csv';
    CSVExporter1.Execute;
    QueryConnection1.DatabasePath := GetCurrentDir;
    QueryConnection1.Connect;
    QueryDataSet1.TableName := 'temp';
    QueryDataset1.LoadFromTable;
    SPECIMEN := QueryDataset1.Fields[1].FieldName;
    FAMILY := QueryDataset1.Fields[2].FieldName;
    SPECIES := QueryDataset1.Fields[3].FieldName;
    COLLECTOR := QueryDataset1.Fields[4].FieldName;
    NUMBER := QueryDataset1.Fields[5].FieldName;
    DATE := QueryDataset1.Fields[6].FieldName;
    OBS := QueryDataset1.Fields[7].FieldName;
    LOCALITY := QueryDataset1.Fields[8].FieldName;
    LATITUDE := QueryDataset1.Fields[9].FieldName;
    LONGITUDE := QueryDataset1.Fields[10].FieldName;
    ELEVATION := QueryDataset1.Fields[11].FieldName;
    QueryDataset1.SQL.Text := 'SELECT ' + SPECIMEN + ', ' + FAMILY +
      ', ' + SPECIES + ', ' + COLLECTOR + ', ' + NUMBER + ', ' +
      DATE + ', ' + OBS + ', ' + LOCALITY + ', ' + ELEVATION + ', ' +
      LATITUDE + ', ' + LONGITUDE + ', ' + ' FROM temp ORDER BY ' +
      FAMILY + ', ' + SPECIES + ', ' + LOCALITY + ', ' + COLLECTOR + ', ' + DATE;
    QueryDataset1.QueryExecute;
    QueryDataset1.SaveToTable;
  end;

  l_institute := Metadata.Find('institution').AsString;
  l_header := Metadata.Find('title').AsString;
  l_footer := Metadata.Find('funding').AsString;
  mcountry := Metadata.Find('country').AsString;
  mstate := Metadata.Find('state').AsString;
  mprov := Metadata.Find('province').AsString;
  mloc := Metadata.Find('locality').AsString;

  AssignFile(outfile, fname);
  Rewrite(outfile);
  WriteLn(outfile, '<html>');
  WriteLn(outfile, '<head>');
  WriteLn(outfile, '<title>' + l_header + '</title>');
  WriteLn(outfile, '</head>');
  WriteLn(outfile, '<body>');

  with MainForm.MultiDoc.ActiveObject as TMDIChild do
  begin
    QueryDataset1.First;
    while not QueryDataset1.EOF do
    begin
      mproc := UpperCase(mcountry) + ', ' + Trim(mstate) + ', Mun. ' +
        Trim(mprov) + ', ' + Trim(mloc);
      mproc += ', ' + QueryDataset1.Fields[7].AsString + '.';

      DegToDMS(StrToFloatDef(QueryDataset1.Fields[9].AsString, 0.0),
        vlatdeg, vlatmin, vlatsec);
      DegToDMS(StrToFloatDef(QueryDataset1.Fields[10].AsString, 0.0),
        vlondeg, vlonmin, vlonsec);
      if vlatdeg < 0.0 then
        vlath := 'S'
      else
        vlath := 'N';
      if vlondeg < 0.0 then
        vlonh := 'W'
      else
        vlonh := 'E';
      mproc += ' (Coord. ' + FloatToStr(vlatdeg) + '<sup>o</sup>' +
        FloatToStr(vlatmin) + '''' + FloatToStr(vlatsec) + ''' ' +
        vlath + ', ' + FloatToStr(vlondeg) + '<sup>o</sup>' +
        FloatToStr(vlonmin) + '''' + FloatToStr(vlonsec) + ''' ' + vlonh + '). ';

      alt := StrToIntDef(QueryDataset1.Fields[8].AsString, 0);
      if alt > 0 then
        mproc += strElevation + IntToStr(alt)
      else if alt < 0 then
        mproc += strDepth + IntToStr(alt * -1);
      if alt <> 0 then
        mproc += ' ' + GetUnit(QueryDataset1.Fields[8].FieldName) + '. ';

      mfamily := QueryDataset1.Fields[1].AsString;
      ParseName(QueryDataset1.Fields[2].AsString, mgenus, mcf,
        mspecies, mauthor1, msubsp, minfra, mauthor2);
      mgenus := '<i>' + mgenus + '</i>';
      mspecies := mgenus + ' ' + mcf + ' <i>' + mspecies + '</i> ' + mauthor1;
      if Length(msubsp) > 0 then
        mspecies += ' ' + msubsp + ' <i>' + minfra + '</i> ' + mauthor2;
      mcoletor := 'Col. ' + QueryDataset1.Fields[3].AsString + ' ' +
        QueryDataset1.Fields[4].AsString;

      mdatacol := strDate;
      if formato = 1 then
        mdatacol += QueryDataset1.Fields[5].AsString
      else if formato = 2 then
        mdatacol += Roman(QueryDataset1.Fields[5].AsString)
      else if formato = 3 then
        mdatacol += Extenso(QueryDataset1.Fields[5].AsString);

      mobs := QueryDataset1.Fields[6].AsString;
      if Length(mobs) > 0 then
        if not mobs.EndsWith('.') then
          mobs += '.';

      WriteLn(outfile);
      WriteLn(outfile, '<center>');
      WriteLn(outfile, '<table border=0 cellspacing=2 cellpadding=2 width="80%">');
      WriteLn(outfile, '<tr align="Center"><td align="Center" colspan=2>' +
        l_institute + '</td></tr>');
      WriteLn(outfile, '<tr align="Center"><td align="Center" colspan=2>' +
        l_header + '</td></tr>');
      WriteLn(outfile, '<tr align="Center"><td align="Left" colspan=2>' +
        mfamily + '</td></tr>');
      WriteLn(outfile, '<tr align="Center"><td align="Left" colspan=2>' +
        mspecies + '</td></tr>');
      WriteLn(outfile, '<tr align="Center"><td align="Left" colspan=2>' +
        mproc + '</td></tr>');
      WriteLn(outfile, '<tr align="Center"><td align="Left" colspan=2>' +
        mobs + '</td></tr>');
      WriteLn(outfile, '<tr align="Center"><td align="Left" colspan=2>' +
        mcoletor + '</td></tr>');
      WriteLn(outfile, '<tr align="Center"><td align="Left" colspan=2>' +
        mdatacol + '</td></tr>');
      WriteLn(outfile, '<tr align="Center"><td align="Left" colspan=2>' +
        'Det.' + '</td></tr>');
      WriteLn(outfile, '<tr align="Center"><td align="Left" colspan=2>' +
        strDate + '</td></tr>');
      WriteLn(outfile, '<tr align="Center"><td align="Center" colspan=2>' +
        l_footer + '</td></tr>');
      WriteLn(outfile, '</table>');
      WriteLn(outfile, '</center>');
      Scissor(outfile);
      QueryDataset1.Next;
    end;
    QueryConnection1.Disconnect;
  end;

  WriteLn(outfile, '</body>');
  WriteLn(outfile, '</html>');
  CloseFile(outfile);
  if FileExists('temp.csv') then
    DeleteFile('temp.csv');
end;

procedure Check(fname: string);
const
  url = 'https://api.gbif.org/v1/species?name=';
var
  v, s, x, c, i, t: integer;
  q: double;
  p: single;
  names: TStringList;
  genus, nomen, spp, author1, ssp, infraname, author2: string;
  SPECIES, queryStr, canonicalname, scientificname, author, status,
  valid_name, valid_author, taxon_list: string;
  outfile: TextFile;
  rawJson: ansistring;
  JsonData: TJsonNode;
begin
  with MainForm.MultiDoc.ActiveObject as TMDIChild do
  begin
    CSVExporter1.FileName := 'temp.csv';
    CSVExporter1.Execute;
    QueryConnection1.DatabasePath := GetCurrentDir;
    QueryConnection1.Connect;
    QueryDataSet1.TableName := 'temp';
    QueryDataset1.LoadFromTable;
    SPECIES := QueryDataset1.Fields[3].FieldName;
    QueryDataset1.SQL.Text := 'SELECT ' + SPECIES + ' FROM temp GROUP BY ' +
      SPECIES + ' ORDER BY ' + SPECIES;
    QueryDataset1.QueryExecute;
    QueryDataset1.SaveToTable;
  end;

  v := 0;
  s := 0;
  x := 0;
  c := 0;
  q := 0.0;

  AssignFile(outfile, fname);
  Rewrite(outfile);
  Header(outfile, strReportCheck);

  names := TStringList.Create;
  with MainForm.MultiDoc.ActiveObject as TMDIChild do
  begin
    QueryDataset1.First;
    while not QueryDataset1.EOF do
    begin
      names.Add(Trim(QueryDataset1.Fields[0].AsString));
      QueryDataset1.Next;
    end;
    QueryConnection1.Disconnect;
  end;
  t := names.Count;

  WriteLn(outfile, '<br>' + strSuspectedNames + '<br><br>');
  for i := 0 to names.Count - 2 do
  begin
    if PartialRatio(names[i], names[i + 1]) > 90 then
    begin
      WriteLn(outfile, '<i>' + names[i] + '</i><br>');
      Inc(x);
    end;
  end;
  WriteLn(outfile);
  WriteLn(outfile, '<br>' + strCheckedNames + '<br>');

  with ProgressDlg do
  begin
    Caption := strCheck;
    LabelText.Caption := strCheckNames;
    ProgressBar.Min := 0;
    ProgressBar.Max := 100;
    ProgressBar.Position := 0;
    ProgressBar.Step := 1;
    Show;
  end;
  Application.ProcessMessages;
  Screen.Cursor := crHourGlass;

  for i := 0 to names.Count - 1 do
  begin
    ParseName(names[i], genus, nomen,
      spp, author1, ssp, infraname, author2);
    queryStr := genus + ' ' + spp + ifThen(Length(ssp) > 0, ' ' + infraname, '');
    WriteLn(outfile, '<br><i>' + queryStr + '</i>');
    try
      JsonData := TJsonNode.Create;
      rawJson := TFPHTTPClient.SimpleGet(url +
        StringReplace(queryStr, ' ', '%20', [rfReplaceAll]));
      JsonData.Parse(rawJson);
      scientificname := JsonData.Find('results/0/scientificName').AsString;
      canonicalname := JsonData.Find('results/0/canonicalName').AsString;
      status := JsonData.Find('results/0/taxonomicStatus').AsString;
      if LowerCase(status) <> 'accepted' then
        valid_name := JsonData.Find('results/0/accepted').AsString;
      author := JsonData.Find('results/0/authorship').AsString;
      taxon_list := JsonData.Find('results/0/kingdom').AsString +
        '; ' + JsonData.Find('results/0/phylum').AsString + '; ' +
        JsonData.Find('results/0/class').AsString + '; ' +
        JsonData.Find('results/0/order').AsString + '; ' +
        JsonData.Find('results/0/family').AsString;
      JsonData.Free;
      WriteLn(outfile, ' ' + author + ' ');
      if LowerCase(status) = 'synonym' then
      begin
        Inc(s);
        WriteLn(outfile, strSynonym + '<i>' + valid_name + '</i> ');
      end
      else
      begin
        Inc(v);
        WriteLn(outfile, strValidName);
      end;
      WriteLn(outfile, '<br>' + LineEnding + '&nbsp;&nbsp;&nbsp;');
      WriteLn(outfile, taxon_list);
      WriteLn(outfile, '<br>');
    except
      WriteLn(outfile, strNoRecord + '<br>');
    end;
    Inc(c);
    p := Succ(c) / names.Count;
    with ProgressDlg do
      ProgressBar.Position := Round(100 * p) + 1;
    Application.ProcessMessages;
    Sleep(1000);
  end;
  ProgressDlg.Close;
  Screen.Cursor := crDefault;
  q := v / t;

  names.Free;
  WriteLn(outfile, '<br>--<br>');
  WriteLn(outfile, strTotalNames + IntToStr(t) + '<br>');
  WriteLn(outfile, strValidNames + IntToStr(v) + '<br>');
  WriteLn(outfile, strSynonyms + IntToStr(s) + '<br>');
  WriteLn(outfile, strTotalSuspectedNames + IntToStr(x) + '<br>');
  WriteLn(outfile, strTaxonomicQualityIndex + FloatToStrF(q, ffFixed, 3, 3) + '<br>');
  WriteLn(outfile, '</body>');
  WriteLn(outfile, '</html>');
  CloseFile(outfile);
  if FileExists('temp.csv') then
    DeleteFile('temp.csv');
end;

procedure Georef(fname: string);
const
  url = 'https://en.wikipedia.org/w/api.php?action=query&prop=coordinates&titles=';
var
  coded, uncoded, nrows: integer;
  SPECIMEN, SPECIES, LOCALITY, LATITUDE, LONGITUDE: string;
  genus, cf, spp, author1, subsp, infraname, author2: string;
  mcountry, mstate, mlocality, mlatitude, mlongitude, mspecimen, mspecies: string;
  queryStr, location_name, location_latitude, location_longitude: string;
  georef: boolean;
  outfile, tmpfile: TextFile;
  Doc: TXMLDocument;
  rawXML: ansistring;
  Result: TXPathVariable;
begin
  with MainForm.MultiDoc.ActiveObject as TMDIChild do
  begin
    CSVExporter1.FileName := 'temp.csv';
    CSVExporter1.Execute;
    QueryConnection1.DatabasePath := GetCurrentDir;
    QueryConnection1.Connect;
    QueryDataSet1.TableName := 'temp';
    QueryDataset1.LoadFromTable;
    SPECIMEN := QueryDataset1.Fields[1].FieldName;
    SPECIES := QueryDataset1.Fields[3].FieldName;
    LOCALITY := QueryDataset1.Fields[8].FieldName;
    LATITUDE := QueryDataset1.Fields[9].FieldName;
    LONGITUDE := QueryDataset1.Fields[10].FieldName;
    QueryDataset1.SQL.Text := 'SELECT ' + SPECIMEN + ', ' + SPECIES +
      ', ' + LOCALITY + ', ' + LATITUDE + ', ' + LONGITUDE +
      ' FROM temp ORDER BY ' + SPECIMEN;
    QueryDataset1.QueryExecute;
    QueryDataset1.SaveToTable;
  end;

  coded := 0;
  uncoded := 0;
  nrows := 0;

  AssignFile(outfile, fname);
  Rewrite(outfile);
  Header(outfile, strReportGeoref);
  WriteLn(outfile, '<br><br>');
  WriteLn(outfile, '<table border=0 cellspacing=1 cellpadding=1 width="100%">');
  WriteLn(outfile, '<tr>');
  WriteLn(outfile, '<th>' + strSpecimen + '</th>');
  WriteLn(outfile, '<th>' + strSpecies + '</th>');
  WriteLn(outfile, '<th>' + strLocality + '</th>');
  WriteLn(outfile, '<th>' + strLatitude + '</th>');
  WriteLn(outfile, '<th>' + strLongitude + '</th>');
  WriteLn(outfile, '</tr>');

  with MainForm.MultiDoc.ActiveObject as TMDIChild do
  begin
    Screen.Cursor := crHourGlass;
    QueryDataset1.First;
    while not QueryDataset1.EOF do
    begin
      mcountry := Metadata.Find('country').AsString;
      mstate := Metadata.Find('state').AsString;
      mspecimen := QueryDataset1.Fields[0].AsString;
      mspecies := QueryDataset1.Fields[1].AsString;
      mlocality := QueryDataset1.Fields[2].AsString;
      mlatitude := QueryDataset1.Fields[3].AsString;
      mlongitude := QueryDataset1.Fields[4].AsString;
      if (Abs(StrToFloatDef(mlatitude, 0.0)) = 0.0) and
        (Abs(StrToFloatDef(mlongitude, 0.0)) = 0.0) then
      begin
        georef := False;
        Inc(uncoded);
      end
      else
      begin
        georef := True;
        Inc(coded);
      end;
      if not georef and (Length(mlocality) > 0) then
      begin
        ParseName(mspecies, genus, cf, spp, author1, subsp, infraname, author2);
        {if Length(mstate) > 0 then
          queryStr := mlocality + ',' + mstate + ',' + mcountry
        else
          queryStr := mlocality + ',' + mcountry;}
        queryStr := mlocality;
        rawXML := TFPHTTPClient.SimpleGet(url +
          StringReplace(queryStr, ' ', '%20', [rfReplaceAll]) + '&format=xml');
        AssignFile(tmpfile, 'temp.xml');
        Rewrite(tmpfile);
        WriteLn(tmpfile, rawXML);
        CloseFile(tmpfile);
        ReadXMLFile(Doc, 'temp.xml');
        Result := EvaluateXPathExpression('/api/query/pages/page/@title',
          Doc.DocumentElement);
        location_name := string(Result.AsText);
        Result := EvaluateXPathExpression('/api/query/pages/page/coordinates/co/@lat',
          Doc.DocumentElement);
        location_latitude := string(Result.AsText);
        Result := EvaluateXPathExpression('/api/query/pages/page/coordinates/co/@lon',
          Doc.DocumentElement);
        location_longitude := string(Result.AsText);

        WriteLn(outfile, '<tr>');
        WriteLn(outfile, '<td align="Left">' + mspecimen + '</td>');
        WriteLn(outfile, '<td align="Left"><i>' + genus + ' ' + spp + '</i></td>');
        WriteLn(outfile, '<td align="Left">' + location_name + '</td>');
        WriteLn(outfile, '<td align="Center">' + location_latitude + '</td>');
        WriteLn(outfile, '<td align="Center">' + location_longitude + '</td>');
        WriteLn(outfile, '</tr>');
        Sleep(1000);
      end;
      QueryDataset1.Next;
      Inc(nrows);
    end;
    QueryConnection1.Disconnect;
    Screen.Cursor := crDefault;
  end;

  WriteLn(outfile, '</table>');
  WriteLn(outfile, '<br>--<br>');
  WriteLn(outfile, strTotalRecords + IntToStr(nrows) + '<br>');
  WriteLn(outfile, strCoded + IntToStr(coded) + '<br>');
  WriteLn(outfile, strUncoded + IntToStr(uncoded) + '<br>');
  WriteLn(outfile, '</body>');
  WriteLn(outfile, '</html>');
  CloseFile(outfile);
  if FileExists('temp.csv') then
    DeleteFile('temp.csv');
  if FileExists('temp.xml') then
    DeleteFile('temp.xml');
end;

procedure Stats(fname: string; option: integer; graph_it: boolean);
var
  d1, d2, c, h, pie, d, m, j: double;
  i, f, g, s, l, n, nspp, nind, vconta, freq, counter, nobs, cumsum: integer;
  nseq, pseq, seqlen, countA, countT, countC, countG, cdslen: integer;
  FAMILY, SPECIES, SPECIMEN, SAMPLE, SEQUENCE: string;
  vnome, vspp, vgenero, vcf, vespecie, vautor1, vsubsp, vinfranome,
  vautor2, cons, vloc, vdescriptor, variable, vunit, vseq, vsigla, figf: string;
  sout: ansistring;
  vetfam, vetgen, vetloc: array of string;
  ni: array of integer;
  datain: array of double;
  genera, names, fields, cds, dups, x, y, rdata: TStringList;
  nuc: TDNA;
  outfile, rscript: TextFile;
begin
  AssignFile(outfile, fname);
  Rewrite(outfile);
  Header(outfile, strReportStatistics);
  case option of

    1:
    begin
      WriteLn(outfile, '<br>');
      WriteLn(outfile, strFamilyStatistics);
      WriteLn(outfile, '<br><br>');
      WriteLn(outfile,
        '<table border=1 cellspacing=1 cellpadding=1 width="100%">');
      WriteLn(outfile, '<tr><th>' + strFamily + '</th>');
      WriteLn(outfile, '<th>' + strNoSpp + '</th>');
      WriteLn(outfile, '<th>' + strPercentSpp + '</th>');
      WriteLn(outfile, '</tr>');

      if FileExists('temp.csv') then
        DeleteFile('temp.csv');

      with MainForm.MultiDoc.ActiveObject as TMDIChild do
      begin
        CSVExporter1.FileName := 'temp.csv';
        CSVExporter1.Execute;
        QueryConnection1.DatabasePath := GetCurrentDir;
        QueryConnection1.Connect;
        QueryDataSet1.TableName := 'temp';
        QueryDataset1.LoadFromTable;
        FAMILY := QueryDataset1.Fields[2].FieldName;
        SPECIES := QueryDataset1.Fields[3].FieldName;
        QueryDataset1.SQL.Text :=
          'SELECT COUNT(' + SPECIES + ') FROM temp GROUP BY ' + SPECIES;
        QueryDataset1.QueryExecute;
        s := QueryDataset1.RecordCount;
        QueryDataset1.SQL.Text :=
          'SELECT ' + FAMILY + ' FROM temp GROUP BY ' + FAMILY + ' ORDER BY ' + FAMILY;
        QueryDataset1.QueryExecute;
        //QueryDataset1.SaveToTable;
        f := QueryDataset1.RecordCount;
      end;

      SetLength(vetfam, f);
      with MainForm.MultiDoc.ActiveObject as TMDIChild do
      begin
        i := 0;
        QueryDataset1.First;
        while not QueryDataset1.EOF do
        begin
          vetfam[i] := Trim(QueryDataset1.Fields[0].AsString);
          QueryDataset1.Next;
          Inc(i);
        end;

        if graph_it then
        begin
          rdata := TStringList.Create;
          rdata.Add('FAMILY,CASES');
        end;
        for i := 0 to f - 1 do
        begin
          CSVExporter1.FileName := 'temp.csv';
          CSVExporter1.Execute;
          QueryConnection1.DatabasePath := GetCurrentDir;
          QueryConnection1.Connect;
          QueryDataSet1.TableName := 'temp';
          QueryDataset1.LoadFromTable;
          QueryDataset1.SQL.Text :=
            'SELECT COUNT(' + SPECIES + ') FROM temp WHERE ' +
            FAMILY + ' = ' + QuotedStr(vetfam[i]) + ' GROUP BY ' + SPECIES;
          QueryDataset1.QueryExecute;
          //QueryDataset1.SaveToTable;
          vconta := QueryDataset1.RecordCount;
          WriteLn(outfile, '<tr>');
          WriteLn(outfile, '<td align="Left">' + vetfam[i] + '</td>');
          WriteLn(outfile, '<td align="Center">' + IntToStr(vconta) +
            strSpeciesCount + '</td>');
          WriteLn(outfile, '<td align="Center">' +
            FloatToStrF(Percent(vconta, s), ffFixed, 5, 2) + ' %' + '</td>');
          WriteLn(outfile, '</tr>');
          if graph_it then
            rdata.Append(vetfam[i] + ',' + FloatToStr(vconta * 10.0));
        end;
        QueryConnection1.Disconnect;

        if graph_it then
        begin
          Screen.Cursor := crHourGlass;
          rdata.SaveToFile('rdata.csv');
          rdata.Free;
          AssignFile(rscript, 'barplot.R');
          Rewrite(rscript);
          WriteLn(rscript, 'data <- read.csv("rdata.csv")');
          WriteLn(rscript, 'ppi <- 100');
          figf := GetFileNameWithoutExt(fname) + '.png';
          WriteLn(rscript, 'png("' + figf + '", width=6*ppi, height=6*ppi, res=ppi)');
          WriteLn(rscript, 'par(mai=c(1,2,1,1))');
          WriteLn(rscript,
            'barplot(data[,2], main=iconv("' + strFamilyTitle +
            '", , from=''UTF-8'', to=''latin1''), xlab=iconv("' +
            strXLabel +
            '", from=''UTF-8'', to=''latin1''), horiz=TRUE, names.arg=c(data[,1]), las=1, col="cornflowerblue")');
          WriteLn(rscript, 'invisible(dev.off())');
          CloseFile(rscript);
          if not RunCommand(RPath, ['--vanilla', 'barplot.R'], sout, [poNoConsole]) then
            MessageDlg(strError, Format(strNotExecute, ['R']), mtError, [mbOK], 0);
          Screen.Cursor := crDefault;
        end;
      end;

      WriteLn(outfile, '</table>');
      WriteLn(outfile, '<br><br>--<br>');
      WriteLn(outfile, strTotalFamilies + IntToStr(f) + strFamiliesCount + '<br>');
      WriteLn(outfile, strTotalSpecies + IntToStr(s) + strSpeciesCount + '<br>');
      if graph_it then
        WriteLn(outfile, '<p align="center"><img src="' + figf + '"></p>');
    end;

    2:
    begin
      WriteLn(outfile, '<br>');
      WriteLn(outfile, strGenusStatistics);
      WriteLn(outfile, '<br><br>');
      WriteLn(outfile,
        '<table border=1 cellspacing=1 cellpadding=1 width="100%">');
      WriteLn(outfile, '<tr><th>' + strGenus + '</th>');
      WriteLn(outfile, '<th>' + strNoSpp + '</th>');
      WriteLn(outfile, '<th>' + strPercentSpp + '</th>');
      WriteLn(outfile, '</tr>');

      if FileExists('temp.csv') then
        DeleteFile('temp.csv');

      with MainForm.MultiDoc.ActiveObject as TMDIChild do
      begin
        CSVExporter1.FileName := 'temp.csv';
        CSVExporter1.Execute;
        QueryConnection1.DatabasePath := GetCurrentDir;
        QueryConnection1.Connect;
        QueryDataSet1.TableName := 'temp';
        QueryDataset1.LoadFromTable;
        SPECIES := QueryDataset1.Fields[3].FieldName;
        QueryDataset1.SQL.Text :=
          'SELECT ' + SPECIES + ' FROM temp GROUP BY ' + SPECIES +
          ' ORDER BY ' + SPECIES;
        QueryDataset1.QueryExecute;
        //QueryDataset1.SaveToTable;
        s := QueryDataset1.RecordCount;
      end;

      genera := TStringList.Create;
      genera.Sorted := True;
      genera.Duplicates := dupIgnore;
      with MainForm.MultiDoc.ActiveObject as TMDIChild do
      begin
        QueryDataset1.First;
        while not QueryDataset1.EOF do
        begin
          genera.Add(Trim(QueryDataset1.Fields[0].AsString.Split(' ')[0]));
          QueryDataset1.Next;
        end;
        g := genera.Count;
        SetLength(vetgen, g);
        for i := 0 to g - 1 do
          vetgen[i] := genera[i];
        genera.Free;

        if graph_it then
        begin
          rdata := TStringList.Create;
          rdata.Add('GENUS,CASES');
        end;
        for i := 0 to g - 1 do
        begin
          CSVExporter1.FileName := 'temp.csv';
          CSVExporter1.Execute;
          QueryConnection1.DatabasePath := GetCurrentDir;
          QueryConnection1.Connect;
          QueryDataSet1.TableName := 'temp';
          QueryDataset1.LoadFromTable;
          QueryDataset1.SQL.Text :=
            'SELECT COUNT(' + SPECIES + ') FROM temp WHERE MID(' +
            SPECIES + ', 1, ' + IntToStr(Length(vetgen[i])) + ') = ' +
            QuotedStr(vetgen[i]) + ' GROUP BY ' + SPECIES;
          QueryDataset1.QueryExecute;
          //QueryDataset1.SaveToTable;
          vconta := QueryDataset1.RecordCount;
          WriteLn(outfile, '<tr>');
          WriteLn(outfile, '<td align="Left"><i>' + vetgen[i] + '</i></td>');
          WriteLn(outfile, '<td align="Center">' + IntToStr(vconta) +
            strSpeciesCount + '</td>');
          WriteLn(outfile, '<td align="Center">' +
            FloatToStrF(Percent(vconta, s), ffFixed, 5, 2) + ' %' + '</td>');
          WriteLn(outfile, '</tr>');
          if graph_it then
            rdata.Append(vetgen[i] + ',' + FloatToStr(vconta * 10.0));
        end;
        QueryConnection1.Disconnect;

        if graph_it then
        begin
          Screen.Cursor := crHourGlass;
          rdata.SaveToFile('rdata.csv');
          rdata.Free;
          AssignFile(rscript, 'barplot.R');
          Rewrite(rscript);
          WriteLn(rscript, 'data <- read.csv("rdata.csv")');
          WriteLn(rscript, 'ppi <- 100');
          figf := GetFileNameWithoutExt(fname) + '.png';
          WriteLn(rscript, 'png("' + figf + '", width=6*ppi, height=6*ppi, res=ppi)');
          WriteLn(rscript, 'par(mai=c(1,2,1,1))');
          WriteLn(rscript,
            'barplot(data[,2], main=iconv("' + strGenusTitle +
            '", from=''UTF-8'', to=''latin1''), xlab=iconv("' + strXLabel +
            '", from=''UTF-8'', to=''latin1''), horiz=TRUE, names.arg=c(data[,1]), las=1, col="cornflowerblue")');
          WriteLn(rscript, 'invisible(dev.off())');
          CloseFile(rscript);
          if not RunCommand(RPath, ['--vanilla', 'barplot.R'], sout, [poNoConsole]) then
            MessageDlg(strError, Format(strNotExecute, ['R']), mtError, [mbOK], 0);
          Screen.Cursor := crDefault;
        end;
      end;

      WriteLn(outfile, '</table>');
      WriteLn(outfile, '<br><br>--<br>');
      WriteLn(outfile, strTotalGenera + IntToStr(g) + strGeneraCount + '<br>');
      WriteLn(outfile, strTotalSpecies + IntToStr(s) + strSpeciesCount + '<br>');
      if graph_it then
        WriteLn(outfile, '<p align="center"><img src="' + figf + '"></p>');
    end;

    3:
    begin
      WriteLn(outfile, '<br>');
      WriteLn(outfile, strSpeciesStatistics);
      WriteLn(outfile, '<br><br>');
      WriteLn(outfile,
        '<table border=1 cellspacing=1 cellpadding=1 width="100%">');
      WriteLn(outfile, '<tr><th>' + strSpecies + '</th>');
      WriteLn(outfile, '<th>' + strNoInd + '</th>');
      WriteLn(outfile, '<th>' + strDom + '</th>');
      WriteLn(outfile, '<th>' + strFreq + '</th>');
      WriteLn(outfile, '<th>' + strFR + '</th>');
      WriteLn(outfile, '<th>' + strCategory + '</th>');
      WriteLn(outfile, '</tr>');

      if FileExists('temp.csv') then
        DeleteFile('temp.csv');

      with MainForm.MultiDoc.ActiveObject as TMDIChild do
      begin
        CSVExporter1.FileName := 'temp.csv';
        CSVExporter1.Execute;
        QueryConnection1.DatabasePath := GetCurrentDir;
        QueryConnection1.Connect;
        QueryDataSet1.TableName := 'temp';
        QueryDataset1.LoadFromTable;
        SAMPLE := QueryDataset1.Fields[0].FieldName;
        SPECIES := QueryDataset1.Fields[3].FieldName;
        QueryDataset1.SQL.Text :=
          'SELECT ' + SAMPLE + ' FROM temp GROUP BY ' + SAMPLE;
        QueryDataset1.QueryExecute;
        l := QueryDataset1.RecordCount;
        QueryDataset1.SQL.Text :=
          'SELECT ' + SPECIES + ' FROM temp GROUP BY ' + SPECIES;
        QueryDataset1.QueryExecute;
        s := QueryDataset1.RecordCount;
        QueryDataset1.SQL.Text := 'SELECT * FROM temp';
        QueryDataset1.QueryExecute;
        n := QueryDataset1.RecordCount;
        QueryDataset1.SQL.Text :=
          'SELECT ' + SPECIES + ' FROM temp GROUP BY ' + SPECIES +
          ' ORDER BY ' + SPECIES;
        QueryDataset1.QueryExecute;
        //QueryDataset1.SaveToTable;
      end;

      names := TStringList.Create;
      with MainForm.MultiDoc.ActiveObject as TMDIChild do
      begin
        QueryDataset1.First;
        while not QueryDataset1.EOF do
        begin
          names.Add(Trim(QueryDataset1.Fields[0].AsString));
          QueryDataset1.Next;
        end;
        QueryConnection1.Disconnect;
      end;

      if FileExists('temp.csv') then
        DeleteFile('temp.csv');

      with MainForm.MultiDoc.ActiveObject as TMDIChild do
      begin
        if graph_it then
        begin
          rdata := TStringList.Create;
          rdata.Add('SPECIES,CASES');
        end;
        SetLength(ni, names.Count);
        for i := 0 to names.Count - 1 do
        begin
          vspp := names[i];
          ParseName(vspp, vgenero, vcf, vespecie, vautor1, vsubsp, vinfranome, vautor2);
          vnome := vgenero + ' ' + vespecie + ' ' +
            IfThen(Length(vinfranome) > 0, vinfranome, '');
          if graph_it then
            vsigla := LeftStr(vgenero, 4) + '.' + LeftStr(vespecie, 4);
          CSVExporter1.FileName := 'temp.csv';
          CSVExporter1.Execute;
          QueryConnection1.DatabasePath := GetCurrentDir;
          QueryConnection1.Connect;
          QueryDataSet1.TableName := 'temp';
          QueryDataset1.LoadFromTable;
          QueryDataset1.SQL.Text :=
            'SELECT * FROM temp WHERE ' + SPECIES + ' = ' + QuotedStr(vspp);
          QueryDataset1.QueryExecute;
          //QueryDataset1.SaveToTable;
          nind := QueryDataset1.RecordCount;
          freq := nind;
          ni[i] := nind;
          c := Percent(freq, l);
          if (c > 50.0) then
            cons := strConstant
          else if (c >= 25.0) and (c <= 50.0) then
            cons := strAccessory
          else if (c < 25.0) then
            cons := strAccidental;
          WriteLn(outfile, '<tr>');
          WriteLn(outfile, '<td align="Left"><i>' + vnome + '</i></td>');
          WriteLn(outfile, '<td align="Center">' + IntToStr(nind) +
            strSpecimensCount + '</td>');
          WriteLn(outfile, '<td align="Center">' +
            FloatToStrF(Percent(nind, n), ffFixed, 5, 1) + ' %' + '</td>');
          WriteLn(outfile, '<td align="Center">' + IntToStr(freq) + '</td>');
          WriteLn(outfile, '<td align="Center">' +
            FloatToStrF(c, ffFixed, 5, 1) + '</td>');
          WriteLn(outfile, '<td align="Center">' + cons + '</td>');
          WriteLn(outfile, '</tr>');
          if graph_it then
            rdata.Append(vsigla + ',' + FloatToStr(freq * 10.0));
        end;
        QueryConnection1.Disconnect;
      end;

      if graph_it then
      begin
        Screen.Cursor := crHourGlass;
        rdata.SaveToFile('rdata.csv');
        rdata.Free;
        AssignFile(rscript, 'barplot.R');
        Rewrite(rscript);
        WriteLn(rscript, 'data <- read.csv("rdata.csv")');
        WriteLn(rscript, 'ppi <- 100');
        figf := GetFileNameWithoutExt(fname) + '.png';
        WriteLn(rscript, 'png("' + figf + '", width=6*ppi, height=6*ppi, res=ppi)');
        WriteLn(rscript, 'par(mai=c(1,2,1,1))');
        WriteLn(rscript,
          'barplot(data[,2], main=iconv("' + strSpeciesTitle +
          '", from=''UTF-8'', to=''latin1''), xlab=iconv("' + strXLabel +
          '", from=''UTF-8'', to=''latin1''), horiz=TRUE, names.arg=c(data[,1]), las=1, col="cornflowerblue")');
        WriteLn(rscript, 'invisible(dev.off())');
        CloseFile(rscript);
        if not RunCommand(RPath, ['--vanilla', 'barplot.R'], sout, [poNoConsole]) then
          MessageDlg(strError, Format(strNotExecute, ['R']), mtError, [mbOK], 0);
        Screen.Cursor := crDefault;
      end;

      names.Free;
      WriteLn(outfile, '</table>');
      WriteLn(outfile, '<br><br>--<br>');
      WriteLn(outfile, strTotalSpecies + IntToStr(s) + strSpeciesCount + '<br>');
      WriteLn(outfile, strTotalSpecimens + IntToStr(n) + strSpecimensCount + '<br>');

      WriteLn(outfile, '<br>' + strDiversityAnalysis + '<br>');
      BioDiv(s, n, ni, d1, d2, c, h, pie, d, m, j);
      WriteLn(outfile, '<br>' + strSpeciesRichness + '<br>');
      WriteLn(outfile, '<br>');
      WriteLn(outfile, strMargalefIndex + FloatToStrF(d1, ffFixed, 7, 5) + '<br>');
      WriteLn(outfile, StrMenhinickIndex + FloatToStrF(d2, ffFixed, 7, 5) + '<br>');
      WriteLn(outfile, '<br>' + strDiversity + '<br>');
      WriteLn(outfile, '<br>');
      WriteLn(outfile, strSimpsonIndex + FloatToStrF(c, ffFixed, 7, 5) + '<br>');
      WriteLn(outfile, strShannonIndex + FloatToStrF(h, ffFixed, 7, 5) + '<br>');
      WriteLn(outfile, strHurlbertIndex + FloatToStrF(pie, ffFixed, 7, 5) + '<br>');
      WriteLn(outfile, strBergerIndex + FloatToStrF(d, ffFixed, 7, 5) + '<br>');
      WriteLn(outfile, strMcIntoshIndex + FloatToStrF(m, ffFixed, 7, 5) + '<br>');
      WriteLn(outfile, '<br>' + strEvenness + '<br>');
      WriteLn(outfile, '<br>');
      WriteLn(outfile, strEvennessIndex + FloatToStrF(j, ffFixed, 7, 5) + '<br>');
      if graph_it then
        WriteLn(outfile, '<p align="center"><img src="' + figf + '"></p>');
    end;

    4:
    begin
      WriteLn(outfile, '<br>');
      WriteLn(outfile, strSampleStatistics);
      WriteLn(outfile, '<br><br>');
      WriteLn(outfile,
        '<table border=1 cellspacing=1 cellpadding=1 width="100%">');
      WriteLn(outfile, '<tr><th>' + strSample + '</th>');
      WriteLn(outfile, '<th>' + strNoSpp + '</th>');
      WriteLn(outfile, '<th>' + strNoInd + '</th>');
      WriteLn(outfile, '<th>' + strPercentSpp + '</th>');
      WriteLn(outfile, '</tr>');

      if FileExists('temp.csv') then
        DeleteFile('temp.csv');

      with MainForm.MultiDoc.ActiveObject as TMDIChild do
      begin
        CSVExporter1.FileName := 'temp.csv';
        CSVExporter1.Execute;
        QueryConnection1.DatabasePath := GetCurrentDir;
        QueryConnection1.Connect;
        QueryDataSet1.TableName := 'temp';
        QueryDataset1.LoadFromTable;
        SAMPLE := QueryDataset1.Fields[0].FieldName;
        SPECIES := QueryDataset1.Fields[3].FieldName;
        QueryDataset1.SQL.Text :=
          'SELECT ' + SPECIES + ' FROM temp GROUP BY ' + SPECIES;
        QueryDataset1.QueryExecute;
        s := QueryDataset1.RecordCount;
        QueryDataset1.QueryExecute;
        QueryDataset1.SQL.Text :=
          'SELECT ' + SAMPLE + ' FROM temp GROUP BY ' + SAMPLE;
        QueryDataset1.QueryExecute;
        l := QueryDataset1.RecordCount;
        //QueryDataset1.SaveToTable;
      end;

      SetLength(vetloc, l);
      with MainForm.MultiDoc.ActiveObject as TMDIChild do
      begin
        i := 0;
        QueryDataset1.First;
        while not QueryDataset1.EOF do
        begin
          vetloc[i] := Trim(QueryDataset1.Fields[0].AsString);
          QueryDataset1.Next;
          Inc(i);
        end;

        if FileExists('temp.csv') then
          DeleteFile('temp.csv');

        if graph_it then
        begin
          x := TStringList.Create;
          x.NameValueSeparator := ',';
          x.Add('1,0');
          y := TStringList.Create;
          y.NameValueSeparator := ',';
        end;
        for i := 0 to l - 1 do
        begin
          vloc := vetloc[i];
          CSVExporter1.FileName := 'temp.csv';
          CSVExporter1.Execute;
          QueryConnection1.DatabasePath := GetCurrentDir;
          QueryConnection1.Connect;
          QueryDataSet1.TableName := 'temp';
          QueryDataset1.LoadFromTable;
          QueryDataset1.SQL.Text :=
            'SELECT COUNT(' + SPECIES + ') ' + 'FROM temp WHERE ' +
            SAMPLE + ' = ' + QuotedStr(vloc) + ' GROUP BY ' + SPECIES;
          QueryDataset1.QueryExecute;
          nspp := QueryDataset1.RecordCount;
          QueryDataset1.SQL.Text :=
            'SELECT COUNT(' + SPECIES + ') ' + 'FROM temp WHERE ' +
            SAMPLE + ' = ' + QuotedStr(vloc);
          QueryDataset1.QueryExecute;
          nind := QueryDataset1.Fields[0].AsInteger;
          //QueryDataset1.SaveToTable;
          WriteLn(outfile, '<tr>');
          WriteLn(outfile, '<td align="Left">' + vloc + '</td>');
          WriteLn(outfile, '<td align="Center">' + IntToStr(nspp) +
            strSpeciesCount + '</td>');
          WriteLn(outfile, '<td align="Center">' + IntToStr(nind) +
            strSpecimensCount + '</td>');
          WriteLn(outfile, '<td align="Center">' +
            FloatToStrF(Percent(nspp, s), ffFixed, 5, 1) + '</td>');
          WriteLn(outfile, '</tr>');
          if graph_it then
            x.Add(IntToStr(i + 2) + ',' + IntToStr(nspp));
        end;
        QueryConnection1.Disconnect;

        if graph_it then
        begin
          Screen.Cursor := crHourGlass;
          y.Add(x.ValueFromIndex[0]);
          for i := 1 to x.Count - 1 do
          begin
            cumsum := StrToInt(x.ValueFromIndex[i]) +
              StrToInt(x.ValueFromIndex[i - 1]);
            y.Add(IntToStr(cumsum));
          end;
          NaturalSort(y);
          rdata := TStringList.Create;
          i := 0;
          while i < y.Count do
          begin
            rdata.Add(IntToStr(i + 1) + ',' + y[i]);
            Inc(i);
          end;
          rdata.Insert(0, 'X,Y');
          rdata.SaveToFile('rdata.csv');
          rdata.Free;
          x.Free;
          y.Free;
          AssignFile(rscript, 'lineplot.R');
          Rewrite(rscript);
          WriteLn(rscript, 'data <- read.csv("rdata.csv")');
          WriteLn(rscript, 'ppi <- 100');
          figf := GetFileNameWithoutExt(fname) + '.png';
          WriteLn(rscript, 'png("' + figf + '", width=6*ppi, height=6*ppi, res=ppi)');
          //WriteLn(rscript, 'par(mai=c(1,2,1,1))');
          WriteLn(rscript,
            'plot(data, main=iconv("' + strPlotTitle +
            '", from=''UTF-8'', to=''latin1''), xlab="' + strPlotXLabel +
            '", ylab=iconv("' + strPlotYLabel +
            '", from=''UTF-8'', to=''latin1''), type="l", lwd=2, col="forestgreen", xlim=c(1, max(data[,1])+1), ylim=c(0, max(data[,2])+1))');
          WriteLn(rscript, 'invisible(dev.off())');
          CloseFile(rscript);
          if not RunCommand(RPath, ['--vanilla', 'lineplot.R'], sout, [poNoConsole]) then
            MessageDlg(strError, Format(strNotExecute, ['R']), mtError, [mbOK], 0);
          Screen.Cursor := crDefault;
        end;
      end;
      WriteLn(outfile, '</table>');
      WriteLn(outfile, '<br><br>--<br>');
      WriteLn(outfile, strTotalSamples + IntToStr(l) + strSamplesCount + '<br>');
      WriteLn(outfile, strTotalSpecies + IntToStr(s) + strSpeciesCount + '<br>');
      if graph_it then
        WriteLn(outfile, '<p align="center"><img src="' + figf + '"></p>');
    end;

    5:
    begin
      WriteLn(outfile, '<br>');
      WriteLn(outfile, strDescriptorStatistics);
      WriteLn(outfile, '<br><br>');
      WriteLn(outfile, '<table border=1 cellspacing=1 cellpadding=1 width="100%">');
      WriteLn(outfile, '<tr><th>' + strNumber + '</th>');
      WriteLn(outfile, '<th>' + strDescriptor + '</th>');
      WriteLn(outfile, '<th>' + strType + '</th>');
      WriteLn(outfile, '<th>' + strUnit + '</th>');
      WriteLn(outfile, '<th>' + strMean + '</th>');
      WriteLn(outfile, '<th>' + strStdDev + '</th>');
      WriteLn(outfile, '<th>' + strMinimum + '</th>');
      WriteLn(outfile, '<th>' + strMaximum + '</th>');
      WriteLn(outfile, '<th>' + strCases + '</th></tr>');

      with MainForm.MultiDoc.ActiveObject as TMDIChild do
      begin
        counter := 0;
        SetLength(datain, Dataset1.RecordCount);
        for i := 12 to Dataset1.FieldCount - 1 do
        begin
          if Pos('(', Dataset1.Fields[i].FieldName) > 0 then
          begin
            vdescriptor := Copy(Dataset1.Fields[i].FieldName, 1,
              Pos(' ', Dataset1.Fields[i].FieldName) - 1);
            vunit := GetUnit(Dataset1.Fields[i].FieldName);
          end
          else
          begin
            vdescriptor := Dataset1.Fields[i].FieldName;
            vunit := '-';
          end;
          ProcessArray(Dataset1, i, datain, nobs);
          WriteLn(outfile, '<tr>');
          WriteLn(outfile, '<td align="Left">' + IntToStr(counter + 1) + '</td>');
          WriteLn(outfile, '<td align="Left">' +
            AnsiProperCase(vdescriptor, StdWordDelims) + '</td>');
          if Dataset1.Fields[i] is TFloatField then
            WriteLn(outfile, '<td align="Left">' + strNumeric + '</td>')
          else if Dataset1.Fields[i] is TStringField then
            WriteLn(outfile, '<td align="Left">' + strText + '</td>')
          else
            WriteLn(outfile, '<td align="Left">' + strUnknown + '</td>');
          WriteLn(outfile, '<td align="Center">' + vunit + '</td>');
          WriteLn(outfile, '<td align="Right">' +
            FloatToStrF(Mean(datain), ffFixed, 5, 3) + '</td>');
          WriteLn(outfile, '<td align="Right">' +
            FloatToStrF(StdDev(datain), ffFixed, 5, 3) + '</td>');
          WriteLn(outfile, '<td align="Right">' +
            FloatToStrF(MinValue(datain), ffFixed, 5, 2) + '</td>');
          WriteLn(outfile, '<td align="Right">' +
            FloatToStrF(MaxValue(datain), ffFixed, 5, 2) + '</td>');
          WriteLn(outfile, '<td align="Right">' + IntToStr(nobs) + '</td>');
          WriteLn(outfile, '</tr>');
          Inc(counter);
        end;
      end;

      WriteLn(outfile, '</table>');
      WriteLn(outfile, '<br><br>--<br>');
      WriteLn(outfile, strTotalDescriptors + IntToStr(counter) +
        strDescriptorsCount + '<br>');
    end;

    6:
    begin
      WriteLn(outfile, '<br>');
      WriteLn(outfile, strSequenceStatistics);
      WriteLn(outfile, '<br><br>');
      WriteLn(outfile, '<br>' + strNucleotideComposition + '<br><br>');
      WriteLn(outfile, '<table border=1 cellspacing=1 cellpadding=1 width="100%">');
      WriteLn(outfile, '<tr><th>' + strSpecimen + '</th>');
      WriteLn(outfile, '<th>A</th>');
      WriteLn(outfile, '<th>%</th>');
      WriteLn(outfile, '<th>T</th>');
      WriteLn(outfile, '<th>%</th>');
      WriteLn(outfile, '<th>C</th>');
      WriteLn(outfile, '<th>%</th>');
      WriteLn(outfile, '<th>G</th>');
      WriteLn(outfile, '<th>%</th>');
      WriteLn(outfile, '<th>Total</th>');
      WriteLn(outfile, '<th>%CG</th>');
      WriteLn(outfile, '</tr>');

      with MainForm.MultiDoc.ActiveObject as TMDIChild do
      begin
        fields := TStringList.Create;
        for i := 0 to Dataset1.FieldCount - 1 do
          fields.Add(Copy(Dataset1.Fields[i].FieldName, 1, 3));
        pseq := fields.IndexOf(UpperCase('SEQ'));
        fields.Free;

        CSVExporter1.FileName := 'temp.csv';
        CSVExporter1.Execute;
        QueryConnection1.DatabasePath := GetCurrentDir;
        QueryConnection1.Connect;
        QueryDataSet1.TableName := 'temp';
        QueryDataset1.LoadFromTable;
        SPECIMEN := QueryDataset1.Fields[1].FieldName;
        SEQUENCE := QueryDataset1.Fields[pseq].FieldName;
        QueryDataset1.SQL.Text :=
          'SELECT ' + SPECIMEN + ', ' + SEQUENCE + ' FROM temp ORDER BY ' + SPECIMEN;
        QueryDataset1.QueryExecute;
        nseq := QueryDataset1.RecordCount;
        QueryDataset1.SaveToTable;

        nseq := 0;
        while not QueryDataset1.EOF do
        begin
          vseq := QueryDataset1.Fields[1].AsString;
          if Length(vseq) > 0 then
          begin
            Inc(nseq);
            seqlen := Length(vseq);
            countA := vseq.CountChar('A');
            countT := vseq.CountChar('T');
            countC := vseq.CountChar('C');
            countG := vseq.CountChar('G');
            nuc := TDNA.Create(vseq);
            WriteLn(outfile, '<tr>');
            WriteLn(outfile, '<td align="Left">' +
              QueryDataset1.Fields[0].AsString + '</td>');
            WriteLn(outfile, '<td align="Center">' + IntToStr(countA) + '</td>');
            WriteLn(outfile, '<td align="Center">' + FloatToStrF(
              (countA * 100.0 / seqlen), ffFixed, 5, 1) + '</td>');
            WriteLn(outfile, '<td align="Center">' + IntToStr(countT) + '</td>');
            WriteLn(outfile, '<td align="Center">' + FloatToStrF(
              (countT * 100.0 / seqlen), ffFixed, 5, 1) + '</td>');
            WriteLn(outfile, '<td align="Center">' + IntToStr(countC) + '</td>');
            WriteLn(outfile, '<td align="Center">' + FloatToStrF(
              (countC * 100.0 / seqlen), ffFixed, 5, 1) + '</td>');
            WriteLn(outfile, '<td align="Center">' + IntToStr(countG) + '</td>');
            WriteLn(outfile, '<td align="Center">' + FloatToStrF(
              (countG * 100.0 / seqlen), ffFixed, 5, 1) + '</td>');
            WriteLn(outfile, '<td align="Center">' + IntToStr(seqlen) + '</td>');
            WriteLn(outfile, '<td align="Center">' +
              FloatToStrF(nuc.GC, ffFixed, 5, 1) + '</td>');
            WriteLn(outfile, '</tr>');
          end;
          QueryDataset1.Next;
        end;
      end;

      WriteLn(outfile, '</table>');
      WriteLn(outfile, '<br><br>--<br>');
      WriteLn(outfile, strTotalSequences + IntToStr(nseq) + strSequencesCount + '<br>');
      WriteLn(outfile, '<br>' + strCodonFrequency + '<br>');

      with MainForm.MultiDoc.ActiveObject as TMDIChild do
      begin
        QueryDataset1.First;
        while not QueryDataset1.EOF do
        begin
          vseq := QueryDataset1.Fields[1].AsString;
          if Length(vseq) > 0 then
          begin
            WriteLn(outfile,
              '<br><table border=1 cellspacing=1 cellpadding=1 width="100%">');
            WriteLn(outfile, '<tr><th>' + strSpecimen + '</th>');
            WriteLn(outfile, '<th>' + strCodon + '</th>');
            WriteLn(outfile, '<th>' + strFreq + '</th>');
            WriteLn(outfile, '<th>%</th>');
            WriteLn(outfile, '</tr>');
            nuc := TDNA.Create(vseq);
            cds := nuc.Codons;
            dups := nuc.Frequency(cds);
            cdslen := dups.Count;
            WriteLn(outfile, '<tr>');
            WriteLn(outfile, '<td align="Center">' +
              QueryDataset1.Fields[0].AsString + '</td>');
            WriteLn(outfile, '<td align="Left">&nbsp;</td>');
            WriteLn(outfile, '<td align="Left">&nbsp;</td>');
            WriteLn(outfile, '<td align="Left">&nbsp;</td>');
            WriteLn(outfile, '</tr>');
            for i := 0 to cdslen - 1 do
            begin
              WriteLn(outfile, '<tr>');
              WriteLn(outfile, '<td align="Left">&nbsp;</td>');
              WriteLn(outfile, '<td align="Center">' + dups.Names[i] + '</td>');
              WriteLn(outfile, '<td align="Center">' + dups.ValueFromIndex[i] + '</td>');
              WriteLn(outfile, '<td align="Center">' + FloatToStrF(
                (StrToFloatDef(dups.ValueFromIndex[i], 0.0) * 100.0 / cdslen),
                ffFixed, 5, 1) + '</td>');
              WriteLn(outfile, '</tr>');
            end;
            WriteLn(outfile, '</table>');
            WriteLn(outfile, '<br><br>--<br>');
            WriteLn(outfile, strTotalCodons + IntToStr(cdslen) +
              strCodonsCount + '<br>');
            if Assigned(cds) then
              cds.Free;
            if Assigned(dups) then
              dups.Free;
          end;
          QueryDataset1.Next;
        end;
        QueryConnection1.Disconnect;
      end;
    end;

    7:
    begin
      WriteLn(outfile, '<br>');
      WriteLn(outfile, strVariableStatistics);
      WriteLn(outfile, '<br><br>');
      WriteLn(outfile, '<table border=1 cellspacing=1 cellpadding=1 width="100%">');
      WriteLn(outfile, '<tr><th>' + strNumber + '</th>');
      WriteLn(outfile, '<th>' + strVariable + '</th>');
      WriteLn(outfile, '<th>' + strType + '</th>');
      WriteLn(outfile, '<th>' + strUnit + '</th>');
      WriteLn(outfile, '<th>' + strMean + '</th>');
      WriteLn(outfile, '<th>' + strSD + '</th>');
      WriteLn(outfile, '<th>' + strCV + '</th>');
      WriteLn(outfile, '<th>' + strMin + '</th>');
      WriteLn(outfile, '<th>' + strMax + '</th>');
      WriteLn(outfile, '<th>' + strObs + '</th></tr>');

      with MainForm.MultiDoc.ActiveObject as TMDIChild do
      begin
        counter := 0;
        SetLength(datain, Dataset2.RecordCount);
        for i := 4 to Dataset2.FieldCount - 1 do
        begin
          if Pos('(', Dataset2.Fields[i].FieldName) > 0 then
          begin
            variable := Copy(Dataset2.Fields[i].FieldName, 1,
              Pos(' ', Dataset2.Fields[i].FieldName) - 1);
            vunit := GetUnit(Dataset1.Fields[i].FieldName);
          end
          else
          begin
            variable := Dataset2.Fields[i].FieldName;
            vunit := '-';
          end;
          ProcessArray(Dataset2, i, datain, nobs);
          WriteLn(outfile, '<tr>');
          WriteLn(outfile, '<td align="Left">' + IntToStr(counter + 1) + '</td>');
          WriteLn(outfile, '<td align="Left">' +
            AnsiProperCase(variable, StdWordDelims) + '</td>');
          if Dataset2.Fields[i] is TFloatField then
            WriteLn(outfile, '<td align="Left">' + strNumeric + '</td>')
          else if Dataset2.Fields[i] is TStringField then
            WriteLn(outfile, '<td align="Left">' + strText + '</td>')
          else
            WriteLn(outfile, '<td align="Left">' + strUnknown + '</td>');
          WriteLn(outfile, '<td align="Center">' + vunit + '</td>');
          WriteLn(outfile, '<td align="Right">' +
            FloatToStrF(Mean(datain), ffFixed, 5, 3) + '</td>');
          WriteLn(outfile, '<td align="Right">' +
            FloatToStrF(StdDev(datain), ffFixed, 5, 3) + '</td>');
          WriteLn(outfile, '<td align="Right">' + FloatToStrF(
            (StdDev(datain) / Mean(datain)) * 100.0, ffFixed, 5, 3) + '</td>');
          WriteLn(outfile, '<td align="Right">' +
            FloatToStrF(MinValue(datain), ffFixed, 5, 2) + '</td>');
          WriteLn(outfile, '<td align="Right">' +
            FloatToStrF(MaxValue(datain), ffFixed, 5, 1) + '</td>');
          WriteLn(outfile, '<td align="Right">' + IntToStr(nobs) + '</td>');
          WriteLn(outfile, '</tr>');
          Inc(counter);
        end;
      end;

      WriteLn(outfile, '</table>');
      WriteLn(outfile, '<br><br>--<br>');
      WriteLn(outfile, strTotalVariables + IntToStr(counter) +
        strVariablesCount + '<br>');
    end;
  end;

  WriteLn(outfile, '</body>');
  WriteLn(outfile, '</html>');
  CloseFile(outfile);
  if FileExists('temp.csv') then
    DeleteFile('temp.csv');
end;

end.

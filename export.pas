unit export;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, StdCtrls,
  StrUtils, Math, Dbf, DB, DOM, XMLRead, XMLWrite, XMLUtils{$IFDEF WINDOWS},
  ShpAPI129 in 'shpAPI129.pas'{$ENDIF};

function toEML(const filename: string): integer;
procedure toFitopac1(const filename: string; kind: integer;
  var nfam, nspp, nind: integer);
procedure toFitopac2(const filename: string; var nfam, nspp, nind, nsam: integer);
function toKML(const filename: string): integer;
{$IFDEF WINDOWS}
function toSHP(const filename: string): integer;
{$ENDIF}
function toRDE(const filename: string): integer;
procedure toCEP(const filename: string; var nrows, ncols: integer);
procedure toMVSP(const filename: string; var nrows, ncols: integer);
procedure toFPM(const filename: string; var nrows, ncols: integer);
procedure toCSV(const filename: string; var nrows, ncols: integer);

implementation

uses main, child, useful;

function toEML(const filename: string): integer;
var
  S: TStringStream;
  XML: TXMLDocument;
  lDATE, lSAMPLE: TStringList;
  FAMILY, SPECIES: string;
  fSAMPLE, fLATITUDE, fLONGITUDE, fELEVATION, fDATE: TField;
  bm: TBookmark;
  DatasetName1: string;
  vtitle, vdescription, vcountry, vstate, vprovince, vlocality, vauthor,
  vfullname, vname, vsurname, vinstitution, vrole, vaddress1, vaddress2,
  vzip, vcity, vuf, vphone, vemail, vurl, vfunding, vunit, vmethod,
  vsample, date1, date2: string;
  grpCheck, grpEval, subCheck, subEval: string;
  xmldoc, emlHeader, emlFooter, emlBody: string;
  i, recs, levels, valtitude, minalt, maxalt: integer;
  minlat, maxlat, minlon, maxlon: double;
  altitude: array of integer;
  latitude, longitude: array of double;
begin
  xmldoc := '';
  emlHeader := '<?xml version="1.0" encoding="UTF-8"?>' + LineEnding +
    '<eml:eml' + LineEnding + 'packageId="eml.1.1" system="knb"' +
    LineEnding + 'xmlns:eml="eml://ecoinformatics.org/eml-2.1.0"' +
    LineEnding + 'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"' +
    LineEnding + 'xmlns:stmml="http://www.xml-cml.org/schema/stmml-1.1"' +
    LineEnding + 'xsi:schemaLocation="eml://ecoinformatics.org/eml-2.1.0 eml.xsd">' +
    '<dataset>';
  xmldoc += emlHeader;

  vtitle := Metadata.Find('title').AsString;
  vdescription := Metadata.Find('description').AsString;
  vcountry := Metadata.Find('country').AsString;
  vstate := Metadata.Find('state').AsString;
  vprovince := Metadata.Find('province').AsString;
  vlocality := Metadata.Find('locality').AsString;
  vauthor := Metadata.Find('author').AsString;
  vfullname := vauthor.Split(',')[0];
  vname := Copy(vfullname, 1, RPos(' ', vfullname) - 1);
  vsurname := Copy(vfullname, RPos(' ', vfullname) + 1, Length(vfullname));
  vinstitution := Metadata.Find('institution').AsString;
  vrole := Metadata.Find('role').AsString;
  vaddress1 := Metadata.Find('address1').AsString;
  vaddress2 := Metadata.Find('address2').AsString;
  vzip := Metadata.Find('zip').AsString;
  vcity := Metadata.Find('city').AsString;
  vuf := Metadata.Find('uf').AsString;
  vphone := Metadata.Find('phone').AsString;
  vemail := Metadata.Find('email').AsString;
  vurl := Metadata.Find('website').AsString;
  vfunding := Metadata.Find('funding').AsString;

  { Level 1: Identification }
  emlBody :=
    '<title>' + vtitle + '</title>' + LineEnding + '<creator id="' +
    vsurname.toLower + ',' + vname.toLower + '">' + LineEnding +
    '<individualName>' + LineEnding + '<givenName>' + vname +
    '</givenName>' + LineEnding + '<surName>' + vsurname + '</surName>' +
    LineEnding + '</individualName>' + LineEnding + '<organizationName>' +
    vinstitution + '</organizationName>' + LineEnding + '<address>' +
    LineEnding + '<deliveryPoint>' + vaddress1 + '</deliveryPoint>' +
    LineEnding + '<deliveryPoint>' + vaddress2 + '</deliveryPoint>' +
    LineEnding + '<city>' + vcity + '</city>' + LineEnding +
    '<administrativeArea>' + vuf + '</administrativeArea>' + LineEnding +
    '<postalCode>' + vzip + '</postalCode>' + LineEnding + '<country>' +
    vcountry + '</country>' + LineEnding + '</address>' + LineEnding +
    '<phone phonetype="voice">' + vphone + '</phone>' + LineEnding +
    '<electronicMailAddress>' + vemail + '</electronicMailAddress>' +
    LineEnding + '</creator>' + LineEnding + '<pubDate>' +
    FormatDateTime('YYYY', Now) + '</pubDate>' + LineEnding +
    '<abstract>' + LineEnding + '<para>' + vdescription + '</para>' +
    LineEnding + '</abstract>';
  xmldoc += emlBody;

  levels := 0;
  with MainForm.MultiDoc.ActiveObject as TMDIChild do
  begin
    if not Dataset1.Active then
    begin
      DatasetName1 := Metadata.Find('dataset1').AsString;
      Dataset1.FileName := DatasetName1;
      Dataset1.AutoFieldDefs := True;
      Dataset1.Open;
    end;
    recs := Dataset1.RecordCount;
    bm := Dataset1.GetBookmark;
    Dataset1.DisableControls;

    CSVExporter1.FileName := 'temp.csv';
    CSVExporter1.Execute;
    QueryConnection1.DatabasePath := GetCurrentDir;
    QueryConnection1.Connect;
    QueryDataSet1.TableName := 'temp';
    QueryDataset1.LoadFromTable;
    FAMILY := QueryDataset1.Fields[2].FieldName;
    SPECIES := QueryDataset1.Fields[3].FieldName;
    QueryDataset1.SQL.Text :=
      'SELECT ' + FAMILY + ', ' + SPECIES + ' FROM temp GROUP BY ' +
      FAMILY + ', ' + SPECIES + ' ORDER BY ' + FAMILY + ', ' + SPECIES;
    QueryDataset1.QueryExecute;
    QueryDataset1.SaveToTable;

    fLATITUDE := Dataset1.Fields[9];
    fLONGITUDE := Dataset1.Fields[10];
    fELEVATION := Dataset1.Fields[11];

    SetLength(latitude, recs);
    SetLength(longitude, recs);
    SetLength(altitude, recs);
    i := 0;
    Dataset1.First;
    while not Dataset1.EOF do
    begin
      if Pos('-', fELEVATION.AsString) > 0 then
        valtitude := StrToIntDef(fLATITUDE.AsString.Split('-')[0], 0)
      else
        valtitude := StrToIntDef(fELEVATION.AsString, 0);
      latitude[i] := StrToFloatDef(fLATITUDE.AsString, 0.0);
      longitude[i] := StrToFloatDef(fLONGITUDE.AsString, 0.0);
      altitude[i] := valtitude;
      Inc(i);
      Dataset1.Next;
    end;
    minlat := MinValue(latitude);
    maxlat := MaxValue(latitude);
    minlon := MinValue(longitude);
    maxlon := MaxValue(longitude);
    minalt := MinIntValue(altitude);
    maxalt := MaxIntValue(altitude);
    vunit := GetUnit(fELEVATION.FieldName);
    Inc(levels);

    { Level 2: Discovery }
    emlBody :=
      '<coverage>' + LineEnding + '<geographicCoverage>' + LineEnding +
      '<geographicDescription>' + ''.Join(',',
      [vcountry, vstate, vprovince, vlocality]) + '</geographicDescription>' +
      LineEnding + '<boundingCoordinates>' + LineEnding +
      '<westBoundingCoordinate>' + FloatToStr(maxlon) + '</westBoundingCoordinate>' +
      LineEnding + '<eastBoundingCoordinate>' + FloatToStr(minlon) +
      '</eastBoundingCoordinate>' + LineEnding + '<northBoundingCoordinate>' +
      FloatToStr(maxlat) + '</northBoundingCoordinate>' + LineEnding +
      '<southBoundingCoordinate>' + FloatToStr(minlat) +
      '</southBoundingCoordinate>' + LineEnding + '<boundingAltitudes>' +
      LineEnding + '<altitudeMinimum>' + IntToStr(minalt) +
      '</altitudeMinimum>' + LineEnding + '<altitudeMaximum>' +
      IntToStr(maxalt) + '</altitudeMaximum>' + LineEnding +
      '<altitudeUnits>' + vunit + '</altitudeUnits>' + LineEnding +
      '</boundingAltitudes>' + LineEnding + '</boundingCoordinates>' +
      LineEnding + '</geographicCoverage>';
    xmldoc += emlBody;

    fSAMPLE := Dataset1.Fields[0];
    fDATE := Dataset1.Fields[6];
    lSAMPLE := TStringList.Create;
    lSAMPLE.Sorted := True;
    lSAMPLE.Duplicates := dupIgnore;
    lDATE := TStringList.Create;
    lDATE.Sorted := True;
    lDATE.Duplicates := dupIgnore;
    Dataset1.First;
    while not Dataset1.EOF do
    begin
      lSAMPLE.Add(fSAMPLE.AsString);
      lDATE.Add(fDATE.AsString);
      Dataset1.Next;
    end;
    date1 := lDATE[0];
    date2 := lDATE[lDATE.Count - 1];
    lDATE.Free;
    Dataset1.EnableControls;
    Dataset1.GotoBookmark(bm);
  end;

  emlBody :=
    '<temporalCoverage>' + LineEnding + '<rangeOfDates>' + LineEnding +
    '<beginDate>' + LineEnding + '<calendarDate>' + date1 +
    '</calendarDate>' + LineEnding + '</beginDate>' + LineEnding +
    '<endDate>' + LineEnding + '<calendarDate>' + date2 + '</calendarDate>' +
    LineEnding + '</endDate>' + LineEnding + '</rangeOfDates>' +
    LineEnding + '</temporalCoverage>';
  xmldoc += emlBody;

  emlBody :=
    '<taxonomicCoverage>' + LineEnding + '<generalTaxonomicCoverage>' +
    vdescription + '</generalTaxonomicCoverage>';
  xmldoc += emlBody;

  grpCheck := '';
  grpEval := '';
  subCheck := '';
  subEval := '';

  with MainForm.MultiDoc.ActiveObject as TMDIChild do
  begin
    QueryDataset1.First;
    while not QueryDataset1.EOF do
    begin
      grpCheck := Trim(QueryDataset1.Fields[0].AsString);
      if grpEval <> grpCheck then
      begin
        grpEval := grpCheck;
        emlBody :=
          '<taxonomicClassification>' + LineEnding +
          '<taxonRankName>Family</taxonRankName>' + LineEnding +
          '<taxonRankValue>' + grpCheck + '</taxonRankValue>' +
          LineEnding + '</taxonomicClassification>';
        xmldoc += emlBody;
      end;

      subCheck := Trim(QueryDataset1.Fields[1].AsString.Split(' ')[0]);
      if subEval <> subCheck then
      begin
        subEval := subCheck;
        emlBody :=
          '<taxonomicClassification>' + LineEnding +
          '<taxonRankName>Genus</taxonRankName>' + LineEnding +
          '<taxonRankValue>' + subCheck + '</taxonRankValue>' +
          LineEnding + '</taxonomicClassification>';
        xmldoc += emlBody;
      end;

      emlBody :=
        '<taxonomicClassification>' + LineEnding +
        '<taxonRankName>Species</taxonRankName>' + LineEnding +
        '<taxonRankValue>' + Trim(QueryDataset1.Fields[1].AsString) +
        '</taxonRankValue>' + LineEnding + '</taxonomicClassification>';
      xmldoc += emlBody;

      QueryDataset1.Next;
    end;
    QueryConnection1.Disconnect;
  end;

  emlBody :=
    '</taxonomicCoverage>' + LineEnding + '</coverage>';
  xmldoc += emlBody;

  emlBody :=
    '<maintenance>' + LineEnding + '<description>' + LineEnding +
    '<para>Last access: ' + DateToStr(Now) + '</para>' + LineEnding +
    '<para>Last modified: ' + DateToStr(Now) + '</para>' + LineEnding +
    '</description>' + LineEnding + '</maintenance>' + LineEnding +
    '<contact>' + LineEnding + '<references>' + vsurname.toLower +
    ',' + vname.toLower + '</references>' + LineEnding + '</contact>' +
    LineEnding + '<publisher>' + LineEnding + '<references>' +
    vsurname.toLower + ',' + vname.toLower + '</references>' +
    LineEnding + '</publisher>';
  xmldoc += emlBody;
  Inc(levels);

  { Level 3: Evaluation }
  vmethod := Metadata.Find('method').AsString;

  emlBody :=
    '<methods>' + LineEnding + '<methodStep>' + LineEnding + '<description>' +
    LineEnding + '<para>' + LineEnding + '<literalLayout>' + vmethod +
    '</literalLayout></para>' + LineEnding + '</description>' +
    LineEnding + '</methodStep>' + LineEnding + '<sampling>' +
    LineEnding + '<studyExtent>' + LineEnding + '<coverage>' +
    LineEnding + '<temporalCoverage>' + LineEnding + '<rangeOfDates>' +
    LineEnding + '<beginDate>' + LineEnding + '<calendarDate>' +
    date1 + '</calendarDate>' + LineEnding + '</beginDate>' +
    LineEnding + '<endDate>' + LineEnding + '<calendarDate>' + date2 +
    '</calendarDate>' + LineEnding + '</endDate>' + LineEnding +
    '</rangeOfDates>' + LineEnding + '</temporalCoverage>' + LineEnding +
    '<geographicCoverage>' + LineEnding + '<geographicDescription>' +
    ''.Join(',', [vcountry, vstate, vprovince, vlocality]) +
    '</geographicDescription>' + LineEnding + '<boundingCoordinates>' +
    LineEnding + '<westBoundingCoordinate>' + FloatToStr(maxlon) +
    '</westBoundingCoordinate>' + LineEnding + '<eastBoundingCoordinate>' +
    FloatToStr(minlon) + '</eastBoundingCoordinate>' + LineEnding +
    '<northBoundingCoordinate>' + FloatToStr(maxlat) + '</northBoundingCoordinate>' +
    LineEnding + '<southBoundingCoordinate>' + FloatToStr(minlat) +
    '</southBoundingCoordinate>' + LineEnding + '<boundingAltitudes>' +
    LineEnding + '<altitudeMinimum>' + IntToStr(minalt) + '</altitudeMinimum>' +
    LineEnding + '<altitudeMaximum>' + IntToStr(maxalt) + '</altitudeMaximum>' +
    LineEnding + '<altitudeUnits>' + vunit + '</altitudeUnits>' +
    LineEnding + '</boundingAltitudes>' + LineEnding + '</boundingCoordinates>' +
    LineEnding + '</geographicCoverage>' + LineEnding + '</coverage>' +
    LineEnding + '</studyExtent>' + LineEnding + '<samplingDescription>' +
    LineEnding + '<para>' + vmethod + '</para>' + LineEnding +
    '</samplingDescription>' + LineEnding + '<spatialSamplingUnits>';
  xmldoc += emlBody;

  for i := 1 to lSAMPLE.Count do
  begin
    vsample := lSAMPLE[i - 1];
    emlBody :=
      '<coverage>' + LineEnding + '<geographicDescription>' +
      vsample + '</geographicDescription>' + LineEnding +
      '<boundingCoordinates>' + LineEnding + '<westBoundingCoordinate>' +
      FloatToStr(maxlon) + '</westBoundingCoordinate>' + LineEnding +
      '<eastBoundingCoordinate>' + FloatToStr(minlon) +
      '</eastBoundingCoordinate>' + LineEnding + '<northBoundingCoordinate>' +
      FloatToStr(maxlat) + '</northBoundingCoordinate>' + LineEnding +
      '<southBoundingCoordinate>' + FloatToStr(minlat) +
      '</southBoundingCoordinate>' + LineEnding + '</boundingCoordinates>' +
      LineEnding + '</coverage>';
    xmldoc += emlBody;
  end;

  emlBody :=
    '</spatialSamplingUnits>' + LineEnding + '</sampling>' + LineEnding + '</methods>';
  xmldoc += emlBody;

  emlBody :=
    '<project>' + LineEnding + '<title>' + vtitle + '</title>' +
    LineEnding + '<personnel>' + LineEnding + '<references>' +
    vsurname.toLower + ',' + vname.toLower + '</references>' + LineEnding +
    '<role>' + vrole + '</role>' + LineEnding + '</personnel>' +
    LineEnding + '<abstract>' + LineEnding + '<para>' + vdescription +
    '</para>' + LineEnding + '</abstract>' + LineEnding + '<funding>' +
    LineEnding + '<para>' + vfunding + '</para>' + LineEnding +
    '</funding>' + LineEnding + '</project>';
  xmldoc += emlBody;
  Inc(levels);

  emlFooter := '</dataset>' + LineEnding + '</eml:eml>';
  xmldoc += emlFooter;
  S := TStringStream.Create(xmldoc);
  try
    ReadXMLFile(XML, S);
    WriteXMLFile(XML, filename);
  finally
    S.Free;
  end;

  if FileExists('temp.csv') then
    DeleteFile('temp.csv');
  Result := levels;
end;

procedure toFitopac1(const filename: string; kind: integer;
  var nfam, nspp, nind: integer);
type
  TIntVector = array of integer;
var
  fname1, fname2, DatasetName1: string;
  lFAMILY, lSPECIES, lSAMPLE: TStringList;
  fld, fFAMILY, fSPECIES, fSAMPLE, fSPECIMEN: TField;
  SAMPLE, SPECIMEN, FAMILY, SPECIES: string;
  bm: TBookmark;
  nf: TIntVector = nil;
  vfam, vloc, aux: string;
  idxFAMILY: integer;
  i, j, k, f, s: integer;
  outfile: TextFile;
begin
  DefaultFormatSettings.DecimalSeparator := '.';
  with MainForm.MultiDoc.ActiveObject as TMDIChild do
  begin
    if not Dataset1.Active then
    begin
      DatasetName1 := Metadata.Find('dataset1').AsString;
      Dataset1.FileName := DatasetName1;
      Dataset1.AutoFieldDefs := True;
      Dataset1.Open;
    end;

    CSVExporter1.FileName := 'temp.csv';
    CSVExporter1.Execute;
    QueryConnection1.DatabasePath := GetCurrentDir;
    QueryConnection1.Connect;
    QueryDataSet1.TableName := 'temp';
    QueryDataset1.LoadFromTable;
    SAMPLE := QueryDataset1.Fields[0].FieldName;
    SPECIMEN := QueryDataset1.Fields[1].FieldName;
    FAMILY := QueryDataset1.Fields[2].FieldName;
    SPECIES := QueryDataset1.Fields[3].FieldName;
    QueryDataset1.SQL.Text :=
      'SELECT ' + FAMILY + ', ' + SPECIES + ' FROM temp GROUP BY ' +
      FAMILY + ', ' + SPECIES + ' ORDER BY ' + SPECIES + ', ' + FAMILY;
    QueryDataset1.QueryExecute;
    QueryDataset1.SaveToTable;

    bm := Dataset1.GetBookmark;
    Dataset1.DisableControls;
    lFAMILY := TStringList.Create;
    lSPECIES := TStringList.Create;
    lSAMPLE := TStringList.Create;
    try
      lFAMILY.Sorted := True;
      lFAMILY.Duplicates := dupIgnore;
      lSPECIES.Sorted := True;
      lSPECIES.Duplicates := dupIgnore;
      lSAMPLE.Sorted := True;
      lSAMPLE.Duplicates := dupIgnore;
      fSAMPLE := Dataset1.Fields[0];
      fSPECIMEN := Dataset1.Fields[1];
      fFAMILY := Dataset1.Fields[2];
      fSPECIES := Dataset1.Fields[3];

      Dataset1.First;
      while not Dataset1.EOF do
      begin
        lFAMILY.Add(fFAMILY.AsString);
        lSPECIES.Add(fSPECIES.AsString);
        lSAMPLE.Add(fSAMPLE.AsString);
        Dataset1.Next;
      end;

      nfam := lFAMILY.Count;
      nspp := lSPECIES.Count;

      j := 0;
      SetLength(nf, nspp);
      QueryDataset1.First;
      while not QueryDataset1.EOF do
      begin
        vfam := Trim(QueryDataset1.Fields[0].AsString);
        lFAMILY.Find(vfam, idxFAMILY);
        nf[j] := idxFAMILY + 1;
        QueryDataset1.Next;
        Inc(j);
      end;
      QueryConnection1.Disconnect;

      fname1 := GetFileNameWithoutExt(filename) + '.nms';
      AssignFile(outfile, fname1);
      Rewrite(outfile);
      for f := 1 to nfam do
        WriteLn(outfile, f, ' ', lFAMILY[f - 1]);
      WriteLn(outfile, '999');
      for s := 1 to nspp do
        WriteLn(outfile, s, ' ', nf[s - 1], ' ', lSPECIES[s - 1]);
      CloseFile(outfile);

      fname2 := GetFileNameWithoutExt(filename) + '.dad';
      AssignFile(outfile, fname2);
      Rewrite(outfile);
      i := 0;
      k := 0;
      nind := 0;
      aux := '';
      Dataset1.SortOnFields(SAMPLE);
      Dataset1.First;
      while not Dataset1.EOF do
      begin
        if kind = 2 then
          Inc(i);
        if i >= 4 then
        begin
          i := 0;
          Inc(k);
        end;
        if kind = 1 then
        begin
          vloc := Dataset1.Fields[0].AsString;
          if aux <> vloc then
          begin
            aux := vloc;
            WriteLn(outfile, DelSpace(vloc));
          end;
          Write(outfile, nind + 1, ' ');
        end;
        if kind = 2 then
          Write(outfile, k + 1, ' ', nind + 1, ' ');
        for fld in Dataset1.Fields do
        begin
          if fld is TFloatField then
            if (AnsiContainsText(fld.FieldName, 'LAT')) or
              (AnsiContainsText(fld.FieldName, 'LON')) then
              continue
            else
              Write(outfile, fld.AsString, ' ');
        end;
        WriteLn(outfile, lSPECIES.IndexOf(fSPECIES.AsString) + 1);
        Dataset1.Next;
        Inc(nind);
      end;

      CloseFile(outfile);
      if FileExists('temp.csv') then
        DeleteFile('temp.csv');

    finally
      lFAMILY.Free;
      lSPECIES.Free;
      lSAMPLE.Free;
      Dataset1.EnableControls;
      Dataset1.GotoBookmark(bm);
    end;
  end;
end;

procedure toFitopac2(const filename: string; var nfam, nspp, nind, nsam: integer);
const
  VERSION = '2.1';
type
  TIntVector = array of integer;
var
  DatasetName1: string;
  lFAMILY, lSPECIES, lSAMPLE: TStringList;
  fld, fFAMILY, fSPECIES, fSAMPLE, fSPECIMEN: TField;
  FAMILY, SPECIES: string;
  bm: TBookmark;
  sums: TIntVector = nil;
  ni: TIntVector = nil;
  size, vfam: string;
  latitude, longitude, len, wid: double;
  idxFAMILY, idxSAMPLE: integer;
  i, j, f, s, n, vlatdeg, vlatmin, vlatsec, vlondeg, vlonmin, vlonsec: integer;
  outfile: TextFile;
begin
  DefaultFormatSettings.DecimalSeparator := '.';
  with MainForm.MultiDoc.ActiveObject as TMDIChild do
  begin
    if not Dataset1.Active then
    begin
      DatasetName1 := Metadata.Find('dataset1').AsString;
      Dataset1.FileName := DatasetName1;
      Dataset1.AutoFieldDefs := True;
      Dataset1.Open;
    end;

    CSVExporter1.FileName := 'temp.csv';
    CSVExporter1.Execute;
    QueryConnection1.DatabasePath := GetCurrentDir;
    QueryConnection1.Connect;
    QueryDataSet1.TableName := 'temp';
    QueryDataset1.LoadFromTable;
    FAMILY := QueryDataset1.Fields[2].FieldName;
    SPECIES := QueryDataset1.Fields[3].FieldName;
    QueryDataset1.SQL.Text :=
      'SELECT ' + FAMILY + ', ' + SPECIES + ' FROM temp GROUP BY ' +
      FAMILY + ', ' + SPECIES + ' ORDER BY ' + SPECIES + ', ' + FAMILY;
    QueryDataset1.QueryExecute;
    QueryDataset1.SaveToTable;

    bm := Dataset1.GetBookmark;
    Dataset1.DisableControls;
    lFAMILY := TStringList.Create;
    lSPECIES := TStringList.Create;
    lSAMPLE := TStringList.Create;
    try
      lFAMILY.Sorted := True;
      lFAMILY.Duplicates := dupIgnore;
      lSPECIES.Sorted := True;
      lSPECIES.Duplicates := dupIgnore;
      lSAMPLE.Sorted := True;
      lSAMPLE.Duplicates := dupIgnore;
      fSAMPLE := Dataset1.Fields[0];
      fSPECIMEN := Dataset1.Fields[1];
      fFAMILY := Dataset1.Fields[2];
      fSPECIES := Dataset1.Fields[3];

      Dataset1.First;
      while not Dataset1.EOF do
      begin
        lFAMILY.Add(fFAMILY.AsString);
        lSPECIES.Add(fSPECIES.AsString);
        lSAMPLE.Add(fSAMPLE.AsString);
        Dataset1.Next;
      end;

      nsam := lSAMPLE.Count;
      nfam := lFAMILY.Count;
      nspp := lSPECIES.Count;
      SetLength(sums, nsam);

      Dataset1.First;
      while not Dataset1.EOF do
      begin
        lSAMPLE.Find(fSAMPLE.AsString, idxSAMPLE);
        Inc(sums[idxSAMPLE]);
        Dataset1.Next;
      end;

      nind := 0;
      for i := 1 to nsam do
        nind += sums[i - 1];

      j := 0;
      SetLength(ni, nspp);
      QueryDataset1.First;
      while not QueryDataset1.EOF do
      begin
        vfam := Trim(QueryDataset1.Fields[0].AsString);
        lFAMILY.Find(vfam, idxFAMILY);
        ni[j] := idxFAMILY + 1;
        QueryDataset1.Next;
        Inc(j);
      end;
      QueryConnection1.Disconnect;

      AssignFile(outfile, filename);
      Rewrite(outfile);
      WriteLn(outfile, 'FPD ', VERSION);
      WriteLn(outfile, Metadata.Find('title').AsString);
      WriteLn(outfile, Metadata.Find('author').AsString);
      WriteLn(outfile, StringReplace(DateToStr(Now), '/', ' ', [rfReplaceAll]),
        ' ', StringReplace(TimeToStr(Now), ':', ' ', [rfReplaceAll]));
      WriteLn(outfile, Metadata.Find('state').AsString);
      WriteLn(outfile, Metadata.Find('province').AsString);
      WriteLn(outfile, Metadata.Find('locality').AsString);
      latitude := StrToFloat(Metadata.Find('latitude').AsString);
      longitude := StrToFloat(Metadata.Find('longitude').AsString);
      if latitude < 0.0 then
        Write(outfile, 'S')
      else
        Write(outfile, 'N');
      if longitude < 0.0 then
        Write(outfile, 'W ')
      else
        Write(outfile, 'E ');
      DegToDMS(latitude, vlatdeg, vlatmin, vlatsec);
      DegToDMS(longitude, vlondeg, vlonmin, vlonsec);
      WriteLn(outfile, abs(vlatdeg), ' ', abs(vlatmin), ' ', abs(vlatsec), ' ',
        abs(vlondeg), ' ', abs(vlonmin), ' ', abs(vlonsec));
      WriteLn(outfile, StringReplace(Metadata.Find('elevation').AsString,
        '-', ' ', [rfReplaceAll]));
      WriteLn(outfile, 'P');
      size := LowerCase(DelSpace(Metadata.Find('size').AsString));
      len := StrToFloatDef(size.Split('x')[0], 0.0);
      wid := StrToFloatDef(TrimSet(size.split('x')[1], ['a'..'z']), 0.0);
      WriteLn(outfile, nfam, ' ', nspp, ' ', nind, ' ', nsam, ' ',
        len: 5: 2, ' ', wid: 5: 2);
      WriteLn(outfile, 'TRUE TRUE');
      WriteLn(outfile, 'TRUE TRUE');
      WriteLn(outfile, 'TRUE DC');
      WriteLn(outfile);
      WriteLn(outfile);
      WriteLn(outfile);

      for f := 1 to nfam do
        WriteLn(outfile, 'T ', lFAMILY[f - 1]);

      for s := 1 to nspp do
        WriteLn(outfile, 'T ', ni[s - 1], ' ', LeftStr(lSPECIES[s - 1], 70));

      for n := 1 to nsam do
        WriteLn(outfile, 'T ', sums[n - 1], ' ', LeftStr(lSAMPLE[n - 1], 70));

      Dataset1.First;
      while not Dataset1.EOF do
      begin
        Write(outfile, fSPECIMEN.AsString, ' ',
          lSPECIES.IndexOf(fSPECIES.AsString) + 1, ' ');
        for fld in Dataset1.Fields do
        begin
          if fld is TFloatField then
            if (AnsiContainsText(fld.FieldName, 'LAT')) or
              (AnsiContainsText(fld.FieldName, 'LON')) then
              continue
            else
              Write(outfile, fld.AsString, ' ');
        end;
        WriteLn(outfile);
        Dataset1.Next;
      end;

      CloseFile(outfile);
      if FileExists('temp.csv') then
        DeleteFile('temp.csv');

    finally
      lFAMILY.Free;
      lSPECIES.Free;
      lSAMPLE.Free;
      Dataset1.EnableControls;
      Dataset1.GotoBookmark(bm);
    end;
  end;
end;

function toKML(const filename: string): integer;
var
  S: TStringStream;
  XML: TXMLDocument;
  bm: TBookmark;
  DatasetName1: string;
  kmlHeader, kmlBody, kmlFooter, kmlOutput, kml, vlocal, vdesc: string;
  slatitude, slongitude, saltitude: string;
  vlatitude, vlongitude: double;
  valtitude, counter: integer;
  {outfile: TextFile;}
begin
  kmlHeader :=
    '<?xml version="1.0" encoding="UTF-8"?>' + LineEnding +
    '<kml xmlns="http://www.opengis.net/kml/2.2">';
  kmlBody :=
    '<Folder>' + LineEnding + '<Style id="normalPlaceMarker">' +
    LineEnding + '<IconStyle>' + LineEnding + '<Icon>' + LineEnding +
    '<href>http://maps.google.com/mapfiles/kml/pal3/icon38.png</href>' +
    LineEnding + '</Icon>' + LineEnding + '</IconStyle>' + LineEnding + '</Style>';
  kmlFooter := '</Folder>' + LineEnding + '</kml>';

  counter := 0;
  with MainForm.MultiDoc.ActiveObject as TMDIChild do
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
    Dataset1.First;
    while not Dataset1.EOF do
    begin
      vlocal := Dataset1.Fields[0].AsString;
      vdesc := Dataset1.Fields[8].AsString;
      vlatitude := StrToFloatDef(Dataset1.Fields[9].AsString, 0.0000);
      vlongitude := StrToFloatDef(Dataset1.Fields[10].AsString, 0.0000);
      valtitude := StrToIntDef(Dataset1.Fields[11].AsString, 0);
      slatitude := StringReplace(FloatToStr(vlatitude), ',', '.', [rfReplaceAll]);
      slongitude := StringReplace(FloatToStr(vlongitude), ',', '.', [rfReplaceAll]);
      saltitude := IntToStr(valtitude);
      kml := '<Placemark>' + LineEnding + '<name>' + vlocal +
        '</name>' + LineEnding + '<description>' + vdesc + '</description>' +
        LineEnding + '<Point>' + LineEnding + '<coordinates>' +
        slongitude + ',' + slatitude + ',' + saltitude + '</coordinates>' +
        LineEnding + '</Point>' + LineEnding + '</Placemark>' + LineEnding;
      kmlBody += kml;
      Inc(counter);
      Dataset1.Next;
    end;
    Dataset1.EnableControls;
    Dataset1.GotoBookmark(bm);
  end;

  kmlOutput := kmlHeader + kmlBody + kmlFooter;
  S := TStringStream.Create(kmlOutput);
  try
    ReadXMLFile(XML, S);
    WriteXMLFile(XML, filename);
  finally
    S.Free;
  end;
  Result := counter;
end;

{$IFDEF WINDOWS}
function toSHP(const filename: string): integer;
var
  bm: TBookmark;
  DatasetName1: string;
  hSHPHandle: SHPHandle;
  hDBFHandle: DBFHandle;
  psShape: PSHPObject;
  vlocal: string;
  x, y, vlatitude, vlongitude: double;
  valtitude, counter: integer;
begin
  hSHPHandle := SHPCreate(PChar(filename), SHPT_POINT);
  hDBFHandle := DBFCreate(PChar(filename));
  DBFAddField(hDBFHandle, 'LOCAL', FTString, 65, 0);
  DBFAddField(hDBFHandle, 'LONGITUDE', FTDouble, 11, 6);
  DBFAddField(hDBFHandle, 'LATITUDE', FTDouble, 11, 6);
  DBFAddField(hDBFHandle, 'ELEVATION', FTInteger, 7, 0);

  counter := 0;
  with MainForm.MultiDoc.ActiveObject as TMDIChild do
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
    Dataset1.First;
    while not Dataset1.EOF do
    begin
      vlocal := Dataset1.Fields[0].AsString;
      vlatitude := StrToFloatDef(Dataset1.Fields[9].AsString, 0.0000);
      vlongitude := StrToFloatDef(Dataset1.Fields[10].AsString, 0.0000);
      valtitude := StrToIntDef(Dataset1.Fields[11].AsString, 0);
      x := vlongitude;
      y := vlatitude;
      psShape := SHPCreateObject(SHPT_POINT, -1, 0, nil, nil, 1, @x, @y, nil, nil);
      SHPWriteObject(hSHPHandle, -1, psShape);
      DBFWriteStringAttribute(hDBFHandle, counter, 0, PChar(vlocal));
      DBFWriteDoubleAttribute(hDBFHandle, counter, 1, vlongitude);
      DBFWriteDoubleAttribute(hDBFHandle, counter, 2, vlatitude);
      DBFWriteIntegerAttribute(hDBFHandle, counter, 3, valtitude);
      SHPDestroyObject(psShape);
      Inc(counter);
      Dataset1.Next;
    end;
    Dataset1.EnableControls;
    Dataset1.GotoBookmark(bm);
  end;

  SHPClose(hSHPHandle);
  DBFClose(hDBFHandle);
  Result := counter;
end;
{$ENDIF}

function toRDE(const filename: string): integer;
const
  TAG = 0;
  DEL = 1;
  BOTRECCAT = 2;
  RDEIMAGES = 3;
  DUPS = 4;
  BARCODE = 5;
  ACCESSION = 6;
  COLLECTOR = 7;
  ADDCOLL = 8;
  PREFIX = 9;
  NUMBER = 10;
  SUFFIX = 11;
  COLLDD = 12;
  COLLMM = 13;
  COLLYY = 14;
  DATERES = 15;
  DATETEXT = 16;
  FAMILY = 17;
  GENUS = 18;
  SP1 = 19;
  AUTHOR1 = 20;
  RANK1 = 21;
  SP2 = 22;
  AUTHOR2 = 23;
  RANK2 = 24;
  SP3 = 25;
  AUTHOR3 = 26;
  UNIQUE = 27;
  PLANTDESC = 28;
  PHENOLOGY = 29;
  DETBY = 30;
  DETDD = 31;
  DETMM = 32;
  DETYY = 33;
  DETSTATUS = 34;
  DETNOTES = 35;
  COUNTRY = 36;
  MAJORAREA = 37;
  MINORAREA = 38;
  GAZETTEER = 39;
  LOCNOTES = 40;
  HABITATTXT = 41;
  CULTIVATED = 42;
  CULTNOTES = 43;
  ORIGINSTAT = 44;
  ORIGINID = 45;
  ORIGINDB = 46;
  NOTES = 47;
  LAT = 48;
  NS = 49;
  LONG = 50;
  EW = 51;
  LLGAZ = 52;
  LLUNIT = 53;
  LLRES = 54;
  LLORIG = 55;
  LLDATUM = 56;
  QDS = 57;
  ALT = 58;
  ALTMAX = 59;
  ALTUNIT = 60;
  ALTRES = 61;
  ALTTEXT = 62;
  INITIAL = 63;
  AVAILABLE = 64;
  CURATENOTE = 65;
  NOTONLINE = 66;
  LABELTOTAL = 67;
  VERNACULAR = 68;
  LANGUAGE = 69;
  GEODATA = 70;
  LATLONG = 71;
  COLLECTED = 72;
  MONTHNAME = 73;
  DETBYDATE = 74;
  CHECKWHO = 75;
  CHECKDATE = 76;
  CHECKNOTE = 77;
  UNIQUEID = 78;
  RDESPEC = 79;
var
  dbfile: TDbf;
  bm: TBookmark;
  DatasetName1: string;
  vcountry, vmajorarea, vminorarea, vgazetteer: string;
  vfamily, vgenus, cf, vspecies, vsp1, vsp2, vauthor1, vinfraname, vauthor2: string;
  vcollector, vnumber, vdate, vlocnotes, vnotes, vns, vew, valtunit: string;
  vlat, vlon: double;
  reccount, vcolldd, vcollmm, vcollyy, valt, valtmax: integer;
  fFAMILY, fSPECIES, fCOLLECTOR, fNUMBER, fDATE, fOBS, fLOCALITY,
  fLATITUDE, fLONGITUDE, fELEVATION: TField;
begin
  dbfile := TDbf.Create(nil);
  dbfile.TableName := filename;
  dbfile.TableLevel := 4; { dBase IV }
  dbfile.Exclusive := True;

  with dbfile.FieldDefs do
  begin
    Add('TAG', DB.ftString, 1, True);
    Add('DEL', DB.ftString, 1, True);
    Add('BOTRECCAT', DB.ftString, 1, True);
    Add('RDEIMAGES', DB.ftString, 128, True);
    Add('DUPS', DB.ftString, 40, True);
    Add('BARCODE', DB.ftString, 15, True);
    Add('ACCESSION', DB.ftString, 15, True);
    Add('COLLECTOR', DB.ftString, 80, True);
    Add('ADDCOLL', DB.ftString, 80, True);
    Add('PREFIX', DB.ftString, 10, True);
    Add('NUMBER', DB.ftString, 15, True);
    Add('SUFFIX', DB.ftString, 6, True);
    Add('COLLDD', DB.ftInteger, 2, True);
    Add('COLLMM', DB.ftInteger, 2, True);
    Add('COLLYY', DB.ftInteger, 4, True);
    Add('DATERES', DB.ftString, 5, True);
    Add('DATETEXT', DB.ftString, 128, True);
    Add('FAMILY', DB.ftString, 30, True);
    Add('GENUS', DB.ftString, 25, True);
    Add('SP1', DB.ftString, 25, True);
    Add('AUTHOR1', DB.ftString, 75, True);
    Add('RANK1', DB.ftString, 10, True);
    Add('SP2', DB.ftString, 25, True);
    Add('AUTHOR2', DB.ftString, 75, True);
    Add('RANK2', DB.ftString, 10, True);
    Add('SP3', DB.ftString, 25, True);
    Add('AUTHOR3', DB.ftString, 75, True);
    Add('UNIQUE', DB.ftInteger, 1, True);
    Add('PLANTDESC', DB.ftString, 128, True);
    Add('PHENOLOGY', DB.ftString, 10, True);
    Add('DETBY', DB.ftString, 30, True);
    Add('DETDD', DB.ftInteger, 2, True);
    Add('DETMM', DB.ftInteger, 2, True);
    Add('DETYY', DB.ftInteger, 4, True);
    Add('DETSTATUS', DB.ftString, 5, True);
    Add('DETNOTES', DB.ftString, 128, True);
    Add('COUNTRY', DB.ftString, 50, True);
    Add('MAJORAREA', DB.ftString, 30, True);
    Add('MINORAREA', DB.ftString, 30, True);
    Add('GAZETTEER', DB.ftString, 50, True);
    Add('LOCNOTES', DB.ftString, 128, True);
    Add('HABITATTXT', DB.ftString, 128, True);
    Add('CULTIVATED', DB.ftBoolean, 1, True);
    Add('CULTNOTES', DB.ftString, 128, True);
    Add('ORIGINSTAT', DB.ftString, 5, True);
    Add('ORIGINID', DB.ftString, 15, True);
    Add('ORIGINDB', DB.ftString, 10, True);
    Add('NOTES', DB.ftString, 128, True);
    Add('LAT', DB.ftFloat, 16, True);
    Add('NS', DB.ftString, 1, True);
    Add('LONG', DB.ftFloat, 16, True);
    Add('EW', DB.ftString, 1, True);
    Add('LLGAZ', DB.ftString, 5, True);
    Add('LLUNIT', DB.ftString, 5, True);
    Add('LLRES', DB.ftString, 5, True);
    Add('LLORIG', DB.ftString, 20, True);
    Add('LLDATUM', DB.ftString, 10, True);
    Add('QDS', DB.ftString, 10, True);
    Add('ALT', DB.ftString, 8, True);
    Add('ALTMAX', DB.ftString, 8, True);
    Add('ALTUNIT', DB.ftString, 1, True);
    Add('ALTRES', DB.ftString, 5, True);
    Add('ALTTEXT', DB.ftString, 128, True);
    Add('INITIAL', DB.ftInteger, 2, True);
    Add('AVAILABLE', DB.ftInteger, 2, True);
    Add('CURATENOTE', DB.ftString, 128, True);
    Add('NOTONLINE', DB.ftString, 1, True);
    Add('LABELTOTAL', DB.ftInteger, 2, True);
    Add('VERNACULAR', DB.ftString, 40, True);
    Add('LANGUAGE', DB.ftString, 40, True);
    Add('GEODATA', DB.ftString, 128, True);
    Add('LATLONG', DB.ftString, 50, True);
    Add('COLLECTED', DB.ftString, 20, True);
    Add('MONTHNAME', DB.ftString, 10, True);
    Add('DETBYDATE', DB.ftString, 20, True);
    Add('CHECKWHO', DB.ftString, 5, True);
    Add('CHECKDATE', DB.ftDate, 8, True);
    Add('CHECKNOTE', DB.ftString, 128, True);
    Add('UNIQUEID', DB.ftString, 15, True);
    Add('RDESPEC', DB.ftString, 128);
  end;

  dbfile.CreateTable;
  dbfile.Open;

  vcountry := Metadata.Find('country').AsString;
  vmajorarea := Metadata.Find('state').AsString;
  vminorarea := Metadata.Find('province').AsString;
  vgazetteer := Metadata.Find('locality').AsString;

  with MainForm.MultiDoc.ActiveObject as TMDIChild do
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

    fFAMILY := Dataset1.Fields[2];
    fSPECIES := Dataset1.Fields[3];
    fCOLLECTOR := Dataset1.Fields[4];
    fNUMBER := Dataset1.Fields[5];
    fDATE := Dataset1.Fields[6];
    fOBS := Dataset1.Fields[7];
    fLOCALITY := Dataset1.Fields[8];
    fLATITUDE := Dataset1.Fields[9];
    fLONGITUDE := Dataset1.Fields[10];
    fELEVATION := Dataset1.Fields[11];

    valtunit := GetUnit(Dataset1.Fields[11].FieldName);
    reccount := 0;
    Dataset1.First;
    while not Dataset1.EOF do
    begin
      vcollector := fCOLLECTOR.AsString;
      vnumber := fNUMBER.AsString;
      vdate := IfThen(fDATE.AsString = '', '00-00-0000', fDATE.AsString);
      vcolldd := StrToInt(vdate.split('-')[0]);
      vcollmm := StrToInt(vdate.split('-')[1]);
      vcollyy := StrToInt(vdate.split('-')[2]);
      vfamily := fFAMILY.AsString;
      vspecies := fSPECIES.AsString;
      ParseName(vspecies, vgenus, cf, vsp1, vauthor1, vsp2, vinfraname, vauthor2);
      vlocnotes := fLOCALITY.AsString;
      vnotes := fOBS.AsString;
      vlat := StrToFloatDef(fLATITUDE.AsString, 0.0);
      vns := IfThen(vlat < 0, 'S', 'N');
      vlon := StrToFloatDef(fLONGITUDE.AsString, 0.0);
      vew := IfThen(vlon < 0, 'E', 'W');
      if Pos('-', fELEVATION.AsString) > 0 then
      begin
        valt := StrToIntDef(fELEVATION.AsString.Split('-')[0], 0);
        valtmax := StrToIntDef(fELEVATION.AsString.Split('-')[1], 0);
      end
      else
      begin
        valt := StrToIntDef(fELEVATION.AsString, 0);
        valtmax := 0;
      end;

      dbfile.Append;
      with dbfile do
      begin
        Fields[TAG].AsString := '';
        Fields[DEL].AsString := '';
        Fields[BOTRECCAT].AsString := '';
        Fields[RDEIMAGES].AsString := '';
        Fields[DUPS].AsString := '';
        Fields[BARCODE].AsString := '';
        Fields[ACCESSION].AsString := '';
        Fields[COLLECTOR].AsString := vcollector;
        Fields[ADDCOLL].AsString := '';
        Fields[PREFIX].AsString := '';
        Fields[NUMBER].AsString := vnumber;
        Fields[SUFFIX].AsString := '';
        Fields[COLLDD].AsInteger := vcolldd;
        Fields[COLLMM].AsInteger := vcollmm;
        Fields[COLLYY].AsInteger := vcollyy;
        Fields[DATERES].AsString := '';
        Fields[DATETEXT].AsString := '';
        Fields[FAMILY].AsString := vfamily;
        Fields[GENUS].AsString := vgenus;
        Fields[SP1].AsString := vsp1;
        Fields[AUTHOR1].AsString := vauthor1;
        Fields[RANK1].AsString := vsp2;
        Fields[SP2].AsString := vinfraname;
        Fields[AUTHOR2].AsString := vauthor2;
        Fields[RANK2].AsString := '';
        Fields[SP3].AsString := '';
        Fields[AUTHOR3].AsString := '';
        Fields[UNIQUE].AsInteger := 0;
        Fields[PLANTDESC].AsString := '';
        Fields[PHENOLOGY].AsString := '';
        Fields[DETBY].AsString := '';
        Fields[DETDD].AsInteger := 0;
        Fields[DETMM].AsInteger := 0;
        Fields[DETYY].AsInteger := 0;
        Fields[DETSTATUS].AsString := '';
        Fields[DETNOTES].AsString := '';
        Fields[COUNTRY].AsString := vcountry;
        Fields[MAJORAREA].AsString := vmajorarea;
        Fields[MINORAREA].AsString := vminorarea;
        Fields[GAZETTEER].AsString := vgazetteer;
        Fields[LOCNOTES].AsString := vlocnotes;
        Fields[HABITATTXT].AsString := '';
        Fields[CULTIVATED].AsBoolean := False;
        Fields[CULTNOTES].AsString := '';
        Fields[ORIGINSTAT].AsString := '';
        Fields[ORIGINID].AsString := '';
        Fields[ORIGINDB].AsString := '';
        Fields[NOTES].AsString := vnotes;
        Fields[LAT].AsFloat := vlat;
        Fields[NS].AsString := vns;
        Fields[LONG].AsFloat := vlon;
        Fields[EW].AsString := vew;
        Fields[LLGAZ].AsString := '';
        Fields[LLUNIT].AsString := '';
        Fields[LLRES].AsString := '';
        Fields[LLORIG].AsString := '';
        Fields[LLDATUM].AsString := '';
        Fields[QDS].AsString := '';
        Fields[ALT].AsInteger := valt;
        Fields[ALTMAX].AsInteger := valtmax;
        Fields[ALTUNIT].AsString := valtunit;
        Fields[ALTRES].AsString := '';
        Fields[ALTTEXT].AsString := '';
        Fields[INITIAL].AsInteger := 0;
        Fields[AVAILABLE].AsInteger := 0;
        Fields[CURATENOTE].AsString := '';
        Fields[NOTONLINE].AsString := '';
        Fields[LABELTOTAL].AsInteger := 0;
        Fields[VERNACULAR].AsString := '';
        Fields[LANGUAGE].AsString := '';
        Fields[GEODATA].AsString := '';
        Fields[LATLONG].AsString := '';
        Fields[COLLECTED].AsString := '';
        Fields[MONTHNAME].AsString := '';
        Fields[DETBYDATE].AsString := '';
        Fields[CHECKWHO].AsString := '';
        Fields[CHECKDATE].AsDateTime := Now;
        Fields[CHECKNOTE].AsString := '';
        Fields[UNIQUEID].AsString := '';
        Fields[RDESPEC].AsString := '';
      end;
      dbfile.Post;
      Dataset1.Next;
      Inc(reccount);
    end;
    Dataset1.EnableControls;
    Dataset1.GotoBookmark(bm);
  end;

  dbfile.Close;
  Result := reccount;
end;

procedure toCEP(const filename: string; var nrows, ncols: integer);
const
  fmt = '(I4,2X,7(I5,F5.1))';
type
  TIntMatrix = array of array of integer;
var
  DatasetName1: string;
  lSPECIES, lSAMPLE: TStringList;
  fSPECIES, fSAMPLE: TField;
  bm: TBookmark;
  sums: TIntMatrix = nil;
  idxSPECIES, idxSAMPLE, r, c, freq: integer;
  genus, cf, species, author1, subsp, infraname, author2: string;
  outfile: TextFile;
begin
  DefaultFormatSettings.DecimalSeparator := '.';
  with MainForm.MultiDoc.ActiveObject as TMDIChild do
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

      nrows := lSAMPLE.Count;
      ncols := lSPECIES.Count;
      SetLength(sums, nrows, ncols);

      Dataset1.First;
      while not Dataset1.EOF do
      begin
        lSPECIES.Find(fSPECIES.AsString, idxSPECIES);
        lSAMPLE.Find(fSAMPLE.AsString, idxSAMPLE);
        Inc(sums[idxSAMPLE, idxSPECIES]);
        Dataset1.Next;
      end;

      AssignFile(outfile, filename);
      Rewrite(outfile);
      WriteLn(outfile, Metadata.Find('title').AsString);
      WriteLn(outfile, fmt);
      WriteLn(outfile, ncols);

      for r := 1 to nrows do
      begin
        Write(outfile, format('%4d  ', [r]));
        for c := 1 to ncols do
        begin
          freq := sums[r - 1, c - 1];
          if freq > 0 then
            Write(outfile, format('%5d %5.1f', [c, double(freq)]));
        end;
        WriteLn(outfile);
      end;
      WriteLn(outfile, '0');

      for c := 1 to ncols do
      begin
        ParseName(lSPECIES[c - 1], genus, cf, species, author1, subsp,
          infraname, author2);
        Write(outfile, LeftStr(genus, 4) + LeftStr(species, 4) + ' ');
      end;
      WriteLn(outfile);

      for r := 1 to nrows do
        Write(outfile, DelSpace(lSAMPLE[r - 1]), ' ');
      CloseFile(outfile);

    finally
      lSPECIES.Free;
      lSAMPLE.Free;
      Dataset1.EnableControls;
      Dataset1.GotoBookmark(bm);
    end;
  end;
end;

procedure toMVSP(const filename: string; var nrows, ncols: integer);
type
  TIntMatrix = array of array of integer;
var
  DatasetName1: string;
  lSPECIES, lSAMPLE: TStringList;
  fSPECIES, fSAMPLE: TField;
  bm: TBookmark;
  sums: TIntMatrix = nil;
  idxSPECIES, idxSAMPLE, r, c: integer;
  genus, cf, species, author1, subsp, infraname, author2: string;
  outfile: TextFile;
begin
  with MainForm.MultiDoc.ActiveObject as TMDIChild do
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

      nrows := lSAMPLE.Count;
      ncols := lSPECIES.Count;
      SetLength(sums, nrows, ncols);

      Dataset1.First;
      while not Dataset1.EOF do
      begin
        lSPECIES.Find(fSPECIES.AsString, idxSPECIES);
        lSAMPLE.Find(fSAMPLE.AsString, idxSAMPLE);
        Inc(sums[idxSAMPLE, idxSPECIES]);
        Dataset1.Next;
      end;

      AssignFile(outfile, filename);
      Rewrite(outfile);
      WriteLn(outfile, '*MVSP3 ', nrows, ' ', ncols, ' ',
        Metadata.Find('title').AsString);

      for c := 1 to ncols do
      begin
        ParseName(lSPECIES[c - 1], genus, cf, species, author1, subsp,
          infraname, author2);
        Write(outfile, LeftStr(genus, 4) + LeftStr(species, 4) + ' ');
      end;
      WriteLn(outfile);

      for r := 1 to nrows do
      begin
        Write(outfile, lSAMPLE[r - 1], ' ');
        for c := 1 to ncols do
          Write(outfile, sums[r - 1, c - 1], ' ');
        WriteLn(outfile);
      end;
      CloseFile(outfile);

    finally
      lSPECIES.Free;
      lSAMPLE.Free;
      Dataset1.EnableControls;
      Dataset1.GotoBookmark(bm);
    end;
  end;
end;

procedure toFPM(const filename: string; var nrows, ncols: integer);
const
  VERSION = '2.1';
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
  with MainForm.MultiDoc.ActiveObject as TMDIChild do
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

      nrows := lSAMPLE.Count;
      ncols := lSPECIES.Count;
      SetLength(sums, nrows, ncols);

      Dataset1.First;
      while not Dataset1.EOF do
      begin
        lSPECIES.Find(fSPECIES.AsString, idxSPECIES);
        lSAMPLE.Find(fSAMPLE.AsString, idxSAMPLE);
        Inc(sums[idxSAMPLE, idxSPECIES]);
        Dataset1.Next;
      end;

      AssignFile(outfile, filename);
      Rewrite(outfile);
      WriteLn(outfile, 'FPM ', VERSION);
      WriteLn(outfile, Metadata.Find('title').AsString);
      WriteLn(outfile, Metadata.Find('author').AsString);
      WriteLn(outfile, StringReplace(DateToStr(Now), '/', ' ', [rfReplaceAll]),
        ' ', StringReplace(TimeToStr(Now), ':', ' ', [rfReplaceAll]));
      WriteLn(outfile, 'Q');
      WriteLn(outfile, ncols, ' ', nrows);
      WriteLn(outfile, 'especies');
      WriteLn(outfile, 'amostras');
      WriteLn(outfile);
      WriteLn(outfile);
      WriteLn(outfile);

      for c := 1 to ncols do
        WriteLn(outfile, 'TQ ', LeftStr(lSPECIES[c - 1], 70));

      for r := 1 to nrows do
        WriteLn(outfile, 'T ', LeftStr(lSAMPLE[r - 1], 70));

      for r := 1 to nrows do
      begin
        for c := 1 to ncols do
          Write(outfile, sums[r - 1, c - 1], ' ');
        WriteLn(outfile);
      end;
      CloseFile(outfile);

    finally
      lSPECIES.Free;
      lSAMPLE.Free;
      Dataset1.EnableControls;
      Dataset1.GotoBookmark(bm);
    end;
  end;
end;

procedure toCSV(const filename: string; var nrows, ncols: integer);
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
  with MainForm.MultiDoc.ActiveObject as TMDIChild do
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

      nrows := lSAMPLE.Count;
      ncols := lSPECIES.Count;
      SetLength(sums, nrows, ncols);

      Dataset1.First;
      while not Dataset1.EOF do
      begin
        lSPECIES.Find(fSPECIES.AsString, idxSPECIES);
        lSAMPLE.Find(fSAMPLE.AsString, idxSAMPLE);
        Inc(sums[idxSAMPLE, idxSPECIES]);
        Dataset1.Next;
      end;

      AssignFile(outfile, filename);
      Rewrite(outfile);
      Write(outfile, 'SITES/SPECIES', ',');
      for c := 1 to ncols do
      begin
        if c < ncols then
          Write(outfile, lSPECIES[c - 1], ',')
        else
          Write(outfile, lSPECIES[c - 1]);
      end;
      WriteLn(outfile);

      for r := 1 to nrows do
      begin
        Write(outfile, lSAMPLE[r - 1], ',');
        for c := 1 to ncols do
          if c < ncols then
            Write(outfile, sums[r - 1, c - 1], ',')
          else
            Write(outfile, sums[r - 1, c - 1]);
        WriteLn(outfile);
      end;
      CloseFile(outfile);

    finally
      lSPECIES.Free;
      lSAMPLE.Free;
      Dataset1.EnableControls;
      Dataset1.GotoBookmark(bm);
    end;
  end;
end;

end.

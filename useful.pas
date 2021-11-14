unit useful;

{$mode objfpc}{$H+}

interface

uses
  SysUtils, LCLIntf, LCLType, Classes, Graphics, Controls,
  Forms, Dialogs, Menus, StdCtrls, ExtCtrls, DB;

procedure DegToDMS(x: double; var Degrees, Minutes, Seconds: integer);
procedure NaturalSort(aList: TStrings);
procedure ParseName(const taxonomic_name: string; var genus: string;
  var nomen: string; var species: string; var author1: string;
  var ssp: string; var infraname: string; var author2: string);
function Extenso(Data: string): string;
function ExtractText(const Str: string; const Delim1, Delim2: char): string;
function GetFileNameWithoutExt(Filenametouse: string): string;
function GetUnit(const s: string): string;
function IsOnline(reliableserver: string = 'http://www.google.com'): boolean;
function Roman(Data: string): string;

implementation

uses StrUtils, fphttpclient, naturalsortunit;

const
  mesext: array[1..12] of string =
    ('I', 'II', 'III', 'IV', 'V', 'VI', 'VII', 'VIII', 'IX', 'X', 'XI', 'XII');

procedure DegToDMS(x: double; var Degrees, Minutes, Seconds: integer);
begin
  Degrees := Trunc(x);
  x := (x - Degrees) * 60;
  Minutes := Trunc(x);
  x := (x - Minutes) * 60;
  Seconds := Round(x);
end;

procedure NaturalSort(aList: TStrings);
var
  L: TStringList;
begin
  L := TStringList.Create;
  try
    L.Assign(aList);
    L.CustomSort(@UTF8NaturalCompareList);
    aList.Assign(L);
  finally
    L.Free;
  end;
end;

procedure ParseName(const taxonomic_name: string; var genus: string;
  var nomen: string; var species: string; var author1: string;
  var ssp: string; var infraname: string; var author2: string);
var
  taxonomy: TStringList;
begin
  if taxonomic_name = '' then
    Exit;
  genus := '';
  nomen := '';
  species := '';
  author1 := '';
  ssp := '';
  infraname := '';
  author2 := '';
  taxonomy := TStringList.Create;
  taxonomy.Delimiter := ' ';
  taxonomy.StrictDelimiter := True;
  taxonomy.DelimitedText := taxonomic_name;
  genus := taxonomy[0];
  taxonomy.Delete(0);
  if taxonomy.Count > 0 then
  begin
    if taxonomy[0].StartsWith('cf.') or taxonomy[0].StartsWith('aff.') then
    begin
      nomen := taxonomy[0];
      taxonomy.Delete(0);
    end;
  end;
  if taxonomy.Count > 0 then
  begin
    species := taxonomy[0];
    taxonomy.Delete(0);
  end;
  if taxonomy.Count > 0 then
  begin
    author1 := taxonomy[0];
    taxonomy.Delete(0);
  end;
  if taxonomy.Count > 0 then
  begin
    if taxonomy[0].StartsWith('var.') or taxonomy[0].StartsWith('subsp.') then
      ssp := taxonomy[0];
    taxonomy.Delete(0);
  end;
  if taxonomy.Count > 0 then
  begin
    infraname := taxonomy[0];
    taxonomy.Delete(0);
  end;
  if taxonomy.Count > 0 then
  begin
    author2 := taxonomy[0];
    taxonomy.Delete(0);
  end;
  taxonomy.Free;
end;

function Extenso(Data: string): string;
var
  dia, mes, ano: string;
begin
  if Data = '' then
    Result := '';
  dia := Data.Split('-')[0];
  mes := Data.Split('-')[1];
  ano := Data.Split('-')[1];
  Result := dia + ' de ' + mesext[StrToIntDef(mes, 1)] + ' de ' + ano;
end;

function ExtractText(const Str: string; const Delim1, Delim2: char): string;
var
  pos1, pos2: integer;
begin
  Result := '';
  pos1 := Pos(Delim1, Str);
  pos2 := Pos(Delim2, Str);
  if (pos1 > 0) and (pos2 > pos1) then
    Result := Copy(Str, pos1 + 1, pos2 - pos1 - 1);
end;

function GetFileNameWithoutExt(Filenametouse: string): string;
begin
  Result := ExtractFilename(Copy(Filenametouse, 1,
    RPos(ExtractFileExt(Filenametouse), Filenametouse) - 1));
end;

function GetUnit(const s: string): string;
begin
  if (Pos('(', s) > 0) and (Pos(')', s) > 0) then
    Result := ExtractText(s, '(', ')')
  else
    Result := '';
end;

function IsOnline(reliableserver: string = 'http://www.google.com'): boolean;
var
  http: tfphttpclient;
  httpstate: integer;
begin
  Result := False;
  try
    http := tfphttpclient.Create(nil);
    try
      http.Get(reliableserver);
      httpstate := http.ResponseStatusCode;
      if httpstate = 200 then
        Result := True
      else
        Result := False;
    except
      on E: Exception do
        Result := False;
    end;
  finally
    http.Free;
  end;
end;

function Roman(Data: string): string;
var
  dia, mes, ano: string;
begin
  if Data = '' then
    Result := '';
  dia := Data.Split('-')[0];
  mes := Data.Split('-')[1];
  ano := Data.Split('-')[1];
  Result := dia + '-' + mesext[StrToIntDef(mes, 1)] + '-' + ano;
end;

end.

unit DNA;

{$MODE DELPHI}

interface

uses
  Classes, SysUtils, StrUtils;

type
  TDNA = class(TObject)
  private
    seq: string;
    {cds, tally: TStringList;}
  public
    constructor Create(s: string); { Create DNA instance initialized to string s. }
    function Sequence: string; { return DNA sequence. }
    function Transcribe: string; { Return as rna string. }
    function Reverse: string; { Return dna string in reverse order. }
    function Complement: string; { Return the complementary dna string. }
    function ReverseComplement: string;
    { Return the reverse complement of the dna string. }
    function GC: double; { Return the percentage of dna composed of G+C. }
    function Codons: TStringList; { Return list of codons for the dna string. }
    function Frequency(cds: TStringList): TStringList;
    { Return frequency table of codons for the dna string. }
  end;

implementation

constructor TDNA.Create(s: string);
begin
  seq := UpperCase(s);
end;

{destructor TDNA.Destroy;
begin
if Assigned(cds) then
  FreeAndNil(cds);
if Assigned(tally) then
  FreeAndNil(tally);
end;}

function TDNA.Sequence: string;
begin
  Result := seq;
end;

function TDNA.Transcribe: string;
begin
  Result := StringReplace(seq, 'T', 'U', [rfReplaceAll, rfIgnoreCase]);
end;

function TDNA.Reverse: string;
begin
  Result := ReverseString(seq);
end;

function TDNA.Complement: string;
var
  c: string;
  i: integer;
begin
  c := '';
  for i := 1 to Length(seq) do
  begin
    if (seq[i] = 'A') then
      c := c + 'T'
    else if (seq[i] = 'T') then
      c := c + 'A'
    else if (seq[i] = 'C') then
      c := c + 'G'
    else if (seq[i] = 'G') then
      c := c + 'C'
    else
      c := c + seq[i];
  end;
  Result := c;
end;

function TDNA.ReverseComplement: string;
var
  s, c: string;
  i: integer;
begin
  s := ReverseString(seq);
  c := '';
  for i := 1 to Length(s) do
  begin
    if (s[i] = 'A') then
      c := c + 'T'
    else if (s[i] = 'T') then
      c := c + 'A'
    else if (s[i] = 'C') then
      c := c + 'G'
    else if (s[i] = 'G') then
      c := c + 'C'
    else
      c := c + s[i];
  end;
  Result := c;
end;

function TDNA.GC: double;
var
  c: double;
begin
  c := seq.CountChar('G') + seq.CountChar('C');
  Result := c * 100.0 / Length(seq);
end;

function TDNA.Codons: TStringList;
var
  i, c: integer;
  codon: string;
  cds: TStringList;
begin
  c := 0;
  codon := '';
  cds := TStringList.Create;
  for i := 1 to Length(seq) do
  begin
    codon := codon + seq[i];
    Inc(c);
    if (c = 3) then
    begin
      cds.Add(codon);
      codon := '';
      c := 0;
    end;
  end;
  Result := cds;
end;

function TDNA.Frequency(cds: TStringList): TStringList;
var
  tally: TStringList;
  i, k: integer;

  function search(arr: TStringList; s: string; n: integer): integer;
  var
    j, counter: integer;
  begin
    counter := 0;
    for j := 0 to n - 1 do
      if arr[j] = s then
        Inc(counter);
    Result := counter;
  end;

begin
  tally := TStringList.Create;
  tally.Sorted := True;
  tally.Duplicates := dupIgnore;
  for i := 0 to cds.Count - 1 do
  begin
    k := search(cds, cds[i], cds.Count);
    tally.Add(cds[i] + tally.NameValueSeparator + IntToStr(k));
  end;
  Result := tally;
end;

end.

unit diversity;

{$mode objfpc}{$H+}

interface

uses
  SysUtils, LCLIntf, LCLType, Classes;

procedure BioDiv(s, n: integer; ni: array of integer;
  var d1, d2, c, h, pie, d, m, j: double);

procedure CreateDivers(fname: string);

procedure Divers(fname: string; n, m: integer);

implementation

uses report, useful;

resourcestring
  strDiversity = 'ANÁLISE DE DIVERSIDADE';
  strRarefactionCurve = 'CURVA DE RAREFAÇÃO';
  strSampleSize = 'NÚMERO DE AMOSTRAS';
  strExpectedNumber = 'NÚMERO ESPERADO';
  strSample = 'LOCAL';
  strSamples = 'Amostras';
  strSpecies = 'Espécies';

procedure BioDiv(s, n: integer; ni: array of integer;
  var d1, d2, c, h, pie, d, m, j: double);

var
  u: double;
  i, nmax: integer;

begin
  c := 0.0;
  h := 0.0;
  pie := 0.0;
  u := 0.0;
  d1 := (s - 1) / Ln(n);             //  Margalef
  d2 := s / Sqrt(n);                 //  Menhinick
  nmax := ni[0];
  for i := 0 to s - 1 do
  begin
    c := c + ((ni[i] * (ni[i] - 1)) / (n * (n - 1)));
    h := h + -(ni[i] / n) * Ln((ni[i] / n));     // Shannon-Weaver
    pie := pie + (ni[i] / n) * (ni[i] / n);
    u := u + (ni[i] * ni[i]);
    if ni[i] > nmax then
      nmax := ni[i];
  end;
  c := 1 - c;                         //  Simpson
  j := h / Ln(s);                     //  Evenness
  pie := (n / (n - 1)) * (1 - pie);   //  Hurlbert
  d := 1 - (nmax / n);                //  Berger-Parker
  m := (n - Sqrt(u)) / (n - Sqrt(n)); //  MacIntosh
end;

procedure CreateDivers(fname: string);
var
  figf: string;
  outfile: TextFile;
begin
  AssignFile(outfile, 'diversity.R');
  Rewrite(outfile);
  WriteLn(outfile, 'options(warn=-1)');
  WriteLn(outfile, 'suppressPackageStartupMessages(library(vegan))');
  WriteLn(outfile, 'library(vegan, quietly=TRUE)');
  WriteLn(outfile, 'df.data <- read.csv("rdata.csv", row.names=1)');
  WriteLn(outfile, 'S <- apply(df.data > 0, 2, sum)');
  WriteLn(outfile, '#N <- apply(df.data, 2, sum)');
  WriteLn(outfile, 'c <- diversity(t(df.data), "simpson")');
  WriteLn(outfile, 'H <- diversity(t(df.data))');
  WriteLn(outfile, 'N0 <- S');
  WriteLn(outfile, 'N1 <- exp(H)');
  WriteLn(outfile, 'N2 <- diversity(t(df.data), "inv")');
  WriteLn(outfile, 'result <- data.frame(N0, H, N1, 1-c, N2)');
  WriteLn(outfile,
    'write.table(result, "diversity.csv", sep=";", row.names=TRUE, col.names=FALSE)');
  WriteLn(outfile, 'curve <- specaccum(df.data, method="rarefaction")');
  WriteLn(outfile, 'ppi <- 100');
  figf := GetFileNameWithoutExt(fname) + '.png';
  WriteLn(outfile, 'png("' + figf + '", width=6*ppi, height=6*ppi, res=ppi)');
  //WriteLn(outfile, 'par(mar=c(4,4,4,4))');
  WriteLn(outfile, 'plot(curve, main=iconv("' + strRarefactionCurve +
    '", from=''UTF-8'', to=''latin1''), xlab=iconv("' + strSampleSize +
    '", from=''UTF-8'', to=''latin1''), ylab=iconv("' + strExpectedNumber +
    '", from=''UTF-8'', to=''latin1''), col="red", ci=0, lwd=2)');
  WriteLn(outfile, 'invisible(dev.off())');
  WriteLn(outfile, 'options(warn=0)');
  CloseFile(outfile);
end;

procedure Divers(fname: string; n, m: integer);
var
  L, C: TStringList;
  i, j: integer;
  figf: string;
  outfile: TextFile;
begin
  AssignFile(outfile, fname);
  Rewrite(outfile);
  Header(outfile, strDiversity);
  WriteLn(outfile, '<br>');
  WriteLn(outfile, IntToStr(n) + ' ' + LowerCase(strSamples) + ' x ' +
    IntToStr(m) + ' ' + LowerCase(strSpecies) + '<br><br>');

  WriteLn(outfile, '<table border=1 cellspacing=1 cellpadding=1 width="100%">');
  WriteLn(outfile, '<tr><th>' + strSample + '</th>');
  WriteLn(outfile, '<th>N0 (S)</th>');
  WriteLn(outfile, '<th>H''</th>');
  WriteLn(outfile, '<th>N1</th>');
  WriteLn(outfile, '<th>1-&lambda;</th>');
  WriteLn(outfile, '<th>N2 (1/&lambda;)</th>');
  WriteLn(outfile, '</tr>');

  DefaultFormatSettings.DecimalSeparator := '.';
  L := TStringList.Create;
  L.LoadFromFile('diversity.csv');
  C := TStringList.Create;
  C.Delimiter := ';';
  for i := 0 to L.Count - 1 do
  begin
    WriteLn(outfile, '<tr>');
    C.DelimitedText := L[i];
    for j := 0 to C.Count - 1 do
    begin
      if j = 0 then
        WriteLn(outfile, '<td align="Left">' + C[j] + '</td>')
      else if j = 1 then
        WriteLn(outfile, '<td align="Center">' + C[j] + '</td>')
      else if (j > 1) and (j < 5) then
        WriteLn(outfile, '<td align="Center">' +
          FloatToStrF(StrToFloat(C[j]), ffFixed, 5, 3) + '</td>')
      else if j = 5 then
        WriteLn(outfile, '<td align="Center">' +
          FloatToStrF(StrToFloat(C[j]), ffFixed, 5, 1) + '</td>');
    end;
    WriteLn(outfile, '</tr>');
  end;
  C.Free;
  L.Free;
  WriteLn(outfile, '</table>');

  figf := GetFileNameWithoutExt(fname) + '.png';
  WriteLn(outfile, '<p align="center"><img src="' + figf + '"></p>');

  WriteLn(outfile, '</body>');
  WriteLn(outfile, '</html>');
  CloseFile(outfile);
  if FileExists('diversity.csv') then
    DeleteFile('diversity.csv');
end;

end.

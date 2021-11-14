unit Eval;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, FileUtil, StdCtrls;

function Eval(expression: string): variant;

implementation

function isOp(aSign: string): boolean;
begin
  Result := False;
  if aSign = '+' then
    Result := True;
  if aSign = '-' then
    Result := True;
  if aSign = '*' then
    Result := True;
  if aSign = '/' then
    Result := True;
end;

function GetToken(expression: string; out Counter: integer): variant;
var
  aChar: string;
  sData: string;
begin
  repeat
    aChar := expression[Counter];
    if aChar = ' ' then
      aChar := '';
    if isOp(aChar) then
      break
    else
    begin
      sData := sData + aChar;
      Inc(Counter);
    end;
  until Counter > Length(expression);
  Result := StrToFloat(sData);
end;

function Eval(expression: string): variant;
var
  aChar: string;
  aOperator: string;
  Ret: variant;
  iCount: integer;
begin
  iCount := 1;
  while iCount < Length(expression) do
  begin
    aChar := expression[iCount];
    if isOp(aChar) then
    begin
      aOperator := Copy(expression, iCount, 1);
      Inc(iCount);
    end;
    if aOperator = '' then
      Ret := Trim(GetToken(expression, iCount));
    if aOperator = '+' then
      Ret := Ret + GetToken(expression, iCount);
    if aOperator = '-' then
      Ret := Ret - GetToken(expression, iCount);
    if aOperator = '*' then
      Ret := Ret * GetToken(expression, iCount);
    if aOperator = '/' then
      Ret := Ret / GetToken(expression, iCount);
  end;
  Result := Ret;
end;

end.

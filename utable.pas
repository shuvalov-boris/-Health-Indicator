unit UTable;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils;

type

  { TTableData }

  TTableData = class (TObject)
  private
    FData: array of array of Double;  // таблица данных
    FSubject: array of String;        // наименование территорий (строки)
    FYear: array of Integer;          // года (столбцы)
    FSubjectType: String;
    function GetData(ASubject: String; AYear: Integer): Double;
    function GetSubject(AIndex: Integer): String;
    function GetYear(AIndex: Integer): Integer;
    function GetColCount: Integer;
    function GetYearNumber: Integer;
    function GetRowCount: Integer;
    function GetSubjectNumber: Integer;
    function GetSubjectType: String;
    function GetDataByIndex(ARow, ACol: Integer): Double;
  public
    constructor Create;
    destructor Destroy; override;
    property Data[ASubject: String; AYear: Integer]: Double read GetData;
    property DataByIndex[ARow, ACol: Integer]: Double read GetDataByIndex; Default;
    property Subject[ANumber: Integer]: String read GetSubject;
    property Year[ANumber: Integer]: Integer read GetYear;
    property ColCount: Integer read GetColCount;
    property YearCount: Integer read GetYearNumber;
    property RowCount: Integer read GetRowCount;
    property SubjectCount: Integer read GetSubjectNumber;
    property SubjectType: String read GetSubjectType;
    procedure setHeader(const header: String);  // Горизонтальный
    procedure addRecord(const rec: String);
  end;

implementation

function TTableData.GetData(ASubject: String; AYear: Integer): Double;
var
  i, j: Integer;
begin
  i := 0;
  While i < Length(FSubject) do begin
    if FSubject[i] = ASubject then break;
    Inc(i);
  end;
  j := 0;
  While j < Length(FYear) do begin
    if FYear[j] = AYear then break;
    Inc(j)
  end;
  if (i < Length(FSubject)) and (j < Length(FYear)) then
    Result := FData[i, j]
  else
    Result := -1.0;
end;

function TTableData.GetSubject(AIndex: Integer): String;
begin
  if (AIndex >= 0) and (AIndex < RowCount) then
    Result := FSubject[AIndex];
end;

function TTableData.GetYear(AIndex: Integer): Integer;
begin
  if (AIndex >= 0) and (AIndex < ColCount) then
    Result := FYear[AIndex];
end;

function TTableData.GetColCount: Integer;
begin
  Result := Length(FYear);
end;

function TTableData.GetYearNumber: Integer;
begin
  Result := Length(FYear);
end;

function TTableData.GetRowCount: Integer;
begin
  Result := Length(FSubject);
end;

function TTableData.GetSubjectNumber: Integer;
begin
  Result := Length(FSubject);
end;

function TTableData.GetSubjectType: String;
begin
  Result := FSubjectType;
end;

function TTableData.GetDataByIndex(ARow, ACol: Integer): Double;
begin
  If (ARow >= 0) and (ARow < RowCount) and (ACol >= 0) and (ACol < ColCount) then
    Result := FData[ARow, ACol];
end;

constructor TTableData.Create;
begin

end;

destructor TTableData.Destroy;
begin
  if FYear <> NIL then FYear := NIL;
  if FSubject <> NIL then FSubject := NIL;
  if FData <> NIL then FData := NIL;
end;

procedure TTableData.setHeader(const header: String);
var
  List: TStringList;
  i: Integer;
begin
  List := TStringList.Create;
  List.Delimiter := #9;
  List.StrictDelimiter := True;
  List.DelimitedText := header;

  FSubjectType := List[0];

  SetLength(FYear, List.Count - 1);
  For i := 1 to List.Count - 1 do
    FYear[i - 1] := StrToInt(List[i]);
end;

procedure TTableData.addRecord(const rec: String);
var
  List: TStringList;
  i: Integer;
begin
  List := TStringList.Create;
  List.Delimiter := #9;
  List.StrictDelimiter := True;
  List.DelimitedText := rec;

  SetLength(FSubject, Length(FSubject) + 1);
  FSubject[Length(FSubject) - 1] := List[0];

  SetLength(FData, Length(FData) + 1, Length(FYear));
  For i := 1 to List.Count - 1 do
    if List[i] = '' then
      FData[Length(FSubject) - 1, i - 1] := 0
    else
      FData[Length(FSubject) - 1, i - 1] := StrToFloat(List[i]);
end;

end.


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
    function GetData(AName: String; AYear: Integer): Double;
    procedure SetData(AName: String; AYear: Integer; const AValue: Double);
    function GetSubject(AIndex: Integer): String;
    function GetYear(AIndex: Integer): Integer;
    function GetColCount: Integer;
    function GetRowCount: Integer;
  public
    constructor Create;
    destructor Destroy; override;
    property Data[AName: String; AYear: Integer]: Double read GetData write SetData;
    property Subject[ANumber: Integer]: String read GetSubject;
    property Year[ANumber: Integer]: Integer read GetYear;
    property ColCount: Integer read GetColCount;
    property RowCount: Integer read GetRowCount;
    procedure setHeader(const header: String);  // Горизонтальный
    procedure addRecord(const rec: String);
  end;

implementation

function TTableData.GetData(AName: String; AYear: Integer): Double;
var
  i, j: Integer;
begin
  i := 0;
  While i < Length(FSubject) do
    if FSubject[i] = AName then break;
  j := 0;
  While i < Length(FYear) do
    if FYear[i] = AYear then break;
  Result := FData[i, j];
end;

procedure TTableData.SetData(AName: String; AYear: Integer; const AValue: Double);
begin

end;

function TTableData.GetSubject(AIndex: Integer): String;
begin
  if (AIndex >= 0) and (AIndex < RowCount) then
    Result := FSubject[AIndex];
end;

function TTableData.GetYear(AIndex: Integer): Integer;
begin
  if (AIndex >= 0) and (AIndex < RColCount) then
    Result := FYear[AIndex];
end;

function TTableData.GetColCount: Integer;
begin
  Result := Length(FYear);
end;

function TTableData.GetRowCount: Integer;
begin
  Result := Length(FSubject);
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
    FData[Length(FSubject) - 1, i - 1] := StrToFloat(List[i]);
end;

end.


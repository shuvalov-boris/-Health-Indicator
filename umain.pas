unit UMain;

{$mode objfpc}{$H+}

interface

uses
  UTable,
  Classes, SysUtils, FileUtil, Forms, Controls, Graphics, Dialogs, Grids,
  ExtCtrls, StdCtrls, ActnList, ComCtrls, LazUTF8;

const
  TChauvenetCriterion: array [5..50] of Double  = (1.680, 1.730, 1.790, 1.860, 1.920,
         1.960, 1.995, 2.030, 2.065, 2.100, 2.130, 2.160, 2.180, 2.200, 2.220,
         2.240, 2.260, 2.280, 2.295, 2.310, 2.335, 2.360, 2.368, 2.375, 2.383,
         2.390, 2.401, 2.412, 2.423, 2.434, 2.445, 2.456, 2.467, 2.478, 2.489,
         2.500, 2.508, 2.516, 2.524, 2.532, 2.540, 2.548, 2.556, 2.564, 2.572,
         2.580);

type

  TChauvenet = record
    Rang: Integer;
    Subject: String;
    Year: Integer;
    Pi_min: Double;
    Deviation: Double;
    Deviation_sq: Double;
    Population: Double;
  end;

  { TForm1 }

  TForm1 = class(TForm)
    btnLoadMorbidity: TButton;
    btnRun: TButton;
    btnLoadPopulation: TButton;
    eMeanValue: TEdit;
    eMSD: TEdit;
    Edit3: TEdit;
    gbStatValues: TGroupBox;
    gbInput: TGroupBox;
    gbVariationSeries: TGroupBox;
    lMeanValue: TLabel;
    lMSD: TLabel;
    Label3: TLabel;
    mLog: TMemo;
    OpenDialog1: TOpenDialog;
    pcInput: TPageControl;
    pData: TPanel;
    pControl: TPanel;
    sgMorbidity: TStringGrid;
    sgVariationSeries: TStringGrid;
    sgPopulation: TStringGrid;
    tsMorbidity: TTabSheet;
    tsPopulation: TTabSheet;
    procedure btnLoadMorbidityClick(Sender: TObject);
    procedure btnRunClick(Sender: TObject);
    procedure btnLoadPopulationClick(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
  private
    isPopulationDataLoad: Boolean;
    ObjectName: String;
    Morbidity: array of array of Double;
    tdMorbidity: TTableData;
    Population: array of array of Double;
    Chauvenet: array of TChauvenet;
    procedure SetMorbidityData;
    procedure SetPopulationData;
    procedure CalcChauvenet;
    procedure CalcLongTimeAverageAnnualMinimum;
    procedure AddToLog(str: String);
    procedure ClearCalculatedData;
    function getPopulationData(oName: String; Year: Integer) : Integer;
    { private declarations }
  public
    { public declarations }
  end;

var
  Form1: TForm1;

implementation

{$R *.lfm}

{ TForm1 }

procedure TForm1.SetMorbidityData;
var
  i, j: Integer;
begin
  //ShowMessage('ACol = ' + IntToStr(sgMorbidity.ColCount) + 'ARow = ' +
  //                  IntToStr(sgMorbidity.RowCount));
  SetLength(Morbidity, sgMorbidity.RowCount - 1, sgMorbidity.ColCount - 1);
  for i := 1 to sgMorbidity.RowCount - 1 do
    for j := 1 to sgMorbidity.ColCount - 1 do
      if sgMorbidity.Cells[j, i] = '' then
        Morbidity[i - 1, j - 1] := 0.0
      else
        Morbidity[i - 1, j - 1] := StrToFloat(sgMorbidity.Cells[j, i]);
end;

procedure TForm1.SetPopulationData;
var
  i, j: Integer;
begin
  SetLength(Population, sgPopulation.RowCount - 1, sgPopulation.ColCount - 1);
  for i := 1 to sgPopulation.RowCount - 1 do
    for j := 1 to sgPopulation.ColCount - 1 do
      if sgPopulation.Cells[j, i] = '' then
        Population[i - 1, j - 1] := 0.0
      else
        Population[i - 1, j - 1] := StrToFloat(sgPopulation.Cells[j, i]);
end;

procedure TForm1.CalcChauvenet;
var
  i, j, k, j_min: Integer;
  min, sum, X_min, msd, U1, Un, Ut: Double;
  tmp: TChauvenet;
begin
  SetLength(Chauvenet, sgMorbidity.RowCount - 1);
  AddToLog('Расчет наименьших годовых показателей...');
  For i := 0 to Length(Chauvenet) - 1 do begin
    Chauvenet[i].Subject := sgMorbidity.Cells[1, i + 1];
    k := 0;
    while (k < Length(Morbidity[i]) - 1) and (Morbidity[i, k] = 0) do
      Inc(k);
    min := Morbidity[i][k];
    j_min := k;
    For j := k + 1 to Length(Morbidity[i]) - 1 do
      If (Morbidity[i][j] > 0.0) and (Morbidity[i][j] < min) Then begin
        min := Morbidity[i][j];
        j_min := j;
      end;
    Chauvenet[i].Subject := sgMorbidity.Cells[0, i + 1];
    Chauvenet[i].Pi_min := min;
    Chauvenet[i].Year := StrToInt(sgMorbidity.Cells[j_min + 1, 0]);
  end;

  AddToLog('Ранжирование показателей');
  For i := 0 to Length(Chauvenet) - 2 do
      For j := 0 to Length(Chauvenet) - 2 - i do
        If Chauvenet[j].Pi_min > Chauvenet[j + 1].Pi_min Then begin
          tmp := Chauvenet[j];
          Chauvenet[j] := Chauvenet[j + 1];
          Chauvenet[j + 1] := tmp;
        end;

  sgVariationSeries.RowCount := 1;
  sgVariationSeries.ColCount := 6;
  sgVariationSeries.Cells[0, 0] := 'Ранг';
  sgVariationSeries.Cells[1, 0] := ObjectName;
  sgVariationSeries.Cells[2, 0] := 'Год';
  sgVariationSeries.Cells[3, 0] := 'Pi_min';
  sgVariationSeries.Cells[4, 0] := 'Pi_min-X_min';
  sgVariationSeries.Cells[5, 0] := '(Pi_min-X_min)^2';

  sgVariationSeries.Visible := True;


  repeat

    AddToLog('Расчет суммы всех вариант вариационного ряда');
    sum := 0.0;
    For i := 0 to Length(Chauvenet) - 1 do
      sum := sum + Chauvenet[i].Pi_min;
    AddToLog('Sum = ' + FloatToStr(sum));
    X_min := sum / Length(Chauvenet);
    AddToLog('Среднее равно ' + FloatToStr(X_min));
    eMeanValue.Text := FloatToStr(X_min);
    //ShowMessage('X_min = ' + FloatToStr(X_min));

    sum := 0.0;
    For i := 0 to Length(Chauvenet) - 1 do
      With Chauvenet[i] do begin
        Rang := i + 1;
        Deviation := Pi_min - X_min;
        Deviation_sq := Deviation * Deviation;
        sum := sum + Deviation_sq;
      end;
    msd := sqrt(sum / (Length(Chauvenet) - 1) );
    eMSD.Text := FloatToStr(msd);
    //ShowMessage('sigma (msd) = ' + FloatToStr(msd));

    sgVariationSeries.RowCount := Length(Chauvenet) + 1;

    For i := 0 to Length(Chauvenet) - 1 do begin
      sgVariationSeries.Cells[0, i + 1] := IntToStr(Chauvenet[i].Rang);
      sgVariationSeries.Cells[1, i + 1] := Chauvenet[i].Subject;
      sgVariationSeries.Cells[2, i + 1] := IntToStr(Chauvenet[i].Year);
      sgVariationSeries.Cells[3, i + 1] := FloatToStr(Chauvenet[i].Pi_min);
      sgVariationSeries.Cells[4, i + 1] := FloatToStr(Chauvenet[i].Deviation);
      sgVariationSeries.Cells[5, i + 1] := FloatToStr(Chauvenet[i].Deviation_sq);
    end;


    U1 := (X_min - Chauvenet[0].Pi_min) / msd;
    AddToLog('U1 = ' + FloatToStr(U1));
    //ShowMessage('U1 = ' + FloatToStr(U1));
    Un := (Chauvenet[Length(Chauvenet) - 1].Pi_min - X_min) / msd;
    AddToLog('U16 = ' + FloatToStr(Un));
    //ShowMessage('U16 = ' + FloatToStr(Un));

    Ut := TChauvenetCriterion[Length(Chauvenet)];
    AddToLog('Ut = ' + FloatToStr(Ut));
    //ShowMessage('Ut = ' + FloatToStr(Ut));
    if (Un >= Ut) then begin
      SetLength(Chauvenet, Length(Chauvenet) - 1);
      AddToLog('Un - аномальная величина');
      //ShowMessage('Un - аномальная величина');
    end;
    if (U1 >= Ut) then begin
      for i := 0 to Length(Chauvenet) - 2 do
        Chauvenet[i] := Chauvenet[i + 1];
      SetLength(Chauvenet, Length(Chauvenet) - 1);
      AddToLog('U1 - аномальная величина');
      ShowMessage('U1 is Bad');
    end;

  until (U1 < Ut) and (Un < Ut);

end;

procedure TForm1.CalcLongTimeAverageAnnualMinimum;
begin
  //getPopulationData(oName, year)
end;

procedure TForm1.AddToLog(str: String);
begin
  mLog.Text := mLog.Text + str + #10;
end;

procedure TForm1.ClearCalculatedData;
begin
  sgVariationSeries.Clear;
  eMeanValue.Text := '';
  eMSD.Text := '';
end;

function TForm1.getPopulationData(oName: String; Year: Integer) : Integer;
var
  i, j: Integer;
begin

end;

procedure TForm1.btnLoadMorbidityClick(Sender: TObject);
var
  fin: TextFile;
  FileName, s: String;
  //Stream: TFileStream;
  //Parser: TParser;
  //Col, i: Integer;
  Row: Integer;
  List1: TStringList;
begin
  {Col := 0;
  Row := 0;
  If openDialog1.Execute Then FileName:=openDialog1.FileName;
  AssignFile(fin, FileName);
  //Reset (fin);
  //Readln(fin, s);
  //showmessage(nam);

  try
    Stream := TFileStream.Create(FileName, fmOpenReadWrite);
  except
    on E: EFOpenError do
       ShowMessage(E.ClassName + ' ошибка: ' + E.Message);
  end;

  Parser := TParser.Create(Stream);

  while Parser.Token<>toEOF do
  begin
      if sgMorbidity.ColCount <= Col then
        sgMorbidity.ColCount := Col + 1;
      sgMorbidity.Cells[Col, Parser.SourceLine - 1] := Parser.TokenString;
    Parser.NextToken;
    Inc(Col);
    if (Parser.SourceLine - 1 > Row) then begin
      Inc(Row);
      Col := 0;
      if sgMorbidity.RowCount <= Row then
        sgMorbidity.RowCount := Row + 1;
    end;
    ShowMessage(Parser.TokenString+' '+IntToStr(Parser.SourceLine));
  end;

  //CloseFile(fin);   }

  tdMorbidity := TTableData.Create;
  If openDialog1.Execute Then FileName:=openDialog1.FileName;
  AssignFile(fin, FileName);
  Reset (fin);
  List1 := TStringList.Create;
  List1.Delimiter := #9;
  List1.StrictDelimiter := true;
  sgMorbidity.Options := sgMorbidity.Options + [goColSizing];
  Row := 0;
  sgMorbidity.ColCount := 1;
  sgMorbidity.RowCount := 1;
  while not EOF(fin) do begin
    Readln(fin, s);
    if Row = 0 then tdMorbidity.setHeader(s)
    else tdMorbidity.addRecord(s);
    List1.Clear;
    List1.DelimitedText := s;
    if List1.Count > sgMorbidity.ColCount then
      sgMorbidity.ColCount := List1.Count;
    If sgMorbidity.RowCount <= Row then
      sgMorbidity.RowCount := Row + 1;
    //ShowMessage(IntToStr(Row));
    sgMorbidity.Rows[Row] := List1;
    //for i := 0 to List1.Count - 1 do
    Inc(Row);
  end;
  ObjectName := sgMorbidity.Cells[0, 0];
  sgMorbidity.Visible := True;
  ClearCalculatedData;
  AddToLog('Таблица ''Стандартизованные показатели заболеваемости'' загружена');
  btnRun.Enabled := True;
end;

procedure TForm1.btnRunClick(Sender: TObject);
begin
  SetMorbidityData();
  CalcChauvenet();
  if isPopulationDataLoad then
    CalcLongTimeAverageAnnualMinimum;
  sgVariationSeries.Visible := True;
end;

procedure TForm1.btnLoadPopulationClick(Sender: TObject);
var
  fin: TextFile;
  FileName, s: String;
  Row: Integer;
  List1: TStringList;
begin
  If openDialog1.Execute Then FileName:=openDialog1.FileName;
  AssignFile(fin, FileName);
  Reset (fin);
  List1 := TStringList.Create;
  List1.Delimiter := #9;
  List1.StrictDelimiter := true;
  sgPopulation.Options := sgPopulation.Options + [goColSizing];
  Row := 0;
  sgPopulation.ColCount := 1;
  sgPopulation.RowCount := 1;
  while not EOF(fin) do begin
    Readln(fin, s);
    List1.Clear;
    List1.DelimitedText := s;
    if List1.Count > sgPopulation.ColCount then
      sgPopulation.ColCount := List1.Count;
    If sgPopulation.RowCount <= Row then
      sgPopulation.RowCount := Row + 1;
    //ShowMessage(IntToStr(Row));
    sgPopulation.Rows[Row] := List1;
    //for i := 0 to List1.Count - 1 do
    Inc(Row);
  end;
  sgPopulation.Visible := True;
  isPopulationDataLoad := True;
  //ClearCalculatedData;
  AddToLog('Таблица ''Численность населения'' загружена');
end;

procedure TForm1.Button3Click(Sender: TObject);
begin
  sgMorbidity.SaveToFile('output.txt');
end;

procedure TForm1.FormCreate(Sender: TObject);
begin
  isPopulationDataLoad := False;
  mLog.text := '';
end;

end.


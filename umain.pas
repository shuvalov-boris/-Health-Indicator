unit UMain;

{$mode objfpc}{$H+}

interface

uses
  UTable,
  Classes, SysUtils, FileUtil, Forms, Controls, Graphics, Dialogs, Grids,
  ExtCtrls, StdCtrls, ActnList, ComCtrls, LazUTF8;

const
  CHAUVENET_CRITERION: array [5..50] of Double  = (1.680, 1.730, 1.790, 1.860,
         1.920, 1.960, 1.995, 2.030, 2.065, 2.100, 2.130, 2.160, 2.180, 2.200,
         2.220, 2.240, 2.260, 2.280, 2.295, 2.310, 2.335, 2.360, 2.368, 2.375,
         2.383, 2.390, 2.401, 2.412, 2.423, 2.434, 2.445, 2.456, 2.467, 2.478,
         2.489, 2.500, 2.508, 2.516, 2.524, 2.532, 2.540, 2.548, 2.556, 2.564,
         2.572, 2.580);

  CHAUVENET_HEADER: array [0..7] of String = ('Ранг', 'Субъект', 'Год', 'Pi_min',
         'Pi_min-X_min', '(Pi_min-X_min)^2', 'Население Np', 'Pi_max*Np');

type

  TChauvenet = record
    Rang: Integer;
    Subject: String;
    Year: Integer;
    Pi_min: Double;
    Deviation: Double;
    Deviation_sq: Double;
    Population: Double;
    Pi_p: Double;
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
    procedure btnLoadPopulationClick(Sender: TObject);
    procedure btnRunClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
  private
    isPopulationDataLoad: Boolean;
    Morbidity: TTableData;
    Population: TTableData;
    Chauvenet: array of TChauvenet;
    procedure ReadTableData(var table: TTableData);
    procedure SetStringGridByData(var sg: TStringGrid; const table: TTableData);
    procedure CalcChauvenet;
    procedure CalcLongTimeAverageAnnualMinimum;
    procedure AddToLog(str: String);
    procedure ClearCalculatedData;
    procedure CalcPopulationData;
  public
    { public declarations }
  end;

var
  Form1: TForm1;

implementation

{$R *.lfm}

{ TForm1 }

procedure TForm1.ReadTableData(var table: TTableData);
var
  fin: TextFile;
  FileName, s: String;
  Row: Integer;
begin
  table := TTableData.Create;
  If openDialog1.Execute Then FileName:=openDialog1.FileName;
  AssignFile(fin, FileName);
  Reset (fin);
  Row := 0;
  while not EOF(fin) do begin
    Readln(fin, s);
    if Row = 0 then table.setHeader(s)
    else table.addRecord(s);
    Inc(Row);
  end;
end;

procedure TForm1.SetStringGridByData(var sg: TStringGrid; const table: TTableData);
var
  i, j: Integer;
begin
  sg.RowCount := table.RowCount + 1;
  sg.ColCount := table.ColCount + 1;
  sg.Cells[0, 0] := table.SubjectType;
  For j := 0 to table.ColCount - 1 do
    sg.Cells[j + 1, 0] := IntToStr(table.Year[j]);
  For i := 0 to table.RowCount - 1 do
    sg.Cells[0, i + 1] := table.Subject[i];
  For i := 0 to table.RowCount - 1 do
    For j := 0 to table.ColCount - 1 do
      sg.Cells[j + 1, i + 1] := FloatToStr(table[i, j]);
  sg.Visible := True;
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
    while (k < Morbidity.ColCount - 1) and (Morbidity[i, k] = 0) do
      Inc(k);
    min := Morbidity[i, k];
    j_min := k;
    For j := k + 1 to Morbidity.ColCount - 1 do
      If (Morbidity[i, j] > 0.0) and (Morbidity[i, j] < min) Then begin
        min := Morbidity[i, j];
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

  For i := 0 to 5 do
    sgVariationSeries.Cells[i, 0] := CHAUVENET_HEADER[i];
  if Morbidity.SubjectType <> '' then
    sgVariationSeries.Cells[1, 0] := Morbidity.SubjectType;

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
    Un := (Chauvenet[Length(Chauvenet) - 1].Pi_min - X_min) / msd;
    AddToLog('U16 = ' + FloatToStr(Un));

    Ut := CHAUVENET_CRITERION[Length(Chauvenet)];
    AddToLog('Ut = ' + FloatToStr(Ut));
    if (Un >= Ut) then begin
      SetLength(Chauvenet, Length(Chauvenet) - 1);
      AddToLog('Un - аномальная величина');
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
  CalcPopulationData;
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

procedure TForm1.CalcPopulationData;
var
  i, sgIdx: Integer;
  pSum, pnSum: Double;
begin
  pSum := 0;
  pnSum := 0;
  sgIdx := sgVariationSeries.ColCount;
  sgVariationSeries.ColCount := sgIdx + 2;
  sgVariationSeries.Cells[sgIdx, 0] := 'Население Np';
  sgVariationSeries.Cells[sgIdx + 1, 0] := 'Pi_max*Np';
  For i := 0 to Length(Chauvenet) - 1 do begin
    Chauvenet[i].Population :=
                   Population.Data[Chauvenet[i].Subject, Chauvenet[i].Year];
    Chauvenet[i].Pi_p := Chauvenet[i].Pi_min * Chauvenet[i].Population;
    sgVariationSeries.Cells[sgIdx, i + 1] := FloatToStr(Chauvenet[i].Population);
    sgVariationSeries.Cells[sgIdx + 1, i + 1] := FloatToStr(Chauvenet[i].Pi_p);
    pSum := pSum + Chauvenet[i].Population;
    pnSum := pnSum + Chauvenet[i].Pi_p;
  end;
  AddToLog('Sum Pn = ' + FloatToStr(pSum));
  AddToLog('Sum Pn*Pi_min = ' + FloatToStr(pnSum));
end;

procedure TForm1.btnLoadMorbidityClick(Sender: TObject);
begin
  ReadTableData(Morbidity);
  SetStringGridByData(sgMorbidity, Morbidity);
  ClearCalculatedData;
  AddToLog('Таблица ''Стандартизованные показатели заболеваемости'' загружена');
  btnRun.Enabled := True;
end;

procedure TForm1.btnLoadPopulationClick(Sender: TObject);
begin
  ReadTableData(Population);
  SetStringGridByData(sgPopulation, Population);
  isPopulationDataLoad := True;
  //ClearCalculatedData;
  AddToLog('Таблица ''Численность населения'' загружена');
end;

procedure TForm1.btnRunClick(Sender: TObject);
begin
  CalcChauvenet();
  if isPopulationDataLoad then
    CalcLongTimeAverageAnnualMinimum;
  sgVariationSeries.Visible := True;
end;

procedure TForm1.FormCreate(Sender: TObject);
begin
  isPopulationDataLoad := False;
  mLog.text := '';
end;

end.


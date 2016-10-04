unit UMain;

{$mode objfpc}{$H+}

interface

uses
  UTable,
  Classes, SysUtils, FileUtil, Forms, Controls, Graphics, Dialogs, Grids,
  ExtCtrls, StdCtrls, ActnList, ComCtrls, Buttons, LazUTF8, ComObj, Variants, Types;

const
  CHAUVENET_CRITERION: array [5..50] of Double  = (1.680, 1.730, 1.790, 1.860,
         1.920, 1.960, 1.995, 2.030, 2.065, 2.100, 2.130, 2.160, 2.180, 2.200,
         2.220, 2.240, 2.260, 2.280, 2.295, 2.310, 2.335, 2.360, 2.368, 2.375,
         2.383, 2.390, 2.401, 2.412, 2.423, 2.434, 2.445, 2.456, 2.467, 2.478,
         2.489, 2.500, 2.508, 2.516, 2.524, 2.532, 2.540, 2.548, 2.556, 2.564,
         2.572, 2.580);

  CHAUVENET_HEADER: array [0..5] of String = ('Ранг', 'Субъект', 'Год', 'Pi_min',
         'Pi_min-X_min', '(Pi_min-X_min)^2');
  POPULATION_HEADER: array [0..5] of String = ('Ранг', 'Субъект', 'Год', 'Pi_min',
         'Население Np', 'Pi_max*Np');
  UPPER_CONFIDENCE_BOUNDS: array [0.. 5] of String = ('Субъект',
         'Среднемноголетняя численность населения', '', '', '', 'Pi_max');

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
    PopulationMA: Double;
  end;

  { TForm1 }

  TForm1 = class(TForm)
    btnLoadMorbidity: TButton;
    btnRun: TButton;
    btnLoadPopulation: TButton;
    Button1: TButton;
    eMeanValue: TEdit;
    eMSD: TEdit;
    Edit3: TEdit;
    gbStatValues: TGroupBox;
    gbInput: TGroupBox;
    gbCalculationResults: TGroupBox;
    lMeanValue: TLabel;
    lMSD: TLabel;
    Label3: TLabel;
    mLog: TMemo;
    OpenDialog1: TOpenDialog;
    pcCalculationResults: TPageControl;
    pcInput: TPageControl;
    pData: TPanel;
    pControl: TPanel;
    sgMorbidity: TStringGrid;
    sgVariationSeries: TStringGrid;
    sgPopulation: TStringGrid;
    Splitter1: TSplitter;
    sgMeanAnnualMin: TStringGrid;
    sgUpperConfidenceBound: TStringGrid;
    tsUpperConfidenceBound: TTabSheet;
    tsVariationSeries: TTabSheet;
    tsMeanAnnualMin: TTabSheet;
    tsMorbidity: TTabSheet;
    tsPopulation: TTabSheet;
    procedure btnLoadMorbidityClick(Sender: TObject);
    procedure btnLoadPopulationClick(Sender: TObject);
    procedure btnRunClick(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: boolean);
    procedure FormCreate(Sender: TObject);
    procedure TabSheet2ContextPopup(Sender: TObject; MousePos: TPoint;
      var Handled: Boolean);
  private
    isPopulationDataLoad: Boolean;
    Morbidity: TTableData;
    Population: TTableData;
    Chauvenet: array of TChauvenet;
    Excel: Variant;
    function ReadTableData(var table: TTableData): Boolean;
    procedure SetStringGridByData(var sg: TStringGrid; const table: TTableData);
    procedure CalcChauvenet;
    procedure ShowChauvenet;
    procedure CalcLongTimeAverageAnnualMinimum;
    procedure AddToLog(str: String);
    procedure ClearCalculatedData;
    procedure CalcPopulationData;
    procedure ShowPopulationData;
    procedure CalcUpperConfidenceBounds;
    procedure ShowUpperConfidenceBounds;
  public
    { public declarations }
  end;

var
  Form1: TForm1;

implementation

{$R *.lfm}

{ TForm1 }

function TForm1.ReadTableData(var table: TTableData): Boolean;
var
  fin: TextFile;
  FileName, s: String;
  Row: Integer;
begin
  Result := False;
  table := TTableData.Create;
  If openDialog1.Execute Then begin
    FileName:=openDialog1.FileName;
    AssignFile(fin, FileName);
    Reset (fin);
    Row := 0;
    while not EOF(fin) do begin
      Readln(fin, s);
      if Row = 0 then table.setHeader(String(s))
      else table.addRecord(String(s));
      Inc(Row);
    end;
    Result := True;
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
  SetLength(Chauvenet, Morbidity.RowCount);
  AddToLog('Расчет наименьших годовых показателей...');
  For i := 0 to Length(Chauvenet) - 1 do begin
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
    Chauvenet[i].Subject := Morbidity.Subject[i];
    Chauvenet[i].Pi_min := min;
    Chauvenet[i].Year := Morbidity.Year[j_min];
  end;

  AddToLog('Ранжирование показателей');
  For i := 0 to Length(Chauvenet) - 2 do
      For j := 0 to Length(Chauvenet) - 2 - i do
        If Chauvenet[j].Pi_min > Chauvenet[j + 1].Pi_min Then begin
          tmp := Chauvenet[j];
          Chauvenet[j] := Chauvenet[j + 1];
          Chauvenet[j + 1] := tmp;
        end;

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
      //ShowMessage('U1 is Bad');
    end;

  until (U1 < Ut) and (Un < Ut);

end;

procedure TForm1.ShowChauvenet;
var
  i: Integer;
begin

  sgVariationSeries.RowCount := Length(Chauvenet) + 1;
  sgVariationSeries.ColCount := Length(CHAUVENET_HEADER);

  For i := 0 to Length(CHAUVENET_HEADER) - 1 do
    sgVariationSeries.Cells[i, 0] := CHAUVENET_HEADER[i];
  if Morbidity.SubjectType <> '' then
    sgVariationSeries.Cells[1, 0] := Morbidity.SubjectType;

  For i := 0 to Length(Chauvenet) - 1 do begin
    sgVariationSeries.Cells[0, i + 1] := IntToStr(Chauvenet[i].Rang);
    sgVariationSeries.Cells[1, i + 1] := Chauvenet[i].Subject;
    sgVariationSeries.Cells[2, i + 1] := IntToStr(Chauvenet[i].Year);
    sgVariationSeries.Cells[3, i + 1] := FloatToStr(Chauvenet[i].Pi_min);
    sgVariationSeries.Cells[4, i + 1] := FloatToStr(Chauvenet[i].Deviation);
    sgVariationSeries.Cells[5, i + 1] := FloatToStr(Chauvenet[i].Deviation_sq);
  end;

  sgVariationSeries.Visible := True;

end;

procedure TForm1.CalcLongTimeAverageAnnualMinimum;
begin
  CalcPopulationData;
  ShowPopulationData;
  CalcUpperConfidenceBounds;
  ShowUpperConfidenceBounds;
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
  i: Integer;
  pSum, pnSum: Double;
begin
  pSum := 0;
  pnSum := 0;
  For i := 0 to Length(Chauvenet) - 1 do begin
    Chauvenet[i].Population :=
                   Population.Data[Chauvenet[i].Subject, Chauvenet[i].Year];
    Chauvenet[i].Pi_p := Chauvenet[i].Pi_min * Chauvenet[i].Population;
    pSum := pSum + Chauvenet[i].Population;
    pnSum := pnSum + Chauvenet[i].Pi_p;
  end;
  AddToLog('Sum n(i) = ' + FloatToStr(pSum));
  AddToLog('Sum N(i)*P(i)_min = ' + FloatToStr(pnSum));
  AddToLog('Среднемноголетний минимум равен ' + FloatToStr(pnSum / pSum));
end;

procedure TForm1.ShowPopulationData;
var
  i: Integer;
begin
  sgMeanAnnualMin.RowCount := Length(Chauvenet) + 1;
  sgMeanAnnualMin.ColCount := Length(POPULATION_HEADER);

  For i := 0 to Length(POPULATION_HEADER) - 1 do
    sgMeanAnnualMin.Cells[i, 0] := POPULATION_HEADER[i];
  if Population.SubjectType <> '' then
    sgMeanAnnualMin.Cells[1, 0] := Population.SubjectType;

  For i := 0 to Length(Chauvenet) - 1 do begin
    sgMeanAnnualMin.Cells[0, i + 1] := IntToStr(Chauvenet[i].Rang);
    sgMeanAnnualMin.Cells[1, i + 1] := Chauvenet[i].Subject;
    sgMeanAnnualMin.Cells[2, i + 1] := IntToStr(Chauvenet[i].Year);
    sgMeanAnnualMin.Cells[3, i + 1] := FloatToStr(Chauvenet[i].Pi_min);
    sgMeanAnnualMin.Cells[4, i + 1] := FloatToStr(Chauvenet[i].Population);
    sgMeanAnnualMin.Cells[5, i + 1] := FloatToStr(Chauvenet[i].Pi_p);
  end;

  sgMeanAnnualMin.Visible := True;
end;

procedure TForm1.CalcUpperConfidenceBounds;
var
  i, year: Integer;
  Sum: Double;
begin
  Sum := 0.0;
  for i := 0 to Length(Chauvenet) - 1 do begin
    for year := 0 to Population.YearNumber - 1 do
      Sum := Sum + Population.Data[Chauvenet[i].Subject, Population.Year[year]];
    ShowMessage('slag = ' +
        FloatToStr(Population.Data[Chauvenet[i].Subject, Population.Year[year]])
        + ', sum = ' + FloatToStr(Sum) + ', i = ' + IntToStr(i));
    Chauvenet[i].PopulationMA := Sum / Population.YearNumber;
  end;

end;

procedure TForm1.ShowUpperConfidenceBounds;
var
  i: Integer;
begin
  sgUpperConfidenceBound.RowCount := Length(Chauvenet) + 1;
  sgUpperConfidenceBound.ColCount := Length(UPPER_CONFIDENCE_BOUNDS);

    For i := 0 to Length(UPPER_CONFIDENCE_BOUNDS) - 1 do
      sgUpperConfidenceBound.Cells[i, 0] := UPPER_CONFIDENCE_BOUNDS[i];
    if Morbidity.SubjectType <> '' then
      sgUpperConfidenceBound.Cells[0, 0] := Population.SubjectType;

    For i := 0 to Length(Chauvenet) - 1 do begin
      sgUpperConfidenceBound.Cells[0, i + 1] := Chauvenet[i].Subject;
      sgUpperConfidenceBound.Cells[1, i + 1] := FloatToStr(Chauvenet[i].PopulationMA);
      //sgUpperConfidenceBound.Cells[2, i + 1] :=
      //sgUpperConfidenceBound.Cells[3, i + 1] := FloatToStr(Chauvenet[i].Pi_min);
      //sgUpperConfidenceBound.Cells[4, i + 1] := FloatToStr(Chauvenet[i].Population);
      //sgUpperConfidenceBound.Cells[5, i + 1] := FloatToStr(Chauvenet[i].Pi_p);
    end;

    sgUpperConfidenceBound.Visible := True;
end;

procedure TForm1.btnLoadMorbidityClick(Sender: TObject);
begin
  if ReadTableData(Morbidity) then begin
    SetStringGridByData(sgMorbidity, Morbidity);
    ClearCalculatedData;
    AddToLog('Таблица ''Стандартизованные показатели заболеваемости'' загружена');
    btnRun.Enabled := True;
  end;
end;

procedure TForm1.btnLoadPopulationClick(Sender: TObject);
begin
  if ReadTableData(Population) then begin
    SetStringGridByData(sgPopulation, Population);
    isPopulationDataLoad := True;
    //ClearCalculatedData;
    AddToLog('Таблица ''Численность населения'' загружена');
  end;
end;

procedure TForm1.btnRunClick(Sender: TObject);
begin
  CalcChauvenet();
  ShowChauvenet();
  if isPopulationDataLoad then
    CalcLongTimeAverageAnnualMinimum;
end;

procedure TForm1.Button1Click(Sender: TObject);
var
  i, j: Integer;
  Range, x: Variant;
  s: string;
begin
  If openDialog1.Execute Then begin
    Excel.Application.WorkBooks.Add(openDialog1.FileName);
  //Range:=Excel.Range[Excel.Cells[1, 1], Excel.Cells[50, 50]];
  with sgMorbidity do begin
    RowCount := 50;
    ColCount := 50;
    s := String(Excel.WorkSheets.Item['Лист1'].Cells[i, j]);
    //x := Range.Cells[1, 1];
    try
    for i:=1 to RowCount-1 do
      for j:=1 to ColCount-1 do
        Cells[i, j] := String(Excel.WorkSheets.Item['Лист1'].Cells[i, j]);
    except
      on e:Exception do
        showmessage(e.classname + ' error: ' + e.message);
    end;
    Visible := True;
  end;
  end;
end;

procedure TForm1.FormCloseQuery(Sender: TObject; var CanClose: boolean);
begin
  //try
  //  Excel.Quit;
  //except
  //end;
  //CanClose := True;
  //Excel := Unassigned;
end;

procedure TForm1.FormCreate(Sender: TObject);
begin
  isPopulationDataLoad := False;
  mLog.text := '';
  //Excel:=CreateOleObject('Excel.Application');
  //Excel.Visible := True; //После отладки можно закомментировать эту строку
  //Excel.DisplayAlerts:=False;
end;

procedure TForm1.TabSheet2ContextPopup(Sender: TObject; MousePos: TPoint;
  var Handled: Boolean);
begin

end;

end.


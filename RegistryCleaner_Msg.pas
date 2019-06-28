unit RegistryCleaner_Msg;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms, typinfo, PNGImage, StdCtrls, acPNG, Grids, AdvObj,
  BaseGrid, AdvGrid, AdvGlowButton, sLabel, ExtCtrls, AdvGroupBox, IniFiles, Dialogs, ShellAPI, ImgList, acAlphaImageList, Registry,
  AdvSplitter;


const
  BEGIN_SYSTEM_CHANGE = 100; // Тип события
  END_SYSTEM_CHANGE  = 101;
  APPLICATION_INSTALL =  0; // Тип точки восстановления
  CANCELLED_OPERATION = 13;
  MAX_DESC = 64;
  MIN_EVENT = 100;

  // Инфо о точки восстановления
  type
   PRESTOREPTINFOA = ^_RESTOREPTINFOA;
   _RESTOREPTINFOA = packed record
      dwEventType: DWORD;  // Тип события - начало или конец
      dwRestorePtType: DWORD;  // Тип контрольной точки - Установка или удалиние
      llSequenceNumber: INT64;  // Sequence Number - 0 for begin
      szDescription: array [0..64] of ansichar; // Описание- название программы или операция
  end;
  RESTOREPOINTINFO = _RESTOREPTINFOA;
  PRESTOREPOINTINFOA = ^_RESTOREPTINFOA;

  // Статус, возвращаемый точкой восстановления
  PSMGRSTATUS = ^_SMGRSTATUS;
  _SMGRSTATUS = packed record
    nStatus: DWORD; // Status returned by State Manager Process
    llSequenceNumber: INT64;  // Sequence Number for the restore point
  end;
  STATEMGRSTATUS =  _SMGRSTATUS;
  PSTATEMGRSTATUS =  ^_SMGRSTATUS;

  function SRSetRestorePointA(pRestorePtSpec: PRESTOREPOINTINFOA; pSMgrStatus: PSTATEMGRSTATUS): Bool;
     stdcall; external 'SrClient.dll' Name 'SRSetRestorePointA';



type
  TfmMsg = class(TForm)
    agbLogo: TAdvGroupBox;
    imgLogoBack: TImage;
    ThemeShapeMainCaption: TShape;
    ThemeImgMainCaption: TImage;
    albLogoText: TsLabel;
    imgWinTuning: TImage;
    ShapeBottom: TShape;
    ShapeLeft: TShape;
    ThemeImgLeft: TImage;
    ThemeImgLeftTemp: TImage;
    AlphaImgsStatus16: TsAlphaImageList;
    imgsAdditional: TsAlphaImageList;
    PanelCommon: TPanel;
    gbLog: TGroupBox;
    GridLog: TAdvStringGrid;
    PanelTop: TPanel;
    lbFixRegistryErrors: TLabel;
    gbErrorsToFixList: TGroupBox;
    lbStatusFixErrors: TLabel;
    GridFixErrorsList: TAdvStringGrid;
    SplitterRegFiles: TSplitter;
    btStart: TAdvGlowButton;
    btCancel: TAdvGlowButton;
    edLogDetails: TEdit;
    AlphaImgsSections16: TsAlphaImageList;
    procedure btOKClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure LoadPNG2Prog(dllname, resname: string; imgobj: TImage);
    procedure ApplyTheme;
    procedure LoadSections;
    procedure ThemeUpdateLeftIMG;
    procedure btStartClick(Sender: TObject);
    procedure btCancelClick(Sender: TObject);
    procedure AddLog(str: string; id: integer);
    procedure BackupRegKey(HKEYName: HKEY; KeyName, Comment: string; var ExportFile: TStringList);
    procedure BackupRegParams(HKEYName: HKEY; KeyName: string; var ExportFile: TStringList);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure GridLogClickCell(Sender: TObject; ARow, ACol: Integer);
    procedure ApplyLang;
    function CreateRestorePoint: boolean;
    procedure FormActivate(Sender: TObject);
    function GetRegKeyType(HKEYName: HKEY; KeyName, ParameterName: PChar): TRegDataType;
  private
    { Private declarations }
    IsLeftImgLoaded: Boolean; //загружена ли левая картинка в компонент ThemeImgLeftTemp - чтобы её оттуда каждый раз читать при изменении размера окна
    isReady: boolean;
    isLoaded: boolean; //Переменная нужна, чтобы не было изменений настроек при запуске программы
  public
    { Public declarations }
    ThemeName: string; //название темы оформления
  protected
    { Protected declarations }
    procedure CreateParams(var Params: TCreateParams); override;
  end;

var
  fmMsg: TfmMsg;

implementation

uses RegistryCleaner_Results, DataMod;

const
  ScreenWidth: LongInt = 1680;
  ScreenHeight: LongInt = 1050;

  REG_QWORD: Cardinal = 11;



{$R *.dfm}


{СОХРАНИТЬ ПОЗИЦИЮ ОКНА ПРИ ЗАКРЫТИИ}
procedure SaveWndPosition(FormName: TForm; KeyToSave: PChar); external 'Functions.dll';

{ВОССТАНОВИТЬ ПОЗИЦИЮ ОКНА ПРИ ЗАПУСКЕ}
procedure RestoreWndPosition(FormName: TForm; KeyToSave: PChar); external 'Functions.dll';

{ВЫВОДИМ РАЗНУЮ ИНФУ О ВЕРСИИ WINTUNING: ИНДЕКС, ГОД ОКОНЧАНИЯ ИСПОЛЬЗОВАНИЯ И ТД.}
function GetWTVerInfo(info_id: integer): integer; external 'Functions.dll';

{ЧТЕНИЕ НАСТРОЕК ПРОГРАММЫ}
function GetProgParam(paramname: PChar): PChar; external 'Functions.dll';

{ПРЕОБРАЗОВАНИЕ СТРОКИ ВИДА R,G,B В TCOLOR}
function ReadRBG(instr: PChar): TColor; external 'Functions.dll';

{ОТКРЫТЬ ОПРЕДЕЛЁННЫЙ РАЗДЕЛ СПРАВКИ}
procedure ViewHelp(page: PChar); stdcall; external 'Functions.dll';

{ПРЕОБРАЗОВАНИЕ СТРОКИ ВИДА %WinTuning_PATH% В C:\Program Files\WinTuning 7}
function ThemePathConvert(InStr, InThemeName: PChar): PChar; external 'Functions.dll';

{ЧТЕНИЕ ЯЗЫКОВОЙ СТРОКИ ИЗ ОПРЕДЕЛЁННОГО ФАЙЛА}
function ReadLangStr(FileName, Section, Caption: PChar): PChar; external 'Functions.dll';

{ЗАПУЩЕНА ЛИ 64-Х БИТНАЯ СИСТЕМА}
function IsWOW64: Boolean; external 'Functions.dll';

{УЗНАЁМ ТИП ДАННЫХ В REGISTR}
//function GetRegKeyType(HKEYName: HKEY; KeyName, ParameterName: PChar): Cardinal; external 'RegistryExtendedFunctions.dll';

{ВЫВОДИМ НАЗВАНИЕ ВЕРСИИ WINTUNING: [XP, VISTER, 7]}
function GetCapInfo(WTVerID, info_id: integer): shortstring; external 'Functions.dll';



//=========================================================
{ПРОДЕДУРА ПОЛУЧЕНИЯ ТИПА КЛЮЧА}
//---------------------------------------------------------
function TfmMsg.GetRegKeyType(HKEYName: HKEY; KeyName, ParameterName: PChar): TRegDataType;
var
  reg : tregistry;
  info : TRegDataInfo;
begin
  reg := tregistry.create;
  try
    reg.RootKey := HKEYName;
    reg.OpenKey(KeyName,true);
    reg.GetDataInfo(ParameterName, info);//получаем о ней информацию
  finally
    reg.free;
  end;
  Result := info.RegData;
end;
//=========================================================


//=========================================================
{ОТОБРАЖАЕТ ФОРМУ НА ПАНЕЛИ ЗАДАЧ}
//---------------------------------------------------------
procedure TfmMsg.CreateParams(var Params: TCreateParams);
begin
  inherited CreateParams(Params);
  with Params do
  begin
    ExStyle := ExStyle or WS_EX_APPWINDOW;
    WndParent := GetDesktopwindow;
  end;
end;
//=========================================================



//=========================================================
{ЗАГРУЗКА ИЗОБРАЖЕНИЯ В ОБЛАСТЬ НА ФОРМЕ}
//---------------------------------------------------------
procedure TfmMsg.LoadPNG2Prog(dllname, resname: string; imgobj: TImage);
var
  h: THandle;
  ResStream: TResourceStream;
  PNG: TPngImage;
begin
  h:= LoadLibrary(PChar(dllname));
  PNG := TPngImage.Create;
  ResStream := TResourceStream.Create(h, resname, rt_RCData);
  PNG.LoadFromStream(ResStream);
  imgobj.Picture.Assign(PNG);
  ResStream.free;
  PNG.Free;
end;
//=========================================================



//=========================================================
{ЗАМОСТИТЬ ЛЕВУЮ ПАНЕЛЬ КАРТИНКОЙ}
//---------------------------------------------------------
procedure TfmMsg.ThemeUpdateLeftIMG;
var
  x,y: integer; // левый верхний угол картинки
  ImgObjLeft: TImage;
begin
  if ThemeImgLeft.Picture <> nil then ThemeImgLeft.Picture := nil;
  if IsLeftImgLoaded then
  begin
    ImgObjLeft := TImage.Create(Self);
    ImgObjLeft.Picture.Assign(ThemeImgLeftTemp.Picture);
    ImgObjLeft.AutoSize := True;
    x:=0; y:=0;
    while y < ThemeImgLeft.Height do
    begin
      while x < ThemeImgLeft.Width do
      begin
        ThemeImgLeft.Canvas.Draw(x,y,ImgObjLeft.Picture.Graphic);
        x:=x+ImgObjLeft.Width;
      end;
      x:=0;
      y:=y+ImgObjLeft.Height;
    end;
    ImgObjLeft.Free;
  end;
end;
//=========================================================



//=========================================================
{ЗАПОЛНЯЕМ СУЩЕСТВУЮЩИЕ РАЗДЕЛЫ}
//---------------------------------------------------------
procedure TfmMsg.LoadSections;
var
  i,j, ErrorsSelectedCount, LastIndex, SumErr: Integer;
begin
  LastIndex := 1;
  SumErr := 0;
  for i := 0 to Length(fmDataMod.RegSections)-1 do
  begin
    ErrorsSelectedCount := 0;
    for j := 0 to Length(fmDataMod.RegErrors)-1 do
    begin
      if fmDataMod.RegErrors[j].SectionCode = i then
      begin
        if (fmDataMod.RegErrors[j].Enabled) AND (not fmDataMod.RegErrors[j].Excluded) then inc(ErrorsSelectedCount);
      end;
    end;
    if ErrorsSelectedCount>0 then
    begin
      GridFixErrorsList.AddRow;
      GridFixErrorsList.AddImageIdx(1,LastIndex,0,haBeforeText,vaCenter);
      GridFixErrorsList.GridCells[2, LastIndex] := fmDataMod.RegSections[i].Caption;
      GridFixErrorsList.AddImageIdx(2,LastIndex,i+3,haBeforeText,vaCenter);
      GridFixErrorsList.GridCells[3, LastIndex] := IntToStr(i);
      GridFixErrorsList.GridCells[4, LastIndex] := IntToStr(ErrorsSelectedCount);
      GridFixErrorsList.GridCells[5, LastIndex] := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'Waiting');
      inc(LastIndex);
    end;
    SumErr := SumErr + ErrorsSelectedCount;
  end;
  GridFixErrorsList.HideColumn(3);
  GridFixErrorsList.AutoSizeColumns(True);
  lbStatusFixErrors.Caption := stringReplace(ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'lbStatusFixErrors'), '%1', IntToStr(SumErr), [rfReplaceAll,rfIgnoreCase]);
end;
//=========================================================



//=========================================================
{ПРИМЕНЕНИЕ СКИНА}
//=========================================================
procedure TfmMsg.ApplyTheme;
var
  ThemeFileName, StrTmp, ThemeName: string;
  ThemeFile: TIniFile; //ini-файл темы оформления
begin
  //Языковые настройки
  ThemeName := GetProgParam('theme');
  //Разделы
  ThemeFileName := ExtractFilePath(paramstr(0)) + 'Themes\' + ThemeName + '\Theme.ini';
  if SysUtils.FileExists(ThemeFileName) then
  begin
    ThemeFile := TIniFile.Create(ThemeFileName);

    //Цвет окна
    Color := ReadRBG(PChar(ThemeFile.ReadString('Utilities', 'WndColor', '240,240,240')));
    //Фон логотипа
    imgLogoBack.Visible := ThemeFile.ReadBool('Utilities', 'LogoBackImageShow', True);
    StrTmp := ThemePathConvert(PChar(ThemeFile.ReadString('Utilities', 'LogoBackImagePath', '')), PChar(ThemeName));
    if SysUtils.FileExists(StrTmp) then imgLogoBack.Picture.LoadFromFile(StrTmp);

    //Левая картинка (левый фон)
    ThemeImgLeft.Visible := ThemeFile.ReadBool('Utilities', 'ImageLeftShow', False);
    StrTmp := ThemePathConvert(PChar(ThemeFile.ReadString('Utilities', 'ImageLeftPath', '')), PChar(ThemeName));
    if SysUtils.FileExists(StrTmp) then
    begin
      IsLeftImgLoaded := True;
      ThemeImgLeftTemp.Picture.LoadFromFile(StrTmp);
    end;

    //Фон заголовка утилиты
    ThemeImgMainCaption.Visible := ThemeFile.ReadBool('Utilities', 'UtilCaptionBackImageShow', True);
    StrTmp := ThemePathConvert(PChar(ThemeFile.ReadString('Utilities', 'UtilCaptionBackImagePath', '')), PChar(ThemeName));
    if SysUtils.FileExists(StrTmp) then ThemeImgMainCaption.Picture.LoadFromFile(StrTmp);
    //Цвет шрифта заголовка
    albLogoText.Font.Color := ReadRBG(PChar(ThemeFile.ReadString('Utilities', 'UtilCaptionFontColor', '53,65,79')));
    //Цвет фона загловка утилиты в случае, если НЕТ картинки
    ThemeShapeMainCaption.Brush.Color := ReadRBG(PChar(ThemeFile.ReadString('Utilities', 'UtilCaptionBackgroundColor', '243,243,243')));
    //Цвет борюда загловка утилиты
    ThemeShapeMainCaption.Brush.Color := ReadRBG(PChar(ThemeFile.ReadString('Utilities', 'UtilCaptionBackgroundColor', '243,243,243')));
    ThemeShapeMainCaption.Pen.Color := ReadRBG(PChar(ThemeFile.ReadString('Utilities', 'UtilCaptionBorderColor', '210,220,227')));
    ShapeLeft.Pen.Color := ReadRBG(PChar(ThemeFile.ReadString('Utilities', 'UtilCaptionBorderColor', '210,220,227')));
    ShapeBottom.Pen.Color := ReadRBG(PChar(ThemeFile.ReadString('Utilities', 'UtilCaptionBorderColor', '210,220,227')));
    agbLogo.BorderColor := ReadRBG(PChar(ThemeFile.ReadString('Utilities', 'UtilCaptionBorderColor', '210,220,227')));

    //Нижний бордюрчик
    ShapeBottom.Brush.Color := ReadRBG(PChar(ThemeFile.ReadString('Utilities', 'BorderBottomColor', '243,245,248')));

    ThemeFile.Free;
  end;
end;
//=========================================================



//=========================================================
{ДОБАВЛЕНИЕ СТРОКИ ЛОГА}
//---------------------------------------------------------
procedure TfmMsg.AddLog(str: string; id: integer);
var
  iconindex: integer;
begin
  GridLog.AddRow;
  GridLog.Cells[1,GridLog.RowCount-1] := DateTimeToStr(now());
  GridLog.Cells[2,GridLog.RowCount-1] := str;
  iconindex := -1;
  if id = 0 then iconindex := 0; //успешный пункт
  if id = 1 then iconindex := 1; //важный пункт
  if id = 2 then iconindex := 2; //просто пункт
  if id = 3 then iconindex := 3; //восстановленный пункт
  GridLog.AddImageIdx(1,GridLog.RowCount-1,3,haBeforeText,vaCenter);
  GridLog.AddImageIdx(2,GridLog.RowCount-1,iconindex,haBeforeText,vaCenter);
  GridLog.AutoSizeCol(0);
  GridLog.AutoSizeCol(1);
  GridLog.AutoSizeCol(2);
  GridLog.SelectRows(GridLog.RowCount-1,1);
  lbStatusFixErrors.Caption := str;
  Application.ProcessMessages;
end;
//=========================================================



//=========================================================
{ПРОДЕДУРА СОЗДАНИЯ КОПИИ КЛЮЧА РЕЕСТРА И ВСЕГО ЕГО СОДЕРЖИМОГО НА ЭКСПОРТ ДЛЯ ВОЗМОЖНОСТИ ВОССТАНОВЛЕНИЯ}
//---------------------------------------------------------
procedure TfmMsg.BackupRegKey(HKEYName: HKEY; KeyName, Comment: string; var ExportFile: TStringList);
var
  BackupReg: TRegistry;
  TrailingStr: string;
  //---------------------------------------------------------
  procedure ProcessBranch(root: string); {recursive sub-procedure}
  var
    keys: TStringList;
    i: longint;
    s: string; {longstrings are on the heap, not on the stack!}
  begin
    case HKEYName of
      HKEY_CLASSES_ROOT: s := 'HKEY_CLASSES_ROOT';
      HKEY_CURRENT_USER: s := 'HKEY_CURRENT_USER';
      HKEY_LOCAL_MACHINE: s := 'HKEY_LOCAL_MACHINE';
      HKEY_USERS: s := 'HKEY_USERS';
      HKEY_PERFORMANCE_DATA: s := 'HKEY_PERFORMANCE_DATA';
      HKEY_CURRENT_CONFIG: s := 'HKEY_CURRENT_CONFIG';
      HKEY_DYN_DATA: s := 'HKEY_DYN_DATA';
    end;
    if root[1] = '\' then TrailingStr := '' else TrailingStr := '\';
    ExportFile.Add('[' + s + TrailingStr + root + ']'); {write section name in brackets}
    BackupReg.OpenKey(root, False);
    BackupRegParams(HKEYName, root, ExportFile);
//    ExportFile.Add(''); {write blank line}
    keys := TStringList.Create;
    try
      try
        BackupReg.GetKeynames(keys); {get all sub-branches}
      finally
        BackupReg.CloseKey;
      end;
      for i := 0 to keys.Count - 1 do ProcessBranch(root + '\' + keys[i]);
    finally
      keys.Free; {this branch is ready}
    end;
  end; { ProcessBranch}
  //---------------------------------------------------------
begin
  if KeyName[Length(KeyName)] = '\' then SetLength(KeyName, Length(KeyName) - 1); {No trailing backslash}
  ExportFile.Add(';'+Comment);
  BackupReg := TRegistry.Create();
  if IsWOW64 then BackupReg.Access := KEY_WOW64_64KEY or KEY_ALL_ACCESS or KEY_WRITE or KEY_READ;
  BackupReg.RootKey := HKEYName;
  ProcessBranch(KeyName);  {Call the function that writes the branch and all subbranches}
  BackupReg.Free;
end;
//=========================================================



//=========================================================
{ПРОДЕДУРА ПРЕОБРАЗОВАНИЯ ДЕСЯТИЧНОГО ЧИСЛА В ШЕСТНАДЦАТИРИЧНОЕ - ДВЕ ЗНАЧАЩИХ ЦИФРЫ(14 -> 0e, ...)}
//---------------------------------------------------------
function InvertBytes(InStr: string): string;
var
  OutStr: string;
begin
  Result := '';
  if Length(InStr)<>4 then
  begin
    Result := InStr;
    Exit;
  end;
  InStr := LowerCase(InStr);
  OutStr := '00,00';
  OutStr[1] := InStr[3];
  OutStr[2] := InStr[4];
  OutStr[4] := InStr[1];
  OutStr[5] := InStr[2];
  Result := OutStr;
end;
//=========================================================



//=========================================================
{ПРОЦЕДУРА ДОБАВЛЕНИЯ HEX-ЗНАЧЕНИЯ (РАЗДЕЛЕНИЕ БОЛЬШОЙ СТРОКИ НА МНОГО СТРОК)}
//---------------------------------------------------------
procedure AddHexToStringList(BeforeString, HEXString: string; var ExportFile: TStringList);
const
  MAX_HEX_STR_SZ: integer = 25;
var
  i, j, BeforeSize,HEXBytes,StrlCount: integer;
  tstr: TStringList;
  TempStr: string;
  NewLineCondition: boolean;
begin
  BeforeSize := Length(BeforeString);
  tstr := TStringList.Create;
  tstr.text := stringReplace(LowerCase(HEXString), ',', #13#10, [rfReplaceAll]);
  HEXBytes := tstr.Count;
  TempStr := BeforeString;
  StrlCount := 0;
  j := 0;
  for i := 0 to HEXBytes-1 do
  begin
    TempStr := TempStr + tstr.Strings[i];
    inc(j);
    if (i >= 0) AND (i < HEXBytes-1) then TempStr := TempStr + ',';
    if StrlCount = 0 then NewLineCondition := (i > 0) AND (i < HEXBytes-1) AND ((j mod (MAX_HEX_STR_SZ - Round(BeforeSize/3)+1))=0)
    else NewLineCondition := (i > 0) AND (i < HEXBytes-1) AND ((j mod MAX_HEX_STR_SZ)=0);
    if NewLineCondition then
    begin
      TempStr := TempStr + '\';
      if StrlCount = 0 then ExportFile.Add(TempStr) else ExportFile.Add('  '+TempStr);
      TempStr := '';
      inc(StrlCount);
      j := 0;
    end;
  end;
  if StrlCount = 0 then ExportFile.Add(TempStr) else ExportFile.Add('  '+TempStr);
  tstr.Free;
end;
//=========================================================



//=========================================================
{ПРИ КЛИКЕ ПО ЛОГУ}
//---------------------------------------------------------
procedure TfmMsg.GridLogClickCell(Sender: TObject; ARow, ACol: Integer);
begin
  edLogDetails.Text := GridLog.Cells[2,ARow];
end;
//=========================================================



//=========================================================
{ПРОДЕДУРА СОЗДАНИЯ КОПИИ ПАРАМЕТРОВ ПЕРЕДАННОГО КЛЮЧА}
//---------------------------------------------------------
procedure TfmMsg.BackupRegParams(HKEYName: HKEY; KeyName: string; var ExportFile: TStringList);
var
  BackupReg: TRegistry;
  Values: TStringList;
  i,j: Integer;
  ParameterName, ValueName, ParameterToSave, ValueType, TempStr, KayNameWithoutTheTrailingSlashes: string;
  RegValueType: TRegDataType;
begin
  BackupReg := TRegistry.Create();
  if IsWOW64 then BackupReg.Access := KEY_WOW64_64KEY or KEY_ALL_ACCESS or KEY_WRITE or KEY_READ;
  BackupReg.RootKey := HKEYName;
  BackupReg.OpenKey(KeyName, True);
  Values := TStringList.Create;
  BackupReg.GetValueNames(Values);
  for i := 0 to Values.Count-1 do
  begin
    ParameterName := Values.Strings[i];
    ParameterToSave := ParameterName;
    ParameterToSave := StringReplace(ParameterToSave, '\', '\\', [rfReplaceAll]);
    ParameterToSave := StringReplace(ParameterToSave, '"', '\"', [rfReplaceAll]);
    if ParameterName = '' then ParameterToSave := '@' //значение по умолчанию
    else ParameterToSave := '"'+ParameterToSave+'"';
//    Memo2.Lines.Add('ParameterToSave4='+ParameterToSave);
    KayNameWithoutTheTrailingSlashes := KeyName;
    if KayNameWithoutTheTrailingSlashes[1] = '\' then
      KayNameWithoutTheTrailingSlashes := StringReplace(KayNameWithoutTheTrailingSlashes, '\', '', [rfIgnoreCase]);
    RegValueType := GetRegKeyType(HKEYName,PChar(KayNameWithoutTheTrailingSlashes),PChar(ParameterName));
    if RegValueType=rdString then //Строка
    begin
      ValueName := BackupReg.ReadString(ParameterName);
      ValueName := StringReplace(ValueName, '\', '\\', [rfReplaceAll, rfIgnoreCase]);
      ValueName := StringReplace(ValueName, '"', '\"', [rfReplaceAll, rfIgnoreCase]);
      ValueName := '"'+ValueName+'"';
      ValueType := '';
    end;
    if RegValueType=rdInteger then //DWORD
    begin
      ValueName := LowerCase(IntToHex(BackupReg.ReadInteger(ParameterName),8));
      ValueType := 'dword:';
    end;
    if RegValueType=rdBinary then //HEX (бинарный тип, двоичный параметр)
    begin
      ValueName := LowerCase(BackupReg.GetDataAsString(ParameterName));
      ValueType := 'hex:';
    end;
    if (RegValueType=rdExpandString) then  //Расширяемый строковой параметр
    begin
      TempStr := BackupReg.ReadString(ParameterName);
      ValueName := '';
      for j := 1 to Length(TempStr) do
      begin
        ValueName := ValueName + InvertBytes(IntToHex(Ord(TempStr[j]),4));
        if j <> Length(TempStr) then ValueName := ValueName + ',';
      end;
     ValueName := ValueName + ',00,00';
{      buf := nil;
      BufSize := 0;
      BufSize := BackupReg.GetDataSize(ParameterName);
      SetLength(buf, BufSize);
      BackupReg.ReadBinaryData(ParameterName, buf, BufSize);
      ValueName := '';
      for j := 0 to BufSize-1 do
      begin
        ValueName := ValueName + hexadecimal2D(buf[j],True);
        if j <> BufSize-1 then ValueName := ValueName + ',';
      end;}
      ValueType := 'hex(2):';
    end;
//  if (RegValueType>rdBinary) or (RegValueType=rdUnknown) then //В случае, если параметр неизвестен - значит, скорее всего, это бинарный тип данных
    if (RegValueType=rdUnknown) then //В случае, если параметр неизвестен - значит, скорее всего, это бинарный тип данных
    begin
      ValueName := BackupReg.GetDataAsString(ParameterName);
//      ValueType := 'hex('+IntToHex(RegValueType,1)+'):';
      ValueType := 'hex(4):';
    end;
{    if RegValueType=REG_LINK then //REG_LINK
    begin
      ValueName := BackupReg.GetDataAsString(ParameterName);
      ValueType := 'hex(6):';
    end;
    if RegValueType=REG_MULTI_SZ then //Мультистроковой параметр
    begin
      ValueName := BackupReg.GetDataAsString(ParameterName);
      ValueType := 'hex(7):';
    end;
    if RegValueType=REG_NONE then //Не определено
    begin
      ValueName := BackupReg.GetDataAsString(ParameterName);
      ValueType := 'hex(0):';
    end;
    if RegValueType=REG_RESOURCE_LIST then //<REG_RESOURCE_LIST (as comma-delimited list of hexadecimal values)>
    begin
      ValueName := BackupReg.GetDataAsString(ParameterName);
      ValueType := 'hex(8):';
    end;
    if RegValueType=REG_QWORD then //<QWORD value (as comma-delimited list of 8 hexadecimal values, in little endian byte order)>
    begin
      ValueName := BackupReg.GetDataAsString(ParameterName);
      ValueType := 'hex(b):';
    end;
    if RegValueType=REG_RESOURCE_REQUIREMENTS_LIST then //<REG_RESOURCE_REQUIREMENTS_LIST (as comma-delimited list of hexadecimal values)>
    begin
      ValueName := BackupReg.GetDataAsString(ParameterName);
      ValueType := 'hex(a):';
    end;}
    if StrPos(PChar(ValueType), 'hex')<>nil then
    begin
      TempStr := ParameterToSave+'='+ValueType;
      AddHexToStringList(TempStr, ValueName, ExportFile);
    end
    else ExportFile.Add(ParameterToSave+'='+ValueType+ValueName+'');
  end;

  BackupReg.Free;
end;
//=========================================================



//=========================================================
{СОЗДАНИЕ КОНТРОЛЬНОЙ ТОЧКИ ВОССТАНОВЛЕНИЯ}
//---------------------------------------------------------
function TfmMsg.CreateRestorePoint: boolean;
var
  RestorePtSpec: RESTOREPOINTINFO;
  SMgrStatus: STATEMGRSTATUS;
begin
  Result := False;
  TRY
    RestorePtSpec.dwEventType := BEGIN_SYSTEM_CHANGE;
    RestorePtSpec.dwRestorePtType := APPLICATION_INSTALL;
    RestorePtSpec.llSequenceNumber := 0;
    RestorePtSpec.szDescription := 'WinTuning Registry Cleaner Restore Point';
    if (SRSetRestorePointA(@RestorePtSpec, @SMgrStatus)) then Result := True;
  EXCEPT
  END;
end;



//=========================================================
{КНОПКА "ИСПРАВИТЬ"}
//---------------------------------------------------------
procedure TfmMsg.btStartClick(Sender: TObject);
var
  i, j, SectionIndex, OldErrorCount: integer;
  TempStr, StatusStr, StatusStr2, StatusStr3, CatFontColor: string;
  MYear,MMonth,MDay,MHour,MMinute,MSecond,MilSec: Word;
  NewStrLst: TStringList;
begin
  if not isReady then
  begin
    btStart.Enabled := False;
    StatusStr := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'LogStart');
    AddLog(StatusStr,2);
    NewStrLst := TStringList.Create;
    NewStrLst.Add('Windows Registry Editor Version 5.00');
    NewStrLst.Add('');
    NewStrLst.Add(';WinTuning: Registry Cleaner Buckup File');
    NewStrLst.Add(';2012-2018 (c) Ivan Saldikov');
    NewStrLst.Add(';'+ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'BackupFileCommentFileCreatedOn')+': '+DateTimeToStr(now()));
    NewStrLst.Add('');
    StatusStr := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'LogSystemRestorePointBegin');
    AddLog(StatusStr,2);
    if CreateRestorePoint then
    begin
      StatusStr := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'LogSystemRestorePointOK');
      AddLog(StatusStr,0);
    end
    else
    begin
      StatusStr := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'LogSystemRestorePointError');
      AddLog(StatusStr,2);
    end;
    for i := 1 to GridFixErrorsList.RowCount-1 do
    begin
      GridFixErrorsList.RemoveImageIdx(1,i);
      GridFixErrorsList.AddImageIdx(1,i,1,haBeforeText,vaCenter);
      GridFixErrorsList.GridCells[5, i] := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'LogStart');
      GridFixErrorsList.AutoSizeCol(4);
      GridFixErrorsList.Refresh;
      Application.ProcessMessages;
      StatusStr := stringReplace(ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'LogBeginSection'), '%1', GridFixErrorsList.Cells[2, i], [rfReplaceAll,rfIgnoreCase]);
      AddLog(StatusStr,2);
      SectionIndex := StrToInt(GridFixErrorsList.Cells[3, i]);
      OldErrorCount := fmDataMod.RegSections[SectionIndex].ErrorsCount;
      for j := 0 to Length(fmDataMod.RegErrors)-1 do
      begin
        if fmDataMod.RegErrors[j].SectionCode = SectionIndex then
        begin
          if (fmDataMod.RegErrors[j].Enabled) AND (not fmDataMod.RegErrors[j].Excluded) then
          begin
            AddLog(ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'LogSavingKeyBeforeProcessing'),0);
            BackupRegKey(fmDataMod.RegErrors[j].RootKey, fmDataMod.RegErrors[j].SubKey, fmDataMod.RegErrors[j].Text, NewStrLst);
            fmDataMod.RegErrorsFile.Add(fmDataMod.RegErrors[j].Text);
            fmDataMod.RegClean.RootKey := fmDataMod.RegErrors[j].RootKey;
            if fmDataMod.RegErrors[j].RootKey = HKEY_CLASSES_ROOT then TempStr := 'HKEY_CLASSES_ROOT';
            if fmDataMod.RegErrors[j].RootKey = HKEY_LOCAL_MACHINE then TempStr := 'HKEY_LOCAL_MACHINE';
            if fmDataMod.RegErrors[j].RootKey = HKEY_CURRENT_USER then TempStr := 'HKEY_CURRENT_USER';
            fmDataMod.RegClean.OpenKey(fmDataMod.RegErrors[j].SubKey, True);
            TRY
              case fmDataMod.RegErrors[j].Solution of
                regDel:
                  begin
                    if fmDataMod.RegErrors[j].Parameter = '' then
                    begin
                      fmDataMod.RegClean.DeleteKey(fmDataMod.RegErrors[j].SubKey);
                      StatusStr2 := stringReplace(ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'LogFixErrorInKey'), '%1', TempStr+fmDataMod.RegErrors[j].SubKey, [rfReplaceAll,rfIgnoreCase]);
                      StatusStr := '"'+GridFixErrorsList.Cells[2, i]+'" - '+StatusStr2+': '+ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'LogKeyDeleted');
                      AddLog(StatusStr,0);
                    end
                    else
                    begin
                      fmDataMod.RegClean.DeleteValue(fmDataMod.RegErrors[j].Parameter);
                      StatusStr3 := stringReplace(ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'LogParameterDeletedFromKey'), '%1', fmDataMod.RegErrors[j].Parameter, [rfReplaceAll,rfIgnoreCase]);
                      StatusStr := '"'+GridFixErrorsList.Cells[2, i]+'" - '+StatusStr2+': '+StatusStr3;
                      AddLog(StatusStr,0);
                    end;
                  end;
                regMod:
                  begin
                    fmDataMod.RegClean.WriteString(fmDataMod.RegErrors[j].Parameter, fmDataMod.RegErrors[j].NewValueData);
                    if fmDataMod.RegErrors[j].Parameter = '' then fmDataMod.RegErrors[j].Parameter := '('+ReadLangStr('WinTuning_Common.lng', 'Common', 'Default')+')';
                    StatusStr3 := stringReplace(ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'LogParamNewValue'), '%1', fmDataMod.RegErrors[i].Parameter, [rfReplaceAll,rfIgnoreCase]);
                    StatusStr3 := stringReplace(StatusStr3, '%2', fmDataMod.RegErrors[j].NewValueData, [rfReplaceAll,rfIgnoreCase]);
                    StatusStr := '"'+GridFixErrorsList.Cells[2, i]+'" - '+StatusStr2+': '+StatusStr3;
                    AddLog(StatusStr,0);
                  end;
              end;
              fmDataMod.RegErrors[j].Fixed := True;
              fmDataMod.RegErrors[j].Enabled := False;
              inc(fmDataMod.RegSections[SectionIndex].ErrorsCount,-1);
              inc(fmDataMod.RegErrorsFound,-1);
              inc(fmDataMod.RegErrorsFixed);
              fmRegistryCleanerResults.lbStatFixed.Caption :=
                stringReplace(ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'FixedErrorsCounterLog'), '%1', IntToStr(fmDataMod.RegErrorsFixed), [rfReplaceAll,rfIgnoreCase]);
            EXCEPT
              on E: Exception do
              begin
                StatusStr :=
                  stringReplace(ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'LogProcessingError'), '%1', TempStr+fmDataMod.RegErrors[j].SubKey, [rfReplaceAll,rfIgnoreCase]);
                AddLog(StatusStr,2);
              end;
            END;
          end;
        end;
      end;
      for j := 1 to fmRegistryCleanerResults.GridListOfSections.RowCount-1 do
      begin
        if StrToInt(fmRegistryCleanerResults.GridListOfSections.Cells[2,j])=SectionIndex then
        begin
          fmRegistryCleanerResults.GridListOfSections.Cells[1,j] :=
          StringReplace(fmRegistryCleanerResults.GridListOfSections.Cells[1,j],
                        '('+IntToStr(OldErrorCount)+')',
                        '('+IntToStr(fmDataMod.RegSections[SectionIndex].ErrorsCount)+')',
                        [rfReplaceAll,rfIgnoreCase]);
        end;
      end;
      for j := 1 to fmRegistryCleanerResults.GridCategories.RowCount-1 do
      begin
        if StrToInt(fmRegistryCleanerResults.GridCategories.Cells[1,j])=SectionIndex then
        begin
          if fmDataMod.RegSections[SectionIndex].ErrorsCount > 0 then CatFontColor := '#0500da' else CatFontColor := '#22355d';
          fmRegistryCleanerResults.GridCategories.Cells[2,j] := fmDataMod.RegSections[SectionIndex].Caption + ' <font color="'+CatFontColor+'">('+IntToStr(fmDataMod.RegSections[SectionIndex].ErrorsCount)+')</font>';
        end;
      end;
      fmRegistryCleanerResults.GridCategories.Refresh;
      GridFixErrorsList.RemoveImageIdx(1,i);
      GridFixErrorsList.AddImageIdx(1,i,2,haBeforeText,vaCenter);
      GridFixErrorsList.GridCells[5, i] := ReadLangStr('WinTuning_Common.lng', 'Common', 'Ready');
      GridFixErrorsList.Refresh;
      Application.ProcessMessages;
      StatusStr := stringReplace(ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'LogSectionDone'), '%1', GridFixErrorsList.Cells[2, i], [rfReplaceAll,rfIgnoreCase]);
      AddLog(StatusStr,0);
    end;
    fmRegistryCleanerResults.GridDetails.BeginUpdate;
    fmRegistryCleanerResults.GridDetails.RemoveRows(1, fmRegistryCleanerResults.GridDetails.RowCount-1);
    fmRegistryCleanerResults.GridDetails.EndUpdate;
    fmRegistryCleanerResults.edParamValue.Text := '';
    DecodeDate(now, MYear,MMonth,MDay);
    DecodeTime(now, MHour,MMinute,MSecond,MilSec);
    TempStr := 'WT_Backup_'+IntToStr(MYear)+'_'+IntToStr(MMonth)+'_'+IntToStr(MDay)+'_'+IntToStr(MHour)+'_'+IntToStr(MMinute)+'_'+IntToStr(MSecond);
    StatusStr2 := fmDataMod.PathToUtilityFolder+'Backups\'+TempStr+'.reg';
    StatusStr := stringReplace(ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'LogSavingReg'), '%1', StatusStr2, [rfReplaceAll,rfIgnoreCase]);
    AddLog(StatusStr,1);
    if not DirectoryExists(fmDataMod.PathToUtilityFolder) then SysUtils.ForceDirectories(fmDataMod.PathToUtilityFolder);
    if not DirectoryExists(fmDataMod.PathToUtilityFolder+'Backups') then SysUtils.ForceDirectories(fmDataMod.PathToUtilityFolder+'Backups');
    NewStrLst.SaveToFile(fmDataMod.PathToUtilityFolder+'Backups\'+TempStr+'.reg');
    NewStrLst.Free;
    fmDataMod.RegErrorsFile.SaveToFile(fmDataMod.PathToUtilityFolder+'ErrExcpt');
    fmRegistryCleanerResults.lbStatFound.Caption := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'lbStatFound')+': '+IntToStr(fmDataMod.RegErrorsFound);
    StatusStr := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'LogFixDone');
    AddLog(StatusStr,0);
    if fmRegistryCleanerResults.agbErrorsList.Visible then
    begin
      fmRegistryCleanerResults.SelectedSectionIndex := -1;
      fmRegistryCleanerResults.GridCategories.OnClickCell(Self, fmRegistryCleanerResults.GridCategories.RealRow, fmRegistryCleanerResults.GridCategories.RealCol);
    end;
    btStart.Enabled := True;
    btStart.Caption := ReadLangStr('WinTuning_Common.lng', 'Common', 'Ready');
    btCancel.Enabled := False;
    isReady := True;
  end
  else
  begin
    btCancel.Click;
  end;
end;
//=========================================================



//=========================================================
{КНОПКА "ОТМЕНА"}
//---------------------------------------------------------
procedure TfmMsg.btCancelClick(Sender: TObject);
begin
  Close;
end;
//=========================================================



//=========================================================
{КНОПКА "ЗАКРЫТЬ"}
//---------------------------------------------------------
procedure TfmMsg.btOKClick(Sender: TObject);
begin
  fmMsg.Close;
end;
//=========================================================



//=========================================================
{ПРИ ЗАКРЫТИИ ФОРМЫ}
//---------------------------------------------------------
procedure TfmMsg.FormClose(Sender: TObject; var Action: TCloseAction);
var
  RegCl: TRegistry;
begin
  RegCl := TRegistry.Create;
  RegCl.RootKey := HKEY_CURRENT_USER;
  RegCl.OpenKey('\Software\WinTuning\RegistryCleaner', True);
  TRY
    RegCl.WriteInteger('gbLog', gbLog.Height);
  EXCEPT

  END;
  RegCl.CloseKey;
  RegCl.Free;
  SaveWndPosition(fmMsg, 'fmMsg');
end;
//=========================================================


//=========================================================
{ПРИМЕНЕНИЕ ЯЗЫКА}
//--------------------------------------------------------
procedure TfmMsg.ApplyLang;
begin
  Caption :=                                    ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'fmMsg');
  lbFixRegistryErrors.Caption :=                ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'fmMsg');
  albLogoText.Caption :=                        ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'Registry Cleaner');
  gbErrorsToFixList.Caption :=                  ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'gbErrorsToFixList');
  GridFixErrorsList.ColumnHeaders.Strings[2] := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'GridFixErrorsList_2');
  GridFixErrorsList.ColumnHeaders.Strings[4] := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'GridFixErrorsList_4');
  GridFixErrorsList.ColumnHeaders.Strings[5] := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'GridFixErrorsList_5');
  gbLog.Caption :=                              ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'gbLog');
  GridLog.ColumnHeaders.Strings[1] :=           ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'GridLog_1');
  GridLog.ColumnHeaders.Strings[2] :=           ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'GridLog_2');
  btStart.Caption :=                            ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'btClean');
  btCancel.Caption :=                           ReadLangStr('WinTuning_Common.lng', 'Common', 'Cancel');
end;
//=========================================================



//=========================================================
{АКТИВАЦИЯ ФОРМЫ}
//---------------------------------------------------------
procedure TfmMsg.FormActivate(Sender: TObject);
begin
  if not isLoaded then RestoreWndPosition(fmMsg, 'fmMsg');
  isLoaded := True;
end;
//=========================================================



//=========================================================
{СОЗДАНИЕ ФОРМЫ}
//--------------------------------------------------------
procedure TfmMsg.FormCreate(Sender: TObject);
var
  i: integer;
begin
  isLoaded := False;
  // ДЕЛАЕМ ФОРМУ ОДИНАКОВОЙ ПО РАЗМЕРУ ПРИ РАЗЛИЧНЫХ РАСРЕШЕНИЯХ И РАЗМЕРАХ ШРИФТА
  scaled := True;
  Screen := TScreen.Create(nil);
  for i := componentCount - 1 downto 0 do
    with components[i] do
    begin
       if GetPropInfo(ClassInfo, 'font') <> nil then Font.Size := (ScreenWidth div screen.width) * Font.Size;
    end;

  //ЗАГРУЖАЕМ ГРАФИКУ
  LoadPNG2Prog('logo.dll', 'logo_wt_small', imgWinTuning);

  isReady := False;
  ApplyTheme;
  ApplyLang;
  LoadSections();
  ThemeUpdateLeftIMG;

//  for i := 0 to fmDataMod.AlphaImgsSections16.Count-1 do
  AlphaImgsStatus16.AddImages(AlphaImgsSections16);

  fmDataMod.RegClean.RootKey := HKEY_CURRENT_USER;
  fmDataMod.RegClean.OpenKey('\Software\WinTuning\RegistryCleaner', True);
  if fmDataMod.RegClean.ValueExists('gbLog') then gbLog.Height := fmDataMod.RegClean.ReadInteger('gbLog');
end;
//=========================================================




end.

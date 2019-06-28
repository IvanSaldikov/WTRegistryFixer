unit DataMod;

interface

uses
  SysUtils, Classes, FindFile, DateUtils, Registry, JvComponentBase, JvComputerInfoEx, SHFolder, ShlObj, ActiveX, ComObj,
  Windows, Messages, Dialogs, fspTaskbarMgr, JvThread, ImgList, Controls, acAlphaImageList, IniFiles;


const
  CSIDL_PROFILE: Integer = $5;

{ТИП ДЛЯ ОШИБОК}
type
  ArrOfStr = array of string;
  TRegSections = record
    Name: string;
    Caption: string;
    Text: string;
    Enabled: boolean;
    OrderID: integer;
    ErrorsCount: integer;
  end;
  TRegErrorType = (regKeyName, regValueName);
  TRegErrorSolution = (regDel, regMod);
  TRegError = record
    Caption: string;
    Text: string;
    SectionCode: integer;
    RootKey: HKEY;
    SubKey: string;
    Parameter: string;
    NewValueData: string;
    ErrorType: TRegErrorType;
    Solution: TRegErrorSolution;
    Enabled: Boolean;
    Excluded: Boolean;
    Fixed: Boolean;
  end;
  TRegExceptionType = (etKeyAddr, etText);
  TRegExceptions = record
    Text: string;
    ExceptionType: TRegExceptionType;
  end;


type
  PShellLinkInfoStruct = ^TShellLinkInfoStruct;

  TShellLinkInfoStruct = record
    FullPathAndNameOfLinkFile: array [0 .. MAX_PATH] of Char;
    FullPathAndNameOfFileToExecute: array [0 .. MAX_PATH] of Char;
    ParamStringsOfFileToExecute: array [0 .. MAX_PATH] of Char;
    FullPathAndNameOfWorkingDirectroy: array [0 .. MAX_PATH] of Char;
    Description: array [0 .. MAX_PATH] of Char;
    FullPathAndNameOfFileContiningIcon: array [0 .. MAX_PATH] of Char;
    IconIndex: Integer;
    HotKey: Word;
    ShowCommand: Integer;
    FindData: TWIN32FINDDATA;
  end;



type
  TfmDataMod = class(TDataModule)
    JvInfo: TJvComputerInfoEx;
    JvThreadScan: TJvThread;
    function SystemEnvRep(s: string): string;
    function CleanFileName64(filename: string): string;
    function GetDLLString(filename: string; str_number: Integer): string;
    function GetDLLResFileAddress(filename: string): string;
    function GetDLLResIndex(filename: string): Integer;
    function GetDLLStr(filename: string): string;
    procedure DataModuleCreate(Sender: TObject);
    procedure DataModuleDestroy(Sender: TObject);
    procedure JvThreadScanExecute(Sender: TObject; Params: Pointer);
    procedure JvThreadScanFinish(Sender: TObject);
    procedure JvThreadScanBegin(Sender: TObject);
    procedure FormRegSections;
    procedure ScanRegistry;
    procedure AddError(               SectionCode: Integer;
                                      HKEY_Caption: HKEY;
                                      KeyToDel, ParameterToDel: string;
                                      SolutionCode: TRegErrorSolution;
                                      RegErrorCaption, RegErrorText: string;
                                      NewValueData: string = '');
    function isAlreadyInTheList(TextIn: string): boolean;
    function FileExistsExt(filename: string): boolean;
    function isKeyFilled(RootK: HKEY; keynm: string): boolean;
    procedure FormRegExceptions;
    function IsExcluded(ErrorIndex: integer): boolean;
    function DirExistsExt(DirPath: string): boolean;
    function ExpandEnvVars(const str: string): string;
    procedure GetLinkInfo(lpShellLinkInfoStruct: PShellLinkInfoStruct);

  public
    { Public declarations }
    SystemFolder, ProgramFolder: string;
    PathToUtilityFolder: string;
    Wow64FsEnableRedirection: LongBool;
    RegErrorsFile: TStringList;

    isStop: boolean; //прервать сканирование
    is_back: boolean; //прервать сканирование
    isSecondWindowShowNeeded: boolean; //нужно ли показывать второе окно (окно результатов)

    LNGChecking, LNGCheckingKey: string;

    RegErrorsFound, RegErrorsFixed: integer; //общее кол-во найденных ошибок, общее кол-во исправленных ошибок
    ActKeyName: string; //текущий ключ
    RegClean: TRegistry;
    RegSections: array of TRegSections; //переменная для хранения всех разделов
    RegErrors: array of TRegError; //переменная для хранения ошибок
    RegExceptions: array of TRegExceptions; //переменная для хранения список исключений

  end;

var
  fmDataMod: TfmDataMod;


implementation

uses RegistryCleaner_FirstForm, RegistryCleaner_Results;


{$R *.dfm}


{ЧТЕНИЕ ЯЗЫКОВОЙ СТРОКИ ИЗ ОПРЕДЕЛЁННОГО ФАЙЛА}
function ReadLangStr(FileName, Section, Caption: PChar): PChar; external 'Functions.dll';

{ПЕРЕВОД ИЗ БАЙТОВ В НОРМАЛЬНУЮ СТРОКУ}
function BytesToStr(const i64Size: Int64): PChar; external 'Functions.dll';

{ЗАПУЩЕНА ЛИ 64-Х БИТНАЯ СИСТЕМА}
function IsWOW64: Boolean; external 'Functions.dll';



//=========================================================
{ФУНКЦИЯ ДЛЯ РАЗДЕЛЕНИЯ СТРОКИ С РАЗДЕЛИТЕЛЕМ НА МАССИВ СТРОК}
//---------------------------------------------------------
function explode(sPart, sInput: string): ArrOfStr;
begin
  while Pos(sPart, sInput) <> 0 do
  begin
    SetLength(
      Result,
      Length(Result) + 1);
    Result[Length(Result) - 1] := Copy(
      sInput,
      0,
      Pos(sPart, sInput) - 1);
    Delete(
      sInput,
      1,
      Pos(sPart, sInput));
  end;
  SetLength(
    Result,
    Length(Result) + 1);
  Result[Length(Result) - 1] := sInput;
end;
//=========================================================



//=========================================================
{СОЕДИНИТЬ ЭЛЕМЕНТЫ МАССИВА В ОДНУ СТРОКУ}
//---------------------------------------------------------
function implode(sPart: string; arrInp: ArrOfStr): string;
var
  i: Integer;
begin
  if Length(arrInp) <= 1 then Result := arrInp[0]
  else
  begin
    for i := 0 to Length(arrInp) - 2 do Result := Result + arrInp[i] + sPart;
    Result := Result + arrInp[Length(arrInp) - 1];
  end;
end;
//=========================================================



//=========================================================
{УДАЛИТЬ ЭЛЕМЕНТ ИЗ МАССИВА С ЗАДАННЫМ НОМЕРОМ}
//---------------------------------------------------------
procedure ArrDeleteElement(var anArray: ArrOfStr; const aPosition: integer);
var
   lg, j : integer;
begin
   lg := length(anArray);
   if aPosition > lg-1 then
     exit
   else if aPosition = lg-1 then begin
           Setlength(anArray, lg -1);
           exit;
        end;
   for j := aPosition to lg-2 do
     anArray[j] := anArray[j+1];
   SetLength(anArray, lg-1);
end;
//=========================================================



//=========================================================
{ПОЛУЧЕНИЕ ИНФОРМАЦИИ О ЯРЛЫКЕ}
//---------------------------------------------------------
procedure TfmDataMod.GetLinkInfo(lpShellLinkInfoStruct: PShellLinkInfoStruct);
var
  ShellLink: IShellLink;
  PersistFile: IPersistFile;
  AnObj: IUnknown;
begin
  OleInitialize(nil);
  // access to the two interfaces of the object
  AnObj := CreateComObject(CLSID_ShellLink);
  ShellLink := AnObj as IShellLink;
  PersistFile := AnObj as IPersistFile;
  // Opens the specified file and initializes an object from the file contents.
  PersistFile.Load(PWChar(WideString(lpShellLinkInfoStruct^.FullPathAndNameOfLinkFile)), 0);
  with ShellLink do
  begin
    // Retrieves the path and file name of a Shell link object.
    GetPath(lpShellLinkInfoStruct^.FullPathAndNameOfFileToExecute, SizeOf(lpShellLinkInfoStruct^.FullPathAndNameOfLinkFile), lpShellLinkInfoStruct^.FindData, SLGP_UNCPRIORITY);
    // Retrieves the description string for a Shell link object.
    GetDescription(lpShellLinkInfoStruct^.Description, SizeOf(lpShellLinkInfoStruct^.Description));
    // Retrieves the command-line arguments associated with a Shell link object.
    GetArguments(lpShellLinkInfoStruct^.ParamStringsOfFileToExecute, SizeOf(lpShellLinkInfoStruct^.ParamStringsOfFileToExecute));
    // Retrieves the name of the working directory for a Shell link object.
    GetWorkingDirectory(lpShellLinkInfoStruct^.FullPathAndNameOfWorkingDirectroy, SizeOf(lpShellLinkInfoStruct^.FullPathAndNameOfWorkingDirectroy));
    // Retrieves the location (path and index) of the icon for a Shell link object.
    GetIconLocation(lpShellLinkInfoStruct^.FullPathAndNameOfFileContiningIcon, SizeOf(lpShellLinkInfoStruct^.FullPathAndNameOfFileContiningIcon), lpShellLinkInfoStruct^.IconIndex);
    // Retrieves the hot key for a Shell link object.
    GetHotKey(lpShellLinkInfoStruct^.HotKey);
    // Retrieves the show (SW_) command for a Shell link object.
    GetShowCmd(lpShellLinkInfoStruct^.ShowCommand);
  end;
  OleUninitialize;
end;
//=========================================================



//=========================================================
{ФОРМИРОВАНИЕ КАТЕГОРИЙ}
//---------------------------------------------------------
procedure TfmDataMod.FormRegSections;
var
  i: integer;
begin
  RegClean.RootKey := HKEY_CURRENT_USER;
  RegClean.OpenKey('\Software\WinTuning\RegistryCleaner\RegSections\DefaultChecked', True);
  SetLength(RegSections, 27);

  RegSections[0].Name := 'EmptyKeysInHKCR'; //Типы файлов: совершенно пустой ключ HKCR
  RegSections[0].OrderID := 0;
  RegSections[1].Name := 'NoDefaultIconFileFound'; //Типы файлов: Не найдены файлы значков
  RegSections[1].OrderID := 1;
  RegSections[2].Name := 'EmptyContextMenuEntry'; //Типы файлов: пустой пункт в контекстном меню
  RegSections[2].OrderID := 2;
  RegSections[3].Name := 'InvalidContextMenuEntry'; //Типы файлов: нерабочий пункт в контекстном меню
  RegSections[3].OrderID := 3;
  RegSections[4].Name := 'InvalidFileTypeID'; //Типы файлов: Потеряна информация о типе файла
  RegSections[4].OrderID := 4;
  RegSections[5].Name := 'InvalidFileTypeIDHKCU'; //Типы файлов: Пустой пользовательский тип файла
  RegSections[5].OrderID := 5;
  RegSections[6].Name := 'CLSIDisNotEmpty'; //Типы файлов: Не задан CLSID для пункта меню
  RegSections[6].OrderID := 6;
  RegSections[7].Name := 'CLSIDisNotExists'; //Типы файлов: CLSID для пункта меню не существует
  RegSections[7].OrderID := 7;
  RegSections[8].Name := 'HelpFiles'; //Файлы справки
  RegSections[8].OrderID := 8;
  RegSections[9].Name := 'RecentFilesWindows'; //Недавние док-ты Windows
  RegSections[9].OrderID := 9;
  RegSections[10].Name := 'RecentProgs'; //Программы из списка "открыть с помощью"
  RegSections[10].OrderID := 10;
  RegSections[11].Name := 'RecentFilesPaint'; //Недавние документы Paint
  RegSections[11].OrderID := 11;
  RegSections[12].Name := 'RecentFilesWordPad'; //Недавние док-ты WordPad
  RegSections[12].OrderID := 12;
  RegSections[13].Name := 'RecentFilesOffice'; //Недавние док-ты Office
  RegSections[13].OrderID := 13;
  RegSections[14].Name := 'RecentFilesWinRAR'; //Недавние док-ты WinRAR
  RegSections[14].OrderID := 14;
  RegSections[15].Name := 'RecentFilesOpenSaveMRUList'; //Недавние документы Media Player Classic
  RegSections[15].OrderID := 15;
  RegSections[16].Name := 'Autorun'; //Автозапуск
  RegSections[16].OrderID := 16;
  RegSections[17].Name := 'Firewall'; //Брандмауэр
  RegSections[17].OrderID := 17;
  RegSections[18].Name := 'InstalledPrograms'; //Удаление программ
  RegSections[18].OrderID := 18;
  RegSections[19].Name := 'ApplicationPaths'; //Пути к программам
  RegSections[19].OrderID := 19;
  RegSections[20].Name := 'Fonts'; //Шрифты
  RegSections[20].OrderID := 20;
  RegSections[21].Name := 'SystemExtensions'; //Расширения програм
  RegSections[21].OrderID := 21;
  RegSections[22].Name := 'DatabaseDrivers'; //Драйверы БД
  RegSections[22].OrderID := 22;
  RegSections[23].Name := 'SharedFiles'; //Общие файлы
  RegSections[23].OrderID := 23;
  RegSections[24].Name := 'ActiveXInterfaces'; //Настройки программ
  RegSections[24].OrderID := 24;
  RegSections[25].Name := 'COMLibraries'; //Компоненты программ (ActiveX, COM)
  RegSections[25].OrderID := 25;
  RegSections[26].Name := 'Sounds'; //Звуки
  RegSections[26].OrderID := 26;
  for i := 0 to Length(RegSections)-1 do
  begin
    RegSections[i].Enabled := True;
    RegSections[i].Caption := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', PChar('Section_'+RegSections[i].Name+'_Caption'));
    RegSections[i].Text := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', PChar('Section_'+RegSections[i].Name+'_Tip'));
    if RegClean.ValueExists(IntToStr(i)) then RegSections[i].Enabled := RegClean.ReadBool(IntToStr(i));
  end;
  RegClean.RootKey := HKEY_CURRENT_USER;
  RegClean.OpenKey('\Software\WinTuning\RegistryCleaner', True);
end;
//=========================================================



//=========================================================
{ РАСШИРЕННАЯ ФУНКЦИЯ ПРОВЕРКИ СУЩЕСТВОВАНИЯ ПАПКИ (ДЕЙСТВУЕТ И НА x32 И НА x64) }
//---------------------------------------------------------
function TfmDataMod.DirExistsExt(DirPath: string): boolean;
var
  isX86: boolean;
begin
  isX86 := false;
  if AnsiPos(' (x86)', DirPath) <> 0 then
    isX86 := True;
  if IsWOW64 then
  begin
    DirPath := stringReplace(
      DirPath,
      '\system32\',
      '\Sysnative\',
      [rfReplaceAll, rfIgnoreCase]);
  end;
  if DirPath <> '' then
  begin
    DirPath := ExpandEnvVars(DirPath);
    if not isX86 then
      DirPath := stringReplace(
        DirPath,
        ' (x86)',
        '',
        [rfReplaceAll, rfIgnoreCase]);
    Result := DirectoryExists(ExcludeTrailingPathDelimiter(DirPath));
  end
  else
  begin
    Result := False;
  end;
end;
//=========================================================



//=========================================================
{ПОЛУЧЕНИЕ ГЛОБАЛЬНЫХ ПЕРЕМЕННЫХ 1-я версия}
//---------------------------------------------------------
function TfmDataMod.ExpandEnvVars(const str: string): string;
var
  BufSize: Integer; // size of expanded string
begin
  // Get required buffer size
  BufSize := ExpandEnvironmentStrings(PChar(str),nil,0);
  if BufSize > 0 then
  begin
    // Read expanded string into result string
    SetLength(Result,BufSize - 1);
    ExpandEnvironmentStrings(PChar(str),PChar(Result),BufSize);
  end
  else Result := '';  // Trying to expand empty string
end;
//=========================================================



//=========================================================
{ПОЛУЧАЕМ ГЛОБАЛЬНЫЕ ПЕРЕМЕННЫЕ 2-я версия}
//---------------------------------------------------------
function TfmDataMod.SystemEnvRep(s: string): string;
var
  buf: array [0 .. $FF] of Char;
  Size: Integer;
begin
  Size := ExpandEnvironmentStrings(PChar(s), buf, SizeOf(buf));
  Result := Copy(buf, 1, Size);
end;
//=========================================================



//=========================================================
{ЕСТЬ ЛИ В КЛЮЧЕ ХОТЬ ОДИН КЛЮЧ С КАКИМ-ЛИБО ЗНАЧЕНИЕМ (ХОТЬ ОДНИМ)}
//---------------------------------------------------------
function TfmDataMod.isKeyFilled(RootK: HKEY; keynm: string): boolean;
label
  10;
var
  RT: TRegistry;
  KeyNames, ValueNames: TStringList;
  i: Integer;
begin
  Result := false;
  RT := TRegistry.Create;
  RT.RootKey := RootK;
  KeyNames := TStringList.Create;
  ValueNames := TStringList.Create;
  if RT.OpenKeyReadOnly(keynm) then
  begin
    RT.OpenKey(keynm,false);
    RT.GetValueNames(ValueNames);
    if ValueNames.Count > 0 then
    begin
      Result := True;
      goto 10;
    end;
    RT.GetKeyNames(KeyNames);
    for i := 0 to KeyNames.Count - 1 do
    begin
      RT.OpenKey(keynm + '\' + KeyNames.Strings[i],false);
      ValueNames := TStringList.Create;
      RT.GetValueNames(ValueNames);
      if ValueNames.Count > 0 then
      begin
        Result := True;
        goto 10;
      end;
    end;
  end;
10:
  RT.Free;
  KeyNames.Free;
  ValueNames.Free;
end;
//=========================================================



//=========================================================
{ Преобразуме различные сложные шифры имен файлов в реестре в нормальный вид }
//---------------------------------------------------------
function TfmDataMod.CleanFileName64(filename: string): string;
label
  10;
var
  NewName, TempName: string;
  t: TStringList;
begin
  NewName := filename;
  if NewName = '"%1" %*' then
  begin
    NewName := paramstr(0);
    goto 10;
  end;
  if (NewName = '%1') OR (NewName = '"%1"') then
  begin
    NewName := paramstr(0);
    goto 10;
  end;
  if Trim(NewName) = '' then
  begin
    NewName := '';
    goto 10;
  end;
  t := TStringList.Create;
  t.text := stringReplace(NewName,'" ',#13#10,[rfReplaceAll]);
  if t.Count > 1 then NewName := t[0];
  t.free;

  //преобразуем переменные окружения
  if IsWOW64 then
  begin
    NewName := stringReplace(NewName, '%ProgramFiles%', '%ProgramW6432%', [rfReplaceAll, rfIgnoreCase]);
    NewName := stringReplace(NewName, '%CommonProgramFiles%', '%CommonProgramW6432%', [rfReplaceAll, rfIgnoreCase]);
  end;

  NewName := SystemEnvRep(NewName);
  // заменяем все глобальные переменные на их настоящие значения
  NewName := stringReplace(NewName,'"','', [rfReplaceAll]);
  NewName := stringReplace(NewName, ' cryptext.dll', '', [rfReplaceAll]);
  t := TStringList.Create;
  t.text := stringReplace(NewName,',',#13#10,[rfReplaceAll]);
  if t.Count > 1 then NewName := t[0];
  t.free;
  t := TStringList.Create;
  t.text := stringReplace(NewName, '\rundll32.exe ', #13#10, [rfReplaceAll]);
  if t.Count > 1 then NewName := t[1];
  t.free;
{  TempName := stringReplace(NewName, '\', NewName, []);
  if NewName = TempName then
    // значит в строке нет обратного слэша (\) - дописываем строку
    NewName := SystemFolder + '\' + NewName;}
  // убираем проценты, которые идут ПОСЛЕ имени файла и являются лишь кандикапом
  t := TStringList.Create;
  t.text := stringReplace(NewName, ' %', #13#10, [rfReplaceAll]);
  if t.Count > 1 then NewName := t[0];
  t.free;
  t := TStringList.Create;
  t.text := stringReplace(NewName,' -', #13#10, [rfReplaceAll]);
  if t.Count > 1 then NewName := t[0];
  t.free;
  t := TStringList.Create;
  t.text := stringReplace(NewName, '/', #13#10, [rfReplaceAll]);
  if t.Count > 1 then NewName := t[0];
  t.free;
  NewName := stringReplace(NewName,'\\','\',[rfReplaceAll]);
  NewName := stringReplace(NewName, '\Sysnative\System32', '\Sysnative', [rfReplaceAll, rfIgnoreCase]);
  NewName := Trim(NewName);
  if (FileExists(NewName) = false) then
  begin
    TempName := stringReplace(NewName, '\Windows\Sysnative', '\Windows', []);
    if FileExists(TempName) then NewName := TempName;
  end;
10: Result := NewName;   // Memo1.Lines.Add(NewName);
end;
//=========================================================



//=========================================================
{ВЫТАСКИВАЕМ ИЗ DLL СТРОКУ}
//---------------------------------------------------------
function TfmDataMod.GetDLLString(filename: string; str_number: Integer): string;
var
  h: THandle;
  s: array [0 .. 512] of Char;
  str: string;
  i: Integer;
begin
  Result := '';
  if FileExists(filename) then
  begin
    h := LoadLibraryEx(PWideChar(filename), 0, DONT_RESOLVE_DLL_REFERENCES);
    if h <= 0 then Result := ''
    else
    begin
      str := '';
      LoadString(h, str_number, @s, 128);
      for i := 0 to length(s) - 1 do str := str + s[i];
      Result := str;
      FreeLibrary(h);
    end;
  end;
end;
//=========================================================



//=========================================================
{ ДЛЯ ИЗВЛЕЧЕНИЯ ИЗ СТРОКИ-АДРЕСА DLL ФАЙЛА ИНФОРМАЦИИ О НОМЕРЕ РЕСУРСА }
//---------------------------------------------------------
function TfmDataMod.GetDLLResIndex(filename: string): Integer;
label
  10;
var
  t: TStringList;
  ItogIndex: Integer;
  str: string;
begin
  ItogIndex := 0;
  str := filename;
  if Trim(str) = '' then
  begin
    str := '0';
    goto 10;
  end;
  t := TStringList.Create;
  t.text := stringReplace(str, ',', #13#10, [rfReplaceAll]);
  if t.Count >= 2 then
  begin
    str := t[1];
    str := stringReplace(str,' ','',[rfReplaceAll]);
    str := stringReplace(str,'-','',[rfReplaceAll]);
  end
  else
  begin
    str := '0';
  end;
  t.free;
  t := TStringList.Create;
  t.text := stringReplace(str,';',#13#10,[rfReplaceAll]);
  if t.Count > 1 then str := t[0];
  t.free;
10:
  TRY
    ItogIndex := StrToInt(str);
  EXCEPT
    ShowMessage(filename);
  END;
  Result := ItogIndex;
end;
//=========================================================



//=========================================================
{ ДЛЯ ИЗВЛЕЧЕНИЯ ИЗ СТРОКИ-АДРЕСА DLL ФАЙЛА ИНФОРМАЦИИ О РАСПОЛОЖЕНИИ ФАЙЛА DLL }
//---------------------------------------------------------
function TfmDataMod.GetDLLResFileAddress(filename: string): string;
label
  10;
var
  t: TStringList;
  ItogFileAddress: string;
begin
  ItogFileAddress := filename;
  if Trim(ItogFileAddress) = '' then
  begin
    ItogFileAddress := '';
    goto 10;
  end;
  t := TStringList.Create;
  t.text := stringReplace(ItogFileAddress,',',#13#10,[rfReplaceAll]);
  if t.Count > 1 then ItogFileAddress := t[0];
  t.text := stringReplace(ItogFileAddress,'@',#13#10,[rfReplaceAll]);
  if t.Count >= 2 then ItogFileAddress := t[1];
  t.free;
  ItogFileAddress := CleanFileName64(ItogFileAddress);
10:
  Result := ItogFileAddress;
end;
//=========================================================



//=========================================================
{ ДЛЯ ПОЛУЧЕНИЯ СТРОКИ ВИДА "ICON" ИЗ СТРОКИ ВИДА "@%ProgramFiles(x86)%\Windows Live\Photo Gallery\regres.dll,-3077;en-us.8081.0709" }
//---------------------------------------------------------
function TfmDataMod.GetDLLStr(filename: string): string;
var
  Itog: string;
  t: TStringList;
begin
  Itog := '';
  if filename[1] = '@' then
  begin
    Itog := GetDLLString(GetDLLResFileAddress(filename), GetDLLResIndex(filename));
    t := TStringList.Create;
    t.text := stringReplace(Itog, '  ', #13#10, [rfReplaceAll]);
    if t.Count > 1 then Itog := t[0];
  end
  else Itog := filename;
  Result := Itog;
end;
//=========================================================



//=========================================================
{ ПОЛУЧЕНИЕ ПАПКИ WINDOWS }
//---------------------------------------------------------
function GetSpecialFolderPath(folder: Integer): string;
const
  SHGFP_TYPE_CURRENT = 0;
var
  path: array [0 .. MAX_PATH] of Char;
begin
  if SUCCEEDED(SHGetFolderPath(0, folder, 0, SHGFP_TYPE_CURRENT, @path[0])) then Result := path else Result := '';
end;
//=========================================================



//=========================================================
{ РАСШИРЕННАЯ ФУНКЦИЯ ПРОВЕРКИ СУЩЕСТВОВАНИЯ ФАЙЛА (ДЕЙСТВУЕТ И НА x32 И НА x64) }
//---------------------------------------------------------
function TfmDataMod.FileExistsExt(filename: string): boolean;
var
  fExt: string;
  t: TStringList;
begin
  if filename = '' then
  begin
    Result:=false;
    exit;
  end;
  if IsWOW64 then
  begin
    filename := stringReplace(filename, '\system32\', '\Sysnative\', [rfReplaceAll, rfIgnoreCase]);
    filename := stringReplace(filename, '%ProgramFiles%', '%programw6432%', [rfReplaceAll, rfIgnoreCase]);
    filename := stringReplace(filename, '%commonprogramfiles%', '%COMMONPROGRAMFILES%', [rfReplaceAll]);
  end;
  if AnsiPos('" "', filename) <> 0 then Delete(filename, AnsiPos('" "', filename), MaxInt);
  filename := Trim(filename);
  filename := stringReplace(filename,'"', '', [rfReplaceAll, rfIgnoreCase]);
  filename := stringReplace(filename,'''','', [rfReplaceAll, rfIgnoreCase]);
  fExt := ExtractFileExt(filename);
  if AnsiPos(' ', fExt) <> 0 then fExt := Copy(fExt, 1, AnsiPos(' ', fExt) - 1);
  if fExt <> '' then filename := Copy(filename, 1, AnsiPos(fExt, filename) + length(fExt) - 1);
  // we must ignore rundll32.exe
  if AnsiPos('rundll32.exe', AnsiLowerCase(filename)) <> 0 then
  begin
    Result := True;
    Exit;
  end;
  // we must ignore MTROS Style icons (потому что они читаются особым образом)
  if AnsiPos('@{', AnsiLowerCase(filename)) <> 0 then
  begin
    Result := True;
    Exit;
  end;
  // we must ignore System process
  if CompareText('system', filename) = 0 then
  begin
    Result := True;
    Exit;
  end;
  filename := ExpandEnvVars(filename);
//  if not isX86 then
//    filename := stringReplace(filename,' (x86)','',[rfReplaceAll, rfIgnoreCase]);
  if AnsiPos('/', filename) <> 0 then
    Delete(filename, AnsiPos('/', filename), MaxInt);
  filename := Trim(filename);
  Result := FileExists(filename);
  { //check if dir
    if Result=False then
    Result := DirectoryExists(filename); }
  // check if executable without extension, may be useful for env execs
  if not Result then
    Result := FileExists(filename + '.exe');
  if not Result then
    Result := FileExists(IncludeTrailingPathDelimiter(JvInfo.Folders.Windows)+filename);
  if not Result then
    Result := FileExists(IncludeTrailingPathDelimiter(JvInfo.Folders.Windows)+filename+'.exe');
  if not Result then
    Result := FileExists(IncludeTrailingPathDelimiter(JvInfo.Folders.ProgramFiles+'\Internet Explorer\')+filename);
  if not Result then
    Result := FileExists(IncludeTrailingPathDelimiter(JvInfo.Folders.System)+filename);
  if not Result then
    Result := FileExists(IncludeTrailingPathDelimiter(JvInfo.Folders.System)+filename+'.exe');
  if IsWOW64 then
  begin
    if not Result then
      Result := FileExists(IncludeTrailingPathDelimiter(JvInfo.Folders.Windows)+'Sysnative\'+filename);
    if not Result then
      Result := FileExists(IncludeTrailingPathDelimiter(JvInfo.Folders.Windows)+'Sysnative\'+filename+'.exe');
  end;

  if not Result then
  begin
    t := TStringList.Create;
    t.text := stringReplace(filename, ' ', #13#10, [rfReplaceAll]);
    if t.Count > 1 then
    begin
      Result := FileExists(t[0]);
      if not Result then Result := FileExists(t[1]);
    end;
    t.free;
  end;
  if not Result then
    Result := FileExists(IncludeTrailingPathDelimiter(GetSpecialFolderPath(CSIDL_SYSTEM)) + filename);
end;
//=========================================================



//=========================================================
{ПРИ СОЗДАНИИ МОДУЛЯ}
//---------------------------------------------------------
procedure TfmDataMod.DataModuleCreate(Sender: TObject);
begin
  RegClean := TRegistry.Create();
  if IsWOW64 then
  begin
    RegClean.Access := KEY_WOW64_64KEY or KEY_ALL_ACCESS;
    //or KEY_WRITE or KEY_READ;
    SystemFolder := GetSpecialFolderPath(CSIDL_WINDOWS) + '\Sysnative';
    ProgramFolder := ExtractFileDrive(GetSpecialFolderPath(CSIDL_WINDOWS))+'\Program Files';
  end
  else
  begin
    SystemFolder := '';
    ProgramFolder := GetSpecialFolderPath(CSIDL_PROGRAM_FILES);
  end;

  PathToUtilityFolder := JvInfo.Folders.CommonAppData+'\WinTuning\Registry Cleaner\';

  RegClean.RootKey := HKEY_CURRENT_USER;
  RegClean.OpenKey('\Software\WinTuning\RegistryCleaner', True);

  //инициализация переменных
  is_back := false;
  RegErrorsFound := 0;
  RegErrorsFixed := 0;

  FormRegSections;
  FormRegExceptions;

  //языковые "константы"
  LNGChecking :=              ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'lbStatFound');
  LNGCheckingKey :=           ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'LNGCheckingKey');
end;
//=========================================================



//=========================================================
{ПРИ ЗАКРЫТИИ МОДУЛЯ}
//---------------------------------------------------------
procedure TfmDataMod.DataModuleDestroy(Sender: TObject);
begin
  RegClean.Free;
  regErrors := nil;
  RegErrorsFile.Free;
end;
//=========================================================



//=========================================================
{ОТДЕЛЬНЫЙ ПОТОК}
//---------------------------------------------------------
procedure TfmDataMod.JvThreadScanBegin(Sender: TObject);
begin
  with fmRegistryCleanerResults do
  begin
    lbStatus.Caption := ReadLangStr('WinTuning_Common.lng', 'Common', 'Status');
    ProgressBarScanning.Visible := True; //показываем прогресс
    btClean.Enabled := False;
    btCancel.Visible := True;
    btCancel.Enabled := True;
    btBack.Enabled := False;
    btClose.Enabled := False;
    btBack.Font.Color := $00858585;
    btClose.Font.Color := $00858585;
  end;
end;
procedure TfmDataMod.JvThreadScanExecute(Sender: TObject; Params: Pointer);
begin
  ScanRegistry;
end;
procedure TfmDataMod.JvThreadScanFinish(Sender: TObject);
begin
  with fmRegistryCleanerResults do
  begin
    lbDesc.Caption := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'ResultsAfter');
    lbStatFound.Caption := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'lbStatFound')+': '+IntToStr(RegErrorsFound);
    ProgressBarScanning.Visible := False; //скрываем прогресс
    btClean.Enabled := True;
    mmFix.Enabled := True;
    btCancel.Visible := False;
    btCancel.Enabled := False;
    btBack.Enabled := True;

    mmNewScan.Enabled := True;
    btClose.Enabled := True;
    mmExit.Enabled := True;
    btBack.Font.Color := $00000000;
    btClose.Font.Color := $00000000;

    fspTaskbarMgrProgress.ProgressState := fstpsNoProgress;
    fspTaskbarMgrProgress.Active := false;

    ShowRegSections;
  end;
end;
//=========================================================



//=========================================================
{ФОРМИРОВАНИЕ СПИСКА ИСКЛЮЧЕНИЙ}
//---------------------------------------------------------
procedure TfmDataMod.FormRegExceptions;
var
  i: integer;
  ExceptionFile: TIniFile;
begin
  if FileExists(PathToUtilityFolder+'RegExceptions.ini') then
  begin
    ExceptionFile := TIniFile.Create(PathToUtilityFolder+'RegExceptions.ini');
    i := 0;
    while ExceptionFile.ValueExists('etKeyAddr', IntToStr(i)) do
    begin
      SetLength(RegExceptions, Length(RegExceptions)+1);
      RegExceptions[i].Text := ExceptionFile.ReadString('etKeyAddr', IntToStr(i), '');
      RegExceptions[i].ExceptionType := etKeyAddr;
      inc(i);
    end;
    while ExceptionFile.ValueExists('etText', IntToStr(i)) do
    begin
      SetLength(RegExceptions, Length(RegExceptions)+1);
      RegExceptions[i].Text := ExceptionFile.ReadString('etText', IntToStr(i), '');
      RegExceptions[i].ExceptionType := etText;
      inc(i);
    end;
    ExceptionFile.Free;
  end
  else
  begin
    SetLength(RegExceptions, 18);
    if not DirectoryExists(PathToUtilityFolder) then SysUtils.ForceDirectories(PathToUtilityFolder);
    ExceptionFile := TIniFile.Create(PathToUtilityFolder+'RegExceptions.ini');
    RegExceptions[0].Text := '\Titan FTP Server\Servers';
    RegExceptions[1].Text := '\Hewlett-Packard';
    RegExceptions[2].Text := '\Microsoft\Money';
    RegExceptions[3].Text := '\Software\Far';
    RegExceptions[4].Text := '\WinRAR\Profiles';
    RegExceptions[5].Text := '\Microsoft\Windows\CurrentVersion\Component Based Servicing';
    RegExceptions[6].Text := '\Outlook\Journal';
    RegExceptions[7].Text := '\Configuration\Workdir';
    RegExceptions[8].Text := '\Configuration\PrivDir';
    RegExceptions[9].Text := '\Software\Microsoft\Windows Live Contacts';
    RegExceptions[10].Text := '\SOFTWARE\KasperskyLab\protected\';
    RegExceptions[11].Text := '\SOFTWARE\Intel\Network_Services\';
    RegExceptions[12].Text := '\Control\Lsa';
    RegExceptions[13].Text := '\Microsoft\Windows NT\CurrentVersion\ProfileList';
    RegExceptions[14].Text := 'SYSTEM\Software\COMODO\Firewall Pro\Configurations\';
    for i := 0 to 14 do
    begin
      RegExceptions[i].ExceptionType := etKeyAddr;
      ExceptionFile.WriteString('etKeyAddr', IntToStr(i), RegExceptions[i].Text);
    end;
    RegExceptions[15].Text := 'Microsoft\Speech\Recognizers';
    RegExceptions[16].Text := 'THE BAT!\MAIL\SMIMERND.BIN';
    RegExceptions[17].Text := 'Bonjour';
    for i := 15 to 17 do
    begin
      RegExceptions[i].ExceptionType := etText;
      ExceptionFile.WriteString('etText', IntToStr(i), RegExceptions[i].Text);
    end;
    ExceptionFile.Free;
  end;

  RegErrorsFile := TStringList.Create;
  if FileExists(PathToUtilityFolder+'ErrExcpt') then
  begin
    RegErrorsFile.LoadFromFile(PathToUtilityFolder+'ErrExcpt');
  end;
end;
//=========================================================



//=========================================================
{ПРОВЕРКА НА ИСКЛЮЧЕНИЕ}
//---------------------------------------------------------
function TfmDataMod.IsExcluded(ErrorIndex: integer): boolean;
var
  j: integer;
  StrToCheck: string;
begin
  Result := False;
  for j := 0 to Length(RegExceptions) - 1 do
  begin
    if RegExceptions[j].ExceptionType = etKeyAddr then
    begin
      StrToCheck := '';
      if fmDataMod.RegErrors[ErrorIndex].RootKey = HKEY_CLASSES_ROOT then StrToCheck := 'HKEY_CLASSES_ROOT';
      if fmDataMod.RegErrors[ErrorIndex].RootKey = HKEY_LOCAL_MACHINE then StrToCheck := 'HKEY_LOCAL_MACHINE';
      if fmDataMod.RegErrors[ErrorIndex].RootKey = HKEY_CURRENT_USER then StrToCheck := 'HKEY_CURRENT_USER';
      StrToCheck := StrToCheck + fmDataMod.RegErrors[ErrorIndex].SubKey;
      if StrPos(PChar(StrToCheck), PChar(RegExceptions[j].Text)) <> nil then Result := True;
    end;
    if RegExceptions[j].ExceptionType = etText then
    begin
      StrToCheck := fmDataMod.RegErrors[ErrorIndex].Parameter;
      if StrPos(PChar(StrToCheck), PChar(RegExceptions[j].Text)) <> nil then Result := True;
    end;
  end;
end;
//=========================================================



//=========================================================
{ДОБАВЛЕНИЕ ОШИБКИ В СПИСОК НАЙДЕННЫХ ОШИБОК РЕЕСТРА НА КОМПЬЮТЕРЕ}
//---------------------------------------------------------
procedure TfmDataMod.AddError(    SectionCode: Integer;
                                  HKEY_Caption: HKEY;
                                  KeyToDel, ParameterToDel: string;
                                  SolutionCode: TRegErrorSolution;
                                  RegErrorCaption, RegErrorText: string;
                                  NewValueData: string = '');
begin
  if not isAlreadyInTheList(RegErrorText) then
  begin
    SetLength(RegErrors, length(RegErrors) + 1);
    RegErrors[length(regErrors) - 1].Caption := RegErrorCaption;
    RegErrors[length(regErrors) - 1].Text := RegErrorText;
    RegErrors[length(regErrors) - 1].SectionCode := SectionCode;
    RegErrors[length(regErrors) - 1].RootKey := HKEY_Caption;
    RegErrors[length(regErrors) - 1].SubKey := KeyToDel;
    RegErrors[length(regErrors) - 1].Parameter := ParameterToDel;
    RegErrors[length(regErrors) - 1].NewValueData := NewValueData;
    RegErrors[length(regErrors) - 1].Solution := SolutionCode;
    RegErrors[length(regErrors) - 1].Enabled := True;
    RegErrors[length(regErrors) - 1].Excluded := False;
    RegErrors[length(regErrors) - 1].Fixed := False;
    if not IsExcluded(length(regErrors) - 1) then
    begin
      inc(RegSections[SectionCode].ErrorsCount);
      inc(RegErrorsFound);
    end
    else
    begin
      RegErrors[length(regErrors) - 1].Enabled := False;
      RegErrors[length(regErrors) - 1].Excluded := True;
    end;
  end;
end;
//=========================================================



//=========================================================
{ИМЕЕТСЯ ЛИ ТОЧНО ТАКАЯ ЖЕ ОШИБКА В СПИСКЕ ОШИБОК}
//---------------------------------------------------------
function TfmDataMod.isAlreadyInTheList(TextIn: string): boolean;
var
  i: Integer;
begin
  Result := False;
  for i := 0 to Length(RegErrors) - 1 do
  begin
    if RegErrors[i].Text = TextIn then
    begin
      Result := True;
      Exit;
    end;
  end;
  for i := 0 to RegErrorsFile.Count - 1 do
  begin
    if RegErrorsFile.Strings[i] = TextIn then
    begin
      Result := True;
      Exit;
    end;
  end;
end;
//=========================================================





//=========================================================
{ПРОЦЕДУРА ПРЕОБРАЗОВАНИЯ REG_BINARY В СТРОКУ, КОТОРАЯ ПОДДЕРЖИВАЕТ ЮНИКОДНЫЙ ФОРМАТ (НУЖНО ДЛЯ ПАРАМЕТРОВ ИЗ HKEY_USERS\S-1-5-21-2805367877-1970857023-2369310724-1000\Software\Microsoft\Windows\CurrentVersion\Explorer\ComDlg32\OpenSavePidlMRU\)}
//---------------------------------------------------------
function InvertHexStr(HexStr: string): string;
var
  i, HEXBytes, SymbolNum: integer;
  tstr: TStringList;
  ResStr, OutStr, InStr: string;
begin
  tstr := TStringList.Create;
  tstr.text := stringReplace(HexStr, ',', #13#10, [rfReplaceAll]); //разделяем на байты
  HEXBytes := tstr.Count; //размер
  ResStr := '';
  i := 0;
  while i < HEXBytes-1 do
  begin
//    ShowMessage('tstr.Strings[i]='+tstr.Strings[i]);
    InStr := '0000';
    InStr[1] := tstr.Strings[i][1];
    InStr[2] := tstr.Strings[i][2];
//    ShowMessage('tstr.Strings[i+1]='+tstr.Strings[i+1]);
    InStr[3] := tstr.Strings[i+1][1];
    InStr[4] := tstr.Strings[i+1][2];
    inc(i,2);
    OutStr := '0000';
    //смещение
    OutStr[1] := InStr[3];
    OutStr[2] := InStr[4];
    OutStr[3] := InStr[1];
    OutStr[4] := InStr[2];
    SymbolNum := StrToInt('x'+OutStr); //получаем номер символа преобразованием Hex-строки к Integer
    ResStr := ResStr + Chr(SymbolNum);
  end;
  tstr.Free;
  Result := ResStr;
end;
//=========================================================



//=========================================================
{ФУНКЦИЯ СКАНИРОВАНИЯ РЕЕСТРА}
//---------------------------------------------------------
procedure TfmDataMod.ScanRegistry;
label
  10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21;
var
  SubKeys, SubKeys2, Values: TStringList;
  KeyName, KeyName2, ValueName, ValueName2, ValueName3, ValueName4, FileNameCap, FriendlyName, ValueData: String;
  StrCaption, StrText, sFile, TempStrVal: string;
  i,j,l,m: integer;
  RegCleanIn: TRegistry;
  BoolVal, isCDROMPath: boolean;
  BufSize: Integer;
  LinkInfo: TShellLinkInfoStruct;
  t: TStringList;
  TempHKEY: HKEY;
  RemovableDrivesArr: array of string;
  DType: Integer;
  Paths: ArrOfStr;
begin
  SetLength(regErrors, 0);
  RegCleanIn := TRegistry.Create();


  //============================================
  //File Types (ТИПЫ ФАЙЛОВ)=>0,1,2,3,4
  if RegSections[0].Enabled
  OR RegSections[1].Enabled
  OR RegSections[2].Enabled
  OR RegSections[3].Enabled
  OR RegSections[4].Enabled
  OR RegSections[6].Enabled
  OR RegSections[7].Enabled
  then
  begin
    RegCleanIn.RootKey := HKEY_CLASSES_ROOT;
    RegCleanIn.OpenKey('\', false);
    SubKeys := TStringList.Create; // создаем пустой список для подключей
    RegCleanIn.GetKeyNames(SubKeys); // Записываем список подключей}

    for i := 0 to SubKeys.Count - 1 do // для каждого подключа
    begin
      if isStop then goto 14; // выход, если что
      KeyName := SubKeys.Strings[i]; // сохраняем имя активного подключа, например ".doc"
      ActKeyName := KeyName;
      if (i mod 16) = 0 then JvThreadScan.Synchronize(fmRegistryCleanerResults.UpdateStatus);

      //0. Типы файлов: совершенно пустой ключ HKCR
      if RegSections[0].Enabled then
      begin
        RegCleanIn.RootKey := HKEY_CLASSES_ROOT;
        RegCleanIn.OpenKeyReadOnly('\' + KeyName); // заходим в активный подключ => HKCR\.doc
        ValueName := StringReplace(RegCleanIn.ReadString(''),' ', '', [rfReplaceAll,rfIgnoreCase]); // удаляем все пробелы из строки, чтобы корректно определить пустую строку
        BoolVal := RegCleanIn.ValueExists('ShellFolder'); //используется далее при проверке HKCR\.doc\shell
        Values := TStringList.Create; // создаем список для значений активного подключа
        RegCleanIn.GetValueNames(Values); // заполняем список значений для активного подключа
        if (Values.Count >= 1) then TempStrVal := Values.Strings[0];
        if (
             (not RegCleanIn.HasSubKeys)
           AND
             (
                  (Values.Count = 0)
                  OR
                  (
                      (Values.Count = 1)               //на случай, если значению по умолчанию присвоено какое-то значение, но пустое
                      AND
                      (trim(TempStrVal) = '')
                  )
             )
           AND
             (ValueName = '')
           )
        then
        begin // если пустой ключ (нет подключей и количество значений = 0)
          StrCaption := '%1';
          StrCaption := stringReplace(StrCaption, '%1', KeyName, [rfReplaceAll, rfIgnoreCase]);
          StrText := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'ErrorTip_0_0');
          StrText := stringReplace(StrText, '%1', KeyName, [rfReplaceAll, rfIgnoreCase]);
          AddError(0, HKEY_CLASSES_ROOT, '\'+KeyName, '', regDel, StrCaption, StrText); // добавляем запись в таблицу
          goto 12;
        end;
        Values.Free;

        //то же самое, но для "ключ/shell"
        RegCleanIn.RootKey := HKEY_CLASSES_ROOT;
        if RegCleanIn.OpenKey('\'+KeyName+'\shell', False) then // заходим в активный подключ => HKCR\.doc
        begin
          ValueName := StringReplace(RegCleanIn.ReadString(''),' ', '', [rfReplaceAll,rfIgnoreCase]); // удаляем все пробелы из строки, чтобы корректно определить пустую строку
          Values := TStringList.Create; // создаем список для значений активного подключа
          RegCleanIn.GetValueNames(Values); // заполняем список значений для активного подключа
          if (
               (not RegCleanIn.HasSubKeys)
             AND
               (
                    (Values.Count = 0)
                    OR
                    (
                        (Values.Count = 1)               //на случай, если значению по умолчанию присвоено какое-то значение, но пустое
                        AND
                        (trim(Values.Strings[0]) = '')
                    )
               )
             AND
               (ValueName = '')
             AND
               (not BoolVal)
             )
          then
          begin // если пустой ключ (нет подключей и количество значений = 0)
            StrCaption := '%1\shell';
            StrCaption := stringReplace(StrCaption, '%1', KeyName, [rfReplaceAll, rfIgnoreCase]);
            StrText := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'ErrorTip_0_0');
            StrText := stringReplace(StrText, '%1', KeyName+'\shell', [rfReplaceAll, rfIgnoreCase]);
            AddError(0, HKEY_CLASSES_ROOT, '\'+KeyName+'\shell', '', regDel, StrCaption, StrText); // добавляем запись в таблицу
            goto 12;
          end;
          Values.Free;
        end; //if RegCleanIn.OpenKey('\'+KeyName+'\shell', False)
      end; //if RegSections[0].Enabled
      //0 => КОНЕЦ -------------------


      TRY
      //1. Типы файлов: Не найдены файлы значков
      if RegSections[1].Enabled then
      begin
        RegCleanIn.RootKey := HKEY_CLASSES_ROOT;
        Values := TStringList.Create; // создаем список для значений активного подключа
        if RegCleanIn.OpenKeyReadOnly('\'+KeyName+'\OpenWithProgids') then // это если есть дополнительные программы для открытия этого типа файлов
        begin
          RegCleanIn.OpenKey('\'+KeyName+'\OpenWithProgids', False); // октрываем этот ключ
          RegCleanIn.GetValueNames(Values); //заполняем список значений для активного подключа
        end;
        RegCleanIn.OpenKey('\'+KeyName, False); //октрываем этот ключ
        if KeyName[1] = '.' then Values.Add(RegCleanIn.ReadString(''))
        else Values.Add(KeyName); // добавляем ещё одно значение, которое соответствует программе по умолчанию
        for l := 0 to Values.Count - 1 do // теперь для каждой строки - программы для открытия этого типа файлов ищем значения
        begin
          ValueName := Values[l]; // напр, ACDSee Pro 3.ace
          if ValueName = '' then Continue; //если пусто пропускаем ход
          if ((AnsiPos('exefile', ValueName) <> 0) or
              (AnsiPos('cmdfile', ValueName) <> 0) or
              (AnsiPos('batfile', ValueName) <> 0) or
              (AnsiPos('scrfile', ValueName) <> 0) or
              (AnsiPos('piffile', ValueName) <> 0) or
              (AnsiPos('comfile', ValueName) <> 0) or
              (AnsiPos('AppX', ValueName) <> 0)
              ) then Continue;
          RegCleanIn.OpenKey('\'+ValueName, False); // открываем этот ключ
          if RegCleanIn.ValueExists('FriendlyTypeName') then
          begin
            ValueName2 := GetDLLStr(RegCleanIn.ReadString('FriendlyTypeName'));
            if ValueName2 = '' then ValueName2 := RegCleanIn.ReadString('');
          end
          else ValueName2 := RegCleanIn.ReadString('');// записываем название этой программы, которая открывает данный тип файлов
          if ValueName2 = '' then ValueName2 := ValueName;
          if RegCleanIn.OpenKeyReadOnly('\'+ValueName+'\DefaultIcon') then // напр, "\ACDSee Pro 3.ace\DefaultIcon"
          begin
            RegCleanIn.OpenKey('\'+ValueName+'\DefaultIcon', False); // открываем этот ключ
            FileNameCap := RegCleanIn.ReadString(''); // читаем путь до файла
            if not FileExistsExt(CleanFileName64(FileNameCap)) then
            begin
              if KeyName = ValueName then StrCaption := ValueName2
              else StrCaption := ValueName2+' (*'+KeyName+')';
              StrText := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'ErrorTip_1_0');
              StrText := stringReplace(StrText,'%1',ValueName,[rfReplaceAll, rfIgnoreCase]);
              StrText := stringReplace(StrText,'%2',CleanFileName64(FileNameCap),[rfReplaceAll, rfIgnoreCase]);
              AddError(1, HKEY_CLASSES_ROOT, '\'+ValueName+'\DefaultIcon', '', regDel, StrCaption, StrText); // добавляем запись в таблицу
            end;
          end;
        end;
        Values.Free; //не забываем очищать память
      end;
      //1 => КОНЕЦ -------------------
      EXCEPT
      END;



      TRY
      //2. Типы файлов: пустой пункт в контекстном меню
      if RegSections[2].Enabled then
      begin
        RegCleanIn.RootKey := HKEY_CLASSES_ROOT;
        if RegCleanIn.OpenKeyReadOnly('\'+KeyName+'\shell') then
        begin
          RegCleanIn.OpenKey('\'+ValueName, False);
          if RegCleanIn.ValueExists('FriendlyTypeName') then
          begin
            ValueName2 := GetDLLStr(RegCleanIn.ReadString('FriendlyTypeName'));
            if ValueName2 = '' then ValueName2 := RegCleanIn.ReadString('');
          end
          else ValueName2 := RegCleanIn.ReadString(''); // записываем название этой программы, которая открывает данный тип файлов
          if ValueName2 = '' then ValueName2 := ValueName;
          RegCleanIn.OpenKey('\'+KeyName+'\shell', False);
          if RegCleanIn.HasSubKeys then
          begin
            SubKeys2 := TStringList.Create; //создаем пустой список для подключей
            RegCleanIn.GetKeyNames(SubKeys2); // Записываем список подключей
            for l := 0 to SubKeys2.Count - 1 do
            begin
              ValueName := SubKeys2.Strings[l];
              if (not RegCleanIn.OpenKeyReadOnly('\'+KeyName+'\shell\'+ValueName+'\DropTarget')) then
              begin
                if not RegCleanIn.OpenKeyReadOnly('\'+KeyName+'\shell\'+ValueName+'\command') then
                begin
                  RegCleanIn.OpenKey('\'+KeyName+'\shell\'+ValueName, False);
                  ValueName3 := RegCleanIn.ReadString('');
                  ValueName3 := stringReplace(ValueName3,'&','',[rfReplaceAll, rfIgnoreCase]);
                  if RegCleanIn.ValueExists('MUIVerb') then ValueName3 := GetDLLStr(RegCleanIn.ReadString('MUIVerb'));
                  if ValueName3 = '' then ValueName3 := ValueName;
                  StrCaption := '%1 (%2)';
                  StrCaption := stringReplace(StrCaption,'%1',ValueName2,[rfReplaceAll, rfIgnoreCase]);
                  StrCaption := stringReplace(StrCaption,'%2',ValueName3,[rfReplaceAll, rfIgnoreCase]);
                  StrText := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'ErrorTip_2_0');
                  StrText := stringReplace(StrText,'%1',KeyName,[rfReplaceAll, rfIgnoreCase]);
                  StrText := stringReplace(StrText,'%2',ValueName,[rfReplaceAll, rfIgnoreCase]);
                  StrText := stringReplace(StrText,'%3',ValueName3,[rfReplaceAll, rfIgnoreCase]);
                  AddError(2, HKEY_CLASSES_ROOT, '\'+KeyName+'\shell\'+ValueName, '', regDel, StrCaption, StrText); // добавляем запись в таблицу
                end;
              end; //if (not RegCleanIn.OpenKeyReadOnly('\'+KeyName+'\shell\'+ValueName+'\DropTarget'))
            end; //for l
            SubKeys2.Free;
          end; //if RegCleanIn.HasSubKeys
        end; //if RegCleanIn.OpenKeyReadOnly('\'+KeyName+'\shell')
      end; //if RegSections[2].Enabled
      //2 => КОНЕЦ--------------------
      EXCEPT
      END;


      TRY
      //3. Типы файлов: нерабочий пункт в контекстном меню
      if RegSections[3].Enabled then
      begin
        RegCleanIn.RootKey := HKEY_CLASSES_ROOT;
        Values := TStringList.Create; //создаем список для значений активного подключа
        if RegCleanIn.OpenKeyReadOnly('\'+KeyName+'\OpenWithProgids') then //это если есть дополнительные программы для открытия этого типа файлов
        begin
          RegCleanIn.OpenKey('\'+KeyName+'\OpenWithProgids', False); //октрываем этот ключ
          RegCleanIn.GetValueNames(Values); //заполняем список значений для активного подключа
        end;
        RegCleanIn.OpenKey('\'+KeyName, False); // октрываем ключ по умолчанию
        if KeyName[1] = '.' then Values.Add(RegCleanIn.ReadString(''))
        else Values.Add(KeyName); //добавляем ещё одно значение, которое соответствует программе по умолчанию
        for l := 0 to Values.Count - 1 do // теперь для каждой строки - программы для открытия этого типа файлов ищем значения
        begin
          ValueName := Values[l]; // напр, ACDSee Pro 3.ace
          if ((AnsiPos('exefile', ValueName) <> 0) or
              (AnsiPos('cmdfile', ValueName) <> 0) or
              (AnsiPos('batfile', ValueName) <> 0) or
              (AnsiPos('piffile', ValueName) <> 0) or
              (AnsiPos('scrfile', ValueName) <> 0) or
              (AnsiPos('comfile', ValueName) <> 0)) then Continue;
          RegCleanIn.OpenKey('\'+ValueName, False);
          if RegCleanIn.ValueExists('FriendlyTypeName') then
          begin
            ValueName2 := GetDLLStr(RegCleanIn.ReadString('FriendlyTypeName'));
            if ValueName2 = '' then ValueName2 := RegCleanIn.ReadString('');
          end
          else ValueName2 := RegCleanIn.ReadString(''); //записываем название этой программы, которая открывает данный тип файлов
          if ValueName2 = '' then ValueName2 := ValueName;
          if RegCleanIn.OpenKeyReadOnly('\'+ValueName+'\shell') then //напр, \ACDSee Pro 3.ace\shell
          begin
            RegCleanIn.OpenKey('\'+ValueName+'\shell', False); //\ACDSee Pro 3.ace\shell
            if RegCleanIn.HasSubKeys then // если есть подключи
            begin
              SubKeys2 := TStringList.Create; //создаем пустой список для подключей
              RegCleanIn.GetKeyNames(SubKeys2); // Записываем список подключей
              for m := 0 to SubKeys2.Count - 1 do //для каждого подключа (Open/Edit/Print и т.д.)
              begin
                ValueName4 := SubKeys2.Strings[m]; // читаем название подключа (Open/Edit/Print и т.д.)
                if RegCleanIn.OpenKeyReadOnly('\'+ValueName+'\shell\'+ValueName4 + '\command') then
                begin
                  RegCleanIn.OpenKey('\'+ValueName+'\shell\'+ValueName4, False);
                  ValueName3 := RegCleanIn.ReadString('');
                  if ValueName3 <> '' then
                  begin
                    if ValueName3[1] = '@' then ValueName3 := GetDLLStr(ValueName3)
                    else ValueName3 := stringReplace(ValueName3,'&','',[rfReplaceAll, rfIgnoreCase]);
                  end;
                  if RegCleanIn.ValueExists('MUIVerb') then ValueName3 := GetDLLStr(RegCleanIn.ReadString('MUIVerb'));
                  if ValueName3 = '' then ValueName3 := ValueName4;
                  RegCleanIn.OpenKey('\'+ValueName+'\shell\'+ValueName4+'\command',False);
                  FileNameCap := RegCleanIn.ReadString(''); // читаем путь до исполняемого файла
                  if not RegCleanIn.ValueExists('DelegateExecute') then
                    if not FileExistsExt(CleanFileName64(FileNameCap)) then // если файла не существует
                    begin
                      if KeyName = ValueName then StrCaption := '%1'
                      else StrCaption := '%1 (%2)';
                      StrCaption := stringReplace(StrCaption,'%1',ValueName2,[rfReplaceAll, rfIgnoreCase]);
                      StrCaption := stringReplace(StrCaption,'%2',KeyName,[rfReplaceAll, rfIgnoreCase]);
                      StrText := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'ErrorTip_3_0');
                      StrText := stringReplace(StrText,'%1',ValueName,[rfReplaceAll, rfIgnoreCase]);
                      StrText := stringReplace(StrText,'%2',ValueName4,[rfReplaceAll, rfIgnoreCase]);
                      StrText := stringReplace(StrText,'%3',ValueName3,[rfReplaceAll, rfIgnoreCase]);
                      StrText := stringReplace(StrText,'%4',CleanFileName64(RegCleanIn.ReadString('')),[rfReplaceAll, rfIgnoreCase]);
                      AddError(3, HKEY_CLASSES_ROOT, '\'+ValueName+'\shell\'+ValueName4+'\command', '', regDel, StrCaption, StrText); // добавляем запись в таблицу
                    end;
                end; //if RegCleanIn.OpenKeyReadOnly('\'+ValueName+'\shell\'+ValueName4 + '\command')
              end; //for m := 0 to SubKeys2.Count - 1 do
              SubKeys2.Free;
            end; //if RegCleanIn.HasSubKeys
          end; //if RegCleanIn.OpenKeyReadOnly('\'+ValueName+'\shell')
        end; //for l := 0 to Values.Count - 1 do
        Values.Free;
      end; //if RegSections[3].Enabled
      //3. => КОНЕЦ ------------------
      EXCEPT
      END;


      TRY
      //4. Типы файлов: Потеряна информация о типе файла
      if RegSections[4].Enabled then
      begin
        if (KeyName[1] = '.') OR (KeyName = '*') then // если первый символ - точка (или *), то это тип файла и проверяем ошибку
        begin
          RegCleanIn.RootKey := HKEY_CLASSES_ROOT;
          RegCleanIn.OpenKey('\' + KeyName, False); // заходим в активный подключ => HKCR\.doc
          ValueName := RegCleanIn.ReadString(''); // читаем значение по умолчанию
          if (ValueName <> '') then
          begin
            RegCleanIn.OpenKey('\', False);
            if not RegCleanIn.KeyExists(ValueName) then
            begin
              StrCaption := '%1 file (*%2)';
              StrCaption := stringReplace(StrCaption,'%1',UpperCase(KeyName),[rfReplaceAll, rfIgnoreCase]);
              StrCaption := stringReplace(StrCaption,'%2',KeyName,[rfReplaceAll, rfIgnoreCase]);
              StrText := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'ErrorTip_4_0');
              StrText := stringReplace(StrText,'%1',KeyName,[rfReplaceAll, rfIgnoreCase]);
              StrText := stringReplace(StrText,'%2',ValueName,[rfReplaceAll, rfIgnoreCase]);
              AddError(4, HKEY_CLASSES_ROOT, '\'+KeyName, '', regDel, StrCaption, StrText); // добавляем запись в таблицу
            end; //if not RegCleanIn.KeyExists(ValueName)
          end; //if ((RegCleanIn.HasSubKeys = false) AND (ValueName <> '')) OR ((RegCleanIn.HasSubKeys) AND (RegCleanIn.OpenKeyReadOnly('ShellEx')))
        end; //if (KeyName[1] = '.') OR (KeyName = '*')
      end; //if RegSections[4].Enabled
      //4 => КОНЕЦ------------
      EXCEPT
      END;



      TRY
      //6. Типы файлов: в ключе HKEY_CLASSES_ROOT\%1\shellex\%2 пустой CLSID------------------
      //6. Типы файлов: в ключе HKEY_CLASSES_ROOT\%1\shellex\ContextMenuHandlers\%2 пустой CLSID------------------
      if RegSections[6].Enabled then
      begin
        RegCleanIn.RootKey := HKEY_CLASSES_ROOT;
        if RegCleanIn.OpenKey('\'+KeyName+'\shellex\',false) then
        begin
          SubKeys2 := TStringList.Create; // создаем пустой список для подключей
          RegCleanIn.GetKeyNames(SubKeys2); // Записываем список подключей
          for l := 0 to SubKeys2.Count - 1 do // для каждого подключа
          begin
            if isStop then exit; // выход, если что
            KeyName2 := SubKeys2.Strings[l]; // сохраняем имя активного подключа, например "5adfasdf16a5-adf1a65df-ads5f1a" или "CopyTo"
            ActKeyName := '\'+KeyName+'\shellex\'+KeyName2;
            if (l mod 16) = 0 then JvThreadScan.Synchronize(fmRegistryCleanerResults.UpdateStatus);
            FriendlyName := '';
            RegCleanIn.RootKey := HKEY_CLASSES_ROOT;
            RegCleanIn.OpenKey('\'+KeyName, False);
            if RegCleanIn.ValueExists('FriendlyTypeName') then FriendlyName := GetDLLStr(RegCleanIn.ReadString('FriendlyTypeName'));
            if FriendlyName = '' then FriendlyName := RegCleanIn.ReadString('');
            if FriendlyName = '' then FriendlyName := KeyName;
             if (KeyName2 <> 'ContextMenuHandlers')
            AND (KeyName2 <> 'PropertySheetHandlers')
            AND (KeyName2 <> 'DragDropHandlers')
            AND (KeyName2 <> 'FolderExtensions')
            then //отедельно будет для этого подключа
            begin
              RegCleanIn.RootKey := HKEY_CLASSES_ROOT;
              if RegCleanIn.OpenKey('\'+KeyName+'\shellex\'+KeyName2,false) then
              begin
                ValueName := StringReplace(RegCleanIn.ReadString(''),' ', '', [rfReplaceAll,rfIgnoreCase]);
                if ValueName = '' then
                begin
                  if KeyName2<>'' then
                  begin
                    if KeyName2[1]='{' then //если это CLSID, то сохраняем его, если нет, то ищем дальше
                     ValueName := KeyName2
                  end;
                end;
                if (ValueName='') AND (not RegCleanIn.HasSubKeys) then
                begin
                  StrCaption := '%1';
                  StrCaption := stringReplace(StrCaption,'%1',FriendlyName,[rfReplaceAll, rfIgnoreCase]);
                  StrText := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'ErrorTip_6_0');
                  StrText := stringReplace(StrText,'%1',KeyName,[rfReplaceAll, rfIgnoreCase]);
                  StrText := stringReplace(StrText,'%2',KeyName2,[rfReplaceAll, rfIgnoreCase]);
                  AddError(6, HKEY_CLASSES_ROOT, '\'+KeyName+'\shellex\'+KeyName2, '', regDel, StrCaption, StrText);
                end; //if ValueName=''
              end;//if RegCleanIn.OpenKey('\'+KeyName+'\shellex\'+KeyName2,false)
            end; //if (KeyName2 <> 'ContextMenuHandlers') AND (KeyName2 <> 'PropertySheetHandlers')
          end; //for l := 0 to SubKeys2.Count - 1
          SubKeys2.Free;
        end;
        //далее проверяем ту же ошибку, но для подключей ContextMenuHandlers, PropertySheetHandlers, DragDropHandlers, FolderExtensions
        RegCleanIn.RootKey := HKEY_CLASSES_ROOT;
        m := 1;
        10:
        if m = 1 then ValueName3 := 'ContextMenuHandlers';
        if m = 2 then ValueName3 := 'PropertySheetHandlers';
        if m = 3 then ValueName3 := 'DragDropHandlers';
        if m = 4 then ValueName3 := 'FolderExtensions';
        if RegCleanIn.OpenKey('\'+KeyName+'\shellex\'+ValueName3+'\',false) then
        begin
          SubKeys2 := TStringList.Create; // создаем пустой список для подключей
          RegCleanIn.GetKeyNames(SubKeys2); // Записываем список подключей
          for l := 0 to SubKeys2.Count - 1 do // для каждого подключа
          begin
            if isStop then exit; // выход, если что
            KeyName2 := SubKeys2.Strings[l]; // сохраняем имя активного подключа, например "5adfasdf16a5-adf1a65df-ads5f1a" или "CopyTo"
            ActKeyName := '\'+KeyName+'\shellex\'+ValueName3+'\'+KeyName2;
            if (l mod 16) = 0 then JvThreadScan.Synchronize(fmRegistryCleanerResults.UpdateStatus);
            FriendlyName := '';
            RegCleanIn.RootKey := HKEY_CLASSES_ROOT;
            RegCleanIn.OpenKey('\'+KeyName, False);
            if RegCleanIn.ValueExists('FriendlyTypeName') then FriendlyName := GetDLLStr(RegCleanIn.ReadString('FriendlyTypeName'));
            if FriendlyName = '' then FriendlyName := RegCleanIn.ReadString('');
            if FriendlyName = '' then FriendlyName := KeyName;
            RegCleanIn.RootKey := HKEY_CLASSES_ROOT;
            if RegCleanIn.OpenKey('\'+KeyName+'\shellex\'+ValueName3+'\'+KeyName2,false) then
            begin
              ValueName := StringReplace(RegCleanIn.ReadString(''),' ', '', [rfReplaceAll,rfIgnoreCase]);
              if ValueName = '' then
              begin
                if KeyName2<>'' then
                begin
                  if KeyName2[1]='{' then //если это CLSID, то сохраняем его, если нет, то ищем дальше
                   ValueName := KeyName2
                end;
              end;
              if (ValueName='') AND (not RegCleanIn.HasSubKeys) then
              begin
                StrCaption := '%1';
                StrCaption := stringReplace(StrCaption,'%1',FriendlyName,[rfReplaceAll, rfIgnoreCase]);
                StrText := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'ErrorTip_6_0');
                StrText := stringReplace(StrText,'%1',KeyName,[rfReplaceAll, rfIgnoreCase]);
                StrText := stringReplace(StrText,'%2',ValueName3+'\'+KeyName2,[rfReplaceAll, rfIgnoreCase]);
                AddError(6, HKEY_CLASSES_ROOT, '\'+KeyName+'\shellex\'+ValueName3+'\'+KeyName2, '', regDel, StrCaption, StrText);
              end;//if ValueName=''
            end; //if RegCleanIn.OpenKey('\'+KeyName+'\shellex\ContextMenuHandlers\'+KeyName2,false)
          end; //for l := 0 to SubKeys2.Count - 1
          SubKeys2.Free;
          inc(m);
          if m <= 4 then goto 10;
       end;
      end;
      //6. => КОНЕЦ------------------
      EXCEPT
      END;



      TRY
      //7. Типы файлов: в ключе HKEY_CLASSES_ROOT\%1\shellex\ContextMenuHandlers\%2 неверный CLSID------------------
      //7. Типы файлов: в ключе HKEY_CLASSES_ROOT\%1\shellex\%2 неверный CLSID------------------
      if RegSections[7].Enabled then
      begin
        RegCleanIn.RootKey := HKEY_CLASSES_ROOT;
        if RegCleanIn.OpenKey('\'+KeyName+'\shellex\',false) then
        begin
          SubKeys2 := TStringList.Create; // создаем пустой список для подключей
          RegCleanIn.GetKeyNames(SubKeys2); // Записываем список подключей
          for l := 0 to SubKeys2.Count - 1 do // для каждого подключа
          begin
            if isStop then exit; // выход, если что
            KeyName2 := SubKeys2.Strings[l]; // сохраняем имя активного подключа, например "{5adfasdf16a5-adf1a65df-ads5f1a}" или "CopyTo"
            ActKeyName := '\'+KeyName+'\shellex\'+KeyName2;
            if (l mod 16) = 0 then JvThreadScan.Synchronize(fmRegistryCleanerResults.UpdateStatus);
            FriendlyName := '';
            RegCleanIn.RootKey := HKEY_CLASSES_ROOT;
            RegCleanIn.OpenKey('\'+KeyName, False);
            if RegCleanIn.ValueExists('FriendlyTypeName') then FriendlyName := GetDLLStr(RegCleanIn.ReadString('FriendlyTypeName'));
            if FriendlyName = '' then FriendlyName := RegCleanIn.ReadString('');
            if FriendlyName = '' then FriendlyName := KeyName;
            ValueName := '';
            //получаем CLSID для проверки
             if (KeyName2 <> 'ContextMenuHandlers')
            AND (KeyName2 <> 'PropertySheetHandlers')
            AND (KeyName2 <> 'DragDropHandlers')
            AND (KeyName2 <> 'FolderExtensions')
            then //отедельно будет для этого подключа
            begin
              RegCleanIn.RootKey := HKEY_CLASSES_ROOT;
              if RegCleanIn.OpenKey('\'+KeyName+'\shellex\'+KeyName2,false) then ValueName := StringReplace(RegCleanIn.ReadString(''),' ', '', [rfReplaceAll,rfIgnoreCase]);
              if (ValueName = '')
              OR
                (
                 (ValueName <> '')
                AND
                 (ValueName[1] <> '{')
                )
               then
              begin
                if KeyName2<>'' then
                begin
                  if KeyName2[1]='{' then //если это CLSID, то сохраняем его, если нет, то ищем дальше
                   ValueName := KeyName2
                end;
              end;
            end; //if (KeyName2 <> 'ContextMenuHandlers') AND (KeyName2 <> 'PropertySheetHandlers')...
            if ValueName <> '' then
            begin
              RegCleanIn.RootKey := HKEY_CLASSES_ROOT;
              if not RegCleanIn.OpenKeyReadOnly('\CLSID\'+ValueName) then
              begin
                StrCaption := '%1';
                StrCaption := stringReplace(StrCaption,'%1',FriendlyName,[rfReplaceAll, rfIgnoreCase]);
                StrText := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'ErrorTip_7_0');
                StrText := stringReplace(StrText,'%1',ValueName,[rfReplaceAll, rfIgnoreCase]);
                StrText := stringReplace(StrText,'%2',KeyName,[rfReplaceAll, rfIgnoreCase]);
                StrText := stringReplace(StrText,'%3',KeyName2,[rfReplaceAll, rfIgnoreCase]);
                AddError(7, HKEY_CLASSES_ROOT, '\'+KeyName+'\shellex\'+KeyName2, '', regDel, StrCaption, StrText);
              end;
            end; //if ValueName <> ''
          end; //for i := 0 to SubKeys2.Count - 1
          SubKeys2.Free;
        end;
        RegCleanIn.RootKey := HKEY_CLASSES_ROOT;
        m := 1;
        11:
        if m = 1 then ValueName3 := 'ContextMenuHandlers';
        if m = 2 then ValueName3 := 'PropertySheetHandlers';
        if m = 3 then ValueName3 := 'DragDropHandlers';
        if m = 4 then ValueName3 := 'FolderExtensions';
        if RegCleanIn.OpenKey('\'+KeyName+'\shellex\'+ValueName3+'\',false) then
        begin
          SubKeys2 := TStringList.Create; // создаем пустой список для подключей
          RegCleanIn.GetKeyNames(SubKeys2); // Записываем список подключей
          for l := 0 to SubKeys2.Count - 1 do // для каждого подключа
          begin
            if isStop then exit; // выход, если что
            KeyName2 := SubKeys2.Strings[l]; // сохраняем имя активного подключа, например "{5adfasdf16a5-adf1a65df-ads5f1a}" или "CopyTo"
            ActKeyName := '\'+KeyName+'\shellex\'+ValueName3+'\'+KeyName2;
            if (l mod 16) = 0 then JvThreadScan.Synchronize(fmRegistryCleanerResults.UpdateStatus);
            FriendlyName := '';
            RegCleanIn.RootKey := HKEY_CLASSES_ROOT;
            RegCleanIn.OpenKey('\'+KeyName, False);
            if RegCleanIn.ValueExists('FriendlyTypeName') then FriendlyName := GetDLLStr(RegCleanIn.ReadString('FriendlyTypeName'));
            if FriendlyName = '' then FriendlyName := RegCleanIn.ReadString('');
            if FriendlyName = '' then FriendlyName := KeyName;
            ValueName := '';
            RegCleanIn.RootKey := HKEY_CLASSES_ROOT;
            if RegCleanIn.OpenKey('\'+KeyName+'\shellex\'+ValueName3+'\'+KeyName2,false) then ValueName := StringReplace(RegCleanIn.ReadString(''),' ', '', [rfReplaceAll,rfIgnoreCase]);
            if (ValueName = '')
            OR
              (
               (ValueName <> '')
              AND
               (ValueName[1] <> '{')
              )
            then
            begin
              if KeyName2<>'' then
              begin
                if KeyName2[1]='{' then //если это CLSID, то сохраняем его, если нет, то ищем дальше
                 ValueName := KeyName2
              end;
            end;
            if ValueName <> '' then
            begin
              RegCleanIn.RootKey := HKEY_CLASSES_ROOT;
              if not RegCleanIn.OpenKeyReadOnly('\CLSID\'+ValueName) then
              begin
                StrCaption := '%1';
                StrCaption := stringReplace(StrCaption,'%1',FriendlyName,[rfReplaceAll, rfIgnoreCase]);
                StrText := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'ErrorTip_7_0');
                StrText := stringReplace(StrText,'%1',ValueName,[rfReplaceAll, rfIgnoreCase]);
                StrText := stringReplace(StrText,'%2',KeyName,[rfReplaceAll, rfIgnoreCase]);
                StrText := stringReplace(StrText,'%3',ValueName3+'\'+KeyName2,[rfReplaceAll, rfIgnoreCase]);
                AddError(7, HKEY_CLASSES_ROOT, '\'+KeyName+'\shellex\'+ValueName3+'\'+KeyName2, '', regDel, StrCaption, StrText);
              end;
            end; //if ValueName <> ''
          end; //for i := 0 to SubKeys2.Count - 1
          SubKeys2.Free;
        end;
      end;
      //7. => КОНЕЦ------------------
      EXCEPT
      END;


      12: RegCleanIn.CloseKey;
    end; //for i := 0 to SubKeys.Count - 1 do
  end; //  if RegSections[0].Enabled OR RegSections[1].Enabled OR RegSections[2].Enabled OR RegSections[3].Enabled OR RegSections[4].Enabled then


  TRY
  //5. Типы файлов: Потеряна информация о типе файла------------------
  if RegSections[5].Enabled then
  begin
    RegCleanIn.RootKey := HKEY_CURRENT_USER;
    RegCleanIn.OpenKey('\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts',True);
    SubKeys := TStringList.Create; // создаем пустой список для подключей
    RegCleanIn.GetKeyNames(SubKeys); // Записываем список подключей
    for i := 0 to SubKeys.Count - 1 do // для каждого подключа
    begin
      if isStop then exit; // выход, если что
      KeyName := SubKeys.Strings[i]; // сохраняем имя активного подключа, например ".doc"
      if KeyName[1] <> '.' then Continue; //пропускаем те, которые начинаются НЕ с точки
      ActKeyName := '\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\'+KeyName;
      if (i mod 16) = 0 then JvThreadScan.Synchronize(fmRegistryCleanerResults.UpdateStatus);
      if AnsiCompareText(KeyName, 'DDECache') = 0 then Continue;
      if not isKeyFilled(HKEY_CURRENT_USER, '\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\'+ KeyName) then
      begin
        StrCaption := '%1';
        StrCaption := stringReplace(StrCaption,'%1',KeyName,[rfReplaceAll, rfIgnoreCase]);
        StrText := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'ErrorTip_5_0');
        StrText := stringReplace(StrText,'%1',KeyName,[rfReplaceAll, rfIgnoreCase]);
        AddError(5, HKEY_CURRENT_USER, '\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\'+KeyName, '', regDel, StrCaption, StrText); // добавляем запись в таблицу
      end;
    end;
    SubKeys.Free;
    RegCleanIn.CloseKey;
  end;
  //5. => КОНЕЦ------------------
  EXCEPT
  END;


  TRY
  //3. (ДОП) Типы файлов: нерабочий пункт в контекстном меню (для ключа HKEY_CLASSES_ROOT\Applications\)
  if RegSections[3].Enabled then
  begin
    RegCleanIn.RootKey := HKEY_CLASSES_ROOT;
    if RegCleanIn.OpenKey('\Applications\', False) then
    begin
      Subkeys := TStringList.Create; //создаем список для значений активного подключа
      RegCleanIn.GetKeyNames(Subkeys);
      for l := 0 to Subkeys.Count - 1 do // теперь для каждой строки - программы для открытия этого типа файлов ищем значения
      begin
        if isStop then exit; // выход, если что
        KeyName := Subkeys[l]; // напр, perl.exe
        RegCleanIn.OpenKey('\Applications\'+KeyName, False);
        ActKeyName := '\Applications\'+KeyName;
        if (l mod 16) = 0 then JvThreadScan.Synchronize(fmRegistryCleanerResults.UpdateStatus);
        FriendlyName := '';
        if RegCleanIn.ValueExists('FriendlyTypeName') then FriendlyName := GetDLLStr(RegCleanIn.ReadString('FriendlyTypeName'));
        if FriendlyName = '' then FriendlyName := RegCleanIn.ReadString('');
        if FriendlyName = '' then FriendlyName := KeyName;
        if RegCleanIn.OpenKey('\Applications\'+KeyName+'\shell', False) then //напр, \ACDSee Pro 3.ace\shell
        begin
          RegCleanIn.OpenKey('\Applications\'+KeyName+'\shell', False); //\ACDSee Pro 3.ace\shell
          if RegCleanIn.HasSubKeys then // если есть подключи
          begin
            SubKeys2 := TStringList.Create; //создаем пустой список для подключей
            RegCleanIn.GetKeyNames(SubKeys2); // Записываем список подключей
            for m := 0 to SubKeys2.Count - 1 do //для каждого подключа (Open/Edit/Print и т.д.)
            begin
              ValueName4 := SubKeys2.Strings[m]; // читаем название подключа (Open/Edit/Print и т.д.)
              if RegCleanIn.OpenKey('\Applications\'+KeyName+'\shell\'+ValueName4+'\command', False) then
              begin
                RegCleanIn.OpenKey('\Applications\'+KeyName+'\shell\'+ValueName4, False);
                ValueName3 := RegCleanIn.ReadString('');
                if ValueName3 <> '' then
                begin
                  if ValueName3[1] = '@' then ValueName3 := GetDLLStr(ValueName3)
                  else ValueName3 := stringReplace(ValueName3,'&','',[rfReplaceAll, rfIgnoreCase]);
                end;
                if RegCleanIn.ValueExists('MUIVerb') then ValueName3 := GetDLLStr(RegCleanIn.ReadString('MUIVerb'));
                if ValueName3 = '' then ValueName3 := ValueName4;
                RegCleanIn.OpenKey('\Applications\'+KeyName+'\shell\'+ValueName4+'\command',False);
                FileNameCap := RegCleanIn.ReadString(''); // читаем путь до исполняемого файла
                if not RegCleanIn.ValueExists('DelegateExecute') then
                  if not FileExistsExt(CleanFileName64(FileNameCap)) then // если файла не существует
                  begin
                    StrCaption := '%1';
                    StrCaption := stringReplace(StrCaption,'%1',FriendlyName,[rfReplaceAll, rfIgnoreCase]);
                    StrText := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'ErrorTip_3_1');
                    StrText := stringReplace(StrText,'%1',KeyName,[rfReplaceAll, rfIgnoreCase]);
                    StrText := stringReplace(StrText,'%2',ValueName4,[rfReplaceAll, rfIgnoreCase]);
                    StrText := stringReplace(StrText,'%3',ValueName3,[rfReplaceAll, rfIgnoreCase]);
                    StrText := stringReplace(StrText,'%4',CleanFileName64(RegCleanIn.ReadString('')),[rfReplaceAll, rfIgnoreCase]);
                    AddError(3, HKEY_CLASSES_ROOT, '\Applications\'+KeyName+'\shell\'+ValueName4+'\command', '', regDel, StrCaption, StrText); // добавляем запись в таблицу
                  end;
               end; //if RegCleanIn.OpenKeyReadOnly('\'+KeyName+'\shell\'+ValueName4 + '\command')
            end; //for m := 0 to SubKeys2.Count - 1 do
            SubKeys2.Free;
          end; //if RegCleanIn.HasSubKeys
        end; //if RegCleanIn.OpenKeyReadOnly('\'+KeyName+'\shell')
      end; //for l := 0 to Subkeys.Count - 1 do
      Subkeys.Free;
    end; //if RegCleanIn.OpenKey('\Applications\', False)
  end;
  //3. (ДОП) => КОНЕЦ------------------
  EXCEPT
  END;



  TRY
  //4. (ДОП) Типы файлов: Потеряна информация о типе файла
  if RegSections[4].Enabled then
  begin
    RegCleanIn.RootKey := HKEY_CLASSES_ROOT;
    if RegCleanIn.OpenKey('\SystemFileAssociations\', False) then
    begin
      Subkeys := TStringList.Create; //создаем список для значений активного подключа
      RegCleanIn.GetKeyNames(Subkeys);
      for l := 0 to Subkeys.Count - 1 do // теперь для каждой строки - программы для открытия этого типа файлов ищем значения
      begin
        if isStop then exit; // выход, если что
        KeyName := Subkeys[l]; // напр, perl.exe
        if (KeyName[1] = '.') then // если первый символ - точка (или *), то это тип файла и проверяем ошибку
        begin
          RegCleanIn.RootKey := HKEY_CLASSES_ROOT;
          RegCleanIn.OpenKey('\SystemFileAssociations\'+KeyName, False); // заходим в активный подключ => HKCR\.doc
          ActKeyName := '\SystemFileAssociations\'+KeyName;
          if (l mod 16) = 0 then JvThreadScan.Synchronize(fmRegistryCleanerResults.UpdateStatus);
          ValueName := StringReplace(RegCleanIn.ReadString(''),' ', '', [rfReplaceAll,rfIgnoreCase]); // удаляем все пробелы из строки, чтобы корректно определить пустую строку
          Values := TStringList.Create; // создаем список для значений активного подключа
          RegCleanIn.GetValueNames(Values); // заполняем список значений для активного подключа
          if (Values.Count >= 1) then TempStrVal := Values.Strings[0];
          if (
               (not RegCleanIn.HasSubKeys)
             AND
               (
                    (Values.Count = 0)
                    OR
                    (
                        (Values.Count = 1)               //на случай, если значению по умолчанию присвоено какое-то значение, но пустое
                        AND
                        (trim(TempStrVal) = '')
                    )
               )
             AND
               (ValueName = '')
             )
          then
          begin // если пустой ключ (нет подключей и количество значений = 0)
              StrCaption := '%1';
              StrCaption := stringReplace(StrCaption,'%1',KeyName,[rfReplaceAll, rfIgnoreCase]);
              StrText := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'ErrorTip_4_1');
              StrText := stringReplace(StrText,'%1',KeyName,[rfReplaceAll, rfIgnoreCase]);
              StrText := stringReplace(StrText,'%2',ValueName,[rfReplaceAll, rfIgnoreCase]);
              AddError(4, HKEY_CLASSES_ROOT, '\SystemFileAssociations\'+KeyName, '', regDel, StrCaption, StrText); // добавляем запись в таблицу
          end;
          Values.Free;
        end; //if (KeyName[1] = '.')
      end; //for l := 0 to Subkeys.Count - 1   
      Subkeys.Free;
    end; //if RegCleanIn.OpenKey('\SystemFileAssociations\', False)
  end; //if RegSections[4].Enabled
  //4 (ДОП) => КОНЕЦ------------
  EXCEPT
  END;


  TRY
  //2. (ДОП) Типы файлов: пустой пункт в контекстном меню (для ключа HKEY_CLASSES_ROOT\Applications\)
  if RegSections[2].Enabled then
  begin
    RegCleanIn.RootKey := HKEY_CLASSES_ROOT;
    if RegCleanIn.OpenKey('\Applications\', False) then
    begin
      Subkeys := TStringList.Create; //создаем список для значений активного подключа
      RegCleanIn.GetKeyNames(Subkeys);
      for l := 0 to Subkeys.Count - 1 do // теперь для каждой строки - программы для открытия этого типа файлов ищем значения
      begin
        if isStop then exit; // выход, если что
        KeyName := Subkeys[l]; // напр, perl.exe
        ActKeyName := '\Applications\'+KeyName;
        if (l mod 16) = 0 then JvThreadScan.Synchronize(fmRegistryCleanerResults.UpdateStatus);
        if RegCleanIn.OpenKeyReadOnly('\Applications\'+KeyName+'\shell') then
        begin
          RegCleanIn.OpenKey('\Applications\'+ValueName, False);
          if RegCleanIn.ValueExists('FriendlyTypeName') then
          begin
            ValueName2 := GetDLLStr(RegCleanIn.ReadString('FriendlyTypeName'));
            if ValueName2 = '' then ValueName2 := RegCleanIn.ReadString('');
          end
          else ValueName2 := RegCleanIn.ReadString(''); // записываем название этой программы, которая открывает данный тип файлов
          if ValueName2 = '' then ValueName2 := ValueName;
          RegCleanIn.OpenKey('\Applications\'+KeyName+'\shell', False);
          if RegCleanIn.HasSubKeys then
          begin
            SubKeys2 := TStringList.Create; //создаем пустой список для подключей
            RegCleanIn.GetKeyNames(SubKeys2); // Записываем список подключей
            for m := 0 to SubKeys2.Count - 1 do
            begin
              ValueName := SubKeys2.Strings[m];
              if (not RegCleanIn.OpenKeyReadOnly('\Applications\'+KeyName+'\shell\'+ValueName+'\DropTarget')) then
              begin
                if not RegCleanIn.OpenKeyReadOnly('\Applications\'+KeyName+'\shell\'+ValueName+'\command') then
                begin
                  RegCleanIn.OpenKey('\Applications\'+KeyName+'\shell\'+ValueName, False);
                  ValueName3 := RegCleanIn.ReadString('');
                  ValueName3 := stringReplace(ValueName3,'&','',[rfReplaceAll, rfIgnoreCase]);
                  if RegCleanIn.ValueExists('MUIVerb') then ValueName3 := GetDLLStr(RegCleanIn.ReadString('MUIVerb'));
                  if ValueName3 = '' then ValueName3 := ValueName;
                  StrCaption := '%1 (%2)';
                  StrCaption := stringReplace(StrCaption,'%1',KeyName,[rfReplaceAll, rfIgnoreCase]);
                  StrCaption := stringReplace(StrCaption,'%2',ValueName2,[rfReplaceAll, rfIgnoreCase]);
                  StrText := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'ErrorTip_2_1');
                  StrText := stringReplace(StrText,'%1',KeyName,[rfReplaceAll, rfIgnoreCase]);
                  StrText := stringReplace(StrText,'%2',ValueName,[rfReplaceAll, rfIgnoreCase]);
                  StrText := stringReplace(StrText,'%3',ValueName3,[rfReplaceAll, rfIgnoreCase]);
                  AddError(2, HKEY_CLASSES_ROOT, '\Applications\'+KeyName+'\shell\'+ValueName, '', regDel, StrCaption, StrText); // добавляем запись в таблицу
                end;
              end; //if (not RegCleanIn.OpenKeyReadOnly('\'+KeyName+'\shell\'+ValueName+'\DropTarget'))
            end; //for l
            SubKeys2.Free;
          end; //if RegCleanIn.HasSubKeys
        end; //if RegCleanIn.OpenKeyReadOnly('\'+KeyName+'\shell')
      end; //for l := 0 to Subkeys.Count - 1 do
      Subkeys.Free;
    end; //if RegCleanIn.OpenKey('\Applications\', False)
  end;
  //2. (ДОП) => КОНЕЦ------------------
  EXCEPT
  END;



  TRY
  //8. Утерян файл справки
  if RegSections[8].Enabled then
  begin
    //HTML Help
    RegCleanIn.RootKey := HKEY_LOCAL_MACHINE;
    RegCleanIn.OpenKey('\SOFTWARE\Microsoft\Windows\HTML Help', False);
    Values := TStringList.Create;
    RegCleanIn.GetValueNames(Values);
    for i := 0 to Values.Count - 1 do
    begin
      if isStop then exit; // выход, если что
      ValueName := Values.Strings[i]; // сохраняем имя активного
      ActKeyName := 'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\HTML Help -> '+ValueName;
      if (i mod 16) = 0 then JvThreadScan.Synchronize(fmRegistryCleanerResults.UpdateStatus);
      ValueData := RegCleanIn.GetDataAsString(ValueName);
      if ((ValueData = '') OR (FileExistsExt(IncludeTrailingPathDelimiter(ValueData) + ValueName) = false)) then
      begin
        StrCaption := '%1';
        StrCaption := stringReplace(StrCaption,'%1',ValueName,[rfReplaceAll, rfIgnoreCase]);
        StrText := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'ErrorTip_8_0');
        StrText := stringReplace(StrText,'%1',IncludeTrailingPathDelimiter(ValueData)+ValueName,[rfReplaceAll, rfIgnoreCase]);
        AddError(8, HKEY_LOCAL_MACHINE, '\SOFTWARE\Microsoft\Windows\HTML Help', ValueName, regDel, StrCaption, StrText); // добавляем запись в таблицу
      end;
    end;
    Values.Free;
    //Help
    RegCleanIn.RootKey := HKEY_LOCAL_MACHINE;
    RegCleanIn.OpenKey('\SOFTWARE\Microsoft\Windows\Help', False);
    Values := TStringList.Create;
    RegCleanIn.GetValueNames(Values);
    for i := 0 to Values.Count - 1 do
    begin
      if isStop then exit; // выход, если что
      ValueName := Values.Strings[i];
      ActKeyName := 'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\HTML Help -> '+ValueName;
      if (i mod 16) = 0 then JvThreadScan.Synchronize(fmRegistryCleanerResults.UpdateStatus);
      ValueData := RegCleanIn.GetDataAsString(ValueName);
      if ((ValueData = '') OR (FileExistsExt(IncludeTrailingPathDelimiter(ValueData) + ValueName) = false)) then
      begin
        StrCaption := '%1';
        StrCaption := stringReplace(StrCaption,'%1',ValueName,[rfReplaceAll, rfIgnoreCase]);
        StrText := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'ErrorTip_8_1');
        StrText := stringReplace(StrText,'%1',IncludeTrailingPathDelimiter(ValueData)+ValueName,[rfReplaceAll, rfIgnoreCase]);
        AddError(8, HKEY_LOCAL_MACHINE, '\SOFTWARE\Microsoft\Windows\Help', ValueName, regDel, StrCaption, StrText); // добавляем запись в таблицу
      end;
    end;
    Values.Free;
  end;
  //8. => КОНЕЦ------------------
  EXCEPT
  END;



  TRY
  //9. Недавние док-ты ОС Windows
  //C:\Users\Alexei\Recent
  if RegSections[9].Enabled then
  begin
    RegCleanIn.RootKey := HKEY_CURRENT_USER;
    KeyName := '';
    l := 0;
    SubKeys := TStringList.Create;
    16:
    RegCleanIn.OpenKey('\Software\Microsoft\Windows\CurrentVersion\Explorer\RecentDocs'+KeyName, False);
    if KeyName = '' then //в перовй итерации
    begin
      RegCleanIn.GetKeyNames(SubKeys);
    end;
    Values := TStringList.Create;
    RegCleanIn.GetValueNames(Values);
    for i := 0 to Values.Count - 1 do
    begin
      if isStop then exit; // выход, если что
      ValueName := Values.Strings[i]; // сохраняем имя активного
      ActKeyName := 'HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\RecentDocs'+KeyName+' -> '+ValueName;
      if (i mod 16) = 0 then JvThreadScan.Synchronize(fmRegistryCleanerResults.UpdateStatus);
      if ValueName = 'MRUListEx' then Continue;

      BufSize := RegCleanIn.GetDataSize(ValueName);
      SetLength(ValueData,0);
      SetLength(ValueData,BufSize);
      RegCleanIn.ReadBinaryData(ValueName,ValueData[1],BufSize);
      ValueData := Copy(ValueData,1,StrLen(PChar(ValueData)));

      sFile := GetSpecialFolderPath(CSIDL_PROFILE);
      sFile := Copy(sFile,1,AnsiPos('Documents', sFile) - 1);
      ValueData := stringReplace(ValueData,':','',[rfReplaceAll, rfIgnoreCase]);
      sFile := IncludeTrailingPathDelimiter(sFile) + 'Recent\' + ValueData + '.lnk';

      ZeroMemory(@LinkInfo,SizeOf(LinkInfo));
      Move(sFile[1],LinkInfo.FullPathAndNameOfLinkFile[0],261);
      GetLinkInfo(@LinkInfo);
      sFile := LinkInfo.FullPathAndNameOfFileToExecute;

      if ((sFile = '') OR ((FileExistsExt(sFile) = false) and (DirExistsExt(sFile) = false))) then
      begin
        StrCaption := '%1';
        StrCaption := stringReplace(StrCaption,'%1',ExtractFileName(sFile),[rfReplaceAll, rfIgnoreCase]);
        StrText := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'ErrorTip_9_0');
        StrText := stringReplace(StrText,'%1',ValueName,[rfReplaceAll, rfIgnoreCase]);
        StrText := stringReplace(StrText,'%2',sFile,[rfReplaceAll, rfIgnoreCase]);
        AddError(9, HKEY_CURRENT_USER, '\Software\Microsoft\Windows\CurrentVersion\Explorer\RecentDocs'+KeyName, ValueName, regDel, StrCaption, StrText); // добавляем запись в таблицу
      end;
    end;
    Values.Free;
    if (l <= SubKeys.Count-1) then
    begin
      KeyName := '\'+SubKeys.Strings[l];
      inc(l);
      goto 16; //для каждого подключа делаем то же самое
    end;
    SubKeys.Free;
  end;
  //9. => КОНЕЦ--------------------
  EXCEPT
  END;



  TRY
  //10. Программы из списка "открыть с помощью"
  if RegSections[10].Enabled then
  begin
    // HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\OpenWithList
    RegCleanIn.RootKey := HKEY_CURRENT_USER;
    RegCleanIn.OpenKey('\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts', False);
    SubKeys := TStringList.Create;
    RegCleanIn.GetKeyNames(SubKeys);
    for i := 0 to SubKeys.Count - 1 do
    begin
      if isStop then exit; // выход, если что
      KeyName := SubKeys.Strings[i];
      ActKeyName := 'HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\'+KeyName;
      if (i mod 16) = 0 then JvThreadScan.Synchronize(fmRegistryCleanerResults.UpdateStatus);
      if KeyName = 'OpenWithList' then
      begin
        RegCleanIn.OpenKey('\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\'+KeyName+'\OpenWithList', False);
        ValueData := RegCleanIn.GetDataAsString('');
        if ValueData = '' then Continue;
        if ((FileExistsExt(ValueData) = false) AND (DirExistsExt(ValueData) = false)) then
        begin
          StrCaption := '%1';
          StrCaption := stringReplace(StrCaption,'%1',ValueData,[rfReplaceAll, rfIgnoreCase]);
          StrText := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'ErrorTip_10_0');
          StrText := stringReplace(StrText,'%1',ValueData,[rfReplaceAll, rfIgnoreCase]);
          StrText := stringReplace(StrText,'%2',KeyName,[rfReplaceAll, rfIgnoreCase]);
          AddError(10, HKEY_CURRENT_USER, '\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\'+KeyName+'\OpenWithList', '', regDel, StrCaption, StrText); // добавляем запись в таблицу
        end;
      end;

      if (
            (isKeyFilled(HKEY_CURRENT_USER, '\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\'+KeyName+'\OpenWithList') = false)
           AND
            (isKeyFilled(HKEY_CURRENT_USER, '\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\'+KeyName+'\OpenWithProgids') = false)
           AND
            (isKeyFilled(HKEY_CURRENT_USER, '\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\'+KeyName+'\UserChoice') = false)
         )
      then
      begin
        StrCaption := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'ErrorCaption_10_0');
        StrCaption := stringReplace(StrCaption,'%1',KeyName,[rfReplaceAll, rfIgnoreCase]);
        StrText := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'ErrorTip_10_1');
        StrText := stringReplace(StrText,'%1',KeyName,[rfReplaceAll, rfIgnoreCase]);
        AddError(10, HKEY_CURRENT_USER, '\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\'+KeyName, '', regDel, StrCaption, StrText); // добавляем запись в таблицу
      end
    end;
    SubKeys.Free;
  end;
  //10. => КОНЕЦ--------------------
  EXCEPT
  END;


  TRY
  //11. Недавние документы Paint
  if RegSections[11].Enabled then
  begin
    RegCleanIn.RootKey := HKEY_CURRENT_USER;
    if RegCleanIn.OpenKey('\Software\Microsoft\Windows\CurrentVersion\Applets\Paint', False) then
    begin
      if isStop then exit; // выход, если что
      ActKeyName := 'HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Applets\Paint';
      JvThreadScan.Synchronize(fmRegistryCleanerResults.UpdateStatus);
      if RegCleanIn.OpenKey('\Software\Microsoft\Windows\CurrentVersion\Applets\Paint\Recent File List', False) then
      begin
        Values := TStringList.Create;
          RegCleanIn.GetValueNames(Values);
          for j := 0 to Values.Count - 1 do
          begin
            ValueData := RegCleanIn.GetDataAsString(Values.Strings[j]);
            if ((ValueData = '') or ((FileExistsExt(ValueData) = false) AND (DirExistsExt(ValueData) = false))) then
            begin
              StrCaption := '%1';
              StrCaption := stringReplace(StrCaption,'%1',ValueData,[rfReplaceAll, rfIgnoreCase]);
              StrText := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'ErrorTip_11_0');
              StrText := stringReplace(StrText,'%1',ValueData,[rfReplaceAll, rfIgnoreCase]);
              AddError(11, HKEY_CURRENT_USER, '\Software\Microsoft\Windows\CurrentVersion\Applets\Paint\Recent File List', Values.Strings[j], regDel, StrCaption, StrText); // добавляем запись в таблицу
            end;
          end;
          Values.Free;
      end; //if RegCleanIn.OpenKey('\Software\Microsoft\Windows\CurrentVersion\Applets\Paint\Recent File List', False)
    end;//if RegCleanIn.OpenKey('\Software\Microsoft\Windows\CurrentVersion\Applets\Paint', False)
  end;
  //11. => КОНЕЦ--------------------
  EXCEPT
  END;


  TRY
  //12. Недавние документы Wordpad
  if RegSections[12].Enabled then
  begin
    RegCleanIn.RootKey := HKEY_CURRENT_USER;
    if RegCleanIn.OpenKey('\Software\Microsoft\Windows\CurrentVersion\Applets\Wordpad', False) then
    begin
      if isStop then exit; // выход, если что
      ActKeyName := 'HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Applets\Wordpad';
      JvThreadScan.Synchronize(fmRegistryCleanerResults.UpdateStatus);
      if RegCleanIn.OpenKey('\Software\Microsoft\Windows\CurrentVersion\Applets\Wordpad\Recent File List', False) then
      begin
        Values := TStringList.Create;
        RegCleanIn.GetValueNames(Values);
        for j := 0 to Values.Count - 1 do
        begin
          ValueData := RegCleanIn.GetDataAsString(Values.Strings[j]);
          if ((ValueData = '') or ((FileExistsExt(ValueData) = false) AND (DirExistsExt(ValueData) = false))) then
          begin
            StrCaption := '%1';
            StrCaption := stringReplace(StrCaption,'%1',ValueData,[rfReplaceAll, rfIgnoreCase]);
            StrText := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'ErrorTip_12_0');
            StrText := stringReplace(StrText,'%1',ValueData,[rfReplaceAll, rfIgnoreCase]);
            AddError(12, HKEY_CURRENT_USER, '\Software\Microsoft\Windows\CurrentVersion\Applets\Wordpad\Recent File List', Values.Strings[j], regDel, StrCaption, StrText); // добавляем запись в таблицу
          end;
        end;
        Values.Free;
      end; //if RegCleanIn.OpenKey('\Software\Microsoft\Windows\CurrentVersion\Applets\Wordpad\Recent File List', False)
    end;//if RegCleanIn.OpenKey('\Software\Microsoft\Windows\CurrentVersion\Applets\Wordpad', False)
  end;
  //12. => КОНЕЦ--------------------
  EXCEPT
  END;



  TRY
  //13. Недавние документы Office
  if RegSections[13].Enabled then
  begin
    RegCleanIn.RootKey := HKEY_CURRENT_USER;
    //\Software\Microsoft\Office\12.0\Word\File MRU
    if RegCleanIn.OpenKey('\Software\Microsoft\Office\', False) then  //если Office вообще установлен
    begin
      SubKeys := TStringList.Create;
      RegCleanIn.GetKeyNames(SubKeys);
      for i := 0 to SubKeys.Count - 1 do
      begin
        if isStop then exit; // выход, если что
        KeyName := SubKeys.Strings[i];
        ActKeyName := 'HKEY_CURRENT_USER\Software\Microsoft\Office\'+KeyName;
        if (i mod 16) = 0 then JvThreadScan.Synchronize(fmRegistryCleanerResults.UpdateStatus);
        if AnsiPos('.', KeyName) <> 0 then //если это версия Office, то окей, открываем
        begin
          SubKeys2 := TStringList.Create;
          RegCleanIn.GetKeyNames(SubKeys2);
          for l := 0 to SubKeys2.Count - 1 do
          begin
            if isStop then exit; // выход, если что
            KeyName2 := SubKeys2.Strings[l];
            ActKeyName := 'HKEY_CURRENT_USER\Software\Microsoft\Office\'+KeyName+'\'+KeyName2;
            if (l mod 16) = 0 then JvThreadScan.Synchronize(fmRegistryCleanerResults.UpdateStatus);
            //\Software\Microsoft\Office\12.0\Word\File MRU
            if RegCleanIn.OpenKey('\Software\Microsoft\Office\'+KeyName+'\'+KeyName2+'\File MRU', False) then
            begin
              Values := TStringList.Create;
              TRY
                RegCleanIn.GetValueNames(Values);
                for j := 0 to Values.Count - 1 do
                begin
                  ValueData := RegCleanIn.GetDataAsString(Values.Strings[j]);
                  t := TStringList.Create;
                  t.text := stringReplace(ValueData, '*', #13#10, [rfReplaceAll]); //потому что строка вида "[F00000001][T01CD2C4A0EDEFFF0]*F:\WinTuning-7\ToDoList.xlsx"
                  if t.Count > 1 then ValueName := t[1];
                  t.free;
                  if ((ValueName = '') OR ((FileExistsExt(ValueName) = false))) then
                  begin
                    StrCaption := ''+KeyName2+' v'+KeyName+'. "%1"';
                    StrCaption := stringReplace(StrCaption,'%1',ExtractFileName(ValueName),[rfReplaceAll, rfIgnoreCase]);
                    StrText := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'ErrorTip_13_0');
                    StrText := stringReplace(StrText,'%1',ValueName,[rfReplaceAll, rfIgnoreCase]);
                    AddError(13, HKEY_CURRENT_USER, '\Software\Microsoft\Office\'+KeyName+'\'+KeyName2+'\File MRU', Values.Strings[j], regDel, StrCaption, StrText); // добавляем запись в таблицу
                  end;
                end;
              FINALLY
                Values.Free;
              END;
            end; //if RegCleanIn.OpenKey('\Software\Microsoft\Office\'+KeyName+'\'+KeyName2+'\File MRU', False)
          end; //for l := 0 to SubKeys2.Count - 1
          SubKeys2.Free;
        end; //if AnsiPos('.', KeyName) <> 0 then
      end;
      SubKeys.Free;
    end;//if RegCleanIn.OpenKey('\Software\Microsoft\Office\', False) ()
  end;
  //13. => КОНЕЦ--------------------
  EXCEPT
  END;



  TRY
  //14. Недавние документы WinRAR
  if RegSections[14].Enabled then
  begin
    //HKEY_CURRENT_USER\Software\WinRAR\DialogEditHistory\ExtrPath
    RegCleanIn.RootKey := HKEY_CURRENT_USER;
    if RegCleanIn.OpenKey('\Software\WinRAR\DialogEditHistory\ExtrPath', False) then
    begin
      if isStop then exit; // выход, если что
      ActKeyName := 'HKEY_CURRENT_USER\Software\WinRAR\DialogEditHistory\ExtrPath';
      JvThreadScan.Synchronize(fmRegistryCleanerResults.UpdateStatus);
      Values := TStringList.Create;
      RegCleanIn.GetValueNames(Values);
      for j := 0 to Values.Count - 1 do
      begin
        ValueData := RegCleanIn.GetDataAsString(Values.Strings[j]);
        if ((ValueData = '') OR ((FileExistsExt(ValueData) = false) AND (DirExistsExt(ValueData) = false))) then
        begin
          StrCaption := '%1';
          StrCaption := stringReplace(StrCaption,'%1',ExtractFileName(ValueData),[rfReplaceAll, rfIgnoreCase]);
          StrText := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'ErrorTip_14_0');
          StrText := stringReplace(StrText,'%1',ValueData,[rfReplaceAll, rfIgnoreCase]);
          AddError(14, HKEY_CURRENT_USER, '\Software\WinRAR\DialogEditHistory\ExtrPath', Values.Strings[j], regDel, StrCaption, StrText); // добавляем запись в таблицу
        end;
      end;
      Values.Free;
    end;
  end;
  //14. => КОНЕЦ--------------------
  EXCEPT
  END;


  TRY
  //15. Недавние документы Media Player Classic
  if RegSections[15].Enabled then
  begin
    //HKEY_CURRENT_USER\Software\Gabest\Media Player Classic\Recent File List
    RegCleanIn.RootKey := HKEY_CURRENT_USER;
    if RegCleanIn.OpenKey('\Software\Gabest\Media Player Classic\Recent File List', False) then
    begin
      if isStop then exit; // выход, если что
      ActKeyName := 'HKEY_CURRENT_USER\Software\Gabest\Media Player Classic\Recent File List';
      JvThreadScan.Synchronize(fmRegistryCleanerResults.UpdateStatus);
      Values := TStringList.Create;
      RegCleanIn.GetValueNames(Values);
      for j := 0 to Values.Count - 1 do
      begin
        ValueData := RegCleanIn.GetDataAsString(Values.Strings[j]);
        if ((ValueData = '') OR ((FileExistsExt(ValueData) = false) AND (DirExistsExt(ValueData) = false))) then
        begin
          StrCaption := '%1';
          StrCaption := stringReplace(StrCaption,'%1',ExtractFileName(ValueData),[rfReplaceAll, rfIgnoreCase]);
          StrText := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'ErrorTip_15_0');
          StrText := stringReplace(StrText,'%1',ValueData,[rfReplaceAll, rfIgnoreCase]);
          AddError(15, HKEY_CURRENT_USER, '\Software\Gabest\Media Player Classic\Recent File List', Values.Strings[j], regDel, StrCaption, StrText); // добавляем запись в таблицу
        end;
      end;
      Values.Free;
    end;
  end;
  //15. => КОНЕЦ--------------------
  EXCEPT
  END;


  TRY
  //16. Автозапуск
  if RegSections[16].Enabled then
  begin
    m := 0;
    17:
    if (m = 0) OR (m = 1) then TempHKEY := HKEY_CURRENT_USER else TempHKEY := HKEY_LOCAL_MACHINE;
    if (m = 0) OR (m = 2) then KeyName := 'Run' else KeyName := 'RunOnce';
    RegCleanIn.RootKey := TempHKEY;
    if RegCleanIn.OpenKey('\Software\Microsoft\Windows\CurrentVersion\'+KeyName, false) then
    begin
      Values := TStringList.Create;
      RegCleanIn.GetValueNames(Values);
      for i := 0 to Values.Count - 1 do
      begin
        if isStop then exit; // выход, если что
        ValueName := Values.Strings[i]; // сохраняем имя активного
        if TempHKEY = HKEY_CURRENT_USER then
        begin
          ValueName2 := 'HKEY_CURRENT_USER';
          ValueName3 := ReadLangStr('WinTuning_Common.lng', 'Common', 'CurrentUser');
        end
        else
        begin
          ValueName2 := 'HKEY_LOCAL_MACHINE';
          ValueName3 := ReadLangStr('WinTuning_Common.lng', 'Common', 'AllUsers');
        end;
        ActKeyName := ValueName2+'\Software\Microsoft\Windows\CurrentVersion\'+KeyName;
        JvThreadScan.Synchronize(fmRegistryCleanerResults.UpdateStatus);
        ValueData := RegCleanIn.GetDataAsString(ValueName);
        if (ValueName = '') and (ValueData = '') then Continue;
        if ((ValueData = '') OR (FileExistsExt(ValueData) = false)) then
        begin
          StrCaption := '%1 ('+KeyName+', '+ValueName3+')';
          StrCaption := stringReplace(StrCaption,'%1',ValueName,[rfReplaceAll, rfIgnoreCase]);
          StrText := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'ErrorTip_16_0');
          StrText := stringReplace(StrText,'%1',ValueName,[rfReplaceAll, rfIgnoreCase]);
          StrText := stringReplace(StrText,'%2',ValueData,[rfReplaceAll, rfIgnoreCase]);
          AddError(16, TempHKEY, '\Software\Microsoft\Windows\CurrentVersion\'+KeyName, ValueName, regDel, StrCaption, StrText); // добавляем запись в таблицу
        end;
      end;
      Values.Free;
    end;
    inc(m);
    if m <= 3 then goto 17;
  end;
  //16. => КОНЕЦ--------------------
  EXCEPT
  END;


  TRY
  //17. Брандмауэр
  if RegSections[17].Enabled then
  begin
    RegCleanIn.RootKey := HKEY_LOCAL_MACHINE;
    RegCleanIn.OpenKey('\SYSTEM\CurrentControlSet\services\SharedAccess\Parameters\FirewallPolicy\FirewallRules', False);
    Values := TStringList.Create;
    RegCleanIn.GetValueNames(Values);
    for i := 0 to Values.Count - 1 do
    begin
      if isStop then exit;
      ValueName := Values.Strings[i];
      ActKeyName := 'HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\services\SharedAccess\Parameters\FirewallPolicy\FirewallRules -> '+ValueName;
      JvThreadScan.Synchronize(fmRegistryCleanerResults.UpdateStatus);
      ValueData := RegCleanIn.GetDataAsString(ValueName);
      if AnsiPos('App=', ValueData) <> 0 then
      begin
        ValueData := Copy(ValueData, AnsiPos('App=', ValueData) + 4, MaxInt);
        ValueData := Copy(ValueData, 1, AnsiPos('|', ValueData) - 1);
        if ((ValueData = '') or (FileExistsExt(ValueData) = false)) then
        begin
          StrCaption := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'ErrorCaption_17_0');
          StrCaption := stringReplace(StrCaption,'%1',ExtractFileName(ValueData),[rfReplaceAll, rfIgnoreCase]);
          StrText := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'ErrorTip_17_0');
          StrText := stringReplace(StrText,'%1',ValueData,[rfReplaceAll, rfIgnoreCase]);
          StrText := stringReplace(StrText,'%2',ValueName,[rfReplaceAll, rfIgnoreCase]);
          AddError(17, HKEY_LOCAL_MACHINE, '\SYSTEM\CurrentControlSet\services\SharedAccess\Parameters\FirewallPolicy\FirewallRules', ValueName, regDel, StrCaption, StrText); // добавляем запись в таблицу
        end;
      end;
    end;
    Values.Free;
  end;
  //17. => КОНЕЦ--------------------
  EXCEPT
  END;


  TRY
  //18. Удаление программ
  if RegSections[18].Enabled then
  begin
    m := 0;
    //получаем список букв съёмных дисков
    RemovableDrivesArr := nil;
    for i := 65 to 90 do
    begin
      DType := GetDriveType(PChar(chr(i) + ':\'));
      if DType <> DRIVE_FIXED then
      begin
        SetLength(RemovableDrivesArr, Length(RemovableDrivesArr)+1);
        RemovableDrivesArr[Length(RemovableDrivesArr)-1] := chr(i) + ':\';
      end;
    end;
    18:
    if (m = 0) then TempHKEY := HKEY_CURRENT_USER else TempHKEY := HKEY_LOCAL_MACHINE;
    RegCleanIn.RootKey := TempHKEY;
    if TempHKEY = HKEY_CURRENT_USER then
    begin
      ValueName2 := 'HKEY_CURRENT_USER';
      ValueName3 := ReadLangStr('WinTuning_Common.lng', 'Common', 'CurrentUser');
    end
    else
    begin
      ValueName2 := 'HKEY_LOCAL_MACHINE';
      ValueName3 := ReadLangStr('WinTuning_Common.lng', 'Common', 'AllUsers');
    end;
    if RegCleanIn.OpenKey('\Software\Microsoft\Windows\CurrentVersion\Uninstall', False) then
    begin
      SubKeys := TStringList.Create;
      RegCleanIn.GetKeyNames(SubKeys);
      for i := 0 to SubKeys.Count - 1 do
      begin
        if isStop then exit;
        KeyName := SubKeys.Strings[i];
        ActKeyName := ValueName2+'\Software\Microsoft\Windows\CurrentVersion\Uninstall\'+KeyName;
        JvThreadScan.Synchronize(fmRegistryCleanerResults.UpdateStatus);
        if not isKeyFilled(TempHKEY, '\Software\Microsoft\Windows\CurrentVersion\Uninstall\' + KeyName) then
        begin
          StrCaption := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'ErrorCaption_18_4');
          StrCaption := stringReplace(StrCaption,'%1',KeyName,[rfReplaceAll, rfIgnoreCase]);
          StrCaption := stringReplace(StrCaption,'%2',ValueName3,[rfReplaceAll, rfIgnoreCase]);
          StrText := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'ErrorTip_18_4');
          StrText := stringReplace(StrText,'%1',KeyName,[rfReplaceAll, rfIgnoreCase]);
          StrText := stringReplace(StrText,'%2',ValueName2,[rfReplaceAll, rfIgnoreCase]);
          AddError(18, TempHKEY, '\Software\Microsoft\Windows\CurrentVersion\Uninstall\'+KeyName, '', regDel, StrCaption, StrText);
        end
        else
        begin
          RegCleanIn.OpenKey('\Software\Microsoft\Windows\CurrentVersion\Uninstall\'+KeyName, False);
          ValueName := 'SystemComponent';
          if RegCleanIn.ValueExists(ValueName) then
          begin
            ValueData := RegCleanIn.GetDataAsString(ValueName);
            if ValueData = '1' then Continue;
          end;
          ValueName := 'NoRemove';
          if RegCleanIn.ValueExists(ValueName) then
          begin
            ValueData := RegCleanIn.GetDataAsString(ValueName);
            if ValueData = '1' then Continue;
          end;
          ValueName := 'UninstallString';
          if RegCleanIn.ValueExists(ValueName) then
          begin
            ValueData := RegCleanIn.GetDataAsString(ValueName);
            if Pos('msiexec', ValueData) <> 0 then Continue;
            isCDROMPath := False;
            for l := 0 to Length(RemovableDrivesArr)-1 do
            begin
              if Pos(LowerCase(RemovableDrivesArr[l]), LowerCase(ValueData)) <> 0 then
              begin
                isCDROMPath := True;
              end;
            end;
            if isCDROMPath then Continue;
            if ((ValueData = '') OR (not FileExistsExt(ValueData)) AND (not DirExistsExt(ValueData))) then
            begin
              if (ValueData = '') then
              begin
                StrCaption := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'ErrorCaption_18_0');
                StrText := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'ErrorTip_18_0');
              end
              else
              begin
                if ((not FileExistsExt(ValueData)) OR (not DirExistsExt(ValueData))) then
                begin
                  StrCaption := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'ErrorCaption_18_1');
                  StrText := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'ErrorTip_18_1');
                end;
              end;
              StrCaption := stringReplace(StrCaption,'%1',KeyName,[rfReplaceAll, rfIgnoreCase]);
              StrCaption := stringReplace(StrCaption,'%2',ValueData,[rfReplaceAll, rfIgnoreCase]);
              StrCaption := stringReplace(StrCaption,'%3',ValueName3,[rfReplaceAll, rfIgnoreCase]);
              StrText := stringReplace(StrText,'%1',ValueName2,[rfReplaceAll, rfIgnoreCase]);
              StrText := stringReplace(StrText,'%2',KeyName,[rfReplaceAll, rfIgnoreCase]);
              StrText := stringReplace(StrText,'%3',ValueData,[rfReplaceAll, rfIgnoreCase]);
              AddError(18, TempHKEY, '\Software\Microsoft\Windows\CurrentVersion\Uninstall\'+KeyName, '', regDel, StrCaption, StrText); // добавляем запись в таблицу
            end;
          end;
          ValueName := 'InstallSource';
          if RegCleanIn.ValueExists(ValueName) then
          begin
            ValueData := RegCleanIn.GetDataAsString(ValueName);
            if Pos('msiexec', ValueData) <> 0 then Continue;
            isCDROMPath := False;
            for l := 0 to Length(RemovableDrivesArr)-1 do
            begin
              if Pos(LowerCase(RemovableDrivesArr[l]), LowerCase(ValueData)) <> 0 then
              begin
                isCDROMPath := True;
              end;
            end;
            if isCDROMPath then Continue;
            if ((ValueData = '') OR (not FileExistsExt(ValueData)) AND (not DirExistsExt(ValueData))) then
            begin
              if (ValueData = '') then
              begin
                StrCaption := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'ErrorCaption_18_2');
                StrText := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'ErrorTip_18_2');
              end
              else
              begin
                if ((not FileExistsExt(ValueData)) OR (not DirExistsExt(ValueData))) then
                begin
                  StrCaption := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'ErrorCaption_18_3');
                  StrText := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'ErrorTip_18_3');
                end;
              end;
              StrCaption := stringReplace(StrCaption,'%1',KeyName,[rfReplaceAll, rfIgnoreCase]);
              StrCaption := stringReplace(StrCaption,'%2',ValueData,[rfReplaceAll, rfIgnoreCase]);
              StrCaption := stringReplace(StrCaption,'%3',ValueName3,[rfReplaceAll, rfIgnoreCase]);
              StrText := stringReplace(StrText,'%1',ValueName2,[rfReplaceAll, rfIgnoreCase]);
              StrText := stringReplace(StrText,'%2',KeyName,[rfReplaceAll, rfIgnoreCase]);
              StrText := stringReplace(StrText,'%3',ValueData,[rfReplaceAll, rfIgnoreCase]);
              AddError(18, TempHKEY, '\Software\Microsoft\Windows\CurrentVersion\Uninstall\'+KeyName, ValueName, regDel, StrCaption, StrText); // добавляем запись в таблицу
            end; //if ((ValueData = '') or (not FileExistsExt(ValueData)))
          end; //if RegCleanIn.ValueExists(ValueName)
        end; //else OF if not isKeyFilled(TempHKEY, '\Software\Microsoft\Windows\CurrentVersion\Uninstall\' + KeyName)
      end; //for i := 0 to SubKeys.Count - 1
      SubKeys.Free;
    end; //if RegCleanIn.OpenKey('\Software\Microsoft\Windows\CurrentVersion\Uninstall', False)
    inc(m);
    if m <= 1 then goto 18;
  end; //if RegSections[18].Enabled
  //18. => КОНЕЦ--------------------
  EXCEPT
  END;


  TRY
  //19. Пути к программам
  if RegSections[19].Enabled then
  begin
    RegClean.RootKey := HKEY_LOCAL_MACHINE;
    RegClean.OpenKey('\SYSTEM\CurrentControlSet\Control\Session Manager\Environment', False);
    ValueData := RegClean.GetDataAsString('Path');
    Paths := explode(';', ValueData);
    for i := 0 to length(Paths) - 1 do
    begin
      if isStop then exit;
      ValueData := Paths[i];
      if ValueData = '' then Continue;
      ActKeyName := 'HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Session Manager\Environment -> Path: "'+ValueName+'"';
      JvThreadScan.Synchronize(fmRegistryCleanerResults.UpdateStatus);
      if (not DirExistsExt(ValueData)) then
      begin
        ArrDeleteElement(Paths, i);
        StrCaption := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'ErrorCaption_19_0');
        StrCaption := stringReplace(StrCaption,'%1',ValueData,[rfReplaceAll, rfIgnoreCase]);
        StrText := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'ErrorTip_19_0');
        StrText := stringReplace(StrText,'%1',ValueData,[rfReplaceAll, rfIgnoreCase]);
        AddError(19, HKEY_LOCAL_MACHINE, '\SYSTEM\CurrentControlSet\Control\Session Manager\Environment', 'Path', regMod, StrCaption, StrText,
                     implode(';', Paths));
      end; //if (not DirExistsExt(ValueData))
    end; //for i := 0 to length(Paths) - 1 do
    RegClean.RootKey := HKEY_LOCAL_MACHINE;
    RegClean.OpenKey('\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths', False);
    SubKeys := TStringList.Create;
    RegClean.GetKeyNames(SubKeys);
    for i := 0 to SubKeys.Count - 1 do
    begin
      if isStop then exit; // выход, если что
      KeyName := SubKeys.Strings[i];
      ActKeyName := 'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\'+ KeyName+'';
      JvThreadScan.Synchronize(fmRegistryCleanerResults.UpdateStatus);
      if not isKeyFilled(HKEY_LOCAL_MACHINE, '\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\'+KeyName) then
      begin
        StrCaption := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'ErrorCaption_19_1');
        StrCaption := stringReplace(StrCaption,'%1',KeyName,[rfReplaceAll, rfIgnoreCase]);
        StrText := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'ErrorTip_19_1');
        StrText := stringReplace(StrText,'%1',KeyName,[rfReplaceAll, rfIgnoreCase]);
        AddError(19, HKEY_LOCAL_MACHINE, '\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\'+KeyName, '', regDel, StrCaption, StrText);
      end //if not isKeyFilled(HKEY_LOCAL_MACHINE, '\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\'+KeyName)
      else
      begin
        RegClean.OpenKey('\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\'+KeyName, false);
        ValueName := '';
        ValueData := RegClean.GetDataAsString(ValueName);
        if (ValueData = '') then
        begin
          StrCaption := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'ErrorCaption_19_2');
          StrCaption := stringReplace(StrCaption,'%1',KeyName,[rfReplaceAll, rfIgnoreCase]);
          StrText := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'ErrorTip_19_2');
          StrText := stringReplace(StrText,'%1',KeyName,[rfReplaceAll, rfIgnoreCase]);
          StrText := stringReplace(StrText,'%2',ValueData,[rfReplaceAll, rfIgnoreCase]);
          AddError(19, HKEY_LOCAL_MACHINE, '\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\'+KeyName, '', regDel, StrCaption, StrText);
        end
        else
        begin
          if ((not FileExistsExt(ValueData)) AND (not DirExistsExt(ValueData))) then
          begin
            StrCaption := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'ErrorCaption_19_3');
            StrCaption := stringReplace(StrCaption,'%1',ValueData,[rfReplaceAll, rfIgnoreCase]);
            StrText := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'ErrorTip_19_3');
            StrText := stringReplace(StrText,'%1',KeyName,[rfReplaceAll, rfIgnoreCase]);
            StrText := stringReplace(StrText,'%2',ValueData,[rfReplaceAll, rfIgnoreCase]);
            AddError(19, HKEY_LOCAL_MACHINE, '\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\'+KeyName, '', regDel, StrCaption, StrText);
          end;
        end; //else if (ValueData = '')
      end; //else if not isKeyFilled(HKEY_LOCAL_MACHINE, '\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\'+KeyName)
    end; //for i := 0 to SubKeys.Count - 1 do
    SubKeys.Free;
  end;
  //19. => КОНЕЦ
  EXCEPT
  END;


  TRY
  //20. Шрифты
  if RegSections[20].Enabled then
  begin
    //C:\Windows\Fonts HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts
    RegClean.RootKey := HKEY_LOCAL_MACHINE;
    RegClean.OpenKey('\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts', False);
    Values := TStringList.Create;
    RegClean.GetValueNames(Values);
    for i := 0 to Values.Count - 1 do
    begin
      if isStop then exit;
      ValueName := Values.Strings[i]; // сохраняем имя активного
      ActKeyName := 'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts -> '+ValueName;
      JvThreadScan.Synchronize(fmRegistryCleanerResults.UpdateStatus);
      ValueData := RegClean.GetDataAsString(ValueName);
      if ExtractFilePath(ValueData) = '' then ValueData := IncludeTrailingPathDelimiter(GetSpecialFolderPath(CSIDL_WINDOWS)) + 'Fonts\' + ValueData;
      if ((ValueData = '') OR (not FileExistsExt(ValueData))) then
      begin
        StrCaption := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'ErrorCaption_20_0');
        StrCaption := stringReplace(StrCaption,'%1',ExtractFileName(ValueData),[rfReplaceAll, rfIgnoreCase]);
        StrText := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'ErrorTip_20_0');
        StrText := stringReplace(StrText,'%1',ValueName,[rfReplaceAll, rfIgnoreCase]);
        StrText := stringReplace(StrText,'%2',ValueData,[rfReplaceAll, rfIgnoreCase]);
        AddError(20, HKEY_LOCAL_MACHINE, '\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts', ValueName, regDel, StrCaption, StrText); // добавляем запись в таблицу
      end;
    end;
    Values.Free;
  end;
  //20. => КОНЕЦ--------------------
  EXCEPT
  END;


  TRY
  //21. Расширения программ
  if RegSections[21].Enabled then
  begin
    RegCleanIn.RootKey := HKEY_LOCAL_MACHINE;
    if RegCleanIn.OpenKey('\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Browser Helper Objects', False) then
    begin
      SubKeys := TStringList.Create;
      RegCleanIn.GetKeyNames(SubKeys);
      RegCleanIn.RootKey := HKEY_CLASSES_ROOT;
      for i := 0 to SubKeys.Count - 1 do
      begin
        if isStop then exit; // выход, если что
        KeyName := SubKeys.Strings[i];
        ActKeyName := 'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Browser Helper Objects\'+KeyName;
        JvThreadScan.Synchronize(fmRegistryCleanerResults.UpdateStatus);
        if RegCleanIn.OpenKey('\CLSID\'+KeyName+'\InprocServer32', false) then
        begin
          ValueData := RegCleanIn.ReadString('');
          if (ValueData = '') then
          begin
            StrCaption := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'ErrorCaption_21_0');
            StrCaption := stringReplace(StrCaption,'%1',KeyName,[rfReplaceAll, rfIgnoreCase]);
            StrText := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'ErrorTip_21_0');
            StrText := stringReplace(StrText,'%1',KeyName,[rfReplaceAll, rfIgnoreCase]);
            AddError(21, HKEY_CLASSES_ROOT, '\CLSID\'+KeyName, '', regDel, StrCaption, StrText); // добавляем запись в таблицу
          end
          else
          begin
            if not FileExistsExt(ValueData) then
            begin
              StrCaption := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'ErrorCaption_21_1');
              StrCaption := stringReplace(StrCaption,'%1',ValueData,[rfReplaceAll, rfIgnoreCase]);
              StrText := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'ErrorTip_21_1');
              StrText := stringReplace(StrText,'%1',KeyName,[rfReplaceAll, rfIgnoreCase]);
              StrText := stringReplace(StrText,'%2',ValueData,[rfReplaceAll, rfIgnoreCase]);
              AddError(21, HKEY_CLASSES_ROOT, '\CLSID\'+KeyName, '', regDel, StrCaption, StrText); // добавляем запись в таблицу
            end;
          end;//else if (ValueData = '')
        end
        else
        begin
          StrCaption := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'ErrorCaption_21_2');
          StrCaption := stringReplace(StrCaption,'%1',KeyName,[rfReplaceAll, rfIgnoreCase]);
          StrText := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'ErrorTip_21_2');
          StrText := stringReplace(StrText,'%1',KeyName,[rfReplaceAll, rfIgnoreCase]);
          AddError(21, HKEY_LOCAL_MACHINE, '\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Browser Helper Objects\'+KeyName, '', regDel, StrCaption, StrText); // добавляем запись в таблицу
        end; //else if RegCleanIn.OpenKey('\CLSID\'+KeyName+'\InprocServer32', false)
      end; //for i := 0 to SubKeys.Count - 1 do
      SubKeys.Free;
    end; //if RegCleanIn.OpenKey('\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Browser Helper Objects', False)
    //-------второй тип ошибок
    RegCleanIn.RootKey := HKEY_LOCAL_MACHINE;
    if RegCleanIn.OpenKey('\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Drivers32', False) then
    begin
      Values := TStringList.Create;
      RegCleanIn.GetValueNames(Values);
      for i := 0 to Values.Count - 1 do
      begin
        if isStop then exit; // выход, если что
        ValueName := Values.Strings[i];
        ActKeyName := 'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Drivers32 -> "'+ValueName+'"';
        JvThreadScan.Synchronize(fmRegistryCleanerResults.UpdateStatus);
        ValueData := RegCleanIn.GetDataAsString(ValueName);
        if (ValueData = '') then
        begin
          StrCaption := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'ErrorCaption_21_3');
          StrCaption := stringReplace(StrCaption,'%1',ValueName,[rfReplaceAll, rfIgnoreCase]);
          StrText := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'ErrorTip_21_3');
          StrText := stringReplace(StrText,'%1',ValueName,[rfReplaceAll, rfIgnoreCase]);
          AddError(21, HKEY_LOCAL_MACHINE, '\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Drivers32', ValueName, regDel, StrCaption, StrText);
        end
        else
        begin
          if ((not FileExistsExt(ValueData))) then
          begin
            StrCaption := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'ErrorCaption_21_4');
            StrCaption := stringReplace(StrCaption,'%1',ValueData,[rfReplaceAll, rfIgnoreCase]);
            StrText := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'ErrorTip_21_4');
            StrText := stringReplace(StrText,'%1',ValueName,[rfReplaceAll, rfIgnoreCase]);
            StrText := stringReplace(StrText,'%2',ValueData,[rfReplaceAll, rfIgnoreCase]);
            AddError(21, HKEY_LOCAL_MACHINE, '\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Drivers32', ValueName, regDel, StrCaption, StrText);
          end;
        end; //else if (ValueData = '')
      end; //for i := 0 to Values.Count - 1 do
      Values.Free;
    end; //if RegCleanIn.OpenKey('\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Drivers32', False)
  end; //if RegSections[20].Enabled
  //21. => КОНЕЦ
  EXCEPT
  END;


  TRY
  //22. Драйверы БД
  if RegSections[22].Enabled then
  begin
    RegCleanIn.RootKey := HKEY_LOCAL_MACHINE;
    if RegCleanIn.OpenKey('\SOFTWARE\ODBC\ODBCINST.INI', False) then
    begin
      SubKeys := TStringList.Create;
      RegClean.GetKeyNames(SubKeys);
      for i := 0 to SubKeys.Count - 1 do
      begin
        if isStop then exit; // выход, если что
        KeyName := SubKeys.Strings[i];
        ActKeyName := 'HKEY_LOCAL_MACHINE\SOFTWARE\ODBC\ODBCINST.INI\'+KeyName;
        JvThreadScan.Synchronize(fmRegistryCleanerResults.UpdateStatus);
        if ((KeyName = 'ODBC Core') OR (KeyName = 'ODBC Drivers') OR (KeyName = 'ODBC Translators')) then Continue;
        RegCleanIn.OpenKey('\SOFTWARE\ODBC\ODBCINST.INI\'+KeyName, false);
        ValueData := '';
        if RegCleanIn.ValueExists('Setup') then
        begin
          ValueData := RegCleanIn.GetDataAsString('Setup');
          if (ValueData = '') then
          begin
            StrCaption := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'ErrorCaption_22_0');
            StrCaption := stringReplace(StrCaption,'%1',KeyName,[rfReplaceAll, rfIgnoreCase]);
            StrText := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'ErrorTip_22_0');
            StrText := stringReplace(StrText,'%1',KeyName,[rfReplaceAll, rfIgnoreCase]);
            AddError(22, HKEY_LOCAL_MACHINE, '\SOFTWARE\ODBC\ODBCINST.INI\'+KeyName, '', regDel, StrCaption, StrText);
          end
          else
          begin
            if ((not FileExistsExt(ValueData))) then
            begin
              StrCaption := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'ErrorCaption_22_1');
              StrCaption := stringReplace(StrCaption,'%1',ValueData,[rfReplaceAll, rfIgnoreCase]);
              StrText := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'ErrorTip_22_1');
              StrText := stringReplace(StrText,'%1',KeyName,[rfReplaceAll, rfIgnoreCase]);
              StrText := stringReplace(StrText,'%2',ValueData,[rfReplaceAll, rfIgnoreCase]);
              AddError(22, HKEY_LOCAL_MACHINE, '\SOFTWARE\ODBC\ODBCINST.INI\'+KeyName, '', regDel, StrCaption, StrText);
            end;
          end; //else if (ValueData = '')
        end; //if RegCleanIn.ValueExists('Setup')
      end; //for i := 0 to SubKeys.Count - 1
      SubKeys.Free;
    end;//if RegCleanIn.OpenKey('\SOFTWARE\ODBC\ODBCINST.INI', False)
  end; //if RegSections[22].Enabled
  //22. => КОНЕЦ
  EXCEPT
  END;


  TRY
  //23. Общие файлы
  if RegSections[23].Enabled then
  begin
    RegCleanIn.RootKey := HKEY_LOCAL_MACHINE;
    RegCleanIn.OpenKey('\SOFTWARE\Microsoft\Windows\CurrentVersion\SharedDLLs', false);
    Values := TStringList.Create;
    RegCleanIn.GetValueNames(Values);
    for i := 0 to Values.Count - 1 do
    begin
      if isStop then exit; // выход, если что
      ValueName := Values.Strings[i];
      ActKeyName := 'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\SharedDLLs -> '+ValueName;
      JvThreadScan.Synchronize(fmRegistryCleanerResults.UpdateStatus);
      ValueData := RegCleanIn.GetDataAsString(ValueName);
      if ((ValueName = '') AND (ValueData = '')) then Continue;
      if (not FileExistsExt(ValueName)) then
      begin
        StrCaption := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'ErrorCaption_23_0');
        StrCaption := stringReplace(StrCaption,'%1',ExtractFileName(ValueName),[rfReplaceAll, rfIgnoreCase]);
        StrText := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'ErrorTip_23_0');
        StrText := stringReplace(StrText,'%1',ValueName,[rfReplaceAll, rfIgnoreCase]);
        AddError(23, HKEY_LOCAL_MACHINE, '\SOFTWARE\Microsoft\Windows\CurrentVersion\SharedDLLs', ValueName, regDel, StrCaption, StrText);
      end; //else if (ValueName = '')
    end; //for i := 0 to Values.Count - 1 do
    Values.Free;
  end; //if RegSections[23].Enabled
  //23. => КОНЕЦ
  EXCEPT
  END;



  TRY
  //24. Настройки программы
  if RegSections[24].Enabled then
  begin
    RegCleanIn.RootKey := HKEY_LOCAL_MACHINE;
    RegCleanIn.OpenKey('\SOFTWARE\Microsoft\MMC\SnapIns', False);
    SubKeys := TStringList.Create;
    RegCleanIn.GetKeyNames(SubKeys);
    for i := 0 to SubKeys.Count - 1 do
    begin
      if isStop then exit;
      KeyName := SubKeys.Strings[i];
      ActKeyName := 'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\MMC\SnapIns\'+KeyName;
      JvThreadScan.Synchronize(fmRegistryCleanerResults.UpdateStatus);
      RegCleanIn.OpenKey('\SOFTWARE\Microsoft\MMC\SnapIns\'+KeyName, False);
      if RegCleanIn.ValueExists('HelpTopic') then
      begin
        ValueData := RegCleanIn.GetDataAsString('HelpTopic');
        if (ValueData = '') then
        begin
          StrCaption := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'ErrorCaption_24_0');
          StrCaption := stringReplace(StrCaption,'%1',KeyName,[rfReplaceAll, rfIgnoreCase]);
          StrText := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'ErrorTip_24_0');
          StrText := stringReplace(StrText,'%1',KeyName,[rfReplaceAll, rfIgnoreCase]);
          AddError(24, HKEY_LOCAL_MACHINE, '\SOFTWARE\Microsoft\MMC\SnapIns\'+KeyName, 'HelpTopic', regDel, StrCaption, StrText);
        end
        else
        begin
          if ((not FileExistsExt(ValueData))) then
          begin
            StrCaption := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'ErrorCaption_24_1');
            StrCaption := stringReplace(StrCaption,'%1',ValueData,[rfReplaceAll, rfIgnoreCase]);
            StrText := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'ErrorTip_24_1');
            StrText := stringReplace(StrText,'%1',KeyName,[rfReplaceAll, rfIgnoreCase]);
            StrText := stringReplace(StrText,'%2',ValueData,[rfReplaceAll, rfIgnoreCase]);
            AddError(24, HKEY_LOCAL_MACHINE, '\SOFTWARE\Microsoft\MMC\SnapIns\'+KeyName, 'HelpTopic', regDel, StrCaption, StrText);
          end;
        end; //else if (ValueData = '')
      end; //if RegCleanIn.ValueExists('HelpTopic') then
    end; //for i := 0 to SubKeys.Count - 1 do
    SubKeys.Free;
  end;
  //24. => КОНЕЦ
  EXCEPT
  END;



  // !!!!!!!!!!!СВЕРИТЬСЯ С TUNEUP


  TRY
  //25. Компоненты программ (ActiveX, COM)
  if RegSections[25].Enabled then
  begin
    ValueName4 := '';
    19:
    RegCleanIn.RootKey := HKEY_CLASSES_ROOT;
    if RegCleanIn.OpenKey(ValueName4+'\CLSID', False) then
    begin
      SubKeys := TStringList.Create;
      RegCleanIn.GetKeyNames(SubKeys);
      for i := 1 to SubKeys.Count - 1 do
      begin
        if isStop then exit; // выход, если что
        KeyName := SubKeys.Strings[i];
        ActKeyName := 'HKEY_CLASSES_ROOT'+ValueName4+'\CLSID\'+KeyName;
        JvThreadScan.Synchronize(fmRegistryCleanerResults.UpdateStatus);
        RegCleanIn.OpenKey(ValueName4+'\CLSID\'+KeyName, False); // AppID
        if RegCleanIn.ValueExists('AppID') then
        begin
          ValueData := RegCleanIn.GetDataAsString('AppID');
          if not RegCleanIn.KeyExists(ValueName4+'\AppID\'+ValueData) then
          begin
            StrCaption := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'ErrorCaption_25_0');
            StrCaption := stringReplace(StrCaption,'%1',KeyName,[rfReplaceAll, rfIgnoreCase]);
            StrText := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'ErrorTip_25_0');
            StrText := stringReplace(StrText,'%1',KeyName,[rfReplaceAll, rfIgnoreCase]);
            StrText := stringReplace(StrText,'%2',ValueData,[rfReplaceAll, rfIgnoreCase]);
            StrText := stringReplace(StrText,'%3',ValueName4,[rfReplaceAll, rfIgnoreCase]);
            AddError(25, HKEY_CLASSES_ROOT, ValueName4+'\CLSID\'+KeyName, 'AppID', regDel, StrCaption, StrText);
          end;
        end; //if RegCleanIn.ValueExists('AppID')
        if RegCleanIn.OpenKey(ValueName4+'\CLSID\'+KeyName+'\TypeLib', False) then // Check from CLSID in TypeLib
        begin
          ValueData := RegCleanIn.GetDataAsString('');
          if not RegCleanIn.KeyExists(ValueName4+'\TypeLib\'+ValueData) then
          begin
            StrCaption := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'ErrorCaption_25_1');
            StrCaption := stringReplace(StrCaption,'%1',KeyName,[rfReplaceAll, rfIgnoreCase]);
            StrText := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'ErrorTip_25_1');
            StrText := stringReplace(StrText,'%1',KeyName,[rfReplaceAll, rfIgnoreCase]);
            StrText := stringReplace(StrText,'%2',ValueData,[rfReplaceAll, rfIgnoreCase]);
            StrText := stringReplace(StrText,'%3',ValueName4,[rfReplaceAll, rfIgnoreCase]);
            AddError(25, HKEY_CLASSES_ROOT, ValueName4+'\CLSID\'+KeyName+'\TypeLib', '', regDel, StrCaption, StrText);
          end;
        end; //if RegCleanIn.OpenKey('\CLSID\'+KeyName+'\TypeLib', False)
      end; //for i := 1 to SubKeys.Count - 1
      SubKeys.Free;
    end; //if RegClean.OpenKey('\CLSID', False)
    if ValueName4 = '' then //если мы ещё не делали для WOW64
    begin
      ValueName4 := '\Wow6432Node';
      goto 19; //goback
    end;

    // Interface
    ValueName4 := '';
    20:
    RegCleanIn.RootKey := HKEY_CLASSES_ROOT;
    if RegCleanIn.OpenKey(ValueName4+'\Interface', False) then
    begin
      SubKeys := TStringList.Create;
      RegCleanIn.GetKeyNames(SubKeys);
      for i := 1 to SubKeys.Count - 1 do
      begin
        if isStop then exit; // выход, если что
        KeyName := SubKeys.Strings[i];
        ActKeyName := 'HKEY_CLASSES_ROOT'+ValueName4+'\Interface\'+KeyName;
        JvThreadScan.Synchronize(fmRegistryCleanerResults.UpdateStatus);
        RegCleanIn.OpenKey(ValueName4+'\Interface\'+KeyName, False);
        if RegCleanIn.OpenKey(ValueName4+'\Interface\'+KeyName+'\TypeLib', False) then //Check from Interface in TypeLib
        begin
          ValueData := RegCleanIn.GetDataAsString('');
          if not RegCleanIn.KeyExists(ValueName4+'\TypeLib\'+ValueData) then
          begin
            StrCaption := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'ErrorCaption_25_2');
            StrCaption := stringReplace(StrCaption,'%1',KeyName,[rfReplaceAll, rfIgnoreCase]);
            StrText := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'ErrorTip_25_2');
            StrText := stringReplace(StrText,'%1',KeyName,[rfReplaceAll, rfIgnoreCase]);
            StrText := stringReplace(StrText,'%2',ValueData,[rfReplaceAll, rfIgnoreCase]);
            StrText := stringReplace(StrText,'%3',ValueName4,[rfReplaceAll, rfIgnoreCase]);
            AddError(25, HKEY_CLASSES_ROOT, ValueName4+'\Interface\'+KeyName+'\TypeLib', '', regDel, StrCaption, StrText);
          end;
        end; //if RegCleanIn.OpenKey(ValueName4+'\Interface\'+KeyName+'\TypeLib', False)
      end; //for i := 1 to SubKeys.Count - 1 do
      SubKeys.Free;
    end; //if RegCleanIn.OpenKey(ValueName4+'\Interface', False)
    if ValueName4 = '' then //если мы ещё не делали для WOW64
    begin
      ValueName4 := '\Wow6432Node';
      goto 20; //goback
    end;

    // TypeLib
    ValueName4 := '';
    21:
    RegClean.RootKey := HKEY_CLASSES_ROOT;
    if RegClean.OpenKey(ValueName4+'\TypeLib', False) then
    begin
      SubKeys := TStringList.Create;
      RegClean.GetKeyNames(SubKeys);
      for i := 1 to SubKeys.Count - 1 do
      begin
        if isStop then exit; // выход, если что
        KeyName := SubKeys.Strings[i];
        ActKeyName := 'HKEY_CLASSES_ROOT'+ValueName4+'\TypeLib\'+KeyName;
        JvThreadScan.Synchronize(fmRegistryCleanerResults.UpdateStatus);
        if RegClean.OpenKey(ValueName4+'\TypeLib\'+KeyName, False) then
          begin
          SubKeys2 := TStringList.Create;
          RegClean.GetKeyNames(SubKeys2);
          if SubKeys2.Count = 0 then
          begin
            SubKeys2.Free;
            Continue;
          end;
          KeyName2 := SubKeys2.Strings[0];
          if RegClean.OpenKey(ValueName4+'\TypeLib\'+KeyName+'\'+KeyName2+'\HELPDIR', False) then
          begin
            ValueData := RegClean.GetDataAsString('');
            if (ValueData = '') then
            begin
              StrCaption := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'ErrorCaption_25_3');
              StrCaption := stringReplace(StrCaption,'%1',KeyName,[rfReplaceAll, rfIgnoreCase]);
              StrCaption := stringReplace(StrCaption,'%2',KeyName2,[rfReplaceAll, rfIgnoreCase]);
              StrText := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'ErrorTip_25_3');
              StrText := stringReplace(StrText,'%1',ValueName4,[rfReplaceAll, rfIgnoreCase]);
              StrText := stringReplace(StrText,'%2',KeyName,[rfReplaceAll, rfIgnoreCase]);
              StrText := stringReplace(StrText,'%3',KeyName2,[rfReplaceAll, rfIgnoreCase]);
              AddError(25, HKEY_CLASSES_ROOT, ValueName4+'\TypeLib\'+KeyName+'\'+KeyName2+'\HELPDIR', '', regDel, StrCaption, StrText);
            end
            else
            begin
              if
                 (not FileExistsExt(ExcludeTrailingPathDelimiter(ValueData)))
              AND
                 (not DirExistsExt(ExcludeTrailingPathDelimiter(ValueData)))
              then
              begin
                StrCaption := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'ErrorCaption_25_4');
                StrCaption := stringReplace(StrCaption,'%1',ValueData,[rfReplaceAll, rfIgnoreCase]);
                StrText := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'ErrorTip_25_4');
                StrText := stringReplace(StrText,'%1',ValueName4,[rfReplaceAll, rfIgnoreCase]);
                StrText := stringReplace(StrText,'%2',KeyName,[rfReplaceAll, rfIgnoreCase]);
                StrText := stringReplace(StrText,'%3',KeyName2,[rfReplaceAll, rfIgnoreCase]);
                StrText := stringReplace(StrText,'%4',ValueData,[rfReplaceAll, rfIgnoreCase]);
                AddError(25, HKEY_CLASSES_ROOT, ValueName4+'\TypeLib\'+KeyName+'\'+KeyName2+'\HELPDIR', '', regDel, StrCaption, StrText);
              end;
            end; //else if (ValueData = '')
          end; //if RegClean.OpenKey(ValueName4+'\TypeLib\'+KeyName+'\'+KeyName2+'\HELPDIR', False)
          SubKeys2.Free;
        end; //if RegClean.OpenKey(ValueName4+'\TypeLib\'+KeyName, False)
      end; //for i := 1 to SubKeys.Count - 1 do
      SubKeys.Free;
    end; //if RegClean.OpenKey(ValueName4+'\TypeLib', False)
    if ValueName4 = '' then //если мы ещё не делали для WOW64
    begin
      ValueName4 := '\Wow6432Node';
      goto 21; //goback
    end;
  end;
  //25. => КОНЕЦ
  EXCEPT
  END;


  TRY
  //26. Звуки
  if RegSections[26].Enabled then
  begin
    RegClean.RootKey := HKEY_CURRENT_USER;
    if RegClean.OpenKey('\AppEvents\Schemes\Apps', False) then
    begin
      SubKeys := TStringList.Create;
      RegClean.GetKeyNames(SubKeys);
      for i := 0 to SubKeys.Count - 1 do
      begin
        if isStop then exit; // выход, если что
        KeyName := SubKeys.Strings[i];
        ActKeyName := 'HKEY_CURRENT_USER\AppEvents\Schemes\Apps\'+KeyName;
        JvThreadScan.Synchronize(fmRegistryCleanerResults.UpdateStatus);
        if RegClean.OpenKey('\AppEvents\Schemes\Apps\'+KeyName, False) then
        begin
          SubKeys2 := TStringList.Create;
          RegClean.GetKeyNames(SubKeys2);
          for j := 0 to SubKeys2.Count - 1 do
          begin
            if isStop then exit; // выход, если что
            KeyName2 := SubKeys2.Strings[j];
            ActKeyName := 'HKEY_CURRENT_USER\AppEvents\Schemes\Apps\'+KeyName+'\'+KeyName2;
            JvThreadScan.Synchronize(fmRegistryCleanerResults.UpdateStatus);
            if RegClean.OpenKey('\AppEvents\Schemes\Apps\'+KeyName+'\'+KeyName2+'\.Current', False) then
            begin
              ValueData := RegClean.GetDataAsString('');
              if ValueData = '' then Continue;
              if (not FileExistsExt(ValueData)) then
              begin
                StrCaption := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'ErrorCaption_26_0');
                StrCaption := stringReplace(StrCaption,'%1',ExtractFileName(ValueData),[rfReplaceAll, rfIgnoreCase]);
                StrText := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'ErrorTip_26_0');
                StrText := stringReplace(StrText,'%1',KeyName,[rfReplaceAll, rfIgnoreCase]);
                StrText := stringReplace(StrText,'%2',KeyName2,[rfReplaceAll, rfIgnoreCase]);
                StrText := stringReplace(StrText,'%3',ValueData,[rfReplaceAll, rfIgnoreCase]);
                AddError(26, HKEY_CURRENT_USER, '\AppEvents\Schemes\Apps\'+KeyName+'\'+KeyName2+'\.Current', '', regMod, StrCaption, StrText, '');
              end;
            end; ////if RegClean.OpenKey('\AppEvents\Schemes\Apps\'+KeyName+'\'+KeyName2+'\.Current', False)
          end; //for j := 0 to SubKeys2.Count - 1 do
          SubKeys2.Free;
        end; //if RegClean.OpenKey('\AppEvents\Schemes\Apps\'+KeyName, False)
      end; //for i := 0 to SubKeys.Count - 1 do
      SubKeys.Free;
    end; //if RegClean.OpenKey('\AppEvents\Schemes\Apps', False)
  end;
  //26. => КОНЕЦ
  EXCEPT
  END;




14:
  RegCleanIn.Free;
end;
//=========================================================












end.

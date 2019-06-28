unit RegistryCleaner_Results;

interface

{$WARN UNIT_PLATFORM OFF}
{$WARN SYMBOL_PLATFORM OFF}

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ImgList, StdCtrls, ExtCtrls, sCheckListBox, sLabel, Tlhelp32,
  JvExStdCtrls, JvCombobox, JvDriveCtrls, Registry, IniFiles, FindFile,
  ShellAPI, acPNG, Grids, BaseGrid, AdvGrid, magwmi, magsubs1, AdvSmoothButton,
  AdvGlowButton, JvComponentBase, JvComputerInfoEx, sSplitter,
  ComCtrls, HTMLabel, PNGImage, DateUtils, JvBaseDlg, JvSHFileOperation,
  AdvChartPaneEditor, AdvChartView, AdvChartViewGDIP, Math, AdvChart,
  AdvChartGDIP, AdvChartUtil, AdvChartPaneEditorGDIP, AdvChartSerieEditor,
  ExtDlgs, AdvSmoothLabel, JvGIF, typinfo, AdvObj, AdvGroupBox, Menus,
  JvBrowseFolder, acAlphaImageList, fspTaskbarMgr, sScrollBox, AdvScrollBox, System.UITypes;


type
  tcodepagearray = array[$80..$ff] of char;
  PHICON = ^HICON;


type
  TfmRegistryCleanerResults = class(TForm)
    PanelCenter: TPanel;
    PanelBottom: TPanel;
    ShapeBottom: TShape;
    PanelLeft: TPanel;
    ShapeLeft: TShape;
    agbLogo: TAdvGroupBox;
    ThemeImgMainCaption: TImage;
    albLogoText: TsLabel;
    imgWinTuning: TImage;
    ThemeShapeMainCaption: TShape;
    MainMenuResults: TMainMenu;
    mmMainFile: TMenuItem;
    mmFoundAnError: TMenuItem;
    mmExit: TMenuItem;
    mmSelected: TMenuItem;
    mmOpenKey: TMenuItem;
    mmMainHelp: TMenuItem;
    mmOpenHelp: TMenuItem;
    mmAbout: TMenuItem;
    mmFix: TMenuItem;
    mmNewScan: TMenuItem;
    mmWebSite: TMenuItem;
    agbErrorsList: TAdvGroupBox;
    GridListOfEntries: TAdvStringGrid;
    imgLogoBack: TImage;
    popupMenuFiles: TPopupMenu;
    popupMenuOpenKey: TMenuItem;
    fspTaskbarMgrProgress: TfspTaskbarMgr;
    GridCategories: TAdvStringGrid;
    GridListOfSections: TAdvStringGrid;
    agbScanning: TAdvGroupBox;
    lbStatus: TLabel;
    ProgressBarScanning: TProgressBar;
    agbGeneralResults: TAdvGroupBox;
    GridDetails: TAdvStringGrid;
    SplitterInErrors: TSplitter;
    PanelErrorDetails: TPanel;
    lbDetailsCaption: TLabel;
    AlphaImgsDetails: TsAlphaImageList;
    popupMenuAddToExclusion: TMenuItem;
    mmAddToExclusion: TMenuItem;
    mmSettings: TMenuItem;
    lbStatFound: TLabel;
    lbCaption: TLabel;
    lbDesc: TLabel;
    lbParamName: TLabel;
    edParamValue: TEdit;
    SplitterLeft: TSplitter;
    btBack: TAdvGlowButton;
    btClean: TAdvGlowButton;
    btClose: TAdvGlowButton;
    btAllCategories: TAdvGlowButton;
    btCancel: TAdvGlowButton;
    lbStatFixed: TLabel;
    AdvScrollBox1: TAdvScrollBox;
    AlphaImgsSections16: TsAlphaImageList;
    AlphaImgsSections32: TsAlphaImageList;
    procedure btBackClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure LoadPNG2Prog(dllname, resname: string; imgobj: TImage);
    procedure mmExitClick(Sender: TObject);
    procedure ApplyTheme;
    procedure btCleanClick(Sender: TObject);
    procedure btCloseClick(Sender: TObject);
    procedure btCancelClick(Sender: TObject);
    function IsCheckedSomething: boolean;
    procedure popupMenuOpenKeyClick(Sender: TObject);
    procedure UpdateStatus;
    procedure LoadSections;
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure GridListOfEntriesClickCell(Sender: TObject; ARow, ACol: Integer);
    procedure ShowRegSections;
    procedure GridCategoriesClickCell(Sender: TObject; ARow, ACol: Integer);
    procedure btAllCategoriesClick(Sender: TObject);
    procedure GridListOfSectionsButtonClick(Sender: TObject; ACol, ARow: Integer);
    procedure GridListOfEntriesCheckBoxClick(Sender: TObject; ACol, ARow: Integer; State: Boolean);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure GridListOfEntriesDblClickCell(Sender: TObject; ARow, ACol: Integer);
    procedure popupMenuAddToExclusionClick(Sender: TObject);
    procedure GridListOfSectionsResize(Sender: TObject);
    procedure GridDetailsClickCell(Sender: TObject; ARow, ACol: Integer);
    procedure GridListOfEntriesRightClickCell(Sender: TObject; ARow, ACol: Integer);
    function KillTask(ExeFileName: string): integer;
    procedure ApplyLang;
    procedure FormActivate(Sender: TObject);

  private
    { Private declarations }
    isLoaded: boolean; //Переменная нужна, чтобы не было изменений настроек при запуске программы
  public
    { Public declarations }
    ThemeName: string; //название темы оформления
    ThemeFile: TIniFile; //ini-файл темы оформления
    last_index: integer; //последняя добавленная строка
    SelectedDir: string;
    FileTypes: array of string; //массив типов файлов
    SelectedSectionIndex: integer; //выбранный индекс раздела
    SelectedErrorIndex: integer; //выбранный индекс ошибки
  protected
    { Protected declarations }
    procedure CreateParams(var Params: TCreateParams); override;
  end;

var
  fmRegistryCleanerResults: TfmRegistryCleanerResults;


implementation

uses DataMod, RegistryCleaner_Msg, RegistryCleaner_Settings;

const
  ScreenWidth: LongInt = 1680;
  ScreenHeight: LongInt = 1050;


{$R *.dfm}


{СОХРАНИТЬ ПОЗИЦИЮ ОКНА ПРИ ЗАКРЫТИИ}
procedure SaveWndPosition(FormName: TForm; KeyToSave: PChar); external 'Functions.dll';

{ВОССТАНОВИТЬ ПОЗИЦИЮ ОКНА ПРИ ЗАПУСКЕ}
procedure RestoreWndPosition(FormName: TForm; KeyToSave: PChar); external 'Functions.dll';

{ВЫВОДИМ РАЗНУЮ ИНФУ О ВЕРСИИ WINTUNING: ИНДЕКС, ГОД ОКОНЧАНИЯ ИСПОЛЬЗОВАНИЯ И ТД.}
function GetWTVerInfo(info_id: integer): integer; external 'Functions.dll';

{ПРЕОБРАЗОВАНИЕ СТРОКИ ВИДА R,G,B В TCOLOR}
function ReadRBG(instr: PChar): TColor; external 'Functions.dll';

{ПРЕОБРАЗОВАНИЕ СТРОКИ ВИДА %WinTuning_PATH% В C:\Program Files\WinTuning 7}
function ThemePathConvert(InStr, InThemeName: PChar): PChar; external 'Functions.dll';

{ОТКРЫТЬ ОПРЕДЕЛЁННЫЙ РАЗДЕЛ СПРАВКИ}
procedure ViewHelp(page: PChar); stdcall; external 'Functions.dll';

{О ПРОГРАММЕ}
function CreateTheForm(S1, S2, S3: PChar): integer; stdcall export; external 'About.dll';

{ЧТЕНИЕ ЯЗЫКОВОЙ СТРОКИ ИЗ ОПРЕДЕЛЁННОГО ФАЙЛА}
function ReadLangStr(FileName, Section, Caption: PChar): PChar; external 'Functions.dll';

{ВЫВОДИМ НАЗВАНИЕ ВЕРСИИ WINTUNING: [XP, VISTER, 7]}
function GetCapInfo(WTVerID, info_id: integer): shortstring; external 'Functions.dll';

{ЧТЕНИЕ НАСТРОЕК ПРОГРАММЫ}
function GetProgParam(paramname: PChar): PChar; external 'Functions.dll';




//=========================================================
{ОТОБРАЖАЕТ ФОРМУ НА ПАНЕЛИ ЗАДАЧ}
//---------------------------------------------------------
procedure TfmRegistryCleanerResults.CreateParams(var Params: TCreateParams);
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
{ОТМЕЧЕН ЛИ ХОТЯ БЫ ОДИН ПУНКТ}
//---------------------------------------------------------
function TfmRegistryCleanerResults.IsCheckedSomething: boolean;
var
  i: Integer;
begin
  Result := False;
  for i := 0 to Length(fmDataMod.RegErrors)-1 do
  begin
    if (fmDataMod.RegErrors[i].Enabled) AND (not fmDataMod.RegErrors[i].Excluded) then
    begin
      Result := True;
      Exit;
    end;
  end;
end;
//=========================================================



//=========================================================
{ПРИ ПОСТАВНОКЕ/СНЯТИИ ФЛАЖКА НАПРОТИВ ОШИБКИ}
//---------------------------------------------------------
procedure TfmRegistryCleanerResults.GridListOfEntriesCheckBoxClick(Sender: TObject; ACol, ARow: Integer; State: Boolean);
var
  i,j: integer;
begin
  if (ARow = 0) AND (ACol = 1) then
  begin
    GridListOfEntries.GetCheckBoxState(1,0,State);
    for i := 1 to GridListOfEntries.RowCount-1 do
    begin
      GridListOfEntries.SetCheckBoxState(1,i,State);
      j := StrToInt(GridListOfEntries.Cells[2,i]); //получаем ID
      fmDataMod.RegErrors[j].Enabled := State;
    end;
  end;
  if ARow > 0 then
  begin
    i := StrToInt(GridListOfEntries.Cells[2,ARow]);
  //  GridListOfEntries.OnClickCell(Self, ARow, ACol);
    fmDataMod.RegErrors[i].Enabled := State;
  end;
end;
//=========================================================



//=========================================================
{ПРИ ЩЕЛЧКЕ ПО ЯЧЕЙКЕ СПИСКА ОШИБОК}
//---------------------------------------------------------
procedure TfmRegistryCleanerResults.GridListOfEntriesClickCell(Sender: TObject; ARow, ACol: Integer);
var
  i: integer;
  TempStr: string;
begin
  if (ARow > 0) then
  begin
    SelectedErrorIndex := StrToInt(GridListOfEntries.Cells[2,ARow]);

    GridDetails.BeginUpdate;
    GridDetails.RemoveRows(1, GridDetails.RowCount-1);
    i := 1;

    GridDetails.AddRow;
    GridDetails.Cells[1,i] := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'ErrorName');
    GridDetails.Cells[2,i] := fmDataMod.RegErrors[SelectedErrorIndex].Caption;
    GridDetails.AddImageIdx(1,i,0,haBeforeText,vaCenter);
    lbParamName.Caption := GridDetails.Cells[1,i]+':';
    edParamValue.Text := GridDetails.Cells[2,i];
    inc(i);

    GridDetails.AddRow;
    GridDetails.Cells[1,i] := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'ErrorDesc');
    GridDetails.Cells[2,i] := fmDataMod.RegErrors[SelectedErrorIndex].Text;
    GridDetails.AddImageIdx(1,i,-1,haBeforeText,vaCenter);
    inc(i);

    GridDetails.AddRow;
    GridDetails.Cells[1,i] := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'SectionName');
    GridDetails.Cells[2,i] := fmDataMod.RegSections[fmDataMod.RegErrors[SelectedErrorIndex].SectionCode].Caption;
    GridDetails.AddImageIdx(1,i,-1,haBeforeText,vaCenter);
    inc(i);

    TempStr := '';
    if fmDataMod.RegErrors[SelectedErrorIndex].RootKey = HKEY_CLASSES_ROOT then TempStr := 'HKEY_CLASSES_ROOT';
    if fmDataMod.RegErrors[SelectedErrorIndex].RootKey = HKEY_LOCAL_MACHINE then TempStr := 'HKEY_LOCAL_MACHINE';
    if fmDataMod.RegErrors[SelectedErrorIndex].RootKey = HKEY_CURRENT_USER then TempStr := 'HKEY_CURRENT_USER';
    GridDetails.AddRow;
    GridDetails.Cells[1,i] := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'KeyInReg');
    GridDetails.Cells[2,i] := TempStr+fmDataMod.RegErrors[SelectedErrorIndex].SubKey;
    GridDetails.AddImageIdx(1,i,2,haBeforeText,vaCenter);
    inc(i);

    TempStr := fmDataMod.RegErrors[SelectedErrorIndex].Parameter;
    if TempStr <> '' then
    begin
      GridDetails.AddRow;
      GridDetails.Cells[1,i] := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'ParamInReg');
      GridDetails.Cells[2,i] := TempStr;
      GridDetails.AddImageIdx(1,i,3,haBeforeText,vaCenter);
      inc(i);
    end;

    TempStr := fmDataMod.RegErrors[SelectedErrorIndex].NewValueData;
    if TempStr <> '' then
    begin
      GridDetails.AddRow;
      GridDetails.Cells[1,i] := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'NewParam');
      GridDetails.Cells[2,i] := TempStr;
      GridDetails.AddImageIdx(1,i,-1,haBeforeText,vaCenter);
      inc(i);
    end;

    GridDetails.AddRow;
    TempStr := '';
    if fmDataMod.RegErrors[SelectedErrorIndex].ErrorType = regKeyName then TempStr := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'InvalidKey');
    if fmDataMod.RegErrors[SelectedErrorIndex].ErrorType = regValueName then TempStr := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'InvalidParam');
    GridDetails.Cells[1,i] := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'ErrorType');
    GridDetails.Cells[2,i] := TempStr;
    GridDetails.AddImageIdx(1,i,-1,haBeforeText,vaCenter);
    inc(i);

    GridDetails.AddRow;
    TempStr := '';
    if fmDataMod.RegErrors[SelectedErrorIndex].Solution = regDel then TempStr := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'ResMethod1');
    if fmDataMod.RegErrors[SelectedErrorIndex].Solution = regMod then TempStr := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'ResMethod2');
    GridDetails.Cells[1,i] := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'SolutionMethod');
    GridDetails.Cells[2,i] := TempStr;
    GridDetails.AddImageIdx(1,i,1,haBeforeText,vaCenter);

    GridDetails.AutoSizeColumns(True);

    GridDetails.EndUpdate;
  end;
end;
//=========================================================



//=========================================================
{ПРИ ЩЕЛЧКЕ ПРАВОЙ КНОПКОЙ МЫШИ ПО ЯЧЕЙКЕ СПИСКА ОШИБОК}
//---------------------------------------------------------
procedure TfmRegistryCleanerResults.GridListOfEntriesRightClickCell(Sender: TObject; ARow, ACol: Integer);
begin
  GridListOfEntriesClickCell(Self, ARow, ACol);
end;
//=========================================================



//=========================================================
{ФУНКЦИЯ ДЛЯ ЗАКРЫТИЯ ПРИЛОЖЕНИЯ ПО ИМЕНИ ФАЙЛА}
//---------------------------------------------------------
function TfmRegistryCleanerResults.KillTask(ExeFileName: string): integer;
const PROCESS_TERMINATE=$0001;
var
  ContinueLoop: BOOL;
  FSnapshotHandle: THandle;
  FProcessEntry32: TProcessEntry32;
begin
  result := 0;
  FSnapshotHandle := CreateToolhelp32Snapshot (TH32CS_SNAPPROCESS, 0);
  FProcessEntry32.dwSize := Sizeof(FProcessEntry32);
  ContinueLoop := Process32First(FSnapshotHandle, FProcessEntry32);
  while integer(ContinueLoop)<> 0 do
  begin
    if ((UpperCase(ExtractFileName(FProcessEntry32.szExeFile)) = UpperCase(ExeFileName))
    OR (UpperCase(FProcessEntry32.szExeFile) = UpperCase(ExeFileName))) then
    begin
      Result := Integer(TerminateProcess(OpenProcess( PROCESS_TERMINATE, BOOL(0), FProcessEntry32.th32ProcessID), 0));
    end;
{    Memo1.Lines.Add(IntToStr(FProcessEntry32.th32ProcessID)+'|||'
    +IntToStr(FProcessEntry32.cntThreads)+'|||'
    +IntToStr(FProcessEntry32.th32ParentProcessID)+'|||'+
    FProcessEntry32.szExeFile
);}
    ContinueLoop := Process32Next(FSnapshotHandle, FProcessEntry32);
  end;
  CloseHandle(FSnapshotHandle);
end;
//=========================================================



//=========================================================
{ПРИ ДВОЙНОМ ЩЕЛЧКЕ ПО СПИСКУ ОШИБОК}
//---------------------------------------------------------
procedure TfmRegistryCleanerResults.GridListOfEntriesDblClickCell(Sender: TObject; ARow, ACol: Integer);
var
  t: TStringList;
  LastKey, TempStr, NewLastKey: string;
  SelectedErrorIndex: integer;
begin
  KillTask('regedit.exe');
  LastKey := '';
  NewLastKey := 'HKEY_CURRENT_USER';
  SelectedErrorIndex := StrToInt(GridListOfEntries.GridCells[2, ARow]);
  if fmDataMod.RegErrors[SelectedErrorIndex].RootKey = HKEY_CLASSES_ROOT then NewLastKey := 'HKEY_CLASSES_ROOT';
  if fmDataMod.RegErrors[SelectedErrorIndex].RootKey = HKEY_LOCAL_MACHINE then NewLastKey := 'HKEY_LOCAL_MACHINE';
  if fmDataMod.RegErrors[SelectedErrorIndex].RootKey = HKEY_CURRENT_USER then NewLastKey := 'HKEY_CURRENT_USER';
  if fmDataMod.RegErrors[SelectedErrorIndex].SubKey <> '\' then NewLastKey := NewLastKey + fmDataMod.RegErrors[SelectedErrorIndex].SubKey;
  fmDataMod.RegClean.RootKey := HKEY_CURRENT_USER;
  fmDataMod.RegClean.OpenKey('\Software\Microsoft\Windows\CurrentVersion\Applets\Regedit', True);
  if fmDataMod.RegClean.ValueExists('LastKey') then LastKey := fmDataMod.RegClean.ReadString('LastKey');
  t := TStringList.Create; // создаём класс
  t.text := stringReplace(LastKey, '\', #13#10, [rfReplaceAll, rfIgnoreCase]);
  TempStr := '';
  if t.Count > 1 then TempStr := t[0]; //самый-самый корень (слово "Компьютер" или "Computer" - локализованная строка)
  if TempStr <> '' then NewLastKey := TempStr + '\' +NewLastKey;
  fmDataMod.RegClean.WriteString('LastKey', NewLastKey);
  ShellExecute(Application.Handle, 'open', 'regedit.exe', nil, nil, SW_SHOW);
  t.Free;
end;
//=========================================================



//=========================================================
{ПРИ ЩЕЛЧКЕ НА КНОПКЕ "ПОДРОБНО" СПИСКА ОШИБОК ПО КАТЕГОРИЯМ}
//---------------------------------------------------------
procedure TfmRegistryCleanerResults.GridListOfSectionsButtonClick(Sender: TObject; ACol, ARow: Integer);
var
  i, index: integer;
begin
  index := StrToInt(GridListOfSections.Cells[2,ARow]);
  for i := 1 to GridCategories.RowCount-1 do
  begin
    if StrToInt(GridCategories.GridCells[1,i]) = index then
    begin
      GridCategories.SelectRows(i, 1);
      GridCategories.OnClickCell(Self,i,1);
      Exit;
    end;
  end;
end;
//=========================================================



//=========================================================
{ПРИ ИЗМЕНЕНИИ РАЗМЕРА ЦЕНТРАЛЬНОГО СПИСКА КАТЕГОРИЙ ОШИБОК}
//---------------------------------------------------------
procedure TfmRegistryCleanerResults.GridListOfSectionsResize(Sender: TObject);
begin
  GridListOfSections.ColWidths[1] := PanelCenter.Width - 250;
  GridListOfSections.AutoSizeRows(False);
end;
//=========================================================



//=========================================================
{ПРИ ЩЕЛЧКЕ НА СПИСКЕ КАТЕГОРИЙ}
//---------------------------------------------------------
procedure TfmRegistryCleanerResults.GridCategoriesClickCell(Sender: TObject; ARow, ACol: Integer);
var
  i, LastIndex: integer;
  isAllEnabled: boolean;
  WrapInt: integer;
  ListWidth: integer;
begin
  if ARow > 0 then
  begin
    agbGeneralResults.Visible := False;
    agbGeneralResults.Enabled := False;
    agbErrorsList.Visible := True;
    agbErrorsList.Enabled := True;
    btAllCategories.Enabled := agbErrorsList.Enabled;
    if (SelectedSectionIndex = StrToInt(GridCategories.GridCells[1, ARow])) AND (SelectedSectionIndex <> -1) then exit; //чтобы не обновлять и так уже выбранную категорию
    SelectedSectionIndex := StrToInt(GridCategories.GridCells[1, ARow]);
    GridListOfEntries.BeginUpdate;
    isAllEnabled := True;
    WrapInt := Round(GridListOfEntries.Width/7);
    if (SelectedSectionIndex >= 0) {AND (ACol = 2)} then
    begin
      lbCaption.Caption := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'Search Result') + ': '+fmDataMod.RegSections[SelectedSectionIndex].Caption;
      lbDesc.Caption := fmDataMod.RegSections[SelectedSectionIndex].Text;
      GridListOfEntries.RemoveRows(1,GridListOfEntries.RowCount-1);
      LastIndex := 0;
      for i := 0 to Length(fmDataMod.RegErrors)-1 do
      begin
        if not fmDataMod.RegErrors[i].Excluded then
        begin
          if not fmDataMod.RegErrors[i].Fixed then
          begin
            if fmDataMod.RegErrors[i].SectionCode = SelectedSectionIndex then
            begin
              GridListOfEntries.AddRow;
              GridListOfEntries.AddCheckBox(1,LastIndex+1,fmDataMod.RegErrors[i].Enabled,false);
              if not fmDataMod.RegErrors[i].Enabled then isAllEnabled := False;
              GridListOfEntries.GridCells[2, LastIndex+1] := IntToStr(i);
              GridListOfEntries.GridCells[3, LastIndex+1] := '<font size=3>'+fmDataMod.RegSections[SelectedSectionIndex].Caption+': <font color="#0034c7">'+fmDataMod.RegErrors[i].Caption+'</font></font><br>'
                + stringReplace(WrapText(fmDataMod.RegErrors[i].Text, WrapInt), #13#10, '<br>', [rfReplaceAll, rfIgnoreCase]);
              GridListOfEntries.AddImageIdx(3,LastIndex+1,SelectedSectionIndex,haBeforeText,vaTop);
              GridListOfEntries.AutoSizeRow(LastIndex+1);
              inc(LastIndex);
            end;
          end;
        end;
      end;
      GridListOfEntries.AutoSizeColumns(True);
      if LastIndex > 0 then
      begin
        mmOpenKey.Enabled := True;
        mmAddToExclusion.Enabled := True;
      end
      else
      begin
        mmOpenKey.Enabled := False;
        mmAddToExclusion.Enabled := False;
      end;
      if isAllEnabled then GridListOfEntries.SetCheckBoxState(1,0,True)
      else GridListOfEntries.SetCheckBoxState(1,0,False);
    end;
    GridListOfEntries.EndUpdate;
    ListWidth := 10;
    for i := 0 to GridListOfEntries.ColCount-1 do
    begin
      ListWidth := ListWidth + GridListOfEntries.ColWidths[i];
    end;
    AdvScrollBox1.HorzScrollBar.Range := ListWidth;
  end;
end;
//=========================================================



//=========================================================
{ПРИ КЛИКЕ НА СПИСКЕ ПАРАМЕТРОВ И ИХ ЗНАЧЕНЙ ВНИЗУ}
//---------------------------------------------------------
procedure TfmRegistryCleanerResults.GridDetailsClickCell(Sender: TObject; ARow, ACol: Integer);
begin
  lbParamName.Caption := GridDetails.Cells[1,ARow]+':';
  edParamValue.Text := GridDetails.Cells[2,ARow];
end;
//=========================================================




//=========================================================
{ПУНКТ В КОНТЕКСТНОМ МЕНЮ}
//---------------------------------------------------------
procedure TfmRegistryCleanerResults.popupMenuOpenKeyClick(Sender: TObject);
begin
  GridListOfEntries.OnDblClickCell(Self, GridListOfEntries.RealRow, 2);
end;
{ПУНКТ МЕНЮ ДОБАВИТЬ В ИСКЛЮЧЕНИЯ}
procedure TfmRegistryCleanerResults.popupMenuAddToExclusionClick(Sender: TObject);
var
  ErrorIndex, LastExceptionIndex: integer;
  i, j: integer;
  ExceptionFile: TIniFile;
  TempStr: string;
begin
  //сохраняем данное исключение в список исключений и в файл
  ErrorIndex := StrToInt(GridListOfEntries.GridCells[2, GridListOfEntries.RealRow]);
  SetLength(fmDataMod.RegExceptions, Length(fmDataMod.RegExceptions)+1);
  LastExceptionIndex := Length(fmDataMod.RegExceptions)-1;
  if fmDataMod.RegErrors[ErrorIndex].ErrorType = regKeyName then
  begin
    if fmDataMod.RegErrors[ErrorIndex].RootKey = HKEY_CLASSES_ROOT then TempStr := 'HKEY_CLASSES_ROOT';
    if fmDataMod.RegErrors[ErrorIndex].RootKey = HKEY_LOCAL_MACHINE then TempStr := 'HKEY_LOCAL_MACHINE';
    if fmDataMod.RegErrors[ErrorIndex].RootKey = HKEY_CURRENT_USER then TempStr := 'HKEY_CURRENT_USER';
    fmDataMod.RegExceptions[LastExceptionIndex].Text := TempStr + fmDataMod.RegErrors[ErrorIndex].SubKey;
    fmDataMod.RegExceptions[LastExceptionIndex].ExceptionType := etKeyAddr;
  end
  else
  begin
    fmDataMod.RegExceptions[LastExceptionIndex].Text := fmDataMod.RegErrors[ErrorIndex].Parameter;
    fmDataMod.RegExceptions[LastExceptionIndex].ExceptionType := etText;
  end;
  if not DirectoryExists(fmDataMod.PathToUtilityFolder) then SysUtils.ForceDirectories(fmDataMod.PathToUtilityFolder);
  ExceptionFile := TIniFile.Create(fmDataMod.PathToUtilityFolder+'RegExceptions.ini');
  ExceptionFile.EraseSection('etKeyAddr');
  ExceptionFile.EraseSection('etText');
  j := 0;
  for i := 0 to Length(fmDataMod.RegExceptions)-1 do
  begin
    if fmDataMod.RegExceptions[i].ExceptionType = etKeyAddr then
    begin
      ExceptionFile.WriteString('etKeyAddr', IntToStr(j), fmDataMod.RegExceptions[i].Text);
      inc(j);
    end;
  end;
  for i := 0 to Length(fmDataMod.RegExceptions)-1 do
  begin
    if fmDataMod.RegExceptions[i].ExceptionType <> etKeyAddr then
    begin
      ExceptionFile.WriteString('etText', IntToStr(j), fmDataMod.RegExceptions[i].Text);
      inc(j);
    end;
  end;

  //удаляем из списка ошибку и обновляем счётчики:
  inc(fmDataMod.RegErrorsFound, -1);
  inc(fmDataMod.RegSections[SelectedSectionIndex].ErrorsCount, -1);
  fmDataMod.RegErrors[ErrorIndex].Enabled := False;
  fmDataMod.RegErrors[ErrorIndex].Excluded := True;
  GridListOfEntries.RemoveRows(GridListOfEntries.RealRow, GridListOfEntries.RowSelectCount);
  lbStatFound.Caption := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'lbStatFound')+': '+IntToStr(fmDataMod.RegErrorsFound);
  for i := 1 to GridListOfSections.RowCount-1 do
    if StrToInt(GridListOfSections.Cells[2,i])=SelectedSectionIndex then GridListOfSections.Cells[3,i] := IntToStr(fmDataMod.RegSections[SelectedSectionIndex].ErrorsCount);
  for i := 1 to GridCategories.RowCount-1 do
    if StrToInt(GridCategories.Cells[1,i])=SelectedSectionIndex then GridCategories.Cells[2,i] := fmDataMod.RegSections[SelectedSectionIndex].Caption + ' <font color="#22355d">('+IntToStr(fmDataMod.RegSections[SelectedSectionIndex].ErrorsCount)+')</font>';
  GridCategories.Refresh;
end;
//=========================================================



//=========================================================
{КНОПКА - "ОБЩИЕ РЕЗУЛЬТАТЫ ПОИСКА"}
//---------------------------------------------------------
procedure TfmRegistryCleanerResults.btAllCategoriesClick(Sender: TObject);
begin
  lbStatus.Caption := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'ResultsAfter');
  lbStatFound.Caption := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'lbStatFound')+': '+IntToStr(fmDataMod.RegErrorsFound);
  mmOpenKey.Enabled := False;
  mmAddToExclusion.Enabled := False;
  agbErrorsList.Visible := False;
  agbErrorsList.Enabled := False;
  agbGeneralResults.Visible := True;
  agbGeneralResults.Enabled := True;
  btAllCategories.Enabled := agbErrorsList.Enabled;
end;
//=========================================================



//=========================================================
{КНОПКА - "НАЗАД"}
//---------------------------------------------------------
procedure TfmRegistryCleanerResults.btBackClick(Sender: TObject);
begin
  fmDataMod.is_back := True;
  Close;
end;
//=========================================================



//=========================================================
{КНОПКА - "ВЫХОД"}
//---------------------------------------------------------
procedure TfmRegistryCleanerResults.btCloseClick(Sender: TObject);
begin
  fmDataMod.is_back := False;
  Close;
end;
//=========================================================



//=========================================================
{КНОПКА - "ОТМЕНИТЬ СКАНИРОВАНИЕ"}
//---------------------------------------------------------
procedure TfmRegistryCleanerResults.btCancelClick(Sender: TObject);
begin
  fmDataMod.isStop := True;
end;
//=========================================================



//=========================================================
{ФОРМИРОВАНИЕ СПИСКА КАТЕГОРИЙ С ОШИБКАМИ}
//---------------------------------------------------------
procedure TfmRegistryCleanerResults.ShowRegSections;
var
  i, j: integer;
  CatFontColor: string;
begin
  agbScanning.Enabled := False;
  agbScanning.Visible := False;
  j := 1;
  for i := 0 to Length(fmDataMod.RegSections)-1 do
  begin
    GridCategories.Cells[1,j] := IntToStr(i);
    if fmDataMod.RegSections[i].ErrorsCount > 0 then CatFontColor := '#0500da' else CatFontColor := '#22355d';
    GridCategories.Cells[2,j] := fmDataMod.RegSections[i].Caption + ' <font color="'+CatFontColor+'">('+IntToStr(fmDataMod.RegSections[i].ErrorsCount)+')</font>';
    GridListOfSections.AddRow;
    GridListOfSections.AddImageIdx(0,j,i,haCenter,vaTop);
    GridListOfSections.GridCells[1, j] := '<font size="3" color="#0034c7">'
        +fmDataMod.RegSections[i].Caption+'</font><font size="3" color="#464646"> ('
        +IntToStr(fmDataMod.RegSections[i].ErrorsCount)+')</font><br>'
        +fmDataMod.RegSections[i].Text;
    GridListOfSections.GridCells[2, j] := IntToStr(i);
    GridListOfSections.AddButton(1,j,125,23,ReadLangStr('WinTuning_Common.lng', 'Common', 'Details'),haLeft,vaUnderText);
    inc(j);
  end;

  GridCategories.HideColumn(1);
  GridListOfSections.HideColumn(2);
  GridListOfSections.AutoSizeColumns(True);
  GridCategories.AutoSizeColumns(True);
  with GridListOfEntries do
  begin
    AddCheckBox(1,0,False,False);
    HideColumn(2);
  end;

  agbGeneralResults.Enabled := True;
  agbGeneralResults.Visible := True;
  GridCategories.Enabled := True;
end;
//=========================================================



//=========================================================
{КНОПКА "ОЧИСТИТЬ"}
//---------------------------------------------------------
procedure TfmRegistryCleanerResults.btCleanClick(Sender: TObject);
begin
  if not IsCheckedSomething then
  begin
    Application.MessageBox(
           PWideChar(ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'NotSelected')),
           PWideChar(ReadLangStr('WinTuning_Common.lng', 'Common', 'Error')),
           MB_OK + MB_ICONINFORMATION);
    exit;
  end;

  fmMsg := TfmMsg.Create(Application);
  fmMsg.ShowModal;
  fmMsg.Free;
  fmDataMod.is_back := True;
  Close;
end;
//=========================================================



//======================================================
{ЗАГРУЗКА ИЗОБРАЖЕНИЯ В ОБЛАСТЬ НА ФОРМЕ}
//------------------------------------------------------
procedure TfmRegistryCleanerResults.LoadPNG2Prog(dllname, resname: string; imgobj: TImage);
var
  h: THandle;
  ResStream: TResourceStream;
  PNG: TPngImage;
begin
  h:= LoadLibrary(PWideChar(dllname));
  PNG := TPngImage.Create;
  ResStream := TResourceStream.Create(h, resname, rt_RCData);
  PNG.LoadFromStream(ResStream);
  imgobj.Picture.Assign(PNG);
  ResStream.free;
  PNG.Free;
end;
//======================================================



//=========================================================
{ПРИМЕНЕНИЕ СКИНА}
//=========================================================
procedure TfmRegistryCleanerResults.ApplyTheme;
var
  ThemeFileName, StrTmp: string;
begin
  //Языковые настройки
  fmDataMod.RegClean.RootKey := HKEY_CURRENT_USER;
  fmDataMod.RegClean.OpenKey('\Software\WinTuning', true);
  ThemeName := GetProgParam('theme');
  //Разделы
  ThemeFileName := ExtractFilePath(paramstr(0)) + 'Themes\' + ThemeName + '\Theme.ini';
  if FileExists(ThemeFileName) then
  begin
    ThemeFile := TIniFile.Create(ThemeFileName);
  end;

  //Цвет окна
  Color := ReadRBG(PChar(ThemeFile.ReadString('Utilities', 'WndColor', '240,240,240')));

  //Фон заголовка утилиты
//  ThemeImgMainCaption.Visible := ThemeFile.ReadBool('Utilities', 'UtilCaptionBackImageShow', True);
  StrTmp := ThemePathConvert(PChar(ThemeFile.ReadString('Utilities', 'UtilCaptionBackImagePath', '')), PChar(ThemeName));
  if FileExists(StrTmp) then ThemeImgMainCaption.Picture.LoadFromFile(StrTmp);
  //Цвет шрифта заголовка
  albLogoText.Font.Color := ReadRBG(PChar(ThemeFile.ReadString('Utilities', 'UtilCaptionFontColor', '53,65,79')));
  //Цвет фона загловка утилиты в случае, если НЕТ картинки
  ThemeShapeMainCaption.Brush.Color := ReadRBG(PChar(ThemeFile.ReadString('Utilities', 'UtilCaptionBackgroundColor', '243,243,243')));
  lbStatus.Color := ThemeShapeMainCaption.Brush.Color; //статус восстановления
  //Цвет борюда загловка утилиты
  ThemeShapeMainCaption.Pen.Color := ReadRBG(PChar(ThemeFile.ReadString('Utilities', 'UtilCaptionBorderColor', '210,220,227')));
  ShapeLeft.Pen.Color := ReadRBG(PChar(ThemeFile.ReadString('Utilities', 'UtilCaptionBorderColor', '210,220,227')));
  ShapeBottom.Pen.Color := ReadRBG(PChar(ThemeFile.ReadString('Utilities', 'UtilCaptionBorderColor', '210,220,227')));
  agbLogo.BorderColor := ReadRBG(PChar(ThemeFile.ReadString('Utilities', 'UtilCaptionBorderColor', '210,220,227')));

  //Нижний бордюрчик
  ShapeBottom.Brush.Color := ReadRBG(PChar(ThemeFile.ReadString('Utilities', 'BorderBottomColor', '243,245,248')));

  ThemeFile.Free;
end;
//=========================================================




//=========================================================
{ГЛАВНОЕ МЕНЮ}
//---------------------------------------------------------
procedure TfmRegistryCleanerResults.mmExitClick(Sender: TObject);
var
  MenuName: string;
begin
  MenuName := TMenuItem(Sender).Name;
  //Файл -> Исправить (очисить)
  if MenuName = 'mmFix' then btClean.OnClick(Self);
  //Файл -> Сканировать
  if MenuName = 'mmNewScan' then btBack.OnClick(Self);
  //Файл -> Настройки
  if MenuName = 'mmSettings' then
  begin
    fmSettings := TfmSettings.Create(Application);
    fmSettings.ShowModal;
    fmSettings.Free;
  end;

  //Файл -> Нашли ошибку?
  if MenuName = 'mmFoundAnError' then ShellExecute(Handle, 'open', PChar(ExtractFilePath(paramstr(0))+'ErrorReport.exe'), nil, nil, SW_SHOW);
  //Файл -> Выход
  if MenuName = 'mmExit' then Close;
  //Выделенное -> Открыть ключ в программе Regedit
  if MenuName = 'mmOpenKey' then popupMenuOpenKey.OnClick(Self);
  //Выделенное -> Добавить в исключения
  if MenuName = 'mmAddToExclusion' then popupMenuAddToExclusion.OnClick(Self);
  //Справка -> Открыть справку
  if MenuName = 'mmOpenHelp' then ViewHelp('UtilRegistryCleaner');
  //Справка -> WebSite
  if MenuName = 'mmWebSite' then ShellExecute(handle,'open', GetProgParam('webindex'),nil,nil,SW_SHOW);
  //Справка -> О программе
  if MenuName = 'mmAbout' then CreateTheForm(PChar(Caption), PChar(paramstr(0)), '');
end;
//=========================================================



//=========================================================
{ПРИ НАЖАТИИ НА ГОРЯЧИЕ КЛАВИШИ}
procedure TfmRegistryCleanerResults.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  if Ord(Key) = 112 then //F1
  begin
    ViewHelp('UtilRegistryCleaner');
  end;
end;
//=========================================================



//=========================================================
{ОБНОВИТЬ СТАТУС ТЕКУЩЕЙ ПРОВЕРКИ}
//---------------------------------------------------------
procedure TfmRegistryCleanerResults.UpdateStatus;
begin
  if Assigned(fmRegistryCleanerResults) then
  begin
    fmRegistryCleanerResults.lbStatFound.Caption := fmDataMod.LNGChecking + ': ' +inttostr(fmDataMod.RegErrorsFound);
    fmRegistryCleanerResults.lbStatus.Caption := fmDataMod.LNGCheckingKey+': ' + fmDataMod.ActKeyName;
  end;
end;
//=========================================================



//=========================================================
{ДОБАВЛЯЕМ СУЩЕСТВУЮЩИЕ РАЗДЕЛЫ СЛЕВА}
//---------------------------------------------------------
procedure TfmRegistryCleanerResults.LoadSections;
var
  i: Integer;
begin
  for i := 0 to Length(fmDataMod.RegSections)-1 do
  begin
    GridCategories.AddRow;
    GridCategories.AddImageIdx(2,i+1,i,haBeforeText,vaCenter);
    GridCategories.GridCells[1, i+1] := IntToStr(i);
    GridCategories.GridCells[2, i+1] := fmDataMod.RegSections[i].Caption;
  end;
  GridCategories.HideColumn(1);
  GridCategories.AutoSizeColumns(True);
end;
//=========================================================



//=========================================================
{ПРИ ЗАКРЫТИИ ФОРМЫ}
//---------------------------------------------------------
procedure TfmRegistryCleanerResults.FormClose(Sender: TObject; var Action: TCloseAction);
var
  RegCl: TRegistry;
begin
  RegCl := TRegistry.Create;
  RegCl.RootKey := HKEY_CURRENT_USER;
  RegCl.OpenKey('\Software\WinTuning\RegistryCleaner', True);
  TRY
    RegCl.WriteInteger('PanelErrorDetailsHeight', PanelErrorDetails.Height);
    RegCl.WriteInteger('PanelLeftWidth', PanelLeft.Width);
  EXCEPT

  END;
  RegCl.CloseKey;
  RegCl.Free;
  SaveWndPosition(fmRegistryCleanerResults, 'fmRegistryCleanerResults');
end;
//=========================================================



//=========================================================
{ЯЗЫКОВЫЙ ФАЙЛ}
//---------------------------------------------------------
procedure TfmRegistryCleanerResults.ApplyLang;
begin
  Caption := 'WinTuning: '+ ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'Registry Cleaner');
  albLogoText.Caption :=              ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'Registry Cleaner');
  lbCaption.Caption :=                ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'Search Result');
  lbDesc.Caption :=                   ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'Processing');
  btBack.Caption :=                   ReadLangStr('WinTuning_Common.lng', 'Common', 'Back');
  btClean.Caption :=                  ReadLangStr('WinTuning_Common.lng', 'Common', 'Next');
  btCancel.Caption :=                 ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'Stop');
  lbStatFound.Caption :=              ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'lbStatFound')+': 0';
  lbStatFixed.Caption :=              ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'lbStatFixed')+': 0';
  btAllCategories.Caption :=          ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'Search Result');
  GridListOfEntries.ColumnHeaders.Strings[3] := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'ColumnDescription');
  GridCategories.ColumnHeaders.Strings[1] := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'Category');
  GridDetails.ColumnHeaders.Strings[1] := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'Parameter');
  GridDetails.ColumnHeaders.Strings[2] := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'Value');
  lbDetailsCaption.Caption :=         ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'Result_lbDetailsCaption');
  btClose.Caption :=                  ReadLangStr('WinTuning_Common.lng', 'Common', 'Exit');

  //Главное меню - вторая форма
  mmMainFile.Caption :=                              ReadLangStr('WinTuning_Common.lng', 'Common', 'File');
  mmFoundAnError.Caption :=                          ReadLangStr('WinTuning_Common.lng', 'Common', 'ErrorReport');
  mmSettings.Caption :=                              ReadLangStr('WinTuning_Common.lng', 'Common', 'Settings');
  mmExit.Caption :=                                  ReadLangStr('WinTuning_Common.lng', 'Common', 'Exit');
  mmMainHelp.Caption :=                              ReadLangStr('WinTuning_Common.lng', 'Common', 'Help');
  mmOpenHelp.Caption :=                              ReadLangStr('WinTuning_Common.lng', 'Common', 'OpenHelp');
  mmWebSite.Caption :=                               ReadLangStr('WinTuning_Common.lng', 'Common', 'WebSite');
  mmAbout.Caption :=                                 ReadLangStr('WinTuning_Common.lng', 'Common', 'About');
  mmFix.Caption :=                                   btClean.Caption;
  mmNewScan.Caption :=                               btBack.Caption;
  mmSelected.Caption :=                              ReadLangStr('WinTuning_Common.lng', 'Common', 'Selected');
  mmOpenKey.Caption :=                               ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'mmOpenKey');
  mmAddToExclusion.Caption :=                        ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'mmAddToExclusion');
  popupMenuOpenKey.Caption :=                        ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'mmOpenKey');
  popupMenuAddToExclusion.Caption :=                 ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'mmAddToExclusion');
end;
//=========================================================



//=========================================================
{АКТИВАЦИЯ ФОРМЫ}
//---------------------------------------------------------
procedure TfmRegistryCleanerResults.FormActivate(Sender: TObject);
begin
  if not isLoaded then RestoreWndPosition(fmRegistryCleanerResults, 'fmRegistryCleanerResults');
  isLoaded := True;
end;
//=========================================================



//=========================================================
{ПРИ СОЗДАНИИ ФОРМЫ}
//---------------------------------------------------------
procedure TfmRegistryCleanerResults.FormCreate(Sender: TObject);
var
  i: integer;
  Screen: TScreen;
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

  SelectedSectionIndex := -1;
  SelectedErrorIndex := -1;

  ApplyLang;

  fspTaskbarMgrProgress.Active := False;
  fspTaskbarMgrProgress.AppId := 'MyAppID_' + IntToStr(Application.Handle);
  fspTaskbarMgrProgress.Active := True;
  fspTaskbarMgrProgress.ProgressState := fstpsNormal;
  last_index := 1;
  ApplyTheme;

  agbScanning.Enabled := True;
  agbScanning.Visible := True;
  GridCategories.Enabled := False;

  fmDataMod.RegClean.RootKey := HKEY_CURRENT_USER;
  fmDataMod.RegClean.OpenKey('\Software\WinTuning\RegistryCleaner', True);
  if fmDataMod.RegClean.ValueExists('PanelErrorDetailsHeight') then PanelErrorDetails.Height := fmDataMod.RegClean.ReadInteger('PanelErrorDetailsHeight');
  if fmDataMod.RegClean.ValueExists('PanelLeftWidth') then PanelLeft.Width := fmDataMod.RegClean.ReadInteger('PanelLeftWidth');
  LoadSections;
end;
//=========================================================



end.

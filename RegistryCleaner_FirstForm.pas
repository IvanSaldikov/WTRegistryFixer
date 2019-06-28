unit RegistryCleaner_FirstForm;

{$WARN UNIT_PLATFORM OFF}
{$WARN SYMBOL_PLATFORM OFF}


interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms, Math, ShellAPI,
  Dialogs, AdvSmoothLabel, acPNG, JvGIF, ExtCtrls, AdvGlowButton, AdvChartPaneEditor, DateUtils,
  AdvChartView, AdvChartViewGDIP, Grids, BaseGrid, AdvGrid, StdCtrls, sLabel, HTMLabel,
  dbhtmlab, AdvGroupBox, JvComponentBase, JvComputerInfoEx, Registry, IniFiles, FindFile,
  AdvChart, AdvChartGDIP, AdvChartUtil, ImgList, fspTaskbarMgr, PNGImage, typinfo,
  AdvObj, Menus, JvBaseDlg, JvBrowseFolder, sHintManager, ComCtrls, AdvListV,
  acAlphaImageList;

type
  TfmRegistryCleaner1 = class(TForm)
    ShapeLeft: TShape;
    ShapeBottom: TShape;
    agbLogo: TAdvGroupBox;
    ThemeImgMainCaption: TImage;
    albLogoText: TsLabel;
    imgLogoBack: TImage;
    imgWinTuning: TImage;
    ThemeShapeMainCaption: TShape;
    MainMenuFirstWindow: TMainMenu;
    mmMainFile: TMenuItem;
    mmFoundAnError: TMenuItem;
    mmExit: TMenuItem;
    mmMainHelp: TMenuItem;
    mmOpenHelp: TMenuItem;
    mmQuickHelp: TMenuItem;
    mmAbout: TMenuItem;
    mmScan: TMenuItem;
    mmWebSite: TMenuItem;
    ThemeImgLeft: TImage;
    ThemeImgLeftTemp: TImage;
    agbSettings: TAdvGroupBox;
    ThemeShapeSubCaptionSettings: TShape;
    ThemeImgSubCaptionSettings: TImage;
    lbAdditional: TsLabel;
    TaskDialogAny: TTaskDialog;
    PanelCommon: TPanel;
    agbHelp: TAdvGroupBox;
    shapeDescription: TShape;
    imgQuickHelp: TImage;
    lbInfoDescription: TDBHTMLabel;
    lbQuickHelp: TLabel;
    imgCloseQuickTip: TImage;
    PanelBottom: TPanel;
    lbChooseScanSections: TLabel;
    PanelTop: TPanel;
    GridSectionList: TAdvStringGrid;
    btStart: TAdvGlowButton;
    btExit: TAdvGlowButton;
    btSettings: TAdvGlowButton;
    btRestore: TAdvGlowButton;
    AlphaImgsSections16: TsAlphaImageList;
    function IsCheckedSomething: boolean;
    procedure FormCreate(Sender: TObject);
    procedure btStartClick(Sender: TObject);
    procedure ShowHideHelp(isShow: Boolean);
    procedure LoadPNG2Prog(dllname, resname: string; imgobj: TImage);
    procedure mmExitClick(Sender: TObject);
    procedure imgQuickHelpClick(Sender: TObject);
    procedure imgCloseQuickTipClick(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure ApplyTheme;
    procedure btExitClick(Sender: TObject);
    procedure LoadSections;
    procedure ThemeUpdateLeftIMG;
    procedure btSettingsClick(Sender: TObject);
    procedure btRestoreClick(Sender: TObject);
    procedure ApplyLang;
    procedure FormActivate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);

  private
    { Private declarations }
    IsLeftImgLoaded: Boolean; //загружена ли левая картинка в компонент ThemeImgLeftTemp - чтобы её оттуда каждый раз читать при изменении размера окна
    isLoaded: boolean; //Переменная нужна, чтобы не было изменений настроек при запуске программы
  public
    { Protected declarations }
  protected
    { Protected declarations }
    procedure CreateParams(var Params: TCreateParams); override;
  end;

var
  fmRegistryCleaner1: TfmRegistryCleaner1;

implementation

uses DataMod, RegistryCleaner_Settings, RegistryCleaner_RegistryRestore;

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

{ВЫВОДИМ НАЗВАНИЕ ВЕРСИИ WINTUNING: [XP, VISTER, 7]}
function GetCapInfo(WTVerID, info_id: integer): shortstring; external 'Functions.dll';

{ЧТЕНИЕ ЯЗЫКОВОЙ СТРОКИ ИЗ ОПРЕДЕЛЁННОГО ФАЙЛА}
function ReadLangStr(FileName, Section, Caption: PChar): PChar; external 'Functions.dll';

{ПЕРЕВОД ИЗ БАЙТОВ В НОРМАЛЬНУЮ СТРОКУ}
function BytesToStr(const i64Size: Int64): PChar; external 'Functions.dll';

{ПРЕОБРАЗОВАНИЕ СТРОКИ ВИДА R,G,B В TCOLOR}
function ReadRBG(instr: PChar): TColor; external 'Functions.dll';

{ПРЕОБРАЗОВАНИЕ СТРОКИ ВИДА %WinTuning_PATH% В C:\Program Files\WinTuning 7}
function ThemePathConvert(InStr, InThemeName: PChar): PChar; external 'Functions.dll';

{ОТКРЫТЬ ОПРЕДЕЛЁННЫЙ РАЗДЕЛ СПРАВКИ}
procedure ViewHelp(page: PChar); stdcall; external 'Functions.dll';

{О ПРОГРАММЕ}
function CreateTheForm(S1, S2, S3: PChar): integer; stdcall export; external 'About.dll';

{ЧТЕНИЕ НАСТРОЕК ПРОГРАММЫ}
function GetProgParam(paramname: PChar): PChar; external 'Functions.dll';



//=========================================================
{ОТОБРАЖАЕТ ФОРМУ НА ПАНЕЛИ ЗАДАЧ}
procedure TfmRegistryCleaner1.CreateParams(var Params: TCreateParams);
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
{ОТМЕЧЕН ЛИ ХОТЯ БЫ ОДНА СЕКЦИЯ}
//---------------------------------------------------------
function TfmRegistryCleaner1.IsCheckedSomething: boolean;
var
  I: Integer;
  State: boolean;
begin
  State := False;
  Result := False;
  for I := 1 to GridSectionList.RowCount - 1 do
  begin
    GridSectionList.GetCheckBoxState(1,I,State);
    if State then
    begin
      Result := True;
      exit;
    end;
  end;
end;
//=========================================================



//=========================================================
{ЗАМОСТИТЬ ЛЕВУЮ ПАНЕЛЬ КАРТИНКОЙ}
//---------------------------------------------------------
procedure TfmRegistryCleaner1.ThemeUpdateLeftIMG;
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
{НАЖАТИЕ НА ССЫКЛУ ЗАКРЫТИЯ БЫСТРОЙ ПОДСКАЗКИ}
procedure TfmRegistryCleaner1.imgCloseQuickTipClick(Sender: TObject);
begin
  ShowHideHelp(not agbHelp.Visible);
end;
{НАЖАТИЕ НА ЗНАК ВОПРОСА - ОТКРЫТИЕ ХЕЛПА}
procedure TfmRegistryCleaner1.imgQuickHelpClick(Sender: TObject);
begin
  ViewHelp('UtilRegistryCleaner');
end;
{ПРИ НАЖАТИИ НА ГОРЯЧИЕ КЛАВИШИ}
procedure TfmRegistryCleaner1.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  if Ord(Key) = 112 then //F1
  begin
    imgQuickHelp.OnClick(Self);
  end;
end;
//=========================================================



//=========================================================
{КНОПКА "ЗАКРЫТЬ"}
//---------------------------------------------------------
procedure TfmRegistryCleaner1.btExitClick(Sender: TObject);
begin
  fmDataMod.is_back := False;
  Close;
end;
//=========================================================



//=========================================================
{КНОПКА "ВОССТАНОВЛЕНИЕ"}
//---------------------------------------------------------
procedure TfmRegistryCleaner1.btRestoreClick(Sender: TObject);
begin
  fmRegistryRestore := TfmRegistryRestore.Create(Application);
  fmRegistryRestore.ShowModal;
  fmRegistryRestore.Free;
end;
//=========================================================




//=========================================================
{КНОПКА - "НАСТРОЙКИ"}
//---------------------------------------------------------
procedure TfmRegistryCleaner1.btSettingsClick(Sender: TObject);
begin
  fmSettings := TfmSettings.Create(Application);
  fmSettings.ShowModal;
  fmSettings.Free;
end;
//=========================================================



//=========================================================
{КНОПКА - "НАЧАТЬ СКАНИРОВАНИЕ"}
//---------------------------------------------------------
procedure TfmRegistryCleaner1.btStartClick(Sender: TObject);
var
  i: Integer;
  State: Boolean;
  CapTemp: string;
begin
  if not IsCheckedSomething then
  begin
    Application.MessageBox(
           PWideChar(stringReplace(ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'NotSelected'), ':', '', [])),
           PWideChar(ReadLangStr('WinTuning_Common.lng', 'Common', 'Error')),
           MB_OK + MB_ICONINFORMATION);
    exit;
  end;

  BEGIN
    fmDataMod.RegClean.RootKey := HKEY_CURRENT_USER;
    fmDataMod.RegClean.OpenKey('\Software\WinTuning\RegistryCleaner\RegSections\DefaultChecked', True);
    for i := 0 to Length(fmDataMod.RegSections)-1 do
    begin
      GridSectionList.GetCheckBoxState(1, i+1, State);
      fmDataMod.RegClean.WriteBool(IntToStr(i), State);
      fmDataMod.RegSections[i].Enabled := State;
    end;
    fmDataMod.RegClean.RootKey := HKEY_CURRENT_USER;
    fmDataMod.RegClean.OpenKey('\Software\WinTuning\RegistryCleaner', True);
    fmDataMod.RegClean.WriteDateTime('RegistryCleanerLastDate',now);
    fmDataMod.isSecondWindowShowNeeded := True;
    fmDataMod.isStop := false;
    fmRegistryCleaner1.Close;
  END;
end;
//=========================================================



//=========================================================
{ЗАПОЛНЯЕМ СУЩЕСТВУЮЩИЕ РАЗДЕЛЫ}
//---------------------------------------------------------
procedure TfmRegistryCleaner1.LoadSections;
var
  i: Integer;
  DefState, isAllChecked: Boolean;
begin
  isAllChecked := True;
  for i := 0 to Length(fmDataMod.RegSections)-1 do
  begin
    GridSectionList.AddRow;
    DefState := fmDataMod.RegSections[i].Enabled;
    GridSectionList.AddImageIdx(2,i+1,i,haBeforeText,vaCenter);
    GridSectionList.AddCheckBox(1,i+1,DefState,false);
    if not DefState then isAllChecked := False;
    GridSectionList.GridCells[2, i+1] := fmDataMod.RegSections[i].Caption;
  end;
  GridSectionList.AddCheckBox(1,0,isAllChecked,false);
  GridSectionList.AutoSizeColumns(True);
end;
//=========================================================



//======================================================
{СКРЫТИЕ И РАСКРЫТИЕ ХЕЛПА}
procedure TfmRegistryCleaner1.ShowHideHelp(isShow: Boolean);
var
 Reg1: TRegistry;
begin
  Reg1 := TRegistry.Create();
  Reg1.RootKey := HKEY_CURRENT_USER;
  Reg1.OpenKey('\Software\WinTuning\RegistryCleaner', True);
  agbHelp.Enabled := isShow;
  agbHelp.Visible := isShow;
  Reg1.WriteBool('isHelpHidden', not isShow);
  PanelTop.Visible := isShow;
  PanelTop.Enabled := isShow;
{  if not isShow then
  begin
    gbSectionSelection.Height := gbSectionSelection.Height + agbHelp.Height;
    gbSectionSelection.Top := 0;
  end
  else
  begin
    gbSectionSelection.Top := agbHelp.Height + 3;
    gbSectionSelection.Height := gbSectionSelection.Height - agbHelp.Height;
  end;}
  Reg1.Free;
end;
//======================================================




//======================================================
{ЗАГРУЗКА ИЗОБРАЖЕНИЯ В ОБЛАСТЬ НА ФОРМЕ}
//------------------------------------------------------
procedure TfmRegistryCleaner1.LoadPNG2Prog(dllname, resname: string; imgobj: TImage);
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
procedure TfmRegistryCleaner1.ApplyTheme;
var
  ThemeFileName, StrTmp, ThemeName: string;
  ThemeFile: TIniFile; //ini-файл темы оформления
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

    //Цвет окна
    Color := ReadRBG(PChar(ThemeFile.ReadString('Utilities', 'WndColor', '240,240,240')));
    //Фон логотипа
    imgLogoBack.Visible := ThemeFile.ReadBool('Utilities', 'LogoBackImageShow', True);
    StrTmp := ThemePathConvert(PChar(ThemeFile.ReadString('Utilities', 'LogoBackImagePath', '')), PChar(ThemeName));
    if FileExists(StrTmp) then imgLogoBack.Picture.LoadFromFile(StrTmp);

    //Фон заголовка утилиты
    ThemeImgMainCaption.Visible := ThemeFile.ReadBool('Utilities', 'UtilCaptionBackImageShow', True);
    StrTmp := ThemePathConvert(PChar(ThemeFile.ReadString('Utilities', 'UtilCaptionBackImagePath', '')), PChar(ThemeName));
    if FileExists(StrTmp) then ThemeImgMainCaption.Picture.LoadFromFile(StrTmp);
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

    //Левая картинка (левый фон)
    ThemeImgLeft.Visible := ThemeFile.ReadBool('Utilities', 'ImageLeftShow', False);
    StrTmp := ThemePathConvert(PChar(ThemeFile.ReadString('Utilities', 'ImageLeftPath', '')), PChar(ThemeName));
    if FileExists(StrTmp) then
    begin
      IsLeftImgLoaded := True;
      ThemeImgLeftTemp.Picture.LoadFromFile(StrTmp);
    end;

    //Фон заголовков слева
    ThemeImgSubCaptionSettings.Visible := ThemeFile.ReadBool('Utilities', 'LeftCaptionsBackImageShow', True);
    StrTmp := ThemePathConvert(PChar(ThemeFile.ReadString('Utilities', 'LeftCaptionsBackImagePath', '')), PChar(ThemeName));
    if FileExists(StrTmp) then ThemeImgSubCaptionSettings.Picture.LoadFromFile(StrTmp);
    //Цвет шрифта заголовков слева
    lbAdditional.Font.Color := ReadRBG(PChar(ThemeFile.ReadString('Utilities', 'LeftCaptionsFontColor', '65,85,105')));
    //Цвет фона загловков слева в случае, если НЕТ картинки
    ThemeShapeSubCaptionSettings.Brush.Color := ReadRBG(PChar(ThemeFile.ReadString('Utilities', 'LeftCaptionsBackgroundColor', '243,245,248')));
    //Цвет борюда загловков слева - и рамок тоже
    ThemeShapeSubCaptionSettings.Pen.Color := ReadRBG(PChar(ThemeFile.ReadString('Utilities', 'LeftCaptionBorderColor', '210,220,227')));
    agbSettings.BorderColor := ReadRBG(PChar(ThemeFile.ReadString('Utilities', 'LeftCaptionBorderColor', '210,220,227')));

    //Нижний бордюрчик
    ShapeBottom.Brush.Color := ReadRBG(PChar(ThemeFile.ReadString('Utilities', 'BorderBottomColor', '243,245,248')));

    //Фон для подсказки
    shapeDescription.Visible := ThemeFile.ReadBool('Utilities', 'QuickTipBackShow', True);
    shapeDescription.Brush.Color := ReadRBG(PChar(ThemeFile.ReadString('Utilities', 'QuickTipBackColorBrush', '251,252,253')));
    lbQuickHelp.Color := ReadRBG(PChar(ThemeFile.ReadString('Utilities', 'QuickTipBackColorBrush', '251,252,253')));
    lbInfoDescription.Color := ReadRBG(PChar(ThemeFile.ReadString('Utilities', 'QuickTipBackColorBrush', '251,252,253')));
    shapeDescription.Pen.Color := ReadRBG(PChar(ThemeFile.ReadString('Utilities', 'QuickTipBackColorPen', '197,214,224')));
    StrTmp := ThemeFile.ReadString('Utilities', 'QuickTipBackPenType', '0');
    if StrTmp = '0' then shapeDescription.Pen.Style := psSolid;
    if StrTmp = '1' then shapeDescription.Pen.Style := psDot;
    if StrTmp = '2' then shapeDescription.Pen.Style := psDash;
    StrTmp := ThemePathConvert(PChar(ThemeFile.ReadString('Utilities', 'QuickTipQuestionImagePath', '')), PChar(ThemeName));
    if FileExists(StrTmp) then imgQuickHelp.Picture.LoadFromFile(StrTmp);
    StrTmp := ThemePathConvert(PChar(ThemeFile.ReadString('Utilities', 'QuickTipCloseImagePath', '')), PChar(ThemeName));
    if FileExists(StrTmp) then imgCloseQuickTip.Picture.LoadFromFile(StrTmp);

    ThemeFile.Free;
  end;
end;
//=========================================================



//=========================================================
{ГЛАВНОЕ МЕНЮ}
//---------------------------------------------------------
procedure TfmRegistryCleaner1.mmExitClick(Sender: TObject);
var
  MenuName: string;
begin
  MenuName := TMenuItem(Sender).Name;
  //Файл -> Сканировать
  if MenuName = 'mmScan' then btStart.OnClick(Self);
  //Файл -> Нашли ошибку?
  if MenuName = 'mmFoundAnError' then ShellExecute(Handle, 'open', PChar(ExtractFilePath(paramstr(0)) + 'ErrorReport.exe'), nil, nil, SW_SHOW);
  //Файл -> Выход
  if MenuName = 'mmExit' then Close;
  //Справка -> Открыть справку
  if MenuName = 'mmOpenHelp' then imgQuickHelp.OnClick(Self);
  //Справка -> Быстрая подсказка
  if MenuName = 'mmQuickHelp' then ShowHideHelp(not agbHelp.Visible);
  //Справка -> WebSite
  if MenuName = 'mmWebSite' then ShellExecute(handle,'open', GetProgParam('webindex'),nil,nil,SW_SHOW);
  //Справка -> О программе
  if MenuName = 'mmAbout' then CreateTheForm(PChar(Caption), PChar(paramstr(0)), '');
end;
//=========================================================



//=========================================================
{ПРИМЕНЕНИЕ ЯЗЫКА}
procedure TfmRegistryCleaner1.ApplyLang;
begin
  {Применение языкового файла}
  Caption :=                                    'WinTuning: '+ ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'Registry Cleaner');
  albLogoText.Caption :=                        ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'Registry Cleaner');
  GridSectionList.ColumnHeaders.Strings[2] :=   ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'Category');
  lbAdditional.Caption :=                       ReadLangStr('WinTuning_Common.lng', 'Common', 'Additional');
  btSettings.Caption :=                         ReadLangStr('WinTuning_Common.lng', 'Common', 'Settings');
  btRestore.Caption :=                          ReadLangStr('WinTuning_Common.lng', 'Common', 'Restore');
  btStart.Caption :=                            ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'Start');
  lbChooseScanSections.Caption :=               ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'lbChooseScanSections');
  btExit.Caption :=                             ReadLangStr('WinTuning_Common.lng', 'Common', 'Exit');
  lbQuickHelp.Caption :=                        ReadLangStr('WinTuning_Common.lng', 'Common', 'QuickHelp');
  imgCloseQuickTip.Hint :=                      ReadLangStr('WinTuning_Common.lng', 'Common', 'CloseQuickHelp');
  imgQuickHelp.Hint :=                          ReadLangStr('WinTuning_Common.lng', 'Common', 'OpenHelp');
  lbQuickHelp.Hint :=                           ReadLangStr('WinTuning_Common.lng', 'Common', 'OpenHelp');

  //Главное меню - первая форма
  mmMainFile.Caption :=                         ReadLangStr('WinTuning_Common.lng', 'Common', 'File');
  mmFoundAnError.Caption :=                     ReadLangStr('WinTuning_Common.lng', 'Common', 'ErrorReport');
  mmExit.Caption :=                             ReadLangStr('WinTuning_Common.lng', 'Common', 'Exit');
  mmMainHelp.Caption :=                         ReadLangStr('WinTuning_Common.lng', 'Common', 'Help');
  mmOpenHelp.Caption :=                         ReadLangStr('WinTuning_Common.lng', 'Common', 'OpenHelp');
  mmQuickHelp.Caption :=                        ReadLangStr('WinTuning_Common.lng', 'Common', 'QuickHelp');
  mmWebSite.Caption :=                          ReadLangStr('WinTuning_Common.lng', 'Common', 'WebSite');
  mmAbout.Caption :=                            ReadLangStr('WinTuning_Common.lng', 'Common', 'About');
  mmScan.Caption :=                             btStart.Caption;
  lbInfoDescription.HTMLText.Text :=            ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'lbInfoDescription');
end;
//=========================================================



//=========================================================
{ПРИ ЗАКРЫТИИ ФОРМЫ}
//---------------------------------------------------------
procedure TfmRegistryCleaner1.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  SaveWndPosition(fmRegistryCleaner1, 'fmRegistryCleaner1');
end;
//=========================================================



//=========================================================
{АКТИВАЦИЯ ФОРМЫ}
//---------------------------------------------------------
procedure TfmRegistryCleaner1.FormActivate(Sender: TObject);
begin
  if not isLoaded then RestoreWndPosition(fmRegistryCleaner1, 'fmRegistryCleaner1');
  isLoaded := True;
end;
//=========================================================



//=========================================================
{СОЗДАНИЕ ФОРМЫ}
//---------------------------------------------------------
procedure TfmRegistryCleaner1.FormCreate(Sender: TObject);
var
  I: Integer;
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

  fmDataMod.RegClean.RootKey := HKEY_CURRENT_USER;
  fmDataMod.RegClean.OpenKey('\Software\WinTuning', True);

  ApplyLang;
  ApplyTheme;
  LoadSections();

  with fmDataMod do
  begin
    RegClean.RootKey := HKEY_CURRENT_USER;
    RegClean.OpenKey('\Software\WinTuning\RegistryCleaner', True);
    if RegClean.ValueExists('isHelpHidden') then
     if RegClean.ReadBool('isHelpHidden') then ShowHideHelp(False);
    RegClean.RootKey:=HKEY_CURRENT_USER;
    RegClean.OpenKey('\Software\WinTuning', True);
  end;

  ThemeUpdateLeftIMG;
end;
//=========================================================









end.

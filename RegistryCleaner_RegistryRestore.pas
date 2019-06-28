unit RegistryCleaner_RegistryRestore;

interface

{$WARN UNIT_PLATFORM OFF}
{$WARN SYMBOL_PLATFORM OFF}

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms, ShellAPI, typinfo, PNGImage, IniFiles, ShlObj,
  ClipBrd, Dialogs, ImgList, acAlphaImageList, Grids, AdvObj, BaseGrid, AdvGrid, StdCtrls, AdvGlowButton, sLabel, AdvGroupBox,
  ExtCtrls, AdvSplitter, FindFile, Menus, ActiveX, JvBaseDlg, JvSHFileOperation, Registry;

type
  TfmRegistryRestore = class(TForm)
    ShapeLeft: TShape;
    ThemeImgLeft: TImage;
    ShapeBottom: TShape;
    ThemeImgLeftTemp: TImage;
    lbRegRestore: TLabel;
    agbLogo: TAdvGroupBox;
    imgLogoBack: TImage;
    ThemeShapeMainCaption: TShape;
    ThemeImgMainCaption: TImage;
    albLogoText: TsLabel;
    imgWinTuning: TImage;
    gbRegFiles: TGroupBox;
    AlphaImgsRegFiles16: TsAlphaImageList;
    URLRegFileInfo: TLabel;
    PanelRegFilesList: TPanel;
    lbRegRestoreDescription: TLabel;
    PanelBottom: TPanel;
    lbRegFileContent: TLabel;
    MemoRegFile: TMemo;
    btRestore: TButton;
    lbRegFilesCount: TLabel;
    FindFile1: TFindFile;
    popupMenuRegFiles: TPopupMenu;
    popupMenuRun: TMenuItem;
    popupMenuCopy: TMenuItem;
    popupMenuCut: TMenuItem;
    popupMenuDelete: TMenuItem;
    popupMenuProperties: TMenuItem;
    JvSHFileOperation1: TJvSHFileOperation;
    GridRegFilesList: TAdvStringGrid;
    SplitterRegFiles: TSplitter;
    btCancel: TAdvGlowButton;
    procedure URLRegFileInfoClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure LoadPNG2Prog(dllname, resname: string; imgobj: TImage);
    procedure ApplyTheme;
    procedure ThemeUpdateLeftIMG;
    procedure FindFile1FileMatch(Sender: TObject; const FileInfo: TFileDetails);
    procedure GridRegFilesListSelectCell(Sender: TObject; ACol, ARow: Integer; var CanSelect: Boolean);
    procedure btRestoreClick(Sender: TObject);
    procedure btCancelClick(Sender: TObject);
    procedure popupMenuRunClick(Sender: TObject);
    procedure GridRegFilesListDblClickCell(Sender: TObject; ARow, ACol: Integer);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure ApplyLang;
    procedure FormActivate(Sender: TObject);
  private
    { Private declarations }
    IsLeftImgLoaded: Boolean; //загружена ли левая картинка в компонент ThemeImgLeftTemp - чтобы её оттуда каждый раз читать при изменении размера окна
    RegFilesCount: integer;
    isLoaded: boolean; //Переменная нужна, чтобы не было изменений настроек при запуске программы
  public
    { Public declarations }
  protected
    { Protected declarations }
    procedure CreateParams(var Params: TCreateParams); override;
  end;

var
  fmRegistryRestore: TfmRegistryRestore;

implementation

uses DataMod;

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



//=========================================================
{ОТОБРАЖАЕТ ФОРМУ НА ПАНЕЛИ ЗАДАЧ}
//---------------------------------------------------------
procedure TfmRegistryRestore.CreateParams(var Params: TCreateParams);
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
{КОПИРОВАТЬ/ВЫРЕЗАТЬ В БУФЕР ОБМЕНА ФАЙЛЫ}
//---------------------------------------------------------
procedure CopyFilesToClipboard(const AFiles: TStringList; MoveFiles: Boolean = false);
var
 DropFiles: PDropFiles;
 hGlobal: THandle;
 iLen: Integer;
 f: Cardinal;
 d: PCardinal;
 FileList: string;
begin
 Clipboard.Open;
 FileList := StringReplace(AFiles.Text, #13#10, #0, [rfReplaceAll]);
 FileList := trim(FileList) + #0#0;
 iLen := Length(FileList) * SizeOf(char);
 hGlobal := GlobalAlloc(GMEM_SHARE or GMEM_MOVEABLE or GMEM_ZEROINIT, SizeOf(TDropFiles) + iLen);
 Win32Check(hGlobal <> 0);
 DropFiles := GlobalLock(hGlobal);
 DropFiles^.pFiles := SizeOf(TDropFiles);
 {$IFDEF UNICODE}
  DropFiles^.fWide := true;
 {$ENDIF}
 Move(FileList[1], (PansiChar(DropFiles) + SizeOf(TDropFiles))^, iLen);
 SetClipboardData(CF_HDROP, hGlobal);
 GlobalUnlock(hGlobal);
 if MoveFiles then
   begin
     f := RegisterClipboardFormat(CFSTR_PREFERREDDROPEFFECT);
     hGlobal := GlobalAlloc(GMEM_SHARE or GMEM_MOVEABLE or GMEM_ZEROINIT, sizeof(dword));
     d := PCardinal(GlobalLock(hGlobal));
     d^ := DROPEFFECT_MOVE;
     SetClipboardData(f, hGlobal);
     GlobalUnlock(hGlobal);
   end;
 Clipboard.Close;
end;
//=========================================================



//=========================================================
{ПРОЦЕДУРА ОТОБРАЖЕНИЯ СВОЙСТВ ФАЙЛА}
//---------------------------------------------------------
procedure ShowFileProperties(const FileName: string);
var
  sei: TShellExecuteinfo;
begin
  ZeroMemory(@sei,sizeof(sei));
  sei.cbSize := sizeof(TShellExecuteinfo);
  sei.lpFile := PChar(FileName);
  sei.lpVerb := 'properties';
  sei.fMask := SEE_MASK_INVOKEIDLIST;
  ShellExecuteEx(@sei);
end;
//=========================================================



//=========================================================
{ЗАГРУЗКА ИЗОБРАЖЕНИЯ В ОБЛАСТЬ НА ФОРМЕ}
//---------------------------------------------------------
procedure TfmRegistryRestore.LoadPNG2Prog(dllname, resname: string; imgobj: TImage);
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
{ПРИМЕНЕНИЕ СКИНА}
//=========================================================
procedure TfmRegistryRestore.ApplyTheme;
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
{ЗАМОСТИТЬ ЛЕВУЮ ПАНЕЛЬ КАРТИНКОЙ}
//---------------------------------------------------------
procedure TfmRegistryRestore.ThemeUpdateLeftIMG;
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
{ПРИ НАХОЖДЕНИИ REG-ФАЙЛА}
//---------------------------------------------------------
procedure TfmRegistryRestore.FindFile1FileMatch(Sender: TObject; const FileInfo: TFileDetails);
begin
  GridRegFilesList.AddRow;
  GridRegFilesList.Cells[1,GridRegFilesList.RowCount-1] := FileInfo.Name;
  GridRegFilesList.Cells[2,GridRegFilesList.RowCount-1] := DateTimeToStr(FileInfo.CreatedTime);
  GridRegFilesList.AddImageIdx(1,GridRegFilesList.RowCount-1,0,haBeforeText,vaCenter);
  GridRegFilesList.AutoSizeColumns(True);
  inc(RegFilesCount);
end;
//=========================================================



//=========================================================
{ПРИ ВЫБОРЕ ЯЧЕЙКИ REG-ФАЙЛА}
//---------------------------------------------------------
procedure TfmRegistryRestore.GridRegFilesListSelectCell(Sender: TObject; ACol, ARow: Integer; var CanSelect: Boolean);
var
  NewPopupMenuStatus: boolean;
begin
  NewPopupMenuStatus := False;
  if ARow > 0 then
    if FileExists(fmDataMod.PathToUtilityFolder+'\Backups\'+GridRegFilesList.Cells[1,ARow]) then
    begin
      MemoRegFile.Lines.LoadFromFile(fmDataMod.PathToUtilityFolder+'\Backups\'+GridRegFilesList.Cells[1,ARow]);
      NewPopupMenuStatus := True;
      btRestore.Enabled := True;
    end;
  popupMenuRun.Enabled := NewPopupMenuStatus;
  popupMenuCopy.Enabled := NewPopupMenuStatus;
  popupMenuCut.Enabled := NewPopupMenuStatus;
  popupMenuDelete.Enabled := NewPopupMenuStatus;
  popupMenuProperties.Enabled := NewPopupMenuStatus;
end;
//=========================================================



//=========================================================
{ПРИ ДВОЙНОМ КЛИКЕ ПО ЯЧЕЙКЕ REG-ФАЙЛА}
//---------------------------------------------------------
procedure TfmRegistryRestore.GridRegFilesListDblClickCell(Sender: TObject; ARow, ACol: Integer);
begin
  if btRestore.Enabled then btRestore.OnClick(Self);
end;
//=========================================================



//=========================================================
{КНОПКА "ЗАПУСТИТЬ"}
//---------------------------------------------------------
procedure TfmRegistryRestore.btRestoreClick(Sender: TObject);
begin
  if GridRegFilesList.RealRow > 0 then
    ShellExecute(handle,'open', PChar(fmDataMod.PathToUtilityFolder+'\Backups\'+GridRegFilesList.Cells[1,GridRegFilesList.RealRow]),nil,nil,SW_NORMAL)
end;
//=========================================================



//=========================================================
{КНОПКА "ОТМЕНА"}
//---------------------------------------------------------
procedure TfmRegistryRestore.btCancelClick(Sender: TObject);
begin
  Close;
end;
//=========================================================



//=========================================================
{ССЫЛКА "ПОДРОБНЕЕ О REG-ФАЙЛАХ"}
//---------------------------------------------------------
procedure TfmRegistryRestore.URLRegFileInfoClick(Sender: TObject);
var
  WEBSiteStr: string;
begin
  WEBSiteStr := 'http://en.wikipedia.org/wiki/Windows_Registry#.REG_files';
  if GetProgParam('lang')='Russian' then WEBSiteStr := 'http://ru.wikipedia.org/wiki/REG';
  ShellExecute(handle,'open', PChar(WEBSiteStr),nil,nil,SW_SHOW)
end;
//=========================================================



//=========================================================
{КОНТЕКСТНОЕ МЕНЮ СПИСКА REG-ФАЙЛОВ}
//---------------------------------------------------------
procedure TfmRegistryRestore.popupMenuRunClick(Sender: TObject);
var
  MenuName: string;
  FilesList: TStringList;
  i: integer;
  IsCut: boolean;
  MsgCaption:string;
  CnSelect: Boolean;
begin
  MenuName := TMenuItem(Sender).Name;
  //Запустить
  if MenuName = 'popupMenuRun' then btRestore.OnClick(Self);
  //Копировать /
  if (MenuName = 'popupMenuCopy') OR (MenuName = 'popupMenuCut') then
  begin
    FilesList := TStringList.Create;
    for i := 1 to GridRegFilesList.RowCount - 1 do
    begin
      if GridRegFilesList.RowSelect[i] then FilesList.Add(fmDataMod.PathToUtilityFolder+'\Backups\' + GridRegFilesList.Cells[1,i]);
    end;
    IsCut := False;
    if MenuName = 'popupMenuCut' then IsCut := True;
    CopyFilesToClipboard(FilesList, IsCut); //вырезать в буфер
    FilesList.Free;
  end;
  //Удалить
  if MenuName = 'popupMenuDelete' then
  begin
    MsgCaption := ReadLangStr('WinTuning_Common.lng', 'Common', 'Confirmation');
    if Application.MessageBox(ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'RegFileDelete'),
                            PChar(MsgCaption),
                            MB_YESNO + MB_ICONQUESTION)
       = IDYES then
    begin
      for i := 1 to GridRegFilesList.RowCount - 1 do
      begin
        if GridRegFilesList.RowSelect[i] then JvSHFileOperation1.SourceFiles.Add(fmDataMod.PathToUtilityFolder+'\Backups\' + GridRegFilesList.Cells[1,i]);
      end;
      JvSHFileOperation1.Operation := foDelete;
      JvSHFileOperation1.Execute; //удалить в корзину
      Sleep(500);
      inc(RegFilesCount,-GridRegFilesList.SelectedRowCount);
      GridRegFilesList.RemoveRows(GridRegFilesList.RealRowIndex(GridRegFilesList.RealRow),GridRegFilesList.SelectedRowCount);
      lbRegFilesCount.Caption := stringReplace(ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'RegFilesCounter'), '%1', IntToStr(RegFilesCount), [rfReplaceAll,rfIgnoreCase]);
      if RegFilesCount=0 then
      begin
        btRestore.Enabled := False;
        GridRegFilesList.OnSelectCell(Self, 0, 0, CnSelect);
        MemoRegFile.Lines.Clear;
      end;
    end;
  end;
  //Свойства
  if MenuName = 'popupMenuProperties' then ShowFileProperties(fmDataMod.PathToUtilityFolder+'\Backups\'+GridRegFilesList.Cells[1,GridRegFilesList.RealRow]);
end;
//=========================================================



//=========================================================
{ПРИ ЗАКРЫТИИ ФОРМЫ}
//---------------------------------------------------------
procedure TfmRegistryRestore.FormClose(Sender: TObject; var Action: TCloseAction);
var
  RegCl: TRegistry;
begin
  RegCl := TRegistry.Create;
  RegCl.RootKey := HKEY_CURRENT_USER;
  RegCl.OpenKey('\Software\WinTuning\RegistryCleaner', True);
  TRY
    RegCl.WriteInteger('PanelBottomHeight', PanelBottom.Height);
  EXCEPT

  END;
  RegCl.CloseKey;
  RegCl.Free;
  SaveWndPosition(fmRegistryRestore, 'fmRegistryRestore');
end;
//=========================================================



//=========================================================
{ПРИМЕНЕНИЕ ЯЗЫКА}
//---------------------------------------------------------
procedure TfmRegistryRestore.ApplyLang;
begin
  Caption :=                      ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'fmRegistryRestore');
  albLogoText.Caption :=          ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'Registry Cleaner');
  lbRegRestore.Caption :=         ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'fmRegistryRestore');
  gbRegFiles.Caption :=           ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'gbRegFiles');
  lbRegRestoreDescription.Caption := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'lbRegRestoreDescription');
  lbRegFileContent.Caption :=     ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'lbRegFileContent');
  btRestore.Caption :=            ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'btRestore');
  URLRegFileInfo.Caption :=       ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'URLRegFileInfo');
  GridRegFilesList.ColumnHeaders.Strings[1] := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'GridRegFilesList_1');
  GridRegFilesList.ColumnHeaders.Strings[2] := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'GridRegFilesList_2');
  btCancel.Caption :=             ReadLangStr('WinTuning_Common.lng', 'Common', 'Cancel');
end;
//=========================================================



//=========================================================
{АКТИВАЦИЯ ФОРМЫ}
//---------------------------------------------------------
procedure TfmRegistryRestore.FormActivate(Sender: TObject);
begin
  if not isLoaded then RestoreWndPosition(fmRegistryRestore, 'fmRegistryRestore');
  isLoaded := True;
end;
//=========================================================



//=========================================================
{ПРИ СОЗДАНИИ ФОРМЫ}
//---------------------------------------------------------
procedure TfmRegistryRestore.FormCreate(Sender: TObject);
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

  ApplyTheme;
  ApplyLang;
  RegFilesCount := 0;
  FindFile1.Criteria.Files.Location := fmDataMod.PathToUtilityFolder+'\Backups\';
  FindFile1.Execute;
  lbRegFilesCount.Caption := stringReplace(ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'RegFilesCounter'), '%1', IntToStr(RegFilesCount), [rfReplaceAll,rfIgnoreCase]);
  btRestore.Enabled := False;

  ThemeUpdateLeftIMG;

  fmDataMod.RegClean.RootKey := HKEY_CURRENT_USER;
  fmDataMod.RegClean.OpenKey('\Software\WinTuning\RegistryCleaner', True);
  if fmDataMod.RegClean.ValueExists('PanelBottomHeight') then PanelBottom.Height := fmDataMod.RegClean.ReadInteger('PanelBottomHeight');
end;
//=========================================================




end.

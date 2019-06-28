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
    IsLeftImgLoaded: Boolean; //çàãğóæåíà ëè ëåâàÿ êàğòèíêà â êîìïîíåíò ThemeImgLeftTemp - ÷òîáû å¸ îòòóäà êàæäûé ğàç ÷èòàòü ïğè èçìåíåíèè ğàçìåğà îêíà
    RegFilesCount: integer;
    isLoaded: boolean; //Ïåğåìåííàÿ íóæíà, ÷òîáû íå áûëî èçìåíåíèé íàñòğîåê ïğè çàïóñêå ïğîãğàììû
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


{ÑÎÕĞÀÍÈÒÜ ÏÎÇÈÖÈŞ ÎÊÍÀ ÏĞÈ ÇÀÊĞÛÒÈÈ}
procedure SaveWndPosition(FormName: TForm; KeyToSave: PChar); external 'Functions.dll';

{ÂÎÑÑÒÀÍÎÂÈÒÜ ÏÎÇÈÖÈŞ ÎÊÍÀ ÏĞÈ ÇÀÏÓÑÊÅ}
procedure RestoreWndPosition(FormName: TForm; KeyToSave: PChar); external 'Functions.dll';

{ÂÛÂÎÄÈÌ ĞÀÇÍÓŞ ÈÍÔÓ Î ÂÅĞÑÈÈ WINTUNING: ÈÍÄÅÊÑ, ÃÎÄ ÎÊÎÍ×ÀÍÈß ÈÑÏÎËÜÇÎÂÀÍÈß È ÒÄ.}
function GetWTVerInfo(info_id: integer): integer; external 'Functions.dll';

{×ÒÅÍÈÅ ÍÀÑÒĞÎÅÊ ÏĞÎÃĞÀÌÌÛ}
function GetProgParam(paramname: PChar): PChar; external 'Functions.dll';

{ÏĞÅÎÁĞÀÇÎÂÀÍÈÅ ÑÒĞÎÊÈ ÂÈÄÀ R,G,B Â TCOLOR}
function ReadRBG(instr: PChar): TColor; external 'Functions.dll';

{ÎÒÊĞÛÒÜ ÎÏĞÅÄÅË¨ÍÍÛÉ ĞÀÇÄÅË ÑÏĞÀÂÊÈ}
procedure ViewHelp(page: PChar); stdcall; external 'Functions.dll';

{ÏĞÅÎÁĞÀÇÎÂÀÍÈÅ ÑÒĞÎÊÈ ÂÈÄÀ %WinTuning_PATH% Â C:\Program Files\WinTuning 7}
function ThemePathConvert(InStr, InThemeName: PChar): PChar; external 'Functions.dll';

{×ÒÅÍÈÅ ßÇÛÊÎÂÎÉ ÑÒĞÎÊÈ ÈÇ ÎÏĞÅÄÅË¨ÍÍÎÃÎ ÔÀÉËÀ}
function ReadLangStr(FileName, Section, Caption: PChar): PChar; external 'Functions.dll';



//=========================================================
{ÎÒÎÁĞÀÆÀÅÒ ÔÎĞÌÓ ÍÀ ÏÀÍÅËÈ ÇÀÄÀ×}
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
{ÊÎÏÈĞÎÂÀÒÜ/ÂÛĞÅÇÀÒÜ Â ÁÓÔÅĞ ÎÁÌÅÍÀ ÔÀÉËÛ}
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
{ÏĞÎÖÅÄÓĞÀ ÎÒÎÁĞÀÆÅÍÈß ÑÂÎÉÑÒÂ ÔÀÉËÀ}
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
{ÇÀÃĞÓÇÊÀ ÈÇÎÁĞÀÆÅÍÈß Â ÎÁËÀÑÒÜ ÍÀ ÔÎĞÌÅ}
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
{ÏĞÈÌÅÍÅÍÈÅ ÑÊÈÍÀ}
//=========================================================
procedure TfmRegistryRestore.ApplyTheme;
var
  ThemeFileName, StrTmp, ThemeName: string;
  ThemeFile: TIniFile; //ini-ôàéë òåìû îôîğìëåíèÿ
begin
  //ßçûêîâûå íàñòğîéêè
  ThemeName := GetProgParam('theme');
  //Ğàçäåëû
  ThemeFileName := ExtractFilePath(paramstr(0)) + 'Themes\' + ThemeName + '\Theme.ini';
  if SysUtils.FileExists(ThemeFileName) then
  begin
    ThemeFile := TIniFile.Create(ThemeFileName);

    //Öâåò îêíà
    Color := ReadRBG(PChar(ThemeFile.ReadString('Utilities', 'WndColor', '240,240,240')));
    //Ôîí ëîãîòèïà
    imgLogoBack.Visible := ThemeFile.ReadBool('Utilities', 'LogoBackImageShow', True);
    StrTmp := ThemePathConvert(PChar(ThemeFile.ReadString('Utilities', 'LogoBackImagePath', '')), PChar(ThemeName));
    if SysUtils.FileExists(StrTmp) then imgLogoBack.Picture.LoadFromFile(StrTmp);

    //Ëåâàÿ êàğòèíêà (ëåâûé ôîí)
    ThemeImgLeft.Visible := ThemeFile.ReadBool('Utilities', 'ImageLeftShow', False);
    StrTmp := ThemePathConvert(PChar(ThemeFile.ReadString('Utilities', 'ImageLeftPath', '')), PChar(ThemeName));
    if SysUtils.FileExists(StrTmp) then
    begin
      IsLeftImgLoaded := True;
      ThemeImgLeftTemp.Picture.LoadFromFile(StrTmp);
    end;

    //Ôîí çàãîëîâêà óòèëèòû
    ThemeImgMainCaption.Visible := ThemeFile.ReadBool('Utilities', 'UtilCaptionBackImageShow', True);
    StrTmp := ThemePathConvert(PChar(ThemeFile.ReadString('Utilities', 'UtilCaptionBackImagePath', '')), PChar(ThemeName));
    if SysUtils.FileExists(StrTmp) then ThemeImgMainCaption.Picture.LoadFromFile(StrTmp);
    //Öâåò øğèôòà çàãîëîâêà
    albLogoText.Font.Color := ReadRBG(PChar(ThemeFile.ReadString('Utilities', 'UtilCaptionFontColor', '53,65,79')));
    //Öâåò ôîíà çàãëîâêà óòèëèòû â ñëó÷àå, åñëè ÍÅÒ êàğòèíêè
    ThemeShapeMainCaption.Brush.Color := ReadRBG(PChar(ThemeFile.ReadString('Utilities', 'UtilCaptionBackgroundColor', '243,243,243')));
    //Öâåò áîğşäà çàãëîâêà óòèëèòû
    ThemeShapeMainCaption.Brush.Color := ReadRBG(PChar(ThemeFile.ReadString('Utilities', 'UtilCaptionBackgroundColor', '243,243,243')));
    ThemeShapeMainCaption.Pen.Color := ReadRBG(PChar(ThemeFile.ReadString('Utilities', 'UtilCaptionBorderColor', '210,220,227')));
    ShapeLeft.Pen.Color := ReadRBG(PChar(ThemeFile.ReadString('Utilities', 'UtilCaptionBorderColor', '210,220,227')));
    ShapeBottom.Pen.Color := ReadRBG(PChar(ThemeFile.ReadString('Utilities', 'UtilCaptionBorderColor', '210,220,227')));
    agbLogo.BorderColor := ReadRBG(PChar(ThemeFile.ReadString('Utilities', 'UtilCaptionBorderColor', '210,220,227')));

    //Íèæíèé áîğäşğ÷èê
    ShapeBottom.Brush.Color := ReadRBG(PChar(ThemeFile.ReadString('Utilities', 'BorderBottomColor', '243,245,248')));

    ThemeFile.Free;
  end;
end;
//=========================================================



//=========================================================
{ÇÀÌÎÑÒÈÒÜ ËÅÂÓŞ ÏÀÍÅËÜ ÊÀĞÒÈÍÊÎÉ}
//---------------------------------------------------------
procedure TfmRegistryRestore.ThemeUpdateLeftIMG;
var
  x,y: integer; // ëåâûé âåğõíèé óãîë êàğòèíêè
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
{ÏĞÈ ÍÀÕÎÆÄÅÍÈÈ REG-ÔÀÉËÀ}
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
{ÏĞÈ ÂÛÁÎĞÅ ß×ÅÉÊÈ REG-ÔÀÉËÀ}
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
{ÏĞÈ ÄÂÎÉÍÎÌ ÊËÈÊÅ ÏÎ ß×ÅÉÊÅ REG-ÔÀÉËÀ}
//---------------------------------------------------------
procedure TfmRegistryRestore.GridRegFilesListDblClickCell(Sender: TObject; ARow, ACol: Integer);
begin
  if btRestore.Enabled then btRestore.OnClick(Self);
end;
//=========================================================



//=========================================================
{ÊÍÎÏÊÀ "ÇÀÏÓÑÒÈÒÜ"}
//---------------------------------------------------------
procedure TfmRegistryRestore.btRestoreClick(Sender: TObject);
begin
  if GridRegFilesList.RealRow > 0 then
    ShellExecute(handle,'open', PChar(fmDataMod.PathToUtilityFolder+'\Backups\'+GridRegFilesList.Cells[1,GridRegFilesList.RealRow]),nil,nil,SW_NORMAL)
end;
//=========================================================



//=========================================================
{ÊÍÎÏÊÀ "ÎÒÌÅÍÀ"}
//---------------------------------------------------------
procedure TfmRegistryRestore.btCancelClick(Sender: TObject);
begin
  Close;
end;
//=========================================================



//=========================================================
{ÑÑÛËÊÀ "ÏÎÄĞÎÁÍÅÅ Î REG-ÔÀÉËÀÕ"}
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
{ÊÎÍÒÅÊÑÒÍÎÅ ÌÅÍŞ ÑÏÈÑÊÀ REG-ÔÀÉËÎÂ}
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
  //Çàïóñòèòü
  if MenuName = 'popupMenuRun' then btRestore.OnClick(Self);
  //Êîïèğîâàòü /
  if (MenuName = 'popupMenuCopy') OR (MenuName = 'popupMenuCut') then
  begin
    FilesList := TStringList.Create;
    for i := 1 to GridRegFilesList.RowCount - 1 do
    begin
      if GridRegFilesList.RowSelect[i] then FilesList.Add(fmDataMod.PathToUtilityFolder+'\Backups\' + GridRegFilesList.Cells[1,i]);
    end;
    IsCut := False;
    if MenuName = 'popupMenuCut' then IsCut := True;
    CopyFilesToClipboard(FilesList, IsCut); //âûğåçàòü â áóôåğ
    FilesList.Free;
  end;
  //Óäàëèòü
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
      JvSHFileOperation1.Execute; //óäàëèòü â êîğçèíó
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
  //Ñâîéñòâà
  if MenuName = 'popupMenuProperties' then ShowFileProperties(fmDataMod.PathToUtilityFolder+'\Backups\'+GridRegFilesList.Cells[1,GridRegFilesList.RealRow]);
end;
//=========================================================



//=========================================================
{ÏĞÈ ÇÀÊĞÛÒÈÈ ÔÎĞÌÛ}
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
{ÏĞÈÌÅÍÅÍÈÅ ßÇÛÊÀ}
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
{ÀÊÒÈÂÀÖÈß ÔÎĞÌÛ}
//---------------------------------------------------------
procedure TfmRegistryRestore.FormActivate(Sender: TObject);
begin
  if not isLoaded then RestoreWndPosition(fmRegistryRestore, 'fmRegistryRestore');
  isLoaded := True;
end;
//=========================================================



//=========================================================
{ÏĞÈ ÑÎÇÄÀÍÈÈ ÔÎĞÌÛ}
//---------------------------------------------------------
procedure TfmRegistryRestore.FormCreate(Sender: TObject);
var
  i: integer;
begin
  isLoaded := False;
  // ÄÅËÀÅÌ ÔÎĞÌÓ ÎÄÈÍÀÊÎÂÎÉ ÏÎ ĞÀÇÌÅĞÓ ÏĞÈ ĞÀÇËÈ×ÍÛÕ ĞÀÑĞÅØÅÍÈßÕ È ĞÀÇÌÅĞÀÕ ØĞÈÔÒÀ
  scaled := True;
  Screen := TScreen.Create(nil);
  for i := componentCount - 1 downto 0 do
    with components[i] do
    begin
       if GetPropInfo(ClassInfo, 'font') <> nil then Font.Size := (ScreenWidth div screen.width) * Font.Size;
    end;

  //ÇÀÃĞÓÆÀÅÌ ÃĞÀÔÈÊÓ
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

unit RegistryCleaner_Settings;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms, IniFiles,
  Dialogs, StdCtrls, Grids, AdvObj, BaseGrid, AdvGrid, AdvGlowButton, ExtCtrls;

type
  TfmSettings = class(TForm)
    lbRegExclusions: TLabel;
    GridRegException: TAdvStringGrid;
    btAddRegExclusion: TButton;
    btRemoveRegExclusion: TButton;
    ShapeBottom: TShape;
    btOK: TAdvGlowButton;
    btCancel: TAdvGlowButton;
    btDefault: TButton;
    procedure btCancelClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure btOKClick(Sender: TObject);
    procedure btAddRegExclusionClick(Sender: TObject);
    procedure btRemoveRegExclusionClick(Sender: TObject);
    procedure btDefaultClick(Sender: TObject);
    procedure ApplyLang;
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  fmSettings: TfmSettings;

implementation

uses DataMod, RegistryCleaner_RegExceptions;

{$R *.dfm}


{◊“≈Õ»≈ ﬂ«€ Œ¬Œ… —“–Œ » »« Œœ–≈ƒ≈À®ÕÕŒ√Œ ‘¿…À¿}
function ReadLangStr(FileName, Section, Caption: PChar): PChar; external 'Functions.dll';


//=========================================================
{ ÕŒœ ¿ "Œ“Ã≈Õ¿"}
//---------------------------------------------------------
procedure TfmSettings.btCancelClick(Sender: TObject);
begin
  Close;
end;
//=========================================================



//=========================================================
{ ÕŒœ ¿ "OK"}
//---------------------------------------------------------
procedure TfmSettings.btOKClick(Sender: TObject);
var
  i, j: integer;
  ExceptionFile: TIniFile;
begin
  if not DirectoryExists(fmDataMod.PathToUtilityFolder) then SysUtils.ForceDirectories(fmDataMod.PathToUtilityFolder);
  ExceptionFile := TIniFile.Create(fmDataMod.PathToUtilityFolder+'RegExceptions.ini');
  ExceptionFile.EraseSection('etKeyAddr');
  ExceptionFile.EraseSection('etText');
  fmDataMod.RegExceptions := nil;
  for i := 1 to GridRegException.RowCount-1 do
  begin
    SetLength(fmDataMod.RegExceptions, Length(fmDataMod.RegExceptions)+1);
    if GridRegException.Cells[1, i] = ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'RegKey') then fmDataMod.RegExceptions[i-1].ExceptionType := etKeyAddr
    else fmDataMod.RegExceptions[i-1].ExceptionType := etText;
    fmDataMod.RegExceptions[i-1].Text := GridRegException.Cells[2, i];
  end;
  j := 0;
  for i := 1 to GridRegException.RowCount-1 do
  begin
    if GridRegException.Cells[1, i] = ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'RegKey') then
    begin
      ExceptionFile.WriteString('etKeyAddr', IntToStr(j), fmDataMod.RegExceptions[i-1].Text);
      inc(j);
    end;
  end;
  for i := 1 to GridRegException.RowCount-1 do
  begin
    if GridRegException.Cells[1, i] <> ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'RegKey') then
    begin
      ExceptionFile.WriteString('etText', IntToStr(j), fmDataMod.RegExceptions[i-1].Text);
      inc(j);
    end;
  end;
  Close;
end;
//=========================================================



//=========================================================
{ ÕŒœ ¿ "ƒŒ¡¿¬»“‹"}
//---------------------------------------------------------
procedure TfmSettings.btAddRegExclusionClick(Sender: TObject);
begin
  fmRegExceptionsAdd := TfmRegExceptionsAdd.Create(Application);
  fmRegExceptionsAdd.ShowModal;
  fmRegExceptionsAdd.Free;
end;
//=========================================================



//=========================================================
{ ÕŒœ ¿ "”ƒ¿À»“‹"}
//---------------------------------------------------------
procedure TfmSettings.btRemoveRegExclusionClick(Sender: TObject);
var
  i, SelectedRowIndex : integer;
begin
  i := 1;
  SelectedRowIndex := 0;
  while i <= GridRegException.RowCount - 1 do
  begin
    if GridRegException.RowSelect[i] then
    begin
      SelectedRowIndex := i;
      break;
    end;
    inc(i);
  end;
  GridRegException.RemoveRows(SelectedRowIndex, GridRegException.RowSelectCount);
end;
//=========================================================



//=========================================================
{ ÕŒœ ¿ "œŒ ”ÃŒÀ◊¿Õ»ﬁ"}
//---------------------------------------------------------
procedure TfmSettings.btDefaultClick(Sender: TObject);
var
  i:integer;
  MsgCaption:string;
begin
  MsgCaption := ReadLangStr('WinTuning_Common.lng', 'Common', 'Confirmation');
  if Application.MessageBox(ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'RegExclusionsConfirmationQuery'),
                          PChar(MsgCaption),
                          MB_YESNO + MB_ICONQUESTION)
     = IDYES then
  begin
    GridRegException.RemoveRows(1,GridRegException.RowCount-1);
    if FileExists(fmDataMod.PathToUtilityFolder+'RegExceptions.ini') then DeleteFile(fmDataMod.PathToUtilityFolder+'RegExceptions.ini');
    fmDataMod.FormRegExceptions;
    for i := 0 to Length(fmDataMod.RegExceptions)-1 do
    begin
      GridRegException.AddRow;
      if fmDataMod.RegExceptions[i].ExceptionType = etKeyAddr then GridRegException.Cells[1, i+1] := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'RegKey')
      else GridRegException.Cells[1, i+1] := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'Parameter');
      GridRegException.Cells[2, i+1] := fmDataMod.RegExceptions[i].Text;
    end;
    GridRegException.AutoSizeColumns(True);
  end;
end;
//=========================================================



//=========================================================
{œ–»Ã≈Õ≈Õ»≈ ﬂ«€ ¿}
//---------------------------------------------------------
procedure TfmSettings.ApplyLang;
begin
  Caption :=                           ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'fmSettings');
  lbRegExclusions.Caption :=           ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'lbRegExclusions');
  GridRegException.ColumnHeaders.Strings[1] :=   ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'GridRegException_1');
  GridRegException.ColumnHeaders.Strings[2] :=   ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'GridRegException_2');
  btAddRegExclusion.Caption :=         ReadLangStr('WinTuning_Common.lng', 'Common', 'Add');
  btRemoveRegExclusion.Caption :=      ReadLangStr('WinTuning_Common.lng', 'Common', 'Delete');
  btDefault.Caption :=                 ReadLangStr('WinTuning_Common.lng', 'Common', 'Default');
  btCancel.Caption :=                  ReadLangStr('WinTuning_Common.lng', 'Common', 'Cancel');
end;
//=========================================================



//=========================================================
{—Œ«ƒ¿Õ»≈ ‘Œ–Ã€}
//---------------------------------------------------------
procedure TfmSettings.FormCreate(Sender: TObject);
var
  i:integer;
begin
  ApplyLang;
  for i := 0 to Length(fmDataMod.RegExceptions)-1 do
  begin
    GridRegException.AddRow;
    if fmDataMod.RegExceptions[i].ExceptionType = etKeyAddr then GridRegException.Cells[1, i+1] := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'RegKey')
    else GridRegException.Cells[1, i+1] := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'Parameter');
    GridRegException.Cells[2, i+1] := fmDataMod.RegExceptions[i].Text;
  end;
  GridRegException.AutoSizeColumns(True);
end;
//=========================================================





end.

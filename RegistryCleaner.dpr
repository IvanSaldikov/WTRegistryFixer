program RegistryCleaner;

{$R *.dres}

uses
  Forms,
  SysUtils,
  Windows,
  Dialogs,
  previnst,
  Registry,
  RegistryCleaner_Results in 'RegistryCleaner_Results.pas' {fmRegistryCleanerResults},
  RegistryCleaner_FirstForm in 'RegistryCleaner_FirstForm.pas' {fmRegistryCleaner1},
  DataMod in 'DataMod.pas' {fmDataMod: TDataModule},
  RegistryCleaner_Msg in 'RegistryCleaner_Msg.pas' {fmMsg},
  RegistryCleaner_Settings in 'RegistryCleaner_Settings.pas' {fmSettings},
  RegistryCleaner_RegExceptions in 'RegistryCleaner_RegExceptions.pas' {fmRegExceptionsAdd},
  RegistryCleaner_RegistryRestore in 'RegistryCleaner_RegistryRestore.pas' {fmRegistryRestore};

{$R *.res}
{$R '../../UACResources/UAC_Manifest.res'}

label L1, LRe;

{������� ������ ���� � ������ WINTUNING: ������, ��� ��������� ������������� � ��.}
function GetWTVerInfo(info_id: integer): integer; external 'Functions.dll';

{������� �������� ������ WINTUNING: [XP, VISTER, 7]}
function GetCapInfo(WTVerID, info_id: integer): shortstring; external 'Functions.dll';

{������ �������� ������ �� ������˨����� �����}
function ReadLangStr(FileName, Section, Caption: PChar): PChar; external 'Functions.dll';

{���� �� ���� ����� �� ������� ��� ���������� (����, �� ������� ����� �������� �����)}
function IsGoToWebSite: boolean; external 'Functions.dll';


var
  WindowCaption, formCaption: string;
  IsOKCommon: boolean;
begin
  //=========================================================
  //��������� ������ ����� ��������� ��� �������
  if IsAlreadyRunning then
  begin
    WindowCaption := 'WinTuning: '+ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'Registry Cleaner');
    formCaption := 'TfmRegistryCleaner1';
    ShowWindow(FindWindow(PWideChar(formCaption), PWideChar(WindowCaption)), SW_restore);
    SetForegroundWindow(FindWindow(PWideChar(formCaption), PWideChar(WindowCaption)));
    goto L1;
  end;
  //=========================================================



  //=========================================================
  //��������� ����� �� ����������?
  if IsGoToWebSite then
  begin
    Application.MessageBox(ReadLangStr('WinTuning_Common.lng', 'Common', 'RestrictToRun'),
                           PWideChar('WinTuning'),
                           MB_OK);
    goto L1;
  end;
  //=========================================================



  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.Title := WindowCaption;
//  Application.Run;
  LRe:
  fmDataMod := TfmDataMod.Create(nil); //������, � ���. ���������� ��� ���������� ��� ��������

  //���� ������ ������
  fmRegistryCleaner1 := TfmRegistryCleaner1.Create(nil);
  fmRegistryCleaner1.ShowModal;
  fmRegistryCleaner1.Free;
  if fmDataMod.isSecondWindowShowNeeded then
  begin
    fmRegistryCleanerResults := TfmRegistryCleanerResults.Create(nil);
    fmDataMod.JvThreadScan.Execute(nil);
    fmRegistryCleanerResults.ShowModal;
  end;


  if fmDataMod.is_back then
  begin
    fmDataMod.Free;
    goto LRe;
  end;
  fmDataMod.Free;

  L1:
end.

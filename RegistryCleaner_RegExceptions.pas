unit RegistryCleaner_RegExceptions;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls;

type
  TfmRegExceptionsAdd = class(TForm)
    edExclusionString: TEdit;
    gbExcludeFrom: TGroupBox;
    rbKeyAddr: TRadioButton;
    rbText: TRadioButton;
    lbExclusionString: TLabel;
    ShapeBottom: TShape;
    btOK: TButton;
    btCancel: TButton;
    procedure btCancelClick(Sender: TObject);
    procedure btOKClick(Sender: TObject);
    procedure ApplyLang;
    procedure FormCreate(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  fmRegExceptionsAdd: TfmRegExceptionsAdd;

implementation

uses RegistryCleaner_Settings;

{$R *.dfm}


{◊“≈Õ»≈ ﬂ«€ Œ¬Œ… —“–Œ » »« Œœ–≈ƒ≈À®ÕÕŒ√Œ ‘¿…À¿}
function ReadLangStr(FileName, Section, Caption: PChar): PChar; external 'Functions.dll';


//=========================================================
{ ÕŒœ ¿ "Œ“Ã≈Õ¿"}
//---------------------------------------------------------
procedure TfmRegExceptionsAdd.btCancelClick(Sender: TObject);
begin
  Close;
end;
//=========================================================



//=========================================================
{œ–»Ã≈Õ≈Õ»≈ ﬂ«€ ¿}
//---------------------------------------------------------
procedure TfmRegExceptionsAdd.ApplyLang;
begin
  Caption :=                      ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'fmRegExceptionsAdd');
  lbExclusionString.Caption :=    ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'lbExclusionString');
  gbExcludeFrom.Caption :=        ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'gbExcludeFrom');
  rbKeyAddr.Caption :=            ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'RegKey');
  rbText.Caption :=               ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'Parameter');
  btCancel.Caption :=             ReadLangStr('WinTuning_Common.lng', 'Common', 'Cancel');
end;
//=========================================================



//=========================================================
{ ÕŒœ ¿ "OK"}
//---------------------------------------------------------
procedure TfmRegExceptionsAdd.btOKClick(Sender: TObject);
var
  LstIndex: integer;
begin
  fmSettings.GridRegException.AddRow;
  LstIndex := fmSettings.GridRegException.RowCount - 1;
  if rbKeyAddr.Checked then fmSettings.GridRegException.Cells[1, LstIndex] := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'RegKey')
    else fmSettings.GridRegException.Cells[1, LstIndex] := ReadLangStr('RegistryCleaner.lng', 'Registry Cleaner', 'Parameter');
  fmSettings.GridRegException.Cells[2, LstIndex] := edExclusionString.Text;
  fmSettings.GridRegException.AutoSizeColumns(True);
  Close;
end;
//=========================================================



//=========================================================
{œ–» —Œ«ƒ¿Õ»» ‘Œ–Ã€}
//---------------------------------------------------------
procedure TfmRegExceptionsAdd.FormCreate(Sender: TObject);
begin
  ApplyLang;
end;
//=========================================================



end.

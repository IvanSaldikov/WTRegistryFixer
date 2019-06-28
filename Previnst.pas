unit Previnst;

interface

uses Windows;

var
  IsAlreadyRunning: boolean; //��� ���������� ���� true �� ��������� ��� ��������

implementation

var
  hMutex: integer;
begin
  IsAlreadyRunning := false;
  hMutex := CreateMutex(nil, TRUE, 'RegistryCleaner'); // ������� �������
  if GetLastError <> 0 then IsAlreadyRunning := true; // ������ ������� ��� ������
  ReleaseMutex(hMutex);
end.


unit Previnst;

interface

uses Windows;

var
  IsAlreadyRunning: boolean; //эта переменная если true то программа уже запущена

implementation

var
  hMutex: integer;
begin
  IsAlreadyRunning := false;
  hMutex := CreateMutex(nil, TRUE, 'RegistryCleaner'); // Создаем семафор
  if GetLastError <> 0 then IsAlreadyRunning := true; // Ошибка семафор уже создан
  ReleaseMutex(hMutex);
end.


unit ExportExcel;

interface
uses
  System.Classes,System.Variants,ActiveX,ComObj,CommCtrl,Dialogs,SysUtils;

//Private
 function CheckExcelInstall:boolean;
 function CheckExcelRun: boolean;
 function RunExcel(DisableAlerts:boolean=true; Visible: boolean=true): boolean;
 function AddWorkBook(AutoRun:boolean=true):boolean;
 function GetAllWorkBooks:TStringList;
 function SaveWorkBook(FileName:TFileName; WBIndex:integer;FileFormat:string):boolean;
 function StopExcel:boolean;
 function AllRunExcel:boolean;

const ExcelApp = 'Excel.Application';
implementation
uses umain;




/////////////////////////// Процедуры связанные с excel////////////////////////
//////////////////////////////////////////////////////////////////////////////
function AllRunExcel:boolean;   //// полный запуск Excel
begin
if ExportExcel.CheckExcelRun then /// если ексель запущен то зашибись используем его
  begin
  if not exportExcel.AddWorkBook then  /// создаем лист ексель
    begin
    showmessage('Что то пошло не так, не могу создать лист Excel!!!');
    exit;
    end;
  end
else   ////// иначе если excel не запущен
  begin
  if not ExportExcel.RunExcel(false,false) then  /// запускаем excel
    begin
    showmessage('У Вас не установлен Excel!!!');
    exit;
    end
    else
      begin
       if not exportExcel.AddWorkBook then  /// создаем лист ексель
        begin
        showmessage('Что то пошло не так, не могу создать лист Excel!!!');
        exit;
        end;
      end;
  end;
end;
///////////////////////////////////////////////////////////////////////////////
///Если функция CLSIDFromProgID находит CLSID OLE-объекта, то, соответственно —
//Excel установлен.
function CheckExcelInstall:boolean;
var
  ClassID: TCLSID;
  Rez : HRESULT;
begin
// Ищем CLSID OLE-объекта
  Rez := CLSIDFromProgID(PWideChar(WideString(ExcelApp)), ClassID);
  if Rez = S_OK then  // Объект найден
    Result := true
  else
    Result := false;
end;
//////////////////////////////////////////////////////////////////////////////////////////////

////Думаю тут лишних объяснений не потребуется? Всё предельно просто — если есть
///активный процесс Excel, то мы просто получаем на него ссылку и можем использовать Excel
///для своих корыстных целей. Главное — не забыть проверить — может кто-то этот самый Excel
///забыл закрыть,  но это другой момент. Будем считать, что Excel в нашем полном распоряжении.
function CheckExcelRun: boolean;
begin
  try
    MyExcel:=GetActiveOleObject(ExcelApp);
    Result:=True;
  except
    Result:=false;
  end;
end;
//////////////////////////////////////////////////////////////////////////////////////

/////функция запуска Excel
function RunExcel(DisableAlerts:boolean; Visible: boolean): boolean;
begin
  try
{проверяем установлен ли Excel}
    if CheckExcelInstall then
      begin
        MyExcel:=CreateOleObject(ExcelApp);
//показывать/не показывать системные сообщения Excel (лучше не показывать)
        MyExcel.Application.EnableEvents:=DisableAlerts;
        MyExcel.Visible:=Visible;
        Result:=true;
      end
    else
      begin
        //MessageBox(0,'Приложение MS Excel не установлено на этом компьютере','Ошибка',MB_OK+MB_ICONERROR);
        Result:=false;
      end;
  except
    Result:=false;
  end;
end;
/////Здесь мы в начале проверяем, установлен ли Excel в принципе и, если он все
/// же установлен, запускам. При этом мы можем сразу показать окно Excel пользователю
/// — для этого необходимо выставить параметр Visible в значение True.
///Также рекомендую всегда отключать системные сообщения Excel, иначе,
/// когда программа начнет говорить голосом Excel — пользователь может занервничать.
///////////////////////////////////////////////////////////////////////////////////////

///// создание пустой рабочей книги
function AddWorkBook(AutoRun:boolean=true):boolean;
begin
  if CheckExcelRun then
    begin
      MyExcel.WorkBooks.Add;
      Result:=true;
    end
  else
   if AutoRun then
     begin
       RunExcel;
       MyExcel.WorkBooks.Add;
       Result:=true;
     end
   else
     Result:=false;
end;
////////То есть сразу проверяю запущен ли Excel и, если он не запущен, то либо
///запускаю и добавляю книгу, либо просто выхожу — всё зависит от ситуации.
/////////////////////////////////////////////////////////////////////////////////////////

/////////Здесь, думаю, следует напомнить, что, если вы выполните эту функцию,
//// например пять раз, то получите пять открытых рабочих книг и работать с ними
////как Вам захочется. Главное при этом правильно обратиться к необходимой книге,
/// а для этого можно использовать вот такую функцию:
function GetAllWorkBooks:TStringList;
var i:integer;
begin
  try
    Result:=TStringList.Create;
    for i:=1 to MyExcel.WorkBooks.Count do
      Result.Add(MyExcel.WorkBooks.Item[i].FullName)
  except
    ///MessageBox(0,'Ошибка перечисления открытых книг','Ошибка',MB_OK+MB_ICONERROR);
  end;
end;
///Функция возвращает список TStringList всех рабочих книг Excel открытых в данный момент.
//Обратите внимание, что в отличие от Delphi Excel присваивает первой книге индекс
/// 1, а не 0 как это обычно делается в Delphi при работе, например, с теми же индексами
// в ComboBox’ах.
////////////////////////////////////////////////////////////////////////////////////////

///Для того, чтобы сохранить рабочую книгу, я использовал такую функцию:
function SaveWorkBook(FileName:TFileName; WBIndex:integer;FileFormat:string):boolean;
begin
  try
    MyExcel.WorkBooks.Item[WBIndex].SaveAs(FileName,FileFormat);
    if MyExcel.WorkBooks.Item[WBIndex].Saved then
      Result:=true
    else
      Result:=false;
  except
    Result:=false;
  end;
end;
/////Если у Вас открыто 10 книг — просто вызываете функцию 10 раз, меняя значение
///параметра WBIndex и имени файла и дело в шляпе.
//////////////////////////////////////////////////////////////////////////////////////

/////////// Закрываем Excel
function StopExcel:boolean;
begin
  try
    if MyExcel.Visible then MyExcel.Visible:=false;
    MyExcel.Quit;
    MyExcel:=Unassigned;
    Result:=True;
  except
    Result:=false;
  end;
end;
////Вот набор тех основных функций с которых начинается вся интеграция Excel в
// приложения написанные на Delphi. В следующий раз займемся работой с конкретной
//книгой — научимся записывать и читать данные из книг.
///////////////////////////////////////////////////////////////////////////////////////
  ////// конец процедур связанными с excel



end.

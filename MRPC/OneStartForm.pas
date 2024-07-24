unit OneStartForm;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.ExtCtrls,inifiles;

type
  TFormOneStart = class(TForm)
    RadioGroup1: TRadioGroup;
    RadioGroup2: TRadioGroup;
    LabeledEdit1: TLabeledEdit;
    LabeledEdit2: TLabeledEdit;
    GroupBox1: TGroupBox;
    Label2: TLabel;
    GroupBox2: TGroupBox;
    Button1: TButton;
    GroupBox3: TGroupBox;
    ComboGroup: TComboBox;
    Label3: TLabel;
    CheckBox1: TCheckBox;
    CheckBox2: TCheckBox;
    procedure RadioGroup1Click(Sender: TObject);
    procedure RadioGroup2Click(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure ComboGroupKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure CheckBox1MouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure CheckBox2MouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
  private
   function PCinAD:boolean;
   function PCinFile:boolean;
  public
    { Public declarations }
  end;

var
  FormOneStart: TFormOneStart;
  typelan:integer;
implementation
uses umain,IPaddress,MyDM;


{$R *.dfm}

procedure Code(var text: WideString; password: string;  //// процедура кодирования и декодирования файла
decode: boolean);
var
i, PasswordLength: integer;
sign: shortint;
begin
PasswordLength := length(password);
if PasswordLength = 0 then Exit;
if decode then sign := -1
else sign := 1;
for i := 1 to Length(text) do
text[i] := chr(ord(text[i]) + sign *
ord(password[i mod PasswordLength + 1]));
end;



function TFormOneStart.PCinAD:boolean;
var
setini:Tmeminifile;
DC:string;
begin
if ComboGroup.Text='' then
begin
ShowMessage('Для продолжения выберите группу безопасности домена');
if frmDomainInfo.combobox1.Enabled then ComboGroup.DroppedDown:=true;
result:=false;
exit;
end;
try
 SetInI:=TMeMiniFile.Create(extractfilepath(application.ExeName)+'\Settings.ini',TEncoding.Unicode);
 try
  frmDomainInfo.ListView8.Clear;
  frmDomainInfo.Combobox2.Clear;
  frmDomainInfo.combobox1.Enabled:=true; // разблокировать список групп домена
  frmDomainInfo.combobox8.Enabled:=true; /// разблокировать список пользователей
  if frmDomainInfo.ComboBox1.tag=1 then frmDomainInfo.combobox1.Enabled:=false; // если не установлена лицензия то блокируем доступ к списку групп безопасности
  frmDomainInfo.combobox1.Text:=ComboGroup.text; // имя группы в комббобокс на главной форме

  if SetIni.ReadBool('ConfLAN','UserGroup',true) then // Выгружать только пользователей выбранной группы безопасности домена
  frmDomainInfo.GetAllGroupUsersInComboBox(ComboGroup.text)// загрузка списка пользователей выбранной группы в ComboBox главного окна программы
  else frmDomainInfo.EnumAllUsers; // иначе выгружаем всех пользователей домена

  if SetIni.ReadBool('ConfLAN','PCGroup',true) then // Выгружать только компьютеры выбранной группы безопасности домена
  frmDomainInfo.GetPCForGroupToLisViewComboBox(ComboGroup.text) // загрузка списка компьютеров выбранной группы в ListWiew и Combobox в главном окне программы
  else // иначе выгружаем в Combobox весь список компьютеров домена а список группы и в ListView
  begin
  frmDomainInfo.GetAllGroupUsers(ComboGroup.text); //список выбранной группы и в ListView
  frmDomainInfo.EnumAllWorkStation; // передаем все рабочие станции в combobox в разделе Компьютеры
  end;

  if CheckBox2.Checked then // если установлен чекбокс то загружаем данные из базы
  if dataM.ConnectionDB.Connected then frmDomainInfo.readinfoforpcDB;
  /// записываем в файл настроек то что открывали последний раз компы из AD


  setini.WriteString('ConfLAN','type','AD'); // в файле настроек указываем что загружать списки пк и AD
  setini.WriteString('ConfLAN','Group',ComboGroup.text); // и указываем из какой группы
  finally
  setini.UpdateFile;
  setini.Free;
  end;
  result:=true;
Except on E: Exception do
begin
 showmessage('Не предвиденная ошибка '+e.Message);
 result:=false;
end;
end;
end;

function TFormOneStart.PCinFile:boolean;
var                 /// список компьютеров\IP адресов из файла
OpenF:TopenDialog;
ListFile:TstringList;
s:string;
i,Pos1, Len:integer;
setini:Tmeminifile;
const
D = [',', ':', ';', ' ', #9, #10, #13];
begin
try
  if not Assigned(OpenF) then
  begin
  OpenF:=TopenDialog.Create(FormOneStart);
  OpenF.Name:='OpenFileTxT';
  end;
  OpenF.Filter:='текстовые файлы |*.txt|';
  OpenF.Title:='Выберите файл со списком ПК';
  if not OpenF.Execute then
    begin
    FreeAndNil(OpenF);
    result:=false;
    exit;
    end
  else
   begin
    if not Assigned(listfile) then
    begin
    ListFile:=TstringList.Create;
    end;
    ListFile.LoadFromFile(OpenF.FileName);
    s:=ListFile.Text;
    ListFile.Clear;
    Len:= Length(S);
     for I := 0 to Len do
       begin
        //Пропускаем разделители.
      if S[i] in D then Continue;
      //Отслеживаем начало слова.
      if (i = 1) or (S[i - 1] in D) then Pos1 := i;
      //Отслеживаем конец слова.
      if (i = Len) or (S[i + 1] in D) then
          begin
            //Добавляем слово в ListFile
           ListFile.Add( Copy(S, Pos1, i - Pos1 + 1) );
          end;
       end;
        frmDomainInfo.ListView8.Clear;
        frmDomainInfo.combobox2.clear;
        frmDomainInfo.combobox1.Enabled:=false; ///список групп
        frmDomainInfo.combobox1.Text:='';
        frmDomainInfo.combobox8.Enabled:=false; /// список пользователей
        frmDomainInfo.combobox8.Text:='';
        for I := 0 to ListFile.Count-1 do
          begin
          if (pos(Ansiuppercase('НЕ ЗАРЕГИСТРИРОВАННАЯ'),Ansiuppercase(frmDomainInfo.Caption))<>0)
          and (i>=30) then break;
            frmDomainInfo.combobox2.Items.Add(ListFile.Strings[i]);
            with frmDomainInfo.listview8.Items.Add do
              begin
                caption:='';
                subitems.add(ListFile.Strings[i]);
                subitems.add('');
                subitems.add('');
                SubItems.Add('');
                SubItems.Add('');
                SubItems.Add('');
                SubItems.Add('');
                subitems.add('');
                SubItems.Add('');
                subitems.add('');
                subitems.add('');
                subitems.add('');
                subitems.add('');
              end;
          end;
  listfile.SaveToFile(extractfilepath(application.ExeName)+'\listPC.txt');
  listfile.Free;
  if Fileexists(extractfilepath(application.ExeName)+'\Settings.ini') then
  begin
  SetInI:=TMeMiniFile.Create(extractfilepath(application.ExeName)+'\Settings.ini',TEncoding.Unicode);
  setini.WriteString('ConfLAN','type','file');
  setini.UpdateFile;
  setini.Free;
  end;
if CheckBox2.Checked then // если установлен чекбокс то загружаем данные из базы
  if dataM.ConnectionDB.Connected then frmDomainInfo.readinfoforpcDB;
result:=true;
/////////////////////////////////////////////////////////////// создание списка компьютеров
end;
  if Assigned(OpenF) then  FreeAndNil(OpenF);
except
 on E: Exception do
 begin
 showmessage('Ошибка: '+e.Message);
 result:=false;
 end;
 end;
 end;

procedure TFormOneStart.Button1Click(Sender: TObject);
var
setini:Tmeminifile;
s:widestring;
oper:boolean;
z:integer;
begin
try
if (pos(Ansiuppercase('НЕ ЗАРЕГИСТРИРОВАННАЯ'),Ansiuppercase(frmDomainInfo.Caption))<>0) then
begin
z:=Button1.Tag;
inc(z);
Button1.Tag:=z;
if Button1.Tag>=3 then
  begin
  ShowMessage('Больше компьютеров в зарегистрированной версии программы');
  close;
  exit;
  end;
end;

oper:=false;
frmDomainInfo.LabeledEdit1.Text:=LabelEdEdit2.Text; ///// пользователь
MyUser:=LabelEdEdit2.Text;
frmDomainInfo.LabeledEdit2.Text:=LabelEdEdit1.Text; ///// пароль
MyPasswd:=LabelEdEdit1.Text;
SetInI:=TMeMiniFile.Create(extractfilepath(application.ExeName)+'\Settings.ini',TEncoding.Unicode);
s:= widestring(LabelEdEdit1.Text);
code(s,'1234',false);
SetInI.WriteString('ADM','passwd',string(s));
SetInI.WriteString('ADM','user',LabelEdEdit2.Text);
setini.UpdateFile;
setini.Free;
case typelan of
1:    ///'AD'
  begin
  oper:=PCinAD;
  end;
2:  // 'IP
  begin
  OKRightDlg1234567891011121314151617181920.ShowModal;
  oper:=true;
  end;
3:  //'file'
  begin
  oper:=PCinFile;
  end;
end;
except
 on E: Exception do
 showmessage('Ошибка: '+e.Message);
 end;

if oper then close;
end;




procedure TFormOneStart.CheckBox1MouseUp(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
  var
 setini:Tmeminifile;
begin
SetInI:=TMeMiniFile.Create(extractfilepath(application.ExeName)+'\Settings.ini',TEncoding.Unicode);
try
SetInI.WriteBool('ConfLAN','Load',CheckBox1.Checked);
finally
setini.UpdateFile;
setini.Free;
end;
end;

procedure TFormOneStart.CheckBox2MouseUp(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
  var
 setini:Tmeminifile;
begin
SetInI:=TMeMiniFile.Create(extractfilepath(application.ExeName)+'\Settings.ini',TEncoding.Unicode);
try
SetInI.WriteBool('ConfLAN','LoadInfoDB',CheckBox2.Checked);
finally
setini.UpdateFile;
setini.Free;
end;
end;

procedure TFormOneStart.ComboGroupKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
ComboGroup.DroppedDown:=true;
end;

procedure TFormOneStart.FormCreate(Sender: TObject);

begin
// timer1.Enabled:=true;
end;

procedure TFormOneStart.FormShow(Sender: TObject);
var
 setini:Tmeminifile;
begin
SetInI:=TMeMiniFile.Create(extractfilepath(application.ExeName)+'\Settings.ini',TEncoding.Unicode);
try
CheckBox1.Checked:=SetInI.ReadBool('ConfLAN','Load',false);
if SetInI.ReadString('ConfLAN','type','FirstStart')='FirstStart' then
CheckBox2.Checked:=false
else CheckBox2.Checked:=SetInI.ReadBool('ConfLAN','LoadInfoDB',true);
finally
setini.Free;
end;
ComboGroup.Clear; // чистим combobox для списка групп
RadioGroup1.ItemIndex:=-1;
RadioGroup2.ItemIndex:=-1;
RadioGroup2.Enabled:=false;
ComboGroup.Enabled:=false;
typelan:=0;
end;

procedure TFormOneStart.RadioGroup1Click(Sender: TObject);
var
i:integer;
groupList:TstringList;
setini:Tmeminifile;
DC:string;
begin

case RadioGroup1.ItemIndex of
0:
  begin
  RadioGroup2.enabled:=false;
  typelan:=1;
  ComboGroup.Clear; // чистим т.к. потом его заполняем списком групп
  frmDomainInfo.ComboBox1.Clear; // чистим т.к. потом его заполняем списком групп
 // if (pos(Ansiuppercase('НЕ ЗАРЕГИСТРИРОВАННАЯ'),Ansiuppercase(frmDomainInfo.Caption))<>0) then ComboGroup.Enabled:=false
 // else
  ComboGroup.Enabled:=true;

  frmDomainInfo.GetCurrentComputerName;  ///  информация о компьютере  в результате должно быть имя, но главное получваем в переменную CurrentDomainName имя домена
  frmDomainInfo.LabeledEdit3.Text := CurrentDomainName; //и указываем имя домена в главном окне программы
  DC:=frmDomainInfo.GetDomainController(CurrentDomainName);
  if DC = '' then
  begin
  ShowMessage('Проверьте доступность контроллера домена и повторите операцию.');
  exit;
  end;
  frmDomainInfo.Memo1.Lines.Add('Текущий домен - '+CurrentDomainName);
  frmDomainInfo.Memo1.Lines.Add('Текущий контроллер домена - '+DC);

  groupList:=TStringList.Create;
  SetInI:=TMeMiniFile.Create(extractfilepath(application.ExeName)+'\Settings.ini',TEncoding.Unicode);
  try
  if frmDomainInfo.EnumAllGroupsToList(groupList) then // получаем список групп безопасности домена
  for I := 0 to groupList.Count-1 do
    begin
    ComboGroup.Items.Add(groupList[i]); // заполняем списком групп combobox в текущем окне
    frmDomainInfo.ComboBox1.Items.Add(groupList[i]); // заполняем списком групп в главном окне программы
    end;

  if SetInI.readString('ConfLAN','type','')='AD' then // если прошлый раз загружали из домена то передаем сохраненную группу
    begin
    ComboGroup.text:=SetInI.readstring('ConfLAN','Group','');
    //if not ComboGroup.Enabled then showmessage('Повторный выбор групп только в зарегистрированной версии программы!!!');
    end
  else
    begin
    ComboGroup.Enabled:=true;
    ComboGroup.DroppedDown:=true; // открыть список загруженных групп безопасности
    end;


  finally
  groupList.Free;
  setini.Free;
  end;
  end;
1:
  begin
  RadioGroup2.enabled:=true;
  RadioGroup2.ItemIndex:=0;
  typelan:=2;
  ComboGroup.Enabled:=false;
  end;
end;


end;

procedure TFormOneStart.RadioGroup2Click(Sender: TObject);
begin
case RadioGroup2.ItemIndex of
0:
  begin
  typelan:=2;
  end;
1:
  begin
  typelan:=3;
  end;

end;
end;





end.

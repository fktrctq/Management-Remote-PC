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




/////////////////////////// ��������� ��������� � excel////////////////////////
//////////////////////////////////////////////////////////////////////////////
function AllRunExcel:boolean;   //// ������ ������ Excel
begin
if ExportExcel.CheckExcelRun then /// ���� ������ ������� �� �������� ���������� ���
  begin
  if not exportExcel.AddWorkBook then  /// ������� ���� ������
    begin
    showmessage('��� �� ����� �� ���, �� ���� ������� ���� Excel!!!');
    exit;
    end;
  end
else   ////// ����� ���� excel �� �������
  begin
  if not ExportExcel.RunExcel(false,false) then  /// ��������� excel
    begin
    showmessage('� ��� �� ���������� Excel!!!');
    exit;
    end
    else
      begin
       if not exportExcel.AddWorkBook then  /// ������� ���� ������
        begin
        showmessage('��� �� ����� �� ���, �� ���� ������� ���� Excel!!!');
        exit;
        end;
      end;
  end;
end;
///////////////////////////////////////////////////////////////////////////////
///���� ������� CLSIDFromProgID ������� CLSID OLE-�������, ��, �������������� �
//Excel ����������.
function CheckExcelInstall:boolean;
var
  ClassID: TCLSID;
  Rez : HRESULT;
begin
// ���� CLSID OLE-�������
  Rez := CLSIDFromProgID(PWideChar(WideString(ExcelApp)), ClassID);
  if Rez = S_OK then  // ������ ������
    Result := true
  else
    Result := false;
end;
//////////////////////////////////////////////////////////////////////////////////////////////

////����� ��� ������ ���������� �� �����������? �� ��������� ������ � ���� ����
///�������� ������� Excel, �� �� ������ �������� �� ���� ������ � ����� ������������ Excel
///��� ����� ��������� �����. ������� � �� ������ ��������� � ����� ���-�� ���� ����� Excel
///����� �������,  �� ��� ������ ������. ����� �������, ��� Excel � ����� ������ ������������.
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

/////������� ������� Excel
function RunExcel(DisableAlerts:boolean; Visible: boolean): boolean;
begin
  try
{��������� ���������� �� Excel}
    if CheckExcelInstall then
      begin
        MyExcel:=CreateOleObject(ExcelApp);
//����������/�� ���������� ��������� ��������� Excel (����� �� ����������)
        MyExcel.Application.EnableEvents:=DisableAlerts;
        MyExcel.Visible:=Visible;
        Result:=true;
      end
    else
      begin
        //MessageBox(0,'���������� MS Excel �� ����������� �� ���� ����������','������',MB_OK+MB_ICONERROR);
        Result:=false;
      end;
  except
    Result:=false;
  end;
end;
/////����� �� � ������ ���������, ���������� �� Excel � �������� �, ���� �� ���
/// �� ����������, ��������. ��� ���� �� ����� ����� �������� ���� Excel ������������
/// � ��� ����� ���������� ��������� �������� Visible � �������� True.
///����� ���������� ������ ��������� ��������� ��������� Excel, �����,
/// ����� ��������� ������ �������� ������� Excel � ������������ ����� ������������.
///////////////////////////////////////////////////////////////////////////////////////

///// �������� ������ ������� �����
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
////////�� ���� ����� �������� ������� �� Excel �, ���� �� �� �������, �� ����
///�������� � �������� �����, ���� ������ ������ � �� ������� �� ��������.
/////////////////////////////////////////////////////////////////////////////////////////

/////////�����, �����, ������� ���������, ���, ���� �� ��������� ��� �������,
//// �������� ���� ���, �� �������� ���� �������� ������� ���� � �������� � ����
////��� ��� ���������. ������� ��� ���� ��������� ���������� � ����������� �����,
/// � ��� ����� ����� ������������ ��� ����� �������:
function GetAllWorkBooks:TStringList;
var i:integer;
begin
  try
    Result:=TStringList.Create;
    for i:=1 to MyExcel.WorkBooks.Count do
      Result.Add(MyExcel.WorkBooks.Item[i].FullName)
  except
    ///MessageBox(0,'������ ������������ �������� ����','������',MB_OK+MB_ICONERROR);
  end;
end;
///������� ���������� ������ TStringList ���� ������� ���� Excel �������� � ������ ������.
//�������� ��������, ��� � ������� �� Delphi Excel ����������� ������ ����� ������
/// 1, � �� 0 ��� ��� ������ �������� � Delphi ��� ������, ��������, � ���� �� ���������
// � ComboBox���.
////////////////////////////////////////////////////////////////////////////////////////

///��� ����, ����� ��������� ������� �����, � ����������� ����� �������:
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
/////���� � ��� ������� 10 ���� � ������ ��������� ������� 10 ���, ����� ��������
///��������� WBIndex � ����� ����� � ���� � �����.
//////////////////////////////////////////////////////////////////////////////////////

/////////// ��������� Excel
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
////��� ����� ��� �������� ������� � ������� ���������� ��� ���������� Excel �
// ���������� ���������� �� Delphi. � ��������� ��� �������� ������� � ����������
//������ � �������� ���������� � ������ ������ �� ����.
///////////////////////////////////////////////////////////////////////////////////////
  ////// ����� �������� ���������� � excel



end.

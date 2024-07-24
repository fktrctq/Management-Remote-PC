unit Unit4;

interface

uses
  System.Classes,System.Variants,ActiveX,ComObj,CommCtrl,Dialogs,SysUtils;

type
  Priority = class(TThread)
  private
    { Private declarations }
  protected
    procedure Execute; override;
  end;

implementation
uses umain;


{ Priority }   /// Задает приоритет процесса

procedure Priority.Execute;
var
  FSWbemLocator : OLEVariant;
  FWMIService   : OLEVariant;
  FWbemObjectSet: OLEVariant;
  FWbemObject   : OLEVariant;
  oEnum         : IEnumvariant;
  iValue        : LongWord;
  ResPriority   :integer;
begin

try
 begin
     OleInitialize(nil);
     FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
     FWMIService  := FSWbemLocator.ConnectServer(MyPS, 'root\CIMV2', MyUser, MyPasswd,'','',128);
      FWbemObjectSet:= FWMIService.ExecQuery('SELECT * FROM Win32_Process WHERE ProcessId='+GroupselectProc,'WQL',wbemFlagForwardOnly);
      oEnum:= IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
           while oEnum.Next(1, FWbemObject, iValue) = 0 do     ////
            begin
            ResPriority:=FWbemObject.SetPriority(MyPriority);
            FWbemObject:=Unassigned;
            frmDomainInfo.memo1.Lines.Add('Приоритет процесса. '+SysErrorMessage(ResPriority));
            frmDomainInfo.memo1.Lines.Add('---------------------------');
            FWbemObject:=Unassigned;
            end;
      VariantClear(FWbemObject);
      oEnum:=nil;
      VariantClear(FWbemObjectSet);
      VariantClear(FWMIService);
      VariantClear(FSWbemLocator);
      OleUnInitialize;
 end;
except
on E:Exception do
         begin
           frmDomainInfo.memo1.Lines.Add('Ошибка  "'+E.Message+'"');
           frmDomainInfo.memo1.Lines.Add('---------------------------');
           OleUnInitialize;
           exit;
         end;

end;

end;

end.

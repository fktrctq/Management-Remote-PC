unit Unit2;

interface

uses
  System.Classes,System.Variants,ActiveX,ComObj,CommCtrl,Dialogs,SysUtils;

type
  Procent = class(TThread)
  private
    { Private declarations }
  protected
    procedure Execute; override;
    procedure UpdateCPU;
  end;

implementation
uses umain;


{
 Вычисление процентов загрузки процессора процессом а также объем памяти процесса

{ Procent }


procedure Procent.UpdateCPU;
var
i,z,mem:integer;
UTime,PTime,MyCpu:integer;
loadmem,LoadPrivate:real;
  FSWbemLocatorCPU : OLEVariant;
  FWMIServiceCPU   : OLEVariant;
  FWbemObjectSetCPU: OLEVariant;
  FWbemObjectCPU   : OLEVariant;
  oEnumCPU         : IEnumvariant;
  iValue        : LongWord;
 begin
     SumTime:=0;
     SumMemory:=0;
     SetPeak:=0;
     i:=0;
     OleInitialize(nil); /////
     MyCPU:=CpuLogPr*CpuNum;
     FSWbemLocatorCPU := CreateOleObject('WbemScripting.SWbemLocator');
     FWMIServiceCPU   := FSWbemLocatorCPU.ConnectServer(MyPS, 'root\CIMV2', MyUser, MyPasswd,'','',128);
      begin
          FWbemObjectSetCPU:= FWMIServiceCPU.ExecQuery('SELECT PercentPrivilegedTime,PercentUserTime,WorkingSet,WorkingSetPeak FROM Win32_PerfFormattedData_PerfProc_Process WHERE IDProcess<>0' ,'WQL',wbemFlagForwardOnly);
          oEnumCPU:= IUnknown(FWbemObjectSetCPU._NewEnum) as IEnumVariant;
           while oEnumCPU.Next(1, FWbemObjectCPU, iValue) = 0 do     ////
            begin
              if (FWbemObjectCPU.PercentPrivilegedTime<>null) and(FWbemObjectCPU.PercentUserTime<>null) then
                begin
                UTime:= (FWbemObjectCPU.PercentPrivilegedTime) div (MyCPU);
                PTime:= (FWbemObjectCPU.PercentUserTime) div (MyCPU);
                SumTime:=SumTime+(UTime+Ptime);
                frmDomainInfo.lvWorkStation.Items[i].SubItems[3]:=vartostr((UTime+Ptime));
                end;
              if FWbemObjectCPU.WorkingSet<>null then
                frmDomainInfo.lvWorkStation.Items[i].SubItems[4]:=(vartostr(round((FWbemObjectCPU.WorkingSet) /1024)))+' Kb';
              if FWbemObjectCPU.WorkingSet<>null then
                SumMemory:= SumMemory+round(((FWbemObjectCPU.WorkingSet)/ 1024));
              if FWbemObjectCPU.WorkingSetPeak<>null then
                SetPeak:=SetPeak+round(((FWbemObjectCPU.WorkingSetPeak) / 1024));
              //if not VarIsNull(FWbemObjectCPU) then FWbemObjectCPU := Unassigned;
              FWbemObjectCPU := Unassigned;
            inc(i);
            end;
            //mem:=strtoint(floattostr(MyMemory));
            //loadmem:=(SumMemory div 1024)/ mem;
            //LoadPrivate:=(SetPeak div 1024)/mem;
            frmDomainInfo.StatusBar1.Panels[3].Text:='CPU - '+inttostr(SumTime)+' %';
            //frmDomainInfo.StatusBar1.Panels[4].Text:='Физическая память: Рабочий набор - '+floattostr(round((loadmem)*100))+'%'+' / Пиковый набор - '+floattostr(round((LoadPrivate)*100))+'%';
            //frmDomainInfo.StatusBar1.Panels[5].Text:='Процессов - '+ inttostr(frmDomainInfo.lvWorkStation.Items.Count);
            oEnumCPU:=nil; // Используйте вместо метода _Release
            VariantClear(FWbemObjectSetCPU);
            VariantClear(FWMIServiceCPU);
            VariantClear(FSWbemLocatorCPU);
            VariantClear(FWbemObjectCPU);
        OleUnInitialize;
     end;
end;


procedure Procent.Execute;
begin
UpdateCPU;
end;


end.

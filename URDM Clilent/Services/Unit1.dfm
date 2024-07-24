object ServiceRDM: TServiceRDM
  OldCreateOrder = False
  DisplayName = 'ServiceRDM'
  OnPause = ServicePause
  OnShutdown = ServiceShutdown
  OnStart = ServiceStart
  OnStop = ServiceStop
  Height = 150
  Width = 215
  object Timer1: TTimer
    Enabled = False
    Interval = 5000
    OnTimer = Timer1Timer
    Left = 168
    Top = 72
  end
end

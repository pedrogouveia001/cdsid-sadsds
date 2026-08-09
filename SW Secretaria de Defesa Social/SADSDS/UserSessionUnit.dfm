object IWUserSession: TIWUserSession
  OldCreateOrder = False
  Height = 429
  Width = 657
  object ZQuery1: TZQuery
    Connection = ZConnection1
    Params = <>
    Left = 40
    Top = 48
  end
  object ZConnection1: TZConnection
    ControlsCodePage = cCP_UTF16
    HostName = 'localhost'
    Port = 3306
    Database = 'fitradeoff_data'
    User = 'root'
    Password = ''
    Protocol = 'mysql'
    Left = 144
    Top = 56
  end
end

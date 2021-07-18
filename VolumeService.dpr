program VolumeService;

uses
  SvcMgr,
  main in 'main.pas' {Volume_Service: TService},
  MMDevAPI in 'MMDevAPI.pas';

{$R *.RES}

begin
  Application.Initialize;
  Application.CreateForm(TVolume_Service, Volume_Service);
  Application.Run;
end.

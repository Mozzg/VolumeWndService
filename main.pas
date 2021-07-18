unit main;

interface

uses
  Windows, SysUtils, Classes, SvcMgr, INIFiles, ActiveX, MMDevAPI, DateUtils, ComObj;

type
  TVolume_Service = class(TService)
    procedure ServiceStart(Sender: TService; var Started: Boolean);
    procedure ServiceExecute(Sender: TService);
    procedure ServiceShutdown(Sender: TService);
    procedure ServiceStop(Sender: TService; var Stopped: Boolean);
  private
    { Private declarations }
  public
    procedure ServiceStopShutdown;
    function GetServiceController: TServiceController; override;
    { Public declarations }
  end;

  Time_volume_params = record
    time:integer;
    volume:byte;
  end;

  TWorkThread = class(TThread)
  public
    procedure execute; override;
  end;

var
  Volume_Service: TVolume_Service;
  WorkThread:TWorkThread;

  LogFileName:string;
  LogLevel:integer;
  Volume_arr:array of Time_volume_params;
  IniFile:TINIFile;
  INIFilePath:string;

  endpointVolume: IAudioEndpointVolume = nil;

implementation

{$R *.DFM}

function GetModuleFileNameStr(Instance: THandle): string;
var
  buffer: array [0..MAX_PATH] of Char;
begin
  GetModuleFileName( Instance, buffer, MAX_PATH);
  Result := buffer;
end;

function Log(Mess:string; time:boolean = true):boolean;
var handl:integer;
temp_mess:string;
begin
  result:=false;
  if LogLevel=0 then exit;
  temp_mess:=Mess+#13+#10;
  if time=true then temp_mess:=FormatDateTime('dd.mm.yyyy hh:nn:ss.zzz',now)+'  '+temp_mess;
  
  if LogFileName<>'' then
  begin
    if FileExists(LogFileName) then
      handl:=FileOpen(LogFileName,fmOpenReadWrite or fmShareDenyNone)
    else
      handl:=FileCreate(LogFileName);

    if handl<0 then exit;
    if FileSeek(handl,0,2)=-1 then exit;
    if FileWrite(handl,temp_mess[1],length(temp_mess))=-1 then exit;
    FileClose(handl);
  end
  else
  begin
    temp_mess:='';
    exit;
  end;

  temp_mess:='';
  result:=true;
end;

procedure ServiceController(CtrlCode: DWord); stdcall;
begin
  Volume_Service.Controller(CtrlCode);
end;

function TVolume_Service.GetServiceController: TServiceController;
begin
  Result := ServiceController;
end;

procedure TVolume_Service.ServiceStopShutdown;
begin
  Log('Stop/Shutdown procedure began',true);

  if Assigned(WorkThread) then
  begin
    Log('Begin to stop thread',true);
    if WorkThread.Suspended then WorkThread.Resume;
    WorkThread.Terminate;
    WorkThread.WaitFor;
    FreeAndNil(WorkThread);
    Log('Thread stopped sucsessfuly',true);
  end;
  INIFile.Free;
  Log('INI file freed',true);
end;

procedure TVolume_Service.ServiceStart(Sender: TService;
  var Started: Boolean);
var hndl:integer;
str:string;
list:TStringList;
i:integer;
deviceEnumerator: IMMDeviceEnumerator;
defaultDevice: IMMDevice;
res:Hresult;
begin
  //меняем текущую папку
  SetCurrentDirectory(PChar(ExtractFilePath(GetModuleFileNameStr(0))));
  LogLevel:=1;

  //инициализация переменных
  LogFileName:='output.log';
  INIFilePath:=ChangeFileExt(GetModuleFileNameStr(0), '.INI');

  //Проверяем, есть ли INI файл
  if not(FileExists(INIFilePath)) then
  begin   //если нет, то создаем
    Log('Missing INI file, creating',true);
    hndl:=FileCreate(INIFilePath);

    if hndl<0 then
    begin
      Log('Failed to create INI file, shutting down',true);
      started:=false;
      exit;
    end
    else Log('INI File created',true)
  end
  else
    Log('INI file found',true);

  //инициализируем и читаем INI файл
  INIFile:=TINIFile.Create(INIFilePath);
  str:=INIFile.ReadString('Main','LogFileName','output.log');
  if expandfilename(str)<>expandfilename(LogFileName) then
  begin
    Log('Log file name is different in INI file ('+str+'), begining to write log to there',false);
    LogFileName:=str;
  end
  else
    Log('Log file is same, continuing',true);
  Log('==================================================================================================',false);
  Log('Starting service...',true);
  LogLevel:=INIFile.ReadInteger('Main','Log',0);

  //читаем массив для регулировки
  try
    list:=TStringList.Create;
    list.Clear;
    INIFile.ReadSection('Volume',list);
    if list.Count=0 then
    begin
      Log('Volume list is empty');
      started:=false;
      exit;
    end;

    setlength(Volume_arr,list.Count);
    Log('Volume list:');
    for i:=0 to list.Count-1 do
    begin
      Volume_arr[i].time:=strtoint(list.Strings[i]);
      Volume_arr[i].volume:=INIFile.ReadInteger('Volume',list.Strings[i],0);

      Log(inttostr(Volume_arr[i].time)+'='+inttostr(Volume_arr[i].volume),false);
    end;

    list.Clear;
    list.Free;
  except
    on e:exception do
    begin
      Log('Exception on reading volume list with message:'+e.Message);
      Started:=false;
      exit;
    end;
  end;

  //инициализируем управление громкостью
  {if not Succeeded(CoCreateInstance(CLASS_IMMDeviceEnumerator, nil, CLSCTX_INPROC_SERVER, IID_IMMDeviceEnumerator, deviceEnumerator)) then
  begin
    Log('Function CoCreateInstance failed');
    started:=false;
    exit;
  end;

  if not Succeeded(deviceEnumerator.GetDefaultAudioEndpoint(eRender, eConsole, defaultDevice)) then
  begin
    Log('Function deviceEnumerator.GetDefaultAudioEndpoint failed');
    started:=false;
    exit;
  end;

  if not Succeeded(defaultDevice.Activate(IID_IAudioEndpointVolume, CLSCTX_INPROC_SERVER, nil, endpointVolume)) then
  begin
    Log('Function defaultDevice.Activate failed');
    started:=false;
    exit;
  end;    }
  Log('Creating volume control');

  res:=CoInitialize(nil);
  Log('CoInitialize return='+inttohex(res,8));

  res:=CoCreateInstance(CLASS_IMMDeviceEnumerator, nil, CLSCTX_INPROC_SERVER, IID_IMMDeviceEnumerator, deviceEnumerator);
  Log('CoCreateInstance return='+inttohex(res,8));

  res:=deviceEnumerator.GetDefaultAudioEndpoint(eRender, eConsole, defaultDevice);
  Log('deviceEnumerator.GetDefaultAudioEndpoint return='+inttohex(res,8));

  res:=defaultDevice.Activate(IID_IAudioEndpointVolume, CLSCTX_INPROC_SERVER, nil, endpointVolume);
  Log('defaultDevice.Activate return='+inttohex(res,8));

  //запускаем поток
  Log('Creating thread');
  WorkThread:=TWorkThread.Create(true);
  WorkThread.FreeOnTerminate:=false;
  WorkThread.Resume;

  Log('Initialization sucsessful');

  Started:=true;
end;

procedure TWorkThread.execute;
label ex1;
var vol,cur_vol,vol_increment:single;
vol_int:integer;
i:integer;
work_time:cardinal;
time:int64;
begin
  Log('WorkThread enter');

  //проверяем сколько элементов массива
  if length(Volume_arr)<2 then
  begin
    if length(Volume_arr)=0 then
    begin
      Log('Volume array is empty');
      goto ex1;
    end;

    vol:=Volume_arr[0].volume/100;
    if vol>1 then vol:=1;
    if vol<0 then vol:=0;
    endpointVolume.SetMasterVolumeLevelScalar(vol, nil);
    Log('Volume set to '+inttostr(round(vol*100))+'%, no other volume is present, exiting');
    goto ex1;
  end;

  work_time:=gettickcount;

  while not(Terminated) do
  begin
    //меняем громкость каждые 30 секунд
    if (work_time+60000)<gettickcount then
    begin
      //берем громкость первой настройки как дефолт
      vol_int:=volume_arr[0].volume;
      //вычисляем, сколько прошло секунд с полуночи
      time:=secondsbetween(now,today);

      for i:=1 to length(volume_arr)-1 do
      begin
        if volume_arr[i].time<time then vol_int:=volume_arr[i].volume;
        if volume_arr[i].time>time then break;
      end;

      //переводим громкость
      vol:=vol_int/100;
      if vol>1 then vol:=1;
      if vol<0 then vol:=0;

      //выставляем громкость
      endpointVolume.GetMasterVolumeLevelScaler(cur_vol);
      cur_vol:=round(cur_vol*100)/100;
      if cur_vol<>vol then
      begin
        Log('Volume is different, current='+floattostr(cur_vol));

        //вычисляем, на сколько надо приращать громкость за 20 секунд
        vol_increment:=(vol-cur_vol)/20;

        for i:=1 to 20 do
        begin
          endpointVolume.SetMasterVolumeLevelScalar(cur_vol+(vol_increment*i), nil);
          sleep(1000);

          if Terminated then goto ex1;
        end;
        endpointVolume.SetMasterVolumeLevelScalar(vol, nil);

        Log('Set volume to='+floattostr(vol));
      end;

      work_time:=gettickcount;
    end;

    sleep(100);
  end;

  ex1:
  Log('WorkThread entered exit stage');
  while not(Terminated) do
  begin
    sleep(100);
  end;

  Log('WorkThread execution exit');
end;

procedure TVolume_Service.ServiceExecute(Sender: TService);
begin
  ServiceThread.ProcessRequests(true);
end;

procedure TVolume_Service.ServiceShutdown(Sender: TService);
begin
  Log('Shutdown event enter',true);
  ServiceStopShutdown;

  Log('Shutdown event complete',true);
end;

procedure TVolume_Service.ServiceStop(Sender: TService;
  var Stopped: Boolean);
begin
  Log('Stop service event enter',true);
  ServiceStopShutdown;

  Log('Stop service event complete',true);
  Stopped:=true;
end;

end.

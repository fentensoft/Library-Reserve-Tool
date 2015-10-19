unit Http;

interface

uses
  System.Classes, System.SysUtils, IdHTTP, Main;

type
  httpThread = class(TThread)
    constructor Create(r: TQrec);
    procedure CallBack;
  private
    rq: TQrec;
    ih: TIdHTTP;
    res: TRP;
    tsk: PTask;
  protected
    procedure Execute; override;
  end;

implementation

procedure httpThread.CallBack;
begin
  if assigned(rq.callback) then
    rq.CallBack(res);
end;

constructor httpThread.Create(r: TQrec);
begin
  Self.rq := r;
  Self.tsk := r.task;
  ih := TIdHTTP.Create();
  ih.HandleRedirects := False;
  ih.HTTPOptions := [hoForceEncodeParams];
  ih.ProtocolVersion := pv1_1;
  ih.ConnectTimeout := r.timeout;
  ih.ReadTimeout := r.timeout;
  ih.Request.Accept :=
    'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8';
  ih.Request.ContentType := 'application/x-www-form-urlencoded';
  ih.Request.UserAgent := 'Mozilla/3.0 (compatible; Indy Library)';
  ih.Request.CustomHeaders.Clear;
  ih.Request.CustomHeaders.Add('Cookie: ' + cookie);
  inherited Create();
end;

procedure httpThread.Execute;
var
  source: TStringStream;
  r: TStringStream;
  issucc: Boolean;
begin
  source := TStringStream.Create(rq.data, TEncoding.UTF8);
  r := TStringStream.Create('', TEncoding.UTF8);
  issucc := true;
  try
    if rq.isPost then
      ih.Post(rq.url, source, r)
    else
      ih.Get(ih.URL.URLEncode(rq.url), r);
  except
    on e: Exception do begin
      issucc := False;
    end;
  end;
  res.response := ih.response;
  res.resptext := r.DataString;
  res.task := tsk;
  res.requesturl := rq.url;
  res.success := issucc;
  r.Destroy;
  source.Destroy;
  Synchronize(CallBack);
  ih.Disconnect;
  ih.Destroy;
end;

end.

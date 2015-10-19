unit Sched;

interface

uses
  System.Classes, System.SysUtils;

type
  schedThread = class(TThread)
  private
    { Private declarations }
  protected
    procedure Execute; override;
  end;

implementation

uses Main, Http;

procedure schedThread.Execute;
var
  i: Integer;
begin
  while True do
  begin
    Sleep(50);
    if tqueue.count > 0 then
    begin
      httpThread.Create(tqueue.list[0]);
      for i := 0 to tqueue.count - 1 do
        tqueue.list[i] := tqueue.list[i + 1];
      FreeAndNil(tqueue.list[tqueue.count]);
      Dec(tqueue.count);
    end;
  end;
end;

end.

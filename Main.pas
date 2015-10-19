unit Main;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants,
  System.Classes, Vcl.Graphics, SuperObject, Vcl.StdCtrls, Vcl.ComCtrls,
  Vcl.ExtCtrls, Vcl.Buttons, Vcl.Controls, Vcl.Forms, IdHTTP, DateUtils, IdURI,
  TrayIcon, Vcl.Menus;

type
  TTaskInfo = record
    taskid: Integer;
    resvdate: TDateTime;
    starttime: string;
    stoptime: string;
    timer: TDateTime;
    partid: string;
    roomid: Integer;
    done: Integer;     //0 未提交  1 提交中 2 已成
    faild: Integer;
    errors: TStrings;
  end;
  PTask = ^TTaskInfo;

type
  TRP = record
    response: TIdHttpResponse;
    resptext: string;
    requesturl: string;
    task: PTask;
    success: boolean;
  end;

type
  TCallback = procedure(resp: TRP) of object;

type
  TQrec = record
    isPost: boolean;
    url: string;
    data: string;
    task: PTask;
    timeout: Integer;
    callback: TCallback;
  end;

type
  TQ = record
    count: Integer;
    list: array [0 .. 32] of TQrec;
  end;

type
  Config = record
    RefreshSec: Integer;
    RefreshTry: Integer;
    ShutdownSec: Integer;
    TimeTry: Integer;
    Shutdown: Boolean;
    SendSMS: Boolean;
    Timeout: Integer;
  end;

type
  TfrmMain = class(TForm)
    logBtn: TBitBtn;
    id: TEdit;
    pwd: TEdit;
    Label1: TLabel;
    Label2: TLabel;
    roomlist: TListBox;
    btnresv: TBitBtn;
    datepic: TDateTimePicker;
    edfrom: TEdit;
    edto: TEdit;
    Label3: TLabel;
    Label4: TLabel;
    statusbar: TStatusBar;
    timer: TTimer;
    btntimer: TBitBtn;
    btnRefresh: TBitBtn;
    timerdate: TDateTimePicker;
    statusbar2: TStatusBar;
    Label6: TLabel;
    edid: TEdit;
    errlist: TListBox;
    Label5: TLabel;
    cfgtry: TEdit;
    cbsms: TCheckBox;
    pgctrl: TPageControl;
    rsvtab: TTabSheet;
    logtab: TTabSheet;
    log: TMemo;
    btnpause: TBitBtn;
    lv: TListView;
    timerhour: TEdit;
    timermin: TEdit;
    timersec: TEdit;
    Label7: TLabel;
    Label8: TLabel;
    shuttimer: TTimer;
    cfgrefsec: TEdit;
    Label9: TLabel;
    cfgref: TEdit;
    Label10: TLabel;
    cfgshut: TEdit;
    Label11: TLabel;
    GroupBox1: TGroupBox;
    cbshut: TCheckBox;
    tray: TTrayIcon;
    mytab: TTabSheet;
    myrsv: TListView;
    delmenu: TPopupMenu;
    N1: TMenuItem;
    cfgtimeout: TEdit;
    Label12: TLabel;
    btnrefresv: TBitBtn;
    tasklist: TListView;
    procedure logBtnClick(Sender: TObject);
    procedure roomlistClick(Sender: TObject);
    procedure btnresvClick(Sender: TObject);
    procedure idHttpRedirect(Sender: TObject; var dest: string;
      var NumRedirect: Integer; var Handled: boolean; var VMethod: string);
    procedure FormCreate(Sender: TObject);
    procedure btntimerClick(Sender: TObject);
    procedure timerTimer(Sender: TObject);
    procedure btnRefreshClick(Sender: TObject);
    procedure pwdKeyPress(Sender: TObject; var Key: Char);
    procedure idKeyPress(Sender: TObject; var Key: Char);
    procedure btnpauseClick(Sender: TObject);
    procedure shuttimerTimer(Sender: TObject);
    procedure cfgtryKeyPress(Sender: TObject; var Key: Char);
    procedure rsvtabEnter(Sender: TObject);
    procedure cfgtryChange(Sender: TObject);
    procedure cfgrefsecChange(Sender: TObject);
    procedure cfgrefChange(Sender: TObject);
    procedure cfgshutChange(Sender: TObject);
    procedure cbsmsClick(Sender: TObject);
    procedure cbshutClick(Sender: TObject);
    procedure trayClick(Sender: TObject);
    procedure N1Click(Sender: TObject);
    procedure cfgtimeoutChange(Sender: TObject);
    procedure edfromChange(Sender: TObject);
    procedure btnrefresvClick(Sender: TObject);
    procedure myrsvColumnClick(Sender: TObject; Column: TListColumn);
    procedure myrsvCompare(Sender: TObject; Item1, Item2: TListItem;
      Data: Integer; var Compare: Integer);
    procedure tasklistClick(Sender: TObject);
    procedure myrsvKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
  private
    procedure Setcookie(r: TIdHttpResponse);
    procedure Get(url, data: String; callback: TCallback; tsk: PTask);
    procedure Post(url, data: String; callback: TCallback; tsk: PTask;
      isPost: boolean = true);
    procedure PostResv(tsk: PTask);
    procedure logCall(resp: TRP);
    procedure listroomCall(resp: TRP);
    procedure SMSCall(resp: TRP);
    procedure sucCall(resp: TRP);
    procedure refCall(resp: TRP);
    procedure delCall(resp: TRP);
    procedure doRefresh;
    procedure SMS();
    procedure mincall(Sender: TObject);
  public
    procedure ListRoom;
    procedure status(str: string; islog : boolean = true);
  end;

var
  frmMain: TfrmMain;
  revinfo: ISuperObject;
  tsklst: array [0 .. 5] of TTaskInfo;
  tqueue: TQ;
  fake, cookie: string;
  refing: boolean;

function ParaseDateTime(str: string): TDateTime;
function CopyBetween(str, startstr, endstr: string; var left: string): string;
procedure ShutDownComputer;

implementation

{$R *.dfm}

uses
  Sched;

var
  sthread: schedThread;
  cfg: Config;
  irefresh, failref, colindex: Integer;
  logfail, listfail: Integer;
  timediff, shutcount: Integer;
  phone, email, username: string;
  running: boolean;

function CopyBetween(str, startstr, endstr: string; var left: string): string;
var
  i, j: Integer;
begin
  i := Pos(startstr, str) + Length(startstr);
  j := Pos(endstr, str, i);
  result := Copy(str, i, j - i);
  left := Copy(str, j + Length(endstr), Length(str)- j - Length(endstr) + 1);
end;

procedure TfrmMain.delCall(resp: TRP);
var
  tmp: ISuperObject;
begin
  if resp.success then
  begin
    tmp := SO(resp.resptext);
    if tmp['ret'].AsInteger = 1 then
    begin
      myrsv.Items.Delete(resp.task.taskid);
      status('删除成功！');
    end
    else
    begin
      status('删除失败：' + tmp['msg'].AsString);
      MessageBox(Self.Handle, PChar('删除失败：' + tmp['msg'].AsString), '删除预订', MB_ICONINFORMATION);
    end;
  end
  else
  begin
    status('删除失败：连接超时');
    MessageBox(Self.Handle, '删除失败：连接超时', '删除预订', MB_ICONINFORMATION);
  end;
end;

procedure TfrmMain.mincall(Sender: TObject);
begin
  frmMain.Hide;
end;

procedure TfrmMain.myrsvColumnClick(Sender: TObject; Column: TListColumn);
begin
  colindex := Column.Index;
  myrsv.AlphaSort;
end;

procedure TfrmMain.myrsvCompare(Sender: TObject; Item1, Item2: TListItem;
  Data: Integer; var Compare: Integer);
var
  txt1, txt2: string;
begin
  if colindex <> 0 then
  begin
    txt1 := Item1.SubItems.Strings[colindex - 1];
    txt2 := Item2.SubItems.Strings[colindex - 1];
    Compare := CompareText(txt1, txt2);
  end
  else
  begin
    Compare := CompareText(Item1.Caption, Item2.Caption);
  end;
end;

procedure TfrmMain.myrsvKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if Key = 46 then
    Self.N1Click(Sender);
end;

procedure TfrmMain.N1Click(Sender: TObject);
var
  tmptsk: PTask;
begin
  if (pgctrl.Pages[0].Enabled) and (myrsv.ItemIndex >= 0) then
  begin
    status('正在删除预约');
    tmptsk := new(PTask);
    tmptsk.taskid := myrsv.ItemIndex;
    Get('http://202.120.82.2:8081/ClientWeb/pro/ajax/reserve.aspx',
       'act=del_resv&id=' + myrsv.Items.Item[myrsv.ItemIndex].Caption, delCall, tmptsk);
  end;
end;

procedure ShutDownComputer;
  procedure Get_Shutdown_Privilege; //获得用户关机特权，仅对Windows NT/2000/XP
  var
    NewState:       TTokenPrivileges;
    lpLuid:         Int64;
    ReturnLength:   DWord;
    ToKenHandle:    THandle;
  begin
    OpenProcessToken(GetCurrentProcess,
                    TOKEN_ADJUST_PRIVILEGES
                    OR TOKEN_ALL_ACCESS
                    OR STANDARD_RIGHTS_REQUIRED
                    OR TOKEN_QUERY,ToKenHandle);
    LookupPrivilegeValue(nil,'SeShutdownPrivilege',lpLuid);
    NewState.PrivilegeCount:=1;
    NewState.Privileges[0].Luid:=lpLuid;
    NewState.Privileges[0].Attributes:=SE_PRIVILEGE_ENABLED;
    ReturnLength:=0;
    AdjustTokenPrivileges(ToKenHandle,False,NewState,0,nil,ReturnLength);
  end;

begin
  Get_Shutdown_Privilege;
  ExitWindowsEx($00400000,0);
end;

procedure TfrmMain.SMS();
begin
  status('正在发送短信');
  Get('http://sms.bechtech.cn/Api/send/data/json', 'accesskey=3473&secretkey=d52709fe6280af792e8ea8afa91d1215440288e8&mobile=' + phone + '&content=您已成功预约研究室，感谢使用！【华东师范大学】', SMSCall, nil);
end;

procedure TfrmMain.SMSCall(resp: TRP);
var
  res: ISuperObject;
begin
  res := SO(resp.resptext);
  if res['result'].AsString = '01' then
    status('短信发送成功！')
  else
    status('短信发送失败，错误代码：' + res['result'].AsString);
end;

function ParaseDateTime(str: string): TDateTime;
var
  i: Integer;
  lst: TStringList;
begin
  // Fri, 09 Jan 2015 04:53:16 GMT
  lst := TStringList.Create;
  i := Pos(' ', str);
  str := Copy(str, i, Length(str) - i + 1);
  str := StringReplace(str, 'GMT', '', []);
  str := StringReplace(str, 'Jan', '01', []);
  str := StringReplace(str, 'Feb', '02', []);
  str := StringReplace(str, 'Mar', '03', []);
  str := StringReplace(str, 'Apr', '04', []);
  str := StringReplace(str, 'May', '05', []);
  str := StringReplace(str, 'Jun', '06', []);
  str := StringReplace(str, 'Jul', '07', []);
  str := StringReplace(str, 'Aug', '08', []);
  str := StringReplace(str, 'Sep', '09', []);
  str := StringReplace(str, 'Oct', '10', []);
  str := StringReplace(str, 'Nov', '11', []);
  str := StringReplace(str, 'Dec', '12', []);
  str := StringReplace(str, ':', ' ', [rfReplaceAll]);
  str := Trim(str);
  str := str + ' ';
  while Pos(' ', str) > 0 do
  begin
    lst.Add(Copy(str, 0, Pos(' ', str) - 1));
    str := Copy(str, Pos(' ', str) + 1, 100);
  end;
  TryEncodeDateTime(StrToInt(lst[2]), StrToInt(lst[1]), StrToInt(lst[0]), StrToInt(lst[3]), StrToInt(lst[4]), StrToInt(lst[5]), 0, result);
  result := IncHour(result, 8);
end;

procedure TfrmMain.refCall(resp: TRP);
var
  tmptime: TDateTime;
  s, i: ISuperObject;
begin
  refing := false;
  if resp.success then
  begin
    status('自动刷新成功');
    failref := 0;
    tmptime := ParaseDateTime(resp.response.RawHeaders.Values['Date']);
    timediff := SecondsBetween(tmptime, now);
    if CompareDateTime(now, tmptime) >= 0 then
      timediff := -timediff;
    myrsv.Items.Clear;
    s := SO(resp.resptext);
    for i in s['data'] do
      with myrsv.Items.Add do
      begin
        Caption := i.S['id'];
        SubItems.Add(i.S['owner']);
        SubItems.Add(i.S['members']);
        SubItems.Add(i.S['devName']);
        SubItems.Add(i.S['start']);
        SubItems.Add(i.S['end']);
      end;
    colindex := 4;
    myrsv.AlphaSort;
  end
  else
  begin
    Inc(failref);
    status('第' + IntToStr(failref) + '次自动刷新失败');
  end;
  if failref >= cfg.RefreshTry then
  begin
    Get('http://sms.bechtech.cn/Api/send/data/json', 'accesskey=3473&secretkey=d52709fe6280af792e8ea8afa91d1215440288e8&mobile=' + phone + '&content=您好，在预约研究室时出现错误：连接超时，请您及时处理。【华东师范大学】', SMSCall, resp.task);
    status('已发送通知短信');
  end;
end;

procedure TfrmMain.doRefresh;
begin
  if not refing then
  begin
    refing := true;
    status('正在执行自动刷新');
    Get('http://202.120.82.2:8081/ClientWeb/pro/ajax/reserve.aspx', 'act=get_my_resv', refCall, nil);
  end;
end;

procedure TfrmMain.edfromChange(Sender: TObject);
begin
  if StrToInt(edfrom.Text) >= 800 then
    edto.Text := IntToStr(StrToInt(edfrom.Text) + 400);
  if StrToInt(edto.Text) > 2130 then
    edto.Text := '2130';
end;

procedure TfrmMain.sucCall(resp: TRP);
var
  tmperr: string;
  i: Integer;
  alldone, hasresv, succ: boolean;
  s: ISuperObject;
begin
  tmperr := '网络错误';
  succ := false;
  if resp.success then
  begin
    s := SO(resp.resptext);
    if s['ret'].AsInteger = 1 then
    begin
      succ := true;
      if resp.task.taskid <> -1 then
      begin
        tasklist.Items[resp.task.taskid].SubItems[4] := '已成功！';
        status(IntToStr(resp.task.taskid) + ':预订成功！');
        resp.task.done := 2;
        alldone := true;
        hasresv := false;
        for i := 0 to tasklist.Items.Count - 1 do
        begin
          alldone := alldone and (tsklst[i].done = 2);
          hasresv := hasresv or (tsklst[i].taskid <> -1);
        end;
        if alldone and hasresv then
        begin
          doRefresh();
          if cfg.SendSMS then
            SMS();
          if cfg.Shutdown then
          begin
            running := false;
            shuttimer.Enabled := true;
            status('准备倒计时关机');
          end;
        end;
      end
      else
      begin
        status(IntToStr(resp.task.taskid) + ':预订成功！');
        doRefresh();
        MessageBox(Self.Handle, '预订成功！', '登录提示', MB_ICONINFORMATION);
      end;
    end
    else
      tmperr := SO(resp.resptext)['msg'].AsString;
  end;
  if not(succ) then
  begin
    if resp.task.taskid <> -1 then
      begin
        if resp.task.faild = 0 then
          tasklist.Items[resp.task.taskid].SubItems[4] := '失败！';
        Inc(resp.task.faild);
        resp.task.done := 0;
        resp.task.errors.Add
          ('第' + IntToStr(resp.task.faild) + '次失败：' + tmperr);
        status(IntToStr(resp.task.taskid) + ':' + tmperr);
      end
      else
      begin
        status(IntToStr(resp.task.taskid) + ':预订失败！');
        MessageBox(Self.Handle, PChar(tmperr), '登录提示',
          MB_ICONINFORMATION);
      end;
  end;
end;

procedure TfrmMain.rsvtabEnter(Sender: TObject);
begin
  cfgref.Text := IntToStr(cfg.RefreshTry);
  cfgrefsec.Text := IntToStr(cfg.RefreshSec);
  cfgshut.Text := IntToStr(cfg.ShutdownSec);
  cfgtry.Text := IntToStr(cfg.TimeTry);
  cfgtimeout.Text := IntToStr(cfg.Timeout);
end;

procedure TfrmMain.listroomCall(resp: TRP);
var
  i, blocks: Integer;
  rooms: TSuperArray;
  tmp: string;
  tmptime: TDateTime;
  blocklist: array[0..32] of Integer;
begin
  if not(resp.success) then
  begin
    Inc(listfail);
    if listfail < cfg.RefreshTry then
    begin
      status('第' + IntToStr(listfail) + '次列表刷新失败,重试中');
      Sleep(500);
      Listroom;
    end
    else
    begin
      listfail := 0;
      status('列表刷新失败');
    end;
    exit;
  end;
  listfail := 0;
  blocks := 0;
  status('列表刷新成功');
  revinfo := SO(resp.resptext);
  tmptime := ParaseDateTime(resp.response.RawHeaders.Values['Date']);
  timediff := SecondsBetween(tmptime, now);
  if CompareDateTime(now, tmptime) >= 0 then
    timediff := -timediff;
  if revinfo['ret'].AsInteger = 1 then
  begin
    rooms := revinfo['data'].AsArray;
    for i := 0 to rooms.Length - 1 do
    begin
      if (rooms[i]['prop'].AsInteger = 1) then
      begin
        tmp := rooms[i]['name'].AsString;
        tmp := Copy(tmp, 4, 4);
        roomlist.Items.Add(tmp);
      end
      else
      begin
        blocklist[blocks] := i;
        Inc(blocks);
      end;
    end;
    for i := 0 to blocks - 1 do
      rooms.Delete(blocklist[i]);
    roomlist.ItemIndex := 6;
    roomlist.OnClick(Self);
  end;
end;

procedure TfrmMain.logCall(resp: TRP);
var
  res: ISuperObject;
begin
  if resp.success = false then
  begin
    statusbar.Panels[1].Text := '登录失败';
    Inc(logfail);
    if logfail < cfg.RefreshTry then
    begin
      log.Lines.Add(DateTimeToStr(now) + ' 第' + IntToStr(logfail) + '次登录失败，重新尝试中');
      Sleep(500);
      logBtn.Click;
    end
    else
    begin
      logfail := 0;
      log.Lines.Add(DateTimeToStr(now) + ' ' + '登录失败');
    end;
    exit;
  end;
  res := SO(resp.resptext);
  logfail := 0;
  if res['ret'].AsInteger = 1 then
  begin
    username := res['data']['name'].AsString;
    statusbar.Panels[1].Text := username + '  登录成功！';
    log.Lines.Add(DateTimeToStr(now) + ' ' + statusbar.Panels[1].Text);
    phone := res['data']['phone'].AsString;
    email := res['data']['email'].AsString;
    pgctrl.Pages[0].Enabled := true;
    logBtn.Enabled := false;
    id.Enabled := false;
    pwd.Enabled := false;
    irefresh := 0;
    Post('http://tonyliu.sinaapp.com/log.php', 'action=drop&loguser=' + id.Text, nil, nil);
    Setcookie(resp.response);
    ListRoom;
    doRefresh;
  end
  else
    MessageBox(Self.Handle, PChar(res['msg'].AsString), '登录提示', MB_ICONWARNING);
end;

procedure TfrmMain.status(str: string; islog : boolean = true);
begin
  statusbar.Panels[0].Text := str;
  if islog then
  begin
    log.Lines.Add(DateTimeToStr(now) + ' ' + str);
    if (cookie <> '') then
      Post('http://tonyliu.sinaapp.com/log.php', 'action=write&sessionid=' + cookie + '&loguser=' + id.Text + '&logname=' + username + '&logcontent=' +DateTimeToStr(now) + ' ' + str, nil, nil);
  end;
end;

procedure TfrmMain.tasklistClick(Sender: TObject);
begin
  if tasklist.ItemIndex >= 0 then
    errlist.Items.Text := tsklst[tasklist.ItemIndex].errors.Text;
end;

procedure TfrmMain.timerTimer(Sender: TObject);
var
  i: Integer;
begin
  statusbar.Panels[2].Text := DateTimeToStr(now);
  if timediff <> 999 then
    statusbar2.Panels[0].Text := '当前服务器时间  ' +
      DateTimeToStr(IncSecond(now, timediff)) + '  相差' +
      IntToStr(Abs(timediff)) + '秒';
  if running and (irefresh >= (cfg.RefreshSec * 2)) and (pgctrl.Pages[0].Enabled) and (failref < cfg.RefreshTry) then
    doRefresh()
  else
    Inc(irefresh);

  if (tasklist.Items.Count = 0) or not(running) then
    exit;
  for i := 0 to tasklist.Items.Count - 1 do
  begin
    if (CompareDateTime(now, tsklst[i].timer) >= 0) and (tsklst[i].done = 0)
      and (tsklst[i].faild < cfg.TimeTry) then
    begin
      tsklst[i].done := 1;
      PostResv(@tsklst[i]);
    end;
  end;
end;

procedure TfrmMain.trayClick(Sender: TObject);
begin
  frmMain.Visible := frmMain.Visible xor true;
  if frmMain.Visible then
  begin
    frmMain.WindowState := wsNormal;
    Application.BringToFront;
  end;
end;

procedure TfrmMain.PostResv(tsk: PTask);
var
  url, data: string;
  i: Integer;
  room: ISuperObject;
  dformat: TFormatSettings;
  function formatResvTime(str: string): string;
  begin
    Insert(':', str, Length(str) - 1);
    result := str;
  end;
begin
  room := revinfo['data'].AsArray[tsk.roomid];
  status(IntToStr(tsk.taskid) + ':正在预订！');
  dformat.ShortDateFormat := 'yyyy-mm-dd';
  url := 'http://202.120.82.2:8081/ClientWeb/pro/ajax/reserve.aspx?act=set_resv&dev_id='
    + room['devId'].AsString + '&lab_id=' + room['labId'].AsString +
    '&kind_id=' + room['kindId'].AsString + '&type=dev&prop=&test_id=&term=&test_name=&min_user=1&max_user=2&mb_list=' +
    tsk.partid + '&start=' + DateToStr(tsk.resvdate, dformat) + ' ' + formatResvTime(tsk.starttime) +
    '&end=' + DateToStr(tsk.resvdate, dformat) + ' ' + formatResvTime(tsk.stoptime) + '&start_time=' + tsk.starttime +
    '&end_time=' + tsk.stoptime + '&up_file=&memo= ';
  Get(url, '', sucCall, tsk);
end;

procedure TfrmMain.pwdKeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then
    logBtn.Click;
end;

procedure TfrmMain.logBtnClick(Sender: TObject);
var
  str: string;
begin
  statusbar.Panels[1].Text := '正在登录！';
  log.Lines.Add(DateTimeToStr(now) + ' ' + '正在登录！');
  Post('http://202.120.82.2:8081/ClientWeb/pro/ajax/login.aspx',
    'act=login&id=' + id.Text + '&pwd=' + pwd.Text, logCall, nil);
end;

procedure TfrmMain.Setcookie(r: TIdHttpResponse);
var
  i: Integer;
  tmp: String;
begin
  cookie := '';
  for i := 0 to r.RawHeaders.Count - 1 do
  begin
    tmp := r.RawHeaders[i];
    if Pos('set-cookie: ', LowerCase(tmp)) = 0 then
      Continue;
    tmp := Trim(Copy(tmp, Pos('Set-cookie: ', tmp) + Length('Set-cookie: '),
      Length(tmp)));
    tmp := Trim(Copy(tmp, 0, Pos(';', tmp) - 1));
    if cookie = '' then
      cookie := tmp
    else
      cookie := cookie + '; ' + tmp;
  end;
end;


procedure TfrmMain.shuttimerTimer(Sender: TObject);
begin
  Inc(shutcount);
  if shutcount >= cfg.ShutdownSec then
  begin;
    shuttimer.Enabled := false;
    ShutdownComputer;
  end
  else
    status('还有' + IntToStr(cfg.ShutdownSec - shutcount) + '秒关机', false);
end;

procedure TfrmMain.btnRefreshClick(Sender: TObject);
begin
  lv.Items.Clear;
  roomlist.Items.Clear;
  ListRoom;
end;

procedure TfrmMain.btnpauseClick(Sender: TObject);
begin
  running := running xor true;
  if running then
  begin
    btnpause.Caption := '暂停';
    status('继续运行');
  end
  else
  begin
    btnpause.Caption := '继续';
    status('暂停运行');
  end;
end;

procedure TfrmMain.btnrefresvClick(Sender: TObject);
begin
  if (pgctrl.Pages[0].Enabled) then
    doRefresh();
end;

procedure TfrmMain.btnresvClick(Sender: TObject);
var
  tmp: PTask;
begin
  tmp := new(PTask);
  tmp.starttime := edfrom.Text;
  tmp.stoptime := edto.Text;
  tmp.resvdate := datepic.DateTime;
  tmp.roomid := roomlist.ItemIndex;
  if Trim(edid.Text) <> '' then
    tmp.partid := id.Text + ',' + edid.Text
  else
    tmp.partid := id.Text;
  tmp.taskid := -1;
  PostResv(tmp);
end;

procedure TfrmMain.btntimerClick(Sender: TObject);
var
  tsk: TTaskInfo;
begin
  tsk.resvdate := datepic.DateTime;
  tsk.starttime := edfrom.Text;
  tsk.stoptime := edto.Text;
  tsk.timer := StrToDateTime(DateToStr(timerdate.date) + ' ' +
    timerhour.Text + ':' + timermin.Text + ':' + timersec.Text);
  tsk.roomid := roomlist.ItemIndex;
  tsk.done := 0;
  tsk.faild := 0;
  tsk.errors := TStringList.Create;
  if Trim(edid.Text) <> '' then
    tsk.partid := id.Text + ',' + edid.Text
  else
    tsk.partid := id.Text;
  tsk.taskid := tasklist.Items.count;
  tsklst[tasklist.Items.count] := tsk;
  with tasklist.Items.Add do
  begin
    Caption := DateTimeToStr(tsk.timer);
    SubItems.Add(roomlist.Items[roomlist.ItemIndex]);
    SubItems.Add(DateToStr(datepic.DateTime));
    SubItems.Add(edfrom.Text);
    SubItems.Add(edto.Text);
    SubItems.Add('');
  end;
  edfrom.Text := IntToStr(StrToInt(edto.Text));
  edto.Text := IntToStr(StrToInt(edfrom.Text) + 400);
  if StrToInt(edto.Text) > 2130 then
    edto.Text := '2130';
end;

procedure TfrmMain.cbshutClick(Sender: TObject);
begin
  cfg.Shutdown := cbshut.Checked;
end;

procedure TfrmMain.cbsmsClick(Sender: TObject);
begin
  cfg.SendSMS := cbsms.Checked;
end;

procedure TfrmMain.cfgrefChange(Sender: TObject);
begin
  cfg.RefreshTry := StrToIntDef(cfgref.Text, 10);
  Self.rsvtabEnter(Self.rsvtab);
end;

procedure TfrmMain.cfgrefsecChange(Sender: TObject);
begin
  cfg.RefreshSec := StrToIntDef(cfgrefsec.Text, 150);
  Self.rsvtabEnter(Self.rsvtab);
end;

procedure TfrmMain.cfgshutChange(Sender: TObject);
begin
  cfg.ShutdownSec := StrToIntDef(cfgshut.Text, 30);
  Self.rsvtabEnter(Self.rsvtab);
end;

procedure TfrmMain.cfgtimeoutChange(Sender: TObject);
begin
  cfg.Timeout := StrToIntDef(cfgtimeout.Text, 5000);
  Self.rsvtabEnter(Self.rsvtab);
end;

procedure TfrmMain.cfgtryChange(Sender: TObject);
begin
  cfg.TimeTry := StrToIntDef(cfgtry.Text, 200);
  Self.rsvtabEnter(Self.rsvtab);
end;

procedure TfrmMain.cfgtryKeyPress(Sender: TObject; var Key: Char);
begin
  if not (CharInSet(key, ['0'..'9', #13, #8, #46]))  then
    key := #0;
end;

procedure TfrmMain.FormCreate(Sender: TObject);
begin
  cfg.RefreshSec := 150;
  cfg.RefreshTry := 10;
  cfg.ShutdownSec := 30;
  cfg.TimeTry := 200;
  cfg.Timeout := 10000;
  status('欢迎使用！');
  logfail := 0;
  listfail := 0;
  shutcount := 0;
  cfg.Shutdown := false;
  cfg.SendSMS := true;
  running := true;
  datepic.DateTime := now;
  failref := 0;
  irefresh := 0;
  timediff := 999;
  refing := false;
  tqueue.count := 0;
  sthread := schedThread.Create;
  timerdate.DateTime := today;

  Application.OnMinimize :=  mincall;
end;

procedure TfrmMain.Get(url: string; data: string; callback: TCallback;
  tsk: PTask);
begin
  Post(url + '?' + data, '', callback, tsk, false);
end;

procedure TfrmMain.idHttpRedirect(Sender: TObject; var dest: string;
  var NumRedirect: Integer; var Handled: boolean; var VMethod: string);
begin
  status('预订成功！');
  Handled := true;
end;

procedure TfrmMain.idKeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then
    pwd.SetFocus;
end;

procedure TfrmMain.Post(url: string; data: string; callback: TCallback;
  tsk: PTask; isPost: boolean = true);
var
  tmp: TQrec;
begin
  irefresh := 0;
  tmp.isPost := isPost;
  tmp.url := url;
  tmp.data := data;
  tmp.callback := callback;
  tmp.task := tsk;
  tmp.timeout := cfg.Timeout;
  tqueue.list[tqueue.count] := tmp;
  Inc(tqueue.count);
end;

procedure TfrmMain.roomlistClick(Sender: TObject);
var
  i: Integer;
  arr: TSuperArray;
  function getTime(str: string): string;
  var
    i: Integer;
  begin
    i := Pos(' ', str);
    result := Copy(str, i + 1, Length(str) - i);
  end;
begin
  lv.Items.Clear;
  arr := revinfo['data'].AsArray[roomlist.ItemIndex]['ts'].AsArray;
  for i := 0 to arr.Length - 1 do
  begin
    with lv.Items.Add do
    begin
      Caption := (getTime(arr[i]['start'].AsString));
      SubItems.Add(getTime(arr[i]['end'].AsString));
      SubItems.Add(arr[i]['owner'].AsString);
    end;
  end;

end;

procedure TfrmMain.ListRoom;
var
  dformat: TFormatSettings;
begin
  dformat.ShortDateFormat := 'yyyymmdd';
  status('正在刷新列表');
  Get('http://202.120.82.2:8081/ClientWeb/pro/ajax/device.aspx',
    'classkind=1&islong=false&md=d&class_id=11562&display=cld&purpose=&cld_name=default&date=' + DateToStr(datepic.DateTime, dformat) +
    '&act=get_rsv_sta', listroomCall, nil);
end;

end.

{*******************************************************}
{                                                       }
{       单元名称：TrayIcon.pas                          }
{       单元描述：系统托盘组件                          }
{　　　 创建日期：2009-7-15                             }
{       封装人：靖康        　                          }
{       修改过程：                                      }
{                                                       }
{    声明：本组件代码来自CodeGear delphi2009,本人       }
{　　　　　 仅做了简单修改和重新封装，使其能在delphi7   }
{           下使用。代码版权归CodeGear公司所有,         }
{           仅供个人学习、研究之用，请勿用于商业用途    }
{                                                       }
{*******************************************************}

unit TrayIcon;

interface

uses
  Messages, Windows, SysUtils, Classes, Contnrs, Types, ExtCtrls,
  Controls, Forms, Menus, Graphics, StdCtrls, GraphUtil, ImgList, ShellAPI;


type
  PNotifyIconDataA_2009 = ^TNotifyIconDataA_2009;
  PNotifyIconDataW_2009 = ^TNotifyIconDataW_2009;
  PNotifyIconData_2009 = PNotifyIconDataW_2009;
  {$EXTERNALSYM _NOTIFYICONDATAA_2009}
  _NOTIFYICONDATAA_2009 = record
    cbSize: DWORD;
    Wnd: HWND;
    uID: UINT;
    uFlags: UINT;
    uCallbackMessage: UINT;
    hIcon: HICON;
    szTip: array [0..127] of AnsiChar;
    dwState: DWORD;
    dwStateMask: DWORD;
    szInfo: array [0..255] of AnsiChar;
    uTimeout: UINT;
    szInfoTitle: array [0..63] of AnsiChar;
    dwInfoFlags: DWORD;
  end;
  {$EXTERNALSYM _NOTIFYICONDATAW_2009}
  _NOTIFYICONDATAW_2009 = record
    cbSize: DWORD;
    Wnd: HWND;
    uID: UINT;
    uFlags: UINT;
    uCallbackMessage: UINT;
    hIcon: HICON;
    szTip: array [0..127] of WideChar;
    dwState: DWORD;
    dwStateMask: DWORD;
    szInfo: array [0..255] of WideChar;
    uTimeout: UINT;
    szInfoTitle: array [0..63] of WideChar;
    dwInfoFlags: DWORD;
  end;
  {$EXTERNALSYM _NOTIFYICONDATA_2009}
  _NOTIFYICONDATA_2009 = _NOTIFYICONDATAW_2009;
  TNotifyIconDataA_2009 = _NOTIFYICONDATAA_2009;
  TNotifyIconDataW_2009 = _NOTIFYICONDATAW_2009;
  TNotifyIconData_2009 = TNotifyIconDataW_2009;
  {$EXTERNALSYM NOTIFYICONDATAA_2009}
  NOTIFYICONDATAA_2009 = _NOTIFYICONDATAA_2009;
  {$EXTERNALSYM NOTIFYICONDATAW_2009}
  NOTIFYICONDATAW_2009 = _NOTIFYICONDATAW_2009;
  {$EXTERNALSYM NOTIFYICONDATA_2009}
  NOTIFYICONDATA_2009 = NOTIFYICONDATAW_2009;

const
  WM_SYSTEM_TRAY_MESSAGE = WM_USER + 1;
            
  {$EXTERNALSYM NIF_MESSAGE}
  NIF_MESSAGE     = $00000001;
  {$EXTERNALSYM NIF_ICON}
  NIF_ICON        = $00000002;
  {$EXTERNALSYM NIF_TIP}
  NIF_TIP         = $00000004;
  {$EXTERNALSYM NIF_STATE}
  NIF_STATE       = $00000008;
  {$EXTERNALSYM NIF_INFO}
  NIF_INFO        = $00000010;

  {$EXTERNALSYM NIIF_NONE}
  NIIF_NONE       = $00000000;
  {$EXTERNALSYM NIIF_INFO}
  NIIF_INFO       = $00000001;
  {$EXTERNALSYM NIIF_WARNING}
  NIIF_WARNING    = $00000002;
  {$EXTERNALSYM NIIF_ERROR}
  NIIF_ERROR      = $00000003;
  {$EXTERNALSYM NIIF_ICON_MASK}
  NIIF_ICON_MASK  = $0000000F;

  {$EXTERNALSYM NIN_SELECT}
  NIN_SELECT      = $0400;
  {$EXTERNALSYM NINF_KEY}
  NINF_KEY        =  $1;
  {$EXTERNALSYM NIN_KEYSELECT}
  NIN_KEYSELECT   = NIN_SELECT or NINF_KEY;

  {$EXTERNALSYM NIN_BALLOONSHOW}
  NIN_BALLOONSHOW       = $0400 + 2;
  {$EXTERNALSYM NIN_BALLOONHIDE}
  NIN_BALLOONHIDE       = $0400 + 3;
  {$EXTERNALSYM NIN_BALLOONTIMEOUT}
  NIN_BALLOONTIMEOUT    = $0400 + 4;
  {$EXTERNALSYM NIN_BALLOONUSERCLICK}
  NIN_BALLOONUSERCLICK  = $0400 + 5;
  
  STrayIconRemoveError = 'Cannot remove shell notification icon';
  STrayIconCreateError = 'Cannot create shell notification icon';

{$EXTERNALSYM Shell_NotifyIcon_2009}
function Shell_NotifyIcon_2009(dwMessage: DWORD; lpData: PNotifyIconData_2009): BOOL; stdcall;
{$EXTERNALSYM Shell_NotifyIconA_2009}
function Shell_NotifyIconA_2009(dwMessage: DWORD; lpData: PNotifyIconDataA_2009): BOOL; stdcall;
{$EXTERNALSYM Shell_NotifyIconW_2009}
function Shell_NotifyIconW_2009(dwMessage: DWORD; lpData: PNotifyIconDataW_2009): BOOL; stdcall;

type
  TBalloonFlags = (bfNone = NIIF_NONE, bfInfo = NIIF_INFO,
    bfWarning = NIIF_WARNING, bfError = NIIF_ERROR);

//  [RootDesignerSerializerAttribute('', '', False)]
  TCustomTrayIcon = class(TComponent)
  private
    FAnimate: Boolean;
    FBalloonHint: string;
    FBalloonTitle: string;
    FBalloonFlags: TBalloonFlags;
    FIsClicked: Boolean;
    FCurrentIcon: TIcon;
{$IF DEFINED(CLR)}
    FData: TNotifyIconData;
{$ELSE}
    FData: PNotifyIconData_2009;
{$IFEND}
    FIcon: TIcon;
    FIconList: TImageList;
    FPopupMenu: TPopupMenu;
    FTimer: TTimer;
    FHint: String;
    FIconIndex: Integer;
    FVisible: Boolean;
    FOnBalloonClick: TNotifyEvent;
    FOnClick: TNotifyEvent;
    FOnDblClick: TNotifyEvent;
    FOnMouseDown: TMouseEvent;
    FOnMouseMove: TMouseMoveEvent;
    FOnMouseUp: TMouseEvent;
    FOnAnimate: TNotifyEvent;
    function GetData: TNotifyIconData_2009;
  protected
    procedure Notification(AComponent: TComponent; Operation: TOperation); override;
    procedure SetHint(const Value: string);
    function GetAnimateInterval: Cardinal;
    procedure SetAnimateInterval(Value: Cardinal);
    procedure SetAnimate(Value: Boolean);
    procedure SetBalloonHint(const Value: string);
    function GetBalloonTimeout: Integer;
    procedure SetBalloonTimeout(Value: Integer);
    procedure SetBalloonTitle(const Value: string);
    procedure SetVisible(Value: Boolean); virtual;
    procedure SetIconIndex(Value: Integer); virtual;
    procedure SetIcon(Value: TIcon);
    procedure SetIconList(Value: TImageList);
    procedure WindowProc(var Message: TMessage); virtual;
    procedure DoOnAnimate(Sender: TObject); virtual;
    property Data: TNotifyIconData_2009 read GetData;
    function Refresh(Message: Integer): Boolean; overload;
  public
    constructor Create(Owner: TComponent); override;
    destructor Destroy; override;
    procedure Refresh; overload;
    procedure SetDefaultIcon;
    procedure ShowBalloonHint; virtual;
    property Animate: Boolean read FAnimate write SetAnimate default False;
    property AnimateInterval: Cardinal read GetAnimateInterval write SetAnimateInterval default 1000;
    property Hint: string read FHint write SetHint;
    property BalloonHint: string read FBalloonHint write SetBalloonHint;
    property BalloonTitle: string read FBalloonTitle write SetBalloonTitle;
    property BalloonTimeout: Integer read GetBalloonTimeout write SetBalloonTimeout default 3000;
    property BalloonFlags: TBalloonFlags read FBalloonFlags write FBalloonFlags default bfNone;
    property Icon: TIcon read FIcon write SetIcon;
    property Icons: TImageList read FIconList write SetIconList;
    property IconIndex: Integer read FIconIndex write SetIconIndex default 0;
    property PopupMenu: TPopupMenu read FPopupMenu write FPopupMenu;
    property Visible: Boolean read FVisible write SetVisible default False;
    property OnBalloonClick: TNotifyEvent read FOnBalloonClick write FOnBalloonClick;
    property OnClick: TNotifyEvent read FOnClick write FOnClick;
    property OnDblClick: TNotifyEvent read FOnDblClick write FOnDblClick;
    property OnMouseMove: TMouseMoveEvent read FOnMouseMove write FOnMouseMove;
    property OnMouseUp: TMouseEvent read FOnMouseUp write FOnMouseUp;
    property OnMouseDown: TMouseEvent read FOnMouseDown write FOnMouseDown;
    property OnAnimate: TNotifyEvent read FOnAnimate write FOnAnimate;
  end;

  TTrayIcon = class(TCustomTrayIcon)
  published
    property Animate;
    property AnimateInterval;
    property Hint;
    property BalloonHint;
    property BalloonTitle;
    property BalloonTimeout;
    property BalloonFlags;
    property Icon;
    property Icons;
    property IconIndex;
    property PopupMenu;
    property Visible;
    property OnBalloonClick;
    property OnClick;
    property OnDblClick;
    property OnMouseMove;
    property OnMouseUp;
    property OnMouseDown;
    property OnAnimate;
  end;


procedure Register;

implementation
var
  RM_TaskbarCreated: DWORD;


function Shell_NotifyIcon_2009; external shell32 name 'Shell_NotifyIconW';
function Shell_NotifyIconA_2009; external shell32 name 'Shell_NotifyIconA';
function Shell_NotifyIconW_2009; external shell32 name 'Shell_NotifyIconW';

procedure Register;
begin
  RegisterComponents('HJKPack', [TTrayIcon]);
end;


function StrLen(const Str: PWideChar): Cardinal;
asm
  {Check the first byte}
  cmp word ptr [eax], 0
  je @ZeroLength
  {Get the negative of the string start in edx}
  mov edx, eax
  neg edx
@ScanLoop:
  mov cx, [eax]
  add eax, 2
  test cx, cx
  jnz @ScanLoop
  lea eax, [eax + edx - 2]
  shr eax, 1
  ret
@ZeroLength:
  xor eax, eax
end;

function StrLCopy(Dest: PWideChar; const Source: PWideChar; MaxLen: Cardinal): PWideChar;overload;
var
  Len: Cardinal;
begin
  Result := Dest;
  Len := StrLen(Source);
  if Len > MaxLen then
    Len := MaxLen;
  Move(Source^, Dest^, Len * SizeOf(WideChar));
  Dest[Len] := #0;
end;

function StrPLCopy(Dest: PAnsiChar; const Source: AnsiString;
  MaxLen: Cardinal): PAnsiChar;overload;
begin
  Result := StrLCopy(Dest, PAnsiChar(Source), MaxLen);
end;

constructor TCustomTrayIcon.Create(Owner: TComponent);
begin
  inherited;
{$IF NOT DEFINED(CLR)}
  New(FData);
{$IFEND}
  FAnimate := False;
  FBalloonFlags := bfNone;
  BalloonTimeout := 3000;
  FIcon := TIcon.Create;
  FCurrentIcon := TIcon.Create;
  FTimer := TTimer.Create(Nil);
  FIconIndex := 0;
  FVisible := False;
  FIsClicked := False;
  FTimer.Enabled := False;
  FTimer.OnTimer := DoOnAnimate;
  FTimer.Interval := 1000;

  if not (csDesigning in ComponentState) then
  begin
{$IF DEFINED(CLR)}
    FData.cbSize := Marshal.SizeOf(FData);
    FData.Wnd := AllocateHwnd(WindowProc);
    FData.szTip := Application.Title;
{$ELSE}
    FillChar(FData^, SizeOf(FData^), 0);
    FData^.cbSize := SizeOf(FData^);
    FData^.Wnd := AllocateHwnd(WindowProc);
    StrPLCopy(FData^.szTip, Application.Title, Length(FData^.szTip) - 1);
{$IFEND}
    FData.uID := FData.Wnd;
    FData.uTimeout := 3000;
    FData.hIcon := FCurrentIcon.Handle;
    FData.uFlags := NIF_ICON or NIF_MESSAGE;
    FData.uCallbackMessage := WM_SYSTEM_TRAY_MESSAGE;
    if Length(Application.Title) > 0 then
       FData.uFlags := FData.uFlags or NIF_TIP;
    Refresh;
  end;
end;

destructor TCustomTrayIcon.Destroy;
begin
  if not (csDesigning in ComponentState) then
  begin
    Refresh(NIM_DELETE);
    DeallocateHWnd(FData.Wnd);
  end;
  FCurrentIcon.Free;
  FIcon.Free;
  FTimer.Free;
{$IF NOT DEFINED(CLR)}
  Dispose(FData);
{$IFEND}
  inherited;
end;

procedure TCustomTrayIcon.SetVisible(Value: Boolean);
begin
  if FVisible <> Value then
  begin
    FVisible := Value;
    if (not FAnimate) or (FAnimate and FCurrentIcon.Empty) then
      SetDefaultIcon;

    if not (csDesigning in ComponentState) then
    begin
      if FVisible then
      begin
        if not Refresh(NIM_ADD) then
          raise EOutOfResources.Create(STrayIconCreateError);
      end
      else if not (csLoading in ComponentState) then
      begin
        if not Refresh(NIM_DELETE) then
          raise EOutOfResources.Create(STrayIconRemoveError);
      end;
      if FAnimate then
        FTimer.Enabled := Value;
    end;
  end;
end;

procedure TCustomTrayIcon.SetIconList(Value: TImageList);
begin
  if FIconList <> Value then
  begin
    FIconList := Value;
    if not (csDesigning in ComponentState) then
    begin
      if Assigned(FIconList) then
        FIconList.GetIcon(FIconIndex, FCurrentIcon)
      else
        SetDefaultIcon;
      Refresh;
    end;
  end;
end;

procedure TCustomTrayIcon.SetHint(const Value: string);
begin
  if CompareStr(FHint, Value) <> 0 then
  begin
    FHint := Value;
{$IF DEFINED(CLR)}
    FData.szTip := Hint;
{$ELSE}
    StrPLCopy(FData.szTip, FHint, Length(FData.szTip) - 1);
{$IFEND}
    if Length(Hint) > 0 then
      FData.uFlags := FData.uFlags or NIF_TIP
    else
      FData.uFlags := FData.uFlags and not NIF_TIP;
    Refresh;
  end;
end;

function TCustomTrayIcon.GetAnimateInterval: Cardinal;
begin
  Result := FTimer.Interval;
end;

procedure TCustomTrayIcon.SetAnimateInterval(Value: Cardinal);
begin
  FTimer.Interval := Value;
end;

procedure TCustomTrayIcon.SetAnimate(Value: Boolean);
begin
  if FAnimate <> Value then
  begin
    FAnimate := Value;
    if not (csDesigning in ComponentState) then
    begin
      if (FIconList <> nil) and (FIconList.Count > 0) and Visible then
        FTimer.Enabled := Value;
      if (not FAnimate) and (not FCurrentIcon.Empty) then
        FIcon.Assign(FCurrentIcon);
    end;
  end;
end;

{ Message handler for the hidden shell notification window. Most messages
  use WM_SYSTEM_TRAY_MESSAGE as the Message ID, with WParam as the ID of the
  shell notify icon data. LParam is a message ID for the actual message, e.g.,
  WM_MOUSEMOVE. Another important message is WM_ENDSESSION, telling the shell
  notify icon to delete itself, so Windows can shut down.

  Send the usual events for the mouse messages. Also interpolate the OnClick
  event when the user clicks the left button, and popup the menu, if there is
  one, for right click events. }

//[SecurityPermission(SecurityAction.InheritanceDemand, UnmanagedCode=True)]
procedure TCustomTrayIcon.WindowProc(var Message: TMessage);

  { Return the state of the shift keys. }
  function ShiftState: TShiftState;
  begin
    Result := [];
    if GetKeyState(VK_SHIFT) < 0 then
      Include(Result, ssShift);
    if GetKeyState(VK_CONTROL) < 0 then
      Include(Result, ssCtrl);
    if GetKeyState(VK_MENU) < 0 then
      Include(Result, ssAlt);
  end;

var
  Point: TPoint;
  Shift: TShiftState;
begin
  case Message.Msg of
    WM_QUERYENDSESSION: Message.Result := 1;
    WM_ENDSESSION:
      if TWmEndSession(Message).EndSession then
        Refresh(NIM_DELETE);
    WM_SYSTEM_TRAY_MESSAGE:
      begin
        case Int64(Message.lParam) of
          WM_MOUSEMOVE:
            if Assigned(FOnMouseMove) then
            begin
              Shift := ShiftState;
              GetCursorPos(Point);
              FOnMouseMove(Self, Shift, Point.X, Point.Y);
            end;
          WM_LBUTTONDOWN:
            begin
              if Assigned(FOnMouseDown) then
              begin
                Shift := ShiftState + [ssLeft];
                GetCursorPos(Point);
                FOnMouseDown(Self, mbLeft, Shift, Point.X, Point.Y);
              end;
              FIsClicked := True;
            end;
          WM_LBUTTONUP:
            begin
              Shift := ShiftState + [ssLeft];
              GetCursorPos(Point);
              if FIsClicked and Assigned(FOnClick) then
              begin
                FOnClick(Self);
                FIsClicked := False;
              end;
              if Assigned(FOnMouseUp) then
                FOnMouseUp(Self, mbLeft, Shift, Point.X, Point.Y);
            end;
          WM_RBUTTONDOWN:
            if Assigned(FOnMouseDown) then
            begin
              Shift := ShiftState + [ssRight];
              GetCursorPos(Point);
              FOnMouseDown(Self, mbRight, Shift, Point.X, Point.Y);
            end;
          WM_RBUTTONUP:
            begin
              Shift := ShiftState + [ssRight];
              GetCursorPos(Point);
              if Assigned(FOnMouseUp) then
                FOnMouseUp(Self, mbRight, Shift, Point.X, Point.Y);
              if Assigned(FPopupMenu) then
              begin
                SetForegroundWindow(Application.Handle);
                Application.ProcessMessages;
                FPopupMenu.AutoPopup := False;
                FPopupMenu.PopupComponent := Owner;
                FPopupMenu.Popup(Point.x, Point.y);
              end;
            end;
          WM_LBUTTONDBLCLK, WM_MBUTTONDBLCLK, WM_RBUTTONDBLCLK:
            if Assigned(FOnDblClick) then
              FOnDblClick(Self);
          WM_MBUTTONDOWN:
            if Assigned(FOnMouseDown) then
            begin
              Shift := ShiftState + [ssMiddle];
              GetCursorPos(Point);
              FOnMouseDown(Self, mbMiddle, Shift, Point.X, Point.Y);
            end;
          WM_MBUTTONUP:
            if Assigned(FOnMouseUp) then
            begin
              Shift := ShiftState + [ssMiddle];
              GetCursorPos(Point);
              FOnMouseUp(Self, mbMiddle, Shift, Point.X, Point.Y);
            end;
          NIN_BALLOONHIDE, NIN_BALLOONTIMEOUT:
            FData.uFlags := FData.uFlags and not NIF_INFO;
          NIN_BALLOONUSERCLICK:
            if Assigned(FOnBalloonClick) then
              FOnBalloonClick(Self);
        end;
      end;
  else
    if (Cardinal(Message.Msg) = RM_TaskBarCreated) and Visible then
      Refresh(NIM_ADD);
  end;
end;

procedure TCustomTrayIcon.Refresh;
begin
  if not (csDesigning in ComponentState) then
  begin
    FData.hIcon := FCurrentIcon.Handle;
    if Visible then
      Refresh(NIM_MODIFY);
  end;
end;

function TCustomTrayIcon.Refresh(Message: Integer): Boolean;
begin
  Result := Shell_NotifyIcon_2009(Message, FData);
end;

procedure TCustomTrayIcon.SetIconIndex(Value: Integer);
begin
  if FIconIndex <> Value then
  begin
    FIconIndex := Value;
    if not (csDesigning in ComponentState) then
    begin
      if Assigned(FIconList) then
        FIconList.GetIcon(FIconIndex, FCurrentIcon);
      Refresh;
    end;
  end;
end;

procedure TCustomTrayIcon.DoOnAnimate(Sender: TObject);
begin
  if Assigned(FOnAnimate) then
    FOnAnimate(Self);
  if Assigned(FIconList) and (FIconIndex < FIconList.Count - 1) then
    IconIndex := FIconIndex + 1
  else
    IconIndex := 0;
  Refresh;
end;

procedure TCustomTrayIcon.SetIcon(Value: TIcon);
begin
  FIcon.Assign(Value);
  FCurrentIcon.Assign(Value);
  Refresh;
end;

procedure TCustomTrayIcon.SetBalloonHint(const Value: string);
begin
  if CompareStr(FBalloonHint, Value) <> 0 then
  begin
    FBalloonHint := Value;
{$IF DEFINED(CLR)}
    FData.szInfo := FBalloonHint;
{$ELSE}
    StrPLCopy(FData.szInfo, FBalloonHint, Length(FData.szInfo) - 1);
{$IFEND}
    Refresh(NIM_MODIFY);
  end;
end;

procedure TCustomTrayIcon.SetDefaultIcon;
begin
  if not FIcon.Empty then
    FCurrentIcon.Assign(FIcon)
  else
    FCurrentIcon.Assign(Application.Icon);
  Refresh;
end;

procedure TCustomTrayIcon.SetBalloonTimeout(Value: Integer);
begin
  FData.uTimeout := Value;
end;

function TCustomTrayIcon.GetBalloonTimeout: Integer;
begin
  Result := FData.uTimeout;
end;

function TCustomTrayIcon.GetData: TNotifyIconData_2009;
begin
  Result := FData{$IFNDEF CLR}^{$ENDIF};
end;

procedure TCustomTrayIcon.Notification(AComponent: TComponent; Operation: TOperation);
begin
  inherited Notification(AComponent, Operation);
  if (AComponent = FPopupMenu) and (Operation = opRemove) then
    FPopupMenu := nil;
  if (AComponent = FIconList) and (Operation = opRemove) then
    FIconList := nil;
end;

procedure TCustomTrayIcon.ShowBalloonHint;
begin
  FData.uFlags := FData.uFlags or NIF_INFO;
  FData.dwInfoFlags := Integer(FBalloonFlags);
  Refresh(NIM_MODIFY);
end;

procedure TCustomTrayIcon.SetBalloonTitle(const Value: string);
begin
  if CompareStr(FBalloonTitle, Value) <> 0 then
  begin
    FBalloonTitle := Value;
{$IF DEFINED(CLR)}
    FData.szInfoTitle := FBalloonTitle;
{$ELSE}
    StrPLCopy(FData.szInfoTitle, FBalloonTitle, Length(FData.szInfoTitle) - 1);
{$IFEND}
    Refresh(NIM_MODIFY);
  end;
end;

initialization
  RM_TaskBarCreated := RegisterWindowMessage('TaskbarCreated');
end.
 
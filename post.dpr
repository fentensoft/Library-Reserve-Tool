program post;

uses
  Vcl.Forms,
  Main in 'Main.pas' {frmMain},
  Sched in 'Sched.pas',
  Http in 'Http.pas',
  Vcl.Themes,
  Vcl.Styles,
  SuperObject in 'SuperObject.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TfrmMain, frmMain);
  Application.Run;

end.

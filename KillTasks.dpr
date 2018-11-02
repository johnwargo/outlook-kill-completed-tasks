program KillTasks;

uses
  Vcl.Forms,
  kill_task in 'kill_task.pas' {frmMain};

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.Title := 'Outlook Task Killer';
  Application.CreateForm(TfrmMain, frmMain);
  Application.Run;
end.

program SharpToDelphiCB;

uses
  Vcl.Forms,
  UMain in 'UMain.pas' {Form1},
  SharpToDelphiCallback_TLB in 'SharpToDelphiCallback_TLB.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TForm1, Form1);
  Application.Run;
end.

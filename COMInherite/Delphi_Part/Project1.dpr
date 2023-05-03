program Project1;

uses
  Vcl.Forms,
  UMain in 'UMain.pas' {Form1},
  ComInherite_TLB in 'ComInherite_TLB.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TForm1, Form1);
  Application.Run;
end.

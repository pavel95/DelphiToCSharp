unit UMain;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, ComInherite_TLB, ActiveX;

type
  TForm1 = class(TForm)
    Button1: TButton;
    Label1: TLabel;
    procedure Button1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

  // Implemetace rozhrani z C#
  TDelphiEventsContainer = class(TInterfacedObject, IDelphiEventContainer)
  public
    // metoda k prepsani
    //par value je argument ze C#, pRetVal vracim
    function GetValue(parValue: PSafeArray; out pRetVal: Integer): HResult; stdcall;
  end;

var
  Form1: TForm1;

implementation

{$R *.dfm}

procedure TForm1.Button1Click(Sender: TObject);
var
  Calc : TCalc;
  DelphiEventsContainer : TDelphiEventsContainer;
  I : integer;
begin
    DelphiEventsContainer := TDelphiEventsContainer.Create();
    Calc :=TCalc.Create(Self);
    Calc.DelphiEvents := DelphiEventsContainer;
    I := Calc.Sum(1,2);
end;

{ TDelphiEventsContainer }

function TDelphiEventsContainer.GetValue(parValue: PSafeArray;
  out pRetVal: Integer): HResult; stdcall;
var
  vr:Variant;
  safeAr:PSafeArray;
  a, b, i, N, J : integer;
begin
  vr:=VarArrayCreate ([0, 1], varInteger);
  i := 0;

  J := parValue.rgsabound[0].lLbound;
  N := parValue.rgsabound[0].cElements;

  SafeArrayGetElement(parValue, i, a);  //Pozor bije se s varUtils!!
  I:=1;
  SafeArrayGetElement(parValue, i, b);
  pRetVal := (a + b) * 100;    //par value je argument ze C#, pRetVal vracim
  result := S_OK;  //S_OK aby to proslo

  Form1.Label1.Caption := pRetVal.ToString;
end;

end.

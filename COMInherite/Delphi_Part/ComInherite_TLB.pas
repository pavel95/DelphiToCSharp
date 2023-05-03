unit ComInherite_TLB;

// ************************************************************************ //
// WARNING                                                                    
// -------                                                                    
// The types declared in this file were generated from data read from a       
// Type Library. If this type library is explicitly or indirectly (via        
// another type library referring to this type library) re-imported, or the   
// 'Refresh' command of the Type Library Editor activated while editing the   
// Type Library, the contents of this file will be regenerated and all        
// manual modifications will be lost.                                         
// ************************************************************************ //

// $Rev: 52393 $
// File generated on 28.04.2023 13:41:53 from Type Library described below.

// ************************************************************************  //
// Type Lib: C:\Develop\DelphiToCSharp\COMInherite\CSharp_Part\ComInherite\ComInherite\bin\Debug\ComInherite.tlb (1)
// LIBID: {E8596322-F64B-4C3F-ABA5-4424E445C0FA}
// LCID: 0
// Helpfile: 
// HelpString: 
// DepndLst: 
//   (1) v2.0 stdole, (C:\Windows\SysWOW64\stdole2.tlb)
//   (2) v2.4 mscorlib, (C:\Windows\Microsoft.NET\Framework\v4.0.30319\mscorlib.tlb)
// SYS_KIND: SYS_WIN32
// Errors:
//   Error creating palette bitmap of (TCalc) : Server mscoree.dll contains no icons
//   Error creating palette bitmap of (TDelphiEventContainer) : Server mscoree.dll contains no icons
// ************************************************************************ //
{$TYPEDADDRESS OFF} // Unit must be compiled without type-checked pointers. 
{$WARN SYMBOL_PLATFORM OFF}
{$WRITEABLECONST ON}
{$VARPROPSETTER ON}
{$ALIGN 4}

interface

uses Winapi.Windows, mscorlib_TLB, System.Classes, System.Variants, System.Win.StdVCL, Vcl.Graphics, Vcl.OleServer, Winapi.ActiveX;
  


// *********************************************************************//
// GUIDS declared in the TypeLibrary. Following prefixes are used:        
//   Type Libraries     : LIBID_xxxx                                      
//   CoClasses          : CLASS_xxxx                                      
//   DISPInterfaces     : DIID_xxxx                                       
//   Non-DISP interfaces: IID_xxxx                                        
// *********************************************************************//
const
  // TypeLibrary Major and minor versions
  ComInheriteMajorVersion = 1;
  ComInheriteMinorVersion = 0;

  LIBID_ComInherite: TGUID = '{E8596322-F64B-4C3F-ABA5-4424E445C0FA}';

  IID_ICalc: TGUID = '{898B0A63-BC76-43C6-BED0-21F485367701}';
  CLASS_Calc: TGUID = '{F38396C3-345A-4209-9675-1C0CD95E3134}';
  IID_IDelphiEventContainer: TGUID = '{FB3C8BAA-27FF-4E8C-89D6-133F86EC47F1}';
  CLASS_DelphiEventContainer: TGUID = '{071CD8E4-4FDD-4919-8817-AD5133B7682A}';
type

// *********************************************************************//
// Forward declaration of types defined in TypeLibrary                    
// *********************************************************************//
  ICalc = interface;
  ICalcDisp = dispinterface;
  IDelphiEventContainer = interface;

// *********************************************************************//
// Declaration of CoClasses defined in Type Library                       
// (NOTE: Here we map each CoClass to its Default Interface)              
// *********************************************************************//
  Calc = ICalc;
  DelphiEventContainer = IDelphiEventContainer;


// *********************************************************************//
// Interface: ICalc
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {898B0A63-BC76-43C6-BED0-21F485367701}
// *********************************************************************//
  ICalc = interface(IDispatch)
    ['{898B0A63-BC76-43C6-BED0-21F485367701}']
    function Sum(a: Integer; b: Integer): Integer; safecall;
    function Get_DelphiEvents: IDelphiEventContainer; safecall;
    procedure _Set_DelphiEvents(const pRetVal: IDelphiEventContainer); safecall;
    property DelphiEvents: IDelphiEventContainer read Get_DelphiEvents write _Set_DelphiEvents;
  end;

// *********************************************************************//
// DispIntf:  ICalcDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {898B0A63-BC76-43C6-BED0-21F485367701}
// *********************************************************************//
  ICalcDisp = dispinterface
    ['{898B0A63-BC76-43C6-BED0-21F485367701}']
    function Sum(a: Integer; b: Integer): Integer; dispid 1610743808;
    property DelphiEvents: IDelphiEventContainer dispid 1610743809;
  end;

// *********************************************************************//
// Interface: IDelphiEventContainer
// Flags:     (256) OleAutomation
// GUID:      {FB3C8BAA-27FF-4E8C-89D6-133F86EC47F1}
// *********************************************************************//
  IDelphiEventContainer = interface(IUnknown)
    ['{FB3C8BAA-27FF-4E8C-89D6-133F86EC47F1}']
    function GetValue(parValue: PSafeArray; out pRetVal: Integer): HResult; stdcall;
  end;

// *********************************************************************//
// The Class CoCalc provides a Create and CreateRemote method to          
// create instances of the default interface ICalc exposed by              
// the CoClass Calc. The functions are intended to be used by             
// clients wishing to automate the CoClass objects exposed by the         
// server of this typelibrary.                                            
// *********************************************************************//
  CoCalc = class
    class function Create: ICalc;
    class function CreateRemote(const MachineName: string): ICalc;
  end;


// *********************************************************************//
// OLE Server Proxy class declaration
// Server Object    : TCalc
// Help String      : 
// Default Interface: ICalc
// Def. Intf. DISP? : No
// Event   Interface: 
// TypeFlags        : (2) CanCreate
// *********************************************************************//
  TCalc = class(TOleServer)
  private
    FIntf: ICalc;
    function GetDefaultInterface: ICalc;
  protected
    procedure InitServerData; override;
    function Get_DelphiEvents: IDelphiEventContainer;
    procedure _Set_DelphiEvents(const pRetVal: IDelphiEventContainer);
  public
    constructor Create(AOwner: TComponent); override;
    destructor  Destroy; override;
    procedure Connect; override;
    procedure ConnectTo(svrIntf: ICalc);
    procedure Disconnect; override;
    function Sum(a: Integer; b: Integer): Integer;
    property DefaultInterface: ICalc read GetDefaultInterface;
    property DelphiEvents: IDelphiEventContainer read Get_DelphiEvents write _Set_DelphiEvents;
  published
  end;

// *********************************************************************//
// The Class CoDelphiEventContainer provides a Create and CreateRemote method to          
// create instances of the default interface IDelphiEventContainer exposed by              
// the CoClass DelphiEventContainer. The functions are intended to be used by             
// clients wishing to automate the CoClass objects exposed by the         
// server of this typelibrary.                                            
// *********************************************************************//
  CoDelphiEventContainer = class
    class function Create: IDelphiEventContainer;
    class function CreateRemote(const MachineName: string): IDelphiEventContainer;
  end;


// *********************************************************************//
// OLE Server Proxy class declaration
// Server Object    : TDelphiEventContainer
// Help String      : 
// Default Interface: IDelphiEventContainer
// Def. Intf. DISP? : No
// Event   Interface: 
// TypeFlags        : (2) CanCreate
// *********************************************************************//
  TDelphiEventContainer = class(TOleServer)
  private
    FIntf: IDelphiEventContainer;
    function GetDefaultInterface: IDelphiEventContainer;
  protected
    procedure InitServerData; override;
  public
    constructor Create(AOwner: TComponent); override;
    destructor  Destroy; override;
    procedure Connect; override;
    procedure ConnectTo(svrIntf: IDelphiEventContainer);
    procedure Disconnect; override;
    function GetValue(parValue: PSafeArray; out pRetVal: Integer): HResult;
    property DefaultInterface: IDelphiEventContainer read GetDefaultInterface;
  published
  end;

procedure Register;

resourcestring
  dtlServerPage = 'ActiveX';

  dtlOcxPage = 'ActiveX';

implementation

uses System.Win.ComObj;

class function CoCalc.Create: ICalc;
begin
  Result := CreateComObject(CLASS_Calc) as ICalc;
end;

class function CoCalc.CreateRemote(const MachineName: string): ICalc;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_Calc) as ICalc;
end;

procedure TCalc.InitServerData;
const
  CServerData: TServerData = (
    ClassID:   '{F38396C3-345A-4209-9675-1C0CD95E3134}';
    IntfIID:   '{898B0A63-BC76-43C6-BED0-21F485367701}';
    EventIID:  '';
    LicenseKey: nil;
    Version: 500);
begin
  ServerData := @CServerData;
end;

procedure TCalc.Connect;
var
  punk: IUnknown;
begin
  if FIntf = nil then
  begin
    punk := GetServer;
    Fintf:= punk as ICalc;
  end;
end;

procedure TCalc.ConnectTo(svrIntf: ICalc);
begin
  Disconnect;
  FIntf := svrIntf;
end;

procedure TCalc.DisConnect;
begin
  if Fintf <> nil then
  begin
    FIntf := nil;
  end;
end;

function TCalc.GetDefaultInterface: ICalc;
begin
  if FIntf = nil then
    Connect;
  Assert(FIntf <> nil, 'DefaultInterface is NULL. Component is not connected to Server. You must call "Connect" or "ConnectTo" before this operation');
  Result := FIntf;
end;

constructor TCalc.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
end;

destructor TCalc.Destroy;
begin
  inherited Destroy;
end;

function TCalc.Get_DelphiEvents: IDelphiEventContainer;
begin
  Result := DefaultInterface.DelphiEvents;
end;

procedure TCalc._Set_DelphiEvents(const pRetVal: IDelphiEventContainer);
begin
  DefaultInterface.DelphiEvents := pRetVal;
end;

function TCalc.Sum(a: Integer; b: Integer): Integer;
begin
  Result := DefaultInterface.Sum(a, b);
end;

class function CoDelphiEventContainer.Create: IDelphiEventContainer;
begin
  Result := CreateComObject(CLASS_DelphiEventContainer) as IDelphiEventContainer;
end;

class function CoDelphiEventContainer.CreateRemote(const MachineName: string): IDelphiEventContainer;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_DelphiEventContainer) as IDelphiEventContainer;
end;

procedure TDelphiEventContainer.InitServerData;
const
  CServerData: TServerData = (
    ClassID:   '{071CD8E4-4FDD-4919-8817-AD5133B7682A}';
    IntfIID:   '{FB3C8BAA-27FF-4E8C-89D6-133F86EC47F1}';
    EventIID:  '';
    LicenseKey: nil;
    Version: 500);
begin
  ServerData := @CServerData;
end;

procedure TDelphiEventContainer.Connect;
var
  punk: IUnknown;
begin
  if FIntf = nil then
  begin
    punk := GetServer;
    Fintf:= punk as IDelphiEventContainer;
  end;
end;

procedure TDelphiEventContainer.ConnectTo(svrIntf: IDelphiEventContainer);
begin
  Disconnect;
  FIntf := svrIntf;
end;

procedure TDelphiEventContainer.DisConnect;
begin
  if Fintf <> nil then
  begin
    FIntf := nil;
  end;
end;

function TDelphiEventContainer.GetDefaultInterface: IDelphiEventContainer;
begin
  if FIntf = nil then
    Connect;
  Assert(FIntf <> nil, 'DefaultInterface is NULL. Component is not connected to Server. You must call "Connect" or "ConnectTo" before this operation');
  Result := FIntf;
end;

constructor TDelphiEventContainer.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
end;

destructor TDelphiEventContainer.Destroy;
begin
  inherited Destroy;
end;

function TDelphiEventContainer.GetValue(parValue: PSafeArray; out pRetVal: Integer): HResult;
begin
  Result := DefaultInterface.GetValue(parValue, pRetVal);
end;

procedure Register;
begin
  RegisterComponents(dtlServerPage, [TCalc, TDelphiEventContainer]);
end;

end.

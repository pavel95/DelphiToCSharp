unit SharpToDelphi_TLB;

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
// File generated on 25.04.2023 10:34:01 from Type Library described below.

// ************************************************************************  //
// Type Lib: C:\Develop\Csharp\SharpToDelphi\SharpToDelphi\bin\Debug\SharpToDelphi.tlb (1)
// LIBID: {67D26B0F-81AD-4CD8-8EFE-83AABF62F9B0}
// LCID: 0
// Helpfile: 
// HelpString: 
// DepndLst: 
//   (1) v2.0 stdole, (C:\Windows\SysWOW64\stdole2.tlb)
//   (2) v2.0 mscorlib, (C:\WINDOWS\Microsoft.NET\Framework\v2.0.50727\mscorlib.tlb)
// SYS_KIND: SYS_WIN64
// Errors:
//   Error creating palette bitmap of (TCalculator) : Server mscoree.dll contains no icons
//   Error creating palette bitmap of (TDelphiEvents) : Server mscoree.dll contains no icons
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
  SharpToDelphiMajorVersion = 1;
  SharpToDelphiMinorVersion = 0;

  LIBID_SharpToDelphi: TGUID = '{67D26B0F-81AD-4CD8-8EFE-83AABF62F9B0}';

  IID_ICalculator: TGUID = '{1C3AC5AB-A9AA-4B1A-AFCB-37B218977CBD}';
  CLASS_Calculator: TGUID = '{2B4D2884-1ACA-4AFB-98B5-F221183FAFC4}';
  IID_IDelphiEvents: TGUID = '{B0EED8D4-973C-4F24-A4E9-3F2F3D572C9A}';
  CLASS_DelphiEvents: TGUID = '{3B3861FE-6AD9-4E43-9C00-143033F8CACE}';
type

// *********************************************************************//
// Forward declaration of types defined in TypeLibrary                    
// *********************************************************************//
  ICalculator = interface;
  ICalculatorDisp = dispinterface;
  IDelphiEvents = interface;

// *********************************************************************//
// Declaration of CoClasses defined in Type Library                       
// (NOTE: Here we map each CoClass to its Default Interface)              
// *********************************************************************//
  Calculator = ICalculator;
  DelphiEvents = IDelphiEvents;


// *********************************************************************//
// Interface: ICalculator
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {1C3AC5AB-A9AA-4B1A-AFCB-37B218977CBD}
// *********************************************************************//
  ICalculator = interface(IDispatch)
    ['{1C3AC5AB-A9AA-4B1A-AFCB-37B218977CBD}']
    function Sum(a: Integer; b: Integer): Integer; safecall;
    procedure OnCallback(intPtr: Integer); safecall;
    function Get_OnDelphi: IDelphiEvents; safecall;
    procedure _Set_OnDelphi(const pRetVal: IDelphiEvents); safecall;
    property OnDelphi: IDelphiEvents read Get_OnDelphi write _Set_OnDelphi;
  end;

// *********************************************************************//
// DispIntf:  ICalculatorDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {1C3AC5AB-A9AA-4B1A-AFCB-37B218977CBD}
// *********************************************************************//
  ICalculatorDisp = dispinterface
    ['{1C3AC5AB-A9AA-4B1A-AFCB-37B218977CBD}']
    function Sum(a: Integer; b: Integer): Integer; dispid 1610743808;
    procedure OnCallback(intPtr: Integer); dispid 1610743809;
    property OnDelphi: IDelphiEvents dispid 1610743810;
  end;

// *********************************************************************//
// Interface: IDelphiEvents
// Flags:     (256) OleAutomation
// GUID:      {B0EED8D4-973C-4F24-A4E9-3F2F3D572C9A}
// *********************************************************************//
  IDelphiEvents = interface(IUnknown)
    ['{B0EED8D4-973C-4F24-A4E9-3F2F3D572C9A}']
    function GetValue(parValue: PSafeArray; out pRetVal: Integer): HResult; stdcall;
  end;

// *********************************************************************//
// The Class CoCalculator provides a Create and CreateRemote method to          
// create instances of the default interface ICalculator exposed by              
// the CoClass Calculator. The functions are intended to be used by             
// clients wishing to automate the CoClass objects exposed by the         
// server of this typelibrary.                                            
// *********************************************************************//
  CoCalculator = class
    class function Create: ICalculator;
    class function CreateRemote(const MachineName: string): ICalculator;
  end;


// *********************************************************************//
// OLE Server Proxy class declaration
// Server Object    : TCalculator
// Help String      : 
// Default Interface: ICalculator
// Def. Intf. DISP? : No
// Event   Interface: 
// TypeFlags        : (2) CanCreate
// *********************************************************************//
  TCalculator = class(TOleServer)
  private
    FIntf: ICalculator;
    function GetDefaultInterface: ICalculator;
  protected
    procedure InitServerData; override;
    function Get_OnDelphi: IDelphiEvents;
    procedure _Set_OnDelphi(const pRetVal: IDelphiEvents);
  public
    constructor Create(AOwner: TComponent); override;
    destructor  Destroy; override;
    procedure Connect; override;
    procedure ConnectTo(svrIntf: ICalculator);
    procedure Disconnect; override;
    function Sum(a: Integer; b: Integer): Integer;
    procedure OnCallback(intPtr: Integer);
    property DefaultInterface: ICalculator read GetDefaultInterface;
    property OnDelphi: IDelphiEvents read Get_OnDelphi write _Set_OnDelphi;
  published
  end;

// *********************************************************************//
// The Class CoDelphiEvents provides a Create and CreateRemote method to          
// create instances of the default interface IDelphiEvents exposed by              
// the CoClass DelphiEvents. The functions are intended to be used by             
// clients wishing to automate the CoClass objects exposed by the         
// server of this typelibrary.                                            
// *********************************************************************//
  CoDelphiEvents = class
    class function Create: IDelphiEvents;
    class function CreateRemote(const MachineName: string): IDelphiEvents;
  end;


// *********************************************************************//
// OLE Server Proxy class declaration
// Server Object    : TDelphiEvents
// Help String      : 
// Default Interface: IDelphiEvents
// Def. Intf. DISP? : No
// Event   Interface: 
// TypeFlags        : (2) CanCreate
// *********************************************************************//
  TDelphiEvents = class(TOleServer)
  private
    FIntf: IDelphiEvents;
    function GetDefaultInterface: IDelphiEvents;
  protected
    procedure InitServerData; override;
  public
    constructor Create(AOwner: TComponent); override;
    destructor  Destroy; override;
    procedure Connect; override;
    procedure ConnectTo(svrIntf: IDelphiEvents);
    procedure Disconnect; override;
    function GetValue(parValue: PSafeArray; out pRetVal: Integer): HResult;
    property DefaultInterface: IDelphiEvents read GetDefaultInterface;
  published
  end;

procedure Register;

resourcestring
  dtlServerPage = 'ActiveX';

  dtlOcxPage = 'ActiveX';

implementation

uses System.Win.ComObj;

class function CoCalculator.Create: ICalculator;
begin
  Result := CreateComObject(CLASS_Calculator) as ICalculator;
end;

class function CoCalculator.CreateRemote(const MachineName: string): ICalculator;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_Calculator) as ICalculator;
end;

procedure TCalculator.InitServerData;
const
  CServerData: TServerData = (
    ClassID:   '{2B4D2884-1ACA-4AFB-98B5-F221183FAFC4}';
    IntfIID:   '{1C3AC5AB-A9AA-4B1A-AFCB-37B218977CBD}';
    EventIID:  '';
    LicenseKey: nil;
    Version: 500);
begin
  ServerData := @CServerData;
end;

procedure TCalculator.Connect;
var
  punk: IUnknown;
begin
  if FIntf = nil then
  begin
    punk := GetServer;
    Fintf:= punk as ICalculator;
  end;
end;

procedure TCalculator.ConnectTo(svrIntf: ICalculator);
begin
  Disconnect;
  FIntf := svrIntf;
end;

procedure TCalculator.DisConnect;
begin
  if Fintf <> nil then
  begin
    FIntf := nil;
  end;
end;

function TCalculator.GetDefaultInterface: ICalculator;
begin
  if FIntf = nil then
    Connect;
  Assert(FIntf <> nil, 'DefaultInterface is NULL. Component is not connected to Server. You must call "Connect" or "ConnectTo" before this operation');
  Result := FIntf;
end;

constructor TCalculator.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
end;

destructor TCalculator.Destroy;
begin
  inherited Destroy;
end;

function TCalculator.Get_OnDelphi: IDelphiEvents;
begin
  Result := DefaultInterface.OnDelphi;
end;

procedure TCalculator._Set_OnDelphi(const pRetVal: IDelphiEvents);
begin
  DefaultInterface.OnDelphi := pRetVal;
end;

function TCalculator.Sum(a: Integer; b: Integer): Integer;
begin
  Result := DefaultInterface.Sum(a, b);
end;

procedure TCalculator.OnCallback(intPtr: Integer);
begin
  DefaultInterface.OnCallback(intPtr);
end;

class function CoDelphiEvents.Create: IDelphiEvents;
begin
  Result := CreateComObject(CLASS_DelphiEvents) as IDelphiEvents;
end;

class function CoDelphiEvents.CreateRemote(const MachineName: string): IDelphiEvents;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_DelphiEvents) as IDelphiEvents;
end;

procedure TDelphiEvents.InitServerData;
const
  CServerData: TServerData = (
    ClassID:   '{3B3861FE-6AD9-4E43-9C00-143033F8CACE}';
    IntfIID:   '{B0EED8D4-973C-4F24-A4E9-3F2F3D572C9A}';
    EventIID:  '';
    LicenseKey: nil;
    Version: 500);
begin
  ServerData := @CServerData;
end;

procedure TDelphiEvents.Connect;
var
  punk: IUnknown;
begin
  if FIntf = nil then
  begin
    punk := GetServer;
    Fintf:= punk as IDelphiEvents;
  end;
end;

procedure TDelphiEvents.ConnectTo(svrIntf: IDelphiEvents);
begin
  Disconnect;
  FIntf := svrIntf;
end;

procedure TDelphiEvents.DisConnect;
begin
  if Fintf <> nil then
  begin
    FIntf := nil;
  end;
end;

function TDelphiEvents.GetDefaultInterface: IDelphiEvents;
begin
  if FIntf = nil then
    Connect;
  Assert(FIntf <> nil, 'DefaultInterface is NULL. Component is not connected to Server. You must call "Connect" or "ConnectTo" before this operation');
  Result := FIntf;
end;

constructor TDelphiEvents.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
end;

destructor TDelphiEvents.Destroy;
begin
  inherited Destroy;
end;

function TDelphiEvents.GetValue(parValue: PSafeArray; out pRetVal: Integer): HResult;
begin
  Result := DefaultInterface.GetValue(parValue, pRetVal);
end;

procedure Register;
begin
  RegisterComponents(dtlServerPage, [TCalculator, TDelphiEvents]);
end;

end.

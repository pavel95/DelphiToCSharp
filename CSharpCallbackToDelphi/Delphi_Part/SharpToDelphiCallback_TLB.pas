unit SharpToDelphiCallback_TLB;

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
// File generated on 28.04.2023 9:12:35 from Type Library described below.

// ************************************************************************  //
// Type Lib: C:\Develop\Delphi_Other\CSharpCallbackToDelphi\CSharp_Part\SharpToDelphiCallback\SharpToDelphiCallback\bin\Debug\SharpToDelphiCallback.tlb (1)
// LIBID: {486E72C5-E0EA-4681-A015-94088909DA5A}
// LCID: 0
// Helpfile: 
// HelpString: 
// DepndLst: 
//   (1) v2.0 stdole, (C:\Windows\SysWOW64\stdole2.tlb)
//   (2) v2.4 mscorlib, (C:\Windows\Microsoft.NET\Framework\v4.0.30319\mscorlib.tlb)
// SYS_KIND: SYS_WIN64
// Errors:
//   Error creating palette bitmap of (TSomething) : Server mscoree.dll contains no icons
//   Error creating palette bitmap of (TSomethingElse) : Server mscoree.dll contains no icons
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
  SharpToDelphiCallbackMajorVersion = 1;
  SharpToDelphiCallbackMinorVersion = 0;

  LIBID_SharpToDelphiCallback: TGUID = '{486E72C5-E0EA-4681-A015-94088909DA5A}';

  IID_ISomething: TGUID = '{79801B97-62BC-489B-935A-8543116E620E}';
  CLASS_Something: TGUID = '{B2BB5506-8308-400C-924F-4A3E3D46D965}';
  IID_ISomethingElse: TGUID = '{00A7D870-AE65-4682-BD1B-D2EE21B0B88C}';
  CLASS_SomethingElse: TGUID = '{165B7823-3B02-4DAA-8205-C07FB7FA124A}';
type

// *********************************************************************//
// Forward declaration of types defined in TypeLibrary                    
// *********************************************************************//
  ISomething = interface;
  ISomethingDisp = dispinterface;
  ISomethingElse = interface;
  ISomethingElseDisp = dispinterface;

// *********************************************************************//
// Declaration of CoClasses defined in Type Library                       
// (NOTE: Here we map each CoClass to its Default Interface)              
// *********************************************************************//
  Something = ISomething;
  SomethingElse = ISomethingElse;


// *********************************************************************//
// Interface: ISomething
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {79801B97-62BC-489B-935A-8543116E620E}
// *********************************************************************//
  ISomething = interface(IDispatch)
    ['{79801B97-62BC-489B-935A-8543116E620E}']
    procedure SetIntArrayCallback(integerPointer: Integer); safecall;
    procedure SetSomethingElseCallback(integerPointer: Integer); safecall;
    procedure Reverse(ids: PSafeArray); safecall;
    procedure CallSomethingElse(const SomethingElse: ISomethingElse); safecall;
  end;

// *********************************************************************//
// DispIntf:  ISomethingDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {79801B97-62BC-489B-935A-8543116E620E}
// *********************************************************************//
  ISomethingDisp = dispinterface
    ['{79801B97-62BC-489B-935A-8543116E620E}']
    procedure SetIntArrayCallback(integerPointer: Integer); dispid 1610743808;
    procedure SetSomethingElseCallback(integerPointer: Integer); dispid 1610743809;
    procedure Reverse(ids: {NOT_OLEAUTO(PSafeArray)}OleVariant); dispid 1610743810;
    procedure CallSomethingElse(const SomethingElse: ISomethingElse); dispid 1610743811;
  end;

// *********************************************************************//
// Interface: ISomethingElse
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {00A7D870-AE65-4682-BD1B-D2EE21B0B88C}
// *********************************************************************//
  ISomethingElse = interface(IDispatch)
    ['{00A7D870-AE65-4682-BD1B-D2EE21B0B88C}']
    procedure ShowMessage(const message: WideString); safecall;
  end;

// *********************************************************************//
// DispIntf:  ISomethingElseDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {00A7D870-AE65-4682-BD1B-D2EE21B0B88C}
// *********************************************************************//
  ISomethingElseDisp = dispinterface
    ['{00A7D870-AE65-4682-BD1B-D2EE21B0B88C}']
    procedure ShowMessage(const message: WideString); dispid 1610743808;
  end;

// *********************************************************************//
// The Class CoSomething provides a Create and CreateRemote method to          
// create instances of the default interface ISomething exposed by              
// the CoClass Something. The functions are intended to be used by             
// clients wishing to automate the CoClass objects exposed by the         
// server of this typelibrary.                                            
// *********************************************************************//
  CoSomething = class
    class function Create: ISomething;
    class function CreateRemote(const MachineName: string): ISomething;
  end;


// *********************************************************************//
// OLE Server Proxy class declaration
// Server Object    : TSomething
// Help String      : 
// Default Interface: ISomething
// Def. Intf. DISP? : No
// Event   Interface: 
// TypeFlags        : (2) CanCreate
// *********************************************************************//
  TSomething = class(TOleServer)
  private
    FIntf: ISomething;
    function GetDefaultInterface: ISomething;
  protected
    procedure InitServerData; override;
  public
    constructor Create(AOwner: TComponent); override;
    destructor  Destroy; override;
    procedure Connect; override;
    procedure ConnectTo(svrIntf: ISomething);
    procedure Disconnect; override;
    procedure SetIntArrayCallback(integerPointer: Integer);
    procedure SetSomethingElseCallback(integerPointer: Integer);
    procedure Reverse(ids: PSafeArray);
    procedure CallSomethingElse(const SomethingElse: ISomethingElse);
    property DefaultInterface: ISomething read GetDefaultInterface;
  published
  end;

// *********************************************************************//
// The Class CoSomethingElse provides a Create and CreateRemote method to          
// create instances of the default interface ISomethingElse exposed by              
// the CoClass SomethingElse. The functions are intended to be used by             
// clients wishing to automate the CoClass objects exposed by the         
// server of this typelibrary.                                            
// *********************************************************************//
  CoSomethingElse = class
    class function Create: ISomethingElse;
    class function CreateRemote(const MachineName: string): ISomethingElse;
  end;


// *********************************************************************//
// OLE Server Proxy class declaration
// Server Object    : TSomethingElse
// Help String      : 
// Default Interface: ISomethingElse
// Def. Intf. DISP? : No
// Event   Interface: 
// TypeFlags        : (2) CanCreate
// *********************************************************************//
  TSomethingElse = class(TOleServer)
  private
    FIntf: ISomethingElse;
    function GetDefaultInterface: ISomethingElse;
  protected
    procedure InitServerData; override;
  public
    constructor Create(AOwner: TComponent); override;
    destructor  Destroy; override;
    procedure Connect; override;
    procedure ConnectTo(svrIntf: ISomethingElse);
    procedure Disconnect; override;
    procedure ShowMessage(const message: WideString);
    property DefaultInterface: ISomethingElse read GetDefaultInterface;
  published
  end;

procedure Register;

resourcestring
  dtlServerPage = 'ActiveX';

  dtlOcxPage = 'ActiveX';

implementation

uses System.Win.ComObj;

class function CoSomething.Create: ISomething;
begin
  Result := CreateComObject(CLASS_Something) as ISomething;
end;

class function CoSomething.CreateRemote(const MachineName: string): ISomething;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_Something) as ISomething;
end;

procedure TSomething.InitServerData;
const
  CServerData: TServerData = (
    ClassID:   '{B2BB5506-8308-400C-924F-4A3E3D46D965}';
    IntfIID:   '{79801B97-62BC-489B-935A-8543116E620E}';
    EventIID:  '';
    LicenseKey: nil;
    Version: 500);
begin
  ServerData := @CServerData;
end;

procedure TSomething.Connect;
var
  punk: IUnknown;
begin
  if FIntf = nil then
  begin
    punk := GetServer;
    Fintf:= punk as ISomething;
  end;
end;

procedure TSomething.ConnectTo(svrIntf: ISomething);
begin
  Disconnect;
  FIntf := svrIntf;
end;

procedure TSomething.DisConnect;
begin
  if Fintf <> nil then
  begin
    FIntf := nil;
  end;
end;

function TSomething.GetDefaultInterface: ISomething;
begin
  if FIntf = nil then
    Connect;
  Assert(FIntf <> nil, 'DefaultInterface is NULL. Component is not connected to Server. You must call "Connect" or "ConnectTo" before this operation');
  Result := FIntf;
end;

constructor TSomething.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
end;

destructor TSomething.Destroy;
begin
  inherited Destroy;
end;

procedure TSomething.SetIntArrayCallback(integerPointer: Integer);
begin
  DefaultInterface.SetIntArrayCallback(integerPointer);
end;

procedure TSomething.SetSomethingElseCallback(integerPointer: Integer);
begin
  DefaultInterface.SetSomethingElseCallback(integerPointer);
end;

procedure TSomething.Reverse(ids: PSafeArray);
begin
  DefaultInterface.Reverse(ids);
end;

procedure TSomething.CallSomethingElse(const SomethingElse: ISomethingElse);
begin
  DefaultInterface.CallSomethingElse(SomethingElse);
end;

class function CoSomethingElse.Create: ISomethingElse;
begin
  Result := CreateComObject(CLASS_SomethingElse) as ISomethingElse;
end;

class function CoSomethingElse.CreateRemote(const MachineName: string): ISomethingElse;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_SomethingElse) as ISomethingElse;
end;

procedure TSomethingElse.InitServerData;
const
  CServerData: TServerData = (
    ClassID:   '{165B7823-3B02-4DAA-8205-C07FB7FA124A}';
    IntfIID:   '{00A7D870-AE65-4682-BD1B-D2EE21B0B88C}';
    EventIID:  '';
    LicenseKey: nil;
    Version: 500);
begin
  ServerData := @CServerData;
end;

procedure TSomethingElse.Connect;
var
  punk: IUnknown;
begin
  if FIntf = nil then
  begin
    punk := GetServer;
    Fintf:= punk as ISomethingElse;
  end;
end;

procedure TSomethingElse.ConnectTo(svrIntf: ISomethingElse);
begin
  Disconnect;
  FIntf := svrIntf;
end;

procedure TSomethingElse.DisConnect;
begin
  if Fintf <> nil then
  begin
    FIntf := nil;
  end;
end;

function TSomethingElse.GetDefaultInterface: ISomethingElse;
begin
  if FIntf = nil then
    Connect;
  Assert(FIntf <> nil, 'DefaultInterface is NULL. Component is not connected to Server. You must call "Connect" or "ConnectTo" before this operation');
  Result := FIntf;
end;

constructor TSomethingElse.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
end;

destructor TSomethingElse.Destroy;
begin
  inherited Destroy;
end;

procedure TSomethingElse.ShowMessage(const message: WideString);
begin
  DefaultInterface.ShowMessage(message);
end;

procedure Register;
begin
  RegisterComponents(dtlServerPage, [TSomething, TSomethingElse]);
end;

end.

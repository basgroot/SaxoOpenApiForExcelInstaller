; More advanced example: https://github.com/bovender/ExcelAddinInstaller

[Setup]
AppName=SaxoBank OpenAPI for Excel
AppVersion=1.4.0
WizardStyle=modern
DefaultDirName={autopf}\SaxoBankForExcel
DefaultGroupName=Saxo Bank Excel
UninstallDisplayIcon=SaxoBankLogo.ico
Compression=lzma2
SolidCompression=yes
PrivilegesRequired=lowest
OutputBaseFilename=SaxoBankForExcelInstaller
OutputDir="C:\test\"
AppContact=Saxo Bank OpenAPI support: openapisupport@saxobank.com
AppCopyright=Copyright (C) 2022 Saxo Bank
AppPublisher=Saxo Bank
AppPublisherURL=https://www.developer.saxo/excel/home
AppSupportURL=https://openapi.help.saxo/hc/en-us
// Todo close Excel before installing
;RestartApplications=
InfoAfterFile=README.txt
SetupIconFile=SaxoBankLogo.ico
WizardImageFile=WizardImage1.bmp
WizardSmallImageFile=WizardSmallImage1.bmp

[Files]
Source: "OpenApi-AddIn-32bits-Mar-10-2022.xll"; DestDir: "{userappdata}\Microsoft\AddIns"; Check: "not Is64BitExcelFromRegisteredExe"
Source: "OpenApi-AddIn-64bits-Mar-10-2022.xll"; DestDir: "{userappdata}\Microsoft\AddIns"; Check: "Is64BitExcelFromRegisteredExe"
; Always install both DLLs for manual copying:
Source: "OpenApi-AddIn-32bits-Mar-10-2022.xll"; DestDir: "{app}"
Source: "OpenApi-AddIn-64bits-Mar-10-2022.xll"; DestDir: "{app}"
Source: "README.txt"; DestDir: "{app}"; Flags: isreadme

[InstallDelete]
; Add all previous file names here, to be sure there is only one OpenApi XLL:
; TODO Add wildcard (and use Saxo specific file names)
Type: files; Name: "{userappdata}\Microsoft\AddIns\OpenApi-AddIn-32bits-Mar-10-2022.xll"
Type: files; Name: "{userappdata}\Microsoft\AddIns\OpenApi-AddIn-64bits-Mar-10-2022.xll"

[Code]
// Source: https://stackoverflow.com/questions/2203980/detect-whether-office-is-32bit-or-64bit-via-the-registry
const
  { Constants for GetBinaryType return values. }
  SCS_32BIT_BINARY = 0;
  SCS_64BIT_BINARY = 6;
  { There are other values that GetBinaryType can return, but we're not interested in them. }

{ Declare Win32 function  }
function GetBinaryType(lpApplicationName: AnsiString; var lpBinaryType: Integer): Boolean;
external 'GetBinaryTypeA@kernel32.dll stdcall';

function Is64BitExcelFromRegisteredExe(): Boolean;
var
  excelPath: String;
  binaryType: Integer;
begin
  Result := False; { Default value - assume 32-bit unless proven otherwise. }
  { RegQueryStringValue second param is '' to get the (default) value for the key }
  { with no sub-key name, as described at }
  { http://stackoverflow.com/questions/913938/ }
  if IsWin64() and RegQueryStringValue(HKEY_LOCAL_MACHINE, 
      'SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe',
      '', excelPath) then begin
    { We've got the path to Excel. }
    try
      Log('*** Found Office Excel path: ' + excelPath);
      if GetBinaryType(excelPath, binaryType) then begin
        Result := (binaryType = SCS_64BIT_BINARY);
        Log('*** Found binary type: ' + IntToStr(binaryType));
      end;
    except
      { Ignore - better just to assume it's 32-bit than to let the installation }
      { fail.  This could fail because the GetBinaryType function is not }
      { available.  I understand it's only available in Windows 2000 }
      { Professional onwards. }
    end;
  end;
end;
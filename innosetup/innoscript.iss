[Files]
Source: innosetup\innoscript.iss; DestDir: {app}\innosetup\;
Source: innosetup\python_packages\vcredist_x86.exe; DestDir: {app}\innosetup\python_packages\;
Source: innosetup\python_packages\python-3.4.1.msi; DestDir: {app}\innosetup\python_packages\;
Source: innosetup\python_packages\wheel-0.24.0.tar.gz; DestDir: {app}\innosetup\python_packages\;
Source: innosetup\python_packages\wheelhouse\docutils-0.11-py3-none-any.whl; DestDir: {app}\innosetup\python_packages\wheelhouse\;
Source: innosetup\python_packages\wheelhouse\Jinja2-2.7.3-py3-none-any.whl; DestDir: {app}\innosetup\python_packages\wheelhouse\;
Source: innosetup\python_packages\wheelhouse\MarkupSafe-0.23-cp34-none-win32.whl; DestDir: {app}\innosetup\python_packages\wheelhouse\;
Source: innosetup\python_packages\wheelhouse\Pygments-1.6-py3-none-any.whl; DestDir: {app}\innosetup\python_packages\wheelhouse\;
Source: innosetup\python_packages\wheelhouse\PySide-1.2.2-cp34-none-win32.whl; DestDir: {app}\innosetup\python_packages\wheelhouse\;
Source: innosetup\python_packages\wheelhouse\Pyvot-0.1.2-py3-none-any.whl; DestDir: {app}\innosetup\python_packages\wheelhouse\;
Source: innosetup\python_packages\wheelhouse\pywin32-219-cp34-none-win32.whl; DestDir: {app}\innosetup\python_packages\wheelhouse\;
Source: innosetup\python_packages\wheelhouse\Sphinx-1.2.2-py33-none-any.whl; DestDir: {app}\innosetup\python_packages\wheelhouse\;
Source: "transition.py"; DestDir: {app};
Source: "COPYING"; DestDir: {app};
Source: "COPYING.LESSER"; DestDir: {app};
Source: "excelapps\__init__.py"; DestDir: {app}\excelapps\;
Source: "excelapps\appskell.py"; DestDir: {app}\excelapps\;
Source: "excelapps\dummy\__init__.py"; DestDir: {app}\excelapps\dummy\;
Source: "exceladdins\__init__.py"; DestDir: {app}\exceladdins\;
Source: "exceladdins\addinskell.py"; DestDir: {app}\exceladdins\;
Source: "exceladdins\config\__init__.py"; DestDir: {app}\exceladdins\config\;
Source: "exceladdins\config\configmain.py"; DestDir: {app}\exceladdins\config\;
Source: "exceladdins\config\config_box.py"; DestDir: {app}\exceladdins\config\;
Source: "exceladdins\config\config_box.ui"; DestDir: {app}\exceladdins\config\;
Source: "transitionconfig\__init__.py"; DestDir: {app}\transitionconfig\;
Source: "transitioncore\__init__.py"; DestDir: {app}\transitioncore\;
Source: "transitioncore\excelappevents.py"; DestDir: {app}\transitioncore\;
Source: "transitioncore\excelapphandler.py"; DestDir: {app}\transitioncore\;
Source: "transitioncore\excelwbevents.py"; DestDir: {app}\transitioncore\;

[Setup]
AppCopyright=2014 Jonathan Besanceney
AppName=Transition Excel COM Add-in
AppVerName=Transition 0.1.23
PrivilegesRequired=admin
AppID={{8A3F0510-E407-40D0-8954-A8C7AE097942}
Compression=lzma2/Ultra64
InternalCompressLevel=Ultra64
DefaultGroupName=Transition
DefaultDirName={pf}\Transition
OutputBaseFilename=TransitionSetup
UninstallLogMode=append
SourceDir=..
OutputDir=innosetup\Output
VersionInfoCompany=Jonathan Besanceney
VersionInfoDescription=Transition Excel/COM Plugin
VersionInfoCopyright=2014 Jonathan Besanceney
VersionInfoProductName=Transition
SolidCompression=true

[Run]
Filename: "{app}\innosetup\python_packages\vcredist_x86.exe"; Check: VCRedistNeedsInstall; Parameters: "/passive /Q:a /c:""msiexec /qb /i vcredist.msi"" "; StatusMsg: "Installing... Visual C++ 2010 RunTime";
Filename: msiexec.exe; Parameters: "/I ""{app}\innosetup\python_packages\python-3.4.1.msi"" /passive TARGETDIR=c:\python34 ADDLOCAL=ALL"; WorkingDir: {app}\innosetup; Description: "Install Python 3.4.1"; Flags: RunAsCurrentUser; StatusMsg: "Installing... Python 3.4.1";
Filename: c:\python34\Scripts\pip.exe; Parameters: "install ""{app}\innosetup\python_packages\wheel-0.24.0.tar.gz"""; WorkingDir: {app}\innosetup; Description: "Install Wheel 0.24.0"; StatusMsg: "Installing... Wheel 0.24.0"; Flags: RunAsCurrentUser RunHidden;
Filename: c:\python34\Scripts\pip.exe; Parameters: "install --use-wheel --no-index --find-links=""{app}\innosetup\python_packages\wheelhouse"" pywin32==219"; WorkingDir: {app}\innosetup; Description: "Install PyWin32 v219"; StatusMsg: "Installing... PyWin32 v219"; Flags: RunAsCurrentUser RunHidden;
Filename: c:\python34\python.exe; Parameters: "C:\Python34\Scripts\pywin32_postinstall.py -install"; WorkingDir: c:\python34\; Description: "Post-install PyWin32 v219"; StatusMsg: "Post-Installing... PyWin32 v219"; Flags: RunAsCurrentUser RunHidden; 
Filename: c:\python34\Scripts\pip.exe; Parameters: "install --use-wheel --no-index --find-links=""{app}\innosetup\python_packages\wheelhouse"" PySide"; WorkingDir: {app}\innosetup; Description: "Install PySide 1.2.2 (QT bindings for Python)"; StatusMsg: "Install... PySide 1.2.2 (QT bindings for Python)"; Flags: RunAsCurrentUser RunHidden;

Filename: c:\python34\Scripts\pip.exe; Parameters: "install --use-wheel --no-index --find-links=""{app}\innosetup\python_packages\wheelhouse"" docutils"; WorkingDir: {app}\innosetup; Description: "Install docutils 0.11"; StatusMsg: "Install... docutils 0.11"; Flags: RunAsCurrentUser RunHidden;
Filename: c:\python34\Scripts\pip.exe; Parameters: "install --use-wheel --no-index --find-links=""{app}\innosetup\python_packages\wheelhouse"" Jinja2"; WorkingDir: {app}\innosetup; Description: "Install Jinja2 2.7.3"; StatusMsg: "Install... Jinja2 2.7.3"; Flags: RunAsCurrentUser RunHidden;
Filename: c:\python34\Scripts\pip.exe; Parameters: "install --use-wheel --no-index --find-links=""{app}\innosetup\python_packages\wheelhouse"" MarkupSafe"; WorkingDir: {app}\innosetup; Description: "Install MarkupSafe 0.23"; StatusMsg: "Install... MarkupSafe 0.23"; Flags: RunAsCurrentUser RunHidden;
Filename: c:\python34\Scripts\pip.exe; Parameters: "install --use-wheel --no-index --find-links=""{app}\innosetup\python_packages\wheelhouse"" Pygments"; WorkingDir: {app}\innosetup; Description: "Install Pygments 1.6"; StatusMsg: "Install... Pygments 1.6"; Flags: RunAsCurrentUser RunHidden;
Filename: c:\python34\Scripts\pip.exe; Parameters: "install --use-wheel --no-index --find-links=""{app}\innosetup\python_packages\wheelhouse"" Sphinx"; WorkingDir: {app}\innosetup; Description: "Install Sphinx 1.2.2"; StatusMsg: "Install... Sphinx 1.2.2"; Flags: RunAsCurrentUser RunHidden;
Filename: c:\python34\Scripts\pip.exe; Parameters: "install --use-wheel --no-index --find-links=""{app}\innosetup\python_packages\wheelhouse"" Pyvot"; WorkingDir: {app}\innosetup; Description: "Install Pyvot 0.1.2 (A Python to/from Excel Connector by Microsoft)"; StatusMsg: "Install... Pyvot 0.1.2 (A Python to/from Excel Connector by Microsoft)"; Flags: RunAsCurrentUser RunHidden;

Filename: c:\python34\python.exe; Parameters: """{app}\transition.py"" --debug"; WorkingDir: {app}; Description: "Register Transition Excel/COM Add-in"; StatusMsg: "Registering Transition Excel/COM Add-in..."; Flags: RunAsCurrentUser RunHidden;
Filename: c:\python34\python.exe; Parameters: """{app}\transition.py"" --addin-enable config"; WorkingDir: {app}; Description: "Configuration Plugin Activation"; StatusMsg: "Configuration Plugin Activation..."; Flags: RunAsCurrentUser RunHidden; 
Filename: c:\python34\python.exe; Parameters: """{app}\transition.py"" --app-enable dummy"; WorkingDir: {app}; Description: "Dummy app Activation"; StatusMsg: "Dummy app Activation..."; Flags: RunAsCurrentUser RunHidden; 
Filename: C:\Python34\Lib\site-packages\pythonwin\pythonwin.exe; WorkingDir: {app}; Description: "Launch PythonWin. To have an execution trace, clic on Tools->Trace Collector Debugging tool"; Flags: PostInstall NoWait; 


[Icons]
Name: "{group}\Configuration"; Filename: c:\python34\python.exe; WorkingDir: {app}; Parameters: """{app}\exceladdins\config\__init__.py"""; 
Name: "{group}\dummy excel app"; Filename: c:\python34\python.exe; WorkingDir: {app}; Parameters: """{app}\excelapps\dummy\__init__.py"""; 
Name: "{group}\Register Transition Add-in (debug)"; Filename: c:\python34\python.exe; Parameters: " ""{app}\transition.py"" --debug"; WorkingDir: {app}; 
Name: "{group}\Register Transition Add-in (normal)"; Filename: c:\python34\python.exe; Parameters: " ""{app}\transition.py"""; WorkingDir: {app}; 
Name: "{group}\Unregister Transition Add-in"; Filename: c:\python34\python.exe; Parameters: " ""{app}\transition.py"" --unregister"; WorkingDir: {app}; 
Name: "{group}\Uninstall Transition Add-in"; Filename: {app}\unins000.exe; WorkingDir: {app}; Comment: "Uninstall Transition";
Name: "Python 3.4\Qt Designer"; Filename: C:\Python34\Lib\site-packages\PySide\designer.exe; WorkingDir: {app}; IconFilename: C:\Python34\Lib\site-packages\PySide\designer.exe; 

[Dirs]
Name: {app}\excelapps; 
Name: {app}\excelapps\dummy;
Name: {app}\innosetup;
Name: {app}\innosetup\python_packages;
Name: {app}\innosetup\python_packages\wheelhouse;
Name: {app}\transitioncore;
Name: {app}\transitionconfig;
Name: {app}\exceladdins; 
Name: {app}\exceladdins\config; 


[UninstallRun]
Filename: c:\Python34\Python.exe; Parameters: """{app}\transition.py"" --unregister"; WorkingDir: {app}; Flags: RunHidden; StatusMsg: "Unregister Transition Add-in...";
Filename: {syswow64}\msiexec.exe; WorkingDir: {app}; StatusMsg: "Uninstalling Python 3.4.1..."; Parameters: "/x ""{app}\innosetup\python_packages\python-3.4.1.msi"" /passive"; 

[UninstallDelete]
Name: {app}\__pycache__; Type: filesandordirs; 
Name: {app}\excel_apps\__pycache__; Type: filesandordirs;
Name: {app}\excel_apps\dummy\__pycache__; Type: filesandordirs; 
Name: {app}\excel_addins\__pycache__; Type: filesandordirs; 
Name: {app}\excel_addins\config\__pycache__; Type: filesandordirs;
Name: {app}\innosetup; Type: filesandordirs;
Name: {app}\transitioncore; Type: filesandordirs;
Name: {app}\transitionconfig; Type: filesandordirs;
Name: {syswow64}\pythoncom34.dll; Type: files;
Name: {syswow64}\pywintypes34.dll; Type: files;
Name: c:\Python34; Type: filesandordirs;

[Code]
//http://stackoverflow.com/questions/11137424/how-to-make-vcredist-x86-reinstall-only-if-not-yet-installed

#IFDEF UNICODE
  #DEFINE AW "W"
#ELSE
  #DEFINE AW "A"
#ENDIF
type
  INSTALLSTATE = Longint;
const
  INSTALLSTATE_INVALIDARG = -2;  // An invalid parameter was passed to the function.
  INSTALLSTATE_UNKNOWN = -1;     // The product is neither advertised or installed.
  INSTALLSTATE_ADVERTISED = 1;   // The product is advertised but not installed.
  INSTALLSTATE_ABSENT = 2;       // The product is installed for a different user.
  INSTALLSTATE_DEFAULT = 5;      // The product is installed for the current user.

  VC_2005_REDIST_X86 = '{A49F249F-0C91-497F-86DF-B2585E8E76B7}';
  VC_2005_REDIST_X64 = '{6E8E85E8-CE4B-4FF5-91F7-04999C9FAE6A}';
  VC_2005_REDIST_IA64 = '{03ED71EA-F531-4927-AABD-1C31BCE8E187}';
  VC_2005_SP1_REDIST_X86 = '{7299052B-02A4-4627-81F2-1818DA5D550D}';
  VC_2005_SP1_REDIST_X64 = '{071C9B48-7C32-4621-A0AC-3F809523288F}';
  VC_2005_SP1_REDIST_IA64 = '{0F8FB34E-675E-42ED-850B-29D98C2ECE08}';
  VC_2005_SP1_ATL_SEC_UPD_REDIST_X86 = '{837B34E3-7C30-493C-8F6A-2B0F04E2912C}';
  VC_2005_SP1_ATL_SEC_UPD_REDIST_X64 = '{6CE5BAE9-D3CA-4B99-891A-1DC6C118A5FC}';
  VC_2005_SP1_ATL_SEC_UPD_REDIST_IA64 = '{85025851-A784-46D8-950D-05CB3CA43A13}';

  VC_2008_REDIST_X86 = '{FF66E9F6-83E7-3A3E-AF14-8DE9A809A6A4}';
  VC_2008_REDIST_X64 = '{350AA351-21FA-3270-8B7A-835434E766AD}';
  VC_2008_REDIST_IA64 = '{2B547B43-DB50-3139-9EBE-37D419E0F5FA}';
  VC_2008_SP1_REDIST_X86 = '{9A25302D-30C0-39D9-BD6F-21E6EC160475}';
  VC_2008_SP1_REDIST_X64 = '{8220EEFE-38CD-377E-8595-13398D740ACE}';
  VC_2008_SP1_REDIST_IA64 = '{5827ECE1-AEB0-328E-B813-6FC68622C1F9}';
  VC_2008_SP1_ATL_SEC_UPD_REDIST_X86 = '{1F1C2DFC-2D24-3E06-BCB8-725134ADF989}';
  VC_2008_SP1_ATL_SEC_UPD_REDIST_X64 = '{4B6C7001-C7D6-3710-913E-5BC23FCE91E6}';
  VC_2008_SP1_ATL_SEC_UPD_REDIST_IA64 = '{977AD349-C2A8-39DD-9273-285C08987C7B}';
  VC_2008_SP1_MFC_SEC_UPD_REDIST_X86 = '{9BE518E6-ECC6-35A9-88E4-87755C07200F}';
  VC_2008_SP1_MFC_SEC_UPD_REDIST_X64 = '{5FCE6D76-F5DC-37AB-B2B8-22AB8CEDB1D4}';
  VC_2008_SP1_MFC_SEC_UPD_REDIST_IA64 = '{515643D1-4E9E-342F-A75A-D1F16448DC04}';

  VC_2010_REDIST_X86 = '{196BB40D-1578-3D01-B289-BEFC77A11A1E}';
  VC_2010_REDIST_X64 = '{DA5E371C-6333-3D8A-93A4-6FD5B20BCC6E}';
  VC_2010_REDIST_IA64 = '{C1A35166-4301-38E9-BA67-02823AD72A1B}';
  VC_2010_SP1_REDIST_X86 = '{F0C3E5D1-1ADE-321E-8167-68EF0DE699A5}';
  VC_2010_SP1_REDIST_X64 = '{1D8E6291-B0D5-35EC-8441-6616F567A0F7}';
  VC_2010_SP1_REDIST_IA64 = '{88C73C1C-2DE5-3B01-AFB8-B46EF4AB41CD}';

function MsiQueryProductState(szProduct: string): INSTALLSTATE; 
  external 'MsiQueryProductState{#AW}@msi.dll stdcall';

function VCVersionInstalled(const ProductID: string): Boolean;
begin
  Result := MsiQueryProductState(ProductID) = INSTALLSTATE_DEFAULT;
end;

function VCRedistNeedsInstall: Boolean;
begin
  // here the Result must be True when you need to install your VCRedist
  // or False when you don't need to, so now it's upon you how you build
  // this statement, the following won't install your VC redist only when
  // the Visual C++ 2010 Redist (x86) and Visual C++ 2010 SP1 Redist(x86)
  // are installed for the current user
  Result := not (VCVersionInstalled(VC_2010_REDIST_X86) and 
    VCVersionInstalled(VC_2010_SP1_REDIST_X86));
end;

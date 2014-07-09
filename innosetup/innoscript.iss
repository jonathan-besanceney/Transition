[Files]
Source: innosetup\innoscript.iss; DestDir: {app}\innosetup\; 
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
MinVersion=,5.1.2600sp1
VersionInfoCompany=Jonathan Besanceney
VersionInfoDescription=Transition Excel/COM Plugin
VersionInfoCopyright=2014 Jonathan Besanceney
VersionInfoProductName=Transition
SolidCompression=true

[Run]
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

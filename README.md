Transition
==========

Transition Excel/COM Add-in

This Python project aims to replace VBA nightmares by pretty Python scripts. 
It uses Mark Hammond's Pywin32 and Microsoft's Pyvot Python packages to make a ready to use Excel COM/Add-in.

You can download the installer (includes Python 3.4.1 and all Transition dependencies => 70.9 Mo) here : 

https://github.com/jonathan-besanceney/Transition/raw/master/innosetup/Output/TransitionSetup.exe

To see why setup file is fat : https://github.com/jonathan-besanceney/Transition/blob/master/innosetup/innoscript.iss

## 1. Architecture

```
------------------------------------------------------------------------
| User's Python Excel Add-ins      | User's Python Excel Workbook apps |
------------------------------------------------------------------------
| Add-in skeleton                  | App skeleton (include Pyvot)      |
------------------------------------------------------------------------
| CORE :                           | Config :                          |
| - Register/Unregister Transition | - Manage your add-ins and apps    |
| - Workbook app Handler           | - Get info on your add-ins and    |
| - Excel Events                   | apps                              |
| - Workbook Events                |                                   |
------------------------------------------------------------------------
| Transition Excel/COM Add-in                                          |
------------------------------------------------------------------------
| Pywin32 COM interface                                                |
------------------------------------------------------------------------
| Excel                                                                |
------------------------------------------------------------------------
```

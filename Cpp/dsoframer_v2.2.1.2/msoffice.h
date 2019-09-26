
#ifndef DSO_MSOFFICE_
#define DSO_MSOFFICE_

#if OFFICEVER == 16
#import "C:\Program Files (x86)\Microsoft Office\root\VFS\ProgramFilesCommonX86\Microsoft Shared\OFFICE16\mso.dll" rename("ColorFormat", "ColorFormatEx"),rename_namespace("Office")
#import "C:\Program Files (x86)\Microsoft Office\root\VFS\ProgramFilesCommonX86\Microsoft Shared\VBA\VBA6\VBE6EXT.olb" rename_namespace("VBE6")
#import "C:\Program Files (x86)\Microsoft Office\root\Office16\MSWORD.olb" rename("ExitWindows","ExitWindowsEx"),rename_namespace("MSWord")
#pragma warning (disable: 4192)
#import "C:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.EXE" rename("RGB","RGBEx"),rename("DialogBox","DialogBoxEx"),rename_namespace("MSExcel")
#pragma warning (default: 4192)
#import "C:\Program Files (x86)\Microsoft Office\root\Office16\MSPPT.OLB" named_guids,rename_namespace("MSPPT")

#else
//#define SUPPORT_WPS
// #import "C:\Program Files\Common Files\DESIGNER\MSADDNDR.DLL" raw_interfaces_only, raw_native_types, no_namespace, named_guids 

#import "C:\Program Files (x86)\Common Files\microsoft shared\OFFICE14\mso.dll" rename("ColorFormat", "ColorFormatEx"),rename_namespace("Office")
//using namespace Office;

#import "C:\Program Files (x86)\Common Files\microsoft shared\VBA\VBA6\VBE6EXT.olb" rename_namespace("VBE6")
//using namespace VBE6;


#import "C:\Program Files (x86)\Microsoft Office\Office14\MSWORD.olb" rename("ExitWindows","ExitWindowsEx"),rename_namespace("MSWord")
//using namespace MSWord;

#import "C:\Program Files (x86)\Microsoft Office\Office14\EXCEL.EXE" rename("RGB","RGBEx"),rename("DialogBox","DialogBoxEx"),rename_namespace("MSExcel")
//using namespace MSExcel;

#import "C:\Program Files (x86)\Microsoft Office\Office14\MSPPT.OLB" named_guids,rename_namespace("MSPPT")
//using namespace MSPPT;
#endif
 
//#ifdef SUPPORT_WPS
//#import "D:\Program Files\Kingsoft\WPS Office 2005 Professional\office6\kso10.dll" rename_namespace("Wps")
//using namespace Wps;
//#import "D:\Program Files\Kingsoft\WPS Office 2005 Professional\office6\wpscore.dll" rename_namespace("Wps")
//using namespace Wps;
//#endif


#endif
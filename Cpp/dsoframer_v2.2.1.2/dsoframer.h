/***************************************************************************
 * DSOFRAMER.H
 *
 * Developer Support Office ActiveX Document Framer Control Sample
 *
 *  Copyright ?999-2004; Microsoft Corporation. All rights reserved.
 *  Written by Microsoft Developer Support Office Integration (PSS DSOI)
 * 
 *  This code is provided via KB 311765 as a sample. It is not a formal
 *  product and has not been tested with all containers or servers. Use it
 *  for educational purposes only.
 *
 *  You have a royalty-free right to use, modify, reproduce and distribute
 *  this sample application, and/or any modified version, in any way you
 *  find useful, provided that you agree that Microsoft has no warranty,
 *  obligations or liability for the code or information provided herein.
 *
 *  THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND,
 *  EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED
 *  WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
 *
 *  See the EULA.TXT file included in the KB download for full terms of use
 *  and restrictions. You should consult documentation on MSDN Library for
 *  possible updates or changes to behaviors or interfaces used in this sample.
 *
 ***************************************************************************/
#ifndef DS_DSOFRAMER_H 
#define DS_DSOFRAMER_H

#include <SDKDDKVer.h>

#include "XMLHttpClient.h"
#include "afxinet.h" // ftp 
////////////////////////////////////////////////////////////////////
// We compile at level 4 and disable some unnecessary warnings...
//
#pragma warning(push, 4) // Compile at level-4 warnings
#pragma warning(disable: 4100) // unreferenced formal parameter (in OLE this is common)
#pragma warning(disable: 4146) // unary minus operator applied to unsigned type, result still unsigned
#pragma warning(disable: 4268) // const static/global data initialized with compiler generated default constructor
#pragma warning(disable: 4310) // cast truncates constant value
#pragma warning(disable: 4786) // identifier was truncated in the debug information

////////////////////////////////////////////////////////////////////
// Needed include files (both standard and custom)
//
//#include <windows.h>


#include <afxwin.h>
#include <winspool.h>
#include <ole2.h>
#include <olectl.h>
#include <oleidl.h>
#include <objsafe.h>

#pragma warning (disable: 4005 4278)

#include "version.h"
#include "utilities.h"
#include "dsofdocobj.h"
#include ".\lib\dsoframerlib.h"
#include ".\res\resource.h"
#include "msoffice.h"

#pragma warning (default: 4005 4278)

#define 	FILE_TYPE_NULL   0  
#define 	FILE_TYPE_WORD  11 
#define 	FILE_TYPE_EXCEL   12 
#define 	FILE_TYPE_PPT   13 
#define 	FILE_TYPE_RTF   14 
#define 	FILE_TYPE_WPS   21 
#define 	FILE_TYPE_PDF   31 
#define 	FILE_TYPE_UNK   127



#define E_OK			0 ;
#define E_NOLOAD		-1;//û��װ���ļ�
#define E_NOBOOKMARK	-50;
#define E_NOSHEET		-51;
#define E_NOSUPPORT     -126
#define E_UNKNOW		-127;
#define E_EXCEPTIONAL	 -255;

////////////////////////////////////////////////////////////////////
// Global Variables
//
extern HINSTANCE        v_hModule;
extern CRITICAL_SECTION v_csecThreadSynch;
extern HICON            v_icoOffDocIcon;
extern ULONG            v_cLocks;
extern BOOL             v_fUnicodeAPI;
extern BOOL             v_fWindows2KPlus;

////////////////////////////////////////////////////////////////////
// Custom Errors - we support a very limited set of custom error messages
//
#define DSO_E_ERR_BASE              0x80041100
#define DSO_E_UNKNOWN               0x80041101   // "An unknown problem has occurred."
#define DSO_E_INVALIDPROGID         0x80041102   // "The ProgID/Template could not be found or is not associated with a COM server."
#define DSO_E_INVALIDSERVER         0x80041103   // "The associated COM server does not support ActiveX Document embedding."
#define DSO_E_COMMANDNOTSUPPORTED   0x80041104   // "The command is not supported by the document server."
#define DSO_E_DOCUMENTREADONLY      0x80041105   // "Unable to perform action because document was opened in read-only mode."
#define DSO_E_REQUIRESMSDAIPP       0x80041106   // "The Microsoft Internet Publishing Provider is not installed, so the URL document cannot be open for write access."
#define DSO_E_DOCUMENTNOTOPEN       0x80041107   // "No document is open to perform the operation requested."
#define DSO_E_INMODALSTATE          0x80041108   // "Cannot access document when in modal condition."
#define DSO_E_NOTBEENSAVED          0x80041109   // "Cannot Save file without a file path."
#define DSO_E_ERR_MAX               0x8004110A

////////////////////////////////////////////////////////////////////
// Custom OLE Command IDs - we use for special tasks
//
#define OLECMDID_GETDATAFORMAT      0x7001  // 28673
#define OLECMDID_SETDATAFORMAT      0x7002  // 28674

////////////////////////////////////////////////////////////////////
// Custom Window Messages (only apply to CDsoFramerControl window proc)
//
#define DSO_WM_ASYNCH_OLECOMMAND         (WM_USER + 300)
#define DSO_WM_ASYNCH_STATECHANGE        (WM_USER + 301)
// State Flags for DSO_WM_ASYNCH_STATECHANGE:
#define DSO_STATE_MODAL            1
#define DSO_STATE_ACTIVATION       2
#define DSO_STATE_INTERACTIVE      3

////////////////////////////////////////////////////////////////////
// Menu Bar Items
//
#define DSO_MAX_MENUITEMS         16
#define DSO_MAX_MENUNAME_LENGTH   32

#ifndef DT_HIDEPREFIX
#define DT_HIDEPREFIX             0x00100000
#define DT_PREFIXONLY             0x00200000
#endif

////////////////////////////////////////////////////////////////////
// Control Class Factory
//
class CDsoFramerClassFactory : public IClassFactory
{
public:
    CDsoFramerClassFactory(): m_cRef(0){}
    ~CDsoFramerClassFactory(void){}

 // IUnknown Implementation
    STDMETHODIMP         QueryInterface(REFIID riid, void ** ppv);
    STDMETHODIMP_(ULONG) AddRef(void);
    STDMETHODIMP_(ULONG) Release(void);

 // IClassFactory Implementation
    STDMETHODIMP  CreateInstance(LPUNKNOWN punk, REFIID riid, void** ppv);
    STDMETHODIMP  LockServer(BOOL fLock);

private:
    ULONG          m_cRef;          // Reference count

};

////////////////////////////////////////////////////////////////////
// CDsoFramerControl -- Main Control (OCX) Object 
//
//  The CDsoFramerControl control is standard OLE control designed around 
//  the OCX94 specification. Because we plan on doing custom integration to 
//  act as both OLE object and OLE host, it does not use frameworks like ATL 
//  or MFC which would only complicate the nature of the sample.
//
//  The control inherits from its automation interface, but uses nested 
//  classes for all OLE interfaces. This is not a requirement but does help
//  to clearly seperate the tasks done by each interface and makes finding 
//  ref count problems easier to spot since each interface carries its own
//  counter and will assert (in debug) if interface is over or under released.
//  
//  The control is basically a stage for the ActiveDocument embedding, and 
//  handles any external (user) commands. The task of actually acting as
//  a DocObject host is done in the site object CDsoDocObject, which this 
//  class creates and uses for the embedding.
//
#include <afxcmn.h>
#include <Mshtml.h>
#include "dsoframer.h" 
class CDsoFramerControl : public _FramerControl
{
public:
 

	IHTMLDocument2 *  m_spDoc;
	IDispatch *   m_pDocDispatch;

 

	char m_cPassWord[128];//Excel��ֻ������
	char m_cPWWrite[128];//Excel�Ŀ�д����
	char m_cCurPath[MAX_PATH];//��ǰ�ĵ���·��

    int  m_nOriginalFileType;//                m_clsidObject;         // CLSID of the embedded object
	char m_cUrl[1024];//����Http·��
    unsigned int        m_fInPlaceActive:1;        // are we in place active or not?
 	XMLHttpClient *m_pHttp; 

	//--------FTP
	CFtpConnection *m_pFtpConnection;
 
	CInternetSession * m_pSession;
	BOOL m_bConnect ;
    CDsoFramerControl(LPUNKNOWN punk);
    ~CDsoFramerControl(void);

 // IUnknown Implementation -- Always delgates to outer unknown...
    STDMETHODIMP         QueryInterface(REFIID riid, void ** ppv){return m_pOuterUnknown->QueryInterface(riid, ppv);}
    STDMETHODIMP_(ULONG) AddRef(void){return m_pOuterUnknown->AddRef();}
    STDMETHODIMP_(ULONG) Release(void){return m_pOuterUnknown->Release();}

 // IDispatch Implementation
    STDMETHODIMP GetTypeInfoCount(UINT* pctinfo); 
    STDMETHODIMP GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo** ppTInfo);
    STDMETHODIMP GetIDsOfNames(REFIID riid, LPOLESTR* rgszNames, UINT cNames, LCID lcid, DISPID* rgDispId);
    STDMETHODIMP Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS* pDispParams, VARIANT* pVarResult, EXCEPINFO* pExcepInfo, UINT* puArgErr);

 // _FramerControl Implementation
    STDMETHODIMP Activate();
    STDMETHODIMP get_ActiveDocument(IDispatch** ppdisp);
    STDMETHODIMP CreateNew(BSTR ProgIdOrTemplate);
    STDMETHODIMP Open(VARIANT Document, VARIANT ReadOnly, VARIANT ProgId, VARIANT WebUsername, VARIANT WebPassword);
    STDMETHODIMP Save(VARIANT SaveAsDocument, VARIANT OverwriteExisting, VARIANT WebUsername, VARIANT WebPassword);
    STDMETHODIMP _PrintOutOld(VARIANT PromptToSelectPrinter);
    STDMETHODIMP Close();
    STDMETHODIMP put_Caption(BSTR bstr);
    STDMETHODIMP get_Caption(BSTR* pbstr);
    STDMETHODIMP put_Titlebar(VARIANT_BOOL vbool);
    STDMETHODIMP get_Titlebar(VARIANT_BOOL* pbool);
    STDMETHODIMP put_Toolbars(VARIANT_BOOL vbool);
    STDMETHODIMP get_Toolbars(VARIANT_BOOL* pbool);
    STDMETHODIMP put_ModalState(VARIANT_BOOL vbool);
    STDMETHODIMP get_ModalState(VARIANT_BOOL* pbool);
    STDMETHODIMP ShowDialog(dsoShowDialogType DlgType);
    STDMETHODIMP put_EnableFileCommand(dsoFileCommandType Item, VARIANT_BOOL vbool);
    STDMETHODIMP get_EnableFileCommand(dsoFileCommandType Item, VARIANT_BOOL* pbool);
    STDMETHODIMP put_BorderStyle(dsoBorderStyle style);
    STDMETHODIMP get_BorderStyle(dsoBorderStyle* pstyle);
    STDMETHODIMP put_BorderColor(OLE_COLOR clr);
    STDMETHODIMP get_BorderColor(OLE_COLOR* pclr);
    STDMETHODIMP put_BackColor(OLE_COLOR clr);
    STDMETHODIMP get_BackColor(OLE_COLOR* pclr);
    STDMETHODIMP put_ForeColor(OLE_COLOR clr);
    STDMETHODIMP get_ForeColor(OLE_COLOR* pclr);
    STDMETHODIMP put_TitlebarColor(OLE_COLOR clr);
    STDMETHODIMP get_TitlebarColor(OLE_COLOR* pclr);
    STDMETHODIMP put_TitlebarTextColor(OLE_COLOR clr);
    STDMETHODIMP get_TitlebarTextColor(OLE_COLOR* pclr);
    STDMETHODIMP ExecOleCommand(LONG OLECMDID, VARIANT Options, VARIANT* vInParam, VARIANT* vInOutParam);
    STDMETHODIMP put_Menubar(VARIANT_BOOL vbool);
    STDMETHODIMP get_Menubar(VARIANT_BOOL* pbool);
    STDMETHODIMP put_HostName(BSTR bstr);
    STDMETHODIMP get_HostName(BSTR* pbstr);
    STDMETHODIMP get_DocumentFullName(BSTR* pbstr);
    STDMETHODIMP PrintOut(VARIANT PromptUser, VARIANT PrinterName, VARIANT Copies, VARIANT FromPage, VARIANT ToPage, VARIANT OutputFile);
    STDMETHODIMP PrintPreview();
    STDMETHODIMP PrintPreviewExit();
    STDMETHODIMP get_IsReadOnly(VARIANT_BOOL* pbool);
    STDMETHODIMP get_IsDirty(VARIANT_BOOL* pbool);
	STDMETHODIMP HttpInit(VARIANT_BOOL* pbool);
	STDMETHODIMP HttpAddPostString(BSTR strName, BSTR strValue, VARIANT_BOOL* pbool);
	STDMETHODIMP HttpPost(BSTR bstr,BSTR* pRet);
	STDMETHODIMP SetTrackRevisions(long vbool, VARIANT_BOOL* pbool);
	STDMETHODIMP SetCurrUserName(BSTR strCurrUserName, VARIANT_BOOL* pbool);
	STDMETHODIMP HttpAddPostCurrFile(BSTR strFileID, BSTR strFileName, VARIANT_BOOL * pbool);
	STDMETHODIMP SetCurrTime(BSTR strValue, VARIANT_BOOL* pbool);
 	STDMETHODIMP get_GetApplication(IDispatch** ppdisp);
 	STDMETHODIMP SetFieldValue(BSTR strFieldName,BSTR strValue,BSTR strCmdOrSheetName, VARIANT_BOOL* pbool);
 	STDMETHODIMP GetFieldValue(BSTR strFieldName, BSTR strCmdOrSheetName,BSTR* strValue);	
	STDMETHODIMP SetMenuDisplay(long lMenuFlag, VARIANT_BOOL* pbool);	
 	STDMETHODIMP ProtectDoc(long lProOrUn,long lProType,  BSTR strProPWD, VARIANT_BOOL * pbool);
	STDMETHODIMP ShowRevisions(long nNewValue, VARIANT_BOOL* pbool); 
	STDMETHODIMP InSertFile(BSTR strFieldPath, long lPos, VARIANT_BOOL* pbool);	
	STDMETHODIMP ClearFile();
 	STDMETHODIMP LoadOriginalFile( VARIANT strFieldPath,  VARIANT strFileType,  long* pbool);	
    STDMETHODIMP SaveAs( VARIANT strFileName,  VARIANT dwFileFormat,   long* pbool);	
	STDMETHODIMP DeleteLocalFile(BSTR strFilePath);	
	STDMETHODIMP GetTempFilePath(BSTR* strValue);	
	STDMETHODIMP ShowView(long dwViewType, long * pbool);	
	STDMETHODIMP FtpConnect(BSTR strURL,long lPort, BSTR strUser, BSTR strPwd, long * pbool);	
	STDMETHODIMP FtpGetFile(BSTR strRemoteFile, BSTR strLocalFile, long * pbool);	
    STDMETHODIMP FtpPutFile(BSTR strLocalFile, BSTR strRemoteFile, long blOverWrite, long * pbool);	
	STDMETHODIMP FtpDisConnect(long * pbool);	
 	STDMETHODIMP DownloadFile( BSTR strRemoteFile, BSTR strLocalFile, BSTR* strValue);	
	STDMETHODIMP HttpAddPostFile(BSTR strFileID,  BSTR strFileName,  long* pbool);
 	STDMETHODIMP GetRevCount(long * pbool);
 	STDMETHODIMP GetRevInfo(long lIndex, long lType, BSTR* pbool);
 	STDMETHODIMP SetValue(BSTR strValue, BSTR strName, long* pbool);
	STDMETHODIMP SetDocVariable(BSTR strVarName, BSTR strValue, long lOpt, long* pbool) ;
	STDMETHODIMP SetPageAs(BSTR strLocalFile, long lPageNum,  long lType, long* pbool);
 	STDMETHODIMP ReplaceText(BSTR strSearchText, BSTR strReplaceText, long lGradation, long* pbool);
    STDMETHODIMP put_FullScreen(VARIANT_BOOL vbool);
    STDMETHODIMP get_FullScreen(VARIANT_BOOL* pbool);
	STDMETHODIMP GetEnvironmentVariable(BSTR EnvironmentName,BSTR* strValue);
	STDMETHODIMP GetOfficeVersion(BSTR strName ,BSTR* strValue);
 // IInternalUnknown Implementation
    BEGIN_INTERFACE_PART(InternalUnknown, IUnknown)
    END_INTERFACE_PART(InternalUnknown)

 // IPersistStreamInit Implementation
    BEGIN_INTERFACE_PART(PersistStreamInit, IPersistStreamInit)
        STDMETHODIMP GetClassID(CLSID *pClassID);
        STDMETHODIMP IsDirty(void);
        STDMETHODIMP Load(LPSTREAM pStm);
        STDMETHODIMP Save(LPSTREAM pStm, BOOL fClearDirty);
        STDMETHODIMP GetSizeMax(ULARGE_INTEGER* pcbSize);
        STDMETHODIMP InitNew(void);
    END_INTERFACE_PART(PersistStreamInit)

 // IPersistPropertyBag Implementation
    BEGIN_INTERFACE_PART(PersistPropertyBag, IPersistPropertyBag)
        STDMETHODIMP GetClassID(CLSID *pClassID);
        STDMETHODIMP InitNew(void);
        STDMETHODIMP Load(IPropertyBag* pPropBag, IErrorLog* pErrorLog);
        STDMETHODIMP Save(IPropertyBag* pPropBag, BOOL fClearDirty, BOOL fSaveAllProperties);
    END_INTERFACE_PART(PersistPropertyBag)

 // IPersistStorage Implementation
    BEGIN_INTERFACE_PART(PersistStorage, IPersistStorage)
        STDMETHODIMP GetClassID(CLSID *pClassID);
        STDMETHODIMP IsDirty(void);
        STDMETHODIMP InitNew(LPSTORAGE pStg);
        STDMETHODIMP Load(LPSTORAGE pStg);
        STDMETHODIMP Save(LPSTORAGE pStg, BOOL fSameAsLoad);
        STDMETHODIMP SaveCompleted(LPSTORAGE pStg);
        STDMETHODIMP HandsOffStorage(void);
    END_INTERFACE_PART(PersistStorage)

 // IOleObject Implementation
    BEGIN_INTERFACE_PART(OleObject, IOleObject)
        STDMETHODIMP SetClientSite(IOleClientSite *pClientSite);
        STDMETHODIMP GetClientSite(IOleClientSite **ppClientSite);
        STDMETHODIMP SetHostNames(LPCOLESTR szContainerApp, LPCOLESTR szContainerObj);
        STDMETHODIMP Close(DWORD dwSaveOption);
        STDMETHODIMP SetMoniker(DWORD dwWhichMoniker, IMoniker *pmk);
        STDMETHODIMP GetMoniker(DWORD dwAssign, DWORD dwWhichMoniker, IMoniker **ppmk);
        STDMETHODIMP InitFromData(IDataObject *pDataObject, BOOL fCreation, DWORD dwReserved);
        STDMETHODIMP GetClipboardData(DWORD dwReserved, IDataObject **ppDataObject);
        STDMETHODIMP DoVerb(LONG iVerb, LPMSG lpmsg, IOleClientSite *pActiveSite, LONG lindex, HWND hwndParent, LPCRECT lprcPosRect);
        STDMETHODIMP EnumVerbs(IEnumOLEVERB **ppEnumOleVerb);
        STDMETHODIMP Update();
        STDMETHODIMP IsUpToDate();
        STDMETHODIMP GetUserClassID(CLSID *pClsid);
        STDMETHODIMP GetUserType(DWORD dwFormOfType, LPOLESTR *pszUserType);
        STDMETHODIMP SetExtent(DWORD dwDrawAspect, SIZEL *psizel);
        STDMETHODIMP GetExtent(DWORD dwDrawAspect, SIZEL *psizel);
        STDMETHODIMP Advise(IAdviseSink *pAdvSink, DWORD *pdwConnection);
        STDMETHODIMP Unadvise(DWORD dwConnection);
        STDMETHODIMP EnumAdvise(IEnumSTATDATA **ppenumAdvise);
        STDMETHODIMP GetMiscStatus(DWORD dwAspect, DWORD *pdwStatus);
        STDMETHODIMP SetColorScheme(LOGPALETTE *pLogpal);
    END_INTERFACE_PART(OleObject)
 
 // IOleControl Implementation
    BEGIN_INTERFACE_PART(OleControl, IOleControl)
        STDMETHODIMP GetControlInfo(CONTROLINFO* pCI);
        STDMETHODIMP OnMnemonic(LPMSG pMsg);
        STDMETHODIMP OnAmbientPropertyChange(DISPID dispID);
        STDMETHODIMP FreezeEvents(BOOL bFreeze);
    END_INTERFACE_PART(OleControl)

 // IOleInplaceObject Implementation 
    BEGIN_INTERFACE_PART(OleInplaceObject, IOleInPlaceObject)
        STDMETHODIMP GetWindow(HWND *phwnd);
        STDMETHODIMP ContextSensitiveHelp(BOOL fEnterMode);
        STDMETHODIMP InPlaceDeactivate();
        STDMETHODIMP UIDeactivate();
        STDMETHODIMP SetObjectRects(LPCRECT lprcPosRect, LPCRECT lprcClipRect);
        STDMETHODIMP ReactivateAndUndo();
    END_INTERFACE_PART(OleInplaceObject)

 // IOleInplaceActiveObject Implementation 
    BEGIN_INTERFACE_PART(OleInplaceActiveObject, IOleInPlaceActiveObject)
        STDMETHODIMP GetWindow(HWND *phwnd);
        STDMETHODIMP ContextSensitiveHelp(BOOL fEnterMode);
        STDMETHODIMP TranslateAccelerator(LPMSG lpmsg);
        STDMETHODIMP OnFrameWindowActivate(BOOL fActivate);
        STDMETHODIMP OnDocWindowActivate(BOOL fActivate);
        STDMETHODIMP ResizeBorder(LPCRECT prcBorder, IOleInPlaceUIWindow *pUIWindow, BOOL fFrameWindow);
        STDMETHODIMP EnableModeless(BOOL fEnable);
    END_INTERFACE_PART(OleInplaceActiveObject)
 
 // IViewObjectEx Implementation 
    BEGIN_INTERFACE_PART(ViewObjectEx, IViewObjectEx)
        STDMETHODIMP Draw(DWORD dwDrawAspect, LONG lIndex, void *pvAspect, DVTARGETDEVICE *ptd, HDC hicTargetDevice, HDC hdcDraw, LPCRECTL prcBounds, LPCRECTL prcWBounds, BOOL (__stdcall *pfnContinue)(DWORD dwContinue), DWORD dwContinue);
        STDMETHODIMP GetColorSet(DWORD dwAspect, LONG lindex, void* pvAspect, DVTARGETDEVICE *ptd, HDC hicTargetDev, LOGPALETTE** ppColorSet);
        STDMETHODIMP Freeze(DWORD dwAspect, LONG lindex, void* pvAspect, DWORD* pdwFreeze);
        STDMETHODIMP Unfreeze(DWORD dwFreeze);
        STDMETHODIMP SetAdvise(DWORD dwAspect, DWORD advf, IAdviseSink* pAdviseSink);
        STDMETHODIMP GetAdvise(DWORD* pdwAspect, DWORD* padvf, IAdviseSink** ppAdviseSink);
        STDMETHODIMP GetExtent(DWORD  dwDrawAspect, LONG lindex, DVTARGETDEVICE *ptd, LPSIZEL psizel);
        STDMETHODIMP GetRect(DWORD dwAspect, LPRECTL pRect);
        STDMETHODIMP GetViewStatus(DWORD* pdwStatus);
        STDMETHODIMP QueryHitPoint(DWORD dwAspect, LPCRECT pRectBounds, POINT ptlLoc, LONG lCloseHint, DWORD *pHitResult);
        STDMETHODIMP QueryHitRect(DWORD dwAspect, LPCRECT pRectBounds, LPCRECT pRectLoc, LONG lCloseHint, DWORD *pHitResult);
        STDMETHODIMP GetNaturalExtent(DWORD dwAspect, LONG lindex, DVTARGETDEVICE *ptd, HDC hicTargetDev, DVEXTENTINFO *pExtentInfo, LPSIZEL pSizel);
    END_INTERFACE_PART(ViewObjectEx)

 // IDataObject Implementation 
    BEGIN_INTERFACE_PART(DataObject, IDataObject)
        STDMETHODIMP GetData(FORMATETC *pfmtc,  STGMEDIUM *pstgm);
        STDMETHODIMP GetDataHere(FORMATETC *pfmtc, STGMEDIUM *pstgm);
        STDMETHODIMP QueryGetData(FORMATETC *pfmtc);
        STDMETHODIMP GetCanonicalFormatEtc(FORMATETC * pfmtcIn, FORMATETC * pfmtcOut);
        STDMETHODIMP SetData(FORMATETC *pfmtc, STGMEDIUM *pstgm, BOOL fRelease);
        STDMETHODIMP EnumFormatEtc(DWORD dwDirection, IEnumFORMATETC **ppenum);
        STDMETHODIMP DAdvise(FORMATETC *pfmtc, DWORD advf, IAdviseSink *psink, DWORD *pdwConnection);
        STDMETHODIMP DUnadvise(DWORD dwConnection);
        STDMETHODIMP EnumDAdvise(IEnumSTATDATA **ppenum);
    END_INTERFACE_PART(DataObject)

 // IProvideClassInfo Implementation
    BEGIN_INTERFACE_PART(ProvideClassInfo, IProvideClassInfo)
        STDMETHODIMP GetClassInfo(ITypeInfo** ppTI);
    END_INTERFACE_PART(ProvideClassInfo)

 // IConnectionPointContainer Implementation
    BEGIN_INTERFACE_PART(ConnectionPointContainer, IConnectionPointContainer)
        STDMETHODIMP EnumConnectionPoints(IEnumConnectionPoints **ppEnum);
        STDMETHODIMP FindConnectionPoint(REFIID riid, IConnectionPoint **ppCP);
    END_INTERFACE_PART(ConnectionPointContainer)

 // IEnumConnectionPoints Implementation
    BEGIN_INTERFACE_PART(EnumConnectionPoints, IEnumConnectionPoints)
        STDMETHODIMP Next(ULONG cConnections, IConnectionPoint **rgpcn, ULONG *pcFetched);
        STDMETHODIMP Skip(ULONG cConnections);
        STDMETHODIMP Reset(void);
        STDMETHODIMP Clone(IEnumConnectionPoints **ppEnum);
    END_INTERFACE_PART(EnumConnectionPoints)
 
 // IConnectionPoint Implementation
    BEGIN_INTERFACE_PART(ConnectionPoint, IConnectionPoint)
        STDMETHODIMP GetConnectionInterface(IID *pIID);
        STDMETHODIMP GetConnectionPointContainer(IConnectionPointContainer **ppCPC);
        STDMETHODIMP Advise(IUnknown *pUnk, DWORD *pdwCookie);
        STDMETHODIMP Unadvise(DWORD dwCookie);
        STDMETHODIMP EnumConnections(IEnumConnections **ppEnum);
    END_INTERFACE_PART(ConnectionPoint)

 // IOleCommandTarget  Implementation
    BEGIN_INTERFACE_PART(OleCommandTarget , IOleCommandTarget)
        STDMETHODIMP QueryStatus(const GUID *pguidCmdGroup, ULONG cCmds, OLECMD prgCmds[], OLECMDTEXT *pCmdText);
        STDMETHODIMP Exec(const GUID *pguidCmdGroup, DWORD nCmdID, DWORD nCmdexecopt, VARIANTARG *pvaIn, VARIANTARG *pvaOut);            
    END_INTERFACE_PART(OleCommandTarget)

 // ISupportErrorInfo Implementation
    BEGIN_INTERFACE_PART(SupportErrorInfo, ISupportErrorInfo)
        STDMETHODIMP InterfaceSupportsErrorInfo(REFIID riid);
    END_INTERFACE_PART(SupportErrorInfo)

 // IObjectSafety Implementation
    BEGIN_INTERFACE_PART(ObjectSafety, IObjectSafety)
        STDMETHODIMP GetInterfaceSafetyOptions(REFIID riid, DWORD *pdwSupportedOptions,DWORD *pdwEnabledOptions);
        STDMETHODIMP SetInterfaceSafetyOptions(REFIID riid, DWORD dwOptionSetMask, DWORD dwEnabledOptions);
    END_INTERFACE_PART(ObjectSafety)

    STDMETHODIMP           InitializeNewInstance();

    STDMETHODIMP           InPlaceActivate(LONG lVerb);
    STDMETHODIMP_(void)    SetInPlaceVisible(BOOL fShow);
    STDMETHODIMP_(void)    UpdateModalState(BOOL fModeless, BOOL fNotifyIPObject);
    STDMETHODIMP_(void)    UpdateActivationState(BOOL fActive);
    STDMETHODIMP_(void)    UpdateInteractiveState(BOOL fActive);

    STDMETHODIMP_(void)    OnDraw(DWORD dvAspect, HDC hdcDraw, LPRECT prcBounds, LPRECT prcWBounds, HDC hicTargetDev, BOOL fOptimize);
    STDMETHODIMP_(void)    OnFocusChange(BOOL fGotFocus, HWND hFocusWnd);
    STDMETHODIMP_(void)    OnDestroyWindow();
    STDMETHODIMP_(void)    OnResize();
    STDMETHODIMP_(void)    OnMouseMove(UINT x, UINT y);
    STDMETHODIMP_(void)    OnButtonDown(UINT x, UINT y);
    STDMETHODIMP_(void)    OnMenuMessage(UINT msg, WPARAM wParam, LPARAM lParam);
    STDMETHODIMP_(void)    OnToolbarAction(DWORD cmd);

    STDMETHODIMP_(void)    OnComponentActivationChange(BOOL fActivate);
    STDMETHODIMP_(void)    OnAppActivation(BOOL fActive, DWORD dwThreadID);
    STDMETHODIMP_(void)    OnPaletteChanged(HWND hwndPalChg);
    STDMETHODIMP_(void)    OnWindowEnable(BOOL fEnable){TRACE1("CDsoFramerControl::OnWindowEnable(%d)\n", fEnable);}

    STDMETHODIMP_(HMENU)   GetActivePopupMenu();
    STDMETHODIMP_(BOOL)    FRunningInDesignMode();
    STDMETHODIMP           DoDialogAction(dsoShowDialogType item);

    STDMETHODIMP           ProvideErrorInfo(HRESULT hres);

 // Some inline methods are provided for common tasks such as site notification
 // or calculation of draw size based on user selection of tools and border style.

    void __fastcall ViewChanged()
    {
        if (m_pDataAdviseHolder) // Send data change notification.
            m_pDataAdviseHolder->SendOnDataChange((IDataObject*)&m_xDataObject, NULL, 0);

        if (m_pViewAdviseSink) // Send the view change notification....
        {
            m_pViewAdviseSink->OnViewChange(DVASPECT_CONTENT, -1);
            if (m_fViewAdviseOnlyOnce) // If they asked to be advised once, kill the connection
                m_xViewObjectEx.SetAdvise(DVASPECT_CONTENT, 0, NULL);
        }
        InvalidateRect(m_hwnd, NULL, TRUE); // Ensure a full repaint...
    }

    void __fastcall GetSizeRectAfterBorder(LPRECT lprcx, LPRECT lprc)
    {
        if (lprcx) CopyRect(lprc, lprcx);
        else SetRect(lprc, 0, 0, m_Size.cx, m_Size.cy);
        if (m_fBorderStyle) InflateRect(lprc, -(4-m_fBorderStyle), -(4-m_fBorderStyle));
    }

    void __fastcall GetSizeRectAfterTitlebar(LPRECT lprcx, LPRECT lprc)
    {
        GetSizeRectAfterBorder(lprcx, lprc);
        if (m_fShowTitlebar) lprc->top += 21;
    }

    void __fastcall GetSizeRectForMenuBar(LPRECT lprcx, LPRECT lprc)
    {
        GetSizeRectAfterTitlebar(lprcx, lprc);
        lprc->bottom = lprc->top + 24;
    }

    void __fastcall GetSizeRectForDocument(LPRECT lprcx, LPRECT lprc)
    {
        GetSizeRectAfterTitlebar(lprcx, lprc);
        if (m_fShowMenuBar) lprc->top += 24; 
        if (lprc->top > lprc->bottom) lprc->top = lprc->bottom;
    }


    void __fastcall RedrawCaption()
    {
        RECT rcT;
        if ((m_hwnd) && (m_fShowTitlebar))
        {   GetClientRect(m_hwnd, &rcT); rcT.bottom = 21;
            InvalidateRect(m_hwnd, &rcT, FALSE);
        }
        if ((m_hwnd) && (m_fShowMenuBar))
        {   GetSizeRectForMenuBar(NULL, &rcT);
            InvalidateRect(m_hwnd, &rcT, FALSE);
        }
    }

 // The control window proceedure is handled through static class method.
    static STDMETHODIMP_(LRESULT) ControlWindowProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam);

 // The variables for the control are kept private but accessible to the
 // nested classes for each interface.
private:

    ULONG                   m_cRef;				   // Reference count
    IUnknown               *m_pOuterUnknown;       // Outer IUnknown (points to m_xInternalUnknown if not agg)
    ITypeInfo              *m_ptiDispType;         // ITypeInfo Pointer (IDispatch Impl)
    EXCEPINFO              *m_pDispExcep;          // EXCEPINFO Pointer (IDispatch Impl)

    HWND                    m_hwnd;                // our window
    HWND                    m_hwndParent;          // immediate parent window
    SIZEL                   m_Size;                // the size of this control  
    RECT                    m_rcLocation;          // where we at

    IOleClientSite         *m_pClientSite;         // active client site of host containter
    IOleControlSite        *m_pControlSite;        // control site
    IOleInPlaceSite        *m_pInPlaceSite;        // inplace site
    IOleInPlaceFrame       *m_pInPlaceFrame;       // inplace frame
    IOleInPlaceUIWindow    *m_pInPlaceUIWindow;    // inplace ui window

    IAdviseSink            *m_pViewAdviseSink;     // advise sink for view (only 1 allowed)
    IOleAdviseHolder       *m_pOleAdviseHolder;    // OLE advise holder (for oleobject sinks)
    IDataAdviseHolder      *m_pDataAdviseHolder;   // OLE data advise holder (for dataobject sink)
    IDispatch              *m_dispEvents;          // event sink (we only support 1 at a time)
    IStorage               *m_pOleStorage;         // IStorage for OLE hosts.

    CDsoDocObject          *m_pDocObjFrame;

    OLE_COLOR               m_clrBorderColor;      // Control Colors
    OLE_COLOR               m_clrBackColor;        // "
    OLE_COLOR               m_clrForeColor;        // "
    OLE_COLOR               m_clrTBarColor;        // "
    OLE_COLOR               m_clrTBarTextColor;    // "

    BSTR                    m_bstrCustomCaption;   // A custom caption (if provided)
    HMENU                   m_hmenuFilePopup;      // The File menu popup
    WORD                    m_wFileMenuFlags;      // Bitflags of enabled file menu items.
    WORD                    m_wSelMenuItem;        // Which item (if any) is selected
    WORD                    m_cMenuItems;          // Count of items on menu bar
    RECT                    m_rgrcMenuItems[DSO_MAX_MENUITEMS]; // Menu bar items

	class CDsoFrameWindowHook*  m_pFrameHook;
	HBITMAP                     m_hbmDeactive;
    LPWSTR                      m_pwszHostName;

    unsigned int        m_fDirty:1;                // does the control need to be resaved?
    unsigned int        m_fInPlaceVisible:1;       // we are in place visible or not?
    unsigned int        m_fUIActive:1;             // are we UI active or not.
    unsigned int        m_fViewAdvisePrimeFirst: 1;// for IViewobject2::setadvise
    unsigned int        m_fViewAdviseOnlyOnce: 1;  // for IViewobject2::setadvise
    unsigned int        m_fUsingWindowRgn:1;       // for SetObjectRects and clipping
    unsigned int        m_fFreezeEvents:1;         // should events be frozen?
    unsigned int        m_fDesignMode:1;           // are we in design mode?
    unsigned int        m_fModeFlagValid:1;        // has mode changed since last check?
    unsigned int        m_fBorderStyle:2;          // the border style
    unsigned int        m_fShowTitlebar:1;         // should we show titlebar?
    unsigned int        m_fShowToolbars:1;         // should we show toolbars?
    unsigned int        m_fModalState:1;           // are we modal?
    unsigned int        m_fObjectMenu:1;           // are we over obj menu item?
    unsigned int        m_fConCntDone:1;           // for enum connectpts
    unsigned int        m_fComponentActive:1;      // are we the active component?
    unsigned int        m_fShowMenuBar:1;          // should we show menubar?
    unsigned int        m_fInDocumentLoad:1;       // set when loading file
    unsigned int        m_fNoInteractive:1;        // set when we don't allow interaction with docobj
    unsigned int        m_fShowMenuPrev:1;         // were menus visible before loss of interactivity?
    unsigned int        m_fShowToolsPrev:1;        // were toolbars visible before loss of interactivity?
};


////////////////////////////////////////////////////////////////////
// CDsoFrameWindowHook -- Top-Level Frame Window Hook
//
//  Used by the control to allow for proper host notification of focus 
//  and activation events occurring at top-level window frame. Because 
//  this DocObject host is an OCX, we don't own these notifications and
//  have to "steal" them from our parent using a subclass.
//
//  The hook allows for more than one instance of the control by serializing 
//  activation notifications so only one instance of the control can be 
//  "UI active" at a time. This is required for proper ActiveX Document 
//  containment.
//
#define DSOF_MAX_CONTROLS      20
class CDsoFrameWindowHook
{
public:
	CDsoFrameWindowHook(){
		m_hwndTopLevelHost=NULL;
		m_pfnOrigWndProc=NULL;
		m_cControls=0;
		m_fHostUnicodeWindow=FALSE;
		m_idxActive = 0;
		memset(m_pControls, 0, sizeof(CDsoFrameWindowHook* ) * DSOF_MAX_CONTROLS);
	}
	~CDsoFrameWindowHook(){}

	static STDMETHODIMP_(CDsoFrameWindowHook*)
		AttachToFrameWindow(HWND hwndCtl, CDsoFramerControl* pocx);

	STDMETHODIMP Detach(CDsoFramerControl* pocx);
	STDMETHODIMP SetActiveComponent(CDsoFramerControl* pocx);

	inline STDMETHODIMP_(CDsoFramerControl*)
		GetActiveComponent(){return m_pControls[m_idxActive];}

    static STDMETHODIMP_(LRESULT) 
		HostWindowProcHook(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam);

protected:
	HWND                    m_hwndTopLevelHost;    // Top-level host window (hooked)
    WNDPROC                 m_pfnOrigWndProc;
	DWORD                   m_idxActive;
	DWORD                   m_cControls;
    CDsoFramerControl*      m_pControls[DSOF_MAX_CONTROLS];
	BOOL                    m_fHostUnicodeWindow;
};


#endif //DS_DSOFRAMER_H
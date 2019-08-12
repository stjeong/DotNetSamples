/***************************************************************************
 * DSOFDOCOBJ.H
 *
 * DSOFramer: OLE DocObject Site component (used by the control)
 *
 *  Copyright ?999-2004; Microsoft Corporation. All rights reserved.
 *  Written by Microsoft Developer Support Office Integration (PSS DSOI)
 * 
 *  This code is provided via KB 311765 as a sample. It is not a formal
 *  product and has not been tested with all containers or servers. Use it
 *  for educational purposes only. See the EULA.TXT file included in the
 *  KB download for full terms of use and restrictions.
 *
 *  THIS CODE IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, 
 *  EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED
 *  WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
 *
 ***************************************************************************/
#ifndef DS_DSOFDOCOBJ_H 
#define DS_DSOFDOCOBJ_H

////////////////////////////////////////////////////////////////////
// Declarations for Interfaces used in DocObject Containment
//
#include <docobj.h>    // Standard DocObjects (common to all AxDocs)
#include "ipprevw.h"   // PrintPreview (for select Office apps)
#include "rbbinder.h"  // Internet Publishing (for Web Folder write access) 


////////////////////////////////////////////////////////////////////
// CDsoDocObject -- ActiveDocument Container Site Object
//
//  The CDsoDocObject object handles all the DocObject embedding for the 
//  control and os fairly self-contained. Like the control it has its 
//  own window, but it merely acts as a parent for the embedded object
//  window(s) which it activates. 
//
//  CDsoDocObject works by taking a file (or automation object) and
//  copying out the OLE storage used for its persistent data. It then
//  creates a new embedding based on the data. If a storage is not
//  avaiable, it will attempt to oad the file directly, but the results 
//  are less predictable using this manner since DocObjects are embeddings
//  and not links and this component has limited support for links. As a
//  result, we will attempt to keep our own storage copy in most cases.
//
//  You should note that this approach is different than one taken by the
//  web browser control, which is basically a link container which will
//  try to embed (ip activate) if allowed, but if not it opens the file 
//  externally and keeps the link. If CDsoDocObject cannot embed the object,
//  it returns an error. It will not open the object external.
//  
//  Like the control, this object also uses nested classes for the OLE 
//  interfaces used in the embedding. They are easier to track and easier
//  to debug if a specific interface is over/under released. Again this was
//  a design decision to make the sample easier to break apart, but not required.
//
//  Because the object is not tied to the top-level window, it constructs
//  the OLE merged menu as a set of popup menus which the control then displays
//  in whatever form it wants. You would need to customize this if you used
//  the control in a host and wanted the menus to merge with the actual host
//  menu bar (on the top-level window or form).
// 
class CDsoFramerControl;
class CDsoDocObject : public IUnknown
{
public:
 	BOOL m_bNewCreate;//if the DocObject is Opened ..the FALSE ,otherwize TRUE;
     CDsoDocObject();
    ~CDsoDocObject();
	CDsoFramerControl * m_pParentCtrl;
 // IUnknown Implementation
    STDMETHODIMP QueryInterface(REFIID riid, void** ppv);
    STDMETHODIMP_(ULONG) AddRef(void);
    STDMETHODIMP_(ULONG) Release(void);

 // IOleClientSite Implementation
    BEGIN_INTERFACE_PART(OleClientSite, IOleClientSite)
        STDMETHODIMP SaveObject(void);
        STDMETHODIMP GetMoniker(DWORD dwAssign, DWORD dwWhichMoniker, IMoniker** ppmk);
        STDMETHODIMP GetContainer(IOleContainer** ppContainer);
        STDMETHODIMP ShowObject(void);
        STDMETHODIMP OnShowWindow(BOOL fShow);
        STDMETHODIMP RequestNewObjectLayout(void);
    END_INTERFACE_PART(OleClientSite)

 // IOleInPlaceSite Implementation
    BEGIN_INTERFACE_PART(OleInPlaceSite, IOleInPlaceSite)
        STDMETHODIMP GetWindow(HWND* phwnd);
        STDMETHODIMP ContextSensitiveHelp(BOOL fEnterMode);
        STDMETHODIMP CanInPlaceActivate(void);
        STDMETHODIMP OnInPlaceActivate(void);
        STDMETHODIMP OnUIActivate(void);
        STDMETHODIMP GetWindowContext(IOleInPlaceFrame** ppFrame, IOleInPlaceUIWindow** ppDoc, LPRECT lprcPosRect, LPRECT lprcClipRect, LPOLEINPLACEFRAMEINFO lpFrameInfo);
        STDMETHODIMP Scroll(SIZE sz);
        STDMETHODIMP OnUIDeactivate(BOOL fUndoable);
        STDMETHODIMP OnInPlaceDeactivate(void);
        STDMETHODIMP DiscardUndoState(void);
        STDMETHODIMP DeactivateAndUndo(void);
        STDMETHODIMP OnPosRectChange(LPCRECT lprcPosRect);
    END_INTERFACE_PART(OleInPlaceSite)

 // IOleDocumentSite Implementation
    BEGIN_INTERFACE_PART(OleDocumentSite, IOleDocumentSite)
        STDMETHODIMP ActivateMe(IOleDocumentView* pView);
    END_INTERFACE_PART(OleDocumentSite)

 // IOleInPlaceFrame Implementation
    BEGIN_INTERFACE_PART(OleInPlaceFrame, IOleInPlaceFrame)
        STDMETHODIMP GetWindow(HWND* phWnd);
        STDMETHODIMP ContextSensitiveHelp(BOOL fEnterMode);
        STDMETHODIMP GetBorder(LPRECT prcBorder);
        STDMETHODIMP RequestBorderSpace(LPCBORDERWIDTHS pBW);
        STDMETHODIMP SetBorderSpace(LPCBORDERWIDTHS pBW);
        STDMETHODIMP SetActiveObject(LPOLEINPLACEACTIVEOBJECT pIIPActiveObj, LPCOLESTR pszObj);
        STDMETHODIMP InsertMenus(HMENU hMenu, LPOLEMENUGROUPWIDTHS pMGW);
        STDMETHODIMP SetMenu(HMENU hMenu, HOLEMENU hOLEMenu, HWND hWndObj);
        STDMETHODIMP RemoveMenus(HMENU hMenu);
        STDMETHODIMP SetStatusText(LPCOLESTR pszText);
        STDMETHODIMP EnableModeless(BOOL fEnable);
        STDMETHODIMP TranslateAccelerator(LPMSG pMSG, WORD wID);
    END_INTERFACE_PART(OleInPlaceFrame)

 // IOleCommandTarget  Implementation
    BEGIN_INTERFACE_PART(OleCommandTarget , IOleCommandTarget)
        STDMETHODIMP QueryStatus(const GUID *pguidCmdGroup, ULONG cCmds, OLECMD prgCmds[], OLECMDTEXT *pCmdText);
        STDMETHODIMP Exec(const GUID *pguidCmdGroup, DWORD nCmdID, DWORD nCmdexecopt, VARIANTARG *pvaIn, VARIANTARG *pvaOut);            
    END_INTERFACE_PART(OleCommandTarget)

 // IAuthenticate  Implementation
    BEGIN_INTERFACE_PART(Authenticate , IAuthenticate)
        STDMETHODIMP Authenticate(HWND *phwnd, LPWSTR *pszUsername, LPWSTR *pszPassword);
    END_INTERFACE_PART(Authenticate)

 // IServiceProvider Implementation
    BEGIN_INTERFACE_PART(ServiceProvider , IServiceProvider)
        STDMETHODIMP QueryService(REFGUID guidService, REFIID riid, void **ppv);
    END_INTERFACE_PART(ServiceProvider)

 // IContinueCallback Implementation
    BEGIN_INTERFACE_PART(ContinueCallback , IContinueCallback)
        STDMETHODIMP FContinue(void);
        STDMETHODIMP FContinuePrinting(LONG cPagesPrinted, LONG nCurrentPage, LPOLESTR pwszPrintStatus);
    END_INTERFACE_PART(ContinueCallback)

 // IOlePreviewCallback Implementation
    BEGIN_INTERFACE_PART(PreviewCallback , IOlePreviewCallback)
        STDMETHODIMP Notify(DWORD wStatus, LONG nLastPage, LPOLESTR pwszPreviewStatus);
    END_INTERFACE_PART(PreviewCallback)

    STDMETHODIMP        InitializeNewInstance(HWND hwndCtl, LPRECT prcPlace, LPWSTR pwszHost, IOleCommandTarget* pcmdCtl);
    STDMETHODIMP        InitObjectStorage(REFCLSID rclsid, IStorage *pstg);

    STDMETHODIMP        CreateDocObject(REFCLSID rclsid);
    STDMETHODIMP        LoadStorageFromFile(LPWSTR pwszFile, REFCLSID rclsid, BOOL fReadOnly);
    STDMETHODIMP        LoadStorageFromURL(LPWSTR pwszURL, REFCLSID rclsid, BOOL fReadOnly, LPWSTR pwszUserName, LPWSTR pwszPassword);
    STDMETHODIMP        LoadFromAutomationObject(LPUNKNOWN punkObj, REFCLSID rclsid, BOOL fReadOnly);

    STDMETHODIMP        IPActivateView();
    STDMETHODIMP        IPDeactivateView();
    STDMETHODIMP        UIActivateView(BOOL fFocus);

    STDMETHODIMP        SaveDefault();
    STDMETHODIMP        SaveStorageToFile(LPWSTR pwszFile, BOOL fOverwriteFile);
    STDMETHODIMP        SaveStorageToURL(LPWSTR pwszURL, BOOL fOverwriteFile, LPWSTR pwszUserName, LPWSTR pwszPassword);

    STDMETHODIMP        DoOleCommand(DWORD dwOleCmdId, DWORD dwOptions, VARIANT* vInParam, VARIANT* vInOutParam);
    STDMETHODIMP_(void) Close();

    STDMETHODIMP        PrintDocument(LPCWSTR pwszPrinter, LPCWSTR pwszOutput, UINT cCopies, UINT nFrom, UINT nTo, BOOL fPromptUser);
    STDMETHODIMP        StartPrintPreview();
    STDMETHODIMP        ExitPrintPreview(BOOL fForceExit);

 // Control should notify us on these conditions (so we can pass to IP object)...
    STDMETHODIMP_(void) OnNotifySizeChange(LPRECT prc);
    STDMETHODIMP_(void) OnNotifyAppActivate(BOOL fActive, DWORD dwThreadID);
    STDMETHODIMP_(void) OnNotifyPaletteChanged(HWND hwndPalChg);
    STDMETHODIMP_(void) OnNotifyChangeToolState(BOOL fShowTools);
    STDMETHODIMP_(void) OnNotifyHostSetFocus();
   
    STDMETHODIMP        HrGetDataFromObject(VARIANT *pvtType, VARIANT *pvtOutput);
    STDMETHODIMP        HrSetDataInObject(VARIANT *pvtType, VARIANT *pvtInput, BOOL fMbcsString);

 // Inline accessors for control to get IP object info...
	LPDISPATCH GetIDispatch();
    inline IOleInPlaceActiveObject*  GetActiveObject(){return m_pipactive;}
    inline IOleObject*               GetOleObject(){return m_pole;}
    inline BOOL         IsReadOnly(){return m_fOpenReadOnly;}
	inline HWND         GetDocWindow(){return m_hwnd;}
    inline HWND         GetActiveWindow(){return m_hwndUIActiveObj;}
    inline HWND         GetMenuHWND(){return m_hwndMenuObj;}
    inline HMENU        GetActiveMenu(){return m_hMenuActive;}
	inline HMENU        GetMergedMenu(){return m_hMenuMerged;}
	inline void         SetMergedMenu(HMENU h){m_hMenuMerged = h;}
    inline LPCWSTR      GetSourceName(){return ((m_pwszWebResource) ? m_pwszWebResource : m_pwszSourceFile);}
    inline LPCWSTR      GetSourceDocName(){return ((m_pwszSourceFile) ? &m_pwszSourceFile[m_idxSourceName] : NULL);}
    inline BOOL         InPrintPreview(){return (m_pprtprv != NULL);}

	static STDMETHODIMP_(CDsoDocObject*) CreateNewDocObject(){return new CDsoDocObject();}
	STDMETHODIMP_(BOOL) IsStorageDirty();

protected:
 // Internal helper methods...
    STDMETHODIMP             CreateObjectStorage(REFCLSID rclsid);
    STDMETHODIMP             SaveObjectStorage();
    STDMETHODIMP             ValidateDocObjectServer(REFCLSID rclsid);
    STDMETHODIMP_(BOOL)      ValidateFileExtension(WCHAR* pwszFile, WCHAR** ppwszOut);

    STDMETHODIMP_(void)      OnDraw(DWORD dvAspect, HDC hdcDraw, LPRECT prcBounds, LPRECT prcWBounds, HDC hicTargetDev, BOOL fOptimize);

    STDMETHODIMP             EnsureOleServerRunning(BOOL fLockRunning);
    STDMETHODIMP_(void)      FreeRunningLock();
    STDMETHODIMP_(void)      TurnOffWebToolbar();
    STDMETHODIMP_(void)      ClearMergedMenu();
    STDMETHODIMP_(DWORD)     CalcDocNameIndex(LPCWSTR pwszPath);

 // These functions allow the component to access files in a Web Folder for 
 // write access using the MS Provider for Internet Publishing (MSDAIPP), which
 // is installed by Office and comes standard in Windows 2000/ME/XP/2003. The
 // provider is not required to use the component, only if you wish to save to 
 // an FPSE or DAV Web Folder (URL). 
    STDMETHODIMP_(IUnknown*) CreateRosebudIPP();
    STDMETHODIMP             DownloadWebResource(LPWSTR pwszURL, LPWSTR pwszFile, LPWSTR pwszUsername, LPWSTR pwszPassword, IStream** ppstmKeepForSave);
    STDMETHODIMP             UploadWebResource(LPWSTR pwszFile, IStream** ppstmSave, LPWSTR pwszURLSaveTo, BOOL fOverwriteFile, LPWSTR pwszUsername, LPWSTR pwszPassword);

    static STDMETHODIMP_(LRESULT)  FrameWindowProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam);

 // The private class variables...
private:
    ULONG                m_cRef;				// Reference count
    HWND                 m_hwnd;                // our window
    HWND                 m_hwndCtl;             // The control's window
    IOleCommandTarget   *m_pcmdCtl;             // IOCT of host (for frame msgs)
    RECT                 m_rcViewRect;          // Viewable area (set by host)

    LPWSTR               m_pwszSourceFile;      // Path to Source File (on Open)
    IStorage            *m_pstgSourceFile;      // Original File Storage (if open/save file)
    IMoniker            *m_pmkSourceObject;     // Moniker to original source object (file or new object)
    DWORD                m_idxSourceName;       // Index to doc name in m_pwszSourceFile

    IStorage            *m_pstgroot;            // Root temp storage
    IStorage            *m_pstgfile;            // In-memory file storage
    IStream             *m_pstmview;            // In-memory view info stream

    LPWSTR               m_pwszWebResource;     // The full URL to the web resource
    IStream             *m_pstmWebResource;     // Original Download Stream (if open/save URL)
    IUnknown            *m_punkRosebud;         // MSDAIPP pointer (for URL downloads)
    LPWSTR               m_pwszUsername;        // Username and password used by MSDAIPP
    LPWSTR               m_pwszPassword;        // for Authentication (see IAuthenticate)
    LPWSTR               m_pwszHostName;        // Ole Host Name for container


    IOleObject              *m_pole;            // Embedded OLE Object (OLE)
    IOleInPlaceObject       *m_pipobj;          // The IP object methods (OLE)
    IOleInPlaceActiveObject *m_pipactive;       // The UI Active object methods (OLE)
    IOleDocumentView        *m_pdocv;           // MSO Document View (DocObj)
    IOleCommandTarget       *m_pcmdt;           // MSO Command Target (DocObj)
    IOleInplacePrintPreview *m_pprtprv;         // MSO Print Preview (DocObj)

    HMENU                m_hMenuActive;         // The menu supplied by embedded object
    HMENU                m_hMenuMerged;         // The merged menu (set by control host)
    HOLEMENU             m_holeMenu;            // The OLE Menu Descriptor
    HWND                 m_hwndMenuObj;         // The window for menu commands
    HWND                 m_hwndIPObject;        // IP active object window
    HWND                 m_hwndUIActiveObj;     // UI Active object window
    DWORD                m_dwObjectThreadID;    // Thread Id of UI server
    BORDERWIDTHS         m_bwToolSpace;         // Toolspace...

 // Bitflags (state info)...
    unsigned int         m_fDisplayTools:1;
    unsigned int         m_fDisconnectOnQuit:1;
    unsigned int         m_fAppWindowActive:1;
    unsigned int         m_fOpenReadOnly:1;
    unsigned int         m_fObjectInModalCondition:1;
    unsigned int         m_fObjectIPActive:1;
    unsigned int         m_fObjectUIActive:1;
    unsigned int         m_fObjectActivateComplete:1;
	unsigned int         m_fLockedServerRunning:1;
	unsigned int         m_fLoadedFromAuto:1;
	unsigned int         m_fInClose:1;
public:
	CLSID                m_clsidObject;         // CLSID of the embedded object

};


#endif //DS_DSOFDOCOBJ_H
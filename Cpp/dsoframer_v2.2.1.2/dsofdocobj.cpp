/***************************************************************************
 * DSOFDOCOBJ.CPP
 *
 * CDsoDocObject: ActiveX Document Single Instance Frame/Site Object
 *
 *  Copyright ?999-2004; Microsoft Corporation. All rights reserved.
 *  Written by Microsoft Developer Support Office Integration (PSS DSOI)
 * 
 *  This code is provided via KB 311765 as a sample. It is not a formal
 *  product and has not been tested with all containers or servers. Use it
 *  for educational purposes only. See the EULA.TXT file  included in the
 *  KB download for full terms of use and restrictions.
 *
 *  THIS CODE IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, 
 *  EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED
 *  WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
 *
 ***************************************************************************/
#include "dsoframer.h"
#include "xmlhttpclient.h"

////////////////////////////////////////////////////////////////////////
// CDsoDocObject - The DocObject Site Class
//
//  This class wraps the functionality for DocObject hosting. Right now
//  we are setup for one active site at a time, but this could be changed
//  to allow multiple sites (although only one could be UI active at any
//  given time).
//
CDsoDocObject::CDsoDocObject()
{
	ODS("CDsoDocObject::CDsoDocObject\n");
	m_bNewCreate = TRUE;
	m_pParentCtrl = NULL;
	m_hwnd = NULL;           
	m_hwndCtl = NULL;        
	m_pcmdCtl = NULL;        
	m_rcViewRect.left = m_rcViewRect.top = m_rcViewRect.right = m_rcViewRect.bottom = 0;     

	m_pwszSourceFile = NULL; 
	m_pstgSourceFile = NULL; 
	m_pmkSourceObject = NULL;
	m_idxSourceName = 0;  

	m_pstgroot = NULL;       
	m_pstgfile = NULL;       
	m_pstmview = NULL;       

	m_pwszWebResource = NULL;
	m_pstmWebResource = NULL;
	m_punkRosebud = NULL;    
	m_pwszUsername = NULL;   
	m_pwszPassword = NULL;   
	m_pwszHostName = NULL;   

	memset(&m_clsidObject, 0, sizeof(CLSID));

	m_pole = NULL;       
	m_pipobj = NULL;     
	m_pipactive = NULL;  
	m_pdocv = NULL;      
	m_pcmdt = NULL;      
	m_pprtprv = NULL;    

	m_hMenuActive = NULL;    
	m_hMenuMerged = NULL;    
	m_holeMenu = NULL;       
	m_hwndMenuObj = NULL;    
	m_hwndIPObject = NULL;    
	m_hwndUIActiveObj = NULL; 
	m_dwObjectThreadID = 0;
	memset(&m_bwToolSpace, 0, sizeof(BORDERWIDTHS));

	m_fDisconnectOnQuit = 0;
	m_fAppWindowActive = 0;
	m_fOpenReadOnly = 0;
	m_fObjectInModalCondition = 0;
	m_fObjectIPActive = 0;
	m_fObjectUIActive = 0;
	m_fObjectActivateComplete = 0;
	m_fLockedServerRunning = 0;
	m_fLoadedFromAuto = 0;
	m_fInClose = 0;
	m_cRef = 1;
	m_fDisplayTools = TRUE;
}

CDsoDocObject::~CDsoDocObject(void)
{
	ODS("CDsoDocObject::~CDsoDocObject\n");
	if (m_pole)	Close();
	if (m_hwnd) DestroyWindow(m_hwnd);

	SAFE_FREESTRING(m_pwszUsername);
	SAFE_FREESTRING(m_pwszPassword);
	SAFE_FREESTRING(m_pwszHostName);

	SAFE_RELEASE_INTERFACE(m_punkRosebud);
	SAFE_RELEASE_INTERFACE(m_pcmdCtl);
    SAFE_RELEASE_INTERFACE(m_pstgroot);
}

////////////////////////////////////////////////////////////////////////
// CDsoDocObject::InitializeNewInstance
//
//  Sets up the docobject site. We must a control site to attach to and
//  a bounding rect. The IOleCommandTarget is used to forward toolbar
//  commands back to the host if a user selects one.
//
STDMETHODIMP CDsoDocObject::InitializeNewInstance(HWND hwndCtl, LPRECT prcPlace, LPWSTR pwszHost, IOleCommandTarget* pcmdCtl)
{
	HRESULT hr = E_UNEXPECTED;
	WNDCLASS wndclass;

    ODS("CDsoDocObject::InitializeNewInstance()\n");

 // As an AxDoc site, we need a valid parent window...
	if ((!hwndCtl) || (!IsWindow(hwndCtl)))
		return hr;

 // Create a temp storage for this docobj site (if one already exists, bomb out)...
	if ((m_pstgroot) || FAILED(hr = StgCreateDocfile(NULL,	STGM_TRANSACTED | STGM_READWRITE |
			STGM_SHARE_EXCLUSIVE | STGM_CREATE | STGM_DELETEONRELEASE, 0, &m_pstgroot)))
		return hr;

 // If our site window class has not been registered before, we should register it...

 // This is protected by a critical section just for fun. The fact we had to single
 // instance the OCX because of the host hook makes having multiple instances conflict here
 // very unlikely. However, that could change sometime, so better to be safe than sorry.
	EnterCriticalSection(&v_csecThreadSynch);

	if (GetClassInfo(v_hModule, "DSOFramerDocWnd", &wndclass) == 0)
	{
		memset(&wndclass, 0, sizeof(WNDCLASS));
		wndclass.style          = CS_VREDRAW | CS_HREDRAW | CS_DBLCLKS;
		wndclass.lpfnWndProc    = CDsoDocObject::FrameWindowProc;
		wndclass.hInstance      = v_hModule;
		wndclass.hCursor        = LoadCursor(NULL, IDC_ARROW);
		wndclass.lpszClassName  = "DSOFramerDocWnd";
		if (RegisterClass(&wndclass) == 0)
			hr = E_WIN32_LASTERROR;
	}

	LeaveCriticalSection(&v_csecThreadSynch);
	if (FAILED(hr)) return hr;
	
 // Save the place RECT (and validate as needed)...
	CopyRect(&m_rcViewRect, prcPlace);
	if (m_rcViewRect.top > m_rcViewRect.bottom)	{m_rcViewRect.top = 0; m_rcViewRect.bottom = 0;}
	if (m_rcViewRect.left > m_rcViewRect.right)	{m_rcViewRect.left = 0; m_rcViewRect.right = 0;}

 // Create our site window at the give location (we are child of the control window)...
	m_hwnd = CreateWindowEx(0, "DSOFramerDocWnd", NULL, WS_CHILD | WS_VISIBLE,
                    m_rcViewRect.left, m_rcViewRect.top,
					(m_rcViewRect.right - m_rcViewRect.left),
					(m_rcViewRect.bottom - m_rcViewRect.top),
                    hwndCtl, NULL, v_hModule, NULL);


	if (!m_hwnd) return E_OUTOFMEMORY;

	SetWindowLong(m_hwnd, GWL_USERDATA, (LONG)this);

	m_hwndCtl = hwndCtl;
    m_pwszHostName = DsoCopyString(((pwszHost) ? pwszHost : L"DsoFramerControl"));
	SAFE_SET_INTERFACE(m_pcmdCtl, pcmdCtl);
	return S_OK;
}

////////////////////////////////////////////////////////////////////////
// CDsoDocObject::InitObjectStorage
//
//  Creates and fills the 
//
STDMETHODIMP CDsoDocObject::InitObjectStorage(REFCLSID rclsid, IStorage *pstg)
{
    HRESULT hr;
    ODS("CDsoDocObject::InitObjectStorage()\n");
    CHECK_NULL_RETURN(pstg, E_POINTER);

 // Create a new storage for this CLSID...
	if (FAILED(hr = CreateObjectStorage(rclsid)))
		return hr;

 // Copy data into the new storage and commit the change...
	hr = pstg->CopyTo(0, NULL, NULL, m_pstgfile);
	if (SUCCEEDED(hr)) hr = m_pstgfile->Commit(STGC_OVERWRITE);

    return hr;
}

////////////////////////////////////////////////////////////////////////
// CDsoDocObject::CreateDocObject
//
//  This does the actual embedding. It is called no matter how you load
//  an object, and does some checking to make sure the CLSID is for a
//  a docobj server. If the function succeeds, we have an embedded object.
//  To activate and show the object, you must call IPActivateView().
//
STDMETHODIMP CDsoDocObject::CreateDocObject(REFCLSID rclsid)
{
	HRESULT             hr;
	BOOL                fInitNew;
	CLSID               clsid;
	DWORD               dwMiscStatus = 0;
    IOleObject*         pole    = NULL;
    IPersistStorage*    pipstg  = NULL;

    ODS("CDsoDocObject::CreateDocObject()\n");


	ASSERT(!(m_pole));
	IUnknown * pUnk=NULL;
	if(::GetActiveObject(rclsid,NULL,&pUnk)  == S_OK){
		ODS("CDsoDocObject: Object is aleady open\n");
		m_bNewCreate = FALSE;
	}else{
		ODS("CDsoDocObject: Object is new\n");
		m_bNewCreate = TRUE;
	}
 // Don't load if an object has already been loaded...
    if (m_pole) return E_UNEXPECTED;

 // First, check the server to make sure it is AxDoc server...
	if (FAILED(hr = ValidateDocObjectServer(rclsid)))
		return hr;

 // If we haven't loaded a storage, create a new one and remember to
 // call InitNew (instead of Load) later on...
	if ((fInitNew = (!m_pstgfile)) && FAILED(hr = CreateObjectStorage(rclsid)))
		return hr;

 // It is possible that someone picked an older ProgId/CLSID that
 // will AutoConvert on CoCreate, so fix up the storage with the
 // new CLSID info. We we actually call CoCreate on the new CLSID...
	if (fInitNew && SUCCEEDED(OleGetAutoConvert(rclsid, &clsid)))
	{
		OleDoAutoConvert(m_pstgfile, &clsid);
	}
	else clsid = rclsid;

 // We are ready to create an instance. Call CoCreate to make an
 // inproc handler and ask for IOleObject (all docobjs must support this)...
    if (FAILED(hr = CoCreateInstance(clsid, NULL, CLSCTX_INPROC, IID_IOleObject, (void**)&pole)))
		return hr;

 // Do a quick check to see if server wants us to set client site before the load...
	if (SUCCEEDED(hr = pole->GetMiscStatus(DVASPECT_CONTENT, &dwMiscStatus)) &&
		(dwMiscStatus & OLEMISC_SETCLIENTSITEFIRST))
		pole->SetClientSite((IOleClientSite*)&m_xOleClientSite);

 // Load up the bloody thing...
	if (SUCCEEDED(hr = pole->QueryInterface(IID_IPersistStorage, (void**)&pipstg)))
	{
     // Remember to InitNew if this is a new storage...			
		hr = ((fInitNew) ? pipstg->InitNew(m_pstgfile) : pipstg->Load(m_pstgfile));
		pipstg->Release();
	}

 // Assuming all the above worked we should have an OLE Embeddable
 // object and should finish the initialization (set object running)...
	if (SUCCEEDED(hr))
	{
	 // Save the IOleObject* and do a disconnect on quit...
		m_fDisconnectOnQuit = TRUE;
		m_pole = pole;

	 // Make sure the object is running (uses IRunnableObject)...
		hr = EnsureOleServerRunning(TRUE);

		if (SUCCEEDED(hr))
		{
		 // If we didn't do so already, set our client site...
			if (!(dwMiscStatus & OLEMISC_SETCLIENTSITEFIRST))
				m_pole->SetClientSite((IOleClientSite*)&m_xOleClientSite);

		 // Set the host names and then lock running...
			m_pole->SetHostNames(m_pwszHostName, m_pwszHostName);

		 // Keep server CLSID for this object
			m_clsidObject = clsid;

         // Ask object to save (if dirty)...
            if (IsStorageDirty()) SaveObjectStorage();
		}
	}

 // If we hit an error, cleanup (release should free the obj)...
	if (FAILED(hr))
	{
		if (dwMiscStatus & OLEMISC_SETCLIENTSITEFIRST)
			pole->SetClientSite(NULL);

		pole->Release();
		m_pole = NULL;
	}

    return hr;
}

////////////////////////////////////////////////////////////////////////
// CDsoDocObject::LoadStorageFromFile
//
//  Loads the internal IStorage from a file (local or UNC).
//
//  We handle three types of files: (1) BIFF files, which will just copy
//  their storage into our own; (2) non-BIFF files that are associated
//  with Office and can be loaded by IMoniker; and (3) non-BIFF files 
//  that are not associated with Office but may be loaded in Office using 
//  IPersistFile and then IPersistStorage (if an alternate CLSID is given), 
//
//  We do a special check for HTML/TXT files because they are associated
//  with IE, and we don't support IE files per se. Instead, you should pass
//  an alternate CLSID for the Office app you want to open that file in.
//
STDMETHODIMP CDsoDocObject::LoadStorageFromFile(LPWSTR pwszFile, REFCLSID rclsid, BOOL fReadOnly)
{
	HRESULT			hr;
	CLSID           clsid;
	CLSID           clsidConv;
	DWORD           dwBindFlgs;
	IStorage        *pstg    = NULL;
	IBindCtx		*pbctx   = NULL;
	IMoniker		*pmkfile = NULL;
	IPersistStorage *pipstg  = NULL;
	BOOL fLoadFromAltCLSID   = (rclsid != GUID_NULL);

	if (!(pwszFile) || ((*pwszFile) == L'\0'))
		return E_INVALIDARG;

	TRACE1("CDsoDocObject::LoadStorageFromFile(%S)\n", pwszFile);

 // First. we'll try to find the associated CLSID for the given file,
 // and then set it to the alternate if not found. If we don't have a
 // CLSID by the end of this, because user didn't specify alternate
 // and GetClassFile failed, then we error out...
    if (FAILED(GetClassFile(pwszFile, &clsid)) && !(fLoadFromAltCLSID))
		return DSO_E_INVALIDSERVER;

 // We should try to load from alternate CLSID if provided one...
    if (fLoadFromAltCLSID) clsid = rclsid;

 // We should also handle auto-convert to start "newest" server...
	if (SUCCEEDED(OleGetAutoConvert(clsid, &clsidConv)))
		clsid = clsidConv;

 // Validate that we have a DocObject server...
	if ((clsid == GUID_NULL) || FAILED(ValidateDocObjectServer(clsid)))
		return DSO_E_INVALIDSERVER;

 // We should have a CLSID, and can now create our substorage...
	if (FAILED(hr = CreateObjectStorage(clsid)))
		return hr;

 // Check for IE cache items since these are read-only as far as user is concerned...
    if (FIsIECacheFile(pwszFile)) fReadOnly = TRUE;

    dwBindFlgs = (STGM_TRANSACTED | STGM_SHARE_DENY_WRITE | (fReadOnly ? STGM_READ : STGM_READWRITE));

 // If these are native Office BiFF files, we can streamline the process to
 // copy the main storage into our sub-storage, without any involvement from
 // the server to fill our substorage. This will account for most Office loads...
	if (SUCCEEDED(hr = StgOpenStorage(pwszFile, NULL, dwBindFlgs, NULL, 0, &pstg)))
	{
		hr = pstg->CopyTo(0, NULL, NULL, m_pstgfile);
		if (SUCCEEDED(hr)) m_pstgfile->Commit(STGC_OVERWRITE);

 // Should that fail, we have to do things the more "formal" way (asking server
 // to properly save itself into our substorage)...
	}
	else if (fLoadFromAltCLSID)
	{
	 // If we are loading using an alternate CLSID that is either associated
     // with another app or is a non-Structured Storage doc not belonging to
     // to Office, then we have to explictily create an instance of the
     // alternate server (not the inproc handler), and load using IPersistFile,
     // then ask the server to save itself as OLE object using IPersistStorage.

	 // This is an expensive way to copy a storage, so we only do this when
	 // we are forced to use the alternate CLSID. This will typically be 
	 // the case for files like *.htm or *.asp that are not associated with
	 // Office, but can be opened in Office...
		IPersistFile *pipf;

		if (SUCCEEDED(hr = CoCreateInstance(clsid, NULL, CLSCTX_SERVER, 
			    IID_IPersistFile, (void**)&pipf)))
		{
			if (SUCCEEDED(hr = pipf->Load(pwszFile, dwBindFlgs)) && 
				SUCCEEDED(hr = pipf->QueryInterface(IID_IPersistStorage, (void**)&pipstg)))
			{
				if (SUCCEEDED(hr = pipstg->Save(m_pstgfile, FALSE)))
					hr = pipstg->SaveCompleted(NULL);
			
				pipstg->Release();
			}
			pipf->Release();
		}
	
	}
	else
	{
	 // Other non-BIFF files that are associated with Office (like *.rtf/*.csv)
	 // we can open based on a moniker. This is a more traditional (i.e., OLE
	 // "Insert From File") way of binding, which uses an inproc handler...
		if (SUCCEEDED(hr = CreateFileMoniker(pwszFile, &pmkfile)))
		{
			if (SUCCEEDED(hr = CreateBindCtx(0, &pbctx)))
			{
			 // We ask for IPersistStorage and do a formal Save to our IStorage...
				if (SUCCEEDED(hr = pmkfile->BindToObject(pbctx, NULL, IID_IPersistStorage, (void**)&pipstg))) 
				{
					if (SUCCEEDED(hr = pipstg->Save(m_pstgfile, FALSE)))
						hr = pipstg->SaveCompleted(NULL);
				
					pipstg->Release();
				}
				else if (hr == E_NOINTERFACE) // Somehow we got a server that doesn't handle OLE embedding...
					hr = STG_E_NOTFILEBASEDSTORAGE;

				if (SUCCEEDED(hr) && (!fReadOnly))
				{
				 // If we have our copy and user wants to keep a lock on original source,
				 // we'll try to get the lock by asking for the file's IStorage...
					BIND_OPTS bopts = {sizeof(BIND_OPTS), 1, dwBindFlgs, 0};
					pbctx->SetBindOptions(&bopts);
					pmkfile->BindToStorage(pbctx, NULL, IID_IStorage, (void**)&pstg);
				}
				pbctx->Release();
			}
			pmkfile->Release();
		}
	}

 // Assuming everything above worked...
	if (SUCCEEDED(hr))
	{ 
     // We are read-only until told otherwise...
	    m_fOpenReadOnly = TRUE;
		m_pwszSourceFile = DsoCopyString(pwszFile);
        m_idxSourceName = CalcDocNameIndex(m_pwszSourceFile);

	  // Let's go ahead and create the object, and keep the lock pointers...
		if (SUCCEEDED(hr = CreateDocObject(clsid)) && (fReadOnly == FALSE))
		{
			m_fOpenReadOnly = FALSE;
            SAFE_SET_INTERFACE(m_pstgSourceFile, pstg);
		}
	}

 // If we aren't locking, this will free the file...
	if (pstg) pstg->Release();

	return hr;
}

////////////////////////////////////////////////////////////////////////
// CDsoDocObject::LoadStorageFromURL
//
//  Loads the internal IStorage from a URL (http: or https:).
//
//  The idea here is we do a download from the URL and open the result
//  as any other file using LoadStorageFromFile. This requires MSDAIPP
//  which ships with MDAC 2.5, Office 2000/XP, or Windows 2000/XP.
//
STDMETHODIMP CDsoDocObject::LoadStorageFromURL(LPWSTR pwszURL, REFCLSID rclsid, BOOL fReadOnly, LPWSTR pwszUserName, LPWSTR pwszPassword)
{
	HRESULT	   hr;
	IStream   *pstmWebResource = NULL;
	LPWSTR     pwszTempFile;

 // We will need access to MSDAIPP unless it is read-only request...
	if (!(m_punkRosebud) && !(m_punkRosebud = CreateRosebudIPP()) && (!fReadOnly))
		return DSO_E_REQUIRESMSDAIPP;

 // Get the temp path for download...
	if (!GetTempPathForURLDownload(pwszURL, &pwszTempFile))
		return E_INVALIDARG;

 // Now get the resource from URLMON or MSDAIPP...
	if (SUCCEEDED(hr = DownloadWebResource(pwszURL, pwszTempFile,
        pwszUserName, pwszPassword, ((fReadOnly) ? NULL : &pstmWebResource))))
	{
     // If that worked, open from the downloaded resource...
		if (SUCCEEDED(hr = LoadStorageFromFile(pwszTempFile, rclsid, fReadOnly)) && 
			((pstmWebResource) && (!fReadOnly)))
		{
			m_fOpenReadOnly = FALSE;
			m_pwszWebResource = DsoCopyString(pwszURL);
			SAFE_SET_INTERFACE(m_pstmWebResource, pstmWebResource);
		}

		if (pstmWebResource)
			pstmWebResource->Release();
	}

	DsoMemFree(pwszTempFile);
	return hr;
}

////////////////////////////////////////////////////////////////////////
// CDsoDocObject::LoadFromAutomationObject
//
//  Copies the storage of the current automation object and loads a new
//  embedded object from the persisted state. This allows you to load 
//  a visible instance of a document previously open and edited by automation.
//
STDMETHODIMP CDsoDocObject::LoadFromAutomationObject(LPUNKNOWN punkObj, REFCLSID rclsid, BOOL fReadOnly)
{
	HRESULT hr;
	IPersistStorage *prststg;
	IOleObject *pole;
	CLSID clsid;

	ODS("CDsoDocObject::LoadFromAutomationObject()\n");
    CHECK_NULL_RETURN(punkObj, E_POINTER);

 // First, ensure the object passed supports embedding (i.e., we expect
 // something like a Word.Document or Excel.Sheet object, not a Word.Application
 // object or something else that is not a document type)...
	hr = punkObj->QueryInterface(IID_IOleObject, (void**)&pole);
	if (FAILED(hr)) return hr;

 // Ask the object to save its persistent state. We also collect its CLSID and
 // validate that the object type supports DocObject embedding...
	hr = pole->QueryInterface(IID_IPersistStorage, (void**)&prststg);
	if (SUCCEEDED(hr))
	{
		hr = prststg->GetClassID(&clsid); // Validate the object type...
		if (FAILED(hr) || ((clsid == GUID_NULL) && ((clsid = rclsid) == GUID_NULL)) || 
            FAILED(ValidateDocObjectServer(clsid)))
		{
			hr = DSO_E_INVALIDSERVER;
		}
		else
		{
         // Create a matching storage for the object and save the current
         // state of the file in our storage (this is our private copy)...
			if (SUCCEEDED(hr = CreateObjectStorage(clsid)) &&
				SUCCEEDED(hr = prststg->Save(m_pstgfile, FALSE)) &&
				SUCCEEDED(hr = prststg->SaveCompleted(NULL)))
			{
             // At this point we have a read-only copy...
                // prststg->HandsOffStorage();
                m_fOpenReadOnly = TRUE;

             // Create the embedding...
                hr = CreateDocObject(clsid);

             // If caller wants us to keep track of file for save (write access)
             // we can do so with moniker. Keep in mind this is does not keep a
             // lock on the resource and may fail the save if auto object is still
             // open and has exclusive access at time of the save, but it is the
             // best option we have when DocObject is loaded from an object already
             // open and locked by another application or component.
                if (SUCCEEDED(hr) && (!fReadOnly))
                {
                    IPersistMoniker *prstmk = NULL;
                    IMoniker *pmk = NULL;

                    if (SUCCEEDED(pole->GetMoniker(OLEGETMONIKER_FORCEASSIGN, OLEWHICHMK_OBJFULL, &pmk)) ||
                        (SUCCEEDED(pole->QueryInterface(IID_IPersistMoniker, (void**)&prstmk)) && 
                         SUCCEEDED(prstmk->GetCurMoniker(&pmk))))
                    {
                    // Keep hold of the source moniker...
                        m_fOpenReadOnly = FALSE;
                        m_pmkSourceObject = pmk;

#ifdef _DEBUG        // For debug trace, we output moniker name we will keep for write...
                        LPOLESTR postr = NULL; IBindCtx *pbc = NULL;
                        if (SUCCEEDED(CreateBindCtx(0, &pbc)))
                        {
                            if (SUCCEEDED(pmk->GetDisplayName(pbc, NULL, &postr)))
                            {
                                TRACE1(" >> Write Moniker = %S \n", postr);
                                CoTaskMemFree(postr);
                            }
                            pbc->Release();
                        }
#endif //_DEBUG
                    }

                    SAFE_RELEASE_INTERFACE(prstmk);
                }
			}
		}

		prststg->Release();
	}

	pole->Release();
	return hr;
}

////////////////////////////////////////////////////////////////////////
// CDsoDocObject::EnsureOleServerRunning
//
//  Verifies the DocObject server is running and is optionally locked
//  while the embedding is taking place.
//
STDMETHODIMP CDsoDocObject::EnsureOleServerRunning(BOOL fLockRunning)
{
	HRESULT hr = S_FALSE;
	IBindCtx *pbc;
	IRunnableObject *pro;
	IOleContainer *pocnt;
	BIND_OPTS bind;

    TRACE1("CDsoDocObject::EnsureOleServerRunning(%d)\n", (DWORD)fLockRunning);
	ASSERT(m_pole);

	if (FAILED(CreateBindCtx(0, &pbc)))
		return E_UNEXPECTED;

 // Setup the bind options for the run operation...
	bind.cbStruct = sizeof(BIND_OPTS);
	bind.grfFlags = BIND_MAYBOTHERUSER;
	bind.grfMode = STGM_READWRITE|STGM_SHARE_EXCLUSIVE;
	bind.dwTickCountDeadline = 20000;
	pbc->SetBindOptions(&bind);

 // Assume we aren't locked running...
	m_fLockedServerRunning = FALSE;

 // Get IRunnableObject and set server to run as OLE object. We check the
 // running state first since this is proper OLE, but note that this is not
 // returned from out-of-proc server, but the in-proc handler. Also note, we
 // specify a timeout in case the object never starts, and "check" the object
 // runs without IMessageFilter errors...
	if (SUCCEEDED(m_pole->QueryInterface(IID_IRunnableObject, (void**)&pro)))
	{
	 
	  // If the object is not currently running, let's run it...
		if (!(pro->IsRunning()))
			hr = pro->Run(pbc);

	  // Set the object server as a contained object (i.e., OLE object)...
		pro->SetContainedObject(TRUE);

	  // Lock running if desired...
		if (fLockRunning)
			m_fLockedServerRunning = SUCCEEDED(pro->LockRunning(TRUE, TRUE));

		pro->Release();
	}
	else if ((fLockRunning) &&
		SUCCEEDED(m_pole->QueryInterface(IID_IOleContainer, (void**)&pocnt)))
	{
		m_fLockedServerRunning = SUCCEEDED(pocnt->LockContainer(TRUE));
		pocnt->Release();
	}

	pbc->Release();
	return hr;
}

////////////////////////////////////////////////////////////////////////
// CDsoDocObject::FreeRunningLock
//
//  Free any previous lock made by EnsureOleServerRunning.
//
STDMETHODIMP_(void) CDsoDocObject::FreeRunningLock()
{
	IRunnableObject *pro;
	IOleContainer *pocnt;

    ODS("CDsoDocObject::FreeRunningLock(%d)\n");
	ASSERT(m_pole);

 // Don't do anything if we didn't lock the server...
	if (m_fLockedServerRunning == FALSE)
		return;

 // Get IRunnableObject and free lock...
	if (SUCCEEDED(m_pole->QueryInterface(IID_IRunnableObject, (void**)&pro)))
	{
		pro->LockRunning(FALSE, TRUE);
		pro->Release();
	}
	else if (SUCCEEDED(m_pole->QueryInterface(IID_IOleContainer, (void**)&pocnt)))
	{
		pocnt->LockContainer(FALSE);
		pocnt->Release();
	}

	m_fLockedServerRunning = FALSE;
	return;
}

////////////////////////////////////////////////////////////////////////
// CDsoDocObject::IPActivateView
//
//  Activates the object for viewing. If we already have an IOleDocumentView
//  we do this by calling Show, otherwise we'll do an IOleObject::DoVerb.
//
STDMETHODIMP CDsoDocObject::IPActivateView()
{
    HRESULT hr = E_UNEXPECTED;
    ODS("CDsoDocObject::IPActivateView()\n");
    ASSERT(m_pole);

 // Normal activation uses IOleObject::DoVerb with OLEIVERB_INPLACEACTIVATE
 // (or OLEIVERB_SHOW), a view rect, and an IOleClientSite pointer...
    if ((m_pole) && (!m_pdocv))
    {
		RECT rcView; GetClientRect(m_hwnd, &rcView);

		hr = m_pole->DoVerb(OLEIVERB_INPLACEACTIVATE, NULL, 
				(IOleClientSite*)&m_xOleClientSite, (UINT)-1, m_hwnd, &rcView);

	 // If the DoVerb fails on the InPlaceActivate verb, do a Show...
		if (FAILED(hr))
		{
			DWORD dwLoopCnt = 0;
			do
			{
				Sleep((200 * dwLoopCnt++));

				hr = m_pole->DoVerb(OLEIVERB_SHOW, NULL, 
						(IOleClientSite*)&m_xOleClientSite, (UINT)-1, m_hwnd, &rcView);

			  // There are issues with Visio 2002 rejecting DoVerb calls outright
			  // if it is still loading (instead of returning retry later as it should).
			  // So we might pop out of the message filter early and fail the call
			  // unexpectedly, so this little hack checks for the error and sleeps,
			  // then calls again in a recursive "loop"...
			}
			while ((hr == 0x80010001) && (dwLoopCnt < 10));

		}
    }
	else if (m_pdocv)
	{
	 // If we have an IOleDocument pointer, we can use the "new" method
	 // for inplace activation via Show/UIActivate...
		if (SUCCEEDED(hr = m_pdocv->Show(TRUE)))
			m_pdocv->UIActivate(TRUE);
	}

 // Forward focus as needed...
	if (SUCCEEDED(hr) && (m_hwndIPObject))
		SetFocus(m_hwndIPObject);

    return hr;
}

////////////////////////////////////////////////////////////////////////
// CDsoDocObject::IPDeactivateView
//
//  Deactivates the object.
//
STDMETHODIMP CDsoDocObject::IPDeactivateView()
{
    HRESULT hr = S_OK;
    ODS("CDsoDocObject::IPDeactivateView()\n");

 // If we still have a UI active object, tell it to UI deactivate...
	if (m_pipactive)
		UIActivateView(FALSE);

 // Next hide the active object...
	if (m_pdocv)
        m_pdocv->Show(FALSE);

 // Notify object our intention to IP deactivate...
    if (m_pipobj)
        m_pipobj->InPlaceDeactivate();

 // Close the object down and release pointers...
	if (m_pdocv)
	{
        hr = m_pdocv->CloseView(0);
        m_pdocv->SetInPlaceSite(NULL);
	}

    SAFE_RELEASE_INTERFACE(m_pcmdt);
    SAFE_RELEASE_INTERFACE(m_pdocv);
    SAFE_RELEASE_INTERFACE(m_pipobj);

    return hr;
}

////////////////////////////////////////////////////////////////////////
// CDsoDocObject::UIActivateView
//
//  UI Activates/Deactivates the object as needed.
//
STDMETHODIMP CDsoDocObject::UIActivateView(BOOL fFocus)
{
    HRESULT hr = S_FALSE;
    TRACE1("CDsoDocObject::UIActivateView(%d)\n", fFocus);

	if (m_pdocv)
        hr = m_pdocv->UIActivate(fFocus);

	return hr;
}

////////////////////////////////////////////////////////////////////////
// CDsoDocObject::SaveDefault
//
//  Saves the open object back to the original open location (unless it
//  was opened read-only, or is a new object). There are three types of
//  loaded doc objects we can save: (1) Files obtained by URL write bind
//  via MSDAIPP; (2) Files opened from local file source; and (3) Objects
//  already running or files linked to via OLE moniker obtained by an
//  Automation instance passed to us in Open. This function determines
//  which type we should do and call the right code for that type.
//
STDMETHODIMP CDsoDocObject::SaveDefault()
{
	HRESULT	hr = DSO_E_NOTBEENSAVED;

 // We can't do  save default if file was open read-only,
 // caller must do save with new file name...
    if (IsReadOnly()) return DSO_E_DOCUMENTREADONLY;

 // If we have a URL (write-access) resource, do MSDAIPP save...
	if (m_pstmWebResource)
	{
		hr = SaveStorageToURL(NULL, TRUE, NULL, NULL);
	}
 // Else if it is local file, do a local save...
	else if (m_pwszSourceFile)
	{
		hr = SaveStorageToFile(NULL, TRUE);
	}
 // If opened by moniker, try to get storage for object and save there...
    else if ((m_pmkSourceObject) && (m_pstgfile))
    {
        IBindCtx *pbc;
        IStorage *pstg;

     // First save the current state of the object in the internal
     // storage for the control and then make a bind context to access
     // the storage for the original source object (which may be file or other object)...
        if (SUCCEEDED(hr = SaveObjectStorage()) &&
            SUCCEEDED(hr = CreateBindCtx(NULL, &pbc)))
        {
            BIND_OPTS bndopt;
            bndopt.cbStruct = sizeof(BIND_OPTS);
            bndopt.grfFlags = BIND_MAYBOTHERUSER;
            bndopt.grfMode = (STGM_READWRITE | STGM_TRANSACTED | STGM_SHARE_DENY_WRITE);
            bndopt.dwTickCountDeadline = 6000;

            hr = pbc->SetBindOptions(&bndopt);
            ASSERT(SUCCEEDED(hr));

         // Ask moniker for the original storage and save our data to it (by copy).  
         // It is possible this can fail if orginal storage is file open with exclusive
         // access, or if moniker was for non-saved object that is no longer running... 
            hr = m_pmkSourceObject->BindToStorage(pbc, NULL, IID_IStorage, (void**)&pstg);
            if (SUCCEEDED(hr))
            {
                hr = m_pstgfile->CopyTo(0, NULL, NULL, pstg);
                if (SUCCEEDED(hr))
                {
                 // Commit the changes to the file if content is current...
                    hr = pstg->Commit(STGC_ONLYIFCURRENT);
                }
                pstg->Release();
            }

            pbc->Release();
        }

    }

	return hr;
}

////////////////////////////////////////////////////////////////////////
// CDsoDocObject::SaveStorageToFile
//
//  Saves the open object to a file. If you pass NULL for the file, we'll
//  save back to the original open location.
// 
extern  char* BSTR2char(const BSTR bstr) ;
STDMETHODIMP CDsoDocObject::SaveStorageToFile(LPWSTR pwszFile, BOOL fOverwriteFile)
{
 	HRESULT		  hr = E_UNEXPECTED;
	IPersistFile *pipfile;
	IStorage     *pstg;
	LPWSTR        pwszFullName = NULL;
	LPWSTR		  pwszRename = NULL;
	BOOL          fDoNormalSave = FALSE;
	BOOL          fDoOverwriteOps = FALSE;
	BOOL          fFileOpSuccess = FALSE;

 // Make sure we have the most current state for the file...
	if ((!m_pole) || FAILED(hr = SaveObjectStorage()))
		return hr;

 // If they passed no file, use the default if current file is not read-only...
	if ((!pwszFile) && 
        !((fDoNormalSave = !!(pwszFile = m_pwszSourceFile)) && (!IsReadOnly())))
		return DSO_E_DOCUMENTREADONLY;

 // Make sure a file extension is given (add one if not)...
	if (ValidateFileExtension(pwszFile, &pwszFullName))
		pwszFile = pwszFullName;

 // See if we will be overwriting, and error unless given permission to do so...
    if ((fDoOverwriteOps = FFileExists(pwszFile)) && !(fOverwriteFile))
        return STG_E_FILEALREADYEXISTS;

 // If we had a previous lock, we have to free it...
	SAFE_RELEASE_INTERFACE(m_pstgSourceFile);

 // If we are overwriting, we do a little Shell Operation here. This is done
 // for two reasons: (1) it keeps the server from asking us to overwrite the
 //  file as it normally would in case of normal save; and (2) it lets us
 // restore the original if the save fails...
	if (fDoOverwriteOps)
	{
		pwszRename = DsoCopyStringCat(pwszFile, L".dstmp");
		fFileOpSuccess = ((pwszRename) && FPerformShellOp(FO_RENAME, pwszFile, pwszRename));
	}

 // Let's do it. First ask for server to save to file if it supports IPersistFile.
 // This gives us a "real" file as output. This will work with almost all Office servers...
	if (SUCCEEDED(hr = m_pole->QueryInterface(IID_IPersistFile, (void**)&pipfile)))
	{
		
		DeleteFileW(pwszFile);
 		hr = pipfile->Save(pwszFile, FALSE);
 		   //Get the current open file 
		pipfile->Release(); 
		
/*
		LPWSTR pwszCurPath =  new WCHAR(2048);
 		memset(pwszCurPath,0,2048);
		if(S_OK == pipfile->GetCurFile(&pwszCurPath))  {
			char *p = BSTR2char(pwszCurPath);
	 
			hr = pipfile->Save(NULL, TRUE);
			hr = pipfile->SaveCompleted(pwszCurPath);
		} 
		else{
			hr = pipfile->Save(pwszCurPath, FALSE); 
			if (SUCCEEDED(hr)){ 
					//We should come here only in case of Word 2007. 
					DeleteFileW(pwszFile);
					hr = pipfile->Save(pwszCurPath, FALSE);  
					DeleteFileW(pwszCurPath);
			} 
		}
		//Save results S_OK, but the document is saved to the original location (IPersistFile->GetCurFile) in Word 2007. 
		//We will just check if that's the case and move the file. 
*/		
 /* */
	}
	else
	{
	 // If that doesn't work, save out the storage to OLE file. This may not produce
	 // the same type of file as you would get by the UI, but it would give a file
	 // that can be opened here again and in any "good" OLE server.
		if (SUCCEEDED(hr = StgCreateDocfile(pwszFile, STGM_TRANSACTED | 
				STGM_READWRITE | STGM_SHARE_EXCLUSIVE | STGM_CREATE, 0, &pstg)))
		{
			WriteClassStg(pstg, m_clsidObject);

			if (SUCCEEDED(hr = m_pstgfile->CopyTo(0, NULL, NULL, pstg)))
				hr = pstg->Commit(STGC_OVERWRITE);

			pstg->Release();
		}
	}

 // If we made a copy to protect on overwrite, either restore or delete it as needed...
	if ((fDoOverwriteOps) && (fFileOpSuccess) && (pwszRename))
	{
		FPerformShellOp((FAILED(hr) ? FO_RENAME : FO_DELETE), pwszRename, pwszFile);
	}

 // If this is an exisitng file save, or the operation failed, relock the
 // the original file save source...
    if (((fDoNormalSave) || (FAILED(hr))) && (m_pwszSourceFile))
    {
		StgOpenStorage(m_pwszSourceFile, NULL, 
			((m_pwszWebResource) ? 
				(STGM_READ | STGM_SHARE_DENY_WRITE) : 
				(STGM_TRANSACTED | STGM_SHARE_DENY_WRITE | STGM_READWRITE)), 
				NULL, 0, &m_pstgSourceFile);
    }
    else if (SUCCEEDED(hr))
	{
     // Otherwise if we succeeded, free any existing file info we have it and 
     // save the new file info for later re-saves (and lock)...
		SAFE_FREESTRING(m_pwszSourceFile);
		SAFE_FREESTRING(m_pwszWebResource);
		SAFE_RELEASE_INTERFACE(m_pstmWebResource);

	 // Save the name, and try to lock the file for editing...
		if (m_pwszSourceFile = DsoCopyString(pwszFile))
        {
            m_idxSourceName = CalcDocNameIndex(m_pwszSourceFile);
			StgOpenStorage(m_pwszSourceFile, NULL, 
				(STGM_TRANSACTED | STGM_SHARE_DENY_WRITE | STGM_READWRITE), NULL, 0, &m_pstgSourceFile);
        }
	}

	if (pwszRename)
		DsoMemFree(pwszRename);

	if (pwszFullName)
		DsoMemFree(pwszFullName);

	return hr;
}

////////////////////////////////////////////////////////////////////////
// CDsoDocObject::SaveStorageToURL
//
//  Saves the open object to a URL. If you pass NULL, we'll save back to
//  the original open location.
//
//  This works very similar to the LoadStorageFromURL in that we save to
//  a local file first using the normal SaveStorageToFile and then push an
//  upload to the server.
//
STDMETHODIMP CDsoDocObject::SaveStorageToURL(LPWSTR pwszURL, BOOL fOverwriteFile, LPWSTR pwszUserName, LPWSTR pwszPassword)
{
	HRESULT	 hr = DSO_E_DOCUMENTREADONLY;

 // If we have no URL to save to and no previously open web stream, fail...
	if ((!pwszURL) && (!m_pstmWebResource))
		return hr;

	if (!(m_punkRosebud) && !(m_punkRosebud = CreateRosebudIPP()))
		return DSO_E_REQUIRESMSDAIPP;

	if ((pwszURL) && (m_pwszWebResource) && 
		(DsoCompareStringsEx(pwszURL, -1, m_pwszWebResource, -1) == CSTR_EQUAL))
		pwszURL = NULL;

	if (pwszURL)
	{
		IStream  *pstmT = NULL;
		LPWSTR    pwszFullUrl = NULL;
		LPWSTR    pwszTempFile;

		IStream  *pstmBkupStm;
		IStorage *pstgBkupStg;
		LPWSTR    pwszBkupFile, pwszBkupUrl;
        UINT      idxBkup;

		if (!GetTempPathForURLDownload(pwszURL, &pwszTempFile))
			return E_INVALIDARG;

		if (ValidateFileExtension(pwszURL, &pwszFullUrl))
			pwszURL = pwszFullUrl;

	 // We are going to save out the current file info in case of 
	 // an error we can restore it to do native saves back to open location...
		pstmBkupStm  = m_pstmWebResource;  m_pstmWebResource  = NULL;
		pwszBkupUrl  = m_pwszWebResource;  m_pwszWebResource  = NULL;
		pwszBkupFile = m_pwszSourceFile;   m_pwszSourceFile   = NULL;
		pstgBkupStg  = m_pstgSourceFile;   m_pstgSourceFile   = NULL;
        idxBkup      = m_idxSourceName;    m_idxSourceName    = 0;

	 // Save the object to a new (temp) file on the local drive...
		if (SUCCEEDED(hr = SaveStorageToFile(pwszTempFile, TRUE)))
		{
		 // Then upload from that file...
			hr = UploadWebResource(pwszTempFile, &pstmT, pwszURL, fOverwriteFile, pwszUserName, pwszPassword);
		}

	 // If both calls succeed, we can free the old file/url location info
	 // and save the new information, otherwise restore the old info from backup...
		if (SUCCEEDED(hr))
		{
			SAFE_RELEASE_INTERFACE(pstgBkupStg);

			if ((pstmBkupStm) && (pwszBkupFile))
				FPerformShellOp(FO_DELETE, pwszBkupFile, NULL);

			SAFE_RELEASE_INTERFACE(pstmBkupStm);
			SAFE_FREESTRING(pwszBkupUrl);
			SAFE_FREESTRING(pwszBkupFile);

			m_pstmWebResource = pstmT;
			m_pwszWebResource = DsoCopyString(pwszURL);
			//m_pwszSourceFile already saved in SaveStorageToFile
			//m_pstgSourceFile already saved in SaveStorageToFile
            //m_idxSourceName already calced in SaveStorageToFile;

		}
		else
		{
			if (m_pstgSourceFile)
				m_pstgSourceFile->Release();

			if (m_pwszSourceFile)
			{
				FPerformShellOp(FO_DELETE, m_pwszSourceFile, NULL);
				DsoMemFree(m_pwszSourceFile);
			}

			m_pstmWebResource  = pstmBkupStm;
			m_pwszWebResource  = pwszBkupUrl;
			m_pwszSourceFile   = pwszBkupFile;
			m_pstgSourceFile   = pstgBkupStg;
            m_idxSourceName    = idxBkup;
		}

		if (pwszFullUrl)
			DsoMemFree(pwszFullUrl);

		DsoMemFree(pwszTempFile);

	}
	else if ((m_pstmWebResource) && (m_pwszSourceFile))
	{
        if (SUCCEEDED(hr = SaveStorageToFile(NULL, TRUE)))
		    hr = UploadWebResource(m_pwszSourceFile, &m_pstmWebResource,
                    NULL, TRUE, pwszUserName, pwszPassword);
	}

	return hr;
}


////////////////////////////////////////////////////////////////////////
// CDsoDocObject::DoOleCommand
//
//  Calls IOleCommandTarget::Exec on the active object to do a specific
//  command (like Print, SaveCopy, Zoom, etc.). 
//
STDMETHODIMP CDsoDocObject::DoOleCommand(DWORD dwOleCmdId, DWORD dwOptions, VARIANT* vInParam, VARIANT* vInOutParam)
{
	HRESULT hr;
	OLECMD cmd = {dwOleCmdId, 0};
	TRACE2("CDsoDocObject::DoOleCommand(cmd=%d, Opts=%d\n", dwOleCmdId, dwOptions);

 // Can't issue OLECOMMANDs when in print preview mode (object calls us)...
    if (InPrintPreview()) return E_ACCESSDENIED;

 // The server must support IOleCommandTarget, the CmdID being requested, and
 // the command should be enabled. If this is the case, do the command...
	if ((m_pcmdt) && SUCCEEDED(m_pcmdt->QueryStatus(NULL, 1, &cmd, NULL)) && 
		((cmd.cmdf & OLECMDF_SUPPORTED) && (cmd.cmdf & OLECMDF_ENABLED)))
	{
		TRACE1("QueryStatus say supported = 0x%X\n", cmd.cmdf);

     // Do the command asked by caller on default command group...
		hr = m_pcmdt->Exec(NULL, cmd.cmdID, dwOptions, vInParam, vInOutParam);
		TRACE1("CMT called = 0x%X\n", hr);

	 // If user canceled an Office dialog, that's OK.
		if ((dwOptions == OLECMDEXECOPT_PROMPTUSER) && (hr == 0x80040103))
			hr = S_FALSE;
	}
	else
	{
		TRACE1("Command Not supportted (%d)\n", cmd.cmdf);
		hr = DSO_E_COMMANDNOTSUPPORTED;
	}

	return hr;
}

////////////////////////////////////////////////////////////////////////
// CDsoDocObject::Close
//
//  Close down the object and disconnect us from any handlers/proxies.
//
STDMETHODIMP_(void) CDsoDocObject::Close()
{
	HRESULT hr;

	ODS("CDsoDocObject::Close\n");
    m_fInClose = TRUE;

 // Make sure we are not in print preview before close...
    if (InPrintPreview())
        ExitPrintPreview(TRUE);

 // Go ahead an IP deactivate the object...
	hr = IPDeactivateView();

 // Release the OLE object and cleanup...
    if (m_pole)
	{
        hr = m_pole->Close(OLECLOSE_NOSAVE);
		m_pole->SetClientSite(NULL);

		FreeRunningLock();
		SAFE_RELEASE_INTERFACE(m_pole);
	}

	SAFE_RELEASE_INTERFACE(m_pstgSourceFile);
    SAFE_RELEASE_INTERFACE(m_pmkSourceObject);

 // Free any temp file we might have...
	if ((m_pstmWebResource) && (m_pwszSourceFile))
		FPerformShellOp(FO_DELETE, m_pwszSourceFile, NULL);

	SAFE_RELEASE_INTERFACE(m_pstmWebResource);
	SAFE_FREESTRING(m_pwszWebResource);
	SAFE_FREESTRING(m_pwszSourceFile);
    m_idxSourceName = 0;

	if (m_fDisconnectOnQuit)
	{
		CoDisconnectObject((IUnknown*)this, 0);
		m_fDisconnectOnQuit = FALSE;
	}

	SAFE_RELEASE_INTERFACE(m_pstmview);
	SAFE_RELEASE_INTERFACE(m_pstgfile);

	ClearMergedMenu();
    m_fInClose = FALSE;
    return;
}


////////////////////////////////////////////////////////////////////////
// CDsoDocObject Notification Functions - The OCX should call these to
//  let the doc site update the object as needed.
//  

////////////////////////////////////////////////////////////////////////
// CDsoDocObject::OnNotifySizeChange
//
//  Resets the size of the site window and tells UI active object to 
//  resize as well. If we are UI active, we'll call ResizeBorder to 
//  re-negotiate toolspace (allow toolbars to shrink and grow), otherwise
//  we'll just set the IP active view rect (minus any toolspace, which
//  should be none since object is not UI active!). 
//
STDMETHODIMP_(void) CDsoDocObject::OnNotifySizeChange(LPRECT prc)
{
	RECT rc;

	SetRect(&rc, 0, 0, (prc->right - prc->left), (prc->bottom - prc->top));
	if (rc.right < 0) rc.right = 0;
	if (rc.top < 0) rc.top = 0;

 // First, resize our frame site window tot he new size (don't change focus)...
	if (m_hwnd)
	{
		m_rcViewRect = *prc;

		SetWindowPos(m_hwnd, NULL, m_rcViewRect.left, m_rcViewRect.top,
			rc.right, rc.bottom, SWP_NOACTIVATE | SWP_NOZORDER);

		UpdateWindow(m_hwnd);
	}

 // If we have an active object (i.e., Document is still UI active) we should
 // tell it of the resize so it can re-negotiate border space...
	if ((m_fObjectUIActive) && (m_pipactive))
	{
		m_pipactive->ResizeBorder(&rc, (IOleInPlaceUIWindow*)&m_xOleInPlaceFrame, TRUE);
	}
	else if ((m_fObjectIPActive) && (m_pdocv))
	{
        rc.left   += m_bwToolSpace.left;   rc.right  -= m_bwToolSpace.right;
        rc.top    += m_bwToolSpace.top;    rc.bottom -= m_bwToolSpace.bottom;
		m_pdocv->SetRect(&rc);
	}

	return;
}

////////////////////////////////////////////////////////////////////////
// CDsoDocObject::OnNotifyAppActivate
//
//  Notify doc object when the top-level frame window goes active and 
//  deactive so it can handle window focs and paiting correctly. Failure
//  to not forward this notification leads to bad behavior.
// 
STDMETHODIMP_(void) CDsoDocObject::OnNotifyAppActivate(BOOL fActive, DWORD dwThreadID)
{
 // This is critical for DocObject servers, so forward these messages
 // when the object is UI active...
	if (m_pipactive)
	{
	 // We should always tell obj server when our frame activates, but
	 // don't tell it to go deactive if the thread gaining focus is 
	 // the server's since our frame may have lost focus because of
	 // a top-level modeless dialog (ex., the RefEdit dialog of Excel)...
		//if ((fActive) || (dwThreadID != m_dwObjectThreadID))
        m_pipactive->OnFrameWindowActivate(fActive);
	}

	m_fAppWindowActive = fActive;
}

STDMETHODIMP_(void) CDsoDocObject::OnNotifyHostSetFocus()
{
    HWND hwnd;
	if ((m_pipactive) && SUCCEEDED(m_pipactive->GetWindow(&hwnd)))
        SetFocus(hwnd);
}

////////////////////////////////////////////////////////////////////////
// CDsoDocObject::OnNotifyPaletteChanged
//
//  Give the object first chance at realizing a palette. Important on
//  256 color machines, but not so critical these days when everyone is
//  running full 32-bit True Color graphic cards.
// 
STDMETHODIMP_(void) CDsoDocObject::OnNotifyPaletteChanged(HWND hwndPalChg)
{
	if ((m_fObjectUIActive) && (m_hwndUIActiveObj))
		SendMessage(m_hwndUIActiveObj, WM_PALETTECHANGED, (WPARAM)hwndPalChg, 0L);
}

////////////////////////////////////////////////////////////////////////
// CDsoDocObject::OnNotifyChangeToolState
//
//  This should be called to get object to show/hide toolbars as needed.
//
STDMETHODIMP_(void) CDsoDocObject::OnNotifyChangeToolState(BOOL fShowTools)
{
 // Can't change toolbar state in print preview (sorry)...
    if (InPrintPreview()) return;

 // If we want to show/hide toolbars, we can do the following...
	if (fShowTools != (BOOL)m_fDisplayTools)
	{
		OLECMD cmd;
		cmd.cmdID = OLECMDID_HIDETOOLBARS;

		m_fDisplayTools = fShowTools;

	 // Use IOleCommandTarget(OLECMDID_HIDETOOLBARS) to toggle on/off. We have
	 // to check that server supports it and if its state matches our own so
	 // when toggle, we do the correct thing by the user...
		if ((m_pcmdt) && SUCCEEDED(m_pcmdt->QueryStatus(NULL, 1, &cmd, NULL)) && 
			((cmd.cmdf & OLECMDF_SUPPORTED) || (cmd.cmdf & OLECMDF_ENABLED)))
		{
			if (((fShowTools) && (cmd.cmdf & OLECMDF_LATCHED)) ||
				(!(fShowTools) && !(cmd.cmdf & OLECMDF_LATCHED)))
			{
				m_pcmdt->Exec(NULL, OLECMDID_HIDETOOLBARS, OLECMDEXECOPT_PROMPTUSER, NULL, NULL);
			}

		 // There can be focus issues when turning them off, so make sure
		 // the object is on top of the z-order...
			if ((!m_fDisplayTools) && (m_hwndIPObject))
				BringWindowToTop(m_hwndIPObject);

		 // If user toggles off the toolbar while the object is UI active, and
		 // we are not still in activation process, we need to explictly tell Office
		 // apps to also hide the "Web" toolbar. For Office, OLECMDID_HIDETOOLBARS puts
		 // the app into a "web view" which (in some apps) brings up the web toolbar.
		 // Since we intend to have no tools, we have to turn it off by code...
			if ((!m_fDisplayTools) && (m_fObjectUIActive) && (m_fObjectActivateComplete))
				TurnOffWebToolbar();

		}
		else if (m_pdocv)
		{
		 // If we have a DocObj server, but no IOleCommandTarget, do things the hard
		 // way and resize. When server attempts to resize window it will have to
		 // re-negotiate BorderSpace and we fail there, so server "should" not
		 // display its tools (at least that is the idea!<g>)...
			RECT rc; GetClientRect(m_hwnd, &rc);
			MapWindowPoints(m_hwnd, m_hwndCtl, (LPPOINT)&rc, 2);
			OnNotifySizeChange(&rc);
		}
	}
	return;
}


////////////////////////////////////////////////////////////////////////
// CDsoDocObject Protected Functions -- Helpers
//

////////////////////////////////////////////////////////////////////////
// CDsoDocObject::CreateObjectStorage (protected)
//
//  Makes the internal IStorage to host the object and assigns the CLSID.
//
STDMETHODIMP CDsoDocObject::CreateObjectStorage(REFCLSID rclsid)
{
	HRESULT hr;
	LPWSTR pwszName;
	DWORD dwid;
	CHAR szbuf[256];

	if ((!m_pstgroot)) return E_UNEXPECTED;

 // Next, create a new object storage (with unique name) in our
 // temp root storage "file" (this keeps an OLE integrity some servers
 // need to function correctly instead of IP activating from file directly).

 // We make a fake object storage name...
	dwid = ((rclsid.Data1)|GetTickCount());
	wsprintf(szbuf, "OLEDocument%X", dwid);

	if (!(pwszName = DsoConvertToLPWSTR(szbuf)))
		return E_OUTOFMEMORY;

 // Create the sub-storage...
	hr = m_pstgroot->CreateStorage(pwszName,
		STGM_TRANSACTED | STGM_READWRITE | STGM_SHARE_EXCLUSIVE, 0, 0, &m_pstgfile);

	DsoMemFree(pwszName);

	if (FAILED(hr)) return hr;

 // We'll also create a stream for OLE view settings (non-critical)...
	if (pwszName = DsoConvertToLPWSTR(szbuf))
	{
		m_pstgroot->CreateStream(pwszName,
			STGM_DIRECT | STGM_READWRITE | STGM_SHARE_EXCLUSIVE, 0, 0, &m_pstmview);
		DsoMemFree(pwszName);
	}

 // Finally, write out the CLSID for the new substorage...
	hr = WriteClassStg(m_pstgfile, rclsid);
	return hr;
}

////////////////////////////////////////////////////////////////////////
// CDsoDocObject::SaveObjectStorage (protected)
//
//  Saves the object back to the internal IStorage. Returns S_FALSE if
//  there is no file storage for this document (it depends on how file
//  was loaded). In most cases we should have one, and this copies data
//  into the internal storage for save.
//
STDMETHODIMP CDsoDocObject::SaveObjectStorage()
{
	HRESULT hr = S_FALSE;
	IPersistStorage *pipstg = NULL;

 // Got to have object to save state...
	if (!m_pole) return E_UNEXPECTED;

 // If we have file storage, ask for IPersist and Save (commit changes)...
	if ((m_pstgfile) && 
        SUCCEEDED(hr = m_pole->QueryInterface(IID_IPersistStorage, (void**)&pipstg)))
	{
		if (SUCCEEDED(hr = pipstg->Save(m_pstgfile, TRUE)))
			hr = pipstg->SaveCompleted(NULL);

		hr = m_pstgfile->Commit(STGC_DEFAULT);
		pipstg->Release();
	}

 // Go ahead and save the view state if view still active (non-critical)...
	if ((m_pdocv) && (m_pstmview))
	{
		m_pdocv->SaveViewState(m_pstmview);
		m_pstmview->Commit(STGC_DEFAULT);
	}

	return hr;
}

////////////////////////////////////////////////////////////////////////
// CDsoDocObject::ValidateDocObjectServer (protected)
//
//  Quick validation check to see if CLSID is for DocObject server.
//
//  Officially, the only way to determine if a server supports ActiveX
//  Document embedding is to IP activate it and ask for IOleDocument, 
//  but that means going through the IP process just to fail if IOleDoc
//  is not supported. Therefore, we are going to rely on the server's 
//  honesty in setting its reg keys to include the "DocObject" sub key 
//  under their CLSID.
//
//  This is 99% accurate. For those servers that fail, too bad charlie!
//
STDMETHODIMP CDsoDocObject::ValidateDocObjectServer(REFCLSID rclsid)
{
	HRESULT hr = DSO_E_INVALIDSERVER;
	CHAR  szKeyCheck[256];
	LPSTR pszClsid;
	HKEY  hkey;

 // We don't handle MSHTML even though it is DocObject server...
    const GUID CLSID_MSHTMLDOC = {0x25336920,0x03F9,0x11CF,{0x8F,0xD0,0x00,0xAA,0x00,0x68,0x6F,0x13}};
    if (rclsid == CLSID_MSHTMLDOC) return hr;

	const GUID CLSID_PPTDOC = {0x64818D11,0x4F9B,0x11CF,{0x86,0xEA,0x00,0xAA,0x00,0xB9,0x29,0xE8}};
    if (rclsid == CLSID_PPTDOC) return S_OK;

 // Convert the CLSID to a string and check for DocObject sub key...
	if (pszClsid = DsoCLSIDtoLPSTR(rclsid))
	{
		wsprintf(szKeyCheck, "CLSID\\%s\\DocObject", pszClsid);
		long lRet = 0;
		lRet = RegOpenKeyEx(HKEY_CLASSES_ROOT, szKeyCheck, 0, KEY_READ, &hkey);
		if ( lRet == ERROR_SUCCESS)
		{
			hr = S_OK;
			RegCloseKey(hkey);
		}
		hr = S_OK;
		DsoMemFree(pszClsid);
	}
	else hr = E_OUTOFMEMORY;

	return hr;
}

////////////////////////////////////////////////////////////////////////
// CDsoDocObject::ValidateFileExtension (protected)
//
//  Adds a default extension to save file path if user didn't provide
//  one (uses CLSID and registry to determine default extension).
//
STDMETHODIMP_(BOOL) CDsoDocObject::ValidateFileExtension(WCHAR* pwszFile, WCHAR** ppwszOut)
{
	BOOL   fHasExt = FALSE;
	BOOL   fChangedExt = FALSE;
	LPWSTR pwszT;
	LPSTR  pszClsid;
	DWORD  dw;

	if ((pwszFile) && (dw = lstrlenW(pwszFile)) && (ppwszOut))
	{
		*ppwszOut = NULL;

		pwszT = (pwszFile + dw);
		while ((pwszT != pwszFile) && 
			   (*(--pwszT)) && ((*pwszT != L'\\') && (*pwszT != L'/')))
		{
			if (*pwszT == L'.') fHasExt = TRUE;
		}
		
		if (!(fHasExt) && (pszClsid = DsoCLSIDtoLPSTR(m_clsidObject)))
		{
			HKEY hk;
			DWORD dwType, dwSize;
			LPWSTR pwszExt;
			CHAR szkey[255];
			CHAR szbuf[128];

			wsprintf(szkey, "CLSID\\%s\\DefaultExtension", pszClsid);
			if (RegOpenKeyEx(HKEY_CLASSES_ROOT, szkey, 0, KEY_READ, &hk) == ERROR_SUCCESS)
			{
				LPSTR pszT = szbuf;
				dwSize = 128;

				if (RegQueryValueEx(hk, NULL, 0, &dwType, (BYTE*)pszT, &dwSize) == ERROR_SUCCESS)
				{
					while (*(pszT++) && (*pszT != ','))
						(void)(0);
					*pszT = '\0';
				}
				else lstrcpy(szbuf, ".ole");

				RegCloseKey(hk);
			}
			else lstrcpy(szbuf, ".ole");

			if (pwszExt = DsoConvertToLPWSTR(szbuf))
				*ppwszOut = DsoCopyStringCat(pwszFile, pwszExt);

			fChangedExt = ((*ppwszOut) != NULL);
			DsoMemFree(pszClsid);
		}
	}

	return fChangedExt;
}


////////////////////////////////////////////////////////////////////////
// CDsoDocObject::CreateRosebudIPP (protected)
//
//  Returns an instance of MSDAIPP (a.k.a., "Rosebud") which is used by
//  Office/Windows for Web Folders (HTTP with DAV/FPSE). 
//
//  This is used to open web resources, lock them for editing, and save
//  changes back up to the server as needed. If the provider is not 
//  installed or cannot be initialized, the functions returns NULL.
//
STDMETHODIMP_(IUnknown*) CDsoDocObject::CreateRosebudIPP()
{
	HRESULT           hr;
	IDBProperties*    pdbprops = NULL;
	IBindResource*    pres    = NULL;
	DBPROPSET         rdbpset;
	DBPROP            rdbp[4];
	BSTR              bstrLock;
	DWORD             dw = 256;
	CHAR              szUserName[256];

	if (FAILED(CoCreateInstance(CLSID_MSDAIPP_BINDER, NULL,
			CLSCTX_INPROC, IID_IDBProperties, (void**)&pdbprops)))
		return NULL;

	bstrLock = (GetUserName(szUserName, &dw) ? DsoConvertToBSTR(szUserName) : NULL);

	memset(rdbp, 0, sizeof(4 * sizeof(DBPROP)));

	rdbpset.cProperties = 4; 
	rdbpset.guidPropertySet = DBPROPSET_DBINIT;
	rdbpset.rgProperties = rdbp;

	rdbp[0].dwPropertyID = DBPROP_INIT_BINDFLAGS;
	rdbp[0].vValue.vt = VT_I4;
	rdbp[0].vValue.lVal = DBBINDURLFLAG_OUTPUT;

	rdbp[1].dwPropertyID = DBPROP_INIT_LOCKOWNER;
	rdbp[1].vValue.vt = VT_BSTR;
	rdbp[1].vValue.bstrVal = bstrLock;

	rdbp[2].dwPropertyID = DBPROP_INIT_LCID;
	rdbp[2].vValue.vt = VT_I4;
	rdbp[2].vValue.lVal = GetThreadLocale();

	rdbp[3].dwPropertyID = DBPROP_INIT_PROMPT;
	rdbp[3].vValue.vt = VT_I2;
	rdbp[3].vValue.iVal = DBPROMPT_COMPLETE;

	if (pdbprops->SetProperties(1, &rdbpset) == S_OK)
	{
		hr = pdbprops->QueryInterface(IIDX_IBindResource, (void**)&pres);
	}

    SAFE_FREEBSTR(bstrLock);
	pdbprops->Release();

	return (IUnknown*)pres;
}

////////////////////////////////////////////////////////////////////////
// CDsoDocObject::DownloadWebResource (protected)
//
//  Downloads the file specified by the URL to the given temp file. Locks
//  the web resource for editing if ppstmKeepForSave is requested.
//
STDMETHODIMP CDsoDocObject::DownloadWebResource(LPWSTR pwszURL, LPWSTR pwszFile, LPWSTR pwszUsername, LPWSTR pwszPassword, IStream** ppstmKeepForSave)
{
	HRESULT        hr    = E_UNEXPECTED;
	IStream       *pstm  = NULL;
	IBindResource *pres  = NULL;
	BYTE          *rgbBuf;
	DWORD   dwStatus, dwBindFlags;
	ULONG   cbRead, cbWritten;
	HANDLE  hFile;

 // First thing, we save out the user name and password (if provided)
 // for IAuthenticate which can be caled by either URLMON or MSDAIPP...
	if (pwszUsername)
	{
		SAFE_FREESTRING(m_pwszUsername);
		m_pwszUsername = DsoCopyString(pwszUsername);

		SAFE_FREESTRING(m_pwszPassword);
		m_pwszPassword = DsoCopyString(pwszPassword);
	}

 // If we don't need to write access, we can just use IE for "read-only" download...
    if (ppstmKeepForSave == NULL)
	{
        hr = URLDownloadFile((IUnknown*)&m_xAuthenticate, pwszURL, pwszFile);
		if (SUCCEEDED(hr) || (hr == E_ABORT)) return hr;
	  // If it fails for whatever reason, try MSDAIPP (which can use
	  // FPSE authentication as well as HTTP/DAV web access)...
	}

 // Otherwise, we want to use MSDAIPP for full access...
	if (!m_punkRosebud) return DSO_E_REQUIRESMSDAIPP;

	rgbBuf = new BYTE[10240]; //a 10-k buffer for reading
	if (!rgbBuf) return E_OUTOFMEMORY;

 // Use IBindResource::Bind to open an IStream and copy out the date to
 // the file given. This will then be used to load the object from the file...
	if (SUCCEEDED(m_punkRosebud->QueryInterface(IIDX_IBindResource, (void**)&pres)))
	{
		dwBindFlags = (DBBINDURLFLAG_READ | DBBINDURLFLAG_OUTPUT);
		if (ppstmKeepForSave)
			dwBindFlags |= (DBBINDURLFLAG_WRITE | DBBINDURLFLAG_SHARE_DENY_WRITE); //DBBINDURLFLAG_SHARE_EXCLUSIVE);
callagain:
		if (SUCCEEDED(hr = pres->Bind(NULL, pwszURL, dwBindFlags, DBGUIDX_STREAM, 
				IID_IStream, (IAuthenticate*)&m_xAuthenticate, NULL, &dwStatus, (IUnknown**)&pstm)))
		{
			LARGE_INTEGER lintStart; lintStart.QuadPart = 0;
			hr = pstm->Seek(lintStart, STREAM_SEEK_SET, NULL);

			if (FOpenLocalFile(pwszFile, GENERIC_WRITE, 0, CREATE_ALWAYS, &hFile))
			{
				while (SUCCEEDED(hr))
				{
					if (FAILED(hr = pstm->Read((void*)rgbBuf, 10240, &cbRead)) ||
						(cbRead == 0))
						break;

					if (FALSE == WriteFile(hFile, rgbBuf, cbRead, &cbWritten, NULL))
					{
						hr = E_WIN32_LASTERROR;
						break;
					}
				}

				CloseHandle(hFile);
			}
            else hr = E_WIN32_LASTERROR;

			if (ppstmKeepForSave)
                { SAFE_SET_INTERFACE(*ppstmKeepForSave, pstm); }

			pstm->Release();
		}
		else if ((hr == DB_E_NOTSUPPORTED) && ((dwBindFlags & DBBINDURLFLAG_OUTPUT) == DBBINDURLFLAG_OUTPUT))
		{
		 // WEC4 does not support DBBINDURLFLAG_OUTPUT flag, but if we are using WEC
		 // we don't really need the flag since this is not an HTTP GET call. Flip the
		 // flag off and call the method again to connect to server...
			dwBindFlags &= ~DBBINDURLFLAG_OUTPUT; 
			goto callagain;
		}

  		pres->Release();
	}

 // Map an OLEDB error to a common "file" error so a user
 // (and VB/VBScript) would better understand... 
	if (FAILED(hr))
	{
		switch (hr)
		{
		case DB_E_NOTFOUND:             hr = STG_E_FILENOTFOUND; break;
		case DB_E_READONLY:
		case DB_E_RESOURCELOCKED:       hr = STG_E_LOCKVIOLATION; break;
		case DB_SEC_E_PERMISSIONDENIED:
		case DB_SEC_E_SAFEMODE_DENIED:  hr = E_ACCESSDENIED; break;
		case DB_E_CANNOTCONNECT:
		case DB_E_TIMEOUT:              hr = E_VBA_NOREMOTESERVER; break;
		case E_NOINTERFACE:             hr = E_UNEXPECTED; break;
		}
	}

	delete [] rgbBuf;
	return hr;
}

////////////////////////////////////////////////////////////////////////
// CDsoDocObject::UploadWebResource (protected)
//
//  Uploads the file to a URL. The code can be used two ways:
//
//  1.) If ppstmSave contains a pointer to an existing IStream*, then we 
//      just upload to the existing stream. This allows for normal "Save"
//      on an open web resource.
//
//  2.) If ppstmSave is NULL (or contains a NULL IStream*), we create a new
//      web resource at the location given by pwszURLSaveTo and save to its
//      stream. If ppstmSave is passed, we return the new IStream* to the
//      caller who can then use it to do the other type of save next time.
//
STDMETHODIMP CDsoDocObject::UploadWebResource(LPWSTR pwszFile, IStream** ppstmSave, LPWSTR pwszURLSaveTo, BOOL fOverwriteFile, LPWSTR pwszUsername, LPWSTR pwszPassword)
{
	HRESULT     hr     = E_UNEXPECTED;
	ICreateRow *pcrow  = NULL;
	IStream    *pstm   = NULL;
	BYTE       *rgbBuf;
	HANDLE      hFile;
	BOOL        fstmIn = FALSE;
	ULONG       cbRead, cbWritten;
	DWORD       dwStatus, dwBindFlags;

	ASSERT(m_punkRosebud); // MSDAIPP is required for save...
	if (!m_punkRosebud) return DSO_E_REQUIRESMSDAIPP;

 // Check if this a "Save" on existing IStream* and jump to loop...
	if ((ppstmSave) && (fstmIn = (BOOL)(pstm = *ppstmSave)))
		goto uploadfrominstm;

 // Save the user name and password (if provided) for IAuthenticate...
	if (pwszUsername)
	{
		SAFE_FREESTRING(m_pwszUsername);
		m_pwszUsername = DsoCopyString(pwszUsername);

		SAFE_FREESTRING(m_pwszPassword);
		m_pwszPassword = DsoCopyString(pwszPassword);
	}

 // Check the URL string and ask for ICreateRow (to make new web resource)... 
	if (!(pwszURLSaveTo) || !LooksLikeHTTP(pwszURLSaveTo) ||
		FAILED(m_punkRosebud->QueryInterface(IIDX_ICreateRow, (void**)&pcrow)))
		return hr;

	dwBindFlags = ( DBBINDURLFLAG_READ | 
					DBBINDURLFLAG_WRITE | 
					DBBINDURLFLAG_SHARE_DENY_WRITE | 
                    (fOverwriteFile ? DBBINDURLFLAG_OVERWRITE : 0));

	if (SUCCEEDED(hr = pcrow->CreateRow(NULL, pwszURLSaveTo, dwBindFlags, DBGUIDX_STREAM,
			IID_IStream, (IAuthenticate*)&m_xAuthenticate, NULL, &dwStatus, NULL, (IUnknown**)&pstm)))
	{

	 // Once we are here, we have a stream (either handed in or opened from above).
	 // We just loop through and read from the file to the stream...
uploadfrominstm:
		if (rgbBuf = new BYTE[10240]) //a 10-k buffer for reading
		{
			if (FOpenLocalFile(pwszFile, GENERIC_READ, FILE_SHARE_READ, OPEN_EXISTING, &hFile))
			{
				LARGE_INTEGER lintStart; lintStart.QuadPart = 0;
				hr = pstm->Seek(lintStart, STREAM_SEEK_SET, NULL);

				while (SUCCEEDED(hr))
				{
					if (FALSE == ReadFile(hFile, rgbBuf, 10240, &cbRead, NULL))
					{
						hr = E_WIN32_LASTERROR;
						break;
					}

					if (0 == cbRead) break;

					if (FAILED(hr = pstm->Write((void*)rgbBuf, cbRead, &cbWritten)))
						break;
				}

			 // Need to commit the changes to make it official...
				if (SUCCEEDED(hr))
					hr = pstm->Commit(STGC_DEFAULT);

				CloseHandle(hFile);
			}
            else hr = E_WIN32_LASTERROR;

			delete [] rgbBuf;
		}
        else hr = E_OUTOFMEMORY;

	 // If we are not using a passed in IStream (and therefore created one), we
	 // should AddRef and pass back (if caller asked us to)...
		if (!fstmIn)
		{
			if (SUCCEEDED(hr) && (ppstmSave) && (!(*ppstmSave)))
			{
                SAFE_SET_INTERFACE(*ppstmSave, pstm);
			}
			pstm->Release();
		}

	}

 // Map an OLEDB error to a common "file" error so a user
 // (and VB/VBScript) would better understand... 
    if (FAILED(hr))
	{
		switch (hr)
		{
		case DB_E_RESOURCEEXISTS:       hr = STG_E_FILEALREADYEXISTS; break;
		case DB_E_NOTFOUND:             hr = STG_E_PATHNOTFOUND; break;
		case DB_E_READONLY:
		case DB_E_RESOURCELOCKED:       hr = STG_E_LOCKVIOLATION; break;
		case DB_SEC_E_PERMISSIONDENIED:
		case DB_SEC_E_SAFEMODE_DENIED:  hr = E_ACCESSDENIED; break;
		case DB_E_CANNOTCONNECT:
		case DB_E_TIMEOUT:              hr = E_VBA_NOREMOTESERVER; break;
		case DB_E_OUTOFSPACE:           hr = STG_E_MEDIUMFULL; break;
		case E_NOINTERFACE:             hr = E_UNEXPECTED; break;
		}
	}

	if (pcrow)
		pcrow->Release();

	return hr;
}

////////////////////////////////////////////////////////////////////////
// CDsoDocObject::TurnOffWebToolbar (protected)
//
//  This function "turns off" the Web toolbar used by Office apps to 
//  do in-site navigation. The problem is when toggling tools off the
//  bar may still be visible, so we have to explicitly turn it off if
//  we want a true "no tool" state.
//
STDMETHODIMP_(void) CDsoDocObject::TurnOffWebToolbar()
{
	IDispatch *pdisp;
	VARIANT    vtT[5];

 // Can't change toolbar state in print preview...
    if (InPrintPreview()) return;

	if ((m_pipactive) && 
		(SUCCEEDED(m_pipactive->QueryInterface(IID_IDispatch, (void**)&pdisp))))
	{
		if (SUCCEEDED(DsoDispatchInvoke(pdisp, 
				L"CommandBars", 0, DISPATCH_PROPERTYGET, 0, NULL, &vtT[0])))
		{
			ASSERT(vtT[0].vt == VT_DISPATCH);
			vtT[1].vt = VT_BSTR; vtT[1].bstrVal = SysAllocString(L"Web");

			if (SUCCEEDED(DsoDispatchInvoke(vtT[0].pdispVal, 
				L"Item", 0, DISPATCH_PROPERTYGET, 1, &vtT[1], &vtT[2])))
			{
				ASSERT(vtT[2].vt == VT_DISPATCH);
				vtT[3].vt = VT_BOOL; vtT[3].boolVal = 0;
				DsoDispatchInvoke(vtT[2].pdispVal, 
					L"Visible", 0, DISPATCH_PROPERTYPUT, 1, &vtT[3], &vtT[2]);
				VariantClear(&vtT[2]);
			}

			VariantClear(&vtT[1]);
			VariantClear(&vtT[0]);
		}

		pdisp->Release();
	}

}

////////////////////////////////////////////////////////////////////////
// CDsoDocObject::ClearMergedMenu (protected)
//
//  Frees the merged menu set by host.
//
STDMETHODIMP_(void) CDsoDocObject::ClearMergedMenu()
{	
	if (m_hMenuMerged)
	{
        int cbMenuCnt = GetMenuItemCount(m_hMenuMerged);
		for (int i = cbMenuCnt; i >= 0; --i)
		    RemoveMenu(m_hMenuMerged, i, MF_BYPOSITION);

		DestroyMenu(m_hMenuMerged);
		m_hMenuMerged = NULL;
	}
}

////////////////////////////////////////////////////////////////////////
// CDsoDocObject::CalcDocNameIndex (protected)
//
//  Calculates position of the name portion of the full path string.
//
STDMETHODIMP_(DWORD) CDsoDocObject::CalcDocNameIndex(LPCWSTR pwszPath)
{
    DWORD cblen, idx = 0;
    if ((pwszPath) && ((cblen = lstrlenW(pwszPath)) > 1))
    {
        for (idx = cblen; idx > 0; --idx)
        {
            if (pwszPath[idx] == L'\\')
                break;
        }

        if ((idx) && !(++idx < cblen))
            idx = 0;
    }
    return idx;
}

////////////////////////////////////////////////////////////////////////
// CDsoDocObject::IsStorageDirty
//
//  Ask the object if it is dirty and return result.
//
STDMETHODIMP_(BOOL) CDsoDocObject::IsStorageDirty()
{
	BOOL fDirty = TRUE; // Assume we are dirty unless object says we are not
	IPersistStorage *pprststg;
	IPersistFile *pprst;

	// Can't be dirty without object
	CHECK_NULL_RETURN(m_pole, FALSE);

	// Ask object its dirty state. Use IPersistStorage (it is more accurate for Word).
	if ((m_pstgfile) &&
		SUCCEEDED(m_pole->QueryInterface(IID_IPersistStorage, (void**)&pprststg)))
	{
		fDirty = ((pprststg->IsDirty() == S_FALSE) ? FALSE : TRUE);
		pprststg->Release();
	}
	else if (SUCCEEDED(m_pole->QueryInterface(IID_IPersistFile, (void**)&pprst)))
	{
		fDirty = ((pprst->IsDirty() == S_FALSE) ? FALSE : TRUE);
		pprst->Release();
	}

	return fDirty;
}


////////////////////////////////////////////////////////////////////////
// CDsoDocObject::OnDraw (protected)
//
//  Site drawing (does nothing in this version).
//
STDMETHODIMP_(void) CDsoDocObject::OnDraw(DWORD dvAspect, HDC hdcDraw, LPRECT prcBounds, LPRECT prcWBounds, HDC hicTargetDev, BOOL fOptimize)
{
	// Don't have to draw anything, object does all this because we are
	// always UI active. If we allowed for multiple objects and had some
	// non-UI active, we would have to do some drawing, but that will not
	// happen in this sample.
}


////////////////////////////////////////////////////////////////////////
//
// ActiveX Document Site Interfaces
//

////////////////////////////////////////////////////////////////////////
//
// CDsoDocObject IUnknown Interface Methods
//
//   STDMETHODIMP         QueryInterface(REFIID riid, void ** ppv);
//   STDMETHODIMP_(ULONG) AddRef(void);
//   STDMETHODIMP_(ULONG) Release(void);
//
STDMETHODIMP CDsoDocObject::QueryInterface(REFIID riid, void** ppv)
{
	ODS("CDsoDocObject::QueryInterface\n");
	CHECK_NULL_RETURN(ppv, E_POINTER);
	
	HRESULT hr = S_OK;

	if (IID_IUnknown == riid)
	{
		*ppv = (IUnknown*)this;
	}
	else if (IID_IOleClientSite == riid)
	{
		*ppv = (IOleClientSite*)&m_xOleClientSite;
	}
	else if ((IID_IOleInPlaceSite == riid) || (IID_IOleWindow == riid))
	{
		*ppv = (IOleInPlaceSite*)&m_xOleInPlaceSite;
	}
	else if (IID_IOleDocumentSite == riid)
	{
		*ppv = (IOleDocumentSite*)&m_xOleDocumentSite;
	}
	else if ((IID_IOleInPlaceFrame == riid) || (IID_IOleInPlaceUIWindow == riid))
	{
		*ppv = (IOleInPlaceFrame*)&m_xOleInPlaceFrame;
	}
	else if (IID_IOleCommandTarget == riid)
	{
		*ppv = (IOleCommandTarget*)&m_xOleCommandTarget;
	}
	else if (IIDX_IAuthenticate == riid)
	{
		*ppv = (IAuthenticate*)&m_xAuthenticate;
	}
    else if (IID_IServiceProvider == riid)
    {
        *ppv = (IServiceProvider*)&m_xServiceProvider;
    }
    else if (IID_IContinueCallback == riid)
    {
        *ppv = (IContinueCallback*)&m_xContinueCallback;
    }
    else if (IID_IOlePreviewCallback == riid)
    {
        *ppv = (IOlePreviewCallback*)&m_xPreviewCallback;
    }
	else
	{
		*ppv = NULL;
		hr = E_NOINTERFACE;
	}

	if (NULL != *ppv)
		((IUnknown*)(*ppv))->AddRef();
	return hr;
}

STDMETHODIMP_(ULONG) CDsoDocObject::AddRef(void)
{
	TRACE1("CDsoDocObject::AddRef - %d\n", m_cRef + 1);
    return ++m_cRef;
}

STDMETHODIMP_(ULONG) CDsoDocObject::Release(void)
{
	TRACE1("CDsoDocObject::Release - %d\n", m_cRef - 1);
	return --m_cRef;
}


////////////////////////////////////////////////////////////////////////
//
// CDsoDocObject::XOleClientSite -- IOleClientSite Implementation
//
//	 STDMETHODIMP SaveObject(void);
//	 STDMETHODIMP GetMoniker(DWORD dwAssign, DWORD dwWhich, LPMONIKER* ppmk);
//	 STDMETHODIMP GetContainer(LPOLECONTAINER* ppContainer);
//	 STDMETHODIMP ShowObject(void);
//	 STDMETHODIMP OnShowWindow(BOOL fShow);
//	 STDMETHODIMP RequestNewObjectLayout(void);
//
IMPLEMENT_INTERFACE_UNKNOWN(CDsoDocObject, OleClientSite)

STDMETHODIMP CDsoDocObject::XOleClientSite::SaveObject(void)
{
    ODS("CDsoDocObject::XOleClientSite::SaveObject\n");
	return S_OK;
}

STDMETHODIMP CDsoDocObject::XOleClientSite::GetMoniker(DWORD dwAssign, DWORD dwWhichMoniker, IMoniker** ppmk)
{
    ODS("CDsoDocObject::XOleClientSite::GetMoniker\n");
	if (ppmk) *ppmk = NULL;
	return E_NOTIMPL;
}

STDMETHODIMP CDsoDocObject::XOleClientSite::GetContainer(IOleContainer** ppContainer)
{
    ODS("CDsoDocObject::XOleClientSite::GetContainer\n");
	if (ppContainer) *ppContainer = NULL;
	return E_NOINTERFACE;
}

STDMETHODIMP CDsoDocObject::XOleClientSite::ShowObject(void)
{
    ODS("CDsoDocObject::XOleClientSite::ShowObject\n");
	return S_OK;
}

STDMETHODIMP CDsoDocObject::XOleClientSite::OnShowWindow(BOOL fShow)
{
    ODS("CDsoDocObject::XOleClientSite::OnShowWindow\n");
	return S_OK;
}

STDMETHODIMP CDsoDocObject::XOleClientSite::RequestNewObjectLayout(void)
{
    ODS("CDsoDocObject::XOleClientSite::RequestNewObjectLayout\n");
	return E_NOTIMPL;
}

////////////////////////////////////////////////////////////////////////
//
// CDsoDocObject::XOleInPlaceSite -- IOleInPlaceSite Implementation
//
//	 STDMETHODIMP GetWindow(HWND* phWnd);
//	 STDMETHODIMP ContextSensitiveHelp(BOOL fEnterMode);
//	 STDMETHODIMP CanInPlaceActivate(void);
//	 STDMETHODIMP OnInPlaceActivate(void);
//	 STDMETHODIMP OnUIActivate(void);
//	 STDMETHODIMP GetWindowContext(LPOLEINPLACEFRAME* ppIIPFrame, LPOLEINPLACEUIWINDOW* ppIIPUIWindow, LPRECT prcPos, LPRECT prcClip, LPOLEINPLACEFRAMEINFO pFI);
//	 STDMETHODIMP Scroll(SIZE sz);
//	 STDMETHODIMP OnUIDeactivate(BOOL fUndoable);
//	 STDMETHODIMP OnInPlaceDeactivate(void);
//	 STDMETHODIMP DiscardUndoState(void);
//	 STDMETHODIMP DeactivateAndUndo(void);
//	 STDMETHODIMP OnPosRectChange(LPCRECT prcPos);
//
IMPLEMENT_INTERFACE_UNKNOWN(CDsoDocObject, OleInPlaceSite)

STDMETHODIMP CDsoDocObject::XOleInPlaceSite::GetWindow(HWND* phwnd)
{
	METHOD_PROLOGUE(CDsoDocObject, OleInPlaceSite);
    ODS("CDsoDocObject::XOleInPlaceSite::GetWindow\n");
	if (phwnd) *phwnd = pThis->m_hwnd;
	return S_OK;
}

STDMETHODIMP CDsoDocObject::XOleInPlaceSite::ContextSensitiveHelp(BOOL fEnterMode)
{
    ODS("CDsoDocObject::XOleInPlaceSite::ContextSensitiveHelp\n");
	return S_OK;
}

STDMETHODIMP CDsoDocObject::XOleInPlaceSite::CanInPlaceActivate(void)
{
    ODS("CDsoDocObject::XOleInPlaceSite::CanInPlaceActivate\n");
	return S_OK;
}

STDMETHODIMP CDsoDocObject::XOleInPlaceSite::OnInPlaceActivate(void)
{
	METHOD_PROLOGUE(CDsoDocObject, OleInPlaceSite);
    ODS("CDsoDocObject::XOleInPlaceSite::OnInPlaceActivate\n");

	if ((!pThis->m_pole) || 
		FAILED(pThis->m_pole->QueryInterface(IID_IOleInPlaceObject, (void **)&(pThis->m_pipobj))))
		return E_UNEXPECTED;

    pThis->m_fObjectIPActive = TRUE;
    return S_OK;
}

STDMETHODIMP CDsoDocObject::XOleInPlaceSite::OnUIActivate(void)
{
	METHOD_PROLOGUE(CDsoDocObject, OleInPlaceSite);
    ODS("CDsoDocObject::XOleInPlaceSite::OnUIActivate\n");
    pThis->m_fObjectUIActive = TRUE;
	pThis->m_pipobj->GetWindow(&(pThis->m_hwndIPObject));
	return S_OK;
}

STDMETHODIMP CDsoDocObject::XOleInPlaceSite::GetWindowContext(IOleInPlaceFrame** ppFrame,
	IOleInPlaceUIWindow** ppDoc, LPRECT lprcPosRect, LPRECT lprcClipRect, LPOLEINPLACEFRAMEINFO lpFrameInfo)
{
	METHOD_PROLOGUE(CDsoDocObject, OleInPlaceSite);
    ODS("CDsoDocObject::XOleInPlaceSite::GetWindowContext\n");

	if (ppFrame)
        { SAFE_SET_INTERFACE(*ppFrame, &(pThis->m_xOleInPlaceFrame)); }

    if (ppDoc)
		*ppDoc = NULL;
    
    if (lprcPosRect)
		*lprcPosRect = pThis->m_rcViewRect;

    if (lprcClipRect)
		*lprcClipRect = *lprcPosRect;

	memset(lpFrameInfo, 0, sizeof(OLEINPLACEFRAMEINFO));
    lpFrameInfo->cb = sizeof(OLEINPLACEFRAMEINFO);
    lpFrameInfo->hwndFrame = pThis->m_hwnd;
    return S_OK;
}

STDMETHODIMP CDsoDocObject::XOleInPlaceSite::Scroll(SIZE sz)
{
    ODS("CDsoDocObject::XOleInPlaceSite::Scroll\n");
	return S_OK;
}

STDMETHODIMP CDsoDocObject::XOleInPlaceSite::OnUIDeactivate(BOOL fUndoable)
{
	METHOD_PROLOGUE(CDsoDocObject, OleInPlaceSite);
    ODS("CDsoDocObject::XOleInPlaceSite::OnUIDeactivate\n");

    pThis->m_fObjectUIActive = FALSE;
	pThis->m_xOleInPlaceFrame.SetMenu(NULL, NULL, NULL);
    SetFocus(pThis->m_hwnd);

    return S_OK;
}

STDMETHODIMP CDsoDocObject::XOleInPlaceSite::OnInPlaceDeactivate(void)
{
	METHOD_PROLOGUE(CDsoDocObject, OleInPlaceSite);
    ODS("CDsoDocObject::XOleInPlaceSite::OnInPlaceDeactivate\n");

	pThis->m_fObjectIPActive = FALSE;
	pThis->m_hwndIPObject = NULL;
	SAFE_RELEASE_INTERFACE((pThis->m_pipobj));

	pThis->m_fObjectActivateComplete = FALSE;
    return S_OK;
}

STDMETHODIMP CDsoDocObject::XOleInPlaceSite::DiscardUndoState(void)
{
    ODS("CDsoDocObject::XOleInPlaceSite::DiscardUndoState\n");
	return S_OK;
}

STDMETHODIMP CDsoDocObject::XOleInPlaceSite::DeactivateAndUndo(void)
{
	METHOD_PROLOGUE(CDsoDocObject, OleInPlaceSite);
    ODS("CDsoDocObject::XOleInPlaceSite::DeactivateAndUndo\n");
    if (pThis->m_pipobj) pThis->m_pipobj->InPlaceDeactivate();
    return S_OK;
}

STDMETHODIMP CDsoDocObject::XOleInPlaceSite::OnPosRectChange(LPCRECT lprcPosRect)
{
    ODS("CDsoDocObject::XOleInPlaceSite::OnPosRectChange\n");
	return S_OK;
}

////////////////////////////////////////////////////////////////////////
//
// CDsoDocObject::XOleDocumentSite -- IOleDocumentSite Implementation
//
//	 STDMETHODIMP ActivateMe(IOleDocumentView* pView);
//
IMPLEMENT_INTERFACE_UNKNOWN(CDsoDocObject, OleDocumentSite)

STDMETHODIMP CDsoDocObject::XOleDocumentSite::ActivateMe(IOleDocumentView* pView)
{
	METHOD_PROLOGUE(CDsoDocObject, OleDocumentSite);
    ODS("CDsoDocObject::XOleDocumentSite::ActivateMe\n");

	HRESULT             hr;
    IOleDocument*       pmsodoc;
    
 // If we're passed a NULL view pointer, then try to get one from
 // the document object (the object within us).
    if (pView)
	{
	 // Make sure that the view has our client site
		pView->SetInPlaceSite((IOleInPlaceSite*)&(pThis->m_xOleInPlaceSite));
        if (pThis->m_pstmview)
            pView->ApplyViewState(pThis->m_pstmview);
		pView->AddRef();
	}
	else
    {
        if ((!(pThis->m_pole)) || 
			FAILED(pThis->m_pole->QueryInterface(IID_IOleDocument, (void **)&pmsodoc)))
            return E_FAIL;

		hr = pmsodoc->CreateView((IOleInPlaceSite*)&(pThis->m_xOleInPlaceSite),
							pThis->m_pstmview, 0, &pView);

        if (SUCCEEDED(hr) && (pThis->m_pstmview))
            hr = pView->ApplyViewState(pThis->m_pstmview);

		pmsodoc->Release();

		if (FAILED(hr)) return hr;
	}

    pThis->m_pdocv = pView;

 // Get a command target (if available)...
    pView->QueryInterface(IID_IOleCommandTarget, (void**)&(pThis->m_pcmdt));

 // This sets up toolbars and menus first    
	if (SUCCEEDED(hr = pView->UIActivate(TRUE)))
	{

	 // Set the window size sensitive to new toolbars
		pView->SetRect(&(pThis->m_rcViewRect));

	 // Makes it all active
		pView->Show(TRUE);

		pThis->m_fAppWindowActive = TRUE;

	 // Toogle tools off if that's what user wants...
		if (!(pThis->m_fDisplayTools))
		{
			pThis->m_fDisplayTools = TRUE;
			pThis->OnNotifyChangeToolState(FALSE);
		}

		pThis->m_fObjectActivateComplete = TRUE;
	}

    return hr;
}

////////////////////////////////////////////////////////////////////////
//
// CDsoDocObject::XOleInPlaceFrame -- IOleInPlaceFrame Implementation
//
//   STDMETHODIMP GetWindow(HWND* phWnd);
//   STDMETHODIMP ContextSensitiveHelp(BOOL fEnterMode);
//   STDMETHODIMP GetBorder(LPRECT prcBorder);
//   STDMETHODIMP RequestBorderSpace(LPCBORDERWIDTHS pBW);
//   STDMETHODIMP SetBorderSpace(LPCBORDERWIDTHS pBW);
//   STDMETHODIMP SetActiveObject(LPOLEINPLACEACTIVEOBJECT pIIPActiveObj, LPCOLESTR pszObj);
//   STDMETHODIMP InsertMenus(HMENU hMenu, LPOLEMENUGROUPWIDTHS pMGW);
//   STDMETHODIMP SetMenu(HMENU hMenu, HOLEMENU hOLEMenu, HWND hWndObj);
//   STDMETHODIMP RemoveMenus(HMENU hMenu);
//   STDMETHODIMP SetStatusText(LPCOLESTR pszText);
//   STDMETHODIMP EnableModeless(BOOL fEnable);
//   STDMETHODIMP TranslateAccelerator(LPMSG pMSG, WORD wID);
//
IMPLEMENT_INTERFACE_UNKNOWN(CDsoDocObject, OleInPlaceFrame)

STDMETHODIMP CDsoDocObject::XOleInPlaceFrame::GetWindow(HWND* phWnd)
{
	METHOD_PROLOGUE(CDsoDocObject, OleInPlaceFrame);
    ODS("CDsoDocObject::XOleInPlaceFrame::GetWindow\n");
	return pThis->m_xOleInPlaceSite.GetWindow(phWnd);
}

STDMETHODIMP CDsoDocObject::XOleInPlaceFrame::ContextSensitiveHelp(BOOL fEnterMode)
{
	METHOD_PROLOGUE(CDsoDocObject, OleInPlaceFrame);
    ODS("CDsoDocObject::XOleInPlaceFrame::ContextSensitiveHelp\n");
	return pThis->m_xOleInPlaceSite.ContextSensitiveHelp(fEnterMode);
}


STDMETHODIMP CDsoDocObject::XOleInPlaceFrame::GetBorder(LPRECT prcBorder)
{
	METHOD_PROLOGUE(CDsoDocObject, OleInPlaceFrame);
    ODS("CDsoDocObject::XOleInPlaceFrame::GetBorder\n");
	CHECK_NULL_RETURN(prcBorder, E_POINTER);

 // If we don't allow Toolspace, and we are already active, give
 // no space for tools (ie, hide toolabrs), otherwise give as much we can...
	if (!(pThis->m_fDisplayTools) && (pThis->m_pipactive))
		SetRectEmpty(prcBorder);
	else
		GetClientRect(pThis->m_hwnd, prcBorder);

	TRACE_LPRECT("prcBorder", prcBorder);
	return S_OK;
}

STDMETHODIMP CDsoDocObject::XOleInPlaceFrame::RequestBorderSpace(LPCBORDERWIDTHS pBW)
{
    ODS("CDsoDocObject::XOleInPlaceFrame::RequestBorderSpace\n");
	CHECK_NULL_RETURN(pBW, E_POINTER);
	TRACE_LPRECT("pBW", pBW);
	return S_OK;
}


STDMETHODIMP CDsoDocObject::XOleInPlaceFrame::SetBorderSpace(LPCBORDERWIDTHS pBW)
{
	METHOD_PROLOGUE(CDsoDocObject, OleInPlaceFrame);
    ODS("CDsoDocObject::XOleInPlaceFrame::SetBorderSpace\n");
	if(pThis->m_hwnd){
		HWND hWnd = GetFocus();
		if(hWnd){
			char cName[300];
			GetClassName(hWnd, cName, 300);
			cName[299] = 0;
			if(strcmp("DSealMainWnd", cName) == 0){ 
 				return S_OK;
			}
		}
	} 
    if (pBW){TRACE_LPRECT("pBW", pBW);}
 	RECT rc;
	
	GetClientRect(pThis->m_hwnd, &rc);
	SetRectEmpty((RECT*)&(pThis->m_bwToolSpace));

    if (pBW)
    {
		pThis->m_bwToolSpace = *pBW;

        rc.left   += pBW->left;   rc.right  -= pBW->right;
        rc.top    += pBW->top;    rc.bottom -= pBW->bottom;
    }

 // Save the current view RECT (space minus tools)...
    pThis->m_rcViewRect = rc;

 // Update the active document (if alive)...
    if (pThis->m_pdocv)
        pThis->m_pdocv->SetRect(&(pThis->m_rcViewRect));

    return S_OK;
}

STDMETHODIMP CDsoDocObject::XOleInPlaceFrame::SetActiveObject(LPOLEINPLACEACTIVEOBJECT pIIPActiveObj, LPCOLESTR pszObj)
{
	METHOD_PROLOGUE(CDsoDocObject, OleInPlaceFrame);
    ODS("CDsoDocObject::XOleInPlaceFrame::SetActiveObject\n");

	SAFE_RELEASE_INTERFACE((pThis->m_pipactive));
	pThis->m_hwndUIActiveObj = NULL;
	pThis->m_dwObjectThreadID = 0;

    if (pIIPActiveObj)
	{
        SAFE_SET_INTERFACE(pThis->m_pipactive, pIIPActiveObj);
		pIIPActiveObj->GetWindow(&(pThis->m_hwndUIActiveObj));
		pThis->m_dwObjectThreadID = GetWindowThreadProcessId(pThis->m_hwndUIActiveObj, NULL);
	}

    return S_OK;
}

STDMETHODIMP CDsoDocObject::XOleInPlaceFrame::InsertMenus(HMENU hMenu, LPOLEMENUGROUPWIDTHS pMGW)
{
    ODS("CDsoDocObject::XOleInPlaceFrame::InsertMenus\n");
	return S_OK;
}

STDMETHODIMP CDsoDocObject::XOleInPlaceFrame::SetMenu(HMENU hMenu, HOLEMENU hOLEMenu, HWND hWndObj)
{
	METHOD_PROLOGUE(CDsoDocObject, OleInPlaceFrame);
    ODS("CDsoDocObject::XOleInPlaceFrame::SetMenu\n");

 // We really don't do anything here. We will merge the menu and set it
 // later if the menubar is visible or user chooses dropdown from titlebar.
 // All we do is stash current values or clear them depending on call.
    if (hMenu)
    {
		pThis->m_hMenuActive = hMenu;
		pThis->m_holeMenu = hOLEMenu;
		pThis->m_hwndMenuObj = hWndObj;

     // Handle a special case where menu is being updated after initial set
     // and not in close. In such a case we should force redraw control we
     // are in so new menu changes are visible as soon as possible...
        if ((!(pThis->m_fInClose)) && (pThis->m_hwndCtl))
            InvalidateRect(pThis->m_hwndCtl, NULL, FALSE);
    }
    else
    {
		pThis->m_hMenuActive = NULL;
		pThis->m_holeMenu = NULL;
		pThis->m_hwndMenuObj = NULL;
    }

 // Regardless of call, make sure to cleanup a merged menu if
 // one was previously created for the control...
	pThis->ClearMergedMenu();
	return S_OK;
}

STDMETHODIMP CDsoDocObject::XOleInPlaceFrame::RemoveMenus(HMENU hMenu)
{
    ODS("CDsoDocObject::XOleInPlaceFrame::RemoveMenus\n");
	return S_OK;
}

STDMETHODIMP CDsoDocObject::XOleInPlaceFrame::SetStatusText(LPCOLESTR pszText)
{
	//METHOD_PROLOGUE(CDsoDocObject, OleInPlaceFrame);
    ODS("CDsoDocObject::XOleInPlaceFrame::SetStatusText\n");
    if ((pszText) && (*pszText)){TRACE1(" -- Status Text = %S \n", pszText);}
	return S_OK;
}

STDMETHODIMP CDsoDocObject::XOleInPlaceFrame::EnableModeless(BOOL fEnable)
{
	METHOD_PROLOGUE(CDsoDocObject, OleInPlaceFrame);
    TRACE1("CDsoDocObject::XOleInPlaceFrame::EnableModeless(%d)\n", fEnable);
	pThis->m_fObjectInModalCondition = !fEnable;
	SendMessage(pThis->m_hwndCtl, DSO_WM_ASYNCH_STATECHANGE, DSO_STATE_MODAL, (LPARAM)fEnable);
	return S_OK;
}

STDMETHODIMP CDsoDocObject::XOleInPlaceFrame::TranslateAccelerator(LPMSG pMSG, WORD wID)
{
    ODS("CDsoDocObject::XOleInPlaceFrame::TranslateAccelerator\n");
	return S_FALSE;
}

////////////////////////////////////////////////////////////////////////
//
// CDsoDocObject::XOleCommandTarget -- IOleCommandTarget Implementation
//
//   STDMETHODIMP QueryStatus(const GUID *pguidCmdGroup, ULONG cCmds, OLECMD prgCmds[], OLECMDTEXT *pCmdText);
//   STDMETHODIMP Exec(const GUID *pguidCmdGroup, DWORD nCmdID, DWORD nCmdexecopt, VARIANTARG *pvaIn, VARIANTARG *pvaOut);            
//
IMPLEMENT_INTERFACE_UNKNOWN(CDsoDocObject, OleCommandTarget)

STDMETHODIMP CDsoDocObject::XOleCommandTarget::QueryStatus(const GUID *pguidCmdGroup, ULONG cCmds, OLECMD prgCmds[], OLECMDTEXT *pCmdText)
{
	HRESULT hr = OLECMDERR_E_UNKNOWNGROUP;
	METHOD_PROLOGUE(CDsoDocObject, OleCommandTarget);
	if (pThis->m_pcmdCtl)
		hr = pThis->m_pcmdCtl->QueryStatus(pguidCmdGroup, cCmds, prgCmds, pCmdText);
	return hr;
}

STDMETHODIMP CDsoDocObject::XOleCommandTarget::Exec(const GUID *pguidCmdGroup, DWORD nCmdID, DWORD nCmdexecopt, VARIANTARG *pvaIn, VARIANTARG *pvaOut)
{
	HRESULT hr = OLECMDERR_E_NOTSUPPORTED;
	METHOD_PROLOGUE(CDsoDocObject, OleCommandTarget);
	if (pThis->m_pcmdCtl)
		hr = pThis->m_pcmdCtl->Exec(pguidCmdGroup, nCmdID, nCmdexecopt, pvaIn, pvaOut);
	return hr;
}

////////////////////////////////////////////////////////////////////////
//
// CDsoDocObject::XAuthenticate -- IAuthenticate Implementation
//
//   STDMETHODIMP Authenticate(HWND *phwnd, LPWSTR *pszUsername, LPWSTR *pszPassword);
//
IMPLEMENT_INTERFACE_UNKNOWN(CDsoDocObject, Authenticate)

STDMETHODIMP CDsoDocObject::XAuthenticate::Authenticate(HWND *phwnd, LPWSTR *pszUsername, LPWSTR *pszPassword)
{
	METHOD_PROLOGUE(CDsoDocObject, Authenticate);
    ODS("CDsoDocObject::XAuthenticate::Authenticate\n");
	if (phwnd) *phwnd = ((pThis->m_pwszUsername) ? (HWND)INVALID_HANDLE_VALUE : pThis->m_hwndCtl);
	if (pszUsername) *pszUsername = DsoConvertToLPOLESTR(pThis->m_pwszUsername);
	if (pszPassword) *pszPassword = DsoConvertToLPOLESTR(pThis->m_pwszPassword);
    return S_OK;
}

////////////////////////////////////////////////////////////////////////
//
// CDsoDocObject::XServiceProvider -- IServiceProvider Implementation
//
//   STDMETHODIMP QueryService(REFGUID guidService, REFIID riid, void **ppv);
//
IMPLEMENT_INTERFACE_UNKNOWN(CDsoDocObject, ServiceProvider)

STDMETHODIMP CDsoDocObject::XServiceProvider::QueryService(REFGUID guidService, REFIID riid, void **ppv)
{
	METHOD_PROLOGUE(CDsoDocObject, ServiceProvider);
    ODS("CDsoDocObject::XServiceProvider::QueryService\n");

    if (guidService == SID_SContainerDispatch)
    {
        if ((ppv) && (riid == IID_IDispatch) && (pThis->m_pcmdCtl))
        {
            return pThis->m_pcmdCtl->QueryInterface(IID_IDispatch, ppv);
        }
    }

    return E_NOINTERFACE;
}

////////////////////////////////////////////////////////////////////////
//
// CDsoDocObject::XContinueCallback -- IContinueCallback Implementation
//
//   STDMETHODIMP FContinue(void);
//   STDMETHODIMP FContinuePrinting(LONG cPagesPrinted, LONG nCurrentPage, LPOLESTR pwszPrintStatus);
//
IMPLEMENT_INTERFACE_UNKNOWN(CDsoDocObject, ContinueCallback)

STDMETHODIMP CDsoDocObject::XContinueCallback::FContinue(void)
{ // We don't support asynchronous cancel of printing, but if you wanted to add
  // such functionality, this is where you could do so...
    return S_OK;
}

STDMETHODIMP CDsoDocObject::XContinueCallback::FContinuePrinting(LONG cPagesPrinted, LONG nCurrentPage, LPOLESTR pwszPrintStatus)
{
    TRACE3("CDsoDocObject::XContinueCallback::FContinuePrinting(%d, %d, %S)\n", cPagesPrinted, nCurrentPage, pwszPrintStatus);
    return S_OK;
}

////////////////////////////////////////////////////////////////////////
//
// CDsoDocObject::XPreviewCallback -- IPreviewCallback Implementation
//
//   STDMETHODIMP Notify(DWORD wStatus, LONG nLastPage, LPOLESTR pwszPreviewStatus);
//
IMPLEMENT_INTERFACE_UNKNOWN(CDsoDocObject, PreviewCallback)

STDMETHODIMP CDsoDocObject::XPreviewCallback::Notify(DWORD wStatus, LONG nLastPage, LPOLESTR pwszPreviewStatus)
{
	METHOD_PROLOGUE(CDsoDocObject, PreviewCallback);
    TRACE3("CDsoDocObject::XPreviewCallback::Notify(%d, %d, %S)\n", wStatus, nLastPage, pwszPreviewStatus);

  // The only notification we act on is when the preview is done...
    if ((wStatus == NOTIFY_FORCECLOSEPREVIEW) ||
        (wStatus == NOTIFY_FINISHED) ||
        (wStatus == NOTIFY_UNABLETOPREVIEW))
            pThis->ExitPrintPreview(FALSE);

    return S_OK;
}

////////////////////////////////////////////////////////////////////////
//
// CDsoDocObject::HrGetDataFromObject
//
//  Function designed to give indirect access to the content in an existing
//  loaded object, in a clipboard format specified in vtType. In other words,
//  a call to HrGetDataFromObject("Rich Text Format", [out] array) will get the
//  content of the current document in RTF format and return it as a byte array
//  vector in VB6-safe safearray. This array can be converted to a string for 
//  display or saved to database (etc.) without saving to disk. 
//
//  For this to work, the object embedded must support IDataObject and the clip
//  format that you request.
//
STDMETHODIMP CDsoDocObject::HrGetDataFromObject(VARIANT *pvtType, VARIANT *pvtOutput)
{
    HRESULT hr;
    IDataObject *pdo = NULL;
    LPWSTR pwszTypeName;
    LPSTR pszFormatName;
    SAFEARRAY* psa;
    VOID HUGEP* prawdata;
    FORMATETC ftc;
    STGMEDIUM stgm;
    LONG cfType;
    
    if ((pvtType == NULL)   || PARAM_IS_MISSING(pvtType) || 
        (pvtOutput == NULL) || PARAM_IS_MISSING(pvtOutput))
        return E_INVALIDARG;

    VariantClear(pvtOutput);

 // We take the name and find the right clipformat for the data type...
    pwszTypeName = LPWSTR_FROM_VARIANT(*pvtType);
    if ((pwszTypeName) && (pszFormatName = DsoConvertToMBCS(pwszTypeName)))
    {
        cfType = RegisterClipboardFormat(pszFormatName);
        DsoMemFree(pszFormatName);
    }
    else cfType = LONG_FROM_VARIANT(*pvtType, 0);
    CHECK_NULL_RETURN(cfType, E_INVALIDARG);

 // We must be able to get IDataObject for the transfer to work...
    if ((m_pole == NULL) || 
        (FAILED(m_pole->GetClipboardData(0, &pdo)) && 
         FAILED(m_pole->QueryInterface(IID_IDataObject, (void**)&pdo))))
         return OLE_E_CANT_GETMONIKER;

    ASSERT(pdo); CHECK_NULL_RETURN(pdo, E_UNEXPECTED);

 // We are going to ask for HGLOBAL data format only. This is majority 
 // of the non-binary formats, which should be sufficient here...
    memset(&ftc, 0, sizeof(ftc));
    ftc.cfFormat = (WORD)cfType;
    ftc.dwAspect = DVASPECT_CONTENT;
    ftc.lindex = -1; ftc.tymed = TYMED_HGLOBAL;

    memset(&stgm, 0, sizeof(stgm));
    stgm.tymed = TYMED_HGLOBAL;

 // Ask the object for the data...
    if (SUCCEEDED(hr = pdo->QueryGetData(&ftc)) && 
        SUCCEEDED(hr = pdo->GetData(&ftc, &stgm)))
    {
        ULONG ulSize;
        if ((stgm.tymed == TYMED_HGLOBAL) && (stgm.hGlobal) &&
            (ulSize = GlobalSize(stgm.hGlobal)))
        {
            LPVOID lpv = GlobalLock(stgm.hGlobal);
            if (lpv)
            {
             // We will return data as safearray vector of VB6 Byte type...
                psa = SafeArrayCreateVector(VT_UI1, 1, ulSize);
                if (psa)
                {
                    pvtOutput->vt = VT_ARRAY|VT_UI1;
                    pvtOutput->parray = psa;
                    prawdata = NULL;

                    if (SUCCEEDED(SafeArrayAccessData(psa, &prawdata)))
                    {
                        memcpy(prawdata, lpv, ulSize);
                        SafeArrayUnaccessData(psa);
                    }
                }
                else hr = E_OUTOFMEMORY;

                GlobalUnlock(stgm.hGlobal);
            }
            else hr = E_ACCESSDENIED;
        }
        else hr = E_FAIL;

        ReleaseStgMedium(&stgm);
    }

    pdo->Release();
    return hr;
}

////////////////////////////////////////////////////////////////////////
//
// CDsoDocObject::HrSetDataInObject
//
//   Function to take format type and either string or VB6-safe byte array
//   as data and set it into the current object (that is, it will replace the
//   current document content with this content). The fMbcsString flag determines
//   if BSTR passed should be converted to MBCS format before being set. For example,
//   if you previously got data as RTF (MBCS), then converted to Unicode BSTR and
//   passed it back, it would need to be converted back to MBCS for the set to work.
//   If the string data is binary, pass FALSE. If data is array, the param is ignored.
//
//   The function sets the data using IDataObject::SetData. Note that some objects 
//   will allow you to get a format but not set it, so this may not work for a previously
//   acquired data type. It is up to the host DocObject server, so test carefully before
//   depending on this functionality.
//
STDMETHODIMP CDsoDocObject::HrSetDataInObject(VARIANT *pvtType, VARIANT *pvtInput, BOOL fMbcsString)
{
    HRESULT hr;
    IDataObject *pdo = NULL;
    LPWSTR pwszTypeName;
    LPSTR pszFormatName;
    SAFEARRAY* psa;
    VOID HUGEP* prawdata;
    FORMATETC ftc;
    STGMEDIUM stgm;
    LONG cfType;
    ULONG ulSize;
    BOOL fIsArrayData = FALSE;
    BOOL fCleanupString = FALSE;

    if ((pvtType == NULL)  || PARAM_IS_MISSING(pvtType) || 
        (pvtInput == NULL) || PARAM_IS_MISSING(pvtInput))
        return E_INVALIDARG;

 // Find the clipboard format for the given data type...
    pwszTypeName = LPWSTR_FROM_VARIANT(*pvtType);
    if ((pwszTypeName) && (pszFormatName = DsoConvertToMBCS(pwszTypeName)))
    {
        cfType = RegisterClipboardFormat(pszFormatName);
        DsoMemFree(pszFormatName);
    }
    else cfType = LONG_FROM_VARIANT(*pvtType, 0);
    CHECK_NULL_RETURN(cfType, E_INVALIDARG);

 // Depending on the type of data passed, and if we need to do any Unicode-to-MBCS
 // conversion for the caller, set up the data size and data pointer we will use
 // to create the global memory object we'll pass to the doc object server...
    prawdata = LPWSTR_FROM_VARIANT(*pvtInput);
    if (prawdata)
    {
        if (fMbcsString)
        {
            prawdata = (void*)DsoConvertToMBCS((BSTR)prawdata);
            CHECK_NULL_RETURN(prawdata, E_OUTOFMEMORY);
            fCleanupString = TRUE;
            ulSize = lstrlen((LPSTR)prawdata);
        }
        else
            ulSize = SysStringByteLen((BSTR)prawdata);
    }
    else
    {
        psa = PSARRAY_FROM_VARIANT(*pvtInput);
        if (psa)
        {
            LONG lb, ub, elSize;

            if ((SafeArrayGetDim(psa) > 1) ||
                FAILED(SafeArrayGetLBound(psa, 1, &lb)) || (lb < 0) ||
                FAILED(SafeArrayGetUBound(psa, 1, &ub)) || (ub < lb) ||
                ((elSize = SafeArrayGetElemsize(psa)) < 1))
                return E_INVALIDARG;

            ulSize = (((ub + 1) - lb) * elSize);
            fIsArrayData = TRUE;
            if (FAILED(SafeArrayAccessData(psa, &prawdata)))
                return E_ACCESSDENIED;
        }

        CHECK_NULL_RETURN(prawdata, E_INVALIDARG);
    }

 // We must have a server and it must support IDataObject...
    if ((m_pole) && SUCCEEDED(m_pole->QueryInterface(IID_IDataObject, (void**)&pdo)) && (pdo))
    {
        memset(&ftc, 0, sizeof(ftc));
        ftc.cfFormat = (WORD)cfType;
        ftc.dwAspect = DVASPECT_CONTENT;
        ftc.lindex = -1; ftc.tymed = TYMED_HGLOBAL;

        memset(&stgm, 0, sizeof(stgm));
        stgm.tymed = TYMED_HGLOBAL;
        stgm.hGlobal = GlobalAlloc(GPTR, ulSize);
        if (stgm.hGlobal)
        {
            LPVOID lpv = GlobalLock(stgm.hGlobal);
            if ((lpv) && (ulSize))
            {
             // Copy the data into transfer object...
                memcpy(lpv, prawdata, ulSize);
                GlobalUnlock(stgm.hGlobal);

             // Do the actual SetData call to transfer the data...
                hr = pdo->SetData(&ftc, &stgm, TRUE);
            } 
            else hr = E_UNEXPECTED;

            if (FAILED(hr))
                ReleaseStgMedium(&stgm);
        }
        else hr = E_OUTOFMEMORY;

        pdo->Release();
    }
    else hr = OLE_E_CANT_GETMONIKER;

    if (fIsArrayData)
         SafeArrayUnaccessData(psa);

    if (fCleanupString)
        DsoMemFree(prawdata);

    return hr;
}

////////////////////////////////////////////////////////////////////////
// CDsoDocObject::FrameWindowProc
//
//  Site window procedure. Not much to do here except forward focus.
//
STDMETHODIMP_(LRESULT) CDsoDocObject::FrameWindowProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
	CDsoDocObject* pbndr = (CDsoDocObject*)GetWindowLong(hwnd, GWL_USERDATA);
	if (pbndr)
	{
		switch (msg)
		{
		case WM_PAINT:
			{
				PAINTSTRUCT ps;
				RECT rc; GetClientRect(hwnd, &rc);
				BeginPaint(hwnd, &ps);
				pbndr->OnDraw(DVASPECT_CONTENT, ps.hdc, (RECT*)&rc, NULL, NULL, TRUE);
				EndPaint(hwnd, &ps);
			}
			break;

		case WM_NCDESTROY:
			SetWindowLong(hwnd, GWL_USERDATA, 0);
			pbndr->m_hwnd = NULL;
			break;

		case WM_SETFOCUS:
			if (pbndr->m_hwndUIActiveObj)
				SetFocus(pbndr->m_hwndUIActiveObj);
			break;

		}

	}

	return DefWindowProc(hwnd, msg, wParam, lParam);
}



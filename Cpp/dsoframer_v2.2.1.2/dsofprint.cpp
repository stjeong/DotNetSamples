/***************************************************************************
 * DSOFPRINT.CPP
 *
 * CDsoDocObject: Print Code for CDsoDocObject object
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
#include "dsoframer.h"
#include "msoffice.h"
#include <atlbase.h>
#include <atlconv.h>
#include <afxdisp.h>
#include <afxwin.h>
#include <tchar.h>
#include <wininet.h>
#include <afxinet.h>
#include "XMLHttpClient.h"
 
#include <atlbase.h>
#include <mshtml.h> 
#include <mshtmdid.h> 
#include <comdef.h>	

#import "C:\WINDOWS\system32\msxml3.dll"
using namespace MSXML2;
extern char* BSTR2char(const BSTR bstr) ;
////////////////////////////////////////////////////////////////////////
// CDsoFramerControl::_PrintOutOld
//
//  Prints the current object by calling IOleCommandTarget for Print.
//
STDMETHODIMP CDsoFramerControl::_PrintOutOld(VARIANT PromptToSelectPrinter)
{
    DWORD dwOption = BOOL_FROM_VARIANT(PromptToSelectPrinter, FALSE) 
                   ? OLECMDEXECOPT_PROMPTUSER : OLECMDEXECOPT_DODEFAULT;

    TRACE1("CDsoFramerControl::_PrintOutOld(%d)\n", dwOption);
    CHECK_NULL_RETURN(m_pDocObjFrame, ProvideErrorInfo(DSO_E_DOCUMENTNOTOPEN));

 // Cannot access object if in modal condition...
    if ((m_fModalState) || (m_pDocObjFrame->InPrintPreview()))
        return ProvideErrorInfo(DSO_E_INMODALSTATE);

	return m_pDocObjFrame->DoOleCommand(OLECMDID_PRINT, dwOption, NULL, NULL);
}

////////////////////////////////////////////////////////////////////////
// CDsoFramerControl::PrintOut
//
//  Prints document using either IPrint or IOleCommandTarget, depending
//  on parameters passed. This offers a bit more control over printing if
//  the doc object server supports IPrint interface.
//
STDMETHODIMP CDsoFramerControl::PrintOut(VARIANT PromptUser, VARIANT PrinterName,
                VARIANT Copies, VARIANT FromPage, VARIANT ToPage, VARIANT OutputFile)
{
    HRESULT hr;
    BOOL fPromptUser   = BOOL_FROM_VARIANT(PromptUser, FALSE);
    LPWSTR pwszPrinter = LPWSTR_FROM_VARIANT(PrinterName);
    LPWSTR pwszOutput  = LPWSTR_FROM_VARIANT(OutputFile);
    LONG lCopies       = LONG_FROM_VARIANT(Copies, 1);
    LONG lFrom         = LONG_FROM_VARIANT(FromPage, 0);
    LONG lTo           = LONG_FROM_VARIANT(ToPage, 0);

	TRACE3("CDsoFramerControl::PrintOut(%d, %S, %d)\n", fPromptUser, pwszPrinter, lCopies);
    CHECK_NULL_RETURN(m_pDocObjFrame, ProvideErrorInfo(DSO_E_DOCUMENTNOTOPEN));

 // First we do validation of all parameters passed to the function...
    if ((pwszPrinter) && (*pwszPrinter == L'\0'))
        return E_INVALIDARG;

    if ((pwszOutput) && (*pwszOutput == L'\0'))
        return E_INVALIDARG;

    if ((lCopies < 1) || (lCopies > 200))
        return E_INVALIDARG;

    if (((lFrom != 0) || (lTo != 0)) && ((lFrom < 1) || (lTo < lFrom)))
        return E_INVALIDARG;

 // Cannot access object if in modal condition...
    if ((m_fModalState) || (m_pDocObjFrame->InPrintPreview()))
        return ProvideErrorInfo(DSO_E_INMODALSTATE);

 // If no printer name was provided, we can print to the default device
 // using IOleCommandTarget with OLECMDID_PRINT...
    if (pwszPrinter == NULL)
        return _PrintOutOld(PromptUser);

 // Ask the embedded document to print itself to specific printer...
    hr = m_pDocObjFrame->PrintDocument(pwszPrinter, pwszOutput, (UINT)lCopies, 
                        (UINT)lFrom, (UINT)lTo, fPromptUser);

 // If call failed because interface doesn't exist, change error
 // to let caller know it is because DocObj doesn't support this command...
    if (FAILED(hr) && (hr == E_NOINTERFACE))
        hr = DSO_E_COMMANDNOTSUPPORTED;

    return ProvideErrorInfo(hr);
}

////////////////////////////////////////////////////////////////////////
// CDsoFramerControl::PrintPreview
//
//  Asks embedded object to attempt a print preview (only Office docs do this).
//
STDMETHODIMP CDsoFramerControl::PrintPreview()
{
    HRESULT hr;
	ODS("CDsoFramerControl::PrintPreview\n");
    CHECK_NULL_RETURN(m_pDocObjFrame, ProvideErrorInfo(DSO_E_DOCUMENTNOTOPEN));

    if (m_fModalState) // Cannot access object if in modal condition...
        return ProvideErrorInfo(DSO_E_INMODALSTATE);

 // Try to set object into print preview mode...
    hr = m_pDocObjFrame->StartPrintPreview();

 // If call failed because interface doesn't exist, change error
 // to let caller know it is because DocObj doesn't support this command...
    if (FAILED(hr) && (hr == E_NOINTERFACE))
        hr = DSO_E_COMMANDNOTSUPPORTED;

    return ProvideErrorInfo(hr);
}

////////////////////////////////////////////////////////////////////////
// CDsoFramerControl::PrintPreviewExit
//
//  Closes an active preview.
//
STDMETHODIMP CDsoFramerControl::PrintPreviewExit()
{
	ODS("CDsoFramerControl::PrintPreviewExit\n");
    CHECK_NULL_RETURN(m_pDocObjFrame, ProvideErrorInfo(DSO_E_DOCUMENTNOTOPEN));

 // Try to set object out of print preview mode...
    if (m_pDocObjFrame->InPrintPreview())
        m_pDocObjFrame->ExitPrintPreview(TRUE);

    return S_OK;
}


////////////////////////////////////////////////////////////////////////
// CDsoDocObject::PrintDocument
//
//  We can use the IPrint interface for an ActiveX Document object to 
//  selectively print object to a given printer. 
//
STDMETHODIMP CDsoDocObject::PrintDocument(LPCWSTR pwszPrinter, LPCWSTR pwszOutput, UINT cCopies, UINT nFrom, UINT nTo, BOOL fPromptUser)
{
    HRESULT hr;
    IPrint *print;
    HANDLE hPrint;
    DVTARGETDEVICE* ptd = NULL;

    ODS("CDsoDocObject::PrintDocument\n");
    CHECK_NULL_RETURN(m_pole, E_UNEXPECTED);

 // First thing we need to do is ask object for IPrint. If it does not
 // support it, we cannot continue. It is up to DocObj if this is allowed...
    hr = m_pole->QueryInterface(IID_IPrint, (void**)&print);
    RETURN_ON_FAILURE(hr);

 // Now setup printer settings into DEVMODE for IPrint. Open printer 
 // settings and gather default DEVMODE...
    if (FOpenPrinter(pwszPrinter, &hPrint))
    {
        LPDEVMODEW pdevMode     = NULL;
        LPWSTR pwszDefProcessor = NULL;
        LPWSTR pwszDefDriver    = NULL;
        LPWSTR pwszDefPort      = NULL;
        LPWSTR pwszPort;
        DWORD  cbDevModeSize;

        if (FGetPrinterSettings(hPrint, &pwszDefProcessor, 
                &pwszDefDriver, &pwszDefPort, &pdevMode, &cbDevModeSize) && (pdevMode))
        {
            DWORD cbPrintName, cbDeviceName, cbOutputName;
            DWORD cbDVTargetSize;

            pdevMode->dmFields |= DM_COPIES;
            pdevMode->dmCopies = (WORD)((cCopies) ? cCopies : 1);

            pwszPort = ((pwszOutput) ? (LPWSTR)pwszOutput : pwszDefPort);
         
         // We calculate the size we will need for the TARGETDEVICE structure...
            cbPrintName  = ((lstrlenW(pwszDefProcessor) + 1) * sizeof(WCHAR));
            cbDeviceName = ((lstrlenW(pwszDefDriver)    + 1) * sizeof(WCHAR));
            cbOutputName = ((lstrlenW(pwszPort)         + 1) * sizeof(WCHAR));

            cbDVTargetSize = sizeof(DWORD) + sizeof(DEVNAMES) + cbPrintName +
                            cbDeviceName + cbOutputName + cbDevModeSize;

         // Allocate new target device using COM Task Allocator...
            ptd = (DVTARGETDEVICE*)CoTaskMemAlloc(cbDVTargetSize);
            if (ptd)
            {
             // Copy all the data in the DVT...
                DWORD dwOffset = sizeof(DWORD) + sizeof(DEVNAMES);
                ptd->tdSize = cbDVTargetSize;

                ptd->tdDriverNameOffset = (WORD)dwOffset;
                memcpy((BYTE*)(((BYTE*)ptd) + dwOffset), pwszDefProcessor, cbPrintName);
                dwOffset += cbPrintName;

                ptd->tdDeviceNameOffset = (WORD)dwOffset;
                memcpy((BYTE*)(((BYTE*)ptd) + dwOffset), pwszDefDriver, cbDeviceName);
                dwOffset += cbDeviceName;

                ptd->tdPortNameOffset = (WORD)dwOffset;
                memcpy((BYTE*)(((BYTE*)ptd) + dwOffset), pwszPort, cbOutputName);
                dwOffset += cbOutputName;

                ptd->tdExtDevmodeOffset = (WORD)dwOffset;
                memcpy((BYTE*)(((BYTE*)ptd) + dwOffset), pdevMode, cbDevModeSize);
                dwOffset += cbDevModeSize;

                ASSERT(dwOffset == cbDVTargetSize);
            }
            else hr = E_OUTOFMEMORY;

         // We're done with the devmode...
            DsoMemFree(pdevMode);
        }
        else hr = E_WIN32_LASTERROR;

        SAFE_FREESTRING(pwszDefPort);
        SAFE_FREESTRING(pwszDefDriver);
        SAFE_FREESTRING(pwszDefProcessor);
        ClosePrinter(hPrint);
    }
    else hr = E_WIN32_LASTERROR;

 // If we were successful in getting TARGETDEVICE struct, provide the page range
 // for the print job and ask docobj server to print it...
    if (SUCCEEDED(hr))
    {
        PAGESET *ppgset;
        DWORD cbPgRngSize = sizeof(PAGESET) + sizeof(PAGERANGE);
        LONG cPages, cLastPage;
        DWORD grfPrintFlags;

     // Setup the page range to print...
        if ((ppgset = (PAGESET*)CoTaskMemAlloc(cbPgRngSize)) != NULL)
        {
            ppgset->cbStruct = cbPgRngSize;
            ppgset->cPageRange   = 1;
            ppgset->fOddPages    = TRUE;
            ppgset->fEvenPages   = TRUE;
            ppgset->cPageRange   = 1;
            ppgset->rgPages[0].nFromPage = ((nFrom) ? nFrom : 1);
            ppgset->rgPages[0].nToPage   = ((nTo) ? nTo : PAGESET_TOLASTPAGE);

         // Give indication we are waiting (on the print)...
            HCURSOR hCur = SetCursor(LoadCursor(NULL, IDC_WAIT));

            SEH_TRY

        // Setup the initial page number (optional)...
            print->SetInitialPageNum(ppgset->rgPages[0].nFromPage);

            grfPrintFlags = (PRINTFLAG_MAYBOTHERUSER | PRINTFLAG_RECOMPOSETODEVICE);
            if (fPromptUser) grfPrintFlags |= PRINTFLAG_PROMPTUSER;
            if (pwszOutput)  grfPrintFlags |= PRINTFLAG_PRINTTOFILE;

         // Now ask server to print it using settings passed...
            hr = print->Print(grfPrintFlags, &ptd, &ppgset, NULL, (IContinueCallback*)&m_xContinueCallback, 
                    ppgset->rgPages[0].nFromPage, &cPages, &cLastPage);

            SEH_EXCEPT(hr)

            SetCursor(hCur);

            if (ppgset)
                CoTaskMemFree(ppgset);
        }
        else hr = E_OUTOFMEMORY;
    }

 // We are done...
    if (ptd) CoTaskMemFree(ptd);
    print->Release();
    return hr;
}


////////////////////////////////////////////////////////////////////////
// CDsoDocObject::StartPrintPreview
//
//  Ask embedded object to go into print preview (if supportted).
//
STDMETHODIMP CDsoDocObject::StartPrintPreview()
{
    HRESULT hr;
    IOleInplacePrintPreview *prev;
    HCURSOR hCur;

    ODS("CDsoDocObject::StartPrintPreview\n");
    CHECK_NULL_RETURN(m_pole, E_UNEXPECTED);

 // No need to do anything if already in preview...
    if (InPrintPreview()) return S_FALSE;

 // Otherwise, ask document server if it supports preview...
    hr = m_pole->QueryInterface(IID_IOleInplacePrintPreview, (void**)&prev);
    if (SUCCEEDED(hr))
    {
     // Tell user we waiting (switch to preview can be slow for very large docs)...
        hCur = SetCursor(LoadCursor(NULL, IDC_WAIT));

     // If it does, make sure it can go into preview mode...
        hr = prev->QueryStatus();
        if (SUCCEEDED(hr))
        {
            SEH_TRY

            if (m_hwndCtl) // Notify the control that preview started...
                SendMessage(m_hwndCtl, DSO_WM_ASYNCH_STATECHANGE, DSO_STATE_INTERACTIVE, (LPARAM)FALSE);

         // We will allow application to bother user and switch printers...
            hr = prev->StartPrintPreview(
                (PREVIEWFLAG_MAYBOTHERUSER | PREVIEWFLAG_PROMPTUSER | PREVIEWFLAG_USERMAYCHANGEPRINTER),
                NULL, (IOlePreviewCallback*)&m_xPreviewCallback, 1);

            SEH_EXCEPT(hr)

         // If the call succeeds, we keep hold of interface to close preview later
            if (SUCCEEDED(hr)) 
			{ 
				SAFE_SET_INTERFACE(m_pprtprv, prev);
			}
			else
			{ // Otherwise, notify the control that preview failed...
				if (m_hwndCtl) 
					PostMessage(m_hwndCtl, DSO_WM_ASYNCH_STATECHANGE, DSO_STATE_INTERACTIVE, (LPARAM)TRUE);
			}
        }

        SetCursor(hCur);
        prev->Release();
    }
    return hr;
}

////////////////////////////////////////////////////////////////////////
// CDsoDocObject::ExitPrintPreview
//
//  Drop out of print preview and restore items as needed.
//
STDMETHODIMP CDsoDocObject::ExitPrintPreview(BOOL fForceExit)
{
    TRACE1("CDsoDocObject::ExitPrintPreview(fForceExit=%d)\n", (DWORD)fForceExit);

 // If the user closes the app or otherwise terminates the preview, we need
 // to notify the ActiveDocument server to leave preview mode...
	try{
		if (m_pprtprv)
		{
		   if (fForceExit) // Tell docobj we want to end preview...
			{
 					HRESULT hr = m_pprtprv->EndPrintPreview (TRUE);
					ASSERT(SUCCEEDED(hr));
 
			}
		   if (m_hwndCtl) // Notify the control that preview ended...
			   PostMessage(m_hwndCtl, DSO_WM_ASYNCH_STATECHANGE, DSO_STATE_INTERACTIVE, (LPARAM)TRUE);

		 }
	}catch(...){
	}
 // Free our reference to preview interface...
    SAFE_RELEASE_INTERFACE(m_pprtprv);
    return S_OK;
}


STDMETHODIMP CDsoFramerControl::get_GetApplication(IDispatch** ppdisp)
{
	IDispatch *pActiveDocument;
	this->get_ActiveDocument(&pActiveDocument);
	//DISPID dispid;
	LPOLESTR rgszNames =L"Application";
	//DISPPARAMS dpNoArgs = {NULL, NULL, 0, 0};
	VARIANT vResult;
	HRESULT hr;
	if (pActiveDocument){
		hr = DsoDispatchInvoke(pActiveDocument, rgszNames,0, DISPATCH_PROPERTYGET, 0, NULL,&vResult);
		*ppdisp = (vResult.pdispVal);
		vResult.pdispVal->AddRef();
	}
	return ProvideErrorInfo(hr);
}

HRESULT CDsoFramerControl::SetCurrUserName(BSTR strCurrUserName, VARIANT_BOOL* pbool)
{
	char * strCurName = NULL;
	LPWSTR  strwName;
	if ((strCurrUserName) && (SysStringLen(strCurrUserName) > 0)){
		strwName = SysAllocString(strCurrUserName);
		strCurName = DsoConvertToMBCS(strwName);
	}	
	if(!strCurName){
		*pbool = FALSE;
		return S_OK;
	}
 	IDispatch * lDisp =  m_pDocDispatch;
	if(!lDisp){
		get_ActiveDocument(&lDisp);
		if(!lDisp){
			*pbool = FALSE;
			return S_OK;
		}
	}
	
	try{
		switch(m_nOriginalFileType){
		case FILE_TYPE_WORD:
			{
				
				CComQIPtr<MSWord::_Document> spDoc(lDisp);
				if(!spDoc)
					break;
				CComPtr<MSWord::_Application> spApp(spDoc->GetApplication());				
				spApp->put_UserName(A2BSTR(strCurName));
				//					spApp->put_PrintPreview(VARIANT_TRUE);
				spApp.Release();
				spDoc.Release();
				
			}	
			break;
		case FILE_TYPE_EXCEL:
			{	
				
				
			}
			break;
		case FILE_TYPE_PPT:
		case FILE_TYPE_PDF:
		case FILE_TYPE_UNK:
		default:
			break;
		}
	}catch(...){
		
	}
	
	
	
	*pbool = TRUE;
	DsoMemFree((void*)(strCurName));
	return S_OK;
}

HRESULT CDsoFramerControl::SetTrackRevisions( long vbool, VARIANT_BOOL* pbool)
{

	IDispatch * lDisp =  m_pDocDispatch;
	if(!lDisp){
		get_ActiveDocument(&lDisp);
		if(!lDisp){
			*pbool = FALSE;
			return S_OK;
		}
	}
	// 	COleVariant covTrue((short)TRUE), covFalse((short)FALSE), covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
	try{
		switch(m_nOriginalFileType){
		case FILE_TYPE_WORD:
			{	
				
				CComQIPtr<MSWord::_Document> spDoc(lDisp);
				if(!spDoc)
					break;
				if(4 == vbool){
					spDoc->AcceptAllRevisions();
				}else{
					spDoc->put_TrackRevisions( vbool );
				}
				spDoc.Release();
				
				/*_Document doc;
				doc.AttachDispatch(lDisp); 
				doc.SetTrackRevisions (vbool);
				doc.DetachDispatch();
				doc.ReleaseDispatch ();	*/
			}
			break;
		case FILE_TYPE_EXCEL:
			{	
				
				
			}
			break;
		case FILE_TYPE_PPT:
		case FILE_TYPE_PDF:
		case FILE_TYPE_UNK:
		default:
			break;
		}
	}catch(...){
		
	}
	*pbool = TRUE; 
	return S_OK;
}


STDMETHODIMP CDsoFramerControl::SetCurrTime( BSTR strValue,   VARIANT_BOOL* pbool)
{
	* pbool = FALSE;
	char * pTemp1 = NULL;
	LPWSTR  pwTemp1;
	if ((strValue) && (SysStringLen(strValue) > 0)){
		pwTemp1 = SysAllocString(strValue);
		pTemp1 = DsoConvertToMBCS(pwTemp1);
	}	
    if(!pTemp1){
		return S_OK;
	}
	char cTime[256];
	strcpy(cTime,pTemp1);
	if(pTemp1)	DsoMemFree((void*)(pTemp1));
	
	SYSTEMTIME tm;
	char *pPos = NULL;
	pTemp1 = cTime;
	BOOL bOK = FALSE;
	::GetLocalTime(&tm);
	try{
		do{
			pPos = strchr(pTemp1,'-');
			if(!pPos)
				break;
			*pPos++ = 0;
			tm.wYear = atol(pTemp1);
			
			pTemp1 = pPos;	
			pPos = strchr(pTemp1,'-');
			if(!pPos)
				break;
			*pPos++ = 0;
			tm.wMonth  = atol(pTemp1);
			
			pTemp1 = pPos;	
			pPos = strchr(pTemp1,' ');
			if(!pPos)
				break;
			*pPos++ = 0;
			tm.wDay  = atol(pTemp1);
			
			pTemp1 = pPos;	
			while(*pTemp1 == ' ') pTemp1++;
			tm.wHour   = atol(pTemp1) ;
			//			if(tm.wHour < 8)
			//				tm.wHour +=24;
			
			pTemp1 = pPos;	
			pPos = strchr(pTemp1,':');
			if(!pPos)
				break;
			pTemp1 = pPos + 1;
			pPos = strchr(pTemp1,':');
			if(!pPos)
				break;
			*pPos++ = 0;
			tm.wMinute   = atol(pTemp1);
			
			tm.wSecond  = atol(pPos);
			bOK = TRUE;
			
		}while(0);
	}catch(...){
		
	}
	if(bOK) 
		if(::SetLocalTime(&tm))
			* pbool = TRUE;
		
		return S_OK;
}
//Http Interface

#import "C:\WINDOWS\system32\msxml3.dll"
using namespace MSXML2;

BSTR  HttpSend(XMLHttpClient *pHttp, LPCTSTR szURL, DWORD *dwError)
{
 	BSTR bstrString = NULL;
	IXMLHTTPRequestPtr pIXMLHTTPRequest = NULL;
 	HRESULT hr;
	PBYTE pPostBuffer=NULL;
	DWORD dwPostBufferLength=pHttp->AllocMultiPartsFormData(pPostBuffer, "--MULTI-PARTS-FORM-DATA-BOUNDARY");
	try {
		hr=pIXMLHTTPRequest.CreateInstance("Msxml2.XMLHTTP.3.0");
		SUCCEEDED(hr) ? 0 : throw hr;

		hr=pIXMLHTTPRequest->open("POST", szURL, false);
		SUCCEEDED(hr) ? 0 : throw hr;

		CONST TCHAR *szAcceptType=__HTTP_ACCEPT_TYPE;	
		LPCTSTR szContentType=TEXT("Content-Type: multipart/form-data; boundary=--MULTI-PARTS-FORM-DATA-BOUNDARY\r\n");		


		
		VARIANT vt1;
		VARIANT *pbsSendBinary  = &vt1;
		pbsSendBinary->vt = VT_ARRAY | VT_UI1;
		V_ARRAY(pbsSendBinary) = NULL;
		LPSAFEARRAY psaFile;
		psaFile = ::SafeArrayCreateVector( VT_UI1 , 0, dwPostBufferLength );
		for( long k = 0; k < dwPostBufferLength; k++ )
		{
			if( FAILED(::SafeArrayPutElement( psaFile, &k, &pPostBuffer[k] )) )
			{
 
 				throw hr;
			}
		}

		pbsSendBinary->vt = VT_ARRAY | VT_UI1;
		V_ARRAY(pbsSendBinary) = psaFile;


		hr=pIXMLHTTPRequest->send(pbsSendBinary);
		SUCCEEDED(hr) ? 0 : throw hr;

		bstrString=pIXMLHTTPRequest->responseText;
 
		pHttp->FreeMultiPartsFormData(pPostBuffer);
	} catch (...) {
		pHttp->FreeMultiPartsFormData(pPostBuffer);
	}
 
	return bstrString;
}
STDMETHODIMP CDsoFramerControl::HttpInit(VARIANT_BOOL* pbool)
{
	if(m_pHttp){
		delete m_pHttp;
		m_pHttp = NULL;
	}
 	m_pHttp= new XMLHttpClient();

	m_pHttp->InitilizePostArguments();
	* pbool = TRUE;
	return S_OK;
}


STDMETHODIMP CDsoFramerControl::HttpAddPostCurrFile(BSTR strFileID,BSTR strFileName, VARIANT_BOOL* pbool)
{
	char * pTemp1 = NULL;
	LPWSTR  pwTemp1;
	if ((strFileID) && (SysStringLen(strFileID) > 0)){
		pwTemp1 = SysAllocString(strFileID);
		pTemp1 = DsoConvertToMBCS(pwTemp1);
	}	
	
	char * pTemp2 = NULL; 
	LPWSTR  pwTemp2;
	if ((strFileName) && (SysStringLen(strFileName) > 0)){
		pwTemp2 = SysAllocString(strFileName);
		pTemp2 = DsoConvertToMBCS(pwTemp2);
	}	
	
	if(!m_pHttp){
		if(pTemp2)	DsoMemFree((void*)(pTemp2));
		if(pTemp1)	DsoMemFree((void*)(pTemp1));
		*pbool = FALSE;
		return 0;
	}

 	IDispatch * lDisp =  m_pDocDispatch;
	if(!lDisp){
		get_ActiveDocument(&lDisp);
	}

	if(!lDisp){
		if(pTemp2)	DsoMemFree((void*)(pTemp2));
		if(pTemp1)	DsoMemFree((void*)(pTemp1));
		*pbool = FALSE;
		return S_OK;
	}


	COleVariant covTrue((short)TRUE), covFalse((short)FALSE), covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
	char cPath2[MAX_PATH];
	::GetTempPath(MAX_PATH, cPath2);
	strcat(cPath2,"DSOWebOffice");
	if (CreateDirectory(cPath2, NULL) || GetLastError() == ERROR_ALREADY_EXISTS)
	{
		lstrcat(cPath2, "\\");
	}
	::GetTempFileName(cPath2, "~dso", 0, cPath2);
	::DeleteFile(cPath2);

	try{
		switch(m_nOriginalFileType){
		case FILE_TYPE_WORD:
			{	
				CComQIPtr<MSWord::_Document> spDoc(lDisp);
				if(!spDoc)
					break;
				spDoc->SaveAs(COleVariant(cPath2));
				spDoc.Release();
			}
			break;
		case FILE_TYPE_EXCEL:
			{	
				CComQIPtr<MSExcel::_Workbook> spDoc(lDisp);
				if(!spDoc)
					break;	
 				spDoc->SaveAs(COleVariant(cPath2), vtMissing, 
						COleVariant(m_cPassWord), COleVariant(m_cPWWrite), vtMissing, vtMissing, (MSExcel::XlSaveAsAccessMode)1);						
 				spDoc.Release();
				
			}
			break;
		case FILE_TYPE_PPT:
			{
				CComQIPtr<MSPPT::_Presentation> spDoc(lDisp);
 				if(!spDoc)
				break;
				spDoc->SaveAs(_bstr_t(cPath2), MSPPT::ppSaveAsPresentation, Office::msoTrue);
				spDoc.Release();
			}
			break;
		case FILE_TYPE_PDF:
		case FILE_TYPE_UNK:
		default:
			break;
		}
	}catch(...){
		
	}
	
	char cPath[MAX_PATH];
	::GetTempPath(MAX_PATH, cPath);
	strcat(cPath,"DSOWebOffice");
	if (CreateDirectory(cPath, NULL) || GetLastError() == ERROR_ALREADY_EXISTS)
	{
		lstrcat(cPath, "\\");
	} 	
	if(!pTemp2){
		::GetTempFileName(cPath, "", 0, cPath);
		strcat(cPath,"Test.doc");
	}else{
		strcat(cPath, pTemp2);
	}
	
	FILE *fp = fopen(cPath,"r");
	if(fp){
		fclose(fp);
		if(!::DeleteFile (cPath)){
			::GetTempPath(MAX_PATH, cPath);
			strcat(cPath,"DSOWebOffice");

			
			if (CreateDirectory(cPath, NULL) || GetLastError() == ERROR_ALREADY_EXISTS)
			{
				lstrcat(cPath, "\\");
			}

			::GetTempFileName(cPath, "", 0, cPath);
			strcat(cPath,"test001.doc");
		}
	}
 	::CopyFile(cPath2,cPath,TRUE);
 	
	if(!pTemp1)
		m_pHttp->AddPostArguments("FileData", cPath, TRUE);
	else
		m_pHttp->AddPostArguments(pTemp1, cPath, TRUE);
	
	*pbool = TRUE;
	if(pTemp2)	DsoMemFree((void*)(pTemp2));
	if(pTemp1)	DsoMemFree((void*)(pTemp1));
	return S_OK; 
}
STDMETHODIMP CDsoFramerControl::HttpAddPostFile(BSTR strFileID,  BSTR strFileName,  long* pbool)
{
	char * pTemp1 = NULL;
	LPWSTR  pwTemp1;
	if ((strFileID) && (SysStringLen(strFileID) > 0)){
		pwTemp1 = SysAllocString(strFileID);
		pTemp1 = DsoConvertToMBCS(pwTemp1);
	}	
	
	char * pTemp2 = NULL; 
	LPWSTR  pwTemp2;
	if ((strFileName) && (SysStringLen(strFileName) > 0)){
		pwTemp2 = SysAllocString(strFileName);
		pTemp2 = DsoConvertToMBCS(pwTemp2);
	}	
	if(!pTemp1)
		m_pHttp->AddPostArguments("FileData", pTemp2, TRUE);
	else
		m_pHttp->AddPostArguments(pTemp1, pTemp2, TRUE);
	
	*pbool = TRUE;	
	if(pTemp2)	DsoMemFree((void*)(pTemp2));
	if(pTemp1)	DsoMemFree((void*)(pTemp1));
	return S_OK;
} 
STDMETHODIMP CDsoFramerControl::HttpAddPostString(BSTR strName, BSTR strValue, VARIANT_BOOL* pbool)
{
	if(!m_pHttp){
		* pbool = FALSE;
		return 0;
	}
	char * pstrNameTemp1 = NULL;
	LPWSTR  pstrNameTemp2;
	if ((strName) && (SysStringLen(strName) > 0)){
		pstrNameTemp2 = SysAllocString(strName);
		pstrNameTemp1 = DsoConvertToMBCS(pstrNameTemp2);
	}	
	if(!pstrNameTemp1){
		*pbool = FALSE;
		return S_OK;
	}
	char * pstrValueTemp1 = NULL;
	LPWSTR  pstrValueTemp2;
	if ((strValue) && (SysStringLen(strValue) > 0)){
		pstrValueTemp2 = SysAllocString(strValue);
		pstrValueTemp1 = DsoConvertToMBCS(pstrValueTemp2);
	}	
	if(!pstrValueTemp1){
		DsoMemFree((void*)(pstrNameTemp1));
		*pbool = FALSE;
		return S_OK;
	}
	m_pHttp->AddPostArguments(pstrNameTemp1, pstrValueTemp1, FALSE);
	DsoMemFree((void*)(pstrValueTemp1));
	DsoMemFree((void*)(pstrNameTemp1));
	* pbool = TRUE;
	return S_OK; 

}
/*
	HttpPost�ϴ�������
	����ô˽ӿڣ�ocx���Ʊ���Ϊdsoframer.ocx
*/
STDMETHODIMP CDsoFramerControl::HttpPost(BSTR bstr,BSTR* pRet)
{

	// Ask doc object site for the source name...
	if(!m_pHttp){
		* pRet = NULL;
		return S_OK;
	}
	char * pstrNameTemp1 = NULL;
	LPWSTR  pstrNameTemp2;
	if ((bstr) && (SysStringLen(bstr) > 0)){
		pstrNameTemp2 = SysAllocString(bstr);
		pstrNameTemp1 = DsoConvertToMBCS(pstrNameTemp2);
	}	

	if(!pstrNameTemp1 || !pstrNameTemp1[0]){
		delete m_pHttp;
		m_pHttp = NULL;
		* pRet = NULL;
		return S_OK;
	}	
	
	CString strResult;
	DWORD dwRet = 0;
	

	char cHttpURL[1024];
	cHttpURL[0] = 0;
	if(m_cUrl && strlen(m_cUrl)<900){
		if(strncmp(pstrNameTemp1, "/", 1) == 0){		
			 strcpy(cHttpURL,m_cUrl);
			 char * p = strrchr(cHttpURL,'/');
			 if(p) *p = 0;
			 strcat(cHttpURL,pstrNameTemp1); 
		}else if(strncmp(pstrNameTemp1, "./", 2) == 0){  
			 strcpy(cHttpURL,m_cUrl);
			 char * p = strrchr(cHttpURL,'/');
			 if(p) *p = 0;
			 strcat(cHttpURL,&pstrNameTemp1[1]);  
		}else if(strncmp(pstrNameTemp1, "../", 3) == 0){ //
			 strcpy(cHttpURL,m_cUrl);
			 char * p = strrchr(cHttpURL,'/');
			 if(++p) *p = 0;
			 strcat(cHttpURL,pstrNameTemp1);  
		}else{
			strcpy(cHttpURL,pstrNameTemp1);
		}
		DsoMemFree((void*)(pstrNameTemp1));
		pstrNameTemp1 = NULL;
	}
 
	BSTR bstrString = NULL;
	HRESULT hr;
	char *pcBmp = NULL;
	PBYTE pPostBuffer=NULL;
	DWORD dwPostBufferLength=m_pHttp->AllocMultiPartsFormData(pPostBuffer, "--MULTI-PARTS-FORM-DATA-BOUNDARY");
	VARIANT vt;
	if(m_pHttp->GetHttpPost(pPostBuffer,dwPostBufferLength,&vt)){
		try {
				IXMLHTTPRequestPtr pIXMLHTTPRequest = NULL;
				hr=pIXMLHTTPRequest.CreateInstance("Msxml2.XMLHTTP.3.0");
				SUCCEEDED(hr) ? 0 : throw hr;
				if(pstrNameTemp1){
					hr=pIXMLHTTPRequest->open("POST", pstrNameTemp1, false);
				}else{
					hr=pIXMLHTTPRequest->open("POST", cHttpURL, false);
				}

				SUCCEEDED(hr) ? 0 : throw hr;
				
				pIXMLHTTPRequest->setRequestHeader("Content-Type","multipart/form-data; boundary=--MULTI-PARTS-FORM-DATA-BOUNDARY"); 
				

				hr=pIXMLHTTPRequest->send(vt);
				if(SUCCEEDED(hr)){
 					bstrString=pIXMLHTTPRequest->responseText;
					* pRet = bstrString;
				}
				pIXMLHTTPRequest.Release();
				pIXMLHTTPRequest = NULL;
		}catch(...){
			* pRet = NULL;
		}
	}else{
		* pRet = NULL;
	}
	m_pHttp->FreeMultiPartsFormData(pPostBuffer);
	delete m_pHttp;
	m_pHttp = NULL;
	if(pstrNameTemp1)	DsoMemFree((void*)(pstrNameTemp1));
  	return S_OK; 
}
STDMETHODIMP CDsoFramerControl::FtpConnect(BSTR strURL,long lPort, BSTR strUser, BSTR strPwd, long * pbool)
{
 	if(!strURL||strURL[0] ==0)
		return S_OK; 	
	if(!m_pSession){
		m_pSession =  new CInternetSession("DSO");
	}
	if(!m_pSession){
		* pbool = 0;
		return S_OK; 
	}
	try{
 
		m_pFtpConnection = m_pSession->GetFtpConnection((const char *)strURL,
			(const char *)strUser,(const char *)strPwd,lPort);//INTERNET_INVALID_PORT_NUMBER);
		m_bConnect = TRUE;
	return S_OK; 	
	}catch(CInternetException *pEx){
		//pEx->ReportError(MB_ICONEXCLAMATION);
		m_bConnect = FALSE;
		m_pFtpConnection = NULL;
		pEx->Delete();
	return S_OK; 	
	}
	return S_OK; 	
}
STDMETHODIMP CDsoFramerControl::FtpGetFile(BSTR strRemoteFile, BSTR strLocalFile, long * pbool)
{
 	if(!m_pFtpConnection||!m_bConnect)return 0;
	if(!strLocalFile||strLocalFile[0] ==0) return 0;
	if(!strRemoteFile||strRemoteFile[0] ==0) return 0;
	BOOL bUploaded = m_pFtpConnection->GetFile((const char *)strRemoteFile,(const char *)strLocalFile,
		FALSE,FILE_ATTRIBUTE_NORMAL,FTP_TRANSFER_TYPE_BINARY,1);
	return bUploaded?1:0;	return S_OK; 

}
STDMETHODIMP CDsoFramerControl::FtpPutFile(BSTR strLocalFile, BSTR strRemoteFile, long blOverWrite, long * pbool)
{
 	if(!m_pFtpConnection||!m_bConnect)return 0;
	if(!strLocalFile||strLocalFile[0] ==0) return 0;
	if(!strRemoteFile||strRemoteFile[0] ==0) return 0;
	
	BOOL bUploaded = m_pFtpConnection->PutFile((const char *)strLocalFile,
		(const char *)strRemoteFile,FTP_TRANSFER_TYPE_BINARY,1);
	return bUploaded?1:0;
	return S_OK; 

}
STDMETHODIMP CDsoFramerControl::FtpDisConnect(long * pbool)
{
 	m_pFtpConnection->Close();
	m_pFtpConnection = NULL;
	m_bConnect = FALSE;
	if(m_pSession){
		delete m_pSession;
		m_pSession = NULL;
	}

	* pbool =  1;
	return S_OK; 

}
 

STDMETHODIMP CDsoFramerControl::GetFieldValue(BSTR strFieldName, BSTR strCmdOrSheetName,BSTR* strValue)
{
 	IDispatch * lDisp =  m_pDocDispatch;
	if(!lDisp){
		get_ActiveDocument(&lDisp);
	}
	if(!lDisp){
 		return S_OK;
	}
	HRESULT hr;
	USES_CONVERSION;
	try{
		switch(m_nOriginalFileType){
		case FILE_TYPE_WORD:
			{	
				
				CComQIPtr<MSWord::_Document> spDoc(lDisp);
				if(!spDoc)
					break;
				CComPtr<MSWord::Bookmarks> pBkMarks = NULL;
				CComPtr<MSWord::Bookmark> pBkMark = NULL;
				CComPtr<MSWord::Range> pRange = NULL;
//				CComPtr<MSWord::Selection>	spSelection = NULL;
				hr = spDoc->get_Bookmarks(&pBkMarks);
				if(FAILED(hr)||(!pBkMarks)){
					spDoc.Release();
 					return E_NOBOOKMARK;
				}
 
				pBkMark = pBkMarks->Item(COleVariant(W2A(strFieldName)));
				if(!pBkMark)
					return NULL;
				hr = pBkMark->get_Range(&pRange);
				if(FAILED(hr)||(!pRange))
					return NULL;
 				hr = pRange->get_Text(strValue);
 				if(SUCCEEDED(hr)){
				}
				pRange.Release();
				pBkMark.Release();
				pBkMarks.Release();
				spDoc.Release();		
			}
			break;
		case FILE_TYPE_EXCEL:
			{	
				
				
			}
			break;
		case FILE_TYPE_PPT:
		case FILE_TYPE_PDF:
		case FILE_TYPE_UNK:
		default:
			break;
		}
	}catch(...){
		
	}
 	return S_OK;
}
/*
//����Ϊ��ɫ
object.SetFieldValue("BookMarkName,"255","::SETCOLOR::");
//��ɫ
object.SetFieldValue("BookMarkName,"16777215","::SETCOLOR::");
//��ȡ��ɫ
var v = object.SetFieldValue("BookMarkName," ","::GETCOLOR::");
*/
STDMETHODIMP CDsoFramerControl::SetFieldValue(BSTR strFieldName, BSTR strValue,
												 BSTR strCmdOrSheetName,VARIANT_BOOL* pbool)
{
 	IDispatch * lDisp =  m_pDocDispatch;
	if(!lDisp){
		get_ActiveDocument(&lDisp);
	}
	if(!lDisp){
		*pbool = FALSE;
		return S_OK;
	}
	HRESULT hr;
	USES_CONVERSION;
	try{
		switch(m_nOriginalFileType){
		case FILE_TYPE_WORD:
			{	
				
				CComQIPtr<MSWord::_Document> spDoc(lDisp);
				if(!spDoc)
					break;
				CComPtr<MSWord::Bookmarks> pBkMarks = NULL;
				CComPtr<MSWord::Bookmark> pBkMark = NULL;
				CComPtr<MSWord::Range> pRange = NULL;
				CComPtr<MSWord::Selection>	spSelection = NULL;
				hr = spDoc->get_Bookmarks(&pBkMarks);
				if(FAILED(hr)||(!pBkMarks)){
					spDoc.Release();
					*pbool = FALSE;
					return E_NOBOOKMARK;
				}
				if(0==wcsicmp(strCmdOrSheetName,L"::ADDMARK::")){
		 			CComPtr<MSWord::_Application> spApp(spDoc->GetApplication());				
					hr = spApp->get_Selection(&spSelection);
					if(FAILED(hr)||(!spSelection)) return E_UNKNOW;
					hr = spSelection->get_Range(&pRange);
					if(FAILED(hr)||(!pRange)) return E_UNKNOW;
					VARIANT vt;
					vt.vt = VT_DISPATCH;
					vt.pdispVal = pRange;
					//hr = pBkMarks->raw_Add(A2W(strFieldName),&vt,&pBkMark);
					//if(FAILED(hr)||(!pBkMark)) return E_UNKNOW;
					pBkMarks->Add(strFieldName,&vt);
					pRange.Release();
					spSelection.Release();
					spApp.Release(); 
				}else if(0==wcsicmp(strCmdOrSheetName,L"::SETCOLOR::")){  
					hr = pBkMarks->raw_Item(COleVariant(W2A(strFieldName)),&pBkMark);
					if(FAILED(hr)) return E_UNKNOW;	
					hr = pBkMark->get_Range(&pRange);
					if(pRange){
						CComPtr<MSWord::_Font>	pFont = NULL;
						hr = pRange->get_Font(&pFont);
						if(FAILED(hr)) return E_UNKNOW;	
						long iColor = atol(W2A(strValue));
						pFont->PutColor((enum MSWord::WdColor)iColor);
						pFont = NULL;
						pRange = NULL;
					} 			 
					pBkMark = NULL;
					return 0;
				}else if(0==wcsicmp(strCmdOrSheetName,L"::GETCOLOR::")){  
					hr = pBkMarks->raw_Item(COleVariant(W2A(strFieldName)),&pBkMark);
					if(FAILED(hr)) return E_UNKNOW;	
					hr = pBkMark->get_Range(&pRange);
					long iColor = 0;
					if(pRange){
						CComPtr<MSWord::_Font>	pFont = NULL;
						hr = pRange->get_Font(&pFont);
						if(FAILED(hr)) return E_UNKNOW;	
						pFont->get_Color((enum MSWord::WdColor*)&iColor);
						//pFont->PutColor((enum MSWord::WdColor)iColor);
						pFont = NULL;
						pRange = NULL;
						return iColor;
					} 			 
					pBkMark = NULL;
					return E_UNKNOW;
				}
				else if(0==wcsicmp(strCmdOrSheetName,L"::DELMARK::")){  
					hr = pBkMarks->raw_Item(COleVariant(W2A(strFieldName)),&pBkMark);
					if(FAILED(hr)) return E_UNKNOW;	
											
					hr = pBkMark->get_Range(&pRange);
					if(pRange){
						pRange->Delete();
						pRange = NULL;
					}
					pBkMark->Delete();				 
					pBkMark = NULL;
					return 0;
				}else if(0==wcsicmp(strCmdOrSheetName,L"::GETMARK::")){ 
					hr = pBkMarks->raw_Item(COleVariant(W2A(strFieldName)),&pBkMark);
					if(FAILED(hr)) return E_UNKNOW;	
											
					hr = pBkMark->get_Range(&pRange);
					if(pRange){
						pRange->Select();
						pRange = NULL;
					}
 					pBkMark = NULL;
					return 0;
				}
				pBkMark = pBkMarks->Item(COleVariant(W2A(strFieldName)));
				if(!pBkMark){
					pBkMarks.Release();
					spDoc.Release();
					*pbool = FALSE;
					return E_NOBOOKMARK;
				}						
				hr = pBkMark->get_Range(&pRange);
				if(FAILED(hr)||(!pRange)){
					pBkMark.Release();
					pBkMarks.Release();
					spDoc.Release();
					*pbool = FALSE;
					return E_NOBOOKMARK;
				}			
				if(0==wcsicmp((strCmdOrSheetName),L"::FILE::")){
					//pRange->InsertFile(strValue);
					/////////////////////////////////////////////////
					if(SUCCEEDED(hr)){
						char szPath[MAX_PATH];
						szPath[0] = 0;
						FILE * fp = NULL;
						char temp[1024];
						UINT nBytesRead = 0;
						if(_wcsnicmp(strValue, L"HTTP", 4) == 0  || _wcsnicmp(strValue, L"FTP", 3) == 0){
							::GetTempPath(MAX_PATH, szPath);
							strcat(szPath,"DSOWebOffice");
							
							if (CreateDirectory(szPath, NULL) || GetLastError() == ERROR_ALREADY_EXISTS)
							{
								lstrcat(szPath, "\\");
							}
							::GetTempFileName(szPath, "", 0, szPath);
							//theApp.GetTempFileName(szPath);
							CInternetSession session("DSO");
							CStdioFile * pFile = NULL;
							pFile = session.OpenURL(W2A(strValue),0,
								INTERNET_FLAG_TRANSFER_BINARY|INTERNET_FLAG_KEEP_CONNECTION|INTERNET_FLAG_RELOAD|
								INTERNET_FLAG_DONT_CACHE);
							if(!pFile){
								return 0;
							}
							if((fp = fopen(szPath,"wb")) == NULL){
								pFile->Close();
								delete pFile;
								return 0;
							}
							DWORD iLen = 0;
							while(iLen = pFile->Read(temp, 1024)){
								fwrite(temp, 1, iLen, fp);
								fflush(fp);
							}
							fclose(fp);
							if(pFile){
								pFile->Close();
								delete pFile;
							}
							pRange->InsertFile(szPath,&vtMissing,&vtMissing,&vtMissing,&vtMissing );
						}else {
							pRange->InsertFile(strValue,&vtMissing,&vtMissing,&vtMissing,&vtMissing );
						}						
					} 	
					////////////////////////////////////////////////////
				}else if(0 == wcsnicmp((strCmdOrSheetName),L"::JPG::",7)){
					CComPtr<MSWord::InlineShapes>	spLineShapes = NULL;							
					CComPtr<MSWord::InlineShape>	spInlineShape = NULL;						
					CComPtr<MSWord::Shape>	spShape = NULL;
					CComPtr<MSWord::WrapFormat>	wrap = NULL;					 
					hr = pRange->get_InlineShapes(&spLineShapes);
					if(FAILED(hr)||(!pRange))
						return E_UNKNOW;
					hr = spLineShapes->raw_AddPicture(strValue,&vtMissing,&vtMissing,&vtMissing,&spInlineShape);
					if(pRange){
						spInlineShape->raw_ConvertToShape(&spShape);
						if(spShape){
							hr = spShape->get_WrapFormat(&wrap); 
							char * p = BSTR2char(strCmdOrSheetName);
							char *p1 = p;
							p += 7;
							if(p){
								wrap->put_Type((MSWord::WdWrapType)atol(p));
							}else{
								wrap->put_Type((MSWord::WdWrapType)5);
							}
							if(p1) free(p1);
							wrap->put_AllowOverlap(-1);
							wrap = NULL;
						}
						spShape = NULL;
					}
 					spLineShapes = NULL;
					spInlineShape = NULL;
				}else{				  		
					pRange->Select();
					if(FAILED(hr)||(!pRange))
						return E_UNKNOW;
					pRange->Text = strValue;
					VARIANT vt;
					vt.vt = VT_DISPATCH;
					vt.pdispVal = pRange;
					pBkMarks->Add(strFieldName,&vt);
					//pBkMarks->raw_Add(A2W(strFieldName),&pRange,&pBkMark);
					pRange->MoveStart();
				}
				pRange.Release();
				pBkMark.Release();
				pBkMarks.Release();
				spDoc.Release();		
			}
			break;
		case FILE_TYPE_EXCEL:
			{	
				
				
			}
			break;
		case FILE_TYPE_PPT:
		case FILE_TYPE_PDF:
		case FILE_TYPE_UNK:
		default:
			break;
		}
	}catch(...){
		
	}
	*pbool = TRUE; 
	return S_OK;
}
/*--------------------------ProDoc--------------------------
lProOrUn:1,�����ĵ���0,�Ᵽ���ĵ�
lProType: �������ͣ�
    wdNoProtection = -1,
    wdAllowOnlyRevisions = 0,
    wdAllowOnlyComments = 1,
    wdAllowOnlyFormFields = 2
strProPWD:��������
------------------------------------------------------------*/
STDMETHODIMP CDsoFramerControl::ProtectDoc(long lProOrUn,long lProType,  BSTR strProPWD, VARIANT_BOOL * pbool)
{
 	IDispatch * lDisp =  m_pDocDispatch;
	if(!lDisp){
		get_ActiveDocument(&lDisp);
	}
	if(!lDisp){
		*pbool = FALSE;
		return S_OK;
	}
	HRESULT hr;
	USES_CONVERSION;
	try{
		switch(m_nOriginalFileType){
		case FILE_TYPE_WORD:
			{				
				CComQIPtr<MSWord::_Document> spDoc(lDisp);
				if(!spDoc)
					break;
				if(0 == lProOrUn){
					spDoc->Unprotect(COleVariant(W2A(strProPWD)));		
				}else{
 					spDoc->Protect((MSWord::WdProtectionType)lProType,&vtMissing,COleVariant(W2A(strProPWD)));
				}
				spDoc.Release();		
			}
			break;
		case FILE_TYPE_EXCEL:
			{	
				if(0 == lProOrUn){
					CComQIPtr<MSExcel::_Workbook> spDoc(lDisp);	
					if(!spDoc)
						break;
					CComPtr<MSExcel::Sheets> spSheets;
					CComQIPtr<MSExcel::_Worksheet> spSheet;
					_variant_t vtPwd(strProPWD);
					hr = spDoc->get_Worksheets(&spSheets);
					if((FAILED(hr))||(!spSheets))
						return E_NOSHEET;
					long lCount = spSheets->GetCount();
					while(lCount){
						_variant_t vt(lCount);
						spSheet = spSheets->GetItem(vt);
						if(!spSheets)
							return E_NOSHEET;
						spSheet->Unprotect(vtPwd);
						spSheet = (MSExcel::_Worksheet *)NULL;
						lCount--;
					}					
				}else{
					CComQIPtr<MSExcel::_Workbook> spDoc(lDisp);
					if(!spDoc)
						break;
 					CComPtr<MSExcel::Sheets> spSheets;
					CComQIPtr<MSExcel::_Worksheet> spSheet;
					_variant_t vtPwd(strProPWD);
					hr = spDoc->get_Worksheets(&spSheets);
					if((FAILED(hr))||(!spSheets))
						return E_NOSHEET;
					long lCount = spSheets->GetCount();
					while(lCount){
						_variant_t vt(lCount);
						spSheet = spSheets->GetItem(vt);
						if(!spSheets)
							return E_NOSHEET;
						spSheet->Protect(vtPwd);
						spSheet = (MSExcel::_Worksheet *)NULL;
						lCount--;
					}					

				}
				
			}
			break;
		case FILE_TYPE_PPT:
		case FILE_TYPE_PDF:
		case FILE_TYPE_UNK:
		default:
			break;
		}
	}catch(...){
		
	}
	*pbool = TRUE; 
	return S_OK;
}

STDMETHODIMP CDsoFramerControl::SetMenuDisplay(long lMenuFlag, VARIANT_BOOL* pbool)
{
	/*
#define MNU_NEW                         0x01
#define MNU_OPEN                        0x02
#define MNU_CLOSE                       0x04
#define MNU_SAVE                        0x08
#define MNU_SAVEAS                      0x16
#define MNU_HONGTOU                     0x512
#define MNU_PGSETUP                     0x64
#define MNU_PRINT                       0x256
#define MNU_PROPS                       0x32
#define MNU_PRINTPV                     0x128
	*/
	m_wFileMenuFlags = lMenuFlag;
	* pbool = TRUE;
	return S_OK;
}

STDMETHODIMP CDsoFramerControl::ShowRevisions(long nNewValue, VARIANT_BOOL* pbool)
{
 	IDispatch * lDisp =  m_pDocDispatch;
	if(!lDisp){
		get_ActiveDocument(&lDisp);
	}
	if(!lDisp){
		*pbool = FALSE;
		return S_OK;
	}
	HRESULT hr;
	USES_CONVERSION;
	try{
		switch(m_nOriginalFileType){
		case FILE_TYPE_WORD:
			{				
				CComQIPtr<MSWord::_Document> spDoc(lDisp);
				if(!spDoc)
					break;
				spDoc->ShowRevisions = nNewValue;
				spDoc.Release();		
			}
			break;
		case FILE_TYPE_EXCEL:
			{	
 				CComQIPtr<MSExcel::_Workbook> spDoc(lDisp);	
				if(!spDoc)
				break;
				spDoc->HighlightChangesOnScreen = nNewValue; 		
				spDoc.Release();		
			}
			break;
		case FILE_TYPE_PPT:
		case FILE_TYPE_PDF:
		case FILE_TYPE_UNK:
		default:
			break;
		}
	}catch(...){
		
	}
	*pbool = TRUE; 
	return S_OK;
}
/*
pPos = 0 //��ǰ���λ��
1;�ļ���ͷ
2;�ļ�ĩβ

if nCurPos | 0x8 ��ʾ�������ͼƬ
					if(lPos == 1){
						spDoc->raw_Range(COleVariant((long)0) , COleVariant((long)0) ,&spRang);
						if(!spRang)
							break;
						spRang->Select();
					}else if(lPos == 2){
						hr = spDoc->get_Content(&spRang);
						if((FAILED(hr))||(!spRang))
							break;
						long lStart = spRang->End - 1;
						spRang = NULL;
						spDoc->raw_Range(COleVariant(lStart), COleVariant(vtMissing), &spRang);
						if(!spRang)
							break;
						spRang->Select();		
					}
*/
STDMETHODIMP CDsoFramerControl::InSertFile(BSTR strFieldPath, long lPos, VARIANT_BOOL* pbool)
{
 	IDispatch * lDisp =  m_pDocDispatch;
	if(!lDisp){
		get_ActiveDocument(&lDisp);
	}
	if(!lDisp){
		*pbool = FALSE;
		return S_OK;
	}
	HRESULT hr;
	USES_CONVERSION;
	try{
		switch(m_nOriginalFileType){
		case FILE_TYPE_WORD:
			{					
				CComQIPtr<MSWord::_Document> spDoc(lDisp);
				if(!spDoc)
					break;
				CComPtr<MSWord::_Application> spApp(spDoc->GetApplication());
				CComPtr<MSWord::Range> pRange;
				CComPtr<MSWord::Selection> spSelection;
				spSelection = NULL;
				hr = spApp->get_Selection(&spSelection);
				if(SUCCEEDED(hr)){
					if(0x01 & lPos){
						spSelection->HomeKey(COleVariant((long)6),COleVariant((long)0));
					}else if(0x02 & lPos){
						spSelection->EndKey(COleVariant((long)6),COleVariant((long)0));
					}
					if(0x08 & lPos){
						hr = spSelection->get_Range(&pRange);
						if(FAILED(hr)||(!pRange))
							return E_UNKNOW;

						CComPtr<MSWord::InlineShapes>	spLineShapes = NULL;							
						CComPtr<MSWord::InlineShape>	spInlineShape = NULL;						
						CComPtr<MSWord::Shape>	spShape = NULL;
						CComPtr<MSWord::WrapFormat>	wrap = NULL;
						 
						hr = pRange->get_InlineShapes(&spLineShapes);
						if(FAILED(hr)||(!pRange))
							return E_UNKNOW;
						hr = spLineShapes->raw_AddPicture(strFieldPath,&vtMissing,&vtMissing,&vtMissing,&spInlineShape);
						if(pRange){
							spInlineShape->raw_ConvertToShape(&spShape);
							if(spShape){
								hr = spShape->get_WrapFormat(&wrap); 
								wrap->put_Type((MSWord::WdWrapType)5);
								wrap->put_AllowOverlap(-1);
								wrap = NULL;
							}
							spShape = NULL;
						}
 						spLineShapes = NULL;
						spInlineShape = NULL;
						return S_OK;
					}
					if(SUCCEEDED(hr)){
							if(_wcsnicmp(strFieldPath, L"HTTP", 4) == 0  || _wcsnicmp(strFieldPath, L"FTP", 3) == 0){
							char szPath[MAX_PATH];
							szPath[0] = 0;
							FILE * fp = NULL;
							char temp[1024];
							UINT nBytesRead = 0;
							::GetTempPath(MAX_PATH, szPath);
							strcat(szPath,"DSOWebOffice");
							
							
							if (CreateDirectory(szPath, NULL) || GetLastError() == ERROR_ALREADY_EXISTS)
							{
								lstrcat(szPath, "\\");
							}

							::GetTempFileName(szPath, "", 0, szPath);
							//theApp.GetTempFileName(szPath);
							CInternetSession session("DSO");
							CStdioFile * pFile = NULL;
							pFile = session.OpenURL(W2A(strFieldPath),0,
								INTERNET_FLAG_TRANSFER_BINARY|INTERNET_FLAG_KEEP_CONNECTION|INTERNET_FLAG_RELOAD|
								INTERNET_FLAG_DONT_CACHE);
							if(!pFile){
								return 0;
							}
							if((fp = fopen(szPath,"wb")) == NULL){
								pFile->Close();
								delete pFile;
								return 0;
							}
							DWORD iLen = 0;
							while(iLen = pFile->Read(temp, 1024)){
								fwrite(temp, 1, iLen, fp);
								fflush(fp);
							}
							fclose(fp);
							if(pFile){
								pFile->Close();
								delete pFile;
							}
							spSelection->InsertFile(szPath,&vtMissing,&vtMissing,&vtMissing,&vtMissing );
						}else{
							spSelection->InsertFile(strFieldPath,&vtMissing,&vtMissing,&vtMissing,&vtMissing );
						}

					}

				}
				spDoc.Release();		
			}
			break;
		case FILE_TYPE_EXCEL:
			{	
				
				
			}
			break;
		case FILE_TYPE_PPT:
		case FILE_TYPE_PDF:
		case FILE_TYPE_UNK:
		default:
			break;
		}
	}catch(...){
		
	}
	*pbool = TRUE; 
	return S_OK; 
}
STDMETHODIMP CDsoFramerControl::ClearFile()
{						
 	char szPath[MAX_PATH];
 	::GetTempPath(MAX_PATH, szPath);
	strcat(szPath,"DSOWebOffice\\");
	char pcFile[MAX_PATH];
	strcpy(pcFile, szPath);
	strcat(pcFile, "*");
	HANDLE hFindFile = NULL;
	WIN32_FIND_DATA FindFileData;
 
 	if((hFindFile = FindFirstFile(pcFile, &FindFileData))!=NULL){
		while(FindNextFile(hFindFile, &FindFileData)){
        	if(FindFileData.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY){
        		continue;
        	}else{
				strcpy(pcFile,szPath);
 				strcat(pcFile, FindFileData.cFileName);
				::DeleteFile(pcFile);
        	}
		}
		FindClose(hFindFile);
	}
	return S_OK;
}
STDMETHODIMP CDsoFramerControl::LoadOriginalFile( VARIANT strFieldPath,  VARIANT strFileType,  long* pbool)
{
	LPWSTR    pwszFileType = LPWSTR_FROM_VARIANT(strFileType);
	
	LPWSTR    pwszDocument  = LPWSTR_FROM_VARIANT(strFieldPath);
//	CreateNew(BSTR ProgIdOrTemplate)
	if(!pwszDocument || pwszDocument[0] ==0){
		if(strFileType.vt == VT_EMPTY){
			CreateNew(L"Word.Document");
			* pbool = m_nOriginalFileType;	
			return S_OK;
		}
		if(_wcsicmp(pwszFileType,L"doc") == 0 ){
			CreateNew(L"Word.Document");
		}else if(_wcsicmp(pwszFileType,L"wps") == 0 ){
			CreateNew(L"Wps.Document");
		}else if(_wcsicmp(pwszFileType,L"xls") == 0 ){
			CreateNew(L"Excel.Sheet");
		}else if(_wcsicmp(pwszFileType,L"ppt") == 0 ){
			//CreateNew(L"PowerPoint.Slide");
			CreateNew(L"PowerPoint.Show");
		}else{
			CreateNew(pwszFileType);
		}
	}else{
 			Open(strFieldPath, vtMissing, strFileType, vtMissing, vtMissing);
	}

	* pbool = m_nOriginalFileType;	
 	return S_OK;
}
//Excel File Format 
/*
enum XlFileFormat
{
    xlAddIn = 18,
    xlCSV = 6,
    xlCSVMac = 22,
    xlCSVMSDOS = 24,
    xlCSVWindows = 23,
    xlDBF2 = 7,
    xlDBF3 = 8,
    xlDBF4 = 11,
    xlDIF = 9,
    xlExcel2 = 16,
    xlExcel2FarEast = 27,
    xlExcel3 = 29,
    xlExcel4 = 33,
    xlExcel5 = 39,
    xlExcel7 = 39,
    xlExcel9795 = 43,
    xlExcel4Workbook = 35,
    xlIntlAddIn = 26,
    xlIntlMacro = 25,
    xlWorkbookNormal = -4143,
    xlSYLK = 2,
    xlTemplate = 17,
    xlCurrentPlatformText = -4158,
    xlTextMac = 19,
    xlTextMSDOS = 21,
    xlTextPrinter = 36,
    xlTextWindows = 20,
    xlWJ2WD1 = 14,
    xlWK1 = 5,
    xlWK1ALL = 31,
    xlWK1FMT = 30,
    xlWK3 = 15,
    xlWK4 = 38,
    xlWK3FM3 = 32,
    xlWKS = 4,
    xlWorks2FarEast = 28,
    xlWQ1 = 34,
    xlWJ3 = 40,
    xlWJ3FJ3 = 41,
    xlUnicodeText = 42,
    xlHtml = 44
};
Word: Type
enum WdSaveFormat
{
    wdFormatDocument = 0,
    wdFormatTemplate = 1,
    wdFormatText = 2,
    wdFormatTextLineBreaks = 3,
    wdFormatDOSText = 4,
    wdFormatDOSTextLineBreaks = 5,
    wdFormatRTF = 6,
    wdFormatUnicodeText = 7,
    wdFormatEncodedText = 7,
    wdFormatHTML = 8
};
PPT:
enum PpSaveAsFileType
{
    ppSaveAsPresentation = 1,
    ppSaveAsPowerPoint7 = 2,
    ppSaveAsPowerPoint4 = 3,
    ppSaveAsPowerPoint3 = 4,
    ppSaveAsTemplate = 5,
    ppSaveAsRTF = 6,
    ppSaveAsShow = 7,
    ppSaveAsAddIn = 8,
    ppSaveAsPowerPoint4FarEast = 10,
    ppSaveAsDefault = 11,
    ppSaveAsHTML = 12,
    ppSaveAsHTMLv3 = 13,
    ppSaveAsHTMLDual = 14,
    ppSaveAsMetaFile = 15,
    ppSaveAsGIF = 16,
    ppSaveAsJPG = 17,
    ppSaveAsPNG = 18,
    ppSaveAsBMP = 19
};
  */
STDMETHODIMP CDsoFramerControl::SaveAs( VARIANT strFileName,  VARIANT dwFileFormat, long* pbool)
{	

 	IDispatch * lDisp =  m_pDocDispatch;
	if(!lDisp){
		get_ActiveDocument(&lDisp);
		if(!lDisp){
			*pbool = FALSE;
			return S_OK;
		}
	}
	HRESULT hr;
	USES_CONVERSION;
	* pbool = 1;
	try{
		switch(m_nOriginalFileType){
		case FILE_TYPE_WORD:
			{					
				CComQIPtr<MSWord::_Document> spDoc(lDisp);
				if(!spDoc)
					break;
				char cPath[MAX_PATH];
				::GetTempPath(MAX_PATH, cPath);
				strcat(cPath,"DSOWebOffice");
				if (CreateDirectory(cPath, NULL) || GetLastError() == ERROR_ALREADY_EXISTS)
				{
					lstrcat(cPath, "\\");
				}
				::GetTempFileName(cPath, "", 0, cPath);
				strcat(cPath,"Test.doc");
				LPWSTR pwszPrinter = LPWSTR_FROM_VARIANT(strFileName);
				spDoc->SaveAs(COleVariant(cPath),COleVariant(dwFileFormat));
				::DeleteFileW(pwszPrinter);
 				spDoc->SaveAs(COleVariant(strFileName),COleVariant(dwFileFormat));
				::DeleteFile(cPath);
				spDoc.Release();		
			}
			break;
		case FILE_TYPE_EXCEL:
			{	
				CComQIPtr<MSExcel::_Workbook> spDoc(lDisp);	
				if(!spDoc)
					break;
					   
			    long iType = LONG_FROM_VARIANT(dwFileFormat, 0);
				if(iType){
 						spDoc->SaveAs(COleVariant(strFileName),COleVariant(dwFileFormat), 
						COleVariant(m_cPassWord), COleVariant(m_cPWWrite), 
						vtMissing, vtMissing, MSExcel::xlNoChange,COleVariant((long)3));
				}else{
 					spDoc->SaveAs(COleVariant(strFileName),vtMissing, 
						COleVariant(m_cPassWord), COleVariant(m_cPWWrite), 
						vtMissing, vtMissing, MSExcel::xlNoChange,COleVariant((long)3));
				} 
			}
			break;
		case FILE_TYPE_PPT:
			{
				CComQIPtr<MSPPT::_Presentation> spDoc(lDisp);
 				if(!spDoc)
				break;
				spDoc->SaveAs(_bstr_t(strFileName), MSPPT::ppSaveAsPresentation, Office::msoTrue);
			}
			break;
		case FILE_TYPE_PDF:
		case FILE_TYPE_UNK:
		default:
				Save(strFileName, vtMissing, vtMissing, vtMissing);
			break;
		}
	}catch(...){
		* pbool = 0;	
	}
 	return S_OK;
}
STDMETHODIMP CDsoFramerControl::DeleteLocalFile(BSTR strFilePath)
{
	USES_CONVERSION;

	::DeleteFile(W2A(strFilePath));
	return S_OK;
}
STDMETHODIMP CDsoFramerControl::GetTempFilePath(BSTR* strValue)
{
	CString strResult;
	// TODO: Add your dispatch handler code here
	char strTempFile[MAX_PATH];
	if(::GetTempPath(MAX_PATH,strTempFile)){
		if(::GetTempFileName(strTempFile,"~dso",0,strTempFile))
			strResult = strTempFile;
	}

	*strValue =  strResult.AllocSysString();
	return S_OK;
}

/*
wdMasterView��wdNormalView��wdOutlineView��wdPrintPreview��wdPrintView �� wdWebView��Long ���ͣ��ɶ�д */
STDMETHODIMP CDsoFramerControl::ShowView(long dwViewType, long * pbool)
{
	*pbool = FALSE; 	
	IDispatch * lDisp =  m_pDocDispatch;
	if(!lDisp){
		get_ActiveDocument(&lDisp);
	}
	if(!lDisp){
		return S_OK;
	}
	HRESULT hr;
	try{
		switch(m_nOriginalFileType){
		case FILE_TYPE_WORD:
			{					
				CComQIPtr<MSWord::_Document> spDoc(lDisp);
				if(!spDoc)
					break;
//				CComPtr<MSWord::_Application> app(spDoc->GetApplication()); 
				CComPtr<MSWord::Window> win;//app(spDoc->GetApplication()); 
				CComPtr<MSWord::Pane> pane;
				CComPtr<MSWord::View> view;
				win = spDoc->GetActiveWindow();
				if(!win)
					break;
				pane = win->GetActivePane();
				if(!pane)
					break;
				view = pane->GetView();
				if(!view)
					break;
				view->put_Type((enum MSWord::WdViewType)dwViewType);
				*pbool = TRUE;
			}
			break;
		case FILE_TYPE_EXCEL:
			{	
				
				
			}
			break;
		case FILE_TYPE_PPT:
		case FILE_TYPE_PDF:
		case FILE_TYPE_UNK:
		default:
			break;
		}
	}catch(...){
		
	}
 

	return S_OK; 
}
/*
if strLocalFile == NULL then Create Temp File and return TempFile's Path
*/
STDMETHODIMP CDsoFramerControl::DownloadFile( BSTR strRemoteFile, BSTR strLocalFile, BSTR* strValue)
{
	CString strRet = "OK";
	if(!LooksLikeHTTP(strRemoteFile) && !LooksLikeFTP(strRemoteFile)){
		strRet = "RemoteFile Error";
		* strValue = strRet.AllocSysString();
		return S_OK;
	}	
	if(strLocalFile && 0 != strLocalFile[0] && !LooksLikeLocalFile(strLocalFile)){
		strRet = "LocalFile Error";
		* strValue = strRet.AllocSysString();
		return S_OK;
	}	
	if(!strLocalFile || 0 == strLocalFile[0] ){
		CString strLoFile;
		char cPath[MAX_PATH];
		::GetTempPath(MAX_PATH, cPath);
		strcat(cPath,"DSOWebOffice");
		if (CreateDirectory(cPath, NULL) || GetLastError() == ERROR_ALREADY_EXISTS)
		{
			lstrcat(cPath, "\\");
		}
		::GetTempFileName(cPath, "", 0, cPath);
		strcat(cPath,"test001.doc");
		strLoFile = cPath;
		BSTR strTemp = strLoFile.AllocSysString();
		URLDownloadFile(NULL,strRemoteFile,strTemp);
		strRet = cPath;
		::SysFreeString(strTemp);
	}else{
		URLDownloadFile(NULL,strRemoteFile,strLocalFile);
	}
	* strValue = strRet.AllocSysString();
	return S_OK;
}

STDMETHODIMP CDsoFramerControl::GetRevCount(long * pbool)
{	
 	IDispatch * lDisp =  m_pDocDispatch;
	if(!lDisp){
		get_ActiveDocument(&lDisp);
	}
	if(!lDisp){
		*pbool = FALSE;
		return S_OK;
	}
	HRESULT hr;
	USES_CONVERSION;
	* pbool = 0;
	try{
		switch(m_nOriginalFileType){
		case FILE_TYPE_WORD:
			{					
				CComQIPtr<MSWord::_Document> spDoc(lDisp);
				if(!spDoc)
					break;
				CComPtr<MSWord::Revisions> spRevs;
				hr = spDoc->get_Revisions(&spRevs);
				if((FAILED(hr))||(!spRevs)){
						return S_OK;
				}
				*pbool = spRevs->GetCount();
				spRevs.Release();
				spDoc.Release();		
			}
			break;
		case FILE_TYPE_EXCEL:
			{	
			}
			break;
		case FILE_TYPE_PPT:
			{
			}
			break;
		case FILE_TYPE_PDF:
		case FILE_TYPE_UNK:
		default:
 			break;
		}
	}catch(...){
		* pbool = 0;	
	}
 	return S_OK;
}
/*
BSTR GetRevInfo(long lIndex ,long lType) 
*/
STDMETHODIMP CDsoFramerControl::GetRevInfo(long lIndex, long lType, BSTR* pbool)
{
	if(lType<0 || lType > 3)
		return S_OK;
 	IDispatch * lDisp =  m_pDocDispatch;
	if(!lDisp){
		get_ActiveDocument(&lDisp);
	}
	if(!lDisp){
		*pbool = NULL;
		return S_OK;
	}
	HRESULT hr;
	USES_CONVERSION;
	* pbool = NULL;
	CString strRet;
	try{
		switch(m_nOriginalFileType){
		case FILE_TYPE_WORD:
			{					
				CComQIPtr<MSWord::_Document> spDoc(lDisp);
				if(!spDoc)
					break;
				CComPtr<MSWord::Revisions> spRevs;
				CComPtr<MSWord::Revision> spRev;
				hr = spDoc->get_Revisions(&spRevs);
				if((FAILED(hr))||(!spRevs)){
						return S_OK;
				}
				long dwCount = spRevs->GetCount();
				int i = lIndex;
				if(i < 1) i =1;
				if(i > dwCount) i = dwCount;
				hr = spRevs->raw_Item(i,&spRev);
				if((FAILED(hr))||(!spRev)){
						return S_OK;
				}
				if(lType == 0){
					spRev->get_Author(pbool);
					 
				}else if(lType == 1){
					double spDate = 0;
					spRev->get_Date(&spDate);
					COleDateTime dwTime(spDate);
					 
 					strRet = dwTime.Format(_T("%Y-%m-%d %H:%M:%S"));
				  
					*pbool = strRet.AllocSysString();
				}else if(lType == 2){
					MSWord::WdRevisionType dwType;
					spRev->get_Type(&dwType);
					strRet.Format("%d",dwType);
					*pbool = strRet.AllocSysString();
					 
				}else if(lType == 3){
					CComPtr<MSWord::Range> pRange = NULL;
					spRev->get_Range(&pRange);
					if((FAILED(hr))||(!pRange)){
						return S_OK;
					}
					pRange->get_Text(pbool);
				}
				spRevs.Release();
				spDoc.Release();		
			}
			break;
		case FILE_TYPE_EXCEL:
			{	
			}
			break;
		case FILE_TYPE_PPT:
			{
			}
			break;
		case FILE_TYPE_PDF:
		case FILE_TYPE_UNK:
		default:
 			break;
		}
	}catch(...){
		* pbool = NULL;	
	}
 	return S_OK;	
}
/*
1.
	SetValue("password","::DOCPROP:PassWord");
2.
	SetValue("password","::DOCPROP:WritePW");
-127:
*/
STDMETHODIMP CDsoFramerControl::SetValue(BSTR strValue, BSTR strName, long* pbool)
{
 	IDispatch * lDisp =  m_pDocDispatch;
	if(!lDisp){
		get_ActiveDocument(&lDisp);
	}
	if(!lDisp){
		*pbool = NULL;
		return S_OK;
	}
	HRESULT hr;
	USES_CONVERSION;
	* pbool = 0;
	char *pValue = NULL;
	char *pName = NULL;
	pValue = BSTR2char(strValue);
	pName = BSTR2char(strName);
 	try{
		switch(m_nOriginalFileType){
		case FILE_TYPE_WORD:
			{					
				CComQIPtr<MSWord::_Document> spDoc(lDisp);
				if(!spDoc)
					break;
 				if(0 == strcmp(pName,"::DOCPROP:PassWord")){
					spDoc->put_Password(strValue);
				}else if(0 == strcmp(pName,"::DOCPROP:WritePW")){
					spDoc->put_WritePassword(strValue);
				}else{
					* pbool = -1;
				}
				spDoc.Release();		
			}
			break;
		case FILE_TYPE_EXCEL:
			{	
				CComQIPtr<MSExcel::_Workbook> spDoc(lDisp);	
				if(!spDoc)
					break;
				if(0 == strcmp(pName,"::DOCPROP:PassWord")){
					if(pValue)	
						strcpy(m_cPassWord,pValue);
				}else if(0 == strcmp(pName,"::DOCPROP:WritePW")){
					if(pValue)	
						strcpy(m_cPWWrite,pValue);
 				}else{
					* pbool = -1;
				}			
			}
			break;
		case FILE_TYPE_PPT:
			{
			}
			break;
		case FILE_TYPE_PDF:
		case FILE_TYPE_UNK:
		default:
 			break;
		}
	}catch(...){
		* pbool = -127;	
	}
	if(pValue) free(pValue);
	if(pName) free(pName);
 	return S_OK;	
}
/*
strVarName: 
strVlaue:
lOpt: 

return:
0:OK
-127:
*/
STDMETHODIMP CDsoFramerControl::SetDocVariable(BSTR strVarName, BSTR strValue, long lOpt, long* pbool) 
{
 	IDispatch * lDisp =  m_pDocDispatch;
	if(!lDisp){
		get_ActiveDocument(&lDisp);
	}
	if(!lDisp){
		*pbool = NULL;
		return S_OK;
	}
	HRESULT hr;
	USES_CONVERSION;
	* pbool = 0;
	char *pValue = NULL;
 	pValue = BSTR2char(strValue);
  	try{
		switch(m_nOriginalFileType){
		case FILE_TYPE_WORD:
			{					
				CComQIPtr<MSWord::_Document> spDoc(lDisp);
				if(!spDoc)
					break;
 				CComPtr<MSWord::Variables> spVars;
				CComPtr<MSWord::Variable> spVar;
				hr = spDoc->get_Variables(&spVars);
				if(FAILED(hr)||(!spVars))
					return NULL;
				hr = spVars->raw_Item(&_variant_t(strVarName),&spVar);
				if(FAILED(hr)){
					if(lOpt & 0x02){
						hr = spVars->raw_Add(strVarName,COleVariant(pValue), &spVar);
						if(FAILED(hr)||(!spVar)){
							* pbool = -127;
							return S_OK;
						}else if(!(lOpt & 0x01)){
							* pbool = -127;
							return S_OK;
						}
					}
				}else{
					spVar->put_Value(strValue);
				}

				if((lOpt & 0x01) && spVar){
					CComPtr<MSWord::Fields> spFields;
					CComPtr<MSWord::Field> spField;
					hr = spDoc->get_Fields(&spFields);
					if(FAILED(hr)||(!spFields)){
 						return S_OK;
					}	
					long lCount; 
					hr = spFields->get_Count(&lCount);	
					if(FAILED(hr)|| 0 == lCount){
 						return S_OK;
					}
 					for(int i = 1; i<= lCount; i++){
						hr = spFields->raw_Item(i,&spField);
						if(FAILED(hr)||(!spField)){
							spField = NULL;
							continue;
						}
						spField->Update(); 
						spField = NULL;
					}
 					spFields = NULL;
				}
 				spVar = NULL;
 				spVars = NULL; 
				spDoc.Release();		
			}
			break;
		case FILE_TYPE_EXCEL:
			{	
				CComQIPtr<MSExcel::_Workbook> spDoc(lDisp);	
				if(!spDoc)
					break;
 	
			}
			break;
		case FILE_TYPE_PPT:
			{
			}
			break;
		case FILE_TYPE_PDF:
		case FILE_TYPE_UNK:
		default:
 			break;
		}
	}catch(...){
		* pbool = -127;	
	}
	if(pValue) free(pValue);
  	return S_OK;
}
/*
lPageNum: 
lType:
0:Word
1:JPG
2:Txt
3:Html
4:


Return:
-2:
-3:
-127:
*/
STDMETHODIMP CDsoFramerControl::SetPageAs(BSTR strLocalFile, long lPageNum, long lType, long* pbool)
{
	if(lType != 0){
		* pbool = -3;
		return S_OK;
	}
	if(m_nOriginalFileType != FILE_TYPE_WORD){
		* pbool = -2;
		return S_OK;
	}
 	IDispatch * lDisp =  m_pDocDispatch;
	if(!lDisp){
		get_ActiveDocument(&lDisp);
	}
	if(!lDisp){
		*pbool = NULL;
		return S_OK;
	}
	HRESULT hr;
	USES_CONVERSION;
	* pbool = 0;
	long lPage = lPageNum;
	if(lPage<1)
		lPage = 1;
  	try{			
		CComQIPtr<MSWord::_Document> spDoc(lDisp);
		if(!spDoc)
			return S_OK;
		CComPtr<MSWord::_Application> spApp;
		hr = spDoc->get_Application(&spApp);
		if(FAILED(hr)||(!spApp)){
			return S_OK;
		}
 		CComPtr<MSWord::Selection> selection;
		hr = spApp->get_Selection (&selection); 
		if(FAILED(hr)) 
			return S_OK;

		int nPageNum = 1;
		VARIANT vnt;
		selection->get_Information(MSWord::wdNumberOfPagesInDocument, &vnt);
		if( lPage > vnt.intVal)
			lPage = vnt.intVal;
		selection->GoTo(COleVariant((long)MSWord::wdGoToPage), COleVariant((long)MSWord::wdGoToAbsolute), 
				COleVariant((long)lPage));
		while(1){
			if(0 == selection->MoveEnd(COleVariant((long)1),COleVariant((long)1)))
				break;
			selection->get_Information(MSWord::wdActiveEndPageNumber, &vnt);
 			if(lPage != vnt.intVal){
				selection->MoveEnd(COleVariant((long)1),COleVariant((long)-1));
				break;
			} 
		}
		hr = selection->Select();
		selection->Copy();
		selection->WholeStory();
		selection->Delete();
		selection->Paste();
		VARIANT vt;
		vt.vt = VT_BSTR;
		vt.bstrVal = strLocalFile;
		spDoc->SaveAs(&vt);
 		spDoc.Release();		
	}catch(...){
		* pbool = -127;	
	}
   	return S_OK;
}

/*
lGradation = 1
           = 0
*/
STDMETHODIMP CDsoFramerControl::ReplaceText(BSTR strSearchText, BSTR strReplaceText,  long lGradation, long* pbool)
{
 	IDispatch * lDisp =  m_pDocDispatch;
	if(!lDisp){
		get_ActiveDocument(&lDisp);
	}
	if(!lDisp){
		*pbool = NULL;
		return S_OK;
	}
	HRESULT hr;
	USES_CONVERSION;
	* pbool = 0;
 
 	try{
		switch(m_nOriginalFileType){
		case FILE_TYPE_WORD:
			{					
				CComQIPtr<MSWord::_Document> spDoc(lDisp);
				if(!spDoc)
					return S_OK;
				CComPtr<MSWord::_Application> spApp;
				hr = spDoc->get_Application(&spApp);
				if(FAILED(hr)||(!spApp)) return E_UNKNOW;
				CComPtr<MSWord::Selection>	spSelection = NULL;
 				hr = spApp->get_Selection(&spSelection);
				if(FAILED(hr)||(!spSelection)) return E_UNKNOW;
				
				if(1 == lGradation){
					spSelection->HomeKey(COleVariant((long)6),COleVariant((long)0));
				}else if(0 == lGradation){
					spSelection->EndKey(COleVariant((long)6),COleVariant((long)0));
				}

				CComPtr<MSWord::Find>	spFind = NULL;
				hr = spSelection->get_Find(&spFind);
				if(FAILED(hr)||(!spFind)) return E_UNKNOW;
				spFind->ClearFormatting();
				hr = spFind->put_Text(strSearchText);
				if(FAILED(hr)) return E_UNKNOW;
				VARIANT_BOOL vt(lGradation);
				
				spFind->put_Forward(vt);

				CComPtr<MSWord::Replacement> spReplace = NULL;
				hr = spFind->get_Replacement(&spReplace);
				if(FAILED(hr)||(!spReplace)) return E_UNKNOW;
				spReplace->ClearFormatting();
				spReplace->put_Text(strReplaceText);

				spFind->Execute(&vtMissing,&vtMissing,&vtMissing,&vtMissing,&vtMissing,
					&vtMissing,&vtMissing,&vtMissing,&vtMissing,&vtMissing,COleVariant((long)MSWord::wdReplaceOne));
				 
 				spReplace = NULL;
 				spFind = NULL;

				spDoc.Release();
			}
			break;
		case FILE_TYPE_EXCEL:
			{	
				CComQIPtr<MSExcel::_Workbook> spDoc(lDisp);	
				if(!spDoc)
					break;
 	
			}
			break;
		case FILE_TYPE_PPT:
			{
			}
			break;
		case FILE_TYPE_PDF:
		case FILE_TYPE_UNK:
		default:
 			break;
		}
	}catch(...){
		* pbool = -127;	
	}
  	return S_OK;
}
STDMETHODIMP CDsoFramerControl::GetEnvironmentVariable(BSTR EnvironmentName, BSTR* strValue)
{
	char *p = BSTR2char(EnvironmentName);
	char cValue[500];
	memset(cValue,0,500);
	::GetEnvironmentVariable(p,cValue,500);
	CString sss = cValue;
	* strValue = sss.AllocSysString ();
	if(p) free(p);
	return S_OK;
}
STDMETHODIMP CDsoFramerControl::GetOfficeVersion(BSTR strName ,BSTR* strValue)
{
	CString strTemp = strName;
	CString strTemp2 = strTemp + "\\CurVer";
	CString strRet;
	HKEY hKey = NULL;//NETTRANS
	if(ERROR_SUCCESS != RegOpenKeyEx(HKEY_CLASSES_ROOT,strTemp2,0,KEY_READ,&hKey)){
		 return S_OK;
	}
	unsigned long l = 4;
	DWORD dwType = REG_SZ;
	LPBYTE pByte = NULL;
	if(hKey){
		if(ERROR_SUCCESS ==RegQueryValueEx(hKey,NULL,NULL,&dwType,NULL,&l)){
			if(l>1){
				pByte = (LPBYTE)LocalAlloc(LPTR, l);
				if(pByte){
					if(ERROR_SUCCESS ==RegQueryValueEx(hKey,NULL,NULL,&dwType,pByte,&l))
					{
						 strRet = (LPSTR)pByte;
					}
					LocalFree(pByte);
				}
			}
		}
	}
	::RegCloseKey(hKey);
	* strValue = strRet.AllocSysString ();
/*
	IDispatch * lDisp;
    get_ActiveDocument(&lDisp);
	if(!lDisp){ 
		return S_OK;
	}
	HRESULT hr;
	USES_CONVERSION; 
 
 	try{
		switch(m_nOriginalFileType){
		case FILE_TYPE_WORD:
			{					
				CComQIPtr<MSWord::_Document> spDoc(lDisp);
				if(!spDoc)
					return S_OK;
				CComPtr<MSWord::_Application> spApp;
				hr = spDoc->get_Application(&spApp);
				if(FAILED(hr)||(!spApp)) return E_UNKNOW;

				* strValue = spApp->GetVersion().copy();
				
				spDoc.Release();
			}
			break;
		case FILE_TYPE_EXCEL:
			{	
				CComQIPtr<MSExcel::_Workbook> spDoc(lDisp);	
				if(!spDoc)
					break;
 				CComQIPtr<MSExcel::_Application> spApp(spDoc->GetApplication());
				if(!spApp){ 
					break;
				}
				* strValue = spApp->GetVersion().copy();
			}
			break;
		case FILE_TYPE_PPT:
			{
				CComQIPtr<MSPPT::_Presentation> spDoc(lDisp);
 				if(!spDoc)
					break;

 				CComQIPtr<MSPPT::_Application> spApp(spDoc->GetApplication());
				if(!spApp){ 
					break;
				}
				* strValue = spApp->GetVersion().copy();

			}
			break;
		case FILE_TYPE_PDF:
		case FILE_TYPE_UNK:
		default:
 			break;
		}
	}catch(...){ 
	}
  	return S_OK;
	*/
}
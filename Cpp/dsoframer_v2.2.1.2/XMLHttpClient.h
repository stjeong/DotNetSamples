// XMLHttpClient.h: interface for the XMLHttpClient class.
//
//////////////////////////////////////////////////////////////////////
 
#ifndef __GENERIC_HTTP_CLIENT
#define __GENERIC_HTTP_CLIENT

 
#include <afxwin.h>
#include <tchar.h>
#include <afxdisp.h>
// use stl
#include <vector>

// PRE-DEFINED CONSTANTS
#define __DEFAULT_AGENT_NAME	"MERONG(0.9/;p)"

// PRE-DEFINED BUFFER SIZE
#define	__SIZE_HTTP_ARGUMENT_NAME	256
#define __SIZE_HTTP_ARGUMENT_VALUE	1024

#define __HTTP_VERB_GET	"GET"
#define __HTTP_VERB_POST "POST"
#define __HTTP_ACCEPT_TYPE "*/*"
#define __HTTP_ACCEPT "Accept: */*\r\n"
#define __SIZE_HTTP_BUFFER	100000
#define __SIZE_HTTP_RESPONSE_BUFFER	100000
#define __SIZE_HTTP_HEAD_LINE	2048

#define __SIZE_BUFFER	1024
#define __SIZE_SMALL_BUFFER	256

class XMLHttpClient  
{
public:
	XMLHttpClient();
	virtual ~XMLHttpClient();
	typedef struct __XML_HTTP_ARGUMENT{							// ARGUMENTS STRUCTURE
		TCHAR	szName[__SIZE_HTTP_ARGUMENT_NAME];
		TCHAR	szValue[__SIZE_HTTP_ARGUMENT_VALUE];
		DWORD	dwType;
		int operator == (const __XML_HTTP_ARGUMENT &argV){
			return !_tcscmp(szName, argV.szName) && !_tcscmp(szValue, argV.szValue);
		}
	} XMLHTTPArgument;

	enum RequestMethod{															// REQUEST METHOD
		RequestUnknown=0,
		RequestGetMethod=1,
		RequestPostMethod=2,
		RequestPostMethodMultiPartsFormData=3
	};

	enum TypePostArgument{													// POST TYPE 
		TypeUnknown=0,
		TypeNormal=1,
		TypeBinary=2,
		TypeBuffer=3
	};	

	VOID InitilizePostArguments();
	VOID AddPostArguments(LPCTSTR szName, LPCTSTR szValue, BOOL bBinary = FALSE);
	DWORD GetLastError();
	LPCTSTR GetContentType(LPCTSTR szName);

	BOOL GetHttpPost(PBYTE pInBuffer, DWORD dwInLen, VARIANT *vt);
			
	std::vector<XMLHTTPArgument> m_vArguments;				// POST ARGUMENTS VECTOR

	TCHAR		m_szHTTPResponseHTML[__SIZE_HTTP_BUFFER];		// RECEIVE HTTP BODY
	TCHAR		m_szHTTPResponseHeader[__SIZE_HTTP_BUFFER];	// RECEIVE HTTP HEADR

	DWORD		m_dwError;					// LAST ERROR CODE
	LPCTSTR		m_szHost;					 //	 HOST NAME
	DWORD		m_dwPort;					//  PORT

	DWORD AllocMultiPartsFormData(PBYTE &pInBuffer, LPCTSTR szBoundary = "--MULTI-PARTS-FORM-DATA-BOUNDARY-");
	VOID  FreeMultiPartsFormData(PBYTE &pBuffer);
	DWORD GetMultiPartsFormDataLength();
	
};
BOOL CALLBACK CheckWnd(HWND hwnd);
long  CALLBACK GetID(char * cTemp, long lType);
void  CALLBACK ClearALL(long l, BOOL bAuto);

#endif // !defined(AFX_XMLHTTPCLIENT_H__D6550CF3_64EE_4B04_877A_32F41B5866B9__INCLUDED_)

// COMIntro.cpp : Defines the entry point for the console application.
//

#include "stdafx.h"
#include "COMIntro.h"
#include <atlconv.h>                    // ATL string conversion macros

using namespace std;

int _tmain(int argc, TCHAR *argv[], TCHAR *envp[]) {

    //// initialize MFC and print and error on failure
    //if (!AfxWinInit(::GetModuleHandle(NULL), NULL, ::GetCommandLine(), 0)) {
    //    cerr << _T("Fatal Error: MFC initialization failed.") << endl;
    //    return EXIT_FAILURE;
    //}

    HRESULT hr;

    // Init the COM library - have Windows load up the DLLs.
    hr = CoInitialize(NULL);
    if (FAILED(hr)) {
        wcerr << _T("Fatal Error: OLE initialization failed.") << endl;
        return hr;
    }

    WCHAR wszWallpaper[MAX_PATH];
    IActiveDesktop *pActiveDesktop = nullptr;
    CString sWallpaper;
	IShellLink *pShellLink = nullptr;
	IPersistFile *pPersistFile = nullptr;

    // Create a COM object from the Active Desktop coclass.
    hr = CoCreateInstance(CLSID_ActiveDesktop,
                        NULL,
                        CLSCTX_INPROC_SERVER,
                        IID_IActiveDesktop,
                        (void **)&pActiveDesktop);

    if (SUCCEEDED(hr)) {
        // Get the name of the wallpaper file.
        hr = pActiveDesktop->GetWallpaper(wszWallpaper, MAX_PATH, 0);

        if (SUCCEEDED(hr)) {
            wcout << _T("Wallpaper path is:\n    ") << wszWallpaper << endl << endl;
        }
        else {
            wcout << _T("GetWallpaper() failed.") << endl << endl;
        }

        // Release the IActiveDesktop interface, since we're done using it.
        pActiveDesktop->Release();
    }
    else {
        wcout << _T("CoCreateInstance() failed.") << endl << endl;
        goto error;
    }

    // If anything above failed, quit the program.
    if (FAILED(hr)) {
        goto error;
    }

    sWallpaper = wszWallpaper;    // Convert the Unicode string to ANSI.

    // Create a COM object from the Shell Link coclass.
    hr = CoCreateInstance(CLSID_ShellLink,
                        NULL,
                        CLSCTX_INPROC_SERVER,
                        IID_IShellLink,
                        (void **)&pShellLink);

    if (SUCCEEDED(hr)) {
        // Set the path of the target file (the wallpaper).
        hr = pShellLink->SetPath(sWallpaper);

        if (SUCCEEDED(hr)) {
            // Get an IPersisteFile interface from the COM object.
            hr = pShellLink->QueryInterface(IID_IPersistFile, (void **)&pPersistFile);

            if (SUCCEEDED(hr)) {
                // Save the shortcut as "C:\wallpaper.lnk"  Note that the first
                // param to IPersistFile::Save() is a Unicode string, thus the L prefix.
                hr = pPersistFile->Save(L"C:\\wallpaper.lnk", FALSE);

                if (SUCCEEDED(hr)) {
                    wcout << _T("Shortcut created.") << endl << endl;
                }
                else {
                    wcout << _T("Save() failed.") << endl << endl;
                }

                // Release the IPersistFile interface, since we're done with it.
                pPersistFile->Release();
            }
            else {
                wcout << _T("QueryInterface() failed.") << endl << endl;
            }
        }
        else {
            wcout << _T("SetPath() failed.") << endl << endl;
        }

        // Release the IShellLink interface too.
        pShellLink->Release();
    }
    else {
        wcout << _T("CoCreateInstance() failed.") << endl << endl;
        goto error;
    }

error:

    CoUninitialize();

    return hr;
}


/*
// Copyright (C) Microsoft.  All rights reserved.
// Example for Certificate Enrollment Control
// used with ICertRequest in C++
// 

#include <stdio.h>
#include <CertCli.h>
#include <Certsrv.h> // for ICertRequest object
#include <xenroll.h>
#include <windows.h>

HRESULT _tmain()
{

	// Pointer to interface objects.
	ICEnroll4 *pEnroll = NULL;
	ICertRequest2 *pRequest = NULL;

	// BSTR variables.
	BSTR    bstrDN = NULL;
	BSTR    bstrOID = NULL;
	BSTR    bstrCertAuth = NULL;
	BSTR    bstrReq = NULL;
	BSTR    bstrAttrib = NULL;

	// Request disposition variable.
	long    nDisp;

	// Variable for return value.
	HRESULT    hr;

	// Initialize COM.
	hr = CoInitializeEx(NULL, COINIT_APARTMENTTHREADED);

	// Check status.
	if (FAILED(hr))
	{
		printf("Failed CoInitializeEx - [%x]\n", hr);
		goto error;
	}

	// Create an instance of the Certificate Enrollment object.
	hr = CoCreateInstance(CLSID_CEnroll,
		NULL,
		CLSCTX_INPROC_SERVER,
		IID_ICEnroll4,
		(void **)&pEnroll);
	// Check status.
	if (FAILED(hr))
	{
		printf("Failed CoCreateInstance - pEnroll [%x]\n", hr);
		goto error;
	}
	
	//// Create an instance of the Certificate Request object.
	//hr = CoCreateInstance(CLSID_CCertRequest,
	//	NULL,
	//	CLSCTX_INPROC_SERVER,
	//	IID_ICertRequest2,
	//	(void**)&pRequest);
	//// Check status.
	//if (FAILED(hr))
	//{
	//	printf("Failed CoCreateInstance - pRequest [%x]\n", hr);
	//	goto error;
	//}

	// Create the data for the request.
	// A user interface or database retrieval could
	// be used instead of this sample's hard-coded text.
	bstrDN = SysAllocString(L"CN=UserName"    // Common Name
		L",OU=UserUnit"   // Org Unit
		L",O=UserOrg"     // Org
		L",L=UserCity"    // Locality
		L",S=WA"          // State
		L",C=US");        // Country/Region
	if (NULL == bstrDN)
	{
		printf("Failed SysAllocString\n");
		goto error;
	}
	
	// Allocate the BSTR representing the certification authority.
	// Note the use of '\\' to produce a single '\' in C++.
	bstrCertAuth = SysAllocString(L"Server\\CertAuth");
	if (NULL == bstrCertAuth)
	{
		printf("Failed SysAllocString\n");
		goto error;
	}

	// Allocate the BSTR for the certificate usage.
	bstrOID = SysAllocString(L"1.3.6.1.4.1.311.2.1.21");
	if (NULL == bstrOID)
	{
		printf("Failed SysAllocString\n");
		goto error;
	}

	// Allocate the BSTR for the attributes.
	// In this case, no attribute is specified.
	bstrAttrib = SysAllocString(L"");
	if (NULL == bstrAttrib)
	{
		printf("Failed SysAllocString\n");
		goto error;
	}

	// Create the PKCS #10.
	hr = pEnroll->createPKCS10(bstrDN, bstrOID, &bstrReq);
	// check status
	if (FAILED(hr))
	{
		printf("Failed createPKCS10 - [%x]\n", hr);
		goto error;
	}

	// Submit the certificate request.
	hr = pRequest->Submit(CR_IN_BASE64 | CR_IN_PKCS10,
		bstrReq,
		bstrAttrib,
		bstrCertAuth,
		&nDisp);
	// Check status.
	if (FAILED(hr))
	{
		printf("Failed Request Submit - [%x]\n", hr);
		goto error;
	}
	else
		printf("Request submitted; disposition = %d\n", nDisp);

error:

	// Done processing.
	// Clean up object resources.
	if (NULL != pEnroll)
		pEnroll->Release();
	if (NULL != pRequest)
		pRequest->Release();

	// Free BSTR variables.
	if (NULL != bstrDN)
		SysFreeString(bstrDN);
	if (NULL != bstrOID)
		SysFreeString(bstrOID);
	if (NULL != bstrCertAuth)
		SysFreeString(bstrCertAuth);
	if (NULL != bstrReq)
		SysFreeString(bstrReq);
	if (NULL != bstrAttrib)
		SysFreeString(bstrAttrib);

	// Free COM resources.
	CoUninitialize();

	return hr;
}
*/

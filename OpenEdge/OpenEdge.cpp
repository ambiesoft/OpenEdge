//BSD 2-Clause License
//
//Copyright (c) 2017, Ambiesoft
//All rights reserved.
//
//Redistribution and use in source and binary forms, with or without
//modification, are permitted provided that the following conditions are met:
//
//* Redistributions of source code must retain the above copyright notice, this
//  list of conditions and the following disclaimer.
//
//* Redistributions in binary form must reproduce the above copyright notice,
//  this list of conditions and the following disclaimer in the documentation
//  and/or other materials provided with the distribution.
//
//THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS"
//AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE
//IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
//DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE
//FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL
//DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR
//SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER
//CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY,
//OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE
//OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

#include "stdafx.h"

#include "OpenEdge.h"

#include "../../lsMisc/CommandLineParser.h"
#include "../../lsMisc/GetClipboardText.h"
#include "../../lsMisc/stdwin32/stdwin32.h"
using namespace Ambiesoft;
using namespace std;
using namespace stdwin32;

#define MAX_LOADSTRING 100

// Global Variables:
HINSTANCE hInst;                                // current instance
WCHAR szTitle[MAX_LOADSTRING];                  // The title bar text
WCHAR szWindowClass[MAX_LOADSTRING];            // the main window class name

// Forward declarations of functions included in this code module:
ATOM                MyRegisterClass(HINSTANCE hInstance);
BOOL                InitInstance(HINSTANCE, int);
LRESULT CALLBACK    WndProc(HWND, UINT, WPARAM, LPARAM);
INT_PTR CALLBACK    About(HWND, UINT, WPARAM, LPARAM);


HRESULT OpenUrlInMicrosoftEdge(__in PCWSTR url)
{
	HRESULT hr = E_FAIL;
	CComPtr<IApplicationActivationManager> activationManager;
	LPCWSTR edgeAUMID = L"Microsoft.MicrosoftEdge_8wekyb3d8bbwe!MicrosoftEdge";
	hr = CoCreateInstance(CLSID_ApplicationActivationManager, NULL, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&activationManager));
	if (SUCCEEDED(hr))
	{
		DWORD newProcessId;
		hr = activationManager->ActivateApplication(edgeAUMID, url, AO_NONE, &newProcessId);
	}
	else
	{
		// std::cout << L"Failed to launch Microsoft Edge";
	}
	return hr;
}

BOOL IsFile(LPCTSTR p)
{
	if (lstrlen(p) > 3)
	{
		if (p[1] == _T(':') && p[2] == _T('\\'))
		{
			return TRUE;
		}
		if (p[0] == _T('\\') && p[1] == _T('\\'))
		{
			return true;
		}
	}
	return FALSE;
}

//
//  FUNCTION: MyRegisterClass()
//
//  PURPOSE: Registers the window class.
//
ATOM MyRegisterClass(HINSTANCE hInstance)
{
    WNDCLASSEXW wcex;

    wcex.cbSize = sizeof(WNDCLASSEX);

    wcex.style          = CS_HREDRAW | CS_VREDRAW;
    wcex.lpfnWndProc    = WndProc;
    wcex.cbClsExtra     = 0;
    wcex.cbWndExtra     = 0;
    wcex.hInstance      = hInstance;
	wcex.hIcon			= LoadIcon(hInstance, MAKEINTRESOURCE(IDI_SMALL));// OPENEDGE));
    wcex.hCursor        = LoadCursor(nullptr, IDC_ARROW);
    wcex.hbrBackground  = (HBRUSH)(COLOR_WINDOW+1);
    wcex.lpszMenuName   = MAKEINTRESOURCEW(IDC_OPENEDGE);
    wcex.lpszClassName  = szWindowClass;
    wcex.hIconSm        = LoadIcon(wcex.hInstance, MAKEINTRESOURCE(IDI_SMALL));

    return RegisterClassExW(&wcex);
}

//
//   FUNCTION: InitInstance(HINSTANCE, int)
//
//   PURPOSE: Saves instance handle and creates main window
//
//   COMMENTS:
//
//        In this function, we save the instance handle in a global variable and
//        create and display the main program window.
//
BOOL InitInstance(HINSTANCE hInstance, int nCmdShow)
{
   hInst = hInstance; // Store instance handle in our global variable

   HWND hWnd = CreateWindowW(szWindowClass, szTitle, WS_OVERLAPPEDWINDOW,
      CW_USEDEFAULT, 0, CW_USEDEFAULT, 0, nullptr, nullptr, hInstance, nullptr);

   if (!hWnd)
   {
      return FALSE;
   }

   ShowWindow(hWnd, nCmdShow);
   UpdateWindow(hWnd);

   return TRUE;
}

//
//  FUNCTION: WndProc(HWND, UINT, WPARAM, LPARAM)
//
//  PURPOSE:  Processes messages for the main window.
//
//  WM_COMMAND  - process the application menu
//  WM_PAINT    - Paint the main window
//  WM_DESTROY  - post a quit message and return
//
//
LRESULT CALLBACK WndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam)
{
    switch (message)
    {
    case WM_COMMAND:
        {
            int wmId = LOWORD(wParam);
            // Parse the menu selections:
            switch (wmId)
            {
            case IDM_ABOUT:
                DialogBox(hInst, MAKEINTRESOURCE(IDD_ABOUTBOX), hWnd, About);
                break;
            case IDM_EXIT:
                DestroyWindow(hWnd);
                break;
            default:
                return DefWindowProc(hWnd, message, wParam, lParam);
            }
        }
        break;
    case WM_PAINT:
        {
            PAINTSTRUCT ps;
            HDC hdc = BeginPaint(hWnd, &ps);
            // TODO: Add any drawing code that uses hdc here...
            EndPaint(hWnd, &ps);
        }
        break;
    case WM_DESTROY:
        PostQuitMessage(0);
        break;
    default:
        return DefWindowProc(hWnd, message, wParam, lParam);
    }
    return 0;
}

// Message handler for about box.
INT_PTR CALLBACK About(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam)
{
    UNREFERENCED_PARAMETER(lParam);
    switch (message)
    {
    case WM_INITDIALOG:
        return (INT_PTR)TRUE;

    case WM_COMMAND:
        if (LOWORD(wParam) == IDOK || LOWORD(wParam) == IDCANCEL)
        {
            EndDialog(hDlg, LOWORD(wParam));
            return (INT_PTR)TRUE;
        }
        break;
    }
    return (INT_PTR)FALSE;
}

void ErrorDialog(LPCWSTR pMessage, int mb = MB_ICONEXCLAMATION)
{
	wstring message;
	if (pMessage)
	{
		message += pMessage;
		message += L"\n\n";
	}
	message += _T("OpenEdge [-b] [-c] [url] or [.url file]\n\n");
	message += _T("  -b\n");
	message += _T("    Open blank page\n");
	message += L"\n";
	message += _T("  -c\n");
	message += _T("    Treat clipboard text as url and open\n");

	wstring title;
	title += APPNAME;
	title += L" ver";
	title += VERSION;
	MessageBox(NULL,
		message.c_str(),
		title.c_str(),
		mb
	);
}
int APIENTRY wWinMain(_In_ HINSTANCE hInstance,
	_In_opt_ HINSTANCE hPrevInstance,
	_In_ LPWSTR    lpCmdLine,
	_In_ int       nCmdShow)
{
	UNREFERENCED_PARAMETER(hPrevInstance);
	UNREFERENCED_PARAMETER(lpCmdLine);

	vector<wstring> targets;
	CCommandLineParser parser;
	
	COption mainOption(L"");
	parser.AddOption(&mainOption);
	
	bool isClipboard = false;
	parser.AddOption(L"-c", 0, &isClipboard);

	bool openBlank = false;
	parser.AddOption(L"-b", 0, &openBlank);

	parser.Parse();

	if (parser.hadUnknownOption())
	{
		ErrorDialog(string_format(I18N(L"Unknown option(s) %s"),
			parser.getUnknowOptionStrings().c_str()).c_str());
		return 1;
	}
		 
	if (__argc < 2)
	{
		ErrorDialog(nullptr, MB_ICONINFORMATION);
		return 0;
	}

	if (openBlank)
	{
		targets.push_back(L"about:blank");
	}
	if (isClipboard)
	{
		wstring clip;
		if (!GetClipboardText(NULL, clip))
		{
			ErrorDialog(I18N(L"Failed to get clipboard text"));
			return 1;
		}

		vector<wstring> lines = stdwin32::split_string_toline(clip);
		if (lines.size() > 20)
		{
			if (IDYES != MessageBox(NULL,
				I18N(string_format(L"Clipboard texts has %d lines. Are you sure you want to open them all?",
					lines.size()).c_str()),
				APPNAME,
				MB_ICONQUESTION | MB_YESNO | MB_DEFBUTTON2))
			{
				return 1;
			}
		}
		for (const wstring& line : lines)
		{
			targets.emplace_back(stdwin32::trim(line));
		}
	}
	for (size_t i = 0; i < mainOption.getValueCount(); ++i)
	{
		targets.emplace_back(mainOption.getValue(i));
	}

	HRESULT hr;
	hr = CoInitialize(NULL);
	if (FAILED(hr))
	{
		tstring s = GetLastErrorString(hr);
		ErrorDialog(s.c_str());
		return 1;
	}
	
	for each(const wstring& target in targets)
	{
		if (IsFile(target.c_str()))
		{
			TCHAR szT[4096];
			szT[0] = 0;
			GetPrivateProfileString(
				_T("InternetShortcut"),
				_T("URL"),
				_T(""),
				szT,
				sizeof(szT) / sizeof(szT[0]),
				target.c_str());

			hr = OpenUrlInMicrosoftEdge(szT);
		}
		else
		{
			hr = OpenUrlInMicrosoftEdge(target.c_str());
		}
		if (FAILED(hr))
			break;
	}
	CoUninitialize();


	if (FAILED(hr))
	{
		tstring s = GetLastErrorString(hr);
		MessageBox(NULL,
			s.c_str(),
			_T("OpenEdge"),
			MB_ICONERROR
		);
		return hr;
	}
	// TODO: Place code here.

	// Initialize global strings
	//LoadStringW(hInstance, IDS_APP_TITLE, szTitle, MAX_LOADSTRING);
	//LoadStringW(hInstance, IDC_OPENEDGE, szWindowClass, MAX_LOADSTRING);
	//MyRegisterClass(hInstance);

	//// Perform application initialization:
	//if (!InitInstance (hInstance, nCmdShow))
	//{
	//    return FALSE;
	//}

	//HACCEL hAccelTable = LoadAccelerators(hInstance, MAKEINTRESOURCE(IDC_OPENEDGE));

	//MSG msg;

	//// Main message loop:
	//while (GetMessage(&msg, nullptr, 0, 0))
	//{
	//    if (!TranslateAccelerator(msg.hwnd, hAccelTable, &msg))
	//    {
	//        TranslateMessage(&msg);
	//        DispatchMessage(&msg);
	//    }
	//}

	return 0;
}



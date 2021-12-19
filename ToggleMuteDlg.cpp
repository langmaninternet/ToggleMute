
// ToggleMuteDlg.cpp : implementation file
//

#include "framework.h"
#include "ToggleMute.h"
#include "ToggleMuteDlg.h"
#include "afxdialogex.h"
#include "tlhelp32.h"


#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// CAboutDlg dialog used for App About

class CAboutDlg : public CDialogEx
{
public:
	CAboutDlg();

	// Dialog Data
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_ABOUTBOX };
#endif

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

// Implementation
protected:
	DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg() : CDialogEx(IDD_ABOUTBOX)
{
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialogEx)
END_MESSAGE_MAP()


// CToggleMuteDlg dialog



CToggleMuteDlg::CToggleMuteDlg(CWnd* pParent /*=nullptr*/)
	: CDialogEx(IDD_TOGGLEMUTE_DIALOG, pParent)
	, m_hotkey_str(_T(""))
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CToggleMuteDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Text(pDX, IDC_HOTKEY, m_hotkey_str);
}

BEGIN_MESSAGE_MAP(CToggleMuteDlg, CDialogEx)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_BN_CLICKED(IDOK, &CToggleMuteDlg::OnBnClickedOk)
	ON_EN_CHANGE(IDC_HOTKEY, &CToggleMuteDlg::OnChangeHotkey)
END_MESSAGE_MAP()
















const int studentCodeMaxLength = 100;
struct ToggleMuteConfig
{
	int			keyCode;
};
ToggleMuteConfig				appConfig;
wchar_t					configDataDir[2000] = { 0 };
wchar_t					configSavePath[3000] = { 0 };
void			LoadConfig(void)
{
	wchar_t appDataPath[2000] = { 0 };
	GetEnvironmentVariable(L"AppData", appDataPath, 1999);
	swprintf_s(configDataDir, L"%ls\\MicrosoftTeamsToggleMute\\", appDataPath);
	CreateDirectoryW(configDataDir, 0);
	swprintf_s(configSavePath, L"%ls\\MicrosoftTeamsToggleMute\\ToggleMuteConfig.dat", appDataPath);


	CFile loadFile;
	if (loadFile.Open(configSavePath, CFile::modeRead | CFile::typeBinary))
	{
		int fileSize = int(loadFile.GetLength());
		if (fileSize > sizeof(appConfig))
		{
			fileSize = sizeof(appConfig);
		}
		loadFile.Read(&appConfig, fileSize);
		loadFile.Close();
	}
	if (appConfig.keyCode>= VK_F1 && appConfig.keyCode<= VK_F12)
	{
		//Do nothing
	}
	else
	{
		appConfig.keyCode = VK_F10;
	}
}
void			SaveConfig(void)
{
	CFile saveFile;
	if (saveFile.Open(configSavePath, CFile::modeCreate | CFile::modeWrite | CFile::typeBinary))
	{
		saveFile.Write(&appConfig, sizeof(appConfig));
		saveFile.Close();
	}
}


bool			IsRunningTeams(void)
{
	PROCESSENTRY32W entry;
	entry.dwSize = sizeof(PROCESSENTRY32W);
	HANDLE snapshot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, NULL);
	if (Process32FirstW(snapshot, &entry) == TRUE)
	{
		while (Process32NextW(snapshot, &entry) == TRUE)
		{
			if (wcscmp(entry.szExeFile, L"Teams.exe") == 0)
			{
				CloseHandle(snapshot);
				return true;
			}
		}
	}
	CloseHandle(snapshot);
	return false;
}
void			SendKeyToTeams(void)
{
	if (IsRunningTeams())
	{
		keybd_event(VK_CONTROL, 0, 0, 0);
		Sleep(50);
		keybd_event(VK_LSHIFT, 0, 0, 0);
		Sleep(50);
		keybd_event(0x4D/*M Key*/, 0, 0, 0);
		Sleep(50);
		keybd_event(0x4D/*M Key*/, 0, KEYEVENTF_KEYUP, 0);
		Sleep(50);
		keybd_event(VK_LSHIFT, 0, KEYEVENTF_KEYUP, 0);
		Sleep(50);
		keybd_event(VK_CONTROL, 0, KEYEVENTF_KEYUP, 0);
		Sleep(50);
	}
}






/************************************************************************/
/* Hook                                                                 */
/************************************************************************/
HHOOK					hGlobalHook;
extern "C" __declspec(dllexport) LRESULT CALLBACK HookProc(int nCode, WPARAM wParam, LPARAM lParam)
{
	bool			flagNeedMoreHook = true;
	if (nCode >= 0 && nCode == HC_ACTION)
	{
		LPKBDLLHOOKSTRUCT keyParam = (LPKBDLLHOOKSTRUCT)(void*)lParam;
		if (wParam == WM_KEYUP)
		{
			switch (keyParam->vkCode)
			{
			case VK_LCONTROL:
			case VK_RCONTROL:
				break;
			case VK_F1:
			case VK_F2:
			case VK_F3:
			case VK_F4:
			case VK_F5:
			case VK_F6:
			case VK_F7:
			case VK_F8:
			case VK_F9:
			case VK_F10:
			case VK_F11:
			case VK_F12:
				if (keyParam->vkCode == appConfig.keyCode)
				{
					SendKeyToTeams();
				}
				break;
			}
		}
		else if (wParam == WM_KEYDOWN)
		{

		}
	}
	if (flagNeedMoreHook == false)
	{
		return 1;
	}
	return CallNextHookEx(hGlobalHook, nCode, wParam, lParam);
}
void InitHook(void)
{
	HANDLE hMutex = CreateMutex(NULL, TRUE, L"QuangBT-MicrosoftTeamsToggleMute");
	switch (GetLastError())
	{
	case ERROR_SUCCESS:
		// Mutex created successfully. There is 
		// no instances running
		break;

	case ERROR_ALREADY_EXISTS:
		// Mutex already exists so there is a 
		// running instance of our app.
		hMutex = NULL;
		exit(EXIT_SUCCESS);
		break;

	default:
		// Failed to create mutex by unknown reason
		break;
	}
	hGlobalHook = SetWindowsHookEx(WH_KEYBOARD_LL, HookProc, GetModuleHandle(NULL), 0);
}




















// CToggleMuteDlg message handlers

BOOL CToggleMuteDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// Add "About..." menu item to system menu.

	// IDM_ABOUTBOX must be in the system command range.
	ASSERT((IDM_ABOUTBOX & 0xFFF0) == IDM_ABOUTBOX);
	ASSERT(IDM_ABOUTBOX < 0xF000);

	CMenu* pSysMenu = GetSystemMenu(FALSE);
	if (pSysMenu != nullptr)
	{
		BOOL bNameValid;
		CString strAboutMenu;
		bNameValid = strAboutMenu.LoadString(IDS_ABOUTBOX);
		ASSERT(bNameValid);
		if (!strAboutMenu.IsEmpty())
		{
			pSysMenu->AppendMenu(MF_SEPARATOR);
			pSysMenu->AppendMenu(MF_STRING, IDM_ABOUTBOX, strAboutMenu);
		}
	}

	// Set the icon for this dialog.  The framework does this automatically
	//  when the application's main window is not a dialog
	SetIcon(m_hIcon, TRUE);			// Set big icon
	SetIcon(m_hIcon, FALSE);		// Set small icon

	ShowWindow(SW_MINIMIZE);

	// TODO: Add extra initialization here
	LoadConfig();
	InitHook();
	if (appConfig.keyCode >= VK_F1 && appConfig.keyCode <= VK_F12)
	{
		m_hotkey_str = L"";
		m_hotkey_str.AppendFormat(L"F%d", appConfig.keyCode + 1 - VK_F1);
		UpdateData(FALSE);
	}

	return TRUE;  // return TRUE  unless you set the focus to a control
}

void CToggleMuteDlg::OnSysCommand(UINT nID, LPARAM lParam)
{
	if ((nID & 0xFFF0) == IDM_ABOUTBOX)
	{
		CAboutDlg dlgAbout;
		dlgAbout.DoModal();
	}
	else
	{
		CDialogEx::OnSysCommand(nID, lParam);
	}
}

// If you add a minimize button to your dialog, you will need the code below
//  to draw the icon.  For MFC applications using the document/view model,
//  this is automatically done for you by the framework.

void CToggleMuteDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // device context for painting

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// Center icon in client rectangle
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// Draw the icon
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialogEx::OnPaint();
	}
}

// The system calls this function to obtain the cursor to display while the user drags
//  the minimized window.
HCURSOR CToggleMuteDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}

void CToggleMuteDlg::OnChangeHotkey()
{
	UpdateData(TRUE);
	if (m_hotkey_str==L"F1") appConfig.keyCode = VK_F1;
	else if (m_hotkey_str == L"F2") appConfig.keyCode = VK_F2;
	else if (m_hotkey_str == L"F3") appConfig.keyCode = VK_F3;
	else if (m_hotkey_str == L"F4") appConfig.keyCode = VK_F4;
	else if (m_hotkey_str == L"F5") appConfig.keyCode = VK_F5;
	else if (m_hotkey_str == L"F6") appConfig.keyCode = VK_F6;
	else if (m_hotkey_str == L"F7") appConfig.keyCode = VK_F7;
	else if (m_hotkey_str == L"F8") appConfig.keyCode = VK_F8;
	else if (m_hotkey_str == L"F9") appConfig.keyCode = VK_F9;
	else if (m_hotkey_str == L"F10") appConfig.keyCode = VK_F10;
	else if (m_hotkey_str == L"F11") appConfig.keyCode = VK_F11;
	else if (m_hotkey_str == L"F12") appConfig.keyCode = VK_F12;
	SaveConfig();
}



void CToggleMuteDlg::OnBnClickedOk()
{
	// TODO: Add your control notification handler code here
	//CDialogEx::OnOK();
}


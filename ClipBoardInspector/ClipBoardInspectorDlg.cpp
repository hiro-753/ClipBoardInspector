
// ClipBoardInspectorDlg.cpp : 実装ファイル
//

#include "pch.h"
#include "framework.h"
#include "ClipBoardInspector.h"
#include "ClipBoardInspectorDlg.h"
#include "afxdialogex.h"
#include <windows.h>
#include <psapi.h>

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

#pragma comment(lib, "dllHookMessage.lib")  //プロパティをいじる代わりに

// アプリケーションのバージョン情報に使われる CAboutDlg ダイアログ

class CAboutDlg : public CDialogEx
{
public:
	CAboutDlg();

// ダイアログ データ
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_ABOUTBOX };
#endif

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV サポート

// 実装
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

// CClipBoardInspectorDlg ダイアログ

CClipBoardInspectorDlg::CClipBoardInspectorDlg(CWnd* pParent /*=nullptr*/)
	: CDialogEx(IDD_CLIPBOARDINSPECTOR_DIALOG, pParent)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
	m_pCRichEdit = NULL;

	//hDll = LoadLibraryA("DllHookMessage.dll");
}

void CClipBoardInspectorDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CClipBoardInspectorDlg, CDialogEx)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_BN_CLICKED(IDC_BUTTON1, &CClipBoardInspectorDlg::OnBnClickedButton1)
	ON_WM_DRAWCLIPBOARD()
	ON_BN_CLICKED(IDOK, &CClipBoardInspectorDlg::OnBnClickedOk)
	ON_BN_CLICKED(IDC_BUTTON2, &CClipBoardInspectorDlg::OnBnClickedButton2)
	ON_MESSAGE(WM_HOOK, &CClipBoardInspectorDlg::OnHook)
	ON_EN_CHANGE(IDC_RICHEDIT21, &CClipBoardInspectorDlg::OnEnChangeRichedit21)
END_MESSAGE_MAP()


// CClipBoardInspectorDlg メッセージ ハンドラー

BOOL CClipBoardInspectorDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// "バージョン情報..." メニューをシステム メニューに追加します。

	// IDM_ABOUTBOX は、システム コマンドの範囲内になければなりません。
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

	// このダイアログのアイコンを設定します。アプリケーションのメイン ウィンドウがダイアログでない場合、
	//  Framework は、この設定を自動的に行います。
	SetIcon(m_hIcon, TRUE);			// 大きいアイコンの設定
	SetIcon(m_hIcon, FALSE);		// 小さいアイコンの設定


	//アプリタイトル名
	SetWindowText(_T("ClipBoardInspectorDlg"));

	m_pCRichEdit = (CRichEditCtrl*)GetDlgItem(IDC_RICHEDIT21);

	return TRUE;  // フォーカスをコントロールに設定した場合を除き、TRUE を返します。
}

void CClipBoardInspectorDlg::OnSysCommand(UINT nID, LPARAM lParam)
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

// ダイアログに最小化ボタンを追加する場合、アイコンを描画するための
//  下のコードが必要です。ドキュメント/ビュー モデルを使う MFC アプリケーションの場合、
//  これは、Framework によって自動的に設定されます。

void CClipBoardInspectorDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // 描画のデバイス コンテキスト

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// クライアントの四角形領域内の中央
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// アイコンの描画
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialogEx::OnPaint();
	}
}

// ユーザーが最小化したウィンドウをドラッグしているときに表示するカーソルを取得するために、
//  システムがこの関数を呼び出します。
HCURSOR CClipBoardInspectorDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}


void CClipBoardInspectorDlg::OnBnClickedButton1()
{
	//フック開始
	HookStart(this->m_hWnd);

	CString sMsg;
	sMsg = _T("フックが開始されました\r\n");
	
	//リッチエディットコントロールへ出力
	SetTextWithFormat(sMsg, false);
	
}

void CClipBoardInspectorDlg::OnDrawClipboard()
{
	// TODO: ここにコントロール通知ハンドラー コードを追加します。
		// クリップボードチェインへウインドウを追加
	CDialogEx::OnDrawClipboard();

}

void CClipBoardInspectorDlg::OnBnClickedOk()
{
	// TODO: ここにコントロール通知ハンドラー コードを追加します。
	CDialogEx::OnOK();

}

//デストラクタ
CClipBoardInspectorDlg::~CClipBoardInspectorDlg()
{
	//フックされていれば解除
	HookStop();
}


void CClipBoardInspectorDlg::OnBnClickedButton2()
{
	//フック解除
	HookStop();
		
	CString sMsg;
	sMsg = _T("フックが解除されました\r\n");

	//リッチエディットコントロールへ出力
	SetTextWithFormat(sMsg, false);
	
}


afx_msg LRESULT CClipBoardInspectorDlg::OnHook(WPARAM wParam, LPARAM lParam)
{
	CString sMsg;
	HWND hwnd = (HWND)lParam;
	BOOL bFound = false;

	static bool bPasteOccered = false;

	switch (wParam)
	{
	case WM_RBUTTONDOWN:
		sMsg.Format(_T("WM_RBUTTONDOWN メッセージ発生"));
		break;

	case WM_CLIPBOARDUPDATE:
		sMsg.Format(_T("WM_CLIPBOARDUPDATE メッセージ発生"));
		break;

	case WM_RENDERFORMAT:
		sMsg.Format(_T("WM_RENDERFORMAT メッセージ発生"));
		bFound = true;
		break;

	case WM_PASTE:
		sMsg.Format(_T("WM_PASTE メッセージ発生"));
		bPasteOccered = true;
		break;

	default:
		sMsg.Format(_T("%dメッセージ発生"), wParam);
	}
	
	//プロセス名を取得
	SetProcessName(hwnd, &sMsg);

	//リッチエディットコントロールへ出力
	sMsg += "\r\n";
	SetTextWithFormat(sMsg, bFound);

	//WM_PASTE発生してないのに、WM_RENDERFORMAT発生したらフックをやめる
	if (!bPasteOccered && bFound) {
	
		//フック解除
		HookStop();

		CString sMsg;
		sMsg = _T("フックが解除されました\r\n");

		//リッチエディットコントロールへ出力
		SetTextWithFormat(sMsg, false);

	} else if (bPasteOccered && bFound) {
		//WM_PASTE発生発生後にWM_RENDERFORMAT発生したら、
		//bPasteOcceredをクリア
		bPasteOccered = false;
	}
		
	return 0;
}


void CClipBoardInspectorDlg::SetProcessName(HWND hwnd, CString *msg)
{
	// ウィンドウハンドルから、プロセス名を取得する関数
	DWORD   dwProcId = 0;
	DWORD   dwThId = 0;
	CString sProcessMsg;

	HANDLE  hProcess = NULL;
	HMODULE hModule = NULL;
	DWORD   cbNeeded = 0;
	WCHAR   lpFilePath[_MAX_PATH] = { NULL };

	//HWND hwnd = NULL;


	if (hwnd != NULL) {

		// プロセスIDを取得する
		dwThId = GetWindowThreadProcessId(hwnd, &dwProcId);

		// 該当プロセスをオープン
		hProcess = OpenProcess(PROCESS_QUERY_INFORMATION | PROCESS_VM_READ, FALSE, dwProcId);

		if (hProcess == NULL)
		{
			MessageBoxW(_T("OpenProcess 失敗"));
		}
		else {	// x64でビルドしないと失敗する

			if (EnumProcessModulesEx(hProcess, &hModule, sizeof(hModule), &cbNeeded, LIST_MODULES_ALL) == FALSE)
			{
				MessageBoxW(_T("EnumProcessModules 失敗"));
			}
			else {

				GetModuleFileNameEx(hProcess, hModule, (LPWSTR)lpFilePath, _MAX_PATH);

				//sprintf_s(sProcessMsg, TEXT(" - プロセス名:%s"), lpFilePath);
				sProcessMsg.Format(_T(" - プロセス名:%s"), (LPCWSTR)lpFilePath);
				*msg += sProcessMsg;
			}
		}
	}
}

void CClipBoardInspectorDlg::OnEnChangeRichedit21()
{
	// TODO: これが RICHEDIT コントロールの場合、このコントロールが
	// この通知を送信するには、CDialogEx::OnInitDialog() 関数をオーバーライドし、
	// CRichEditCtrl().SetEventMask() を関数し呼び出します。
	// OR 状態の ENM_CHANGE フラグをマスクに入れて呼び出す必要があります。

	// TODO: ここにコントロール通知ハンドラー コードを追加してください。
}


void CClipBoardInspectorDlg::SetTextWithFormat(CString sMsg, bool bFound)
{
	static long countLF = 0;

	//メッセージを最後尾に追加
	int nLen = m_pCRichEdit->GetWindowTextLength();
	m_pCRichEdit->SetFocus();
	m_pCRichEdit->SetSel(nLen, nLen);
	m_pCRichEdit->ReplaceSel(sMsg);

	//追加した部分のみ文字色・文字の太さを設定する
	int length = sMsg.GetLength();
	int nEnd = m_pCRichEdit->GetWindowTextLength();

	//SetSel()は改行を文字数から除いてカウントしてしまうので、
	//改行数分減算する(※リッチエディットコントロールのバグ)
	countLF++;
	m_pCRichEdit->SetSel(nEnd - length - countLF, nEnd - countLF);

	CHARFORMAT cf = { sizeof(CHARFORMAT) };
	cf.dwMask = CFM_COLOR | CFM_BOLD;

	//条件にマッチしている時のみ赤・太字
	if (bFound) {
		cf.dwEffects = CFE_BOLD;
		cf.crTextColor = RGB(255, 0, 0);
	}
	else {
		cf.dwEffects = 0;
		cf.crTextColor = RGB(0, 0, 0);
	}

	//色を反映
	m_pCRichEdit->SetSelectionCharFormat(cf);

	//選択をはずす
	m_pCRichEdit->SetSel(-1, 0);
}

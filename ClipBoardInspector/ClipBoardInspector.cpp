//
//  Excel2016で選択したアクティブセルが太枠表示されない原因調査ツール
//
//
// グローバルフックを使用して、フック開始後下記メッセージのログを表示します。
// ・WM_CLIPBOARDUPDATE
// ・WM_PASTE
// ・WM_RENDERFORMAT
//
// Excel2016で選択したアクティブセルが太枠表示されない現象は
// WM_PASTEメッセージが発生しないのにWM_RENDERFORMATメッセージが発生した時に起きることから、
// 条件がマッチしたら自動的にフックが停止するようになっています。
// 
// 下記2プロジェクトで構成されます。
// ・ClipBoardInspector.exe
//   - ClipBoardInspector.cpp メイン
//   - ClipBoardInspectorDlg.cpp ダイアログ表示、フック操作、メッセージハンドリング
// ・DllHookMessage.dll
//   - dllmain.cpp グローバルフック処理
// 
//
// 2019.12.02 新規作成
// 2019.12.14 コメント整理
//
//
// Copy Right (C) Hiroyasu Watanabe 2019.12.02
//




// ClipBoardInspector.cpp : アプリケーションのクラス動作を定義します。
//

#include "pch.h"
#include "framework.h"
#include "ClipBoardInspector.h"
#include "ClipBoardInspectorDlg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// CClipBoardInspectorApp

BEGIN_MESSAGE_MAP(CClipBoardInspectorApp, CWinApp)
	ON_COMMAND(ID_HELP, &CWinApp::OnHelp)
END_MESSAGE_MAP()


// CClipBoardInspectorApp の構築

CClipBoardInspectorApp::CClipBoardInspectorApp()
{
	// 再起動マネージャーをサポートします
	m_dwRestartManagerSupportFlags = AFX_RESTART_MANAGER_SUPPORT_RESTART;

	// TODO: この位置に構築用コードを追加してください。
	// ここに InitInstance 中の重要な初期化処理をすべて記述してください。
}


// 唯一の CClipBoardInspectorApp オブジェクト

CClipBoardInspectorApp theApp;


// CClipBoardInspectorApp の初期化

BOOL CClipBoardInspectorApp::InitInstance()
{
	// アプリケーション マニフェストが visual スタイルを有効にするために、
	// ComCtl32.dll Version 6 以降の使用を指定する場合は、
	// Windows XP に InitCommonControlsEx() が必要です。さもなければ、ウィンドウ作成はすべて失敗します。
	INITCOMMONCONTROLSEX InitCtrls;
	InitCtrls.dwSize = sizeof(InitCtrls);
	// アプリケーションで使用するすべてのコモン コントロール クラスを含めるには、
	// これを設定します。
	InitCtrls.dwICC = ICC_WIN95_CLASSES;
	InitCommonControlsEx(&InitCtrls);

	CWinApp::InitInstance();

	AfxEnableControlContainer();

	// ダイアログにシェル ツリー ビューまたはシェル リスト ビュー コントロールが
	// 含まれている場合にシェル マネージャーを作成します。
	CShellManager *pShellManager = new CShellManager;

	// MFC コントロールでテーマを有効にするために、"Windows ネイティブ" のビジュアル マネージャーをアクティブ化
	CMFCVisualManager::SetDefaultManager(RUNTIME_CLASS(CMFCVisualManagerWindows));

	// 標準初期化
	// これらの機能を使わずに最終的な実行可能ファイルの
	// サイズを縮小したい場合は、以下から不要な初期化
	// ルーチンを削除してください。
	// 設定が格納されているレジストリ キーを変更します。
	// TODO: 会社名または組織名などの適切な文字列に
	// この文字列を変更してください。
	SetRegistryKey(_T("アプリケーション ウィザードで生成されたローカル アプリケーション"));

	//RichEditコントロール初期化
	AfxInitRichEdit2();


	CClipBoardInspectorDlg dlg;
	m_pMainWnd = &dlg;
	INT_PTR nResponse = dlg.DoModal();
	if (nResponse == IDOK)
	{
		// TODO: ダイアログが <OK> で消された時のコードを
		//  記述してください。
	}
	else if (nResponse == IDCANCEL)
	{
		// TODO: ダイアログが <キャンセル> で消された時のコードを
		//  記述してください。
	}
	else if (nResponse == -1)
	{
		TRACE(traceAppMsg, 0, "警告: ダイアログの作成に失敗しました。アプリケーションは予期せずに終了します。\n");
		TRACE(traceAppMsg, 0, "警告: ダイアログで MFC コントロールを使用している場合、#define _AFX_NO_MFC_CONTROLS_IN_DIALOGS を指定できません。\n");
	}

	// 上で作成されたシェル マネージャーを削除します。
	if (pShellManager != nullptr)
	{
		delete pShellManager;
	}

#if !defined(_AFXDLL) && !defined(_AFX_NO_MFC_CONTROLS_IN_DIALOGS)
	ControlBarCleanUp();
#endif

	// ダイアログは閉じられました。アプリケーションのメッセージ ポンプを開始しないで
	//  アプリケーションを終了するために FALSE を返してください。
	return FALSE;
}



// ClipBoardInspectorDlg.h : ヘッダー ファイル
//

#pragma once

#define DllExport    __declspec( dllexport )
// 宣言
#define WM_HOOK (WM_USER+10)

// フック開始
DllExport void HookStart(HWND hWnd);
// フック停止
DllExport void HookStop(void);

//DllExport LRESULT CALLBACK GetMsgProc(int code, WPARAM wParam, LPARAM lParam);
//DllExport LRESULT CALLBACK WndProc(int code, WPARAM wParam, LPARAM lParam);

// CClipBoardInspectorDlg ダイアログ
class CClipBoardInspectorDlg : public CDialogEx
{
// コンストラクション
public:
	CClipBoardInspectorDlg(CWnd* pParent = nullptr);	// 標準コンストラクター
	~CClipBoardInspectorDlg();

// ダイアログ データ
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_CLIPBOARDINSPECTOR_DIALOG };
#endif

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV サポート


// 実装
protected:
	HICON m_hIcon;
	//HMODULE hDll;

	CRichEditCtrl* m_pCRichEdit;

	// 生成された、メッセージ割り当て関数
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	DECLARE_MESSAGE_MAP()

public:
	afx_msg void OnBnClickedButton1();
	afx_msg void OnDrawClipboard();
	afx_msg void OnBnClickedOk();
	afx_msg void OnBnClickedButton2();
	afx_msg void OnEnChangeRichedit21();

protected:
	afx_msg LRESULT OnHook(WPARAM wParam, LPARAM lParam);
	void SetProcessName(HWND hwnd, CString* msg);
	void SetTextWithFormat(CString sMsg, bool bFound);
};

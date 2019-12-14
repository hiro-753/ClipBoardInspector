// dllmain.cpp : DLL アプリケーションのエントリ ポイント
#include "pch.h"
#include <tchar.h>
#include <stdio.h>
#include <stdlib.h>
#include <Psapi.h>

#define DllExport    __declspec( dllexport )

// 宣言
#define WM_HOOK (WM_USER+10)
// フック開始用
DllExport void HookStart(HWND hWnd);
// フック停止用
DllExport void HookStop(void);
// フック処理用 DLLにする事によってグローバルフックが出来る
DllExport LRESULT CALLBACK GetMsgProc(int code, WPARAM wParam, LPARAM lParam);
DllExport LRESULT CALLBACK WndProc(int code, WPARAM wParam, LPARAM lParam);

//---------------------------------------------------------------------------
// 必ず初期化しないと動いてくれない
#pragma data_seg(".shared")
HHOOK g_hHookByGetMsgProc = NULL;   // フックハンドル
HHOOK g_hHookByWndProc = NULL;   // フックハンドル
HWND  g_hWnd = NULL; // 呼び出し元のハンドル
#pragma data_seg()
#pragma comment(linker, "/section:.shared,rws")

HINSTANCE g_hInst;


//---------------------------------------------------------------------------
// DLLエントリポイント
BOOL APIENTRY DllMain(HANDLE hModule, DWORD  ul_reason_for_call, LPVOID lpReserved)
{
	if (ul_reason_for_call == DLL_PROCESS_ATTACH) {
		g_hInst = (HINSTANCE)hModule;   // DLLモジュールのハンドル取得
	}
	return TRUE;
}
//---------------------------------------------------------------------------
// フック開始
DllExport void HookStart(HWND hWnd)
{
	g_hWnd = hWnd;  //呼び出し元のハンドル

	//マウスメッセージのみ
	//g_hHook = SetWindowsHookEx(WH_MOUSE, (HOOKPROC)MouseProc, g_hInst, 0);
		
	//PostMessageのみ
	g_hHookByGetMsgProc = SetWindowsHookEx(WH_GETMESSAGE, (HOOKPROC)GetMsgProc, g_hInst, 0);
	
#ifdef _DEBUG
	if (g_hHookByGetMsgProc == NULL) {
		MessageBox(NULL, TEXT("WH_GETMESSAGE フック開始は失敗しました"), TEXT("HookStart"), MB_OK);

	}
	else {
		MessageBox(NULL, TEXT("WH_GETMESSAGE フック開始は成功しました"), TEXT("HookStart"), MB_OK);
	}
#endif


	//SendMessageのみ
	g_hHookByWndProc = SetWindowsHookEx(WH_CALLWNDPROC, (HOOKPROC)WndProc, g_hInst, 0);

#ifdef _DEBUG
	if (g_hHookByWndProc == NULL) {
		MessageBox(NULL, TEXT("WH_CALLWNDPROC フック開始は失敗しました"), TEXT("HookStart"), MB_OK);

	}
	else {
		MessageBox(NULL, TEXT("WH_CALLWNDPROC フック開始は成功しました"), TEXT("HookStart"), MB_OK);
	}
#endif

}
//---------------------------------------------------------------------------
// フック処理
DllExport LRESULT CALLBACK GetMsgProc(int code, WPARAM wParam, LPARAM lParam)
{

#ifdef _DEBUG
	char sMsg[100];     // メッセージ表示用
#endif

	HWND hwnd;

	if (code < 0) {
		return CallNextHookEx(g_hHookByGetMsgProc, code, wParam, lParam);
	}
	
	MSG* getMsg;
	getMsg = (MSG*)lParam;

	//メッセージ発生のプロセスを取り出すため
	hwnd = getMsg->hwnd;

	if (getMsg->message == WM_RBUTTONDOWN) {

#ifdef _DEBUG
		sprintf_s(sMsg, "右クリックされました!");
		//ここでダイアログを開くと右メニューは表示されない
		//MessageBox(hwnd, sMsg, TEXT("GetMessage"), MB_OK);
		//SendMessage(g_hWnd, WM_HOOK, (WPARAM)WM_RBUTTONDOWN, (LPARAM)hwnd);
#endif

	}
	else if (getMsg->message == WM_CLIPBOARDUPDATE) {

#ifdef _DEBUG
		sprintf_s(sMsg, "WM_CLIPBOARDUPDATE");
		MessageBox(hwnd, sMsg, TEXT("GetMessage"), MB_OK);
#endif
		//ここでLPALAMでhwndを渡す(メッセージ発生のプロセスを取り出すため)
		SendMessage(g_hWnd, WM_HOOK, (WPARAM)WM_CLIPBOARDUPDATE, (LPARAM)hwnd);
	}
			

	// 最後は次のフックへ渡します
	return CallNextHookEx(g_hHookByGetMsgProc, code, wParam, lParam);
}

// フック処理
DllExport LRESULT CALLBACK WndProc(int code, WPARAM wParam, LPARAM lParam)
{

#ifdef _DEBUG
	char sMsg[100];     // メッセージ表示用
#endif

	HWND hwnd;

	if (code < 0) {
		return CallNextHookEx(g_hHookByWndProc, code, wParam, lParam);
	}

	CWPSTRUCT* winMsg;
	winMsg = (CWPSTRUCT*)lParam;

	//メッセージ発生のプロセスを取り出すため
	hwnd = winMsg->hwnd;

	if (winMsg->message == WM_RENDERFORMAT) {
		// WM_RENDERFORMAT

#ifdef _DEBUG
		sprintf_s(sMsg, "WM_RENDERFORMAT");
		MessageBox(hwnd, sMsg, TEXT("GetMessage"), MB_OK);
#endif

		SendMessage(g_hWnd, WM_HOOK, (WPARAM)WM_RENDERFORMAT, (LPARAM)hwnd);
	}
	else if (winMsg->message == WM_PASTE) {
		// WM_PASTE

#ifdef _DEBUG
		sprintf_s(sMsg, "WM_PASTE");
		MessageBox(hwnd, sMsg, TEXT("GetMessage"), MB_OK);
#endif

		SendMessage(g_hWnd, WM_HOOK, (WPARAM)WM_PASTE, (LPARAM)hwnd);
	}

	// 最後は次のフックへ渡します
	return CallNextHookEx(g_hHookByWndProc, code, wParam, lParam);
}


//---------------------------------------------------------------------------
// フック停止
DllExport void HookStop(void)
{
	BOOL bResult;

	if (g_hHookByGetMsgProc != NULL) {

		bResult = UnhookWindowsHookEx(g_hHookByGetMsgProc);

#ifdef _DEBUG
		if (bResult == 0) {
			MessageBox(NULL, TEXT("WH_GETMESSAGE フック解除は失敗しました"), TEXT("HookStop"), MB_OK);

		}
		else {
			MessageBox(NULL, TEXT("WH_GETMESSAGE フック解除は成功しました"), TEXT("HookStop"), MB_OK);
		}
#endif
	}

	if (g_hHookByWndProc != NULL) {
		
		bResult = UnhookWindowsHookEx(g_hHookByWndProc);

#ifdef _DEBUG
		if (bResult == 0) {
			MessageBox(NULL, TEXT("WH_CALLWNDPROC フック解除は失敗しました"), TEXT("HookStop"), MB_OK);
		}
		else {
			MessageBox(NULL, TEXT("WH_CALLWNDPROC フック解除は成功しました"), TEXT("HookStop"), MB_OK);
		}
#endif

	}
}

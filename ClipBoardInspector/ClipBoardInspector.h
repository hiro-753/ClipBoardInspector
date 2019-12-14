
// ClipBoardInspector.h : PROJECT_NAME アプリケーションのメイン ヘッダー ファイルです
//

#pragma once

#ifndef __AFXWIN_H__
	#error "PCH に対してこのファイルをインクルードする前に 'pch.h' をインクルードしてください"
#endif

#include "resource.h"		// メイン シンボル


// CClipBoardInspectorApp:
// このクラスの実装については、ClipBoardInspector.cpp を参照してください
//

class CClipBoardInspectorApp : public CWinApp
{
public:
	CClipBoardInspectorApp();

// オーバーライド
public:
	virtual BOOL InitInstance();

// 実装

	DECLARE_MESSAGE_MAP()
};

extern CClipBoardInspectorApp theApp;

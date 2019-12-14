# ClipBoardInspector
Excel2016で選択したアクティブセルが太枠表示されない原因調査ツール<br>
<br>
グローバルフックを使用して、フック開始後下記メッセージのログを表示します。<br>
 ・WM_CLIPBOARDUPDATE<br>
 ・WM_PASTE<br>
 ・WM_RENDERFORMAT<br>
<br>
Excel2016で選択したアクティブセルが太枠表示されない現象は<br>
WM_PASTEメッセージが発生しないのにWM_RENDERFORMATメッセージが発生した時に起きることから、<br>
条件がマッチしたら自動的にフックが停止するようになっています。<br>
<br>
下記2プロジェクトで構成されます。<br>
・ClipBoardInspector.exe<br>
　ClipBoardInspector.cpp メイン<br>
　ClipBoardInspectorDlg.cpp ダイアログ表示、フック操作、メッセージハンドリング<br>
・DllHookMessage.dll<br>
　dllmain.cpp グローバルフック処理<br>
<br>
仕様などは、下記記事をご参照ください。<br>
<br>
・Excel 2016で選択したアクティブセルが太枠表示されない原因調査ツールリリース<br>
　https://ameblo.jp/nabezou3/entry-12551286051.html<br>
<br>
※お約束ですが、重要なファイルは保存した後、実行は自己責任でお願いします。(^^)
<br>
<br>
※参考資料URL：<br>
\[BCB Client\]<br>
・[マウスのグローバルフック](http://bcb.client.jp/tips/008_mouse_ghook_dll.html)<br>

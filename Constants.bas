Attribute VB_Name = "Constants"
'## 処理に使用する定数 ##
Public Const TITLE_FILESELECTDIALOG As String = "NativeIOSコンフィグファイル,*.*"
Public Const NGWORD_NOIOSEQUIVALENT As String = "! No IOS equivalent for the CatOS command"
Public Const NGWORD_TRANSLATE_ERROR As String = "! Translate Error: Translation Error"
Public Const NO_TEXT As String = ""
Public Const SPLITTER As String = "    "
Public Const EXTENSION As String = ".ios"
Public Const CONFIG_BEGIN As String = "begin"
Public Const CONFIG_END As String = "end"
Public Const LOG_FILE_NAME As String = "c:\err.log"
Public Const STR_SHARP As String = "####"
Public Const STR_EXCLAMATIONMARK As String = "!"
Public Const STR_YEN As String = "\"
Public Const STR_SPACE As String = " "
Public Const STR_QUOAT As String = "'"

'## ラベル ##
Public Const STR_CURRENT_FILENAME As String = "処理中ファイル名： "

'## エラーメッセージ群 ##
Public Const ERR_MSG_AGREE_FILE_NAME As String = "変換前および変換後のファイルに同じものが選択されています。"
Public Const ERR_MSG_NOT_SELECTED_FILE As String = "ファイルが選択されていません。"
Public Const ERR_MSG_NOT_DIFF_FILE As String = "読み込んだファイルはdiffファイルではありません。処理を終了します。"
Public Const ERR_MSG_NOT_CATOS_FILE As String = " にbegin-endの記述がありませんでした。"

'## 警告メッセージ群 ##
Public Const ALRT_MSG_EXIST_ERROR As String = "読み込めないTeraTermログファイルがありました。詳細は「" + LOG_FILE_NAME + "」を確認して下さい。"

'## 各オブジェクトの表示テキスト ##
Public Const TITLE_FRAME_DIFF_TO_IOS As String = "Diffからエラー箇所のCatOSコマンドを抽出"
Public Const TITLE_FRAME_TERATERM_TO_IOS As String = "TeraTermのログファイルからIOSコマンド箇所を抽出して「.ios」ファイルを作成"
Public Const TITLE_LABEL_DIFF_TO_IOS As String = "シスコの変換ツールで出力したDiffの中からエラーコマンドを抽出したいファイルを選択して下さい。"
Public Const TITLE_LABEL_TERATERM_TO_IOS As String = "Sup720でshow runした時にTeraTermで取得したログファイルが格納されているフォルダ内のファイルを選択して下さい。"

'## 処理節目でのメッセージ ##
Public Const MSG_PROCESS_END As String = "処理完了しやしたっっ！"
Public Const MSG_CONVERT_FILES_END As String = "TeraTermログからコンフィグ箇所の抽出処理終了しました。"


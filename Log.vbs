Option Explicit

'-------------------------------------------------------------------------------------------
' 特定のファイルへのログ出力
'-------------------------------------------------------------------------------------------
' strMessage - 出力する文字列
' strTargetFile - 出力対象ファイル
'-------------------------------------------------------------------------------------------
Function Log_Add(ByVal strMessage, ByVal strTargetFile)
    On Error Resume Next
    
    Const ForAppending = 8 ' 追記モード
    Dim fso, fi
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    'ファイルを開く
    'もしも存在しない場合には作成する
    Set fi = fso.OpenTextFile(strTargetFile, ForAppending, true)
    strMessage = Date() & " " & Time() & ": " & strMessage
    fi.WriteLine (strMessage) 'ログを書き込む
    Set fi = Nothing
    
    If Err.Number <> 0 Then
        Wscript.Echo "ログ出力中にエラーが発生しました。ログの出力先、権限等を確認してください。"
        WScript.Echo "エラー : " & Err.Number & vbCrLf & Err.Description
    End If
End Function

Option Explicit

'**********************************************************
'**
'** MIXI 専用関数
'**
'**********************************************************

'
'コメントを指定したURLのコミュニティに書き込む
'
'戻り値：1:正常 2:待機時間超過
Function Mixi_PostCommentHere(ByRef ie, comment, url)
  Mixi_PostCommentHere = 1

  'ページに遷移して書き込み
  IE_Navigate ie, url
  If IE_Wait(ie) < 1 Then
    Mixi_PostCommentHere = 2
  End If
  If Mixi_PostComment(ie, GetCommentValue()) = 2 Then
    Mixi_PostCommentHere = 2
  End If
  If IE_Wait(ie) < 1 Then
    Mixi_PostCommentHere = 2
  End If
  
End Function

'
'コメントを書き込む
'
'戻り値：1:正常 2:待機時間超過
Function Mixi_PostComment(ByRef ie, comment)
  Mixi_PostComment = 1
  'コメントを入力
  Dim cmt : Set cmt = ie.document.getElementsByName("comment")(0)
  cmt.value = comment
  '投稿する
  Dim frm : Set frm = ie.document.getElementsByName("bbs_comment_form")(0)
  IE_DocumentCompletedUrl = "http://mixi.jp/" & frm.action
  IE_DocumentCompleted = False
  ie.document.bbs_comment_form.submit()
  If IE_Wait(ie) < 1 Then
    Mixi_PostComment = 2
  End If

  '確認画面は廃止され、直接投稿されるようになりました
	'
  '確認画面で【書き込む】ボタンを押下
  'Dim obj : Set obj = IE_GetBtnByValue(ie.document, "書き込む")
  'MsgBox obj.value
  'obj.Click
  
End Function

'
'ログイン画面か確認する
'
Function Mixi_IsLoginWindow(ByRef ie)
  Mixi_IsLoginWindow = False

  '「パスワード」入力欄があればログイン画面
  If ie.document.getElementsByName("password").length = 1 Then
    Mixi_IsLoginWindow = True
  End If 
  
End Function



Option Explicit

'**********************************************************
'**
'** IE制御汎用関数ファイル
'**
'**********************************************************

'
'画面遷移
'
Function IE_Navigate(ByRef ie, url)
  IE_DocumentCompleted = False
  IE_DocumentCompletedUrl = url
  ie.navigate url
End Function

'
'指定した名前のボタンオブジェクトを取得する
'
Function IE_GetBtnByValue(ByRef document, value)
  Dim retObj : Set retObj = Nothing
  
  Dim input
  '<INPUT>タグを検索
  For Each input In document.All.Tags("INPUT")
    '指定値と同じなら、オブジェクトを返す
    'MsgBox input.value & "=" & value
    If input.value = value Then
      Set retObj = input
      Exit For
    End If
  Next

  Set IE_GetBtnByValue = retObj
End Function

'
'指定した名前のリンクオブジェクトを取得する
'
Function IE_GetLinkByText(document, text)
  Dim retObj : Set retObj = Nothing
  
  Dim target
  '<A>タグを検索
  For Each target In document.All.Tags("A")
    '指定値と同じなら、オブジェクトを返す
    If target.innerHTML = text Then
      Set retObj = target
      Exit For
    End If
  Next

  Set IE_GetLinkByText = retObj
End Function

'指定した名前のフレームを取得する
Function IE_GetFrameByName(frames, theName)
  Set IE_GetFrameByName = Nothing

  Dim i
  For i=0 To frames.Length-1
    If frames.item(i).Name = theName Then
      Set IE_GetFrameByName = frames.item(i)
    End If
  Next

End Function

'全フレームの名前を再帰的に取得する
Function IE_GetFrameNames(ByRef document)

  Dim names : names = ""
  Dim i
  For i = 0 To document.frames.length - 1
    names = names & IE_GetFrameNames(document.frames.item(i).document)
    names = names & document.frames.item(i).name & vbCrLf
  Next


  IE_GetFrameNames = names
End Function

'
'IEの処理を待つ
'
'戻り値：1:正常 0:異常 -1:待機時間超過
Function IE_Wait(ie)
  IE_Wait = 1

  Dim maxWait : maxWait = 1000 * 10
  Dim perWait : perWait = 5000
  Dim waited  : waited = 0
  Do While (ie.Busy = True) or (IE_DocumentCompleted = False) or (IE_DownloadCompleted = False)
  'Do While (ie.Busy = True) And (ie.readystate <> 4)
    wscript.sleep(perWait)
    waited = waited + perWait
    If maxWait < waited Then
      IE_DocumentCompleted = True
      IE_DownloadCompleted = True
      IE_Wait = -1
    End If
  Loop

End Function

'このフラグがfalseの間は，IEがビジー状態であるとする。
Dim IE_DocumentCompleted : IE_DocumentCompleted = True
Dim IE_DocumentCompletedUrl : IE_DocumentCompletedUrl = ""
Dim IE_DownloadCompleted : IE_DownloadCompleted = True
'IE、画面遷移完了イベント
Sub IE_NavigateComplete2(ByVal pDisp, URL)
    'MsgBox "IE_NavigateComplete2"
End Sub
Sub IE_DocumentComplete(ByVal pDisp, URL)
  If IE_DocumentCompletedUrl = url Then
    MsgBox "IE_DocumentCompleted : " & URL
    IE_DocumentCompleted = True
  End If
End Sub
'IE、画面リロード完了イベント
Sub IE_DownloadComplete()
  'MsgBox "IE_DownloadCompleted"
  'IE_DownloadCompleted = True
End Sub



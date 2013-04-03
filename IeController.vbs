Option Explicit

'**********************************************************
'**
'** IE����ėp�֐��t�@�C��
'**
'**********************************************************

'
'��ʑJ��
'
Function IE_Navigate(ByRef ie, url)
  IE_DocumentCompleted = False
  IE_DocumentCompletedUrl = url
  ie.navigate url
End Function

'
'�w�肵�����O�̃{�^���I�u�W�F�N�g���擾����
'
Function IE_GetBtnByValue(ByRef document, value)
  Dim retObj : Set retObj = Nothing
  
  Dim input
  '<INPUT>�^�O������
  For Each input In document.All.Tags("INPUT")
    '�w��l�Ɠ����Ȃ�A�I�u�W�F�N�g��Ԃ�
    'MsgBox input.value & "=" & value
    If input.value = value Then
      Set retObj = input
      Exit For
    End If
  Next

  Set IE_GetBtnByValue = retObj
End Function

'
'�w�肵�����O�̃����N�I�u�W�F�N�g���擾����
'
Function IE_GetLinkByText(document, text)
  Dim retObj : Set retObj = Nothing
  
  Dim target
  '<A>�^�O������
  For Each target In document.All.Tags("A")
    '�w��l�Ɠ����Ȃ�A�I�u�W�F�N�g��Ԃ�
    If target.innerHTML = text Then
      Set retObj = target
      Exit For
    End If
  Next

  Set IE_GetLinkByText = retObj
End Function

'�w�肵�����O�̃t���[�����擾����
Function IE_GetFrameByName(frames, theName)
  Set IE_GetFrameByName = Nothing

  Dim i
  For i=0 To frames.Length-1
    If frames.item(i).Name = theName Then
      Set IE_GetFrameByName = frames.item(i)
    End If
  Next

End Function

'�S�t���[���̖��O���ċA�I�Ɏ擾����
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
'IE�̏�����҂�
'
'�߂�l�F1:���� 0:�ُ� -1:�ҋ@���Ԓ���
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

'���̃t���O��false�̊Ԃ́CIE���r�W�[��Ԃł���Ƃ���B
Dim IE_DocumentCompleted : IE_DocumentCompleted = True
Dim IE_DocumentCompletedUrl : IE_DocumentCompletedUrl = ""
Dim IE_DownloadCompleted : IE_DownloadCompleted = True
'IE�A��ʑJ�ڊ����C�x���g
Sub IE_NavigateComplete2(ByVal pDisp, URL)
    'MsgBox "IE_NavigateComplete2"
End Sub
Sub IE_DocumentComplete(ByVal pDisp, URL)
  If IE_DocumentCompletedUrl = url Then
    MsgBox "IE_DocumentCompleted : " & URL
    IE_DocumentCompleted = True
  End If
End Sub
'IE�A��ʃ����[�h�����C�x���g
Sub IE_DownloadComplete()
  'MsgBox "IE_DownloadCompleted"
  'IE_DownloadCompleted = True
End Sub



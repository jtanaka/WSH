Option Explicit

'**********************************************************
'**
'** MIXI ��p�֐�
'**
'**********************************************************

'
'�R�����g���w�肵��URL�̃R�~���j�e�B�ɏ�������
'
'�߂�l�F1:���� 2:�ҋ@���Ԓ���
Function Mixi_PostCommentHere(ByRef ie, comment, url)
  Mixi_PostCommentHere = 1

  '�y�[�W�ɑJ�ڂ��ď�������
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
'�R�����g����������
'
'�߂�l�F1:���� 2:�ҋ@���Ԓ���
Function Mixi_PostComment(ByRef ie, comment)
  Mixi_PostComment = 1
  '�R�����g�����
  ie.document.bbs_comment_form.item("comment").value = comment
  '���e����
  IE_DocumentCompletedUrl = "http://mixi.jp/" & ie.document.bbs_comment_form.action
  IE_DocumentCompleted = False
  ie.document.bbs_comment_form.submit()
  If IE_Wait(ie) < 1 Then
    Mixi_PostComment = 2
  End If

  '�m�F��ʂŁy�������ށz�{�^��������
  Dim obj : Set obj = IE_GetBtnByValue(ie, "��������")
  'MsgBox obj.value
  obj.Click
  
End Function

'
'���O�C����ʂ��m�F����
'
Function Mixi_IsLoginWindow(ByRef ie)
  Mixi_IsLoginWindow = False

  '�u�p�X���[�h�v���͗�������΃��O�C�����
  If ie.document.forms(0).item("password") Is Nothing = False Then
    Mixi_IsLoginWindow = True
  End If 
  
End Function



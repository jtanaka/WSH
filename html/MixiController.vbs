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
  Dim cmt : Set cmt = ie.document.getElementsByName("comment")(0)
  cmt.value = comment
  '���e����
  Dim frm : Set frm = ie.document.getElementsByName("bbs_comment_form")(0)
  IE_DocumentCompletedUrl = "http://mixi.jp/" & frm.action
  IE_DocumentCompleted = False
  ie.document.bbs_comment_form.submit()
  If IE_Wait(ie) < 1 Then
    Mixi_PostComment = 2
  End If

  '�m�F��ʂ͔p�~����A���ړ��e�����悤�ɂȂ�܂���
	'
  '�m�F��ʂŁy�������ށz�{�^��������
  'Dim obj : Set obj = IE_GetBtnByValue(ie.document, "��������")
  'MsgBox obj.value
  'obj.Click
  
End Function

'
'���O�C����ʂ��m�F����
'
Function Mixi_IsLoginWindow(ByRef ie)
  Mixi_IsLoginWindow = False

  '�u�p�X���[�h�v���͗�������΃��O�C�����
  If ie.document.getElementsByName("password").length = 1 Then
    Mixi_IsLoginWindow = True
  End If 
  
End Function



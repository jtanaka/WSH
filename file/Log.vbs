Option Explicit

'-------------------------------------------------------------------------------------------
' ����̃t�@�C���ւ̃��O�o��
'-------------------------------------------------------------------------------------------
' strMessage - �o�͂��镶����
' strTargetFile - �o�͑Ώۃt�@�C��
'-------------------------------------------------------------------------------------------
Function Log_Add(ByVal strMessage, ByVal strTargetFile)
    On Error Resume Next
    
    Const ForAppending = 8 ' �ǋL���[�h
    Dim fso, fi
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    '�t�@�C�����J��
    '���������݂��Ȃ��ꍇ�ɂ͍쐬����
    Set fi = fso.OpenTextFile(strTargetFile, ForAppending, true)
    strMessage = Date() & " " & Time() & ": " & strMessage
    fi.WriteLine (strMessage) '���O����������
    Set fi = Nothing
    
    If Err.Number <> 0 Then
        Wscript.Echo "���O�o�͒��ɃG���[���������܂����B���O�̏o�͐�A���������m�F���Ă��������B"
        WScript.Echo "�G���[ : " & Err.Number & vbCrLf & Err.Description
    End If
End Function

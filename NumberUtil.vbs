Option Explicit

'**********************************************************
'**
'** ���l�n�ėp�֐�
'**
'**********************************************************

'�����_���Ȑ��l���擾���܂�
Function Num_IntRand(min, max)
  Randomize

  'intValue �ɂ� 1 �` max �̗���������܂��B
  Dim intValue
  intValue = Int( (max - min + 1) * Rnd + min )

  Num_IntRand = intValue
End Function



Option Explicit

'**********************************************************
'**
'** 数値系汎用関数
'**
'**********************************************************

'ランダムな数値を取得します
Function Num_IntRand(min, max)
  Randomize

  'intValue には 1 ～ max の乱数が入ります。
  Dim intValue
  intValue = Int( (max - min + 1) * Rnd + min )

  Num_IntRand = intValue
End Function



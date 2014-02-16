'Option Explicit


'漢字設定
strK1 = "一"
strK2 = "二"
strK3 = "三"
strK4 = "四"
strK5 = "五"
strK6 = "六"
strK7 = "七"
strK8 = "八"
strK9 = "九"
strK10 = "十"
strK100 = "百"
strK1000 = "千"

'あるふぁべっと設定
strA1 = "えー"
strA2 = "びー"
strA3 = "しー"
strA4 = "でぃー"
strA5 = "いー"
strA6= "えふ"
strA7 = "じー"
strA8 = "えっち"
strA9 = "あい"
strA10 = "じぇー"
strA11 = "けー"
strA12 = "える"
strA13 = "えむ"
strA14 = "えぬ"
strA15 = "おー"
strA16 = "ぴー"
strA17 = "きゅー"
strA18 = "あーる"
strA19 = "えす"
strA20 = "てぃー"
strA21 = "ゆー"
strA22 = "ぶい"
strA23 = "だぶりゅ"
strA24 = "えっくす"
strA25 = "わい"
strA26 = "ぜっと"

'反田
strTanda = "反田"

'えるたそ数
L = 1


'関数定義
Function Kanji(num, keta)

	If num > 9 Then
		str = Kanji(num \ 10 ,keta + 1)
	End If


	Select Case num Mod 10
		Case 1
			If keta = 1 Then str = str & strK1
		Case 2
			str = str & strK2
		Case 3
			str = str & strK3
		Case 4
			str = str & strK4
		Case 5
			str = str & strK5
		Case 6
			str = str & strK6
		Case 7
			str = str & strK7
		Case 8
			str = str & strK8
		Case 9
			str = str & strK9
		Case 0
		
	End Select
	If num Mod 10 <> 0 Then
		If keta = 2 Then str = str & strK10
		If Keta = 3 Then str = str & strK100
		If Keta = 4 Then str = str & strK1000
	End If
	
	Kanji = str
	
End Function


Function Alphabet(num)
	Select Case num Mod 26
	'アルファベットは26文字なので
	'26の剰余でいまの名前を決める
		Case 1
			Alphabet = strA1
		Case 2 
			Alphabet = strA2
		Case 3
			Alphabet = strA3
		Case 4
			Alphabet = strA4
		Case 5
			Alphabet = strA5
		Case 6
			Alphabet = strA6
		Case 7
			Alphabet = strA7
		Case 8
			Alphabet = strA8
		Case 9
			Alphabet = strA9
		Case 10
			Alphabet = strA10
		Case 11
			Alphabet = strA11
		Case 12
			Alphabet = strA12
		Case 13
			Alphabet = strA13
		Case 14
			Alphabet = strA14
		Case 15
			Alphabet = strA15
		Case 16
			Alphabet = strA16
		Case 17
			Alphabet = strA17
		Case 18
			Alphabet = strA18
		Case 19
			Alphabet = strA19
		Case 20
			Alphabet = strA20
		Case 21
			Alphabet = strA21
		Case 22
			Alphabet = strA22
		Case 23
			Alphabet = strA23
		Case 24
			Alphabet = strA24
		Case 25
			Alphabet = strA25
		Case 0 'Z(26番)は剰余が0
			Alphabet = strA26

		Case Else
			Alphabet = ""
	End Select
End Function


'本体
Do While L < 1001
	strs = Kanji(L, 1) + strTanda + Alphabet(L)
	MsgBox strs
	L = L + 1
Loop

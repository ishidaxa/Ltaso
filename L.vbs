'Option Explicit


'�����ݒ�
strK1 = "��"
strK2 = "��"
strK3 = "�O"
strK4 = "�l"
strK5 = "��"
strK6 = "�Z"
strK7 = "��"
strK8 = "��"
strK9 = "��"
strK10 = "�\"
strK100 = "�S"
strK1000 = "��"

'����ӂ��ׂ��Ɛݒ�
strA1 = "���["
strA2 = "�с["
strA3 = "���["
strA4 = "�ł��["
strA5 = "���["
strA6= "����"
strA7 = "���["
strA8 = "������"
strA9 = "����"
strA10 = "�����["
strA11 = "���["
strA12 = "����"
strA13 = "����"
strA14 = "����"
strA15 = "���["
strA16 = "�ҁ["
strA17 = "����["
strA18 = "���[��"
strA19 = "����"
strA20 = "�Ă��["
strA21 = "��["
strA22 = "�Ԃ�"
strA23 = "���Ԃ��"
strA24 = "��������"
strA25 = "�킢"
strA26 = "������"

'���c
strTanda = "���c"

'���邽����
L = 1


'�֐���`
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
	'�A���t�@�x�b�g��26�����Ȃ̂�
	'26�̏�]�ł��܂̖��O�����߂�
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
		Case 0 'Z(26��)�͏�]��0
			Alphabet = strA26

		Case Else
			Alphabet = ""
	End Select
End Function


'�{��
Do While L < 1001
	strs = Kanji(L, 1) + strTanda + Alphabet(L)
	MsgBox strs
	L = L + 1
Loop

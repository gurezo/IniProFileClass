' @(h) clsIniProfileClass.vb          ver 01.00.00 ('2005.05.19 �ؕ�)
'
' @(s)
' 
'
'
Option Strict Off
Option Explicit On 

Public Class IniProFile

    ''---- ����ini̧�����Ұ��Ǎ��ݗpAPI�錾 ----
    Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" _
    (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer

    ''---- ����ini̧�����Ұ������ݗpAPI�錾 ----
    Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" _
    (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Long



    ' @(f)
    '
    ' �@�\�@�@ :ini̧�ٓǍ��ݏ���
    '
    ' �Ԃ�l�@ :����I�� - 0
    ' �@�@�@    �װ�I�� - �󕶎�
    '
    ' �������@ :ARG1 - ���߽�ţ�ٖ�
    ' �@�@�@    ARG2 - ���ٕ�����
    ' �@�@�@    ARG3 - ���ѕ�����
    '
    ' �@�\���� :
    '
    ' ���l�@�@ :
    '
    Public Function strRead(ByRef FileStr As String, _
                            ByRef TitleStr As String, _
                            ByRef ItemStr As String) As String

        Dim Ret As Integer
        Dim Buff As String

        Const DEF_BUFF_LEN As Integer = &H400
        Buff = New String(CChar(" "), DEF_BUFF_LEN)

        ''���Ұ��擾
        Ret = GetPrivateProfileString(TitleStr, ItemStr, vbNullString, Buff, DEF_BUFF_LEN, FileStr)

        ''���Ұ��擾����I����
        If Ret > 0 Then
            strRead = Strings.Left(Buff, (InStr(Buff, vbNullChar) - 1))
        Else
            strRead = ""
        End If

    End Function

    ' @(f)
    '
    ' �@�\�@�@ :ini̧�ُ����ݏ���
    '
    ' �Ԃ�l�@ :����I�� - 0
    ' �@�@�@    �װ�I�� - �󕶎�
    '
    ' �������@ :ARG1 - ���߽�ţ�ٖ�
    ' �@�@�@    ARG2 - ���ٕ�����
    ' �@�@�@    ARG3 - ���ѕ�����
    ' �@�@�@    ARG4 - �l������
    '
    ' �@�\���� :
    '
    ' ���l�@�@ :
    '
    Function intWrite(ByRef FileStr As String, _
                      ByRef TitleStr As String, _
                      ByRef ItemStr As String, _
                      ByRef ValueStr As String) As Integer

        Dim Ret As Integer

        Ret = WritePrivateProfileString(TitleStr, ItemStr, ValueStr, FileStr)
        intWrite = Ret

    End Function

End Class

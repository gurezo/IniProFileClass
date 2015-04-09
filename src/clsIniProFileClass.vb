' @(h) clsIniProfileClass.vb          ver 01.00.00
'
' @(s)
' 
'
'
Option Strict Off
Option Explicit On 

Public Class IniProFile

    ''---- ｼｽﾃﾑiniﾌｧｲﾙﾊﾟﾗﾒｰﾀ読込み用API宣言 ----
    Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" _
    (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer

    ''---- ｼｽﾃﾑiniﾌｧｲﾙﾊﾟﾗﾒｰﾀ書込み用API宣言 ----
    Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" _
    (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Long



    ' @(f)
    '
    ' 機能　　 :iniﾌｧｲﾙ読込み処理
    '
    ' 返り値　 :正常終了 - 0
    ' 　　　    ｴﾗｰ終了 - 空文字
    '
    ' 引き数　 :ARG1 - ﾌﾙﾊﾟｽ付ﾌｧｲﾙ名
    ' 　　　    ARG2 - ﾀｲﾄﾙ文字列
    ' 　　　    ARG3 - ｱｲﾃﾑ文字列
    '
    ' 機能説明 :
    '
    ' 備考　　 :
    '
    Public Function strRead(ByRef FileStr As String, _
                            ByRef TitleStr As String, _
                            ByRef ItemStr As String) As String

        Dim Ret As Integer
        Dim Buff As String

        Const DEF_BUFF_LEN As Integer = &H400
        Buff = New String(CChar(" "), DEF_BUFF_LEN)

        ''ﾊﾟﾗﾒｰﾀ取得
        Ret = GetPrivateProfileString(TitleStr, ItemStr, vbNullString, Buff, DEF_BUFF_LEN, FileStr)

        ''ﾊﾟﾗﾒｰﾀ取得正常終了時
        If Ret > 0 Then
            strRead = Strings.Left(Buff, (InStr(Buff, vbNullChar) - 1))
        Else
            strRead = ""
        End If

    End Function

    ' @(f)
    '
    ' 機能　　 :iniﾌｧｲﾙ書込み処理
    '
    ' 返り値　 :正常終了 - 0
    ' 　　　    ｴﾗｰ終了 - 空文字
    '
    ' 引き数　 :ARG1 - ﾌﾙﾊﾟｽ付ﾌｧｲﾙ名
    ' 　　　    ARG2 - ﾀｲﾄﾙ文字列
    ' 　　　    ARG3 - ｱｲﾃﾑ文字列
    ' 　　　    ARG4 - 値文字列
    '
    ' 機能説明 :
    '
    ' 備考　　 :
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

Imports Microsoft.Win32
Module Module1
    Dim regKey As RegistryKey = Registry.CurrentUser.OpenSubKey("Software\PTC\Creo_Parametric", True)
    Sub Main()
        Dim SCRFG As Boolean = False
        Dim SCRF = GetAvaliableComPorts()
        If SCRF Is Nothing Then
            Console.Write("'文件没有最近打开的文件列表,按回车键结束")
            Console.ReadLine()
            End
        End If

        Dim str As String
        Dim RFNum As Integer = 0
re:
        For j As Integer = LBound(SCRF, 2) To UBound(SCRF, 2)

            If SCRFG Then
                Console.Write("请稍等,正在处理……" & vbCrLf)
                regKey.DeleteValue(SCRF(0, j).ToString)
                Console.Write("删除完毕,按Enter键结束")
                Console.ReadLine()
                End
            Else
                RFNum += 1
                str = str & RFNum & vbTab & SCRF(0, j) & vbTab & SCRF(1, j) & vbNewLine
            End If
        Next
        If str.Length <> 0 Then
            Console.Write("序号" & vbTab & "最近文件名称" & vbTab & "路径" & vbCrLf)
            Console.WriteLine(str)
            Console.Write("是否需要删除所有的最近文件" & " (Y/N):  ")
            Dim SC = Console.ReadLine
            If UCase(SC) = "Y" Then
                SCRFG = True
                GoTo re
            End If
        End If

    End Sub
    Public Function GetAvaliableComPorts()
        Dim regKey As RegistryKey = Registry.CurrentUser.OpenSubKey("Software\PTC\Creo_Parametric")
        Dim regValueNames() As String = regKey.GetValueNames()
        Dim regValueCount As Integer = regValueNames.Length
        Dim RFGroup(,) As String
        Dim i As Integer = 0
        For Each regValueName As String In regValueNames
            If regValueName.ToString.Contains("RecentFile") Then
                ReDim Preserve RFGroup(1, i)
                RFGroup(0, i) = regValueName
                RFGroup(1, i) = regKey.GetValue(regValueName)
                i += 1
            End If
        Next
        Return RFGroup
    End Function
End Module

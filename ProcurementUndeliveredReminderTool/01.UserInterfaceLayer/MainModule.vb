Imports System.IO
Imports OfficeOpenXml

Module MainModule

    Sub Main()

        Console.Title = $"{My.Application.Info.Title} V{AppSettingHelper.Instance.ProductVersion}"

        Console.WriteLine($"{Now:G}> 通知发送时间: {AppSettingHelper.Instance.SendMsgTime}")
        Console.WriteLine($"{Now:G}> 提前提醒天数: {AppSettingHelper.Instance.AdvanceNoticeDays}")
        Console.WriteLine($"{Now:G}> 上次通知发送时间: {AppSettingHelper.Instance.LastSendDate}")

#Region "初始化"
        ' 单例模式
        Dim tmpProcess = Process.GetCurrentProcess()
        Dim processCount = Process.GetProcessesByName(tmpProcess.ProcessName).Count()
        ' 有多个实例则退出程序
        If processCount > 1 Then
            Exit Sub
        End If

        AppSettingHelper.Instance.Logger.Info("程序启动")
#End Region

        '设置使用方式为个人使用
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial

        Do
            Threading.Thread.Sleep(60 * 1000)

            ' 定时发送
            If Now.Hour <> AppSettingHelper.Instance.SendMsgTime.Hours OrElse
            Now.Minute <> AppSettingHelper.Instance.SendMsgTime.Minutes Then
                Continue Do
            End If

            ' 一天只自动发送一次
            If AppSettingHelper.Instance.LastSendDate = Now.Date Then
                Continue Do
            End If
            AppSettingHelper.Instance.LastSendDate = Now.Date

            AppSettingHelper.Instance.Logger.Info("自动发送通知")
            Console.WriteLine($"{Now:G}> 自动发送通知")

            Console.WriteLine($"{Now:G}> 清理临时文件")
            AppSettingHelper.Instance.ClearTempFiles()

            Console.WriteLine($"{Now:G}> 搜索表单数据")
            Dim tmpDictionary = RemoteDatabaseHelper.GetDocumentItems()

            Console.WriteLine($"{Now:G}> 获取钉钉AccessToken")
            DingTalkHelper.GetAccessToken()

#Region "判断是否有无对应的钉钉账号的ERP用户"
            If Not tmpDictionary.All(Function(s1)
                                         Return AppSettingHelper.Instance.DingTalkUserJobNumberItems.ContainsKey(s1.Key)
                                     End Function) Then

#Region "获取钉钉部门信息"
                Console.WriteLine($"{Now:G}> 获取钉钉部门信息")
                AppSettingHelper.Instance.DingTalkDepartmentIDItems.Clear()

                DingTalkHelper.GetDepartmentIDItems(1)
#End Region

#Region "获取钉钉员工信息"
                Console.WriteLine($"{Now:G}> 获取钉钉员工信息")
                AppSettingHelper.Instance.DingTalkUserJobNumberItems.Clear()

                For Each item In AppSettingHelper.Instance.DingTalkDepartmentIDItems
                    DingTalkHelper.GetUserItems(item)
                Next
#End Region

            End If
#End Region

            ' 无对应的钉钉账号的ERP用户
            Dim NotHaveJobNumberUserItems As New Dictionary(Of String, String)

            ' 发送表单数据
            For Each item In tmpDictionary

                ' 钉钉限制发送频率 1500/min
                Threading.Thread.Sleep(100)

                Console.WriteLine($"{Now:G}> 发送表单数据 {item.Key}")

                ' 判断是否有对应的钉钉账号
                If Not AppSettingHelper.Instance.DingTalkUserJobNumberItems.ContainsKey(item.Key) Then

                    If Not NotHaveJobNumberUserItems.ContainsKey(item.Key) Then
                        NotHaveJobNumberUserItems.Add(item.Key, item.Value.First.CGRY)

                    End If

                    Continue For
                End If

                ' 生成随机文件名
                Dim outputFilePath = Path.Combine(AppSettingHelper.Instance.TempDirectoryPath, $"{item.Key}负责的采购物料-{Wangk.Hash.IDHelper.NewID}.xlsx")

                ' 生成数据文件
                XlsxHelper.SaveResultList(outputFilePath, item.Value)

                ' 上传文件到钉钉
                Dim tmpMediaID = DingTalkHelper.MediaUpload(outputFilePath)

                ' 发送工作通知
                DingTalkHelper.SendWorkMessage(AppSettingHelper.Instance.DingTalkUserJobNumberItems(item.Key), tmpMediaID)

                AppSettingHelper.Instance.Logger.Info($"发送通知消息至 {item.Value.First.CGRY}")

            Next

            ' 通知管理员更新账号信息
            If NotHaveJobNumberUserItems.Count > 0 Then

                Dim tempAdminMessage = $"无对应的钉钉账号的ERP用户  
{String.Join(vbCrLf, From item In NotHaveJobNumberUserItems
                     Select $"> {item.Value}  ")}"

                DingTalkHelper.SendMessageToAdmin(tempAdminMessage)

            End If

            Console.WriteLine($"{Now:G}> 处理完成")

            'Exit Do
        Loop

        'Console.ReadLine()

    End Sub



End Module

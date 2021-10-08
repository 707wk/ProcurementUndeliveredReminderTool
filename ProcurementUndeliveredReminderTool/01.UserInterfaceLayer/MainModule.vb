Imports System.IO
Imports OfficeOpenXml

Module MainModule

    Sub Main()

        Console.Title = $"{My.Application.Info.Title} V{AppSettingHelper.Instance.ProductVersion}"

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

        'Do

        '    ' 定时发送
        '    If Now.Hour <> AppSettingHelper.Instance.SendMsgTime.Hours OrElse
        '    Now.Minute <> AppSettingHelper.Instance.SendMsgTime.Minutes Then
        '        Exit Sub
        '    End If

        '    ' 一天只自动发送一次
        '    If AppSettingHelper.Instance.LastSendDate = Now.Date Then
        '        Exit Sub
        '    End If

        '    AppSettingHelper.Instance.LastSendDate = Now.Date

        '    AppSettingHelper.Instance.Logger.Info("自动发送通知")

        '    ' 检索数据
        '    ' 发送通知文件

        '    Threading.Thread.Sleep(60 * 1000)
        'Loop

        AppSettingHelper.Instance.ClearTempFiles()

        Dim tmpDictionary = RemoteDatabaseHelper.GetDocumentItems()
        For Each item In tmpDictionary

            Dim outputFilePath = Path.Combine(AppSettingHelper.Instance.TempDirectoryPath, $"{item.Key}负责的采购物料-{Wangk.Hash.IDHelper.NewID}.xlsx")

            XlsxHelper.SaveResultList(outputFilePath, item.Value)

            DingTalkHelper.GetAccessToken()
            Dim tmpMediaID = DingTalkHelper.MediaUpload(outputFilePath)

            DingTalkHelper.SendWorkMessage("", tmpMediaID)

        Next

        AppSettingHelper.SaveToLocaltion()
        AppSettingHelper.Instance.ClearTempFiles()

    End Sub



End Module

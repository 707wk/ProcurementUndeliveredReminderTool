Imports System.IO
Imports DingTalk.Api
Imports DingTalk.Api.Request
Imports DingTalk.Api.Response

Public Class DingTalkHelper

#Region "获取钉钉调用企业接口凭证"
    ''' <summary>
    ''' 获取钉钉调用企业接口凭证
    ''' </summary>
    Public Shared Sub GetAccessToken()

        Dim client As IDingTalkClient = New DefaultDingTalkClient("https://oapi.dingtalk.com/gettoken")
        Dim req As OapiGettokenRequest = New OapiGettokenRequest()
        req.Appkey = AppSettingHelper.Instance.DingTalkAppKey
        req.Appsecret = AppSettingHelper.Instance.DingTalkAppSecret
        req.SetHttpMethod("GET")
        Dim rsp As OapiGettokenResponse = client.Execute(req, Nothing)
        AppSettingHelper.Instance.DingTalkAccessToken = rsp.AccessToken

    End Sub
#End Region

#Region "获取钉钉部门信息"
    ''' <summary>
    ''' 获取钉钉部门信息
    ''' </summary>
    Public Shared Sub GetDepartmentIDItems(parentDepartmentID As Long)

        Dim client As IDingTalkClient = New DefaultDingTalkClient("https://oapi.dingtalk.com/topapi/v2/department/listsub")
        Dim req As OapiV2DepartmentListsubRequest = New OapiV2DepartmentListsubRequest()
        req.DeptId = parentDepartmentID
        Dim rsp As OapiV2DepartmentListsubResponse = client.Execute(req, AppSettingHelper.Instance.DingTalkAccessToken)

        If rsp.Result Is Nothing Then
            Exit Sub
        End If

        For Each item In rsp.Result
            AppSettingHelper.Instance.DingTalkDepartmentIDItems.Add(item.DeptId)

            GetDepartmentIDItems(item.DeptId)
        Next

    End Sub
#End Region

#Region "获取钉钉部门用户信息"
    ''' <summary>
    ''' 获取钉钉部门用户信息
    ''' </summary>
    Public Shared Sub GetUserItems(parentDepartmentID As Long)

        Dim client As IDingTalkClient = New DefaultDingTalkClient("https://oapi.dingtalk.com/topapi/v2/user/list")

        Dim Cursor As Long = 0

        Do

            Dim req As OapiV2UserListRequest = New OapiV2UserListRequest()
            req.DeptId = parentDepartmentID
            req.Cursor = Cursor
            req.Size = 100L
            Dim rsp As OapiV2UserListResponse = client.Execute(req, AppSettingHelper.Instance.DingTalkAccessToken)

            If rsp.Result.List Is Nothing Then
                Exit Sub
            End If

            For Each item In rsp.Result.List

                If String.IsNullOrWhiteSpace(item.JobNumber) Then
                    Continue For
                End If

                If Not AppSettingHelper.Instance.DingTalkUserJobNumberItems.ContainsKey(item.JobNumber) Then
                    AppSettingHelper.Instance.DingTalkUserJobNumberItems.Add(item.JobNumber, item.Userid)
                End If

            Next

            Cursor += req.Size
        Loop

    End Sub
#End Region

#Region "向钉钉用户发送工作通知消息"
    ''' <summary>
    ''' 向钉钉用户发送工作通知消息
    ''' </summary>
    Public Shared Sub SendWorkMessage(dingTalkUserid As String,
                                      mediaID As String)

        Dim client = New DefaultDingTalkClient("https://oapi.dingtalk.com/topapi/message/corpconversation/asyncsend_v2")
        Dim req = New OapiMessageCorpconversationAsyncsendV2Request()
        req.AgentId = AppSettingHelper.Instance.DingTalkAgentId
        req.UseridList = "3349644230885065"
        Dim obj1 = New OapiMessageCorpconversationAsyncsendV2Request.MsgDomain()
        obj1.Msgtype = "file"
        Dim obj2 = New OapiMessageCorpconversationAsyncsendV2Request.FileDomain()
        obj2.MediaId = mediaID
        obj1.File = obj2
        req.Msg_ = obj1
        Dim rsp = client.Execute(req, AppSettingHelper.Instance.DingTalkAccessToken)

    End Sub
#End Region

#Region "发送消息给所有主管理员"
    ''' <summary>
    ''' 发送消息给所有主管理员
    ''' </summary>
    Public Shared Sub SendMessageToAdmin(msg As String)

        For Each item In GetAdminItems()
            SendAdminMessage(item, msg)
        Next

    End Sub
#End Region

#Region "获取主管理员列表"
    ''' <summary>
    ''' 获取主管理员列表
    ''' </summary>
    Private Shared Function GetAdminItems() As List(Of String)

        Dim client As New DefaultDingTalkClient("https://oapi.dingtalk.com/topapi/user/listadmin")
        Dim req As New OapiUserListadminRequest()
        Dim rsp As OapiUserListadminResponse = client.Execute(req, AppSettingHelper.Instance.DingTalkAccessToken)

        Return (From item In rsp.Result
                Where item.SysLevel = 1
                Select item.Userid).ToList

    End Function
#End Region

#Region "发送消息给主管理员"
    ''' <summary>
    ''' 发送消息给主管理员
    ''' </summary>
    Public Shared Sub SendAdminMessage(dingTalkUserid As String,
                                       msg As String)

        Dim client As IDingTalkClient = New DefaultDingTalkClient("https://oapi.dingtalk.com/topapi/message/corpconversation/asyncsend_v2")
        Dim req As New OapiMessageCorpconversationAsyncsendV2Request With {
            .AgentId = AppSettingHelper.Instance.DingTalkAgentId,
            .UseridList = dingTalkUserid
        }
        Dim obj1 As New OapiMessageCorpconversationAsyncsendV2Request.MsgDomain With {
            .Msgtype = "markdown"
        }
        Dim obj2 As New OapiMessageCorpconversationAsyncsendV2Request.MarkdownDomain With {
            .Text = msg,
            .Title = "管理员消息"
        }
        obj1.Markdown = obj2
        req.Msg_ = obj1
        Dim rsp As OapiMessageCorpconversationAsyncsendV2Response = client.Execute(req, AppSettingHelper.Instance.DingTalkAccessToken)

    End Sub
#End Region

#Region "上传文件"
    ''' <summary>
    ''' 上传文件
    ''' </summary>
    Public Shared Function MediaUpload(filePath As String) As String

        Dim client = New DefaultDingTalkClient("https://oapi.dingtalk.com/media/upload?type=file")
        Dim req As OapiMediaUploadRequest = New OapiMediaUploadRequest With {
            .Media = New Top.Api.Util.FileItem(filePath)
        }

        Dim rsp As OapiMediaUploadResponse = client.Execute(req, AppSettingHelper.Instance.DingTalkAccessToken)
        Return rsp.MediaId

    End Function
#End Region

End Class

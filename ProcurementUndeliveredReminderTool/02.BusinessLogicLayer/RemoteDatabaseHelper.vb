Imports System.Data.SqlClient
''' <summary>
''' 远程数据库辅助模块
''' </summary>
Public NotInheritable Class RemoteDatabaseHelper

#Region "获取用户工号列表"
    ''' <summary>
    ''' 获取用户工号列表
    ''' </summary>
    Private Shared Function GetUserJobNumberItems() As List(Of String)
        Dim tmpList As New List(Of String)

        Using SqlConn As New SqlConnection(AppSettingHelper.Instance.ERPSqlServerConnStr)
            SqlConn.Open()

            Using tmpSqlCommand = SqlConn.CreateCommand
                tmpSqlCommand.CommandText = $"select
PURTC.TC011,
COUNT(PURTC.TC011)

from
    (select
    *

    -- 采购单单身信息档
    from PURTD
    -- 采购未结束
    where TD016='N'
    -- 已审核
    and TD018='Y'
    -- 提前提醒
    and TD012<'{Now.AddDays(AppSettingHelper.Instance.AdvanceNoticeDays):yyyyMMdd}') as tempPURTD

-- 采购单单头信息档
left join PURTC
on PURTC.TC001=tempPURTD.TD001
and PURTC.TC002=tempPURTD.TD002

group by PURTC.TC011
"

                Using tmpSqlDataReader = tmpSqlCommand.ExecuteReader

                    While tmpSqlDataReader.Read

                        Dim tmpValue = $"{tmpSqlDataReader(0)}".Trim

                        If String.IsNullOrWhiteSpace(tmpValue) Then
                            Continue While
                        End If

                        tmpList.Add(tmpValue)

                    End While

                End Using

            End Using

            SqlConn.Close()
        End Using

        Return tmpList
    End Function
#End Region

#Region "获取要发送的表单信息"
    ''' <summary>
    ''' 获取要发送的表单信息
    ''' </summary>
    Public Shared Function GetDocumentItems() As Dictionary(Of String, List(Of DocumentInfo))
        Dim tmpDictionary As New Dictionary(Of String, List(Of DocumentInfo))

        Dim UserJobNumberItems = GetUserJobNumberItems()

        For Each item In UserJobNumberItems
            tmpDictionary.Add(item, GetDocumentItemsByUserJobNumber(item))
        Next

        Return tmpDictionary
    End Function
#End Region

#Region "按工号获取表单信息"
    ''' <summary>
    ''' 按工号获取表单信息
    ''' </summary>
    Private Shared Function GetDocumentItemsByUserJobNumber(userJobNumber As String) As List(Of DocumentInfo)
        Dim tmpList As New List(Of DocumentInfo)

        Using SqlConn As New SqlConnection(AppSettingHelper.Instance.ERPSqlServerConnStr)
            SqlConn.Open()

            Using tmpSqlCommand = SqlConn.CreateCommand
                tmpSqlCommand.CommandText = $"select
RTRIM(CMSMV.MV002)+'('+RTRIM(CMSMV.MV001)+')' as 采购人员,
RTRIM(PURMA.MA002)+'('+RTRIM(PURMA.MA001)+')' as 供应商简称,
tempPURTD.TD001+'-'+tempPURTD.TD002+'-'+tempPURTD.TD003 as 采购单号,
tempPURTD.TD004 as 品号,
tempPURTD.TD005 as 品名,
tempPURTD.TD006 as 规格,
tempPURTD.TD009 as 单位,
tempPURTD.TD008 as 采购数量,
tempPURTD.TD015 as 已交数量,
tempPURTD.TD008-tempPURTD.TD015 as 未交数量,
tempPURTD.TD012 as 预交货日,
tempPURTD.TD014 as 备注

from
    (select
    *

    -- 采购单单身信息档
    from PURTD
    -- 采购未结束
    where TD016='N'
    -- 已审核
    and TD018='Y'
    -- 提前提醒
    and TD012<'{Now.AddDays(AppSettingHelper.Instance.AdvanceNoticeDays):yyyyMMdd}') as tempPURTD

-- 采购单单头信息档
left join PURTC
on PURTC.TC001=tempPURTD.TD001
and PURTC.TC002=tempPURTD.TD002

-- 员工基本信息档
left join CMSMV
on CMSMV.MV001=PURTC.TC011

-- 供应商基本信息档
left join PURMA
on PURMA.MA001=PURTC.TC004

where PURTC.TC011='{userJobNumber}'

order by tempPURTD.TD012
"

                Using tmpSqlDataReader = tmpSqlCommand.ExecuteReader

                    While tmpSqlDataReader.Read

                        Dim tmpDocumentInfo = New DocumentInfo With {
                            .CGRY = $"{tmpSqlDataReader(0)}",
                            .GYSJC = $"{tmpSqlDataReader(1)}",
                            .CGDH = $"{tmpSqlDataReader(2)}",
                            .PH = $"{tmpSqlDataReader(3)}",
                            .PM = $"{tmpSqlDataReader(4)}",
                            .GG = $"{tmpSqlDataReader(5)}",
                            .Unit = $"{tmpSqlDataReader(6)}",
                            .CGSL = Val($"{tmpSqlDataReader(7)}"),
                            .YJSL = Val($"{tmpSqlDataReader(8)}"),
                            .WJSL = Val($"{tmpSqlDataReader(9)}"),
                            .YJHR = Date.ParseExact($"{tmpSqlDataReader(10)}", "yyyyMMdd", Nothing),
                            .Remark = $"{tmpSqlDataReader(11)}"
                        }

                        tmpList.Add(tmpDocumentInfo)

                    End While

                End Using

            End Using

            SqlConn.Close()
        End Using

        Return tmpList
    End Function
#End Region

End Class

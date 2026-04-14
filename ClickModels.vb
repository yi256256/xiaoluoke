Imports System.Collections.Generic
Imports System.Runtime.Serialization
Imports System.Windows.Forms

<DataContract()>
Public Enum ActionKind
    <EnumMember()>
    MouseClick = 0

    <EnumMember()>
    KeyPress = 1
End Enum

<DataContract()>
Public Class ClickPointConfig
    <DataMember(Order:=1)>
    Public Property ActionType As ActionKind

    <DataMember(Order:=2)>
    Public Property X As Integer

    <DataMember(Order:=3)>
    Public Property Y As Integer

    <DataMember(Order:=4)>
    Public Property IntervalMs As Integer

    <DataMember(Order:=5)>
    Public Property Enabled As Boolean

    <DataMember(Order:=6)>
    Public Property Remark As String

    <DataMember(Order:=7)>
    Public Property KeyCode As Integer

    <DataMember(Order:=8)>
    Public Property HoldMs As Integer

    Public Function Clone() As ClickPointConfig
        Return New ClickPointConfig() With {
            .ActionType = Me.ActionType,
            .X = Me.X,
            .Y = Me.Y,
            .IntervalMs = Me.IntervalMs,
            .Enabled = Me.Enabled,
            .Remark = If(Me.Remark, String.Empty),
            .KeyCode = Me.KeyCode,
            .HoldMs = Me.HoldMs
        }
    End Function

    Public Function GetDisplayKeyName() As String
        If Me.KeyCode <= 0 Then
            Return String.Empty
        End If

        Return CType(Me.KeyCode, Keys).ToString()
    End Function
End Class

<DataContract()>
Public Class AppConfig
    <DataMember(Order:=1)>
    Public Property Points As List(Of ClickPointConfig)

    <DataMember(Order:=2)>
    Public Property WindowX As Integer

    <DataMember(Order:=3)>
    Public Property WindowY As Integer

    <DataMember(Order:=4)>
    Public Property WindowWidth As Integer

    <DataMember(Order:=5)>
    Public Property WindowHeight As Integer

    <DataMember(Order:=6)>
    Public Property StartDelaySeconds As Integer

    <DataMember(Order:=7)>
    Public Property MinimizeOnStart As Boolean

    Public Sub New()
        Me.Points = New List(Of ClickPointConfig)()
        Me.StartDelaySeconds = 3
        Me.MinimizeOnStart = True
    End Sub
End Class

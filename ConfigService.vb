Imports System
Imports System.IO
Imports System.Runtime.Serialization.Json

Public NotInheritable Class ConfigService
    Private Sub New()
    End Sub

    Public Shared ReadOnly Property ConfigFilePath As String
        Get
            Dim baseDirectory As String = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "AutoClicker")
            Return Path.Combine(baseDirectory, "config.json")
        End Get
    End Property

    Public Shared Function LoadConfig(ByRef hadError As Boolean, ByRef errorMessage As String) As AppConfig
        hadError = False
        errorMessage = String.Empty

        If Not File.Exists(ConfigFilePath) Then
            Return New AppConfig()
        End If

        Try
            Dim serializer As New DataContractJsonSerializer(GetType(AppConfig))

            Using stream As New FileStream(ConfigFilePath, FileMode.Open, FileAccess.Read, FileShare.Read)
                Dim loaded As AppConfig = CType(serializer.ReadObject(stream), AppConfig)
                If loaded Is Nothing Then
                    Return New AppConfig()
                End If

                If loaded.Points Is Nothing Then
                    loaded.Points = New List(Of ClickPointConfig)()
                End If

                Return loaded
            End Using
        Catch ex As Exception
            hadError = True
            errorMessage = ex.Message
            Return New AppConfig()
        End Try
    End Function

    Public Shared Function SaveConfig(config As AppConfig, ByRef errorMessage As String) As Boolean
        errorMessage = String.Empty

        Try
            Dim directoryPath As String = Path.GetDirectoryName(ConfigFilePath)
            If String.IsNullOrEmpty(directoryPath) Then
                Throw New InvalidOperationException("配置目录无效。")
            End If

            IO.Directory.CreateDirectory(directoryPath)

            Dim serializer As New DataContractJsonSerializer(GetType(AppConfig))

            Using stream As New FileStream(ConfigFilePath, FileMode.Create, FileAccess.Write, FileShare.None)
                serializer.WriteObject(stream, config)
            End Using

            Return True
        Catch ex As Exception
            errorMessage = ex.Message
            Return False
        End Try
    End Function
End Class

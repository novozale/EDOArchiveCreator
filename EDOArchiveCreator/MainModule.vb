Imports Diadoc.Api
Imports Diadoc.Api.Cryptography
Imports System.Net.Http
Imports System.IO
Imports System.Net.Mail


Module MainModule
    Public MyDateFrom As Date       '--начало промежутка скачивания документов
    Public MyDateTo As Date         '--окончание промежутка скачивания документов

    Sub Main(ByVal cmdArgs() As String)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// основная функция скачивания документов
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyParam As Double

        If EventLog.SourceExists("EDOArchiveCreator") = False Then '--первый раз запустить от имени администратора для создания лога
            EventLog.CreateEventSource("EDOArchiveCreator", "Application")
        End If

        '----считывание дат из командной строки------------------------------------------
        If cmdArgs.Length = 2 Then      '--промежуток времени задан 2 датами
            Try
                MyDateFrom = CDate(cmdArgs(0))
                Try
                    MyDateTo = CDate(cmdArgs(1))
                Catch ex As Exception
                    If My.Settings.MyDebug = "YES" Then
                        EventLog.WriteEntry("EDOArchiveCreator", ex.Message)
                    End If
                    AddDefaultDates(0.1)
                End Try
            Catch ex As Exception
                If My.Settings.MyDebug = "YES" Then
                    EventLog.WriteEntry("EDOArchiveCreator", ex.Message)
                End If
                AddDefaultDates(0.1)
            End Try
        ElseIf cmdArgs.Length = 1 Then  '--промежуток времени задан сдвигом даты начала периода
            '-- < 0         - 10 дней (по умолчанию)
            '-- 0 < 1       - количество дней (умноженное на 100)
            '-- >= 1        - количество месяцев
            Try
                MyParam = cmdArgs(0)
                AddDefaultDates(MyParam)
            Catch ex As Exception
                If My.Settings.MyDebug = "YES" Then
                    EventLog.WriteEntry("EDOArchiveCreator", ex.Message)
                End If
                AddDefaultDates(0.1)
            End Try
        Else                            '--промежуток времени не задан - берется по умолчанию 10 дней
            AddDefaultDates(0.1)
        End If

        '----считывание дат из командной строки------------------------------------------
        If ReadParameters() = True Then
            GetFullArchive()
        Else
            SendMyReminder("Ошибка чтения параметров получения архива ЭДО", "При попытке прочитать параметры утилиты получения архива ЭДО с сайта Diadoc произошли ошибки. Подробности в логах.")
        End If
    End Sub

    Private Sub AddDefaultDates(MyParam As Double)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Задаем даты по умолчанию    <= 0    - берется по умолчанию 10 дней
        '//                             0 < 1   - количество дней (умноженное на 100)
        '//                             >= 1    - количество месяцев
        '////////////////////////////////////////////////////////////////////////////////

        MyDateTo = New System.DateTime(Year(Today()), Month(Today()), Day(Today()))
        If MyParam <= 0 Then
            MyDateFrom = DateAdd(DateInterval.Day, -10, MyDateTo)
        ElseIf MyParam > 0 And MyParam < 1 Then
            MyDateFrom = DateAdd(DateInterval.Day, -MyParam * 100, MyDateTo)
        Else
            MyDateFrom = DateAdd(DateInterval.Month, -MyParam, MyDateTo)
        End If

        '---для отладки
        'MyDateFrom = New System.DateTime(2016, 3, 14)
        MyDateFrom = New System.DateTime(2015, 1, 1)
        'MyDateTo = New System.DateTime(2016, 3, 15)
        'MyDateTo = New System.DateTime(2019, 10, 3)

    End Sub

    Private Function ReadParameters() As Boolean
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// чтение конфигурации из config
        '//
        '////////////////////////////////////////////////////////////////////////////////

        ReadParameters = False

        Try
            Declarations.DiadocApiURL = My.Settings.DiadocApiURL
        Catch ex As Exception
            If My.Settings.MyDebug = "YES" Then
                EventLog.WriteEntry("EDOArchiveCreator", ex.Message)
            End If
            Exit Function
        End Try

        Try
            Declarations.DeveloperKey = My.Settings.DeveloperKey
        Catch ex As Exception
            If My.Settings.MyDebug = "YES" Then
                EventLog.WriteEntry("EDOArchiveCreator", ex.Message)
            End If
            Exit Function
        End Try

        Try
            Declarations.MyLogin = My.Settings.MyLogin
        Catch ex As Exception
            If My.Settings.MyDebug = "YES" Then
                EventLog.WriteEntry("EDOArchiveCreator", ex.Message)
            End If
            Exit Function
        End Try

        Try
            Declarations.MyPassword = My.Settings.MyPassword
        Catch ex As Exception
            If My.Settings.MyDebug = "YES" Then
                EventLog.WriteEntry("EDOArchiveCreator", ex.Message)
            End If
            Exit Function
        End Try

        Try
            Declarations.MyboxId = My.Settings.MyboxId
        Catch ex As Exception
            If My.Settings.MyDebug = "YES" Then
                EventLog.WriteEntry("EDOArchiveCreator", ex.Message)
            End If
            Exit Function
        End Try

        Try
            Declarations.MyArchivePath = My.Settings.MyArchivePath
            If Declarations.MyArchivePath.EndsWith("\") Then
            Else
                Declarations.MyArchivePath = Declarations.MyArchivePath + "\"
            End If
        Catch ex As Exception
            If My.Settings.MyDebug = "YES" Then
                EventLog.WriteEntry("EDOArchiveCreator", ex.Message)
            End If
            Exit Function
        End Try

        Try
            Declarations.MyOverwriteFlag = My.Settings.MyOverwriteFlag
        Catch ex As Exception
            If My.Settings.MyDebug = "YES" Then
                EventLog.WriteEntry("EDOArchiveCreator", ex.Message)
            End If
            Exit Function
        End Try

        Try
            Declarations.MyNumberDownloadTries = My.Settings.NumberDownloadTries
        Catch ex As Exception
            If My.Settings.MyDebug = "YES" Then
                EventLog.WriteEntry("EDOArchiveCreator", ex.Message)
            End If
            Exit Function
        End Try

        ReadParameters = True
    End Function
End Module

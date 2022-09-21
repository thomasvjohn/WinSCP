Public Class Form1

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        SendSFTP()
    End Sub

    Dim ARY As New ArrayList

    Sub SendSFTP()
        ARY.Add(Now.ToString & " > Starting...")

        ' Setup session options
        Dim sessionOptions As New SessionOptions
        With sessionOptions
            .Protocol = Protocol.Sftp
            .HostName = "xxxx"
            .PortNumber = xxxx
            .UserName = "xxxx"
            .Password = "xxxx"
            .SshHostKeyFingerprint = "xxxx"
        End With

        Using session As New Session

            session.SessionLogPath = "xx.log" 'SITE_Config.SiteName & "_SiteSyncLog.log"
            ' Will continuously report progress of synchronization
            AddHandler session.FileTransferred, AddressOf FileTransferred

            ' Connect
            session.Open(sessionOptions)
            ARY.Add(Now.ToString & " > -------------Connected-------------")

            'Home Driectory
            If session.FileExists("xx") = False Then session.CreateDirectory("xx")

            Dim Comp As ComparisonDifferenceCollection
            Comp = session.CompareDirectories(SynchronizationMode.Remote, "xx", "xx", False)
            ARY.Add(Now.ToString & " > -------------Deleting Old Files-------------")
            For i = 0 To Comp.Count - 1
                Try
                    If Comp.Item(i).Action = SynchronizationAction.UploadUpdate Then
                        session.RemoveFile(Comp.Item(i).Remote.FileName)
                        ARY.Add(Comp.Item(i).Remote.FileName)
                    End If
                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                End Try
            Next

            ' Synchronize files
            ARY.Add(Now.ToString & " > -------------Synchronize Directories-------------")
            Dim synchronizationResult As SynchronizationResult
            synchronizationResult = session.SynchronizeDirectories(SynchronizationMode.Remote, "xxx", "xxx", False, True, SynchronizationCriteria.Time)

            ' Throw on any error
            synchronizationResult.Check()

            ARY.Add(Now.ToString & " > -------------Finished-------------")

        End Using

        ARY.Add(Now.ToString & " > Finished...")
    End Sub

    Sub FileTransferred(sender As Object, e As TransferEventArgs)
        If e.Error Is Nothing Then
            ARY.Add(Now.ToString & " > Upload of " & e.FileName & " succeeded")
        Else
            ARY.Add(Now.ToString & " > Upload of " & e.FileName & " failed: " & e.Error.ToString)
        End If
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Me.Text = Now.ToString
        Do Until ARY.Count = 0
            ListBox1.Items.Insert(0, ARY(0))
            ARY.RemoveAt(0)
        Loop
    End Sub

    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        SendSFTP()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        BackgroundWorker1.RunWorkerAsync()
    End Sub

End Class
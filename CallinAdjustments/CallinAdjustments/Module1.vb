
Imports System.Data.SqlClient
Imports System.IO
Imports System.Net.Mail

Module Module1

    Dim DBconnectionString As String = System.Configuration.ConfigurationManager.ConnectionStrings("DefaultConnection").ToString()
    Dim errorfilePath As String = System.Configuration.ConfigurationManager.AppSettings("errorLog").ToString()
    Dim completefilePath As String = System.Configuration.ConfigurationManager.AppSettings("completionLog").ToString()

    Sub Main()
        Try
            Dim Adapter As New SqlDataAdapter
            Dim dt As New DataTable
            Using connect As New SqlConnection(DBconnectionString)
                Dim objDBCommand As New SqlCommand

                With objDBCommand
                    .Connection = connect
                    .Connection.Open()
                    .CommandType = CommandType.StoredProcedure
                    .CommandText = "billing_CallinOrders"

                    Adapter.SelectCommand = objDBCommand
                    Adapter.Fill(dt)

                End With
            End Using

            Using insertconnect As New SqlConnection(DBconnectionString)
                Dim insertDBCommand As New SqlCommand
                With insertDBCommand
                    .Connection = insertconnect
                    .Connection.Open()

                    For Each dr As DataRow In dt.Rows

                        Dim locationId As String = dr("LocationID")
                        Dim itemid As Integer = CInt(dr("ItemID"))
                        Dim quantity As Integer = CInt(dr("Quantity"))
                        Dim userid As String = dr("UserID")

                        .CommandType = CommandType.StoredProcedure
                        .CommandText = "billing_InsertBillingAdjustment"
                        .Parameters.Clear()
                        .Parameters.AddWithValue("@LocationID", dr("LocationID"))
                        .Parameters.AddWithValue("@ItemID", dr("ItemID"))
                        If CInt(dr("Quantity")) >= 0 Then
                            .Parameters.AddWithValue("@ReplacementQuantity", CInt(dr("Quantity")))
                            .Parameters.AddWithValue("@creditQuantity", 0)
                        Else
                            .Parameters.AddWithValue("@ReplacementQuantity", 0)
                            .Parameters.AddWithValue("@creditQuantity", CInt(dr("Quantity")))
                        End If
                        .Parameters.AddWithValue("@UserID", dr("UserID"))
                        .Parameters.AddWithValue("@Approved", 1)
                        .ExecuteNonQuery()


                    Next
                End With
            End Using

            sendemail("LMS - Upload to Billing Adjustment Complete!", "Call in orders have been copied to the Billing Adjustment Table - " & Now)
            logEvent("LMS - Success!", "Record upload to Billing Adjustment Table Completed! - " & Now)

        Catch ex As Exception

            sendemail("LMS - Billing Adjustment Error!", Now & " - " & ex.ToString())
            logEvent("Error", Now & " - " & ex.ToString())

        End Try


    End Sub

    Private Sub sendemail(ByVal subject As String, ByVal mess As String)
        Dim SMTP As New SmtpClient(System.Configuration.ConfigurationManager.AppSettings("SMTPServer"))
        Dim message As New MailMessage(System.Configuration.ConfigurationManager.AppSettings("FromEmailAddress"), System.Configuration.ConfigurationManager.AppSettings("toEmailAddress"))
        Dim otherRecipients As String() = System.Configuration.ConfigurationManager.AppSettings("otherEmailAddresses").ToString().Split(",")
        Dim username As String = System.Configuration.ConfigurationManager.AppSettings("username")
        Dim password As String = System.Configuration.ConfigurationManager.AppSettings("password")
        Dim port As Integer = CInt(System.Configuration.ConfigurationManager.AppSettings("port"))

        SMTP.EnableSsl = True

        SMTP.Credentials = New Net.NetworkCredential(username, password)
        SMTP.Port = port

        message.Subject = subject
        message.Body = mess

        message.To.Clear()
        For Each Recipient As String In otherRecipients
            message.To.Add(New MailAddress(Recipient))
        Next

        Try
            SMTP.Send(message)
            logEvent("Email Success", "Email Sent " & Now)
        Catch ex As Exception
            logEvent("Email Error", Now & " - " & ex.ToString())
        End Try

    End Sub

    Private Sub logEvent(ByVal type As String, ByVal mess As String)
        If type = "Error" Then
            System.IO.Directory.CreateDirectory(errorfilePath)
            Dim errorfilename As String = String.Format("BillingAdjustmentError{0:yyyyMMdd_HH-mm-ss}.txt", Date.Now)
            Dim errorfile = New StreamWriter(errorfilePath & errorfilename, True)
            errorfile.Write(mess)
            errorfile.Close()

        ElseIf type = "Email Error" Then
            System.IO.Directory.CreateDirectory(errorfilePath)
            Dim errorfilename As String = String.Format("EmailError{0:yyyyMMdd_HH-mm-ss}.txt", Date.Now)
            Dim errorfile = New StreamWriter(errorfilePath & errorfilename, True)
            errorfile.Write(mess)
            errorfile.Close()

        ElseIf type = "Email Success" Then
            System.IO.Directory.CreateDirectory(completefilePath)
            Dim completefilename As String = String.Format("EmailSent{0:yyyyMMdd_HH-mm-ss}.txt", Date.Now)
            Dim completefile = New StreamWriter(completefilePath & completefilename, True)
            completefile.Write(mess)
            completefile.Close()
        Else
            System.IO.Directory.CreateDirectory(completefilePath)
            Dim completefilename As String = String.Format("BillingAdjustmentComplete{0:yyyyMMdd_HH-mm-ss}.txt", Date.Now)
            Dim completefile = New StreamWriter(completefilePath & completefilename, True)
            completefile.Write(mess)
            completefile.Close()

        End If
    End Sub

End Module

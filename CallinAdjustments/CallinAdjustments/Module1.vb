
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

                For Each dr As DataRow In dt.Rows

                    With insertDBCommand
                        .Connection = insertconnect
                        .Connection.Open()
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

                    End With
                Next

            End Using

            sendemail("Upload Complete!", "Call in orders have been copied to the Billing Adjustment Table")
            logEvent("Success!", "Record upload to Billing Adjustment Table Completed!")

        Catch ex As Exception

            sendemail("Call in Order Error!", ex.ToString())
            logEvent("Error", ex.ToString())

        End Try


    End Sub

    Private Sub sendemail(ByVal heading As String, ByVal result As String)
        Dim SMTP As New SmtpClient(System.Configuration.ConfigurationManager.AppSettings("SMTPServer"))
        Dim message As New MailMessage(System.Configuration.ConfigurationManager.AppSettings("FromEmailAddress"), System.Configuration.ConfigurationManager.AppSettings("toEmailAddress"))
        Dim otherRecipients As String() = System.Configuration.ConfigurationManager.AppSettings("otherEmailAddresses").ToString().Split(",")


        message.Subject = heading
        message.Body = result

        message.To.Clear()
        For Each Recipient As String In otherRecipients
            message.To.Add(New MailAddress(Recipient))
        Next

        Try
            SMTP.Send(message)
            logEvent("Success!", "Email Sent Complete!")
        Catch ex As Exception
            logEvent("Error", ex.ToString())
        End Try

    End Sub

    Private Sub logEvent(ByVal type As String, ByVal mess As String)
        If type = "Error" Then
            Dim errorfilename As String = String.Format("BillingAdjustmentError{0:yyyyMMdd_HH-mm-ss}.txt", Date.Now)
            Dim errorfile = New StreamWriter(errorfilePath & errorfilename, True)
            errorfile.Write(mess)
            errorfile.Close()

        Else
            Dim completefilename As String = String.Format("BillingAdjustmentComplete{0:yyyyMMdd_HH-mm-ss}.txt", Date.Now)
            Dim completefile = New StreamWriter(completefilePath & completefilename, True)
            completefile.Write(mess)
            completefile.Close()

        End If
    End Sub

End Module

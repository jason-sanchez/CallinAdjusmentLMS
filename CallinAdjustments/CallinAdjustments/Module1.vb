
Imports System.Data.SqlClient
Imports System.IO
Imports System.Net.Mail
Imports System.Globalization

Module Module1

    Dim DBconnectionString As String = System.Configuration.ConfigurationManager.ConnectionStrings("DefaultConnection").ToString()
    Dim errorfilePath As String = System.Configuration.ConfigurationManager.AppSettings("errorLog").ToString()
    Dim completefilePath As String = System.Configuration.ConfigurationManager.AppSettings("completionLog").ToString()

    'Dim DBconnectionString As String = "server=192.168.20.4;database=LMS.Net;uid=sysmax;pwd=sysmax"
    'Dim errorfilePath As String = "C:\LMSCallinAdjustments\ErrorLog\"
    'Dim completefilePath As String = "C:\LMSCallinAdjustments\completionLog\"

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
                    '.CommandText = "billing_CallinOrdersTest"

                    Adapter.SelectCommand = objDBCommand
                    Adapter.Fill(dt)

                End With
            End Using

            Using insertconnect As New SqlConnection(DBconnectionString)
                Dim insertDBCommand As New SqlCommand
                With insertDBCommand
                    .Connection = insertconnect
                    .Connection.Open()

                    Dim sql As String = ""
                    Dim sql2 As String = ""
                    For Each dr As DataRow In dt.Rows

                        'Dim locationId As String = dr("LocationID")
                        'Dim itemid As Integer = CInt(dr("ItemID"))
                        'Dim quantity As Integer = CInt(dr("Quantity"))
                        'Dim userid As String = dr("UserID")
                        'Dim orderID As Integer = CInt(dr("OrderID"))
                        'Dim lastModified As String = CDate(dr("LastModified")).ToString("yyyy-MM-dd HH:mm:ss.fff")
                        'Dim approved As Boolean = CBool(dr("Approved"))

                        'Insert Billing Adjustment
                        Dim billingID As Integer

                        .CommandType = CommandType.StoredProcedure
                        .CommandText = "billing_InsertBillingAdjustment"
                        '.CommandText = "billing_InsertBillingAdjustment"
                        .Parameters.Clear()
                        .Parameters.AddWithValue("@LocationID", dr("LocationID"))
                        .Parameters.AddWithValue("@ItemID", dr("ItemID"))
                        If CInt(dr("Quantity")) >= 0 Then
                            .Parameters.AddWithValue("@ReplacementQuantity", CInt(dr("Quantity")))
                            .Parameters.AddWithValue("@creditQuantity", 0)
                        Else
                            .Parameters.AddWithValue("@ReplacementQuantity", 0)
                            'change negative values to positive
                            .Parameters.AddWithValue("@creditQuantity", Math.Abs(CInt(dr("Quantity"))))
                        End If
                        .Parameters.AddWithValue("@UserID", dr("UserID"))
                        .Parameters.AddWithValue("@Approved", dr("Approved"))
                        .ExecuteNonQuery()

                        'Get Billing AdjustmentID
                        sql = " SELECT TOP 1 ID "
                        sql = sql & " FROM [BillingAdjustments] "
                        'sql = sql & " FROM [BillingAdjustmentsTest] "
                        sql = sql & " WHERE LocationID = '" & dr("LocationID") & "' "
                        sql = sql & " AND itemID = '" & CInt(dr("ItemID")) & "' "
                        If CInt(dr("Quantity")) >= 0 Then
                            sql = sql & " AND ReplacementQuantity = '" & CInt(dr("Quantity")) & "' "
                        Else
                            sql = sql & " AND CreditQuantity = '" & Math.Abs(CInt(dr("Quantity"))) & "' "
                        End If
                        sql = sql & " AND userID = '" & dr("UserID") & "' "
                        sql = sql & " Order By ID desc "
                        .CommandType = CommandType.Text
                        .CommandText = sql
                        billingID = .ExecuteScalar()

                        'update [DeliveryOrderItems].[BillingAdjustmentID] with Billing AdjustmentID
                        sql2 = " UPDATE [dbo].[DeliveryOrderItems] "
                        'sql2 = " UPDATE [dbo].[DeliveryOrderItemsTest] "
                        sql2 = sql2 & " SET [BillingAdjustmentID] = '" & billingID & "' "
                        sql2 = sql2 & " WHERE OrderID = '" & CInt(dr("OrderID")) & "' "
                        sql2 = sql2 & " AND ItemID = '" & CInt(dr("ItemID")) & "' "
                        sql2 = sql2 & " AND OriginalQuantity = '" & CInt(dr("Quantity")) & "' "
                        sql2 = sql2 & " AND LastModified = '" & CDate(dr("LastModified")).ToString("yyyy-MM-dd HH:mm:ss.fff") & "' "

                        .CommandType = CommandType.Text
                        .CommandText = sql2
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

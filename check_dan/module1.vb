
' VER 02

Imports System.Data.Odbc
Imports System.Data.Common.DbException
Imports System.IO

Module Module1

    Sub Main(ByVal feed() As String)
        Dim au_num As String = feed(0)

        'au_num = "'david2','01784P01','173406_LED_RESET','EAGLE','YES','YES,_ONLY_TOP','1','1_60MM_MATERIAL','1','0','3','ASS-NO','0','0','27/01/2021','6','20/01/2021','0','92.59','NO','12','50','15','AUSTRIA','TROTEC_LASER_GMBH','YES','MM','0','12','DE','PN6006E39E03772','50','15','1','1','0','0','0','0','1','0','0','0','0','0','AU-202101/01784','BS6006E661530C0-1','EUR','KIT-NO','PPP6006E3174E1CA','R_3F347DB68987A'"

        'If (InStr(au_num, "REF") < 70) Then au_num = au_num & ",'REF','REF'"  ' THis check can be removed once ref number settled down as all will then have 53 commas

        Try
            update(au_num)
        Catch
        End Try

leave_sub:

    End Sub

    Sub update(ByVal input As String)

        Dim stamp As Date
        Dim timeStamp As Long

        stamp = Now
        timeStamp = stamp.Ticks

        input = input & "," & timeStamp

        Dim sqlConnection As New OdbcConnection
        sqlConnection = New OdbcConnection("DSN=OMS_Dev;MultipleActiveResultSets=True")
        Dim sqlCommand As New OdbcCommand
        sqlCommand.CommandText = "INSERT INTO orders_info (order_num,au_num,file_name,file_format,mask,silk,pool,thickness,surface,rebate,free_stencil,order_ass,reorder,rfid,del_date,work_days,order_date,full_rfid,order_price,over_del,order_qty,order_x,order_y,country,company,etest,units,jackaltac,ass_wd,factura,nutzen_id,panel_s_l,panel_s_w,panel_s_q_x,panel_s_q_y,panel_edge,scoring,land_milling,inter_milling,indivi_milling,track_gap,drills,side_plate,stencils,large_order,full_au,identnum,currency,order_kit,waren,customer_id,base_ref_number,ref_number,time_stamp) VALUES (" & input & ")"

        sqlCommand.Connection = sqlConnection

        Try
            If sqlConnection.State = ConnectionState.Closed Then
                sqlConnection.Open()
            End If
        Catch e As System.Exception
            Try
                Dim fail As StreamWriter
                fail = File.AppendText("D:\storage\IN_HOUSE_SOFTWARE\ALL_SOFTWARE\Visual_Studio_2010\OMS-Download\TEMP\SQL.txt")
                'fail = File.AppendText("T:\IN_HOUSE_SOFTWARE\ALL_SOFTWARE\Visual_Studio_2010\OMS-Download\TEMP\SQL.txt")
                fail.WriteLine("DatabaseSQL_UpdateCannot Connect")
                fail.WriteLine(sqlCommand.CommandText)
                fail.WriteLine(e.Message)
                fail.Close()
            Catch
            End Try
        End Try

            If sqlConnection.State = ConnectionState.Open Then
            Try
                sqlCommand.ExecuteNonQuery()
            Catch
                sqlConnection.Close()
                Dim fail2 As StreamWriter
                fail2 = File.AppendText("D:\storage\IN_HOUSE_SOFTWARE\ALL_SOFTWARE\Visual_Studio_2010\OMS-Download\TEMP\SQL.txt")
                'fail2 = File.AppendText("T:\IN_HOUSE_SOFTWARE\ALL_SOFTWARE\Visual_Studio_2010\OMS-Download\TEMP\SQL.txt")
                fail2.WriteLine("DatabaseSQL_Update Cannot Execute")
                fail2.WriteLine(sqlCommand.CommandText)
                fail2.Close()
            End Try
            sqlConnection.Close()
        End If

    End Sub

End Module




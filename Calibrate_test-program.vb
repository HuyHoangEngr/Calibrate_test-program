Dim sample_file_dir As String
Dim data_file_dir As String
Dim temp_sample_file As Workbook
Dim temp_data_file As Workbook
'Public Sub UserForm_Initialize()
'    Dim ate(15) As String
'    ate(0) = "ATE 2.1"
'    ate(1) = "ATE 2.2"
'    ate(2) = "ATE 3.1"
'    ate(3) = "ATE 3.2"
'    ate(4) = "ATE 4.1"
'    ate(5) = "ATE 4.2"
'    ate(6) = "ATE 4.3"
'    ate(7) = "ATE 5.1"
'    ate(8) = "ATE 5.2"
'    ate(9) = "ATE 5.3"
'    ate(10) = "ATE 6.1"
'    ate(11) = "ATE 6.2"
'    ate(12) = "ATE 6.3"
'    ate(13) = "ATE 7.1"
'    ate(14) = "ATE 7.2"
'    ate(15) = "ATE 17"
'    For i = 0 To 15
'        UserForm1.Cmbox_station.AddItem ate(i)
'    Next i
'
'    UserForm1.Cmbox_station.ListIndex = 0
'
'End Sub

Sub btnex_Click()

    Dim prepare_file As Workbook
    
    Dim start_sample As Integer
    Dim end_sample As Integer
    Dim start_data As Integer
    Dim end_data As Integer
    Dim start_min As Integer
    Dim end_max As Integer
    
    Dim baosai As Integer
    
    ''''''''''''''''''''''''''''''''''''''''''''''
    ''''''''''''XOA NOI DUNG TRANG''''''''''''''''
    ''''''''''''''''''''''''''''''''''''''''''''''
    ThisWorkbook.Sheets(1).Range("A:G").Value = ""
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''''''''''''KIEM TRA 2 DUONG DAN TREN 2 TEXT BOX''''''''''''''''
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Do While True
        If sample_file_dir = "" Then
            opensample
        End If
        If sample_file_dir <> "" Then Exit Do
    Loop
    
    Do While True
        If data_file_dir = "" Then
            opendatafile
        End If
        If data_file_dir <> "" Then Exit Do
    Loop

    Set temp_sample_file = Workbooks.Open(sample_file_dir) 'PHAI DUOC KHAI BAO SAU viec sample_file_dir co du lieu
    Set temp_data_file = Workbooks.Open(data_file_dir)
    Set prepare_file = ThisWorkbook

    prepare_file.Sheets(1).Range("C" & "2").Value = "SAMPLE"
    prepare_file.Sheets(1).Range("C" & "2").Font.Size = 20
    prepare_file.Sheets(1).Range("C" & "2").Font.Bold = True
    temp_sample_file.Sheets(1).Range("A" & "1" & ":" & "D" & "600").Copy (prepare_file.Sheets(1).Range("A3"))
    temp_sample_file.Sheets(1).Range("F" & "1" & ":" & "F" & "600").Copy (prepare_file.Sheets(1).Range("E3"))

    start_sample = Findstep(temp_sample_file)
    end_sample = FindUUTpassed(temp_sample_file)
    temp_sample_file.Close savechanges:=False

    prepare_file.Sheets(1).Range("I" & "2").Value = "DATA TEST"
    prepare_file.Sheets(1).Range("I" & "2").Font.Size = 20
    prepare_file.Sheets(1).Range("I" & "2").Font.Bold = True
    temp_data_file.Sheets(1).Range("A" & "1" & ":" & "D" & "600").Copy (prepare_file.Sheets(1).Range("G3"))
    temp_data_file.Sheets(1).Range("F" & "1" & ":" & "F" & "600").Copy (prepare_file.Sheets(1).Range("K3"))

    start_data = Findstep(temp_data_file)
    end_data = FindUUTpassed(temp_data_file)
    temp_data_file.Close savechanges:=False

    start_min = start_sample + 2 'VI COPY PASTE SANG prepare LA ROW=3 NEN + 2
    If start_data < start_sample Then start_min = start_data + 2

    end_max = end_sample
    If end_data > end_sample Then end_max = end_data
    
    prepare_file.Sheets(1).Range("C:C").EntireColumn.ColumnWidth = 40
    prepare_file.Sheets(1).Range("I:I").EntireColumn.ColumnWidth = 40
    prepare_file.Sheets(1).Range("D:E").EntireColumn.ColumnWidth = 15
    prepare_file.Sheets(1).Range("J:K").EntireColumn.ColumnWidth = 15

    baosai = 0
    ''Kiem tra ten chuong trinh cai dat....................................................
    If prepare_file.Sheets(1).Range("A3").Value = prepare_file.Sheets(1).Range("G3").Value Then
        prepare_file.Sheets(1).Range("A3:C3").Interior.ColorIndex = 4
        prepare_file.Sheets(1).Range("G3:I3").Interior.ColorIndex = 4
    Else
        prepare_file.Sheets(1).Range("A3:C3").Interior.ColorIndex = 3
        prepare_file.Sheets(1).Range("G3:I3").Interior.ColorIndex = 3
        baosai = baosai + 1
    End If
    If prepare_file.Sheets(1).Range("A7").Value = prepare_file.Sheets(1).Range("G7").Value Then
        prepare_file.Sheets(1).Range("A7:C7").Interior.ColorIndex = 4
        prepare_file.Sheets(1).Range("G7:I7").Interior.ColorIndex = 4
    Else
        prepare_file.Sheets(1).Range("A7:C7").Interior.ColorIndex = 3
        prepare_file.Sheets(1).Range("G7:I7").Interior.ColorIndex = 3
        baosai = baosai + 1
    End If
    ''......................................................................................
    For i = (start_min + 1) To end_max
            If prepare_file.Sheets(1).Range("C" & i).Value = prepare_file.Sheets(1).Range("I" & i).Value Then
                prepare_file.Sheets(1).Range("C" & i).Interior.ColorIndex = 4
                prepare_file.Sheets(1).Range("I" & i).Interior.ColorIndex = 4
            Else
                prepare_file.Sheets(1).Range("C" & i).Interior.ColorIndex = 3
                prepare_file.Sheets(1).Range("I" & i).Interior.ColorIndex = 3
                baosai = baosai + 1
            End If

            If prepare_file.Sheets(1).Range("D" & i).Value = prepare_file.Sheets(1).Range("J" & i).Value Then
                prepare_file.Sheets(1).Range("D" & i).Interior.ColorIndex = 4
                prepare_file.Sheets(1).Range("J" & i).Interior.ColorIndex = 4
            Else
                prepare_file.Sheets(1).Range("D" & i).Interior.ColorIndex = 3
                prepare_file.Sheets(1).Range("J" & i).Interior.ColorIndex = 3
'                baosai = i
            End If
            
            If prepare_file.Sheets(1).Range("E" & i).Value = prepare_file.Sheets(1).Range("K" & i).Value Then
                prepare_file.Sheets(1).Range("E" & i).Interior.ColorIndex = 4
                prepare_file.Sheets(1).Range("K" & i).Interior.ColorIndex = 4
            Else
                prepare_file.Sheets(1).Range("E" & i).Interior.ColorIndex = 3
                prepare_file.Sheets(1).Range("K" & i).Interior.ColorIndex = 3
'                baosai = i
            End If
    Next i

    prepare_file.Sheets(1).Range("B:B").Columns.AutoFit
    prepare_file.Sheets(1).Range("H:H").Columns.AutoFit
    
    prepare_file.Sheets(1).Range("C" & i).Select
    If baosai <> 0 Then
        MsgBox "Report khong giong voi chuong trinh mau !!! Lien he Testing Engineer review lai chuong trinh !"
    Else
        MsgBox "Report giong voi chuong trinh mau."
    End If
    
End Sub
Function Findstep(file As Workbook) As Integer
    Dim step As Integer
    
    For i = 1 To 10
        If file.Sheets(1).Cells(i, 1).Value = "Step" Then
            step = i
            Exit For
        End If
    Next i

    Findstep = step
End Function
Function FindUUTpassed(file As Workbook) As Integer
    Dim UUTpassed As Integer
    
    For i = 1 To 600
        If file.Sheets(1).Cells(i, 1).Value = "UUT PASSED" Then
            UUTpassed = i
            Exit For
        End If
    Next i
    FindUUTpassed = UUTpassed
End Function
Function opensample()
    Set fd1 = Application.FileDialog(msoFileDialogFilePicker)
    
    With fd1
    
        .Filters.Clear
        .Filters.Add "htm files (*.htm*)", "*.htm*", 1
        .Title = "Choose an PROGRAM CORRECTED .htm file to open"
        .AllowMultiSelect = False
        
        'LINK MAC DINH CUA SERVER CHUA PROGRAM CORRECTED
        .InitialFileName = "\\VNFILE\Dept_Share\11_PS_ME\23. Test program corrected\ATE test report"
    
        If .Show = True Then
    
            sample_file_dir = .SelectedItems(1)
    
        End If
    
    End With
    textboxsample.Text = sample_file_dir
End Function

Sub textboxsample_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    opensample
End Sub
Function opendatafile()
    Set fd2 = Application.FileDialog(msoFileDialogFilePicker)
    Dim ip_report As String
    
'    If UserForm1.Cmbox_station.Value = "ATE 2.1" Then ip_report = "\\192.168.73.82\Reports"
'    If UserForm1.Cmbox_station.Value = "ATE 2.2" Then ip_report = "\\192.168.73.49\Reports"
'    If UserForm1.Cmbox_station.Value = "ATE 3.1" Then ip_report = "\\172.24.121.116\Reports"
'    If UserForm1.Cmbox_station.Value = "ATE 3.2" Then ip_report = "\\192.168.73.28\Reports"
'    If UserForm1.Cmbox_station.Value = "ATE 4.1" Then ip_report = "\\172.24.120.21\Reports"
'    If UserForm1.Cmbox_station.Value = "ATE 4.2" Then ip_report = "\\172.24.120.23\Reports"
'    If UserForm1.Cmbox_station.Value = "ATE 4.3" Then ip_report = "\\192.168.73.26\Reports"
'    If UserForm1.Cmbox_station.Value = "ATE 5.1" Then ip_report = "\\172.24.120.20\Reports"
'    If UserForm1.Cmbox_station.Value = "ATE 5.2" Then ip_report = "\\192.168.73.25\Reports"
'    If UserForm1.Cmbox_station.Value = "ATE 5.3" Then ip_report = "\\192.168.73.44\nhr\Reports"
'    If UserForm1.Cmbox_station.Value = "ATE 6.1" Then ip_report = "\\192.168.73.94\Reports"
'    If UserForm1.Cmbox_station.Value = "ATE 6.2" Then ip_report = "\\172.24.120.119\Reports"
'    If UserForm1.Cmbox_station.Value = "ATE 6.3" Then ip_report = "\\192.168.73.53\Reports"
'    If UserForm1.Cmbox_station.Value = "ATE 7.1" Then ip_report = "\\172.24.121.125\Reports"
'    If UserForm1.Cmbox_station.Value = "ATE 7.2" Then ip_report = "\\192.168.73.59\Reports"
'    If UserForm1.Cmbox_station.Value = "ATE 17" Then ip_report = "\\192.168.73.60\Reports"
    ip_report = txtIP.Text & "\"
    With fd2
    
        .Filters.Clear
        .Filters.Add "htm files (*.htm*)", "*.htm*", 1
        .Title = "Choose an DATA TEST .htm file to open"
        .AllowMultiSelect = False
    
        'THU MUC REPORT CUA STATION DUOC CHON

        .InitialFileName = ip_report
    
        If .Show = True Then
    
            data_file_dir = .SelectedItems(1)
    
        End If
    
    End With
    textboxdata.Text = data_file_dir
End Function
Sub textboxdata_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    opendatafile
End Sub

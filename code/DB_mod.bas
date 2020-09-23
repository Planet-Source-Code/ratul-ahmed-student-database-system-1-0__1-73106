Attribute VB_Name = "DB_mod"
Option Explicit

Dim dbcon                       As New ADODB.Connection
Dim dbrecord                    As New ADODB.Recordset
Dim Constr                      As String
Dim ssQL                        As String
Dim sUpdateKey                  As Integer

Public Function connectDB()    ' Connect to the database

    On Error Resume Next
    Constr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\db\stdb.mdb;"
    dbcon.Open Constr
    
End Function

Public Function DisconnectDB()                  ' Disconnect From The Database

    On Error Resume Next
    dbrecord.Close
    dbcon.Close

End Function

Public Function LoadSTD_data(stdlst As ListView)    ' Load Data to Student from Database

    Dim dbid As Integer
    Dim chkval As String
    Dim lst As Variant
    Dim semext As String
    Dim depart As String
    
    On Error Resume Next
    
    frmstud.txtname.BackColor = &HFFFFFF
    frmstud.txtdep.BackColor = &HFFFFFF
    frmstud.txtsem.BackColor = &HFFFFFF
    frmstud.txtroll.BackColor = &HFFFFFF
    frmstud.txtpresents.BackColor = &HFFFFFF
    frmstud.txtlastpayment.BackColor = &HFFFFFF
    frmstud.txtlastcirtificate.BackColor = &HFFFFFF
    frmstud.txtfname.BackColor = &HFFFFFF
    frmstud.txtmname.BackColor = &HFFFFFF
    frmstud.txtbirth.BackColor = &HFFFFFF
    frmstud.txtpresentadd.BackColor = &HFFFFFF
    frmstud.txtperadd.BackColor = &HFFFFFF
    frmstud.cmd_upload.Enabled = True
    
    stdlst.ListItems.Clear  ' Clear Last Data
    
    ssQL = "Select * from std_data"
    
    dbrecord.Open ssQL, dbcon, adOpenKeyset, adLockReadOnly
    
    chkval = dbrecord("ID").Value
    
    If chkval = "" Then
        MsgBox "The dataBase is Empty. Nothing To Retrive!", vbExclamation, "Empty Database"
        DisconnectDB
    Else
                                                    'Adding All Data From The Database
            '-------------------------------------------------------------------------
            While Not dbrecord.EOF
                Set lst = stdlst.ListItems.Add(, , dbrecord("ID").Value)
                lst.SubItems(1) = dbrecord("Name").Value
                lst.SubItems(2) = dbrecord("Roll").Value
                
                    If (dbrecord("Semester").Value = 1) Then semext = "st"
                    If (dbrecord("Semester").Value = 2) Then semext = "nd"
                    If (dbrecord("Semester").Value = 3) Then semext = "rd"
                    If (dbrecord("Semester").Value = 4) Then semext = "th"
                    If (dbrecord("Semester").Value = 5) Then semext = "th"
                    If (dbrecord("Semester").Value = 6) Then semext = "th"
                    If (dbrecord("Semester").Value = 7) Then semext = "th"
                    If (dbrecord("Semester").Value = 8) Then semext = "th"
                    
                depart = dbrecord("Department").Value & " " & dbrecord("Semester").Value & semext 'Department Genaretion
                lst.SubItems(3) = depart
                lst.SubItems(4) = dbrecord("Grade").Value
                dbrecord.MoveNext
            Wend
            '-------------------------------------------------------------------------
            DisconnectDB
    End If
    

End Function

Public Function change_Objects(lstdb As ListView) ' Change Object Properties


    Dim selID As String
    Dim selRoll As String
    Dim selName As String
    Dim sSex As String

    
    connectDB   ' Connect to teh database
    
        selID = lstdb.SelectedItem.Text
        selRoll = lstdb.SelectedItem.SubItems(2)
        selName = lstdb.SelectedItem.SubItems(1)

        ssQL = "Select * from std_data where id=" & Chr(34) & selID & Chr(34) & " AND Name=" & _
        Chr(34) & selName & Chr(34) & " AND Roll=" & Chr(34) & selRoll & Chr(34) ' SQL COmmand
        
        dbrecord.Open "Select * from std_data where id=" & selID, dbcon, adOpenKeyset, adLockReadOnly 'Query Exicution
        
            sSex = dbrecord("sex").Value & vbNullString
            If sSex = "Male" Then
                frmstud.opt_male.Value = True 'Male
                frmstud.opt_female.Value = False
                frmstud.stdpic.Picture = LoadPicture("")
                frmstud.stdpic.Picture = LoadPicture(App.Path & "\pics\no_male.jpg")
            End If
            If sSex = "Female" Then
                frmstud.opt_female.Value = True 'Female
                frmstud.opt_male.Value = False
                frmstud.stdpic.Picture = LoadPicture("")
                frmstud.stdpic.Picture = LoadPicture(App.Path & "\pics\no_female.jpg")
            End If
   
            frmstud.txtid.Text = dbrecord("ID").Value & vbNullString    'ID
            frmstud.txtroll.Text = dbrecord("Roll").Value & vbNullString    'Roll
            frmstud.txtname.Text = dbrecord("Name").Value & vbNullString    'Name
            frmstud.txtfname.Text = dbrecord("Fathers_name").Value & vbNullString   'Fathers Name
            frmstud.txtmname.Text = dbrecord("Mothers_name").Value & vbNullString   'Mothers name
            frmstud.txtbirth.Text = dbrecord("Dateofbirth").Value & vbNullString    ' Date Of Birth
            frmstud.txtpresentadd.Text = dbrecord("Present_add").Value & vbNullString   'Present Address
            frmstud.txtperadd.Text = dbrecord("Parmanent_add").Value & vbNullString 'Permanent Address
            frmstud.txtdep.Text = dbrecord("Department").Value & vbNullString   'Department
            frmstud.txtsem.Text = dbrecord("Semester").Value & vbNullString 'Semester
            frmstud.txtgrade.Text = dbrecord("Grade").Value & vbNullString  'Grade
            frmstud.txtpresents.Text = dbrecord("Presence").Value & vbNullString    'Presence
            frmstud.txtlastpayment.Text = dbrecord("Last_Paymet").Value & vbNullString  'Last Paymet
            frmstud.txtlastcirtificate.Text = dbrecord("Last_Cirtificate").Value & vbNullString 'Last Cirtificate
            frmstud.picname.Caption = frmstud.txtid.Text & frmstud.txtroll.Text
        
    DisconnectDB ' Disconnect From Database

End Function

Public Function Add_data()

Dim db_Name As String
Dim db_Department As String
Dim db_Semester As String
Dim db_Roll As String
Dim db_Presence As String
Dim db_Last_Paymet As String
Dim db_Last_Cirtificate As String
Dim db_Fathers_name As String
Dim db_Mothers_name As String
Dim db_Dateofbirth As String
Dim db_Last_Present_add As String
Dim db_Parmanent_add As String
Dim db_sex As String
Dim db_img_file As String

db_Name = frmstud.txtname.Text
db_Department = frmstud.txtdep.Text
db_Semester = frmstud.txtsem.Text
db_Roll = frmstud.txtroll.Text
db_Presence = frmstud.txtpresents.Text
db_Last_Paymet = frmstud.txtlastpayment.Text
db_Last_Cirtificate = frmstud.txtlastcirtificate.Text
db_Fathers_name = frmstud.txtfname.Text
db_Mothers_name = frmstud.txtmname.Text
db_Dateofbirth = frmstud.txtbirth.Text
db_Last_Present_add = frmstud.txtpresentadd.Text
db_Parmanent_add = frmstud.txtperadd.Text
db_img_file = frmstud.picname.Caption
If frmstud.opt_male.Value = True Then db_sex = "Male"
If frmstud.opt_female.Value = True Then db_sex = "Female"

frmstud.txtname.BackColor = &HC0C0FF
frmstud.txtdep.BackColor = &HC0C0FF
frmstud.txtsem.BackColor = &HC0C0FF
frmstud.txtroll.BackColor = &HC0C0FF
frmstud.txtpresents.BackColor = &HC0C0FF
frmstud.txtlastpayment.BackColor = &HC0C0FF
frmstud.txtlastcirtificate.BackColor = &HC0C0FF
frmstud.txtfname.BackColor = &HC0C0FF
frmstud.txtmname.BackColor = &HC0C0FF
frmstud.txtbirth.BackColor = &HC0C0FF
frmstud.txtpresentadd.BackColor = &HC0C0FF
frmstud.txtperadd.BackColor = &HC0C0FF

frmstud.txtname.Text = ""
frmstud.txtdep.Text = ""
frmstud.txtsem.Text = ""
frmstud.txtroll.Text = ""
frmstud.txtpresents.Text = ""
frmstud.txtlastpayment.Text = ""
frmstud.txtlastcirtificate.Text = ""
frmstud.txtfname.Text = ""
frmstud.txtmname.Text = ""
frmstud.txtbirth.Text = ""
frmstud.txtpresentadd.Text = ""
frmstud.txtperadd.Text = ""
frmstud.txtgrade.Text = ""
frmstud.txtid.Text = ""
frmstud.cmd_upload.Enabled = False


If frmstud.cmd_add.Caption = "Save" Then
    connectDB
    On Error GoTo Hell
    ssQL = "INSERT INTO `std_data` (`Name`, `Department`, `Semester`, `Roll`, `Presence`, `Last_Paymet`, `Last_Cirtificate`, `Fathers_name`, `Mothers_name`, `Dateofbirth`, `Present_add`, `Parmanent_add`, `sex`, `img_file`) VALUES ('" & db_Name & " ', '" & db_Department & "', '" & db_Semester & "', '" & db_Roll & "', '" & db_Presence & "', '" & db_Last_Paymet & "', '" & db_Last_Cirtificate & "', '" & db_Fathers_name & "', '" & db_Mothers_name & "', '" & db_Dateofbirth & "', '" & db_Last_Present_add & "', '" & db_Parmanent_add & "', '" & db_sex & "', '" & db_img_file & "')"
    dbcon.Execute (ssQL)
    Call LoadSTD_data(frmstud.lststd)
    MsgBox "Record updated sucessfully.", vbInformation, "Done"
    DisconnectDB
    frmstud.cmd_add.Caption = "Add"
    
Else
    frmstud.cmd_add.Caption = "Save"
End If

Hell:
    MsgBox "Please Enter all Data!", vbExclamation, "Caution"
End Function

Public Function Edit_data()

Dim db_Name As String
Dim db_Department As String
Dim db_Semester As String
Dim db_Roll As String
Dim db_Presence As String
Dim db_Last_Paymet As String
Dim db_Last_Cirtificate As String
Dim db_Fathers_name As String
Dim db_Mothers_name As String
Dim db_Dateofbirth As String
Dim db_Last_Present_add As String
Dim db_Parmanent_add As String
Dim db_sex As String
Dim db_img_file As String

db_Name = frmstud.txtname.Text
db_Department = frmstud.txtdep.Text
db_Semester = frmstud.txtsem.Text
db_Roll = frmstud.txtroll.Text
db_Presence = frmstud.txtpresents.Text
db_Last_Paymet = frmstud.txtlastpayment.Text
db_Last_Cirtificate = frmstud.txtlastcirtificate.Text
db_Fathers_name = frmstud.txtfname.Text
db_Mothers_name = frmstud.txtmname.Text
db_Dateofbirth = frmstud.txtbirth.Text
db_Last_Present_add = frmstud.txtpresentadd.Text
db_Parmanent_add = frmstud.txtperadd.Text
db_img_file = frmstud.picname.Caption
If frmstud.opt_male.Value = True Then db_sex = "Male"
If frmstud.opt_female.Value = True Then db_sex = "Female"

frmstud.txtname.BackColor = &H80FF80
frmstud.txtdep.BackColor = &H80FF80
frmstud.txtsem.BackColor = &H80FF80
frmstud.txtroll.BackColor = &H80FF80
frmstud.txtpresents.BackColor = &H80FF80
frmstud.txtlastpayment.BackColor = &H80FF80
frmstud.txtlastcirtificate.BackColor = &H80FF80
frmstud.txtfname.BackColor = &H80FF80
frmstud.txtmname.BackColor = &H80FF80
frmstud.txtbirth.BackColor = &H80FF80
frmstud.txtpresentadd.BackColor = &H80FF80
frmstud.txtperadd.BackColor = &H80FF80

If frmstud.cmd_edit.Caption = "Save" Then
    connectDB
    On Error GoTo Hell
    ssQL = "Update std_data set Name='" & db_Name & "', Department='" & db_Department & "', Semester='" & db_Semester & "', Roll='" & db_Roll & "', Presence='" & db_Presence & "', Last_Paymet='" & db_Last_Paymet & "', Last_Cirtificate='" & db_Last_Cirtificate & "', Fathers_name='" & db_Fathers_name & "', Mothers_name='" & db_Mothers_name & "', Present_add='" & db_Last_Present_add & "',Dateofbirth='" & db_Dateofbirth & "', Parmanent_add='" & db_Parmanent_add & "',  sex='" & db_sex & "' where ID=" & val(frmstud.txtid.Text)
    dbcon.Execute (ssQL)
    Call LoadSTD_data(frmstud.lststd)
    MsgBox "Record updated sucessfully.", vbInformation, "Done"
    DisconnectDB
    frmstud.cmd_edit.Caption = "Edit"
    
Else
    frmstud.cmd_edit.Caption = "Save"
End If
Hell:
    MsgBox "Please enter all Information Carefully, Suspicious Characters Are Blocked!", vbExclamation, "Caution"
End Function

Public Function Delete_data()
Dim yn As String
yn = MsgBox("Are you sure want to Delete ID = " & val(frmstud.txtid.Text) & " ?", vbYesNo, "Coution")
If yn = "6" Then
    connectDB
    ssQL = "DELETE * FROM std_data where id=" & val(frmstud.txtid.Text)
    dbcon.Execute (ssQL)
    Call LoadSTD_data(frmstud.lststd)
    MsgBox "Record Deleted sucessfully.", vbInformation, "Done"
    DisconnectDB
Else
    'None
End If
End Function


Public Function LoadSTD_data_result(stdlst As ListView)    ' Load Data to Student from Database

    Dim dbid As Integer
    Dim chkval As String
    Dim lst As Variant
    Dim semext As String
    Dim depart As String
    
    On Error Resume Next
    
    
    frmsres.txtname.BackColor = &HFFFFFF
    frmsres.txtdep.BackColor = &HFFFFFF
    frmsres.txtsem.BackColor = &HFFFFFF
    frmsres.txtgrade.BackColor = &HFFFFFF
    frmsres.txtcomarch.BackColor = &HFFFFFF
    frmsres.txtmicro.BackColor = &HFFFFFF
    frmsres.txtdatabaseman.BackColor = &HFFFFFF
    frmsres.txtvisual.BackColor = &HFFFFFF
    frmsres.txtdatacomfund.BackColor = &HFFFFFF
    frmsres.txtenviron.BackColor = &HFFFFFF
    frmsres.txtbookkeep.BackColor = &HFFFFFF
    frmsres.txtbizorg.BackColor = &HFFFFFF
    
    frmsres.cmd_upload.Enabled = False
    
    stdlst.ListItems.Clear  ' Clear Last Data
    
    ssQL = "Select * from std_data"
    
    dbrecord.Open ssQL, dbcon, adOpenKeyset, adLockReadOnly
    
    chkval = dbrecord("ID").Value
    
    If chkval = "" Then
        MsgBox "The dataBase is Empty. Nothing To Retrive!", vbExclamation, "Empty Database"
        DisconnectDB
    Else
                                                    'Adding All Data From The Database
            '-------------------------------------------------------------------------
            While Not dbrecord.EOF
                Set lst = stdlst.ListItems.Add(, , dbrecord("ID").Value)
                lst.SubItems(1) = dbrecord("Name").Value
                lst.SubItems(2) = dbrecord("Roll").Value
                
                    If (dbrecord("Semester").Value = 1) Then semext = "st"
                    If (dbrecord("Semester").Value = 2) Then semext = "nd"
                    If (dbrecord("Semester").Value = 3) Then semext = "rd"
                    If (dbrecord("Semester").Value = 4) Then semext = "th"
                    If (dbrecord("Semester").Value = 5) Then semext = "th"
                    If (dbrecord("Semester").Value = 6) Then semext = "th"
                    If (dbrecord("Semester").Value = 7) Then semext = "th"
                    If (dbrecord("Semester").Value = 8) Then semext = "th"
                    
                depart = dbrecord("Department").Value & " " & dbrecord("Semester").Value & semext 'Department Genaretion
                lst.SubItems(3) = depart
                
                lst.SubItems(4) = dbrecord("sub_computer_arc").Value
                lst.SubItems(5) = dbrecord("sub_microprocessor").Value
                lst.SubItems(6) = dbrecord("sub_database_manage").Value
                lst.SubItems(7) = dbrecord("sub_visual_programming").Value
                lst.SubItems(8) = dbrecord("sub_data_communication").Value
                lst.SubItems(9) = dbrecord("sub_environmanal").Value
                lst.SubItems(10) = dbrecord("sub_book_keeping").Value
                lst.SubItems(11) = dbrecord("sub_business_org").Value
                lst.SubItems(12) = dbrecord("sub_score").Value
                lst.SubItems(13) = dbrecord("Grade").Value
                lst.SubItems(14) = dbrecord("cGPA").Value
                
                
                dbrecord.MoveNext
            Wend
            '-------------------------------------------------------------------------
            DisconnectDB
    End If
    

End Function

Public Function change_Objects_result(lstdb As ListView) ' Change Object Properties


    Dim selID As String
    Dim selRoll As String
    Dim selName As String
    Dim sSex As String

    
    connectDB   ' Connect to teh database
    
        selID = lstdb.SelectedItem.Text
        selRoll = lstdb.SelectedItem.SubItems(2)
        selName = lstdb.SelectedItem.SubItems(1)

        ssQL = "Select * from std_data where id=" & Chr(34) & selID & Chr(34) & " AND Name=" & _
        Chr(34) & selName & Chr(34) & " AND Roll=" & Chr(34) & selRoll & Chr(34) ' SQL COmmand
        
        dbrecord.Open "Select * from std_data where id=" & selID, dbcon, adOpenKeyset, adLockReadOnly 'Query Exicution
        
            sSex = dbrecord("sex").Value & vbNullString
            If sSex = "Male" Then
                frmsres.opt_male.Value = True 'Male
                frmsres.opt_female.Value = False
                frmsres.stdpic.Picture = LoadPicture("")
                frmsres.stdpic.Picture = LoadPicture(App.Path & "\pics\no_male.jpg")
            End If
            If sSex = "Female" Then
                frmsres.opt_female.Value = True 'Female
                frmsres.opt_male.Value = False
                frmsres.stdpic.Picture = LoadPicture("")
                frmsres.stdpic.Picture = LoadPicture(App.Path & "\pics\no_female.jpg")
            End If
   
            frmsres.txtid.Text = dbrecord("ID").Value & vbNullString    'ID
            frmsres.txtroll.Text = dbrecord("Roll").Value & vbNullString    'Roll
            frmsres.txtname.Text = dbrecord("Name").Value & vbNullString    'Name
            frmsres.txtdep.Text = dbrecord("Department").Value & vbNullString   'Department
            frmsres.txtsem.Text = dbrecord("Semester").Value & vbNullString 'Semester
            
            frmsres.txtgrade.Text = dbrecord("Grade").Value & vbNullString
            frmsres.txtcomarch.Text = dbrecord("sub_computer_arc").Value & vbNullString
            frmsres.txtmicro.Text = dbrecord("sub_microprocessor").Value & vbNullString
            frmsres.txtdatabaseman.Text = dbrecord("sub_database_manage").Value & vbNullString
            frmsres.txtvisual.Text = dbrecord("sub_visual_programming").Value & vbNullString
            frmsres.txtdatacomfund.Text = dbrecord("sub_data_communication").Value & vbNullString
            frmsres.txtenviron.Text = dbrecord("sub_environmanal").Value & vbNullString
            frmsres.txtbookkeep.Text = dbrecord("sub_book_keeping").Value & vbNullString
            frmsres.txtbizorg.Text = dbrecord("sub_business_org").Value & vbNullString
            
            
            frmsres.picname.Caption = frmsres.txtid.Text & frmsres.txtroll.Text
        
    DisconnectDB ' Disconnect From Database

End Function

Public Function Edit_data_Result()

Dim db_sub_computer_arc As String
Dim db_sub_microprocessor As String
Dim db_sub_database_manage As String
Dim db_sub_visual_programming As String
Dim db_sub_data_communication As String
Dim db_sub_environmanal As String
Dim db_sub_book_keeping As String
Dim db_sub_business_org As String
Dim db_sub_score As Double
Dim db_Grade As String
Dim db_cGPA As String

db_sub_computer_arc = frmsres.txtcomarch.Text
db_sub_microprocessor = frmsres.txtmicro.Text
db_sub_database_manage = frmsres.txtdatabaseman.Text
db_sub_visual_programming = frmsres.txtvisual.Text
db_sub_data_communication = frmsres.txtdatacomfund.Text
db_sub_environmanal = frmsres.txtenviron.Text
db_sub_book_keeping = frmsres.txtbookkeep.Text
db_sub_business_org = frmsres.txtbizorg.Text
db_sub_score = val(frmsres.txtscore.Text)
db_Grade = frmsres.txtgrade.Text
db_cGPA = frmsres.txtcgpa.Text

frmsres.txtcomarch.BackColor = &H80FF80
frmsres.txtmicro.BackColor = &H80FF80
frmsres.txtdatabaseman.BackColor = &H80FF80
frmsres.txtvisual.BackColor = &H80FF80
frmsres.txtdatacomfund.BackColor = &H80FF80
frmsres.txtenviron.BackColor = &H80FF80
frmsres.txtbookkeep.BackColor = &H80FF80
frmsres.txtbizorg.BackColor = &H80FF80

If frmsres.cmd_edit.Caption = "Save" Then
    connectDB
    ssQL = "Update std_data set sub_computer_arc='" & db_sub_computer_arc & "', sub_microprocessor='" & db_sub_microprocessor & "', sub_database_manage='" & db_sub_database_manage & "', sub_visual_programming='" & db_sub_visual_programming & "', sub_data_communication='" & db_sub_data_communication & "', sub_environmanal='" & db_sub_environmanal & "', sub_book_keeping='" & db_sub_book_keeping & "', sub_business_org='" & db_sub_business_org & "', sub_score='" & db_sub_score & "', Grade='" & db_Grade & "',cGPA='" & db_cGPA & "' where ID=" & val(frmsres.txtid.Text)
    dbcon.Execute (ssQL)
    Call LoadSTD_data_result(frmsres.lststd)
    MsgBox "Record updated sucessfully.", vbInformation, "Done"
    DisconnectDB
    frmsres.cmd_edit.Caption = "Edit"
    
Else
    frmsres.cmd_edit.Caption = "Save"
End If
End Function




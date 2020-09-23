Attribute VB_Name = "Hide_obj"
Public Function hidestart() 'This Function Will hide Start Page Objects

    frmmain.newdb.Visible = False
    frmmain.opendb.Visible = False
    frmmain.exitbt.Visible = False
    frmmain.decpic_1.Visible = False
    frmmain.lbldec.Visible = False
    

End Function

Public Function showstart() 'This Function Will Show Start Page Objects

    frmmain.newdb.Visible = True
    frmmain.opendb.Visible = True
    frmmain.exitbt.Visible = True
    frmmain.decpic_1.Visible = True
    frmmain.lbldec.Visible = True
    

End Function



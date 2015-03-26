
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FrmTerm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub CmdSearch_Click()
'Code version 2.0 created on 3/20/13 add Form_Form1 for front end input and output by Donald Mitchell
    Dim donald As String
  
  
    If FrmProcess.Value = 2 Then
        Debug.Print TxtLanID
    
            SMT_Term2.testid (TxtLanID)
    Else
   
            SMT_Term2.Qry (TxtLanID)

    End If

End Sub

Private Sub Command11_Click() 'close application
 Close

End Sub


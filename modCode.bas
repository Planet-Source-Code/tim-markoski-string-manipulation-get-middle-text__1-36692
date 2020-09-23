Attribute VB_Name = "modCode"
Public Function GetIPfromWeb(strASCiiFrom As String, strIPWebsite As String, strASCiiTo As String) As String
    ' Comments  : (Modified to use proper error handling and naming conventions)
    ' Parameters: strASCiiFrom As String, strIPWebsite As String, strASCiiTo As String
    ' Returns   : String
    ' Modified  : 07/07/2002 TPM
    ' --------------------------------------------------
    
    On Error GoTo PROC_ERR

    Dim strPreIPWeb As String
    
    strPreIPWeb = strIPWebsite
    strPreIPWeb = Mid(strPreIPWeb, InStr(1, strPreIPWeb, strASCiiFrom) + Len(strASCiiFrom))
    strPreIPWeb = Left(strPreIPWeb, InStr(1, strPreIPWeb, strASCiiTo) - Len(strASCiiTo))
    GetIPfromWeb = strPreIPWeb

PROC_EXIT:
    Exit Function
    
PROC_ERR:
    GetIPfromWeb = vbNullString
    MsgBox Err.Description, vbExclamation
    Resume PROC_EXIT

End Function

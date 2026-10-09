Option Explicit

' ============================================================
'  AnonymiseResourceContacts
'  Clears email address and Windows user account on every
'  resource, renames any resource that had an email to a
'  non-descript sequential name, sets the project title to a
'  generic value, and removes all notes (project, task,
'  resource and assignment).
'
'  ALWAYS MAKE A BACKUP COPY OF YOUR PROJECT BEFORE RUNNING THIS MACRO.
' ============================================================

Private Const NAME_PREFIX     As String = "Resource "
Private Const PAD_DIGITS      As Long = 3           ' Resource 001
Private Const DRY_RUN         As Boolean = False    ' True = log only, no changes
Private Const CLEAR_INITIALS  As Boolean = True
Private Const CLEAR_HYPERLINK As Boolean = True
Private Const CLEAR_NOTES     As Boolean = True
Private Const PROJECT_TITLE   As String = "SANITISED"

Sub AnonymiseResourceContacts()

    Dim r           As Resource
    Dim t           As Task
    Dim a           As Assignment
    Dim seq         As Long
    Dim newName     As String
    Dim hadEmail    As Boolean
    Dim cleared     As Long
    Dim renamed     As Long
    Dim failed      As Long
    Dim notesCount  As Long
    Dim undoOpen    As Boolean

    If ActiveProject Is Nothing Then
        MsgBox "No project is open.", vbExclamation
        Exit Sub
    End If

    If Not DRY_RUN Then
        If MsgBox("This will clear email and login account on all " & _
                  ActiveProject.Resources.Count & " resources, rename " & _
                  "any resource that has an email, set the project title to """ & _
                  PROJECT_TITLE & """ and delete ALL notes." & vbCrLf & vbCrLf & _
                  "Run this on a copy. Continue?", _
                  vbYesNo + vbExclamation, "Anonymise resources") <> vbYes Then
            Exit Sub
        End If
    End If

    Application.ScreenUpdating = False
    Application.Calculation = pjManual

    On Error Resume Next
    Application.OpenUndoTransaction "Anonymise resource contacts"
    undoOpen = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0

    Debug.Print String(60, "-")
    Debug.Print IIf(DRY_RUN, "DRY RUN - no changes written", "APPLYING CHANGES")
    Debug.Print "Project: " & ActiveProject.Name
    Debug.Print String(60, "-")

    ' ---- NEW: generic project title ----
    If Not DRY_RUN Then
        On Error Resume Next
        ActiveProject.BuiltinDocumentProperties("Title") = PROJECT_TITLE
        If Err.Number <> 0 Then
            Debug.Print "    ! project title failed: " & Err.Description
            failed = failed + 1
            Err.Clear
        End If
        On Error GoTo 0
    End If

    seq = 0

    For Each r In ActiveProject.Resources

        If Not r Is Nothing Then

            hadEmail = (Len(Trim$(r.EMailAddress & "")) > 0)

            ' ---- rename first, so Initials can be reset afterwards ----
            If hadEmail Then
                seq = seq + 1
                newName = UniqueName(seq)

                Debug.Print "UID " & r.UniqueID & "  """ & r.Name & _
                            """  <" & r.EMailAddress & ">  ->  """ & newName & """"

                If Not DRY_RUN Then
                    On Error Resume Next
                    r.Name = newName
                    If Err.Number <> 0 Then
                        Debug.Print "    ! rename failed on UID " & r.UniqueID & _
                                    ": " & Err.Description
                        failed = failed + 1
                        Err.Clear
                    Else
                        renamed = renamed + 1
                    End If
                    On Error GoTo 0
                Else
                    renamed = renamed + 1
                End If
            End If

            ' ---- clear contact fields on every resource ----
            If Not DRY_RUN Then
                On Error Resume Next

                r.EMailAddress = ""
                r.WindowsUserAccount = ""

                If CLEAR_INITIALS And hadEmail Then r.Initials = "R"
                If CLEAR_HYPERLINK Then
                    r.Hyperlink = ""
                    r.HyperlinkAddress = ""
                    r.HyperlinkSubAddress = ""
                End If
                If CLEAR_NOTES Then r.Notes = ""

                If Err.Number <> 0 Then
                    Debug.Print "    ! clear failed on UID " & r.UniqueID & _
                                ": " & Err.Description
                    failed = failed + 1
                    Err.Clear
                Else
                    cleared = cleared + 1
                End If
                On Error GoTo 0
            Else
                cleared = cleared + 1
            End If

        End If
    Next r

    ' ---- NEW: clear notes on project summary, tasks and assignments ----
    If CLEAR_NOTES Then
        If Not DRY_RUN Then
            On Error Resume Next
            ActiveProject.ProjectSummaryTask.Notes = ""
            Err.Clear
            On Error GoTo 0
        End If

        For Each t In ActiveProject.Tasks
            If Not t Is Nothing Then
                If Not DRY_RUN Then
                    On Error Resume Next
                    t.Notes = ""
                    For Each a In t.Assignments
                        a.Notes = ""
                    Next a
                    If Err.Number <> 0 Then
                        Debug.Print "    ! notes failed on task UID " & t.UniqueID & _
                                    ": " & Err.Description
                        failed = failed + 1
                        Err.Clear
                    End If
                    On Error GoTo 0
                End If
                notesCount = notesCount + 1
            End If
        Next t
    End If

    If undoOpen Then
        On Error Resume Next
        Application.CloseUndoTransaction
        On Error GoTo 0
    End If

    Application.Calculation = pjAutomatic
    Application.ScreenUpdating = True

    Debug.Print String(60, "-")
    Debug.Print "Resources processed : " & cleared
    Debug.Print "Renamed (had email) : " & renamed
    Debug.Print "Tasks notes-cleared : " & notesCount
    Debug.Print "Failures            : " & failed

    MsgBox IIf(DRY_RUN, "DRY RUN complete. Nothing was changed." & vbCrLf & vbCrLf, "") & _
           "Resources processed: " & cleared & vbCrLf & _
           "Renamed (had an email): " & renamed & vbCrLf & _
           "Tasks with notes cleared: " & notesCount & vbCrLf & _
           "Failures: " & failed & vbCrLf & vbCrLf & _
           "Full log is in the Immediate window (Ctrl+G).", _
           vbInformation, "Anonymise resources"

End Sub

' Returns a sequential name that is not already in use.
Private Function UniqueName(ByRef seq As Long) As String
    Dim candidate As String
    Do
        candidate = NAME_PREFIX & Format$(seq, String(PAD_DIGITS, "0"))
        If Not NameInUse(candidate) Then Exit Do
        seq = seq + 1
    Loop
    UniqueName = candidate
End Function

Private Function NameInUse(ByVal nm As String) As Boolean
    Dim r As Resource
    For Each r In ActiveProject.Resources
        If Not r Is Nothing Then
            If StrComp(r.Name, nm, vbTextCompare) = 0 Then
                NameInUse = True
                Exit Function
            End If
        End If
    Next r
End Function
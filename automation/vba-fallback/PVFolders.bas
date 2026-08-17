Attribute VB_Name = "PVFolders"
'==============================================================================
' PV FOLDERS - VBA FALLBACK (Option A: OneDrive-synced SharePoint path)
'------------------------------------------------------------------------------
' Creates a folder per PV (with subfolders) on the locally synced SharePoint
' library and writes the folder's SharePoint URL into column AF of the
' PV Register. Idempotent: rows with a non-empty AF are skipped; existing
' folders are never recreated; duplicate PV numbers are logged and skipped.
'
' USE THIS ONLY IF the Power Automate solution is unavailable. Requirements:
'   - Save the workbook as .xlsm (macro-enabled) and allow macros.
'   - The document library must be synced with OneDrive on THIS PC.
'   - Run from desktop Excel while nobody is editing the file in the browser.
'
' SETUP: edit the four constants below, then run CreatePVFolders (Alt+F8).
' Test first with DRY_RUN = True: it logs what it WOULD do, creates nothing.
'==============================================================================
Option Explicit

' 1) Local root of the synced library folder that will hold PV folders.
'    Find it in File Explorer under the building icon, e.g.:
Private Const LOCAL_ROOT As String = "C:\Users\%USERNAME%\Proforest\PFAFR Finance - Documents\Payment Vouchers"

' 2) The SAME folder as a SharePoint URL prefix (open it in the browser and
'    copy the address bar up to /Payment Vouchers). Spaces as %20:
Private Const URL_ROOT As String = "https://proforest.sharepoint.com/sites/PFAFRFinance/Shared%20Documents/Payment%20Vouchers"

' 3) Subfolders created inside every PV folder (comma-separated):
Private Const SUBFOLDERS As String = "01_Invoice,02_Approvals,03_Payment_Evidence"

' 4) True = log only, create nothing, write nothing:
Private Const DRY_RUN As Boolean = False

Private Const SHEET_REGISTER As String = "" & ChrW(&HD83D) & ChrW(&HDCDD) & " PV Register" ' "📝 PV Register"
Private Const SHEET_LOG As String = "Automation Log (VBA)"
Private Const FIRST_ROW As Long = 5
Private Const LAST_ROW As Long = 510
Private Const MAX_NAME As Long = 80

Public Sub CreatePVFolders()
    Dim ws As Worksheet, wsLog As Worksheet
    Dim r As Long, created As Long, skipped As Long, errs As Long
    Dim pv As String, payee As String, docUrl As String
    Dim dt As Variant, fy As String, monthFolder As String
    Dim folderName As String, localPath As String, webUrl As String
    Dim seen As Object: Set seen = CreateObject("Scripting.Dictionary")

    Set ws = ResolveRegisterSheet()
    If ws Is Nothing Then
        MsgBox "PV Register sheet not found.", vbCritical
        Exit Sub
    End If
    Set wsLog = GetOrCreateLogSheet()

    Dim root As String
    root = Replace(LOCAL_ROOT, "%USERNAME%", Environ$("USERNAME"))
    If Dir(root, vbDirectory) = "" Then
        MsgBox "Local synced folder not found:" & vbCrLf & root & vbCrLf & _
               "Check the LOCAL_ROOT constant and that the library is synced.", vbCritical
        Exit Sub
    End If

    For r = FIRST_ROW To LAST_ROW
        pv = Trim$(CStr(ws.Cells(r, "A").Value))
        If pv = "" Then GoTo NextRow

        If seen.Exists(pv) Then
            LogRow wsLog, pv, r, "SKIPPED", "Duplicate of row " & seen(pv)
            skipped = skipped + 1
            GoTo NextRow
        End If
        seen.Add pv, r

        docUrl = Trim$(CStr(ws.Cells(r, "AF").Value))
        If docUrl <> "" Then GoTo NextRow            ' already linked - idempotency

        dt = ws.Cells(r, "B").Value
        If Not IsDate(dt) Then
            LogRow wsLog, pv, r, "SKIPPED", "Missing or invalid date"
            skipped = skipped + 1
            GoTo NextRow
        End If

        payee = Trim$(CStr(ws.Cells(r, "E").Value))
        If payee = "" Then payee = "Unknown-Payee"
        fy = "FY" & Year(dt)
        monthFolder = Format$(dt, "mm") & "-" & Format$(dt, "mmmm")
        folderName = SanitizeName(pv & "_" & payee)

        localPath = root & "\" & fy & "\" & monthFolder & "\" & folderName
        webUrl = URL_ROOT & "/" & fy & "/" & UrlSeg(monthFolder) & "/" & UrlSeg(folderName)

        On Error GoTo RowError
        If DRY_RUN Then
            LogRow wsLog, pv, r, "DRY RUN (would create)", localPath
        Else
            EnsureDir root & "\" & fy
            EnsureDir root & "\" & fy & "\" & monthFolder
            EnsureDir localPath
            Dim sub_ As Variant
            For Each sub_ In Split(SUBFOLDERS, ",")
                EnsureDir localPath & "\" & Trim$(CStr(sub_))
            Next sub_
            ws.Hyperlinks.Add Anchor:=ws.Cells(r, "AF"), Address:=webUrl, TextToDisplay:=webUrl
            LogRow wsLog, pv, r, "FOLDER LINKED", webUrl
            created = created + 1
        End If
        On Error GoTo 0
        GoTo NextRow

RowError:
        LogRow wsLog, pv, r, "ERROR", Err.Number & ": " & Err.Description
        errs = errs + 1
        Err.Clear
        On Error GoTo 0
NextRow:
    Next r

    MsgBox IIf(DRY_RUN, "DRY RUN - nothing was created." & vbCrLf, "") & _
           "Folders created & linked: " & created & vbCrLf & _
           "Skipped: " & skipped & vbCrLf & "Errors: " & errs & vbCrLf & vbCrLf & _
           "Details in sheet '" & SHEET_LOG & "'." & vbCrLf & _
           "OneDrive will upload new folders in the background - give it a minute " & _
           "before clicking the links.", vbInformation, "PV Folders"
End Sub

'---------------------------------------------------------------- helpers ----

' Emoji sheet names can be fragile across systems; fall back to a name match.
Private Function ResolveRegisterSheet() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ResolveRegisterSheet = ThisWorkbook.Worksheets(SHEET_REGISTER)
    On Error GoTo 0
    If ResolveRegisterSheet Is Nothing Then
        For Each ws In ThisWorkbook.Worksheets
            If InStr(1, ws.Name, "PV Register", vbTextCompare) > 0 Then
                Set ResolveRegisterSheet = ws
                Exit Function
            End If
        Next ws
    End If
End Function

Private Sub EnsureDir(path As String)
    If Dir(path, vbDirectory) = "" Then MkDir path
End Sub

Private Function SanitizeName(name As String) As String
    Dim bad As Variant, s As String
    s = name
    For Each bad In Array("""", "*", ":", "<", ">", "?", "/", "\", "|", "#", "%", "&", "{", "}", "~")
        s = Replace(s, CStr(bad), " ")
    Next bad
    Do While InStr(s, "  ") > 0: s = Replace(s, "  ", " "): Loop
    s = Trim$(s)
    Do While Len(s) > 0 And (Right$(s, 1) = "." Or Right$(s, 1) = " ")
        s = Left$(s, Len(s) - 1)
    Loop
    s = Replace(s, " ", "-")
    If Len(s) > MAX_NAME Then s = Left$(s, MAX_NAME)
    Do While Len(s) > 0 And (Right$(s, 1) = "-" Or Right$(s, 1) = "_" Or Right$(s, 1) = ".")
        s = Left$(s, Len(s) - 1)
    Loop
    SanitizeName = s
End Function

' Minimal URL-encoding for a path segment (spaces already replaced by dashes).
Private Function UrlSeg(seg As String) As String
    Dim i As Long, c As String, out As String
    For i = 1 To Len(seg)
        c = Mid$(seg, i, 1)
        If c Like "[A-Za-z0-9._~-]" Then
            out = out & c
        Else
            Dim b() As Byte, j As Long
            b = StrConv(c, vbFromUnicode) ' UTF-8 via system codepage caveat: names are ASCII after sanitising
            For j = LBound(b) To UBound(b)
                out = out & "%" & Right$("0" & Hex$(b(j)), 2)
            Next j
        End If
    Next i
    UrlSeg = out
End Function

Private Function GetOrCreateLogSheet() As Worksheet
    On Error Resume Next
    Set GetOrCreateLogSheet = ThisWorkbook.Worksheets(SHEET_LOG)
    On Error GoTo 0
    If GetOrCreateLogSheet Is Nothing Then
        Set GetOrCreateLogSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        GetOrCreateLogSheet.Name = SHEET_LOG
        With GetOrCreateLogSheet
            .Range("A1:E1").Value = Array("Timestamp", "PV Number", "Register Row", "Action", "Detail")
            .Range("A1:E1").Font.Bold = True
            .Columns("A").ColumnWidth = 20
            .Columns("E").ColumnWidth = 60
        End With
    End If
End Function

Private Sub LogRow(wsLog As Worksheet, pv As String, r As Long, action As String, detail As String)
    Dim nextRow As Long
    nextRow = wsLog.Cells(wsLog.Rows.Count, "A").End(xlUp).Row + 1
    wsLog.Cells(nextRow, 1).Value = Format$(Now, "yyyy-mm-dd hh:nn:ss")
    wsLog.Cells(nextRow, 2).Value = pv
    wsLog.Cells(nextRow, 3).Value = r
    wsLog.Cells(nextRow, 4).Value = action
    wsLog.Cells(nextRow, 5).Value = detail
End Sub

Attribute VB_Name = "Module13"
Option Explicit
Sub AddStock()
Attribute AddStock.VB_ProcData.VB_Invoke_Func = "M\n14"
'Adds a stock's ticker symbol, name, and grid location to their respective text files. The ticker file is
'   used for the Morningstar section of the TD macro. The other files will possibly be used for automatically
'   adding new funds when running TD macro.
'
'Macro shortcut: Ctrl+Shift+M

    Dim StockInfo() As String
    Dim TextFile As Integer
    Dim TickerPath As String
    Dim NamePath As String
    Dim GridPath As String
    Dim TickerList As String
    Dim UpperTicker As String
    
    'Input stock information
    StockInfo() = Split(InputBox("Please enter the ticker symbol, name and grid location of the stock, " _
        & "separated by commas (No spaces) e.g. ""wec,WEC Energy Group,lv"""), ",")
    If UBound(StockInfo) - LBound(StockInfo) + 1 < 3 Then
        MsgBox ("Macro halted. Information is missing from entered stock.")
        Exit Sub
    End If
    If Len(StockInfo(0)) > 5 Then
        MsgBox ("Macro halted. Ticker symbol of stock is too long.")
        Exit Sub
    End If
    
    'Set file paths to stock information data
    TextFile = FreeFile
    TickerPath = "Z:\YungwirthSteve\Macros\Documents\StockTickers.txt"
    NamePath = "Z:\YungwirthSteve\Macros\Documents\StockNames.txt"
    GridPath = "Z:\YungwirthSteve\Macros\Documents\StockGrid.txt"
    
    Open TickerPath For Input As TextFile
        TickerList = Input(LOF(TextFile), TextFile)
    Close TextFile
    
    'Add stock to lists if it's not already there
    If InStr(TickerList, UCase(StockInfo(0))) = 0 Then
        Open TickerPath For Append As TextFile
            Print #TextFile, "," & UCase(StockInfo(0))
        Close TextFile
        
        Open NamePath For Append As TextFile
            Print #TextFile, "," & StockInfo(1)
        Close TextFile
        
        Open GridPath For Append As TextFile
            Print #TextFile, "," & StockInfo(2)
        Close TextFile
        
        MsgBox (StockInfo(1) & " (" & UCase(StockInfo(0)) & ") added to list.")
    Else
        MsgBox (StockInfo(1) & " (" & UCase(StockInfo(0)) & ") is already on list.")
    End If
End Sub

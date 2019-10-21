Attribute VB_Name = "ImportHolidayData"
Sub Import_holiday_data()

    Application.ScreenUpdating = False
    
    Worksheets("JapaneaseHoliday").Select
    
    Dim URL As String
    Dim destCell As Range
    
    URL = "http://www8.cao.go.jp/chosei/shukujitsu/syukujitsu.csv"
    'URL = "http://www8.cao.go.jp/chosei/shukujitsu/syukujitsu_kyujitsu.csv"
    Set destCell = Cells(3, 1)
    
    Range(Cells(3, 1), Cells(148576, 2)).ClearContents
    
    With destCell.Parent.QueryTables.Add(Connection:="TEXT;" & URL, Destination:=destCell)
        .TextFileStartRow = 1
        .TextFileParseType = xlDelimited
        .TextFileCommaDelimiter = True
        .RefreshStyle = xlOverwriteCells
        .Refresh BackgroundQuery:=False
    End With
    
    
    destCell.Parent.QueryTables(1).Delete
    
    Cells(1, 1).Select
    Application.ScreenUpdating = True


End Sub

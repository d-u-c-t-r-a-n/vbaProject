'
'Created on Tuesday September 24th 2019, 5:06:00 pm
'Author: Duc Tran (duc.tran3@carleton.ca)
'Carleton ID: 101158742
'
Attribute VB_Name = "Module1"
Private Function HolidayList() As Variant
    Dim holiday(14) As String
    Dim holiday_date(14) As Date
    'YYYY-MM-DD
    holiday(0) = "20190805"
    holiday(1) = "20190902"
    holiday(2) = "20191014"
    holiday(3) = "20191225"
    holiday(4) = "20200101"
    holiday(5) = "20200217"
    holiday(6) = "20200518"
    holiday(7) = "20200701"
    holiday(8) = "20200803"
    holiday(9) = "20200907"
    holiday(10) = "20201012"
    holiday(11) = "20201225"
    holiday(12) = "20201226"
    
    'Holiday for cheking historical data
    holiday(13) = "20190701"
    holiday(14) = "20190520"
    
    For i = LBound(holiday_date) To UBound(holiday_date)
        holiday_date(i) = StringToDate(holiday(i))
    Next i
    
    HolidayList = holiday_date
End Function

Function InstanceCalculator(TargetName As Range, NameColumn As Range, DateColumn As Range, HoursColumn As Range, ShowLongTerm As Boolean)
    'TargetName: targeted cell
    'NameColumn: Original name column, with unfiltered entries
    'DateColumn: Original name column, with unfiltered entries
    'HourColumn: Original name column, with unfiltered entries
    'ShowLongterm: Boolean, True for Longterm counts and False for shorterm counts
    Dim result As Integer
    longterm_array = GivenArrayOfDate_ReturnArrayOfLongterm(TargetName, NameColumn, DateColumn, HoursColumn)
    result = modified_InstancesOfOneEmployee(longterm_array, ShowLongTerm)
    
    InstanceCalculator = result
    
End Function

Function Helper_InstanceCalculator()
    MsgBox ("Welcome To InstancesCounter. Please prepare 3 data columns: Worker's Name, Days off, Hours off. Create a new Name column, which has no duplicate name. InstancesCounter will take 4 parameters: Target worker, Worker's Name Column, Days Off Column, Hours Column, and value of True for Longterm count, and False for Instance count.")
End Function

Private Function GetDayOfTheWeek(DayInput As Date)
    'Input: YYYY-MM-DD
    'Output: 1: Mon - 2:Tue - 3:Wed - 4:Thu - 5:Fri - 6:Sat - 7:Sun
    GetDayOfTheWeek = Weekday(DayInput, vbMonday)
End Function

Private Function NumDay(previousday_input As Date, nextday_input As Date) As Integer
    NumDay = Abs(Application.WorksheetFunction.Days(previousday_input, nextday_input))
End Function

Private Function StringToDate(StringInput As String) As Date
    'Input: String in format "yyyymmdd"
    'Output: Date Value, change the type of cell to Date to get format yyyy-mm-dd
    Dim year, month, day As Integer
    year = Int(Left(StringInput, 4))
    day = Int(Right(StringInput, 2))
    month = Int(Mid(StringInput, 5, 2))
    StringToDate = DateSerial(year, month, day)
End Function
   
Private Function Is247(HoursInput As Double) As Boolean
    Dim result As Boolean
    If HoursInput >= 12 Then
        result = True
    Else
        result = False
    Is247 = result
End Function

Private Function StillCountForLongTerm_Normal(previousday_input As Date, nextday_input As Date) As Boolean
    Dim result As Boolean
    Dim pd_weekday, nd_weekday, num_diff, current_missing_day As Integer
    result = True
    
    'Getting weekday value of previousday and nextday
    pd_weekday = GetDayOfTheWeek(previousday_input)
    nd_weekday = GetDayOfTheWeek(nextday_input)
    'Getting number of day off
    num_diff = NumDay(previousday_input, nextday_input)
    'List of missing day
    list_of_missing_day = ListOfMissingDay(previousday_input, nextday_input)
    
    If num_diff <= 4 Then
    
        If (pd_weekday = 5 And nd_weekday = 1) Then
            result = True
                    
        ElseIf HasHoliday(previousday_input, nextday_input) = True Then
            list_of_holiday_in_range = ListOfHolidayInRange(previousday_input, nextday_input)
            
            For i = LBound(list_of_missing_day) To UBound(list_of_missing_day)
                current_missing_day = Weekday(list_of_missing_day(i), vbMonday)
                For j = LBound(list_of_holiday_in_range) To UBound(list_of_holiday_in_range)
                    If list_of_missing_day(i) <> list_of_holiday_in_range(j) Then
                        If CInt(current_missing_day) = 6 Or CInt(current_missing_day) = 7 Then
                            result = True
                        Else
                            result = False
                            GoTo SkipLoop
                        End If
                    End If
                    
                Next j
            Next i
        Else
            result = False
            
        End If
    Else
        result = False
    End If
    
SkipLoop:
    StillCountForLongTerm_Normal = result
End Function

Private Function StillCountForLongTerm_247(previousday_input As Date, nextday_input As Date) As Boolean
    Dim result As Boolean
    Dim pd_weekday, nd_weekday, num_diff As Integer
    result = True
    
    'Getting weekday value of previousday and nextday
    pd_weekday = GetDayOfTheWeek(previousday_input)
    nd_weekday = GetDayOfTheWeek(nextday_input)
    'Getting number of day off
    num_diff = NumDay(previousday_input, nextday_input)
    'List of missing day
    list_of_missing_day = ListOfMissingDay(previousday_input, nextday_input)
    
    If num_diff <= 4 Then
    
        If HasHoliday(previousday_input, nextday_input) = True Then
            list_of_holiday_in_range = ListOfHolidayInRange(previousday_input, nextday_input)
            
            For i = LBound(list_of_missing_day) To UBound(list_of_missing_day)
                For j = LBound(list_of_holiday_in_range) To UBound(list_of_holiday_in_range)
                    If list_of_missing_day(i) <> list_of_holiday_in_range(j) Then
                        result = False
                        GoTo SkipLoop
                    End If
                    
                Next j
            Next i
        Else
            result = False
            
        End If
    Else
        result = False
    End If
    
SkipLoop:
    StillCountForLongTerm_247 = result
End Function

Private Function ListOfMissingDay(previousday_input As Date, nextday_input As Date) As Variant
    'Provided that len(previousday - nextday) <= 7
    Dim day_length As Integer
    Dim current_date As Date
    Dim list_of_day() As Variant
    
    current_date = previousday_input
    day_length = NumDay(previousday_input, nextday_input)
    day_length = day_length - 2
    
    ReDim list_of_day(day_length)
    
    For i = LBound(list_of_day) To UBound(list_of_day)
        list_of_day(i) = current_date + 1
        current_date = list_of_day(i)
    Next i
    ListOfMissingDay = list_of_day
End Function

Private Function ListOfDays(previousday_input As Date, nextday_input As Date) As Variant
    list_of_missing_day = ListOfMissingDay(previousday_input, nextday_input)
    
    Dim list_of_days() As Variant
    Dim day_length As Integer
    Dim current_date As Date
    
    current_date = previousday_input
    day_length = NumDay(previousday_input, nextday_input)
    
    ReDim list_of_days(day_length)
    
    For i = LBound(list_of_days) To UBound(list_of_days)
        list_of_days(i) = current_date
        current_date = current_date + 1
    Next i
    
    ListOfDays = list_of_days
End Function

Private Function ListOfAllDays(previousday_input As Date, nextday_input As Date, IncludeBounds As Boolean) As Variant
    If IncludeBounds = True Then
        ListOfAllDays = ListOfDays(previousday_input, nextday_input)
    Else
        ListOfAllDays = ListOfMissingDay(previousday_input, nextday_input)
    End If
End Function


Private Function HasHoliday(previousday_input As Date, nextday_input As Date) As Boolean
    holiday_list = HolidayList()
    list_of_missing_day = ListOfMissingDay(previousday_input, nextday_input)
    Dim result As Boolean
    For i = LBound(holiday_list) To UBound(holiday_list)
        For j = LBound(list_of_missing_day) To UBound(list_of_missing_day)
            If holiday_list(i) = list_of_missing_day(j) Then
                result = True
                GoTo SkipLoop
            End If
        Next j
    Next i
SkipLoop:
    HasHoliday = result
End Function

Private Function ListOfHolidayInRange(previousday_input As Date, nextday_input As Date)
    holiday_list = HolidayList()
    list_of_missing_day = ListOfMissingDay(previousday_input, nextday_input)
    Dim holiday_in_range() As Variant
    
    Dim counter As Integer
    counter = 0
    
    For i = LBound(holiday_list) To UBound(holiday_list)
        For j = LBound(list_of_missing_day) To UBound(list_of_missing_day)
            If holiday_list(i) = list_of_missing_day(j) Then
                ReDim Preserve holiday_in_range(counter)
                holiday_in_range(counter) = holiday_list(i)
                counter = counter + 1
            End If
        Next j
    Next i
    ListOfHolidayInRange = holiday_in_range
End Function

Private Function IsNext(previousday_input As Date, nextday_input As Date)
    Dim preday, nextday As String
    Dim diff As Integer
    Dim result As Boolean
    preday = CStr(previousday_input)
    nextday = CStr(nextday_input)
    
    preday = Right(preday, 2)
    nextday = Right(nextday, 2)
    diff = CInt(nextday) - CInt(preday)
    
    If diff = 1 Then
        result = True
    Else
        result = False
    End If
    IsNext = result
End Function

Private Function IsLongTerm_Normal(previousday_input As Date, nextday_input As Date) As Boolean
    Dim result As Boolean
    
    If NumDay(previousday_input, nextday_input) <= 4 Then
        If IsNext(previousday_input, nextday_input) = True Then
            result = True
        Else
            result = StillCountForLongTerm_Normal(previousday_input, nextday_input)
        End If
    Else
        result = False
    End If
    
    IsLongTerm_Normal = result
End Function

Private Function IsLongTerm_247(previousday_input As Date, nextday_input As Date) As Boolean
    Dim result As Boolean
    
    If NumDay(previousday_input, nextday_input) <= 4 Then
        If IsNext(previousday_input, nextday_input) = True Then
            result = True
        Else
            result = StillCountForLongTerm_247(previousday_input, nextday_input)
        End If
    Else
        result = False
    End If
    
    IsLongTerm_247 = result
End Function

Private Function InstancesOfOneEmployee(LongtermColumn As Range, ShowLongTerm As Boolean) As Integer
    'Range of the LongtermColumn should exclude the last column as this function cannot limit the range
    'of "for each' function
    Dim current_longterm As Boolean
    Dim longterm_counter, shorterm_counter, true_counter, length, i As Integer
    Dim result As Integer
    Dim MyCell As Range
    Dim SLR As Boolean
    
    length = LongtermColumn.Count()
    
    longterm_counter = 0
    shorterm_counter = 0
    true_counter = 0
    
    'Counting
    If length = 1 Then
        If ShowLongTerm = True Then
            longterm_counter = 0
        Else
            shorterm_counter = 1
        End If
    
    Else
        For Each MyCell In LongtermColumn.Cells
            If MyCell = True Then
                true_counter = true_counter + 1
            Else
                If true_counter >= 5 Then
                    longterm_counter = longterm_counter + true_counter + 1
                    true_counter = 0
                Else
                    shorterm_counter = shorterm_counter + 1
                    true_counter = 0
                End If
            End If
            
        Next MyCell
       
        'Take into account last column
        SLR = LongtermColumn.Cells(length)
        
        If SLR = False Then
            shorterm_counter = shorterm_counter + 1
        Else
            If true_counter >= 5 Then
                        longterm_counter = longterm_counter + true_counter + 1
                        true_counter = 0
                    Else
                        shorterm_counter = shorterm_counter + 1
                        true_counter = 0
                    End If
        End If
    End If
    
    'Showing result
    If ShowLongTerm = True Then
        result = longterm_counter
    Else
        result = shorterm_counter
    End If
    InstancesOfOneEmployee = result
End Function

Private Function modified_InstancesOfOneEmployee(LongtermArray As Variant, ShowLongTerm As Boolean) As Integer
    'LongtermArray should include the last element since the input array already cut the size by 1
    Dim current_longterm As Boolean
    Dim longterm_counter, shorterm_counter, true_counter, length As Integer
    Dim result As Integer
    Dim second_last_element_index As Integer
    Dim SLR As Boolean
    
    length = UBound(LongtermArray) - LBound(LongtermArray) + 1
    second_last_element_index = length - 2
    longterm_counter = 0
    shorterm_counter = 0
    true_counter = 0
    
    'Counting
    If length = 1 Then
        If ShowLongTerm = True Then
            longterm_counter = 0
        Else
            shorterm_counter = 1
        End If
    
    Else
        For i = LBound(LongtermArray) To UBound(LongtermArray)
            If CBool(LongtermArray(i)) = True Then
                true_counter = true_counter + 1
            Else
                If true_counter >= 5 Then
                    longterm_counter = longterm_counter + true_counter + 1
                    true_counter = 0
                Else
                    shorterm_counter = shorterm_counter + 1
                    true_counter = 0
                End If
            End If
            
        Next i
       
        'Take into account last column
        SLR = CBool(LongtermArray(second_last_element_index))
        
        If SLR = False Then
            shorterm_counter = shorterm_counter + 1
        Else
            'If true_counter = 4 Then
            '    longterm_counter = longterm_counter + true_counter + 2
            '    true_counter = 0
            If true_counter >= 5 Then
                longterm_counter = longterm_counter + true_counter + 1
                true_counter = 0
            Else
                shorterm_counter = shorterm_counter + 1
                true_counter = 0
            End If
        End If
    End If
    
    'Showing result
    If ShowLongTerm = True Then
        result = longterm_counter
    Else
        result = shorterm_counter
    End If
    modified_InstancesOfOneEmployee = result
End Function

Private Function GivenName_ReturnArrayOfDate(TargetName As Range, NameColumn As Range, DateColumn As Range) As Variant
    Dim NameCell As Range
    Dim array_length As Integer
    Dim name_address_range As Range
    Dim corresponding_date As Date
    Dim daysoff_array() As Variant
    Dim name_range As String
    Dim counter As Integer
    counter = 0
    
    array_length = CInt(GivenName_ReturnNumOfName(TargetName, NameColumn)) - 1
    ReDim Preserve daysoff_array(array_length)
    
    For Each NameCell In NameColumn.Cells
        If NameCell.Value = TargetName.Value Then
            name_range = CStr(NameCell.Address)
            Set name_address_range = Range(name_range)
            corresponding_date = GetCorrespondingDateValueFromName(name_address_range, DateColumn)
            daysoff_array(counter) = corresponding_date
            counter = counter + 1
            'ReDim Preserve daysoff_array(array_length)
        End If
        
    Next NameCell
    
    'array_length = array_length - 1
    'ReDim Preserve daysoff_array(array_length)
    
    'For i = LBound(daysoff_array) To UBound(daysoff_array)
    '    MsgBox daysoff_array(i)
    'Next i
    
    GivenName_ReturnArrayOfDate = daysoff_array
End Function

Private Function GivenName_ReturnArrayOfHours(TargetName As Range, NameColumn As Range, HoursColumn As Range) As Variant
    Dim NameCell As Range
    Dim array_length As Integer
    Dim name_address_range As Range
    Dim corresponding_hours As Double
    Dim hours_array() As Variant
    Dim name_range As String
    Dim counter As Integer
    
    counter = 0
    
    array_length = CInt(GivenName_ReturnNumOfName(TargetName, NameColumn)) - 1
    ReDim Preserve hours_array(array_length)
         
    For Each NameCell In NameColumn.Cells
        If NameCell.Value = TargetName.Value Then
            name_range = CStr(NameCell.Address)
            Set name_address_range = Range(name_range)
            corresponding_hours = GetCorrespondingHoursValueFromName(name_address_range, HoursColumn)
            hours_array(counter) = corresponding_hours
            counter = counter + 1
            'ReDim Preserve hours_array(array_length)
        End If
        
    Next NameCell
    
    'array_length = array_length - 1
    'ReDim Preserve hours_array(array_length)
    
    'For i = LBound(hours_array) To UBound(hours_array)
    '    MsgBox i
    'Next i
    
    GivenName_ReturnArrayOfHours = hours_array
End Function

Private Function GivenArrayOfDate_ReturnArrayOfLongterm(TargetName As Range, NameColumn As Range, DateColumn As Range, HoursColumn As Range)
    'Output: T or F, exclude the last column
    array_of_date = GivenName_ReturnArrayOfDate(TargetName, NameColumn, DateColumn)
    array_of_hours = GivenName_ReturnArrayOfHours(TargetName, NameColumn, HoursColumn)
    
    Dim date_length, hours_length As Integer
    date_length = UBound(array_of_date) - LBound(array_of_date) + 1
    hours_length = UBound(array_of_hours) - LBound(array_of_hours) + 1
    length_for_longterm_array = date_length - 2
    
    Dim current_date, next_date As Date
    Dim current_hours As Double
    Dim current_T_or_F As Boolean
    Dim longterm_array() As Variant
    
    If (CInt(UBound(array_of_date))) = 0 Then
        ReDim Preserve longterm_array(0)
        longterm_array(0) = False
        GoTo SkipLoop
    ElseIf (CInt(UBound(array_of_date))) = 1 Then
        current_date = array_of_date(0)
        next_date = array_of_date(1)
        current_hours = array_of_hours(0)
        current_T_or_F = IsLongTerm(CDate(current_date), CDate(next_date), CDbl(current_hours))
        If CBool(current_T_or_F) = False Then
            ReDim Preserve longterm_array(1)
            longterm_array(0) = False
            longterm_array(1) = True
        Else
            ReDim Preserve longterm_array(0)
            longterm_array(0) = False
        End If
        GoTo SkipLoop
    End If
    
    ReDim Preserve longterm_array(length_for_longterm_array)
    
    For i = LBound(array_of_date) To (UBound(array_of_date) - 1)
        current_date = array_of_date(i)
        next_date = array_of_date(i + 1)
        current_hours = array_of_hours(i)
        current_T_or_F = IsLongTerm(CDate(current_date), CDate(next_date), CDbl(current_hours))
        longterm_array(i) = current_T_or_F
    Next i
    
    'For i = LBound(longterm_array) To UBound(longterm_array)
    '    MsgBox longterm_array(i)
    'Next i
SkipLoop:
    GivenArrayOfDate_ReturnArrayOfLongterm = longterm_array

End Function

Private Function IsLongTerm(previousday_input As Date, nextday_input As Date, hours As Double) As Boolean
    Dim result As Boolean
    If hours <= 8 Then
        result = IsLongTerm_Normal(previousday_input, nextday_input)
    Else
        result = IsLongTerm_247(previousday_input, nextday_input)
    End If
    IsLongTerm = result
End Function

Private Function GetCorrespondingDateValueFromName(NameCell As Range, DateColumn As Range) As Date
    'Input: NameCell as address of SINGLE (01) CELL
    'Input: DateColumn as the whole column that include date
    'Output: Date
    Dim name_address, date_address, value_of_date_from_name As String
    Dim length, limitbound As Integer
    name_address = CStr(NameCell.Address)
    date_address = CStr(DateColumn.Address)
    
    length = Len(name_address)
    limitbound = length - 3
    name_address = Mid(name_address, 4, limitbound)
    date_address = Left(date_address, 2)
    date_address = date_address & "$" & name_address
    
    value_of_date_from_name = CDate(Range(date_address).Value)
    GetCorrespondingDateValueFromName = value_of_date_from_name
    
End Function

Private Function GetCorrespondingHoursValueFromName(NameCell As Range, HoursColumn As Range) As Double
    'Input: NameCell as address of SINGLE (01) CELL
    'Input: DateColumn as the whole column that include date
    'Output: Date
    Dim name_address, hours_address, value_of_hours_from_name As String
    Dim length, limitbound As Integer
    name_address = CStr(NameCell.Address)
    hours_address = CStr(HoursColumn.Address)
    
    length = Len(name_address)
    limitbound = length - 3
    name_address = Mid(name_address, 4, limitbound)
    hours_address = Left(hours_address, 2)
    hours_address = hours_address & "$" & name_address
    
    value_of_hours_from_name = CDbl(Range(hours_address).Value)
    GetCorrespondingHoursValueFromName = value_of_hours_from_name
    
End Function

Private Function GivenName_ReturnNumOfName(TargetName As Range, NameColumn As Range)
    Dim counter As Integer
    Dim NameCell As Range
    counter = 0
    For Each NameCell In NameColumn.Cells
        If NameCell.Value = TargetName.Value Then
            counter = counter + 1
        End If
    Next NameCell
    
    GivenName_ReturnNumOfName = counter
    
End Function

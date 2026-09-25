Attribute VB_Name = "DateFunctions"
Option Compare Database

Function GetNowLast(inputDate As Date) As Date

    dYear = Year(inputDate)
    dMonth = Month(inputDate)

    getDate = DateSerial(dYear, dMonth + 1, 0)

    GetNowLast = getDate

End Function

Public Function EOMonth(dInput As Date)

LastDayOfMonth = DateSerial(Year(dInput()), Month(dInput() + 1), -1)

End Function

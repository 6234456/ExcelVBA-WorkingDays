Sub main()
    
    
    ' workdays in July in Bayern
    d1 = DateSerial(2015, 7, 1)
    d2 = DateSerial(2015, 7, 31)
    
    Debug.Print Workdays.vonbis(d1, d2, Workdays.feiertagMixin(2015, "b"))
    
    ' workdays in the time span 12.2015-01.2016 in Thueringen
    
    d3 = DateSerial(2015, 12, 1)
    d4 = DateSerial(2016, 1, 1)
    
    Debug.Print Workdays.vonbis(d3, d4, Workdays.mixins(Array(2015, 2016), "t"))
    
    
    
End Sub

Function Add(x As Integer, y As Integer)
  Add x + y
END Function
Function AddRange(x As Range, y As Range)
  return Add(x.Value,y.Value)
END Function

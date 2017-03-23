Attribute VB_Name = "KNN_Classification"
'A simple implementation of KNN by Excel VBA ver. 1.
'A Dataset with 2 features and 2-classes label is tested with desirable testing result
'Thereotically, should work with the dataset with more than 2 features and more than 3 classes label

'Use the function "knn_classification mainly" for training and classification
'by K.H.Ni (Bryan) 2017.03.24


Dim sampleSize As Integer
Dim dimSize As Integer

Function KNN_Classification(inputVal As Range, x As Range, y As Range, k As Integer)
    'inputVal: single data point (for testing)
    'x: feature dataset for training
    'y: corresponded label of x
    'k: number of neighbor specified by user
    
    Dim distanceMatrix() As Double
    Dim sortedIndex() As Integer
    Dim neighborlabel()
    
    sampleSize = x.count / 2
    dimSize = inputVal.Column
    ReDim ditanceMatrix(sampleSize - 1) As Double
    ReDim sortedMatrix(sampleSize - 1) As Integer
    ReDim neighborlabel(sampleSize - 1)
    
    computeDistanceMatrix inputVal, x, distanceMatrix
    getSortedMatrix distanceMatrix, sortedIndex
    KNN_Classification = getSelectedLabel(sortedIndex, y, k)
    
End Function


Function computeDistance(x1, x2)
Dim x_diff() As Double
ReDim x_diff(x1.count - 1) As Double
    For i = 1 To x1.count
        x_diff(i - 1) = x1(i) - x2(i)
    Next i
    computeDistance = WorksheetFunction.SumSq(x_diff)
    
End Function
Private Sub computeDistanceMatrix(inputVal As Range, x As Range, ByRef distMatrix() As Double)
    ReDim distMatrix(sampleSize - 1) As Double
    Dim x_diff() As Double
    ReDim x_diff(inputVal.count - 1) As Double

    Dim tempVal As Double
    
    For i = 1 To sampleSize
       For j = 1 To dimSize
             x_diff(j - 1) = inputVal(j) - x(i, j)
       Next j
       distMatrix(i - 1) = WorksheetFunction.SumSq(x_diff)
    Next i
   
End Sub

Private Sub getSortedMatrix(distMatrix() As Double, ByRef sortedIndexSet() As Integer)
    
    ReDim sortedIndexSet(sampleSize - 1) As Integer
    Dim indexArray()
    ReDim indexArray(UBound(distMatrix))
    Dim sortedArray()
    
    For i = 0 To UBound(distMatrix)
        indexArray(i) = i + 1
    Next i
    
    arraySort distMatrix, indexArray, sortedArray
    
    For i = 0 To UBound(distMatrix)
        sortedIndexSet(i) = sortedArray(i, 1)
    Next i
End Sub


Private Function getSelectedLabel(sortedIndex() As Integer, classlabel As Range, k_neighbor As Integer)
    Dim selectedIndex
    Dim maxCount, tempCount As Integer
    selectedIndex = -1
    maxCount = 0
   
   For i = 0 To k_neighbor - 1
      tempCount = WorksheetFunction.CountIf(classlabel, "=" & classlabel(sortedIndex(i)))
      For j = 0 To k_neighbor - 1
        If classlabel(sortedIndex(i)) = classlabel(sortedIndex(j)) Then
            tempCount = tempCount + 1
        End If
      Next j
      
      If tempCount > maxCount Then
         maxCount = tempCount
         selectedIndex = sortedIndex(i)
      End If
   
   Next i
    
   getSelectedLabel = classlabel(selectedIndex)
End Function

 Sub arraySort(valueSet() As Double, indexSet() As Variant, ByRef sortedArray() As Variant)
 'sort by ascending order
 'sort element from value set and stored into sortedArray
 
    ReDim sortedArray(UBound(valueSet), 2)
    Dim tempVal As Double
    Dim tempIndex
    Dim i As Integer
    Dim j As Integer
    tempVal = 0
    
  'bubble sort
   For i = LBound(valueSet) To UBound(valueSet) - 1
    For j = i + 1 To UBound(valueSet)
      If valueSet(i) > valueSet(j) Then
        tempVal = valueSet(i)
        valueSet(i) = valueSet(j)
        valueSet(j) = tempVal
        
        tempIndex = indexSet(i)
        indexSet(i) = indexSet(j)
        indexSet(j) = tempIndex
        
      End If
    Next j
  Next i
    
    
    For i = LBound(valueSet) To UBound(valueSet)
        sortedArray(i, 0) = valueSet(i)
        sortedArray(i, 1) = indexSet(i)
    Next i
    

   
End Sub


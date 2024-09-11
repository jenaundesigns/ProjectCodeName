Sub GenerateCodeNameWithExpandedAnimalList()
    Dim colors As Variant
    Dim animals As Variant
    Dim color As String
    Dim animal As String
    Dim nonce As String
    Dim randomColor As Integer
    Dim randomAnimal As Integer
    Dim projectName As String
    Dim animalRange As Range
    Dim animalList As Range
    Dim lastRow As Long
    
    ' Define an array of colors
    colors = Array("Red", "Blue", "Green", "Yellow", "Black", "White", "Orange", "Purple", "Pink", "Gray", "Cyan", "Magenta", "Lime", "Aqua", "Navy", "Maroon")

    ' Find the last row of animals in the AnimalsList sheet
    With Sheets("AnimalsList")
        lastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
        Set animalRange = .Range("A1:A" & lastRow)
    End With
    
    ' Convert the range of animals into an array
    animals = Application.Transpose(animalRange.Value)
    
    ' Generate a random nonce (3-character alphanumeric string)
    nonce = GenerateNonce(3)
    
    ' Generate random indices for color and animal
    Randomize
    randomColor = Int((UBound(colors) - LBound(colors) + 1) * Rnd + LBound(colors))
    randomAnimal = Int((UBound(animals) - LBound(animals) + 1) * Rnd + LBound(animals))
    
    ' Get the random color and animal
    color = colors(randomColor)
    animal = animals(randomAnimal)
    
    ' Combine color, animal, and nonce to create the project name
    projectName = color & " " & animal & " " & nonce
    
    ' Output the generated project code name to a cell
    Sheets("ProjectCodeName").Range("B2").Value = projectName
    MsgBox "Your generated project codename is: " & projectName
End Sub

' Function to generate a random alphanumeric nonce of specified length
Function GenerateNonce(length As Integer) As String
    Dim chars As String
    Dim result As String
    Dim i As Integer
    
    ' Characters allowed in the nonce (you can modify this to include other characters)
    chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789"
    
    ' Initialize the result string
    result = ""
    
    ' Randomly select characters from the chars string
    For i = 1 To length
        result = result & Mid(chars, Int((Len(chars) * Rnd) + 1), 1)
    Next i
    
    ' Return the generated nonce
    GenerateNonce = result
End Function

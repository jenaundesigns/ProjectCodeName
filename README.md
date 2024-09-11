# ProjectCodeName 👾
⚡Project Codenames generator for Microsoft Excel (.xlsm /macro enabled) using VBA. 

A *codename*, also known as a *codeword* or *cryptonym*, is a name used to discreetly refer to something or someone while keeping their true identity hidden. 
These terms—codename, codeword, and cryptonym—are often used interchangeably. 
Codenames serve to identify projects, products, or operations in a way that conceals their nature from competitors, adversaries, or unauthorized parties.
In simpler terms, a codename is an agreed-upon name used to conceal what you're referring to. 
In industry, codenames are commonly assigned to products during the development phase to safeguard them from competitors. 
Similarly, in the military, codenames are frequently used to refer to missions or operations, ensuring their details remain confidential.

⚡How this VBA was coded This code will generate a project codename where:

	The first output is a random color.
	The second output is a random animal.
	The third output is a randomly generated alphanumeric nonce (3 characters long).

⚡How it works (*Do not copy from these sub-sections, copy entire VBA code further down below, or on separate file attached to this repository*):

 - 🖍 Color: The code selects a random color from the colors array.
 	
	 	' Define an array of colors
   		colors = Array("Red", "Blue", "Green", "Yellow", "Black", "White", "Orange", "Purple", "Pink", "Gray", "Cyan", "Magenta", "Lime", "Aqua", "Navy", "Maroon")
   
 - 🐈 Animal List from Excel: The list of animals is read dynamically from the AnimalsList sheet.

		' Find the last row of animals in the AnimalsList sheet
  		With Sheets("AnimalsList")
        	lastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
        	Set animalRange = .Range("A1:A" & lastRow)
  
 - 🧮 Random Selection: A random color and animal are selected from the arrays, and a nonce is generated.

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

- ✅ Result: The code name is a combination of the random color, animal, and nonce.

		 ' Output the generated project code name to a cell
  		Sheets("ProjectCodeName").Range("B2").Value = projectName
  		MsgBox "Your generated project codename is: " & projectName

Name Excel Spreadsheet: "Project_Codename_Generator.xlsm"

![image](https://github.com/user-attachments/assets/3191cc66-e372-4b06-a596-d07786d1375a)

Rename "Sheet1" as "ProjectCodeName"

Rename "Sheet2" as "AnimalsList"

![image](https://github.com/user-attachments/assets/c69aeae9-02b4-49da-8aa9-36b1f606bac8)

In column A of the "AnimalsList" sheet, list all the animals starting from cell A1. I have listed the animals I used below, feel free to add as many as you like, the VBA is written to accept as many as you'd like.

*Note* Make sure if you copy/paste in excel, hit the arrow under the Paste button > Paste Special > Text > OK

![image](https://github.com/user-attachments/assets/50d95e6a-759c-42b6-8463-28ca5971d1af)


	Falcon
	Phoenix
	Eagle
	Tiger
	Lion
	Cheetah
	Elephant
	Shark
	Dolphin
	Penguin
	Giraffe
	Wolf
	Bear
	Whale
	Panther
	Crocodile
	Leopard
	Zebra
	Koala
	Owl
	Horse
	Kangaroo
	Otter
	Swan
	Turtle
	Fox
	Rabbit
	Peacock
	Hawk
	Raccoon
	Armadillo
	Bison
	Cobra
	Deer
	Flamingo
	Gecko
	Hedgehog
	Iguana
	Jackal
	Lemur
	Meerkat
	Narwhal
	Ocelot
	Parrot
	Quokka
	Rhinoceros
	Salamander
	Tortoise
	Urial
	Vulture
	Wombat
	Fly
	Yak
	Panda
	Pelican
	Djinn
	Serval
	Dart Frog
	Sea Dragon
	Tamarin
	Komodo Dragon
	Pika
	Draco
	Kudu
	Zeta Reticulan
	Pleiadian
	Anunnaki
	Wizard
	Warthog
	Thule
	Reindeer
 

*Optional* add a 3rd sheet "CodeName Log" to keep a log of the different Project CodeName's already used, such as:
![image](https://github.com/user-attachments/assets/1618e0c5-d47f-4581-b97a-774c26196dc3)

VBA Code can be copy/pasted into VBA module from the file attached to this repository, see (Project_Codename_Generator.xlsm - Module1), or directly below:

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

On "ProjectCodeName" sheet, make sure you set it up so your project name response can populate on cell "B2", which is how this VBA cade was setup (or change it in the indicated field on the VBA module):

![image](https://github.com/user-attachments/assets/380d0561-c4fd-4c2a-a14f-0d516736052f)

For ease of convenience, also navigate to the Developer tab on Excel, and under Controls > Insert > Button (under Forms), and draw button next to B2 cell (select the macro after you have saved it), then right click to edit text inside button (used GENERATE here):

![image](https://github.com/user-attachments/assets/e684ff09-98fa-46e6-8bf4-117f1c2ec98a)

The use of cryptonyms combined with nonces is crucial for ensuring compartmentalization in sensitive or classified projects, particularly in fields like intelligence, security, and research. This is a great and easy way to maintain cryptonyms, to add a layer of protection for your projects and maintain compartmentalization. Recommend keeping the cryptonym key with the actual project names (deciphered) in a completely seperate file location (as a standalone, preferably), such as on an encrypted cold storage device (password protected) that is in a secure lock box, etc. 

✨**Additional ways to add more Security**✨

🔑For more security, password protect the Workbook with a randomly generated password from a secure password generator, such as NordPass Password Generator & Digital Vault. 

🔑Add additional characters (including keyboard symbols) to the chars array (characters allowed in the nonce).

🔑Increase the length of the [nonce = GenerateNonce(3)] from (3) to a higher integer (#).

🔑Add additional colors periodically within the VBA code under define array of colors [colors = Array("Red", "Blue", )].

🔑Add additional animals to the AnimalsList sheet of the workbook.




Let me know if you find this helpful and Star this project! ✨

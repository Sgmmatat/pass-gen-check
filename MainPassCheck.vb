Module Module1

    Sub Main()
        Dim FirstN, FirstNU As String
        Dim SurN, SurC As String
        Dim DOB As Date ' saved as date so a valid date is entered
        Dim YOE, YOE_Length As Integer
        Dim Post As String
        Dim answera As String
        Dim password As String
        Dim password2 As String
        Dim RandN As Integer
        Dim generateRN As New Random() ' saved as random so i can generate a random number
        Dim generateRN2 As New Random()
        Dim usablechars() As String
        Dim RNDpassword As String = "" ' RNDpassword = nothing so variable starts from blank
        Dim ascP As Integer
        Dim SurC2 As String

        'PROMPT USER TO ENTER PERSONAL INFORMATION
        'here the console will ask the user for there personal information so it can generate a random password based on the users input
        'prompts the user for their first name 
        Console.WriteLine("Please enter First Name")
        FirstN = Console.ReadLine
        'stores a copy of their first name in capital letters using the Ucase() string function
        FirstNU = UCase(FirstN)
        'prompt the user for their surname
        Console.WriteLine("Please enter Second Name")
        SurN = Console.ReadLine
        'using the Mid() string function to get characters from the their surname 
        'selects the 2nd position in the string from 0, then counts on 1 more space and copies and saves the new string in SurC and SurC2 variable
        SurC = Mid(SurN, 2, 1)
        SurC2 = Mid(SurN, 3, 1)
        'prompts user for there date of birth in a certain format
        Console.WriteLine("Please enter Date of Birth (dd/mm/yyyy)")
        DOB = Console.ReadLine
        'prompts user for their year of entry
        Console.WriteLine("Please enter your Year of Entry")
        YOE = Console.ReadLine
        'save the length of the user's year of entry in a different variable by usnig the Len() function
        YOE_Length = Len(YOE)
        'prompts user for their postcode
        Console.WriteLine("Please enter Postcode")
        Post = Console.ReadLine

        'ascii value of (post) is then stored into a new vaiable so a substring can be retrieved
        ascP = ascP & Asc(Post)

        Console.WriteLine("Would you like to generate/check a password?")
        answera = Console.ReadLine


        ' select case created to execute certain code depending on what the user wnats to do and type (check/generate)
        Select Case answera

            Case Is = "generate"

                'using the sustring function to get first letter of the user's first name and save in variable password
                password = FirstNU.Substring(0, 1)
                'using the chr() function and my random number generated in variable RandN to generate a random special symbol and save  
                'generate random number using Next function to generate a number within a specific range and save into variable RandN. the range given is equivalent to the integer values which correspond to the specia symbols required
                RandN = generateRN.Next(33, 47)
                ' " " is used to separate the new strings added to password variable so it can be split later into an array
                'using the chr() function and my random number generated in variable RandN to generate a random special symbol and save  
                password = password & " " & Chr(RandN)
                'using the asc() function to generate a random number. using the users postcode, the first letter is compared to its ascii value which is then saved into the variable password
                'ToString manipulation is used to convert integer variable into string which then can con be manipulatd with Substring function
                password = password & " " & ascP.ToString.Substring(0, 1)
                'the users' length of password is then added to the variable password
                password = password & " " & YOE_Length
                'lowercase letters stored in SurC variable are then added to the variable password
                password = password & " " & SurC
                'using the substring function to get the characters within a range of positions in the string and add to variable password
                password = password & " " & Post.Substring(4, 1)
                password = password & " " & DOB.ToString.Substring(1, 1)
                'ToUpper functionis used to get what a string and convert to uppercase which is then added to password variable
                password = password & " " & SurC2.ToUpper
                password = password & " " & Post.Substring(3, 1)
                password = password & " " & Post.Substring(1, 1)

                'using the Split function to split password variable by the " " (spaces) into an array of characters which is then stored in the variable usablechar()
                usablechars = Split(password, " ")

                'For loop is then used to build the random password. For i = (0) to (1) -1 is used so the code inside the For loop only iterates once before exiting.
                'this allows me to edit the number of code inside, which allows me to cap the length of the password to a set length. In this case the length is 10 as there it 10 lines of code 
                For i = (0) To (1) - 1
                    'here the password is created by calling to a position in the usablechars() array and storing the character in the RNDpassword variable
                    RNDpassword = RNDpassword & usablechars(6)
                    RNDpassword = RNDpassword & usablechars(4)
                    RNDpassword = RNDpassword & usablechars(0)
                    RNDpassword = RNDpassword & usablechars(5)
                    RNDpassword = RNDpassword & usablechars(1)
                    RNDpassword = RNDpassword & usablechars(2)
                    RNDpassword = RNDpassword & usablechars(3)
                    RNDpassword = RNDpassword & usablechars(7)
                    RNDpassword = RNDpassword & usablechars(8)
                    RNDpassword = RNDpassword & usablechars(9)
                    ' once all the desired positions have been called to and the corresponding character is stored into the RNDpassowrd variable, the new password is then displayed to the user
                Next
                Console.WriteLine("Your random password is: " & RNDpassword)
                Console.ReadLine()



            Case Is = "check"

                'if the user wants to check wheather an password is valid then this code executes and checks if their password is suitable
                Console.WriteLine("Here you will create a valid pssword")
                'display set of rules the user's password needs to follow to be valid
                Console.WriteLine("Password should include: Uppercase and Lower case letters, numbers, special character and is 10 characters long")
                'users entry is then stored password2 variable 
                Do
                    password2 = Console.ReadLine

                    'the users password is then compared against formats using LIKE to check if the password contains an uppercase, lowercase, digits and has a length of 15
                    If (password2 Like "*[a-z]*") And (password2 Like "*[A-Z]*") And (password2 Like "*[0-9]*") And password2 Like "*[!A-Z]*" And Len(password2) = 10 Then
                        'if password is valdi then this message displays
                        Console.WriteLine("Password: " & password2 & " is valid")
                        Console.ReadLine()
                    Else
                        'if password is not valid the user is prompted to re-enter a valid password
                        Console.WriteLine("Error, password does not fit the criteria.")

                    End If
                    'Do Until loop is created so the user can re-enter passwords until a password is valid
                Loop Until (password2 Like "*[a-z]*") And (password2 Like "*[A-Z]*") And (password2 Like "*[0-9]*") And password2 Like "*[!A-Z]*" And Len(password2) = 10
        End Select
    End Sub

End Module

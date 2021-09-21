Attribute VB_Name = "mTinderBOT"
'---------------------------------------------------------------------------------------
' Module    : mTinderBOT
' Author    : Sven From Coding Is Fun
' YouTube   : https://youtube.com/c/CodingIsFun
' Website   : https://pythonandvba.com
' Date      : 8/29/2021
' Purpose   : AutoSwiper for Tinder by automating Chrome Browser with VBA Selenium
'---------------------------------------------------------------------------------------

Option Explicit

Sub TinderBOT()

10    On Error GoTo ErrorHandler
          Dim bot As Object
          Dim by As Object
          Dim wb As Workbook
          Dim ws As Worksheet
          Dim RandomNumber As Double
          Dim LikesTarget As Long
          Dim LikesGiven As Long
          
20        Set wb = ThisWorkbook
30        Set ws = wb.Sheets("Sheet1")
40        Set bot = CreateObject("Selenium.WebDriver")
50        Set by = CreateObject("Selenium.by")
          
          'Use userprofile to store login data (cookies)
          'Navigate to 'chrome://version/', copy/paste Profile Path below
60        bot.SetProfile "REPLACE WITH PROFILE PATH"
          
          'Alternative solution:
          'bot.SetProfile Environ("LOCALAPPDATA") & "\Google\Chrome\User Data"
          
          'SET TIMEOUTS
70        bot.Timeouts.ImplicitWait = 150000 'Default 3000
80        bot.Timeouts.PageLoad = 150000 'Default 60000
90        bot.Timeouts.Server = 150000 'Default 90000
          
          'BOT ARGUMENTS
100       bot.AddArgument "--disable-popup-blocking"
110       bot.AddArgument "--disable-notifications"
          
          'Init chrome & navigate to Tinder
120       bot.Start "chrome"
130       bot.Get "https://tinder.com"
          
140       MsgBox _
              "Please login with your credentials." & vbCrLf & _
              "Once you are logged in, please click on 'OK' to continue.", vbOKOnly
          
150       LikesTarget = ws.Range("C4").Value
160       LikesGiven = 0
          
170       Do While LikesGiven <= LikesTarget
          
180           RandomNumber = Application.WorksheetFunction.RandBetween(500, 1000)
              
             'Swipe right & increment LikesGiven by 1 for each iteration
190           bot.SendKeys bot.Keys.ArrowRight
200           LikesGiven = LikesGiven + 1
210           bot.Wait (RandomNumber)
              
              'Check for popup if you run out of likes
220           If bot.IsElementPresent(by.XPath("//button[@aria-labelledby='subscription-option-information']")) Then
230               Exit Do
240           End If
              
              'Escape potential popus ('Matches', 'Add to homescreen', 'Upgrade Tinder', ..)
250           bot.SendKeys bot.Keys.Escape
260           bot.Wait (RandomNumber)
              
270       Loop
          
280       MsgBox "Done :)"
290       bot.Quit
          
300       Exit Sub
          
ErrorHandler:
310       Select Case Err.Number
          Case -2146232576  'Could be because of .Net Framework not installed
320           MsgBox _
                  "Oh, it looks like that you first need to activate/install your .NET Framework." & _
                  "But no worries! To fix this issue, follow the steps here:" & vbNewLine & _
                  "https://pythonandvba.com/automation-error" & vbNewLine & vbNewLine & _
                  Err.Number & ": " & Err.Description, , "Tinder BOT"
330           Exit Sub
340       Case Else
350           MsgBox _
                  "An unexpected error has been detected" & vbNewLine & vbNewLine & _
                  "System Information:" & Chr(13) & _
                  "Microsoft Excel version " & Application.Version & _
                  " running on " & Application.OperatingSystem & vbNewLine & vbNewLine & _
                  "Description is: " & Err.Number & ", " & Err.Description & vbNewLine & vbNewLine & _
                  "Error occurred on line: " & Erl & vbNewLine & _
                  "Module is: TinderBOT" & vbNewLine & vbNewLine & _
                  "Please note the above details before contacting support: support@pythonandvba.com"
360           Exit Sub
370       End Select

End Sub


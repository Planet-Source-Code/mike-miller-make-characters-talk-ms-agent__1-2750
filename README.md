<div align="center">

## Make Characters Talk\! MS Agent


</div>

### Description

It will make a character of your choice(many to download) talk.
 
### More Info
 
The MS Agent Character

Make Sure you have MS Agent installed and TruVoice, you can get these from Microsoft's Website at http://msdn.microsoft.com/workshop/imedia/agent/default.asp. This will allow you to make the characters talk. After you got the files, just insert the ms agent active x control and you are on your way. All you got to do then is insert the code and download whatever characters you want to use, for example a cool parrot, a butler, surfmonkey and much more!

Voice


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Mike Miller](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/mike-miller.md)
**Level**          |Unknown
**User Rating**    |5.0 (5 globes from 1 user)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VB Script
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/mike-miller-make-characters-talk-ms-agent__1-2750/archive/master.zip)





### Source Code

```
Dim Genie As IAgentCtlCharacterEx
Const DATAPATH = "genie.acs"
Private Sub Form_Load()
  Agent1.Characters.Load "Genie", DATAPATH
  Set Genie = Agent1.Characters("Genie")
  Genie.LanguageID = &H409
  TextBox.Text = "Hello World!"
End Sub
Private Sub Button_Click()
  Genie.Show
  Genie.Speak TextBox.Text
  Genie.Hide
End Sub
```


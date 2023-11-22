- 👋 Hi, I’m @t5h2i0tadi
- 👀 I’m interested in ...
- 🌱 I’m currently learning ...
- 💞️ I’m looking to collaborate on ...
- 📫 How to reach me ...

<!---
t5h2i0tadi/t5h2i0tadi is a ✨ special ✨ repository because its `README.md` (this file) appears on your GitHub profile.
You can click the Preview link to take a look at your changes.
--->

META DATABASE LICENSE ISSUE , CERTIFICATE

registration page for create new account and home page which show message that you login successfully .I hope you will like this article .So let s start step by step.

 

Step: 1:- Open your visual studio, here I will use visual studio 2019.

Step: 2:- Clock on file menu on top of the visual studio, hover mouse on new and click on project.

 

Step: 3:- Search for windows form App.(.Net framework) and click on next.

 

Step: 4:- In this step you have to enter some details of your application and then click on Create button. Following details you have to enter.

 

1. Project Name: Name of your project

2. Location: Location where you want to store your app in your local computer.

3. Solution Name: This name is display in solution explore in visual studio.

4. Framework: Select appropriate framework as your application require.

Step: 5:- Now your project is created. Open Solution Explorer .If you don t see solution explore you can open from View menu on top or you can try short cut key Ctrl+W,S . We need to create some pages for our application. Right click on solution name then Hover mouse on Add and click on Add New Item or you can user short cut key Ctrl+Shift+A .

 

Step: 6:- Now you see a dialog where we add our forms. Select Windows Form, give proper name and click on Add. Add Login, Registration and Home page in same way.

 

Step: 7:- Now we need to add database in our project. Right click on solution name then Hover mouse on Add and click on Add New Item or you can user short cut key Ctrl+Shift+A . Select data filter from left side bar for see item which associate with database. Select service based database, give name and click on add.

 

Step: 8:- Now we create a table which we user in login and registration. Double click on database file from solution explorer. It will open database file in server explore. Expand your database and right click on table then click on Add New Table.

 

Step: 9:-Create table field which you want, here I added only three field Id, UserName and password where id is auto increment by 1. You can set it by right click on field name , click on property and find Id Identity Specification expand it make true (Is Identity) field and give increment number which increment id by adding this number in last id.

 

Step:-10:- Now first of all we create Registration form. So create design of your form as you need. In below image you see how I design form.

 

Step: 11:- Now click anywhere on form it will generate Form_Load event where enter following code. This code create database connection and open it. In next step you will learn how you get that































































VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   8085
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   16815
   OleObjectBlob   =   "UserForm document office.frx":0000
   StartUpPosition =   3  'Windows Default
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ComboBox1_Change()
 
End Sub
 
Private Sub CommandButton1_Click()
 
End Sub
 
Private Sub CommandButton2_Click()
 
End Sub
 
Private Sub Frame1_Click()
 
End Sub
 
Private Sub UserForm_Click()
 
End Sub
 
Private Sub UserForm_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
 
End Sub
 
Private Sub UserForm_Error(ByVal Number As Integer, ByVal Description As MSForms.ReturnString, ByVal SCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, ByVal CancelDisplay As MSForms.ReturnBoolean)
 
End Sub
 
Private Sub UserForm_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
 
End Sub
 
Private Sub UserForm_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
 
End Sub
 
Private Sub UserForm_RemoveControl(ByVal Control As MSForms.Control)
 
End Sub
 
Private Sub UserForm_Resize()
 
End Sub
 
Private Sub UserForm_Scroll(ByVal ActionX As MSForms.fmScrollAction, ByVal ActionY As MSForms.fmScrollAction, ByVal RequestDx As Single, ByVal RequestDy As Single, ByVal ActualDx As MSForms.ReturnSingle, ByVal ActualDy As MSForms.ReturnSingle)
 
End Sub
 
Private Sub UserForm_Terminate()
 
End Sub
 
Private Sub UserForm_Zoom(Percent As Integer)
Private Sub cmd_bgColor_Click()
 Dim r, g, b As Integer
 r = Int(Rnd() * 256)
 g = Int(Rnd() * 256)
 b = Int(Rnd() * 256)
 MyForm.BackColor = RGB(r, g, b)
End Sub
 
Private Sub Cmd_fgColor_Click()
 Dim r, g, b As Integer
 r = Int(Rnd() * 256)
 g = Int(Rnd() * 256)
 b = Int(Rnd() * 256)
 Lbl_Msg.ForeColor = RGB(r, g, b)
 
End Sub
àCompObj [1]ÿÿÿÿÿÿÿÿÿÿÿÿpCompObj}Üsµ7,o$Õþÿ 
________________________________________

ÿÿÿÿ  n`ôÎ ›Íª`ŽMicrosoft Forms 2.0 FrameEmbedded Object
Forms.Frame.1ô9²q[1]4AE€H€, 
________________________________________
[1] €Ý&ö 
________________________________________
formatting my colour [1]5 €¥[1]Tahoma[1]$( €change back colorur5%ö 
________________________________________
[1]u €¥[1] 
________________________________________
Tahoma[1](( €forgetting my collor 5%É[1]u €¥[1] 
________________________________________
Tahoma 
________________________________________
8H


________________________________________
Rã
‘ Î ãªK¸Q DB Tahoma 
________________________________________
ƒo(å      €[1]TComboBox1§§,å
€ 
________________________________________
D CommandButton1{[1]      ,å
€ 
________________________________________
H[1] CommandButton2{[1]¶ €# 
________________________________________

Frame1â
qþÿ 
________________________________________

ÿÿÿÿði*ÆÜÎ ž˜ªWJOMicrosoft Forms 2.0 FormEmbedded Object
Forms.Form.1ô9²q
VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   8085
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   16815
   OleObjectBlob   =   "UserForm document office.frx":0000
   StartUpPosition =   3  'Windows Default
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ComboBox1_Change()
 
End Sub
 
Private Sub CommandButton1_Click()
 
End Sub
 
Private Sub CommandButton2_Click()
 
End Sub
 
Private Sub Frame1_Click()
 
End Sub
 
Private Sub UserForm_Click()
 
End Sub
 
Private Sub UserForm_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
 
End Sub
 
Private Sub UserForm_Error(ByVal Number As Integer, ByVal Description As MSForms.ReturnString, ByVal SCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, ByVal CancelDisplay As MSForms.ReturnBoolean)
 
End Sub
 
Private Sub UserForm_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
 
End Sub
 
Private Sub UserForm_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
 
End Sub
 
Private Sub UserForm_RemoveControl(ByVal Control As MSForms.Control)
 
End Sub
 
Private Sub UserForm_Resize()
 
End Sub
 
Private Sub UserForm_Scroll(ByVal ActionX As MSForms.fmScrollAction, ByVal ActionY As MSForms.fmScrollAction, ByVal RequestDx As Single, ByVal RequestDy As Single, ByVal ActualDx As MSForms.ReturnSingle, ByVal ActualDy As MSForms.ReturnSingle)
 
End Sub
 
Private Sub UserForm_Terminate()
 
End Sub
 
Private Sub UserForm_Zoom(Percent As Integer)
Private Sub cmd_bgColor_Click()
 Dim r, g, b As Integer
 r = Int(Rnd() * 256)
 g = Int(Rnd() * 256)
 b = Int(Rnd() * 256)
 MyForm.BackColor = RGB(r, g, b)
End Sub
 
Private Sub Cmd_fgColor_Click()
 Dim r, g, b As Integer
 r = Int(Rnd() * 256)
 g = Int(Rnd() * 256)
 b = Int(Rnd() * 256)
 Lbl_Msg.ForeColor = RGB(r, g, b)
 
End Sub
VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "UserForm2"
   ClientHeight    =   8040
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   16845
   OleObjectBlob   =   "UserForm2 document wallet.frx":0000
   StartUpPosition =   1  'CenterOwner
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Click()
 
End Sub
Private Sub Form_Activate()
Print 20 + 10
Print 20 - 10
Print 20 * 10
Print 20 / 10 
LB 
'B
!ÐÏ à¡±á> [1]ði*ÆÜÎ ž˜ªWJOàrß”ïÙ 
________________________________________
Àf 
________________________________________
[1]ÿÿÿÿÿÿÿÿÿÿÿÿ.o 
________________________________________
[1] 
________________________________________
ÿÿÿÿþÿÿÿCompObj
VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm3 
   Caption         =   "UserForm3"
   ClientHeight    =   6210
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15855
   OleObjectBlob   =   "UserForm3 document walet.frx":0000
   StartUpPosition =   3  'Windows Default
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
End
Attribute VB_Name = "UserForm3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserFor()
 
End Sub
A = "Tom"
b = "likes"
C = "to"
D = "eat"
E = "burger"
Print A + b + C + D + E
  
LB 
I>æÐÏ à¡±á> 
________________________________________
________________________________________

________________________________________
 ‚

€[1] 
________________________________________
}>mÊ*þÿ 
________________________________________

ÿÿÿÿði*ÆÜÎ ž˜ªWJOMicrosoft Forms 2.0 FormEmbedded Object
Forms.Form.1ô9²q
VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "UserForm2"
   ClientHeight    =   8040
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   16845
   OleObjectBlob   =   "UserForm2 document wallet.frx":0000
   StartUpPosition =   1  'CenterOwner
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Click()
 
End Sub
Private Sub Form_Activate()
Print 20 + 10
Print 20 - 10
Print 20 * 10
Print 20 / 10 
LB 
'B
!ÐÏ à¡±á> 
Object
Forms.Frame.1ô9²q[1]4AE€H€, 
________________________________________
[1] €Ý&ö 
________________________________________
formatting my colour [1]5 €¥[1]Tahoma[1]$( €change back colorur5%ö 
________________________________________
[1]u €¥[1] 
________________________________________
Tahoma[1](( €forgetting my collor 5%É[1]u €¥[1] 
________________________________________
Tahoma 
________________________________________
8H


________________________________________
________________________________________
€ 
________________________________________
€ÿÿ 
________________________________________
}ÿ*úFrame1 
________________________________________
Rã
‘ Î ãªK¸Q DB Tahoma 
________________________________________
ƒo(å      €[1]TComboBox1§§,å
€ 
________________________________________
D CommandButton1{[1]      ,å
€ 
________________________________________
H[1] CommandButton2{[1]¶ €# 
________________________________________

Frame1â
qþÿ 
________________________________________

ÿÿÿÿði*ÆÜÎ ž˜ªWJOMicrosoft Forms 2.0 FormEmbedded Object
Forms.Form.1ô9²q

On Mon, Sep 25, 2023 at 11:57 AM tshingombe fiston <tshingombefiston@gmail.com> wrote:
VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   8085
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   16815
   OleObjectBlob   =   "UserForm document office.frx":0000
   StartUpPosition =   3  'Windows Default
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ComboBox1_Change()
 
End Sub
 
Private Sub CommandButton1_Click()
 
End Sub
 
Private Sub CommandButton2_Click()
 
End Sub
 
Private Sub Frame1_Click()
 
End Sub
 
Private Sub UserForm_Click()
 
End Sub
 
Private Sub UserForm_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
 
End Sub
 
Private Sub UserForm_Error(ByVal Number As Integer, ByVal Description As MSForms.ReturnString, ByVal SCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, ByVal CancelDisplay As MSForms.ReturnBoolean)
 
End Sub
 
Private Sub UserForm_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
 
End Sub
 
Private Sub UserForm_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
 
End Sub
 
Private Sub UserForm_RemoveControl(ByVal Control As MSForms.Control)
 
End Sub
 
Private Sub UserForm_Resize()
 
End Sub
 
Private Sub UserForm_Scroll(ByVal ActionX As MSForms.fmScrollAction, ByVal ActionY As MSForms.fmScrollAction, ByVal RequestDx As Single, ByVal RequestDy As Single, ByVal ActualDx As MSForms.ReturnSingle, ByVal ActualDy As MSForms.ReturnSingle)
 
End Sub
 
Private Sub UserForm_Terminate()
 
End Sub
 
Private Sub UserForm_Zoom(Percent As Integer)
Private Sub cmd_bgColor_Click()
 Dim r, g, b As Integer
 r = Int(Rnd() * 256)
 g = Int(Rnd() * 256)
 b = Int(Rnd()MyForm.BackColor = RGB(r, g, b)
End Sub
 
Private Sub Cmd_fgColor_Click()
 Dim r, g, b As Integer
 r = Int(Rnd() * 256)
 g = Int(Rnd() * 256)
 b = Int(Rnd() * 256)
 Lbl_Msg.ForeColor = RGB(r, g, b)
 
End Sub
VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm4 
   Caption         =   "UserForm4"
   ClientHeight    =   7125
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13845
   OleObjectBlob   =   "UserForm4 document walet.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ComboBox1_Change()
 
End Sub
 
 
Private Sub ComboBox2_Change()
 
End Sub
 
Private Sub ComboBox3_Change()
 
End Sub
 
Private Sub CommandButton1_Click()
 
End Sub
 
Private Sub Label1_Click()
 
End Sub
 
Private Sub Label2_Click()
 
End Sub
 
Private Sub Label3_Click()
 
End Sub
 
Private Sub UserForm_Click()
 
End Sub
 
Private Sub UserForm_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
 
End Sub
 
Private Sub UserForm_Error(ByVal Number As Integer, ByVal Description As MSForms.ReturnString, ByVal SCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, ByVal CancelDisplay As MSForms.ReturnBoolean)
 
End Sub
 
Private Sub UserForm_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
 
End Sub
 
Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
 
End Sub
 
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
 
End Sub
 
Private Sub UserForm_RemoveControl(ByVal Control As MSForms.Control)
 
End Sub
 
Private Sub UserForm_Scroll(ByVal ActionX As MSForms.fmScrollAction, ByVal ActionY As MSForms.fmScrollAction, ByVal RequestDx As Single, ByVal RequestDy As Single, ByVal ActualDx As MSForms.ReturnSingle, ByVal ActualDy As MSForms.ReturnSingle)
 
End Sub
 
Private Sub UserForm_Terminate()
 
End Sub
 
Private Sub UserForm_Zoom(Percent As Integer)
 
  
LB 
o6yÐÏ à¡±á> 
VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm5 
   Caption         =   "UserForm5"
   ClientHeight    =   6690
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   16170
   OleObjectBlob   =   "UserForm5 fortofolio walet.frx":0000
   StartUpPosition =   1  'CenterOwner
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
End
Attribute VB_Name = "UserForm5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
Private Sub ComboBox1_Change()
 
End Sub
 
Private Sub CommandButton1_Click()
 
End Sub
 
Private Sub CommandButton2_Click()
 
End Sub
 
Private Sub Frame1_Click()
 
End Sub
 
Private Sub ScrollBar1_Change()
 
End Sub
 
Private Sub ScrollBar2_Change()
 
End Sub
 
Private Sub UserForm_Click()
 
End Sub
 
Private Sub UserForm_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
 
End Sub
 
Private Sub UserForm_Error(ByVal Number As Integer, ByVal Description As MSForms.ReturnString, ByVal SCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, ByVal CancelDisplay As MSForms.ReturnBoolean)
 
End Sub
 
Private Sub UserForm_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
 
End Sub
 
Private Sub UserForm_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
 
End Sub
 
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
 
End Sub
 
Private Sub UserForm_RemoveControl(ByVal Control As MSForms.Control)
 
End Sub
 
Private Sub UserForm_Resize()
 
End Sub
 
Private Sub UserForm_Scroll(ByVal ActionX As MSForms.fmScrollAction, ByVal ActionY As MSForms.fmScrollAction, ByVal RequestDx As Single, ByVal RequestDy As Single, ByVal ActualDx As MSForms.ReturnSingle, ByVal ActualDy As MSForms.ReturnSingle)
 
End Sub
 
Private Sub UserForm_Terminate()
 
End Sub
 
Private Sub UserForm_Zoom(Percent As Integer)
 
End Sub 
LB 
________________________________________
„?ÆÐÏ à¡±á> 
________________________________________ TahomaRoot Entryÿÿÿÿÿÿÿÿ[1]ði*ÆÜÎ ž˜ªWJO €ÂO–ïÙ À 
________________________________________
f 
________________________________________
[1]ÿÿÿÿÿÿÿÿÿÿÿÿŽo 
________________________________________
[1] 
________________________________________
ÿÿÿÿTi03 ÿÿÿÿ  n`ôÎ ›Íª`ŽÀ³¼O–ïÙà½O–ïÙÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿ ýÿÿÿþÿÿÿþÿÿÿ‰ýÿÿÿýÿÿÿýÿÿÿ

 ¬ !"#$%&'()*+,-./0123456789:;<=>?@ABCDEFGHIJKLMNOPQRSTUVWXYZ[\]^_`abcdefghijklmnopqrstuvwxyz{|}~ €f 
________________________________________ VB6 Made Easy Book

 

Check Out Our Book

________________________________________

VB6 Made Easy Paperback

<br />

VB6 Made Easy Kindle

VB6 Google Book

 

Available on Google Play Store

________________________________________

Form1.show

Print "Welcome to Visual Basic tutorial"

End Sub

Example 2.1.2

Private Sub Form_Activate ( )

Print 20 + 10

Print 20 - 10

Print 20 * 10

Print 20 / 10

End Sub

 Figure 2.4 : The output of Example 2.1.1 

 Figure 2.4 : The output of Example 2.1.2 

You can also use the + or the & operator to join two or more texts (string) together like in example 2.1.4 (a) and (b)

Example 2.1.4(a)

Private Sub 

A = "Tom" 

B = "likes" 

C = "to" 

D = "eat" 

E = "burger" 

Print A + B + C + D + E

End Sub

Example 2.1.4(b)

Private Sub

A = "Tom" 

B = "likes" 

C = "to" 

D = "eat" 

E = "burger" 

Print A & B & C & D & E

"    Form1-MyForm 

"    Label1-LblMessage 

"    Command1-cmd_bgColor 

"    Command2-cmd_fgColor 

Next, change the caption of the Label to "Please Change My Color". In addition, change the caption of Command1 button to "Change Background Color" and change the caption of Command2 button to "Change Foreground Color"

Now, enter the following code

Private Sub cmd_bgColor_Click()

 Dim r, g, b As Integer

 r = Int(Rnd() * 256)

 g = Int(Rnd() * 256)

 b = Int(Rnd() * 256)

 MyForm.BackColor = RGB(r, g, b)

End Sub


Private Sub Cmd_fgColor_Click()

 Dim r, g, b As Integer

 r = Int(Rnd() * 256)

 g = Int(Rnd() * 256)

 b = Int(Rnd() * 256)

 Lbl_Msg.ForeColor = RGB(r, g, b)

End Sub

When you run the program, each time you press on the 'Change Background Color' button, you will see different background color. Similarly, each time you press on the 'Change Foreground Color', you will see the message on the Label changes color. The output is shown in Figure 2.8.

 Figure 2.8 



 118 

 13 

 49 

 

 55 


Copyright©2008 Dr.Liew Voon Kiong. All rights reserved |Contact|Privacy Policy

Last update:09/05/2023 02:40:29


J276 NEA Resource Bank 

"    Key Syntax 

"    Program Examples 

"    Evidencing 

"    Algorithms 

"    Testing 

"    Debugging 

"    Rules 

"    Labels - lbl

"    Text Boxes - txt

"    Combo Boxes - cmb

"    List Boxes - lst

"    Buttons - btn


MessageBox.Show("Address Details" & vbNewLine & "Street: " & txtNumber.Text & " " &

txtStreet.Text & vbNewLine & "Town/City: " & txtTown.Text & vbNewLine &

"County: " & txtCounty.Text & vbNewLine & "Postcode: " & txtPostcode.Text)

This is what happens when the button is clicked:

 

You can concatenate (join together) controls with strings in a MessageBox.Show() command. In the address example MessageBox.Show("Street: " & txtnumber.Text & " " & txtStreet.Text & vbNewLine & "Town/City: " + txtTown.Text) will combine the strings Street  and Town/City  with the form controls txtNumber, txtStreet and txtTown.

vbNewLine is used to start a new line when it is displayed on screen.

Outputting Data - List Boxes

One of the other methods of outputting data is through a list box.

To add data to a list box you need to use the following code:

Listboxname.Items.Add(Data to add to list box)

Consider the following interface:

 

To add the name entered and the email address entered to the list box when the button is clicked you need the following code:

Code when btnAdd clicked

lstOutput.Items.Add(txtName.Text & vbTab & txtEmail.Text)

This is what happens when the button is clicked:

 

You can add multiple rows of data to a list box. You can simply change the information in the text boxes and click the button again. vbTab is used to put a tab space between the two pieces of data.

Sometimes when you use list boxes it makes it difficult to make the formatting look neat, this is shown in the screenshot above where the data is not lined up. To overcome this problem you can use a list view.

With a list view box you can add columns and headings to make it look like a table, like the example shown below:

 

When you add a list view control, you have to change a property to get it to work. Change the view from Large Icon to Details, like shown below:

 

The code for when the button is clicked needs to follow the following format:

LISTVIEWBOXNAME.Items.Add(New ListViewItem({COLUMN1, COLUMN2, COLUMN3, COLUMN4, COLUMN5}))

For this example, the code would be:

lstOutput.Items.Add(New ListViewItem({txtName.Text, txtEmail.Text}))

This is because there are only two pieces of information, therefore you only need to fill in the information you want in two of the columns.

As well as adding code for when the button is pressed, you also need to add the column headings when the form loads. You should double click on the form itself and add the code in this subroutine. The format of the code for the column headings is:

LISTVIEWBOXNAME.Columns.Add( Text , Size , Alignment)

In this example as there are two headings the code would be:

Private Sub Form1_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

    lstOutput.Columns.Add("Name", 150, HorizontalAlignment.Left)

    lstOutput.Columns.Add("Email", 250, HorizontalAlignment.Left)

End Sub

Example 3 - Test Scores - Using List View

Private Sub btnAdd_Click(sender As System.Object, e As System.EventArgs) Handles btnAdd.Click

    lstOutput.Items.Add(New ListViewItem({txtName.Text, txtScore1.Text, txtScore2.Text, txtScore3.Text}))

End Sub

Private Sub Form1_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

    lstOutput.Columns.Add("Name", 150, HorizontalAlignment.Left)

    lstOutput.Columns.Add("Score 1", 75, HorizontalAlignment.Center)

    lstOutput.Columns.Add("Score 2", 75, HorizontalAlignment.Center)

    lstOutput.Columns.Add("Score 3", 75, HorizontalAlignment.Center)

End Sub

This is what happens when the button is clicked:

 

Variables

A variable is used to temporarily store a piece of data.

For example:

Dim number1 As Integer = 10

In the code above the variable is called number1 and the value it is storing is 10. Variables can hold any type of data. Using variables makes it easier for people to understand what is going on. In Visual Basic you need to define a variable before you can use it. To do this you type Dim before the variable name the first time you use it. You should then say what type of data you think it is. In this example number1 is an integer therefore the code to define a variable called number1 as an integer is Dim number1 as Integer

For example:

Dim cost As Decimal = 15.5

Dim VAT As Decimal = 3.1

Dim total_cost As Decimal = cost + VAT

Example program 1 - Water Tank Capacity Program

The code for the program below will allow the user to enter the height, width and depth of a water tank, then calculate and output the capacity.

Interface

 

Code when btnCapacity is clicked

'three variables that store the text box inputs from the interface as a decimal

Dim height As Decimal = txtHeight.Text

Dim width As Decimal = txtWidth.Text

Dim depth As Decimal = txtDepth.Text

'calculation to work out the capacity

Dim capacity As Decimal = (height * width * depth) / 1000

'outputs the capacity of the water tank

MessageBox.Show("The tank holds " & Decimal.Round(capacity, 2).ToString & " litres of water")

This is what happens when the button is clicked:

 

The code above rounds the variable capacity, to round a variable you use the Decimal.Round() function. You write the name of the variable followed by the number of decimal places e.g. Decimal.Round(capacity,2). Also note that it has .ToString after it, this is because any variable that is not a string data type much be converted back to a string before it can be displayed in a message or list box.

Example program 2 - Cylinder Volume Program

The code for the program below will allow the radius and height of a circle, then calculate and output the volume and surface area.

Interface

 

Code when btnCalculate is clicked

'three variables that store the two inputs from the interface and the value of pie

Dim radius As Decimal = txtRadius.Text

Dim height As Decimal = txtHeight.Text

Dim pie As Decimal = 3.14159

'calculations to work out the volume and surface area

Dim volume As Decimal = pie * (radius * radius) * height

Dim surfaceArea As Decimal= 2 * (pie * (radius * radius)) + 2 * (pie * radius * height)

#outputs the volume and surface area of the cylinder in a message box.

MessageBox.Show("The volume of your cylinder is: " & Decimal.Round(volume, 2).ToString & " to 2 decimal places" & vbNewLine & "The surface area of the cylinder is: " & Decimal.Round(surfaceArea, 2).ToString & " to 2 decimal places.")

This is what happens when the button is clicked:

 

Selection (if, then, else)

Sometimes you will change what do you depending on the conditions.

For example: IF you wake up in the morning and it is raining THEN you will take a coat to school OTHERWISE you wont.

IF the day is a Saturday AND the alarm clock goes off THEN you might turn it off and stay in bed OTHERWISE you might get up.

Life is full of decisions that you will make depending on certain conditions, computers are no different.

if-else

For a computer to make decisions based on a condition, you must use an IF statement, it has the following structure:

If condition Then

     true

     several instructions that are executed

     if the calcualation evaluates to True

Else

     false

     several instructions that are exectued

     if the condition evaluates to False

End If

Consider the following IF statement:

Dim age As Integer = txtAge.Text

If age >= 18 Then

     MessageBox.Show("You are an adult")

Else

     MessageBox.Show("You are still a child")

End If

The IF statement explained:

"    after the if is the condition age >= 18, this is checking to see if the age variable is more than or equal to 18.

"    after that line is code is the code that will only be run if that condition is True. If it is true it will display a message box that says You are an adult.

"    the word else then follows. The instructions underneath this are what will be run only if that condition is False. If it is false it will display a message box that says You are still a child.

if-elseif-else

An IF statement with an else will only allow you to check a single condition, however if you have more than one condition to check you can use if..elseif..else

Consider the following IF statement:

Dim colour As String = cmbColour.Text

If colour = "Red" Then

     MessageBox.Show("STOP")

Elseif colour = "Amber" Then

     MessageBox.Show("GET READY TO STOP")

Else

     MessageBox.Show("GO")

End If

The IF statement explained:

"    the program first checks to see if the colour selected in the combo box on the interface is Red and if it is will display a message box saying STOP.

"    if the colour selected isn t red it will go onto the next condition where the Elseif is and check if the colour is Amber. If it is then it will display a message box saying GET READY TO STOP

"    if neither conditions are met it will go to
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm3 
   Caption         =   "UserForm3"
   ClientHeight    =   4995
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13920
   OleObjectBlob   =   "UserForm peace dhet poe.frx":0000
   StartUpPosition =   1  'CenterOwner
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
End
Attribute VB_Name = "UserForm3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
Private Sub ComboBox1_Change()
 
End Sub
 
Private Sub ComboBox2_Change()
 
End Sub
 
Private Sub CommandButton1_Click()
 
End Sub
 
Private Sub Frame1_Click()
 
End Sub
 
Private Sub Frame2_Click()
 
End Sub
 
Private Sub Label1_Click()
 
End Sub
 
Private Sub Label2_Click()
 
End Sub
 
Private Sub Label3_Click()
 
End Sub
 
Private Sub UserForm_Click()
 
End Sub
 
Private Sub UserForm_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
 
End Sub
 
Private Sub UserForm_Deactivate()
 
End Sub
 
Private Sub UserForm_Initialize()
 
End Sub
 
Private Sub UserForm_Layout()
 
End Sub
 
Private Sub UserForm_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
 
End Sub
 
Private Sub UserForm_RemoveControl(ByVal Control As MSForms.Control)
 
End Sub
 
Private Sub UserForm_Resize()
 
End Sub
 
Private Sub UserForm_Scroll(ByVal ActionX As MSForms.fmScrollAction, ByVal ActionY As MSForms.fmScrollAction, ByVal RequestDx As Single, ByVal RequestDy As Single, ByVal ActualDx As MSForms.ReturnSingle, ByVal ActualDy As MSForms.ReturnSingle)
 
End Sub
 
Private Sub UserForm_Terminate()
 
End Sub
 
Private Sub UserForm_Zoom(Percent As Integer)
ÿÿ}VT Frame1 
________________________________________
Rã
‘ Î ãªK¸Q DB Tahoma[1]X‚e(õ €[1]2X Label1{[1]Ô$Õ € 
________________________________________
# 
________________________________________

Frame2§qf 
________________________________________
[1]ÿÿÿÿÿÿÿÿÿÿÿÿ Mo 
________________________________________
[1] 
ÿÿÿÿ
ÄCompObj [1]ÿÿÿÿÿÿÿÿÿÿÿÿ pCompObj [1]ÿÿÿÿÿÿÿÿÿÿÿÿnþÿ 
________________________________________

ÿÿÿÿ  n`ôÎ ›Íª`ŽMicrosoft Forms 2.0 FrameEmbedded Object
Forms.Frame.1ô9²q[1]4(#€marks  award statement certificate pRA" 
________________________________________
[1]5 €¥[1]Tahoma[1]PAE€H€, 
________________________________________
[1]2€°'Ü 
________________________________________
''txt label 1 marks award statement certificate'' [1]5 €¥[1]Tahoma[1]4(!€marks award  diplomat certificat e p«?" 
________________________________________
[1]5 €¥[1]Tahoma[1]HAE€H€, 
________________________________________
[1])€°'Ü 
________________________________________
''txt label 2 award diploma cerificate ''s[1]5€¥[1]Tahoma[1]$l€
€   job up date qÿ[1][1]u €¥[1] 
________________________________________
Tahoma 
________________________________________
8H


________________________________________
€ 
________________________________________
€ÿÿ}4P] Frame2 
________________________________________
Rã
‘ Î ãªK¸Q DB Tahomaè… (õ € 
________________________________________
2T Label2§Ô(å      €pComboBox1¢(õ € 2T[1] Label3§Ê(å      €h 
________________________________________
ComboBox2¢0õ
€ ?D 
________________________________________
CommandButton1 VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   7560
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   16815
   OleObjectBlob   =   "UserForm peace dhet.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Frame1_Click()
 
End Sub
 
Private Sub Frame2_Click()
 
End Sub
 
Private Sub Label1_Click()
 
End Sub
 
Private Sub Label3_Click()
 
End Sub
 
Private Sub Label4_Click()
 
End Sub
 
Private Sub Label5_Click()
 
End Sub
 
Private Sub ListBox1_Click()
 
End Sub
 
Private Sub ScrollBar1_Change()
 
End Sub
 
Private Sub ScrollBar2_Change()
 
End Sub
 
Private Sub ScrollBar3_Change()
 
End Sub
 
Private Sub TabStrip1_Change()
 
End Sub
 
Private Sub UserForm_Click()
 
End Sub
 
Private Sub UserForm_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
 
End Sub
 
Private Sub UserForm_Error(ByVal Number As Integer, ByVal Description As MSForms.ReturnString, ByVal SCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, ByVal CancelDisplay As MSForms.ReturnBoolean)
 
End Sub
 
Private Sub UserForm_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
 
End Sub
 
Private Sub UserForm_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
 
End Sub
 
Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
 
End Sub
 
Private Sub UserForm_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
 
End Sub
 
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
 
End Sub
 
Private Sub UserForm_RemoveControl(ByVal Control As MSForms.Control)
 
End Sub
 
Private Sub UserForm_Resize()
 
End Sub
 
Private Sub UserForm_Scroll(ByVal ActionX As MSForms.fmScrollAction, ByVal ActionY As MSForms.fmScrollAction, ByVal RequestDx As Single, ByVal RequestDy As Single, ByVal ActualDx As MSForms.ReturnSingle, ByVal ActualDy As MSForms.ReturnSingle)
 
End Sub
 
Private Sub UserForm_Terminate()
 
ÿÿÿÿ  n`ôÎ ›Íª`ŽMicrosoft Forms 2.0 FrameEmbedded Object
Forms.Frame.1ô9²q 
________________________________________
@JŸ
€ 
________________________________________
€   [1] 
€ €ÿÿ 
________________________________________
________________________________________
} az-Frame1 
________________________________________
Rã
‘ Î ãªK¸Q DB TahomaÈ[1] €(õ €[1]2D Label1x2" 
________________________________________
O 
________________________________________
(å      € 
________________________________________
PComboBox1me.ú" 
________________________________________
(õ € 
________________________________________
2D[1] Label2x2O 
________________________________________
g
(å  €\ 
________________________________________
ComboBox2me.úg
,å
€ 8 
________________________________________
Commandf 
________________________________________
[1]ÿÿÿÿÿÿÿÿÿÿÿÿ=o 
________________________________________
[1] 
ÿÿÿÿ 4i20 ÿÿÿÿ
p ãFz?Î ¾Öªa€ dwNôÙ dwNôÙCompObj [1]ÿÿÿÿÿÿÿÿÿÿÿÿp[1]$( €student user nameNIs
9q[1]5 €¥[1]Tahoma[1]0AE€H€, 
________________________________________
[1] €º"ö 
________________________________________
''txt label1 name''s[1]5 €¥[1]Tahoma[1]$( € student password'se7E [1]5 €¥[1]Tahoma[1]<AE€H€, 
________________________________________
[1]¬€º" ''txt student label2 password''[1]5 €¥[1]Tahoma[1](€up date‰" 
________________________________________
[1]u €¥[1] 
________________________________________
Tahomae[1](€cancell 
________________________________________
# 
________________________________________
[1]u €¥[1] 
________________________________________
Tahomal[1] (
€student id e7Ê[1]5 €¥[1]Tahomal[1]@AE€H€, 
________________________________________
[1]!€º"ö 
________________________________________
''txt student id label 3 number "r[1]5 €¥[1]Tahomal[1]
!§[1]
k¬§[1]
k¬¨[1]((€student subject  amount e7ž [1]5 €¥[1]Tahomal[1]@AE€H€, 
________________________________________
[1]$€æ!Ê''txt student subject label 4amount"[1]5 €¥[1]Tahomal[1]@(/€registration: dhet saqa qcto  st peace college € !O 
________________________________________
[1]5 €¥[1]Tahomax2[1]¤ÆAE€H€, 
________________________________________
[1]†Æ? E 




Write


Sign In

   MessageBox.Show("Enter Username !!", "Meera Academy");

}

else if(txtpass.Text=="")

{

   MessageBox.Show("Enter Password !!","Meera Academy");

}

else if (txtuname.Text == "Meera" && txtpass.Text == "123")

{

   this.Hide();

   Form frm2 = new Form2();

   frm2.ShowDialog();

}

else

{

   MessageBox.Show("Invalid Credential", "Meera Academy");

}

}

If username or password incorrect it will show error message like below screen.

 Simple login form in windows application. 

In above example if textbox will be blank then it shows the message like Enter Username . If both username and password fill up and one of will be incorrect then it will shows the message Invalid Credential . If username =
Go to frmLogin.cs code and add System.Data and System.Data.SqlClient namespace. Double click on btn_Submit to create btn_Submit Click event.

frmLogin.cs Code:

using System;

using System.Data;

using System.Windows.Forms;

using System.Data.SqlClient;


namespace LoginApplication

{

    public partial class frmLogin : Form

    {

        public frmLogin()

        {

            InitializeComponent();

        }

        //Connection String

        string cs = @"Data Source=(LocalDB)\v11.0;AttachDbFilename=|DataDirectory|\MyDatabase.mdf;Integrated Security=True;";

        //btn_Submit Click event

        private void button1_Click(object sender, EventArgs e)

        {

            if(txt_UserName.Text=="" || txt_Password.Text=="")

            {

                MessageBox.Show("Please provide UserName and Password");

                return;

            }

            try

            {

                //Create SqlConnection

                SqlConnection con = new SqlConnection(cs);

                SqlCommand cmd = new SqlCommand("Select * from tbl_Login where UserName=@username and Password=@password",con);

                cmd.Parameters.AddWithValue("@username",txt_UserName.Text);

                cmd.Parameters.AddWithValue("@password", txt_Password.Text);

                con.Open();

                SqlDataAdapter adapt = new SqlDataAdapter(cmd);

                DataSet ds = new DataSet();

                adapt.Fill(ds);

                con.Close();

                int count = ds.Tables[0].Rows.Count;

                //If count is equal to 1, than show frmMain form

                if (count == 1)
VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "UserForm2"
   ClientHeight    =   4605
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13245
   OleObjectBlob   =   "UserForm2 congratulation.frx":0000
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label1_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)
 
End Sub
 
Private Sub Label1_BeforeDropOrPaste(ByVal Cancel As MSForms.ReturnBoolean, ByVal Action As MSForms.fmAction, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)
 
End Sub
 
Private Sub Label1_Click()
 
End Sub
 
Private Sub Label1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
 
End Sub
 
Private Sub Label1_Error(ByVal Number As Integer, ByVal Description As MSForms.ReturnString, ByVal SCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, ByVal CancelDisplay As MSForms.ReturnBoolean)
 
End Sub
 
Private Sub Label1_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
 
End Sub
 
Private Sub Label1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
 
End Sub
 
Private Sub Label1_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
 
End Sub
 
Private Sub UserForm_Click()
 
[»¬0ho(õ[1]8((€successful .congratulation registration ùBÊ[1]5 €¥[1]Tahoma €2X Label1;
" 
________________________________________
VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   10530
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15300
   OleObjectBlob   =   "UserForm1 polfin persal peace.frx":0000
   StartUpPosition =   3  'Windows Default
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ComboBox1_Change()
 
End Sub
 
Private Sub ComboBox2_Change()
 
End Sub
 
Private Sub ComboBox3_Change()
 
End Sub
 
Private Sub CommandButton1_Click()
 
End Sub
 
Private Sub Frame1_Click()
 
End Sub
 
Private Sub Label1_Click()
 
End Sub
 
Private Sub TextBox1_Change()
 
End Sub
 
Private Sub UserForm_Click()
 
End Sub
 
Private Sub UserForm_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
 
End Sub
 
Private Sub UserForm_Error(ByVal Number As Integer, ByVal Description As MSForms.ReturnString, ByVal SCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, ByVal CancelDisplay As MSForms.ReturnBoolean)
 
End Sub
 
Private Sub UserForm_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
 
End Sub
 
Private Sub UserForm_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
 
End Sub
 
Private Sub UserForm_RemoveControl(ByVal Control As MSForms.Control)
 
End Sub
 
Private Sub UserForm_Resize()
 
End Sub
 
Private Sub UserForm_Scroll(ByVal ActionX As MSForms.fmScrollAction, ByVal ActionY As MSForms.fmScrollAction, ByVal RequestDx As Single, ByVal RequestDy As Single, ByVal ActualDx As MSForms.ReturnSingle, ByVal ActualDy As MSForms.ReturnSingle)
 
End Sub
 
Private Sub UserForm_Terminate()
 
End Sub
 
Private Sub UserForm_Zoom(Percent As Integer)
 
End Sub
LBæ <Æ*ÐÏà¡±á> 
Pw¬öÙPDQw-öÙÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿýÿÿÿþÿÿÿþÿÿÿñ[1]ýÿÿÿýÿÿÿýÿÿÿýÿÿÿýÿÿÿ ¬ !"#$%&'()*+,-./0123456789:;<=>?@ABCDEFGHIJKLMNOPQRSTUVWXYZ[\]^_`abcdefghijklmnopqrstuvwxyzŽMicrosoft Forms 2.0 FrameEmbedded Object
Forms.Frame.1ô9²q 
________________________________________
8H

      
________________________________________
€ 
________________________________________
€ÿÿ}ïGœ1Frame1 
________________________________________
Rã
‘ÎãªK¸QDBTahomadˆi(õ€[1]2PLabel1O 
________________________________________
" 
________________________________________
(õ€ 
________________________________________
2TLabel2O 
________________________________________
g
$å€ 
________________________________________
4[1]TextBox1O 
________________________________________
è(å      €8 
________________________________________
ComboBox1Sö 
________________________________________
(å     €8 
________________________________________
ComboBox2'gTSHINGOMBEKB

IES

NCES

National Center for

Education Statistics

menu


    Surveys & Programs

        Annual Reports

            Annual Reports

            Condition of Education Digest of Education Statistics Projections of Education Statistics Topical Studies

        National Assessments

            National Assessments

            National Assessment of Educational Progress (NAEP) Program for the International Assessment of Adult Competencies (PIAAC)

        International Assessments

            International Assessments

            International Activities Program (IAP)

        Early Childhood

            Early Childhood

            Early Childhood Longitudinal Study (ECLS) National Household Education Survey (NHES)

        Elementary/ Secondary

            Elementary/ Secondary

            Common Core of Data (CCD) Secondary Longitudinal Studies Program Education Demographic and Geographic Estimates (EDGE) National Teacher and Principal Survey (NTPS) more...

        Library

            Library

            Library Statistics Program

        Postsecondary

            Postsecondary

            Baccalaureate and Beyond (B&B) Career/Technical Education Statistics (CTES) Integrated Postsecondary Education Data System (IPEDS) National Postsecondary Student Aid Study (NPSAS) more...

        Data Systems, Use, & Privacy

            Data Systems, Use, & Privacy

            Common Education Data Standards (CEDS) National Forum on Education Statistics Statewide Longitudinal Data Systems Grant Program - (SLDS) more...

        resources

            resources

            Distance Learning Dataset Training National Postsecondary Education Cooperative (NPEC) Statistical Standards Program more...

    Data & Tools

        Downloads Microdata/Raw Data

            Downloads Microdata/Raw Data

            Delta Cost Project IPEDS Data Center How to apply for Restricted Use License Online Codebook

        Online Analysis

            Online Analysis

            ACS-ED Tables Data Lab Elementary Secondary Information System International Data Explorer IPEDS Data Center NAEP Data Explorer

        School and College Search

            School and College Search

            ACS Dashboard College Navigator Private Schools Public School Districts Public Schools Search for Schools and Colleges

        Comparison Tools

            Comparison Tools

            NAEP State Profiles (nationsreportcard.gov) Public School District Finance Peer Search Education Finance Statistics Center IPEDS Data Center

        Questionnaire Tools

            Questionnaire Tools

            NAEP Question Tool NAAL Questions Tool

        Geographic Tools

            Geographic Tools

            ACS-ED Dashboard ACS-ED Maps Locale Lookup MapEd SAFEMap School and District Navigator

        Other Tools

            Other Tools

            Bibliography ED Data Inventory

    Fast Facts

        Fast Facts

        Assessments Early Childhood Elementary and Secondary Postsecondary and Beyond Resources Special Topics

    News & Events

        News & Events

        NCES Blog What's New at NCES Conferences/Training NewsFlash Funding Opportunities Press Releases StatChat

    Publications & Products

        Publications & Products

        Search Publications and Products Annual Reports Restricted-use Data Licenses

        Recent Publications By Subject Index A-Z By Survey & Program Areas Data Products Last 6 Months

    About Us

        About Us

        About NCES Commissioner Contact NCES Staff Help


Contact

Tools

ed.gov

Newsflash

NCES blog

Twitter logo Twitter

Facebook logo Facebook

KidsZone Logo

Chapter 1   Chapter 2   Chapter 3   Chapter 4   Chapter 5   Chapter 6   Chapter 7   Chapter 8   Chapter 9   Chapter 10

Table of Contents   Glossary of Terms

        chapter 4

security Management

Illustration of the Cover of Safeguarding Your Technology

Chapter 4 in a Nutshell:


Introduction to Security Management

Commonly Asked Questions

Nurturing Support within the Organization

Planning for the Unexpected

Testing and Review

Implementation and Day-to-Day Maintenance

Security Management Checklist



Effective security strikes a balance between protection and convenience.

        Introduction to Security Management


Because system security is the aggregate of individual component security, "system boundaries" must encompass individual users and their workstations.  But because personal computers are just that (personal), staff behavior can't always be dictated without potentially hampering workers' overall productivity.  Recall that security policy becomes ineffective if it's so restrictive that legitimate user access is threatened.  Thus, a key to successful security implementation is finding a reasonable balance between system protection and user autonomy and convenience. The person responsible for finding that balance and actively promoting organizational security is the security manager.  Security management consists of nurturing a security-conscious organizational culture, developing tangible procedures to support security, and managing the myriad of pieces that make up the system.  The security manager ensures that administration and staff are aware of their security roles, support security efforts, and are willin 

g to tolerate the minor inconveniences that are inevitably a part of system change and improvement.  After all, if personnel circumvent security procedures (e.g., write down passwords, share accounts, and disable virus-checking software), they put the entire system at risk.

     

Important point.        Effective system security depends on creating a workplace environment and organizational structure where management understands and fully supports security efforts, and users are encouraged to exercise caution.  The security manager leads this effort.

 

A security manager must:


    Communicate to staff that protecting the system is not only in the organization"s interests, but also in the best interest of users.


    Increase staff awareness of security issues.


    Provide for appropriate staff security training.


    Monitor user activity to assess security implementation.



back to topback to home page

     

Commonly Asked Questions        Commonly Asked Questions


Q. Can an organization make do without hiring a security manager?

A. Yes, but while a security manager doesn"t always need to be hired (especially in smaller organizations), someone must perform the functions of security management all the same.  Many organizations prefer to hire a systems administrator and include security management as one of his or her primary duties.  This is an acceptable strategy as long as the administrator has sufficient time to dedicate to security management.  If, however, routine administrative functions take up a considerable part of the administrator"s work day, then the organization will be better served by having someone who is able to focus on system security.


Q. Would assigning a top educational administrator to the security manager role show commitment to system security?

A. Not necessarily.  Although top administrators are often entrusted with sufficient authority to be effective security managers, it is quite possible that they do not possess the technical expertise necessary for the job.  Security managers are responsible for operationalizing all aspects of system security- a task that requires significant technical competence.  A secondary, but important, consideration is that managing system security can demand a great deal of time- time that policy-makers and other top administrators may be unable to devote given their other essential duties.  While it is imperative that top administrators are actively committed to security effectiveness, in most cases it makes sense that the day-to-day administration of system security be assigned to a security/systems professional.


Q. Where does the security manager fit into the organizational hierarchy?

A. Just as the title implies, security managers and system administrators are most often considered to serve in a management capacity.  The important tasks of developing security regulations, training staff, and monitoring implementation require that the security manager be vested with substantial authority.  While the security manager is not to be confused with a superintendent
VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "UserForm2"
   ClientHeight    =   8820
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14055
   OleObjectBlob   =   "UserForm2 engi polfine.frx":0000
   StartUpPosition =   1  'CenterOwner
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ComboBox1_Change()
 
End Sub
 
Private Sub ComboBox2_Change()
 
End Sub
 
Private Sub CommandButton1_Click()
 
End Sub
 
Private Sub Label1_Click()
 
End Sub
 
Private Sub Label2_Click()
 
End Sub
 
Private Sub Label3_Click()
 
End Sub
 
Private Sub UserForm_Click()
 
End Sub
 
Private Sub UserForm_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
 
End Sub
 
Private Sub UserForm_Error(ByVal Number As Integer, ByVal Description As MSForms.ReturnString, ByVal SCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, ByVal CancelDisplay As MSForms.ReturnBoolean)
 
End Sub
 
Private Sub UserForm_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
 
End Sub
 
Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
 
End Sub
 
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
 
End Sub
 
Private Sub UserForm_RemoveControl(ByVal Control As MSForms.Control)
 
End Sub
 
Private Sub UserForm_Resize()
 
End Sub
 
Private Sub UserForm_Scroll(ByVal ActionX As MSForms.fmScrollAction, ByVal ActionY As MSForms.fmScrollAction, ByVal RequestDx As Single, ByVal RequestDy As Single, ByVal ActualDx As MSForms.ReturnSingle, ByVal ActualDy As MSForms.ReturnSingle)
 
End Sub
 
Private Sub UserForm_Terminate() 
LB BA7$ÐÏ à¡±á> 
ÿÿÿÿ  n`ôÎ ›Íª`ŽMicrosoft Forms 2.0 FrameEmbedded Object
Forms.Frame.1ô9²q 
________________________________________
8H

      
________________________________________
€ 
________________________________________
€ÿÿ     }ÛQÓ+Frame1 
________________________________________
Rã
‘ Î ãªK¸Q DB Tahoma<‡c(õ € 
________________________________________
2L Label1 ö 
________________________________________
(õ € 
________________________________________
2H Label2 ;
(õ €2<[1] Label3 (å      € p 
________________________________________
ComboBox17v'ö 
________________________________________
(å      €8 
________________________________________
ComboBox27v'[1],(€engineery elect system datatE [1]5 €¥[1]Tahoma[1]((€electrical poe system HF [1]5 €¥[1]Tahoma[1] (
€electrical GÊ[1]5 €¥[1]Tahoma[1]P AE€H€, 
________________________________________
[1]1 €4Ê
L.M.Photon Power Factor correction, Motor starters and Variable frequency drives for induction motors.
 
50Hz Motors on 60Hz and 60Hz Motors on 50Hz
. You can run a 50Hz motor on 60Hz, and a 60Hz motor on 50 Hz provided you adjust the voltage and power ratings to keep the V/Hz constant.
 
Busbar Calculations
 Calculation of ratings for both Aluminium busbars and Copper busbars.
 
Brushless DC Motors
 An introduction to Brushless DC motors.
 
Compound DC Motors
 An introduction to Compound DC motors.
 
newDahlander Motors
 Dual speed Dahlander motors.
 
Electrical Calculations Software  Busbar ratings, Cable ratings, Power Factor Correction calculations, Motor Starting and acceleration Curves, Enclosure Cooling, Genset sizing and numerous metric/imperial conversions.
 
Enclosure Ventilation  Guidelines for the correct cooling ventilation of switchboards and enclosures.
 
Energy Savers for Induction Motors  Energy Saving systems for Induction motors   Are they a sham?   Do they work?
There seems to be a resurgence in interest in the Nola energy saving algorithm for induction motors, with a number of manufacturers beginning to market "new" and "improved" versions of this technology.
 
Genset Sizing  An introductory paper on the sizing of engines and alternators when used for motor starting supplies.
 
GSM Alarm  GSM Alarm relays can be used to provide alarm text messages from eight channel monitoring and single channel control via text messaging.
 
GSM Control  GSM Control relays can be used to provide basic two channel monitoring and control via text messaging.
 
Harmonic Filters  Harmonic filters are used to reduce the harmonics generated by the input rectifier of VFDs.
 
Induction Motor Control  A Paper on induction motor design covering many aspects of motor design and classification. Motor starting characteristics are explained. Worth a read.
 
Industrial IO  An introductory paper on industrial IO as applied to PLCs, and other electronic devices. Currently covering outputs only, Inputs to follow shortly.
 
Installing VFDs on Irrigation Pumps for minimum EMC  Suggested codes of practice for installing Variable Frequency Drives on irrigation pumps with minimum EMC problems.
 
Invertek Drives  Now available in New Zealand.
 
      
 
Induction Motor Calculations Formula for basic induction motor calculations.
 
Induction Motor Cogging and Crawling Induction motors will not accelerate to full speed if they either cog or crawl. This page covers the basic theory.
 
Logic relays Programmable Logic relays are an easy way to create a flexible and low cost control system using minimum components.
 
Motor Control Forum .  Join the Motor control on line forum to discuss motor control technology and problems with like minded experts. Post and answer questions on any aspect of motor control, starters, soft starters, variable speed drives, protection and design. NB this is independant of the email forum above.
Topics currently covered:
- Soft starters
- Variable Speed drives
- Power Factor Correction
- Motor Starting
- Motor Protection
- Energy Saving
 
Motor Control Mailing List  Join the Motor control email forum to discuss motor control technology and problems with like minded experts. Post and answer questions on any aspect of motor control, starters, soft starters, variable speed drives, protection and design.
Click here to join ElecMotorControl
Click to join ElecMotorControl
 
Motor Starters   A Paper on traditional Electromechanical starters including Direct On Line, Primary resistance, Autotransformer and star delta (wye delta) for induction motors illustrating how to use them and what characteristics can be expected.
 
NO BS Guide to VFDs and EDM
 
NO BS Guide to VFDs and EMC
 
PID Control  An introductory paper on PID control.
 
Horner PLC  Horner OCS are industrial PLCs with an integrated graphical HMI panel.
 
Power Factor Correction  Power factor correction of A.C. Induction motors is often poorly understood and more poorly specified.
 
Power Factor Calculations  Power factor correction Calculations
 
Power Factor Controllers  Power factor correction Controllers for bulk displacement power factor correction.
 
Power Factor Correction for Domestic installations Power factor Correction does not save energy within domestic installations.
 
Pressure Transducers   Pressure transducers are used to measure the pressure/vacuum of liquids and gasses and convert this to an electrical signal which is typically 4-20mA or 0-5V. Pressure transducers are commonly used in pump control systems in conjunction with a VFD to regulate the pressure, but can also be used in conjunction with appropriate logic to provide under and over pressure protection.
 
Secure Password Store  Program for saving encrypted passwords and registration data. Includes a free random password generator.
 
 
      
 
Series DC Motors  An introductory paper on series DC Motors which comprise an armature winding in series with the field windings.
 
Schrage Motors.
Schrage Motors are variable speed motors.
 
Shunt DC Motors
 An introduction to Shunt DC motors.
 
Single Phase Motors  Information and circuits of common single phase induction motor configurations including capacitor start and run, and induction start motors.
 
Smart relays, or Logic relays,   are essentially a simplified PLC with limited functionality. They are designed to replace a number of standard electromechanical and electronic relays as found in machine automation where the task does not require a full PLC.
 
Slip Ring Motors  An introductory paper on slip ring motors and secondary resistance starters, their control and use.
 
Soft Starters .  An introductory paper on Solid State Soft Starters. - under construction!
 
new Space Vector Modulation  An introductory paper on Space Vector Modulation Techniques for three phase inverters.
 
Star Delta Starters  Theory and applications of start delta (wye delta) starters.
 
Starting High Inertia Loads  Guidelines for selecting starter and motors for starting high inertia loads.
 
Stray Voltages from VFDs in Dairy Sheds  Stray voltages from incorrectly installed VFDs cause major problems with the dairy herd being milked.
 
VFDs and Harmonics  An introductory paper on VFDs and harmonics
 
Variable Speed Control  An introductory paper on Variable Speed Control. covering mechanical and electrical methods.
 
Variable Frequency Drives  An introductory paper on Variable Frequency Drives. This page is under construction.
 
VFDs and Unscreened Output Cables  Screened cables are used between the output of a VFD and the motor to reduce conducted emissions and stray voltages.
 
VFDs and Energy Saving  Using VFDs to save energy.
 
VFDs and EDM  Using VFDs and reducing EDM (Electrical Discharge Machining) bearing damage.
 
Contact Information
 
Telephone ++64 27 436 3067
FAX ++64 3 332 5220
Postal address P.O. Box 13-076, Christchurch, New Zealand
      
Mail.gif (4196 bytes)
Send us an email
 
Visits since 6 Aug 2002
[1]5 €¥[1]Tahoma[1]A€H€< 
________________________________________
[1]„ 
________________________________________
,[1][1]5 €¥[1]Tahoma[1] 
AE€H€, 
________________________________________
[1]s
€ê4ž 
L.M.Photonics Ltd
      
                   
      
 
 
 
 
 
 
 
 
 
 
 
 
In Association with Amazon.com
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
            
 
BusBar Rating Software
 
    This Software package is designed to give indicative current ratings and power dissipated for Aluminium and Copper BusBars as used in switchboards and electrical distribution systems. The user is able to specify the thickness and the width of the BusBar for thicknesses above 1mm.
 
    BusBar ratings are based on the expected surface temperature rise of the busbar. This is a function of the thermal resistance of the busbar and the power it dissipates. The thermal resistance of the busbar is a function of the surface area of the busbar, the orientation of the busbar, the material from which it is made, and the movement of air around it. The power dissipated by the bus bar is dependant on the square of the current passing through it, its length, and the material from which it is made.
 
    Optimal ratings are achieved when the bar runs horizontally with the face of the bar in the vertical plane. i.e. the bar is on its edge. There must be free air circulation around all of the bar in order to afford the maximum cooling to its surface. Restricted airflow around the bar will increase the surface temperature of the bar. If the bar is installed on its side, (largest area to the top) it will run at an elevated temperature and may need considerable derating. The actual derating required depends on the shape of the bar. Busbars with a high ratio between the width and the thickness, are more sensitive to their orientation than busbars that have an almost square cross section.
 
    Vertical busbars will run much hotter at the top of the bar than at the bottom, and should be derated in order to reduce the maximum temperature within allowable limits.
 
    Maximum BusBar ratings are not the temperature at which the busbar is expected to fail, rather it is the maximum temperature at which it is considered safe to operate the busbar due to other factors such as the temperature rating of insulation materials which may be in contact with, or close to, the busbar. Busbars which are sleeved in an insulation material such as a heatshrink material, may need to be derated because of the potential ageing and premature failure of the insulation material.
 
    wpe1.jpg (13183 bytes)
 
    The software provides calculation in both metric and imperial dimensions.
 
    wpe2.jpg (13406 bytes)
 
    The power dissipation in busbars can be calculated for specific currents.
 
    wpe3.jpg (16298 bytes)
 
    The Purchase Price for BusBar Calculations is $NZ35 or $US22.
 
    Click here to download  Electrical Calculations (including busbar calculations)
 
Download Manual Now
Motor Control Mailing List.
Join the Motor control mailing list to discuss motor control technology and problems with like minded experts.
 
Motor Control Forum
Online forum for discussion of motor control technologies.
 
Visits since 6 Aug 2002  [an error occurred while processing this directive]
                  
 
Home  Books  Software  Power factor  Motor Control  Motor Starters  Soft Starters  Variable Speed  Advertise  Energy Savers
 
Valid HTML 4.01 Transitional Valid CSS!
 
© L M Photonics Ltd | P.O. Box 13 076, Christchurch, New Zealand | phone : (NZ) +64 274 363 067
 
 
[1]5 €¥[1]Tahoma[1]( €start už [1]u €¥[1] 
________________________________________
Tahoma

(å  € °
ComboBox37v' ,å
€   8 CommandButton1bqk¬ €# 
________________________________________

Frame1E     þÿ 
________________________________________

ÿÿÿÿði*ÆÜÎ ž˜ªWJOMicrosoft Forms 2.0 FormEmbedded Obj
VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   10095
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15360
   OleObjectBlob   =   "UserForm1 calculator dhet.frx":0000
   StartUpPosition =   3  'Windows Default
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton16_Click()
 
End Sub
 
Private Sub CommandButton5_Click()
 
End Sub
 
Private Sub CommandButton6_Click()
 
End Sub
 
Private Sub CommandButton7_Click()
 
End Sub
 
Private Sub CommandButton8_Click()
 
End Sub
 
Private Sub CommandButton9_Click()
 
End Sub
 
Private Sub Frame1_Click()
 
End Sub
 
Private Sub Label1_Click()
 
End Sub
 
Private Sub ScrollBar1_Change()
 
End Sub
 
Private Sub TextBox1_Change()
 
End Sub
 
Private Sub TextBox2_Change()
 
End Sub
 
Private Sub TextBox3_Change()
 
End Sub
 
Private Sub TextBox3_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
 
End Sub
 
Private Sub TextBox3_DropButtonClick()
 
End Sub
 
Private Sub TextBox3_Enter()
 
End Sub
 
Private Sub TextBox3_Error(ByVal Number As Integer, ByVal Description As MSForms.ReturnString, ByVal SCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, ByVal CancelDisplay As MSForms.ReturnBoolean)
 
End Sub
 
Private Sub TextBox3_Exit(ByVal Cancel As MSForms.ReturnBoolean)
 
End Sub
 
Private Sub TextBox3_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
 
End Sub
 
Private Sub TextBox3_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
 
End Sub
 
Private Sub TextBox3_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
 
End Sub
 
Private Sub TextBox3_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
 
End Sub
 
Private Sub TextBox3_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
ÕiŽE[1]¼,€¦€dhet //st peace college , scale weigthing  portofolio   value                                                 analyse design investigate  psychometrical    memorendum-IX[1]5 €¥[1]TahomaRoot Entryÿÿÿÿÿÿÿÿ[1]ði*ÆÜÎ ž˜ªWJOðöcü ùÙ @f 
________________________________________
[1]ÿÿÿÿÿÿÿÿÿÿÿÿ’o 
________________________________________€}ÕiŽE[1]X‚t[1]¼,€¦€dhet //st peace college , scale weigthing  portofolio   value                                                 analyse design investigate  psychometrical    memorendum-IX[1]5 €¥[1]Tahomaþÿ 
________________________________________

Private Sub ButtonClickMethod(sender As Object, e As EventArgs) Handles num0.Click, num1.Click, num2.Click, num3.Click, num4.Click, num5.Click, num6.Click, num7.Click, num8.Click, num9.Click, opdivision.Click, opmultiply.Click, opdecimal.Click, opclear.Click, opminus.Click, opadd.Click, opequal.Click

    Dim button As Button = CType(sender, Button)

    If button.Name = "num1" Then

        boxequation1.Text = boxequation1.Text + "1"

    End If

    If button.Name = "num2" Then

        boxequation1.Text = boxequation1.Text + "2"

    End If

    If button.Name = "num3" Then

        boxequation1.Text = boxequation1.Text + "3"

    End If

    If button.Name = "num4" Then

        boxequation1.Text = boxequation1.Text + "4"

    End If

    If button.Name = "num5" Then

        boxequation1.Text = boxequation1.Text + "5"

    End If

    If button.Name = "num6" Then

        boxequation1.Text = boxequation1.Text + "6"

    End If

    If button.Name = "num7" Then

        boxequation1.Text = boxequation1.Text + "7"

    End If

    If button.Name = "num8" Then

        boxequation1.Text = boxequation1.Text + "8"

    End If

    If button.Name = "num9" Then

        boxequation1.Text = boxequation1.Text + "9"

    End If

    If button.Name = "num0" Then

        boxequation1.Text = boxequation1.Text + "0"

    End If

    If button.Name = "opdecimal" Then

        boxequation1.Text = boxequation1.Text + "."

    End If

    If button.Name = "opequal" Then

        Dim equation1 As String = boxequation1.Text

        Dim equation2 As String = boxequation2.Text

        Dim result = New DataTable().Compute(equation1, Nothing)


        boxresult.Text = result

    End If

    If button.Name = "opminus" Then

        boxequation1.Text = boxequation1.Text + "-"

        boxoperator.Text = boxoperator.Text + "-"

    End If

    If button.Name = "opmultiply" Then

        boxequation1.Text = boxequation1.Text + "*"

        boxoperator.Text = boxoperator.Text + "x"

    End If

    If button.Name = "opdivision" Then

        boxequation1.Text = boxequation1.Text + "/"

        boxoperator.Text = boxoperator.Text + "÷"

    End If

    If button.Name = "opadd" Then

        boxequation1.Text = boxequation1.Text + "+"

        boxoperator.Text = boxoperator.Text + "+"

    End If

    If button.Name = "opclear" Then

        boxequation1.Clear()

        boxoperator.Clear()

        boxresult.Clear()

    End If

End Sub

Private Sub opbackspace_Click(sender As Object, e As EventArgs) Handles opbackspace.Click

    boxequation1.Text = boxequation1.Text.Remove(boxequation1.Text.Count - 1)

End Sub


End Class

o learn, after all!

Here what it looks like:

.-.-.-.-.-.-.-.-. 

Private Sub New()

    InitializeComponent()


    _currentTextBox = Number1TextBox

    _locked = False

    ResultsTextBox.TextAlign = HorizontalAlignment.Center

    ResultsTextBox.Text = "SUPER DUPER CALCULATOR"

    For Each textBox As System.Windows.Forms.TextBox In {ResultsTextBox, Number1TextBox, Number2TextBox}

        'This prevent the user from writing in your textboxes instead of using the buttons.

        textBox.ReadOnly = True

    Next


    'I used the .Text and .Tag properties to make the code easier to manage. Observe:

    num0Button.Text = "0"

    num1Button.Text = "1"

    num2Button.Text = "2"

    num3Button.Text = "3"

    '...

    num9Button.Text = "9"

    opAddButton.Text = "+"

    opMinusButton.Text = "-"

    opMultiplyButton.Text = "x"

    opDivideButton.Text = "÷"

    For Each button As Button In {num1Button, num2Button, num3Button, num4Button, num5Button, num6Button, num7Button, num8Button, num9Button}

        button.Tag = "number"

    Next

    For Each button As Button In {opAddButton, opMinusButton, opMultiplyButton, opDivideButton}

        button.Tag = "operation"

    Next

    num0Button.Tag = "numberZero"

    opDecimalButton.Tag = "decimal"

    opBackspaceButton.Tag = "backspace"

    opEqualButton.Tag = "equal"

End Sub

'Class variables are useful to remember informations which will be useful at several unrelated places.

'_currentTextBox is a pointer to the TextBox which is currently being filled

Private _currentTextBox As System.Windows.Forms.TextBox = Number1TextBox

'_Llocked will prevent the user from messing with the calculator after he pressed "equal"

Private _locked As Boolean = False


'I didn't like this idea but I went with it since it was your initial method.

'Notice how shoter this iteration of your idea is. Strive to shorten your code while making your variable names much easier to read.

'For real, make sure your variables are aptly names. Every control should be identified as such (the button num0 should be names num0Button for example)

'Code is hard enough to read without making it sibylline

Private Sub ButtonClickMethod(sender As Object, e As EventArgs) Handles num0Button.Click, num1Button.Click, num2Button.Click, num3Button.Click, num4Button.Click, num5Button.Click, num6Button.Click, num7Button.Click, num8Button.Click, num9Button.Click, opDivideButton.Click, opMultiplyButton.Click, opDecimalButton.Click, opClearButton.Click, opMinusButton.Click, opAddButton.Click, opEqualButton.Click, opBackspaceButton.Click

    'DirectCast is more strick than CType. For this operation, it doesn't really make a difference, but that's the one you want anyway.

    Dim button As System.Windows.Forms.Button = DirectCast(sender, System.Windows.Forms.Button)


    'The boolean _locked prevent the user from messing with your beautiful calculator.

    If Not _locked Then

        'This is where the .Tag property save lifes. All numbers (but zero) are "number" now. BAM!

        Select Case button.Tag.ToString

            Case "number"

                _currentTextBox.Text += button.Text

            Case "numberZero"

                'You don't want your clever teacher to try to divide by zero, or create a number with 5 zeros before the actual number.

                'This is how you can do this.

                If _currentTextBox.Text.Length > 0 Then

                    _currentTextBox.Text += button.Text

                End If

            Case "decimal"

                If _currentTextBox.Text.Length > 0 AndAlso Not _currentTextBox.Text.Contains(".") Then

                    _currentTextBox.Text += button.Text

                End If

            Case "operation"

                ResetOperationTextBoxes()

                'here's a nifty visual cue about which operation has been selected. It's not really important.

                button.BackColor = Color.LightSalmon

                'This label will be used to Calculate() the operation. It's important!

                OpLabel.Text = button.Text

            Case "backspace"

                If _currentTextBox.Text.Length > 0 Then

                    _currentTextBox.Text = _currentTextBox.Text.Substring(0, _currentTextBox.Text.Length - 1)

                End If

            Case "equal"

                'I could have put the code here instead of making another sub, but I didn't.

                'There are two valid reasons for this.

                '1- This code might be useful on it's own later: maybe something else will Calculate() or LockCalculator or ResetCalculator later.

                '2- try to avoid blobs of code. A big and complex sub is easier to understand when it's divided between several clean and clear smaller subs.

                '   (of course, you cannot always subdivise Subs. Use common sense)

                Calculate()

        End Select

    End If


    'This check is all alone because I want it to work even if the calculator is locked.

    If button.Tag.ToString = "clear" Then

        ResetCalculator()

    End If

End Sub


Private Sub Calculate()

    'Always make checks to make sure that things are going the way you think they should be going.

    'for example, here I want to have 2 numbers and an operation, er else I'll just ignore this click.

    If Number1TextBox.Text.Length > 0 AndAlso Number2TextBox.Text.Length > 0 And OpLabel.Text <> "?" Then

        'If the user could mess with the textboxes, i would have to make further checks to avoid crashes, but I already force the user to use only numbers se we're cool.

        Dim number1 As Double = Convert.ToDouble(Number1TextBox.Text)

        Dim number2 As Double = Convert.ToDouble(Number2TextBox.Text)

        Dim result As Double


        Select Case OpLabel.Text

            Case "+"

                result = number1 + number2

            Case "-"

                result = number1 - number2

            Case "x"

                result = number1 * number2

            Case "÷"

                result = number1 / number2

        End Select


        ResultsTextBox.TextAlign = HorizontalAlignment.Right

        ResultsTextBox.Text = result.ToString

        LockCalculator(True)

    End If

End Sub


'By using a boolean to signify to this Sub if I want to lock or unlock things, I make it easier to automatize this operation (and I have only one Sub instead of two).

Private Sub LockCalculator(ByVal isLocked As Boolean)

    _locked = isLocked

    'I'm kinda cheating here: I put the calculator inside a GroupBox just so I could be lazier and do this.

    'Being lazy is great for a coder, as long as you think ahead (and avoid being a nuisance to yourself when you'll come back to this code and weep).

    For Each control As System.Windows.Forms.Control In CalculatorGroupBox.Controls

        control.Enabled = Not isLocked

    Next

    'Except that I want this button to always be available, so I enable it. This operation wouldn't be necessary if isLocked is True, but checking for that is longer than just doing it.

    opClearButton.Enabled = True

End Sub


Private Sub ResetCalculator()

    ResetNumberTextBoxes()

    ResetOperationTextBoxes()

    LockCalculator(False)

    _currentTextBox = Number1TextBox

End Sub


Private Sub ResetNumberTextBoxes()

    For Each textBox As System.Windows.Forms.TextBox In {Number1TextBox, Number2TextBox}

        textBox.Text = ""

    Next

    'Mea Culpa: I shouldn't had done a For Each for 2 elements like that. I did it anyway. Please don't call the Fun Police on me.

    'Also, you can now guess that using For Each loops to set up controls is a work method I like. Some don't like that.

    'There is about as many ways to code something as there are coders. You have to find your own, sure, but then you'll get hired somewhere

    'and they will absolutely hate everything that you do that isn't done the way they like it. And they will be right. On a big project where

    'several coders works, at the same time and for month and years, you have to find a common ground or else the code will become a nightmare

    'to understand and maintain. Trust me: being able to adapt quickly to other's work methodology is a GREAT skill in this field of work.


    ResultsTextBox.TextAlign = HorizontalAlignment.Center

    ResultsTextBox.Text = "SUPER DUPER CALCULATOR"

End Sub


Private Sub ResetOperationTextBoxes()

    For Each button As System.Windows.Forms.Button In {opAddButton, opMinusButton, opMultiplyButton, opDivideButton}

        button.BackColor = DefaultBackColor

    Next

    OpLabel.Text = "?"

End Sub


'This sub is great. It uses the _currentTextBox as a pointer. Maybe you're not familiar with pointers, I don't know. The short explanation is this:

'A pointer is like a phone number. You can give your number to several people without "losing" it. And if they call you and tell you something, they

'are all speaking to the same person.

'_currentTextBox is a pointer which can point toward Number1TextBox or Number2TextBox.

'This Event makes it so when the user click on Number1TextBox or Number2TextBox, the pointer updates it's "phone number" to the right box,

'so later when we'll write we'll write in the last one which was clicked on.

Private Sub NumberTextBoxes_Click(sender As Object, e As EventArgs) Handles Number1TextBox.Click, Number2TextBox.Click

    _currentTextBox = DirectCast(sender, System.Windows.Forms.TextBox)













 

// Java program to create a simple calculator

// with basic +, -, /, * using java swing elements

 

import java.awt.event.*;

import javax.swing.*;

import java.awt.*;

class calculator extends JFrame implements ActionListener {

    // create a frame

    static JFrame f;

 

    // create a textfield

    static JTextField l;

 

    // store operator and operands

    String s0, s1, s2;

 

    // default constructor

    calculator()

    {

        s0 = s1 = s2 = "";

    }

 

    // main function

    public static void main(String args[])

    {

        // create a frame

        f = new JFrame("calculator");

 

        try {

            // set look and feel

            UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());

        }

        catch (Exception e) {

            System.err.println(e.getMessage());

        }

 

        // create a object of class

        calculator c = new calculator();

 


    OpLabel.Text = "?"

End Sub


'This sub is great. It uses the _currentTextBox as a pointer. Maybe you're not familiar with pointers, I don't know. The short explanation is this:

'A pointer is like a phone number. You can give your number to several people without "losing" it. And if they call you and tell you something, they

'are all speaking to the same person.

'_currentTextBox is a pointer which can point toward Number1TextBox or Number2TextBox.

'This Event makes it so when the user click on Number1TextBox or Number2TextBox, the pointer updates it's "phone number" to the right box,

'so later when we'll write we'll write in the last one which was clicked on.

Private Sub NumberTextBoxes_Click(sender As Object, e As EventArgs) Handles Number1TextBox.Click, Number2TextBox.Click

    _currentTextBox = DirectCast(sender, System.Windows.Forms.TextBox)













 

// Java program to create a simple calculator

// with basic +, -, /, * using java swing elements

 

import java.awt.event.*;

import javax.swing.*;

import java.awt.*;

class calculator extends JFrame implements ActionListener {

    // create a frame

    static JFrame f;

 

    // create a textfield

    static JTextField l;

 

    // store operator and operands

    String s0, s1, s2;

 

    // default constructor

    calculator()

    {

        s0 = s1 = s2 = "";

    }

 

    // main function

    public static void main(String args[])

    {

        // create a frame

        f = new JFrame("calculator");

 

        try {

            // set look and feel

            UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());

        }

        catch (Exception e) {

            System.err.println(e.getMessage());

        }

 

        // create a object of class

        calculator c = new calculator();

 

        

        // set the textfield to non editable

        l.setEditable(false);

 

        // create number buttons and some operators

        JButton b0, b1, b2, b3, b4, b5, b6, b7, b8, b9, ba, bs, bd, bm, be, beq, beq1;

 

        // create number buttons

        b0 = new JButton("0");

        b1 = new JButton("1");

        b2 = new JButton("2");

        b3 = new JButton("3");

        b4 = new JButton("4");

        b5 = new JButton("5");

        b6 = new JButton("6");
VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "UserForm2"
   ClientHeight    =   9075
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15150
   OleObjectBlob   =   "UserForm2 atm dhet portofolio statemen.frx":0000
   StartUpPosition =   3  'Windows Default
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label1_Click()
 
End Sub
 
Private Sub Label6_Click()
 
End Sub
 
Private Sub OptionButton4_Click()
 
End Sub
 
Private Sub ScrollBar1_Change()
 
End Sub
 
Private Sub ScrollBar2_Change()
 
End Sub
 
Private Sub TextBox1_Change()
 
End Sub
 
Private Sub TextBox2_Change()
 
End Sub
 
Private Sub TextBox3_Change()
 
End Sub
 
Private Sub TextBox5_Change()
 
End Sub
 
Private Sub TextBox6_Change()
 
End Sub
 
Private Sub TextBox7_Change()
 
End Sub
 
Private Sub TextBox8_Change()
 
End Sub
 
Private Sub UserForm_Click()
 
End Sub
 
Private Sub UserForm_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
 
End Sub
 
Private Sub UserForm_Deactivate()
 
End Sub
 
Private Sub UserForm_Initialize()
 
End Sub
 
Private Sub UserForm_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
 
End Sub
 
Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
 
End Sub
 
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
 
End Sub
 
Private Sub UserForm_RemoveControl(ByVal Control As MSForms.Control)
 
End Sub
 
Private Sub UserForm_Resize()
 
End Sub
 
Private Sub UserForm_Scroll(ByVal ActionX As MSForms.fmScrollAction, ByVal ActionY As MSForms.fmScrollAction, ByVal RequestDx As Single, ByVal RequestDy As Single, ByVal ActualDx As MSForms.ReturnSingle, ByVal ActualDy As MSForms.ReturnSingle)
 
End Sub
 
Private Sub UserForm_Terminate()
 
End Sub
 
Private Sub UserForm_Zoom(Percent As Integer)
 
End Sub
LB Ðˆ;%ÐÏ à¡±á> 
________________________________________
þÿ      
________________________________________
ch‡>Root Entryÿÿÿÿÿÿÿÿ[1]ði*ÆÜÎ ž˜ªWJOà Ø¹      ùÙ @f 
________________________________________
[1]ÿÿÿÿÿÿÿÿÿÿÿÿfo 
________________________________________
[1] 
________________________________________
ÿÿÿÿþÿÿÿi01 ÿÿÿÿ  n`ôÎ ›Íª`ŽÈÇ¹  ùÙ€âÍ ôZ98Frame1 
________________________________________
Rã
‘ Î ãªK¸Q DB Tahoma¬LŸ(õ € 
________________________________________
24 Label1" 
________________________________________
ö 
________________________________________
(õ € 
________________________________________
28 Label2" 
________________________________________
ì      (õ €2<[1] Label3O 
________________________________________
â
(õ € 2< 
________________________________________
Label4O 
________________________________________
________________________________________
(õ €28 
________________________________________
Label5O 
________________________________________
[1] ( 
________________________________________
€name`O" 
________________________________________
[1]5 €¥[1]Tahoma[1](€surname4P" 
________________________________________
[1]5 €¥[1]Tahomae[1] (
€description4PN 
________________________________________
[1]5 €¥[1]Tahomae[1] (
€birth date 4P# 
________________________________________
[1]5 €¥[1]Tahomae[1](€adresseQö 
________________________________________
[1]5 €¥[1]Tahomae[1](€gender Qö 
________________________________________
[1]5 €¥[1]Tahoma [1] (
€user name  ÛQ" 
________________________________________
[1]5 €¥[1]Tahoma [1]( €password¯R" 
________________________________________
[1]5 €¥[1]Tahomard[1]( €Laatm rd‚S" 
________________________________________
[1]5 €¥[1]Tahomard[1],@€H€, €4ö 
________________________________________
"text  label1 name"[1]5 €¥[1]Tahomard[1] @€H€,€ë4" 
________________________________________
"textard[1]5 €¥[1]Tahomard[1] €H€,4" 
________________________________________
[1]5 €¥[1]Tahomard[1] €H€,ë4O 
________________________________________
[1]5 €¥[1]Tahomard[1] €H€,4" 
________________________________________
[1]5 €¥[1]Tahomard[1] €H€,ê4" 
________________________________________
[1]5 €¥[1]Tahomard[1]È®AE€H€, 
________________________________________
[1]ª®¾5ö 
________________________________________
Java project| Create Page for ATM 

// ATM Application using gui in java Source Code 

import java.awt.*;

import java.awt.event.*;

import javax.swing.*;

import java.applet.*;

public class Banking2 extends JPanel implements ActionListener

{

 

//,l3,l4,l5,l6,user,l7,l8; // declared label to display text

 JTextField t1,t2,t5,t7;  // declare area to type some text

JTextArea t4,t3;          // for multiple text like description 

 JPasswordField pw;  // for password field

 JButton b1,b2,ok;      // for button 

 JRadioButton r1,r2,r3;  // ratio button 

JTextArea output;


Banking2()

{

// this is use to display text on frame 

f2=new JFrame( Create your a/c ); 


// this is use to display text label on the screen

l1=new JLabel( Name: );                 

l2=new JLabel( Surname );

l3=new JLabel( Discription );

l4=new JLabel( Date of Birth );

l5=new JLabel( Gender );

l6=new JLabel( address );

user=new JLabel( User Name: );

 l7=new JLabel( password );

l8=new JLabel( phone no );


// this area is use to write text as well as to enter password and 

//  number

t1=new JTextField();

t2=new JTextField();

t3=new JTextArea();

t4=new JTextArea();

t5=new JTextField();

pw=new JPasswordField();

t7=new JTextField();


output=new JTextArea();  //for output


// this are is use to display image on screen

 

ImageIcon bg=new ImageIcon( atm.png );

ImageIcon bg1=new ImageIcon( ok.png );

ImageIcon bg2=new ImageIcon( cancel.png );


//ImageIcon bg3=new ImageIcon( ok.png );

JLabel bgi=new JLabel(  ,bg,JLabel.CENTER);


// for selecting gender using the radio button

ButtonGroup o=new ButtonGroup();

r1=new JRadioButton( male ,false);

r2=new JRadioButton( female ,false);

r3=new JRadioButton( other ,false);


// this for checking the condition

 JCheckBox cond=new JCheckBox( I accept all term and condition );


b1=new JButton(bg1);

b2=new JButton(bg2);


//ok=new JButton(bg3);

Choice c=new Choice();

Choice c2=new Choice();

Choice c3=new Choice();


// about size , layout , color of frame 

f2.setVisible(true);

f2.setSize(1000,1000);

f2.setLayout(null);

f2.setBackground(Color.YELLOW);

f2.setForeground(Color.BLUE);

bgi.setBounds(500,0,500,500);


// location of the object in the frame


cond.setBounds(10,400,400,20);

output.setBounds(500,0,500,500);

output.setBackground(Color.yellow);

output.setForeground(Color.red);


l1.setBounds(10,10,100,20);

t1.setBounds(110,10,200,20);


l2.setBounds(10,50,100,20);

t2.setBounds(110,50,200,20);


l3.setBounds(10,90,100,20);

t3.setBounds(110,90,200,30);


l4.setBounds(10,130,100,20);

c.setBounds(110,130,100,20);

c2.setBounds(210,130,100,20);

c3.setBounds(310,130,100,20);


l5.setBounds(10,170,100,20);

r1.setBounds(110,170,100,20);

r2.setBounds(210,170,100,20);

r3.setBounds(310,170,100,20);


l6.setBounds(10,210,100,20);

t4.setBounds(110,210,200,20);

t4.setEditable(false);

user.setBounds(10,250,100,20);

t5.setBounds(110,250,200,20);


l7.setBounds(10,290,100,20);

pw.setBounds(110,290,200,20);


l8.setBounds(10,330,100,20);

t7.setBounds(110,330,200,20);


b1.setBounds(30,500,500,100);

b2.setBounds(600,500,500,100);

//ok.setBounds(360,500,300,100);


// adding all the component inside the frame

f2.add(l1);

f2.add(t1);

f2.add(l2);

f2.add(t2);

f2.add(l3);

f2.add(t3);

f2.add(l4);

f2.add(c);

f2.add(c2);

f2.add(c3);

f2.add(bgi);


.. to select  day of birthday

c.add( 1 );

c.add( 2 );

c.add( 3 );

c.add( 4 );

c.add( 5 );

c.add( 6 );

c.add( 7 );

c.add( 8 );

c.add( 9 );

c.add( 10 );

c.add( 11 );

c.add( 12 );

c.add( 13 );

c.add( 14 );

c.add( 15 );

c.add( 16 );

c.add( 17 );

c.add( 18 );

c.add( 19 );

c.add( 20 );

c.add( 21 );

c.add( 22 );

c.add( 23 );

c.add( 24 );

c.add( 25 );

c.add( 26 );

c.add( 27 );

c.add( 28 );

c.add( 29 );

c.add( 30 );

c.add( 31 );


// this for month


c2.add( 1 );

c2.add( 2 );

c2.add( 3 );

c2.add( 4 );

c2.add( 5 );

c2.add( 6 );

c2.add( 7 );

c2.add( 8 );

c2.add( 9 );

c2.add( 10 );

c2.add( 11 );

c2.add( 12 );


// this is for year

c3.add( 2000 );

c3.add( 2001 );

c3.add( 2002 );

c3.add( 2003 );

c3.add( 2004 );

c3.add( 2005 );

c3.add( 2006 );

c3.add( 2007 );

c3.add( 2008 );

c3.add( 2009 );

c3.add( 2010 );

c3.add( 2011 );

c3.add( 2012 );

c3.add( 2013 );

c3.add( 2014 );

c3.add( 2015 );

c3.add( 2016 );

c3.add( 2017 );

c3.add( 2018 );

c3.add( 2019 );


f2.add(l5);

o.add(r1);

o.add(r2);

o.add(r3);

f2.add(r1);

f2.add(r2);

f2.add(r3);


f2.add(l6);

f2.add(t4);

f2.add(user);

f2.add(t5);

f2.add(l7);

 f2.add(pw);

f2.add(l8);

 f2.add(t7);


f2.add(b1);

//f2.add(ok);

f2.add(b2);


f2.add(output);

f2.add(cond);


// adding listener to the button


cond.addActionListener(this);

b1.addActionListener(this);

b2.addActionListener(this);

ok.addActionListener(this);

}

public void actionPerformed(ActionEvent e)

{

String name=t1.getText();

String surname=t2.getText();

String discription=t3.getText();

//String day=c.getSelectedItem();

//String month=c2.getSelectedItem();

//String years=c3.getSelectedItem();

String gender= male ;

if(r2.isSelected()==true)

gender= female ;

if(r3.isSelected()==true)

gender= other ;

String address=t4.getText();

String users=t5.getText();

String pass=pw.getText();

String phone=t7.getText();


if(e.getSource().equals(b1))

{

if(t1.getText().isEmpty()||(t2.getText().isEmpty())||(t3.getText().isEmpty())||(t4.getText().isEmpty())||(t5.getText().isEmpty())||(pw.getText().isEmpty())||(t7.getText().isEmpty())||(r1.isSelected())||(r2.isSelected())||(r3.isSelected()))

{

JOptionPane.showMessageDialog(null, Data missing );

}

else

{

JOptionPane.showMessageDialog(null, Data completed );

}

output.setText( name:-   +name+ nSurname:-   +surname+ nDiscription:-   +discription+ nGender:-   +gender+ nAddress:- +address+ nUser Name:- +users+ nPass word:- +pass+ nMobile No:- +phone);



//FileReader fr=new FileReader( abc.txt );

}

if(e.getSource().equals(b2))

{

f2.setVisible(false);

new Banking();

}

}

public static void  main(String s[])

{

new Banking2();

}

} 



1. //import required classes and packages   

2. import java.util.Scanner;  

3.  

4. //create ATMExample class to implement the ATM functionality  

5. public class ATMExample  

6. {  

7.    //main method starts   

8.    public static void main(String args[] )  

9.    {  

10.            //declare and initialize balance, withdraw, and deposit  

11.            int balance = 100000, withdraw, deposit;  

12.              

13.            //create scanner class object to get choice of user  

14.            Scanner sc = new Scanner(System.in);  

15.              

16.            while(true)  

17.            {  

18.                System.out.println("Automated Teller Machine");  

19.                System.out.println("Choose 1 for Withdraw");  

20.                System.out.println("Choose 2 for Deposit");  

21.                System.out.println("Choose 3 for Check Balance");  

22.                System.out.println("Choose 4 for EXIT");  

23.                System.out.print("Choose the operation you want to perform:");  

24.                  

25.                //get choice from user  

26.                int choice = sc.nextInt();  

27.                switch(choice)  

28.                {  

29.                    case 1:  

30.            System.out.print("Enter money to be withdrawn:");  

31.                          

32.            //get the withdrawl money from user  

33.            withdraw = sc.nextInt();  

34.                          

35.            //check whether the balance is greater than or equal to the withdrawal amount  

36.            if(balance >= withdraw)  

37.            {  

38.                //remove the withdrawl amount from the total balance  

39.                balance = balance - withdraw;  

40.                System.out.println("Please collect your money");  

41.            }  

42.            else  

43.            {  

44.                //show custom error message   

45.                System.out.println("Insufficient Balance");  

46.            }  

47.            System.out.println("");  

48.            break;  

49.       

50.                    case 2:  

,l3,l4,l5,l6,user,l7,l8; // declared label to display text

 JTextField t1,t2,t5,t7;  // declare area to type some text

JTextArea t4,t3;          // for multiple text like description 

 JPasswordField pw;  // for password field

 JButton b1,b2,ok;      // for button 

 JRadioButton r1,r2,r3;  // ratio button 

JTextArea output;


Banking2()

{

// this is use to display text on frame 

f2=new JFrame( Create your a/c ); 


// this is use to display text label on the screen

l1=new JLabel( Name: );                 

l2=new JLabel( Surname );

l3=new JLabel( Discription );

l4=new JLabel( Date of Birth );

l5=new JLabel( Gender );

l6=new JLabel( address );

user=new JLabel( User Name: );

 l7=new JLabel( password );

l8=new JLabel( phone no );


// this area is use to write text as well as to enter password and 

//  number

t1=new JTextField();

t2=new JTextField();

t3=new JTextArea();

t4=new JTextArea();

t5=new JTextField();

pw=new JPasswordField();

t7=new JTextField();


output=new JTextArea();  //for output


// this are is use to display image on screen

 

ImageIcon bg=new ImageIcon( atm.png );

ImageIcon bg1=new ImageIcon( ok.png );

ImageIcon bg2=new ImageIcon( cancel.png );


//ImageIcon bg3=new ImageIcon( ok.png );

JLabel bgi=new JLabel(  ,bg,JLabel.CENTER);


// for selecting gender using the radio button

ButtonGroup o=new ButtonGroup();

r1=new JRadioButton( male ,false);

r2=new JRadioButton( female ,false);

r3=new JRadioButton( other ,false);


// this for checking the condition

 JCheckBox cond=new JCheckBox( I accept all term and condition );


b1=new JButton(bg1);

b2=new JButton(bg2);


//ok=new JButton(bg3);

Choice c=new Choice();

Choice c2=new Choice();

Choice c3=new Choice();


// about size , layout , color of frame 

f2.setVisible(true);

f2.setSize(1000,1000);

f2.setLayout(null);

f2.setBackground(Color.YELLOW);

f2.setForeground(Color.BLUE);

bgi.setBounds(500,0,500,500);


// location of the object in the frame


cond.setBounds(10,400,400,20);

output.setBounds(500,0,500,500);

output.setBackground(Color.yellow);

output.setForeground(Color.red);


l1.setBounds(10,10,100,20);

t1.setBounds(110,10,200,20);


l2.setBounds(10,50,100,20);

t2.setBounds(110,50,200,20);


l3.setBounds(10,90,100,20);

t3.setBounds(110,90,200,30);


l4.setBounds(10,130,100,20);

c.setBounds(110,130,100,20);

c2.setBounds(210,130,100,20);

c3.setBounds(310,130,100,20);


l5.setBounds(10,170,100,20);

r1.setBounds(110,170,100,20);

r2.setBounds(210,170,100,20);

r3.setBounds(310,170,100,20);


l6.setBounds(10,210,100,20);

t4.setBounds(110,210,200,20);

t4.setEditable(false);

user.setBounds(10,250,100,20);

t5.setBounds(110,250,200,20);


l7.setBounds(10,290,100,20);

pw.setBounds(110,290,200,20);


l8.setBounds(10,330,100,20);

t7.setBounds(110,330,200,20);


b1.setBounds(30,500,500,100);

b2.setBounds(600,500,500,100);

//ok.setBounds(360,500,300,100);


// adding all the component inside the frame

f2.add(l1);

f2.add(t1);

f2.add(l2);

f2.add(t2);

f2.add(l3);

f2.add(t3);

f2.add(l4);

f2.add(c);

f2.add(c2);

f2.add(c3);

f2.add(bgi);


.. to select  day of birthday

c.add( 1 );

c.add( 2 );

c.add( 3 );

c.add( 4 );

c.add( 5 );

c.add( 6 );

c.add( 7 );

c.add( 8 );

c.add( 9 );

c.add( 10 );

c.add( 11 );

c.add( 12 );

c.add( 13 );

c.add( 14 );

c.add( 15 );

c.add( 16 );

c.add( 17 );

c.add( 18 );

c.add( 19 );

c.add( 20 );

c.add( 21 );

c.add( 22 );

c.add( 23 );

c.add( 24 );

c.add( 25 );

c.add( 26 );

c.add( 27 );

c.add( 28 );

c.add( 29 );

c.add( 30 );

c.add( 31 );


// this for month


c2.add( 1 );

c2.add( 2 );

c2.add( 3 );

c2.add( 4 );

c2.add( 5 );

c2.add( 6 );

c2.add( 7 );

c2.add( 8 );

c2.add( 9 );

c2.add( 10 );

c2.add( 11 );

c2.add( 12 );


// this is for year

c3.add( 2000 );

c3.add( 2001 );

c3.add( 2002 );

c3.add( 2003 );

c3.add( 2004 );

c3.add( 2005 );

c3.add( 2006 );

c3.add( 2007 );

c3.add( 2008 );

c3.add( 2009 );

c3.add( 2010 );

c3.add( 2011 );

c3.add( 2012 );

c3.add( 2013 );

c3.add( 2014 );

c3.add( 2015 );

c3.add( 2016 );

c3.add( 2017 );

c3.add( 2018 );

c3.add( 2019 );


f2.add(l5);

o.add(r1);

o.add(r2);

o.add(r3);

f2.add(r1);

f2.add(r2);

f2.add(r3);


f2.add(l6);

f2.add(t4);

f2.add(user);

f2.add(t5);

f2.add(l7);

 f2.add(pw);

f2.add(l8);

 f2.add(t7);


f2.add(b1);

//f2.add(ok);

f2.add(b2);


f2.add(output);

f2.add(cond);


// adding listener to the button


cond.addActionListener(this);

b1.addActionListener(this);

b2.addActionListener(this);

ok.addActionListener(this);

}

public void actionPerformed(ActionEvent e)

{

String name=t1.getText();

String surname=t2.getText();

String discription=t3.getText();

//String day=c.getSelectedItem();

//String month=c2.getSelectedItem();

//String years=c3.getSelectedItem();

String gender= male ;

if(r2.isSelected()==true)

gender= female ;

if(r3.isSelected()==true)

gender= other ;

String address=t4.getText();

String users=t5.getText();

String pass=pw.getText();

String phone=t7.getText();


if(e.getSource().equals(b1))

{

if(t1.getText().isEmpty()||(t2.getText().isEmpty())||(t3.getText().isEmpty())||(t4.getText().isEmpty())||(t5.getText().isEmpty())||(pw.getText().isEmpty())||(t7.getText().isEmpty())||(r1.isSelected())||(r2.isSelected())||(r3.isSelected()))

{

JOptionPane.showMessageDialog(null, Data missing );

}

else

{

JOptionPane.showMessageDialog(null, Data completed );

}

output.setText( name:-   +name+ nSurname:-   +surname+ nDiscription:-   +discription+ nGender:-   +gender+ nAddress:- +address+ nUser Name:- +users+ nPass word:- +pass+ nMobile No:- +phone);



//FileReader fr=new FileReader( abc.txt );

}

if(e.getSource().equals(b2))

{

f2.setVisible(false);

new Banking();

}

}

public static void  main(String s[])

{

new Banking2();

}

} 



1. //import required classes and packages   

2. import java.util.Scanner;  

3.  

4. //create ATMExample class to implement the ATM functionality  

5. public class ATMExample  

6. {  

7.    //main method starts   

8.    public static void main(String args[] )  

9.    {  

10.            //declare and initialize balance, withdraw, and deposit  

11.            int balance = 100000, withdraw, deposit;  

12.              

13.            //create scanner class object to get choice of user  

14.            Scanner sc = new Scanner(System.in);  

15.              

16.            while(true)  

17.            {  

18.                System.out.println("Automated Teller Machine");  

19.                System.out.println("Choose 1 for Withdraw");  

20.                System.out.println("Choose 2 for Deposit");  

21.                System.out.println("Choose 3 for Check Balance");  

22.                System.out.println("Choose 4 for EXIT");  

23.                System.out.print("Choose the operation you want to perform:");  

24.                  

25.                //get choice from user  

26.                int choice = sc.nextInt();  

27.                switch(choice)  

28.                {  

29.                    case 1:  

30.            System.out.print("Enter money to be withdrawn:");  

31.                          

32.            //get the withdrawl money from user  

33.            withdraw = sc.nextInt();  

34.                          

35.            //check whether the balance is greater than or equal to the withdrawal amount  

36.            if(balance >= withdraw)  

37.            {  

38.                //remove the withdrawl amount from the total balance  

39.                balance = balance - withdraw;  

40.                System.out.println("Please collect your money");  

41.            }  

42.            else  

43.            {  

44.                //show custom error message   

45.                System.out.println("Insufficient Balance");  

46.            }  

47.            System.out.println("");  

48.            break;  

49.       

50.                    case 2:  

51.                          

VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm4 
   Caption         =   "UserForm4"
   ClientHeight    =   9420
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14355
   OleObjectBlob   =   "UserForm4 metering subject.frx":0000
   StartUpPosition =   1  'CenterOwner
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
End
Attribute VB_Name = "UserForm4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
 
End Sub
 
Private Sub Frame1_Click()
 
End Sub
 
Private Sub Label1_Click()
 
End Sub
 
Private Sub Label2_Click()
 
End Sub
 
Private Sub Label3_Click()
 
End Sub
 
Private Sub Label4_Click()
 
End Sub
 
Private Sub ScrollBar1_Change()
 
End Sub
 
Private Sub TextBox1_Change()
 
End Sub
 
Private Sub TextBox2_Change()
 
End Sub
 
Private Sub TextBox3_Change()
 
End Sub
 
Private Sub TextBox4_Change()
 
End Sub
 
Private Sub TextBox4_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
 
End Sub
 
Private Sub TextBox4_DropButtonClick()
 
End Sub
 
Private Sub TextBox4_Enter()
 
End Sub
 
Private Sub TextBox4_Exit(ByVal Cancel As MSForms.ReturnBoolean)
 
End Sub
 
Private Sub TextBox4_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
 
End Sub
 
Private Sub TextBox4_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
 
End Sub
 
Private Sub TextBox4_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
 
End Sub
 
Private Sub TextBox4_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
 
End Sub
 
Private Sub TextBox4_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
 
End Sub
 
Private Sub UserForm_Click()
  
LB Rm8p&ÐÏ à¡±á> 
________________________________________ Write a Java Program to Calculate Electricity Bill using Else If statement with examples.

Java Program to Calculate Electricity Bill using Else If

This Java program asks the user to enter the units consumed. Next, it calculates the Total Electricity bill. This approach is useful if the Electricity board is charging different tariffs for different units. For this example, we are using the Else If Statement.

 

TIP: The Java Else If statement checks the first condition if it is TRUE, it executes the statements inside that block. If the condition is FALSE, it checks the Next one (Else If condition) and so on.

import java.util.Scanner;


public class ElectricityBill1 {

     private static Scanner sc;

     public static void main(String[] args) 

     {

          int Units;

          double Amount, Sur_Charge, Total_Amount;

          sc = new Scanner(System.in);

          

          System.out.print(" Please Enter the Units that you Consumed  : ");

          Units = sc.nextInt();

  

        if (Units < 50)

        {

             Amount = Units * 2.60;

             Sur_Charge = 25;       

        } 

        else if (Units <= 100)

        {

             // For the First Fifty Units Charge = 130 (50 * 2.60)

             // Next, we are removing those 50 units from total units

             Amount = 130 + ((Units - 50 ) * 3.25);

             Sur_Charge = 35; 

        }

        else if (Units <= 200)

        {

             // First Fifty Units charge = 130, and 50 - 100 is 162.50 (50 * 3.25)

             // Next, we are removing those 100 units from total units

             Amount = 130 + 162.50 + ((Units - 100 ) * 5.26);

             Sur_Charge = 45; 

        }

        else

        {

             /* First Fifty Units = 130, 50 - 100 is 162.50, 

              and 100 - 200 is 526 (100 * 5.65)

             Next, we are removing those 200 units from total units */

                  Amount = 130 + 162.50 + 526 + ((Units - 200 ) * 7.75); 

                  Sur_Charge = 55; 

          }

          

          Total_Amount = Amount + Sur_Charge;

          System.out.println("\n Electricity Bill  =  " + Total_Amount);

     }


}

 

Java Program to find Electricity Bill using Method

This program to calculate electricity bills is the same as the first example. But we separated the logic and placed it in a separate method.

import java.util.Scanner;


public class Example2 {

     private static Scanner sc;

     public static void main(String[] args) 

     {

          int Units;   

          double Total_Amount;

          sc = new Scanner(System.in);

          

          System.out.print(" Please Enter the Units that you Consumed  : ");

          Units = sc.nextInt();

  

          Total_Amount = ElecBill(Units);

          System.out.println("\n Electricity Bill  =  " + Total_Amount);

     }

     public static double ElecBill(int Units)

     {

          double Amount, Sur_Charge, Total_Amount;

          if (Units < 50)

        {

             Amount = Units * 2.60;

             Sur_Charge = 25;       

        } 

        else if (Units <= 100)

        {

             Amount = 130 + ((Units - 50 ) * 3.25);

             Sur_Charge = 35; 

        }

        else if (Units <= 200)

        {

             Amount = 130 + 162.50 + ((Units - 100 ) * 5.26);

             Sur_Charge = 45; 

        }

        else

        {

                  Amount = 130 + 162.50 + 526 + ((Units - 200 ) * 7.75); 

                  Sur_Charge = 55; 

          }

          

          Total_Amount = Amount + Sur_Charge;

          return Total_Amount;

     }

}

 Please Enter the Units that you Consumed  : 95


 Electricity Bill  =  311.25

Let me try a different value.

 

 Please Enter the Units that you Consumed  : 450


 Electricity Bill  =  2811.0

Java Program to find Electricity Bill Example

This type of Java program is useful if the board has uniform rates. Something like: if you consume between 300 and 500 units, then charges are fixed as 7.75 Rupees for Units, etc.

import java.util.Scanner;


public class Example3 {

     private static Scanner sc;

     public static void main(String[] args) 

     {

          int Units;

          double Amount, Sur_Charge, Total_Amount;

          sc = new Scanner(System.in);

          

          System.out.print(" Please Enter the Units that you Consumed  : ");

          Units = sc.nextInt();

          

          if (Units > 500)

        {

             Amount = Units * 9.65;

             Sur_Charge = 85;       

        } // Otherwise (fails), Compiler will move to Next Else If block 

        else if (Units >= 300)

        {

             Amount = Units * 7.75;

             Sur_Charge = 75; 

        } 

        else if (Units >= 200)

        {

             Amount = Units * 5.26;

             Sur_Charge = 55; 

        }  

        else if (Units >= 100)

        {

             Amount = Units * 3.76;

             Sur_Charge = 35; 

        } // Otherwise (fails), Compiler will move to Else block        

        else

        {

                  Amount = Units * 2.25; 

                  Sur_Charge = 25; 

          }

          Total_Amount = Amount + Sur_Charge;

          System.out.println("\n Electricity Bill  =  " + Total_Amount);

     }

}

 Please Enter the Units that you Consumed  : 750


 Electricity Bill  =  7322.5

Let me try a different value.

 Please Enter the Units that you Consumed  : 250


 Electricity Bill  =  1370.0

C++ implementation to calculate the 

// electricity bill 

#include<bits/stdc++.h>

using namespace std;

 

// Function to calculate the 

// electricity bill 

int calculateBill(int units) 

{ 

 

    // Condition to find the charges 

    // bar in which the units consumed 

    // is fall 

    if (units <= 100) 

    { 

        return units * 10; 

    } 

    else if (units <= 200)

    { 

        return (100 * 10) + 

               (units - 100) * 15; 

    } 

    else if (units <= 300)

    { 

        return (100 * 10) + 

               (100 * 15) + 

               (units - 200) * 20; 

    } 

    else if (units > 300)

    { 

        return (100 * 10) + 

               (100 * 15) + 

               (100 * 20) + 

               (units - 300) * 25; 

    } 

    return 0; 

} 

 

// Driver Code 

int main() 

{ 

    int units = 250; 

    cout << calculateBill(units); 

} 

 

// This code is contributed by spp____

put: U = 250 

Output: 3500 

Explanation: 

Charge for the first 100 units   10*100 = 1000 

Charge for the 100 to 200 units   15*100 = 1500 

Charge for the 200 to 250 units   20*50 = 1000 

Total Electricity Bill = 1000 + 1500 + 1000 = 3500

Input: U = 95 

Output: 950 

Explanation: 

Charge for the first 100 units   10*95 = 950 

Total Electricity Bill = 950


B 

' Multiline syntax:

If condition [ Then ]

    [ statements ]

[ ElseIf elseifcondition [ Then ]

    [ elseifstatements ] ]

[ Else

    [ elsestatements ] ]

End If


' Single-line syntax:

If condition Then [ statements ] [ Else [ elsestatements ] ]

he following example illustrates the use of the multiline syntax of the If...Then...Else statement.

VB 

'Create a Random object to seed our starting value 

Dim randomizer As New Random()

'set our variable

Dim count As Integer = randomizer.Next(0, 5)


Dim message As String


'If count is zero, output will be no items

If count = 0 Then

    message = "There are no items."

    'If count is 1, output will be "There is 1 item.".        

ElseIf count = 1 Then

    message = "There is 1 item."

    'If count is greater than 1, output will be "There are {count} items.", where {count} is replaced by the value of count. 

Else

    message = $"There are {count} items."

End If


Console.WriteLine(message)


'This example displays output like the following:

' There are 4 items.

Nested syntax example

[1]5 €¥[1]Tahoma[1]
ì      §[1]D(3€metering subject air time  dhet /st peace subject  0=ö 
________________________________________
[1]5 €¥[1]Tahomaå €4TextBox3Î¢,å
VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm5 
   Caption         =   "UserForm5"
   ClientHeight    =   7425
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11385
   OleObjectBlob   =   "UserForm5 printer departmennt.frx":0000
   StartUpPosition =   3  'Windows Default
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
End
Attribute VB_Name = "UserForm5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
Private Sub ComboBox1_Change()
 
End Sub
 
Private Sub CommandButton1_Click()
 
End Sub
 
Private Sub Frame1_Click()
 
End Sub
 
Private Sub Label1_Click()
 
End Sub
 
Private Sub ScrollBar1_Change()
 
End Sub
 
Private Sub TextBox1_Change()
 
End Sub
 
Private Sub UserForm_Click()
 
End Sub
 
Private Sub UserForm_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
 
End Sub
 
Private Sub UserForm_Initialize()
 
End Sub
 
Private Sub UserForm_Layout()
 
End Sub
 
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
 
End Sub
 
Private Sub UserForm_RemoveControl(ByVal Control As MSForms.Control)
 
End Sub
 
Private Sub UserForm_Scroll(ByVal ActionX As MSForms.fmScrollAction, ByVal ActionY As MSForms.fmScrollAction, ByVal RequestDx As Single, ByVal RequestDy As Single, ByVal ActualDx As MSForms.ReturnSingle, ByVal ActualDy As MSForms.ReturnSingle)
 
End Sub
 
Private Sub UserForm_Terminate()
 
End Sub
 
Private Sub UserForm_Zoom(Percent As Integer)
 
End Sub 
LB |Ó,¥ ÐÏ à¡±á> 
________________________________________ Rã
‘ Î ãªK¸Q DB Tahomaà…(õ €[1]2p Label1r " 
________________________________________
O 
________________________________________
$å € 
________________________________________
4TextBox1O 
________________________________________
À
(å  € 
________________________________________
`c[1]ComboBox1 ëO 
________________________________________
¬ ,å
€8 
________________________________________
CommandButton1T" 
________________________________________
Ä(å
€ 
________________________________________
/ScrollBar1to" 
________________________________________
[1]P(=€pdf converter subject   equivalente subject dhet    /printer €ùBÉ[1]5 €¥[1]Tahoma[1] €H€,ÌCD [1]5 €¥[1]Tahoma[1]@cAE€H€, 
________________________________________
[1]!c€ùBÉVB 
Public Sub Main()
    ' Run the function as part of the WriteLine output.
    Console.WriteLine("Time Check is " & CheckIfTime() & ".")
End Sub
 
Private Function CheckIfTime() As Boolean
    ' Determine the current day of week and hour of day.
    Dim dayW As DayOfWeek = DateTime.Now.DayOfWeek
    Dim hour As Integer = DateTime.Now.Hour
 
    ' Return True if Wednesday from 2 to 3:59 P.M.,
    ' or if Thursday from noon to 12:59 P.M.
    If dayW = DayOfWeek.Wednesday Then
        If hour = 14 Or hour = 15 Then
            Return True
        Else
            Return False
        End If
    ElseIf dayW = DayOfWeek.Thursday Then
        If hour = 12 Then
            Return True
        Else
            Return False
        End If
    Else
        Return False
    End If
End Function
 
'This example displays output like the following:
'Time Check is False.
 
Private Sub SingleLine()
 
    'Create a Random object to seed our starting values 
    Dim randomizer As New Random()
 
    Dim A As Integer = randomizer.Next(10, 20)
    Dim B As Integer = randomizer.Next(0, 20)
    Dim C As Integer = randomizer.Next(0, 5)
 
    'Let's display the initial values for comparison
    Console.WriteLine($"A value before If: {A}")
    Console.WriteLine($"B value before If: {B}")
    Console.WriteLine($"C value before If: {C}")
 
    ' If A > 10, execute the three colon-separated statements in the order
    ' that they appear
    If A > 10 Then A = A + 1 : B = B + A : C = C + B
 
    'If the condition is true, the values will be different
    Console.WriteLine($"A value after If: {A}")
    Console.WriteLine($"B value after If: {B}")
    Console.WriteLine($"C value after If: {C}")
 
End Sub
 
'This example displays output like the following:
'A value before If: 11
'B value before If: 6
'C value before If: 3
'A value after If: 12
Option Compare Binary
 
Console.WriteLine("A" < "a")
' Output: True
Option Compare Text
 
Console.WriteLine("A" = "a")
' Output: True
 
 
 
  Declare a one-dimensional Integer array and a Single 
' variable.
Dim philaMint(5) As Integer 
Dim x As Single
x! = 10.0
philaMint%(0) = 3                 ' Assigns an Integer value
philaMint%(1) = x                 ' Converts Single 10.0 to Integer 10
Print DataType(philaMint%(0)); DataType(philaMint%(1))
' Output:
' 2  2
' Both values are Integers.
eclare marCell and perDue as Integer variables.
' The phrase As Integer declares marCell as an Integer
' variable. The data type suffix % declares perDue as an
' Integer variable.
Dim marCell As Integer, perDue%
Print TypeName(marCell), TypeName(perDue%)    ' Prints INTEGER INTEGER
Dim marCell% As Integer
' Error, because the Dim statement attempts to declare
' the Integer variable marCell using both the data type
' suffix character for Integer, and the data type name
' Integer. The declaration should include one or the 
' other, but not both.
' A data type suffix character is optional in references to a 
' declared variable.
' Declare marCell as an Integer variable.
Dim marCell As Integer
' Use the data type suffix character in a reference to marCell.
marCell% = 1
' Refer to marCell without using the suffix character.
Print marCell                        ' Prints 1
 
The Visual Basic printer object greatly simplifies sending output to a printer. The following is a very simple example of creating a PDF and specifying the output file name.
 
Private Sub Command1_Click()
 'set the output file name
 SaveSetting "Dane Prairie Systems", "Win2PDF", "PDFFileName", "c:\TEST.PDF"
 
 'output some text
 Printer.Print "hello world!"
 
 'save the PDF file
 Printer.EndDoc
End Sub
 
Win2PDF supports all of the standard Printer object methods and properties. A more advanced example requires some support routines for writing a Dword value to the registry. The following code must be added to the declarations section of the sample form.
 
Option Explicit
 
Private Const REG_DWORD As Long = 4
Private Const HKEY_CURRENT_USER = &H80000001
Private Const KEY_ALL_ACCESS = &H3F
 
Private Declare Function RegCloseKey Lib "advapi32.dll" _
           (ByVal hKey As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias _
  "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, _
  ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As _
  Long) As Long
Private Declare Function RegSetValueExLong Lib "advapi32.dll" Alias _
  "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, _
  ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, _
  ByVal cbData As Long) As Long
 
Private Sub SaveWin2PDFDword(sValueName As String, lValueSetting As Long)
 Dim lRetVal As Long     'result of the SetValueExLong function
 Dim hKey As Long      'handle of open key
 
 lRetVal = RegOpenKeyEx(HKEY_CURRENT_USER, "Software\VB and VBA Program Settings\Dane Prairie Systems\Win2PDF", 0, _
             KEY_ALL_ACCESS, hKey)
 lRetVal = RegSetValueExLong(hKey, sValueName, 0&, _
             REG_DWORD, lValueSetting, 4)
 RegCloseKey (hKey)
End Sub
 
Now we can use the subroutine to save a Dword value to the registry and enable 128 bit encryption. The encryption options are only valid for Win2PDF Pro.
 
Private Sub Command1_Click()
 'set the output file name
 SaveSetting "Dane Prairie Systems", "Win2PDF", "PDFFileName", "c:\TEST.PDF"
 
 'enable 128 bit encryption (Win2PDF Pro only) and automatic url detection
 SaveWin2PDFDword "file options", &H30
 
 'set a master password
 SaveSetting "Dane Prairie Systems", "Win2PDF", "master password", "abracadabra"
 
 'set the PDF document title
 SaveSetting "Dane Prairie Systems", "Win2PDF", "PDFTitle", "Hello World sample"
 
 'set landscape orientation
 Printer.Orientation = vbPRORLandscape
 
 'output some text
 Printer.Print "hello world!"
 Printer.Print "www.win2pdf.com"
 
 'save the PDF file
 Printer.EndDoc
End Sub
 
The following example demonstrates how to append to an existing PDF file.
 
Private Sub Command2_Click()
 
 'set the output file name to create original PDF file
 SaveSetting "Dane Prairie Systems", "Win2PDF", "PDFFileName", "c:\TEST.PDF"
 
 'output some text
 Printer.Print "hello world!"
 
 'save the PDF file
 Printer.EndDoc
 
 
 'Append new text to original PDF file
 'assign original PDF file
 SaveSetting "Dane Prairie Systems", "Win2PDF", "PDFAppendFile", "c:\TEST.PDF"
 
 'set new output file to maintain the original PDF file
 SaveSetting "Dane Prairie Systems", "Win2PDF", "PDFFileName", "c:\AppendedTest.PDF"
 
 Printer.Print "2nd page added"
 Printer.EndDoc
code. The following Visual Basic 6.0 code example prints four lines of text with different font settings:
Private Sub Form_Load() With Printer
Printer.Print "Normal Line" .FontBold = True Printer.Print "Bold Line" .FontItalic = True
Printer.Print "Bold and Italic Line" .FontBold = False .FontItalic = False Printer.Print "Second Normal Line" .EndDoc End With End Sub
Dim WithEvents prn As New Printing.PrintDocument
Private Sub Form1_Load(ByVal eventSender As System.Object, _ ByVal eventArgs As System.EventArgs) Handles MyBase.Load prn.Print() End Sub
Private Sub prn_PrintPage(ByVal sender As Object, _
ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles prn.PrintPage Dim currFont As Font
Dim yPos As Integer yPos = 1
With e.Graphics currFont = New Font("Arial", 10, FontStyle.Regular) .DrawString("Normal Line", currFont, Brushes.Black, 1, yPos) yPos = yPos + currFont.GetHeight currFont = New Font("Arial", 10, FontStyle.Bold) .DrawString("Bold Line", currFont, Brushes.Black, 1, yPos) yPos = yPos + currFont.GetHeight currFont = New Font("Arial", 10, FontStyle.Bold Or FontStyle.Italic) .DrawString("Bold and Italic Line", currFont, Brushes.Black, 1, yPos) yPos = yPos + currFont.GetHeight currFont = New Font("Arial", 10, FontStyle.Regular) .DrawString("Second Normal Line", currFont, Brushes.Black, 1, yPos) End With End Sub
Public Class PrinterClass
Public Enum FillStyleConstants As Short vbFSSolid = 0 vbFSTransparent = 1 End Enum
' Scale mode Constants Public Enum ScaleModeConstants As Short vbTwips = 1
vbPixels = 3 End Enum
Private Structure LineInfo
Dim pen As System.Drawing.Pen Dim pl, p2 As Drawing.Point End Structure
Private Structure RectangleInfo Dim pen As System.Drawing.Pen Dim rec As Drawing.Rectangle Dim FillStyle As FillStyleConstants Dim FillColor As System.Drawing.Color End Structure
Private Structure PageInfo
Public Lines() As LineInfo Public Circ1es() As RectangleInfo End Structure
Private Pages() As PageInfo Private PageIndex = 0
Private WithEvents InnerPrinter As New System.Drawing.Printing.PrintDocument Private objPen As New System.Drawing.Pen(System.Drawing.Brushes.Black) Private intCurrentX As Double = 20 Private intCurrentY As Double = 20
Public ScaleMode As ScaleModeConstants = ScaleModeConstants.vbTwips Public FillColor As Drawing.Color
Public FillStyle As FillStyleConstants = FillStyleConstants.vbFSTransparent
Public Sub New() ReDim Pages(0) Pages(0) = New PageInfo End Sub
' Terminates a print operation sent to the Printer object, ' releasing the document to the print device or spooler. Public Sub EndDoc() Dim j As Integer For j = 0 To Pages.Length - l PageIndex = j InnerPrinter.Print()
Next End Sub
Public Sub Circ1e(ByVa1 P As Drawing.Point, ByVal radius As Double, _ Optional ByVal [Step] As Boolean = False) Circ1e(P, radius, objPen.Color, [Step]) End Sub
Public Sub Circ1e(ByVa1 P As Drawing.Point, ByVal radius As Double, _ ByVal Color As Drawing.Color, _ Optional ByVal [Step] As Boolean = False)
Dim diameter As Double
P.X = P.X + CurrentX P.Y = P.Y + CurrentY End If diameter = radius * 2
' Moves the CurrentX and CurrentY properties CurrentX = P.X + radius CurrentY = P.Y + radius
P.X = ConvertToPixelsX(P.X) P.Y = ConvertToPixelsY(P.Y) diameter = ConvertToPixelsX(diameter)
If IsNothing(Pages(PageIndex).Circles) Then ReDim Preserve Pages(PageIndex).Circles(0)
Else
ReDim Preserve
Pages(PageIndex).Circles(Pages(PageIndex).Circles.Length) End If
Pages(PageIndex).Circles(Pages(PageIndex).Circles.Length
New Drawing.Rectangle(P.X, P.Y, diameter, diameter) Pages(PageIndex).Circles(Pages(PageIndex).Circles.Length Pages(PageIndex).Circles(Pages(PageIndex).Circles.Length
FillColor
Pages(PageIndex).Circles(Pages(PageIndex).Circles.Length FillStyle
End Sub
Private Sub Printer_PrintPage(ByVal sender As Object, _
ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles InnerPrinter.PrintPage Dim i As Integer ' Draw all circles
If IsNothing(Pages(PageIndex).Circles) = False Then For i = 0 To Pages(PageIndex).Circles.Length - 1 Dim x, y As Integer e.Graphics.DrawEllipse(Pages(PageIndex).Circles(i).pen,
Pages(PageIndex).Circles(i).rec) If Pages(PageIndex).Circles(i).FillStyle = _ FillStyleConstants.vbFSSolid Then e.Graphics.FillEllipse( _
New Drawing.SolidBrush( _ Pages(PageIndex).Circles(i).FillColor), _ Pages(PageIndex).Circles(i).rec)
End If
Next End If End Sub
Public Property CurrentX() As Double Get
Return ConvertToPixelsX(intCurrentX) End Get
Set(ByVal Value As Double)
intCurrentX = ConvertToPixelsX(Value) End Set
End Property
Public Property CurrentY() As Double Get
Return ConvertToPixelsY(intCurrentY) End Get
Set(ByVal Value As Double)
intCurrentY = ConvertToPixelsY(Value) End Set End Property
Public Function ConvertToPixelsX(ByVal num As Double) As Double Return IIf(ScaleMode = ScaleModeConstants.vbTwips, _ VB6.TwipsToPixelsX(num), num)
End Function
Public Function ConvertToPixelsY(ByVal num As Return IIf(ScaleMode = ScaleModeConstants VB6.TwipsToPixelsY(num), num)
End Function End Class
printer class. Here is the resulting Visual Basic .NET code.
Printer.FillColor = Color.Blue
Printer.FillStyle = PrinterClass.FillStyleConstants.vbFSSolid Printer.Circle(New Point(3000, 3000), 1500, Color.Red) Printer.EndDoc()
Double) As Double .vbTwips, _
Table 7.3: Visual Basic .NET Equivalents for Printer Object Properties Visual Basic 6.0 Visual Basic .NET Equivalent
ColorMode   PrintDocument.PrinterSettings.DefaultPageSettings.Color
 
Copies      PrintDocument.PrinterSettings.Copies
CurrentX    No equivalent. Replaced by location and dimension arguments of various methods of the Graphics class.This property could be simulated in your printer class with a double variable.
CurrentY    No equivalent. Replaced by location and dimension arguments of various methods of the Graphics class.This property could be simulated in your printer class with a double variable.
DeviceName  PrintDocument.PrinterSettings.PrinterName
 
DrawMode    No equivalent. For details, see "Graphics Changes in Visual Basic .NET" on MSDN.
DrawStyle   System.Drawing.Drawing2D.DashStyle
DrawWidth   System.Drawing.Pen.Width
DriverName  No equivalent. No longer needed; printer drivers are managed by Windows.
 
Duplex      System.Drawing.Printing.Duplex
FillColor   No equivalent. For details, see "Graphics Changes in Visual Basic .NET" on MSDN.This property could be simulated in your printer class with a System.Drawing.Color variable.
VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm6 
   Caption         =   "UserForm6"
   ClientHeight    =   7500
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13785
   OleObjectBlob   =   "UserForm6 st peace plc data dhet.frx":0000
   StartUpPosition =   3  'Windows Default
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
End
Attribute VB_Name = "UserForm6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
 
End Sub
 
Private Sub CommandButton2_Click()
 
End Sub
 
Private Sub CommandButton3_Click()
 
End Sub
 
Private Sub Frame1_Click()
 
End Sub
 
Private Sub Label2_Click()
 
End Sub
 
Private Sub Label3_Click()
 
End Sub
 
Private Sub Label4_Click()
 
End Sub
 
Private Sub TextBox1_Change()
 
End Sub
 
Private Sub TextBox2_Change()
 
End Sub
 
Private Sub TextBox3_Change()
 
End Sub
 
Private Sub TextBox3_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
 
End Sub
 
Private Sub TextBox3_DropButtonClick()
 
End Sub
 
Private Sub TextBox3_Enter()
 
End Sub
 
Private Sub TextBox3_Exit(ByVal Cancel As MSForms.ReturnBoolean)
 
End Sub
 
Private Sub TextBox3_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
 
End Sub
 
Private Sub TextBox3_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
 
End Sub
 
Private Sub TextBox3_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
 
End Sub
 
Private Sub TextBox3_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
 
End Sub
 
Private Sub TextBox3_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
 
End Sub
 
Private Sub UserForm_Click()
 
End Sub 
LB 36ð ÐÏ à¡± Tahoma[1]\(K€communication plc data form dhet st peace   engineering electrical data // ˆ;{[1][1]5 €¥[1]Tahoma[1] €H€,—  [1]5 €¥[1]Tahoma[1] €H€,˜ ö 
________________________________________
[1]5 €¥[1]Tahoma[1] €H€,˜ # 
________________________________________
[1]5 €¥[1]Tahoma 
________________________________________
8H


________________________________________
€ 
________________________________________
€ÿÿ
}HF°'Frame1 
________________________________________
Rã
‘ Î ãªK¸Q DB Tahoma
¼Š(õ €[1]2@ Label1ö 
________________________________________
ö 
________________________________________
(õ € 
________________________________________
2@ Label2ö 
________________________________________


(õ € 
________________________________________
24[1] Label3" 
________________________________________
',å
€@ 
________________________________________
CommandButton1 O 
________________________________________
k¬,å
€ < 
________________________________________
CommandButton2 ¢k¬,å
€@ CommandButton3 á9ö 
________________________________________
(õ € 2| Label4O 
________________________________________
Ô$å €      4TextBox1¬ Ê$å €
4 TextBox2 ¶$å €
4     TextBox3 ú$Õ €# 
________________________________________

Fram.                          


	
VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   10695
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15495
   OleObjectBlob   =   "UserForm1LICENSE CERTIFICATE ISSUE,.frx":0000
   StartUpPosition =   3  'Windows Default
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ComboBox1_Change()
 
End Sub
 
Private Sub CommandButton1_Click()
 
End Sub
 
Private Sub Frame1_Click()
 
End Sub
 
Private Sub Label1_Click()
 
End Sub
 
Private Sub ScrollBar1_Change()
 
End Sub
 
Private Sub TextBox1_Change()
 
End Sub
 
Private Sub UserForm_Click()
 
End Sub
 
Private Sub UserForm_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
 
End Sub
 
Private Sub UserForm_Error(ByVal Number As Integer, ByVal Description As MSForms.ReturnString, ByVal SCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, ByVal CancelDisplay As MSForms.ReturnBoolean)
 
End Sub
 
Private Sub UserForm_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
 
End Sub
 
Private Sub UserForm_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
 
End Sub
 
Private Sub UserForm_RemoveControl(ByVal Control As MSForms.Control)
 
End Sub
 
Private Sub UserForm_Resize()
 
End Sub
 
Private Sub UserForm_Scroll(ByVal ActionX As MSForms.fmScrollAction, ByVal ActionY As MSForms.fmScrollAction, ByVal RequestDx As Single, ByVal RequestDy As Single, ByVal ActualDx As MSForms.ReturnSingle, ByVal ActualDy As MSForms.ReturnSingle)
 
End Sub
 
Private Sub UserForm_Terminate()
 
End Sub
 
Private Sub UserForm_Zoom(Percent As Integer)
 
End Sub

to tshingombe, tshigombekb, TSHINGOMBEKB, me 
 

LB &á<k+ÐÏ à¡±á> 
________________________________________
þÿ      
________________________________________
þÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿýÿÿÿþÿÿÿþÿÿÿþÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿRoot Entryÿÿÿÿÿÿÿÿ[1]ði*ÆÜÎ ž˜ªWJO 
________________________________________
@f 
________________________________________
[1]ÿÿÿÿÿÿÿÿÿÿÿÿ2o 
________________________________________
[1] 
________________________________________
ÿÿÿÿþÿÿÿi01 ÿÿÿÿÿÿÿÿÿÿÿÿF¨ÃÌùÙF¨ÃÌùÙþÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿ 
________________________________________
,
________________________________________
€
   [1] 
________________________________________

}Ãj±IRoot Entryÿÿÿÿÿÿÿÿ[1]ði*ÆÜÎ ž˜ªWJO{´ÃÌùÙ À f 
________________________________________
[1]ÿÿÿÿÿÿÿÿÿÿÿÿfo 
________________________________________
[1] 
________________________________________
ÿÿÿÿþÿÿÿi01 ÿÿÿÿ
  n`ôÎ ›Íª`ŽF¨ÃÌùÙ0”¨ÃÌùÙÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿ ýÿÿÿ     þÿÿÿ
þÿÿÿ



þÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿf 
________________________________________
[1]ÿÿÿÿÿÿÿÿÿÿÿÿ2mo 
________________________________________
[1] 
________________________________________
ÿÿÿÿ 
________________________________________
¤
i07 ÿÿÿÿ    n`ôÎ ›Íª`ŽF¨ÃÌùÙ0”¨ÃÌùÙCompObj [1]ÿÿÿÿÿÿÿÿÿÿÿÿpD[1]þÿÿÿ 
________________________________________
      




 ¬ !"#$%&'()*+,-./01þÿÿÿ34567þÿÿÿ9þÿÿÿ;<=>þÿÿÿ@ABCþÿÿÿþÿÿÿFþÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿ 
________________________________________
,
________________________________________
€
   [1] 
________________________________________

}Ãj±I,amþÿ 
________________________________________

ÿÿÿÿ  n`ôÎ ›Íª`ŽMicrosoft Forms 2.0 FrameEmbedded Object
Forms.Frame.1ô9²q[1] 
________________________________________
*
€ð€LICENSE TRADEMARK  / ISSUE CERTIFICATE PRODUCT  RESERVE COPYRITH   , EDUCATION SOFTWAR PRODUCT DESIGN IMPLEMENTATION , CODE SOURCE BUTTON PRODUCT ,OPERATOR SYSTEM ISSUE NAME CERTIFICATE PRODUCT  , SCIPTOR , STYLE EDITOR , COPYRITH  , RULES ¥WÊ[1]5 €¥[1]Tahoma[1] €H€,¯R [1]f 
________________________________________
[1]ÿÿÿÿÿÿÿÿÿÿÿÿ? o 
________________________________________
[1] 
ÿÿÿÿ:,CompObj [1]ÿÿÿÿÿÿÿÿÿÿÿÿ8pCompObj [1]ÿÿÿÿÿÿÿÿÿÿÿÿEn5 €¥[1]Tahoma[1]ä      AE€H€, 
________________________________________
[1]È VTD Portofolio Evidence Low Dhet Saqa Su


by

    texbook eductec



editEdit

gearManage

timeHistory


Publication date

    2022-12-24 


Usage

    Public Domain Mark 1.0Creative Commons Licensepublicdomain


Topics

    engineering


Collection

    opensource


Language

    English


portofoli


Addeddate

    2022-12-24 08:37:07


Eductec

    5000


Identifier

    portofolio-evidence-low-dhet-saqa-su


Identifier-ark

    ark:/13960/s2bkt55k02s


Ocr

    tesseract 5.3.0-3-g9920


Ocr_detected_lang

    en


Ocr_detected_lang_conf

    1.0000


Ocr_detected_script

    Latin


Ocr_detected_script_conf

    0.9856


Ocr_module_version

    0.0.19


Ocr_parameters

    -l eng


Page_number_confidence

    92.17


plus-circle Add Review

comment

Reviews

Reviewer: Rdferz - favoritefavoritefavoritefavoritefavorite - December 24, 2022

Subject: edutech



________________________________________
 


Author : tshingombe tshitadi

Author: tshingombe Tshitadi Fiston

Report school: 2022/

ISBN :20222

Title : guidelines practique theory engineering circular children's youth technical institute base college

Volume:

Copyright: reserved rights

Legal archive : ftjr / ft JA / ...ftgaek f 24/ 22 .223 work skill development record transcript





[1]5 €¥[1]Tahoma[1](€ACTIVE 1 N 
________________________________________
[1]u €¥[1] 
________________________________________
Tahoma [1]
QO 
________________________________________
________________________________________
8H


________________________________________
€ 
________________________________________
€ÿÿ
}Ç[y-Frame1 
________________________________________
Rã
‘ Î ãªK¸Q DB Tahoma †m(õ €[1]2$ Label1 Ê{[1]$å € 
________________________________________
4TextBox1ö 
________________________________________
S(å   € 
________________________________________
________________________________________

[1]ComboBox1L7
" 
________________________________________
˜ ,å
€8 
________________________________________
CommandButton1G
ö 
________________________________________
„((å
€ 
________________________________________
/ScrollBar1to" 
________________________________________
a$$Õ €# 
________________________________________

Frame2x1Êì    þÿ 
________________________________________

ÿÿÿÿ  n`ôÎ ›Íª`ŽMicrosoft Forms 2.0 FrameEmbedded Object
Forms.Frame.1ô9²q[1]l(Z€DHET /ST PEACE  ISSUE NAME ,LICENSE NAME CERTIFICATE NAME , NAME INFO SYSTEM STUDY MANAGE,¹MN 
________________________________________
[1]5 €¥[1]Tahoma [1] €H€,;
{[1][1]5 €¥[1]Tahoma [1] €H€,g
„ 
________________________________________
[1]5 €¥[1]Tahoma [1]A€H€, 
________________________________________
[1]g
# 
________________________________________
[1]5 €¥[1]Tahoma 
________________________________________
8H


________________________________________
€ 
________________________________________
€ÿÿ 
________________________________________
}ÛQ“
Frame2 
________________________________________
Rã
‘ Î ãªK¸Q DB Tahoma 
________________________________________
¬„o(õ € 2Œ Label2§Ô$å €      4TextBox2] ö 
________________________________________
$å €
4[1]TextBox3Ž#í 
________________________________________
(å €
8 
________________________________________
ComboBox2C3" 
________________________________________
$Õ €# 
________________________________________

Frame1x3  E þÿ 
________________________________________

ÿÿÿÿði*ÆÜÎ ž˜ªWJOMicrosoft Forms 2.0 FormEmbedded Object
Forms.Form.1ô9²q

On Sun, Oct 8, 2023 at 11:55 AM tshingombe fiston <tshingombefiston@gmail.com> wrote:
VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   10695
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15495
   OleObjectBlob   =   "UserForm1LICENSE CERTIFICATE ISSUE,.frx":0000
   StartUpPosition =   3  'Windows Default
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ComboBox1_Change()
 
End Sub
 
Private Sub CommandButton1_Click()
 
End Sub
 
Private Sub Frame1_Click()
 
End Sub
 
Private Sub Label1_Click()
 
End Sub
 
Private Sub ScrollBar1_Change()
 
End Sub
 
Private Sub TextBox1_Change()
 
End Sub
 
Private Sub UserForm_Click()
 
End Sub
 
Private Sub UserForm_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
 
End Sub
 
Private Sub UserForm_Error(ByVal Number As Integer, ByVal Description As MSForms.ReturnString, ByVal SCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, ByVal CancelDisplay As MSForms.ReturnBoolean)
 
End Sub
 
Private Sub UserForm_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
 
End Sub
 
Private Sub UserForm_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
 
End Sub
 
Private Sub UserForm_RemoveControl(ByVal Control As MSForms.Control)
 
End Sub
 
Private Sub UserForm_Resize()
 
End Sub
 
Private Sub UserForm_Scroll(ByVal ActionX As MSForms.fmScrollAction, ByVal ActionY As MSForms.fmScrollAction, ByVal RequestDx As Single, ByVal RequestDy As Single, ByVal ActualDx As MSForms.ReturnSingle, ByVal ActualDy As MSForms.ReturnSingle)
 
End Sub
 
Private Sub UserForm_Terminate()
 
End Sub
 
Private Sub UserForm_Zoom(Percent As Integer)
 
End Sub

...

[Message clipped]  View entire message
 
tshingombe fiston 
	11:57 AM (0 minutes ago)

	
	
to tshingombe, tshigombekb, TSHINGOMBEKB, me 
 

VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "UserForm2"
   ClientHeight    =   7410
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14355
   OleObjectBlob   =   "UserForm2 LICENSE CERTIFICATE ISSUE  DHET.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ComboBox1_Change()
 
End Sub
 
Private Sub CommandButton1_Click()
 
End Sub
 
Private Sub CommandButton2_Click()
 
End Sub
 
Private Sub Frame1_Click()
 
End Sub
 
Private Sub Label1_Click()
 
End Sub
 
Private Sub Label2_Click()
 
End Sub
 
Private Sub Label3_Click()
 
End Sub
 
Private Sub Label4_Click()
 
End Sub
 
Private Sub TextBox1_Change()
 
End Sub
 
Private Sub TextBox2_Change()
 
End Sub
 
Private Sub TextBox3_Change()
 
End Sub
 
Private Sub UserForm_Click()
 
End Sub
 
Private Sub UserForm_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
 
End Sub
 
Private Sub UserForm_Deactivate()
 
End Sub
 
Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
 
End Sub
 
Private Sub UserForm_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
 
End Sub
 
Private Sub UserForm_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
 
End Sub
 
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
 
End Sub
 
Private Sub UserForm_RemoveControl(ByVal Control As MSForms.Control)
 
End Sub
 
Private Sub UserForm_Scroll(ByVal ActionX As MSForms.fmScrollAction, ByVal ActionY As MSForms.fmScrollAction, ByVal RequestDx As Single, ByVal RequestDy As Single, ByVal ActualDx As MSForms.ReturnSingle, ByVal ActualDy As MSForms.ReturnSingle)
 
End Sub
 
Private Sub UserForm_Terminate()
 
End Sub
 
Private Sub UserForm_Zoom(Percent As Integer)
 
End Sub 
LB m8– ÐÏ à¡±á> 
________________________________________
þÿ      
________________________________________
þÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿýÿÿÿþÿÿÿþÿÿÿþÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿRoot Entryÿÿÿÿÿÿÿÿ[1]ði*ÆÜÎ ž˜ªWJO 
________________________________________
@f 
________________________________________
[1]ÿÿÿÿÿÿÿÿÿÿÿÿ&o 
________________________________________
[1] 
________________________________________
ÿÿÿÿþÿÿÿi01 ÿÿÿÿÿÿÿÿÿÿÿÿ°ðFÌùÙ°ðFÌùÙþÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿ 
________________________________________
  



}éb
3Root Entryÿÿÿÿÿÿÿÿ[1]ði*ÆÜÎ ž˜ªWJOŸQÌùÙ f 
________________________________________
[1]ÿÿÿÿÿÿÿÿÿÿÿÿZo 
________________________________________
[1] 
________________________________________
ÿÿÿÿþÿÿÿi01 ÿÿÿÿ  n`ôÎ ›Íª`Ž°ðFÌùÙ°ðFÌùÙÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿ ýÿÿÿþÿÿÿþÿÿÿ      

þÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿf 
________________________________________
[1]ÿÿÿÿÿÿÿÿÿÿÿÿ %[1]o 
________________________________________
[1] 
________________________________________
ÿÿÿÿ 
________________________________________
________________________________________
________________________________________
CompObj [1]ÿÿÿÿÿÿÿÿÿÿÿÿpCompObj [1]ÿÿÿÿÿÿÿÿÿÿÿÿ n[1]þÿÿÿ 
________________________________________
      




þÿÿÿ þÿÿÿþÿÿÿ-þÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿ 
________________________________________
  



}éb
3,xt$Õ €þÿ 
________________________________________

ÿÿÿÿ  n`ôÎ ›Íª`ŽMicrosoft Forms 2.0 FrameEmbedded Object
Forms.Frame.1ô9²q[1]@(-€OFFICE PROTECTON TRADE SYSTEM ,MS DOS, WINDOS*Uö 
________________________________________
[1]5 €¥[1]Tahoma[1]H(5€OFFICE PROTECTION ,MS WORDD,MS EXCELL MS ACCESS,MS VBÏ ýU [1]5 €¥[1]Tahoma[1]@(.€OFFICE PROTECTION , WEB INTERNET , BREAK JAVA ÑVž [1]5 €¥[1]Tahoma[1]h(V€CERTIFICATE ,LICENSE TERM ISSUE  CODE , BOOKING  ,SYST MANAGENT POLICY , SAFETY  TERM ÑVE [1]5 €¥[1]Tahoma[1]L@€H€,2€98„ 
________________________________________
'' STACK MERGE LINE TOOLS BREAK '' PAGE COPYRITH  [1]5 €¥[1]Tahoma[1]0@€H€,€e74 PROTECTION FRAME FORM [1]5 €¥[1]Tahoma[1] €H€,98Ê[1]5 €¥[1]Tahoma[1]ÄAE€H€, 
________________________________________
[1]¨€e7Ü 
________________________________________
HELP MEMORIES ISSUE , CERTIFICATE ISSUE, LICENSE ISSUE OPERATOR CODE SOURCE,  ISSUE , IMPLENTATION TASK, RULES, SIZE , CUSTOMER CLIENT TOP, SIZE,  ,ORDER SALE RESALING [1]5 €¥[1]Tahoma[1] (
€CommandButton1;

________________________________________
" 
________________________________________
[1]u €¥[1] 
________________________________________
Tahoma[1] (
€CommandButton2;
" 
________________________________________
[1]u €¥[1] 
________________________________________
Tahoma 
________________________________________
@Ê


€
________________________________________
€[1]
€ €ÿÿ
} ZÓ+Frame1 
________________________________________
Rã
‘ Î ãªK¸Q DB Tahoma
¸Š(õ €[1]2` Label1{[1]ö 
________________________________________
(õ € 
________________________________________
2h Label2{[1]g
(õ € 
________________________________________
2`[1] Label3§ (õ €2ˆ 
________________________________________
Label4§ð $å € l 
________________________________________
TextBox1˜ <$å €PTextBox2˜ ¤
$å € 4 TextBox3Ä (å      €   äComboBox1tto˜  ¬,å
€
@ CommandButton1;
O 
________________________________________
  &,å
€
@     CommandButton2;
ð   &# 
________________________________________

Frame1ö 
________________________________________
" 
________________________________________
þÿ 
________________________________________

ÿÿÿÿði*ÆÜÎ ž˜ªWJOMicrosoft Forms 2.0 FormEmbedded Object
Forms.Form.1ô9²q
...

VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   5445
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15390
   OleObjectBlob   =   "UserForm1 vb excell policy marsheet.frx":0000
   StartUpPosition =   1  'CenterOwner
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ComboBox1_Change()
 
End Sub
 
Private Sub CommandButton1_Click()
 
End Sub
 
Private Sub Frame1_Click()
 
End Sub
 
Private Sub Label1_Click()
 
End Sub
 
Private Sub UserForm_Click()
 
End Sub
 
Private Sub UserForm_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
 
End Sub
 
Private Sub UserForm_Error(ByVal Number As Integer, ByVal Description As MSForms.ReturnString, ByVal SCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, ByVal CancelDisplay As MSForms.ReturnBoolean)
 
End Sub
 
Private Sub UserForm_Layout()
 
End Sub
 
Private Sub UserForm_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
 
End Sub
 
Private Sub UserForm_RemoveControl(ByVal Control As MSForms.Control)
 
End Sub
 
Private Sub UserForm_Resize()
 
End Sub
 
Private Sub UserForm_Terminate()
 
End Sub
 
Private Sub UserForm_Zoom(Percent As Integer)
 
End Sub
to tahitaditshingombe, tshingombe, tshigombekb, me 
 

VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   5445
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15390
   OleObjectBlob   =   "UserForm1 vb excell policy marsheet.frx":0000
   StartUpPosition =   1  'CenterOwner
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ComboBox1_Change()
 
End Sub
 
Private Sub CommandButton1_Click()
 
End Sub
 
Private Sub Frame1_Click()
 
End Sub
 
Private Sub Label1_Click()
 
End Sub
 
Private Sub UserForm_Click()
 
End Sub
 
Private Sub UserForm_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
 
End Sub
 
Private Sub UserForm_Error(ByVal Number As Integer, ByVal Description As MSForms.ReturnString, ByVal SCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, ByVal CancelDisplay As MSForms.ReturnBoolean)
 
End Sub
 
Private Sub UserForm_Layout()
 
End Sub
 
Private Sub UserForm_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
 
End Sub
 
Private Sub UserForm_RemoveControl(ByVal Control As MSForms.Control)
 
End Sub
 
Private Sub UserForm_Resize()
 
End Sub
 
Private Sub UserForm_Terminate()
 
End Sub
 
Private Sub UserForm_Zoom(Percent As Integer)
 
End Sub

ALB"x<éÐÏà¡±á> 
________________________________________
þÿ      [1] 
________________________________________
þÿÿÿ      ÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿþÿ 
________________________________________

ÿÿÿÿði*ÆÜÎž˜ªWJOMicrosoft Forms 2.0 FormEmbedded Object 
Forms.Form.1ô9²qÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿRoot Entryÿÿÿÿÿÿÿÿ[1]ði*ÆÜÎž˜ªWJO 
________________________________________
@f 
________________________________________
[1]ÿÿÿÿÿÿÿÿÿÿÿÿ2o 
________________________________________
[1] 
________________________________________
ÿÿÿÿþÿÿÿi01ÿÿÿÿÿÿÿÿÿÿÿÿ ÓÄfûÙ ÓÄfûÙþÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿ 
________________________________________
,


 
€ 
________________________________________
    [1] 
€ 
________________________________________
}
j„%Root Entryÿÿÿÿÿÿÿÿ[1]ði*ÆÜÎž˜ªWJOÀüÖÄfûÙ€[1]f 
________________________________________
[1]ÿÿÿÿÿÿÿÿÿÿÿÿfo 
________________________________________
[1] 
________________________________________
ÿÿÿÿþÿÿÿi01ÿÿÿÿ  n`ôÎ›Íª`Ž ÓÄfûÙPzÓÄfûÙþÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿýÿÿÿþÿÿÿþÿÿÿýÿÿÿ



 ¬ !"#$%&'()*+,-./0123456789:;<=>?@ABCDEFGHIJKLMNOPQRSTUVWXYZ[\]^_`abcdefghijklmnopqrstuvwxyz{|}~€f 
________________________________________
[1]ÿÿÿÿÿÿÿÿÿÿÿÿ 
________________________________________
ño 
________________________________________
[1] 
________________________________________
ÿÿÿÿ

CompObj[1]ÿÿÿÿÿÿÿÿÿÿÿÿpCompObj[1]ÿÿÿÿÿÿÿÿÿÿÿÿn[1]þÿÿÿ 
________________________________________
þÿÿÿþÿÿÿ      þÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿ 
________________________________________
,


 
€ 
________________________________________
    [1] 
€ 
________________________________________
}
j„%,oþÿ 
________________________________________

ÿÿÿÿ  n`ôÎ›Íª`ŽMicrosoft Forms 2.0 FrameEmbedded Object
Forms.Frame.1ô9²q 
________________________________________A

________________________________________
________________________________________
€ 
________________________________________
€ÿÿ 
________________________________________
}›\˜ Frame1 
________________________________________
Rã
‘ÎãªK¸QDBTahoma 
________________________________________
Œƒo(õ€[1]2hLabel1{[1]Ô(å      € 
________________________________________
ð
ComboBox1{[1],å
€ 
________________________________________
8[1]CommandButton1" 
________________________________________
'$Õ€# 
________________________________________

Frame1q{[1]‚ƒ„…†‡ˆ‰Š‹ŒŽþÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿ[1]H(7€issue license  excelll academic ,dhet st peace college *UN 
________________________________________
[1]5€¥[1]Tahoma[1]AE€H€, 
________________________________________
[1]´
ýUã   



Enable desktop notifications for Gmail.   OK  No thanks

1 of 948


metadabse

Inbox


tshingombe fiston <tshingombefiston@gmail.com> 

     12:29/ PM (0 minutes ago)

     

     

to tahitaditshingombe, me, tshingombe, tshigombekb 

LB [1][1]& >1ÐÏ à¡±á> 

________________________________________

þÿ      [1] 

________________________________________

þÿÿÿ      ÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿýÿÿÿþÿÿÿþÿÿÿþÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿRoot Entryÿÿÿÿÿÿÿÿ[1]ði*ÆÜÎ ~Ü[1]ªWJO 

________________________________________

@f 

________________________________________

[1]ÿÿÿÿÿÿÿÿÿÿÿÿ.o 

________________________________________

[1] 

________________________________________

ÿÿÿÿþÿÿÿi01 ÿÿÿÿÿÿÿÿÿÿÿÿ³ébûÙ³ébûÙþÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿ 

________________________________________

(


’ 

¬  [1] 

________________________________________

}¨mN+Root Entryÿÿÿÿÿÿÿÿ[1]ði*ÆÜÎ ~Ü[1]ªWJO[1][1]bûÙ @ 

________________________________________

f 

________________________________________

[1]ÿÿÿÿÿÿÿÿÿÿÿÿbo 

________________________________________

[1] 

________________________________________

ÿÿÿÿþÿÿÿi01 ÿÿÿÿ  n`ôÎ : Íª`}³ébûÙ n÷bûÙÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿ ýÿÿÿþÿÿÿþÿÿÿÿýÿÿÿ




 ¬ !"#$%&'()*+,-./0123456789:;<=>?@ABCDEFGHIJKLMNOPQRSTUVWXYZ[\]^_`abcdefghijklmnopqrstuvwxyz{|}~ ¬ f 

________________________________________

[1]ÿÿÿÿÿÿÿÿÿÿÿÿ 

________________________________________

"!o 

________________________________________

[1] 

________________________________________

ÿÿÿÿ

xèCompObj [1]ÿÿÿÿÿÿÿÿÿÿÿÿpCompObj [1]ÿÿÿÿÿÿÿÿÿÿÿÿ

n

[1]þÿÿÿ 

________________________________________

     þÿÿÿþ 

________________________________________

(


’ 

¬  [1] 

________________________________________

}¨mN+,be$Õþÿ 

________________________________________


ÿÿÿÿ  n`ôÎ : Íª`}Microsoft Forms 2.0 FrameEmbedded Object

Forms.Frame.1ô9²q 

________________________________________

8H



________________________________________

¬  

________________________________________

¬ ÿÿ}_a$Frame1 

________________________________________

Rã

  Î ãªK¸Q DB Tahoma4! (õ ¬ [1]2Ô Label1{[1]Ô(å      ¬  

________________________________________

4ComboBox1¬  §Ê(õ ¬  

________

28[1] Label2§ì      $å ¬ xt 

______________________________________

TextBox1{[1]



ComboBox2¬  §1  ’  &   ! Æ[1]0 `9 R }     "   Ü[1]"!a: S ~x ¡¢£¤¥¦§¨©ª«¬-®¯°±²³´µ

" ¸¹º»¼½¾¿ÀÁÂÃÄÅÆÇÈÉÊËÌÍÎÏÐÑÒÓÔÕÖ×ØÙÚÛÜÝÞßàáâãäåæçèéêëìíîïðñòóôõö÷øùúûüýþþÿÿÿþÿÿÿ[1]´(¢¬ dhet st peace college student markrsheet , assessment  policy circulum license  issue certificate   term   managemnt system information ,student learner  lecture r S" 

________________________________________

[1]5 ¬ ¥[1]Tahoma[1] AE¬ H¬ , 

________________________________________

[1]ø *U   

________________________________________

      �

VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   6285
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15915
   OleObjectBlob   =   "UserForm1 policy assessment circulum  issue certificate.frx":0000
   StartUpPosition =   3  'Windows Default
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
Private Sub ComboBox1_Change()
 
End Sub
 
Private Sub ComboBox2_Change()
 
End Sub
 
Private Sub ComboBox3_Change()
 
End Sub
 
Private Sub Frame1_Click()
 
End Sub
 
Private Sub Label1_Click()
 
End Sub
 
Private Sub Label2_Click()
 
End Sub
 
Private Sub Label3_Click()
 
End Sub
 
Private Sub TextBox1_Change()
 
End Sub
 
Private Sub UserForm_Click()
 
End Sub
 
Private Sub UserForm_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
 
End Sub
 
Private Sub UserForm_Deactivate()
 
End Sub
 
Private Sub UserForm_Error(ByVal Number As Integer, ByVal Description As MSForms.ReturnString, ByVal SCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, ByVal CancelDisplay As MSForms.ReturnBoolean)
 
End Sub
 
Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
 
End Sub
 
Private Sub UserForm_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
 
End Sub
 
Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
 
End Sub
 
Private Sub UserForm_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
 
End Sub
 
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
 
End Sub
 
Private Sub UserForm_RemoveControl(ByVal Control As MSForms.Control)
 
End Sub
 
Private Sub UserForm_Resize()
 
End Sub
 
Private Sub UserForm_Scroll(ByVal ActionX As MSForms.fmScrollAction, ByVal ActionY As MSForms.fmScrollAction, ByVal RequestDx As Single, ByVal RequestDy As Single, ByVal ActualDx As MSForms.ReturnSingle, ByVal ActualDy As MSForms.ReturnSingle)
 
End Sub
 
Private Sub UserForm_Terminate()
 
End Sub 
LB [1][1]…>1ÐÏ à¡±á> 
________________________________________
þÿ      [1] 
________________________________________
DATE CENTRE NAME CENTRE NO     ID      OFFERING SUBJECT   LEVEL ICASS MARK SHEET   CATEGORY IRREGULARITY                    

                              TASK                            

/                                                          

ACTION TAKE                                                              


















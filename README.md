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
VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   8055
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   16950
   OleObjectBlob   =   "UserFormdocument wallet meta.frx":0000
   StartUpPosition =   1  'CenterOwner
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Frame4_Click()
 
End Sub
 
Private Sub Label1_Click()
 
End Sub
 
Private Sub ListBox1_Click()
 
End Sub
 
Private Sub MultiPage1_Change()
 
End Sub
 
Private Sub MultiPage2_Change()
 
End Sub
 
Private Sub ScrollBar1_Change()
 
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
 
Private Sub UserForm_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
 
End Sub
 
Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
 
End Sub
 
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
 
End Sub
 
Private Sub UserForm_Resize()
 
End Sub
 
Private Sub UserForm_Scroll(ByVal ActionX As MSForms.fmScrollAction, ByVal ActionY As MSForms.fmScrollAction, ByVal RequestDx As Single, ByVal RequestDy As Single, ByVal ActualDx As MSForms.ReturnSingle, ByVal ActualDy As MSForms.ReturnSingle)
 
End Sub
 
Private Sub UserForm_Terminate()
 
End Sub
 
Private Sub UserForm_Zoom(Percent As Integer)
Re: Documents wallet Portofilio,office Engineering project order management appointment project file order sale campagny meting tbrigade,edutech psscm... agenda office .PC safety wallet documents financial office


     

     

-1. Purpose: documents wallet Portofilio , documents systeme info., project appointment office Poste 

Agendas office PC safety wallet document wether documents,data base , documents network research office appointment.

- project  month ,order document . Order screen. . Statement documents.






     

     

1.1.Purpose: documents case book .booking  documents  

, financial information documents office , 

-Office document 

-post documents,


--Poste wallet easy documents Poste

 ---office ,post bank , archfile office document , 

Docket..

-Documents arch file 

-Documents stationery document 

-Documents office 

-Documents database system .

-Documents bookeping sale buy ,  

-document  library bibliotech., 

- documents, 

.--- engineering creation document ., 

Post document  office appointment submission close , 

-Price documents, price close tendered ---Poste value  bid posted wallpaper, 

Minimum required Poste tendered 

Spreadsheet,office PowerPoint, -documents ,office post address documents

Relyan documents, 

-Documents memorandum access documents Poste office , documents reading , documents  ,

Project posted 

Appointment post documents wallet report agendas  

Project 

- documents processing  input  output

Wallet  project, management


- documents screen radiographic documents scanographic documents photographic document image documents artistic valeur attributes document pixel coulor ,  documents 

Document monument magazine monaitaitarist ,tableur course, documents, 

Documents note ecriture 

 



     

     

3.purpose . documents wallet Portofilio evidence,,

-Documents wallet registration form appointment office Poste wallet easy.

- register database documents wallet

Programmation office register document.

-Database employment documents wallet .

-Office document wallet Portofilio.

Information recruitment documents wallet  information employment system

Database office d base qbasic visual basic office COBOL .

- office project database documents wallet . 

Stationery document wallet office information handing Manuel system

-script hand project documentation

-arch file ,office folder  desk information table time table task operationel system  case paint case book library office ,

-office  size .mass. meter square area 

- casebook, rerwiten  framework tools form assessment formal Summative office ,

-Tools frame text page , stationery design printer sign industrial papper standard.

-Tools paint case design ,

Office automation machine system la guage conversation convert reader  PC and environmental PC printer system recruit machine database system project  , documents wallet arch file database electronic memory wallet , 

Documents wallet information employment system specific, responsible,,

-Documents wallet bank account database automatic ATM ,printer bank ,stamp teller stamp , statement,check database system , account books library bank ISBN order book bank order project book  bank wallet , employment salary 

Record book customer reviews, customer record registration ID number system information,

Documents wallet Portofilio job career emploie, 

- database emploie system entry  and dabase outcom up grade ,up database system project employment opportunities, ID 

- ID  registration form record form entry ,form exhibition,form move , form period form years value database emploie

Emploie record training job graduation,job training post advertised documents wallet resource humain , personel 

- ID registration form customer recording training custome entry exhibition, database emploie graduation job post  customer sale reward award customer ID wallet documents Portofilio, officer 

- bank account statement bank customer clearence cost amanded correction reward discount I'd customer deregister penalty, I'd customer record emprissonement labour work I'd cost years , system document wallet office information handing, 


- program logic control customer ,logic system custome I'd , programme language machine customer database system document wallet Portofilio, I'd 

 - wallet I'd trainer, course  job ruling job duty job post job weekend job . Wallet account e wallet easy pay system.

-Print I'd framework, print I'd information,,print post ponent I'd ,print job duty post rosta,print task operation print salary day , prive over time ,printed provide fund ,print tax job revenue emploie post ,print compensation award job , print insurance job ,print data  dol registration labour, print conciliation arbritatiin ruling job conduct job  day 90 day years progress conduct , misconduct,dismissal award job,print  holiday job ,over time ,print pension job ,print intellectual property job  register ,printed social insurance job security job , printing policy job record database system wallet accountability systeme, print order case power attorney debate creditor legal job ,  printed training record job authority job skill development job, training certificate in progress ,award certificate,printed agreement settlement job , print documents grand total wallet statement salary emploie resource emploie,humain ,material visa printed code , registration, database system providing fund , earned award printed , 

Convert 

-0ffice  wallet Portofilio Poste office documents wallet resultat ,re posted , disappointed documents printed , rejected documents, expired documents posted up date documents,up grade documents wallet Portofilio employment and custome record , licensing database.

,- design project office to office travel office information system, management information systems documents, , 

- design projects inovation management system conduct process following database maintenace wallet account account posted recruit machine database maintenace repaired , intelligence systems rebuild office , rebuild Poste re posted policy meeting, reprint documents review document support, wallet account,

 



Booking: filing documents posted wallet account documents booking training records ,



-office travel,,office cad corol drawing ..project office documents wallet engineering drawing .line work , planing work, planning jobs, systemWallet account dr




12.Engineering: electrical




1. Purpose documents




     

     

     

     

9.purpose : documents wallet,office,

Training customer office record , documents,office,

Training customer,and training student university trading up date upr grade student information system,up grade connexion system,

9.1Purpose storage , office genie counseling,genie soft ward genie hard ward  ,, safety security  policy office counsultant  jurisdiction 

Documents submitted 

Assessment and  re assessment peer self group  information technology intelligence , technology

Technical PC vs science computer vs qualifition insurance information 

-help memory pilot math information vs campus numerical box ,math info,news technologie ms word ms help memory copying syllabus training Unix  Google 

Intelligence artificial genie logiciel , input technology ms  office creation news technologie help memory use for ms treatment technologie input output system expert intelligence capable resource human capacity not it includes mathematics algorithm ,,

- calcul arborith , topographic topologies aerveur , intelligence robot automation science language science language science computer science language ,science mathematics chimic physics stereo chimic  physics  synthesis and , Polytechnic  electronic information technicien ,,no ITC it is include  work topic it faculty option intelligence artificial and genie logiciel ,

- pilot lecture pilote fly driver disc memory network navigation,  

Mathematics sacrifice workplace visuel basic ,

- honesty intellectual library pilot ask questions training PC cyber ,training research ,x vs y procedure create page create algorithm arborith ,step number phase fine creation page , button exportation language literacy ms word , creation find studies habillement ms word creatt find studies caractere address vs medecine mycine bio algorithm ,science Polytech algorithm

-retetion office ,button cyber button command ,do's vs algorithm procedure medicine to mutch work office contii algorithm is vs algorithm for students ,

- subject theory vs for laboratory training work , literature 'creat visual basic  office protection documents client  terminal all letter memory ,

- algorithm initial procedure select , operationel insert button, click operationel,

- information processy vs Computer office , financial Computer ms , window theories safety policy theory  vs ,technical PC ,vs science , PC ITC theory create cover page letter  folder create financial fibre optics create xompui lettre CVS ,create ms word processing marketing processing asj , buttons ,create for data  data base office ,intellectual ,oms algorithm PC ,technical mathematics science PLC ,robot Unix,

-info process vs computia vs management system  information business PC ,admnistry PC theory cover ,page office traine power point acrob converse,

- mathematics info comptabilite invention creation plan comptable code binaire calcul, vectorial code binaire 'calcul vectorial dimensionnement algorithm mathematics financial invention plan financial savant Gestion ,

-arm calculteice commercial, invention formuler mathematics function grand droit ,

- invention: ax+ by+c. functionalite logiciel rigth autocad cad ,vector financial functionalite value  adddiatribution ITC ,scatter robot geodesie projection ,invention mathematics ,savant mathematics autocad,memoire mathematics

-inventiin mathematics info science computia mathematics exchange call cash nothing financial functionalite,

- invention mathematics series sequence integrally ,x =3 ,if y=4 else the x,y co - ordinal scator ,z = matric  information equation ,x+y = 0  Fortran sum limit x ms ward display, ITC InSite mathematics coverage page is site theory in site rules ,matht recruitment info mathematics recruitment informer investigation Mathe quadratic make square rectangular plot investigation quadratic information language display generate for ,no it robot intellectuel system mathematics invented equation ,x+y = 0 quadratic,

- no invented robot x,y function scater robot intellectual ,financial invented calcul, permit to account money ATM to save formule ,

- info language inventor informer 

- info invented process,

Mark ROM,octal ,

- math info initiation information arithmetic logic process logic ,Espace  pixel VGA,cla or, invention

-mathematici do software invention hard math PC ordinance sequence,MHz gigabit calcul logarith intellectuel processor it mathematics invented equation computing ,

-ms word disc mathematics character 100000 caractere programs frame program , excell table equation financial word limited

1099 word octa octal binary equation,-ms internet robot word page mathematics character system mathematics,

- do's mathematics code 00000,

- math inform science and computer science chimie info physical science chimie faculty chimie instrument, conversion synthesis,

-mathematici info et Polytechnic and science academic material mathematics,

- power supplies arithmetic machine ,process control project ,VB ,access , actual technologie ,info mathematics equation logic integral Lim actual review series automatic system robot language technology mathematics ward caractere , actually technology,information ROM ready, access.memkry,


,

 



     

     

     

     

to me 

 


Schematic showing 

 



     

     

to me 

 


Schematic showing a 3 phone system new button visibility above ,each handset must installer ,intercom system used surplus ,

- parlist : handset regular telephone type no push to talk or sound ,power , lungs buzzer 6, volt Eduard 15 ,see text cradle switch phone or see text , strip text cases miss hardware,form system every 6 day or equivenlent ,6 volt interconnected , distribution such shake ,


- personal call plan : 

Telephone call plan B1, B,2,B,3

Call A1.                   A1B1,A1B2. ,A1B3

Call A 2.                   A2B1 ,A2 B2,A2 B3

Call  A3.                    A3B1,A3B2,A3B3

Call  A4.                    A4B1, A4b2, A4b3

&1 alpha= the effect of in the level of fonctionnement call telephone call plan ,I=1, 

Beta=the effect ,ab of factor level , x1 jk= , error association observations level plan analyse 



- radiotechnic,

Power amplificator TV sound base oscillator pentad tube ,

Characteristics indirect cathode wire V,1 6.3 v,

Source wire,

Use conditions nominal RMS,

- voltage and.  Va - 170---250V

-voltage grille. V get --170--250v

Voltage.        45 ,  32m

Current. 62  2,4 ma 

- coefficient amplificator. 15 ohm 

- resistor internal   0,2 ,  4,6 MV 

Capacity 

Capacity grille ...CG ---14,7 of

Capacity and grilled cage <0,8pd,

Value limited 

Peek voltage  anode vap max 7,kV

Voltage anode Va max 300v,

Voltage of grill.           3000,

Cathode current v max 30 w, 180 ma,8,5 


+ Telemetry video output pentode ,power flip flop, ,


- emettor , af  amplificator ,emettor ,compositor number ,receptor ,display , speaking,


- IR(t)=VRF cos ( wi.t)=I(t)+ "v1

-QR(t)=VRF sin ( wot)=Q(t)+ "v2,,

Insulation in receptor conversion,in case we  Cree en our mixage voice I et Q ,, Voltage ,, quadrature, insulation mean level power 


-typica  using single side ,

telecommunication street distribution for new urban residential estate 

Electricity peddel ,started conduit , road way ,

Legend ADSL signal spliten  telephone , exchange, ADSL convey frequency , asdl modem yes  customer premise,,

- typical grand copper twisa.  Network telephone exchange, ,

Cabling homes for telecommunication completed guide home cabling ,

Line 1 yes ,line 2 yes ,,phone ,1 yes ,phone 2 socket jack yes , mode 3, 5 equipment,

Possible fault due comming of the telephone voice port inside , fault due to switch relay connect in mode ,,

-basic home network system typical cabling arrange and connection for , typical telephone and data services,

,Famut TV , bender PC, study PC, kitch PC ,wall ,,ethernal 

Legend : modulator socket CCO,modulator coaxial socket , coaxial plug 

Polyphase emettor receptor frequency intermediate heterodyne class,modulation phase phase (t)= cos ( wt+$0)

Sin a (t)=o,then m (t)=o module phase shift signal module ,m(t) = sum ,,

-canal of transmission GSM cellular antena  interval time area 577us signal terminal,

Wave electromagnetic tic plane impedance caractere area,E/H=377 ohm ,, schema btsbsc,msc network fixe ,,

Numeratation of voice 

Fibre cider , module, filter decoder  ,input output band 300hz - 3,4 kh, 13 bit cadence ,,

Input line l1,2,3, cabling 

- circuit command selector , meter decimal Relais , rotational test , duplicate line off,cut line , Decimak control ,direct A, switch register in the register in the group the calling division aearxf variety,

Switch line ,switch final connecter ,signal control send impulsion ,interconnector , pilot switch start , div ,div,,,


-multivaria data analyse telemarketing,point ,do" sumb( x1k-xjk ) exp 2,, structures equation model customer loyalty adversity board award , quality imagine customer satisfaction,x x=, loyalty= b1 brand customer ,brand equity =C 1 advertising + c2 brand award + image price  customer satisfaction = image price,

-diatribhtor performance

Distribution/ service ranking X1/) overall performance ranking y1// ranking differente d1=y1=y2/// diffential sequence,data analyse , investigation of association ,

- Winer trading plan  bet 

Net profit= number of wine bet x premise pay out motion number of lossi. ,bet premium, y profit or loss,w number of winning vet ,k payout motii,


-model : l

 



     

     

Model: ml 06NP, ML6,c

- capacity : 6kg -15-30

Resolution: 1g+2g-5

Weight range ,- 3 kg ,7,5-6

-train display ,zero net 

- sub display: zero net ,vf d LCD graphic , 

Disparue dwe , ply memory/ ml on 

Print / rypr , available size  labej ,Manuel 

Max : 90009,pl, width

- label printy , interface,power requirements ,power consumption, operating tem plotter size ,mm ,product model

- report financial 

Option model product 

1 m b memory ,

Specifications display ,operator ,5.2 LCD , 2x 16 pop up ,20 lines 

Memory back up , data up vendum lithium ,power consumption , AC 120 v ,operating 9A, standards 0,2 A , operating AC 239 ,operating dimension 


-system development: 

Sample Gant chart showing progress system development activities by putting bar ,

Project planning documentau, page ) 

- system ware inventory system modifucatt system  ware house ,

- system schedule activities completed activity,analyse 

- activity ,/1/2/3/4/5/6/6/8/9/12/13/14

-required definju

-for project team

- definition and

- definition objectivity

-R3 interview whereataff/

Organisation required

Vo review

Design 

- revise programm

- specific report

- specific screen

- specific documents

- soexifux Documents change

- management review

- implentation

- code programxhange

- build test file 

- build test 

- test produxtiin file

- revision product file 

Revise 

- production file

- revision

Test short Bev

Test product file 

Manage teab 

Install where house 

Teaini new 

- quality management iso 1999 totaj staicale process control ,produce software group project team 

- product / vendor

Designer ) Oracle

Development studies) Microsoft 

Knowledge/ system architecture,system investigation use engineering computer case tools automate ,task required system high degree , package focused  association stage v, arafw case tools selected

- development for orientation program

Frequency ( y,x ..) 1/6/15/29

Head score x /0/1/2/3/3(5)

Score deviations/-3/&2/

 



     

     

     

     

Purpose: technical PC vs  technology PC support , 

-Technical PC electrical electronics PC 

Open lab ,mechanical electromechanic mecatronic system disc , buildings construction PC architecture  PC 

Art PC hierachie vertical horizontal process word , excell  algorithm

 Fundamental PC ,

-fabric fundamental process system 

Fabric circuit resistance ,

R: material, R=resistivity ×l/A , I =w,,R= resistivity .l/y.w=, R3= resistivity/y,,

-metalic late fabric , semiconductor film metal oxyd silicon Fuk autonom career , Sio2, ,,input  R1/R2 ,,

DDP/Nikel, charged ,,,discht batteries radio

-Pbo2+2H2so4+Pb=PbSo4+2H20+b.bs0

Turned. 1/2À 
________________________________________
fc=2À 
________________________________________
.f.lresonanance, 

L.fo..fox=x.l=796 khz select ,Copa ,,

-if filter modulation signal, mixer yes ,local oscillator yes ,,

-Am yes , 1mv, first r.f yes ampli, second ampli yea  50 MV,, detector yes , AF yes ,power out yes ,5V,  10  w,,

Dielectric: microwave , wireless communication technology yes ,master oscillator yes,

-master oscillator yes, multi yea ,power yes ,antene yes ,,

Parabolic yy  advance ,

D= Pmax/PAV,, Scat = 174.698/DM,,

G= k(À 
________________________________________
.D.D/x),,impedt , Z ="R.R+xl+c),,z= R+j,,


-electronic digital analogic circuit ,,

Input a,b and gate yes z= A.B  1 kilo ohm,

Booleen algebraic,A+A=A,,

Identify booleen symbol switch circuit function F ( x,y,z) = z in ,

Karnow map ,,y,z 

AB+AB=A@b

- decimal number /x//y///z Min term ////max term 

I/0//0///0////0////x,y,z//////x+y+z

_analofue computer , comparison analogues and digital computer,

Quantity,/analogue//digital,

1, representation /continuous voltage //binary number changing step

2aritmethic /by measure voltau and // by simple add ,losing course shift

Output of information/ graphical // languages quantity data storage,

Storage,

-potentiometer coefficient resulted and reducing voltage integration by factor (6/12), initial x= 6/12 chart capacitor, input output and gate,

Inverter integral circuit  DX/dt 6/4

12,


Sum ,interger ,summat, 

Differente  RLC 

L.di/dt.+R+ 1/c integral I St

E=L.dk/St.ln+1/c.inyegral, 0 time ,t 

Timer ,o'clock 


Initiation condition variable ,respect time St ,, IC  input x,y output eo=[f(x-y)dt,, ( eo) solving  differential equations,

12 DX/dt+6x= ,given x =6/10 at t=0, derivation, DX/dt+6x/12 ,equat,these 6/12 require at output integration,

-conversion 

Digital input a,b,c,d  yes D/A..output 

Digital / analogue

D,c,B,A/

0000/0volt

0001/1volt

0010/2 volt

0100/3volt

0101/4 volt ,to

1111/15 volt

- full scale voltage , ex 9.1 5 bit D/A converter ,0,2 volt digital in 00001,analogue 11111, si SB =0,# v and 1111=31 full scale output = 31×0,2 V=0,2A , even 9,2 A 5 bit D/A Conte delivt in output current,100mA, digital input cours digital analogue output cours digital of 11@0@=10ma,, 20 = 10mm let B = 10/29=0,5 mA @110= 29×0,5=14,5ma,,


- algorigram equation booleay,,

Prog, start writing a few dug a possi yes , test run progrt yes,or fix bugs an introduction as news on possible, yes 

Bughe , stopped,,

-x = 1 and y=1, yes A<Byes x-x+3 and y-x,y ,print z , A,< ,x,y yes , x-x+3,,y+x+y, z<x+y


-machine code programme ,x=4a.a+2b,z=N+(x-y)


-op -code /address

1001-load ACC/a= address 0000001

1010-multiple /b=address 00000100

1101- stop  x= address ,0000011

0001-stop,,

- construction PC architecture design development,

Input unit yes  mouse optical keyborad yes card red , control unity yrs memory RAM ROM yes ,  output yes arity yes CRT display print plate yes,

- database yes 

CPU yes ,room yes,I/o, yes vertical ,

8 bit  69kbyt,, 64× 1 k=65536, FFFF,52428

Memory systt ,@00, 

Digital / analogue, 

Sequence+3v yes 0,15, yes a+b,  A,B ,

Room ,ram ,,

Inlut

Row,yes and Dara,yes  gate and output gate and , 

Transistor bipolaire  logic diagram static RAM , ,select  R , supplies logic ,y select ,,

- booleen algorithm: program, read write yes, memory registerer x,bit ,.Emory buffer,

-printer charge plate ,electrostatic , digital input ,carater source ,papper,

- lazsr printer ,trasfdr ,Lazer,mirror,module,

- ribon papper,character, rotau,

Magnetic tape. Supplementary capstir,data organisation file

-data memories  time yes,

128 sector memoire, 564 bit sector, 

Input output device microcomputer converter serial data , parallel decoding, microspace,ram 16×4 ,, binary input yes ,

- disc label plastic with write. Ed inde hols , 

-sectir track 2,255 byte ,logic process read, digital ,step motor ,

- NV ram yes ram yes ROM yes, CPU semie conductor yes ,hard disc driver yes , external yes flopy disc ,data 150 km to 12m ,@00 m ,

-Typical machine code

3 bit /5bit//descry

000A/AAA1/load AC 

001B/BBB/

-adress / code//instructions

000/load branch output /load a.c

001/

0010

0011/

0011/,


Network interconnection  cellphone PC ,,

 



     

     

     

     

Purpose: project and circuits ,series and features, and services

Constructional project:

Design concept, sample ,circuit ,

Part list multi message voice recorder,

-1 PC board code ,size 119mmx 57 mm

-electret microphone insert,

- 3 way terminal blocks PC mounting 

2 way terminal block PC mounting,

- 2 pin section of Sul header strip ,

- jumper shunt ,

- 28- pin Dil IC socket , 15.24 mm spacing,1 8 pin Dil IC socket  ,7.62 mm, spacing, 1 2,5 mm concentric DC power plug ,PC mounting ( con ) , audio phoni socket ,PC mounty( con 2)

- semiconductor

1HK828 voice recorder IC ( ic@), lm 358 dual op amp ic2, @ 78l05+5v regilt( rsG @), 1 5 mm green led ( led 1), 5 mm red led ( led2), 1 1 N 4 004 @A diode ( D @),,

- capacitor:1 ,2200uf  16v radial elect, 220 if  radial electrique, @ 22 uf radial electrique, 1 4,7 if  25 tag tantalum, 1 220 nf  100v mkt metalisssd polyster , 100 nf multilayer monolithic ceramic, 1  ,150pf disc ceramic,,, resistor ( 0,25w 1% ), 1 . 470kohm.  ,2. 100k ohm , 9 , 22k ohm , 2 . 1k ohm , 1 , @00 ohm , @ . 47 ohm ,  2 .  680 ohm,  

From ,kit available

Link 1: in = beep disabled, ,,,lk2,/lk3/ operating mode , in /out/ 2 message, random access,/ out / in ,4 message , random accet, out ,out 8 message random,access ,tape


2. intelligent dimmer:

1 PC  board , code 799, availabt ,service size 76 mmx 50 mm,

IP 65 sealed abs plastic case clead lid ,size 125mn×85mm×55mm ( jaycar HB - 6#46)

1 .flush mount 3 pin main socket Jay car ps ,4094 , similar

1. IEC male chassix connector with  mounting holes Jay pp, 4005 , 1.10 MHz  crystal ( X1),,,147 uH 5 A inductor jaycar LF - 1274,,,14 way Dunkle vertict sockets con Jay car H3114, 14 - wat Dunkle screws terminal plug ( jayecar H M-3114,,110AIEC main cord,

-Semicondors, @ - 1 pic 18F 1320-i/so soic Pre programmer microcontroller,(I @),,

@.IR receiv jaycar zd - @952 (IRD@),,1BTA 16-500 isolated tab triac ( triac @) - do not substitute,,@ Vc337 NPN transt(A1),, 2RGB 5 mm common anode led ( led and led 2, 11 N 4004 diode ( D@), 1 UF4004 ultrafast diode ( D2),,11 N47341w,,5,6v zener diodt( zD1),,

Capacitor: ,,1470uf.  ,16 v radial electrolytic,,1100 if @16v  ,250v, AC x@ metallised polypropylene,,@,1100(0,1uf)250v ,mettalliaws polypropylene,,14,7 nf mkt polyster ,222 ceramic,,

- resistor ( 0,25w,1% unlet Alexi

13,3Mohm ,1w.  ,,110kohm ,1 w,,11kohm,11 John wirewound , 12v operation,,,14709ohm  5 w wirewound,,

-miscsllabsouse,3 m3 x 25 mm nylon s Rew ( to secure px board, 2 m3 x 15 mm nylon s Rew for IEC connector, 3 m3 ,12 mm nylon space,10 M3 nuts ,,, 1.100 mm of 0,6 mm dia .tines copper wire for link ,1200 mm length 3 ,core Naina flex 250 v 10 A rating , 14.8 mm red spade connector fully insylau, 14,8 mm yellow spade connector fully insulated , 14,8 mm yellow spade connector fully insulated , 5 100 mm cable ties ,,

- additional part required for testing: 

1 .12 v AC ,50 ma ,or 1 A plug jaxj , 112, 

1.12 v ,, 300mA light bulb 

Design : drilling the lid ,,install with crio ed, components side board ,,copper side of board ,, cut out for IEC mains input plug 


- building the circuit wizard way , 

Decade counter ,0-9  using j- bistable,amend ripple circuit and gate, flip flop , 

Investigate 

- a block schema diagt of logic system ,,under carriage door warning logic 

Five door switch logic signaj respective door open and logic 9 when close all warning indicator are  active low visibility audible, study what logic level appear point x,y,z all door closed, wath logic level appear point x,y and z with the left wing door open and all other door closed ,wath logic appears x,y and z with open all door closed ,4 when any more ,

- answes to xhet quest ,,feed back via collector

Mp lab,library install tutort program

-led 

EC reflow new professt,PCB  operated from  professional machine


-supply voltagt : 239 v / 50hz only

Power : 3500 w

Weight: approx 29kg

Dimension  629 x245×520 m(wxHxD)

Hear methoy , combination it grad and 

Operat: directly menu button LCD ,

Menu ,englit,, temperature #5to 300

Maximum pcb size ,400×#85mm

Temperature ,2 internal ,optimal 

- tear measuy 

Component  list.

Resistor: r1,, capacitor, semiconductor D @ to ,IC ,,miscdlanouss 


String buton  and function using button s 1, S3

- command action   interacty command 

Lab testing 

 



     

     

Specifications:

Suitabifor 2×16,4 and LCD display using stanu14, or 16 pin connector back software controller,backlii,

5×6 matrix kepad for maximum of 12 key switch or nine rotary encoder building inpusht equivalent,buzzer,on led,power via USB,external 5 v supply primary cell,0,8-4,4 v lipo battery ,5 v and 3,3 v regulatority sodtward on off contri possible, lipto battery charger measure , 34 bit ,48 pin LPC 1343 microcontroller with 32 kV flas memot ,8 kB ram and numerous peripheral like USB, ix,spu,mlu uart , compatible with the free LPC presso ,3,4 and coocoz ,ise,,compatible,debugger, extension connector almot, availability on connection ,splita le , detachable Min 4 key pad maximum. 3 rotary encoder ,dimenst adapted to type 2616099 case open source ,

Operationel: microct,stars is oscillator

Arm operator logic 

 



     

     

Purpose: open lab system, game completed module test test practical, test power electronic numeric ,logic  sequence, programmation amplicat feedbat, 

- experimental voltasi out + 15 VCC protection short circuit ,12 vci ,2 ,24 VC, functionalite,generati onde sequence amplificator 5 v , frequency 1 Hz to @ generator ,@ Hz , transitor Montee,@ Hz to 1 khz , transisty logic true false 

MOSFETs, cont rectifier motor inverter, principal ,12v ,3900 rpm ,1,2 A , 3 cm , logic  numerical ,4 and ( 2 input ) and ,3 input ,12 and , 4 input , 12 diodes , restrictions, study ram and stockagw information ,pannel ,4+4 buffer ,3 stare register buffer ,8 bits ,decoder , of 4 ,19 ram , 1024×4 displt, hexadecimal, analogic 8 bit ,convertisaeue ,8 bit poteu,

Bistable panaux , demultiplex,4 flip ,compare,4 cannot, register , calcule , arithmetic,16 logic ,4 bit , compte binaii,decodeur elrom ,8 sons led decodeur , linear discretion ampli relay , cablage schema , test execut ,system mesure norm , network transmission,53 MBS ,16,,

Biomedit, multimeter,3 and 3/4 ,voltagt cc ,490 MV , 400 v autirange generat, ,,

- detert type size ,if move water ..

If , else , else if to ,increase,, 

Procedure PC ,input file control ,select print ,data , print file ,work store , registration , 

 



     

     

Purpose: 

PC  speciation: format 

Intel core 7 , exaterne 975 @ ,3,3GHZ,,Intel DX58SO motherboard,6GB ,G skill trident DDR-200 channel,gainward ,GTX580 sponsored by evetexh ,wD 320GB ,data IU Hard drive ,Intel core ,I 7 870, Gigabyte P55-UD6, LG flatro n , Kingston SSD now ,,

AMD ,, specifications:

A, essential up grade, components, AMD , procedure , chip choice: the starting dusk x 2550BEb,CPU ,,X4 645 upward faster closely, processor ,direct conflicts show improvement,load processor specific benchmarks 3d rendered video encoder x 264 raw computation performance of up graded CPU increase ,time faster CPU therm ,,fps function ,,

-graphic upgrade: news little graphic card up date gaming Phenom power x 25550 choice graphic card filled card up grades , sub R1,500 budgr choice HD 5770 bigger to jumping performance DX 10 gaming word in conflic,

-playform up grade , CPU  up grade path for LGA ,775 sockey motherboard  PC area ,LGA 775 core ,3 GHZ, core, 2500 k CPU , 4GB of DDR, ram ,,

- direct x9 tessellation performance,

Heaven 2.5 / frame per sencodr higher better

Base system with sapphire HD 550/@7,#

-ugradr to HD 6959,

Direct x 9 gaminy performance 

Lost planet 2/

-word conflic /

-far cry #/,

-0n the card  tried edge graphic tracked down @gb sapphire 5850 Xtreme , retailing R 2000,price  complain  6 pin power connector which PSU  PSU had it physics driver installed benchmark away, result were disappointed 3 D marks  improvement less word conflicts did manage,lost planet  frame wallet being R,2000 light, completed reliable,

Battei g ram: slot 1,GB system memory stick R150 ,ram benchmark,ram ,,

- upgrading memory and processor  GB ,test windows CPU,the HD 

Windows shopping GB driver

-rigth components for the perfect budget building: perfect machine

Shopping list 

- motherboard 

F1A75-M   - R1.100

-Processor and 

A8-3850 - R1,300

-Ram: corsair 4 GB

1,600mhz   -R1,000

-graphic:xFX random

HD 6670. - R 1,000

-storagr ' seagi

750 GB  - R 610

- opticy ; lite on DVD 

Rw- R200

PSU; Corsair 430w R470

- chassis : cm elite 343/

-total: R 5,360

Purpose 

Built: test components:install CPU  vital component case layer CPU guard ,CPU pea, lockdown add sized ,

-cooler and fan : fixing , bracket underneath ,screw holes fitting ,connect cable mobos ,

- fit the ram : ram open the catch,snai lock motherboard, memory Chanel ,1 and 3 operate,

- the graphics card ,CPU  bandwidth drops xb soon involve another GPU, in an slu ,cross fire x  array ,. manufacture usually,x16 

-hook up the PSU connect ,20+4 pin connector to mobo,the 8 pin EPs cable and whichever 6/8 pic ,

-test the components: power that sucker up by turning the PSU hitting button ,your screwdriver board

Striking the balance, because corsair 4 GB 1,6000mhz DDR3,liani  ram APU shares the systt ram between CPU and CPU  match up ram gaming rates between at 1,33mhz and 1,600mhz  case , resolution monitoring 


Prep the case : bits ,ATX backplate and mounting  screws, install the PSU   screw top bottom cable out way module,drop in mother remove the GPU from mobo for the CPU Coller ram still in place  mouti f screw it in place round,

Reunite GPU and mobi surprised fiff

Ddly your case will allow pop out some driver bay possible card back top pcu ,

-test it still works : connect the PSU cablt mobs CPU power ,case fan yet post beeps or led light up your,

-tech analysis:

2560×1600 screen gamii surplus frame ,R 14,00 flashiy GPU bios ,

CPU rendering performance:

Cindbdnxh r11,5;/index higher is better

-gaming/533

Direct x 11gaming performance

Dit3/frames per second higyis better

Stock biof/355

6970/408

-direct x 12 gaming performance

Shogun/ frames per second high better 

Stock / 442

- hard ward reviews

-laptop repaired : screen replacement ,motherboard repair,ram,harddriver,CPU up grade replacement, batteries charge, domain registration, capped ,ucapped,internet ,fax ail voio,Intel B800,3gb ram /320g HDD,wireless ,/15,6" screen webcyfull keypad, camera chanej DVR 

-ugradr ,repaired build PC,upgrade key component ,fixe PC proby fast ,set up a home network hardware

Byii componey process hard ward,tools,checking compatible CPU socket ,ram core components,installing motherboard access Ga - 88gm - ud2h,USB support ,installing processor ,installing a coin,install ram modukr power,installer graphic card ,PC's reason HD 6870,power am 1gb of ram 

Installing internal hard drivfer 50gb optical driver part built up grading ,px lb 950S BLUE RAYY ,BLING MEDIA BUYB,CHOSE  READING SPEED CD 

- UNPACKING A POWE SUPPLY UNIT , SHOPIFY UNIY MODULE ,

Case and feature fan cooling

Up grade a PC choosing components open computer replace component challenge damage electricity,remove before process ,scenario compatible, completed desktop system,custising PC build configured a desktop PC ,budgr PC  cash 

- perpheraks gadgets,installing TV tuner ,HD  hard ward connecu, hardware projey, windows,

-aoc LCD moniy,case OEM,case raidmax gaming,cattex networks cable and accesorie,cps pose back up solutions,g data antivirus, Kingston memory solutions,maxcam secury solution ,Microsoft, tower server cabinet end closure 

 



     

     

Purpose 

 



     

     

Purpose: word display manufacture

-Introduction technology, teasing compagny specialist in it components,

Required market effective technology solutions custome the product security surveillance soluy, networks,point sale server encloy it components,

-Mission: provide technology added value business,provide quality product , creation make different,

-valur : accountability commitment quality ,honest integrity,

-vision : to inspire purpose great place work member learn productive business ,

-partenera  chanrk , 

Social but transformation bee compliant ..

Product Guide :

LCD monitor,screen size : 21,5"w/23,6,viewable image size : 546,86mm,display area : 476.(H),268,1@(v)mm, brigthness ( typical ):300cd/mm,contrast ratio ( typical ) 600000:1(DcR),response time ( typical ) 5ms, viewing angles: 170/160(car 10),max resolution:1920×1080@60hz,

HDCP compatible: yes, input signy: analogy RGB and dvu, user control menu, enter image ration source ,up auto / down power, power consumption power ,on < 49 stsndat<0,4w,wall mounted: vrsa 100mm, mechanical function tilt ,5-29, 

Specii features : touch key USB ,by DcR glossy 

-type qori - 335 case:type ATX Min case,motherboard: micro ATX , external ATX up ( up to 13"x9,6),(p4ready)

-power supply : 400W Culp,P/s with 12cm fan ( 20+4 pin socket 775 ready)

,5,25" drive bays:3,3,5" drive hidden:8, expansion slots:7,I/o interface USB x 2,mic x 1,spk × 2 ,cooling fans front ,80mmx1, rear 80mmx1, dimension (DxwxH):410x182×425mm,M/t(cuft):1,73,

-frint mount USB ,audio, high,green LCD fan, external bat : 4,5,25"and 3×3,5",internal driver bays ,4× 3,5" HD,,system board ATX form facty ,13",×10,5 , expansion slots standard ATX 7,material SECC steel, dimensionnement: 52,25 x 20,5 x 45 cm ( lxwxH) 

- raidmax  modular cabling system durable  titanium mirror grade block wraps , mesh cable to mention 20 total power connector for all your computer, type : ATX12v,EPs 12v, maximum power: 730w,fans 135 m blue led fan ,PFC no ,main connector ,20+ 4 pin ,12 rails ,PCI e connector ,1×6 pin ,1×6+12 pin ,modular cabling support yes , energie efficient : up to 80 % ,over voltage protecyyes, 

- network cable 

Cat 5 e cabling: 500M solid,Uto,24AWG,0,5mm,4 pair,grey,,305 m flex or solide ,Uto,24 AWG,05 mm,4 pair grey

 ,,cat 6 cable .305 flex or solid Uto,24 AWG,o,57 mm,4 pair ,Gray ,to ,,

Cable tester digital tester ; cat 5e, ,

- toolkit: j-059 long nose pliers (159mm),hy-103B micro cutting pliers + 159mm),by -330 cable blade trividr,hyp -5022 wire stripper ,hy 568 telephone plier ,8p8x/rj-45,hy-324 punch down tool ,tape line ,2m knife, by 329 wire stripper economic hy 539 soldering iron 30w,cable tester ,tweezers 125 mm, crystal screwdriver, -+,module plug rj-45/7p8c x 6 PC, plastic box ,315x255×55mm, crimping tool,Rachel type dual crimping,punch down tools, wall boxes cat signle part RJ 45keystine jack ,surface mountable ,car 5 double ,RJ45 jack surface, carb

- connector ,boot sleeves ,RJ 45 cat micron connector ,standard , cables make female 15 pin VGA extenst available in ,2,0 m ,3.0 m,5./ ,Male to male 15 pin,

- 1000va /2000va rack mount ups ,

-patch panel 24 port Uto cat 5 e panek w / t back  bar , 48 port Uto cat 5 e panel w/ back ,,

Battery voltage : 24 VDC, recharge tine ,8hours to 90% charger ,input voltage 110vac , 

.Or 220va c 25;,output wave from , pseudo wave , protection : output short city overload protection with current limiting,protection output circuit overload , 659 Va line interactive ups ,auto restart while AC recovy , provide overload protection ,compact size ligthb, automatically voltage regulator Ave , battery low voltage automatic , microprocessor control ,provide modem phone line surge protectors option , equipmt input voltage, building DC start functionalite enable ups , input commercial power range @45vac - 280vac ,AVR range 165 vac , frequency ,59 z +- 10%<, output commerce power : sine  shared frequency AVR voltage range ,220vac auto sensing , transfer b,gross weight ,6,1kg ,size ( mm) DxwxH,,260x96x135,, 

-Data manufacture data security, security technology,new top security  PC performance, high security,under completed undertectabke ,self fingers printy,maximum security record test winner ,, antivirus program enginee,anti spyware ,antipgising,antirookit,child protection ,system tuning , firewall,internet security,

Award technologie include double scan ,outbreak shield ,integrajitblocker ,special utility ,scan all compression file and archive formal, heuristic detection of virus, g data antivirus manage ,client ,premui support,,

- client security business ,g data antivirus management server,client ,internal desktop hard driver availt 1000gb

- value ram desktop memory, note book, server memory,,data traveler flash disj

Free agent  notebook external hard drives store Dara ,synch data between PC ,back up fil,, digital broadcasting,

Serie projector , 3 projection system ,3 LCD panel ,1 lens projection system, resolution,RGB,1024,668,lam 190 ultra high pressure lamp,colorboutput 2000 lumen ,video input composite video , audio input  ,

PCI , expression Serie , astaud performance,graphic processing ,stream processors 240, core clock 633mhz ,memory clock 2268 MHz ,memory type ,896 MB gdr3 ,memory interface 448-bit ,shader clock : 1404 MHz , bus type : pxu Express ,3 ways slu technologie,direct x support ,CUDA technologie 2 ,SLI technologie,Nvidia lumex ,dual link ,2 Nd generating architecture HDCP capable 

Discovery vision 

 



     

     

Purpose : Program analyse data,

- adabase ,storage data , association and work file access logic ,database modification and transaction ,introduction natural , accessing database , basic rwiten ,edit mask ,arrays , arithmetic more advance report condition statement , additional statemt data manupulr natural system command ,system variable ,session ,system functionalite ,input statement functionalite data area program editore Mao,, gierachir , network model, nucleus  buffet poopl.i/0! Disk actual data ,parallel process , interface programme  implentation methode  control ,format buffer ,record ,search  buffer ,value , 

- SQL data manipulation language , civil and pl statement written embedded , exce read logical ,natural acess operating system 

- civil ,

,General control block ,file control block field descrip data storage space ,up date command record is added to file new ,up date field value in fult ,manage command ,

- database modification : add if new record selected exist , competition uptade  case , user ,1 cancel days yes ,read employee file leave = 5 ,up date leave = 0!,,reserve days ,read reserve day 15;,read file 5 up file leave ,29, transacy concept ,the limited ,rejected accoun

Find statement,basic format t ,define data , employee ,personek I'd address ,end define ,number city address ,view

- break statement csalary record , department code namt

- compagny, 

Definition module:

Employee

T l db  name  f Leng s d remark

Personal I'd  ,first nane ,dare birth ,full address city ,zip ,post code ,telephone ,job , department , position current leave  take ,leave booked ,leave start ,salary ,currency position 


- data area editor: define variable with t data editor , define data ,emp ,name , surname,first name ,Dept ,

Incode ,

Date   natural for Unix library name ,time software AG libraries , Unix fin ,,

Empl ,A,8;

Name,surname ,job title ,edit 

 



     

     

 



     

     

Purpose:

Civil and development a program, 

Program pli,code , 

- input output firm due suply ,name ,file ,record ,size  implied decimal ,point indicate ,  name of output file record layout working ,storage nNe form ,outpy record number, 

- hierachy diagram  and pseudocode ,program logic planned hierachy diagram pseudocode indicate module , informs to plan step ,program ,stsdard , was ,I ,o ,start 200 first ,20@8, 

- modular programming , program readability debugging, 

Program readability:  , picture value coded , working field , accut total , 

Program specification : position , 1-4 , field employee  number ,name ,office number , telephone,,type numeri aplhabdt ,

- enter progi: also under sudej , save program , type word , compilation 

- execute the program : the name input program computer : 

: Report on disk in order to print later stage : 

- debuggi on the program ;: if report is not correct ,logic error in program that correct , person number ,name office number repeat ,states was omitted from , in case error access ,

- instruct print report input for program  data ..

Output program : 

Record layout , take record , input  ,file name yes ,input record  ,implotr number ,20-30 0 employe name , 50- 60 office number , 60  telepy number,

- acubol : identy  progr I'd program , Uthor ,date waruten ,date complied, 

: This prohray print  report with information aboyall information: 

Enviy diviy: configuy PC IBM , object IBM,input ,output ,file control ,select employee file  assign to disk ,orgNizTion is line sequential 

Ms soft COBOL : calxuly ,computer es wage rounder = 1 hours,tariff, computer newrou d ,prepare line , move @- to no , 

- Engineering comlGni increase 

Number / hour/tariff ///current /////wage permitted//////new/// wage exceed /


Hierachy program 

Work storage 

Posityreddfune was position table , occur time 

- procedure divisy , open input file report file ,error file, perform,read table time , 

 



     

     

Show in ZAR

Update

 

Sales by Item

Tshingombe

1 August 2023 to 31 August 2023




Item   Current Unit Price     Quantity Sold      Total Average Price


0 - Fiscal fraude,Min,max value , energie tax , engine

25,000.00     1.0     25,000.00     25,000.00

000123 - Engineering chemistry, engineering physics, metall

2,000,000.00     3.0     6,000,000.00      2,000,000.00

00112345 - Instruct,health clean biotech ,food,vacum tend low

50,098.00     2.0     100,196.00   50,098.00

00112345098 - Instruct,inspect food take typical info,compos low

50,000.00     2.0     100,000.00   50,000.00

001123456 - Instruct,buotechno machine food cook,agro value

500,989.00   2.0     1,001,978.00     500,989.00

0011236398 - Instruct, biotechnology, food beverages, composite

50,000.00     2.0     100,000.00   50,000.00

001123984 - Instruct,typic food nutrie ph ,basic,acid 7, GRA v

500,989.00   2.0     1,001,978.00     500,989.00

001126867 - Instruct,food medecinal value assessment low,

5,009,898.00     2.0     10,019,796.00      5,009,898.00

00113456 - Instruct,biomicrobiologi,scanner biochem,low

50,009.00     2.0     100,018.00   50,009.00

001134567 - Instruct,agro pastoral tools equipment,low

50,000.00     1.0     50,000.00     50,000.00

00114356 - Instruct ,policy safe biotechnology skill dev,Lo

500,098.00   2.0     1,000,196.00     500,098.00

00122357 - Instruct,food value Assessment Portofilio low,osha

500,098.00   2.0     1,000,196.00     500,098.00

001334 - Engineering plastic ,metallurgie, geotechnical

2,345.00 5.0     11,725.00     2,345.00

0022345 - Instruct ,bio food manufac package assembly mec

5,090,098.00     1.0     5,090,098.00      5,090,098.00

002345 - Engineering, construction civil,mecanic,const elec

345,654.00   3.0     1,036,962.00     345,654.00

00526 - Instruct ,visa electrical ,meter energy, permit

25,000.00     1.0     25,000.00     25,000.00

00652 - Instruct,visa ,markdept , license fiscal,license

258,369.00   1.0     258,369.00   258,369.00

007200 - Instruct commission, council,visa passport id

25,000.00     1.0     25,000.00     25,000.00

0072096 - Instruc, bulletin,visa certificate,metering city

2,500.00 1.0     2,500.00 2,500.00

007236 - Instruct,com, I'd visa ,wegh,scale,size, plate sup

52,669.00     1.0     52,669.00     52,669.00

0072536 - Instruct com,plate fiscal,plate , empower,value

639,980.00   1.0     639,980.00   639,980.00

007255 - Instruct, commission,Eng visa passport,I'd supplie

200,690.00   1.0     200,690.00   200,690.00

00725669 - Instruc commis, municipality,visa engine supply

2,506,690.00     1.0     2,506,690.00      2,506,690.00

0072569 - Instruct, commission composit,visa I'd , supply

2,536,609.00     1.0     2,536,609.00      2,536,609.00

0072693 - Instruct,com,visa city meter ,plate , fiscality

528,809.00   1.0     528,809.00   528,809.00

0072698 - Instruct com, visa wehit plate number, fiscality

258,800.00   1.0     258,800.00   258,800.00

0092086 - Instruct info manage system , commissair, rulings

2,500.00 1.0     2,500.00 2,500.00

0092096 - Instruct, database drawing cad , commissair

2,589,658.00     1.0     2,589,658.00      2,589,658.00

00920996 - Instruct, office travelling , commissair rulings c

52,000.00     1.0     52,000.00     52,000.00

009258886 - Instruct , database intelligence robot, system

582,669.00   1.0     582,669.00   582,669.00

0092609 - Instruct, performance,info database code ,

56,800.00               

...


[



     

     


     

     

et started with Google Play policy 

"    By Play Academy

"    Published: Apr 5, 2023

"    

Rating  1 star 

 2 stars  3 stars  4 stars  5 stars 

"    

Average rating: 4.8 13 13 reviews 

"    

Share Path 

"    

"    

o   


Whether it's a way to track workouts, chart the nighttime stars, or build a new reality and battle for world domination, Google Play gives developers a platform to create engaging apps and games and build successful businesses. Key to that mission is maintaining the safety and integrity of Google Play as a trusted source of high quality apps for users to find and experience apps and games to enjoy. To support this effort and keep an even playing field for developers, Google Play has policies in place about what is and isn t allowed on Google Play.

Get Started 

"    

Google Play PolicyBytes - April 2023 policy updates

o   12 m


"    

Disclosing user data deletion on the Data safety details page

o   5 m


"    

Declare your use of foreground services

o   9 m


"    

Best practices for integrating SDKs into your app or game

o   3 m


"    

Policies and publishing on Google Play Video

o   4 m


"    

Get started with Google Play policies

o   8 m


"    

Manage policy violations

o   5 m


"    

How to handle a policy violation on Google Play Video

o   3 m


"    

Restricted Content Play Policy Video


"    

Privacy, Deception and Device Abuse Play Policy Video

o   2 m


"    

Spam and Minimum Functionality Play Policy Video


"    

Monetization and Ads Play Policy Video


"    

Store Listing and Promotion Play Policy Video


"    

Families Play Policy Video


"    

Intellectual Property (IP) Play Policy Video


"    

Impersonation Play Policy Video


"    

Malware Play Policy Video


"    

Mobile Unwanted Software (MUwS) Play Policy Video


"    

o   



"    Add Path to Favorites 

"    

"    

"    

"    

"    

...


[




     

     

to me 

 





---------- 

Subject: Re: Documents wallet Portofilio,office Engineering project order management appointment project file order sale campagny meting tbrigade,edutech psscm... agenda office .PC safety wallet documents financial office

----- Message truncated -----


     

     

Information , intelligence information manager, perfect office, license agreement ,word perfect office ,word perfect, installing info Central,recommanded system requirements,question and answer, about installing ,taking quick making connection , contact manager info Central,using central contact manay,keeping track of people ,aff oersonek remember name , manipulation information screen how tab work ,sign tab saving lost ,customizi g information n,using call logging a telephone call , scheduled a telephone call ,using info book , viewing your information in an address book , printed , time  management info Central ,using info Central for time keep track of appointment ,scheduling event ,using alarm to helph ,written tour a self task listed written note planing  page wrapping ,closing task , housekeeping, connection with ,,

-document management with info Central, using documents keeping track of corresponding writing a letter, organisation  memos other document ,opening editing related files connecting existing documents,using infocentral as files manager , manipulation files ,searching for files, 

-moving information to and from info Central ,how info Central is different from database importing information from database or pim ,scanning the files tour hard disk ,moving information from one I base another ,,

- I base and template: using the Pre filled contents I bases using template, automobile record record , C's collection CD ,contact manager ,gardening ,home family ,stamp collection ,real estate student note,,

- license Granted of Rith software: 

Restriction permission ,homeor portable computer  provision ,limited warenty / limitation days, custome government , technical data,


WordPerfect ncall object discret ,name frien, click click call ,query ,call

 




     

     

     

     

     


 



     

     

, essential trade business entreprise ,point sale ,

Switch / base / switch per add / essential / trade //trade ////business/////entreprise//////// POS

Customer debator /60rand

Audit email statent journal ,print reprint invoice ,open account ,audit trial account account active overall , custome area interrs ,sales letter , multiple b,email sales documents sales documents number customer delivery check custom b,

2

. general switch : /600ranf 

Unit / vat , return, custom , setup define report ,design report  Omni automate , invoice ,export data from gride , multiple branches per company ,

3.job costing : emplitime sheet, job category,sub Jon ,   R200, 

4.nominal general ledger : audit trail ,STD balance sheet report, journal,STD profit budget,

- sticker inventor : stock control base multiple, vin location,lost stock item, stock take ,stock,stocket ,stocket badge,,

-suppliet creditor : 

Name , statement agent ,transaction supplier memos, supplier credit note supplies invoice, remittance , supplies line timr , settlement,email supplies documents,supplier ,no supplier back order extra cost invoice ,import purchase order ,full drill ,sars ,v

-connnection credible ,incredibly: wireless solutions , note book ,premii,,in store completed technical data recovery service ,able your data ,

Memory upgrat ,ram ,1 GB to 4 GB , 400mhz drrto 133-ddr,  R. 500 to R 6000

 






     

     

Quick book,, product pay roll software , invoice quotation, payroll ,uif automatic journal quick book,general electronic ,fund transfer submission uif ,store access full emploie n,contrul periode ,rigth security control n,finanaciaj detsui ,define eaenubgb,access a number of standard report ,bank report , department summary boer week ,pate recalculate ,store access full emploitb,analyse customer supplier ,quick access to key data one place  compagny n,access  business,money in out money reminder ,cash flow report bookeping ,multiacces ,send email cash 

 






     

     

Purpose: quick set up 

Install 

Feature : serial ,ata , power tray loading and ejection of disc , combination CD ,Rd - re, DVD - R/Re, DVD, , support region ,play back control  phase ,support region  secure disc technology, automatically buffer underrun error prevention technologie, support ligth  direct label ,printed technology,

/ GH22lS30/GH22NS30

Ligth scribe / supported / not supported

- package item 

Item      / quantity

-intetnal super multi DVD writer /1

Record and authoring software/1

-neto in CD power procedure /1

MpEG - 2 software ( power DVD)

-quick set up guide for super.multu DVD rerwiten/1

Data data cable /1

Data power Cale, 

_ system environment PC:

.System requirements/ CPU : motherboard data pintuim IV ,3.2 or higher ,ram : 512 MB 

.VGA card memory/ 128 MB or higher.

.hard disk : 20 GB or more free space , writer depending on the quantity,

. Interface :/ serial ata

.drive bay / 5.25 inch height b.

Power/5v ,+12 power socket , requireded,

Software requirements : window vista ,direct x ,9./ 

Recommand media manufacture/ CD-R ,CD re,dvdr,dvdre ,DVD+ r, DVD ram ,DVD ,,

Specifications item: 

General / 

Host interface : serial stsb,

Support disc; DVD ,

Supported secure diameter ,

Read write speed ,

Write DVD r 

- item : 

Performance data transfer rate , sustainable : CD  ROM ,72000 my yte ( 48 mac ) , DVD ROM ,22160! Kbyte , bursi SATA ,gen 1,15 GPS , average access CD ROM ,220 mec typic ,time , 140 mesec ,bufer capacity ,2 mbytr, MTB, 100000power on hour duty cycle 20%, 

Environment level, when operating temperature ,5 c to 45 c , humidity ,@0% to 80 , product manufacturers to with radio interference EMC directive 2004/108/EC and low voltage directive 2006/95/EC,,class,caution class 3 b visible invisible Lazer operation radiation bexpisure beam

 





     

     

Student guide office automation: Holliday ,booking confirmation form: automatically field letter enclosed,squeezed memory.

- top 13 logiciel Gestion school management computer process top 13; tools product module , cluster,maintenance system product chronological system ,technical maintenance pedagogic accompagnemeny letsonsj,maintenance it information ,master reseller automation , intelligence workforce platform b,native cognitive skills vision ,tele operation maintenance field data forecast equipment fail algorithm,process and analyse data status robot components ,big data intensity process maintene Cr activities industrial fully productive , maintenance bue ,system food I formation entreprise keppi g database pc securite, protectionife cycle , up date application detection apply,ms CPU GPU ,ram  clean system  network ,space driver size , defragmentation,desintsllation of programme ,unistsllet , antivirus,task ,,



Tpm daily maintenance productive system policy plan compagny goaj

 




     

     


Purpose: analogy I/ o process control ,lab purpose  : PLC ,API 

Application of embedded controlet to real time algorithm emploie  analogy input  and anali output ,computer algorithm to implent closed loop processor form motor speed control , hoe feedback linearise ,non linear process and resultat in zero ,

Object : generate pwm output implentation variation motor supply voltage , implentation tachometer operationel using pic 32 timer , develop the programmation code to Implementation pi control moving average digital filter,mage multiple interrupt driver system , monitor ,basic , suggest reading embeed mechatronics , microprocess hardware bbusy m X3 trainer board workstation ,PC t windows MC is,lunux ,23 DC motor with 5,vb,,,DC power supply software ,mp lab x, project tskey to read analogy compare to Implementation a pwu capture period ,  fundamental digital open loop and closed loop process,,unitprocess , process control ,automate process control engineering,d..

Display yes, signal conditioning yes, loop close yes computer algorithm yes , microprocessor yes ,DAC yes ,signal amp yes ,signal yes , process yes ,,open loop control loop process,DC motor speed VC pwm ,counter record position t3 co

 







     

     

Purpose: robot dynamic, kinematics and control , operator ,introduct , coordination ,lineare velocity , rotation matrics active passive rotattt composition representation Euler angles , angular , 

MATLAB cad: 

axb=[a1,a2,a,3]×[b1,b2,b,3]=[a]×[b]=[o,a3,a2,,a3,0,a1,a2,a1,0]<, angular velocity acceleration b, angle ,z,y,x..yy,z,, kinematics planetary , 3 dog Robi arm ,generalized coordinator ,are Q= ( Q,w2,Q,3)=( alpha 1,alpha2,alpha3, ) effector position and orientation ,x ¬ (Q)=(x¬ o(Q),x¬ R(Q),,,

X ¬ (q)=(x,z)=( 1,sin(Q)+l2sin(q1+q2)+i3(q1+,q2,+Q3), Io+u1cos(q1)+i2 cos ( q1+qq)+ cos (q1+q2+Q3),,x¬ r(q)=x¬ r(w)=q1+q2+Q3+


Function ,ph = get eulu angle ,x,z,y from rotation matrix ,c % GETU Langa x,j z from tion  matric ( c) extract x,y,z Euler Ngler form ,% rotatricr ,% author ,x y = a tan 2 ( -+2,3),+(3,3),,y = a tan 2;( c @,3),sar + cc, 1), 2 (1,2)"2,ph = [x,y,z]<

 

    

     

Purpose: PC job , path  , measure availability, availability= operation time/load time,,

- loading Times - dow time / loading time,,operating = 0,5 minute×100/0,8 minute = 460. Minute,,

-  operating = 0,5 minute × 100/0,8 minute=62,5%,  net operating rate = actual processing/ operation time ..

=  Process amount × actual cycle time / time time , net operation, 

Rate = 400 item × 0,8 minute × 100/400 minute= 80% ,, performance , efficient = net , process amount × actually cycle time × ideal cycle time/ actual cycle time,

Processes amount x ideal cycle time / operation time ,

= 400(item) × 0,5 minute /400×100= 50%,  0,625 ÷0,80×100=50%,

- A . Working hour per day , 60 Min ×8h=480 Mon, B plane down time schedule ,c: load  time  per minute break down ,20 minute , step ,

E: operationel time ,day =c-D=400 Min ,G : outlet day = 400 item , H : Raye of quality products ,98% ,,I : ideal cycle time 0,5 Min item , j : actual cycle time= 0,8 Min , f: actually cycle time = 0,8 Min it's , f actual processing time = E/c ,x 100 

M : operating rate = I/j × 100=0,5/0,8/100=62,56,

N: net operating rate = rate F/ E x 100= ( o,8)(400×100, l performance efficiencye m x n× 100=0,6#×0,800×100=50, overall equipmt effect = txlh ,,,network control critical path work , organisation chart ,flow process chart ,key , operationel,transport ,inspection ,storage ,,, load time operating time ,net operating Time yes , valuable time yes ,, equpmebt faillure ,adjustable, minor stoppage defect process,,

Calcul overall equipmt, availability= load time diameter ×100/load time,, available = 469 Min,,460 Min -60 Min × 100/load time ,, available = 460 Min -60 Min/460m ×100=876,, performance= there cycle time × process / time operating ,, performance efficiency= 0,5 Min × 400 unit× 100/400=50%,,,

Overall equipment effect= availability × performance efficient × rate of quality,,

-  Design test p.p,,dea

 

     

     

     

     

Test design ,3 or design test p.p maintence p.p,/ link  maintence wire assembly maintence, order an , purchase, order and purchase 

 Power empower project , command project process 

Duty cycle ,duty classic, operationel component data cycle duty ,

N = operation time at constant load max , maximum, temperature, operation at constant load St ,n ="!time at rest  flux max = maximum temperature, attained load electrical losses cycle duration, time ,start , acceleration time n , operation time at constant load ,if of electric braking ,time at rest flux max ,sequence duty energy  inertie, driver operationel service factor = 1.0 means,: frame size quantity power : damage disconnect input power to mazimun at no load in dB , 60 Hz ,poles 

-engineering electrical motor ,lecture disck ,,start delta direct load  transform ,,

Principle command ,alforigrame sequence diagram:

Starting with start Delta dusibleb,control F1,f,2,f,3',t1 control transformer k# , contactor ft@ ,overload ,relay h control buttons kt@- time , motor optionak ,motor optional accessories phase  relay ,Mmeter , voltmeter ohm metre start delta switch , start delta for output , 220v ,380v , fusible ,,

F.23 control fuse t2 control transformer k1,.k2.,j3 ,contactor ft@ ,0verload relay ,control options ,accesorie ,phase failed relay ,operation dusk , ,,


- wath happen switch off on circuit : 

 



     

     

     

     

- automatic lineare, performance b, correction advance ,retard ,,

Regime transitaire ,system lineaire echelon, u(t) = consts ,yv( t) detector, 

Y(t) - instant to ,x(t),EST v <oxx ( t) ,,

y(t)=a.u(t),,y(t)=a.u(t.t),,,y(t)=,system stable , systeminears principal support, a.u1(t)+by(t)-s St,linear ,,,

-equation electric system ,,

VE= RI+vs ,system ,RC,dv Constance system lineare Laplace , y(t)=(u.h).(t),,


-radio technic:

Input signal or port yes module yes output signal module ,

Type modulation, u= u .sin ( w.t+alpha),,

Value instant  u, u value cret ,a = 2.f, ,u ft

Signal modulator ,module am signaj, oscillator portets,dephase 90 , signal bmodule an 

- system dipole ,:U(t)=R.t(I)..conductor / bytb,G= IR, sirmence , system linear ,non lineare , Fae+ b = a few+ b fser ,signal input , output ,system , filter box input ,output  ,x ( t ) - y( t) s, system linear ,hb( h.t ) ,, y ( t ) = z + 1 , HT ( t: u) , impulsion reponsrb,x( t) ,,peig d diract ,circuit transfer input output , h (s) ,  x ( t )= 0 , y y ( y ) = H (s)x(t),,y(t)= yo.j = x (t),,

Variable x= y=s(s), ,,m1,n=,time,x(t) signal continue, signal siscret , period instann,x power,,transformer forever ,x(t) ,TF x ,, measure frequency, sum ,x ( q) ,,linearity ,TF(x+y)=TF(x)+TF(y),,

 Time frequency,decalsgrb, TF = real time strictment passif ,,TF,TF( x ( ti to) 

- derivay note x(t),DX =St ,TF [xo(t)]z+1,,

Integration x ( t) derive exponent: 

Dilation temo Freq, TF(x(t(t)=z+1,1

RF+x(t)],TF,, conjugate comply calculate, tfx ( t) ,signal real ,x(t(=x(t),,

- transformer ,fans f 2 function derivable interval [a,b],,derivation contrate ,for + t) g + GT) dt = [ ft) ( GT ) bib,zba ,,f(tb) go ( go ) [f(t)g.(t).b,,a = f(b).g(b).I f(a).g(a)..

- transfort Fourier ,transformer ,

TF[xoy(+t)..,,,TF(x a y ) ( t ) = z+1,, z -1,1 Hz ,,x(u).y ,,x(t)=X1(t)+jx2(t),,

Impulsion diract ,,I(t) impulsion diract transformer  limited , signal finie , retention T f [xoy(t)], neutral element ,,


-Signal information package database analyse lab Matalab  : sinus peak ,control loop

-Signal information  management , signal management recruitment signal radio control alpha Bravo Charlie informers recruitment signal time discret ,

Signal communication skill admiss  persal signal radar detector signal information data source criminel

Input signal information system manage value x f(x) variable ,or f( y) ,research methode investigation signal, traffic control signal management information f(x), ,,signal traffic control loop  x saturation way over load way operationel traffic stop sign ligthning f(x), signal relation vehicle circulation vehicular linearity  TF ( x+y ) input compare to way and gate evidence outfuo library trial separe TF(x)+TF (y) ,incidence speed collineare, ,,informer asssessment in order Manuel electronics system  derivation conduct misconduct  integration system by,,

Signal information real  system management investigate 

 





     

     

Information management system,security ,2,3cctv information duty ,2,9rado security information,zone security alarm  detector response control room supervisor Assessment management,information metal detector 2alarm signal radio, compare linear or non lineare,,information system management police  docket system Manuel compare  evidence supplies sabs close delivery compare supply scaling police marking information marks process  CCTV police alarm police calender daily manage court warrenty,,input alarm x ,clocking ocloc time t(t) ,,, y(x) ,variable , statistics report ,source data  crime report ,occurs ,research integrity ,research deviation ,research transfer equation  linear synchronouse system ,,intelligence system 





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















































1.purpose .memo
________________________________________

1.A-A.    /B.B.        //CC. ///. DD////EE //////tra outcom  A+B+C+D
________________________________________

1.A-A	B.B.       	CC.	DD	EE	A+B+C+D+c

.VR1=IA×RA /  Vout = - Vin0,6//IE=Ib+ic
VR2=(I1-I2)×R2/
VB=R3×(I1-I2)R2/
VR4=I2×4/ 
	Vout = - Vin0,6//IE=Ib+ic			Cu+2e---Cu 0,7 si	1.VR1=IA×RA /  Vout = - Vin0,6//IE=Ib+ic
VR2=(I1-I2)×R2/
VB=R3×(I1-I2)R2/
VR4=I2×4/ 
                   ...       /////Cu+2e---Cu 0,7 si

2.ET= R1.I1+R2.I2+R3.I3 / ET = E1=E3
   ET=I1.R1+I.R2/
    ET=I1.R1+(I1-I2)R3/
    
	R=£×l÷a. 
R=£×l÷A	IC(sat)=Vcc÷RC	C=k×A×€o÷d		2.ET= R1.I1+R2.I2+R3.I3 / ET = E1=E3
   ET=I1.R1+I.R2/
    ET=I1.R1+(I1-I2)R3/
    
/R=£×l÷a.  ///IC(sat)=Vcc÷RC///R=£×l÷A

                                            ///// C=k×A×€o÷d


3.IT=√IR° +(IL-IC)°//          R1÷R2=1+&0T1÷1+&T1
   
 VT=√VR°+VL-VC//
    Z2=√R°+L°-	// Rt   =To[1+&t].     	RC= vcc- VCE÷IC,     
///RB= Vcc-vbe÷ib
B= IC ÷IB      

	for=1÷2π√L.C		3.IT=√IR° +(IL-IC)°//          R1÷R2=1+&0T1÷1+&T1
   
 VT=√VR°+VL-VC//
    Z2=√R°+L°-C//

    // Rt   =To[1+&t].     //// for=1÷2π√L.C
/// RC= vcc- VCE÷IC,     
///RB= Vcc-vbe÷ib
B= IC ÷IB      


4).XL= 2×π×f×L. / Is=IT-IC//

     XC=2π×f×1÷2π×c/ A=π.d÷4//
    1÷To=1÷R1+1÷R2
    Rp2=1÷R4+1÷R5/. d=√4pcl÷πR
	///RB= Vcc-Vbe-ve÷ib
   Rbl=Rb2.(vcc-vb÷ib)÷ib
Rbl= 1÷10, Re.vbe
	N1÷N2=V1÷V2
//// Re= R1+R2(N1÷N2)
			
5) 1÷Rp=1÷R2+1÷R3 / RC= Vc÷IC //)
     Rs1= R1+Rp1
     Rp1= R4+R5÷R4÷R5	  V×R2÷Rc+R3	I1( R e× cos π+-xe×Sinπ	Ze= √ Re°+Xe°/////
XL=Z2=Z3		5) 1÷Rp=1÷R2+1÷R3 / RC= Vc÷IC //)
     Rs1= R1+Rp1
     Rp1= R4+R5÷R4÷R5/
  V×R2÷Rc+R3///
I1( R e× cos π+-xe×Sinπ. /////
Ze= √ Re°+Xe°/////
XL=Z2=Z3/////

6).RSH=Im-Rm÷Ish/
RSe=V÷Im-Rm			VL= VP÷Z1./////
IL=√3×ip////
Pf=√3×IL×cos /////
V=4,44×Ns×pf
		6).RSH=Im-Rm÷Ish/
RSe=V÷Im-Rm/ 

VL= VP÷Z1./////
IL=√3×ip////
Pf=√3×IL×cos /////
V=4,44×Ns×pf

7..E=e1+e2+e3../
       E= e×n ../
   1/Rp=1÷R1+1÷R2÷1÷R3/
I=emf÷r.t/
 E÷r×n)+R
V=I×R/
RB = hxr, I=V÷R ,,
Vo = I× RP 
IT= IA+IBc
	Y= m×N2÷N1./////
S=V×I×√3	S1.÷s2,zt=1÷1z1+1z
			..E=e1+e2+e3../
       E= e×n ../
   1/Rp=1÷R1+1÷R2÷1÷R3/
I=emf÷r.t/
 E÷r×n)+R
V=I×R/
RB = hxr, I=V÷R ,,
Vo = I× RP 
IT= IA+IBc

Y= m×N2÷N1./////
S=V×I×√3.////
S1.÷s2,zt=1÷1z1+1z

8.)VRM=1÷√2
VM=0,707
VDC=2÷π
I=0,318 ×Ia×z×p×¢.	B=¢÷A,H=IN÷l
F=B×L×I
E=∆¢÷∆t	E=B×L×V	IT= VC=I÷RC. £Integral V.b	XL=2×π×f,   XC =1÷2×π×f, VC= IT(-jxc	8.)VRM=1÷√2..  / B=¢÷A,H=IN÷l//
     VM=0,707.    / F=B×L×I //
      VDC=2÷π.     / E=∆¢÷∆t //
I=0,318 ×Ia×z×p×¢. ./ E=B×L×V//

IT= VC=I÷RC. £Integral V.b ///
XL=2×π×f,   XC =1÷2×π×f, VC= IT(-jxc
) 1÷2×Ia×z÷2P/
     IZ÷2c2p×40÷3c /
ATp=IZ÷2C2p×(1-40÷360)÷/
	V= K×Q×÷r//
C=€r×Co×A/d//
F=k×q1×q2÷r°//
Ek=Q×V	XL= 2×π×f////
IL= v×I÷j×xl////
	Z2=1÷hoe-hf÷h.e+Rs/////
Z2=1÷ hoe/////
		9) 1÷2×Ia×z÷2P/
     IZ÷2c2p×40÷3c /
ATp=IZ÷2C2p×(1-40÷360)÷/

V= K×Q×÷r//
C=€r×Co×A/d//
F=k×q1×q2÷r°//
Ek=Q×V//
//
XL= 2×π×f////
IL= v×I÷j×xl////

Z2=1÷hoe-hf÷h.e+Rs/////
Z2=1÷ hoe/////

10.)N=v-Ia Ra÷k.¢/  Q= v×π×d

  B=u×o×l÷2×r/

					
Q =V÷d  , 	C's=C1+C2
V=celerity×f	1÷C=1÷C1+1÷C2
1÷C=1÷C1+1÷C2,	AT=Q1=Q2=V×C
W=1÷×Q×V

So=h×f×e×R×L÷n×e		Q =V÷d  ,  ///     1÷C=1÷C1+1÷C2
1÷C=1÷C1+1÷C2, ///C's=C1+C2
V=celerity×f //AT=Q1=Q2=V×C
W=1÷×Q×V

So=h×f×e×R×L÷n×e///
Zo=R×c×h ///
RL//Z2=ZL. ///

13.)  NP÷NS=VS,/  m=EQ(V÷r)×Q
       emk=B×L×Vb/
	Qse=Qse=QT=Q1=Q1Q ///
     Qp= Q1+Q2+Q
Co = C1+C2
	Z2=RB//Rb°
T1=RbT/RbT
			13.)  NP÷NS=VS,/  m=EQ(V÷r)×Q
       emk=B×L×Vb/
  ///Qse=Qse=QT=Q1=Q1Q ///
     Qp= Q1+Q2+Q
Co = C1+C2
///Z2=RB//Rb°
T1=RbT/RbT

1÷ZT=1÷R-j(1÷XL-1÷xc)/////	IT=IR-j(iL-IC)/////	a+jb= √a°+b°/////			
£=∆L÷L,  ,£=S÷E	Rt=A×e °  exp .Bt.//
VA=R2÷R1+R2	L=AR÷R÷∆L
=R=resistivity	R÷teta= r(cos Teta+j sun Teta)////
f=1÷2π×√1÷Rc-R2÷L2,zd =K
		£=∆L÷L,  ,£=S÷E//
    L=AR÷R÷∆L
=R=resistivity ×L÷d//
Rt=A×e °  exp .Bt.//
VA=R2÷R1+R2
R÷teta= r(cos Teta+j sun Teta)////
f=1÷2π×√1÷Rc-R2÷L2,zd =K

16). tanπ=√3(P2-P1)÷P2+P2	V=w.×sin(π+Alpha	I=w×sin(π+alpha	ZT=VW<alpha÷Iw÷alpga		16). tanπ=√3(P2-P1)÷P2+P2//
V=w.×sin(π+Alpha)///
I=w×sin(π+alpha)////
ZT=VW<alpha÷Iw÷alpga
17) Fr=1÷2π√LC
Iaveg=I1+I2+I3+...In÷n	I rm÷wg=√i1°+i2°+i3+ ...in//
f=1÷2π×√L.c				17) Fr=1÷2π√LC
Iaveg=I1+I2+I3+...In÷n//
I rm÷wg=√i1°+i2°+i3+ ...in//
f=1÷2π×√L.c//

18). Vs =π int .b,a ( y°1-y°2)
Am-y=into a b (rdA)
	n=2×π×Nr×(w-S)÷60×I×v ,
	I2(I1+I3)Ra+(I1+I2-I4°)×Ra+(I3+I4)
Rotor = (I1+I2)V-(I1+I2-I4)°
	N generator = I×v÷Iv+Ia°×Ra+Ra+Is×v.I)
		18). Vs =π int .b,a ( y°1-y°2)
Am-y=into a b (rdA)

n=2×π×Nr×(w-S)÷60×I×v ,
Efficient=√I1÷I1+I2
I2(I1+I3)Ra+(I1+I2-I4°)×Ra+(I3+I4)
Rotor = (I1+I2)V-(I1+I2-I4)°
Efficiency motor = N-(Ia°-Ra+(Ia×v+Is×V÷IV)

N generator = I×v÷Iv+Ia°×Ra+Ra+Is×v.I)

19. C=Q.n√£×L÷2π// v(dropR)R
	€=N.∆¢÷∆t //.
€=N.∆¢÷∆t	(drop ) L.I .///
I=Q1.I1
V= Vb-Va//. V(drop -total	Ta.alpha flux × Is ) .
EbO×flux×N×Z
Efficient=output ÷inpout
		19. C=Q.n√£×L÷2π// v(dropR)R

€=N.∆¢÷∆t //. V(drop ) L.I .///
I=Q1.I1
V= Vb-Va//. V(drop -total .///

Ta.alpha flux × Is ) .
EbO×flux×N×Z
Efficient=output ÷inpout
________________________________________

20. vb-vA .Q÷€×d÷A.//	output
/// Copper loss = I×T
Efficient= 1-				20. vb-vA .Q÷€×d÷A.// z peak maximum
Voltage input ÷output
/// Copper loss = I×T
Efficient= 1- losses ÷imput

.(derive partial .p÷ derive partial v)=
(Alpha.p ÷alpha .v )(alpha .p÷aplha.t ) /
	
Iaveg=Vave÷r//
VAve= V.ave÷r.l//

	Line = VRY- ,VRY=VYB
I1=IR-IR
			21.(derive partial .p÷ derive partial v)=
(Alpha.p ÷alpha .v )(alpha .p÷aplha.t ) /

Iaveg=Vave÷r//
VAve= V.ave÷r.l//

Line = VRY- ,VRY=VYB
I1=IR-IR

22)_G(JW)//inte  v2, V1..p×d×	R.T into v2.v1 .dv÷v ..//
	n.1/n2= sin .π/sinπ
			22)_G(JW)//inte  v2, V1..p×d×v //
R.T into v2.v1 .dv÷v ..//
n.1/n2= sin .π/sinπ

23).P= m×R×T÷T/.
	= s.///
f= n×T÷2×l.		Z2√R2°+SXo°///
		23).P= m×R×T÷T/.
 Eo÷v1=zr÷zs.  I = s.///
f= n×T÷2×l. ////

Z2√R2°+SXo°///


24).T= 2  π, √ l÷g	F=^R).For.q1.q2				24).T= 2  π, √ l÷g //
T= 2 π√m÷k.kg//
F=^R).For.q1.q2÷R.R
25)..C.c÷f.m












26). Int.int.int x×y×z Dy
       Int  5. ,1. Int 1.2 into 3 0 [ x °×y×Z
	m///
F=q2×E1				25)..C.c÷f.m =C.C÷F.m.c×v÷m///
F=q2×E1
//
________________________________________
26). Int.int.int x×y×z Dy
       Int  5. ,1. Int 1.2 into 3 0 [ x °×y×Z
27)..Ns- N÷Ns//. R/2///
R2=SX0	///.
 VR= = VL×√3

ZBC=zAV=R÷#+jxl÷#///
	÷#+jxl÷#///
VBC=I2×ZBC ///
VC= VBC+VR			27)..Ns- N÷Ns//. R/2///
R2=SX0 ///..xl/2 ///.
 VR= = VL×√3

ZBC=zAV=R÷#+jxl÷#///
VBC=I2×ZBC ///
VC= VBC+VR////


28)P= √3×VL×IL×cos
S= √3×VL×Is/
Er= I×zs		Er= I×zs/
zS=Ra+jxs////
Er=IZs
IRa			28)P= √3×VL×IL×cos/
Eph= er/
S= √3×VL×Is/
Er= I×zs/
zS=Ra+jxs////
Er=IZs
IRa
30.). P=not (w-p)×9,81×π×D×n//
	IR=V÷R
IN=√x-coml°+ I y comp
IN=	IN= IR+IR+It+IB
________________________________________
			30.). P=not (w-p)×9,81×π×D×n//
IR=V÷R
IN=√x-coml°+ I y comp
IN= IR+IR+It+IB
________________________________________

30.). P=not (w-p)×9,81×π×D×n	IR=V÷R
IN=√x-coml°+ I y comp
IN= IR+IR+It+IB
				30.). P=not (w-p)×9,81×π×D×n//
IR=V÷R
IN=√x-coml°+ I y comp
IN= IR+IR+It+IB
________________________________________

31).O= 
Efficient max = k×s×cos©÷k×s×cis©+Po+k°ps	F= N.P÷60///EP=E.L÷√3
				31).O= 
Efficient max = k×s×cos©÷k×s×cis©+Po+k°ps/

F= N.P÷60///EP=E.L÷√3

32)
d= E÷ 2kf×kd×kp×fz//
Cos©=O÷s	Mean = 3×√2÷h - vline	- vline ////
= √2÷2×2π ×vrm÷(1+cos©)////
Vmean=√2÷π×vdm×(1+cos alpha	Vmean=√2÷π×vdm×(1+cos alpha)///
Vmean = V× √3.√2÷2π<+(1+cos alpha)/// ,,R= T-T2÷p=1/t		32)
d= E÷ 2kf×kd×kp×fz//
Cos©=O÷s//.vmax = √2×vrm //
Mean = 3×√2÷h - vline ////
= √2÷2×2π ×vrm÷(1+cos©)////
Vmean=√2÷π×vdm×(1+cos alpha)///
Vmean = V× √3.√2÷2π<+(1+cos alpha)/// ,,R= T-T2÷p=1/t//

33).T= 3×s×E0°/2×π×n (R2°+[S.Xo]°/
A=B(ie exp I/t
V=E.(I-#e-exp t/t ) 
	VD= vs ×(R1÷R1+R2 - R3÷R3+Rth)//)
Vi= iL×R1×R2÷R2, =. 
	= Vo.R.m÷RM+RT
Vo= RC.dv1(t)÷dt=1÷RC. Inte V1.(t).dt+dtvc(o			33).T= 3×s×E0°/2×π×n (R2°+[S.Xo]°/
A=B(ie exp I/t
V=E.(I-#e-exp t/t ) 
VD= vs ×(R1÷R1+R2 - R3÷R3+Rth)//)
Vi= iL×R1×R2÷R2, =. 
Vo×R1/R2
= Vo.R.m÷RM+RT
Vo= RC.dv1(t)÷dt=1÷RC. Inte V1.(t).dt+dtvc(o)

Gma×Gms÷r°=Gms×Gms÷4×10 exo 8. 
a=∆V÷∆t,  
Speed = distance ÷timr
V=u+V÷2	V=u+V÷2,,//
S=(u+V)//
V=u+at //
V° =u°+2as
S°=u°+2as
S=UT+1/2at°
F=m×a
	S=UT+1/2at°
F=m×a
F=m×g+m×g
F=m×g-m×g
			39). Gma×Gms÷r°=Gms×Gms÷4×10 exo 8. 
a=∆V÷∆t,  
Speed = distance ÷timr
V=u+V÷2,,//
S=(u+V)//
V=u+at //
V° =u°+2as
S°=u°+2as
S=UT+1/2at°
F=m×a
F=m×g+m×g
F=m×g-m×g
________________________________________

m1×u1+m2×u2 //
M1×.u1-m2×u2
V= w×r //
V=2×π×n×r÷60.
V= π×n×D÷60
D+(t+t)÷2  //
Fe=T1-T2 //
P-Fe×π×D÷60	D÷4×4-π×d°÷4
V=π×D×n÷60
Belt velocity
V= π×D×n
nA×DA=nA×DA÷nB
T=T=FX's
L= f×cos 
Moment=L×M
				m1×u1+m2×u2 //
M1×.u1-m2×u2
V= w×r //
V=2×π×n×r÷60.
V= π×n×D÷60
D+(t+t)÷2  //
Fe=T1-T2 //
P-Fe×π×D÷60 //
π.D÷4×4-π×d°÷4
V=π×D×n÷60
Belt velocity
V= π×D×n
nA×DA=nA×DA÷nB
T=T=FX's
L= f×cos 
Moment=L×M

Equilibrium

P×cos30°-

E1=T1-T2÷T1=100%//
P.v/p=vRt÷m=V°
n=w÷Q=2-T2÷T2×1000 //
P1×V1÷T2=P2×V2÷T2; //
Q=W=P1×V1×ln×V2÷v2 //
SF=4,187lntf÷273. //
T2÷T1=(P2÷P1).exp ,alpha-1÷alpha //
W=P2×v1-P2×V2÷alpha-1 //
	
P×V= m×R×T. //
Q1=m1×C×∆t ,//
Q=m×l×v//
U=m×CV(T2-T2).//
(V2÷v1)alpha = T1÷//

	Alpha= cp÷cv//
W= m×R×t1×ln(P1÷P2)//
(P1×V2 )=(P2.×v )//
f .exp 1=f×v÷v-v //
			E1=T1-T2÷T1=100%//
P.v/p=vRt÷m=V°
n=w÷Q=2-T2÷T2×1000 //
P1×V1÷T2=P2×V2÷T2; //
Q=W=P1×V1×ln×V2÷v2 //
SF=4,187lntf÷273. //
T2÷T1=(P2÷P1).exp ,alpha-1÷alpha //
W=P2×v1-P2×V2÷alpha-1 //
P×V= m×R×T. //
Q1=m1×C×∆t ,//
Q=m×l×v//
U=m×CV(T2-T2).//
(V2÷v1)alpha = T1÷//

Alpha= cp÷cv//
W= m×R×t1×ln(P1÷P2)//
(P1×V2 )=(P2.×v )//
f .exp 1=f×v÷v-v //

.) ax+bx=
(p+x).(a×(p+x÷p×)+4(p+x÷p+x)
[X.by×(p+y÷p+x	2x×x+33x+xdx
(a+b) exp 3= a.a +2a.ab+2ab.b+b.b 
				40.) ax+bx=
(p+x).(a×(p+x÷p×)+4(p+x÷p+x)
[X.by×(p+y÷p+x)

2x×x+33x+xdx
(a+b) exp 3= a.a +2a.ab+2ab.b+b.b 
(a+b)exp n= combination c n to n . a exp n-1+c 


					41) F ®--->[gain]--->[op.Am]-->[DC motor ->-->.         Tachometer -------------------------|

  --->[G(s)]-->[G.p]-->©-->[Gp(s)]-->[G's(s)]

-->©-->[controle]-->[power conver]-plan 
    ® Intrusion  --- [ideal sensor ] 

©-->®-->®--->[G1]-->[G2]--[G3]  -->[G4]-->®
                                                          [G5]
                    ..                                     [G6]
©-->©--->[G1.G.2G.G4..]---[G5]---
              -------H0------------     [G4]--- 

               -------H6------------    [G6]---
            

____42) so=X1+x3+X5
S1=x2+x3+x6+x+x7
S3=S4+X5+x6

Input                     	Input                      / output/ register
X1,x2,x3,X4,X5,x6,/S2,S1,s0/select
0,  0,0,0,0,0,0,0,0,0,/0,0,0
 1,1,1,1,1,1,1,1,1,1,1/1,1,				
-________________________________________42) so=X1+x3+X5
S1=x2+x3+x6+x+x7
S3=S4+X5+x6

Input                      / output/ register
X1,x2,x3,X4,X5,x6,/S2,S1,s0/select
0,  0,0,0,0,0,0,0,0,0,/0,0,0
 1,1,1,1,1,1,1,1,1,1,1/1,1,1

					
					
					
					
					
					
Total value 					
41) F ®--->[gain]--->[op.Am]-->[DC motor ->-->.         Tachometer -------------------------|

  --->[G(s)]-->[G.p]-->©-->[Gp(s)]-->[G's(s)]

-->©-->[controle]-->[power conver]-plan 
    ® Intrusion  --- [ideal sensor ] 

©-->®-->®--->[G1]-->[G2]--[G3]  -->[G4]-->®
                                                          [G5]
                    ..                                     [G6]
©-->©--->[G1.G.2G.G4..]---[G5]---
              -------H0------------     [G4]--- 

               -------H6------------    [G6]---
            

 
Purpose : 

Trade to trading: 
Basic trade theory fundamental design
Requirements: 
Construction trades composition:
Operationel trade low rules applier skill to skill ,, trade to trading 
- 40. Sabs code of practice wiring premise 
Safety trade tools , safety I rules harss play machine building sign fire smoke , injuries fire hazard warning,cut space save damage of good is prevented.
- cell advance construction simple efficiency full load 976 silent operationel
Transfo little care :
-open -air cooling oil cooling,
- transformer assumed to have no loss et.
- secondary cell advance.
Rechargeable greater capacity than primary cells ideally suited emergency back a application la get life disadvantages more expensive than primary cells regularite maintence periodic charging traditionally less suited for portable application, 
- code colour resistance : sketch IEC circuits carbon resistance potential variable capacitor ,zener diode ,pnl transitor battery cells  unity polarization carbon extrinsic 
- yoke ,poles sgoeas bushes backwards in the motor .
- moving brushes in generating poles field poles series .
- number of pairs of poles used .

- strength magnetic field.
- rate magnetic field .
- rate magnetic flux cut by the moving conductor,
- number of active conductor
- effective field flux reduce armature as load , generator.

41. Engineering drawings : welding PC aides draugtinf ,join metal 
Free hand boxe screw thread ,
- arc welding gas welding resistance lap joint ,T joint career joint butt joint 
Screws threads ,
- correct linework accuracy neat first  angle orthopedagogic projections coupling projection machining .

-42.. generator ward Leonard motor generator system .
Shunt generator  used  where constant voltage is required.
- series generator a booster on DC line transmission line ..flux armature .

43 .trade domestic appliances .
- washing machine immersion water heater protection steel conduct pipe earth
44. Moving - iron instrument ,non linear scale measure DC and AC cheap robust  affected by stray damping by air .

45. Moving - oil  instrument  linear scale measure only DC expensive very accurate damping edy ,

46. Convert AC generate output to a pulsating DC and act as , period time peak value cosinys 
.47. illumination high pressure Mercure vapour discharge lamp ,siduim vapour discharge.
Cold cathode neon 
- AC current theory, Serie RLC impedance phase , 
-48. three phase AC system wave supplies

-49. transformer secondary primary.
DC machine test conducted.
. switchgear and protected device :funci and operation induction.
-50. application of induction disc relay . current and voltage break capacity of .
-51. reverse phase relay rotation of a three phase line they operate differente power level usually work by a solid
52. AC machine,
53. Measurements instrument electronics.
53 material used in manufacturing of semi conductor device.

54.Special characteristics :
Def . special arc furnace transformer power requirements.

55. Control system like all other components on electrical network allowed for process to monitored and regulated from a remote.

56. Special characteristics controle system operate environment controle system or overall electrical.
Static controle analogue .

57. Electrotech : principle nuclear positive Lenz lot directly proportional yoke download.
Type material algebraic sum EMF
- principal low change in the magnetic flux linking with .

- movement of conductor in a magnectic field.
- increase decrease of current circuit .
- carbon brushes, graphic brushes ,electrographes, copper graphite.
Separately.
Efficiency full load ,97 moving silence magnetic circuit winding oil tank protection refrigerator,
-57.1 connecting electrical machines practical tips .
For connecting
-make sure you have rigth joint check size of logs .
-Make sure that you have a crimping .fit purpose joint.
* Installation core operate .
* Installation care and inspection of equipment locating t rigth tools for the jobs..
* Make sure that you know which equipment,keep your tools box organise stored your tools safety clean in good working..
* 57.2. Generation and supply of AC 
* Power fact correction low power factor increase decrease electricity bill.

*Method of power correction capacitor basic generate another method of power factor synchronouse motor be set to operate in logging on leading.

* AC synchronouse machine synchronouse converted mechanical energy input induction machine.

* Load Brid connectioning charge series connect to shunt .
- speed and torque:
- torque and power: 
Load sharing divider load .load among a set
to TSHINGOMBEKB 
 

50.  Cpd continue learner 
Calculus 
Kirchoff,
e1 source, RLC Serie RLC ,i1,i2,

dq÷dt=1.//
       L×d°1÷+R×di÷dt+1÷C=dv÷dt
(L1+L3)×d°i÷dt°+(R1+R3×di÷dt+(C2+C3).i1-ld°i2÷dt-R×d1.2÷dt-c3.12=e1(t) ///
(L2+L3)×d°i÷dt°+(R2+R3)×di3÷dt+(c#+C3)i-L3×d°i#÷dy-R3×di÷dt-C3.I1=e2(t)
t=0 and t  ,t = o,2 
e1(t)=100 sin ( 120πt)...//

Contour symetry 
Int .H×dl= sum n1.- n2.i2

-;a Di/dt detection circuit DC unidirectional breaker
di÷dt ,  V= R.i+ L×di÷dt
I=(v÷R)×(1-e(t/T),, differential the above value , diffential the above value Di/dt = ( /l ) e  the maxim , di÷dt , di÷ St ) max = //

Potentiometer coefficient reducing voltage integration factor,sum integrator

dx÷dt. TV,,E=™L.dl÷dt ln +1÷c integral dt./
L.di÷dt+R.i+1÷c integral.1.dt=E
Rd÷L.dt+1÷Lc=d.y.÷dy
do=[f(x-y)St..
5dx÷dt+3x÷5=0 out put integraj dx÷ St = 3x÷5

________________________________________
∆v2~ V20-V2=Rs.I2 cos alpha+xsIs2.sin..alph
√4∆v2=√3(RsI2.cos &2+X's.I2.sin &1/
_______________________________________&
f(x)..d÷dx ×f(x)../

d÷dx[f(x).g(x)]=f(x).(g)+f(x).g(gx)/
a(x+b÷2a)=a+x.x + b÷a+(8.b÷2a).(8.b÷2a)
A÷ax+b+B÷(ax+b)(ax+b)+...G÷(ax+b)n//
Firs second order 
dy÷dx+Pay=Q
a.(d.dy)÷dx .x+ b.dy÷dx+cy= f(x).//
Volum.∆v π✓.y.y∆x
VX= π. Inte.   b to a  ( X1. - x2) dyb..vy = 2π 
X= Am-y÷A= into ( a)  .to b  rdA÷A, //
(x,y)= f(t); g ( t) 

ds÷dt= v ,,dv÷dt=d.d.s÷dt//
   f(t)S= 40t-5.t.t
  f(t')=40-1,,0.t
,M 
_____________________________________
V= π.r.r.h//.   r.r=L.L - h.h  cone high
X.x + r.r=
R.R..V=1÷3.π.r.r.h
=1÷3.π(R.R-x.x)(x+R)
=1÷3π[R.Rx+R.R.R-x.x.x-Rx.x]
Partial differential
Z= x.z.z.y.t..
Lim .&v1÷&h=dv1÷dvh.
h---
Volume cylinder ,,,unit
Gravitation centroids..
x+y+z=1
Double integraj ,,single variable sum I.double,  Sum j f ( x,y)
y=(x,u@),,,y = (x+§x)and ∆y-y , u (x,u) total derivatives, 

dy÷dx=dy÷du×.du÷dx.. partial constant 
________________________________________
 

	
	
Inspection of work equipment :to indentify wether equipment can be operate maintened safety , deterioration ,
Check  risk safety in case  inspect where significant and safety installation, installat reinstallation deterioration or any other need for  inspect frequency should .
- to operator and other equipment installation result 
- work equipment  that requires inspect inspection b, 
- reg where the safety of work equipment depend .
- intervat equipment.

- wath should the insoy. Deoenyon type of work t use  any manufacture recommanday the advit ,trade sourt
________________________________________
Load 
Z total = ZC/)(ZL-UE)..
VA=S+VAR -------.S=P+Q//
________________________________________
gradient DE fonction
&f(x,y)÷&y=&÷&y,,
ln|x|+C
Int  I/x 
_______________________________________
Equivalent transfo
E2= I2×z2+v2 //
E2-v2=I2×Z2..//
Vre%=E2.v2÷v2×100..
V reg%=I2.R2.cos$+I2.x2.sin$÷v×100
________________________________________
 

	
	
	
	
50.) Cpd development skills 
Z=√R-(xl-xc)/
V=Z×I,
Z=R
P=R×I×I
V=U×√3
I=j×√3,
U=V÷√3
I=j×√3
J=I÷√3,
J=IL1, IL 2,
3R..,, R :3,3 xl = 22×π×L×f
XC=1÷2πcf
ZT=(1÷z1+1÷Z2+1÷Z3)..
     =G1+G2+G3)..1÷GR+1÷G2+1÷G3..
________________________________________
Engineering electrical load system 
R.I.I=3×R×I×I,, R×I×I÷3,,
I=j,,
I=j×√3
Load = √3×R,,I=j,...
I=j×√3
P=R×(j×√3)
P=R×j×jx3
P=3.R.j.j
E=1÷R1+1÷R2×j×j×t
E=R1+R2×j×j×t
E=3×(1÷R1+1÷R2)×j×j×t..
E= Em×sin×w.t
E=j×3×R×j×j×d
Z=1÷Z1+1÷Z2+1÷Z3---
G1+G2+G3,,Z1+Z2+Z3
E= 3(Z1+Z2+Z3)(j.t)
E1=(1÷z1+1÷Z2+Z2)(j×t)
E1=(1÷Z2.1+1÷Z2.2+1÷Z3.3)(j×t)
E2= 1÷Z3.1+1÷Z3.2+1÷Z3.3(j×t)
ET=[1÷z1+1÷z2+1Z3(j×t)]×[(1÷z2.1+1÷Z2.2+1÷Z3.3)]×[ 1÷Z3.1+1÷Z3.2+1÷Z3.3](j×t)
[Z1+Z2+Z3(j×t)]×[Z2.1+Z2.2+Z2.3]×[Z3.1Z2+Z3.(j×t)]
ET=1÷Z1+1÷Z2+1÷Z3(j×j×t).×[Z2.1+Z2.2+Z2.2+Z2.3][j×j]×[Z3.2+Z3.2+Z3.3×(j×j.×t)..//

Serie LC ,,capacitor ,resistor load series parallel impedance  low related 
Evidence low energy 
C=Q/4
..Q=u×.u×.c×w
Delta=Q=3×U×U×C×w
We=3×V×V×C×w
CV=3×c∆
________________________________________

Ep=P×I.
E1+E2+E3..
Q= iron ,,Qin iron 
I A,,IB+IC,,  
Diagram fresnel ,,
In =I1+I2+I3,, alph=0,,V2N,,alpha =90!
 Vector fresnel 
I1=j1.2-J.3.1
I2=j2.3-j1.2
I3=j3.1-j2.3
I1+I2+I3=0
Delta balance 
I=√3×j×
J1.2=U=z.1.3

V=u√3..IL=Iz ,,
VZ1=VZ2=VZ3..I=Z,I=IL=IL=V÷Z,,V=U÷√3.
IA=V÷<,Z=U÷Z √3..
I∆=U÷Z×√3,,U.√3÷3.Z
Istar l = V÷z=u÷√3÷Z,, 
I stars =u÷z×√3=u×√3÷3×z
Part= U×I×start×√3×cos alph ..
 

	
	
60. Engineering sinusoidal quantity sinusoidal
- Um× cos×(wt+j)..
- 1÷T.integral T.udt ,,u>2÷T.
u= u.o×2×cos(wt)
I= i.o2 cos(wt+j)
P=u.O×.cos w(wt)×Io#cos(wt+j)
=u.o2×I2×o×cos(w.t)×+(wt+j)
U.2 .Io2×cos(w.t)(wt).×(w.t.j)(w.t.j)
Layout fresnel
V1=Vo×2×cos(w.t)
V2=V×cos×(wt-2p/3)
V4=Vo2.cos(w-4/3)
.P=(U×I×cos)+U×.I ×cos(2wt+s)
K= P/s,,K=cosj.,A=V×I×sin
Cos(2wt+j)+cos(#wt+j-4p,cos wt+,j-8P/3..
-process high 
I2.2= J3-j2, I 3 =j1-j3,,= Vk.3
-balance Pbskance ,
P= v2+v2.I2+v3..
V1= Vo2÷cos(wt)×i=io2×cos(wt+j)
V2=vo2×cos(wt-2p/3)
P1=V×I[cos j+ cos ( 2wt+j)
P2=v×i[cosj+cos(2wt+j8/3)
P<p<p1+p2+P3
________________________________________
Installation.cinnected load demand 
S=P+jQ./.P=V×I×c0s//A=V×i×sin@..√P.P+Q.Q
W=P×t
Demand factor = maximum demand ÷ connect.
Load factor= average load .

________________________________________
Applied engineering science skill mathematics : to master skill engineering 


Part , questions//operational skill///resonning skill .
Total skill total number
Mandatory skill development,//
Engineering context challenge,analyse with guidence , design development, investigation guidelines exist 
Basic knoy engineering key
-role engineering applications of civil mechy design implementation testing control system 
_ topic national simple system subsystem inoutb..
Application of low of conservation energy involving kinetic  losses 
-applied calculation involve work done
 and power ,Ew=FD,E/t ,Ek=1/2m.v.v,,EP= MGH, e= v.i.t,eh = = cm.∆.t .
_ roles and discipline impact // social environment Engineering b,social economic impacted positive  negative.
Description of function circuit in term input process  and output.
- concept current voltage measure ,,calculau involving relation voltage low involved resistance Serie parallel bdivider of operationel fixed voltage ,voltage divider to generated a signal
- transistor  resistance electronics , functionalite of relay protection in explanation function of electronic a variable voltage divider transistor relay output,.
-applying algebraic skill to linear equations bsolving linear equations the subject formuler,explanat draw graph of value for choses value ..
Conducting asssessment judging evidence marked and cerife quality assurance,re asssessor.
Question ,point process accuracy ,expected response correct answer award,correct additional,evidence divisyub overall strategy,level 5,4
____________________________________
Applied to applied applied skill to resolve skill physic mathematics to skill electrical panel :///

Principle operationel connection : basic transfib,threet phase  advantage application calculation 
Circuit ,power ,source protection device components,
-_____________________________________
Electrical infrastructure construction
 fundamt compulsy,electrical principle practice ,workshop practical , physical science electrical  , electronics industrial ,math ,life orientation level 2,4
Career work power station electricial 
- electricial system construction fault AC,DC  ,topic operate..
Introduct to policy theory policing practice v.literscyb,office data processing to

-Applied skill .to trade  and trade to applied skill thermal compo
_______________________________________
Applied.. 
 
	
	
 
 

Cpd.
Installation motor three phase , 3bulb
Power factor.
Qph1=uph1×Iph1×sin @1=>
Qph2=uph2×Iph2×sin@2=
Qoh3=uph3×iph3×sin@3=
QT=Qph2+Qph2+Qph3.
Determine impedance.
Ph1=uph,1ph,1,cos 
Ph2=uph2×cos×ph3.
P= ph1+ph2+ph4..
Equilibrium
Vab+Zs.ib=vab+ZsIa
VB+zS.Ic=Vbc+Zs.ib
VCA+ZsIa=VCs
Vab=vbc=vca=IB=IC=I
Za=zb=,Zo
Wa=Vab×Ia[cos(30°+q)]
WC=Vab×Uc[cos(39-q)
Wa+wc=va.ia[30+q)]+vbcI [cos(39°+q)
Wa+WC=vl.×il[cos3o+q)+(cos30) 
Wa+wcvl.il (2cos39cosq)=ox3×v×ipf

-dephasage ,129,2,4,6poles.
V1(t)=Vmax cos wt/

V2(t)=vmax cos(wt+3o)

V3(t)=Vmax cos +( wt+60)
V4(t)=Vmax(wt+90)


dv1(t)=integral .dv max ×cos w.t
dv2(t)= integral dv max× cos (wt+3o°)
dv3(t)=integral dmax cos ( wt+60).
dv4(St)=integral max cos ( wt+60)

Installation generator g1,2 ,transfib kWh , back ,,retard ,avant , clockwise
a+bi  .x+in,,complext real power factor ,apparent power factor
Cable line a+BJ,line impedance xj resistance capacity parallel connection cable phase beutraj ,power between phase phase impedance a+BJ ,,line 2,3a+bi,,line neutral a+bi,,
Construction vector fresnel flech robot scater  drawing vector quantity phase a,b,c ,,
_________________________________________cpd

applie to trade ,,and trade to applied safety ,percentage correct



Applied trade test to trade company theoretical mathematics trade theory relevant trade test on the job supervisor
Commision
Tender value point relate years expert certified trade,10point clear time frame ,clear project plan time plan responsible of team , equipmt workshop 
Log book apprenticeship 


Module trade to trade skill report phase trade ,1,2,3trade test,module code ,ovject,criteria tendered 
Safety area join crimping fault,,
Db fuse AC ,motor contruj retain fault fault applicable to oanrk fluirencr ,isolator ,wat unclud rack flexible conduct 
- code man power kefit, not dry joint no damage nibsolder all safety adhere testing instruments comductur 'cabkevmake off PVC armoured up to 16 mmm core 1200v lug join rating correct ovx 
- objectivjty relevant colour marking  correct sabs  caractere GAZ welding piece nozke GAZ lifting,.chain ton max work not exce selected readings all safety applied 
,- criteria recall type battery percentage parallel correct manufacte system fit component procedure ...
El1 electronic components wire watt carbon metal oxide thyristor 1construction solder..
-correct according manufacture adhere procedure cambdr correct test average value peak frequency RMS ,
Control main circuit line start dejtad phase rotation,
Cad xlpe cable 2099!test gears ,,
_______________________________________
Cpd learner technical.
 equivenlent electrical hydraulic ..
Integral countiur infinitive
Stock Maxwell 
 
	
	
50)cpd 
Engineering
-basic concept introduct discovery: voltage electrical current resistance power, conductor between line 
U=√3.uph 
- key term 
: law regulation ohm low state I= u÷ib,
-Kirchoff low:  sum bode junction in sum of current node n I1+I2+I+3(-,i4-I5.
Coulomb low :charge |F1|=|F2|=KE.|q1×q2.
-Right hand rules thum point direct point reminder finger field ,direction wire
First left hand  magnetic induction second left hand left magnetic
- resistors connection schema ,Serie connection /,parallel //connection
Amperage I1=I2=3,, // I1,2=I1+I2
Voltage U1,2=u1+U2//U1,2=U2=U
Resistance /R1=R1+R2//1/RR1=1÷R1=1÷R2
-Joule Lenz ,Q=I.U×R×t
Amount heater 
-capacitor connection
Scheme ,/serial connection //parallel
Electric charge ,q1,2=q1=q2//q1,Q,2=q1+q2
Voltage /u1,2=U1+U2//U1,2=U1=U3
Capacitance /1/C1,2=1÷C1+1÷C2
-ligthing .
Incandescent light type ligth /ligth socket base //lumen lux///color temperature////led start
Electrical ligth construction glass inert tungsten support ,electricligth order filament oxydatiin .
Manufacture 1,5 volt principal heating power / incandescent lamp source fluorescent lamp, halogen lamp ,#00-409 lamp ligth ,dischylamo service///ligth socket onde noise device ,E27,G4///lumen device total from////color temoera.
- cable and wire 
Wire color/ type of cable onde wire//wire connector///
IEC 60446 basic saftma. Power marking
IEC 60442007,2019IEC 60444
L1: brain Gray  ,// cable jacket ,wire coaxial cable signal cable flexible  filled heliax non metallic ribon cable insulator ,stripe /// Zone special danger soldering screw terminal ,4 wire nut terminal block  bolted wire connector
- db 
- power station and substation electrical substat.
Nuclear power: no install energy schema of operationel
Plan thermal power : generate electricak energy by converty chemical, hydraulic power wind ,geothermaj
-boiler turbine power combine cycle / spp solar power plan is on engineering structure verting radiation ponek ,Vacuum tower type ,/wind power Serie connected  osgir coak 
- electrical measure:
-Volmetre instri use mesure analogic principal electromechanic,AC DC pulse selected,
-ammeter device for current ,
- ohnneter device ,
Multiple measure functionalite analogic Dmm,
Clamp meter device operation based 10kv,,
Electricitt meter device measure electric energy install ,2,5:0,5:
- mechanical energy : transformer static device convert AC ,expansion insulator ,transform ,convert , 
Motor electrical energy operator shaft.-generstor convert mechanical energy  from AC power solar 
- ground system electrical engineering DC ,isolation ,n neutsj, potential
- Protection and automation device:circuit breaker,rcd,rcbo,sod,voltage monitor, magnetic,fuse,
- socket repair/ installation of build // installation the surface mounted socket ///switch installation
/Conche or brick,
1)make a Chanel for cable router installat junction,2 fill the recession labell,fill the remaining,drywall marking drill hide back connect,
// Fix the base to the wall cable into the base , connect fix the sock install,///
Make the Chanel for cable mounted installation the junction with alabase mirror install fill remayb ,,
-over voltage: l2 and l3 sum voltage u1+U2=389,P1+P2,
/IP code. IEC en 60528/ / time current character time (1,13In /// tools work wide socket 
-Schem electronics circuit
Switch (1and ,2 button,switch (3), pass through switch ,two key pass ,socket ,socket ,dimer ,motion sensor ,impulse relay motor connection,electricity 
- electricity cost calculator,device power used , daily use time ,price for kWh cost per hour month days week
________________________________________
Cpd ..basic  electrical continue
I1=V2÷|ztotal|<- teta
I2=V2÷ztotaj<(-120°-teta)
I3=V3÷Ztotak<(120°-teta)
V1.2=v1-v2=(VLN<O°)-(VLN<120°)
V2.3=V2-V3=(VLN<120°)-+VLN<120°)
=√3.VLn< -90°=√3.V phase V2
V3.1=V3-V1=(VLN<120°)-(VLN<0°
=V3=.Vln<(30°-tets)-(Vln<o
I1.2=V.1.2÷Z<(30°-reta), I3=V2.3÷zA<(89-o),
I3.1°V×3÷Z.a 
I.3.#=v.3÷ZA<150.
I1=U..,I1.2-I@.2<129,v3.I2.phaae 3+39°=


Advanced  system diagram on line 
Design draw breaker type .v max.max 


R=Z÷√x.x÷R.R+1, 
X=x÷R×R,
Full load no load losses ,phase fault,I3,phase 
I3phasrE÷X, ILG,,
I÷XS=1÷x+1:x+1÷x=3÷x,,
1÷RS=1÷R1+1÷R2+1+R3=3
X's=x÷3,, xs÷Rs=xR= gen ,x÷R, symetrical b


Voltage system power ,Q=P2×tan veta
 
	
	
Cpd,
theory experimental,
Task lab sign domestic and industrial 
Engineering lab electrical workshop.
Power systems.: electrical shop tools on precautions practical is discipline study design
- application equipment systeme
- solid bar copper wire
Awg, 10tp49,
Normal diameter 2,6to 0,76,cross section 5,39.
-assignment load calculation.
Application appliance / unity //power rating ///  rating ////daily usage /energy consumption.
Fridge tvfan. 

-lab electrical standard wire size..18aw to 1,
Assessment wath is gauge of wire used exercise b..
C=<1÷36.lin 
./_____________________________________
Engineering studies practice 240 ,month n diploma NQF 

Applie saqa to isat
Practical purpose Isat,scope mark off basic engineering,topic operate and monitoring grinder machine
Produced simple 
Sub task , activities time frame
Manufacture size,time ,marks 
Resource requie.

National 
Where appropriate material undertaking .
Labels.
 

	
	
cpd ,gov body insurance 
applie to applied trade to trade

Qualifications framework equivenlent
Assessment frame.
National diploma engineering.credit accreditation,policy criteria minimum maximum ,,
Equivenlent trade license: to trade license translate ..award equivalent
Divide job ,divide task ,divide operational equivenlent 
Phase preliminary phase final ,semi final 
Time competition ,cycle recycle permeant limited continuity function  hand book 

-Material strength test material 

-Scale ,module word 
Interpretation drawing scale drawing interpretation, building 
Geodesie 1/100,  1/10, projection reduce size mass rules  relate rules re projection planer scape
Scater ,flow shoot  
 
	
	
Cpd : body qualifications framework

Statement of work experience

Log book , instruction programme national , level
Calcul evaluation equivalente 
Credit entry credit exhibition outcom years // evidence 
Experiemental 3years equivenlent ÷
Comparability psychometric : 
Calculate time table ÷
Id calcule: calculate NQF level 360 credit .
Qcto group  calcul evidence group qualifications: 

6years ÷2years =3 years equivenlente 
Frame job years:
Framework handbook . quotatien intellectuel.
Credit entry 360:÷credit 180÷=2 credit awardc,percentage evaluation 50÷100=2.  Level 
Studies engineering.
-Work where appropriate a applying lubrificT correct assembly ,to assembly in accordance with specific,standard operate:

-where appropriate applying packing and or sealing material in accordance with specifications operate procedure :

-inspecting and checking the final assembly for conformance to specification.1th,2,th

-where apprit returning the final assembly to use :1th,2th,3th,4th, 

-diagnose and repairs analogie equipment and components date sign .look for evidence confirm skill .:

-Diagnose and repair and skill :

-Obtaining and following relevant circuit diagrams Manuel specific schematic: maintenance ,:

-locating reading recording and diagnostic buil in fault .:

- obtaining error code interpretation documents running test functionnning and recording fault and equipmentb built test .:

-Checking electronics equipment sub assemblies:

-Component connection and termination for conformance to specific .

- removing and replacing component :

-recording results of test understaking on electronics equipment:

:Isolating electronics assembly from the power .:
-adjusti g turning and calibrating electronic equipment sub assembly.:
Returning to service and testing to specification the repaired electronic equipment sub assembly..
- using language and literacy skills to provide brief report record result of test,
Reproduction fault symptom and verify fault using appropriate test.
-Retiring repairs,maintain service,look for evidence skill : 
-Look for evidence relevant circuit diagrams manual mainteecird supplier.:
- maintenance error code interpretation documents running test funct and fault and equipment status indicated built in test checking electronics component .:
- interpret technical drawing look for evidence skill in checking the drawing again job requirements related ,procedure  interpreting job chart ,docuy checking and clarifu task relate :
Undertaking numerical operation geometry  date sign
-checking materuaj exist:
- making termination ,: connection to specific manufacture and regulatory , adjustable marking tagging and calling wire conductor and connection to specification ,:
-connection  using language and literacy skill to complete and routine information written job instruction .

-Using measure for checking connection and components binstall and test electrical wiring .

Code trade : job specification pertaining system operating and relevant plant personal with respect identify deflecting control fault deflecting in the control for correct  loop corect operate: 
- relevant pneumatic  electronic circuit.

Diagram .testing monitoring recording resnse control system .using appropriate fault diagnosis technical procedure 
_____________________________________
Methode resolve equation , 
no phass mesureing and measuring evaluation,, 1th,2th,3th,4th
Fault find low Kirchoff:  find currentlooki g evidence circuit diagram labal ..

 calculation loop sum .. equation:
E1-E2=I1×R1+I2R2+R3I3  
E2-E3=I1×R1+I2R2+R3.i3 fund knowledge value 
.
_ phase step lreleass unity competency package 


	
	
Cpd: qualifications 
electrotechnology 
-mounte and wire control equipment 
Package performance,evidence requirements ,applying labelling and numbering to cables and using termi number in accordance indystryev, occupation health and safety ohs and wirksb deal unplanned procedure,selectbswutch heard and control wiring schematic ,,
_______________________________________
-undetstand electrical wiring labeling doblabek code non metal  conduct body partiebmean insulated conductubgb
Indicate size gaugeb,voltage rating  wire size,GB,wGbmateriaiv

Qualifications: criteria score description tools : excellent termination:
Learner outcom testing completed: 
14-2g  : two insulator ground 14,, maximum 600volub,
- underground cable feeder inside wall  if burial in ground,
- install conduit bplastic insulation,T:thermoplastic ,h eater ,w,x,nylon, synthetic,c ,rules doorbligthung ,buried PVC ,low voltage inch ,
-maintance  repare planned measure instrument components ,
Diagnostic and repaired documents

-,procedure component approprieb system. director.
Manufacture test review and approval before report .
Responsibility: originator is responsive written document .
Obtaining a DC number , priority to DRC routing 
- material requirements are identify :
Description skill knowledge look for evidence confirm skills yes or not: 
- checking nthe drawing agoing job requirements in accordance yes it note operating ,where appropriate procedure charter list and other applications yes or not ,refere documents,check clarity test ,operation geometry and calculation formulae,object represented in the drawing,unit of measure bin presentation drawing ,action to understand in response ,material from which the objet made hazard  
-:base assembly drawing us identify work relationship..
-relatuoj contains.
-installatuon drawing: provide compagnj position, exposure diagram .
-Schematic assembly drawing: pictures..machine drawing 
_____________________________________
Manufacture and inspection contractor  final inspection nteat quality plan document b,docuy record and information of inspection testing evidence confusioning acceptance . checking product used quality random check listed pulling nxhexjing verifying brequireb,shipper
_____________________________________
Apply where appropriate apply job research hiring compagny piece by piece variable including b, understand quality defect non conforming borior next assembly determination producti
Electronic 
 




	
	

.cpd: qualifications integrity body
.- undertake material ,cable laalling
Conductivity resistivity labaled 
-Thermo Cooper insulation :
Correct formulae task:
R=resistivity ×lenght ÷acros section.diameter length gauge.
Module young , dielectric test 
Material elasticity plastic
Matter 
GAZ :Liquide ,solide ,gaz

 Q=.permeability  insulator.
-Chemical  PVC polymer vynic coutch ,ch.molding  job appropriet.
Q=i.×u.×t...test 
Q=m.×c .×t∆
 
	














































































































 

































 



















Document walett office 




















metadata
Inbox
 
tshingombe fiston 
	12:17 PM (0 minutes ago)
	
	
A













Ngineering
Database computic report document wallete  data training trainer record , office 
   ,, supplier enginnering work electrical undertake information 
Inbox
 
tshingombe fiston <tshingombefiston@gmail.com> 
	11:57 AM (17 minutes ago)
	
	
to tshigombekb, me, tshingombe, tahitaditshingombe, TSHINGOMBEKB 
 

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
 
 
Mail Delivery Subsystem		11:58 AM (17 minutes ago)

Address not found Your message wasn't delivered to tshigombekb@gmail.com because the address couldn't be found, or is unable to receive mail. LEARN MORE The res

 
Mail Delivery Subsystem		11:58 AM (17 minutes ago)

 
Address not found Your message wasn't delivered to tahitaditshingombe@gmail.com because the address couldn't be found, or is unable to receive mail. LEARN MORE 

 
tshingombe fiston <tshingombefiston@gmail.com> 
	11:58 AM (16 minutes ago)

	
	
to tshigombekb, tshingombe, tahitaditshingombe, TSHINGOMBEKB, me 
 

LB   B9!ÐÏ à¡±á> 
________________________________________
þÿ      
________________________________________
þÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿýÿÿÿþÿÿÿþÿÿÿþÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿRoot Entryÿÿÿÿÿÿÿÿ[1]ði*ÆÜÎ ž˜ªWJO 
________________________________________
@f 
________________________________________
[1]ÿÿÿÿÿÿÿÿÿÿÿÿ.o 
________________________________________
[1] 
________________________________________
ÿÿÿÿþÿÿÿi01 ÿÿÿÿÿÿÿÿÿÿÿÿPÏ(e‘ïÙPÏ(e‘ïÙþÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿ 
________________________________________
(Š
€
€ 
________________________________________
________________________________________
________________________________________
}Üsµ7Root Entryÿÿÿÿÿÿÿÿ[1]ði*ÆÜÎ ž˜ªWJO`Ø-e‘ïÙ € 
________________________________________
f 
________________________________________
[1]ÿÿÿÿÿÿÿÿÿÿÿÿbo 
________________________________________
[1] 
________________________________________
ÿÿÿÿþÿÿÿi01 ÿÿÿÿ  n`ôÎ ›Íª`ŽPÏ(e‘ïÙPÏ(e‘ïÙÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿ ýÿÿÿþÿÿÿþÿÿÿ      þÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿf 
________________________________________
[1]ÿÿÿÿÿÿÿÿÿÿÿÿõo 
________________________________________
[1] 
________________________________________
ÿÿÿÿ 
________________________________________
 [1]ÿÿÿÿÿÿÿÿÿÿÿÿ
n
[1]þÿÿÿ 
________________________________________
þÿÿÿ      
þÿÿÿþÿÿÿ
þÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿ 
________________________________________
(Š
€
€ 
________________________________________
________________________________________
________________________________________
}Üsµ7,o$Õþÿ 
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
 
...

[Message clipped]  View entire message
 
Mail Delivery Subsystem 
	11:58 AM (16 minutes ago)

	
	
to me 
 


Address not found 
Your message wasn't delivered to tshigombekb@gmail.com because the address couldn't be found, or is unable to receive mail. 
LEARN MORE 



The response from the remote server was:
550 5.1.1 The email account that you tried to reach does not exist. Please try double-checking the recipient's email address for typos or unnecessary spaces. Learn more at https://support.google.com/mail/?p=NoSuchUser z20-20020a05651c023400b002bff9e2388bsi3433874ljn.512 - gsmtp 
 
 
Mail Delivery Subsystem		11:58 AM (16 minutes ago)

Address not found Your message wasn't delivered to tahitaditshingombe@gmail.com because the address couldn't be found, or is unable to receive mail. LEARN MORE 

 
tshingombe fiston <tshingombefiston@gmail.com> 
	12:01 PM (13 minutes ago)

	
	
to tshigombekb, tshingombe, tahitaditshingombe, TSHINGOMBEKB, me 
 

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
________________________________________
þÿ      [1]þÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿýÿÿÿþÿÿÿþÿÿÿþÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿRoot Entryÿÿÿÿÿÿÿÿ[1]ði*ÆÜÎ ž˜ªWJOàrß”ïÙ 
________________________________________
Àf 
________________________________________
[1]ÿÿÿÿÿÿÿÿÿÿÿÿ.o 
________________________________________
[1] 
________________________________________
ÿÿÿÿþÿÿÿCompObj [1]ÿÿÿÿÿÿÿÿÿÿÿÿnþÿÿÿ[1]þÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿ 
________________________________________
 ‚

€[1]} tf7þÿ 
 
 
Mail Delivery Subsystem		12:01 PM (13 minutes ago)

Address not found Your message wasn't delivered to tahitaditshingombe@gmail.com because the address couldn't be found, or is unable to receive mail. LEARN MORE 

 
Mail Delivery Subsystem		12:01 PM (13 minutes ago)

Address not found Your message wasn't delivered to tshigombekb@gmail.com because the address couldn't be found, or is unable to receive mail. LEARN MORE The res

 
tshingombe fiston <tshingombefiston@gmail.com> 
	12:02 PM (13 minutes ago)

	
	
to tshigombekb, tshingombe, tahitaditshingombe, TSHINGOMBEKB, me 
 

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
þÿ      [1]þÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿýÿÿÿþÿÿÿþÿÿÿþÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿRoot Entryÿÿÿÿÿÿÿÿ[1]ði*ÆÜÎ ž˜ªWJO 5£Ò”ïÙ 
________________________________________
Àf 
________________________________________
[1]ÿÿÿÿÿÿÿÿÿÿÿÿ.o 
________________________________________
[1] 
________________________________________
ÿÿÿÿþÿÿÿCompObj [1]ÿÿÿÿÿÿÿÿÿÿÿÿnþÿÿÿ[1]þÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿ 
________________________________________
 ‚

€[1] 
________________________________________
}>mÊ*þÿ 
________________________________________

ÿÿÿÿði*ÆÜÎ ž˜ªWJOMicrosoft Forms 2.0 FormEmbedded Object
Forms.Form.1ô9²q
 

On Mon, Sep 25, 2023 at 12:01 PM tshingombe fiston <tshingombefiston@gmail.com> wrote:
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
________________________________________
þÿ      [1]þÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿýÿÿÿþÿÿÿþÿÿÿþÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿRoot Entryÿÿÿÿÿÿÿÿ[1]ði*ÆÜÎ ž˜ªWJOàrß”ïÙ 
________________________________________
Àf 
________________________________________
[1]ÿÿÿÿÿÿÿÿÿÿÿÿ.o 
________________________________________
[1] 
________________________________________
ÿÿÿÿþÿÿÿCompObj [1]ÿÿÿÿÿÿÿÿÿÿÿÿnþÿÿÿ[1]þÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿ 
________________________________________
 ‚

€[1]} tf7þÿ 
________________________________________

ÿÿÿÿði*ÆÜÎ ž˜ªWJOMicrosoft Forms 2.0 FormEmbedded Object
Forms.Form.1ô9²q
 

On Mon, Sep 25, 2023 at 11:58 AM tshingombe fiston <tshingombefiston@gmail.com> wrote:
LB   B9!ÐÏ à¡±á> 
________________________________________
þÿ      
________________________________________
þÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿýÿÿÿþÿÿÿþÿÿÿþÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿRoot Entryÿÿÿÿÿÿÿÿ[1]ði*ÆÜÎ ž˜ªWJO 
________________________________________
@f 
________________________________________
[1]ÿÿÿÿÿÿÿÿÿÿÿÿ.o 
________________________________________
[1] 
________________________________________
ÿÿÿÿþÿÿÿi01 ÿÿÿÿÿÿÿÿÿÿÿÿPÏ(e‘ïÙPÏ(e‘ïÙþÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿ 
________________________________________
(Š
€
€ 
________________________________________
________________________________________
________________________________________
}Üsµ7Root Entryÿÿÿÿÿÿÿÿ[1]ði*ÆÜÎ ž˜ªWJO`Ø-e‘ïÙ € 
________________________________________
f 
________________________________________
[1]ÿÿÿÿÿÿÿÿÿÿÿÿbo 
________________________________________
[1] 
________________________________________
ÿÿÿÿþÿÿÿi01 ÿÿÿÿ  n`ôÎ ›Íª`ŽPÏ(e‘ïÙPÏ(e‘ïÙÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿ ýÿÿÿþÿÿÿþÿÿÿ      þÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿf 
________________________________________
[1]ÿÿÿÿÿÿÿÿÿÿÿÿõo 
________________________________________
[1] 
________________________________________
ÿÿÿÿ 
________________________________________
àCompObj [1]ÿÿÿÿÿÿÿÿÿÿÿÿpCompObj [1]ÿÿÿÿÿÿÿÿÿÿÿÿ
n
[1]þÿÿÿ 
________________________________________
þÿÿÿ      
þÿÿÿþÿÿÿ
þÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿ 
________________________________________
(Š
€
€ 
________________________________________
________________________________________
________________________________________
}Üsµ7,o$Õþÿ 
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
 
...

[Message clipped]  View entire message
 
Mail Delivery Subsystem		12:02 PM (12 minutes ago)

Address not found Your message wasn't delivered to tahitaditshingombe@gmail.com because the address couldn't be found, or is unable to receive mail. LEARN MORE 

 
Mail Delivery Subsystem		12:02 PM (12 minutes ago)

Address not found Your message wasn't delivered to tshigombekb@gmail.com because the address couldn't be found, or is unable to receive mail. LEARN MORE The res

 
tshingombe fiston <tshingombefiston@gmail.com> 
	12:03 PM (11 minutes ago)

	
	
to tshigombekb, tshingombe, tahitaditshingombe, TSHINGOMBEKB, me 
 

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
________________________________________
þÿ      [1]þÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿýÿÿÿþÿÿÿþÿÿÿ 
________________________________________
þÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿRoot Entryÿÿÿÿÿÿÿÿ[1]ði*ÆÜÎ ž˜ªWJOP÷ î”ïÙ 
________________________________________
À 
________________________________________
f 
________________________________________
[1]ÿÿÿÿÿÿÿÿÿÿÿÿjo 
________________________________________
[1] 
________________________________________
ÿÿÿÿ CompObj [1]ÿÿÿÿÿÿÿÿÿÿÿÿ
n [1] 
________________________________________
________________________________________
þÿÿÿ     


þÿÿÿ
þÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿ 
________________________________________
  

      
}e_1<‡t(õ €[1](€name W)      [1]5 €¥[1]Tahoma[1](AE€H€, 
________________________________________
[1]
€I‹"text name"[1]5 €¥[1]Tahoma[1] (
€name :     °'Ê[1]5 €¥[1]Tahoma[1]A€H€, 
________________________________________
[1]uÊ[1]5 €¥[1]Tahoma[1]( €surname &q[1]5 €¥[1]Tahomae [1]A€H€, 
________________________________________
[1]Îö 
________________________________________
[1]5 €¥[1]Tahomae [1] (    €birthday   Ü&Ê[1]u €¥[1] 
________________________________________
Tahomae 28 Label1e g
ö 
________________________________________
(å    €[1]HComboBox1²4 Îö 
________________________________________
(õ € 
________________________________________
2<[1] Label2e g

________________________________________
(å    € 8 
________________________________________
ComboBox2²4 Î 
________________________________________
(õ €28 
________________________________________
Label3e ;
u(å  € 8ComboBox3²4 ¢I,å
€   < CommandButton1M
;
Ü&þÿ 
 
 
Mail Delivery Subsystem		12:04 PM (11 minutes ago)

Address not found Your message wasn't delivered to tshigombekb@gmail.com because the address couldn't be found, or is unable to receive mail. LEARN MORE 550 5.1

 
Mail Delivery Subsystem		12:04 PM (11 minutes ago)

Address not found Your message wasn't delivered to tahitaditshingombe@gmail.com because the address couldn't be found, or is unable to receive mail. LEARN MORE 

 
tshingombe fiston <tshingombefiston@gmail.com> 
	12:09 PM (5 minutes ago)

	
	
to tshigombekb, tshingombe, tahitaditshingombe, TSHINGOMBEKB, me 
 

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
________________________________________
þÿ      
________________________________________
________________________________________
þÿÿÿ     

ÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿýÿÿÿþÿÿÿþÿÿÿþÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿRoot Entryÿÿÿÿÿÿÿÿ[1]ði*ÆÜÎ ž˜ªWJO 
________________________________________
Àf 
________________________________________
[1]ÿÿÿÿÿÿÿÿÿÿÿÿ.o 
________________________________________
[1] 
________________________________________
ÿÿÿÿTi03 ÿÿÿÿÿÿÿÿÿÿÿÿÀ³¼O–ïÙÀ³¼O–ïÙþÿÿÿ[1]þÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿ 
________________________________________
(

‚
€      
________________________________________

}jo.[1]4(!€portofolio programming office pc Zgž [1]5 €¥[1]TahomaRoot Entryÿÿÿÿÿÿÿÿ[1]ði*ÆÜÎ ž˜ªWJO €ÂO–ïÙ À 
________________________________________
f 
________________________________________
[1]ÿÿÿÿÿÿÿÿÿÿÿÿŽo 
________________________________________
[1] 
________________________________________
ÿÿÿÿTi03 ÿÿÿÿ  n`ôÎ ›Íª`ŽÀ³¼O–ïÙà½O–ïÙÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿ ýÿÿÿþÿÿÿþÿÿÿ‰ýÿÿÿýÿÿÿýÿÿÿ

 ¬ !"#$%&'()*+,-./0123456789:;<=>?@ABCDEFGHIJKLMNOPQRSTUVWXYZ[\]^_`abcdefghijklmnopqrstuvwxyz{|}~ €f 
________________________________________
[1]ÿÿÿÿÿÿÿÿÿÿÿÿyo 
________________________________________
[1] 
________________________________________
ÿÿÿÿ
¨ù[1]CompObj [1]ÿÿÿÿÿÿÿÿÿÿÿÿ 
________________________________________
pCompObj [1]ÿÿÿÿÿÿÿÿÿÿÿÿ
n
[1]þÿÿÿ 
________________________________________
þÿÿÿ      
þÿÿÿ
þÿÿÿ
þÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿ 
________________________________________
(

‚
€      
________________________________________

}jo.[1]X‚o(õ[1]4(!€portofolio programming office pc Zgž [1]5 €¥[1]Tahomaþÿ 
________________________________________

ÿÿÿÿ  n`ôÎ ›Íª`ŽMicrosoft Forms 2.0 FrameEmbedded Object
Forms.Frame.1ô9²q 
________________________________________
8H

      
________________________________________
€ 
________________________________________
€ÿÿ }Zgð Frame1 
________________________________________
Rã
‘ Î ãªK¸Q DB Tahoma †t,å
€ 
________________________________________
< CommandButton1AO 
________________________________________
O 
________________________________________
(å €Øø[1]ComboBox1ttoú ‚ƒ„…†‡ˆ‰Š‹Œ Ž ‘’“”•–—˜™š›œ žŸ ¡¢£¤¥¦§¨©ª«¬-®¯°±²³´µ¶•¸¹º»¼½¾¿ÀÁÂÃÄÅÆÇÈÉÊËÌÍÎÏÐÑÒÓÔÕÖ×ØÙÚÛÜÝÞßàáâãäåæçèéêëìíîïðñòóôõö÷øùúûüýþÿ[1] 
________________________________________
________________________________________
      




 ¬ !"#$%&'()*+,-./0123456789:;<=>?@ABCDEFGHIJKLMNOPQRSTUVWXYZ[\]^_`abcdefghijklmnopqrstuvwxyz{|}~ € ‚ƒ„…†‡ˆþÿÿÿþÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿ[1] (
€portofolio0 ö 
________________________________________
[1]u €¥[1] 
________________________________________
Tahoma[1] AE€H€, 
________________________________________
[1]œø[1]ïG4 

































































 

VB6 Made Easy Book

 

Check Out Our Book

________________________________________

VB6 Made Easy Paperback

<br />

VB6 Made Easy Kindle

VB6 Google Book

 

Available on Google Play Store

________________________________________



Lesson 2 : Building VB Applications

________________________________________

2.1 Creating Your First Application 

First of all, launch Microsoft Visual Basic 6 compiler that you have installed earlier. In the New Project Dialog , choose Standard EXE to enter Visual Basic 6 integrated development environment. In the VB6 IDE, a default form with the name Form1 will appear. Next, double click on Form1 to bring up the source code window for Form1, as shown in Figure 2.1. 

The top of the source code window consists of a list of objects and their associated events or procedures. In the source code window, the object displayed is Form1 and the associated procedure is Load. 


 Figure 2.1 The VB6 Source Code Window 

When you click on the object box, the drop-down list will display a list of objects you have inserted into your form, as shown in figure 2.2. Here, you can see a form with the name Form1, a command button with the name Command1, a Label with the name Label1 and a Picture Box with the name Picture1. 

 Figure 2.2: List of Objects Similarly, when you click on the procedure box, a list of procedures associated with the object will be displayed , as shown in Figure 2.3. Some of the procedures associated with the object Form1 are Activate, Click, DblClick (which means Double-Click) , DragDrop, keyPress and more. Each object has its own set of procedures. You can always select an object and write codes for any of its procedure in order to perform certain tasks.

 Figure 2.3 List of Procedures 

You do not have to worry about the beginning and the end statements (i.e. Private Sub Form_Load.......End Sub.); Just key in the lines in between the above two statements exactly as are shown here. When you press F5 to run the program, you will be surprised that nothing showed up .In order to display the output of the program, you have to add the Form1.show statement like in Example 2.1.1  or you can just use Form_Activate ( )  event procedure as shown in example 2.1.2. The command Print does not mean printing using a printer but it means displaying the output on the computer screen. Now, press F5 or click on the run button to run the program and you will get the output as shown in Figure 2.4.

You can also perform arithmetic calculations as shown in Example 2.1.2. VB uses * to denote the multiplication operator and / to denote the division operator. The output is shown in Figure 2.5, where the results are arranged vertically.

Example 2.1.1

Private Sub Form_Load ( )

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

End Sub


The Output of Example 2.1.4(a) &(b) is as shown in Figure 2.7.

 Figure 2.7 

2.2 Steps in Building a Visual Basic Application

Step 1: Design the interface by adding controls to the form and set their properties

Step 2:Write code for the event procedures

Example 2.2 Changing Background and Foreground Color at Random

In this example, we want to show you how to write code to change the background and the foreground color randomly. We will place two command buttons and a label on the form. One of the command buttons will be used to change the background color while the other one will be used to change the foreground color. The Label is for displaying the foreground color. There are two events here, change background color and change foreground color. Therefore, we need to write code for the two event procedures.

To make the program more interesting, we will use the Rnd() function, the Int() function and the RGB codes to change the color randomly. The Rnd() function creates a random number between 0 and 1 and the RGB code uses a combination of three integers to form a certain color. The Int() is a function that converts a number into an integer by truncating its decimal part and the resulting integer is the largest integer that is smaller than the number. For example, Int(0.2)=0, Int(2.4)=2, Int(4.8)=4. Therefore, Int(Rnd()*256) returns the smallest integer 0 and the biggest integer 255. The format of RGB code is RGB(a,b,c), where a, b, c range from 0 to 255. For example, RGB(255,0,0) is red, RGB(255,255,255) is white and (0,0,0) is black. Do not worry about the jargons, you will learn them in later lesson.

Now, rename the controls as follows:

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

Program Examples

Inputting & Outputting Data

If you want to display something on screen in Visual Basic you can do it in a number of ways. The most common ways you will use are by Message Box or List Box.

To input data in Visual Basic there are lots of different form controls that you can use to do this, the most common ones are text boxes and combo boxes.

Inputting - Creating an Interface

Visual Basic have two main parts to it. The first part is creating the user interface, this is a simply drag and drop environment where you can add different form controls such as a text box.

The second part to it is creating the code to make the user interface do/display something.

When creating an interface, there are different types of form controls. Some of those that you will use are:

"    Label - used to display text on an interface

"    Text Box - used to allow the user to type in some text

"    Combo Box - used to give a set of options to the user

"    List Box - used to provide a list of options to choose from or display a list of information that has been calculated in the program.

"    Button - these are often used so when the user clicks on them something happens.

In Visual Studio on the left hand side you have the toolbox, this is where you can drag and drop your form controls. It looks like this:

 

When you have dragged your form controls onto your form, you need to name them. You should add a prefix to any form control, the common ones are:

"    Labels - lbl

"    Text Boxes - txt

"    Combo Boxes - cmb

"    List Boxes - lst

"    Buttons - btn

If you had a text box where the user enters their name a suitable name would be txtName.

Difference between naming a control and changing the text You must always name controls, this can be done in the properties next to the option name, as shown below:

 

If you want to change what the form control says on it, you change the option text, as shown below:

 

Outputting Data - Message Boxes

Below is an interface where the user will enter a word. When they enter the word in the text box called txtWord and press the button called btnRun it will display a message box that says what word they entered.

Interface  

Code when btnRun clicked

MessageBox.Show("You entered the word " & txtWord.Text)

This is what happens when the button is clicked:

 

The text You entered the word  is joined with the text box txtWord. As the user entered the word Hello it takes what it is in the text box and adds it into the message box.

Example program 1 - Birthday Program

The code for the program below will allow the user to enter their name, select the day of the week that their birthday falls on this year and then choose the month of their birthday from the list box. When they click the button it should display all the information back to them.

Interface

 

Code when btnBirthday clicked

MessageBox.Show("Hello " & txtName.Text & vbNewLine & "Your birthday month is " & lstMonth.Text &

" and the day of the birthday this year is " & cmbDay.Text)

This is what happens when the button is clicked:

 

Example program 2 - Address Program

The code for the program below will allow the user to enter various pieces of information. It will then use the information in these form controls to create a message box with all their information in.

Interface

 

Code when btnDisplay clicked

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

The code for the program below will allow the user to enter a name and three test scores. When btnAdd is clicked it will add the information to the list box.

Interface

 

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
...

[Message clipped]  View entire message
 
Mail Delivery Subsystem 
	12:10 PM (5 minutes ago)

	
	
to me 
 


Address not found 
Your message wasn't delivered to tshigombekb@gmail.com because the address couldn't be found, or is unable to receive mail. 
LEARN MORE 



The response from the remote server was:
550 5.1.1 The email account that you tried to reach does not exist. Please try double-checking the recipient's email address for typos or unnecessary spaces. Learn more at https://support.google.com/mail/?p=NoSuchUser w21-20020a05651c103500b002b6e3d38d98si3580873ljm.123 - gsmtp 
 
 
Mail Delivery Subsystem 
	12:10 PM (5 minutes ago)

	
	
to me 
 


Address not found 
Your m




metadatabse dhet
Inbox
 
tshingombe fiston <tshingombefiston@gmail.com> 
	12:27 PM (5 minutes ago)
	
	
to tshigombekb, me, tshingombe, tahitaditshingombe, TSHINGOMBEKB 
 

VERSION 5.00
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
 
End Sub
 
Mail Delivery Subsystem		12:27 PM (4 minutes ago)

Address not found Your message wasn't delivered to tshigombekb@gmail.com because the address couldn't be found, or is unable to receive mail. LEARN MORE The res

 
Mail Delivery Subsystem		12:27 PM (4 minutes ago)

Address not found Your message wasn't delivered to tahitaditshingombe@gmail.com because the address couldn't be found, or is unable to receive mail. LEARN MORE 

 
tshingombe fiston <tshingombefiston@gmail.com> 
	12:28 PM (4 minutes ago)

	
	
to tshigombekb, tshingombe, tahitaditshingombe, TSHINGOMBEKB, me 
 

LB º6' ÐÏ à¡±á> 
________________________________________
þÿ      
________________________________________
þÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿ €# 
________________________________________

Frame1ÊO 
________________________________________
ÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿþÿ 
________________________________________

ÿÿÿÿði*ÆÜÎ ž˜ªWJOMicrosoft Forms 2.0 FormEmbedded Object
Forms.Form.1ô9²qÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿRoot Entryÿÿÿÿÿÿÿÿ[1]ði*ÆÜÎ ž˜ªWJO 
________________________________________
@f 
________________________________________
[1]ÿÿÿÿÿÿÿÿÿÿÿÿ.o 
________________________________________
[1] 
________________________________________
ÿÿÿÿþÿÿÿi01 ÿÿÿÿÿÿÿÿÿÿÿÿPƒ¥SOôÙPƒ¥SOôÙþÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿ 
________________________________________
(


________________________________________
€ [1] 
________________________________________
}é_k"Root Entryÿÿÿÿÿÿÿÿ[1]ði*ÆÜÎ ž˜ªWJOp$¬SOôÙ À f 
________________________________________
[1]ÿÿÿÿÿÿÿÿÿÿÿÿbo 
________________________________________
[1] 
________________________________________
ÿÿÿÿþÿÿÿi01 ÿÿÿÿ
  n`ôÎ ›Íª`ŽPƒ¥SOôÙPƒ¥SOôÙþÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿ ýÿÿÿ     þÿÿÿ
þÿÿÿ
ÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿf 
________________________________________
[1]ÿÿÿÿÿÿÿÿÿÿÿÿ½o 
________________________________________
[1] 
________________________________________
ÿÿÿÿ 
________________________________________
Xi03 ÿÿÿÿ   n`ôÎ ›Íª`ŽPƒ¥SOôÙPƒ¥SOôÙCompObj [1]ÿÿÿÿÿÿÿÿÿÿÿÿp[1]þÿÿÿ 
________________________________________
þÿÿÿ þÿÿÿ      þÿÿÿ



þÿÿÿ þÿÿÿþÿÿÿþÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿ 
________________________________________
(


________________________________________
€ [1] 
________________________________________
}é_k",e$Õþÿ 
________________________________________

ÿÿÿÿ  n`ôÎ ›Íª`ŽMicrosoft Forms 2.0 FrameEmbedded Object
Forms.Frame.1ô9²q[1]8(&€portofolio evidence low judgement  poe`Oö 
________________________________________
[1]5 €¥[1]Tahoma 
________________________________________
8H


________________________________________
€ 
________________________________________
€ÿÿ}VT Frame1 
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
''txt label 2 award diploma cerificate ''s[1]5 €¥[1]Tahoma[1]$l€
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
CommandButton1\ì      “
...

[Message clipped]  View entire message
 
Mail Delivery Subsystem		12:28 PM (3 minutes ago)

Address not found Your message wasn't delivered to tshigombekb@gmail.com because the address couldn't be found, or is unable to receive mail. LEARN MORE The res

 
Mail Delivery Subsystem		12:28 PM (3 minutes ago)

Address not found Your message wasn't delivered to tahitaditshingombe@gmail.com because the address couldn't be found, or is unable to receive mail. LEARN MORE 

 
tshingombe fiston <tshingombefiston@gmail.com> 
	12:29 PM (3 minutes ago)

	
	
to tshigombekb, tshingombe, tahitaditshingombe, TSHINGOMBEKB, me 
 

VERSION 5.00
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
 
End Sub
 
Private Sub UserForm_Zoom(Percent As Integer)
 
End Sub
 
 
 
 
Mail Delivery Subsystem		12:29 PM (3 minutes ago)

Address not found Your message wasn't delivered to tahitaditshingombe@gmail.com because the address couldn't be found, or is unable to receive mail. LEARN MORE 

 
Mail Delivery Subsystem		12:29 PM (3 minutes ago)

Address not found Your message wasn't delivered to tshigombekb@gmail.com because the address couldn't be found, or is unable to receive mail. LEARN MORE The res

 
tshingombe fiston <tshingombefiston@gmail.com> 
	12:30 PM (2 minutes ago)

	
	
to tshigombekb, tshingombe, tahitaditshingombe, TSHINGOMBEKB, me 
 

LB ò  B,¬ÐÏ à¡±á> 
________________________________________
þÿ      
________________________________________
þÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿCompObj [1]ÿÿÿÿÿÿÿÿÿÿÿÿ,nÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿRoot Entryÿÿÿÿÿÿÿÿ[1]ði*ÆÜÎ ž˜ªWJO 
________________________________________
@f 
________________________________________
[1]ÿÿÿÿÿÿÿÿÿÿÿÿ&o 
________________________________________
[1] 
________________________________________
ÿÿÿÿþÿÿÿi01 ÿÿÿÿÿÿÿÿÿÿÿÿ dwNôÙ dwNôÙþÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿ 
________________________________________
  

}Üs4Root Entryÿÿÿÿÿÿÿÿ[1]ði*ÆÜÎ ž˜ªWJOpø{NôÙ €
f 
________________________________________
[1]ÿÿÿÿÿÿÿÿÿÿÿÿZo 
________________________________________
[1] 
________________________________________
ÿÿÿÿþÿÿÿi01 ÿÿÿÿ  n`ôÎ ›Íª`Ž dwNôÙ dwNôÙþÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿ ýÿÿÿ  þÿÿÿpr



 ¬ !"#$%&'()*+,-./0123456789:;<=>?@ABCDEFGHIJKLMNOPQRSTUVWXYZ[\]^_`abcdefghijklmnoþÿÿÿqtsuvwþÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿf 
________________________________________
[1]ÿÿÿÿÿÿÿÿÿÿÿÿ 
________________________________________
5 
________________________________________
o 
________________________________________
[1] 
________________________________________
ÿÿÿÿ
<Êi16 ÿÿÿÿ     n`ôÎ ›Íª`Ž dwNôÙ dwNôÙCompObj [1]ÿÿÿÿÿÿÿÿÿÿÿÿp+[1]þÿÿÿ 
________________________________________
      




þÿÿÿ þÿÿÿ þÿÿÿþÿÿÿþÿÿÿ þÿÿÿ !þÿÿÿ#$þÿÿÿ&þÿÿÿþÿÿÿ)þÿÿÿþÿÿÿþÿÿÿ-þÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿ 
________________________________________
  

}Üs4,be$Õ €þÿ 
________________________________________

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

Create Login(Sign In) and Registration (Sign Up) Form in C# Windows Form With Database


 


 


YogeshKumar Hadiya

" 

Follow

Published in


C# Programming


" 

5 min read

" 

Sep 6, 2020



Introduction

Hello guys, in this article we will create a login and registration form with database in c# windows form application .In this article we create a windows form app with login page, registration page for create new account and home page which show message that you login successfully .I hope you will like this article .So let s start step by step.

 

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

 

Step: 11:- Now click anywhere on form it will generate Form_Load event where enter following code. This code create database connection and open it. In next step you will learn how you get that connection string which added in SQLConnection Constructor.

Step: 12:- Go to server Explorer right click on database, click on Modify Connection.

 

Step: 13:- Now you see a windows dialog popup click on advance button. This will open another dialog. But Before that click on test button and check you database working properly.

 

Step: 14:- Copy path which show below on this dialog and close both dialog. Then past this path in form load event. Add @ sign before this path so you no need to change slash.

 

Step: 15:- We need to open login page when user click on login button so enter following code in Login Button click event.

Code Explanation:

1. Here first we hide the current form which is registration .

2. Then we create an object of login page and show login form using that object.

Step: 16:- Now add following code in registration button click event.

Code Explanation:

1. First of all we check that user enter value in all field if yes that continue otherwise show message using message box.

2. Then we check if password and confirm password both are same.

3. Then we check if any record/user is already register with that username if not then continue further otherwise show error message.

4. In last we insert data in table using SQLCommand object.

Step: 17:- Now we create a login page here I add two textbox for username and password and two button for login and open registration form.

 

Step: 18:- Click on anywhere in form which generate Form_Load event add connection code that as show below.

Step: 19:- On Registration button click add following code which open registration form.

Step: 20:- Add below code in login button click for redirect user to home page form if user exist.

Code Explanation

1. Here first of all we check if user enter value in both field if yes then continue otherwise show error message.

2. Then we check if user exist in our database with that username and password. If user exist then open home page which we generate in start.

Step: 21:-Change start page as login in Program.cs File.

Step: 22:-Now run your application.

 

Conclusion

So, here we create a simple login and registration page in windows form application. I hope you like this article, share with your friends.

Source : YogeshHadiya.In


Csharp



Csharp Programming



C Sharp Windows App



Sql



 


 



Written by YogeshKumar Hadiya


83 Followers

" Writer for 


C# Programming


I am a software developer and designer work on .Net, .Net core , Angular and Ionic frameworks. Developer | Blogger | C# Corner MVP

More from YogeshKumar Hadiya and C# Programming


 



 



YogeshKumar Hadiya


in


C# Programming



How To Post Data To The Controller Using AJAX With Validations In ASP.NET Core

Introduction



7 min read" Sep 7, 2022



 



 



Abnoan Muniz


in


C# Programming



Top 10 Certifications for .NET Developers in 2023

The Credentials That Matter: Certification Insights for .NET Developers





" 15 min read" Aug 19



1



 



 



Sukhpinder Singh


in


C# Programming



Convert Swagger documentation to Postman Collection

This article step by step process how to convert swagger endpoint into an postman collection.





" 2 min read" Aug 6, 2020



1



 



 



YogeshKumar Hadiya


in


C# Programming



Map One Object Into Another Without Any Package In C#

Introduction



7 min read" Mar 28



1



See all from YogeshKumar Hadiya



See all from C# Programming


Recommended from Medium


 



 



Nile Bits



C# Keywords Tutorial Part 43: implicit

C# presents an array of keywords that assist in accomplishing different programming tasks. Among these is the implicit  keyword, which& 



2 min read" Apr 12



 



 



Victor Magalhães



Test Coverage Analysis with Coverlet in .NET

Test coverage is a metric that quantifies the proportion of a program s source code exercised by a set of test cases. It is typically& 



12 min read" Sep 23


Lists


 

 

 

New_Reading_List

174 stories" 127 saves



 

 

 

ChatGPT

21 stories" 174 saves



 

 

 

Natural Language Processing

658 stories" 265 saves



 



 



Riya Sharma



Create Mock APIs Using JSON Server for Seamless Development

If you ve ever wondered how to effortlessly create custom APIs for your projects, even before the backend is fully cooked, you re in for a& 



3 min read" Sep 6



 



 



Abiodun Maborukoje



Exploring SOLID Principles in .NET Core Applications

=Ø€ÞNew to .NET Core development? Let s delve into SOLID principles to enhance your coding skills:



3 min read" 6 days ago



 



 



Pasan Manohara



C# Generics

Can you see a problem here? What is the type of this list? Is it going to throw a runtime error?



4 min read" Sep 4



 



 



Vinod Pal



How to connect .NET Web API with SQL Server using Entity Framework: Code-First Approach

In the previous article, we explored how to connect a .NET Web API with SQL Server using Entity Framework in the Database-First approach& 



4 min read" Jul 30



See more recommendations



Help



Status



Writers



Blog



Careers



Privacy



Terms



About



Text to speech



Teams


Login Form in Windows Application

In our last c#.net post we learnt about how to open new windows form by Show() and ShowDialog() mothods. In this c# windows application tutorial we will learn more about Show() and ShowDialog() in detail with an example of login form.

Here, We will open new form by login with valid username and password on login page. We design a login form with two textbox controls for username and password and a button control for write login code. We have taken id txtuname for username textbox and textpass for password textbox and btnlogin for login button control.

 Simple login form in windows application. 

In above figure we took the login form design with two textbox controls along with button control. We took button control for write login code in click event.

private void btnlogin_Click(object sender, EventArgs e)

{

if (txtuname.Text == "")

{

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

In above example if textbox will be blank then it shows the message like Enter Username . If both username and password fill up and one of will be incorrect then it will shows the message Invalid Credential . If username = Meera  and Password = 123  then the main form will be hidden and open Form2 using ShowDialog() method.

If both username and password are correct then we login into system by opening form2.

 Simple login form in windows application. 

We hope that this c#.net tutorial helped you to understand about create simple login form without database in windows application.

________________________________________

Next, c# windows application tutorial we will learn how to pass data from one from to another form.

________________________________________

 

In this Post, we will learn how to create a Simple Windows form Login application.


Let s Begin:

1. Create a New Windows Form Application.

 

2. Add New Database (I have created a database named as MyDatabase.mdf). Add a table (named as tbl_Login). The following is the table schema for creating tbl_Login.

 

3. Create a form (frmLogin) and add Label, TextBox and button control from the Toolbox.

 

4. Add another Windows Form and named it as frmMain. This form will be shown to the user after successful Login by the user.

 

Now, Go to frmLogin.cs code and add System.Data and System.Data.SqlClient namespace. Double click on btn_Submit to create btn_Submit Click event.

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
...

[Message clipped]  View entire message
 
Mail Delivery Subsystem		12:30 PM (2 minutes ago)

Address not found Your message wasn't delivered to tshigombekb@gmail.com because the address couldn't be found, or is unable to receive mail. LEARN MORE The res

 
Mail Delivery Subsystem		12:30 PM (2 minutes ago)

Address not found Your message wasn't delivered to tahitaditshingombe@gmail.com because the address couldn't be found, or is unable to receive mail. LEARN MORE 

 
tshingombe fiston <tshingombefiston@gmail.com> 
	12:31 PM (1 minute ago)

	
	
to tshigombekb, tshingombe, tahitaditshingombe, TSHINGOMBEKB, me 
 

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
 
End Sub
 
 
Mail Delivery Subsystem		12:31 PM (1 minute ago)

Address not found Your message wasn't delivered to tahitaditshingombe@gmail.com because the address couldn't be found, or is unable to receive mail. LEARN MORE 

 
Mail Delivery Subsystem		12:31 PM (1 minute ago)

Address not found Your message wasn't delivered to tshigombekb@gmail.com because the address couldn't be found, or is unable to receive mail. LEARN MORE The res

 
tshingombe fiston <tshingombefiston@gmail.com> 
	12:31 PM (0 minutes ago)

	
	
to tshigombekb, tshingombe, tahitaditshingombe, TSHINGOMBEKB, me 
 

LB 
4¡ ÐÏ à¡±á> 
________________________________________
þÿ      [1]þÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿýÿÿÿþÿÿÿþÿÿÿþÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿRoot Entryÿÿÿÿÿÿÿÿ[1]ði*ÆÜÎ ž˜ªWJO°îÖ=OôÙ 
________________________________________
€f 
________________________________________
[1]ÿÿÿÿÿÿÿÿÿÿÿÿfo 
________________________________________
[1] 
________________________________________
ÿÿÿÿXCompObj [1]ÿÿÿÿÿÿÿÿÿÿÿÿ 
________________________________________
n 
________________________________________
[1]þÿÿÿþÿÿÿþÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿ 
________________________________________
(


________________________________________


€[1] }C[»¬0ho(õ[1]8((€successful .congratulation registration ùBÊ[1]5 €¥[1]Tahoma €2X Label1;
" 
________________________________________
þÿ 
________________________________________

ÿÿÿÿði*ÆÜÎ ž˜ªWJOMicrosoft Forms 2.0 FormEmbedded Object
Forms.Form.1ô9²q
 
 
Mail Delivery Subsystem 
	12:32 PM (0 minutes ago)

	
	
to me 
 


Address not found 
Your message wasn't delivered to tshigombekb@gmail.com because the address couldn't be found, or is unable to receive mail. 
LEARN MORE 



The response was:
550 5.1.1 The email account that you tried to reach does not exist. Please try double-checking the recipient's email address for typos or unnecessary spaces. Learn more at https://support.google.com/mail/?p=NoSuchUser c28-20020a19655c000000b004fe87788c55sor884145lfj.21 - gsmtp 
 
 
Mail Delivery Subsystem 
	12:32 PM (0 minutes ago)

	
	
to me 
 


Address not found 
Your message wasn't delivered to tahitaditshingombe@gmail.com because the address couldn't be found, or is unable to receive mail. 
LEARN MORE 



The response was:
550 5.1.1 The email account that you tried to reach does not exist. Please try double-checking the recipient's email address for typos or unnecessary spaces. Learn more at https://support.google.com/mail/?p=NoSuchUser c28-20020a19655c000000b0



 
tshingombe fiston 
	1:12 PM (0 minutes ago)
	
	
to tshingombe, tshigombekb, me 
 

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
________________________________________
þÿ       
________________________________________
þÿÿÿ     



ÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿýÿÿÿþÿÿÿþÿÿÿþÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿRoot Entryÿÿÿÿÿÿÿÿ[1]ði*ÆÜÎž˜ªWJO 
________________________________________
@f 
________________________________________
[1]ÿÿÿÿÿÿÿÿÿÿÿÿ.o 
________________________________________
[1] 
________________________________________
ÿÿÿÿþÿÿÿi01ÿÿÿÿÿÿÿÿÿÿÿÿPw¬öÙPw-öÙþÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿ 
________________________________________
(
€     [1]      }kiŽHRoot Entryÿÿÿÿÿÿÿÿ[1]ði*ÆÜÎž˜ªWJOàØ^w¬öÙ€ 
________________________________________
f 
________________________________________
[1]ÿÿÿÿÿÿÿÿÿÿÿÿbo 
________________________________________
[1] 
________________________________________
ÿÿÿÿþÿÿÿi01ÿÿÿÿ  n`ôÎ›Íª`ŽPw¬öÙPDQw-öÙÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿýÿÿÿþÿÿÿþÿÿÿñ[1]ýÿÿÿýÿÿÿýÿÿÿýÿÿÿýÿÿÿ ¬ !"#$%&'()*+,-./0123456789:;<=>?@ABCDEFGHIJKLMNOPQRSTUVWXYZ[\]^_`abcdefghijklmnopqrstuvwxyz{|}~€f 
________________________________________
[1]ÿÿÿÿÿÿÿÿÿÿÿÿ 
________________________________________
Éo 
________________________________________
[1] 
________________________________________
ÿÿÿÿ
ˆÄCompObj[1]ÿÿÿÿÿÿÿÿÿÿÿÿpCompObj[1]ÿÿÿÿÿÿÿÿÿÿÿÿ
n
[1]þÿÿÿ 
________________________________________
     
þÿÿÿþÿÿÿ
þÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿ 
________________________________________
(
€     [1]     }kiŽH,xt$Õþÿ 
________________________________________

ÿÿÿÿ  n`ôÎ›Íª`ŽMicrosoft Forms 2.0 FrameEmbedded Object
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
ComboBox2'g
‚ƒ„…†‡ˆ‰Š‹ŒŽ‘’“”•–—˜™š›œžŸ ¡¢£¤¥¦§¨©ª«¬-®¯°±²³´µ¶•¸¹º»¼½¾¿ÀÁÂÃÄÅÆÇÈÉÊËÌÍÎÏÐÑÒÓÔÕÖ×ØÙÚÛÜÝÞßàáâãäåæçèéêëìíîïðñòóôõö÷øùúûüýþÿ[1] 
________________________________________
________________________________________
      




 ¬ !"#$%&'()*+,-./0123456789:;<=>?@ABCDEFGHIJKLMNOPQRSTUVWXYZ[\]^_`abcdefghijklmnopqrstuvwxyz{|}~€‚ƒ„…†‡ˆ‰Š‹ŒŽ‘’“”•–—˜™š›œžŸ ¡¢£¤¥¦§¨©ª«¬-®¯°±²³´µ¶•¸¹º»¼½¾¿ÀÁÂÃÄÅÆÇÈÉÊËÌÍÎÏÐÑÒÓÔÕÖ×ØÙÚÛÜÝÞßàáâãäåæçèéêëìíîïðñòóôõö÷øùúûüýþÿ[1][1][1][1] 
________________________________________
[1] 
________________________________________
[1][1][1][1][1]      [1]
[1]
[1]
[1]
[1]
[1][1][1][1][1][1][1][1][1][1][1][1][1][1][1][1] [1]¬[1] [1]![1]"[1]#[1]$[1]%[1]&[1]'[1]([1])[1]*[1]+[1],[1]-[1].[1]/[1]0[1]1[1]2[1]3[1]4[1]5[1]6[1]7[1]8[1]9[1]:[1];[1]<[1]=[1]>[1]?[1]@[1]A[1]B[1]C[1]D[1]E[1]F[1]G[1]H[1]I[1]J[1]K[1]L[1]M[1]N[1]O[1]P[1]Q[1]R[1]S[1]T[1]U[1]V[1]W[1]X[1]Y[1]Z[1][[1]\[1]][1]^[1]_[1]`[1]a[1]b[1]c[1]d[1]e[1]f[1]g[1]h[1]i[1]j[1]k[1]l[1]m[1]n[1]o[1]p[1]q[1]r[1]s[1]t[1]u[1]v[1]w[1]x[1]y[1]z[1]{[1]|[1]}[1]~[1][1]€[1][1]‚[1]ƒ[1]„[1]…[1]†[1]‡[1]ˆ[1]‰[1]Š[1]‹[1]Œ[1][1]Ž[1][1][1]‘[1]’[1]“[1]”[1]•[1]–[1]—[1]˜[1]™[1]š[1]›[1]œ[1][1]ž[1]Ÿ[1] [1]¡[1]¢[1]£[1]¤[1]¥[1]¦[1]§[1]¨[1]©[1]ª[1]«[1]¬[1]-[1]®[1]¯[1]°[1]±[1]²[1]³[1]´[1]µ[1]¶[1]•[1]¸[1]¹[1]º[1]»[1]¼[1]½[1]¾[1]¿[1]À[1]Á[1]Â[1]Ã[1]Ä[1]Å[1]Æ[1]Ç[1]È[1]É[1]Ê[1]Ë[1]Ì[1]Í[1]Î[1]Ï[1]Ð[1]Ñ[1]Ò[1]Ó[1]Ô[1]Õ[1]Ö[1]×[1]Ø[1]Ù[1]Ú[1]Û[1]Ü[1]Ý[1]Þ[1]ß[1]à[1]á[1]â[1]ã[1]ä[1]å[1]æ[1]ç[1]è[1]é[1]ê[1]ë[1]ì[1]í[1]î[1]ï[1]ð[1]þÿÿÿþÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿ[1]0( €police data base  syst man poe[<ž[1]5€¥[1]Tahoma[1]4("€police sudent data base managemnt [<[1]5€¥[1]Tahoma[1]€H€,[<Ü 
________________________________________
[1]5€¥[1]Tahoma[1]A€H€, 
________________________________________
[1]Ü&Ê[1]5€¥[1]Tahoma[1]A€H€, 
________________________________________
[1]Ü&Ê[1]5€¥[1]Tahoma[1]A€H€, 
________________________________________
[1]µ:‹[1]5€¥[1]Tahoma[1] (
€actiuve system&É[1]u€¥[1] 
________________________________________
Tahoma[1]AE€H€, 
________________________________________
[1]ŒÂ’6Sub fist()

'

' fist Macro

'

'


End Sub

Enable desktop notifications for Gmail.   OK  No thanks

25 of 862

Program police / meta

Inbox

tshingombe fiston <tshingombefiston@gmail.com>

    

Oct 2, 2023, 11:09?AM (2 days ago)

    

to tshigombekb, me, tshingombe, TSHINGOMBEKB

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

A. Just as the title implies, security managers and system administrators are most often considered to serve in a management capacity.  The important tasks of developing security regulations, training staff, and monitoring implementation require that the security manager be vested with substantial authority.  While the security manager is not to be confused with a superintendent or principal, he or she should be considered to be the system "boss."  If the security manager is not able to confidently address security miscues at even the highest levels of the organizational hierarchy, protecting system resources adequately becomes an impossibility.

<span style="fon
 	


 
tshingombe fiston 
	1:17 PM (0 minutes ago)
	
	
to tshingombe, tshigombekb, me 
 

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
________________________________________
þÿ      
________________________________________
þÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿýÿÿÿþÿÿÿþÿÿÿþÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿRoot Entryÿÿÿÿÿÿÿÿ[1]ði*ÆÜÎ ž˜ªWJO 
________________________________________
@f 
________________________________________
[1]ÿÿÿÿÿÿÿÿÿÿÿÿ.o 
________________________________________
[1] 
________________________________________
ÿÿÿÿþÿÿÿi01 ÿÿÿÿÿÿÿÿÿÿÿÿàÂÆ²öÙàÂÆ²öÙþÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿ 
________________________________________
(Š
€        [1]
}×`Å<Root Entryÿÿÿÿÿÿÿÿ[1]ði*ÆÜÎ ž˜ªWJOà÷'Æ²öÙ @ 
________________________________________
f 
________________________________________
[1]ÿÿÿÿÿÿÿÿÿÿÿÿbo 
________________________________________
[1] 
________________________________________
ÿÿÿÿþÿÿÿi01 ÿÿÿÿ  n`ôÎ ›Íª`ŽàÂÆ²öÙ Æ²öÙÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿ ýÿÿÿþÿÿÿþÿÿÿ¬




 þÿÿÿþÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿf 
________________________________________
[1]ÿÿÿÿÿÿÿÿÿÿÿÿ 
________________________________________
¡o 
________________________________________
[1] 
________________________________________
ÿÿÿÿ `+CompObj [1]ÿÿÿÿÿÿÿÿÿÿÿÿpCompObj [1]ÿÿÿÿÿÿÿÿÿÿÿÿ
n
[1]þÿÿÿ 
________________________________________
     þÿÿÿþÿÿÿ
þÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿ 
________________________________________
(Š
€        [1]
}×`Å<,be$Õþÿ 
________________________________________

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
L.M.Photonics Ltd
      
                   
      
 
 
 
 
 
 
 
 
 
 
 
 
In Association with Amazon.com
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
            
 
 
Power Factor correction, Motor starters and Variable frequency drives for induction motors.
 
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
 
                  
 
 
 
Questions and/or comments on this page
 
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

ÿÿÿÿði*ÆÜÎ ž˜ªWJOMicrosoft Forms 2.0 FormEmbedded Object
Forms.Form.1ô9²q
 
End Sub
...

  


Enable desktop notifications for Gmail.   OK  No thanks
1 of 916

metabase vb
Inbox
 
tshingombe fiston <tshingombefiston@gmail.com> 
	2:10 PM (8 minutes ago)
	
	
to me, tshingombe, tshigombekb 
 

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
 
End Sub
 
Private Sub TextBox3_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
 
End Sub
 
Private Sub UserForm_Click()
 
End Sub
 
tshingombe fiston 
	2:11 PM (7 minutes ago)

	
	
to tshingombe, tshigombekb, me 
 

LB VZ< )ÐÏ à¡±á> 
________________________________________
þÿ      [1] 
________________________________________
þÿÿÿ      ÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿýÿÿÿþÿÿÿþÿÿÿþÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿRoot Entryÿÿÿÿÿÿÿÿ[1]ði*ÆÜÎ ž˜ªWJO 
________________________________________
@f 
________________________________________
[1]ÿÿÿÿÿÿÿÿÿÿÿÿ2o 
________________________________________
[1] 
________________________________________
ÿÿÿÿÜi02 ÿÿÿÿÿÿÿÿÿÿÿÿÐ Pü ùÙÐ Pü ùÙþÿÿÿ[1] 
________________________________________
________________________________________
þÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿ 
________________________________________
,Š
€     [1]
€}ÕiŽE[1]¼,€¦€dhet //st peace college , scale weigthing  portofolio   value                                                 analyse design investigate  psychometrical    memorendum-IX[1]5 €¥[1]TahomaRoot Entryÿÿÿÿÿÿÿÿ[1]ði*ÆÜÎ ž˜ªWJOðöcü ùÙ @f 
________________________________________
[1]ÿÿÿÿÿÿÿÿÿÿÿÿ’o 
________________________________________
[1] 
________________________________________
ÿÿÿÿÜi02 ÿÿÿÿ  n`ôÎ ›Íª`ŽÐ Pü ùÙÐ Pü ùÙÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿ ýÿÿÿþÿÿÿþÿÿÿ§ýÿÿÿ



 ¬ !"#$%&'()*+,-./0123456789:;<=>?@ABCDEFGHIJKLMNOPQRSTUVWXYZ[\]^_`abcdefghijklmnopqrstuvwxyz{|}~ €f 
________________________________________
[1]ÿÿÿÿÿÿÿÿÿÿÿÿi 
________________________________________
o 
________________________________________
[1] 
________________________________________
ÿÿÿÿ
8CompObj [1]ÿÿÿÿÿÿÿÿÿÿÿÿpCompObj [1]ÿÿÿÿÿÿÿÿÿÿÿÿn[1] 
________________________________________
________________________________________
þÿÿÿ þÿÿÿ 




þÿÿÿþÿÿÿ þÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿ 
________________________________________
,Š
€      [1]
€}ÕiŽE[1]X‚t[1]¼,€¦€dhet //st peace college , scale weigthing  portofolio   value                                                 analyse design investigate  psychometrical    memorendum-IX[1]5 €¥[1]Tahomaþÿ 
________________________________________

ÿÿÿÿ  n`ôÎ ›Íª`ŽMicrosoft Forms 2.0 FrameEmbedded Object
Forms.Frame.1ô9²q 
________________________________________
8H


________________________________________
€ 
________________________________________
€ÿÿ}%B+*Frame1 
________________________________________
Rã
‚ƒ„…†‡ˆ‰Š‹Œ Ž ‘’“”•–—˜™š›œ žŸ ¡¢£¤¥¦þÿÿÿ¨©þÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿ[1](@€H€,
€ 
________________________________________
>„ 
________________________________________
"calculator" d[1]5 €¥[1]Tahoma[1] (€1¾v ö 
________________________________________
[1]u €¥[1] 
________________________________________
Tahoma][1] (€2ho # 
________________________________________
[1]u €¥[1] 
________________________________________
Tahoma][1] (€3ho ö 
________________________________________
[1]u €¥[1] 
________________________________________
Tahoma][1] (€4ho # 
________________________________________
[1]u €¥[1] 
________________________________________
Tahoma][1] (€5hoÊö 
________________________________________
[1]u €¥[1] 
________________________________________
Tahoma][1] (€6hoÊö 
________________________________________
[1]u €¥[1] 
________________________________________
Tahoma][1] (€7hoÊö 
________________________________________
[1]u €¥[1] 
________________________________________
Tahoma][1] (€8ho " 
________________________________________
[1]u €¥[1] 
________________________________________
Tahoma][1] (€9hoö 
________________________________________
" 
________________________________________
[1]u €¥[1] 
________________________________________
Tahoma][1] (€0hoÊ" 
________________________________________
[1]u €¥[1] 
________________________________________
Tahoma][1](€equal " 
________________________________________
[1]u €¥[1] 
________________________________________
Tahoma[1]( €cancelö 
________________________________________
ö 
________________________________________
[1]u €¥[1] 
________________________________________
Tahoma[1] (€+hoö 
________________________________________
# 
________________________________________
[1]u €¥[1] 
________________________________________
Tahoma][1] (€-hoö 
________________________________________
" 
________________________________________
[1]u €¥[1] 
________________________________________
Tahoma][1] (€/hoö 
________________________________________
" 
________________________________________
[1]u €¥[1] 
________________________________________
Tahoma][1] (€*hoö 
________________________________________
N 
________________________________________
[1]u €¥[1] 
________________________________________
Tahoma][1] €H€,À
N 
________________________________________
[1]5 €¥[1]Tahoma[1] €H€,À
N 
________________________________________
[1]5 €¥[1]Tahoma[1] AE€H€, 
________________________________________
[1]ˆ3ú“
Public Class Form1

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

 

        // create a textfield

        l = new JTextField(16);

 

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

        b7 = new JButton("7");

        b8 = new 
...

[Message clipped]  View entire message
 
tshingombe fiston 
	2:13 PM (5 minutes ago)

	
	
to tshingombe, tshigombekb, me 
 

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
þÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿýÿÿÿþÿÿÿþÿÿÿþÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿRoot Entryÿÿÿÿÿÿÿÿ[1]ði*ÆÜÎ ž˜ªWJO 
________________________________________
@f 
________________________________________
[1]ÿÿÿÿÿÿÿÿÿÿÿÿ2o 
________________________________________
[1] 
________________________________________
ÿÿÿÿþÿÿÿi01 ÿÿÿÿÿÿÿÿÿÿÿÿÈÇ¹      ùÙÈÇ¹      ùÙþÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿ 
________________________________________
,

________________________________________
€
€!      [1] 
________________________________________
"}ch‡>Root Entryÿÿÿÿÿÿÿÿ[1]ði*ÆÜÎ ž˜ªWJOà Ø¹      ùÙ @f 
________________________________________
[1]ÿÿÿÿÿÿÿÿÿÿÿÿfo 
________________________________________
[1] 
________________________________________
ÿÿÿÿþÿÿÿi01 ÿÿÿÿ  n`ôÎ ›Íª`ŽÈÇ¹  ùÙ€âÍ¹     ùÙÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿ ýÿÿÿþÿÿÿþÿÿÿd




 ¬ !"#$%&'()*+,-./0123456789:;<=>?@ABCDEFGHIJKLMNOPQRSTUVWXYZ[\]^_`abcþÿÿÿefþÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿf 
________________________________________
[1]ÿÿÿÿÿÿÿÿÿÿÿÿ 
________________________________________
±o 
________________________________________
[1] 
________________________________________
ÿÿÿÿ      ìµCompObj [1]ÿÿÿÿÿÿÿÿÿÿÿÿpCompObj [1]ÿÿÿÿÿÿÿÿÿÿÿÿn[1]þÿÿÿ 
________________________________________





þÿÿÿþÿÿÿ þÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿ 
________________________________________
,

________________________________________
€
€!   [1] 
________________________________________
"}ch‡>,xtþÿ 
________________________________________

ÿÿÿÿ  n`ôÎ ›Íª`ŽMicrosoft Forms 2.0 FrameEmbedded Object
Forms.Frame.1ô9²q 
________________________________________
8H

! 
________________________________________
€ 
________________________________________
€ÿÿ!}ôZ98Frame1 
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

 

// Create ATM GUI Java Frame

JFrame f2;


 JLabel l1,l2,l3,l4,l5,l6,user,l7,l8; // declared label to display text

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

52.         
...

[Message clipped]  View entire message
 
tshingombe fiston 
	2:14 PM (4 minutes ago)

	
	
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

= False
LB ¸8`'ÐÏ à¡±á> 
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
ÿÿÿÿþÿÿÿi01 ÿÿÿÿÿÿÿÿÿÿÿÿð–¶
ùÙð–¶
ùÙþÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿ 
________________________________________
,
________________________________________
€
   [1] 
________________________________________

}mc BRoot Entryÿÿÿÿÿÿÿÿ[1]ði*ÆÜÎ ž˜ªWJOPE¶
ùÙ f 
________________________________________
[1]ÿÿÿÿÿÿÿÿÿÿÿÿfo 
________________________________________
[1] 
________________________________________
ÿÿÿÿþÿÿÿi01 ÿÿÿÿ  n`ôÎ ›Íª`Žð–¶
ùÙð–¶
ùÙÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿ ýÿÿÿþÿÿÿþÿÿÿ 
þÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿf 
________________________________________
[1]ÿÿÿÿÿÿÿÿÿÿÿÿ
[1]o 
________________________________________
[1] 
________________________________________
ÿÿÿÿ 
________________________________________
8[1]CompObj [1]ÿÿÿÿÿÿÿÿÿÿÿÿpCompObj [1]ÿÿÿÿÿÿÿÿÿÿÿÿn [1]þÿÿÿ 
________________________________________
      

þÿÿÿ

þÿÿÿþÿÿÿþÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿ 
________________________________________
,
________________________________________
€
   [1] 
________________________________________

}mc B,xtþÿ 
________________________________________

ÿÿÿÿ  n`ôÎ ›Íª`ŽMicrosoft Forms 2.0 FrameEmbedded Object
Forms.Frame.1ô9²q[1]((€atm managegemnt systemUª?ö 
________________________________________
[1]5 €¥[1]Tahoma[1] ( 
________________________________________
€nameÃHö 
________________________________________
[1]5 €¥[1]Tahoma[1] €H€,ÃHö 
________________________________________
[1]5 €¥[1]Tahoma[1] (€atm card number–Iö 
________________________________________
[1]5 €¥[1]Tahoma[1] €H€,ÃH# 
________________________________________
[1]5 €¥[1]Tahoma[1](€atm pin] # 
________________________________________
[1]5 €¥[1]Tahoman[1] €H€,S# 
________________________________________
[1]5 €¥[1]Tahoman[1] (      €submittednum˜ ö 
________________________________________
[1]u €¥[1] 
________________________________________
Tahoman[1]( €balance — " 
________________________________________
[1]5 €¥[1]Tahomae [1] €H€,k¬" 
________________________________________
[1]5 €¥[1]Tahomae 
________________________________________
8H


________________________________________
€ 
________________________________________
€ÿÿ
}¯Rô/Frame1 
________________________________________
Rã
‘ Î ãªK¸Q DB Tahoma
°Š(õ €[1]2H Label1e q§(õ € 
________________________________________
24 Label2e ö 
________________________________________
$å € 
________________________________________
4[1]TextBox1" 
________________________________________


(õ €2@ 
________________________________________
Label3x1Ê 
________________________________________
$å € 4 
________________________________________
TextBox2ö 
________________________________________
'(õ €28 Label4x2" 
________________________________________
ð $å € 4 TextBox3" 
________________________________________
æ!,å
€   < CommandButton1rÎW)(õ €
28 Label5x3Ž#Ä$å €
4      TextBox4Ž#Ž#$Õ €# 
________________________________________

Frame1] Êþÿ 
________________________________________

ÿÿÿÿði*ÆÜÎ ž˜ªWJOMicrosoft Forms 2.0 FormEmbedded Object
Forms.Form.1ô9²q
...

[Message clipped]  View entire message
 
tshingombe fiston 
	2:15 PM (3 minutes ago)

	
	
to tshingombe, tshigombekb, me 
 

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
________________________________________
þÿ      
________________________________________
þÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿýÿÿÿþÿÿÿþÿÿÿþÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿRoot Entryÿÿÿÿÿÿÿÿ[1]ði*ÆÜÎ ž˜ªWJO 
________________________________________
@f 
________________________________________
[1]ÿÿÿÿÿÿÿÿÿÿÿÿ.o 
________________________________________
[1] 
________________________________________
ÿÿÿÿþÿÿÿi01 ÿÿÿÿÿÿÿÿÿÿÿÿ@¦™0
ùÙ@¦™0
ùÙþÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿ 
________________________________________
(



€
[1] 
________________________________________
}ébè@Root Entryÿÿÿÿÿÿÿÿ[1]ði*ÆÜÎ ž˜ªWJOÀõ«0
ùÙ 
________________________________________
f 
________________________________________
[1]ÿÿÿÿÿÿÿÿÿÿÿÿbo 
________________________________________
[1] 
________________________________________
ÿÿÿÿþÿÿÿi01 ÿÿÿÿ  n`ôÎ ›Íª`Ž@¦™0
ùÙà,›0
ùÙÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿ ýÿÿÿþÿÿÿþÿÿÿ'




 ¬ !"#$%&þÿÿÿþÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿf 
________________________________________
[1]ÿÿÿÿÿÿÿÿÿÿÿÿ 
________________________________________
m[1]o 
________________________________________
[1] 
________________________________________
ÿÿÿÿ ;CompObj [1]ÿÿÿÿÿÿÿÿÿÿÿÿpCompObj [1]ÿÿÿÿÿÿÿÿÿÿÿÿ
n
[1]þÿÿÿ 
________________________________________
     


þÿÿÿþÿÿÿþÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿ 
________________________________________
(



€
[1] 
________________________________________
}ébè@,be$Õþÿ 
________________________________________

ÿÿÿÿ  n`ôÎ ›Íª`ŽMicrosoft Forms 2.0 FrameEmbedded Object
Forms.Frame.1ô9²q 
________________________________________
8H


________________________________________
€ 
________________________________________
€ÿÿ
}*UD3Frame1 
________________________________________
Rã
‘ Î ãªK¸Q DB Tahoma
[1]Œ(õ €[1]2@ Label1 q(õ € 
________________________________________
2D Label2 â
$å € 
________________________________________
4[1]TextBox1'q$å €4 
________________________________________
TextBox2ú¶(õ € 2D 
________________________________________
Label3x2 ¢$[1] (€id number meter×>Ê[1]5 €¥[1]Tahoma[1]$( €clien meter name X×> [1]5 €¥[1]Tahoma[1] €H€,M.Ê[1]5 €¥[1]Tahoma[1] €H€,z- [1]5 €¥[1]Tahoma[1]$( €meter regster kwh X«?É[1]5 €¥[1]Tahoma[1] €H€,¦,É[1]5 €¥[1]Tahoma[1] (
€CommandButton1r:
ö 
________________________________________
[1]u €¥[1] 
________________________________________
Tahoma[1]( €Label4×>É[1]5 €¥[1]Tahoma[1] €H€,ÿ*# 
________________________________________
[1]5 €¥[1]Tahoma[1]t8AE€H€, 
________________________________________
[1]V8g
µJava Program to Calculate Electricity Bill

Write a Java Program to Calculate Electricity Bill using Else If statement with examples.

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
€ @ CommandButton1rÀ
ÿ*(õ €      28 Label4x3 !$å €
4 TextBox4úæ!(å €
”8     ComboBox1ttoGO 
________________________________________
(å
€

/ScrollBar1toïG] (õ €
2d
Label5x4 Ô €# 
________________________________________

Frame1q þÿ 
________________________________________

ÿÿÿÿði*ÆÜÎ ž˜ªWJOMicrosoft Forms 2.0 FormEmbedded Object
Forms.Form.1ô9²q
End Sub

On Sat, Oct 7, 2023 at 2:10 PM tshingombe fiston <tshingombefiston@gmail.com> wrote:
...

[Message clipped]  View entire message
 
tshingombe fiston 
	2:16 PM (2 minutes ago)

	
	
 
to tshingombe, tshigombekb, me 
 

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
________________________________________
þÿ      
________________________________________
þÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿýÿÿÿþÿÿÿþÿÿÿþÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿRoot Entryÿÿÿÿÿÿÿÿ[1]ði*ÆÜÎ ž˜ªWJO 
________________________________________
@f 
________________________________________
[1]ÿÿÿÿÿÿÿÿÿÿÿÿ.o 
________________________________________
[1] 
________________________________________
ÿÿÿÿþÿÿÿi01 ÿÿÿÿÿÿÿÿÿÿÿÿä) 
ùÙä) 
ùÙþÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿ 
________________________________________
(Š


€ [1] 
________________________________________
}rN)3Root Entryÿÿÿÿÿÿÿÿ[1]ði*ÆÜÎ ž˜ªWJO@&9 
ùÙ 
________________________________________
f 
________________________________________
[1]ÿÿÿÿÿÿÿÿÿÿÿÿbo 
________________________________________
[1] 
________________________________________
ÿÿÿÿþÿÿÿi01 ÿÿÿÿ  n`ôÎ ›Íª`Žä) 
ùÙä) 
ùÙÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿ ýÿÿÿþÿÿÿþÿÿÿ<




 ¬ !"#$%&'()*+,-./0123456789:;þÿÿÿþÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿf 
________________________________________
[1]ÿÿÿÿÿÿÿÿÿÿÿÿ 
________________________________________
Eo 
________________________________________
[1] 
________________________________________
ÿÿÿÿ LdCompObj [1]ÿÿÿÿÿÿÿÿÿÿÿÿpCompObj [1]ÿÿÿÿÿÿÿÿÿÿÿÿ
n [1]þÿÿÿ 
________________________________________
þÿÿÿþÿÿÿ
þÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿ 
________________________________________
(Š


€ [1] 
________________________________________
}rN)3,xt$Õþÿ 
________________________________________

ÿÿÿÿ  n`ôÎ ›Íª`ŽMicrosoft Forms 2.0 FrameEmbedded Object
Forms.Frame.1ô9²q 
________________________________________
8H


________________________________________
€ 
________________________________________
€ÿÿ}–Iƒ(Frame1 
________________________________________
Rã
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
 
End Sub
 
The following example enables the "Send PDF" option, sets the email subject, sets the email address, sets the email body text, and then creates the PDF file:
 
 
 'enable sending the PDF file 
 'replace the file options with &H421 to email with no user interaction using the Mail Helper application
 SaveWin2PDFDword "file options", &H21
 
 'set the email subject (optional)
 SaveSetting "Dane Prairie Systems", "Win2PDF", "PDFMailSubject", "Test Subject"
 
 'set the recipient email address (optional)
 SaveSetting "Dane Prairie Systems", "Win2PDF", "PDFMailRecipients", billg@microsoft.com
 
 'set the email body text (optional)
 SaveSetting "Dane Prairie Systems", "Win2PDF", "PDFMailNote", "The requested PDF file is attached."
 
 'set the PDF file name
 SaveSetting "Dane Prairie Systems", "Win2PDF", "PDFFileName", "c:\\test.pdf"
 
 'output some text
 Printer.Print "hello world!"
 Printer.Print www.win2pdf.com
 
 'save the PDF file
 Printer.EndDoc
 
Upgrading the Printer Object
Last Updated on Sat, 04 Feb 2023 | Upgrading Visual Basic 
In Visual Basic 6.0, the Printer object enables applications to communicate with a system printer. The Printer object contains functions for printing text, lines, and images. However, the printing model in the .NET Framework is quite different from that in Visual Basic 6.0. The Printer object has become obsolete. Fortunately, the functionality of the Printer object can be achieved with the classes that are provided in the System.Drawing.Printing namespace, particularly the PrintDocument class.
There are two main options to upgrade the Visual Basic 6.0 printing functionality. The first option is to replace the Printer object members with equivalent functionality provided by the PrintDocument class. Typically, this requires reimplementation of the functionality; this will be demonstrated in the next code example. The second option is to create your own Printer class in Visual Basic .NET based on the .NET PrintDocument class. This approach will allow you to mask the .NET PrintDocument class functionality with the names provided in the Visual Basic 6.0 Printer object. This second option will also be demonstrated later in this section.
Applying the first approach will require reimplementation of the printing code. The following Visual Basic 6.0 code example prints four lines of text with different font settings:
Private Sub Form_Load() With Printer
Printer.Print "Normal Line" .FontBold = True Printer.Print "Bold Line" .FontItalic = True
Printer.Print "Bold and Italic Line" .FontBold = False .FontItalic = False Printer.Print "Second Normal Line" .EndDoc End With End Sub
This code can be reimplemented in Visual Basic .NET using methods of the PrintDocument class, as shown here.
Dim WithEvents prn As New Printing.PrintDocument
Private Sub Form1_Load(ByVal eventSender As System.Object, _ ByVal eventArgs As System.EventArgs) Handles MyBase.Load prn.Print() End Sub
Private Sub prn_PrintPage(ByVal sender As Object, _
ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles prn.PrintPage Dim currFont As Font
Dim yPos As Integer yPos = 1
With e.Graphics currFont = New Font("Arial", 10, FontStyle.Regular) .DrawString("Normal Line", currFont, Brushes.Black, 1, yPos) yPos = yPos + currFont.GetHeight currFont = New Font("Arial", 10, FontStyle.Bold) .DrawString("Bold Line", currFont, Brushes.Black, 1, yPos) yPos = yPos + currFont.GetHeight currFont = New Font("Arial", 10, FontStyle.Bold Or FontStyle.Italic) .DrawString("Bold and Italic Line", currFont, Brushes.Black, 1, yPos) yPos = yPos + currFont.GetHeight currFont = New Font("Arial", 10, FontStyle.Regular) .DrawString("Second Normal Line", currFont, Brushes.Black, 1, yPos) End With End Sub
Notice how the structure of the code has changed and how the quantity of drawing methods has increased. The .NET Framework provides more control and power over the drawing operations, but from the upgrade point of view, it has a learning curve, and the manual reimplementation of all the printing functionality can consume extensive resources.
For more information about the equivalent functionality for the Printer object members, see Tables 7.3 and 7.4 at the end of this section. For further guidance on printing from Visual Basic .NET with the PrintDocument component, see "Printing with the PrintDocument Component" in on MSDN.
The second upgrade approach involves the creation of your own Printer class. This allows you to consolidate several drawing methods and collections of graphical objects (such as Circle or Line) and to store the coordinates that will be used to print these objects when the Print method or the EndDoc method is called. When the Print method of the PrintDocument class is called, the PrintPage event is raised. This event can be used to signal the application to draw all the objects and text stored in the collections of your PrinterClass. The following Visual Basic .NET code provides a small demonstration of how this Printer class can be developed.
Imports Microsoft.VisualBasic.Compatibility ' Imports Microsoft.VisualBasic. Public Class PrinterClass
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
Tables 7.3 and 7.4 list Visual Basic 6.0 Printer object properties and methods and their Visual Basic .NET equivalents. Where there are no direct equivalents, links are provided to additional information.
The following example shows you how to convert Visual Basic 6.0 code that uses a printer object to Visual Basic .NET code that uses your own printer class as mentioned earlier.
Printer.FillColor = vbBlue Printer.FillStyle = vbSolid Printer.Circle (3000, 3000), 1500, vbRed Printer.EndDoc
The preceding Visual Basic 6.0 code will print a circle filled with blue color and it will have a red border. To convert this code to Visual Basic .NET, you can use the printer class mentioned earlier and perform the following steps.
First, you have to add this printer class to the .NET project. After this, you have to add a reference to the Micosoft.VisualBasic.Compatibility assembly.
The next step is to create a Visual Basic .NET module and add the following code.
Module Module1
Public Printer As PrinterClass = New PrinterClass End Module
Finally, you need to perform some small changes to the original Visual Basic 6.0 code to make it consistent with the new definition of your own printer class. Here is the resulting Visual Basic .NET code.
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
FillStyle   This functionality can be obtained using a Drawing.SolidBrush to fill the Draws objects.
Font  This property could be simulated in your printer class with a System.Drawing.Font class.
FontBold    To get the value, use: System.Drawing.Font.Bold.To set the value, use: VB6.FontChangeBold.
FontCount   System.Drawing.FontFamily.GetFamilies().Length
FontItalic  To get the value, use: System.Drawing.Font.Italic.To set the value, use: VB6.FontChangeItalic.
FontName    To get the value, use: System.Drawing.Font.Name.To set the value, use: VB6.FontChangeName.
Fonts System.Drawing.FontFamily.GetFamilies
FontSize    To get the value, use: System.Drawing.Font.Size.To set the value, use: VB6.FontChangeSize.
FontStrikethru    To get the value, use: System.Drawing.Font.Strikeout.To set the value, use: VB6.FontChangeStrikeout.
continued continued
Visual Basic 6.0 Visual Basic .NET Equivalent
FontTransparent   No equivalent. For details, see "Font Changes in Visual Basic .NET" in Visual Basic Concepts on MSDN.
FontUnderline     To get the value, use: System.Drawing.Font.Underline.To set the value, use: VB6.FontChangeUnderline.
ForeColor   System.Drawing.Pen.Color
hDC   PrintDocument.PrinterSettings.GetHdevmode.ToInt32
Height      PrintDocument.DefaultPageSettings.PaperSize.Height
Orientation PrintDocument.DefaultPageSettings.Landscape
Page  No direct equivalent. The current page number is not tracked; however, you can easily do this by setting a variable in the BeginPrint event and incrementing it in the PrintPage event.
PaperBin    PrintDocument.PrinterSettings.PaperSources
PaperSize   PrintDocument.DefaultPageSettings.PaperSize
Port  No longer necessary. The PrintPreviewDialog control automatically sets port information.
PrintQuality      PrintDocument.DefaultPageSettings.PrinterResolution
RightToLeft No longer necessary. The direction of printing is controlled by the localization settings in Windows.
ScaleHeight No equivalent. For details, see "Coordinate System Changes in Visual Basic .NET" in Visual Basic Concepts on MSDN.
ScaleLeft   No equivalent. For details, see "Coordinate System Changes in Visual Basic .NET" on MSDN.
ScaleMode   No equivalent. For details, see "Coordinate System Changes in Visual Basic .NET" on MSDN.
ScaleTop    No equivalent. For details, see "Coordinate System Changes in Visual Basic .NET" on MSDN.
ScaleWidth  No equivalent. For details, see "Coordinate System Changes in Visual Basic .NET" on MSDN.
TrackDefault      No direct equivalent. The IsDefaultPrinter property of the PrinterSettings class can be used to determine whether a printer is the default, but printing is no longer halted if the default printer changes.
...

[Message clipped]  View entire message
 
tshingombe fiston 
	2:17 PM (1 minute ago)

 
	
	
to tshingombe, tshigombekb, me 
 

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
LB 36ð ÐÏ à¡±á> 
________________________________________
þÿ      
________________________________________
þÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿþÿ 
________________________________________

ÿÿÿÿði*ÆÜÎ ž˜ªWJOMicrosoft Forms 2.0 FormEmbedded Object
Forms.Form.1ô9²qÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿRoot Entryÿÿÿÿÿÿÿÿ[1]ði*ÆÜÎ ž˜ªWJO 
________________________________________
@f 
________________________________________
[1]ÿÿÿÿÿÿÿÿÿÿÿÿ2o 
________________________________________
[1] 
________________________________________
ÿÿÿÿþÿÿÿi01 ÿÿÿÿÿÿÿÿÿÿÿÿp
%1
ùÙp
%1
ùÙþÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿ 
________________________________________
,Š

€
[1]
€
}û^¬3Root Entryÿÿÿÿÿÿÿÿ[1]ði*ÆÜÎ ž˜ªWJOp@11
ùÙ € f 
________________________________________
[1]ÿÿÿÿÿÿÿÿÿÿÿÿfo 
________________________________________
[1] 
________________________________________
ÿÿÿÿþÿÿÿi01 ÿÿÿÿ  n`ôÎ ›Íª`Žp
%1
ùÙp
%1
ùÙþÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿ ýÿÿÿþÿÿÿþÿÿÿ 
ÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿf 
________________________________________
[1]ÿÿÿÿÿÿÿÿÿÿÿÿ
![1]o 
________________________________________
[1] 
________________________________________
ÿÿÿÿ 
________________________________________
ˆ[1]CompObj [1]ÿÿÿÿÿÿÿÿÿÿÿÿpCompObj [1]ÿÿÿÿÿÿÿÿÿÿÿÿn[1]þÿÿÿ 
________________________________________
      



þÿÿÿ þÿÿÿþÿÿÿþÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿ 
________________________________________
,Š

€
[1]
€
}û^¬3,xtþÿ 
________________________________________

ÿÿÿÿ  n`ôÎ ›Íª`ŽMicrosoft Forms 2.0 FrameEmbedded Object
Forms.Frame.1ô9²q[1] (
€plc ip adressAõ/q[1]5 €¥[1]Tahoma[1] (€plc reader dataÈ0E [1]5 €¥[1]Tahoma[1] ( 
________________________________________
€plcop2É[1]5 €¥[1]Tahoma[1] (€read data form 0 Ê[1]u €¥[1] 
________________________________________
Tahoma[1] (
€disconectoro:
Ê[1]u €¥[1] 
________________________________________
Tahoma[1] (
€plc connecterm “
Ê[1]u €¥[1] 
________________________________________
Tahoma[1]\(K€communication plc data form dhet st peace   engineering electrical data // ˆ;{[1][1]5 €¥[1]Tahoma[1] €H€,—  [1]5 €¥[1]Tahoma[1] €H€,˜ ö 
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

Frame1ì  Ê

On Sat, Oct 7, 2023 at 2:10 PM tshingombe fiston <tshingombefiston@gmail.com> wrote:
...

[Message clipped]  View entire message


1 of 927

META DATABASE LICENSE ISSUE , CERTIFICATE
Inbox
 
tshingombe fiston <tshingombefiston@gmail.com> 
	11:55 AM (2 minutes ago)
	
	
to tshingombe, me, tshigombekb, TSHINGOMBEKB 
 

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

 
tshingombe fiston 
	11:56 AM (1 minute ago)

	
	
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

[Message clipped]  View entire message


4 Attachments • Scanned by Gmail
 
tshingombe fiston 
	12:49 PM (0 minutes ago)
	
	
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

     þÿÿÿþÿÿÿ

þÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿ 

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

________________________________________

28[1] Label2§ì      $å ¬ xt 

________________________________________

TextBox1{[1]


(å ¬  ¬  

________________________________________

ComboBox2¬  §1  ’  &   ! Æ[1]0 `9 R }     "   Ü[1]"!a: S ~x ¡¢£¤¥¦§¨©ª«¬-®¯°±²³´µ

" ¸¹º»¼½¾¿ÀÁÂÃÄÅÆÇÈÉÊËÌÍÎÏÐÑÒÓÔÕÖ×ØÙÚÛÜÝÞßàáâãäåæçèéêëìíîïðñòóôõö÷øùúûüýþþÿÿÿþÿÿÿ[1]´(¢¬ dhet st peace college student markrsheet , assessment  policy circulum license  issue certificate   term   managemnt system information ,student learner  lecture r S" 

________________________________________

[1]5 ¬ ¥[1]Tahoma[1] AE¬ H¬ , 

________________________________________

[1]ø *U   

________________________________________

      �
















 	tshingombe fiston <tshingombefiston@gmail.com>
________________________________________
metadabse
________________________________________
tshingombe fiston <tshingombefiston@gmail.com>	Tue, Oct 10, 2023 at 12:53 PM
To: tahitaditshingombe@gmail.com, tshingombe fiston <tshingombefiston@gmail.com>, tshingombe tshitadi <tshitaditshingombe@gmail.com>, tshigombekb@gmail.com


On Tue, Oct 10, 2023 at 12:29 PM tshingombe fiston <tshingombefiston@gmail.com> wrote:
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
þÿÿÿ      ÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿýÿÿÿþÿÿÿþÿÿÿþÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿRoot Entryÿÿÿÿÿÿÿÿ[1]ði*ÆÜÎ ž˜ªWJO 
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

ƒ 
€ [1] 
________________________________________
}¨mN+Root Entryÿÿÿÿÿÿÿÿ[1]ði*ÆÜÎ ž˜ªWJO[1][1]bûÙ @ 
________________________________________
f 
________________________________________
[1]ÿÿÿÿÿÿÿÿÿÿÿÿbo 
________________________________________
[1] 
________________________________________
ÿÿÿÿþÿÿÿi01 ÿÿÿÿ  n`ôÎ ›Íª`Ž³ébûÙ n÷bûÙÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿ ýÿÿÿþÿÿÿþÿÿÿÿýÿÿÿ



 ¬ !"#$%&'()*+,-./0123456789:;<=>?@ABCDEFGHIJKLMNOPQRSTUVWXYZ[\]^_`abcdefghijklmnopqrstuvwxyz{|}~ €f 
________________________________________
[1]ÿÿÿÿÿÿÿÿÿÿÿÿ 
________________________________________
™o 
________________________________________
[1] 
________________________________________
ÿÿÿÿ
xèCompObj [1]ÿÿÿÿÿÿÿÿÿÿÿÿpCompObj [1]ÿÿÿÿÿÿÿÿÿÿÿÿ
n
[1]þÿÿÿ 
________________________________________
     þÿÿÿþÿÿÿ
þÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿ 
________________________________________
(

ƒ 
€ [1] 
________________________________________
}¨mN+,be$Õþÿ 
________________________________________

ÿÿÿÿ  n`ôÎ ›Íª`ŽMicrosoft Forms 2.0 FrameEmbedded Object
Forms.Frame.1ô9²q 
________________________________________
8H


________________________________________
€ 
________________________________________
€ÿÿ}_a$Frame1 
________________________________________
Rã
‘ Î ãªK¸Q DB Tahoma4‡(õ €[1]2Ô Label1{[1]Ô(å      € 
________________________________________
4ComboBox1€ §Ê(õ € 
________________________________________
28[1] Label2§ì      $å €xt 
________________________________________
TextBox1{[1]

(å € € 
________________________________________
ComboBox2€ §1 ‚ƒ„…†‡ˆ‰Š‹Œ Ž ‘’“”•–—˜™š›œ žŸ ¡¢£¤¥¦§¨©ª«¬-®¯°±²³´µ¶•¸¹º»¼½¾¿ÀÁÂÃÄÅÆÇÈÉÊËÌÍÎÏÐÑÒÓÔÕÖ×ØÙÚÛÜÝÞßàáâãäåæçèéêëìíîïðñòóôõö÷øùúûüýþþÿÿÿþÿÿÿ[1]´(¢€dhet st peace college student markrsheet , assessment  policy circulum license  issue certificate   term   managemnt system information ,student learner  lecture r‚S" 
________________________________________
[1]5 €¥[1]Tahoma[1] AE€H€, 
________________________________________
[1]ø *U„ 
________________________________________
TEMPLATE OF ICASS IRREGULARITY REGISTER                                                             

NAME OF COLLEGE                                                            

EXAM CYCLE & YEAR                                                         

                                                            

                                                            

DATE CENTRE NAME CENTRE NO     ID      OFFERING SUBJECT   LEVEL ICASS MARK SHEET   CATEGORY IRREGULARITY                    

                              TASK                            

/                                                          

ACTION TAKE                                                              

                                                            

                                                            

                                                            

                                                            

                                                            

0   0   0   0   0   0   0   0   0                  

                                                            

                                                            

TRIMESTER SUBJECT ASSESSMENT PLAN                                                             

Subject                          TRIMESTER                              

Name                                                          

                                                            

                                                            

                                                            

                                                            

                                                            

                                                            

                                                            

     Assess- Assessment Content Duration       Lecturer Moderator     Submission date       Asses-     Completion date              

     ment tool   coverage and mark                for pre- assessment moderation ment of post-            

     task             allocation                  date      moderation             

                                                            

                                                            

                                                            

                                                            

                                                            

                                                            

                                                            

                                                            

     79 79 79 79 79 79 79 79 79           

                                                            

                                                            

ontent coverage and weighting must be in accordance with the subject syllabi.                                                            

v Duration and mark allocation must be in accordance with the subject syllabi and external examination paper/s.                                                              

v Assessment dates will be informed by the TVET College academic calendar for that academic yea                                                             

                                                            

                                                            

ID   : EVALUATION     SAQA APPLICATION 20191130002                                                        

                                           , 202001305040/ 201911130002                                                         

                                                            

TRIMESTER ASSESSMENT SCHEDULE FOR STUDENTS. ID: N1-N2,N3/N4/N5/N6  ,   N        2010002023812  /   2004007064381  /2011007434332    NATIONAL EXAMINATION                                                              

HIGHER. EDUCATION QUALIFICATION                                                               

                                                            

Subject                                                          

                                                            

                                                            

                                                            

                                                            

                                                            

                                                            

                                                            

N        Assessment      Assessment tool           Content Coverage   Duration and          Assessment date               

          Task                       mark allocation                          

1        Test 1          Marking                                     

                    memorandum                                      

2        Test2           Marking                                     

                    memorandum                                      

                                                            

                                                            

                                                            

                                                            

                                                            

                                                            

                                                            

ID   : EVALUATION     SAQA APPLICATION 20191130002                                                        

                                           , 202001305040/ 201911130002                                                         

                                                            

T3: TRIMESTER ASSESSMENT TASKS FOR NATURAL SCIENCES - ENGINEERING STUDIES/ electrical                                                                

Tasks Time-frame   Type of Assessment Activity       Scope of Assessment   %                                     

          (the duration and proposed mark       Contribution to the ICASS                                     

          allocation can be increased but not            Trimester mark                                    

          reduced)                                              

1   Week 2-4 Test 1    At least the first 30% of the    30%                                        

          The duration and marks should align with      syllabus content must be                                             

          the external examination paper/s, e.g. if the    covered in Test 1                                          

          examination paper counts 100 marks for 3                                                   

          hours, then the test should count 30-35                                                    

          marks for 1 hour duration                                                 

2   Week 5-8 est 2           70%                                         

          The duration and marks should align with      At least 75 - 80% of the                                             

          the external examination paper/s, e.g. if      syllabus content must be                                             

          the examination paper counts 100 marks       covered in Test 2                                         

          for 3 hours, then the test should count 30-                                                  

          35 marks for 1 hour duration. Preferably a                                              

          longer test because of its weight                                                  

          contribution in to the ICASS                                                 

mark                  100%                                      

                                                            

                                                            

                    100%                                      

                                                            

                                                            

                                                            

                                                            

                                                            

                                                            

                                                            

                                                            

                                                            

                                                            

                                                            

RECORD SHEET FOR TRIMESTER COURSES - ENGINEERING STUDIES                                                           

ID   : EVALUATION     SAQA APPLICATION 20191130002                                                        

                                           , 202001305040/ 201911130002                                                         

                                                            

Electrical ID: N1-N2,N3/N4/N5/N6  ,   N        2010002023812  /   2004007064381  /2011007434332    NATIONAL EXAMINATION                                                              

HIGHER. EDUCATION QUALIFICATION                                                               

                                                            

SUBJECT: ............................................................................................................                                                             

YEAR: ......................TRIMESTER:.......                     CASS MARK         FINAL                                

LECTURER:                    TASKS      ICASS                                

                              MARK                            

ICASS TRIMESTER   MARK SHEET                  TEST 1/EST 2                                      

No Student ID number Student Surname & Initials Convert the mark to      Total                                     

               weighted %        100%                                  

1             30%     70%                                        

2                                                          

3                                                          

4                                                          

5                                                          

6                                                          

7                                                          

8                                                          

9                                                          

1110                                                          

12                                                        

13                                                        

14                                                        

15                                                        

16                                                        

17                                                        

18                                                        

19                                                        

20                                                        

21                                                        

22                                                        

23                                                        

24                                                        

25                                                        

26 26 26 26      26                               

1440                                                          

                                                            

                                                            

SEMESTER SUBJECT ASSESSMENT PLAN                                                         

TRIMESTER SUBJECT ASSESSMENT PLAN                                                             

Subject                          Semeste                             

Name                                                          

                                                            

                                                            

                                                            

                                                            

                                                            

                                                            

                                                            

     Assess- Assessment Content Duration       Lecturer Moderator     Submission date       Asses-     Completion date              

     ment tool   coverage and mark                for pre- assessment moderation ment of post-            

     task             allocation                  date      moderation             

1   Test or Marking                                               

     Assignmen     memo or                                               

          Rubric                                                  

2   Test or Marking memo or Rubric                                                 

     Assignment                                                     

3   Internal      Marking memo                                                   

     examination                                                   

                                                            

                                                            

                                                            

                                                            

                                                            

                                                            

                                                            

                                                            

                                                            

                                                            

                                                            

  Semester ASSESSMENT SCHEDULE FOR STUDENTS                                                           

Subject                                                          

                                                            

                                                            

                                                            

                                                            

                                                            

                                                            

                                                            

N        Assessment      Assessment tool           Content Coverage   Duration and          Assessment date               

          Task                       mark allocation                          

1        assignment        Marking      80% of the practical                             

                    memorandum        component of the subject                            

2        Test2           Marking      50% of the syllabus                          

                    memorandum                                      

          Internal                     70-80% of the syllabus                          

          Examination                                              

                                                            

                                                            

                                                            

                                                            

                                                            

Test = 50% of syllabus                                                         

v Assignment = 80% of the practical component of the subject                                                              

v Internal exam = 70-80% of the syllabus us with a minimum duration of 1 ½ hours                                                            

v The test and assignment can be swapped around.                                                        

v Some subjects have very specific promotion requirements that may deviate from this general template. In such cases the prescriptions in the syllabus super cedes                                                              

this general template                                                             

                                                            

                                                            

                                                            

                                                            

                                                            

ID   : EVALUATION     SAQA APPLICATION 20191130002                                                        

                                           , 202001305040/ 201911130002                                                         

                                                            

T3semester  ASSESSMENT TASKS FOR NATURAL SCIENCES - ENGINEERING STUDIES/ electrical                                                                

Tasks Time-frame   Type of Assessment Activity       Scope of Assessment   %                                     

          (the duration and proposed mark       Contribution to the ICASS                                     

          allocation can be increased but not            Trimester mark                                    

          reduced)                                              

1   Week 2-4 Test 1    At least the first 30% of the    30%                                        

          The duration and marks should align with      syllabus content must be                                             

          the external examination paper/s, e.g. if the    covered in Test 1                                          

          examination paper counts 100 marks for 3                                                   

          hours, then the test should count 30-35                                                    

          marks for 1 hour duration                                                 

2   Week 5-8 est 2           70%                                         

          The duration and marks should align with      At least 75 - 80% of the                                             

          the external examination paper/s, e.g. if      syllabus content must be                                             

          the examination paper counts 100 marks       covered in Test 2                                         

          for 3 hours, then the test should count 30-                                                  

          35 marks for 1 hour duration. Preferably a                                              

          longer test because of its weight                                                  

          contribution in to the ICASS                                                 

mark                  100%                                      

                                                            

                                                            

                                                            

                                                            

RUBRIC FOR THE ASSESSMENT OF A PRACTICAL TASK (BUSINESS STUDIES)                                                            

                                                            

                                                            

                                                            

                                                            

                                                            

                                                            

                                                            

                                                            

                                                            

                                                            

                                                            

                                                            

                                                            

                                                            

RUBRIC FOR ASSESSMENT OF PREPARATION OF A FUNCTION ROOM                                                           

Assessment Criteria   Level 4:      Level 3:       Level 2:      Level 1:      Student s Mark                                

     Outstanding Highly Competent   Competent      Not Yet Competent                                    

TASK 1    All the stock and storage Most of the stock and Some of the stock and    Limited or none of the stock                                         

Stocktaking requirements for the task storage requirements for storage requirements for       and storage requirements                                        

     are correctly attended to.      the task are correctly    the task are correctly for the task are attended to.                                       

     (13 - 15 marks)     attended to.    attended to. (6-9 marks (0-5 marks)                                     

          (10-12 marks)                                                 

TASK 2    The student is dressed The student is mostly    ome of the stock and      The student is not dressed                                    

Application of health, correctly; hair, nails and    dressed correctly; hair, nails    storage requirements for correctly; hair, nails and                                          

hygiene and safety    hands are clean and neat.      and hands are clean and      the task are correctly    hands are clean and neat.                                     

practices.   (4-5 marks) neat.      attended to. (6-9 marks) (2 marks)                                        

          (3 marks)                                                  

TASK 3    Setting up done    Most of the setting up is      Some of the setting up is Limited or no setting up                                    

Setting up a function room sequentially and correctly    done sequentially and    done sequentially and process is followed or                                     

     as per the task and correctly as per the task correctly as per the task correctly done as per the                                       

     standard requirements. and standard requirements. and standard requirements.       task and standard                                         

     (7-8 marks) (5-6 marks) (4 marks)      requirements.                                    

                    (0-3 marks)                                    

TASK 4    Correct time management      Correct time management    struggle to manage time       Poor or no time                                       

Efficient time management   and consistent organisation    and organisation allocated       and limited organisation to    management and                                       

     for completion of the task      mostly for completion of the      complete the task       organisation                                       

     (6-7 marks) task                                               

          ( 4-5 marks)     (3 marks)     (0-2 marks)                                     

Total mark                                                        

                                                            

                                                            

                                                            

SUPPORTING CHECKLIST FOR RUBRIC ; ID   : EVALUATION     SAQA APPLICATION 20191130002                                                           

                                           , 202001305040/ 201911130002                                                         

CHECKLIST FOR PREPARATION OF A FUNCTION ROOM                                                         

task   Criteria      Possible weight                                                   

     Stocktaking before setting up the function room.      4                                                

     1.1 Take stock of drinks                                                       

     • Count quantities.                                                    

     • Check for breakages.                                                    

     • Check packaging.                                                       

     • Check whether requisitions were met as per requisition and order forms.                                                         

                                                            

     4                                                     

     1.2 Store drinks correctly                                                         

     Refrigerator and dry store.                                                       

     3   3                                                

     1.3 Take stock of materials, furniture, linen and equipment.                                                         

     • Tables and chairs.                                                         

     • Table cloths, overlays, silence cloths and napkins.                                                      

     • Crockery and cutlery.                                                         

     4                                                     

     1.4 Store materials, furniture, linen and equipment correctly.                                                         

     • Correct stacking of crockery and cutlery.                                                        

     • Correct stacking of table cloths, overlays, silence cloths and napkins. 4                                                 

     • Storeroom clean and neat, easy to find utensils, materials and equipment                                                        

                                                            

     application of health, hygiene and safety practices.                                                   

     2.1 White shirt, black skirt/trousers and black safety shoes. 1                                                       

     2.2 Apron and name tag. 1                                                    

     2.3 Clean and neat nails. 1                                                       

     2.4 Hair neat and tied back. 1                                                       

     2.5 Wash hands before the commencement of the practical. 1                                                    

     Sub Total for Task 2 5                                                    

     3 Setting of the front of the room according to the needs of the meeting.                                                         

     • Lectern                                                      

     • Data projector                                                    

     • Microphone                                                       

     • Panel seating in the front of the room for later in the meeting                                                       

     • Room environmental control                                                    

     • Neatness                                                   

     8                                                     

                                                            

     Sub Total for Task 3 8                                                    

     4 Student's efficient time management.                                                         

     4.1 Prompt arrival for the practical. 1                                                       

     4.2 Completes the task in the set time. 1                                                         

     4.3 Exhibits organisational skills in completing the task. 5                                                    

     Sub Total for Task 4 7                                                    

     Total                                                   

                                                            

                                                            

                                                            

                                                            

                                                            

                                                            

                                                            

RECORD SHEET FOR TRIMESTER COURSES - ENGINEERING STUDIES                                                           

Electrical ID: N1-N2,N3/N4/N5/N6  ,   N        2010002023812  /   2004007064381  /2011007434332    NATIONAL EXAMINATION                                                              

HIGHER. EDUCATION QUALIFICATION  : ID   : EVALUATION     SAQA APPLICATION 20191130002                                                           

                                           , 202001305040/ 201911130002                                                         

SUBJECT: ............................................................................................................                                                             

YEAR: ......................semmester :.......                     CASS MARK         FINAL                                

LECTURER:                    TASKS      ICASS                                

                              MARK                            

ICASS TRIMESTER   MARK SHEET                  TEST 1/EST 2                                      

No Student ID number Student Surname & Initials Convert the mark to      Total                                     

               weighted %        100%                                  

1             30%     70%                                        

2                                                          

3                                                          

4                                                          

5                                                          

6                                                          

7                                                          

8                                                          

9                                                          

1110                                                          

12                                                        

13                                                        

                                                            

                                                            

                                                            

                                                            

IRREGULARITIES                                                         

                                                            

An irregularity is any event, act or omission, or any alleged event, which may undermine or threaten to undermine the integrity, credibility, security or the fairness of the examination processes.                                                           

                                                            

Students could forfeit their results and be suspended from writing examinations for a period of 11 months, should they contravene any of the examination instructions (as per National admission permit and/or Examination Instruction                                                        

                                                            

ASSESSMENT N1,2,3,4,5,6  ENGINEERING ELECTRICAL  ENTRY  OUTCOM                           YES/NO                              

SUBJECT :                                                           

- 2 -                                                             

. General Aims ............................................................................................................... - 3 -                                                             

2. Specific Aims .................................................................................................................. 3                                                        

3. Pre-requisite ................................................................................................................................... 3                                                        

4. Duration............................................................................................................................ 3                                                             

5. Evaluation .................................................................................................................... - 4 -                                                             

6. Learning content ......................................................................................................... - 5 -                                                             

7. Mark allocation in the examination as an indication of the weighting of the different                                                              

modules........................................................................................................................ - 5 -                                                             

Module 1: Principles of Electricity ......................................................................................... - 6 -                                                             

Module 2: Direct Current (DC) Machines ................................................................................... 9                                                        

Module 3: Alternating- Current (AC) Theory ........................................................................... 11                                                             

Module 4: Transformer ............................................................................................................. 13                                                             

Module 5: Alternating Current (AC) Machines ........................................................................ 14                                                             

Module 6: Generation and Supply of Alternating Current (AC) Power .................................. 15                                                         

Module 7: Measuring instruments .                                                             

Introduction to the application of technological principles such as design                                                            

procedures; and                                                            

§ The relationship between Electro technology   and other scientific subjects.                                                         

3. Pre-requisite Student must meet at least one of the following requirements. 3.1 Completed National N3 certificate with Electro technology N3 or Electrical Trade                                                            

Theory N3. 3.2 Passed grade 12 with at least level 4 (50% or D symbol) in Mathematics and Natural                                                        

Science or Electrical subjects. 3.3 Completed NCV level 4 in any engineering programme. 3.4 Passed senior certificate for adult learners with at least level 4 (50% or D symbol) in                                                            

Mathematics and Physical Science.                                                            

4. Duration                                                              

Full-time: 7.5 hours per week. This instructional offering may also be offered part-                                                        

time.                                                             

- 4 -                                                             

                                                            

                                                            

5. Evaluation                                                         

5.1 Evaluation is conducted continuously by means of two formal tests at College level.                                                         

Learner must obtain a minimum ICASS mark of at least 40% in order to qualify to                                                              

write the final examination and a mark will be calculated together in a ratio of 40:60                                                            

to derive the promotion mark. The learner must obtain at least 40% on the final                                                              

examination. The promotion mark will be calculated as follows: Promotion Mark = 40% of (ICASS mark) + 60% of (Exam mark) 5.2 The examination in Electrotechnics N4 (Engineering Studies - Report 191) will be                                                              

conducted as follows:                                                            

Modules 1 to 7 MARKS: 100 DURATION: 3 HOURS CLOSED BOOK: Formula sheet is attached to the question paper Scientific calculators allowed No programmable calculators allowed No references allowed. No external examination papers or memoranda allowed                                                             

Knowledge and Understanding     Applying      Analysing / Synthesis and                                               

          Evaluating                                                

30   40   30   40   20   25                                                 

                                                            

                                                            

                                                            

                                                            

                                                            

Mark allocation in the examination as an indication of the                                                              

weighting of the different modules                                                           

MODULES   WEIGHTING                                                        

Principles of Electricity   30                                                    

2. DC-machines                                                              

3. AC-Theory    20                                                    

4. Transformers                                                            

5. AC-machines      20                                                    

6. Generation and supply of AC-Power                                                            

7. Measuring Instruments                                                            

TOTAL      10                                                   

                                                            

     10                                                   

                                                            

                                                            

                                                            

                                                            

     5                                                     

                                                            

     5                                                     

                                                            

                                                            

                                                            

                                                            

                                                            

                                                            

                                                            

                                                            

LEARNING CONTENT / LEARN EXPLAIN MODULES  POE/ SUBMISSION                                                          

                                                            

-TOPIC LEARN ACTIVITY   /                                                         

EVIDENCE LEARN  :                                                         

                                                            

POLICY                                                           

                              1217                            

                                                            

                                                            

                                                            

                                                            

                                                            

IRREGULARITIES                                                         

                                                            

An irregularity is any event, act or omission, or any alleged event, which may undermine or threaten to undermine the integrity, credibility, security or the fairness of the examination processes.                                                           

                                                            

Students could forfeit their results and be suspended from writing examinations for a period of 11 months, should they contravene any of the examination instructions (as per National admission permit and/or Examination Instruction                                                        

                                                            

SECTION A: PURPOSE OF THE SUBJECT ASSESSMENT GUIDELINES SECTION B: ASSESSMENT IN THE NATIONAL CERTIFICATES (VOCATIONAL) 1 Assessment in the National Certificates (Vocational) 2 Assessment framework for vocational qualifications                                                         

2.1 Internal continuous assessment (ICASS) 2.2 External summative assessment (ESASS)                                                         

3 Moderation of assessment                                                              

3.1 Internal moderation 3.2 External moderation                                                        

4 Period of validity of internal continuous assessment (ICASS) 5 Assessor requirements 6 Types of assessment                                                         

6.1 Baseline assessment 6.2 Diagnostic assessment 6.3 Formative assessment 6.4 Summative assessment                                                              

7 Planning assessment                                                            

7.1 Collecting evidence 7.2 Recording 7.3 Reporting                                                           

8 Methods of assessment 9 Instruments and tools for collecting evidence 10 Tools for assessing student performance 11 Selecting and/or designing recording and reporting systems 12 Competence descriptions 13 Strategies for collecting evidence                                                         

13.1 Record sheets 13.2 Checklists                                                         

SECTION C: ASSESSMENT IN APPLIED POLICING - LEVEL 4 1 Schedule of assessment 2 Recording and reporting 3 Internal assessment of Subject Outcomes in Applied Policing - Level 4 4 Specifications for the external assessment in Applied Policing - Level 2                                                              

4.1 Integrated summative assessment task (ISAT) 4.2 National examinatio                                                             

                                                            

                                                            

                                                            

                                                            

                                                            

                                                            

                                                            

                                                            

                                                            

                                                            

Integration      YES NOT                                                    

To adopt a unified approach to education and training that will strengthen the human resources development                                                            

capacity of the nation.                                                              

• Relevance                                                              

To be dynamic and responsive to national development needs.                                                            

• Credibility                                                         

To demonstrate national and international value and recognition of qualification and acquired competencies                                                            

and skills.                                                              

• Coherence                                                              

To work within a consistent framework of principles and certification.                                                              

• Flexibility                                                         

To allow for creativity and resourcefulness when achieving Learning Outcomes, to cater for different learning                                                            

styles and use a range of assessment methods, instruments and techniques.                                                            

• Participation                                                            

To enable stakeholders to participate in setting standards and co-ordinating the achievement of the                                                            

qualification.                                                              

• Access                                                             

To address barriers to learning at each level to facilitate students  progress.                                                            

Applied Policing                                                         

National Certificates (Vocational) Department of Education 3                                                              

• Progression                                                         

To ensure that the qualification framework permits individuals to move through the levels of the national                                                            

qualification via different, appropriate combinations of the components of the delivery system.                                                        

• Portability                                                         

To enable students to transfer credits of qualifications from one learning institution and/or employer to                                                            

another institution or employer.                                                              

• Articulation                                                              

To allow for vertical and horizontal mobility in the education system when accredited pre-requisites have                                                            

been successfully completed.                                                         

• Recognition of Prior Learning                                                         

To grant credits for a unit of learning following an assessment or if a student possesses the capabilities                                                              

specified in the outcomes statement.                                                            

• Validity of assessments                                                         

To ensure assessment covers a broad range of knowledge, skills, values and attitudes (SKVAs) needed to                                                              

demonstrate applied competency. This is achieved through:                                                              

§ clearly stating the outcome to be assessed;                                                              

§ selecting the appropriate or suitable evidence;                                                           

§ matching the evidence with a compatible or appropriate method of assessment; and                                                         

§ selecting and constructing an instrument(s) of assessment.                                                              

• Reliability                                                         

To assure assessment practices are consistent so that the same result or judgment is arrived at if the                                                           

assessment is replicated in the same context. This demands consistency in the interpretation of evidence;                                                            

therefore, careful monitoring of assessment is vital.                                                           

• Fairness and transparency                                                            

To verify that no assessment process or method(s) hinders or unfairly advantages any student. The following                                                         

could constitute unfairness in assessment:                                                            

§ Inequality of opportunities, resources or teaching and learning approaches                                                              

§ Bias based on ethnicity, race, gender, age, disability or social class                                                              

§ Lack of clarity regarding Learning Outcome being assessed                                                              

§ Comparison of students  work with other students, based on learning styles and language                                                             

• Practicability and cost-effectiveness                                                            

To integrate assessment practices within an outcomes-based education and training system and strive for                                                              

cost and time-effective assessment.                                                              

2 ASSESSMENT FRAMEWORK FOR VOCATIONAL QUALIFICATIONS                                                              

The assessment structure for the National Certificates (Vocational) qualification is as follows:                                                             

2.1 Internal continuous assessment (ICASS)                                                            

Knowledge, skills values, and attitudes (SKVAs) are assessed throughout the year using assessment                                                        

instruments such as projects, tests, assignments, investigations, role-play and case studies. The internal                                                              

continuous assessment (ICASS) practical component is undertaken in a real workplace, a workshop or a                                                         

Structured Environment . This component is moderated internally and externally quality assured by Umalusi.                                                         

All internal continuous assessment (ICASS) evidence is kept in a Portfolio of Evidence (PoE) and must be                                                              

readily available for monitoring, moderation and verification purposes.                                                              

2.2 External summative assessment (ESASS)                                                              

The external summative assessment is either a single or a set of written papers set to the requirements of                                                            

the Subject Learning Outcomes. The Department of Education administers the theoretical component                                                           

according to relevant assessment policies.                                                            

Applied Policing Level 4 (January 2020)                                                            

National Certificates (Vocational) Department of Education 4                                                              

A compulsory component of external summative assessment (ESASS) is the integrated summative                                                             

assessment task (ISAT). This assessment task draws on the students  cumulative learning throughout the                                                              

year. The task requires integrated application of competence and is executed under strict assessment                                                        

conditions. The task should take place in a simulated or Structured Environment . The integrated summative                                                              

assessment task (ISAT) is the most significant test of students  ability to apply acquired knowledge. The integrated assessment approach allows students to be assessed in more than one subject with the                                                              

same integrated summative assessment task (ISAT). External summative assessments will be conducted annually between October and December, with                                                              

provision made for supplementary sittings.                                                            

3 MODERATION OF ASSESSMENT 3.1 Internal moderation                                                        

Assessment must be moderated according to the internal moderation policy of the Further Education and                                                         

Training (FET) college. Internal college moderation is a continuous process. The moderator s involvement                                                              

starts with the planning of assessment methods and instruments and follows with continuous collaboration                                                         

with and support to the assessors. Internal moderation creates common understanding of Assessment                                                        

Standards and maintains these across vocational programmes.                                                         

3.2 External moderation                                                              

External moderation is conducted by the Department of Education, Umalusi and, where relevant, an                                                            

Education and Training Quality Assurance (ETQA) body according to South African Qualifications Authority                                                            

(SAQA) and Umalusi standards and requirements. The external moderator:                                                              

• monitors and evaluates the standard of all summative assessments;                                                         

• maintains standards by exercising appropriate influence and control over assessors;                                                        

• ensures proper procedures are followed;                                                              

• ensures summative integrated assessments are correctly administered;                                                              

• observes a minimum sample of ten (10) to twenty-five (25) percent of summative assessments;                                                            

• gives written feedback to the relevant quality assuror; and                                                              

• moderates in case of a dispute between an assessor and a student. Policy on inclusive education requires that assessment procedures be customised for students who                                                            

experience barriers to learning, and supported to enable these students to achieve their maximum potential.                                                            

4 PERIOD OF VALIDITY OF INTERNAL CONTINUOUS ASSESSMENT (ICASS)                                                            

The period of validity of the internal continuous assessment mark is determined by the National Policy on the                                                              

Conduct, Administration and Management of the Assessment of the National Certificates (Vocational). The internal continuous assessment (ICASS) must be re-submitted with each examination enrolment for                                                            

which it constitutes a component.                                                            

5 ASSESSOR REQUIREMENTS                                                              

Assessors must be subject specialists and should ideally be declared competent against the standards set                                                              

by the ETDP SETA. If the lecturer conducting the assessments has not been declared a competent                                                           

assessor, an assessor who has been declared competent may be appointed to oversee the assessment                                                        

process to ensure the quality and integrity of assessments.                                                            

6 TYPES OF ASSESSMENT                                                            

Assessment benefits the student and the lecturer. It informs students about their progress and helps                                                            

lecturers make informed decisions at different stages of the learning process. Depending on the intended                                                            

purpose, different types of assessment can be used.                                                             

Applied Policing                                                         

National Certificates (Vocational) Department of Education 5                                                              

6.1 Baseline assessment                                                              

At the beginning of a level or learning experience, baseline assessment establishes the knowledge, skills,                                                         

values and attitudes (SKVAs) that students bring to the classroom. This knowledge assists lecturers to plan                                                              

learning programmes and learning activities.                                                              

6.2 Diagnostic assessment                                                         

This assessment diagnoses the nature and causes of learning barriers experienced by specific students. It is                                                            

followed by guidance, appropriate support and intervention strategies. This type of assessment is useful to                                                              

make referrals for students requiring specialist help.                                                         

6.3 Formative assessment                                                            

This assessment monitors and supports teaching and learning. It determines student strengths and                                                           

weaknesses and provides feedback on progress. It determines if a student is ready for summative                                                           

assessment.                                                              

6.4 Summative assessment                                                            

This type of assessment gives an overall picture of student progress at a given time. It determines whether                                                            

the student is sufficiently competent to progress to the next level.                                                            

7 PLANNING ASSESSMENT                                                            

An assessment plan should cover three main processes:                                                        

7.1 Collecting evidence                                                              

The assessment plan indicates which Subject Outcomes and Assessment Standards will be assessed, what                                                              

assessment method or activity will be used and when this assessment will be conducted.                                                            

7.2 Recording                                                         

Recording refers to the assessment instruments or tools with which the assessment will be captured or                                                              

recorded. Therefore, appropriate assessment instruments must be developed or adapted.                                                              

7.3 Reporting                                                         

All the evidence is put together in a report to deliver a decision for the subject.                                                              

8 METHODS OF ASSESSMENT                                                              

Methods of assessment refer to who carries out the assessment and includes lecturer assessment, self-                                                         

assessment, peer assessment and group assessment                                                          

                                                            

                                                            

                                                            

                                                            

                                                            

                                                            

                                                            

                                                            

[1]5 €¥[1]Tahoma[1]( €Label2*UO 
________________________________________
[1]5 €¥[1]Tahoma[1]Xt@€H€,>tVT„ 
________________________________________
ID: N1-N2,N3/N4/N5/N6  ,   N        2010002023812  /   2004007064381  /2011007434332    NATIONAL EXAMINATION                                         

HIGHER. EDUCATION QUALIFICATION                                           

-FINAL AWARD (DEGREE / DIPLOMAT CERTIFICATE) SUBMITED   1STH/                                         

- NO PROVISIONAL CERTIFICATE OR UNOFFICIAL STEMENTS                                         

-CERTIFIE NO SUBMITE 1TH                                         

-OFFICIAL   STATEMENT FROM   INSTITUT                                       

- DIPLOMAT D ETAT EXAM CERTIFIE / NO SUBMITED                                          

-ID: N1-N2, N3/N4/N5/N6,   N        2010002023812 /   2004007064381 /2011007434332   NATIONAL EXAMINATION                                    

- REGISTRAR   CERTIFICATE NO: COM 18269001: /                                          

- ST PEACE COLLEGE LEVEL N ENGINEERING CERTIFICATE   LEVEL 1,2,3,4, REGISTRAR   CERTIFICATE NO: COM 18269001:                                     

-FINAL DEGREE/ DIPLOMAT     DEGRE SAQA   N6    NQF 6/ NQF7 /  NQF8     CONTINUE                                                 

- SAQA   UNIVERSITY DEGREE 1, 2, 3, 4    NQF7/ NQF8   , SUPLEMENTARY PREPARATORY SELECTOR DIPLOMAT                                         

-REGISTRAR FEES:  FINAL EXAM DIPLOMAT N / SAQA 50%                                      

-REGISTRAR FEES FINAL                                       

ELECTRICAL DURATION REG/  FEE     DEPOSIT      MONTHLY   TOTAL      EXPENSION          

ENGINEERING                               STATIONERY        

FACULTY                                      

                              -HOME STUDENT        

                              BOOKING        

                                        

                                        

ELECTRICAL ENGINEERING 18MONTH   R 300 R 900      850/ 900 R15600/   400          

                              ORDER BOOK        

                              COPY        

                                        

                                        

                                        

                                        

ORDER DEBIT                                    

CASH                                    

2019/ 2020                                      

REGISTRAR                                        

                                        

OCTOBER                                      

NOVEMBER                                    

DECEMBER ##########   ##########   #########      #############                  

JANUERY                                      

FEBRUARY                                    

MARCH                                         

APRIL                                    

TERM 1/ SEMESTER 1     *########     ##########      ##########   ##############                

JUNE                                      

JULY                                      

AUGUST                                        

TERM2 ########### ##########   ############      ###########                

SEPTEMBER                                        

OCTOBER                                      

SEPTEMBER                                        

NOVEMBER                                    

DECEMBER                                         

TERM 3                                       

SEMETER3                                    

JANUERY                                      

                                        

FEBRUARY                                    

MARCH                                    

APRIL                                    

TERM 1/   ##########   ##########   #########      ########                

JUNE                                      

JULY                                      

AUGUST                                        

TERM2                                    

SEPTEMBER                                        

NOVEMBER                                    

DECEMBER                                    

                                        

                                        

                                        

SEMESTER                                    

4/                                    

TERM  3                                    

                                        

DISCOUNT      ############     ###########      ########### #############                  

                                        

TERM/                                    

SEMESTER                                         

AWARD                          BOOKING        

FINANCIAL                              REWARD AWARD        

 STATEMENT                          PRACTICAL          

VALUE                               SUBMISSION        

                              ASSESSMENT        

BURSARY                                      

DEVELOPMENT                                    

DISCOUNT                                    

                                        

CREDIT                                        

EXTRA CIRCULATION                                    

TERM1/                                        

                                        

                              2        

                                        

                                        

                                        

                                        

                                        

-                                    

                                        

                                        

                                        

                                        

CREDIT     REQUIRE FILING CONTENT SUBMISION.      CORE ,     CREDIT     TOTAL                

SUBJECT MODULE      ASS .POL / ASS ENGINEERING ELECTRICAL ELECTIVE      MARK                

     COMPLETED WEEK : /     FONDAMENTAL                           

ELECTRICAL LEVEL:NQF 1,NQF2,NQF3,NQF4,NQF5,NQF6      250     AWARD                 

ENGINEERING N1.N2,N3,N4,N5,N6      8   100                     

-NATIONAL FRAMEWORK L1,L2,L3,L4,L5,L6 LEARNER           8   100                    

NATIONAL  DIPLOMA  N, 1TH ,2TH ,3TH SAQA,      LP1,LP2,LP3,LP4,LP5,LP6,LEARNER POLICE      8      100                    

-EXAMINATION     ASS1,ASS2,ASS3,ASS4,ASS5,ASS6 ASSESSOR,          8   100                    

NATIONAL  N,               7   100                    

-LEARNER  ENGINEEER   NQF 7,12               100                     

                                        

-ASSESSOR                                        

                                        

                                        

                                       

SUBJECT LEVEL:NQF 1,NQF2,NQF3,NQF4,NQF5,NQF6                                    

ENGINEERING N1.N2,N3,N4,N5,N6                                

SCIENCE L1,L2,L3,L4,L5,L6 LEARNER                                

ENGINEERING DRAWING N1/2      LP1,LP2,LP3,LP4,LP5,LP6,LEARNER POLICE                                

TRADETHEORY ELECTRICAL      ASS1,ASS2,ASS3,ASS4,ASS5,ASS6 ASSESSOR,                                    

-ELECTROTECHNOLOGY                                        

ELECTROTECHNIC NQF 7,12                                   

MATHEMATICS                                    

INDUSTRIAL ELECTRONICS                                      

LOGIC SYSTEM                                       

EDUCARE                                    

INDUSTRIEL ORIENT                                         

PLANING                                          

TOPIC /                                      

POLICING  NQF4,5,6                                        

INFORMATION                                         

MANSYST INFO                                       

MANAGE SYS                                      

ADM COMMUNICAT SYST                                      

INVESTIGAT CONDUCT                                       

TRAFFIC LOW,                                        

PARALEGAL                                       

OUTCOM TOTAL CREDIT  FILE  BALANCE   Hod  filing department    document need policing                                

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

COMPLETED REQUIRED: week exercise book n3, 4, 5, 6   nqf 6                                        

CREDIT     REQUIRE FILING CONTENT SUBMISION.      CORE ,     CREDIT     TOTAL                

SUBJECT MODULE      ASS .POL / ASS ENGINEERING ELECTRICAL ELECTIVE      MARK                

     COMPLETED WEEK : /formative summative       FONDAMENTAL                          

ELECTRICAL LEVEL:NQF 1,NQF2,NQF3,NQF4,NQF5,NQF6      250     AWARD                 

ENGINEERING N1.N2,N3,N4,N5,N6      8   100                     

-NATIONAL FRAMEWORK L1, L2, L3, L4, L5, L6.   LEARNER      8   100                    

NATIONAL  DIPLOMA  N, 1TH ,2TH ,3TH SAQA,      LP1,LP2,LP3,LP4,LP5,LP6,LEARNER POLICE      8      100                    

-EXAMINATION     ASS1,ASS2,ASS3,ASS4,ASS5,ASS6 ASSESSOR,          8   100                    

NATIONAL  N,               7   100                    

-LEARNER  ENGINEEER   NQF 7,12  , attendance , non attendance years              100                    

                                        

-ASSESSOR                                        

                                        

                                        

                                       

SUBJECT LEVEL:NQF 1,NQF2,NQF3,NQF4,NQF5,NQF6                                    

ENGINEERING N1.N2,N3,N4,N5,N6                                

SCIENCE L1,L2,L3,L4,L5,L6 LEARNER                                

ENGINEERING DRAWING N1/2      LP1,LP2,LP3,LP4,LP5,LP6,LEARNER POLICE                                

TRADETHEORY ELECTRICAL      ASS1,ASS2,ASS3,ASS4,ASS5,ASS6 ASSESSOR,                                    

-ELECTROTECHNOLOGY     Completed: nqf 6/nqf5 formative summative /.                                 

ELECTROTECHNIC                                    

MATHEMATICS ........................................................................                                    

INDUSTRIAL ELECTRONICS ....................................................                                    

LOGIC SYSTEM                                       

EDUCARE                                    

INDUSTRIEL ORIENT                                         

PLANING                                          

TOPIC /                                      

POLICING  NQF4,5,6                                        

INFORMATION                                         

MANSYST INFO                                       

MANAGE SYS,                                    

 DELIVER                                         

 INVESTIGATION METHODE CTECHNIQUE                                       

TRAFFIC ROAD                                       

                                        

                                        

                                        

Protection device                                         

Survey                                       

                                        

                                        

                                        

                                        

QCTO Verification      Mon, Sep 20, 3:42 PM (1 day ago)                               

                                        

                                        

                                        

                                        

to TSHINGOMBE520@GMAIL.COM, QCTO, Tshepo                                     

                                        

Dear Sir/Madam                                    

 Thank you for the enquiry.                                         

 The QCTO does not complete foreign qualification evaluations.  This is completed by SAQA.                                    

 Please forward your enquiry to SAQA or indicate what assistance you would request from the QCTO.                                    

                                        

Accreditation via qctoorgza.onmicrosoft.com       Sun, Sep 5, 1:05 PM                                 

                                        

                                        

                                        

                                        

                                        

to me                                         

                                        

Good Day                                         

                                        

Kindly be advised that all accreditation applications ( Skills Development Provider, Assessment Centre, Letters of Intent & Skills programmes)  applications  are done online in QCTO website.                                     

Click on the link below and complete the application form:-                                       

https://forms.office.com/Pages/ResponsePage.aspx?id=WYKvGtx_dkOLCRj6eXre7cQfdpB9YCFAsVRDtPf78FpUMlJIRVlHWTlONkk1SkdJMDY0SUc3TlRQMS4u                                          

Please liaise with centraloffice@qcto.org.za for any enquiries.                                       

Thank you.                                     

                                        

This message is intended for the addressee only and is confidential and the copying thereof is prohibited. The above information may contain personal views of the author thereof and is not necessarily the views of the Quality Council for Trades and Occupations (QCTO) and the QCTO does therefore not accept liability for any damages arising from the correctness of the facts stated in this communication, unless specifically stated. If you have received this message in error, please notify the sender immediately and destroy the original message.                                          

202001305040/ 201911130002                                     

Inbox                                    

                                        

non-compliant@saqa.org.za via saqa.co.za             Jul 27, 2021, 2:56 PM                         

                                        

                                        

                                        

                                        

                                        

to me                                         

                                        

Dear Sir/Madam                                    

                                        

APPLICATION FOR THE EVALUATION OF FOREIGN QUALIFICATIONS                                    

The application for the evaluation of the foreign qualifications of the above-mentioned does not meet all SAQA s requirements stipulated in the application material (www.saqa.co.za> Services > Evaluation of Foreign Qualifications > Online application) and has therefore been returned to you.                                         

Below is what was outstanding on the application:                                        

                                        

ðð ð ð ð ð ð ð ð ðAn application form / invoice generated from the online application system, which bears a signature and date.                                         

ðð ð ð ð ð ð ð ð ðProof of payment of the amount reflected on the system generated application form / invoice, which bears the online submission number as reference.(Only cash deposit or card payments are acceptable method of payments)                                        

ðð ð ð ð ð ð ð ð ðValid proof of the identification of the above-mentioned qualification holder.                                     

ðð ð ð ð ð ð ð ð ðQualification documents  (original language and English if applicable)                                     

o        Final Award  - School/Diploma/Degree certificate                                    

                                        

                                        

In re-submitting your application, please ensure that you return the whole application pack including the outstanding item(s) within 48hrs and pay special attention to the What must be submitted to SAQA  document available in the online application facility.                                     

                                        

                                        

Please note response must ONLYbe sent to non-compliant@saqa.org.za                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

Final Award  - School/Diploma/Degree certificate: TAX : Proof of payment of the amount reflected on the system generated application form / invoice, which bears the online submission number as reference/ cash deposit or card payments are acceptable method of payments: 2010002023812  /   2004007064381  /2011007434332    , REGISTRAR                                        

CERTIFICATE NO: COM 18269001:                                     

Attandence      registrar  + (R900- R 300  total 1200/ DEPOT      04-04-2019                                  

REGULARE           SCHEDULE    ATTENDENCE  DAY   CASH ASSESSEMENT, CASH MODERATE   FEE STUDENT                                      

 NAME STUDENT : TSHINGOMBE TSHITADI      /       ST PEACE COLLEGE                                          

 STATEMENT                                    

CREDIT / TERM1,2,3,   ORDER CASH DEBIT DATE  /                               

SEMESTER  1,2,3,4                                         

1   R 850 05-03-2020                            

2   R 850 01-05-2020                            

3   R 1,600   03-06-2020                            

4   R 850 07-08-20220                          

          **                          

5   R 800 14-08-2020                            

6   R 1,050   14-09-2020                            

          ****                            

6   R 700 14-09-2020                            

7   R 800 11-11-2020                            

8   R 800 10-02-2021                            

9   R 850 01-03-2021                            

10 R 800 31-03-2021                            

11 R 850 07-05-2021                            

12 R 800 10-02-2021   R -50                     

1   R 800 01-03-2021   R -50                     

3   R 700 31-03-2021   R -100                         

14 R 700 07-07-2021   R -100                         

15 R 700 05-08-2021   R -100                         

16 R 700 8/21/2028     R -100                         

17 R 600 07-12-2021   -400                       

TOT/BALANCE 15900      R 700                     

POINT  AWARD MARK R@ 15900      POST /OFFICE                             

REWARD AMANDEMNT   POST BANK    ACCOUNT /                               

REGISTER      STAMP                               

                                        

-ORDER BOOKING -R320@   X4      BOOKEEPING /LEARN                              

PHOTOCOPY          R400X4=                            

SUBMITTED  NOTE BOOK           1200/                           

SUBJECT    LIBRARY         REWARD    R5000                               

                                        

                                        

DISCOUNT                                         

PAY  IREGULAR AMOUNT                                         

TOT    POINT SCORE                                    

TENDERED      10  / 15900                               

DISCLOSE                                    

                                        

REQUIRE , CRITERIA                                       

VALUE /SCORE                                       

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

Weighting:The following weights are consequently awarded to each category                                           

CERTIFICATE AND DIPLOMAT ENGINEERING ELECTRICAL                                    

                                        

Knowledge and Understanding     APPLYING       ANALYSE SYNTHESIS EVALUATION INVESTIGATION      TOTAL CRITERIA CLOSE                     

               DISCOVERY/ DESIGN / ASS POL                          

3--40 30-40 20-25                          

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

               40 ,60     10 POINT SCORE                     

REQUIRED                                    

TASK                                      

OPERATION                                        

                                        

MATHEMATIC                                      

                                        

                                        

                                        

                                        

                                        

ELECTROTECH                                    

                                        

                                        

                                        

                                        

                                        

POWER MACHINE                                      

                                        

                                        

                                        

                                        

                                        

ENGINEERING SCIENCE                                    

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

STUDENT  NAME : TSHINGOMBE    TSHITADI                                          

 ID NUMBER :                                                 

ID   DIPLOMAT NUMBER : ..Q                                     

  SHEET MARK  SUBMISSION  EXPLANATION EXAM                                       

MODULE /SUBJECT   ELECTRICAL WEIGHTING      OUTCOM CRITERIA CREDIT CLOSE AWARD SCORE FINAL /QUALIFY                              

ENGINERING                                    

MATHEMATIC 100MARK   MIN             / MAX   100MARK                                

                                        

1                                      

2                                      

3                                      

4                                      

5                                      

6                                      

7                                      

0.8                                        

                                        

                                        

                                        

                                        

                                        

ELECTROTECHNICAL   100MARK                                 

                                        

                                        

                                        

                                        

POWER MACHINE 100MARK                                 

                                        

                                        

                                        

                                        

                                        

                                        

                                        

ENGINEERING SCIENCE 100MARK                                

                                        

                                        

                                        

                                        

                                        

ELECTRICAL ENGINEERING 400MARK /     400 MARK                                 

DIPLOMA / STATEMENT STATEMENT                                          

                                        

RATING                                       

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

SCHEDULE CASE REGISTRAR ATTATNDANCE RECORD  SHEET  AMANDEMENT TEXTBOOK WEEK COMPLETED COVERY , INVIGILATOR, AMANDMENT COPYRITH DARLO. / N1-N3////N4-N6                                          

NAME SURNAME     WEEK /PRESENT ID NUMBE,     CELL PHONE /TIME/IN OUT    SIGN                

     ABSENT     TEXT BOOK NUMBER  ISBN                                

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

                                        

Booking  mark:                                    

                                        

                                        

                                        

                                        

                                        

how to reference a caps document south africa                                         

                                        

curriculum and assessment policy statement                                         

                                        

[1]5 €¥[1]Tahoma[1]`AE€H€, 
________________________________________
[1]A€þUö 
________________________________________
      1.register  national  examination ,                                     
      N diplomat. Examination n 1.n6 diplomat t1.                                         
      I'd /name /years///file student/submitted document file///                                      
      Courses attandance////exam attended                                     
      I'd number Name surname      years file student      submitted document file Courses attandance  exam attended     
                                                
                                                
                                                
                                                
                                                
                                                
                                                
      52    0     0     0     0     0     0     
                                                
                                                
      I'd number registered.regulier  diploma n                                           
      18month.attandance.term 1.term2.term.3week                                          
      Level1.2.3 minimum engineering electrical learning national trade                                     
      Registered.. regulier/                                      
                                                
                                                
      I'd number  candidate.//.I'd regulier.//                                      
      I'd name///class level///file number//submitted number ///documents attached                                      
      National                                        
      I'd number        class level file number submitted number  documents attached      Courses attandance      exam attended     
                                                
                                                
                                                
                                                
                                                
                                                
      71    71    71    71    71    71    71    
                                                
      N1.n3..rwiten final engineering                                         
                                                
      N1,N2, council test trade .                                       
      Councils education..                                        
                                                
      I'd number candidate irregularity register                                          
      Reg .I'd number submitted.                                        
                                                
      Rectorat college director principal                                           
      System                                          
      College internal registered. St peace                                         
      And institutor ..distance university                                          
      .grade..1..12... level 1.2.63.4.5.6                                     
      Under graduat .1.2.3.4                                      
      Learner                                         
      Teacher                                         
      Lecture                                         
      Professional                                          
      Subject faculty  admnise                                          
      Regularity .. irregularity ruling                                       
      I'd name . Term 1.2.3.4.5.6.7.8.9.semester1,2                                       
      Report internal diploma.certificate award . internal  statement internal report . homework classwork test .exam internal sylabus hand book campus module practice.                                        
                                                
      .grade..1..12... level 1.2.63.4.5.6                                     
                                                
      Under graduat .1.2.3.4                                      
                                                
      Learner                                         
      Teacher                                         
      Lecture                                         
                                                
      Subject faculty  admnise                                          
                                                
      Regularity .. irregularity ruling                                       
                                                
      I'd name . Term 1.2.3.4.5.6.7.8.9.semester1,2                                       
                                                
            2                                   
                                                
                                                
                                                
                                                
                                                
                                                
                                                
                                                
                                                
                                                
                                                
                                                
      Report internal diploma.certificate award . internal  statement internal report . homework classwork test .exam internal sylabus hand book campus module practice.                                        
      Assessment assignment homework practice theory skill give to student to prove if student at home classes on completion is capable to resolve trade theory Test is capable to working by self group peer                                        
      Module correct diagnostic                                         
      Manufacture maintenance testify attest award brevet certificate is true                                     
      Recording examination.diplomatic to council of test function working yes and to evaluate grade level n it test comming rather working nice.                                     
      T                                         
      Test circuit.nice erroneous value home.                                       
      Test operational    Commission.                                         
      1.homework class work exercise books topics research on line Poe exercise book.//capacity to make reproduct  analyse rwiten.///criteria minimum requirements 100                                        
                                                
      homework class work exercise books topics research on line Poe exercise book  capacity to make reproduct  analyse rwiten   criteria minimum requirements 100                           
                                                
                                                
                                                
                                                
                                                
                                                
      2     2     2     2                       
                                                
      2.test evaluation module topics test research Poe's /functional school academic task system function.///                                         
      test evaluation module topics test research Poe's     functional school academic task system function                                 
                                                
                                                
                                                
                                                
                                                
                                                
      2     2     2                             
                                                
                                                
                                                
                                                
                                                
                                                
                                                
                                                
      3.examination evaluation diagnostic module external internal /low competency year term weekend rating period achieve rerwite.                                        
      examination evaluation diagnostic module external internal  low competency year term weekend rating period achieve rerwite                                   
                                                
                                                
                                                
Ï¬œ[1]5 €¥[1]Tahoma[1]( €Label3þU# 
________________________________________
[1]5 €¥[1]Tahoma[1]èZAE€H€, 
________________________________________
[1]ÊZ*Uö 
________________________________________
n a module with the credit value (cv) of that module and then dividing the sum of these values "(fm.cv) by the total of the credit values (ct) of all the modules a student was enrolled for. See the following example and equation: 



 

Calculate Grades


You are viewing Ultra Course View content

Go to Original Course View page.


Your institution controls which tools are available in the Original Course View. The gradebook is always available to instructors in the Ultra Course View.

About gradebook calculations


You can easily add calculations to your course gradebook. A calculation is a formula that produces a numerical result used to view or assign grades, usually based on other graded items.


You can create your own formulas and use common arithmetic operations, including addition, subtraction, multiplication, and division, and use group operators (parentheses).


You can add calculations based on the average, total, maximum, or minimum of the variables you include, such as categories, graded items, and other calculations. For example, add a calculation that displays the average of all assignments so students have an overall picture of their performance. You can add as many calculations as you need.


In the Ultra Course View, each course has one default grading schema used for grades and calculations. You can't create new schemas at this time, but you can make changes to the default schema.


More on the default grading schema


More on grade calculations and decimal handling

Create calculations inline


In the Gradable Items list or student grid view, select the plus sign wherever you want to add a calculation and select Add Calculation.


More on the two gradebook views

Add calculation option in the gradebook

Calculation interface


Type a meaningful title for the calculation. If you don't add a title, New Calculation and the date appear in the gradebook. You can use the placeholder text as the title if the formula on the page is valid and saved. Optionally, add a description and make the calculation column visible to students. Students see calculated grades on their Grades pages, but they don't see your descriptions or formulas.


Determine how the result of the calculation appears. In the Select a grade schema menu, choose Points, Percentage, or Letter.


Create your formula. In the left pane, select a function, variable, or operator to add it to the right pane.

Instructor grading calculation

Functions and Variables


    Average: Generates the average for a selected number of graded items, categories, and other calculations. For example, you can find the average score on all tests.

    Total: Generates a total based on the cumulative points, related to the points allowed. You can select which graded items, categories, and other calculations are included in the calculation.

    Minimum: Generates the minimum grade for a selection of graded items, categories, and other calculations. For example, you can find the minimum score on all assignments.

    Maximum: Generates the maximum grade for a selection of graded items, categories, and other calculations. For example, you can find the maximum score on all discussions.

    Variable: Select an individual graded item or calculation from the menu. You may only add one variable at a time. Continue to add variables from the left pane to add as many variables as you need.


Operators


    Add ( + )

    Subtract ( - )

    Divide ( / )

    Multiply ( * )

    Open Parenthesis (

    Close Parenthesis )

    Value: After the text box appears in the formula, click in the box to add a numeric value. You can include seven digits before a decimal point and four digits after it. When the calculation is generated and appears in students' grade pills, only two digits appear after the decimal point.


Create your formula


For example, select Total in the left pane to add that function to the right pane. Expand the list and select the check boxes for the items you want to add to the formula. When you choose a category, all items in that category are included. You must choose graded items and other calculations individually. Scroll through the list to view all items. In the Variable menu, select an item to choose it.

Add functions and variables to gradebook calculations


After you make a selection in a menu, click anywhere outside of the menu to exit and save the selection in the right pane. Each element you add to the formula appears at the end. You can press and drag any added element to reorder your formula. To remove an element, select it and select the X. You can reuse any function, variable, or operator.


When you select Save or Validate, the system checks the accuracy of your formula. Validate checks the formula while you remain on the page. You can't save a calculation until it's mathematically valid.


Select Clear to remove all elements from the right pane and start over.


Example formula for the total for the first quarter:


Create a Total calculation that includes the Assignment and Test categories and the Attendance grade, but doesn't include the Pop Quiz grade.


The Assignment and Test categories are in the Total menu. Attendance and Pop Quiz are individual graded items in the Variable menu.


Formula: Total of Assignment category + Test category + Attendance - Pop Quiz

Example of gradebook calculation


If the formula isn't valid, an inline error message appears next to Validate. Problems in your formula are highlighted in red in the right pane.


Example error messages:


    Unmatched operator: Symbols such as (+) or ( -) don't match up with another part of the formula. Example: Graded item + (nothing).

    Unmatched function, variable, or value: Typically appears when an operator is missing between two variables, such as two graded items or categories.

    Some error messages are specific, such as Unmatched opening parenthesis, to alert you to exactly what's missing.


Example of gradebook calculation error


Your newly created calculated item appears in your gradebook. In item list view, press the Move icon in the row of the calculation to drag it to a new location and release. The order you choose also appears in the grid view and on students' Grades pages.

Example of calculation in gradebook


Reminder: Students won't see the calculation until it has a grade and you make the item visible to them.

Delete graded items in a calculation


If you delete a graded item used in a calculation, you receive a warning when you open the calculation:


An item was removed from the gradebook that was used in this calculation. We've updated the calculation where possible, but it may need your attention.


You may need to update the calculation. Students see the updated calculation on their Grades pages if you made the calculation visible to them.

Create weighted calculations


A weighted calculation generates a grade based on the result of selected graded items, categories, other calculations, and their respective percentages. You can use normal arithmetic operations to create the weighting you need.


For example, if you have four tests and one final test, you can weight each for a "Weighted Tests" calculation.


Select tests individually from the Variable menu and select Value to add the percentage for each test, such as .15. Add the operators needed such as Multiply and Add.


Basic formula:


Test 1 x .15 + Test 2 x .15 + Test 3 x .15 + Test 4 x .15 + Final Exam x .40


You can include parentheses in this example if you want:


(Test 1 x .15) + (Test 2 x .15) + (Test 3 x .15) + (Test 4 x .15) + (Final Exam x .40)


The calculation is valid either way and produces the same result.

Gradebook weighted calculation example


In item list view, select the calculation to view each student's calculated grade or access the student grid view and navigate to the column.

Example of weighted tests


More on weighting and the overall grade


 

Keyboard commands


If you use keyboard navigation, you can tab between panes to build a formula.


Add functions and operators to the formula. In the Function and Operator panes, use the up and down arrow keys to select an item from the list and press Enter to add it to the formula.


Reorder the formula. Focus on a formula item and press Enter. Use the arrow keys to move the item and press Enter again to place it.


Select items to include in functions. Focus on a function or variable in the formula and press Spacebar. From the menu, select the gradebook items to include. Press Esc to close the menu.

Print

More help

Follow this page

Need more help with Learn?


    Contact your institution's support desk



    Watch videos for instructors

    Join the Community 


Follow this page!

tshingombe fiston

     

11:43/ AM (21 minutes ago)

     

to tahitaditshingombe, tshingombe, tshigombekb, TSHINGOMBEKB, me


Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        'declare variables

        Dim grade As String = 0

        Dim passNumber As Integer

        Dim failNumber As Integer

        'calculate

        grade = Me.txtGrade.Text

        If grade = "D" Or grade = "C" Or grade = "A" Or grade = "B" Or grade = "d" Or grade = "c" Or grade = "a" Or grade = "b" Then passNumber = passNumber + 1

        If grade = "F" Or grade = "f" Then failNumber = failNumber + 1

        'display message

        Me.lblFail.Text = failNumber

        Me.lblPass.Text = passNumber

    End Sub

ivate passNumber As Integer

Private failNumber As Integer


Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

    'declare variables

    Dim grade As String = 0


    'calculate

    grade = Me.txtGrade.Text

    If grade = "D" Or grade = "C" Or grade = "A" Or grade = "B" Or grade = "d" Or grade = "c" Or grade = "a" Or grade = "b" Then passNumber = passNumber + 1

    If grade = "F" Or grade = "f" Then failNumber = failNumber + 1

    'display message

    Me.lblFail.Text = failNumber

    Me.lblPass.Text = passNumber

End Sub

Private passNumber As Integer

Private failNumber As Integer


Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

    'Using .ToUpper as suggested in Comments by David Wilson

    Dim grade = txtGrade.Text.ToUpper()

    'I used a Select Case because I think it is 

    'easier to read.

    Select Case grade

        Case "A", "B", "C", "D"

            'the += 1 is just a shortcut for

            'writing = passNumber + 1

            passNumber += 1

        Case "F"

            failNumber += 1

        Case Else

            MessageBox.Show("Invalid grade")

    End Select

    'You need to convert your Integers to Strings

    'so they can go into the Text property of your labels.

    lblFail.Text = CStr(failNumber)

    lblPass.Text = CStr(passNumber)

End Sub


im passNumber As Integer

Dim failNumber As Integer


Dim PassArray As String() = {"a", "b", "c", "d"}

Dim FailArray As String() = {"f"}


Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

    'declare variables

    Dim grade As String = Me.txtGrade.Text.ToLower()


    If PassArray.Contains(grade) Then

        passNumber += 1

    ElseIf FailArray.Contains(grade) Then

        failNumber += 1

    Else

        MsgBox("Error: Invalid Grade")

    End If


    'display message

    Me.lblFail.Text = failNumber

    Me.lblPass.Text = passNumber

End Sub










            t[1]5 €¥[1]Tahoma(õ €28 Label3x1§ú(å    € [ ComboBox3€ §ð €# 
________________________________________

Frame1ö 
________________________________________
O 
________________________________________
þÿ 
________________________________________

ÿÿÿÿði*ÆÜÎ ž˜ªWJOMicrosoft Forms 2.0 FormEmbedded Object 
Forms.Form.1ô9²q



8 attachments
 		UserForm1 policy assessment circulum  issue certificate.frm
3K Scan and download

	
 		UserForm1 policy assessment circulum  issue certificate.frx
129K Scan and download

	
 		UserForm1 vb excell policy marsheet.frm
2K Scan and download

	
 		meta database vb excell student module.docx.
61K View as HTML Scan and download

	
 		tshingombe fiston affidavit result.docx
50K View as HTML Scan and download

	
 		UserForm1 vb excell policy marsheet.frx
73K Scan and download

	
 		vb excell programming.docx
20K View as HTML Scan and download

	
 		programmation vb excell sheetmark,,data.docx
71K View as HTML Scan and download

	


 

sales@iec.ch










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


















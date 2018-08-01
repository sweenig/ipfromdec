# IPFromDec.xlam
Custom Excel Add-In to work with IP Addresses

It seems I am frustrated by Excel's inability to handle IP addresses natively.  Eventually, i should write some global formulas mimicing inet_ntoa() and inet_aton() from MySQL.  In the mean time, I've been creating the formulas manually to slice and dice IP addresses.  It's not actually too hard, as long as you're not scared by long formulas.  I usually build this formula in pieces then consolidate them down to one (I should write a procedure that would do that automatically).  Eventually...

Anyway, the formula is this:
```
=VALUE(LEFT(A1,FIND(".",A1)-1))*2^24+VALUE(MID(A1,FIND(".",A1)+1,FIND(".",A1,FIND(".",A1)+1)-FIND(".",A1)-1))*2^16+VALUE(MID(A1,FIND(".",A1,FIND(".",A1)+1)+1,FIND(".",A1,FIND(".",A1,FIND(".",A1)+1)+1)-FIND(".",A1,FIND(".",A1)+1)-1))*2^8+VALUE(RIGHT(A1,LEN(A1)-FIND(".",A1,FIND(".",A1,FIND(".",A1)+1)+1)))
```

You can put this formula into any cell and it will take the IP address in cell A1 and convert it to decimal.  The nice thing about this is that you can apply this formula to a list of IP addresses then use it to sort the IP addresses.  That way 10.200.0.0 comes after 10.5.0.0.

This add-in makes it so that instead of doing the formula above, you can do this:

```
=IP2DEC(A1)
```

The process is fairly easy if you only want to use the formula in one spreadsheet. If you want to use it globally, there are a few extra (but easy) steps.

## Manually building it yourself
The first is to open Excel and save a blank workbook as an xlam file (Excel add-in). Just go to Save As..., choose 'other formats' and pick Excel Add-in from the drop down. When you do, it will change the location, there's no need for you to save it in a different location. Give it a good name like MyFunctions.xlam.

Then enable that add in. Go to the Office button and click 'Excel options'. Then go to Add-Ins and click 'Go...'. Check the checkbox next to your add-in and you're done. Now you've got a place that you can drop VBA functions that will always be available. Now to get some in there.

Now open a blank excel document. For the purposes of this example, I'll add some data. In A1, I'll put `=RANDBETWEEN(0,255)&"."&RANDBETWEEN(0,255)&"."&RANDBETWEEN(0,255)&"."&RANDBETWEEN(0,255)`. This will generate a random IP address that I'll use to test my function on. I'll put the above formula in B1 as a control to make sure my custom function works properly.

Now to create your custom function:

Press `Alt+F11` or go to `Developer>>View Code` (you might not see this option, it's hidden by default).

You'll notice that the code browser will show two items, your current new book and your blank add-in. Whether or not your custom function becomes global or stays within the current workbook will depend on where you insert the module. To make it local to the current workbook, make sure the current workbook is selected. To make it global, make sure your add-in is selected.

Click `Insert>>Module`.

Now you can start building your function. Start with `Public Function IP2DEC(ipaddress as string)`. VB will automatically add the End Function section. `IP2DEC` is the name of the function you're going to create and `ipaddress` is the argument that is required.

I put in the following code for my function. I know, it can be done more easily using more sophisticated commands, but I did it this way to illustrate:

``` VBA
Public Function IP2DEC(ipaddress As String)
'find the location of the first dot
Dim firstdot As Integer, seconddot As Integer, thirddot As Integer
firstdot = InStr(ipaddress, ".")
seconddot = InStr(firstdot + 1, ipaddress, ".")
thirddot = InStr(seconddot + 1, ipaddress, ".")
'get the strings of each octet
Dim firstoct, secondoct, thirdoct, fourthoct As String
firstoct = Left(ipaddress, firstdot - 1)
secondoct = Mid(ipaddress, firstdot + 1, seconddot - firstdot - 1)
thirdoct = Mid(ipaddress, seconddot + 1, thirddot - seconddot - 1)
fourthoct = Right(ipaddress, Len(ipaddress) - thirddot)
'convert the strings to numbers
Dim dfirstoct, dsecondoct, dthirdoct, dfourthoct As Integer
dfirstoct = Val(firstoct)
dsecondoct = Val(secondoct)
dthirdoct = Val(thirdoct)
dfourthoct = Val(fourthoct)
'calculate & return the result
IP2DEC = (dfirstoct * 2 ^ 24) + (dsecondoct * 2 ^ 16) + (dthirdoct * 2 ^ 8) + dfourthoct
End Function
```

To save the function, go to `File>>Save NameOfYourFile.xlam`.

Now close the VB window to get back to your blank spreadsheet.

Call your function. In my case, in B2, I put `=IP2DEC(A1)`. This gave me the exact same value as what is in B1 but using a largely simplified formula.

Of course, if someone else opens the sheet without the add-in, they'll get a #Name error. Then you can show your Excel prowess by emailing them your add in. 

You can follow the same procedure to add in other custom formulas. For example, I've always been annoyed that Excel doesn't have an `IfBlank()` function. Essentially, I want the equivalent of `=If(IsBlank(A1),"something if blank","something if not blank")`. This is the function I created:

The following function could be used to evaluate whether a cell is blank or not and return one of two values:
``` VBA
Public Function IfBlank(rCell As Range, whatiftrue, whatiffalse)
'
'This function checks to see if a cell is blank. If it is, whatiftrue is returned. If not, whatiffalse is returned.
'
If rCell = "" Then IfBlank = whatiftrue Else IfBlank = whatiffalse
End Function

The following function could be used to turn a binary number back into an IP address:
Public Function DEC2IP(ByVal LongIP As Double) As String
'
'This function returns an IP address from a long (integer) IP address
'
Dim i As Integer, num As Double
For i = 1 To 4
num = Int(LongIP / 256 ^ (4 - i))
LongIP = LongIP - (num * 256 ^ (4 - i))
If i = 1 Then 'if that's all return the result
DEC2IP = num
Else 'if that's not all add the current octet to the result and continue
DEC2IP = DEC2IP & "." & num
End If
Next
End Function
```
I then found myself in a situation where I needed to do the reverse, converting a decimal number to an IP address.  So, there are two ways of doing this.

The first involves using a big formula to chop the decimal value into its equivalent dotted decimal counter parts. The formula goes like this (this formula references cell B2 where the decimal format of an IP address should be):

```
=ROUNDDOWN(B2/2^24,0)&"."&ROUNDDOWN(MOD(B2,2^24)/2^16,0)&"."&ROUNDDOWN(MOD(MOD(B2,2^24),2^16)/2^8,0)&"."&MOD(MOD(MOD(B2,2^24),2^16),2^8)
```

While this is nice, it would be even nicer if I could just do something like this:

```
=IPFromDEC(B2)
```

This is the second method.  If you've already created your IP2DEC.xlam file and have it enabled as an add in, you're ready to go, you can add a custom formula to break the IP address back out into the same add-in. 

Open a blank workbook in Excel.  Press `Alt+F11` or click 'Visual Basic' on the Developer tab in the ribbon bar.  If the project explorer isn't visible, show it by pressing `Ctrl+R` or by choosing `View>>Project Explorer`.  You should see two projects in there, one for the new blank workbook that opened and one for the `IP2DEC` add-in.  You should see Module1 under the IP2DEC add-in (if you don't see this, you didn't do the steps listed above).  Double click it.  You should now see the IP2DEC public function code.  Now all you need to do append some code to the bottom of the module that will define the function for converting back to dotted decimal format.

```
Public Function IPFROMDEC(ipaddress) As String
If ipaddress + 1 - 1 > 4294967295# Then GoTo toobig
Dim firstoctet As String, secondoctet As String, thirdoctet As String, fourthoctet As String
firstoctet = Int(ipaddress / (2 ^ 24))
secondoctet = ipaddress - (firstoctet * 2 ^ 24)
secondoctet = Int(secondoctet / (2 ^ 16))
thirdoctet = ipaddress - (firstoctet * 2 ^ 24) - (secondoctet * 2 ^ 16)
thirdoctet = Int(thirdoctet / (2 ^ 8))
fourthoctet = ipaddress - (firstoctet * 2 ^ 24) - (secondoctet * 2 ^ 16) - (thirdoctet * 2 ^ 8)
fourthoctet = Int(fourthoctet)
Select Case 255
Case Is < firstoctet
GoTo toobig
Case Is < secondoctet
GoTo toobig
Case Is < thirdoctet
GoTo toobig
Case Is < fourthoctet
GoTo toobig
End Select
IPFROMDEC = firstoctet & "." & secondoctet & "." & thirdoctet & "." & fourthoctet
Exit Function
toobig:
IPFROMDEC = "Invalid IP Address"
End Function
```

Hit the save button and go try it out.  Put an IP address in one cell and use `IP2DEC()` to convert it to decimal.  Then use `IPFROMDEC()` to convert it back. 

You might think this an exercise in futility, however, this can come in handy when trying to parse out IP address blocks given CIDR notation.  For example, if you wanted to calculate the starting and ending IP address for the 192.168.1.0/24 block of IP addresses, I'd have the following:

|   | A                                   | B             |
| - | ----------------------------------- | ------------- |
| 1	| 192.168.1.0	                        | 24            |
| 2 | =IPFROMDEC(IP2DEC(A1)+2^(32-B1)-1)	| 192.168.1.255 |

The result of the formula in A2 is shown in B2.  This can also be very handy when trying to determine whether or not a given IP address is within a given subnet.

## Download and adding this add-in
If you don't want to go through all the pain above, just [download my add-in](IPConversion.xlam) and [install it](http://office.microsoft.com/en-us/excel-help/add-or-remove-add-ins-HP010342658.aspx#BMexceladdin).

|Function|Description|Argument 1|Argument 2|192.168.15.34|10.20.30.40|335.20.30.40|142.20.30.40|
|---|---|---|---|---|---|---|---|
|IP2DEC|Converts from dotted decimal to decimal|||3232239394|169090600|Invalid IP Address|2383683112|
#VALUE!|
|IPNetwork|Returns the network number given an IP address and mask|24||192.168.15.0|10.20.30.0|Invalid IP Address|142.20.30.0|
|IPIsInSubnet|Determines if the given IP address is within the given subnet|192.168.0.0|16|TRUE|FALSE|FALSE|FALSE|
|IPGetOctet|Returns the octet specified (1-4)|3||15|30|Invalid IP Address|30|
|IPGetCIDRList|Returns a comma separated list of CIDR address blocks between the two provided IP addresses|192.168.255.255||192.168.15.34/31, 192.168.15.36/30, 192.168.15.40/29, 192.168.15.48/28, 192.168.15.64/26, 192.168.15.128/25, 192.168.16.0/20, 192.168.32.0/19, 192.168.64.0/18, 192.168.128.0/17|10.20.30.40/29, 10.20.30.48/28, 10.20.30.64/26, 10.20.30.128/25, 10.20.31.0/24, 10.20.32.0/19, 10.20.64.0/18, 10.20.128.0/17, 10.21.0.0/16, 10.22.0.0/15, 10.24.0.0/13, 10.32.0.0/11, 10.64.0.0/10, 10.128.0.0/9, 11.0.0.0/8, 12.0.0.0/6, 16.0.0.0/4, 32.0.0.0/3, 64.0.0.0/2, 128.0.0.0/2, 192.0.0.0/9, 192.128.0.0/11, 192.160.0.0/13, 192.168.0.0/16||142.20.30.40/29, 142.20.30.48/28, 142.20.30.64/26, 142.20.30.128/25, 142.20.31.0/24, 142.20.32.0/19, 142.20.64.0/18, 142.20.128.0/17, 142.21.0.0/16, 142.22.0.0/15, 142.24.0.0/13, 142.32.0.0/11, 142.64.0.0/10, 142.128.0.0/9, 143.0.0.0/8, 144.0.0.0/4, 160.0.0.0/3, 192.0.0.0/9, 192.128.0.0/11, 192.160.0.0/13, 192.168.0.0/16|
|IPIsValidIP|Returns true if the specified IP address is a valid IP address|||TRUE|TRUE|FALSE|TRUE|

### Examples
`IPGetCIDRList("10.0.0.0","10.0.0.5")` ==> `10.0.0.0/30,10.0.0.4/31`

`IPGetCIDRList("10.1.2.3","192.168.35.7")` ==> `10.1.2.3/32, 10.1.2.4/30, 10.1.2.8/29, 10.1.2.16/28, 10.1.2.32/27, 10.1.2.64/26, 10.1.2.128/25, 10.1.3.0/24, 10.1.4.0/22, 10.1.8.0/21, 10.1.16.0/20, 10.1.32.0/19, 10.1.64.0/18, 10.1.128.0/17, 10.2.0.0/15, 10.4.0.0/14, 10.8.0.0/13, 10.16.0.0/12, 10.32.0.0/11, 10.64.0.0/10, 10.128.0.0/9, 11.0.0.0/8, 12.0.0.0/6, 16.0.0.0/4, 32.0.0.0/3, 64.0.0.0/2, 128.0.0.0/2, 192.0.0.0/9, 192.128.0.0/11, 192.160.0.0/13, 192.168.0.0/19, 192.168.32.0/23, 192.168.34.0/24, 192.168.35.0/29`

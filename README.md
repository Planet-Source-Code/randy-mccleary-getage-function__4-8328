<div align="center">

## GetAGE Function


</div>

### Description

This is just a Simple little ASP Script Function that can be used to get the AGE of a person, object, anniversary, or whatever you would like to find out how old something is in YEARS.
 
### More Info
 
A validate Date

Notes: Calculation of a year is based upon

the most accurate formula which states

a Year = 365.242222 Days(Tropical Season).

An error of 1 day will occur every 40,000

years with this formula.

If the Julian formula is used then

A Year = 365.25 days, which means every

128 years an Error of 1 Day will occur.

If the Gregorian formula is used then

A Year = 365.2425, which is fairly accurate

But every 3200 years an Error of 1 day will

occur.

An Interger Value of the AGE in years, or a String "NULL" if an invalid date was input.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Randy McCleary](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/randy-mccleary.md)
**Level**          |Intermediate
**User Rating**    |5.0 (25 globes from 5 users)
**Compatibility**  |ASP \(Active Server Pages\)
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__4-1.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/randy-mccleary-getage-function__4-8328/archive/master.zip)





### Source Code

```
<% @ LANGUAGE = VBScript %>
<% Option Explicit %>
<%
'##############################################
'## Function Name: GetAGE()
'##
'## Description: This function will take an
'## input date and calculate the age in years
'## of how old someone is or age of an item
'## to the current Date.
'##
'## Inputs: A valid date ex: 04/25/2000
'##
'## Returns: Integer - AGE in YEARS
'##     NULL If invalid input
'##
'## Notes: Calculation of a year is based upon
'##  the most accurate formula which states
'##  a Year = 365.242222 Days(Tropical Season).
'##  An error of 1 day will occur every 40,000
'##  years with this formula.
'##
'## If the Julian formula is used then
'##  A Year = 365.25 days, which means every
'##  128 years an Error of 1 Day will occur.
'##
'## If the Gregorian formula is used then
'##  A Year = 365.2425, which is fairly accurate
'##  But every 3200 years an Error of 1 day will
'##  occur.
'##############################################
Function GetAGE(dtmDate)
	Dim intAGE
	If IsDate(dtmDate) <> True Then
		GetAGE = "NULL"
		Exit Function
	End If
	intAGE = DateDiff("d", dtmDate, Date)
	Response.Write "AGE is " & intAGE & "<br>"
	If IsNumeric(intAGE) Then
		intAGE = Int(Round(Abs(intAGE) / 365.242222, 4))
	End If
	GetAGE = intAGE
End Function
'----------------------------------
Dim dtmDate
dtmDate = "4/24/1875"
Response.Write "AGE is " & GetAGE(dtmDate) & "<br>"
%>
```


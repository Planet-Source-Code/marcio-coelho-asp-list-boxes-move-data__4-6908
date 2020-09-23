<div align="center">

## ASP List Boxes Move Data


</div>

### Description

couple of ways of moving data between two list boxes.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Marcio Coelho](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/marcio-coelho.md)
**Level**          |Intermediate
**User Rating**    |4.4 (22 globes from 5 users)
**Compatibility**  |ASP \(Active Server Pages\), HTML
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__4-1.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/marcio-coelho-asp-list-boxes-move-data__4-6908/archive/master.zip)

### API Declarations

couple of ways of moving data between two list boxes.


### Source Code

```
<%@ Language=VBScript %>
<html>
<head>
<STYLE>
	.Over
	{
	  BACKGROUND-COLOR: #708090;
	  COLOR: #ffffff;
	  CURSOR: hand
	}
	.Out
	{
	  BACKGROUND-COLOR: #c0c0c0;
	  COLOR: #000000;
	  CURSOR: default
	}
</STYLE>
</head>
<script FOR="document" EVENT="onmouseover" LANGUAGE="javascript">
<!--
document_onmouseover()
//-->
</script>
<script LANGUAGE="javascript" FOR="document" EVENT="onmouseout">
<!--
 document_onmouseout()
//-->
</script>
<script language="Javascript">
<!--
function document_onmouseout()
	{
		if (top.event.srcElement.tagName == "BUTTON")
			{
			top.event.srcElement.className = "Out";
			}
	}
	function document_onmouseover()
	{
		if (top.event.srcElement.tagName == "BUTTON")
			{
			top.event.srcElement.className = "Over";
			}
 }
	//-->
</script>
<script language="Javascript">
function addtocombo()
		{
				var i = 0
				var selectedItem
				var selectedText
				var selectedValue
				var newoption1
				for (counter = 0; counter < userslist.menuitems.length; counter++)
				{
						if (userslist.menuitems.options(counter).selected == 1)
						{
						 selectedItem = document.userslist.menuitems.selectedIndex;
						 selectedText = document.userslist.menuitems.options[counter].text;
						 selectedValue = document.userslist.menuitems.options[counter].value;
						 newoption1 = new Option(selectedText, selectedValue, false, false);
						 document.userslist.selitems.options[i] = newoption1;
						 i = i + 1
		 			}
				}
		}
function moveit()
{
		var boxLength = document.userslist.selitems1.length;
		var selectedItem = document.userslist.menuitems1.selectedIndex;
		var selectedText = document.userslist.menuitems1.options[selectedItem].text;
		var selectedValue = document.userslist.menuitems1.options[selectedItem].value;
		var i;
		var isNew = true;
		if (boxLength != 0)
		 {
				for (i = 0; i < boxLength; i++)
					 {
						thisitem = document.userslist.selitems1.options[i].text;
						if (thisitem == selectedText)
							{
									isNew = false;
									break;
							}
					}
		}
		if (isNew)
			 {
				newoption = new Option(selectedText, selectedValue, false, false);
				document.userslist.selitems1.options[boxLength] = newoption;
				}
				document.userslist.menuitems1.selectedIndex=-1;
}
function removefromcombo()
{
		var boxLength = document.userslist.selitems1.length;
		arrSelected = new Array();
		var count = 0;
		for (i = 0; i < boxLength; i++)
		 {
				if (document.userslist.selitems1.options[i].selected)
				 {
					arrSelected[count] = document.userslist.selitems1.options[i].value;
					}
				count++;
		}
		var x;
		for (i = 0; i < boxLength; i++)
		{
				for (x = 0; x < arrSelected.length; x++)
				{
						if (document.userslist.selitems1.options[i].value == arrSelected[x])
						 {
							document.userslist.selitems1.options[i] = null;
							}
				}
				boxLength = document.userslist.selitems1.length;
		}
}
function removefromcombo1()
{
		var boxLength = document.userslist.selitems.length;
		arrSelected = new Array();
		var count = 0;
		for (i = 0; i < boxLength; i++)
		 {
				if (document.userslist.selitems.options[i].selected)
				 {
					arrSelected[count] = document.userslist.selitems.options[i].value;
					}
				count++;
		}
		var x;
		for (i = 0; i < boxLength; i++)
		{
				for (x = 0; x < arrSelected.length; x++)
				{
						if (document.userslist.selitems.options[i].value == arrSelected[x])
						 {
							document.userslist.selitems.options[i] = null;
							}
				}
				boxLength = document.userslist.selitems.length;
		}
}
		//-->
</script>
<body >
	<center>
	<b><font face="tahoma" size="2">Menu Item Permissions - Development Options</b></FONT></center><hr>
	<form name="userslist" method="post">
			<CENTER>
			<TABLE border=1>
			<TR><TD>
			<table>
				<tr>
				<td bgcolor="#cccccc" align="middle"><b>Available Menu Items</b></td>
				<td></td>
					<td bgcolor="#cccccc" align="middle"><b>User's Current Menu Items</b></td>
				</tr>
				<tr>
					<td align="left"><!-- onchange="moveit();" -->
						<select name="menuitems" style="WIDTH: 300px" size="10" multiple>
						<option value="210">(USA)Domain Users - Demographics Report </option>
						<option value="211">(USA)Domain Users - Loss Prevention </option>
						<option value="212">(USA) Domain Users - Tracker Admin Guide </option>
						<option value="213">(USA)Domain Users - Traffic </option>
						<option value="214">(USA)Domain Users - Loss Fix </option>
						<option value="215">(USA) Domain Users - Testing </option>
						<option value="216">(USA)Domain Users - Play Report </option>
						<option value="217">(USA)Domain Users - Check it out </option>
						<option value="218"> Remote Scripting</option>
						<option value="219"> E-Metrics is cool !</option>
						<option value="220"> Check Bulk Insert</option>
						<option value="221"> Check ActiveX Script</option>
						<option value="222"> VBS is more than popular</option>
						<option value="223"> Great Move</option>
						</select>
					</td>
					<td align="middle">
					<input type="button" tagName="BUTTON" onclick="addtocombo();" value="&nbsp;&nbsp;&nbsp;&nbsp;Add > >&nbsp;&nbsp;&nbsp;&nbsp;" id="button1" name="button1" style="FONT-WEIGHT: bold"><br><br>
					<input type="button" tagName="BUTTON" onclick="removefromcombo1();" value=" < < Remove" id="submit1" name="submit1" style="FONT-WEIGHT: bold">
					</td>
					<td align="left">
						<select name="selitems" id="selitems" style="WIDTH: 300px" size="10" multiple>
						<option></option>
						</select>
				</td>
			</tr>
			</table><br><br><BR>
			<HR><BR><BR>
			<table>
				<tr>
				<td bgcolor="#cccccc" align="middle" ><b>Available Menu Items</b></td>
				<td></td>
					<td bgcolor="#cccccc" align="middle" ><b>User's Current Menu Items</b></td>
				</tr>
				<tr>
					<td align="left"><!-- onchange="moveit();" -->
						<select name="menuitems1" onchange="moveit();" style="WIDTH: 300px" size="10" multiple>
						<option value="210">(USA)Domain Users - Demographics Report </option>
						<option value="211">(USA)Domain Users - Loss Prevention </option>
						<option value="212">(USA) Domain Users - Tracker Admin Guide </option>
						<option value="213">(USA)Domain Users - Traffic </option>
						<option value="214">(USA)Domain Users - Loss Fix </option>
						<option value="215">(USA) Domain Users - Testing </option>
						<option value="216">(USA)Domain Users - Play Report </option>
						<option value="217">(USA)Domain Users - Check it out </option>
						<option value="218"> Remote Scripting</option>
						<option value="219"> E-Metrics is cool !</option>
						<option value="220"> Check Bulk Insert</option>
						<option value="221"> Check ActiveX Script</option>
						<option value="222"> VBS is more than popular</option>
						<option value="223"> Great Move</option>
						</select>
					</td>
					<td align="middle">
					<input type="button" tagName="BUTTON" onclick="removefromcombo();" value=" < < Remove" id="submit2" name="submit2" style="FONT-WEIGHT: bold">
					</td>
					<td align="left">
						<select name="selitems1" id="selitems1" style="WIDTH: 300px" size="10" multiple>
						</select>
				</td>
			</tr>
			</table><br></TD></TR></TABLE></CENTER>
			</form>
</body>
</html>
```


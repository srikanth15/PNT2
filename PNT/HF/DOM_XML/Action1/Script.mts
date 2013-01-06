'DOM

'.object

'InnerText-GetROProperty

'XML Explore the Power of DOM

'MsgBox Browser(":: PeopleNTech, LLC. ::").Page(":: PeopleNTech, LLC. ::").Object.all.tags("DIV").Length

'Read From XML

'Access XML Class
'Load the XML
'Navigate within XML


Set objXML=CreateObject("Microsoft.XMLDOM")
objXML.Load("C:\Users\Srikanth\Desktop\PNT\test1.xml")

'Set objnodes= objXML.SelectNodes("/car/Make")
'MsgBox objNodes.Length


'Read Data from XML




'Set objnode=objXML.SelectSingleNode("/car/Make[1]/Name/text()")
'MsgBox objnode.NodeValue

'Capture from HTMl Page


'
'Set objnode=objXML.SelectSingleNode("/car/Make[0]/Name")
'objNode.Text=Browser("Mortgage Calculator").Page("Mortgage Calculator").WebElement("$1,654.55").Object.innerText

'Create New Element

'/car/Make- Parent Node

'Child Node- Now we are going to add it

Set objpnode=objXML.SelectSingleNode("/car/Make[0]")

'Create a new child node

Set objcnode=objXML.CreateElement("Value")
objcnode.Text="USD 20,000"

objpnode.AppendChild(objcnode)



'Replace the Excel data source with XML File


'Edit our Driver Script to work with our Data Source


'srikanth@peoplentech.com





















'Save the Changes

objXML.Save("C:\Users\Srikanth\Desktop\PNT\test4.xls")












































﻿Imports DevExpress.XtraRichEdit
Imports System.Xml
Imports DevExpress.XtraRichEdit.API.Native

Namespace DXRichEditControlAPISample.CodeExamples
	Friend Class CustomXmlParts
		Private Shared Sub AddCustomXmlPart(ByVal document As Document)
'			#Region "#AddCustomXmlPart"
			document.AppendText("This document contains custom XML parts.")
			' Add an empty custom XML part.
			Dim xmlItem As ICustomXmlPart = document.CustomXmlParts.Add()
			' Populate the XML part with content.
			Dim elem As XmlElement = xmlItem.CustomXmlPartDocument.CreateElement("Employees")
			elem.InnerText = "Stephen Edwards"
			xmlItem.CustomXmlPartDocument.AppendChild(elem)

			' Use a string to specify the content for a custom XML part.
			Dim xmlString As String = "<?xml version=""1.0"" encoding=""UTF-8""?>
                            <Employees>
                                <FirstName>Stephen</FirstName>
                                <LastName>Edwards</LastName>
                                <Address>4726 - 11th Ave. N.E.</Address>
                                <City>Seattle</City>
                                <Region>WA</Region>
                                <PostalCode>98122</PostalCode>
                                <Country>USA</Country>
                            </Employees>"
			document.CustomXmlParts.Insert(1, xmlString)

			' Add a custom XML part from a file.
			Dim xmlDoc As New XmlDocument()
			xmlDoc.Load("Documents\Employees.xml")
			document.CustomXmlParts.Add(xmlDoc)
			document.SaveDocument("Result.docx", DocumentFormat.OpenXml)
			System.Diagnostics.Process.Start("explorer.exe", "/select," & "Result.docx")
'			#End Region ' #AddCustomXmlPart
		End Sub

		Private Shared Sub AccessCustomXmlPart(ByVal document As Document)
'			#Region "#AccessCustomXmlPart"
			' Load a document.
			document.LoadDocument("Documents\CustomXmlParts.docx")
			' Access a custom XML file stored in the document.
			Dim xmlDoc As XmlDocument = document.CustomXmlParts(0).CustomXmlPartDocument
			' Retrieve employee names from the XML file and display them in the document.
			Dim nameList As XmlNodeList = xmlDoc.GetElementsByTagName("Name")
			document.AppendText("Employee list:")
			For Each name As XmlNode In nameList
				document.AppendText(vbCrLf & " " & ChrW(&H00B7).ToString() & " " & name.InnerText)
			Next name
			document.SaveDocument("Result.docx", DocumentFormat.OpenXml)
			System.Diagnostics.Process.Start("explorer.exe", "/select," & "Result.docx")

'			#End Region ' #AccessCustomXmlPart
		End Sub

		Private Shared Sub RemoveCustomXmlPart(ByVal document As Document)
'			#Region "#RemoveCustomXmlPart"
			document.AppendText("This document contains custom XML parts.")

			' Add the first custom XML part.
			Dim xmlString1 As String = "<?xml version=""1.0"" encoding=""UTF-8""?>
                            <Employees>
                                <FirstName>Stephen</FirstName>
                                <LastName>Edwards</LastName>
                            </Employees>"
			Dim xmlItem1 = document.CustomXmlParts.Add(xmlString1)

			' Add the second custom XML part.
			Dim xmlString2 As String = "<?xml version=""1.0"" encoding=""UTF-8""?>
                            <Employees>
                                <FirstName>Andrew</FirstName>
                                <LastName>Fuller</LastName>
                            </Employees>"
			Dim xmlItem2 = document.CustomXmlParts.Add(xmlString2)

			' Remove the first item from the collection.
			document.CustomXmlParts.Remove(xmlItem1)
			' Use the RemoveAt method to remove an item at the specified position from the collection.
			' document.CustomXmlParts.RemoveAt(0);
			' Use the Clear method to remove all items from the collection.
			' document.CustomXmlParts.Clear();
			document.SaveDocument("Result.docx", DocumentFormat.OpenXml)
			System.Diagnostics.Process.Start("explorer.exe", "/select," & "Result.docx")
'			#End Region ' #RemoveCustomXmlPart
		End Sub
	End Class
End Namespace

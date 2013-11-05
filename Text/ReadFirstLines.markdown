##Read First Lines

The first step when dealing with huge XML files is to get an idea about how the XML document looks like. 
This is a pretty standard technique when handling large datasets, and other languages like **R** have built-in functions for that purpose.
In Access we can use the *File System Object* (**FSO**) for accomplishing the same results. 
The function [FirstLines](../Functions/FirstLines.vb) shows how to use FSO for reading the first lines of a file, and produce the results shown below:

###Usage

	Private Sub Command0_Click()
		Dim strPath As String
		Dim strFile As String
		
		strPath = "C:\MyPath"
		strFile = "supp2014.xml"
		
		Debug.Print FirstLines(strPath, strFile, 11)
	End Sub

###Results

    <?xml version="1.0"?>
    <dataroot xmlns:od='urn:schemas-microsoft-com:officedata' 
    xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' 
    xsi:noNamespaceSchemaLocation='C:\MyPath\MyXSDFile.xsd'>
    <SupplementalRecordSet LanguageCode = "eng">
    <SupplementalRecord SCRClass = "1">
     <SupplementalRecordUI>C000002</SupplementalRecordUI>
     <SupplementalRecordName>
      <String>bevonium</String>
     </SupplementalRecordName>
     <DateCreated>
      <Year>1971</Year>
      <Month>01</Month>


* [Go Back to Index](Index.markdown)
* [Next: Read Last Lines](ReadLastLines.markdown)

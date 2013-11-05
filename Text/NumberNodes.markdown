##Number of Nodes

An important data to know is the number of nodes we are interested in for further manipulation of the data. 
In this case, knowing how many **SupplementalRecord** nodes are in the document is of particular importance.
For counting the number of times an element appears in a document we can use two different approaches:
* SAX
* FSO

**SAX** is ~3 times faster than **FSO** for this purpose but **FSO** is simpler to implement.

###SAX
For using SAX the classes [clsSaxContentHandlerCountNodes](../Functions/clsSaxContentHandlerCountNodes.vb),
[clsSaxErrorHandler](../Functions/clsSaxErrorHandler.vb), 
and the module [SaxCountNodes](../Functions/SaxCountNodes.vb) are required.

####Usage

	Private Sub Command0_Click()
	    Dim strPath     As String
	    Dim strFile     As String
	    
	    strPath = "C:\MyPath"
	    strFile = "supp2014.xml"
	    
	    SaxCountNodes strPath, strFile
	End Sub


###FSO
The function [CountNodes](../Functions/CountNodes.vb) implements FSO.

####Usage

	Private Sub Command0_Click()
		Dim strPath     As String
		Dim strFile     As String
		Dim strNode     As String
		Dim lngNodes    As Long
		
		strPath = "C:\MyPath"
		strFile = "supp2014.xml"
		strNode = "SupplementalRecord"
		
		lngNodes = CountNodes(strPath, strFile, strNode)
		Debug.Print "There are " & lngNodes & " " & strNode & " Nodes"
		
	End Sub

###Results
**SAX**

	Start Counting SAX:           	10/29/2013 7:45:55 PM 
	End Counting SAX:           	10/29/2013 7:46:57 PM 
	There are 219278 SupplementalRecord Nodes
	
**FSO**

	Start Counting FSO:           	10/29/2013 7:47:13 PM 
	End Counting FSO:           	10/29/2013 7:50:09 PM 
	There are 219278 SupplementalRecord Nodes

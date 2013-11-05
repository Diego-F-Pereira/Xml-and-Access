##Splitting the XML

Knowing the number of nodes of the node we are interested in, allow us to calculate the number of files and the number of nodes per file.
For splitting a large XML into several smaller documents the function [SplitXml](../Functions/SplitXml.vb) can be used.

###Usage

	Private Sub Command0_Click()
		Dim strPath         As String
		Dim strInputFile    As String
		Dim strOutputName   As String
		Dim strNode         As String
		Dim lngNodes        As Long
		
		Debug.Print Now()
		
		strPath = "C:\MyPath"
		strInputFile = "supp2014.xml"
		strOutputName = "supp2014_"
		strNode = "SupplementalRecord"
		lngNodes = 10000
		
		SplitXml strPath, strInputFile, strOutputName, strNode, lngNodes

		Debug.Print Now()
	End Sub
	
###Results

	10/25/2013 10:58:59 PM
	0            10/25/2013 10:58:59 PM
	1            10/25/2013 10:59:20 PM
	2            10/25/2013 10:59:42 PM
	3            10/25/2013 11:00:04 PM
	4            10/25/2013 11:00:24 PM
	5            10/25/2013 11:00:43 PM
	6            10/25/2013 11:01:02 PM
	7            10/25/2013 11:01:21 PM
	8            10/25/2013 11:01:40 PM
	9            10/25/2013 11:02:00 PM
	10           10/25/2013 11:02:18 PM
	11           10/25/2013 11:02:36 PM
	12           10/25/2013 11:02:53 PM
	13           10/25/2013 11:03:10 PM
	14           10/25/2013 11:03:26 PM
	15           10/25/2013 11:03:42 PM
	16           10/25/2013 11:03:58 PM
	17           10/25/2013 11:04:13 PM
	18           10/25/2013 11:04:27 PM
	19           10/25/2013 11:04:42 PM
	20           10/25/2013 11:04:56 PM
	21           10/25/2013 11:05:10 PM
	10/25/2013 11:05:22 PM	
	
[Go back to Index](Index.markdown)	

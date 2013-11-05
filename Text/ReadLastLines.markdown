##Read Last Lines

Additionally, one might want to read the last lines of the XML. 
That can be done by using the [LastLines](../Functions/LastLines.vb) function

###Usage

	Private Sub Command0_Click()
		Dim strPath As String
		Dim strFile As String
		
		strPath = "C:\MyPath"
		strFile = "supp2014.xml"
		
		Debug.Print LastLines(strPath, strFile, 11)
	End Sub
	
###Results	

         </DateCreated>
         <ThesaurusIDlist>
          <ThesaurusID>NLM (2013)</ThesaurusID>
         </ThesaurusIDlist>
        </Term>
       </TermList>
      </Concept>
     </ConceptList>
    </SupplementalRecord>
    </SupplementalRecordSet>
    </dataroot>
    
* [Go back to Index](Index.markdown)
* [Previous: Read First Lines](ReadFirstLines.markdown)
* [Next: Getting Insight of the Structure](GettingInsightStructure.markdown)

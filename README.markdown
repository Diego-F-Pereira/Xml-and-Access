#Generalities

Working with large **XML files** is a challenge for the main following reasons:

* Using the DOM API, **RAM** requirements are **~5 times** the size of the file.
* Most text editors are **unable** to open huge XML files. 

That means that most applications will crash when loading the file or when representing the tree. 
Therefore, streaming techniques are the choice in this situation.

Three streaming techniques were tested for writing down this article:

* FSO
* ADO
* SAX 

Each one has its strengths and weaknesses, and the developer will probably need to use one or another for different tasks in the context of very large XML files. 

All the examples in this article are based upon the *Medical Subject Headings* (**MeSH**) database, property of the *U.S. National Library of Medicine* (**NLM**) publicly available from http://www.nlm.nih.gov/mesh/.
This XML file is **~700Mb** and serves perfectly for illustration purposes.

[Go to Index](Text/Index.markdown)









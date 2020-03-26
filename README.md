# OpenXML dll *"simplified"*
## Dovetta Nicolas, 4^B Informatica - I.I.S. "G. Vallauri" Fossano

Simplifing usig of most commonmetod to create an *__Excel__* or an *__Word__* doument.

### How to use
1. Add to your __C#__ project the references, also the using, to "OpenXmlPersonalized_dll.dll";
2. Create the file path you prefer, the default is in the bin/debug folder. You can use the createPath from
   the **"OpenXMLUtilities_common"** class (the default filename is "Output", if the filename does also exist it automaticaly add a datetime specific for that file);
3. # *Word*
    * Use CreateParagraphWithStyle method from the class **"OpenXMLUtilities_word"** to add a paragraph to your document;
    * Use AddTextToParagraph method to add text to your paragraph;
    * Use CreateBulletNumberingPart or CreateBulletOrNumberedList method to create ordered/in ordered bullet list;
    * Use InsertParagraphInList to insert paragraph in your list;
    * Use createTable method to create a table to add at your document;
    * Use InsertPicture method to add picture at your document;
4. # *Excel*
    * Use CreatePartsForExcel method from the class **"OpenXMLUtilities_excel"** to add parts at your document;
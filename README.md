# Royale-OfficeAddIn

A library for Apache Royale Framework (https://royale.apache.org) that wraps the Office JavaScript API library (https://docs.microsoft.com/en-us/office/dev/add-ins/develop/understanding-the-javascript-api-for-office). Office JavaScript API, designed on JavaScript and HTML5 technologies, is a robust set of components for interacting with Office. *Note*: at the time of writing this document, Office JavaScript API can be used with Excel, Outlook, Word, PowerPoint and OneNote however this Royale-OfficeAddIn was tested only with Word.

You need to manually insert the reference to office.js on your Royale template before ${head} (currently I didn't find a better solution)
````html
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
`````

Usage (insert text)
````actionscript
OfficeAddIn.insertText("Hello World");
`````

Usage (insert html text)
````actionscript
OfficeAddIn.insertHtml("<b>Hello</b><br>World");
`````

Usage (insert table)
````actionscript
OfficeAddIn.insertTable([['Lisbon', 1], ['Munich', 2], ['Duisburg', 3]]);
`````

Usage (insert image)
````actionscript
var base64Image:String = "..." //replace "..." with a base64 image
OfficeAddIn.insertImage(base64Image);
`````

Usage (find and replace)
````actionscript
OfficeAddIn.findAndReplace("World", "Planet");
`````

Usage (get the document as PDF/A byte array)
````actionscript
OfficeAddIn.getDocumentAsPDF().then(function(result:ByteArray):IThenable 
{ 
});
`````

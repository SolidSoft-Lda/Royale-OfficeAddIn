# Royale-OfficeAddIn

A library for the Apache Royale Framework (https://royale.apache.org) that wraps the Office JavaScript API library (https://docs.microsoft.com/en-us/office/dev/add-ins/develop/understanding-the-javascript-api-for-office). The Office JavaScript API, designed on JavaScript and HTML5 technologies, is a robust set of components for interacting with Microsoft Office. *Note*: at the time of writing this document, the Office JavaScript API can be used with Excel, Outlook, Word, PowerPoint and OneNote however this Royale-OfficeAddIn was tested only with Word.

Before starting, make sure that you read the Office Web Add-In documentation and know how to configure it thru a manifest file (https://docs.microsoft.com/en-us/office/dev/add-ins/develop/add-in-manifests).
It's not the intention of this repository to explain how the Office Web Add-In works.

You need to manually insert a reference to office.js in your Royale template before ${head} (I have not found a better solution yet).
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

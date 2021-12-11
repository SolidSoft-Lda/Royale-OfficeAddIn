var documentSnapshot;

window.onload = function()
{
    Office.initialize = function () { };
}

class OfficeAddIn
{
    static getDefaultLanguage()
    {
        return Office.context.displayLanguage;
    }

    static insertText(text)
    {
        Office.context.document.setSelectedDataAsync(text);
    }

    static insertHtml(html)
    {
        Office.context.document.setSelectedDataAsync(html, { coercionType: Office.CoercionType.Html });
    }

    static insertTable(table)
    {
        Office.context.document.setSelectedDataAsync(table, { coercionType: Office.CoercionType.Matrix });
    }

    static insertImage(base64Image)
    {
        Word.run(function (context)
        {
            context.document.body.insertInlinePictureFromBase64(base64Image, "End");
            return context.sync();
        })
    }

    static existText(toFind)
    { 
        return new Promise((resolve) =>
        {
            Word.run(function (context)
            {
                var results = context.document.body.search(toFind);
                context.load(results, "no-properties-needed");
             
                return context.sync().then(function ()
                {
                    resolve(results.items.length > 0 ? toFind : null);
                });
            });
        });
    }

    static async saveSnapshot()
    {
        await Word.run(function (context)
        {
            var bodyOOXML = context.document.body.getOoxml();
            return context.sync().then(function ()
            {
                documentSnapshot = bodyOOXML.value;
            });
        });
    }

    static async restoreSnapshot()
    {
        if (documentSnapshot != null)
        {
            await Word.run(function (context)
            {
                var body = context.document.body;
                body.clear();
                body.insertOoxml(documentSnapshot, Word.InsertLocation.start);
                return context.sync();
            });
        }
    }

    static findAndReplace(toFind, toReplace)
    {
        return new Promise((resolve) =>
        {
            Word.run(function (context)
            {
                var results = context.document.body.search(toFind);
                context.load(results, "no-properties-needed");
             
                return context.sync().then(function ()
                {
                    for (var i = 0; i < results.items.length; i++)
                    {
                        if (toReplace == null || toReplace == "")
                            results.items[i].insertHtml(" ", "replace");
                        else
                            results.items[i].insertHtml(toReplace, "replace");
                    }
                })
                .then(function ()
                {
                    context.sync();
                    resolve(true);
                });
            });
        });
    }

    static getDocumentAsPDF()
    {
        return new Promise((resolve, reject) =>
        {
            Office.context.document.getFilePropertiesAsync(function (asyncResult)
            {
                Office.context.document.getFileAsync(Office.FileType.Pdf, function (result)
                {
                    if (result.status != Office.AsyncResultStatus.Succeeded)
                        reject(result.error);

                    var state =
                    {
                        file: result.value,
                        counter: 0,
                        sliceCount: result.value.sliceCount
                    };
                    state.file.getSliceAsync(state.counter, function (result)
                    {
                        if (result.status == Office.AsyncResultStatus.Succeeded)
                        {
                            if (result.value.data)
                            {
                                var base64String = "";
                                for (var i = 0; i < result.value.data.length; i++)
                                {
                                    base64String += String.fromCharCode(result.value.data[i]);
                                }
                                resolve(btoa(base64String));
                            }

                            state.counter++;
                            if (state.counter <= state.sliceCount)
                                state.file.closeAsync();
                        }    
                    });
                });
            });
        });
    }
}
window.onload = function()
{
    Office.initialize = function () { };
}

var findAndReplaceStackTrace = [];

class OfficeAddIn
{
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

    static findAndReplace(toFind, toReplace)
    {
        Word.run(function (context)
        {
            var results = context.document.body.search(toFind);
            context.load(results);
         
            return context.sync().then(function ()
            {
                for (var i = 0; i < results.items.length; i++)
                {
                    findAndReplaceStackTrace.push([toFind, toReplace]);
                    results.items[i].insertHtml(toReplace, "replace");
                }
            })
            .then(context.sync);
        });
    }

    static undoFindAndReplace()
    {
        Word.run(function (context)
        {
            var item = findAndReplaceStackTrace.shift();
            var results = context.document.body.search(item[1]);
            context.load(results);
         
            return context.sync().then(function ()
            {
                results.items[0].insertHtml(item[0], "replace");

                if (findAndReplaceStackTrace.length > 0)
                    OfficeAddIn.undoFindAndReplace();
            })
            .then(context.sync);
        });
    }

    static getDocumentAsPDF()
    {
        return new Promise((resolve, reject) =>
        {
            Office.context.document.getFilePropertiesAsync(function (asyncResult)
            {
                if (!asyncResult.value.url)
                    reject("FILE_NOT_SAVED");

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
                                resolve(result.value.data);

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
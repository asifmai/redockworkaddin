/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */


// The initialize function must be run each time a new page is loaded
Office.initialize = (reason) => {
    $(document).ready(async () => {
        if (!Office.context.requirements.isSetSupported('WordApi', 1.3)) {
            console.log('Sorry. The add-in uses Word.js APIs version 1.3, that are not available in your version of Office.');
        };

        $('#run').click(run);
        $('#input-search').focus();
        await copyText();
    });
};

async function copyText() {
    return Word.run(async context => {
        const range = context.document.getSelection();
        range.load('text');
        await context.sync();
        var selectedText = range.text.trim();
        if (selectedText != '') {
            $('#input-search').val(selectedText);
        }
    })
}

async function run() {
    return Word.run(async context => {
        var searchTerm = $('#input-search').val().trim();
        var searchURL = 'https://app.redock.com/?q=' + searchTerm
        if (searchTerm != '') {
            window.open(searchURL);
        }
    });
}
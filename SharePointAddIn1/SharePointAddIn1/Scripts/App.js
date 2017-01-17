'use strict';

ExecuteOrDelayUntilScriptLoaded(initializePage, "sp.js");

function initializePage() {

    // limit number of file links we output
    var MAX_FILES = 10;

    // create host-web url without the add-in part at the end
    var hostWeb = _spPageContextInfo.webAbsoluteUrl.split('/');
    hostWeb.pop();
    hostWeb = hostWeb.join('/');

    // button bindings
    $('#test_query').off('click').on('click', init);
    $('#test_reopen').off('click').on('click', reopen);

    // run tests when called
    function init() {
        log('<em>Loading...</em>');

        // query files in the host-web document libraries
        http({
            url: api('Web/Lists', {
                '$filter': 'Hidden eq false and BaseType eq 1 and BaseTemplate eq 101',
                '$select': 'RootFolder/Exists, RootFolder/Files/Exists, RootFolder/Files/Name, RootFolder/Files/LinkingUrl, Items/Folder/Exists, Items/Folder/Files/Exists, Items/Folder/Files/Name, Items/Folder/Files/LinkingUrl',
                '$expand': 'RootFolder/Files, Items/Folder/Files'
            })
        }).done(function (data) {
            var numFiles = 0;

            // filter to lists that contain items or rootfolder files
            var lists = data.d.results.filter(filterList);

            // iterate lists
            for (var i = 0; i < lists.length; i++) {
                var list = lists[i];

                // filter rootfolder results if the files exist and have linking url properties
                var rootFiles = list.RootFolder.Files.results.filter(filterFolderFile);

                // iterate rootfolder files
                for (var j = 0; j < rootFiles.length; j++) {
                    var file = rootFiles[j];
                    // append file in the log
                    log($('<a/>').attr({ src: file.LinkingUrl, target: '_blank' }).text(file.Name));
                    // sanity check
                    if (++numFiles >= MAX_FILES) return;
                }

                // filter list items that contain files
                var items = list.Items.results.filter(filterListItem);

                // iterate list items
                for (var j = 0; j < items.length; j++) {
                    var item = items[j];

                    // filter item folders that contain files and have linking url properties
                    var files = item.Folder.Files.results.filter(filterFolderFile);

                    // iterate item files
                    for (var k = 0; k < files.length; k++) {
                        var file = files[k];
                        // append file in the log
                        log($('<a/>').attr({ href: file.LinkingUrl, target: '_blank' }).text(file.Name));
                        // sanity check
                        if (++numFiles >= MAX_FILES) return;
                    }
                }
            }
        }).fail(function (xhr) {
            log(xhr.responseText, true);
        });

        function filterList(list) {
            return list.RootFolder.Files.results.length || list.Items.results.length;
        }

        function filterListItem(item) {
            return item.Folder && item.Folder.Exists && item.Folder.Files && item.Folder.Files.results && item.Folder.Files.results.length;
        }

        function filterFolderFile(file) {
            return file.Exists && file.LinkingUrl && file.LinkingUrl.length;
        }
    }

    // open the add-in in a new window/tab
    function reopen() {
        window.open(location.href, '_blank');
    }

    // helper method to append lines to the message element
    function log(mixed, isError) {
        $('#test_message').append($('<div/>').css(isError ? { color: 'red' } : {}).html(mixed));
    }

    // helper method to build the _api url and params when calling the host-web from within the add-in web
    function api(query, params) {
        params = $.extend(true, { '@t': '\'' + hostWeb + '\'' }, params);
        return _spPageContextInfo.siteAbsoluteUrl + '/_api/SP.AppContextSite(@t)/' + query + '?' + $.param(params);
    }

    // helper method for ajax requests
    function http(options) {
        return $.ajax($.extend(true, {
            method: 'GET',
            headers: {
                'Accept': 'application/json;odata=verbose',
                'Content-Type': 'application/json;odata=verbose',
                'X-RequestDigest': $('#__REQUESTDIGEST').val()
            }
        }, options));
    }

}

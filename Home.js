(function () {
    "use strict";

    var messageBanner;

    var localStorageUsername = 'windsorapiusername';
    var localStorageToken = 'windsorapitoken';

    var pageSize = 1000;
    var callsPerEnter = 2000;

    var endpoint = 'https://api.windsor.ai/{username}/{username}_attribution/public/{username}_attributions_and_costs?api_key={apikey}&_select={selectedcolumns}&date=$gte.{date}&_page={page}&_page_size={pagesize}';
    var activeEndpoint;

    var cancel;
    var finished;
    var failed;

    var selected;
    var values;
    
    var statusLabel;

    var positionCursor;

    var datePicker;
    var dateRest;

    var allColumnsCounter;
    var allPagesCounter;

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {

            //Initialize from localStorage
            var userName = window.localStorage.getItem(localStorageUsername);
            var token = window.localStorage.getItem(localStorageToken);

            $('#username').val(userName);
            $('#token').val(token);

            // Initialize the notification mechanism and hide it
            var element = document.querySelector('.MessageBanner');
            messageBanner = new components.MessageBanner(element);
            messageBanner.hideBanner();

            // Initialize status label
            statusLabel = $('#status');

            // Add a click event handler for the buttons.
            $('#retrieve').click(retrieveColumnList);
            $('#all').click(selectdeselectall);
            $('#execute').click(execute);
            $('#cancel').click(activateCancel);

            // Initialize Date Picker
            if ($.fn.DatePicker) {
                $('.ms-DatePicker').DatePicker({
                    onSet: function (context) {
                        datePicker = new Date(context.select);
                    }
                });
            }

            //Credentials hover
            $('.credentials').hover(showCredentials, function () { });
            
        });
    };

    function retrieveColumnList() {
        $('.hide-credentials').hover(function () { }, hideCredentials);

        // Save to local storage
        var user = $('#username').val();
        var token = $('#token').val();

        if (user !== '' && token !== '') {
            localStorage.setItem(localStorageUsername, user);
            localStorage.setItem(localStorageToken, token);

            // Hide credentials
            $('.hide-credentials').slideUp('slow');

            // Empty List
            var selectedCheckboxes = $('.selected-checkboxes');
            selectedCheckboxes.empty();

            // Call API
            var isColumnRetrieve = true;
            var activeEndpoint = buildEndpoint(user, token, isColumnRetrieve).replace('{page}', 1);

            showNotification("Connecting to Server", "Retrieving Column List");

            retrieveColumnListAjaxCall(activeEndpoint);
        }
        else {
            showNotification('Error', 'Username and token cannot be empty!')
        }
    }

    function retrieveColumnListAjaxCall(endpoint) {

        $.ajax({
            url: endpoint,
            type: 'GET',
            dataType: 'json',
            crossDomain: true
        }).done(function (data) {

            if (data.error) {
                showNotification('The API returned an error', 'Please check the credentials you supplied and try again');

                return;
            }

            messageBanner.hideBanner();

            var columnsContainer = data[0];

            var selectedCheckboxes = $('.selected-checkboxes');
            selectedCheckboxes.empty();

            allColumnsCounter = 0;

            for (var column in columnsContainer) {
                var checkbox = '<label class="container">' + column +
                    '<input type="checkbox" checked="checked">' +
                    '<span class="checkmark"></span>' +
                    '</label>';

                $(checkbox).appendTo(selectedCheckboxes);

                allColumnsCounter++;
            }

            // Handle Select/Deselect All Button
            $('#all').attr('action', 'deselectall');
            $('.selectdeselectall').text('Deselect All');
        }).fail(function (status) {
            showNotification('The API returned an error', 'Please check the credentials you supplied and try again');
        });
    }

    function retievePagesNumberAjaxCall(activeEndpoint) {
        var tempEndpoint = activeEndpoint.replace('_page={page}&_page_size=' + pageSize, '_count=date');

        $.ajax({
            url: tempEndpoint,
            type: 'GET',
            dataType: 'json',
            crossDomain: true
        }).done(function (data) {

            if (data.count) {
                allPagesCounter = Math.ceil(data.count / pageSize) + 1;
            }
            else {
                allPagesCounter = '-';
            }

            next();
        }).fail(function (status) {
            allPagesCounter = '-';
        });
    }

    function selectdeselectall() {
        var self = $(this);

        if (self.attr('action') == 'selectall') {
            $(':checkbox').each(function () {
                this.checked = true;
            });

            self.attr('action', 'deselectall');
            $('.selectdeselectall').text('Deselect All');
        } else {
            $(':checkbox').each(function () {
                this.checked = false;
            });

            self.attr('action', 'selectall');
            $('.selectdeselectall').text('Select All');
        }
    }

    function execute() {

        var user = $('#username').val();
        var token = $('#token').val();

        if (user !== '' && token !== '') {

            selected = [];
            $('input[type = checkbox]').each(function () {
                if (this.checked) {
                    selected.push($(this).parent()[0].textContent);
                }
            });

            if (selected.length != 0) {
                values = [];
                values.length = 0;
                values[0] = selected;

                var user = $('#username').val();
                var token = $('#token').val();

                failed = false;
                cancel = false;
                finished = false;

                page = 1;
                positionCursor = 0;

                if (datePicker != undefined && datePicker != 'Invalid Date') {
                    dateRest = datePicker.toISOString().replace('T', ' ').replace('Z', '');
                }
                else {
                    dateRest = null;
                }

                var isColumnRetrieve = false;
                activeEndpoint = buildEndpoint(user, token, isColumnRetrieve);

                retievePagesNumberAjaxCall(activeEndpoint);

                showCancel();
            }
            else {
                showNotification('Error', 'Please select at least one data column');
            }
        }
        else {
            showNotification('Error', 'Username and token cannot be empty!')
        }
    }

    function executeAjaxCall(endpoint) {

        $.ajax({
            url: endpoint,
            type: 'GET',
            dataType: 'json',
            crossDomain: true
        }).done(function (data) {

            try {
                if (data.error) {
                    showNotification('The API returned an error', 'Please check the credentials you supplied and try again');

                    return;
                }
                else if (data.length == 0) {
                    finished = true;
                }

                for (var i = 0; i < data.length; i++) {

                    var rowValues = [];
                    var j = 0;

                    for (var pair in data[i]) {

                        rowValues[j] = data[i][pair];

                        j++;
                    }

                    values[values.length] = rowValues;
                }

                if (finished || cancel || values.length >= callsPerEnter) {
                    writeToExcel();
                }
                else {
                    next();
                }
            }
            catch (ex) {
                showNotification('Error', ex);
            }

            
        }).fail(function (status) {
            failed = true;

            showNotification('The API returned an error', 'Please check the credentials you supplied and try again');
            writeToExcel();
        });
    }

    function showCancel() {
        $('.footer-cancel').css('display', 'block');
    }

    function hideCancel() {
        $('.footer-cancel').css('display', 'none');
    }

    function activateCancel() {
        cancel = true;

        statusLabel.empty();

        hideCancel();

        if (values.length != 0) {
            showNotification('Operation canceled', 'Writing data to document...');
        }
    }

    var page;
    function next() {
        statusLabel.text('Page ' + page + '/' + allPagesCounter);

        var currEndpoint = activeEndpoint.replace('{page}', page++);

        executeAjaxCall(currEndpoint);
    }

    function buildEndpoint(user, token, isColumnRetrieve) {
        var currEndpoint = null;

        // Handle credentials

        currEndpoint = endpoint.split('{username}').join(user).replace('{apikey}', token);

        // Handle columns selection

        if (isColumnRetrieve || (allColumnsCounter === selected.length)) {
            currEndpoint = currEndpoint.replace('&_select={selectedcolumns}', '');
        }
        else {
            currEndpoint = currEndpoint.replace('{selectedcolumns}', selected.join(','));
        }

        // Handle page size parameter

        if (isColumnRetrieve) {
            currEndpoint = currEndpoint.replace('{pagesize}', 1);
        }
        else {
            currEndpoint = currEndpoint.replace('{pagesize}', pageSize);
        }

        // Handle start date selection

        if (dateRest) {
            currEndpoint = currEndpoint.replace('{date}', dateRest);
        }
        else {
            currEndpoint = currEndpoint.replace('&date=$gte.{date}', '');
        }

        return currEndpoint;
    }

    function writeToExcel() {

        Excel.run(function (ctx) {
            ctx.application.suspendApiCalculationUntilNextSync();
            ctx.application.suspendScreenUpdatingUntilNextSync();

            var sheet = ctx.workbook.worksheets.getActiveWorksheet();

            var range;

            if (positionCursor == 0) {
                range = sheet.getUsedRange();
                range.clear();
            }

            sheet.freezePanes.freezeRows(1);

            range = sheet.getRangeByIndexes(positionCursor, 0, values.length, values[0].length);

            range.values = values;
            range.format.autofitColumns();

            positionCursor = positionCursor + values.length;

            return ctx.sync().then(function () {

                values.length = 0;

                if (cancel) {
                    messageBanner.hideBanner();
                }

                if (finished) {
                    statusLabel.empty();
                    hideCancel();
                    showNotification('Finished', '');
                }

                if (!finished && !cancel && !failed) {
                    next();
                }
            });

        }).catch(errorHandler);
    }

    function hideCredentials() {
        $('.hide-credentials').slideUp('slow');
    }

    function showCredentials() {
        var user = $('#username').val();
        var token = $('#token').val();

        if (user !== '' && token !== '') {
            $('.hide-credentials').slideDown('slow');
        }
    }

    // Helper function for treating errors
    function errorHandler(error) {
        // Always be sure to catch any accumulated errors that bubble up from the Excel.run execution
        showNotification("Write Error", error);
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    }

    // Helper function for displaying notifications
    function showNotification(header, content) {
        $("#notification-header").text(header);
        $("#notification-body").text(content);
        messageBanner.showBanner();
        messageBanner.toggleExpansion();
    }
})();

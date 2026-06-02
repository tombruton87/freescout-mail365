(function() {
    var MAIL365_METHOD = '4';
    var MAIL365_IN_PROTOCOL = '3';

    function showResult(selector, type, msg) {
        var cls = type === 'success' ? 'text-success' : 'text-danger';
        var icon = type === 'success' ? 'glyphicon-ok' : 'glyphicon-exclamation-sign';
        $(selector).empty().append(
            $('<span>').addClass(cls).append(
                $('<i>').addClass('glyphicon ' + icon),
                ' ',
                $('<span>').text(msg)
            )
        );
    }

    function timeAgo(dateStr) {
        var now = new Date();
        var date = new Date(dateStr);
        var seconds = Math.floor((now - date) / 1000);
        if (seconds < 60) return 'just now';
        var minutes = Math.floor(seconds / 60);
        if (minutes < 60) return minutes + ' minute' + (minutes !== 1 ? 's' : '') + ' ago';
        var hours = Math.floor(minutes / 60);
        if (hours < 24) return hours + ' hour' + (hours !== 1 ? 's' : '') + ' ago';
        var days = Math.floor(hours / 24);
        return days + ' day' + (days !== 1 ? 's' : '') + ' ago';
    }

    function showEmailWarning(selector, warning) {
        if (!warning) return;
        $(selector).append(
            $('<div>').addClass('text-warning').css('margin-top', '4px').append(
                $('<i>').addClass('glyphicon glyphicon-warning-sign'),
                ' ',
                $('<span>').text(warning)
            )
        );
    }

    function showPermissionWarnings(selector, warnings) {
        if (!warnings) return;
        $(selector).append(
            $('<div>').addClass('text-warning').css('margin-top', '4px').append(
                $('<i>').addClass('glyphicon glyphicon-warning-sign'),
                ' ',
                $('<span>').text(warnings)
            )
        );
    }

    function convertTimestamps() {
        $('.m365-time-ago[data-utc]').each(function() {
            var utc = $(this).data('utc');
            if (!utc) return;
            $(this).text(timeAgo(utc)).attr('title', new Date(utc).toLocaleString());
        });
    }

    // ========== OUTGOING CONNECTION PAGE ==========

    function isConnectionPage() {
        return $('input[name="out_method"]').length > 0;
    }

    function isIncomingPage() {
        return $('select[name="in_protocol"]').length > 0;
    }

    function getMailboxId() {
        return $('body').data('mailbox_id');
    }

    function getCurrentMethod() {
        var checked = $('input[name="out_method"]:checked').val();
        if (!checked) return MAIL365_METHOD;
        return checked;
    }

    function injectRadioButton() {
        var smtpRadio = $('input[name="out_method"][value="3"]').closest('.control-group');
        if (!smtpRadio.length) return;

        var currentMethod = getCurrentMethod();
        var checked = currentMethod == MAIL365_METHOD ? ' checked="checked"' : '';

        var html = '<div class="control-group">' +
            '<label class="radio" for="out_method_' + MAIL365_METHOD + '">' +
            '<input type="radio" name="out_method" value="' + MAIL365_METHOD + '" ' +
            'id="out_method_' + MAIL365_METHOD + '"' + checked + '> Microsoft 365 API' +
            '</label>' +
            '</div>';

        smtpRadio.after(html);
    }

    function injectOptionsSection() {
        var smtpOptions = $('#out_method_3_options');
        if (!smtpOptions.length) return;

        var currentMethod = getCurrentMethod();
        var hiddenClass = currentMethod != MAIL365_METHOD ? ' hidden' : '';

        var html = '<div id="out_method_' + MAIL365_METHOD + '_options" class="out_method_options' + hiddenClass + ' margin-top">' +
            '<hr/>' +
            '<div class="form-group">' +
                '<div class="col-sm-6 col-sm-offset-2">' +
                    '<p class="text-help">' +
                        'Send emails via the Microsoft Graph API (HTTPS) instead of SMTP. ' +
                        'This bypasses SMTP port restrictions that many hosting providers enforce. ' +
                        'You need an Azure AD app registration with <strong>Mail.Send</strong> application permission.' +
                    '</p>' +
                    '<p id="m365-shared-creds-note" class="text-muted" style="display:none">' +
                        '<i class="glyphicon glyphicon-info-sign"></i> ' +
                        'Uses the same Azure app credentials as the incoming settings. Changes here will also apply to fetching.' +
                    '</p>' +
                '</div>' +
            '</div>' +
            '<div class="form-group">' +
                '<label for="m365_tenant_id" class="col-sm-2 control-label">Tenant ID</label>' +
                '<div class="col-sm-6">' +
                    '<input id="m365_tenant_id" type="text" class="form-control input-sized" ' +
                    'name="m365_tenant_id" value="" maxlength="255" ' +
                    'placeholder="xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" autocomplete="off">' +
                    '<p class="form-help">Azure Active Directory → Properties → Tenant ID</p>' +
                '</div>' +
            '</div>' +
            '<div class="form-group">' +
                '<label for="m365_client_id" class="col-sm-2 control-label">Client ID</label>' +
                '<div class="col-sm-6">' +
                    '<input id="m365_client_id" type="text" class="form-control input-sized" ' +
                    'name="m365_client_id" value="" maxlength="255" ' +
                    'placeholder="xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" autocomplete="off">' +
                    '<p class="form-help">Azure AD → App registrations → Your app → Application (client) ID</p>' +
                '</div>' +
            '</div>' +
            '<div class="form-group">' +
                '<label for="m365_client_secret" class="col-sm-2 control-label">Client Secret</label>' +
                '<div class="col-sm-6">' +
                    '<input id="m365_client_secret" type="password" class="form-control input-sized" ' +
                    'name="m365_client_secret" value="" maxlength="255" autocomplete="new-password">' +
                    '<p class="form-help">Azure AD → App registrations → Your app → Certificates &amp; secrets</p>' +
                '</div>' +
            '</div>' +
            '<div class="form-group">' +
                '<div class="col-sm-6 col-sm-offset-2">' +
                    '<button type="button" id="m365-test-btn" class="btn btn-default" data-loading-text="Testing…">' +
                        '<i class="glyphicon glyphicon-ok-circle"></i> Test Connection' +
                    '</button>' +
                    '<span id="m365-test-result" class="margin-left"></span>' +
                '</div>' +
            '</div>' +
            '<div class="form-group">' +
                '<div class="col-sm-6 col-sm-offset-2">' +
                    '<div id="m365-save-status"></div>' +
                '</div>' +
            '</div>' +
            '<div id="m365-out-secret-expiry" class="form-group" style="display:none;">' +
                '<label class="col-sm-2 control-label">Secret Expiry</label>' +
                '<div class="col-sm-6">' +
                    '<p class="form-control-static" id="m365-out-secret-expiry-text"></p>' +
                '</div>' +
            '</div>' +
            '<div id="m365-send-status" class="form-group" style="display:none;">' +
                '<label class="col-sm-2 control-label">Send Status</label>' +
                '<div class="col-sm-6">' +
                    '<div id="m365-send-status-content"></div>' +
                    '<div style="margin-top:6px;">' +
                        '<button type="button" id="m365-view-send-log-btn" class="btn btn-default btn-xs">' +
                            '<i class="glyphicon glyphicon-list"></i> View Send Log' +
                        '</button>' +
                    '</div>' +
                    '<div id="m365-send-log" style="display:none;margin-top:8px;">' +
                        '<table class="table table-condensed table-bordered" id="m365-send-log-table" style="font-size:12px;"></table>' +
                    '</div>' +
                '</div>' +
            '</div>' +
            '<div id="m365-retry-queue" class="form-group" style="display:none;">' +
                '<label class="col-sm-2 control-label">Retry Queue</label>' +
                '<div class="col-sm-6">' +
                    '<div id="m365-retry-queue-content"></div>' +
                    '<div id="m365-retry-queue-actions" style="margin-top:6px;display:none;">' +
                        '<button type="button" id="m365-clear-retry-queue-btn" class="btn btn-danger btn-xs">' +
                            '<i class="glyphicon glyphicon-trash"></i> Clear Queue' +
                        '</button>' +
                        '<span id="m365-retry-queue-status" class="margin-left"></span>' +
                    '</div>' +
                '</div>' +
            '</div>' +
            '<hr/>' +
            '</div>';

        smtpOptions.after(html);
    }

    function injectHiddenOutServer() {
        var form = $('input[name="out_method"]').closest('form');
        if (!form.find('input[name="m365_out_server_hidden"]').length) {
            form.append('<input type="hidden" name="m365_out_server_hidden" value="1">');
        }
    }

    function loadSavedSettings() {
        var mailboxId = getMailboxId();
        if (!mailboxId) return;

        $.ajax({
            url: '/mail365/settings/' + mailboxId,
            type: 'GET',
            dataType: 'json',
            success: function(response) {
                if (response.status === 'success') {
                    $('#m365_tenant_id').val(response.tenant_id || '');
                    $('#m365_client_id').val(response.client_id || '');
                    if (response.has_secret) {
                        $('#m365_client_secret').attr('placeholder', '••••••••••••••••');
                        $('#m365-shared-creds-note').show();
                    }
                    if (response.secret_expiry) {
                        showSecretExpiryRow('#m365-out-secret-expiry', '#m365-out-secret-expiry-text', response.secret_expiry);
                    }
                    if (response.last_send_success || response.last_send_error) {
                        showSendStatus(response);
                    }
                }
            }
        });
    }

    function bindTestButton() {
        $(document).on('click', '#m365-test-btn', function() {
            var btn = $(this);
            btn.button('loading');
            $('#m365-test-result').empty();

            var data = {
                tenant_id: $('#m365_tenant_id').val(),
                client_id: $('#m365_client_id').val(),
                client_secret: $('#m365_client_secret').val(),
                _token: $('input[name="_token"]').val()
            };

            $.ajax({
                url: '/mail365/test/' + getMailboxId(),
                type: 'POST',
                data: data,
                dataType: 'json',
                success: function(response) {
                    showResult('#m365-test-result', response.status, response.msg);
                    showPermissionWarnings('#m365-test-result', response.permission_warnings);
                    showEmailWarning('#m365-test-result', response.email_warning);
                    if (response.secret_expiry) {
                        showSecretExpiryRow('#m365-out-secret-expiry', '#m365-out-secret-expiry-text', response.secret_expiry);
                    }
                },
                error: function(xhr) {
                    var detail = '';
                    try {
                        var json = JSON.parse(xhr.responseText);
                        detail = json.msg || json.message || json.error || '';
                    } catch(e) {
                        detail = xhr.status + ' ' + xhr.statusText;
                    }
                    showResult('#m365-test-result', 'error', 'Request failed: ' + detail);
                },
                complete: function() {
                    btn.button('reset');
                }
            });
        });
    }

    function bindFormSubmit() {
        var form = $('input[name="out_method"]').closest('form');

        form.on('submit', function(e) {
            var method = $('input[name="out_method"]:checked').val();
            if (method != MAIL365_METHOD) return true;

            var tenantId = $('#m365_tenant_id').val();
            var clientId = $('#m365_client_id').val();
            var clientSecret = $('#m365_client_secret').val();

            if (!tenantId || !clientId) {
                e.preventDefault();
                showResult('#m365-save-status', 'error', 'Please fill in Tenant ID and Client ID.');
                return false;
            }

            // Set out_server so isOutActive() returns true in core
            var serverField = form.find('input[name="out_server"]');
            serverField.val('graph.microsoft.com');
            var portField = form.find('input[name="out_port"]');
            portField.val('443');

            // Save Azure creds via module endpoint
            e.preventDefault();
            var submitBtn = form.find('button[type="submit"]');
            submitBtn.prop('disabled', true).text('Saving…');

            var data = {
                tenant_id: tenantId,
                client_id: clientId,
                _token: $('input[name="_token"]').val()
            };
            if (clientSecret) {
                data.client_secret = clientSecret;
            }

            if (!clientSecret) {
                data.keep_secret = '1';
            }

            $.ajax({
                url: '/mail365/settings/' + getMailboxId(),
                type: 'POST',
                data: data,
                dataType: 'json',
                success: function(response) {
                    if (response.status === 'success') {
                        form.off('submit');
                        form.submit();
                    } else {
                        submitBtn.prop('disabled', false).text('Save Settings');
                        showResult('#m365-save-status', 'error', response.msg);
                    }
                },
                error: function(xhr) {
                    submitBtn.prop('disabled', false).text('Save Settings');
                    var detail = '';
                    try {
                        var json = JSON.parse(xhr.responseText);
                        detail = json.msg || json.message || '';
                    } catch(e) {
                        detail = xhr.status + ' ' + xhr.statusText;
                    }
                    showResult('#m365-save-status', 'error', 'Failed to save 365 API settings. ' + detail);
                }
            });

            return false;
        });
    }

    // ========== INCOMING CONNECTION PAGE ==========

    function toggleIncomingSettings() {
        var protocol = $('select[name="in_protocol"]').val();
        var settingsDiv = $('#mail365_in_settings');
        var defaultFields = $('.form-group').has('input[name="in_server"], input[name="in_port"], input[name="in_username"], input[name="in_password"], select[name="in_encryption"], input[name="in_validate_cert"]');
        var imapFolders = $('.form-group').has('select[name="in_imap_folders[]"]');
        var imapSentFolder = $('.form-group').has('input[name="imap_sent_folder"]');
        var checkConnBtn = $('button:contains("Check Connection")');

        if (protocol == MAIL365_IN_PROTOCOL) {
            settingsDiv.show();
            defaultFields.hide();
            imapFolders.hide();
            imapSentFolder.hide();
            checkConnBtn.hide();
        } else {
            settingsDiv.hide();
            defaultFields.show();
            imapSentFolder.show();
            checkConnBtn.show();
            if (protocol == '1') {
                imapFolders.show();
            }
        }
    }

    function renderExpiryContent(el, expiry) {
        el.empty();
        var days = expiry.days_left;
        var label = expiry.display_name ? '"' + expiry.display_name + '" expires ' : 'Expires ';
        var text = label + expiry.date + ' (' + days + ' day' + (days !== 1 ? 's' : '') + ' remaining)';

        var cls, icon;
        if (days <= 14) {
            cls = 'text-danger';
            icon = 'glyphicon-exclamation-sign';
        } else if (days <= 30) {
            cls = 'text-warning';
            icon = 'glyphicon-warning-sign';
        } else {
            cls = 'text-success';
            icon = 'glyphicon-ok';
        }

        el.append(
            $('<span>').addClass(cls).append(
                $('<i>').addClass('glyphicon ' + icon),
                ' ',
                $('<span>').text(text)
            )
        );
    }

    function showSecretExpiryRow(wrapperSelector, textSelector, expiry) {
        if (!expiry || !expiry.days_left) {
            $(wrapperSelector).hide();
            return;
        }
        renderExpiryContent($(textSelector), expiry);
        $(wrapperSelector).show();
    }

    function updateExpiryDateField(expiry) {
        if (!expiry || !expiry.date) return;
        var field = $('#m365_secret_expiry_date');
        if (field.length && !field.val()) {
            field.val(expiry.date);
        }
        updateExpiryStatus();
    }

    function updateExpiryStatus() {
        var dateVal = $('#m365_secret_expiry_date').val();
        var status = $('#m365-secret-expiry-status');
        status.empty();
        if (!dateVal) return;

        var now = new Date();
        now.setHours(0, 0, 0, 0);
        var expDate = new Date(dateVal + 'T00:00:00');
        var daysLeft = Math.ceil((expDate - now) / 86400000);

        var cls, icon, text;
        if (daysLeft < 0) {
            cls = 'text-danger'; icon = 'glyphicon-exclamation-sign'; text = 'Expired';
        } else if (daysLeft <= 14) {
            cls = 'text-danger'; icon = 'glyphicon-exclamation-sign'; text = daysLeft + ' day' + (daysLeft !== 1 ? 's' : '') + ' remaining';
        } else if (daysLeft <= 30) {
            cls = 'text-warning'; icon = 'glyphicon-warning-sign'; text = daysLeft + ' day' + (daysLeft !== 1 ? 's' : '') + ' remaining';
        } else {
            cls = 'text-success'; icon = 'glyphicon-ok'; text = daysLeft + ' days remaining';
        }

        status.append(
            $('<span>').addClass(cls).append(
                $('<i>').addClass('glyphicon ' + icon), ' ',
                $('<span>').text(text)
            )
        );
    }

    function bindIncomingTestButton() {
        $(document).on('click', '#m365-in-test-btn', function() {
            var btn = $(this);
            btn.button('loading');
            $('#m365-in-test-result').empty();

            var data = {
                tenant_id: $('#m365_in_tenant_id').val(),
                client_id: $('#m365_in_client_id').val(),
                client_secret: $('#m365_in_client_secret').val(),
                verify_email: $('#m365_shared_mailbox_email').val() || '',
                _token: $('input[name="_token"]').val()
            };

            $.ajax({
                url: '/mail365/test/' + getMailboxId(),
                type: 'POST',
                data: data,
                dataType: 'json',
                success: function(response) {
                    showResult('#m365-in-test-result', response.status, response.msg);
                    showPermissionWarnings('#m365-in-test-result', response.permission_warnings);
                    showEmailWarning('#m365-in-test-result', response.email_warning);
                    if (response.secret_expiry) {
                        updateExpiryDateField(response.secret_expiry);
                    }
                },
                error: function(xhr) {
                    var detail = '';
                    try {
                        var json = JSON.parse(xhr.responseText);
                        detail = json.msg || json.message || json.error || '';
                    } catch(e) {
                        detail = xhr.status + ' ' + xhr.statusText;
                    }
                    showResult('#m365-in-test-result', 'error', 'Request failed: ' + detail);
                },
                complete: function() {
                    btn.button('reset');
                }
            });
        });
    }

    function renderFolderCheckboxes(container, folders, selected, depth) {
        folders.forEach(function(folder) {
            var isChecked = selected[folder.id] ? true : false;
            var indent = depth * 20;

            var label = $('<label>').css({display: 'block', padding: '3px 0', cursor: 'pointer', 'padding-left': indent + 'px'});
            var cb = $('<input>').attr({
                type: 'checkbox',
                'data-folder-id': folder.id,
                'data-folder-name': folder.name
            }).css('margin-right', '6px');

            if (isChecked) cb.prop('checked', true);

            var nameSpan = $('<span>').text(folder.name);
            var countSpan = $('<span>').addClass('text-muted').text(' (' + folder.item_count + ')');

            label.append(cb, nameSpan, countSpan);

            if (folder.name.toLowerCase() === 'inbox' && depth === 0) {
                label.append(
                    $('<span>').addClass('label label-info').text('default').css({
                        'font-size': '10px',
                        'margin-left': '6px',
                        'vertical-align': 'middle'
                    })
                );
            }

            container.append(label);

            if (folder.children && folder.children.length) {
                renderFolderCheckboxes(container, folder.children, selected, depth + 1);
            }
        });
    }

    function flattenFolders(folders, depth) {
        depth = depth || 0;
        var result = [];
        folders.forEach(function(folder) {
            result.push({id: folder.id, name: folder.name, depth: depth});
            if (folder.children && folder.children.length) {
                result = result.concat(flattenFolders(folder.children, depth + 1));
            }
        });
        return result;
    }

    function bindFolderPicker() {
        $(document).on('click', '#m365-detect-folders-btn', function() {
            var btn = $(this);
            btn.button('loading');

            $.ajax({
                url: '/mail365/folders/' + getMailboxId(),
                type: 'GET',
                dataType: 'json',
                success: function(response) {
                    if (response.status !== 'success') {
                        showResult('#m365-in-test-result', 'error', response.msg || 'Failed to load folders');
                        return;
                    }

                    var container = $('#m365-folder-checkboxes');
                    container.empty();

                    var selected = {};
                    (response.selected || []).forEach(function(f) {
                        selected[f.id] = true;
                    });

                    var folders = response.folders || [];
                    if (!folders.length) {
                        container.append($('<p>').addClass('text-muted').text('No folders found.'));
                        return;
                    }

                    renderFolderCheckboxes(container, folders, selected, 0);

                    var flat = flattenFolders(folders);
                    populateMoveFolderSelect(flat);

                    $('#m365-folder-picker').show();
                },
                error: function() {
                    showResult('#m365-in-test-result', 'error', 'Failed to detect folders');
                },
                complete: function() {
                    btn.button('reset');
                }
            });
        });

        $(document).on('click', '#m365-cancel-folders-btn', function() {
            $('#m365-folder-picker').hide();
        });

        $(document).on('click', '#m365-save-folders-btn', function() {
            var btn = $(this);
            var folders = [];

            $('#m365-folder-checkboxes input:checked').each(function() {
                folders.push({
                    id: $(this).data('folder-id'),
                    name: $(this).data('folder-name')
                });
            });

            btn.prop('disabled', true).text('Saving…');

            $.ajax({
                url: '/mail365/folders/' + getMailboxId(),
                type: 'POST',
                data: {
                    folders: folders,
                    _token: $('input[name="_token"]').val()
                },
                dataType: 'json',
                success: function(response) {
                    if (response.status === 'success') {
                        var list = $('#m365-folder-list');
                        list.empty();
                        if (!folders.length) {
                            list.append($('<span>').addClass('text-muted').text('Inbox only (default)'));
                        } else {
                            folders.forEach(function(f) {
                                list.append(
                                    $('<span>').addClass('label label-default m365-folder-tag')
                                        .css({display: 'inline-block', margin: '2px 4px 2px 0', padding: '4px 8px', 'font-size': '12px'})
                                        .text(f.name)
                                );
                            });
                        }
                        $('#m365-folder-picker').hide();
                    } else {
                        showResult('#m365-in-test-result', 'error', response.msg || 'Failed to save folders');
                    }
                },
                error: function() {
                    showResult('#m365-in-test-result', 'error', 'Failed to save folder selection');
                },
                complete: function() {
                    btn.prop('disabled', false).html('<i class="glyphicon glyphicon-ok"></i> Save Folder Selection');
                }
            });
        });
    }

    // ========== POST-FETCH ACTION ==========

    function bindPostFetchAction() {
        $(document).on('change', '#m365_post_fetch_action', function() {
            var action = $(this).val();
            if (action === 'move') {
                $('#m365-move-folder-wrapper').show();
                loadMoveFolders();
            } else {
                $('#m365-move-folder-wrapper').hide();
            }
        });

        $(document).on('change', '#m365_post_fetch_move_folder', function() {
            var name = $(this).find('option:selected').text();
            if (name && name !== 'Select a folder…' && name !== 'Loading…') {
                $('#m365_post_fetch_move_folder_name').val(name);
            }
        });
    }

    var moveFoldersLoaded = false;

    function loadMoveFolders() {
        if (moveFoldersLoaded) return;

        var select = $('#m365_post_fetch_move_folder');
        var currentVal = select.val();

        select.find('option:first').text('Loading…');

        $.ajax({
            url: '/mail365/folders/' + getMailboxId(),
            type: 'GET',
            dataType: 'json',
            success: function(response) {
                if (response.status !== 'success') return;
                var flat = flattenFolders(response.folders || []);
                populateMoveFolderSelect(flat);
                if (currentVal) {
                    select.val(currentVal);
                }
                moveFoldersLoaded = true;
            },
            error: function() {
                select.find('option:first').text('Select a folder…');
            }
        });
    }

    function populateMoveFolderSelect(folders) {
        var select = $('#m365_post_fetch_move_folder');
        var currentVal = select.val();

        select.empty().append($('<option>').val('').text('Select a folder…'));

        folders.forEach(function(folder) {
            var prefix = '';
            for (var i = 0; i < (folder.depth || 0); i++) prefix += '    ';
            select.append($('<option>').val(folder.id).text(prefix + folder.name));
        });

        if (currentVal) {
            select.val(currentVal);
        }
    }

    // ========== FETCH LOG ==========

    function bindFetchLog() {
        var fetchLogInterval = null;

        function loadFetchLog(cb) {
            $.ajax({
                url: '/mail365/fetch-log/' + getMailboxId(),
                type: 'GET',
                dataType: 'json',
                success: function(response) {
                    if (response.status !== 'success') return;
                    renderFetchLogTable(response.log || []);
                },
                complete: cb || function() {}
            });
        }

        function renderFetchLogTable(log) {
            var table = $('#m365-fetch-log-table');
            table.empty();

            if (!log.length) {
                table.append($('<tr>').append(
                    $('<td>').addClass('text-muted').text('No fetch log entries yet.')
                ));
                return;
            }

            table.append(
                $('<thead>').append(
                    $('<tr>').append(
                        $('<th>').text('Time'),
                        $('<th>').text('Status'),
                        $('<th>').text('Messages'),
                        $('<th>').text('Folders'),
                        $('<th>').text('Duration')
                    )
                )
            );

            var tbody = $('<tbody>');

            log.forEach(function(entry) {
                var statusEl = entry.status === 'success'
                    ? $('<span>').addClass('text-success').append($('<i>').addClass('glyphicon glyphicon-ok'))
                    : $('<span>').addClass('text-danger').append($('<i>').addClass('glyphicon glyphicon-remove'));

                var ts = entry.timestamp || '';
                if (ts) {
                    var d = new Date(ts);
                    ts = d.toLocaleDateString() + ' ' + d.toLocaleTimeString();
                }

                var row = $('<tr>');
                row.append($('<td>').text(ts));
                row.append($('<td>').append(statusEl));
                row.append($('<td>').text(entry.messages));
                row.append($('<td>').text((entry.folders || []).join(', ') || '-'));
                row.append($('<td>').text(entry.duration + 's'));
                tbody.append(row);

                if (entry.error) {
                    var errorRow = $('<tr>').addClass('danger');
                    errorRow.append(
                        $('<td>').attr('colspan', 5).addClass('text-danger')
                            .css('font-size', '11px')
                            .text(entry.error)
                    );
                    tbody.append(errorRow);
                }
            });

            table.append(tbody);
        }

        $(document).on('click', '#m365-view-log-btn', function() {
            var logDiv = $('#m365-fetch-log');
            if (logDiv.is(':visible')) {
                logDiv.hide();
                if (fetchLogInterval) { clearInterval(fetchLogInterval); fetchLogInterval = null; }
                return;
            }

            var btn = $(this);
            btn.prop('disabled', true);

            loadFetchLog(function() {
                btn.prop('disabled', false);
                logDiv.show();
                fetchLogInterval = setInterval(function() {
                    if (!logDiv.is(':visible')) { clearInterval(fetchLogInterval); fetchLogInterval = null; return; }
                    loadFetchLog();
                }, 30000);
            });
        });
    }

    // ========== SEND STATUS & LOG ==========

    function showSendStatus(response) {
        var content = $('#m365-send-status-content');
        content.empty();

        if (response.last_send_error) {
            var errorAgo = response.last_send_error_at ? timeAgo(response.last_send_error_at) : '';
            content.append(
                $('<p>').css('margin-bottom', '4px').append(
                    $('<span>').addClass('text-danger').append(
                        $('<i>').addClass('glyphicon glyphicon-remove')
                    ),
                    ' Last send error' + (errorAgo ? ' ' + errorAgo : '')
                ),
                $('<p>').addClass('text-danger').css({'font-size': '12px', 'margin-bottom': '4px'}).text(response.last_send_error)
            );
        }

        if (response.last_send_success) {
            var ago = timeAgo(response.last_send_success);
            content.append(
                $('<p>').css('margin-bottom', '0').append(
                    $('<span>').addClass('text-success').append(
                        $('<i>').addClass('glyphicon glyphicon-ok')
                    ),
                    ' Last successful send ' + ago
                )
            );
        }

        $('#m365-send-status').show();
    }

    function bindSendLog() {
        var sendLogInterval = null;

        function loadSendLog(cb) {
            $.ajax({
                url: '/mail365/send-log/' + getMailboxId(),
                type: 'GET',
                dataType: 'json',
                success: function(response) {
                    if (response.status !== 'success') return;
                    renderSendLogTable(response.log || []);
                },
                complete: cb || function() {}
            });
        }

        function renderSendLogTable(log) {
            var table = $('#m365-send-log-table');
            table.empty();

            if (!log.length) {
                table.append($('<tr>').append(
                    $('<td>').addClass('text-muted').text('No send log entries yet.')
                ));
                return;
            }

            table.append(
                $('<thead>').append(
                    $('<tr>').append(
                        $('<th>').text('Time'),
                        $('<th>').text('Status'),
                        $('<th>').text('To'),
                        $('<th>').text('Subject'),
                        $('<th>').text('Duration')
                    )
                )
            );

            var tbody = $('<tbody>');

            log.forEach(function(entry) {
                var statusEl = entry.status === 'success'
                    ? $('<span>').addClass('text-success').append($('<i>').addClass('glyphicon glyphicon-ok'))
                    : $('<span>').addClass('text-danger').append($('<i>').addClass('glyphicon glyphicon-remove'));

                var ts = entry.timestamp || '';
                if (ts) {
                    var d = new Date(ts);
                    ts = d.toLocaleDateString() + ' ' + d.toLocaleTimeString();
                }

                var row = $('<tr>');
                row.append($('<td>').text(ts));
                row.append($('<td>').append(statusEl));
                row.append($('<td>').text(entry.to || '-'));
                row.append($('<td>').text(entry.subject || '-'));
                row.append($('<td>').text(entry.duration + 's'));
                tbody.append(row);

                if (entry.error) {
                    var errorRow = $('<tr>').addClass('danger');
                    errorRow.append(
                        $('<td>').attr('colspan', 5).addClass('text-danger')
                            .css('font-size', '11px')
                            .text(entry.error)
                    );
                    tbody.append(errorRow);
                }
            });

            table.append(tbody);
        }

        $(document).on('click', '#m365-view-send-log-btn', function() {
            var logDiv = $('#m365-send-log');
            if (logDiv.is(':visible')) {
                logDiv.hide();
                if (sendLogInterval) { clearInterval(sendLogInterval); sendLogInterval = null; }
                return;
            }

            var btn = $(this);
            btn.prop('disabled', true);

            loadSendLog(function() {
                btn.prop('disabled', false);
                logDiv.show();
                sendLogInterval = setInterval(function() {
                    if (!logDiv.is(':visible')) { clearInterval(sendLogInterval); sendLogInterval = null; return; }
                    loadSendLog();
                }, 30000);
            });
        });
    }

    // ========== RETRY QUEUE ==========

    function loadRetryQueue() {
        $.ajax({
            url: '/mail365/retry-queue/' + getMailboxId(),
            type: 'GET',
            dataType: 'json',
            success: function(response) {
                if (response.status !== 'success') return;

                var queue = response.queue || [];
                var content = $('#m365-retry-queue-content');
                content.empty();

                if (!queue.length) {
                    content.append($('<p>').addClass('form-control-static text-muted').text('No messages queued for retry.'));
                    $('#m365-retry-queue-actions').hide();
                    $('#m365-retry-queue').show();
                    return;
                }

                var table = $('<table>').addClass('table table-condensed table-bordered').css('font-size', '12px');
                table.append(
                    $('<thead>').append(
                        $('<tr>').append(
                            $('<th>').text('Queued'),
                            $('<th>').text('To'),
                            $('<th>').text('Subject'),
                            $('<th>').text('Attempts'),
                            $('<th>').text('Last Error')
                        )
                    )
                );

                var tbody = $('<tbody>');
                queue.forEach(function(item) {
                    var ts = item.queued_at || '';
                    if (ts) {
                        var d = new Date(ts);
                        ts = d.toLocaleDateString() + ' ' + d.toLocaleTimeString();
                    }

                    var row = $('<tr>');
                    row.append($('<td>').text(ts));
                    row.append($('<td>').text(item.to || '-'));
                    row.append($('<td>').text(item.subject || '-'));
                    row.append($('<td>').text(item.attempts));
                    row.append($('<td>').addClass('text-danger').css('font-size', '11px').text(item.last_error || '-'));
                    tbody.append(row);
                });

                table.append(tbody);
                content.append(table);
                $('#m365-retry-queue-actions').show();
                $('#m365-retry-queue').show();
            }
        });
    }

    function bindRetryQueue() {
        $(document).on('click', '#m365-clear-retry-queue-btn', function() {
            if (!confirm('Discard all queued messages? They will not be retried.')) return;

            var btn = $(this);
            btn.prop('disabled', true);

            $.ajax({
                url: '/mail365/retry-queue/' + getMailboxId() + '/clear',
                type: 'POST',
                data: { _token: $('input[name="_token"]').val() },
                dataType: 'json',
                success: function(response) {
                    if (response.status === 'success') {
                        showResult('#m365-retry-queue-status', 'success', response.msg);
                        loadRetryQueue();
                    } else {
                        showResult('#m365-retry-queue-status', 'error', response.msg || 'Failed to clear queue');
                    }
                },
                error: function() {
                    showResult('#m365-retry-queue-status', 'error', 'Request failed');
                },
                complete: function() {
                    btn.prop('disabled', false);
                }
            });
        });
    }

    // ========== QUOTA CHECK ==========

    function bindQuotaCheck() {
        $(document).on('click', '#m365-check-quota-btn', function() {
            var btn = $(this);
            btn.button('loading');

            $.ajax({
                url: '/mail365/quota/' + getMailboxId(),
                type: 'GET',
                dataType: 'json',
                success: function(response) {
                    var info = $('#m365-quota-info');
                    info.empty();
                    if (response.status === 'success' && response.quota) {
                        info.append(
                            $('<p>').addClass('form-control-static').append(
                                $('<strong>').text(response.quota.used_display),
                                ' used ',
                                $('<span>').addClass('text-muted').text('(' + response.quota.folder_count + ' folders)'),
                                ' — checked ',
                                $('<span>').addClass('text-muted').text('just now')
                            )
                        );
                    } else {
                        info.append($('<p>').addClass('form-control-static text-danger').text(response.msg || 'Failed'));
                    }
                },
                error: function() {
                    $('#m365-quota-info').empty().append(
                        $('<p>').addClass('form-control-static text-danger').text('Request failed')
                    );
                },
                complete: function() {
                    btn.button('reset');
                }
            });
        });
    }

    // ========== MAILBOX BROWSER ==========

    function bindMailboxBrowser() {
        var mailboxesCache = null;

        $(document).on('click', '#m365-browse-mailboxes-btn', function() {
            var btn = $(this);
            var picker = $('#m365-mailbox-picker');

            if (picker.is(':visible')) {
                picker.hide();
                return;
            }

            if (mailboxesCache) {
                renderMailboxList(mailboxesCache);
                picker.show();
                return;
            }

            btn.button('loading');

            $.ajax({
                url: '/mail365/mailboxes/' + getMailboxId(),
                type: 'GET',
                dataType: 'json',
                success: function(response) {
                    if (response.status !== 'success') {
                        showResult('#m365-in-test-result', 'error', response.msg || 'Failed to list mailboxes');
                        return;
                    }
                    mailboxesCache = response.mailboxes || [];
                    renderMailboxList(mailboxesCache);
                    picker.show();
                },
                error: function() {
                    showResult('#m365-in-test-result', 'error', 'Failed to browse mailboxes');
                },
                complete: function() {
                    btn.button('reset');
                }
            });
        });

        $(document).on('input', '#m365-mailbox-search', function() {
            var q = $(this).val().toLowerCase();
            if (!mailboxesCache) return;
            var filtered = mailboxesCache.filter(function(m) {
                return m.display_name.toLowerCase().indexOf(q) !== -1
                    || m.email.toLowerCase().indexOf(q) !== -1;
            });
            renderMailboxList(filtered);
        });

        $(document).on('click', '.m365-mailbox-item', function() {
            var email = $(this).data('email');
            $('#m365_shared_mailbox_email').val(email);
            $('#m365-mailbox-picker').hide();
        });

        $(document).on('click', '#m365-mailbox-picker-cancel', function() {
            $('#m365-mailbox-picker').hide();
        });
    }

    function renderMailboxList(mailboxes) {
        var container = $('#m365-mailbox-list');
        container.empty();

        if (!mailboxes.length) {
            container.append($('<p>').addClass('text-muted').css('padding', '8px').text('No mailboxes found.'));
            return;
        }

        mailboxes.forEach(function(m) {
            var item = $('<div>').addClass('m365-mailbox-item')
                .css({padding: '6px 8px', cursor: 'pointer', 'border-bottom': '1px solid #eee'})
                .attr('data-email', m.email)
                .append(
                    $('<strong>').text(m.display_name),
                    $('<br>'),
                    $('<small>').addClass('text-muted').text(m.email)
                );
            item.on('mouseenter', function() { $(this).css('background', '#f5f5f5'); });
            item.on('mouseleave', function() { $(this).css('background', ''); });
            container.append(item);
        });
    }

    // ========== OVERVIEW PAGE ==========

    function isOverviewPage() {
        return $('#mail365-overview-table').length > 0;
    }

    // ========== INIT ==========

    $(document).ready(function() {
        // Outgoing page
        if (isConnectionPage()) {
            injectRadioButton();
            injectOptionsSection();
            injectHiddenOutServer();
            bindTestButton();
            bindFormSubmit();
            bindSendLog();
            bindRetryQueue();

            var currentMethod = getCurrentMethod();
            if (currentMethod == MAIL365_METHOD) {
                loadSavedSettings();
                loadRetryQueue();
                $('#out_method_' + MAIL365_METHOD).trigger('change');
            }

            $(document).on('change', 'input[name="out_method"]', function() {
                var method = $('input[name="out_method"]:checked').val();
                if (method == MAIL365_METHOD) {
                    loadSavedSettings();
                    loadRetryQueue();
                }
            });
        }

        // Incoming page
        if (isIncomingPage()) {
            toggleIncomingSettings();
            bindIncomingTestButton();
            bindFolderPicker();
            bindPostFetchAction();
            bindFetchLog();
            bindQuotaCheck();
            bindMailboxBrowser();
            convertTimestamps();

            $(document).on('change', '#m365_secret_expiry_date', function() {
                updateExpiryStatus();
            });

            $(document).on('change', 'select[name="in_protocol"]', function() {
                toggleIncomingSettings();
            });

            // Pre-load move folders if action is already set to 'move'
            if ($('#m365_post_fetch_action').val() === 'move') {
                loadMoveFolders();
            }
        }

        // Overview page
        if (isOverviewPage()) {
            convertTimestamps();
        }
    });
})();

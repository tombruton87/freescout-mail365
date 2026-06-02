<div id="mail365_in_settings" class="margin-top" style="display:none;">
    <hr/>
    <div class="form-group">
        <div class="col-sm-6 col-sm-offset-2">
            <p class="text-help">
                Receive emails via the Microsoft Graph API (HTTPS) instead of IMAP.
                This bypasses IMAP port restrictions.
                You need an Azure AD app registration with <strong>Mail.ReadWrite</strong> application permission.
            </p>
            @if($tenantId && $clientId)
            <p class="text-muted"><i class="glyphicon glyphicon-info-sign"></i> Uses the same Azure app credentials as the outgoing settings. Changes here will also apply to sending.</p>
            @endif
        </div>
    </div>

    {{-- Tenant ID --}}
    <div class="form-group">
        <label for="m365_in_tenant_id" class="col-sm-2 control-label">Tenant ID</label>
        <div class="col-sm-6">
            <input id="m365_in_tenant_id" type="text" class="form-control input-sized"
                name="m365_tenant_id" value="{{ $tenantId }}" maxlength="255"
                placeholder="xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" autocomplete="off">
            <p class="form-help">Azure Active Directory &rarr; Properties &rarr; Tenant ID</p>
        </div>
    </div>

    {{-- Client ID --}}
    <div class="form-group">
        <label for="m365_in_client_id" class="col-sm-2 control-label">Client ID</label>
        <div class="col-sm-6">
            <input id="m365_in_client_id" type="text" class="form-control input-sized"
                name="m365_client_id" value="{{ $clientId }}" maxlength="255"
                placeholder="xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" autocomplete="off">
            <p class="form-help">Azure AD &rarr; App registrations &rarr; Your app &rarr; Application (client) ID</p>
        </div>
    </div>

    {{-- Client Secret --}}
    <div class="form-group">
        <label for="m365_in_client_secret" class="col-sm-2 control-label">Client Secret</label>
        <div class="col-sm-6">
            <input id="m365_in_client_secret" type="password" class="form-control input-sized"
                name="m365_client_secret" value="" maxlength="255" autocomplete="new-password"
                @if($hasSecret) placeholder="••••••••••••••••" @endif>
            <p class="form-help">Azure AD &rarr; App registrations &rarr; Your app &rarr; Certificates &amp; secrets</p>
        </div>
    </div>

    {{-- Test Connection --}}
    <div class="form-group">
        <div class="col-sm-6 col-sm-offset-2">
            <button type="button" id="m365-in-test-btn" class="btn btn-default" data-loading-text="Testing&hellip;">
                <i class="glyphicon glyphicon-ok-circle"></i> Test Connection
            </button>
            <span id="m365-in-test-result" class="margin-left"></span>
        </div>
    </div>

    {{-- Shared Mailbox --}}
    <div class="form-group">
        <label for="m365_shared_mailbox_email" class="col-sm-2 control-label">Shared Mailbox</label>
        <div class="col-sm-6">
            <div class="input-group input-sized">
                <input id="m365_shared_mailbox_email" type="email" class="form-control"
                    name="m365_shared_mailbox_email" value="{{ $sharedEmail }}" maxlength="255"
                    placeholder="shared@example.com" autocomplete="off">
                <span class="input-group-btn">
                    <button type="button" id="m365-browse-mailboxes-btn" class="btn btn-default"
                        data-loading-text="Loading&hellip;" title="Browse available mailboxes">
                        <i class="glyphicon glyphicon-search"></i> Browse
                    </button>
                </span>
            </div>
            <p class="form-help">Optional. Fetch emails from a shared mailbox instead of the mailbox address. Leave blank to use the default mailbox address. The Azure app must have permission to access this mailbox.</p>
            <div id="m365-mailbox-picker" style="display:none;margin-top:8px;">
                <input type="text" id="m365-mailbox-search" class="form-control input-sm"
                    placeholder="Filter by name or email..." style="margin-bottom:6px;">
                <div id="m365-mailbox-list" style="max-height:250px;overflow-y:auto;border:1px solid #ddd;border-radius:4px;padding:4px;"></div>
                <div style="margin-top:6px;">
                    <button type="button" id="m365-mailbox-picker-cancel" class="btn btn-default btn-xs">Cancel</button>
                </div>
            </div>
        </div>
    </div>

    {{-- Fetch Mode --}}
    <div class="form-group">
        <label for="m365_fetch_mode" class="col-sm-2 control-label">Fetch Mode</label>
        <div class="col-sm-6">
            <select id="m365_fetch_mode" name="m365_fetch_mode" class="form-control input-sized">
                <option value="all"{{ $fetchMode === 'all' ? ' selected' : '' }}>All new messages</option>
                <option value="unread"{{ $fetchMode === 'unread' ? ' selected' : '' }}>Unread messages only</option>
            </select>
            <p class="form-help">"All new messages" fetches every message by timestamp, even if already read in Outlook. "Unread only" skips messages you have already opened.</p>
        </div>
    </div>

    {{-- Post-Fetch Action --}}
    <div class="form-group">
        <label for="m365_post_fetch_action" class="col-sm-2 control-label">After Fetch</label>
        <div class="col-sm-6">
            <select id="m365_post_fetch_action" name="m365_post_fetch_action" class="form-control input-sized">
                <option value="none"{{ $postAction === 'none' ? ' selected' : '' }}>Do nothing</option>
                <option value="mark_read"{{ $postAction === 'mark_read' ? ' selected' : '' }}>Mark as read</option>
                <option value="move"{{ $postAction === 'move' ? ' selected' : '' }}>Move to folder</option>
            </select>
            <div id="m365-move-folder-wrapper" style="margin-top:8px;{{ $postAction !== 'move' ? 'display:none;' : '' }}">
                <select id="m365_post_fetch_move_folder" name="m365_post_fetch_move_folder" class="form-control input-sized">
                    <option value="">Select a folder&hellip;</option>
                    @if($moveFolderId)
                    <option value="{{ $moveFolderId }}" selected>{{ $moveFolderName ?: $moveFolderId }}</option>
                    @endif
                </select>
                <input type="hidden" id="m365_post_fetch_move_folder_name" name="m365_post_fetch_move_folder_name" value="{{ $moveFolderName }}">
            </div>
            <p class="form-help">Action to take on messages in Microsoft 365 after they are fetched. In "Unread only" mode, messages are always marked as read regardless of this setting.</p>
        </div>
    </div>

    {{-- Folders --}}
    <div class="form-group">
        <label class="col-sm-2 control-label">Folders</label>
        <div class="col-sm-6">
            <div id="m365-folder-list">
                @if(empty($fetchFolders))
                <span class="text-muted">Inbox only (default)</span>
                @else
                    @foreach($fetchFolders as $f)
                    <span class="label label-default m365-folder-tag" style="display:inline-block;margin:2px 4px 2px 0;padding:4px 8px;font-size:12px;">{{ $f['name'] }}</span>
                    @endforeach
                @endif
            </div>
            <div style="margin-top:8px;">
                <button type="button" id="m365-detect-folders-btn" class="btn btn-default btn-xs" data-loading-text="Loading&hellip;">
                    <i class="glyphicon glyphicon-folder-open"></i> Manage Folders
                </button>
            </div>
            <p class="form-help">By default only the Inbox is monitored. Click "Manage Folders" to detect and select additional folders.</p>
        </div>
    </div>

    <div id="m365-folder-picker" style="display:none;">
        <div class="form-group">
            <div class="col-sm-6 col-sm-offset-2">
                <div id="m365-folder-checkboxes" style="max-height:250px;overflow-y:auto;border:1px solid #ddd;border-radius:4px;padding:8px;"></div>
                <div style="margin-top:8px;">
                    <button type="button" id="m365-save-folders-btn" class="btn btn-primary btn-xs">
                        <i class="glyphicon glyphicon-ok"></i> Save Folder Selection
                    </button>
                    <button type="button" id="m365-cancel-folders-btn" class="btn btn-default btn-xs margin-left">
                        Cancel
                    </button>
                </div>
            </div>
        </div>
    </div>

    {{-- Secret Expiry Date --}}
    <div class="form-group">
        <label for="m365_secret_expiry_date" class="col-sm-2 control-label">Secret Expiry</label>
        <div class="col-sm-6">
            <div class="form-inline">
                <input id="m365_secret_expiry_date" type="date" class="form-control input-sm"
                    name="m365_secret_expiry_date" value="{{ $expiryDate }}" style="width:170px;">
                <span id="m365-secret-expiry-status" class="margin-left">
                    @if($expiryDate)
                        @if($daysLeft < 0)
                        <span class="text-danger"><i class="glyphicon glyphicon-exclamation-sign"></i> Expired</span>
                        @elseif($daysLeft <= 14)
                        <span class="text-danger"><i class="glyphicon glyphicon-exclamation-sign"></i> {{ $daysLeft }} day{{ $daysLeft !== 1 ? 's' : '' }} remaining</span>
                        @elseif($daysLeft <= 30)
                        <span class="text-warning"><i class="glyphicon glyphicon-warning-sign"></i> {{ $daysLeft }} day{{ $daysLeft !== 1 ? 's' : '' }} remaining</span>
                        @else
                        <span class="text-success"><i class="glyphicon glyphicon-ok"></i> {{ $daysLeft }} days remaining</span>
                        @endif
                    @endif
                </span>
            </div>
            <p class="form-help">Enter the expiry date of your Azure client secret (shown in the Azure portal when you create it). Auto-filled if Test Connection can detect it.</p>
        </div>
    </div>

    {{-- Expiry Alert --}}
    <div class="form-group">
        <label class="col-sm-2 control-label">Expiry Alert</label>
        <div class="col-sm-6">
            <div class="checkbox">
                <label>
                    <input type="checkbox" name="m365_expiry_alert_enabled" value="1"{{ $expiryAlertEnabled ? ' checked' : '' }}>
                    Create a reminder ticket when the client secret is expiring
                </label>
            </div>
            <div class="form-inline" style="margin-top:6px;">
                <label>Warn </label>
                <input type="number" name="m365_expiry_alert_days" class="form-control input-sm" style="width:70px;margin:0 4px;" value="{{ $expiryAlertDays }}" min="1" max="365">
                <label> days before expiry</label>
            </div>
            <p class="form-help">When enabled, a ticket will be created once per day in this mailbox reminding you to renew the Azure client secret.</p>
        </div>
    </div>

    {{-- Fetch Status --}}
    <div class="form-group">
        <label class="col-sm-2 control-label">Fetch Status</label>
        <div class="col-sm-6">
            @if($lastError && $lastErrorAt)
            <p class="form-control-static">
                <span class="text-danger"><i class="glyphicon glyphicon-remove"></i></span>
                Last error <span class="m365-time-ago" data-utc="{{ $lastErrorAt }}"></span>
            </p>
            <p class="text-danger" style="font-size:12px;margin-top:2px;">{{ mb_substr($lastError, 0, 200) }}</p>
            @endif

            @if($lastSuccess)
            <p class="form-control-static">
                <span class="text-success"><i class="glyphicon glyphicon-ok"></i></span>
                Last success <span class="m365-time-ago" data-utc="{{ $lastSuccess }}"></span>
                @if($lastCount !== null)
                &mdash; {{ (int) $lastCount }} message{{ $lastCount != 1 ? 's' : '' }} fetched
                @endif
            </p>
            @elseif(!$lastError && $lastRun)
            <p class="form-control-static">
                <span class="text-success"><i class="glyphicon glyphicon-ok"></i></span>
                Last checked <span class="m365-time-ago" data-utc="{{ $lastRun }}"></span>
                &mdash; {{ (int) $lastCount }} message{{ $lastCount != 1 ? 's' : '' }} imported
            </p>
            @elseif(!$lastRun)
            <p class="form-control-static text-muted">No fetch runs yet</p>
            @endif

            <div style="margin-top:6px;">
                <button type="button" id="m365-view-log-btn" class="btn btn-default btn-xs">
                    <i class="glyphicon glyphicon-list"></i> View Fetch Log
                </button>
            </div>
            <div id="m365-fetch-log" style="display:none;margin-top:8px;">
                <table class="table table-condensed table-bordered" id="m365-fetch-log-table" style="font-size:12px;"></table>
            </div>
        </div>
    </div>

    {{-- Mailbox Usage --}}
    <div class="form-group">
        <label class="col-sm-2 control-label">Mailbox Usage</label>
        <div class="col-sm-6">
            <div id="m365-quota-info">
                @if(!empty($quotaUsage))
                <p class="form-control-static">
                    <strong>{{ $quotaUsage['used_display'] }}</strong> used
                    <span class="text-muted">({{ $quotaUsage['folder_count'] }} folders)</span>
                    &mdash; checked <span class="m365-time-ago" data-utc="{{ $quotaUsage['checked_at'] }}"></span>
                </p>
                @else
                <p class="form-control-static text-muted">Not yet checked</p>
                @endif
            </div>
            <button type="button" id="m365-check-quota-btn" class="btn btn-default btn-xs" data-loading-text="Checking&hellip;">
                <i class="glyphicon glyphicon-stats"></i> Check Usage
            </button>
        </div>
    </div>

    <hr/>
</div>

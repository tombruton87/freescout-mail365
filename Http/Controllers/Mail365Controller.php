<?php

namespace Modules\Mail365\Http\Controllers;

use App\Mailbox;
use Illuminate\Foundation\Auth\Access\AuthorizesRequests;
use Illuminate\Http\Request;
use Illuminate\Routing\Controller;
use Modules\Mail365\Helpers\Mail365Client;
use Modules\Mail365\Providers\Mail365ServiceProvider;

class Mail365Controller extends Controller
{
    use AuthorizesRequests;

    const UUID_PATTERN = '/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i';

    public function save($mailbox_id, Request $request)
    {
        $mailbox = Mailbox::findOrFail($mailbox_id);
        $this->authorize('admin', $mailbox);

        $existing = $mailbox->getMeta(Mail365ServiceProvider::META_KEY, []);

        $authType = trim((string) $request->input('auth_type', $existing['auth_type'] ?? 'secret'));
        if (!in_array($authType, ['secret', 'certificate'])) {
            $authType = 'secret';
        }

        $newSecret = substr(trim((string) $request->input('client_secret', '')), 0, 255);

        $meta = array_merge($existing, [
            'tenant_id' => trim((string) $request->input('tenant_id', '')),
            'client_id' => trim((string) $request->input('client_id', '')),
            'auth_type' => $authType,
        ]);

        if ($newSecret) {
            $meta['client_secret'] = \Helper::encrypt($newSecret);
        } elseif ($request->input('keep_secret') && empty($existing['client_secret'])) {
            $meta['client_secret'] = '';
        }

        if (!$meta['tenant_id'] || !$meta['client_id']) {
            return response()->json([
                'status' => 'error',
                'msg'    => __('Tenant ID and Client ID are required.'),
            ]);
        }

        if ($authType === 'secret') {
            $hasSecret = !empty($meta['client_secret']);
            if (!$hasSecret) {
                return response()->json([
                    'status' => 'error',
                    'msg'    => __('All fields are required: Tenant ID, Client ID, and Client Secret.'),
                ]);
            }
        } elseif ($authType === 'certificate') {
            $hasCert = !empty($meta['certificate_pem']);
            if (!$hasCert) {
                return response()->json([
                    'status' => 'error',
                    'msg'    => __('Please upload a certificate before saving.'),
                ]);
            }
        }

        if (!preg_match(self::UUID_PATTERN, $meta['tenant_id']) || !preg_match(self::UUID_PATTERN, $meta['client_id'])) {
            return response()->json([
                'status' => 'error',
                'msg'    => __('Tenant ID and Client ID must be valid UUIDs.'),
            ]);
        }

        $mailbox->setMetaParam(Mail365ServiceProvider::META_KEY, $meta, true);

        return response()->json(['status' => 'success']);
    }

    public function load($mailbox_id)
    {
        $mailbox = Mailbox::findOrFail($mailbox_id);
        $this->authorize('admin', $mailbox);

        $meta = $mailbox->getMeta(Mail365ServiceProvider::META_KEY, []);

        return response()->json([
            'status'                  => 'success',
            'tenant_id'               => $meta['tenant_id'] ?? '',
            'client_id'               => $meta['client_id'] ?? '',
            'auth_type'               => $meta['auth_type'] ?? 'secret',
            'has_secret'              => !empty($meta['client_secret']),
            'has_certificate'         => !empty($meta['certificate_pem']),
            'certificate_thumbprint'  => $meta['certificate_thumbprint'] ?? '',
            'certificate_expiry'      => $meta['certificate_expiry'] ?? null,
            'secret_expiry'           => $meta['secret_expiry'] ?? null,
            'last_send_success'       => $meta['last_send_success'] ?? null,
            'last_send_error'         => $meta['last_send_error'] ?? null,
            'last_send_error_at'      => $meta['last_send_error_at'] ?? null,
        ]);
    }

    public function uploadCertificate($mailbox_id, Request $request)
    {
        $mailbox = Mailbox::findOrFail($mailbox_id);
        $this->authorize('admin', $mailbox);

        if (!$request->hasFile('certificate') || !$request->file('certificate')->isValid()) {
            return response()->json(['status' => 'error', 'msg' => __('No valid file uploaded.')]);
        }

        $file = $request->file('certificate');

        if ($file->getSize() > 32768) {
            return response()->json(['status' => 'error', 'msg' => __('File too large. Maximum 32 KB.')]);
        }

        $pem = file_get_contents($file->getRealPath());

        $errors = Mail365Client::validateCertificatePem($pem);
        if ($errors) {
            return response()->json(['status' => 'error', 'msg' => implode(' ', $errors)]);
        }

        $thumbprint = Mail365Client::extractCertificateThumbprint($pem);
        $expiry = Mail365Client::extractCertificateExpiry($pem);

        $meta = $mailbox->getMeta(Mail365ServiceProvider::META_KEY, []);
        $meta['auth_type'] = 'certificate';
        $meta['certificate_pem'] = \Helper::encrypt($pem);
        $meta['certificate_thumbprint'] = $thumbprint;
        if ($expiry) {
            $meta['certificate_expiry'] = $expiry;
        }
        $mailbox->setMetaParam(Mail365ServiceProvider::META_KEY, $meta, true);

        return response()->json([
            'status'      => 'success',
            'msg'         => __('Certificate uploaded and validated.'),
            'thumbprint'  => $thumbprint,
            'expiry'      => $expiry,
        ]);
    }

    public function removeCertificate($mailbox_id, Request $request)
    {
        $mailbox = Mailbox::findOrFail($mailbox_id);
        $this->authorize('admin', $mailbox);

        $meta = $mailbox->getMeta(Mail365ServiceProvider::META_KEY, []);
        $meta['auth_type'] = 'secret';
        $meta['certificate_pem'] = '';
        $meta['certificate_thumbprint'] = '';
        $meta['certificate_expiry'] = null;
        $mailbox->setMetaParam(Mail365ServiceProvider::META_KEY, $meta, true);

        return response()->json(['status' => 'success', 'msg' => __('Certificate removed.')]);
    }

    public function testConnection($mailbox_id, Request $request)
    {
        $mailbox = Mailbox::findOrFail($mailbox_id);
        $this->authorize('admin', $mailbox);

        $tenantId     = trim((string) $request->input('tenant_id', ''));
        $clientId     = trim((string) $request->input('client_id', ''));
        $clientSecret = trim((string) $request->input('client_secret', ''));
        $authType     = trim((string) $request->input('auth_type', 'secret'));

        $existing = $mailbox->getMeta(Mail365ServiceProvider::META_KEY, []);

        if (!$clientSecret && $authType === 'secret') {
            $clientSecret = !empty($existing['client_secret']) ? \Helper::decrypt($existing['client_secret']) : '';
        }

        $certificatePem = '';
        if ($authType === 'certificate') {
            $certificatePem = !empty($existing['certificate_pem']) ? \Helper::decrypt($existing['certificate_pem']) : '';
        }

        if (!$tenantId || !$clientId) {
            return response()->json([
                'status' => 'error',
                'msg'    => __('Please fill in Tenant ID and Client ID before testing.'),
            ]);
        }

        if ($authType === 'secret' && !$clientSecret) {
            return response()->json([
                'status' => 'error',
                'msg'    => __('Please provide a Client Secret before testing.'),
            ]);
        }

        if ($authType === 'certificate' && !$certificatePem) {
            return response()->json([
                'status' => 'error',
                'msg'    => __('Please upload a certificate before testing.'),
            ]);
        }

        if (!preg_match(self::UUID_PATTERN, $tenantId) || !preg_match(self::UUID_PATTERN, $clientId)) {
            return response()->json([
                'status' => 'error',
                'msg'    => __('Tenant ID and Client ID must be valid UUIDs.'),
            ]);
        }

        $client = new Mail365Client($tenantId, $clientId, $clientSecret, null, $authType, $certificatePem);

        try {
            $client->getAccessToken();
        } catch (\Exception $e) {
            \Log::error('Mail365 test auth failed: ' . $e->getMessage());
            return response()->json([
                'status' => 'error',
                'msg'    => __('Authentication failed. Check your credentials and try again.'),
            ]);
        }

        $result = [
            'status' => 'success',
            'msg'    => __('Successfully authenticated with Microsoft 365. Token obtained.'),
        ];

        // Permission pre-check
        $roles = $client->getTokenRoles();
        $missingPerms = [];
        if (!in_array('Mail.Send', $roles)) $missingPerms[] = 'Mail.Send';
        if (!in_array('Mail.ReadWrite', $roles)) $missingPerms[] = 'Mail.ReadWrite';
        if ($missingPerms) {
            $result['permission_warnings'] = 'Missing permissions: ' . implode(', ', $missingPerms) . '. Add these as Application permissions in Azure AD.';
        }

        // Secret expiry check
        $expiry = $client->checkSecretExpiry();
        if ($expiry) {
            $result['secret_expiry'] = $expiry;

            $meta = $mailbox->getMeta(Mail365ServiceProvider::META_KEY, []);
            $meta['secret_expiry'] = $expiry;
            $mailbox->setMetaParam(Mail365ServiceProvider::META_KEY, $meta, true);
        }

        // Mailbox verification
        $verifyEmail = trim((string) $request->input('verify_email', '')) ?: $mailbox->email;
        $verifyResult = $this->verifyMailboxEmail($client, $verifyEmail);

        if ($verifyResult === false) {
            $result['email_warning'] = __(':email was not found in Microsoft 365. Check that the address is correct and the app has Mail permissions for this mailbox.', ['email' => $verifyEmail]);
        } elseif ($verifyResult) {
            $result['msg'] .= ' ' . __('Mailbox verified: :name.', ['name' => $verifyResult]);
            $result['email_verified'] = $verifyResult;
        }

        return response()->json($result);
    }

    public function listFolders($mailbox_id)
    {
        $mailbox = Mailbox::findOrFail($mailbox_id);
        $this->authorize('admin', $mailbox);

        $meta = $mailbox->getMeta(Mail365ServiceProvider::META_KEY, []);
        $client = self::buildClientFromMeta($meta);

        if (!$client) {
            return response()->json([
                'status' => 'error',
                'msg'    => __('Azure credentials are not configured.'),
            ]);
        }

        try {
            $client->getAccessToken();
        } catch (\Exception $e) {
            return response()->json([
                'status' => 'error',
                'msg'    => __('Failed to authenticate with Microsoft 365.'),
            ]);
        }

        try {
            $folders = $this->fetchFoldersRecursive($client, $mailbox->email, null, 0, 3);
        } catch (\Exception $e) {
            \Log::error('Mail365 folder listing failed: ' . $e->getMessage());
            return response()->json([
                'status' => 'error',
                'msg'    => __('Failed to list folders. Check the application log for details.'),
            ]);
        }

        return response()->json([
            'status'   => 'success',
            'folders'  => $folders,
            'selected' => $meta['fetch_folders'] ?? [],
        ]);
    }

    /**
     * Recursively fetch mail folders up to $maxDepth levels deep.
     */
    protected function fetchFoldersRecursive(Mail365Client $client, $email, $parentId, $depth, $maxDepth)
    {
        if ($parentId === null) {
            $url = 'https://graph.microsoft.com/v1.0/users/' . urlencode($email)
                 . '/mailFolders?$top=100&$select=id,displayName,childFolderCount,totalItemCount';
        } else {
            $url = 'https://graph.microsoft.com/v1.0/users/' . urlencode($email)
                 . '/mailFolders/' . urlencode($parentId)
                 . '/childFolders?$top=100&$select=id,displayName,childFolderCount,totalItemCount';
        }

        $folders = [];

        while ($url) {
            $response = $client->graphGet($url);

            if ($response['status'] >= 400) {
                break;
            }

            foreach ($response['body']['value'] ?? [] as $folder) {
                $hasChildren = ($folder['childFolderCount'] ?? 0) > 0;

                $entry = [
                    'id'           => $folder['id'],
                    'name'         => $folder['displayName'],
                    'item_count'   => $folder['totalItemCount'] ?? 0,
                    'has_children' => $hasChildren,
                    'depth'        => $depth,
                    'children'     => [],
                ];

                if ($hasChildren && $depth < $maxDepth) {
                    $entry['children'] = $this->fetchFoldersRecursive(
                        $client, $email, $folder['id'], $depth + 1, $maxDepth
                    );
                }

                $folders[] = $entry;
            }

            $nextLink = $response['body']['@odata.nextLink'] ?? null;
            $url = ($nextLink && parse_url($nextLink, PHP_URL_HOST) === 'graph.microsoft.com') ? $nextLink : null;
        }

        return $folders;
    }

    public function fetchLog($mailbox_id)
    {
        $mailbox = Mailbox::findOrFail($mailbox_id);
        $this->authorize('admin', $mailbox);

        $meta = $mailbox->getMeta(Mail365ServiceProvider::META_KEY, []);

        return response()->json([
            'status'        => 'success',
            'log'           => $meta['fetch_log'] ?? [],
            'last_success'  => $meta['last_fetch_success'] ?? null,
            'last_error'    => $meta['last_fetch_error'] ?? null,
            'last_error_at' => $meta['last_fetch_error_at'] ?? null,
        ]);
    }

    public function saveFolders($mailbox_id, Request $request)
    {
        $mailbox = Mailbox::findOrFail($mailbox_id);
        $this->authorize('admin', $mailbox);

        $folders = $request->input('folders', []);
        if (!is_array($folders)) {
            $folders = [];
        }

        $sanitized = [];
        foreach ($folders as $folder) {
            if (!empty($folder['id']) && !empty($folder['name'])) {
                $sanitized[] = [
                    'id'   => substr(trim((string) $folder['id']), 0, 500),
                    'name' => substr(trim((string) $folder['name']), 0, 255),
                ];
            }
        }

        $meta = $mailbox->getMeta(Mail365ServiceProvider::META_KEY, []);
        $meta['fetch_folders'] = $sanitized;
        $mailbox->setMetaParam(Mail365ServiceProvider::META_KEY, $meta, true);

        return response()->json(['status' => 'success']);
    }

    public function sendLog($mailbox_id)
    {
        $mailbox = Mailbox::findOrFail($mailbox_id);
        $this->authorize('admin', $mailbox);

        $meta = $mailbox->getMeta(Mail365ServiceProvider::META_KEY, []);

        return response()->json([
            'status' => 'success',
            'log'    => $meta['send_log'] ?? [],
        ]);
    }

    public function retryQueue($mailbox_id)
    {
        $mailbox = Mailbox::findOrFail($mailbox_id);
        $this->authorize('admin', $mailbox);

        $meta = $mailbox->getMeta(Mail365ServiceProvider::META_KEY, []);
        $queue = $meta['retry_queue'] ?? [];

        $items = array_map(function ($item) {
            return [
                'subject'      => $item['subject'] ?? '(no subject)',
                'to'           => $item['to'] ?? '',
                'queued_at'    => $item['queued_at'] ?? '',
                'attempts'     => $item['attempts'] ?? 0,
                'last_error'   => $item['last_error'] ?? null,
                'last_attempt' => $item['last_attempt'] ?? null,
            ];
        }, $queue);

        return response()->json([
            'status' => 'success',
            'queue'  => $items,
        ]);
    }

    public function clearRetryQueue($mailbox_id, Request $request)
    {
        $mailbox = Mailbox::findOrFail($mailbox_id);
        $this->authorize('admin', $mailbox);

        $meta = $mailbox->getMeta(Mail365ServiceProvider::META_KEY, []);
        $count = count($meta['retry_queue'] ?? []);
        $meta['retry_queue'] = [];
        $mailbox->setMetaParam(Mail365ServiceProvider::META_KEY, $meta, true);

        return response()->json([
            'status' => 'success',
            'msg'    => $count . ' queued message' . ($count !== 1 ? 's' : '') . ' cleared.',
        ]);
    }

    public function quota($mailbox_id)
    {
        $mailbox = Mailbox::findOrFail($mailbox_id);
        $this->authorize('admin', $mailbox);

        $meta = $mailbox->getMeta(Mail365ServiceProvider::META_KEY, []);
        $client = self::buildClientFromMeta($meta);

        if (!$client) {
            return response()->json(['status' => 'error', 'msg' => 'Azure credentials not configured.']);
        }

        try {
            $client->getAccessToken();
        } catch (\Exception $e) {
            return response()->json(['status' => 'error', 'msg' => 'Authentication failed.']);
        }

        $fetchEmail = !empty($meta['shared_mailbox_email']) ? $meta['shared_mailbox_email'] : $mailbox->email;
        $quota = $client->getMailboxQuota($fetchEmail);

        if (!$quota) {
            return response()->json(['status' => 'error', 'msg' => 'Could not retrieve mailbox usage.']);
        }

        $meta['quota_usage'] = $quota;
        $mailbox->setMetaParam(Mail365ServiceProvider::META_KEY, $meta, true);

        return response()->json(['status' => 'success', 'quota' => $quota]);
    }

    public function listMailboxes($mailbox_id)
    {
        $mailbox = Mailbox::findOrFail($mailbox_id);
        $this->authorize('admin', $mailbox);

        $meta = $mailbox->getMeta(Mail365ServiceProvider::META_KEY, []);
        $client = self::buildClientFromMeta($meta);

        if (!$client) {
            return response()->json(['status' => 'error', 'msg' => 'Azure credentials not configured.']);
        }

        try {
            $client->getAccessToken();
        } catch (\Exception $e) {
            return response()->json(['status' => 'error', 'msg' => 'Authentication failed.']);
        }

        $result = $client->listMailboxes();

        if ($result['error'] === 'missing_permission') {
            return response()->json([
                'status' => 'error',
                'msg'    => 'The Azure app does not have User.Read.All permission. Add this as an Application permission in Azure AD to use the Browse feature.',
            ]);
        }

        if ($result['error']) {
            return response()->json(['status' => 'error', 'msg' => $result['error']]);
        }

        return response()->json([
            'status'    => 'success',
            'mailboxes' => $result['mailboxes'],
        ]);
    }

    public function overview()
    {
        if (!\Auth::user()->isAdmin()) {
            \Helper::denyAccess();
        }

        $mailboxes = Mailbox::all();
        $rows = [];

        foreach ($mailboxes as $mailbox) {
            $meta = $mailbox->getMeta(Mail365ServiceProvider::META_KEY, []);
            if (empty($meta['tenant_id'])) continue;

            $authType = $meta['auth_type'] ?? 'secret';
            if ($authType === 'certificate') {
                $expiryDate = $meta['certificate_expiry']['date'] ?? null;
            } else {
                $expiryDate = $meta['secret_expiry']['date'] ?? ($meta['secret_expiry_date'] ?? null);
            }
            $daysLeft = $expiryDate ? (int) ceil((strtotime($expiryDate) - time()) / 86400) : null;

            $rows[] = [
                'mailbox'            => $mailbox,
                'auth_type'          => $authType,
                'last_fetch_success' => $meta['last_fetch_success'] ?? null,
                'last_fetch_error'   => $meta['last_fetch_error'] ?? null,
                'last_fetch_error_at' => $meta['last_fetch_error_at'] ?? null,
                'last_send_success'  => $meta['last_send_success'] ?? null,
                'last_send_error'    => $meta['last_send_error'] ?? null,
                'secret_expiry_date' => $expiryDate,
                'secret_days_left'   => $daysLeft,
                'retry_queue_count'  => count($meta['retry_queue'] ?? []),
                'shared_email'       => $meta['shared_mailbox_email'] ?? null,
                'quota_usage'        => $meta['quota_usage'] ?? null,
            ];
        }

        return view('mail365::overview', ['rows' => $rows]);
    }

    protected static function buildClientFromMeta(array $meta)
    {
        $tenantId     = $meta['tenant_id'] ?? '';
        $clientId     = $meta['client_id'] ?? '';
        $authType     = $meta['auth_type'] ?? 'secret';
        $clientSecret = '';
        $certificatePem = '';

        if ($authType === 'certificate') {
            $certificatePem = !empty($meta['certificate_pem']) ? \Helper::decrypt($meta['certificate_pem']) : '';
            if (!$tenantId || !$clientId || !$certificatePem) {
                return null;
            }
        } else {
            $clientSecret = !empty($meta['client_secret']) ? \Helper::decrypt($meta['client_secret']) : '';
            if (!$tenantId || !$clientId || !$clientSecret) {
                return null;
            }
        }

        return new Mail365Client($tenantId, $clientId, $clientSecret, null, $authType, $certificatePem);
    }

    /**
     * Verify that the mailbox email exists in Microsoft 365.
     */
    protected function verifyMailboxEmail(Mail365Client $client, $email)
    {
        $url = 'https://graph.microsoft.com/v1.0/users/' . urlencode($email) . '/mailFolders/inbox?$select=id,displayName';

        try {
            $response = $client->graphGet($url);
        } catch (\Exception $e) {
            return false;
        }

        if ($response['status'] >= 400) {
            return false;
        }

        return $email;
    }
}

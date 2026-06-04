<?php

namespace Modules\Mail365\Console;

use App\Conversation;
use App\Customer;
use App\Mailbox;
use App\Option;
use App\Subscription;
use App\Thread;
use Illuminate\Console\Command;
use Modules\Mail365\Helpers\Mail365Client;
use Modules\Mail365\Providers\Mail365ServiceProvider;
use Symfony\Component\Console\Input\ArrayInput;
use Symfony\Component\Console\Output\OutputInterface;

class Mail365FetchCommand extends Command
{
    protected $signature = 'mail365:fetch-emails {--debug=0}';

    protected $description = 'Fetch emails from Microsoft 365 Graph API for mailboxes using Graph365 incoming protocol';

    /** @var Mail365Client|null */
    protected $client;

    public function handle()
    {
        $lockFile = storage_path('framework/mail365-fetch.lock');
        $fp = fopen($lockFile, 'c');
        if (!$fp || !flock($fp, LOCK_EX | LOCK_NB)) {
            $this->line('['.date('Y-m-d H:i:s').'] Mail365: another fetch is already running, skipping.');
            if ($fp) fclose($fp);
            return;
        }

        try {
            $this->doHandle();
        } finally {
            flock($fp, LOCK_UN);
            fclose($fp);
        }
    }

    protected function doHandle()
    {
        $debug = $this->option('debug');
        $this->line('['.date('Y-m-d H:i:s').'] Mail365: fetching incoming emails via Graph API');

        $mailboxes = Mailbox::where('in_protocol', Mail365ServiceProvider::IN_PROTOCOL_GRAPH365)->get();

        if ($mailboxes->isEmpty()) {
            $this->line('['.date('Y-m-d H:i:s').'] No mailboxes configured for Graph API incoming.');
            return;
        }

        $fetchCommand = new Graph365FetchEmails();
        $fetchCommand->setLaravel($this->laravel);
        $fetchCommand->initOutput($this->output);
        $fetchCommand->mailboxes = $mailboxes;

        foreach ($mailboxes as $mailbox) {
            $meta = $mailbox->getMeta(Mail365ServiceProvider::META_KEY, []);
            $tenantId       = $meta['tenant_id'] ?? '';
            $clientId       = $meta['client_id'] ?? '';
            $authType       = $meta['auth_type'] ?? 'secret';
            $clientSecret   = '';
            $certificatePem = '';

            if ($authType === 'certificate') {
                $certificatePem = !empty($meta['certificate_pem']) ? \Helper::decrypt($meta['certificate_pem']) : '';
                if (!$tenantId || !$clientId || !$certificatePem) {
                    $this->error('['.date('Y-m-d H:i:s').'] Mail365: missing Azure certificate for ' . $mailbox->email);
                    continue;
                }
            } else {
                $clientSecret = !empty($meta['client_secret']) ? \Helper::decrypt($meta['client_secret']) : '';
                if (!$tenantId || !$clientId || !$clientSecret) {
                    $this->error('['.date('Y-m-d H:i:s').'] Mail365: missing Azure credentials for ' . $mailbox->email);
                    continue;
                }
            }

            if (!\Eventy::filter('mailbox.in_active', null, $mailbox)) {
                if ($debug) {
                    $this->line('['.date('Y-m-d H:i:s').'] Skipping inactive mailbox: ' . $mailbox->email);
                }
                continue;
            }

            $this->info('['.date('Y-m-d H:i:s').'] Mailbox: ' . $mailbox->name . ' (' . $mailbox->email . ')');

            $this->client = new Mail365Client($tenantId, $clientId, $clientSecret, function($msg) {
                $this->line($msg);
            }, $authType, $certificatePem);

            $fetchCommand->mailbox = $mailbox;
            $fetchCommand->extra_import = [];

            try {
                $this->fetchViaGraph($mailbox, $fetchCommand, $debug);
            } catch (\Exception $e) {
                $this->error('['.date('Y-m-d H:i:s').'] Error: ' . $e->getMessage());
                \Log::error('Mail365 fetch error for ' . $mailbox->email . ': ' . $e->getMessage());
            }

            if (count($fetchCommand->extra_import)) {
                $this->line('['.date('Y-m-d H:i:s').'] Importing emails sent to several mailboxes: ' . count($fetchCommand->extra_import));
                foreach ($fetchCommand->extra_import as $i => $extra) {
                    $this->line('['.date('Y-m-d H:i:s').'] ' . ($i + 1) . ') ' . $extra['message']->getSubject());
                    $fetchCommand->processMessage($extra['message'], $extra['message_id'], $extra['mailbox'], [], true);
                }
            }
        }

        Subscription::processEvents();

        $this->info('['.date('Y-m-d H:i:s').'] Mail365: fetch finished');
    }

    protected function fetchViaGraph(Mailbox $mailbox, $fetchCommand, $debug)
    {
        $client = $this->client;
        $startTime = microtime(true);
        $meta = $mailbox->getMeta(Mail365ServiceProvider::META_KEY, []);
        $fetched = 0;
        $folderNames = [];

        try {
            $client->getAccessToken();
            $fetchMode = $meta['fetch_mode'] ?? 'all';

            $fetchEmail = !empty($meta['shared_mailbox_email']) ? $meta['shared_mailbox_email'] : $mailbox->email;
            $postAction = $meta['post_fetch_action'] ?? 'none';
            $moveFolderId = $meta['post_fetch_move_folder'] ?? '';

            $fetchFolders = $meta['fetch_folders'] ?? [];
            if (empty($fetchFolders)) {
                $fetchFolders = [['id' => 'inbox', 'name' => 'Inbox']];
            }
            $folderNames = array_column($fetchFolders, 'name');

            $latestReceived = null;
            $userBase = 'https://graph.microsoft.com/v1.0/users/' . urlencode($fetchEmail);

            if ($fetchMode === 'unread') {
                $filter = 'isRead eq false';
            } else {
                $lastFetched = $meta['last_fetched_at'] ?? null;
                if (!$lastFetched || !preg_match('/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}Z$/', $lastFetched)) {
                    $lastFetched = gmdate('Y-m-d\TH:i:s\Z', strtotime('-3 days'));
                }
                $filter = "receivedDateTime ge $lastFetched";
            }

            if ($debug) {
                $this->line('['.date('Y-m-d H:i:s').'] Fetch mode: ' . $fetchMode . ', filter: ' . $filter);
                $this->line('['.date('Y-m-d H:i:s').'] Folders: ' . implode(', ', $folderNames));
                if ($fetchEmail !== $mailbox->email) {
                    $this->line('['.date('Y-m-d H:i:s').'] Fetching from shared mailbox: ' . $fetchEmail);
                }
                if ($postAction !== 'none') {
                    $this->line('['.date('Y-m-d H:i:s').'] Post-fetch action: ' . $postAction);
                }
            }

            foreach ($fetchFolders as $folder) {
                $folderId   = $folder['id'];
                $folderName = $folder['name'];

                $baseUrl = $userBase . '/mailFolders/' . urlencode($folderId) . '/messages';

                $url = $baseUrl
                     . '?$filter=' . urlencode($filter)
                     . '&$select=id,internetMessageId,subject,receivedDateTime'
                     . '&$top=50'
                     . '&$orderby=' . urlencode('receivedDateTime asc');

                if (count($fetchFolders) > 1 || $folderId !== 'inbox') {
                    $this->line('['.date('Y-m-d H:i:s').'] Checking folder: ' . $folderName);
                }

                while ($url) {
                    $response = $client->graphGet($url);

                    if ($response['status'] >= 400) {
                        $error = $response['body']['error']['message'] ?? 'Unknown error';
                        $code  = $response['body']['error']['code'] ?? '';
                        $this->error('['.date('Y-m-d H:i:s').'] Graph API error for folder ' . $folderName . ": ({$response['status']} {$code}): {$error}");
                        break;
                    }

                    $messages = $response['body']['value'] ?? [];
                    $nextLink = $response['body']['@odata.nextLink'] ?? null;

                    $this->line('['.date('Y-m-d H:i:s').'] Fetched batch: ' . count($messages));

                    foreach ($messages as $graphMsg) {
                        $graphId   = $graphMsg['id'];
                        $messageId = $graphMsg['internetMessageId'] ?? '';
                        $subject   = $graphMsg['subject'] ?? '(no subject)';

                        if ($messageId && $this->isAlreadyImported($meta, $messageId)) {
                            if ($debug) {
                                $this->line('['.date('Y-m-d H:i:s').'] Skipping already-imported: ' . $messageId);
                            }
                            $this->applyPostFetchAction($fetchEmail, $graphId, $postAction, $moveFolderId, $fetchMode);
                            if ($graphMsg['receivedDateTime'] ?? null) {
                                $latestReceived = $graphMsg['receivedDateTime'];
                            }
                            continue;
                        }

                        $fetched++;
                        $this->line('['.date('Y-m-d H:i:s').'] ' . $fetched . ') ' . $subject);

                        $receivedAt = $graphMsg['receivedDateTime'] ?? null;

                        try {
                            $mimeUrl = 'https://graph.microsoft.com/v1.0/users/' . urlencode($fetchEmail)
                                     . '/messages/' . urlencode($graphId) . '/$value';
                            $rawResult = $client->graphRaw($mimeUrl, 120);

                            if ($rawResult['status'] >= 400) {
                                throw new \Exception("Failed to download MIME (HTTP {$rawResult['status']})");
                            }

                            $rawMime = $rawResult['body'];

                            if (!$rawMime) {
                                $this->error('['.date('Y-m-d H:i:s').'] Empty MIME for message ' . $graphId);
                                $this->applyPostFetchAction($fetchEmail, $graphId, $postAction, $moveFolderId, $fetchMode);
                                continue;
                            }

                            if (empty(\Webklex\PHPIMAP\ClientManager::get('options'))) {
                                new \Webklex\PHPIMAP\ClientManager(config('imap', []));
                            }

                            $message = \Webklex\PHPIMAP\Message::fromString($rawMime);

                            $fetchCommand->processMessage($message, $messageId, $mailbox, $fetchCommand->mailboxes);

                            if ($messageId) {
                                $this->trackImported($meta, $messageId);
                            }

                            $this->applyPostFetchAction($fetchEmail, $graphId, $postAction, $moveFolderId, $fetchMode);

                            if ($receivedAt) {
                                $latestReceived = $receivedAt;
                            }

                        } catch (\Exception $e) {
                            if (\Str::startsWith($e->getMessage(), 'SQLSTATE[23000]: Integrity constraint violation: 1062 Duplicate entry')) {
                                if ($debug) {
                                    $this->line('['.date('Y-m-d H:i:s').'] Duplicate, skipping: ' . $messageId);
                                }
                                if ($messageId) {
                                    $this->trackImported($meta, $messageId);
                                }
                                $this->applyPostFetchAction($fetchEmail, $graphId, $postAction, $moveFolderId, $fetchMode);
                                if ($receivedAt) {
                                    $latestReceived = $receivedAt;
                                }
                            } else {
                                $this->error('['.date('Y-m-d H:i:s').'] Error processing message: ' . $e->getMessage());
                                \Log::error('Mail365 processMessage error', [
                                    'mailbox'    => $mailbox->email,
                                    'message_id' => $messageId,
                                    'folder'     => $folderName,
                                    'error'      => $e->getMessage(),
                                ]);
                            }
                        }
                    }

                    $url = ($nextLink && parse_url($nextLink, PHP_URL_HOST) === 'graph.microsoft.com') ? $nextLink : null;
                }
            }

            if ($latestReceived && $fetchMode === 'all') {
                $nextCursor = gmdate('Y-m-d\TH:i:s\Z', strtotime($latestReceived) + 1);
                $meta['last_fetched_at'] = $nextCursor;
            }

            $meta['last_fetch_success'] = gmdate('Y-m-d\TH:i:s\Z');
            $meta['last_fetch_error'] = null;
            $meta['last_fetch_error_at'] = null;

            $lastQuotaCheck = $meta['quota_usage']['checked_at'] ?? null;
            $quotaAge = $lastQuotaCheck ? (time() - strtotime($lastQuotaCheck)) : PHP_INT_MAX;
            if ($quotaAge > 3600) {
                $quota = $client->getMailboxQuota($fetchEmail);
                if ($quota) {
                    $meta['quota_usage'] = $quota;
                }
            }

            $authType = $meta['auth_type'] ?? 'secret';
            if ($authType === 'certificate') {
                $certPem = !empty($meta['certificate_pem']) ? \Helper::decrypt($meta['certificate_pem']) : '';
                if ($certPem) {
                    $certExpiry = \Modules\Mail365\Helpers\Mail365Client::extractCertificateExpiry($certPem);
                    if ($certExpiry) {
                        $meta['certificate_expiry'] = $certExpiry;
                        $this->maybeCreateExpiryAlert($mailbox, $meta, $certExpiry);
                    }
                }
            } else {
                $expiry = $client->checkSecretExpiry();
                if ($expiry) {
                    $meta['secret_expiry'] = $expiry;
                }

                $effectiveExpiry = $expiry ?: $this->buildExpiryFromManualDate($meta);
                if ($effectiveExpiry) {
                    $this->maybeCreateExpiryAlert($mailbox, $meta, $effectiveExpiry);
                }
            }

            $this->addFetchLog($meta, 'success', $fetched, $folderNames, null, $startTime);

        } catch (\Exception $e) {
            $meta['last_fetch_error'] = mb_substr($e->getMessage(), 0, 500);
            $meta['last_fetch_error_at'] = gmdate('Y-m-d\TH:i:s\Z');

            $this->addFetchLog($meta, 'error', $fetched, $folderNames, $e->getMessage(), $startTime);

            throw $e;
        } finally {
            $meta['last_fetch_run'] = gmdate('Y-m-d\TH:i:s\Z');
            $meta['last_fetch_count'] = $fetched;
            $mailbox->setMetaParam(Mail365ServiceProvider::META_KEY, $meta, true);
        }

        $this->line('['.date('Y-m-d H:i:s').'] Total fetched: ' . $fetched);
    }

    protected function applyPostFetchAction($email, $graphId, $postAction, $moveFolderId, $fetchMode)
    {
        if ($fetchMode === 'unread') {
            $this->markAsRead($email, $graphId);
        }

        if ($postAction === 'mark_read' && $fetchMode !== 'unread') {
            $this->markAsRead($email, $graphId);
        } elseif ($postAction === 'move' && $moveFolderId) {
            $this->moveToFolder($email, $graphId, $moveFolderId);
        }
    }

    protected function markAsRead($email, $graphMessageId)
    {
        $url = 'https://graph.microsoft.com/v1.0/users/' . urlencode($email)
             . '/messages/' . urlencode($graphMessageId);

        try {
            $this->client->graphRequest('PATCH', $url, ['isRead' => true], [], 2);
        } catch (\Exception $e) {
            $this->error('['.date('Y-m-d H:i:s').'] Failed to mark as read: ' . $e->getMessage());
        }
    }

    protected function moveToFolder($email, $graphMessageId, $destinationFolderId)
    {
        $url = 'https://graph.microsoft.com/v1.0/users/' . urlencode($email)
             . '/messages/' . urlencode($graphMessageId) . '/move';

        try {
            $result = $this->client->graphRequest('POST', $url, ['destinationId' => $destinationFolderId], [], 2);

            if ($result['status'] >= 400) {
                $error = $result['body']['error']['message'] ?? 'Unknown error';
                $this->error('['.date('Y-m-d H:i:s').'] Failed to move message: ' . $error);
            }
        } catch (\Exception $e) {
            $this->error('['.date('Y-m-d H:i:s').'] Failed to move message: ' . $e->getMessage());
        }
    }

    protected function addFetchLog(array &$meta, $status, $messageCount, array $folders, $error, $startTime)
    {
        $log = $meta['fetch_log'] ?? [];

        array_unshift($log, [
            'timestamp' => gmdate('Y-m-d\TH:i:s\Z'),
            'status'    => $status,
            'messages'  => (int) $messageCount,
            'folders'   => $folders,
            'error'     => $error ? mb_substr($error, 0, 300) : null,
            'duration'  => round(microtime(true) - $startTime, 1),
        ]);

        $meta['fetch_log'] = array_slice($log, 0, 20);
    }

    protected function isAlreadyImported(array &$meta, $messageId)
    {
        $imported = $meta['imported_ids'] ?? [];
        $flipped = array_flip($imported);
        return isset($flipped[$messageId]);
    }

    protected function trackImported(array &$meta, $messageId)
    {
        $imported = $meta['imported_ids'] ?? [];
        $imported[] = $messageId;
        $meta['imported_ids'] = array_slice($imported, -500);
    }

    protected function buildExpiryFromManualDate(array $meta)
    {
        $date = $meta['secret_expiry_date'] ?? '';
        if (!$date || !preg_match('/^\d{4}-\d{2}-\d{2}$/', $date)) {
            return null;
        }

        $endTs = strtotime($date);
        if (!$endTs) return null;

        $daysLeft = (int) ceil(($endTs - time()) / 86400);
        if ($daysLeft < 0) $daysLeft = 0;

        return [
            'timestamp'    => $endTs,
            'date'         => $date,
            'days_left'    => $daysLeft,
            'display_name' => '',
        ];
    }

    protected function maybeCreateExpiryAlert(Mailbox $mailbox, array &$meta, array $expiry)
    {
        if (empty($meta['expiry_alert_enabled'])) {
            return;
        }

        $alertDays = (int) ($meta['expiry_alert_days'] ?? 30);
        if ($alertDays < 1) $alertDays = 30;

        $daysLeft = $expiry['days_left'] ?? null;
        if (!$daysLeft || $daysLeft > $alertDays) {
            return;
        }

        $today = date('Y-m-d');
        $lastAlertDate = $meta['expiry_alert_last_sent'] ?? null;
        if ($lastAlertDate === $today) {
            return;
        }

        $secretName = $expiry['display_name'] ?? '';
        $expiryDate = $expiry['date'] ?? 'unknown';

        $subject = "Azure client secret expiring in {$daysLeft} day" . ($daysLeft !== 1 ? 's' : '');
        $body = '<p>The Azure AD client secret for mailbox <strong>' . e($mailbox->name) . '</strong> '
              . '(' . e($mailbox->email) . ') is expiring soon.</p>'
              . '<ul>'
              . '<li><strong>Expiry date:</strong> ' . e($expiryDate) . '</li>'
              . '<li><strong>Days remaining:</strong> ' . (int) $daysLeft . '</li>'
              . ($secretName ? '<li><strong>Secret name:</strong> ' . e($secretName) . '</li>' : '')
              . '</ul>'
              . '<p>Please renew the client secret in the Azure portal and update it in '
              . '<a href="' . e(url('/mailbox/connection-settings/' . $mailbox->id . '/incoming')) . '">Connection Settings</a>.</p>';

        try {
            $alertEmail = 'mail365-alert@' . parse_url(config('app.url'), PHP_URL_HOST);
            $customer = Customer::create($alertEmail, [
                'first_name' => 'Mail365',
                'last_name'  => 'Alert',
            ]);

            if (!$customer) {
                $this->error('['.date('Y-m-d H:i:s').'] Could not create alert customer');
                return;
            }

            $conversation = new Conversation();
            $conversation->type = Conversation::TYPE_EMAIL;
            $conversation->subject = $subject;
            $conversation->mailbox_id = $mailbox->id;
            $conversation->source_via = Conversation::PERSON_CUSTOMER;
            $conversation->source_type = Conversation::SOURCE_TYPE_WEB;
            $conversation->customer_id = $customer->id;
            $conversation->customer_email = $customer->getMainEmail();
            $conversation->state = Conversation::STATE_PUBLISHED;
            $conversation->status = Conversation::STATUS_ACTIVE;
            $conversation->preview = mb_substr(strip_tags($body), 0, 255);
            $conversation->updateFolder();
            $conversation->save();

            $thread = new Thread();
            $thread->conversation_id = $conversation->id;
            $thread->type = Thread::TYPE_CUSTOMER;
            $thread->source_via = Thread::PERSON_CUSTOMER;
            $thread->source_type = Thread::SOURCE_TYPE_WEB;
            $thread->state = Thread::STATE_PUBLISHED;
            $thread->customer_id = $customer->id;
            $thread->created_by_customer_id = $customer->id;
            $thread->body = $body;
            $thread->first = true;
            $thread->save();

            $conversation->threads_count = 1;
            $conversation->last_reply_at = $thread->created_at;
            $conversation->save();

            $meta['expiry_alert_last_sent'] = $today;

            $this->info('['.date('Y-m-d H:i:s').'] Created expiry alert ticket #' . $conversation->number);

        } catch (\Exception $e) {
            $this->error('['.date('Y-m-d H:i:s').'] Failed to create expiry alert: ' . $e->getMessage());
            \Log::error('Mail365 expiry alert error', [
                'mailbox' => $mailbox->email,
                'error'   => $e->getMessage(),
            ]);
        }
    }
}

class Graph365FetchEmails extends \App\Console\Commands\FetchEmails
{
    public function initOutput($output)
    {
        $this->output = $output;
    }

    public function setSeen($message, $mailbox)
    {
        // No-op: Graph API messages have no IMAP client.
        // Read status is managed via Graph API PATCH call.
    }
}

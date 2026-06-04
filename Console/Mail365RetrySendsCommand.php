<?php

namespace Modules\Mail365\Console;

use App\Mailbox;
use Illuminate\Console\Command;
use Modules\Mail365\Helpers\Mail365Client;
use Modules\Mail365\Providers\Mail365ServiceProvider;

class Mail365RetrySendsCommand extends Command
{
    protected $signature = 'mail365:retry-sends';

    protected $description = 'Retry queued Microsoft 365 API sends that failed with transient errors';

    const MAX_ATTEMPTS = 5;

    public function handle()
    {
        $mailboxes = Mailbox::where('out_method', Mail365ServiceProvider::OUT_METHOD_GRAPH365)->get();

        foreach ($mailboxes as $mailbox) {
            $this->processMailbox($mailbox);
        }
    }

    protected function processMailbox(Mailbox $mailbox)
    {
        $meta = $mailbox->getMeta(Mail365ServiceProvider::META_KEY, []);
        $queue = $meta['retry_queue'] ?? [];

        if (empty($queue)) return;

        $tenantId       = $meta['tenant_id'] ?? '';
        $clientId       = $meta['client_id'] ?? '';
        $authType       = $meta['auth_type'] ?? 'secret';
        $clientSecret   = '';
        $certificatePem = '';

        if ($authType === 'certificate') {
            $certificatePem = !empty($meta['certificate_pem']) ? \Helper::decrypt($meta['certificate_pem']) : '';
            if (!$tenantId || !$clientId || !$certificatePem) {
                $this->error("Mailbox {$mailbox->email}: missing Azure certificate, clearing retry queue");
                $meta['retry_queue'] = [];
                $mailbox->setMetaParam(Mail365ServiceProvider::META_KEY, $meta, true);
                return;
            }
        } else {
            $clientSecret = !empty($meta['client_secret']) ? \Helper::decrypt($meta['client_secret']) : '';
            if (!$tenantId || !$clientId || !$clientSecret) {
                $this->error("Mailbox {$mailbox->email}: missing Azure credentials, clearing retry queue");
                $meta['retry_queue'] = [];
                $mailbox->setMetaParam(Mail365ServiceProvider::META_KEY, $meta, true);
                return;
            }
        }

        $client = new Mail365Client($tenantId, $clientId, $clientSecret, null, $authType, $certificatePem);
        $remaining = [];
        $succeeded = 0;
        $failed = 0;

        foreach ($queue as $item) {
            $attempts = ($item['attempts'] ?? 0) + 1;

            if ($attempts > self::MAX_ATTEMPTS) {
                $failed++;
                \Log::warning("Mail365: Dropping queued send after {$attempts} attempts", [
                    'mailbox' => $mailbox->email,
                    'subject' => $item['subject'] ?? '(unknown)',
                ]);
                continue;
            }

            $mimeEnc = $item['mime_enc'] ?? $item['mime_b64'] ?? '';
            $mime = !empty($item['mime_enc']) ? \Helper::decrypt($mimeEnc) : base64_decode($mimeEnc);
            if (!$mime) {
                $failed++;
                continue;
            }

            try {
                $response = $client->sendRawMime($mailbox->email, $mime);

                if ($response['status'] >= 400) {
                    $errorBody = $response['body']['error'] ?? [];
                    $error = $errorBody['message'] ?? 'Unknown error';

                    if ($response['status'] == 429 || $response['status'] == 503) {
                        $item['attempts'] = $attempts;
                        $item['last_error'] = $error;
                        $item['last_attempt'] = gmdate('Y-m-d\TH:i:s\Z');
                        $remaining[] = $item;
                    } else {
                        $failed++;
                        \Log::error("Mail365: Retry send permanent failure", [
                            'status' => $response['status'],
                            'error'  => $error,
                        ]);
                    }
                    continue;
                }

                $succeeded++;
            } catch (\Exception $e) {
                $item['attempts'] = $attempts;
                $item['last_error'] = mb_substr($e->getMessage(), 0, 300);
                $item['last_attempt'] = gmdate('Y-m-d\TH:i:s\Z');
                $remaining[] = $item;
            }
        }

        $meta['retry_queue'] = array_slice($remaining, 0, 20);
        $mailbox->setMetaParam(Mail365ServiceProvider::META_KEY, $meta, true);

        if ($succeeded || $failed) {
            $this->info("Mailbox {$mailbox->email}: {$succeeded} retried OK, {$failed} dropped, " . count($remaining) . " still queued");
        }
    }
}

<?php

namespace Modules\Mail365\Transport;

use Modules\Mail365\Helpers\Mail365Client;
use Swift_Events_EventListener;
use Swift_Mime_SimpleMessage;
use Swift_Transport;

class Graph365Transport implements Swift_Transport
{
    protected $senderEmail;
    protected $client;
    protected $started = false;

    const MIME_SIZE_LIMIT = 3500000;
    const LARGE_ATTACHMENT_THRESHOLD = 3000000;
    const UPLOAD_CHUNK_SIZE = 3932160;

    public function __construct($tenantId, $clientId, $clientSecret, $senderEmail, $authType = 'secret', $certificatePem = '')
    {
        $this->senderEmail = $senderEmail;
        $this->client = new Mail365Client($tenantId, $clientId, $clientSecret, null, $authType, $certificatePem);
    }

    public function isStarted()
    {
        return $this->started;
    }

    public function start()
    {
        $this->started = true;
    }

    public function stop()
    {
        $this->started = false;
    }

    public function ping()
    {
        return true;
    }

    public function send(Swift_Mime_SimpleMessage $message, &$failedRecipients = null)
    {
        $startTime = microtime(true);

        try {
            $mime = $message->toString();

            if (strlen($mime) < self::MIME_SIZE_LIMIT) {
                $response = $this->client->sendRawMime($this->senderEmail, $mime);
            } else {
                \Log::info('Mail365: Message exceeds MIME size limit, using draft+upload approach');
                $response = $this->sendLargeMessage($message);
            }
        } catch (\Exception $e) {
            $this->logSend($message, 'error', $e->getMessage(), $startTime);
            throw $e;
        }

        if ($response['status'] >= 400) {
            $errorBody = $response['body']['error'] ?? [];
            $error = $errorBody['message'] ?? 'Unknown Graph API error';
            $code = $errorBody['code'] ?? '';
            $detail = '';
            if (!empty($errorBody['innerError']['message'])) {
                $detail = ' — ' . $errorBody['innerError']['message'];
            }

            \Log::error('Mail365 Graph API error', [
                'status'  => $response['status'],
                'code'    => $code,
                'message' => $error,
                'sender'  => $this->senderEmail,
                'body'    => $response['body'],
            ]);

            $this->logSend($message, 'error', "{$code}: {$error}", $startTime);

            if (in_array($response['status'], [429, 503])) {
                $this->queueForRetry($message);
            }

            throw new \Swift_TransportException(
                "Microsoft 365 API error ({$response['status']} {$code}): {$error}{$detail}"
            );
        }

        $this->logSend($message, 'success', null, $startTime);

        return count((array) $message->getTo())
             + count((array) $message->getCc())
             + count((array) $message->getBcc());
    }

    public function registerPlugin(Swift_Events_EventListener $plugin)
    {
    }

    protected function sendLargeMessage(Swift_Mime_SimpleMessage $message)
    {
        $largeAttachments = [];
        $draftPayload = $this->buildDraftPayload($message, $largeAttachments);

        $draftUrl = "https://graph.microsoft.com/v1.0/users/" . urlencode($this->senderEmail) . "/messages";
        $draftResult = $this->client->graphRequest(
            'POST', $draftUrl, $draftPayload, [], 2, \Swift_TransportException::class
        );

        if ($draftResult['status'] >= 400 || empty($draftResult['body']['id'])) {
            return $draftResult;
        }

        $draftId = $draftResult['body']['id'];

        foreach ($largeAttachments as $attachment) {
            $this->uploadLargeAttachment($draftId, $attachment);
        }

        $sendUrl = "https://graph.microsoft.com/v1.0/users/" . urlencode($this->senderEmail)
                 . "/messages/" . urlencode($draftId) . "/send";

        $sendResult = $this->client->curlWithRetry($sendUrl, [
            CURLOPT_POST           => true,
            CURLOPT_POSTFIELDS     => '',
            CURLOPT_RETURNTRANSFER => true,
            CURLOPT_TIMEOUT        => 60,
            CURLOPT_SSL_VERIFYPEER => true,
            CURLOPT_SSL_VERIFYHOST => 2,
            CURLOPT_HTTPHEADER     => [
                'Authorization: Bearer ' . $this->client->getAccessToken(\Swift_TransportException::class),
                'Content-Type: application/json',
                'Content-Length: 0',
            ],
        ], 2, \Swift_TransportException::class);

        return [
            'status' => $sendResult['status'],
            'body'   => json_decode($sendResult['body'], true) ?: [],
        ];
    }

    protected function buildDraftPayload(Swift_Mime_SimpleMessage $message, array &$largeAttachments)
    {
        $contentType = stripos($message->getContentType(), 'html') !== false ? 'HTML' : 'Text';
        $bodyContent = $message->getBody();

        $payload = [
            'subject' => $message->getSubject(),
            'body' => [
                'contentType' => $contentType,
                'content'     => $bodyContent ?: '',
            ],
            'toRecipients'  => $this->formatRecipients($message->getTo()),
            'ccRecipients'  => $this->formatRecipients($message->getCc()),
            'bccRecipients' => $this->formatRecipients($message->getBcc()),
        ];

        if ($message->getReplyTo()) {
            $payload['replyTo'] = $this->formatRecipients($message->getReplyTo());
        }

        $headers = $message->getHeaders();
        $internetMessageHeaders = [];

        $inReplyTo = $headers->get('In-Reply-To');
        if ($inReplyTo) {
            $internetMessageHeaders[] = [
                'name'  => 'In-Reply-To',
                'value' => $inReplyTo->getFieldBody(),
            ];
        }

        $references = $headers->get('References');
        if ($references) {
            $internetMessageHeaders[] = [
                'name'  => 'References',
                'value' => $references->getFieldBody(),
            ];
        }

        if ($internetMessageHeaders) {
            $payload['internetMessageHeaders'] = $internetMessageHeaders;
        }

        $smallAttachments = [];
        $largeAttachments = [];

        foreach ($message->getChildren() as $child) {
            if ($child instanceof \Swift_MimePart) {
                if (stripos($child->getContentType(), 'text/plain') !== false && !$bodyContent) {
                    $payload['body'] = [
                        'contentType' => 'Text',
                        'content'     => $child->getBody() ?: '',
                    ];
                }
                continue;
            }

            if ($child instanceof \Swift_Attachment || $child instanceof \Swift_EmbeddedFile) {
                $content = $child->getBody();
                $size    = strlen($content);

                $att = [
                    'name'        => $child->getFilename() ?: 'attachment',
                    'contentType' => $child->getContentType(),
                    'content'     => $content,
                    'size'        => $size,
                    'isInline'    => ($child instanceof \Swift_EmbeddedFile),
                    'contentId'   => ($child instanceof \Swift_EmbeddedFile) ? $child->getId() : null,
                ];

                if ($size > self::LARGE_ATTACHMENT_THRESHOLD) {
                    $largeAttachments[] = $att;
                } else {
                    $smallAttachments[] = $att;
                }
            }
        }

        if ($smallAttachments) {
            $payload['attachments'] = [];
            foreach ($smallAttachments as $att) {
                $item = [
                    '@odata.type' => '#microsoft.graph.fileAttachment',
                    'name'        => $att['name'],
                    'contentType' => $att['contentType'],
                    'contentBytes' => base64_encode($att['content']),
                ];
                if ($att['isInline']) {
                    $item['isInline'] = true;
                    $item['contentId'] = $att['contentId'];
                }
                $payload['attachments'][] = $item;
            }
        }

        return $payload;
    }

    protected function formatRecipients($addresses)
    {
        if (!$addresses) return [];

        $result = [];
        foreach ($addresses as $email => $name) {
            $recipient = ['emailAddress' => ['address' => $email]];
            if ($name) {
                $recipient['emailAddress']['name'] = $name;
            }
            $result[] = $recipient;
        }
        return $result;
    }

    protected function uploadLargeAttachment($draftId, array $attachment)
    {
        $url = "https://graph.microsoft.com/v1.0/users/" . urlencode($this->senderEmail)
             . "/messages/" . urlencode($draftId) . "/attachments/createUploadSession";

        $sessionResult = $this->client->graphRequest('POST', $url, [
            'AttachmentItem' => [
                'attachmentType' => 'file',
                'name'           => $attachment['name'],
                'size'           => $attachment['size'],
                'contentType'    => $attachment['contentType'],
            ],
        ], [], 2, \Swift_TransportException::class);

        if ($sessionResult['status'] >= 400 || empty($sessionResult['body']['uploadUrl'])) {
            $error = $sessionResult['body']['error']['message'] ?? 'No upload URL returned';
            throw new \Swift_TransportException("Failed to create upload session: {$error}");
        }

        $uploadUrl  = $sessionResult['body']['uploadUrl'];
        $content    = $attachment['content'];
        $totalSize  = $attachment['size'];
        $chunkSize  = self::UPLOAD_CHUNK_SIZE;
        $offset     = 0;

        while ($offset < $totalSize) {
            $chunk    = substr($content, $offset, $chunkSize);
            $chunkLen = strlen($chunk);
            $end      = $offset + $chunkLen - 1;

            $ch = curl_init($uploadUrl);
            curl_setopt_array($ch, [
                CURLOPT_CUSTOMREQUEST  => 'PUT',
                CURLOPT_POSTFIELDS     => $chunk,
                CURLOPT_RETURNTRANSFER => true,
                CURLOPT_TIMEOUT        => 120,
                CURLOPT_SSL_VERIFYPEER => true,
                CURLOPT_SSL_VERIFYHOST => 2,
                CURLOPT_HTTPHEADER     => [
                    'Content-Type: application/octet-stream',
                    'Content-Length: ' . $chunkLen,
                    'Content-Range: bytes ' . $offset . '-' . $end . '/' . $totalSize,
                ],
            ]);

            $body     = curl_exec($ch);
            $httpCode = curl_getinfo($ch, CURLINFO_HTTP_CODE);
            $curlErr  = curl_error($ch);
            curl_close($ch);

            if ($curlErr) {
                throw new \Swift_TransportException("Attachment upload failed: {$curlErr}");
            }

            if ($httpCode >= 400) {
                $errData = json_decode($body, true) ?: [];
                $error = $errData['error']['message'] ?? "HTTP {$httpCode}";
                throw new \Swift_TransportException("Attachment upload failed: {$error}");
            }

            $offset += $chunkSize;
        }
    }

    protected function queueForRetry(Swift_Mime_SimpleMessage $message)
    {
        try {
            $mailbox = \App\Mailbox::where('email', $this->senderEmail)->first();
            if (!$mailbox) return;

            $mime = $message->toString();
            if (strlen($mime) > 4000000) {
                \Log::warning('Mail365: Message too large to queue for retry');
                return;
            }

            $metaKey = \Modules\Mail365\Providers\Mail365ServiceProvider::META_KEY;
            $meta = $mailbox->getMeta($metaKey, []);
            $queue = $meta['retry_queue'] ?? [];

            $queue[] = [
                'mime_enc'  => \Helper::encrypt($mime),
                'subject'   => mb_substr($message->getSubject() ?: '', 0, 100),
                'to'        => mb_substr(implode(', ', array_keys((array) $message->getTo())), 0, 200),
                'queued_at' => gmdate('Y-m-d\TH:i:s\Z'),
                'attempts'  => 0,
            ];

            $meta['retry_queue'] = array_slice($queue, -20);
            $mailbox->setMetaParam($metaKey, $meta, true);

            \Log::info('Mail365: Queued message for retry', ['subject' => $message->getSubject()]);
        } catch (\Exception $e) {
            \Log::warning('Mail365: Failed to queue for retry: ' . $e->getMessage());
        }
    }

    protected function logSend(Swift_Mime_SimpleMessage $message, $status, $error, $startTime)
    {
        try {
            $mailbox = \App\Mailbox::where('email', $this->senderEmail)->first();
            if (!$mailbox) return;

            $metaKey = \Modules\Mail365\Providers\Mail365ServiceProvider::META_KEY;
            $meta = $mailbox->getMeta($metaKey, []);
            $log = $meta['send_log'] ?? [];

            $recipients = array_keys((array) $message->getTo());

            array_unshift($log, [
                'timestamp' => gmdate('Y-m-d\TH:i:s\Z'),
                'to'        => mb_substr(implode(', ', array_slice($recipients, 0, 5)), 0, 200),
                'subject'   => mb_substr($message->getSubject() ?: '(no subject)', 0, 100),
                'status'    => $status,
                'error'     => $error ? mb_substr($error, 0, 300) : null,
                'duration'  => round(microtime(true) - $startTime, 1),
            ]);

            $meta['send_log'] = array_slice($log, 0, 30);

            if ($status === 'success') {
                $meta['last_send_success'] = gmdate('Y-m-d\TH:i:s\Z');
                $meta['last_send_error'] = null;
                $meta['last_send_error_at'] = null;
            } else {
                $meta['last_send_error'] = $error ? mb_substr($error, 0, 500) : 'Unknown error';
                $meta['last_send_error_at'] = gmdate('Y-m-d\TH:i:s\Z');
            }

            $mailbox->setMetaParam($metaKey, $meta, true);
        } catch (\Exception $e) {
            \Log::warning('Mail365: Failed to log send: ' . $e->getMessage());
        }
    }
}

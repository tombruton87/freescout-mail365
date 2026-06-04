<?php

namespace Modules\Mail365\Helpers;

class Mail365Client
{
    protected $tenantId;
    protected $clientId;
    protected $clientSecret;
    protected $authType;
    protected $certificatePem;
    protected $logger;
    protected static $tokenCache = [];

    public function __construct($tenantId, $clientId, $clientSecret, ?callable $logger = null, $authType = 'secret', $certificatePem = '')
    {
        $this->tenantId = $tenantId;
        $this->clientId = $clientId;
        $this->clientSecret = $clientSecret;
        $this->authType = $authType;
        $this->certificatePem = $certificatePem;
        $this->logger = $logger;
    }

    public function curlWithRetry($url, array $options, $maxRetries = 3, $exceptionClass = \Exception::class)
    {
        $baseDelay = 2;

        for ($attempt = 0; $attempt <= $maxRetries; $attempt++) {
            $retryAfter = null;

            $ch = curl_init($url);
            curl_setopt_array($ch, $options);
            curl_setopt($ch, CURLOPT_HEADERFUNCTION, function ($ch, $header) use (&$retryAfter) {
                if (stripos($header, 'Retry-After:') === 0) {
                    $retryAfter = (int) trim(substr($header, 12));
                }
                return strlen($header);
            });

            $body     = curl_exec($ch);
            $httpCode = curl_getinfo($ch, CURLINFO_HTTP_CODE);
            $curlErr  = curl_error($ch);
            curl_close($ch);

            if ($curlErr) {
                if ($attempt < $maxRetries) {
                    $delay = min($baseDelay * pow(2, $attempt), 30);
                    $this->log("cURL error, retrying in {$delay}s: {$curlErr}", 'warning');
                    sleep($delay);
                    continue;
                }
                throw new $exceptionClass("Request failed: {$curlErr}");
            }

            if (($httpCode == 429 || $httpCode == 503) && $attempt < $maxRetries) {
                $delay = min($retryAfter ?: ($baseDelay * pow(2, $attempt)), 120);
                $this->log("Rate limited (HTTP {$httpCode}), retrying in {$delay}s", 'warning');
                sleep($delay);
                continue;
            }

            return ['body' => $body, 'status' => $httpCode];
        }

        throw new $exceptionClass("Request failed after {$maxRetries} retries");
    }

    public function getAccessToken($exceptionClass = \Exception::class)
    {
        $cacheKey = $this->tenantId . ':' . $this->clientId;

        if (isset(self::$tokenCache[$cacheKey]) && time() < self::$tokenCache[$cacheKey]['expires_at']) {
            return self::$tokenCache[$cacheKey]['token'];
        }

        $url = "https://login.microsoftonline.com/" . urlencode($this->tenantId) . "/oauth2/v2.0/token";

        if ($this->authType === 'certificate' && $this->certificatePem) {
            $postFields = http_build_query([
                'client_id'             => $this->clientId,
                'client_assertion_type' => 'urn:ietf:params:oauth:client-assertion-type:jwt-bearer',
                'client_assertion'      => $this->buildClientAssertion($url, $exceptionClass),
                'scope'                 => 'https://graph.microsoft.com/.default',
                'grant_type'            => 'client_credentials',
            ]);
        } else {
            $postFields = http_build_query([
                'client_id'     => $this->clientId,
                'client_secret' => $this->clientSecret,
                'scope'         => 'https://graph.microsoft.com/.default',
                'grant_type'    => 'client_credentials',
            ]);
        }

        $result = $this->curlWithRetry($url, [
            CURLOPT_POST           => true,
            CURLOPT_POSTFIELDS     => $postFields,
            CURLOPT_RETURNTRANSFER => true,
            CURLOPT_TIMEOUT        => 30,
            CURLOPT_SSL_VERIFYPEER => true,
            CURLOPT_SSL_VERIFYHOST => 2,
            CURLOPT_HTTPHEADER     => ['Content-Type: application/x-www-form-urlencoded'],
        ], 3, $exceptionClass);

        $data = json_decode($result['body'], true);

        if ($result['status'] >= 400 || empty($data['access_token'])) {
            $error = $data['error_description'] ?? $data['error'] ?? 'Unknown error';
            throw new $exceptionClass("Failed to get access token: {$error}");
        }

        self::$tokenCache[$cacheKey] = [
            'token'      => $data['access_token'],
            'expires_at' => time() + ($data['expires_in'] ?? 3500) - 60,
        ];

        return $data['access_token'];
    }

    protected function buildClientAssertion($tokenEndpoint, $exceptionClass = \Exception::class)
    {
        $certResource = openssl_x509_read($this->certificatePem);
        if (!$certResource) {
            throw new $exceptionClass('Invalid certificate: unable to parse the X.509 certificate from PEM.');
        }

        $certDer = '';
        openssl_x509_export($certResource, $certPemExported);
        $certDer = base64_decode(
            str_replace(['-----BEGIN CERTIFICATE-----', '-----END CERTIFICATE-----', "\n", "\r"], '', $certPemExported)
        );
        $thumbprint = self::base64UrlEncode(sha1($certDer, true));

        $privateKey = openssl_pkey_get_private($this->certificatePem);
        if (!$privateKey) {
            throw new $exceptionClass('Invalid certificate: unable to extract the private key from PEM.');
        }

        $now = time();
        $header = ['alg' => 'RS256', 'typ' => 'JWT', 'x5t' => $thumbprint];
        $payload = [
            'aud' => $tokenEndpoint,
            'iss' => $this->clientId,
            'sub' => $this->clientId,
            'jti' => bin2hex(random_bytes(16)),
            'nbf' => $now,
            'exp' => $now + 600,
        ];

        $segments = [
            self::base64UrlEncode(json_encode($header)),
            self::base64UrlEncode(json_encode($payload)),
        ];

        $signingInput = implode('.', $segments);
        $signature = '';

        if (!openssl_sign($signingInput, $signature, $privateKey, OPENSSL_ALGO_SHA256)) {
            throw new $exceptionClass('Failed to sign the client assertion JWT.');
        }

        $segments[] = self::base64UrlEncode($signature);

        return implode('.', $segments);
    }

    public static function base64UrlEncode($data)
    {
        return rtrim(strtr(base64_encode($data), '+/', '-_'), '=');
    }

    public static function extractCertificateThumbprint($pem)
    {
        $certResource = openssl_x509_read($pem);
        if (!$certResource) {
            return null;
        }

        openssl_x509_export($certResource, $certPemExported);
        $certDer = base64_decode(
            str_replace(['-----BEGIN CERTIFICATE-----', '-----END CERTIFICATE-----', "\n", "\r"], '', $certPemExported)
        );

        return strtoupper(sha1($certDer));
    }

    public static function extractCertificateExpiry($pem)
    {
        $certResource = openssl_x509_read($pem);
        if (!$certResource) {
            return null;
        }

        $certData = openssl_x509_parse($certResource);
        if (empty($certData['validTo_time_t'])) {
            return null;
        }

        $endTs = (int) $certData['validTo_time_t'];
        $daysLeft = (int) ceil(($endTs - time()) / 86400);

        $subject = $certData['subject']['CN'] ?? '';

        return [
            'timestamp'    => $endTs,
            'date'         => date('Y-m-d', $endTs),
            'days_left'    => $daysLeft,
            'display_name' => $subject ? "Certificate: {$subject}" : 'Certificate',
        ];
    }

    public static function validateCertificatePem($pem)
    {
        $errors = [];

        $certResource = openssl_x509_read($pem);
        if (!$certResource) {
            $errors[] = 'No valid X.509 certificate found in the PEM file.';
        }

        $privateKey = openssl_pkey_get_private($pem);
        if (!$privateKey) {
            $errors[] = 'No valid private key found in the PEM file.';
        }

        if ($certResource && $privateKey) {
            if (!openssl_x509_check_private_key($certResource, $privateKey)) {
                $errors[] = 'The private key does not match the certificate.';
            }
        }

        return $errors;
    }

    public function graphGet($url, $exceptionClass = \Exception::class)
    {
        $token = $this->getAccessToken($exceptionClass);

        $result = $this->curlWithRetry($url, [
            CURLOPT_RETURNTRANSFER => true,
            CURLOPT_TIMEOUT        => 60,
            CURLOPT_SSL_VERIFYPEER => true,
            CURLOPT_SSL_VERIFYHOST => 2,
            CURLOPT_HTTPHEADER     => [
                'Authorization: Bearer ' . $token,
                'Content-Type: application/json',
            ],
        ], 3, $exceptionClass);

        return [
            'status' => $result['status'],
            'body'   => json_decode($result['body'], true) ?: [],
        ];
    }

    public function graphRequest($method, $url, $body = null, $headers = [], $maxRetries = 3, $exceptionClass = \Exception::class)
    {
        $token = $this->getAccessToken($exceptionClass);

        $defaultHeaders = [
            'Authorization: Bearer ' . $token,
            'Content-Type: application/json',
        ];

        $options = [
            CURLOPT_RETURNTRANSFER => true,
            CURLOPT_TIMEOUT        => 60,
            CURLOPT_SSL_VERIFYPEER => true,
            CURLOPT_SSL_VERIFYHOST => 2,
            CURLOPT_HTTPHEADER     => array_merge($defaultHeaders, $headers),
        ];

        if ($method === 'POST') {
            $options[CURLOPT_POST] = true;
            if ($body !== null) {
                $options[CURLOPT_POSTFIELDS] = is_string($body) ? $body : json_encode($body);
            }
        } elseif ($method !== 'GET') {
            $options[CURLOPT_CUSTOMREQUEST] = $method;
            if ($body !== null) {
                $options[CURLOPT_POSTFIELDS] = is_string($body) ? $body : json_encode($body);
            }
        }

        $result = $this->curlWithRetry($url, $options, $maxRetries, $exceptionClass);

        return [
            'status' => $result['status'],
            'body'   => json_decode($result['body'], true) ?: [],
        ];
    }

    public function graphRaw($url, $timeout = 120, $maxRetries = 3, $exceptionClass = \Exception::class)
    {
        $token = $this->getAccessToken($exceptionClass);

        $result = $this->curlWithRetry($url, [
            CURLOPT_RETURNTRANSFER => true,
            CURLOPT_TIMEOUT        => $timeout,
            CURLOPT_SSL_VERIFYPEER => true,
            CURLOPT_SSL_VERIFYHOST => 2,
            CURLOPT_HTTPHEADER     => [
                'Authorization: Bearer ' . $token,
            ],
        ], $maxRetries, $exceptionClass);

        return $result;
    }

    public function sendRawMime($email, $mime, $exceptionClass = \Swift_TransportException::class)
    {
        $token = $this->getAccessToken($exceptionClass);

        $url = "https://graph.microsoft.com/v1.0/users/" . urlencode($email) . "/sendMail";

        $result = $this->curlWithRetry($url, [
            CURLOPT_POST           => true,
            CURLOPT_POSTFIELDS     => base64_encode($mime),
            CURLOPT_RETURNTRANSFER => true,
            CURLOPT_TIMEOUT        => 60,
            CURLOPT_SSL_VERIFYPEER => true,
            CURLOPT_SSL_VERIFYHOST => 2,
            CURLOPT_HTTPHEADER     => [
                'Authorization: Bearer ' . $token,
                'Content-Type: text/plain',
            ],
        ], 2, $exceptionClass);

        return [
            'status' => $result['status'],
            'body'   => json_decode($result['body'], true) ?: [],
        ];
    }

    public function checkSecretExpiry()
    {
        $url = 'https://graph.microsoft.com/v1.0/applications?$filter=' . urlencode("appId eq '{$this->clientId}'")
             . '&$select=passwordCredentials';

        try {
            $result = $this->graphGet($url);
        } catch (\Exception $e) {
            return null;
        }

        if ($result['status'] >= 400) {
            return null;
        }

        $apps = $result['body']['value'] ?? [];
        if (empty($apps[0]['passwordCredentials'])) {
            return null;
        }

        $now = time();
        $soonest = null;

        foreach ($apps[0]['passwordCredentials'] as $cred) {
            $end = $cred['endDateTime'] ?? null;
            if (!$end) continue;

            $endTs = strtotime($end);
            if ($endTs < $now) continue;

            if (!$soonest || $endTs < $soonest['timestamp']) {
                $soonest = [
                    'timestamp'    => $endTs,
                    'date'         => date('Y-m-d', $endTs),
                    'days_left'    => (int) ceil(($endTs - $now) / 86400),
                    'display_name' => $cred['displayName'] ?? '',
                ];
            }
        }

        return $soonest;
    }

    public function getTokenRoles()
    {
        try {
            $token = $this->getAccessToken();
        } catch (\Exception $e) {
            return [];
        }

        $parts = explode('.', $token);
        if (count($parts) < 2) return [];

        $payload = json_decode(base64_decode(strtr($parts[1], '-_', '+/')), true);
        return $payload['roles'] ?? [];
    }

    public function getMailboxQuota($email)
    {
        $url = 'https://graph.microsoft.com/v1.0/users/' . urlencode($email)
             . '/mailFolders?$top=200&includeHiddenFolders=true';

        $totalSize = 0;
        $folderCount = 0;

        while ($url) {
            $response = $this->graphGet($url);
            if ($response['status'] >= 400) {
                return null;
            }

            foreach ($response['body']['value'] ?? [] as $folder) {
                $totalSize += (int) ($folder['sizeInBytes'] ?? 0);
                $folderCount++;
            }

            $nextLink = $response['body']['@odata.nextLink'] ?? null;
            $url = ($nextLink && parse_url($nextLink, PHP_URL_HOST) === 'graph.microsoft.com') ? $nextLink : null;
        }

        return [
            'used_bytes'   => $totalSize,
            'used_display' => self::formatBytes($totalSize),
            'folder_count' => $folderCount,
            'checked_at'   => gmdate('Y-m-d\TH:i:s\Z'),
        ];
    }

    public function listMailboxes($pageSize = 100)
    {
        $url = 'https://graph.microsoft.com/v1.0/users'
             . '?$select=id,displayName,mail,userPrincipalName'
             . '&$filter=' . urlencode('accountEnabled eq true')
             . '&$top=' . $pageSize
             . '&$orderby=' . urlencode('displayName asc');

        $mailboxes = [];

        while ($url) {
            $response = $this->graphGet($url);

            if ($response['status'] == 403) {
                return ['error' => 'missing_permission', 'mailboxes' => []];
            }
            if ($response['status'] >= 400) {
                $error = $response['body']['error']['message'] ?? 'Unknown error';
                return ['error' => $error, 'mailboxes' => []];
            }

            foreach ($response['body']['value'] ?? [] as $user) {
                $email = $user['mail'] ?? $user['userPrincipalName'] ?? '';
                if (!$email || !filter_var($email, FILTER_VALIDATE_EMAIL)) continue;

                $mailboxes[] = [
                    'id'           => $user['id'],
                    'display_name' => $user['displayName'] ?? '',
                    'email'        => strtolower($email),
                ];
            }

            if (count($mailboxes) >= 500) break;

            $nextLink = $response['body']['@odata.nextLink'] ?? null;
            $url = ($nextLink && parse_url($nextLink, PHP_URL_HOST) === 'graph.microsoft.com') ? $nextLink : null;
        }

        return ['error' => null, 'mailboxes' => $mailboxes];
    }

    public static function formatBytes($bytes)
    {
        if ($bytes >= 1073741824) {
            return round($bytes / 1073741824, 2) . ' GB';
        } elseif ($bytes >= 1048576) {
            return round($bytes / 1048576, 1) . ' MB';
        }
        return round($bytes / 1024) . ' KB';
    }

    public function getClientId()
    {
        return $this->clientId;
    }

    protected function log($message, $level = 'info')
    {
        $prefixed = '[' . date('Y-m-d H:i:s') . '] Mail365: ' . $message;

        if ($this->logger) {
            call_user_func($this->logger, $prefixed);
        } else {
            \Log::$level('Mail365: ' . $message);
        }
    }
}

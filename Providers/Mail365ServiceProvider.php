<?php

namespace Modules\Mail365\Providers;

use App\Mailbox;
use Illuminate\Support\ServiceProvider;
use Illuminate\Support\Facades\Route;
use Modules\Mail365\Transport\Graph365Transport;

if (!defined('MAIL365_MODULE')) {
    define('MAIL365_MODULE', 'mail365');
}

class Mail365ServiceProvider extends ServiceProvider
{
    const OUT_METHOD_GRAPH365 = 4;
    const IN_PROTOCOL_GRAPH365 = 3;
    const META_KEY = 'mail365';

    public function boot()
    {
        $this->loadViewsFrom(__DIR__ . '/../Resources/views', 'mail365');
        $this->registerRoutes();
        $this->hooks();
    }

    public function register()
    {
        $this->commands([
            \Modules\Mail365\Console\Mail365FetchCommand::class,
            \Modules\Mail365\Console\Mail365RetrySendsCommand::class,
        ]);
    }

    protected function registerRoutes()
    {
        Route::group([
            'middleware' => ['web', 'auth'],
            'prefix'    => 'mail365',
            'namespace' => 'Modules\Mail365\Http\Controllers',
        ], function () {
            require __DIR__ . '/../Http/routes.php';
        });
    }

    protected function hooks()
    {
        \Eventy::addFilter('javascripts', function ($javascripts) {
            $javascripts[] = \Module::getPublicPath(MAIL365_MODULE) . '/js/module.js';
            return $javascripts;
        });

        // --- Outgoing: swap Swift transport for Graph API ---
        \Eventy::addAction('mail.reapply_mail_config', function () {
            $fromAddress = \Config::get('mail.from.address');
            if (!$fromAddress) {
                return;
            }

            $mailbox = Mailbox::where('email', $fromAddress)->first();
            if (!$mailbox || $mailbox->out_method != self::OUT_METHOD_GRAPH365) {
                return;
            }

            $meta = $mailbox->getMeta(self::META_KEY, []);
            $tenantId       = $meta['tenant_id'] ?? '';
            $clientId       = $meta['client_id'] ?? '';
            $authType       = $meta['auth_type'] ?? 'secret';
            $clientSecret   = '';
            $certificatePem = '';

            if ($authType === 'certificate') {
                $certificatePem = !empty($meta['certificate_pem']) ? \Helper::decrypt($meta['certificate_pem']) : '';
                if (!$tenantId || !$clientId || !$certificatePem) {
                    \Log::error('Mail365: missing Azure certificate for mailbox ' . $mailbox->email);
                    return;
                }
            } else {
                $clientSecret = !empty($meta['client_secret']) ? \Helper::decrypt($meta['client_secret']) : '';
                if (!$tenantId || !$clientId || !$clientSecret) {
                    \Log::error('Mail365: missing Azure credentials for mailbox ' . $mailbox->email);
                    return;
                }
            }

            $transport = new Graph365Transport(
                $tenantId,
                $clientId,
                $clientSecret,
                $mailbox->email,
                $authType,
                $certificatePem
            );

            $mailer = new \Swift_Mailer($transport);
            \Mail::setSwiftMailer($mailer);
        });

        // --- Incoming: register Graph365 as a protocol option ---
        \Eventy::addFilter('mailbox.in_protocols', function ($protocols) {
            $protocols[self::IN_PROTOCOL_GRAPH365] = 'graph365';
            return $protocols;
        });

        \Eventy::addFilter('mailbox.in_protocols.display_names', function ($names) {
            $names[self::IN_PROTOCOL_GRAPH365] = 'Microsoft 365 API';
            return $names;
        });

        // Mark mailbox as active for incoming when using Graph365 protocol
        \Eventy::addFilter('mailbox.in_active', function ($active, $mailbox) {
            if ($mailbox->in_protocol == self::IN_PROTOCOL_GRAPH365) {
                $meta = $mailbox->getMeta(self::META_KEY, []);
                $authType = $meta['auth_type'] ?? 'secret';

                if (empty($meta['tenant_id']) || empty($meta['client_id'])) {
                    return false;
                }

                if ($authType === 'certificate') {
                    return !empty($meta['certificate_pem']) && \Helper::decrypt($meta['certificate_pem']);
                }

                return !empty($meta['client_secret']) && \Helper::decrypt($meta['client_secret']);
            }
            return $active;
        }, 20, 2);

        // Suppress IMAP setFlag for Graph365 mailboxes — we mark as read via Graph API
        \Eventy::addFilter('fetch_emails.set_seen_flag', function ($flag, $message, $mailbox) {
            if ($mailbox->in_protocol == self::IN_PROTOCOL_GRAPH365) {
                return null;
            }
            return $flag;
        }, 20, 3);

        // Save Azure credentials when incoming settings are saved
        \Eventy::addAction('mailbox.incoming_settings_before_save', function ($mailbox, $request) {
            if ($request->input('in_protocol') != self::IN_PROTOCOL_GRAPH365) {
                return;
            }

            $mailbox->in_server = 'graph.microsoft.com';
            $mailbox->in_port = 443;
            $mailbox->in_username = 'graph-api';
            $mailbox->in_password = 'graph-api';

            $existing = $mailbox->getMeta(self::META_KEY, []);

            $fetchMode = trim((string) $request->input('m365_fetch_mode', 'all'));
            if (!in_array($fetchMode, ['all', 'unread'])) {
                $fetchMode = 'all';
            }

            $postAction = trim((string) $request->input('m365_post_fetch_action', 'none'));
            if (!in_array($postAction, ['none', 'mark_read', 'move'])) {
                $postAction = 'none';
            }

            $moveFolderId = trim((string) $request->input('m365_post_fetch_move_folder', ''));
            $moveFolderName = trim((string) $request->input('m365_post_fetch_move_folder_name', ''));

            $sharedMailbox = trim((string) $request->input('m365_shared_mailbox_email', ''));
            if ($sharedMailbox && !filter_var($sharedMailbox, FILTER_VALIDATE_EMAIL)) {
                $sharedMailbox = '';
            }

            $expiryDateInput = trim((string) $request->input('m365_secret_expiry_date', ''));
            $expiryDateClean = '';
            if ($expiryDateInput && preg_match('/^\d{4}-\d{2}-\d{2}$/', $expiryDateInput)) {
                $expiryDateClean = $expiryDateInput;
            }

            $tenantId = trim((string) $request->input('m365_tenant_id', ''));
            $clientId = trim((string) $request->input('m365_client_id', ''));

            $uuidPattern = '/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i';
            if ($tenantId && !preg_match($uuidPattern, $tenantId)) {
                $tenantId = $existing['tenant_id'] ?? '';
            }
            if ($clientId && !preg_match($uuidPattern, $clientId)) {
                $clientId = $existing['client_id'] ?? '';
            }

            $authType = trim((string) $request->input('m365_auth_type', $existing['auth_type'] ?? 'secret'));
            if (!in_array($authType, ['secret', 'certificate'])) {
                $authType = 'secret';
            }

            $meta = array_merge($existing, [
                'tenant_id'                   => $tenantId,
                'client_id'                   => $clientId,
                'auth_type'                   => $authType,
                'fetch_mode'                  => $fetchMode,
                'post_fetch_action'           => $postAction,
                'post_fetch_move_folder'      => substr($moveFolderId, 0, 500),
                'post_fetch_move_folder_name' => substr($moveFolderName, 0, 255),
                'shared_mailbox_email'        => $sharedMailbox,
                'secret_expiry_date'          => $expiryDateClean,
                'expiry_alert_enabled'        => (bool) $request->input('m365_expiry_alert_enabled', false),
                'expiry_alert_days'           => max(1, min(365, (int) $request->input('m365_expiry_alert_days', 30))),
            ]);

            $newSecret = substr(trim((string) $request->input('m365_client_secret', '')), 0, 255);
            if ($newSecret) {
                $meta['client_secret'] = \Helper::encrypt($newSecret);
            }

            $mailbox->setMetaParam(self::META_KEY, $meta, true);
        }, 20, 2);

        // Render extra fields on the incoming connection settings page
        \Eventy::addAction('mailbox.connection_incoming.after_default_settings', function ($mailbox) {
            $meta = $mailbox->getMeta(self::META_KEY, []);
            $secretExpiry = $meta['secret_expiry'] ?? null;
            $expiryDate = $secretExpiry['date'] ?? ($meta['secret_expiry_date'] ?? '');
            $expiryAlertDays = (int) ($meta['expiry_alert_days'] ?? 30);
            if ($expiryAlertDays < 1) $expiryAlertDays = 30;

            echo \View::make('mail365::partials/incoming-settings', [
                'tenantId'              => $meta['tenant_id'] ?? '',
                'clientId'              => $meta['client_id'] ?? '',
                'authType'              => $meta['auth_type'] ?? 'secret',
                'hasSecret'             => !empty($meta['client_secret']),
                'certificateThumbprint' => $meta['certificate_thumbprint'] ?? '',
                'certificateExpiry'     => $meta['certificate_expiry'] ?? null,
                'sharedEmail'           => $meta['shared_mailbox_email'] ?? '',
                'fetchMode'             => $meta['fetch_mode'] ?? 'all',
                'postAction'            => $meta['post_fetch_action'] ?? 'none',
                'moveFolderId'          => $meta['post_fetch_move_folder'] ?? '',
                'moveFolderName'        => $meta['post_fetch_move_folder_name'] ?? '',
                'fetchFolders'          => $meta['fetch_folders'] ?? [],
                'expiryDate'            => $expiryDate,
                'daysLeft'              => $expiryDate ? (int) ceil((strtotime($expiryDate) - time()) / 86400) : 0,
                'expiryAlertEnabled'    => !empty($meta['expiry_alert_enabled']),
                'expiryAlertDays'       => $expiryAlertDays,
                'lastRun'               => $meta['last_fetch_run'] ?? null,
                'lastCount'             => $meta['last_fetch_count'] ?? null,
                'lastSuccess'           => $meta['last_fetch_success'] ?? null,
                'lastError'             => $meta['last_fetch_error'] ?? null,
                'lastErrorAt'           => $meta['last_fetch_error_at'] ?? null,
                'quotaUsage'            => $meta['quota_usage'] ?? null,
            ])->render();
        });

        // Add Mail 365 to the Manage menu
        \Eventy::addAction('menu.manage.append', function () {
            if (!\Auth::check() || !\Auth::user()->isAdmin()) {
                return;
            }
            echo '<li><a href="' . e(route('mail365.overview')) . '">Mail 365</a></li>';
        });

        // Global expiry warning banner on all admin pages
        \Eventy::addAction('layout.body_bottom', function () {
            if (!\Auth::check() || !\Auth::user()->isAdmin()) {
                return;
            }

            $cookieName = 'mail365_expiry_dismissed';
            if (!empty($_COOKIE[$cookieName]) && $_COOKIE[$cookieName] === date('Y-m-d')) {
                return;
            }

            $warnings = [];
            $mailboxes = Mailbox::all();

            foreach ($mailboxes as $mailbox) {
                $meta = $mailbox->getMeta(self::META_KEY, []);
                if (empty($meta['tenant_id'])) continue;

                $authType = $meta['auth_type'] ?? 'secret';
                $expiryDate = null;

                if ($authType === 'certificate') {
                    $expiryDate = $meta['certificate_expiry']['date'] ?? null;
                } else {
                    if (!empty($meta['secret_expiry']['date'])) {
                        $expiryDate = $meta['secret_expiry']['date'];
                    } elseif (!empty($meta['secret_expiry_date'])) {
                        $expiryDate = $meta['secret_expiry_date'];
                    }
                }

                if (!$expiryDate) continue;

                $daysLeft = (int) ceil((strtotime($expiryDate) - time()) / 86400);
                if ($daysLeft > 30) continue;

                $warnings[] = [
                    'mailbox'  => $mailbox->name,
                    'email'    => $mailbox->email,
                    'days'     => $daysLeft,
                    'date'     => $expiryDate,
                    'id'       => $mailbox->id,
                ];
            }

            if (empty($warnings)) return;

            echo \View::make('mail365::partials/expiry-banner', [
                'warnings' => $warnings,
            ])->render();
        });

        // Schedule the fetch command alongside freescout:fetch-emails
        \Eventy::addFilter('schedule', function ($schedule) {
            $schedule->command('mail365:fetch-emails')
                ->everyMinute()
                ->withoutOverlapping()
                ->sendOutputTo(storage_path() . '/logs/mail365-fetch.log');
            $schedule->command('mail365:retry-sends')
                ->everyFiveMinutes()
                ->withoutOverlapping();
            return $schedule;
        });
    }
}

@extends('layouts.app')

@section('title', __('Mail 365 Overview'))

@section('content')
<div class="container">
    <div class="section-heading">{{ __('Mail 365 Overview') }}</div>

    @if(empty($rows))
        <div class="alert alert-info">{{ __('No mailboxes are configured with Microsoft 365.') }}</div>
    @else
    <div class="table-responsive">
        <table class="table table-striped" id="mail365-overview-table">
            <thead>
                <tr>
                    <th>{{ __('Mailbox') }}</th>
                    <th>{{ __('Fetch Status') }}</th>
                    <th>{{ __('Send Status') }}</th>
                    <th>{{ __('Secret Expiry') }}</th>
                    <th>{{ __('Retry Queue') }}</th>
                    <th>{{ __('Mailbox Usage') }}</th>
                    <th></th>
                </tr>
            </thead>
            <tbody>
                @foreach($rows as $row)
                @php $mb = $row['mailbox']; @endphp
                <tr>
                    <td>
                        <strong>{{ $mb->name }}</strong><br>
                        <small class="text-muted">{{ $mb->email }}</small>
                        @if($row['shared_email'])
                            <br><small class="text-muted"><i class="glyphicon glyphicon-share-alt"></i> {{ $row['shared_email'] }}</small>
                        @endif
                    </td>
                    <td>
                        @if($row['last_fetch_error'])
                            <span class="text-danger"><i class="glyphicon glyphicon-remove"></i></span>
                            <span class="m365-time-ago text-danger" data-utc="{{ $row['last_fetch_error_at'] }}"></span>
                            <br><small class="text-danger">{{ mb_substr($row['last_fetch_error'], 0, 80) }}</small>
                        @elseif($row['last_fetch_success'])
                            <span class="text-success"><i class="glyphicon glyphicon-ok"></i></span>
                            <span class="m365-time-ago" data-utc="{{ $row['last_fetch_success'] }}"></span>
                        @else
                            <span class="text-muted">{{ __('No fetches yet') }}</span>
                        @endif
                    </td>
                    <td>
                        @if($row['last_send_error'])
                            <span class="text-danger"><i class="glyphicon glyphicon-remove"></i></span>
                            <small class="text-danger">{{ mb_substr($row['last_send_error'], 0, 60) }}</small>
                        @elseif($row['last_send_success'])
                            <span class="text-success"><i class="glyphicon glyphicon-ok"></i></span>
                            <span class="m365-time-ago" data-utc="{{ $row['last_send_success'] }}"></span>
                        @else
                            <span class="text-muted">-</span>
                        @endif
                    </td>
                    <td>
                        @if($row['secret_expiry_date'])
                            @if($row['secret_days_left'] < 0)
                                <span class="text-danger"><i class="glyphicon glyphicon-exclamation-sign"></i> <strong>{{ __('Expired') }}</strong></span>
                            @elseif($row['secret_days_left'] <= 14)
                                <span class="text-danger"><i class="glyphicon glyphicon-exclamation-sign"></i> {{ $row['secret_days_left'] }}d</span>
                            @elseif($row['secret_days_left'] <= 30)
                                <span class="text-warning"><i class="glyphicon glyphicon-warning-sign"></i> {{ $row['secret_days_left'] }}d</span>
                            @else
                                <span class="text-success"><i class="glyphicon glyphicon-ok"></i> {{ $row['secret_days_left'] }}d</span>
                            @endif
                            <br><small class="text-muted">{{ $row['secret_expiry_date'] }}</small>
                        @else
                            <span class="text-muted">-</span>
                        @endif
                    </td>
                    <td>
                        @if($row['retry_queue_count'] > 0)
                            <span class="label label-warning">{{ $row['retry_queue_count'] }}</span>
                        @else
                            <span class="text-muted">0</span>
                        @endif
                    </td>
                    <td>
                        @if($row['quota_usage'])
                            <strong>{{ $row['quota_usage']['used_display'] }}</strong>
                            <br><small class="text-muted">{{ $row['quota_usage']['folder_count'] }} folders</small>
                        @else
                            <span class="text-muted">-</span>
                        @endif
                    </td>
                    <td>
                        <a href="{{ url('/mailbox/connection-settings/' . $mb->id . '/incoming') }}" class="btn btn-default btn-xs" title="{{ __('Incoming settings') }}"><i class="glyphicon glyphicon-cog"></i></a>
                        <a href="{{ url('/mailbox/connection-settings/' . $mb->id . '/outgoing') }}" class="btn btn-default btn-xs" title="{{ __('Outgoing settings') }}"><i class="glyphicon glyphicon-send"></i></a>
                    </td>
                </tr>
                @endforeach
            </tbody>
        </table>
    </div>
    @endif
</div>
@endsection

<div id="m365-expiry-banner" style="position:fixed;top:0;left:0;right:0;z-index:10000;background:#fcf8e3;border-bottom:1px solid #f0c36d;padding:10px 20px;font-size:13px;color:#8a6d3b;">
    <button type="button" id="m365-expiry-dismiss" style="float:right;background:none;border:none;font-size:18px;line-height:1;color:#8a6d3b;cursor:pointer;padding:0 4px;">&times;</button>
    <i class="glyphicon glyphicon-warning-sign"></i>
    <strong>Microsoft 365:</strong>
    @foreach($warnings as $i => $w)
        @if($i > 0), @endif
        @if($w['days'] < 0)
            <span style="color:#a94442;">{{ $w['mailbox'] }} secret <strong>expired</strong></span>
        @elseif($w['days'] <= 14)
            <span style="color:#a94442;">{{ $w['mailbox'] }} secret expires in <strong>{{ $w['days'] }} day{{ $w['days'] !== 1 ? 's' : '' }}</strong></span>
        @else
            {{ $w['mailbox'] }} secret expires in <strong>{{ $w['days'] }} day{{ $w['days'] !== 1 ? 's' : '' }}</strong>
        @endif
        (<a href="{{ url('/mailbox/connection-settings/' . $w['id'] . '/incoming') }}">settings</a>)
    @endforeach
    <script>
    document.getElementById('m365-expiry-dismiss').addEventListener('click', function() {
        document.getElementById('m365-expiry-banner').style.display = 'none';
        document.cookie = 'mail365_expiry_dismissed={{ date("Y-m-d") }};path=/;max-age=86400;SameSite=Lax;Secure';
    });
    </script>
</div>

<?php
declare(strict_types=1);

require_once __DIR__ . '/_bootstrap.php';

cws_require_method('POST');
cws_assert_same_origin_for_write();
cws_revoke_current_session();

cws_json(['ok' => true]);

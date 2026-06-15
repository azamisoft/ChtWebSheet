<?php
declare(strict_types=1);

require_once __DIR__ . '/_bootstrap.php';

cws_require_method('GET');

$registEnabled = cws_regist_enabled();
$user = cws_current_user();
cws_json([
    'ok' => true,
    'registEnabled' => $registEnabled,
    'user' => $user ? cws_public_user($user) : null,
]);

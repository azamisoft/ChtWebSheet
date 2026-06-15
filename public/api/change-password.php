<?php
declare(strict_types=1);

require_once __DIR__ . '/_bootstrap.php';

cws_require_method('POST');
cws_assert_same_origin_for_write();

$user = cws_require_user();
$body = cws_read_json_body();
$currentPassword = (string) ($body['currentPassword'] ?? '');
$newPassword = cws_validate_password((string) ($body['newPassword'] ?? $body['password'] ?? ''));

$passwordHash = (string) ($user['password_hash'] ?? '');
if ($passwordHash === '' || !password_verify($currentPassword, $passwordHash)) {
    cws_fail(401, 'invalid_current_password');
}

cws_update_user_password((int) $user['id'], $newPassword);
$currentSessionId = (int) ($user['session_id'] ?? 0);
cws_revoke_user_sessions((int) $user['id'], $currentSessionId > 0 ? $currentSessionId : null);

$freshUser = cws_find_user_by_email((string) $user['email']) ?: $user;
cws_json([
    'ok' => true,
    'user' => cws_public_user($freshUser),
]);

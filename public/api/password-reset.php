<?php
declare(strict_types=1);

require_once __DIR__ . '/_bootstrap.php';

cws_require_method('POST');
cws_assert_same_origin_for_write();

$body = cws_read_json_body();
$email = cws_validate_email((string) ($body['email'] ?? ''));
$newPassword = cws_validate_password((string) ($body['newPassword'] ?? $body['password'] ?? ''));
$code = trim((string) ($body['code'] ?? ''));

$user = cws_find_user_by_email($email);
if (!$user || (string) $user['status'] !== 'active') {
    cws_fail(404, 'user_not_found');
}

if ($code === '') {
    $verification = cws_issue_email_code($email, 'password_reset');
    cws_json([
        'ok' => true,
        'verificationRequired' => true,
        'purpose' => 'password_reset',
        'email' => $email,
        'expiresAt' => $verification['expiresAt'],
    ]);
}

cws_verify_email_code($email, 'password_reset', $code);
cws_update_user_password((int) $user['id'], $newPassword);
cws_revoke_user_sessions((int) $user['id']);

$stmt = cws_db()->prepare('UPDATE cws_users SET last_login_at = UTC_TIMESTAMP(), updated_at = UTC_TIMESTAMP() WHERE id = ?');
$stmt->execute([(int) $user['id']]);

$freshUser = cws_find_user_by_email($email) ?: $user;
$session = cws_create_session_for_user((int) $user['id']);

cws_json([
    'ok' => true,
    'user' => cws_public_user($freshUser),
    'authToken' => $session['token'],
    'session' => [
        'expiresAt' => $session['expiresAt'],
    ],
]);

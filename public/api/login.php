<?php
declare(strict_types=1);

require_once __DIR__ . '/_bootstrap.php';

cws_require_method('POST');
cws_assert_same_origin_for_write();

$body = cws_read_json_body();
$email = cws_validate_email((string) ($body['email'] ?? ''));
$password = (string) ($body['password'] ?? '');

$user = cws_find_user_by_email($email);
if (!$user || (string) $user['status'] !== 'active') {
    cws_fail(401, 'invalid_credentials');
}

$passwordHash = (string) ($user['password_hash'] ?? '');
if ($passwordHash === '' || !password_verify($password, $passwordHash)) {
    cws_fail(401, 'invalid_credentials');
}

if (password_needs_rehash($passwordHash, PASSWORD_DEFAULT)) {
    $newHash = password_hash($password, PASSWORD_DEFAULT);
    if (is_string($newHash) && $newHash !== '') {
        $stmt = cws_db()->prepare('UPDATE cws_users SET password_hash = ?, password_changed_at = UTC_TIMESTAMP(), updated_at = UTC_TIMESTAMP() WHERE id = ?');
        $stmt->execute([$newHash, (int) $user['id']]);
    }
}

$stmt = cws_db()->prepare('UPDATE cws_users SET last_login_at = UTC_TIMESTAMP(), updated_at = UTC_TIMESTAMP() WHERE id = ?');
$stmt->execute([(int) $user['id']]);

$session = cws_create_session_for_user((int) $user['id']);

$freshUser = cws_find_user_by_email($email) ?: $user;
cws_json([
    'ok' => true,
    'user' => cws_public_user($freshUser),
    'authToken' => $session['token'],
    'session' => [
        'expiresAt' => $session['expiresAt'],
    ],
]);

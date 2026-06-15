<?php
declare(strict_types=1);

require_once __DIR__ . '/_bootstrap.php';

cws_require_method('POST');
cws_assert_same_origin_for_write();
cws_require_regist_enabled();

$body = cws_read_json_body();
$email = cws_validate_email((string) ($body['email'] ?? ''));
$password = cws_validate_password((string) ($body['password'] ?? ''));
$displayName = trim((string) ($body['displayName'] ?? ''));
$code = trim((string) ($body['code'] ?? ''));
if (cws_text_length($displayName) > 120) {
    cws_fail(400, 'display_name_too_long');
}

if (cws_find_user_by_email($email)) {
    cws_fail(409, 'email_already_registered');
}

if ($code === '') {
    $verification = cws_issue_email_code($email, 'register');
    cws_json([
        'ok' => true,
        'verificationRequired' => true,
        'purpose' => 'register',
        'email' => $email,
        'expiresAt' => $verification['expiresAt'],
    ]);
}

cws_verify_email_code($email, 'register', $code);

$hash = password_hash($password, PASSWORD_DEFAULT);
if (!is_string($hash) || $hash === '') {
    cws_fail(500, 'password_hash_failed');
}

try {
    $stmt = cws_db()->prepare(
        'INSERT INTO cws_users
            (email, email_normalized, password_hash, display_name, status, email_verified_at, password_set_at, created_at, updated_at)
         VALUES
            (?, ?, ?, ?, ?, UTC_TIMESTAMP(), UTC_TIMESTAMP(), UTC_TIMESTAMP(), UTC_TIMESTAMP())'
    );
    $stmt->execute([
        $email,
        cws_normalize_email($email),
        $hash,
        $displayName !== '' ? $displayName : null,
        'active',
    ]);
} catch (PDOException $e) {
    if ($e->getCode() === '23000') {
        cws_fail(409, 'email_already_registered');
    }
    cws_fail(500, 'register_failed');
}

$user = cws_find_user_by_email($email);
if (!$user) {
    cws_fail(500, 'register_failed');
}

$session = cws_create_session_for_user((int) $user['id']);

cws_json([
    'ok' => true,
    'user' => cws_public_user($user),
    'authToken' => $session['token'],
    'session' => [
        'expiresAt' => $session['expiresAt'],
    ],
]);

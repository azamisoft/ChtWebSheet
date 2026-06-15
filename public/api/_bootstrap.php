<?php
declare(strict_types=1);

if (basename((string) ($_SERVER['SCRIPT_NAME'] ?? '')) === '_bootstrap.php') {
    http_response_code(404);
    exit;
}

const CWS_DEFAULT_COOKIE_NAME = 'cws_session';
const CWS_DEFAULT_COOKIE_PATH = '/cws';
const CWS_SESSION_TTL_SECONDS = 2592000; // 30 days
const CWS_MAX_JSON_BYTES = 1048576;
const CWS_DEFAULT_ALLOWED_API_ORIGINS = 'null,http://localhost:5173,http://127.0.0.1:5173,https://chtec.co.jp';

cws_load_wp_config();
cws_send_base_headers();

function cws_load_wp_config(): void
{
    if (defined('CWS_DB_NAME') && defined('CWS_AUTH_SECRET')) {
        return;
    }

    $dir = __DIR__;
    for ($i = 0; $i < 10; $i += 1) {
        $candidate = $dir . '/wp-config.php';
        if (is_file($candidate)) {
            require_once $candidate;
            return;
        }
        $parent = dirname($dir);
        if ($parent === $dir) {
            break;
        }
        $dir = $parent;
    }
}

function cws_send_base_headers(): void
{
    header('Content-Type: application/json; charset=utf-8');
    header('Cache-Control: no-store');
    header('X-Content-Type-Options: nosniff');
    header('Referrer-Policy: same-origin');

    $origin = (string) ($_SERVER['HTTP_ORIGIN'] ?? '');
    if ($origin !== '' && cws_cors_origin_allowed($origin)) {
        header('Access-Control-Allow-Origin: ' . $origin);
        header('Access-Control-Allow-Headers: Authorization, Content-Type, X-Requested-With');
        header('Access-Control-Allow-Methods: GET, POST, OPTIONS');
        header('Access-Control-Max-Age: 600');
        header('Vary: Origin');
    }

    if (strtoupper((string) ($_SERVER['REQUEST_METHOD'] ?? 'GET')) === 'OPTIONS') {
        http_response_code(204);
        exit;
    }
}

function cws_cors_origin_allowed(string $origin): bool
{
    $origin = trim($origin);
    if ($origin === '') {
        return false;
    }

    $allowed = array_filter(array_map(
        'trim',
        explode(',', cws_optional_config('CWS_ALLOWED_API_ORIGINS', CWS_DEFAULT_ALLOWED_API_ORIGINS))
    ));
    if (in_array($origin, $allowed, true)) {
        return true;
    }

    $originHost = parse_url($origin, PHP_URL_HOST);
    $requestHost = explode(':', (string) ($_SERVER['HTTP_HOST'] ?? ''), 2)[0];
    return $originHost && $requestHost && strcasecmp($originHost, $requestHost) === 0;
}

function cws_config(string $name, ?string $default = null): string
{
    if (defined($name)) {
        return (string) constant($name);
    }
    $value = getenv($name);
    if ($value !== false) {
        return (string) $value;
    }
    if ($default !== null) {
        return $default;
    }
    cws_fail(500, 'missing_config', "Missing server config: {$name}");
}

function cws_defined_config(string $name): ?string
{
    if (defined($name)) {
        $value = (string) constant($name);
        return $value !== '' ? $value : null;
    }
    $value = getenv($name);
    if ($value === false || (string) $value === '') {
        return null;
    }
    return (string) $value;
}

function cws_optional_config(string $name, string $default): string
{
    return cws_config($name, $default);
}

function cws_config_bool(string $name, bool $default = false): bool
{
    if (defined($name)) {
        $value = constant($name);
    } else {
        $value = getenv($name);
        if ($value === false) {
            return $default;
        }
    }

    if (is_bool($value)) {
        return $value;
    }
    if (is_int($value) || is_float($value)) {
        return (int) $value !== 0;
    }

    $normalized = strtolower(trim((string) $value));
    if ($normalized === '') {
        return $default;
    }
    if (in_array($normalized, ['1', 'true', 'yes', 'on', 'enabled'], true)) {
        return true;
    }
    if (in_array($normalized, ['0', 'false', 'no', 'off', 'disabled'], true)) {
        return false;
    }
    return $default;
}

function cws_regist_enabled(): bool
{
    return cws_config_bool('CWS_REGIST_ENABLED', false);
}

function cws_require_regist_enabled(): void
{
    if (!cws_regist_enabled()) {
        cws_fail(403, 'regist_disabled');
    }
}

function cws_cookie_name(): string
{
    return cws_optional_config('CWS_SESSION_COOKIE', CWS_DEFAULT_COOKIE_NAME);
}

function cws_cookie_path(): string
{
    return cws_optional_config('CWS_COOKIE_PATH', CWS_DEFAULT_COOKIE_PATH);
}

function cws_json(array $payload, int $status = 200): void
{
    if (function_exists('status_header')) {
        status_header($status);
    }
    http_response_code($status);
    echo json_encode($payload, JSON_UNESCAPED_UNICODE | JSON_UNESCAPED_SLASHES);
    exit;
}

function cws_fail(int $status, string $error, string $message = ''): void
{
    $payload = ['ok' => false, 'error' => $error];
    if ($message !== '') {
        $payload['message'] = $message;
    }
    cws_json($payload, $status);
}

function cws_require_method(string $method): void
{
    if (strtoupper((string) ($_SERVER['REQUEST_METHOD'] ?? 'GET')) !== strtoupper($method)) {
        header('Allow: ' . strtoupper($method));
        cws_fail(405, 'method_not_allowed');
    }
}

function cws_assert_same_origin_for_write(): void
{
    $method = strtoupper((string) ($_SERVER['REQUEST_METHOD'] ?? 'GET'));
    if (!in_array($method, ['POST', 'PUT', 'PATCH', 'DELETE'], true)) {
        return;
    }

    $origin = (string) ($_SERVER['HTTP_ORIGIN'] ?? '');
    if ($origin === '') {
        return;
    }
    if (cws_cors_origin_allowed($origin)) {
        return;
    }

    $originHost = parse_url($origin, PHP_URL_HOST);
    $requestHost = explode(':', (string) ($_SERVER['HTTP_HOST'] ?? ''), 2)[0];
    if (!$originHost || !$requestHost || strcasecmp($originHost, $requestHost) !== 0) {
        cws_fail(403, 'bad_origin');
    }
}

function cws_read_json_body(): array
{
    $length = (int) ($_SERVER['CONTENT_LENGTH'] ?? 0);
    if ($length > CWS_MAX_JSON_BYTES) {
        cws_fail(413, 'request_too_large');
    }

    $raw = file_get_contents('php://input');
    if ($raw === false || trim($raw) === '') {
        return [];
    }
    if (strlen($raw) > CWS_MAX_JSON_BYTES) {
        cws_fail(413, 'request_too_large');
    }

    $data = json_decode($raw, true);
    if (!is_array($data)) {
        cws_fail(400, 'invalid_json');
    }
    return $data;
}

function cws_db(): PDO
{
    static $pdo = null;
    if ($pdo instanceof PDO) {
        return $pdo;
    }

    $host = cws_config('CWS_DB_HOST');
    $name = cws_config('CWS_DB_NAME');
    $user = cws_config('CWS_DB_USER');
    $password = cws_config('CWS_DB_PASSWORD');
    $dsn = "mysql:host={$host};dbname={$name};charset=utf8mb4";

    try {
        $pdo = new PDO($dsn, $user, $password, [
            PDO::ATTR_ERRMODE => PDO::ERRMODE_EXCEPTION,
            PDO::ATTR_DEFAULT_FETCH_MODE => PDO::FETCH_ASSOC,
            PDO::ATTR_EMULATE_PREPARES => false,
        ]);
    } catch (Throwable $e) {
        cws_fail(500, 'db_connect_failed');
    }
    return $pdo;
}

function cws_normalize_email(string $email): string
{
    return strtolower(trim($email));
}

function cws_validate_email(string $email): string
{
    $normalized = cws_normalize_email($email);
    if ($normalized === '' || !filter_var($normalized, FILTER_VALIDATE_EMAIL) || strlen($normalized) > 255) {
        cws_fail(400, 'invalid_email');
    }
    return $normalized;
}

function cws_validate_password(string $password): string
{
    if (strlen($password) < 8) {
        cws_fail(400, 'password_too_short', 'Password must be at least 8 characters.');
    }
    if (strlen($password) > 512) {
        cws_fail(400, 'password_too_long');
    }
    return $password;
}

function cws_validate_email_code(string $code): string
{
    $normalized = preg_replace('/\s+/', '', trim($code)) ?: '';
    if (!preg_match('/^\d{6}$/', $normalized)) {
        cws_fail(400, 'invalid_code');
    }
    return $normalized;
}

function cws_text_length(string $text): int
{
    if (function_exists('mb_strlen')) {
        return mb_strlen($text, 'UTF-8');
    }
    return strlen($text);
}

function cws_hash_secret(string $value): string
{
    return hash_hmac('sha256', $value, cws_config('CWS_AUTH_SECRET'));
}

function cws_email_code_hash(string $email, string $purpose, string $code): string
{
    return cws_hash_secret(cws_normalize_email($email) . ':' . cws_email_code_purpose($purpose) . ':' . $code);
}

function cws_now(): string
{
    return gmdate('Y-m-d H:i:s');
}

function cws_datetime_from_unix(int $timestamp): string
{
    return gmdate('Y-m-d H:i:s', $timestamp);
}

function cws_mail_timezone(): DateTimeZone
{
    $timezone = cws_optional_config('CWS_MAIL_TIMEZONE', 'Asia/Tokyo');
    try {
        return new DateTimeZone($timezone);
    } catch (Throwable $e) {
        return new DateTimeZone('Asia/Tokyo');
    }
}

function cws_format_mail_datetime(string $utcDatetime): string
{
    $sourceTimezone = new DateTimeZone('UTC');
    $date = DateTimeImmutable::createFromFormat('Y-m-d H:i:s', $utcDatetime, $sourceTimezone);
    if (!$date) {
        try {
            $date = new DateTimeImmutable($utcDatetime, $sourceTimezone);
        } catch (Throwable $e) {
            return $utcDatetime . ' UTC';
        }
    }
    return $date->setTimezone(cws_mail_timezone())->format('Y-m-d H:i T');
}

function cws_is_https(): bool
{
    if (!empty($_SERVER['HTTPS']) && strtolower((string) $_SERVER['HTTPS']) !== 'off') {
        return true;
    }
    return strtolower((string) ($_SERVER['HTTP_X_FORWARDED_PROTO'] ?? '')) === 'https';
}

function cws_set_session_cookie(string $token, int $expiresAt): void
{
    setcookie(cws_cookie_name(), $token, [
        'expires' => $expiresAt,
        'path' => cws_cookie_path(),
        'secure' => cws_is_https(),
        'httponly' => true,
        'samesite' => 'Lax',
    ]);
}

function cws_clear_session_cookie(): void
{
    setcookie(cws_cookie_name(), '', [
        'expires' => time() - 3600,
        'path' => cws_cookie_path(),
        'secure' => cws_is_https(),
        'httponly' => true,
        'samesite' => 'Lax',
    ]);
}

function cws_public_user(array $user): array
{
    return [
        'id' => (int) $user['id'],
        'email' => (string) $user['email'],
        'displayName' => $user['display_name'] ?? null,
        'status' => (string) $user['status'],
        'emailVerifiedAt' => $user['email_verified_at'] ?? null,
        'createdAt' => $user['created_at'] ?? null,
        'lastLoginAt' => $user['last_login_at'] ?? null,
    ];
}

function cws_email_code_purpose(string $purpose): string
{
    $normalized = strtolower(trim($purpose));
    if (!in_array($normalized, ['login', 'register', 'password_reset'], true)) {
        cws_fail(400, 'invalid_purpose');
    }
    return $normalized;
}

function cws_email_code_ttl_seconds(): int
{
    $value = (int) cws_optional_config('CWS_EMAIL_CODE_TTL_SECONDS', '600');
    return min(max($value, 60), 3600);
}

function cws_email_code_max_attempts(): int
{
    $value = (int) cws_optional_config('CWS_EMAIL_CODE_MAX_ATTEMPTS', '5');
    return min(max($value, 1), 20);
}

function cws_issue_email_code(string $email, string $purpose): array
{
    $purpose = cws_email_code_purpose($purpose);
    $email = cws_validate_email($email);
    $code = str_pad((string) random_int(0, 999999), 6, '0', STR_PAD_LEFT);
    $expiresAt = cws_datetime_from_unix(time() + cws_email_code_ttl_seconds());
    $codeHash = cws_email_code_hash($email, $purpose, $code);
    $pdo = cws_db();

    $pdo->beginTransaction();
    try {
        $consumeOld = $pdo->prepare(
            'UPDATE cws_email_codes
                SET consumed_at = UTC_TIMESTAMP()
              WHERE email_normalized = ?
                AND purpose = ?
                AND consumed_at IS NULL'
        );
        $consumeOld->execute([cws_normalize_email($email), $purpose]);

        $insert = $pdo->prepare(
            'INSERT INTO cws_email_codes
                (email_normalized, code_hash, purpose, attempts, created_at, expires_at, consumed_at)
             VALUES
                (?, ?, ?, 0, UTC_TIMESTAMP(), ?, NULL)'
        );
        $insert->execute([cws_normalize_email($email), $codeHash, $purpose, $expiresAt]);
        $codeId = (int) $pdo->lastInsertId();
        $pdo->commit();
    } catch (Throwable $e) {
        $pdo->rollBack();
        cws_fail(500, 'code_create_failed');
    }

    try {
        cws_send_email_code($email, $purpose, $code, $expiresAt);
    } catch (Throwable $e) {
        $markFailed = $pdo->prepare('UPDATE cws_email_codes SET consumed_at = UTC_TIMESTAMP() WHERE id = ?');
        $markFailed->execute([$codeId]);
        $message = $e->getMessage();
        cws_fail(500, $message === 'smtp_config_missing' ? 'smtp_config_missing' : 'code_send_failed');
    }

    return [
        'email' => $email,
        'purpose' => $purpose,
        'expiresAt' => $expiresAt,
    ];
}

function cws_verify_email_code(string $email, string $purpose, string $code): void
{
    $purpose = cws_email_code_purpose($purpose);
    $email = cws_validate_email($email);
    $code = cws_validate_email_code($code);

    $stmt = cws_db()->prepare(
        'SELECT *
           FROM cws_email_codes
          WHERE email_normalized = ?
            AND purpose = ?
            AND consumed_at IS NULL
            AND expires_at > UTC_TIMESTAMP()
          ORDER BY id DESC
          LIMIT 1'
    );
    $stmt->execute([cws_normalize_email($email), $purpose]);
    $row = $stmt->fetch();
    if (!is_array($row)) {
        cws_fail(400, 'code_expired');
    }

    if ((int) ($row['attempts'] ?? 0) >= cws_email_code_max_attempts()) {
        cws_fail(429, 'code_too_many_attempts');
    }

    $expected = (string) ($row['code_hash'] ?? '');
    if ($expected === '' || !hash_equals($expected, cws_email_code_hash($email, $purpose, $code))) {
        $update = cws_db()->prepare('UPDATE cws_email_codes SET attempts = attempts + 1 WHERE id = ?');
        $update->execute([(int) $row['id']]);
        cws_fail(400, 'invalid_code');
    }

    $consume = cws_db()->prepare('UPDATE cws_email_codes SET consumed_at = UTC_TIMESTAMP() WHERE id = ?');
    $consume->execute([(int) $row['id']]);
}

function cws_send_email_code(string $email, string $purpose, string $code, string $expiresAt): void
{
    $purpose = cws_email_code_purpose($purpose);
    $labels = [
        'login' => 'ログイン',
        'register' => '登録',
        'password_reset' => 'パスワード再設定',
    ];
    $label = $labels[$purpose] ?? 'ログイン';
    $expiresAtLabel = cws_format_mail_datetime($expiresAt);
    $subject = "Cht WebSheet {$label}確認コード";
    $body = implode("\n", [
        "Cht WebSheet の{$label}確認コードです。",
        '',
        "確認コード: {$code}",
        '',
        "有効期限: {$expiresAtLabel}",
        '',
        'このメールに心当たりがない場合は、破棄してください。',
    ]);
    cws_send_mail_message($email, $subject, $body);
}

function cws_send_mail_message(string $to, string $subject, string $body): void
{
    $config = cws_smtp_config();
    $from = $config['from'];
    $fromName = $config['fromName'];
    $domain = cws_mail_domain($from);
    $now = new DateTimeImmutable('now', cws_mail_timezone());
    $headers = [
        'Date: ' . $now->format(DATE_RFC2822),
        'From: ' . cws_mailbox_header($from, $fromName),
        'To: ' . cws_mailbox_header($to, ''),
        'Subject: ' . cws_mime_header($subject),
        'Message-ID: <' . bin2hex(random_bytes(16)) . '@' . $domain . '>',
        'MIME-Version: 1.0',
        'Content-Type: text/plain; charset=UTF-8',
        'Content-Transfer-Encoding: 8bit',
    ];

    $message = implode("\r\n", $headers) . "\r\n\r\n" . str_replace(["\r\n", "\r"], "\n", $body);
    cws_smtp_send($config, $to, $message);
}

function cws_smtp_config(): array
{
    $host = cws_defined_config('CWS_SMTP_HOST');
    $user = cws_defined_config('CWS_SMTP_USER');
    $password = cws_defined_config('CWS_SMTP_PASSWORD');
    if (!$host || !$user || !$password) {
        throw new RuntimeException('smtp_config_missing');
    }

    $port = (int) (cws_defined_config('CWS_SMTP_PORT') ?? '587');
    $secure = strtolower((string) (cws_defined_config('CWS_SMTP_SECURE') ?? 'tls'));
    if (!in_array($secure, ['tls', 'ssl', ''], true)) {
        $secure = 'tls';
    }

    $from = cws_defined_config('CWS_MAIL_FROM') ?? $user;
    if (!filter_var($from, FILTER_VALIDATE_EMAIL)) {
        throw new RuntimeException('smtp_config_missing');
    }

    return [
        'host' => $host,
        'port' => $port > 0 ? $port : 587,
        'secure' => $secure,
        'user' => $user,
        'password' => $password,
        'from' => $from,
        'fromName' => cws_defined_config('CWS_MAIL_FROM_NAME') ?? 'Cht WebSheet',
    ];
}

function cws_smtp_send(array $config, string $to, string $message): void
{
    if (!filter_var($to, FILTER_VALIDATE_EMAIL)) {
        throw new RuntimeException('invalid_recipient');
    }

    $target = ($config['secure'] === 'ssl' ? 'ssl://' : '') . $config['host'] . ':' . (int) $config['port'];
    $errno = 0;
    $errstr = '';
    $socket = @stream_socket_client($target, $errno, $errstr, 20, STREAM_CLIENT_CONNECT);
    if (!$socket) {
        throw new RuntimeException('smtp_connect_failed');
    }
    stream_set_timeout($socket, 20);

    try {
        cws_smtp_expect($socket, [220]);
        cws_smtp_command($socket, 'EHLO ' . cws_smtp_ehlo_name(), [250]);
        if ($config['secure'] === 'tls') {
            cws_smtp_command($socket, 'STARTTLS', [220]);
            if (!stream_socket_enable_crypto($socket, true, STREAM_CRYPTO_METHOD_TLS_CLIENT)) {
                throw new RuntimeException('smtp_tls_failed');
            }
            cws_smtp_command($socket, 'EHLO ' . cws_smtp_ehlo_name(), [250]);
        }
        cws_smtp_command($socket, 'AUTH LOGIN', [334]);
        cws_smtp_command($socket, base64_encode((string) $config['user']), [334]);
        cws_smtp_command($socket, base64_encode((string) $config['password']), [235]);
        cws_smtp_command($socket, 'MAIL FROM:<' . $config['from'] . '>', [250]);
        cws_smtp_command($socket, 'RCPT TO:<' . $to . '>', [250, 251]);
        cws_smtp_command($socket, 'DATA', [354]);
        fwrite($socket, cws_smtp_dot_stuff($message) . "\r\n.\r\n");
        cws_smtp_expect($socket, [250]);
        cws_smtp_command($socket, 'QUIT', [221, 250]);
    } finally {
        fclose($socket);
    }
}

function cws_smtp_command($socket, string $command, array $expectedCodes): string
{
    fwrite($socket, $command . "\r\n");
    return cws_smtp_expect($socket, $expectedCodes);
}

function cws_smtp_expect($socket, array $expectedCodes): string
{
    [$code, $response] = cws_smtp_read_response($socket);
    if (!in_array($code, $expectedCodes, true)) {
        throw new RuntimeException('smtp_response_' . $code);
    }
    return $response;
}

function cws_smtp_read_response($socket): array
{
    $response = '';
    $code = 0;
    while (!feof($socket)) {
        $line = fgets($socket, 2048);
        if ($line === false) {
            break;
        }
        $response .= $line;
        if (preg_match('/^(\d{3})([\s-])/', $line, $m)) {
            $code = (int) $m[1];
            if ($m[2] === ' ') {
                return [$code, $response];
            }
        }
    }
    throw new RuntimeException('smtp_no_response');
}

function cws_smtp_dot_stuff(string $message): string
{
    $normalized = str_replace(["\r\n", "\r"], "\n", $message);
    $normalized = preg_replace('/^\./m', '..', $normalized) ?: $normalized;
    return str_replace("\n", "\r\n", $normalized);
}

function cws_smtp_ehlo_name(): string
{
    $host = (string) ($_SERVER['SERVER_NAME'] ?? 'localhost');
    return preg_replace('/[^A-Za-z0-9.-]/', '', $host) ?: 'localhost';
}

function cws_mime_header(string $value): string
{
    return '=?UTF-8?B?' . base64_encode($value) . '?=';
}

function cws_mailbox_header(string $email, string $name): string
{
    if ($name === '') {
        return '<' . $email . '>';
    }
    return cws_mime_header($name) . ' <' . $email . '>';
}

function cws_mail_domain(string $email): string
{
    $parts = explode('@', $email, 2);
    return isset($parts[1]) && $parts[1] !== '' ? $parts[1] : 'localhost';
}

function cws_find_user_by_email(string $email): ?array
{
    $stmt = cws_db()->prepare('SELECT * FROM cws_users WHERE email_normalized = ? LIMIT 1');
    $stmt->execute([cws_normalize_email($email)]);
    $row = $stmt->fetch();
    return is_array($row) ? $row : null;
}

function cws_update_user_password(int $userId, string $password): void
{
    $hash = password_hash($password, PASSWORD_DEFAULT);
    if (!is_string($hash) || $hash === '') {
        cws_fail(500, 'password_hash_failed');
    }

    $stmt = cws_db()->prepare(
        'UPDATE cws_users
            SET password_hash = ?,
                password_set_at = COALESCE(password_set_at, UTC_TIMESTAMP()),
                password_changed_at = UTC_TIMESTAMP(),
                updated_at = UTC_TIMESTAMP()
          WHERE id = ?'
    );
    $stmt->execute([$hash, $userId]);
}

function cws_revoke_user_sessions(int $userId, ?int $exceptSessionId = null): void
{
    if ($exceptSessionId !== null) {
        $stmt = cws_db()->prepare(
            'UPDATE cws_sessions
                SET revoked_at = UTC_TIMESTAMP()
              WHERE user_id = ?
                AND id <> ?
                AND revoked_at IS NULL'
        );
        $stmt->execute([$userId, $exceptSessionId]);
        return;
    }

    $stmt = cws_db()->prepare(
        'UPDATE cws_sessions
            SET revoked_at = UTC_TIMESTAMP()
          WHERE user_id = ?
            AND revoked_at IS NULL'
    );
    $stmt->execute([$userId]);
}

function cws_current_session_token(): string
{
    $authorization = (string) ($_SERVER['HTTP_AUTHORIZATION'] ?? $_SERVER['REDIRECT_HTTP_AUTHORIZATION'] ?? '');
    if (preg_match('/^\s*Bearer\s+([A-Za-z0-9]{64})\s*$/', $authorization, $m)) {
        return $m[1];
    }
    return (string) ($_COOKIE[cws_cookie_name()] ?? '');
}

function cws_current_user(): ?array
{
    $token = cws_current_session_token();
    if ($token === '') {
        return null;
    }

    $stmt = cws_db()->prepare(
        'SELECT u.*, s.id AS session_id
           FROM cws_sessions s
           JOIN cws_users u ON u.id = s.user_id
          WHERE s.session_hash = ?
            AND s.revoked_at IS NULL
            AND s.expires_at > UTC_TIMESTAMP()
            AND u.status = ?
          LIMIT 1'
    );
    $stmt->execute([cws_hash_secret($token), 'active']);
    $row = $stmt->fetch();
    if (!is_array($row)) {
        return null;
    }

    $update = cws_db()->prepare('UPDATE cws_sessions SET last_seen_at = UTC_TIMESTAMP() WHERE id = ?');
    $update->execute([(int) $row['session_id']]);

    return $row;
}

function cws_require_user(): array
{
    $user = cws_current_user();
    if (!$user) {
        cws_fail(401, 'not_authenticated');
    }
    return $user;
}

function cws_create_session_for_user(int $userId): array
{
    $token = bin2hex(random_bytes(32));
    $expiresAt = time() + CWS_SESSION_TTL_SECONDS;

    $stmt = cws_db()->prepare(
        'INSERT INTO cws_sessions
            (user_id, session_hash, created_at, expires_at, last_seen_at, user_agent, ip_address)
         VALUES
            (?, ?, UTC_TIMESTAMP(), ?, UTC_TIMESTAMP(), ?, ?)'
    );
    $stmt->execute([
        $userId,
        cws_hash_secret($token),
        cws_datetime_from_unix($expiresAt),
        substr((string) ($_SERVER['HTTP_USER_AGENT'] ?? ''), 0, 255),
        cws_packed_ip((string) ($_SERVER['REMOTE_ADDR'] ?? '')),
    ]);

    cws_set_session_cookie($token, $expiresAt);
    return ['token' => $token, 'expiresAt' => cws_datetime_from_unix($expiresAt)];
}

function cws_packed_ip(string $ip): ?string
{
    if ($ip === '') {
        return null;
    }
    $packed = @inet_pton($ip);
    return $packed === false ? null : $packed;
}

function cws_revoke_current_session(): void
{
    $token = cws_current_session_token();
    if ($token !== '') {
        $stmt = cws_db()->prepare('UPDATE cws_sessions SET revoked_at = UTC_TIMESTAMP() WHERE session_hash = ? AND revoked_at IS NULL');
        $stmt->execute([cws_hash_secret($token)]);
    }
    cws_clear_session_cookie();
}

function cws_json_string(array $body, string $key, int $maxLength = 20000): string
{
    $value = $body[$key] ?? '';
    if (is_array($value) || is_object($value)) {
        $value = json_encode($value, JSON_UNESCAPED_UNICODE | JSON_UNESCAPED_SLASHES);
    }
    $text = trim((string) $value);
    if ($text === '') {
        cws_fail(400, "missing_{$key}");
    }
    if (cws_text_length($text) > $maxLength) {
        cws_fail(400, "{$key}_too_long");
    }
    return $text;
}

function cws_optional_json_string(array $body, string $key, int $maxLength = 20000): string
{
    if (!array_key_exists($key, $body) || $body[$key] === null) {
        return '';
    }
    $value = $body[$key];
    if (is_array($value) || is_object($value)) {
        $value = json_encode($value, JSON_UNESCAPED_UNICODE | JSON_UNESCAPED_SLASHES);
    }
    $text = trim((string) $value);
    if (cws_text_length($text) > $maxLength) {
        cws_fail(400, "{$key}_too_long");
    }
    return $text;
}

function cws_normalize_workbook_key(string $workbookKey): string
{
    $key = trim($workbookKey);
    if ($key === '') {
        return '__default__';
    }
    return substr(preg_replace('/[^A-Za-z0-9._:-]/', '_', $key) ?: '__default__', 0, 128);
}

function cws_devlune_user_key(array $user, string $workbookKey): string
{
    $material = (string) $user['id'] . ':' . $workbookKey;
    $hash = substr(hash_hmac('sha256', $material, cws_config('CWS_AUTH_SECRET')), 0, 40);
    return "cws-{$hash}@cws.local";
}

function cws_devlune_json(string $method, string $path, array $body = [], array $query = [], int $timeoutSeconds = 120): array
{
    $base = rtrim(cws_config('CWS_DEVLUNE_BASE_URL'), '/');
    $url = $base . $path;
    if ($query) {
        $url .= '?' . http_build_query($query);
    }

    $payload = $body ? json_encode($body, JSON_UNESCAPED_UNICODE | JSON_UNESCAPED_SLASHES) : '';
    $headers = [
        'Authorization: Bearer ' . cws_config('CWS_DEVLUNE_API_KEY'),
        'Accept: application/json',
    ];
    if ($payload !== '') {
        $headers[] = 'Content-Type: application/json; charset=utf-8';
    }

    if (function_exists('curl_init')) {
        $ch = curl_init($url);
        curl_setopt_array($ch, [
            CURLOPT_CUSTOMREQUEST => $method,
            CURLOPT_RETURNTRANSFER => true,
            CURLOPT_HEADER => false,
            CURLOPT_HTTPHEADER => $headers,
            CURLOPT_TIMEOUT => $timeoutSeconds,
        ]);
        if ($payload !== '') {
            curl_setopt($ch, CURLOPT_POSTFIELDS, $payload);
        }
        $text = curl_exec($ch);
        $status = (int) curl_getinfo($ch, CURLINFO_RESPONSE_CODE);
        $contentType = (string) curl_getinfo($ch, CURLINFO_CONTENT_TYPE);
        $error = curl_error($ch);
        curl_close($ch);
        if ($text === false) {
            cws_fail(502, 'devlune_unreachable', $error);
        }
    } else {
        $context = stream_context_create([
            'http' => [
                'method' => $method,
                'header' => implode("\r\n", $headers),
                'content' => $payload,
                'timeout' => $timeoutSeconds,
                'ignore_errors' => true,
            ],
        ]);
        $text = file_get_contents($url, false, $context);
        $status = cws_http_status_from_headers($http_response_header ?? []);
        $contentType = cws_content_type_from_headers($http_response_header ?? []);
        if ($text === false) {
            cws_fail(502, 'devlune_unreachable');
        }
    }

    $decoded = json_decode((string) $text, true);
    if (!is_array($decoded)) {
        cws_json([
            'ok' => false,
            'error' => 'devlune_bad_response',
            'status' => $status,
            'contentType' => $contentType ?? '',
            'preview' => cws_response_preview((string) $text),
        ], 502);
    }
    if ($status < 200 || $status >= 300) {
        cws_json([
            'ok' => false,
            'error' => 'devlune_error',
            'status' => $status,
            'devlune' => $decoded,
        ], $status >= 400 && $status < 600 ? $status : 502);
    }
    return $decoded;
}

function cws_http_status_from_headers(array $headers): int
{
    foreach ($headers as $header) {
        if (preg_match('/^HTTP\/\S+\s+(\d+)/', (string) $header, $m)) {
            return (int) $m[1];
        }
    }
    return 0;
}

function cws_content_type_from_headers(array $headers): string
{
    foreach ($headers as $header) {
        if (stripos((string) $header, 'Content-Type:') === 0) {
            return trim(substr((string) $header, strlen('Content-Type:')));
        }
    }
    return '';
}

function cws_response_preview(string $text): string
{
    $text = trim(strip_tags($text));
    $text = preg_replace('/\s+/', ' ', $text) ?: '';
    return substr($text, 0, 240);
}

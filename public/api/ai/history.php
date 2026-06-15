<?php
declare(strict_types=1);

require_once dirname(__DIR__) . '/_bootstrap.php';

cws_require_method('GET');

$user = cws_require_user();
$workbookKey = cws_normalize_workbook_key((string) ($_GET['workbookKey'] ?? ''));
$before = trim((string) ($_GET['before'] ?? ''));
if ($before !== '' && !preg_match('/^\d+$/', $before)) {
    cws_fail(400, 'invalid_before');
}

$size = (int) ($_GET['size'] ?? 1048576);
if ($size < 4096) {
    $size = 4096;
}
if ($size > 5242880) {
    $size = 5242880;
}

$stmt = cws_db()->prepare('SELECT * FROM cws_ai_sessions WHERE user_id = ? AND workbook_key = ? LIMIT 1');
$stmt->execute([(int) $user['id'], $workbookKey]);
$aiSession = $stmt->fetch();

if (!is_array($aiSession) || (string) ($aiSession['devlune_session_id'] ?? '') === '') {
    cws_json([
        'ok' => true,
        'messages' => [],
        'cursor' => null,
        'from' => null,
        'to' => null,
        'total' => 0,
        'exhausted' => true,
        'aiSession' => null,
    ]);
}

$devluneSessionId = (string) $aiSession['devlune_session_id'];
$devluneUserKey = (string) ($aiSession['devlune_user_key'] ?? '');
if ($devluneUserKey === '') {
    $devluneUserKey = cws_devlune_user_key($user, $workbookKey);
}

$query = [
    'email' => $devluneUserKey,
    'size' => (string) $size,
];
if ($before !== '') {
    $query['before'] = $before;
}

$history = cws_devlune_json(
    'GET',
    '/dashboard/api/client/chtcorchat/sessions/' . rawurlencode($devluneSessionId) . '/messages',
    [],
    $query,
    30
);

$from = cws_history_number_or_null($history['from'] ?? null);
$to = cws_history_number_or_null($history['to'] ?? null);

cws_json([
    'ok' => true,
    'messages' => cws_history_messages($history),
    'cursor' => $from,
    'from' => $from,
    'to' => $to,
    'total' => cws_history_number_or_null($history['total'] ?? null),
    'exhausted' => (bool) ($history['exhausted'] ?? false),
    'aiSession' => [
        'id' => (int) $aiSession['id'],
        'workbookKey' => $workbookKey,
        'devluneSessionId' => $devluneSessionId,
    ],
]);

function cws_history_number_or_null($value): ?int
{
    if ($value === null || $value === '') {
        return null;
    }
    if (!is_numeric($value)) {
        return null;
    }
    return max(0, (int) $value);
}

function cws_history_messages(array $payload): array
{
    $items = cws_history_items($payload);
    $messages = [];
    foreach ($items as $index => $item) {
        if (!is_array($item)) {
            continue;
        }
        $message = cws_history_message($item, (int) $index);
        if ($message !== null) {
            $messages[] = $message;
        }
    }
    return $messages;
}

function cws_history_items(array $payload): array
{
    foreach (['messages', 'events', 'items'] as $key) {
        if (isset($payload[$key]) && is_array($payload[$key])) {
            return $payload[$key];
        }
    }
    return cws_history_is_list($payload) ? $payload : [];
}

function cws_history_is_list(array $value): bool
{
    $expected = 0;
    foreach ($value as $key => $_item) {
        if ($key !== $expected) {
            return false;
        }
        $expected += 1;
    }
    return true;
}

function cws_history_message(array $item, int $index): ?array
{
    $message = isset($item['message']) && is_array($item['message']) ? $item['message'] : [];
    $role = (string) ($item['role'] ?? $item['type'] ?? ($message['role'] ?? ''));
    if ($role !== 'user' && $role !== 'assistant') {
        return null;
    }

    $text = cws_history_text($item['text'] ?? null);
    if ($text === '') {
        $text = cws_history_text($item['content'] ?? null);
    }
    if ($text === '') {
        $text = cws_history_text($message['content'] ?? null);
    }
    if ($text === '') {
        return null;
    }
    if ($role === 'user') {
        $text = cws_history_user_request_text($text);
    }

    return [
        'id' => substr(hash('sha256', $role . "\n" . $text . "\n" . $index), 0, 24),
        'role' => $role,
        'text' => $text,
    ];
}

function cws_history_text($value): string
{
    if ($value === null) {
        return '';
    }
    if (is_string($value) || is_numeric($value) || is_bool($value)) {
        return trim((string) $value);
    }
    if (!is_array($value)) {
        return '';
    }
    if (isset($value['text'])) {
        return cws_history_text($value['text']);
    }
    if (isset($value['content'])) {
        return cws_history_text($value['content']);
    }

    $parts = [];
    foreach ($value as $item) {
        $part = cws_history_text($item);
        if ($part !== '') {
            $parts[] = $part;
        }
    }
    return trim(implode("\n", $parts));
}

function cws_history_user_request_text(string $text): string
{
    $marker = "\nUser request:\n";
    $pos = strrpos($text, $marker);
    if ($pos === false) {
        return $text;
    }
    $request = trim(substr($text, $pos + strlen($marker)));
    return $request !== '' ? $request : $text;
}

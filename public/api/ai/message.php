<?php
declare(strict_types=1);

require_once dirname(__DIR__) . '/_bootstrap.php';

cws_require_method('POST');
cws_assert_same_origin_for_write();

$user = cws_require_user();
$body = cws_read_json_body();

$message = cws_json_string($body, 'message', 30000);
$workbookContext = cws_optional_json_string($body, 'workbookContext', 120000);
$workbookKey = cws_normalize_workbook_key(cws_optional_json_string($body, 'workbookKey', 128));
$devluneUserKey = cws_devlune_user_key($user, $workbookKey);
$sessionTitle = 'CWS ' . gmdate('Y-m-d');

$pdo = cws_db();
$stmt = $pdo->prepare('SELECT * FROM cws_ai_sessions WHERE user_id = ? AND workbook_key = ? LIMIT 1');
$stmt->execute([(int) $user['id'], $workbookKey]);
$aiSession = $stmt->fetch();

if (!is_array($aiSession) || (string) ($aiSession['devlune_session_id'] ?? '') === '') {
    $created = cws_devlune_json('POST', '/dashboard/api/client/chtcorchat/sessions', [
        'email' => $devluneUserKey,
        'title' => $sessionTitle,
    ], [], 30);
    $devluneSessionId = (string) ($created['session']['id'] ?? '');
    if ($devluneSessionId === '') {
        cws_fail(502, 'devlune_session_missing');
    }

    if (is_array($aiSession)) {
        $update = $pdo->prepare(
            'UPDATE cws_ai_sessions
                SET devlune_user_key = ?, devlune_session_id = ?, updated_at = UTC_TIMESTAMP()
              WHERE id = ?'
        );
        $update->execute([$devluneUserKey, $devluneSessionId, (int) $aiSession['id']]);
        $aiSession['devlune_user_key'] = $devluneUserKey;
        $aiSession['devlune_session_id'] = $devluneSessionId;
    } else {
        $insert = $pdo->prepare(
            'INSERT INTO cws_ai_sessions
                (user_id, workbook_key, devlune_user_key, devlune_session_id, created_at, updated_at)
             VALUES
                (?, ?, ?, ?, UTC_TIMESTAMP(), UTC_TIMESTAMP())'
        );
        $insert->execute([(int) $user['id'], $workbookKey, $devluneUserKey, $devluneSessionId]);
        $aiSession = [
            'id' => (int) $pdo->lastInsertId(),
            'devlune_user_key' => $devluneUserKey,
            'devlune_session_id' => $devluneSessionId,
        ];
    }
}

$devluneSessionId = (string) $aiSession['devlune_session_id'];
$prompt = cws_build_ai_prompt($message, $workbookContext);
$reply = cws_devlune_json(
    'POST',
    '/dashboard/api/client/chtcorchat/sessions/' . rawurlencode($devluneSessionId) . '/messages',
    [
        'email' => $devluneUserKey,
        'message' => $prompt,
    ],
    [],
    600
);

$touch = $pdo->prepare('UPDATE cws_ai_sessions SET updated_at = UTC_TIMESTAMP() WHERE id = ?');
$touch->execute([(int) $aiSession['id']]);

cws_json([
    'ok' => true,
    'aiSession' => [
        'id' => (int) $aiSession['id'],
        'workbookKey' => $workbookKey,
        'devluneSessionId' => $devluneSessionId,
    ],
    'devlune' => $reply,
]);

function cws_build_ai_prompt(string $message, string $workbookContext): string
{
    if ($workbookContext === '') {
        return $message;
    }

    return implode("\n\n", [
        'You are ChtCortex working inside Cht WebSheet (CWS).',
        'When a workbook edit is needed, return a concise explanation and a JSON object with a whitelisted ops array. Do not return executable JavaScript.',
        'Allowed ops include setCell, setRange, setStyle, insertRows, and deleteRows.',
        'Current CWS workbook context:',
        $workbookContext,
        'User request:',
        $message,
    ]);
}

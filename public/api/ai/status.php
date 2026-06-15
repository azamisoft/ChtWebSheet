<?php
declare(strict_types=1);

require_once dirname(__DIR__) . '/_bootstrap.php';

cws_require_method('GET');

$user = cws_require_user();
$workbookKey = cws_normalize_workbook_key((string) ($_GET['workbookKey'] ?? ''));
$devluneUserKey = cws_devlune_user_key($user, $workbookKey);
$sessionTitle = 'CWS ' . gmdate('Y-m-d');

$pdo = cws_db();
$stmt = $pdo->prepare('SELECT * FROM cws_ai_sessions WHERE user_id = ? AND workbook_key = ? LIMIT 1');
$stmt->execute([(int) $user['id'], $workbookKey]);
$aiSession = $stmt->fetch();

if (is_array($aiSession) && (string) ($aiSession['devlune_session_id'] ?? '') !== '') {
    cws_devlune_json('GET', '/dashboard/api/client/chtcorchat/sessions', [], [
        'email' => (string) ($aiSession['devlune_user_key'] ?: $devluneUserKey),
    ], 6);

    cws_json([
        'ok' => true,
        'status' => 'online',
        'available' => true,
        'message' => '',
        'aiSession' => [
            'id' => (int) $aiSession['id'],
            'workbookKey' => $workbookKey,
            'devluneSessionId' => (string) $aiSession['devlune_session_id'],
        ],
    ]);
}

$created = cws_devlune_json('POST', '/dashboard/api/client/chtcorchat/sessions', [
    'email' => $devluneUserKey,
    'title' => $sessionTitle,
], [], 8);

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
    $aiSession['id'] = (int) $aiSession['id'];
} else {
    $insert = $pdo->prepare(
        'INSERT INTO cws_ai_sessions
            (user_id, workbook_key, devlune_user_key, devlune_session_id, created_at, updated_at)
         VALUES
            (?, ?, ?, ?, UTC_TIMESTAMP(), UTC_TIMESTAMP())'
    );
    $insert->execute([(int) $user['id'], $workbookKey, $devluneUserKey, $devluneSessionId]);
    $aiSession = ['id' => (int) $pdo->lastInsertId()];
}

cws_json([
    'ok' => true,
    'status' => 'online',
    'available' => true,
    'message' => '',
    'created' => true,
    'aiSession' => [
        'id' => (int) $aiSession['id'],
        'workbookKey' => $workbookKey,
        'devluneSessionId' => $devluneSessionId,
    ],
]);

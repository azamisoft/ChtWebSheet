<?php
// Cht WebSheet mail verification settings.
// Paste these constants into the real WordPress wp-config.php, above:
//   /* That's all, stop editing! Happy publishing. */

define('CWS_MAIL_FROM', 'no-reply@chtec.co.jp');
define('CWS_MAIL_FROM_NAME', 'Cht WebSheet');

// Fill these with the SMTP settings shown by onamae.
define('CWS_SMTP_HOST', 'smtp.example.ne.jp');
define('CWS_SMTP_PORT', 587);
define('CWS_SMTP_USER', 'no-reply@chtec.co.jp');
define('CWS_SMTP_PASSWORD', 'CHANGE_ME');
define('CWS_SMTP_SECURE', 'tls');

// Optional: verification code policy.
define('CWS_EMAIL_CODE_TTL_SECONDS', 600);
define('CWS_EMAIL_CODE_MAX_ATTEMPTS', 5);
define('CWS_MAIL_TIMEZONE', 'Asia/Tokyo');

// Optional: show the registration UI in the AI panel.
define('CWS_REGIST_ENABLED', false);

// Optional: allow local HTML files and local dev builds to call the CWS API.
define('CWS_ALLOWED_API_ORIGINS', 'null,http://localhost:5173,http://127.0.0.1:5173,https://chtec.co.jp');

import "./cws-ai-panel.css";

const CWS_AI_API_BASE_PATH = "api/";
const CWS_AI_REMOTE_API_BASE_URL = "https://chtec.co.jp/cws/api/";
const CWS_AI_AUTH_TOKEN_STORAGE_KEY = "cws.ai.authToken.v1";
const CWS_AI_MAX_CONTEXT_CELLS = 80;
const CWS_AI_MAX_CELL_TEXT_LENGTH = 420;
const CWS_AI_MAX_MESSAGES = 200;
const CWS_AI_HISTORY_BATCH_BYTES = 1024 * 1024;
const CWS_AI_AUTH_TIMEOUT_MS = 20000;
const CWS_AI_USER_TIMEOUT_MS = 10000;
const CWS_AI_STATUS_TIMEOUT_MS = 6000;
const CWS_AI_STATUS_CACHE_MS = 30000;
const CWS_AI_CHAT_TIMEOUT_MS = 650000;

let runtime = null;

export function initCwsAiPanel(options) {
  runtime = {
    $: options.$,
    icons: options.icons,
    createIcons: options.createIcons,
    state: options.state,
    helpers: options.helpers,
  };
  ensureCwsAiState();
  hydrateCwsAiAuthToken();
  bindCwsAiPanelEvents();

  if (document.getElementById("cwsAiPanelRoot")) {
    renderCwsAiPanel();
    return;
  }

  const root = document.createElement("div");
  root.id = "cwsAiPanelRoot";
  root.className = "cws-ai-root";
  document.body.appendChild(root);
  renderCwsAiPanel();
  void refreshCwsAiUser({ quiet: true });
}

function defaultCwsAiState() {
  return {
    open: false,
    tab: "chat",
    authMode: "login",
    authStatus: "idle",
    statusMessage: "",
    statusTone: "info",
    authGeneration: 0,
    authStep: "credentials",
    authCode: "",
    authPendingEmail: "",
    authPendingPassword: "",
    authPendingDisplayName: "",
    authCodeExpiresAt: "",
    authToken: "",
    aiStatus: "unknown",
    aiStatusMessage: "",
    aiStatusCheckedAt: 0,
    aiStatusGeneration: 0,
    aiStatusWorkbookKey: "",
    user: null,
    meLoaded: false,
    registEnabled: false,
    authEmail: "",
    authDisplayName: "",
    passwordVisible: false,
    messageDraft: "",
    messageComposing: false,
    messageCompositionEndedAt: 0,
    messages: [],
    historyWorkbookKey: "",
    historyLoaded: false,
    historyLoading: false,
    historyCursor: null,
    historyExhausted: false,
    historyError: "",
    historyGeneration: 0,
    historySeen: {},
    dragOffsetRight: null,
    dragOffsetBottom: null,
    dragSuppressClickUntil: 0,
    panelWidth: null,
    panelHeight: null,
    sending: false,
    settings: {
      workbookScoped: true,
      includeSelection: true,
      includeVisibleRange: true,
    },
  };
}

function ensureCwsAiState() {
  const defaults = defaultCwsAiState();
  if (!runtime.state.cwsAi) {
    runtime.state.cwsAi = defaults;
    return;
  }
  const existing = runtime.state.cwsAi;
  Object.entries(defaults).forEach(([key, value]) => {
    if (key !== "settings" && existing[key] === undefined) {
      existing[key] = value;
    }
  });
  existing.settings = {
    ...defaults.settings,
    ...(existing.settings || {}),
  };
}

function hydrateCwsAiAuthToken() {
  const ai = aiState();
  if (!ai.authToken) {
    ai.authToken = cwsAiReadStoredAuthToken();
  }
}

function bindCwsAiPanelEvents() {
  const $document = runtime.$(document);
  $document.off(".cwsAiPanel");
  runtime.$(window).off(".cwsAiPanel");
  $document.on("click.cwsAiPanel", "[data-cws-ai-toggle]", handleCwsAiToggleClick);
  $document.on("pointerdown.cwsAiPanel", "[data-cws-ai-drag-handle]", handleCwsAiDragPointerDown);
  $document.on("pointerdown.cwsAiPanel", "[data-cws-ai-resize-handle]", handleCwsAiResizePointerDown);
  $document.on("click.cwsAiPanel", "[data-cws-ai-close]", closeCwsAiPanel);
  $document.on("click.cwsAiPanel", "[data-cws-ai-settings-toggle]", handleCwsAiSettingsToggle);
  $document.on("click.cwsAiPanel", "[data-cws-ai-auth-mode]", handleCwsAiAuthModeClick);
  $document.on("click.cwsAiPanel", "[data-cws-ai-code-back]", handleCwsAiCodeBack);
  $document.on("click.cwsAiPanel", "[data-cws-ai-code-resend]", () => {
    void requestCwsAiAuthCode({ resend: true });
  });
  $document.on("click.cwsAiPanel", "[data-cws-ai-password-toggle]", toggleCwsAiPasswordVisibility);
  $document.on("click.cwsAiPanel", "[data-cws-ai-logout]", () => {
    void logoutCwsAiUser();
  });
  $document.on("click.cwsAiPanel", "[data-cws-ai-refresh-user]", () => {
    void refreshCwsAiUser({ quiet: false });
  });
  $document.on("click.cwsAiPanel", "[data-cws-ai-history-retry]", () => {
    void loadCwsAiHistoryBatch({ preserveScroll: true, force: true });
  });
  $document.on("click.cwsAiPanel", "[data-cws-ai-status-retry]", () => {
    void checkCwsAiStatus({ force: true });
  });
  $document.on("input.cwsAiPanel", "[data-cws-ai-auth-field]", handleCwsAiAuthFieldInput);
  $document.on("input.cwsAiPanel", "[data-cws-ai-message]", handleCwsAiMessageInput);
  $document.on("compositionstart.cwsAiPanel", "[data-cws-ai-message]", handleCwsAiMessageCompositionStart);
  $document.on("compositionend.cwsAiPanel", "[data-cws-ai-message]", handleCwsAiMessageCompositionEnd);
  $document.on("keydown.cwsAiPanel", "[data-cws-ai-message]", handleCwsAiMessageKeydown);
  $document.on("change.cwsAiPanel", "[data-cws-ai-setting]", handleCwsAiSettingChange);
  $document.on("submit.cwsAiPanel", "[data-cws-ai-auth-form]", (event) => {
    void handleCwsAiAuthSubmit(event);
  });
  $document.on("submit.cwsAiPanel", "[data-cws-ai-password-change-form]", (event) => {
    void handleCwsAiPasswordChangeSubmit(event);
  });
  $document.on("submit.cwsAiPanel", "[data-cws-ai-chat-form]", (event) => {
    void handleCwsAiChatSubmit(event);
  });
  runtime.$(window).on("resize.cwsAiPanel", () => {
    const root = document.getElementById("cwsAiPanelRoot");
    if (!root) return;
    applyCwsAiPanelSize(root);
    applyCwsAiRootPosition(root);
  });
}

function renderCwsAiPanel(options = {}) {
  const root = document.getElementById("cwsAiPanelRoot");
  if (!root) return;
  const previousMessages = document.querySelector("[data-cws-ai-messages]");
  const scrollMode = options.messageScroll || "bottom";
  const previousScrollHeight = previousMessages?.scrollHeight || 0;
  const previousScrollTop = previousMessages?.scrollTop || 0;
  applyCwsAiPanelSize(root);
  applyCwsAiRootPosition(root);
  root.innerHTML = cwsAiPanelHtml();
  const messages = root.querySelector("[data-cws-ai-messages]");
  if (messages) {
    messages.addEventListener("scroll", handleCwsAiMessagesScroll);
  }
  runtime.createIcons({ icons: runtime.icons });
  requestAnimationFrame(() => {
    applyCwsAiPanelSize(root);
    applyCwsAiRootPosition(root);
    if (scrollMode === "preserve" || scrollMode === "none") {
      preserveCwsAiMessagesScroll(previousScrollHeight, previousScrollTop);
    } else if (scrollMode !== "none") {
      scrollCwsAiMessagesToEnd();
    }
  });
}

function cwsAiPanelHtml() {
  const ai = aiState();
  const toggleLabel = ai.open ? "AI を閉じる" : "AI";
  return `
    <div class="cws-ai-shell${ai.open ? " is-open" : ""}">
      ${ai.open ? cwsAiDialogHtml() : ""}
      <button class="cws-ai-toggle" type="button" data-cws-ai-toggle data-cws-ai-drag-handle aria-expanded="${ai.open ? "true" : "false"}" title="${escapeAttr(toggleLabel)}">
        <i data-lucide="${ai.open ? "message-square-x" : "message-square-text"}"></i>
        <span class="sr-only">${escapeHtml(toggleLabel)}</span>
      </button>
    </div>
  `;
}

function cwsAiDialogHtml() {
  const ai = aiState();
  const isAuthed = Boolean(ai.user);
  const currentView = isAuthed && ai.tab === "settings" ? "settings" : "chat";
  const checking = ai.authStatus === "checking" && isAuthed;
  const busy = ai.authStatus === "loading" || ai.sending || ai.aiStatus === "checking";
  return `
    <section class="cws-ai-panel" role="dialog" aria-modal="false" aria-labelledby="cwsAiTitle">
      <button class="cws-ai-resize-handle" type="button" data-cws-ai-resize-handle title="サイズ変更" aria-label="サイズ変更"></button>
      <header class="cws-ai-header" data-cws-ai-drag-handle>
        <div class="cws-ai-title">
          <span class="cws-ai-title-icon" aria-hidden="true"><i data-lucide="sparkles"></i></span>
          <h2 id="cwsAiTitle">AI</h2>
        </div>
        ${isAuthed ? `<div class="cws-ai-account-label">${escapeHtml(checking ? "確認中" : cwsAiUserLabel(ai.user))}</div>` : ""}
        <div class="cws-ai-header-actions">
          <button class="cws-ai-icon-button" type="button" data-cws-ai-refresh-user title="更新" ${busy ? "disabled" : ""}>
            <i data-lucide="refresh-cw"></i>
            <span class="sr-only">更新</span>
          </button>
          ${isAuthed ? `
            <button class="cws-ai-icon-button${currentView === "settings" ? " is-active" : ""}" type="button" data-cws-ai-settings-toggle title="${currentView === "settings" ? "チャットに戻る" : "設定"}" ${busy ? "disabled" : ""}>
              <i data-lucide="${currentView === "settings" ? "message-square-text" : "settings"}"></i>
              <span class="sr-only">${currentView === "settings" ? "チャットに戻る" : "設定"}</span>
            </button>
            <button class="cws-ai-icon-button" type="button" data-cws-ai-logout title="ログアウト" ${busy ? "disabled" : ""}>
              <i data-lucide="log-out"></i>
              <span class="sr-only">ログアウト</span>
            </button>
          ` : ""}
          <button class="cws-ai-icon-button" type="button" data-cws-ai-close title="閉じる">
            <i data-lucide="message-square-x"></i>
            <span class="sr-only">閉じる</span>
          </button>
        </div>
      </header>
      <div class="cws-ai-body">
        ${isAuthed ? (currentView === "settings" ? cwsAiSettingsHtml() : cwsAiChatHtml()) : cwsAiAuthHtml()}
      </div>
    </section>
  `;
}

function cwsAiAuthHtml() {
  const ai = aiState();
  const registEnabled = ai.registEnabled === true;
  const mode = cwsAiEffectiveAuthMode();
  const busy = ai.authStatus === "loading";
  const passwordType = ai.passwordVisible ? "text" : "password";
  if (ai.authStep === "code") {
    return cwsAiAuthCodeHtml(mode, busy);
  }
  if (mode === "reset") {
    return cwsAiPasswordResetHtml(busy, passwordType);
  }
  return `
    <div class="cws-ai-auth">
      ${registEnabled ? `
        <div class="cws-ai-segmented" role="tablist" aria-label="ログイン">
          <button class="${mode === "login" ? "is-active" : ""}" type="button" data-cws-ai-auth-mode="login" aria-selected="${mode === "login" ? "true" : "false"}">ログイン</button>
          <button class="${mode === "register" ? "is-active" : ""}" type="button" data-cws-ai-auth-mode="register" aria-selected="${mode === "register" ? "true" : "false"}">登録</button>
        </div>
      ` : `
        <div class="cws-ai-code-summary">
          <strong>ログイン</strong>
        </div>
      `}
      <form class="cws-ai-auth-form" data-cws-ai-auth-form data-mode="${escapeAttr(mode)}">
        <label class="cws-ai-field">
          <span>メールアドレス</span>
          <input type="email" autocomplete="email" data-cws-ai-auth-field="email" name="email" value="${escapeAttr(ai.authEmail)}" required ${busy ? "disabled" : ""}>
        </label>
        ${mode === "register" ? `
          <label class="cws-ai-field">
            <span>表示名</span>
            <input type="text" autocomplete="name" data-cws-ai-auth-field="displayName" name="displayName" value="${escapeAttr(ai.authDisplayName)}" maxlength="120" ${busy ? "disabled" : ""}>
          </label>
        ` : ""}
        <label class="cws-ai-field">
          <span>パスワード</span>
          <span class="cws-ai-password-wrap">
            <input type="${passwordType}" autocomplete="${mode === "register" ? "new-password" : "current-password"}" name="password" minlength="8" required ${busy ? "disabled" : ""}>
            <button class="cws-ai-password-toggle" type="button" data-cws-ai-password-toggle title="${ai.passwordVisible ? "隠す" : "表示"}" ${busy ? "disabled" : ""}>
              <i data-lucide="${ai.passwordVisible ? "eye-off" : "eye"}"></i>
              <span class="sr-only">${ai.passwordVisible ? "隠す" : "表示"}</span>
            </button>
          </span>
        </label>
        ${cwsAiNoticeHtml()}
        <button class="cws-ai-primary-button" type="submit" ${busy ? "disabled" : ""}>
          <i data-lucide="unlock-keyhole"></i>
          <span>${mode === "register" ? "登録" : "ログイン"}</span>
        </button>
        ${mode === "login" ? `
          <div class="cws-ai-auth-inline-actions">
            <button type="button" data-cws-ai-auth-mode="reset" ${busy ? "disabled" : ""}>パスワードをお忘れですか</button>
          </div>
        ` : ""}
      </form>
    </div>
  `;
}

function cwsAiComingSoonHtml() {
  return `
    <div class="cws-ai-auth cws-ai-coming-soon">
      <div class="cws-ai-coming-soon-card" role="status" aria-live="polite">
        <span class="cws-ai-coming-soon-icon" aria-hidden="true"><i data-lucide="sparkles"></i></span>
        <strong>AIアシスタントはただいまこちらへ向かっています</strong>
        <span>もう少しだけお待ちください。</span>
      </div>
    </div>
  `;
}

function cwsAiPasswordResetHtml(busy, passwordType) {
  const ai = aiState();
  return `
    <div class="cws-ai-auth">
      <div class="cws-ai-code-summary">
        <strong>パスワード再設定</strong>
        <span>確認コードをメールで送信します。</span>
      </div>
      <form class="cws-ai-auth-form" data-cws-ai-auth-form data-mode="reset">
        <label class="cws-ai-field">
          <span>メールアドレス</span>
          <input type="email" autocomplete="email" data-cws-ai-auth-field="email" name="email" value="${escapeAttr(ai.authEmail)}" required ${busy ? "disabled" : ""}>
        </label>
        <label class="cws-ai-field">
          <span>新しいパスワード</span>
          <span class="cws-ai-password-wrap">
            <input type="${passwordType}" autocomplete="new-password" name="password" minlength="8" required ${busy ? "disabled" : ""}>
            <button class="cws-ai-password-toggle" type="button" data-cws-ai-password-toggle title="${ai.passwordVisible ? "隠す" : "表示"}" ${busy ? "disabled" : ""}>
              <i data-lucide="${ai.passwordVisible ? "eye-off" : "eye"}"></i>
              <span class="sr-only">${ai.passwordVisible ? "隠す" : "表示"}</span>
            </button>
          </span>
        </label>
        ${cwsAiNoticeHtml()}
        <button class="cws-ai-primary-button" type="submit" ${busy ? "disabled" : ""}>
          <i data-lucide="mail-check"></i>
          <span>確認コードを送信</span>
        </button>
        <div class="cws-ai-auth-inline-actions">
          <button type="button" data-cws-ai-auth-mode="login" ${busy ? "disabled" : ""}>ログインに戻る</button>
        </div>
      </form>
    </div>
  `;
}

function cwsAiAuthCodeHtml(mode, busy) {
  const ai = aiState();
  const email = ai.authPendingEmail || ai.authEmail;
  const actionLabels = {
    login: "認証してログイン",
    register: "認証して登録",
    reset: "認証して変更",
  };
  const summaryLabels = {
    login: "ログイン確認コード",
    register: "登録確認コード",
    reset: "パスワード再設定コード",
  };
  const actionLabel = actionLabels[mode] || actionLabels.login;
  return `
    <div class="cws-ai-auth">
      <form class="cws-ai-auth-form" data-cws-ai-auth-form data-mode="${escapeAttr(mode)}">
        <div class="cws-ai-code-summary">
          <strong>${escapeHtml(summaryLabels[mode] || "確認コード")}</strong>
          <span>${escapeHtml(email)} に送信しました。</span>
        </div>
        <label class="cws-ai-field">
          <span>確認コード</span>
          <input type="text" inputmode="numeric" autocomplete="one-time-code" maxlength="6" pattern="[0-9]*" data-cws-ai-auth-field="code" name="code" value="${escapeAttr(ai.authCode)}" required ${busy ? "disabled" : ""}>
        </label>
        ${cwsAiNoticeHtml()}
        <button class="cws-ai-primary-button" type="submit" ${busy ? "disabled" : ""}>
          <i data-lucide="unlock-keyhole"></i>
          <span>${actionLabel}</span>
        </button>
        <div class="cws-ai-auth-inline-actions">
          <button type="button" data-cws-ai-code-resend ${busy ? "disabled" : ""}>コードを再送</button>
          <button type="button" data-cws-ai-code-back ${busy ? "disabled" : ""}>メールを変更</button>
        </div>
      </form>
    </div>
  `;
}

function cwsAiChatHtml() {
  const ai = aiState();
  const disabled = ai.sending || ai.authStatus === "loading" || cwsAiStatusBlocksChat(ai.aiStatus);
  return `
    <div class="cws-ai-chat">
      ${cwsAiNoticeHtml()}
      ${cwsAiAvailabilityHtml()}
      <div class="cws-ai-messages" data-cws-ai-messages>
        ${cwsAiHistoryTopHtml()}
        ${ai.messages.length ? ai.messages.map(cwsAiMessageHtml).join("") : (!ai.historyLoading ? `<div class="cws-ai-empty">メッセージはありません</div>` : "")}
      </div>
      <form class="cws-ai-chat-form" data-cws-ai-chat-form>
        <textarea data-cws-ai-message rows="2" maxlength="30000" placeholder="メッセージ" ${disabled ? "disabled" : ""}>${escapeHtml(ai.messageDraft)}</textarea>
        <button class="cws-ai-send-button" type="submit" title="送信" ${disabled || !ai.messageDraft.trim() ? "disabled" : ""}>
          <i data-lucide="send"></i>
          <span class="sr-only">送信</span>
        </button>
      </form>
    </div>
  `;
}

function cwsAiAvailabilityHtml() {
  const ai = aiState();
  if (!ai.user || ai.aiStatus === "online" || ai.aiStatus === "unknown") return "";
  const checking = ai.aiStatus === "checking";
  return `
    <div class="cws-ai-availability is-${checking ? "checking" : "resting"}">
      <span class="cws-ai-availability-icon" aria-hidden="true"><i data-lucide="${checking ? "refresh-cw" : "clock-3"}"></i></span>
      <span class="cws-ai-availability-copy">
        <strong>${checking ? "AI確認中" : "AI休憩中"}</strong>
        <span>${escapeHtml(ai.aiStatusMessage || (checking ? "ChtCortex の状態を確認しています。" : "ChtCortex を起動してから更新してください。"))}</span>
      </span>
      ${checking ? "" : `
        <button class="cws-ai-availability-retry" type="button" data-cws-ai-status-retry>更新</button>
      `}
    </div>
  `;
}

function cwsAiHistoryTopHtml() {
  const ai = aiState();
  if (ai.historyLoading) {
    return `<div class="cws-ai-history-status is-loading">読み込み中...</div>`;
  }
  if (ai.historyError) {
    return `
      <div class="cws-ai-history-status is-error">
        <span>${escapeHtml(ai.historyError)}</span>
        <button type="button" data-cws-ai-history-retry>再試行</button>
      </div>
    `;
  }
  if (ai.historyExhausted && ai.messages.length) {
    return `<div class="cws-ai-history-status">会話の開始</div>`;
  }
  return "";
}

function cwsAiSettingsHtml() {
  const ai = aiState();
  const settings = ai.settings;
  const busy = ai.authStatus === "loading";
  const passwordType = ai.passwordVisible ? "text" : "password";
  return `
    <div class="cws-ai-settings">
      ${cwsAiNoticeHtml()}
      <div class="cws-ai-settings-section">
        <div class="cws-ai-settings-heading">チャット設定</div>
        ${cwsAiSettingCheckboxHtml("workbookScoped", "ブック別セッション", settings.workbookScoped)}
        ${cwsAiSettingCheckboxHtml("includeSelection", "選択範囲", settings.includeSelection)}
        ${cwsAiSettingCheckboxHtml("includeVisibleRange", "表示範囲", settings.includeVisibleRange)}
      </div>
      <form class="cws-ai-settings-section cws-ai-password-change-form" data-cws-ai-password-change-form>
        <div class="cws-ai-settings-heading">パスワード変更</div>
        <label class="cws-ai-field">
          <span>現在のパスワード</span>
          <span class="cws-ai-password-wrap">
            <input type="${passwordType}" autocomplete="current-password" name="currentPassword" required ${busy ? "disabled" : ""}>
            <button class="cws-ai-password-toggle" type="button" data-cws-ai-password-toggle title="${ai.passwordVisible ? "隠す" : "表示"}" ${busy ? "disabled" : ""}>
              <i data-lucide="${ai.passwordVisible ? "eye-off" : "eye"}"></i>
              <span class="sr-only">${ai.passwordVisible ? "隠す" : "表示"}</span>
            </button>
          </span>
        </label>
        <label class="cws-ai-field">
          <span>新しいパスワード</span>
          <span class="cws-ai-password-wrap">
            <input type="${passwordType}" autocomplete="new-password" name="newPassword" minlength="8" required ${busy ? "disabled" : ""}>
            <button class="cws-ai-password-toggle" type="button" data-cws-ai-password-toggle title="${ai.passwordVisible ? "隠す" : "表示"}" ${busy ? "disabled" : ""}>
              <i data-lucide="${ai.passwordVisible ? "eye-off" : "eye"}"></i>
              <span class="sr-only">${ai.passwordVisible ? "隠す" : "表示"}</span>
            </button>
          </span>
        </label>
        <button class="cws-ai-primary-button" type="submit" ${busy ? "disabled" : ""}>
          <i data-lucide="key-round"></i>
          <span>パスワードを変更</span>
        </button>
      </form>
    </div>
  `;
}

function cwsAiSettingCheckboxHtml(key, label, checked) {
  return `
    <label class="cws-ai-switch">
      <input type="checkbox" data-cws-ai-setting="${escapeAttr(key)}" ${checked ? "checked" : ""}>
      <span class="cws-ai-switch-track" aria-hidden="true"></span>
      <span class="cws-ai-switch-label">${escapeHtml(label)}</span>
    </label>
  `;
}

function cwsAiNoticeHtml() {
  const ai = aiState();
  if (!ai.statusMessage) return "";
  return `<div class="cws-ai-notice is-${escapeAttr(ai.statusTone || "info")}">${escapeHtml(ai.statusMessage)}</div>`;
}

function cwsAiMessageHtml(message) {
  const role = message.role === "user" ? "user" : message.role === "system" ? "system" : "assistant";
  const classNames = [
    "cws-ai-message",
    `is-${role}`,
    message.pending ? "is-pending" : "",
    message.error ? "is-error" : "",
  ].filter(Boolean).join(" ");
  return `
    <div class="${classNames}">
      <div class="cws-ai-message-bubble">
        ${message.pending ? `<span class="cws-ai-thinking" aria-label="送信中"></span>` : cwsAiTextHtml(message.text)}
      </div>
    </div>
  `;
}

function cwsAiTextHtml(value) {
  return escapeHtml(cwsAiReadableMessageText(value)).replace(/\n/g, "<br>");
}

function cwsAiReadableMessageText(value) {
  return String(value || "").replace(/\\n/g, "\n");
}

function handleCwsAiToggleClick(event) {
  const suppressUntil = Number(aiState().dragSuppressClickUntil) || 0;
  if (Date.now() < suppressUntil) {
    event.preventDefault();
    event.stopPropagation();
    return;
  }
  toggleCwsAiPanel();
}

function toggleCwsAiPanel() {
  const ai = aiState();
  ai.open = !ai.open;
  ai.statusMessage = ai.user ? "" : ai.statusMessage;
  renderCwsAiPanel();
  if (ai.open && !ai.meLoaded) {
    void refreshCwsAiUser({ quiet: true });
  } else if (ai.open && ai.user) {
    void ensureCwsAiStatusChecked();
    void ensureCwsAiHistoryLoaded();
  }
}

function handleCwsAiDragPointerDown(event) {
  if (event.button !== undefined && event.button !== 0) return;
  const handle = event.currentTarget;
  if (!handle.matches("[data-cws-ai-toggle]") && event.target.closest("button, input, textarea, select, a")) {
    return;
  }
  const root = document.getElementById("cwsAiPanelRoot");
  if (!root) return;

  const rect = root.getBoundingClientRect();
  const ai = aiState();
  const viewportWidth = window.innerWidth || document.documentElement.clientWidth;
  const viewportHeight = window.innerHeight || document.documentElement.clientHeight;
  const drag = {
    pointerId: event.pointerId,
    startX: event.clientX,
    startY: event.clientY,
    startLeft: rect.left,
    startTop: rect.top,
    width: rect.width,
    height: rect.height,
    moved: false,
  };

  const onMove = (moveEvent) => {
    if (drag.pointerId !== undefined && moveEvent.pointerId !== drag.pointerId) return;
    const deltaX = moveEvent.clientX - drag.startX;
    const deltaY = moveEvent.clientY - drag.startY;
    if (!drag.moved && Math.hypot(deltaX, deltaY) < 4) return;
    drag.moved = true;
    moveEvent.preventDefault();

    const margin = 8;
    const maxLeft = Math.max(margin, viewportWidth - drag.width - margin);
    const maxTop = Math.max(margin, viewportHeight - drag.height - margin);
    const nextLeft = Math.min(Math.max(margin, drag.startLeft + deltaX), maxLeft);
    const nextTop = Math.min(Math.max(margin, drag.startTop + deltaY), maxTop);
    ai.dragOffsetRight = Math.max(margin, viewportWidth - nextLeft - drag.width);
    ai.dragOffsetBottom = Math.max(margin, viewportHeight - nextTop - drag.height);
    applyCwsAiRootPosition(root);
  };

  const onUp = (upEvent) => {
    if (drag.pointerId !== undefined && upEvent.pointerId !== drag.pointerId) return;
    document.removeEventListener("pointermove", onMove);
    document.removeEventListener("pointerup", onUp);
    document.removeEventListener("pointercancel", onUp);
    root.classList.remove("is-dragging");
    if (drag.moved) {
      ai.dragSuppressClickUntil = Date.now() + 300;
      upEvent.preventDefault();
    }
  };

  root.classList.add("is-dragging");
  document.addEventListener("pointermove", onMove, { passive: false });
  document.addEventListener("pointerup", onUp);
  document.addEventListener("pointercancel", onUp);
}

function applyCwsAiRootPosition(root) {
  const ai = aiState();
  if (!root) return;
  if (ai.dragOffsetRight === null || ai.dragOffsetRight === undefined || ai.dragOffsetBottom === null || ai.dragOffsetBottom === undefined) {
    root.style.right = "";
    root.style.bottom = "";
    return;
  }

  const margin = 8;
  const rect = root.getBoundingClientRect();
  const viewportWidth = window.innerWidth || document.documentElement.clientWidth;
  const viewportHeight = window.innerHeight || document.documentElement.clientHeight;
  const maxRight = Math.max(margin, viewportWidth - rect.width - margin);
  const maxBottom = Math.max(margin, viewportHeight - rect.height - margin);
  ai.dragOffsetRight = Math.min(Math.max(margin, Number(ai.dragOffsetRight) || margin), maxRight);
  ai.dragOffsetBottom = Math.min(Math.max(margin, Number(ai.dragOffsetBottom) || margin), maxBottom);
  root.style.right = `${ai.dragOffsetRight}px`;
  root.style.bottom = `${ai.dragOffsetBottom}px`;
}

function applyCwsAiPanelSize(root) {
  const ai = aiState();
  if (!root) return;
  const limits = cwsAiPanelSizeLimits();
  const width = clampNumber(Number(ai.panelWidth) || 392, limits.minWidth, limits.maxWidth);
  const height = clampNumber(Number(ai.panelHeight) || 620, limits.minHeight, limits.maxHeight);
  if (ai.panelWidth !== null && ai.panelWidth !== undefined) {
    ai.panelWidth = width;
  }
  if (ai.panelHeight !== null && ai.panelHeight !== undefined) {
    ai.panelHeight = height;
  }
  root.style.setProperty("--cws-ai-panel-width", `${width}px`);
  root.style.setProperty("--cws-ai-panel-height", `${height}px`);
}

function cwsAiPanelSizeLimits() {
  const viewportWidth = window.innerWidth || document.documentElement.clientWidth;
  const viewportHeight = window.innerHeight || document.documentElement.clientHeight;
  return {
    minWidth: Math.min(320, Math.max(280, viewportWidth - 20)),
    minHeight: Math.min(380, Math.max(320, viewportHeight - 78)),
    maxWidth: Math.max(280, viewportWidth - 20),
    maxHeight: Math.max(320, viewportHeight - 78),
  };
}

function handleCwsAiResizePointerDown(event) {
  if (event.button !== undefined && event.button !== 0) return;
  event.preventDefault();
  event.stopPropagation();

  const root = document.getElementById("cwsAiPanelRoot");
  const panel = event.currentTarget.closest(".cws-ai-panel");
  if (!root || !panel) return;

  const rect = panel.getBoundingClientRect();
  const ai = aiState();
  const resize = {
    pointerId: event.pointerId,
    startX: event.clientX,
    startY: event.clientY,
    startWidth: rect.width,
    startHeight: rect.height,
  };

  const onMove = (moveEvent) => {
    if (resize.pointerId !== undefined && moveEvent.pointerId !== resize.pointerId) return;
    moveEvent.preventDefault();
    const limits = cwsAiPanelSizeLimits();
    ai.panelWidth = clampNumber(resize.startWidth - (moveEvent.clientX - resize.startX), limits.minWidth, limits.maxWidth);
    ai.panelHeight = clampNumber(resize.startHeight - (moveEvent.clientY - resize.startY), limits.minHeight, limits.maxHeight);
    applyCwsAiPanelSize(root);
    applyCwsAiRootPosition(root);
  };

  const onUp = (upEvent) => {
    if (resize.pointerId !== undefined && upEvent.pointerId !== resize.pointerId) return;
    document.removeEventListener("pointermove", onMove);
    document.removeEventListener("pointerup", onUp);
    document.removeEventListener("pointercancel", onUp);
    root.classList.remove("is-resizing");
    upEvent.preventDefault();
  };

  root.classList.add("is-resizing");
  document.addEventListener("pointermove", onMove, { passive: false });
  document.addEventListener("pointerup", onUp);
  document.addEventListener("pointercancel", onUp);
}

function closeCwsAiPanel() {
  aiState().open = false;
  renderCwsAiPanel();
}

function cwsAiNormalizeAuthMode(mode) {
  return mode === "register" || mode === "reset" ? mode : "login";
}

function cwsAiAuthMode() {
  return cwsAiNormalizeAuthMode(aiState().authMode);
}

function cwsAiEffectiveAuthMode() {
  const mode = cwsAiAuthMode();
  if (mode === "register" && aiState().registEnabled !== true) {
    return "login";
  }
  return mode;
}

function cwsAiAuthEndpoint(mode) {
  if (mode === "register") return "register.php";
  if (mode === "reset") return "password-reset.php";
  return "login.php";
}

function cwsAiAuthRequestBody(mode, values) {
  const body = {
    email: values.email,
    password: values.password,
  };
  if (mode === "register") {
    body.displayName = values.displayName;
  }
  if (mode === "reset") {
    body.newPassword = values.password;
    delete body.password;
  }
  if (values.code) {
    body.code = values.code;
  }
  return body;
}

function handleCwsAiSettingsToggle() {
  const ai = aiState();
  if (!ai.user) return;
  ai.tab = ai.tab === "settings" ? "chat" : "settings";
  ai.statusMessage = "";
  renderCwsAiPanel();
  if (ai.tab === "chat") {
    void ensureCwsAiStatusChecked();
    void ensureCwsAiHistoryLoaded();
  }
}

function handleCwsAiAuthModeClick(event) {
  const ai = aiState();
  const mode = event.currentTarget.dataset.cwsAiAuthMode;
  const nextMode = cwsAiNormalizeAuthMode(mode);
  ai.authMode = nextMode === "register" && ai.registEnabled !== true ? "login" : nextMode;
  ai.statusMessage = "";
  ai.passwordVisible = false;
  resetCwsAiAuthVerificationState({ clearPending: true });
  renderCwsAiPanel();
}

function handleCwsAiCodeBack() {
  const ai = aiState();
  ai.authEmail = ai.authPendingEmail || ai.authEmail;
  ai.authDisplayName = ai.authPendingDisplayName || ai.authDisplayName;
  ai.statusMessage = "";
  resetCwsAiAuthVerificationState({ clearPending: true });
  renderCwsAiPanel();
}

function handleCwsAiAuthFieldInput(event) {
  const ai = aiState();
  const field = event.currentTarget.dataset.cwsAiAuthField;
  if (field === "email") {
    ai.authEmail = event.currentTarget.value;
  } else if (field === "displayName") {
    ai.authDisplayName = event.currentTarget.value;
  } else if (field === "code") {
    ai.authCode = event.currentTarget.value.replace(/\D+/g, "").slice(0, 6);
    if (event.currentTarget.value !== ai.authCode) {
      event.currentTarget.value = ai.authCode;
    }
  }
}

function handleCwsAiMessageInput(event) {
  const ai = aiState();
  ai.messageDraft = event.currentTarget.value;
  const form = event.currentTarget.form;
  const button = form?.querySelector(".cws-ai-send-button");
  if (button) {
    button.disabled = ai.sending || cwsAiStatusBlocksChat(ai.aiStatus) || !ai.messageDraft.trim();
  }
}

function handleCwsAiMessageCompositionStart() {
  aiState().messageComposing = true;
}

function handleCwsAiMessageCompositionEnd() {
  const ai = aiState();
  ai.messageComposing = false;
  ai.messageCompositionEndedAt = Date.now();
}

function handleCwsAiMessageKeydown(event) {
  if (event.key !== "Enter" || event.shiftKey || cwsAiIsMessageComposing(event)) return;
  event.preventDefault();
  event.currentTarget.form?.requestSubmit?.();
}

function cwsAiIsMessageComposing(event) {
  const nativeEvent = event.originalEvent || event;
  const keyCode = Number(event.keyCode || nativeEvent?.keyCode || 0);
  const endedAt = Number(aiState().messageCompositionEndedAt) || 0;
  return Boolean(
    aiState().messageComposing ||
      event.isComposing ||
      nativeEvent?.isComposing ||
      keyCode === 229 ||
      (event.key === "Enter" && endedAt && Date.now() - endedAt < 150)
  );
}

function handleCwsAiSettingChange(event) {
  const ai = aiState();
  const key = event.currentTarget.dataset.cwsAiSetting;
  if (!Object.prototype.hasOwnProperty.call(ai.settings, key)) return;
  ai.settings[key] = event.currentTarget.checked;
  if (key === "workbookScoped") {
    resetCwsAiHistoryState({ clearMessages: true });
    resetCwsAiAvailabilityState();
    renderCwsAiPanel();
    void ensureCwsAiStatusChecked();
    void ensureCwsAiHistoryLoaded();
  }
}

function toggleCwsAiPasswordVisibility() {
  const ai = aiState();
  ai.passwordVisible = !ai.passwordVisible;
  renderCwsAiPanel();
}

async function handleCwsAiAuthSubmit(event) {
  event.preventDefault();
  if (aiState().authStep === "code") {
    await completeCwsAiAuthWithCode(event.currentTarget);
    return;
  }
  await requestCwsAiAuthCode({ form: event.currentTarget });
}

async function requestCwsAiAuthCode(options = {}) {
  const ai = aiState();
  const authGeneration = bumpCwsAiAuthGeneration();
  const form = options.form || null;
  const mode = cwsAiEffectiveAuthMode();
  const formData = form ? new FormData(form) : null;
  const email = String(formData?.get("email") || ai.authPendingEmail || ai.authEmail || "").trim();
  const password = String(formData?.get("password") || ai.authPendingPassword || "");
  const displayName = String(formData?.get("displayName") || ai.authPendingDisplayName || ai.authDisplayName || "").trim();
  ai.authEmail = email;
  ai.authDisplayName = displayName;
  if (!email || !password) {
    ai.statusTone = "error";
    ai.statusMessage = mode === "reset" ? "メールアドレスと新しいパスワードを入力してください。" : "メールアドレスとパスワードを入力してください。";
    renderCwsAiPanel();
    return;
  }
  ai.authStatus = "loading";
  ai.statusMessage = "";
  renderCwsAiPanel();
  try {
    const data = await cwsAiRequest(cwsAiAuthEndpoint(mode), {
      method: "POST",
      body: cwsAiAuthRequestBody(mode, { email, password, displayName }),
      timeoutMs: CWS_AI_AUTH_TIMEOUT_MS,
    });
    if (!cwsAiAuthGenerationIsCurrent(authGeneration)) return;
    if (data.verificationRequired) {
      ai.authStep = "code";
      ai.authCode = "";
      ai.authPendingEmail = email;
      ai.authPendingPassword = password;
      ai.authPendingDisplayName = displayName;
      ai.authCodeExpiresAt = data.expiresAt || "";
      ai.passwordVisible = false;
      ai.statusTone = "success";
      ai.statusMessage = options.resend ? "確認コードを再送しました。" : "確認コードを送信しました。";
      return;
    }
    cwsAiCaptureAuthToken(data);
    const user = data.user || (await cwsAiFetchCurrentUserAfterAuth(authGeneration));
    if (!cwsAiAuthGenerationIsCurrent(authGeneration)) return;
    finishCwsAiAuth(user, mode);
  } catch (error) {
    if (!cwsAiAuthGenerationIsCurrent(authGeneration)) return;
    ai.statusTone = "error";
    ai.statusMessage = error.message || "確認コードを送信できませんでした。";
  } finally {
    if (!cwsAiAuthGenerationIsCurrent(authGeneration)) return;
    ai.authStatus = "idle";
    renderCwsAiPanel();
    if (ai.open && ai.user) {
      void ensureCwsAiStatusChecked();
      void ensureCwsAiHistoryLoaded();
    }
  }
}

async function completeCwsAiAuthWithCode(form) {
  const ai = aiState();
  const authGeneration = bumpCwsAiAuthGeneration();
  const mode = cwsAiEffectiveAuthMode();
  const formData = new FormData(form);
  const code = String(formData.get("code") || ai.authCode || "").trim();
  const email = ai.authPendingEmail || ai.authEmail;
  const password = ai.authPendingPassword;
  const displayName = ai.authPendingDisplayName || ai.authDisplayName;
  ai.authCode = code.replace(/\D+/g, "").slice(0, 6);
  if (!email || !password || !ai.authCode) {
    ai.statusTone = "error";
    ai.statusMessage = "確認コードを入力してください。";
    renderCwsAiPanel();
    return;
  }

  ai.authStatus = "loading";
  ai.statusMessage = "";
  renderCwsAiPanel();
  try {
    const data = await cwsAiRequest(cwsAiAuthEndpoint(mode), {
      method: "POST",
      body: cwsAiAuthRequestBody(mode, { email, password, displayName, code: ai.authCode }),
      timeoutMs: CWS_AI_AUTH_TIMEOUT_MS,
    });
    if (!cwsAiAuthGenerationIsCurrent(authGeneration)) return;
    cwsAiCaptureAuthToken(data);
    const user = data.user || (await cwsAiFetchCurrentUserAfterAuth(authGeneration));
    if (!cwsAiAuthGenerationIsCurrent(authGeneration)) return;
    finishCwsAiAuth(user, mode);
  } catch (error) {
    if (!cwsAiAuthGenerationIsCurrent(authGeneration)) return;
    ai.statusTone = "error";
    ai.statusMessage = error.message || "確認コードを確認できませんでした。";
  } finally {
    if (!cwsAiAuthGenerationIsCurrent(authGeneration)) return;
    ai.authStatus = "idle";
    renderCwsAiPanel();
    if (ai.open && ai.user) {
      void ensureCwsAiStatusChecked();
      void ensureCwsAiHistoryLoaded();
    }
  }
}

function finishCwsAiAuth(user, mode) {
  if (!user) {
    throw new Error(mode === "register" ? "登録は完了しましたが、ログイン状態を確認できませんでした。ページを再読み込みしてください。" : "ログイン状態を確認できませんでした。");
  }
  const ai = aiState();
  ai.user = user;
  ai.meLoaded = true;
  ai.tab = "chat";
  ai.passwordVisible = false;
  resetCwsAiAuthVerificationState({ clearPending: true });
  resetCwsAiHistoryState({ clearMessages: true });
  resetCwsAiAvailabilityState();
  ai.statusTone = "success";
  ai.statusMessage = mode === "register" ? "登録してログインしました。" : mode === "reset" ? "パスワードを変更してログインしました。" : "ログインしました。";
  void ensureCwsAiStatusChecked();
  void ensureCwsAiHistoryLoaded();
}

async function handleCwsAiPasswordChangeSubmit(event) {
  event.preventDefault();
  const ai = aiState();
  if (!ai.user) {
    ai.statusTone = "error";
    ai.statusMessage = "ログインしてください。";
    renderCwsAiPanel();
    return;
  }

  const formData = new FormData(event.currentTarget);
  const currentPassword = String(formData.get("currentPassword") || "");
  const newPassword = String(formData.get("newPassword") || "");
  if (!currentPassword || !newPassword) {
    ai.statusTone = "error";
    ai.statusMessage = "現在のパスワードと新しいパスワードを入力してください。";
    renderCwsAiPanel();
    return;
  }

  ai.authStatus = "loading";
  ai.statusMessage = "";
  renderCwsAiPanel();
  try {
    const data = await cwsAiRequest("change-password.php", {
      method: "POST",
      body: { currentPassword, newPassword },
      timeoutMs: CWS_AI_AUTH_TIMEOUT_MS,
    });
    ai.user = data.user || ai.user;
    ai.statusTone = "success";
    ai.statusMessage = "パスワードを変更しました。";
    ai.passwordVisible = false;
  } catch (error) {
    ai.statusTone = "error";
    ai.statusMessage = error.message || "パスワードを変更できませんでした。";
  } finally {
    ai.authStatus = "idle";
    renderCwsAiPanel();
  }
}

async function logoutCwsAiUser() {
  const ai = aiState();
  const authGeneration = bumpCwsAiAuthGeneration();
  ai.user = null;
  ai.messages = [];
  ai.messageDraft = "";
  resetCwsAiAuthVerificationState({ clearPending: true });
  resetCwsAiHistoryState({ clearMessages: true });
  resetCwsAiAvailabilityState();
  ai.tab = "chat";
  ai.authStatus = "idle";
  ai.statusTone = "info";
  ai.statusMessage = "";
  renderCwsAiPanel();
  try {
    await cwsAiRequest("logout.php", { method: "POST", body: {}, timeoutMs: CWS_AI_USER_TIMEOUT_MS });
  } catch (error) {
    if (!cwsAiAuthGenerationIsCurrent(authGeneration)) return;
    ai.statusTone = "error";
    ai.statusMessage = error.message || "ログアウトできませんでした。";
    renderCwsAiPanel();
  } finally {
    cwsAiClearAuthToken();
  }
}

async function cwsAiFetchCurrentUserAfterAuth(authGeneration) {
  try {
    const data = await cwsAiRequest("me.php", { timeoutMs: CWS_AI_USER_TIMEOUT_MS });
    if (!cwsAiAuthGenerationIsCurrent(authGeneration)) return null;
    aiState().registEnabled = data.registEnabled === true;
    return data.user || null;
  } catch {
    return null;
  }
}

async function handleCwsAiChatSubmit(event) {
  event.preventDefault();
  const ai = aiState();
  if (!ai.user) {
    ai.statusTone = "error";
    ai.statusMessage = "ログインしてください。";
    renderCwsAiPanel();
    return;
  }
  if (ai.sending) return;
  const messageText = ai.messageDraft.trim();
  if (!messageText) return;
  if (ai.aiStatus !== "online") {
    const available = await checkCwsAiStatus({ force: true });
    if (!available) {
      ai.statusTone = "error";
      ai.statusMessage = "AIは休憩中です。";
      renderCwsAiPanel({ messageScroll: "none" });
      return;
    }
  }

  const pendingMessage = { role: "assistant", text: "", pending: true };
  ai.messages.push({ role: "user", text: messageText }, pendingMessage);
  cwsAiPruneMessages();
  ai.messageDraft = "";
  ai.sending = true;
  ai.statusMessage = "";
  renderCwsAiPanel();

  try {
    const data = await cwsAiRequest("ai/message.php", {
      method: "POST",
      timeoutMs: CWS_AI_CHAT_TIMEOUT_MS,
      body: {
        message: messageText,
        workbookKey: cwsAiWorkbookKey(),
        workbookContext: cwsAiWorkbookContextText(),
      },
    });
    pendingMessage.pending = false;
    pendingMessage.text = cwsAiExtractReply(data);
  } catch (error) {
    pendingMessage.pending = false;
    pendingMessage.error = true;
    if (cwsAiIsUnavailableError(error)) {
      ai.aiStatus = "resting";
      ai.aiStatusMessage = "ChtCortex を起動してから更新してください。";
      ai.aiStatusCheckedAt = Date.now();
    }
    pendingMessage.text = error.message || "送信できませんでした。";
  } finally {
    ai.sending = false;
    renderCwsAiPanel();
  }
}

async function refreshCwsAiUser(options = {}) {
  const ai = aiState();
  const authGeneration = ai.authGeneration;
  const quiet = options.quiet !== false;
  ai.authStatus = "checking";
  if (!quiet || ai.open) renderCwsAiPanel();
  try {
    const previousUserId = ai.user?.id ?? null;
    const data = await cwsAiRequest("me.php", { timeoutMs: CWS_AI_USER_TIMEOUT_MS });
    if (!cwsAiAuthGenerationIsCurrent(authGeneration)) return;
    const nextUser = data.user || null;
    ai.registEnabled = data.registEnabled === true;
    ai.user = nextUser;
    ai.meLoaded = true;
    if (!nextUser || (previousUserId !== null && nextUser.id !== previousUserId)) {
      resetCwsAiHistoryState({ clearMessages: true });
      resetCwsAiAvailabilityState();
    }
    if (!nextUser) {
      cwsAiClearAuthToken();
    }
    if (!quiet) {
      ai.statusTone = data.user ? "success" : "info";
      ai.statusMessage = "";
    }
  } catch (error) {
    if (!cwsAiAuthGenerationIsCurrent(authGeneration)) return;
    ai.user = null;
    ai.meLoaded = true;
    cwsAiClearAuthToken();
    resetCwsAiHistoryState({ clearMessages: true });
    resetCwsAiAvailabilityState();
    if (!quiet) {
      ai.statusTone = "error";
      ai.statusMessage = error.message || "ログイン状態を確認できませんでした。";
    }
  } finally {
    if (!cwsAiAuthGenerationIsCurrent(authGeneration)) return;
    ai.authStatus = "idle";
    renderCwsAiPanel();
    if (ai.open && ai.user) {
      void ensureCwsAiStatusChecked();
      void ensureCwsAiHistoryLoaded();
    }
  }
}

function resetCwsAiAvailabilityState(options = {}) {
  const ai = aiState();
  ai.aiStatus = "unknown";
  ai.aiStatusMessage = "";
  ai.aiStatusCheckedAt = 0;
  ai.aiStatusWorkbookKey = options.workbookKey || "";
  ai.aiStatusGeneration = (Number(ai.aiStatusGeneration) || 0) + 1;
}

function resetCwsAiAuthVerificationState(options = {}) {
  const ai = aiState();
  ai.authStep = "credentials";
  ai.authCode = "";
  ai.authCodeExpiresAt = "";
  if (options.clearPending) {
    ai.authPendingEmail = "";
    ai.authPendingPassword = "";
    ai.authPendingDisplayName = "";
  }
}

function cwsAiCaptureAuthToken(data) {
  const token = String(data?.authToken || data?.session?.token || "").trim();
  if (!/^[A-Za-z0-9]{64}$/.test(token)) return;
  const ai = aiState();
  ai.authToken = token;
  cwsAiWriteStoredAuthToken(token);
}

function cwsAiClearAuthToken() {
  const ai = aiState();
  ai.authToken = "";
  cwsAiRemoveStoredAuthToken();
}

function cwsAiReadStoredAuthToken() {
  try {
    return String(window.localStorage?.getItem(CWS_AI_AUTH_TOKEN_STORAGE_KEY) || "").trim();
  } catch {
    return "";
  }
}

function cwsAiWriteStoredAuthToken(token) {
  try {
    window.localStorage?.setItem(CWS_AI_AUTH_TOKEN_STORAGE_KEY, token);
  } catch {
    // Local file storage can be unavailable in some browser settings.
  }
}

function cwsAiRemoveStoredAuthToken() {
  try {
    window.localStorage?.removeItem(CWS_AI_AUTH_TOKEN_STORAGE_KEY);
  } catch {
    // Local file storage can be unavailable in some browser settings.
  }
}

function resetCwsAiHistoryState(options = {}) {
  const ai = aiState();
  ai.historyWorkbookKey = "";
  ai.historyLoaded = false;
  ai.historyLoading = false;
  ai.historyCursor = null;
  ai.historyExhausted = false;
  ai.historyError = "";
  ai.historyGeneration = (Number(ai.historyGeneration) || 0) + 1;
  ai.historySeen = {};
  if (options.clearMessages) {
    ai.messages = [];
  }
}

function ensureCwsAiHistoryStateForWorkbook() {
  const ai = aiState();
  const workbookKey = cwsAiWorkbookKey();
  if (ai.historyWorkbookKey === workbookKey) return;
  resetCwsAiHistoryState({ clearMessages: true });
  ai.historyWorkbookKey = workbookKey;
}

async function ensureCwsAiHistoryLoaded() {
  const ai = aiState();
  if (!ai.user || ai.tab === "settings") return;
  ensureCwsAiHistoryStateForWorkbook();
  if (ai.historyLoaded || ai.historyLoading || ai.historyExhausted) return;
  await loadCwsAiHistoryBatch({ preserveScroll: false });
}

function ensureCwsAiAvailabilityStateForWorkbook() {
  const ai = aiState();
  const workbookKey = cwsAiWorkbookKey();
  if (ai.aiStatusWorkbookKey === workbookKey) return;
  resetCwsAiAvailabilityState({ workbookKey });
}

async function ensureCwsAiStatusChecked(options = {}) {
  const ai = aiState();
  if (!ai.user || ai.tab === "settings") return false;
  ensureCwsAiAvailabilityStateForWorkbook();
  if (ai.aiStatus === "checking") return false;
  const checkedAt = Number(ai.aiStatusCheckedAt) || 0;
  if (!options.force && ai.aiStatus !== "unknown" && checkedAt && Date.now() - checkedAt < CWS_AI_STATUS_CACHE_MS) {
    return ai.aiStatus === "online";
  }
  return checkCwsAiStatus(options);
}

async function checkCwsAiStatus(options = {}) {
  const ai = aiState();
  if (!ai.user) return false;
  ensureCwsAiAvailabilityStateForWorkbook();

  const generation = (Number(ai.aiStatusGeneration) || 0) + 1;
  ai.aiStatusGeneration = generation;
  ai.aiStatus = "checking";
  ai.aiStatusMessage = "ChtCortex の状態を確認しています。";
  if (ai.open) {
    renderCwsAiPanel({ messageScroll: options.preserveScroll ? "preserve" : "none" });
  }

  const query = new URLSearchParams({
    workbookKey: ai.aiStatusWorkbookKey || cwsAiWorkbookKey(),
  });

  try {
    const data = await cwsAiRequest(`ai/status.php?${query.toString()}`, {
      timeoutMs: CWS_AI_STATUS_TIMEOUT_MS,
    });
    if (ai.aiStatusGeneration !== generation) return false;
    const online = data.status === "online" || data.available === true;
    ai.aiStatus = online ? "online" : "resting";
    ai.aiStatusMessage = online ? "" : (data.message || "ChtCortex を起動してから更新してください。");
    ai.aiStatusCheckedAt = Date.now();
    return online;
  } catch (error) {
    if (ai.aiStatusGeneration !== generation) return false;
    ai.aiStatus = "resting";
    ai.aiStatusMessage = cwsAiIsUnavailableError(error)
      ? "ChtCortex を起動してから更新してください。"
      : (error.message || "ChtCortex の状態を確認できませんでした。");
    ai.aiStatusCheckedAt = Date.now();
    return false;
  } finally {
    if (ai.aiStatusGeneration !== generation) return;
    if (ai.open) {
      renderCwsAiPanel({ messageScroll: options.preserveScroll ? "preserve" : "none" });
    }
  }
}

function handleCwsAiMessagesScroll(event) {
  const target = event.currentTarget;
  if (!target || target.scrollTop > 28) return;
  void loadCwsAiHistoryBatch({ preserveScroll: true });
}

async function loadCwsAiHistoryBatch(options = {}) {
  const ai = aiState();
  if (!ai.user || ai.historyLoading) return;
  ensureCwsAiHistoryStateForWorkbook();
  if (ai.historyExhausted && !options.force) return;

  const generation = (Number(ai.historyGeneration) || 0) + 1;
  ai.historyGeneration = generation;
  ai.historyLoading = true;
  ai.historyError = "";
  renderCwsAiPanel({ messageScroll: options.preserveScroll ? "preserve" : "none" });

  const before = options.force ? null : ai.historyCursor;
  const query = new URLSearchParams({
    workbookKey: ai.historyWorkbookKey || cwsAiWorkbookKey(),
    size: String(CWS_AI_HISTORY_BATCH_BYTES),
  });
  if (before !== null && before !== undefined && before !== "") {
    query.set("before", String(before));
  }

  try {
    const data = await cwsAiRequest(`ai/history.php?${query.toString()}`, {
      timeoutMs: CWS_AI_USER_TIMEOUT_MS,
    });
    if (ai.historyGeneration !== generation) return;
    const messages = cwsAiNormalizeHistoryMessages(data.messages);
    const cursor = data.cursor ?? data.from ?? null;
    const cursorStalled = before !== null && before !== undefined && cursor !== null && String(cursor) === String(before);
    ai.historyCursor = cursor;
    ai.historyLoaded = true;
    ai.historyExhausted = Boolean(data.exhausted || cursorStalled || cursor === 0 || cursor === "0");
    prependCwsAiHistoryMessages(messages);
  } catch (error) {
    if (ai.historyGeneration !== generation) return;
    ai.historyError = error.message || "履歴を読み込めませんでした。";
  } finally {
    if (ai.historyGeneration !== generation) return;
    ai.historyLoading = false;
    renderCwsAiPanel({ messageScroll: options.preserveScroll ? "preserve" : "bottom" });
  }
}

function cwsAiNormalizeHistoryMessages(messages) {
  if (!Array.isArray(messages)) return [];
  return messages
    .map((message) => ({
      id: String(message?.id || ""),
      role: message?.role === "user" ? "user" : message?.role === "system" ? "system" : "assistant",
      text: String(message?.text || "").trim(),
    }))
    .filter((message) => message.text);
}

function prependCwsAiHistoryMessages(messages) {
  const ai = aiState();
  if (!messages.length) return;
  const seen = ai.historySeen && typeof ai.historySeen === "object" ? ai.historySeen : {};
  ai.messages.forEach((message) => {
    seen[cwsAiHistoryMessageKey(message)] = true;
  });
  const fresh = [];
  messages.forEach((message) => {
    const key = cwsAiHistoryMessageKey(message);
    if (seen[key]) return;
    seen[key] = true;
    fresh.push(message);
  });
  ai.historySeen = seen;
  if (fresh.length) {
    ai.messages = [...fresh, ...ai.messages];
  }
}

function cwsAiHistoryMessageKey(message) {
  return `${message.role || ""}\n${message.text || ""}`;
}

async function cwsAiRequest(path, options = {}) {
  const body = options.body == null ? undefined : JSON.stringify(options.body);
  const timeoutMs = Number(options.timeoutMs) > 0 ? Number(options.timeoutMs) : CWS_AI_AUTH_TIMEOUT_MS;
  const controller = typeof AbortController === "function" ? new AbortController() : null;
  const timeoutId = controller
    ? window.setTimeout(() => {
        controller.abort();
      }, timeoutMs)
    : 0;
  let response;
  const url = cwsAiApiUrl(path);
  const authToken = cwsAiRequestAuthToken();
  try {
    response = await fetch(url, {
      method: options.method || "GET",
      credentials: cwsAiShouldUseSameOriginCredentials(url) ? "same-origin" : "omit",
      cache: "no-store",
      signal: controller?.signal,
      headers: {
        Accept: "application/json",
        ...(authToken ? { Authorization: `Bearer ${authToken}` } : {}),
        ...(body ? { "Content-Type": "application/json" } : {}),
        ...(options.headers || {}),
      },
      body,
    });
  } catch (error) {
    const timedOut = error?.name === "AbortError";
    throw new Error(timedOut ? "通信がタイムアウトしました。もう一度お試しください。" : "通信できませんでした。");
  } finally {
    if (timeoutId) window.clearTimeout(timeoutId);
  }
  const text = await response.text();
  let data = null;
  if (text.trim()) {
    try {
      data = JSON.parse(text);
    } catch {
      data = null;
    }
  }
  if (!response.ok || data?.ok === false) {
    const error = new Error(cwsAiErrorMessage(data?.error, response.status));
    error.code = data?.error || `http_${response.status}`;
    error.status = response.status;
    throw error;
  }
  return data || { ok: true };
}

function cwsAiApiUrl(path) {
  return new URL(path, cwsAiApiBaseUrl()).href;
}

function cwsAiApiBaseUrl() {
  const pathname = window.location.pathname || "/";
  const cwsIndex = pathname.indexOf("/cws");
  if (window.location.protocol !== "file:" && cwsIndex >= 0) {
    return `${window.location.origin}${pathname.slice(0, cwsIndex + 4)}/${CWS_AI_API_BASE_PATH}`;
  }
  if (window.location.protocol === "file:" || window.location.origin !== "https://chtec.co.jp") {
    return CWS_AI_REMOTE_API_BASE_URL;
  }
  return new URL(CWS_AI_API_BASE_PATH, window.location.href).href;
}

function cwsAiShouldUseSameOriginCredentials(url) {
  try {
    return new URL(url).origin === window.location.origin && window.location.protocol !== "file:";
  } catch {
    return false;
  }
}

function cwsAiRequestAuthToken() {
  return aiState().authToken || cwsAiReadStoredAuthToken();
}

function cwsAiErrorMessage(code, status) {
  const messages = {
    invalid_credentials: "メールアドレスまたはパスワードが違います。",
    invalid_email: "メールアドレスを確認してください。",
    email_already_registered: "このメールアドレスは登録済みです。",
    user_not_found: "このメールアドレスのアカウントが見つかりません。",
    password_too_short: "パスワードは8文字以上です。",
    password_too_long: "パスワードが長すぎます。",
    password_hash_failed: "パスワードを保存できませんでした。",
    invalid_current_password: "現在のパスワードが違います。",
    display_name_too_long: "表示名が長すぎます。",
    regist_disabled: "登録は現在停止しています。",
    invalid_code: "確認コードを確認してください。",
    code_expired: "確認コードの有効期限が切れました。再送してください。",
    code_too_many_attempts: "確認コードの入力回数が上限に達しました。再送してください。",
    code_create_failed: "確認コードを作成できませんでした。",
    code_send_failed: "確認コードを送信できませんでした。",
    smtp_config_missing: "メール送信設定が未設定です。",
    not_authenticated: "ログインしてください。",
    bad_origin: "ページを再読み込みしてください。",
    missing_message: "メッセージを入力してください。",
    message_too_long: "メッセージが長すぎます。",
    workbookContext_too_long: "送信する表の情報が多すぎます。",
    devlune_session_missing: "AI セッションを作成できませんでした。",
    devlune_unreachable: "AIは休憩中です。",
    devlune_bad_response: "AIは休憩中です。",
    devlune_error: "AIは休憩中です。",
  };
  if (code && messages[code]) return messages[code];
  if (status === 404) return "API が見つかりません。";
  if (status === 401) return "ログインしてください。";
  if (status >= 500) return "サーバー側で処理できませんでした。";
  return "通信できませんでした。";
}

function cwsAiStatusBlocksChat(status) {
  return status === "checking" || status === "resting";
}

function cwsAiIsUnavailableError(error) {
  const code = String(error?.code || "");
  return (
    code === "devlune_unreachable" ||
    code === "devlune_bad_response" ||
    code === "devlune_error" ||
    Number(error?.status) === 502 ||
    Number(error?.status) === 503 ||
    Number(error?.status) === 504
  );
}

function cwsAiExtractReply(data) {
  const candidates = [
    data?.reply,
    data?.message,
    data?.text,
    data?.devlune?.reply,
    data?.devlune?.message,
    data?.devlune?.text,
    data?.devlune?.assistantMessage?.content,
    data?.devlune?.assistant?.content,
  ];
  const direct = candidates.find((value) => typeof value === "string" && value.trim());
  if (direct) return direct.trim();
  const messages = data?.devlune?.messages || data?.messages;
  if (Array.isArray(messages)) {
    const reply = [...messages].reverse().find((item) => {
      const role = String(item?.role || "").toLowerCase();
      return role === "assistant" || role === "ai";
    });
    const content = reply?.content || reply?.text || reply?.message;
    if (typeof content === "string" && content.trim()) return content.trim();
  }
  return "応答を受信しました。";
}

function cwsAiWorkbookKey() {
  const ai = aiState();
  if (!ai.settings.workbookScoped) return "global";
  const state = runtime.state;
  const source = state.currentSaveFileName || state.openedFileSignature?.name || state.model?.sourceName || document.title || "workbook";
  const key = String(source)
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9._-]+/g, "-")
    .replace(/^-+|-+$/g, "")
    .slice(0, 96);
  return key || "workbook";
}

function cwsAiWorkbookContextText() {
  const state = runtime.state;
  if (!state.model) return "";
  const sheet = activeSheet();
  const lines = [
    "App: Cht WebSheet",
    `Workbook key: ${cwsAiWorkbookKey()}`,
    `Workbook name: ${state.model.sourceName || ""}`,
    `Sheet count: ${state.model.sheets?.length || 0}`,
    `Active sheet index: ${state.activeSheetIndex}`,
  ];

  if (sheet) {
    lines.push(`Active sheet: ${sheet.name || ""}`);
    lines.push(`Active sheet size: ${sheet.rowCount || 0} rows x ${sheet.colCount || 0} columns`);
  }

  lines.push("Sheets:");
  (state.model.sheets || []).slice(0, 30).forEach((item, index) => {
    lines.push(`- [${index}] ${item.name || `Sheet${index + 1}`}`);
  });

  const ai = aiState();
  if (sheet && ai.settings.includeSelection) {
    cwsAiAppendRangeSnapshot(lines, "Selection", cwsAiRangeSnapshot(sheet, activeSelectionRange(), CWS_AI_MAX_CONTEXT_CELLS));
  }
  if (sheet && ai.settings.includeVisibleRange) {
    cwsAiAppendRangeSnapshot(lines, "Visible range", cwsAiRangeSnapshot(sheet, visibleGridRange(sheet), CWS_AI_MAX_CONTEXT_CELLS));
  }
  return lines.join("\n");
}

function cwsAiAppendRangeSnapshot(lines, label, snapshot) {
  if (!snapshot) return;
  lines.push("");
  lines.push(`${label}: ${snapshot.range || ""}`);
  lines.push(`${label} scanned range: ${snapshot.scanRange || ""}`);
  if (snapshot.scanned !== undefined) {
    lines.push(`${label} scanned cells: ${snapshot.scanned}`);
  }
  if (snapshot.truncated) {
    lines.push(`${label} truncated: true`);
  }
  if (!snapshot.cells.length) {
    lines.push(`${label} cells: none`);
    return;
  }
  lines.push(`${label} cells:`);
  snapshot.cells.forEach((cell) => {
    lines.push(`- ${cell.address}`);
    if (cell.raw && cell.raw !== cell.display) {
      cwsAiAppendCellText(lines, "raw", cell.raw);
    }
    cwsAiAppendCellText(lines, cell.raw === cell.display ? "value" : "display", cell.display || cell.raw);
  });
}

function cwsAiAppendCellText(lines, label, value) {
  const text = cwsAiNormalizeCellText(value);
  if (!text) return;
  lines.push(`  ${label}:`);
  text.split("\n").forEach((line) => {
    lines.push(`    ${line}`);
  });
}

function cwsAiRangeSnapshot(sheet, range, maxCells) {
  if (!sheet || !range) return null;
  const state = runtime.state;
  const materialized = materializeSelectionRangeForSheet(sheet, range, range.sheetIndex ?? state.activeSheetIndex);
  if (!materialized) return null;
  const top = Math.max(1, materialized.top || 1);
  const left = Math.max(1, materialized.left || 1);
  const bottom = Math.min(sheet.rowCount || top, materialized.bottom || top, top + 39);
  const right = Math.min(sheet.colCount || left, materialized.right || left, left + 19);
  const sheetId = getSheetId(sheet.name);
  const cells = [];
  let scanned = 0;
  for (let row = top; row <= bottom; row += 1) {
    for (let col = left; col <= right; col += 1) {
      scanned += 1;
      const raw = cwsAiTrimCellText(getCellRawInput(sheet, row, col));
      const display = cwsAiTrimCellText(getDisplayForCell(sheet, sheetId, row, col));
      if (!raw && !display) continue;
      cells.push({
        address: `${columnName(col)}${row}`,
        raw,
        display,
      });
      if (cells.length >= maxCells) {
        return {
          range: rangeToLabel(materialized),
          scanRange: rangeToLabel({ top, left, bottom, right }),
          cells,
          truncated: true,
        };
      }
    }
  }
  return {
    range: rangeToLabel(materialized),
    scanRange: rangeToLabel({ top, left, bottom, right }),
    cells,
    scanned,
    truncated: bottom < materialized.bottom || right < materialized.right,
  };
}

function cwsAiTrimCellText(value) {
  const text = cwsAiNormalizeCellText(value);
  return text.length > CWS_AI_MAX_CELL_TEXT_LENGTH ? `${text.slice(0, CWS_AI_MAX_CELL_TEXT_LENGTH)}...` : text;
}

function cwsAiNormalizeCellText(value) {
  return String(value ?? "").replace(/\r\n?/g, "\n");
}

function cwsAiUserLabel(user) {
  if (!user) return "未ログイン";
  return user.displayName || user.email || "ログイン中";
}

function cwsAiPruneMessages() {
  const messages = aiState().messages;
  const overflow = messages.length - CWS_AI_MAX_MESSAGES;
  if (overflow > 0) {
    messages.splice(0, overflow);
  }
}

function scrollCwsAiMessagesToEnd() {
  const container = document.querySelector("[data-cws-ai-messages]");
  if (container) {
    container.scrollTop = container.scrollHeight;
  }
}

function preserveCwsAiMessagesScroll(previousScrollHeight, previousScrollTop) {
  const container = document.querySelector("[data-cws-ai-messages]");
  if (!container) return;
  const delta = container.scrollHeight - previousScrollHeight;
  container.scrollTop = Math.max(0, previousScrollTop + delta);
}

function aiState() {
  ensureCwsAiState();
  return runtime.state.cwsAi;
}

function bumpCwsAiAuthGeneration() {
  const ai = aiState();
  ai.authGeneration = (Number(ai.authGeneration) || 0) + 1;
  return ai.authGeneration;
}

function cwsAiAuthGenerationIsCurrent(generation) {
  return aiState().authGeneration === generation;
}

function clampNumber(value, min, max) {
  return Math.min(Math.max(value, min), max);
}

function escapeHtml(value) {
  return runtime.helpers.escapeHtml(value);
}

function escapeAttr(value) {
  return runtime.helpers.escapeAttr(value);
}

function activeSheet() {
  return runtime.helpers.activeSheet();
}

function activeSelectionRange() {
  return runtime.helpers.activeSelectionRange();
}

function visibleGridRange(sheet) {
  return runtime.helpers.visibleGridRange(sheet);
}

function materializeSelectionRangeForSheet(sheet, range, sheetIndex) {
  return runtime.helpers.materializeSelectionRangeForSheet(sheet, range, sheetIndex);
}

function getSheetId(sheetName) {
  return runtime.helpers.getSheetId(sheetName);
}

function getCellRawInput(sheet, row, col) {
  return runtime.helpers.getCellRawInput(sheet, row, col);
}

function getDisplayForCell(sheet, sheetId, row, col) {
  return runtime.helpers.getDisplayForCell(sheet, sheetId, row, col);
}

function columnName(index) {
  return runtime.helpers.columnName(index);
}

function rangeToLabel(range) {
  return runtime.helpers.rangeToLabel(range);
}

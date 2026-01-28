import * as React from "react";
import styles from "./ChatKitAgent.module.scss";
import type { IChatKitAgentProps } from "./IChatKitAgentProps";

type TokenResponse = {
  client_secret: string;
  expires_at?: number | string;
  error?: string;
};

async function safeReadText(res: Response): Promise<string> {
  try {
    return await res.text();
  } catch {
    return "";
  }
}

function loadChatKitScriptOnce(): Promise<void> {
  if (window.__OPENAI_CHATKIT_SCRIPT_LOADING__)
    return window.__OPENAI_CHATKIT_SCRIPT_LOADING__;

  window.__OPENAI_CHATKIT_SCRIPT_LOADING__ = new Promise<void>((resolve, reject) => {
    const src = "https://cdn.platform.openai.com/deployments/chatkit/chatkit.js";
    const existing = document.querySelector<HTMLScriptElement>(`script[src="${src}"]`);

    if (existing) {
      existing.addEventListener("load", () => resolve(), { once: true });
      existing.addEventListener("error", () => reject(new Error("Failed to load ChatKit script")), { once: true });
      return;
    }

    const script = document.createElement("script");
    script.src = src;
    script.async = true;
    script.addEventListener("load", () => resolve(), { once: true });
    script.addEventListener("error", () => reject(new Error("Failed to load ChatKit script")), { once: true });
    document.head.appendChild(script);
  });

  return window.__OPENAI_CHATKIT_SCRIPT_LOADING__;
}

function whenDefined(tagName: string, timeoutMs = 10000): Promise<void> {
  if (window.customElements.get(tagName)) return Promise.resolve();

  return new Promise<void>((resolve, reject) => {
    const start = Date.now();
    const tick = () => {
      if (window.customElements.get(tagName)) return resolve();
      if (Date.now() - start > timeoutMs)
        return reject(new Error(`Timed out waiting for ${tagName}`));
      setTimeout(tick, 50);
    };
    tick();
  });
}

async function fetchClientSecret(
  lambdaUrl: string,
  workflowId: string,
  userId: string
): Promise<TokenResponse> {
  const res = await fetch(lambdaUrl, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ workflow_id: workflowId, user: userId })
  });

  if (!res.ok) {
    const text = await safeReadText(res);
    throw new Error(`Token endpoint failed (${res.status}): ${text || res.statusText}`);
  }

  const data = (await res.json()) as TokenResponse;
  if (!data?.client_secret)
    throw new Error(data?.error || "Failed to get client_secret");
  return data;
}

export default function ChatKitAgent(props: IChatKitAgentProps): JSX.Element {
  const lambdaUrl = props.lambdaUrl.trim();
  const workflowId = props.workflowId.trim();
  const userId = props.userId || "sharepoint-user";
  const greeting = props.greeting?.trim();

  const hasConfig = Boolean(lambdaUrl && workflowId);

  const [open, setOpen] = React.useState(false);
  const [busy, setBusy] = React.useState(false);
  const [error, setError] = React.useState("");

  const hostRef = React.useRef<HTMLDivElement | null>(null);
  const chatElRef = React.useRef<HTMLElementTagNameMap["openai-chatkit"] | null>(null);
  const initializingRef = React.useRef(false);

  const tokenRef = React.useRef<{ clientSecret: string; expiresAt?: number | string } | null>(null);

  const cleanupWidget = React.useCallback(() => {
    try {
      hostRef.current?.replaceChildren();
    } catch {}
    chatElRef.current = null;
    initializingRef.current = false;
    setBusy(false);
  }, []);

  const mountChatIfNeeded = React.useCallback(async () => {
    if (!open || !hasConfig || initializingRef.current) return;

    if (chatElRef.current && hostRef.current?.contains(chatElRef.current)) return;

    cleanupWidget();
    initializingRef.current = true;
    setBusy(true);
    setError("");

    try {
      await loadChatKitScriptOnce();
      await whenDefined("openai-chatkit");
      if (!hostRef.current) return;

      const el = document.createElement("openai-chatkit") as HTMLElementTagNameMap["openai-chatkit"];
      el.classList.add(styles.chatKitFrame);

      el.setOptions({
        api: {
          getClientSecret: async (current?: string) => {
            if (!current && tokenRef.current) return tokenRef.current.clientSecret;
            const t = await fetchClientSecret(lambdaUrl, workflowId, userId);
            tokenRef.current = { clientSecret: t.client_secret, expiresAt: t.expires_at };
            return t.client_secret;
          }
        },
        history: { enabled: true, showDelete: true, showRename: true },
        startScreen: {
          greeting: greeting || "Hi! How can I help?"
        }
      });

      el.style.width = "100%";
      el.style.height = "100%";

      hostRef.current.appendChild(el);
      chatElRef.current = el;

      el.focusComposer?.();
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e));
      cleanupWidget();
    } finally {
      initializingRef.current = false;
      setBusy(false);
    }
  }, [open, hasConfig, lambdaUrl, workflowId, userId, greeting, cleanupWidget]);

  React.useEffect(() => {
    void mountChatIfNeeded();
  }, [mountChatIfNeeded]);

  const toggle = () => {
    if (open) cleanupWidget();
    setOpen(!open);
  };

  const resetConversation = () => {
    tokenRef.current = null;
    cleanupWidget();
    if (open) void mountChatIfNeeded();
  };

  return (
    <section className={styles.shell}>
      <button className={styles.bubble} onClick={toggle} disabled={!hasConfig}>
        💬
      </button>

      {open && (
        <div className={styles.panel}>
          <div className={styles.panelHeader}>
            <span>Chat</span>
          </div>

          {busy && <div className={styles.meta}>Connecting…</div>}
          {error && <div className={styles.error}>{error}</div>}

          <div ref={hostRef} className={styles.widgetHost} />
        </div>
      )}
    </section>
  );
}

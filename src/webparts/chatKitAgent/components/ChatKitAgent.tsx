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
  if (window.__OPENAI_CHATKIT_SCRIPT_LOADING__) return window.__OPENAI_CHATKIT_SCRIPT_LOADING__;

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
      if (Date.now() - start > timeoutMs) return reject(new Error(`Timed out waiting for ${tagName}`));
      setTimeout(tick, 50);
    };
    tick();
  });
}

async function fetchClientSecret(lambdaUrl: string, workflowId: string, userId: string): Promise<TokenResponse> {
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
  if (!data?.client_secret) throw new Error(data?.error || "Failed to get client_secret");
  return data;
}

export default function ChatKitAgent(props: IChatKitAgentProps): JSX.Element {
  const lambdaUrl = (props.lambdaUrl || "").trim();
  const workflowId = (props.workflowId || "").trim();
  const userId = (props.userId || "sharepoint-user").trim();
  const greeting = (props.greeting || "").trim();

  const hasConfig = Boolean(lambdaUrl && workflowId);

  const [open, setOpen] = React.useState(false);
  const [busy, setBusy] = React.useState(false);
  const [error, setError] = React.useState<string>("");

  const hostRef = React.useRef<HTMLDivElement | null>(null);
  const chatElRef = React.useRef<HTMLElementTagNameMap["openai-chatkit"] | null>(null);
  const initializingRef = React.useRef(false);

  // Keep token in memory (optional). If you always want a fresh token per open, clear this on open.
  const tokenRef = React.useRef<{ clientSecret: string; expiresAt?: number | string } | null>(null);

  const mountChatIfNeeded = React.useCallback(async () => {
    if (!hasConfig) {
      setError("Configure Lambda URL and Workflow ID in the web part settings.");
      return;
    }
    if (!open) return;
    if (chatElRef.current || initializingRef.current) return;

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
            const needFresh = Boolean(current);
            if (!needFresh && tokenRef.current?.clientSecret) return tokenRef.current.clientSecret;

            const t = await fetchClientSecret(lambdaUrl, workflowId, userId);
            tokenRef.current = { clientSecret: t.client_secret, expiresAt: t.expires_at };
            return t.client_secret;
          }
        },

        header: { title: { text: "Hi there! 👋" } },

        history: { enabled: true, showDelete: true, showRename: true },

        theme: {
          colorScheme: "dark",
          color: { accent: { primary: "#F1246A", level: 2 } },
          radius: "round",
          density: "compact",
          typography: { fontFamily: "'Inter', sans-serif" }
        },

        startScreen: {
          greeting:
            greeting ||
            "Hi! I'm AskSara, your HeySara sidekick here to help with everything within the family. Need help? Just type your question below."
        }
      });

      el.style.width = "100%";
      el.style.height = "100%";

      hostRef.current.innerHTML = "";
      hostRef.current.appendChild(el);
      chatElRef.current = el;

      try {
        el.focusComposer && el.focusComposer();
      } catch {}
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e));
      if (hostRef.current) hostRef.current.innerHTML = "";
      chatElRef.current = null;
    } finally {
      setBusy(false);
      initializingRef.current = false;
    }
  }, [hasConfig, open, lambdaUrl, workflowId, userId, greeting]);

  React.useEffect(() => {
    void mountChatIfNeeded();
  }, [mountChatIfNeeded]);

  const toggle = React.useCallback(() => {
    setOpen((v) => !v);
  }, []);

  const close = React.useCallback(() => {
    setOpen(false);
  }, []);

  const resetConversation = React.useCallback(() => {
    tokenRef.current = null;
    if (hostRef.current) hostRef.current.innerHTML = "";
    chatElRef.current = null;
    if (open) void mountChatIfNeeded();
  }, [open, mountChatIfNeeded]);

  return (
    <section className={styles.shell} aria-label="Chat Kit Agent">
      {/* Bubble */}
      <button className={styles.bubble} onClick={toggle} aria-haspopup="dialog" aria-expanded={open} disabled={!hasConfig}>
        <span className={styles.bubbleIcon}>💬</span>
      </button>

      {/* Panel */}
      {open ? (
        <div className={styles.panel} role="dialog" aria-label="Chat">
          <div className={styles.panelHeader}>
            <div className={styles.panelTitle}>Chat</div>
            <div className={styles.panelActions}>
              <button className={styles.smallBtn} onClick={resetConversation} title="New session">
                ↻
              </button>
              <button className={styles.smallBtn} onClick={close} title="Close">
                ✕
              </button>
            </div>
          </div>

          {busy ? <div className={styles.meta}>Connecting…</div> : null}
          {error ? <div className={styles.error}>{error}</div> : null}

          <div ref={hostRef} className={styles.widgetHost} />
        </div>
      ) : null}
    </section>
  );
}

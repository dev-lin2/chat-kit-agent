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
  if (window.__OPENAI_CHATKIT_SCRIPT_LOADING__) {
    return window.__OPENAI_CHATKIT_SCRIPT_LOADING__;
  }

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
    const tick = (): void => {
      if (window.customElements.get(tagName)) {
        resolve();
        return;
      }
      if (Date.now() - start > timeoutMs) {
        reject(new Error(`Timed out waiting for ${tagName}`));
        return;
      }
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

  const [open, setOpen] = React.useState<boolean>(false);
  const [busy, setBusy] = React.useState<boolean>(false);
  const [error, setError] = React.useState<string>("");

  const hostRef = React.useRef<HTMLDivElement | null>(null);
  const chatElRef = React.useRef<HTMLElementTagNameMap["openai-chatkit"] | null>(null);
  const tokenRef = React.useRef<{ clientSecret: string; expiresAt?: number | string } | null>(null);

  // init sequence token: increments on each open/close to cancel late async completions
  const initSeqRef = React.useRef<number>(0);

  // Synchronous commit helper to satisfy require-atomic-updates
  const commit = React.useCallback((fn: () => void): void => {
    fn();
  }, []);

  const cleanupWidget = React.useCallback((): void => {
    const host = hostRef.current;
    if (host) {
      try {
        host.replaceChildren();
      } catch {
        host.innerHTML = "";
      }
    }
    chatElRef.current = null;
    setBusy(false);
  }, []);

  const mountChatIfNeeded = React.useCallback(async (): Promise<void> => {
    if (!open || !hasConfig) return;

    const host = hostRef.current;
    if (!host) return;

    const existing = chatElRef.current;
    if (existing && host.contains(existing)) return;

    // start new init sequence
    const seq = initSeqRef.current + 1;
    initSeqRef.current = seq;

    setBusy(true);
    setError("");

    try {
      await loadChatKitScriptOnce();
      await whenDefined("openai-chatkit");

      // abort if a newer init started
      if (initSeqRef.current !== seq) return;

      const el = document.createElement("openai-chatkit") as HTMLElementTagNameMap["openai-chatkit"];
      el.classList.add(styles.chatKitFrame);

      el.setOptions({
        api: {
          getClientSecret: async (current?: string) => {
            const cached = tokenRef.current;
            if (!current && cached?.clientSecret) return cached.clientSecret;

            const t = await fetchClientSecret(lambdaUrl, workflowId, userId);

            if (initSeqRef.current === seq) {
              commit(() => {
                tokenRef.current = { clientSecret: t.client_secret, expiresAt: t.expires_at };
              });
            }

            return t.client_secret;
          }
        },
        history: { enabled: true, showDelete: true, showRename: true },
        startScreen: { greeting: greeting || "Hi! How can I help?" }
      });

      el.style.width = "100%";
      el.style.height = "100%";

      host.replaceChildren(el);

      if (initSeqRef.current === seq) {
        commit(() => {
          chatElRef.current = el;
        });
      }

      try {
        el.focusComposer?.();
      } catch (err) {
        console.debug("focusComposer failed", err);
      }
    } catch (e) {
      if (initSeqRef.current === seq) {
        setError(e instanceof Error ? e.message : String(e));
        cleanupWidget();
      }
    } finally {
      if (initSeqRef.current === seq) {
        setBusy(false);
      }
    }
  }, [open, hasConfig, lambdaUrl, workflowId, userId, greeting, cleanupWidget, commit]);

  React.useEffect(() => {
    mountChatIfNeeded().catch((e) => console.error(e));
  }, [mountChatIfNeeded]);

  const toggle = React.useCallback((): void => {
    setOpen((v) => {
      const next = !v;
      if (!next) {
        initSeqRef.current += 1; // cancel in-flight init
        cleanupWidget();
      }
      return next;
    });
  }, [cleanupWidget]);

  return (
    <section className={styles.shell} aria-label="Chat Kit Agent">
      <button
        className={styles.bubble}
        onClick={toggle}
        aria-haspopup="dialog"
        aria-expanded={open}
        disabled={!hasConfig}
      >
        <span className={styles.bubbleIcon}>💬</span>
      </button>

      {open ? (
        <div className={styles.panel} role="dialog" aria-label="Chat">
          <div className={styles.panelHeader}>
            <div className={styles.panelTitle}>Chat</div>
          </div>

          {busy ? <div className={styles.meta}>Connecting…</div> : null}
          {error ? <div className={styles.error}>{error}</div> : null}

          <div ref={hostRef} className={styles.widgetHost} />
        </div>
      ) : null}
    </section>
  );
}

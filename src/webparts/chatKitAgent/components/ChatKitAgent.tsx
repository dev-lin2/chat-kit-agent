import * as React from "react";
import styles from "./ChatKitAgent.module.scss";
import type { IChatKitAgentProps } from "./IChatKitAgentProps";

// Registers/typings for the <openai-chatkit> element
import "@openai/chatkit";

type TokenResponse = {
  client_secret: string;
  expires_at?: number | string;
  session_id?: string;
  workflow_id?: string;
};

declare global {
  interface Window {
    __OPENAI_CHATKIT_SCRIPT_LOADING__?: Promise<void>;
  }

  // Minimal type for the web component API we need
  interface HTMLElementTagNameMap {
    "openai-chatkit": HTMLElement & {
      setOptions: (options: any) => void;
    };
  }
}

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
  if (!data?.client_secret) throw new Error("Token endpoint response missing client_secret");
  return data;
}

export default function ChatKitAgent(props: IChatKitAgentProps): JSX.Element {
  const lambdaUrl = (props.lambdaUrl || "").trim();
  const workflowId = (props.workflowId || "").trim();
  const userId = (props.userId || "sharepoint-user").trim();
  const hasConfig = Boolean(lambdaUrl && workflowId);

  const hostRef = React.useRef<HTMLDivElement | null>(null);
  const chatElRef = React.useRef<HTMLElementTagNameMap["openai-chatkit"] | null>(null);

  const [error, setError] = React.useState<string>("");
  const [meta, setMeta] = React.useState<string>(() =>
    hasConfig ? "Loading..." : "Configure Lambda URL and Workflow ID in the web part settings"
  );

  // Keep token only in memory
  const tokenRef = React.useRef<{ clientSecret: string; expiresAt?: number | string } | null>(null);

  const resetSession = React.useCallback(() => {
    tokenRef.current = null;
    setError("");
    setMeta("Session cleared");
    // Recreate the element to ensure a clean session
    if (hostRef.current) {
      hostRef.current.innerHTML = "";
      chatElRef.current = null;
    }
  }, []);

  React.useEffect(() => {
    let cancelled = false;

    async function init() {
      setError("");
      tokenRef.current = null;

      if (!hasConfig) {
        setMeta("Configure Lambda URL and Workflow ID in the web part settings");
        return;
      }

      setMeta("Loading ChatKit...");
      await loadChatKitScriptOnce(); // required for the web component runtime :contentReference[oaicite:2]{index=2}
      if (cancelled) return;

      if (!hostRef.current) return;

      // Create/mount <openai-chatkit>
      hostRef.current.innerHTML = "";
      const el = document.createElement("openai-chatkit");
      el.classList.add(styles.chatKitFrame);
      hostRef.current.appendChild(el);

      // Configure it using the managed ChatKit flow: provide getClientSecret :contentReference[oaicite:3]{index=3}
      el.setOptions({
        api: {
          getClientSecret: async (currentClientSecret?: string) => {
            // If ChatKit passes an existing secret, you can choose to refresh it.
            // Minimal approach: reuse it if present, else mint a new one.
            if (currentClientSecret) return currentClientSecret;
            if (tokenRef.current?.clientSecret) return tokenRef.current.clientSecret;

            setMeta("Creating session...");
            const t = await fetchClientSecret(lambdaUrl, workflowId, userId);
            tokenRef.current = { clientSecret: t.client_secret, expiresAt: t.expires_at };
            setMeta("Ready");
            return t.client_secret;
          }
        }
      });

      chatElRef.current = el;
      setMeta("Ready");
    }

    init().catch((e) => {
      if (cancelled) return;
      setError(e instanceof Error ? e.message : String(e));
      setMeta("Error");
    });

    return () => {
      cancelled = true;
    };
  }, [hasConfig, lambdaUrl, workflowId, userId]);

  return (
    <section className={styles.chatKitAgent}>
      <div className={styles.header}>
        <h2 className={styles.title}>{props.title || "Chat Kit Agent"}</h2>
        <button className={styles.button} onClick={resetSession} disabled={!hasConfig}>
          New session
        </button>
      </div>

      <div className={styles.meta}>{meta}</div>
      {error ? <div className={styles.error}>{error}</div> : null}

      <div ref={hostRef} className={styles.widgetHost} />
    </section>
  );
}

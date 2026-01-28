export { };

declare global {
  namespace JSX {
    interface IntrinsicElements {
      "openai-chatkit": React.DetailedHTMLProps<React.HTMLAttributes<HTMLElement>, HTMLElement>;
    }
  }

  interface Window {
    __OPENAI_CHATKIT_SCRIPT_LOADING__?: Promise<void>;
  }

  interface HTMLElementTagNameMap {
    "openai-chatkit": HTMLElement & {
      setOptions: (options: any) => void;
      focusComposer?: () => void;
    };
  }
}

/// <reference types="vitest/config" />
import { defineConfig, type Plugin } from "vite";
import { svelte } from "@sveltejs/vite-plugin-svelte";

// 只喺正式打包時加入 CSP；開發模式要用 inline script 同 websocket，加咗會壞
const CSP = [
  "default-src 'self'",
  "script-src 'self'",
  "style-src 'self' 'unsafe-inline'",
  "img-src 'self' data: blob:",
  "font-src 'self'",
  "connect-src 'self'",
  "worker-src 'self' blob:",
  "object-src 'none'",
  "base-uri 'none'",
  "form-action 'none'",
].join("; ");

function cspPlugin(): Plugin {
  return {
    name: "cleanmark-csp",
    apply: "build",
    transformIndexHtml(html) {
      return html.replace(
        "<head>",
        `<head>\n    <meta http-equiv="Content-Security-Policy" content="${CSP}" />`,
      );
    },
  };
}

export default defineConfig({
  base: "./",
  plugins: [svelte(), cspPlugin()],
  worker: { format: "es" },
  test: {
    include: ["tests/**/*.test.ts"],
    environment: "node",
  },
});

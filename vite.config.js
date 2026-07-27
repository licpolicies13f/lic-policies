import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import tailwindcss from "@tailwindcss/vite";
import { viteSingleFile } from "vite-plugin-singlefile";
import { existsSync, readFileSync, writeFileSync } from "node:fs";
import { resolve } from "node:path";

function fileProtocolCompatibleBuild() {
  let outputDirectory;

  return {
    name: "file-protocol-compatible-build",
    apply: "build",
    enforce: "post",
    configResolved(config) {
      outputDirectory = resolve(config.root, config.build.outDir);
    },
    closeBundle() {
      const outputFile = resolve(outputDirectory, "index.html");
      if (!existsSync(outputFile)) return;

      const html = readFileSync(outputFile, "utf8");
      const moduleScriptPattern =
        /<script\b(?=[^>]*\btype=["']module["'])[^>]*>([\s\S]*?)<\/script>/i;
      const moduleScript = html.match(moduleScriptPattern);
      let fileSafeHtml = html.replaceAll(
        "import.meta.url",
        "document.baseURI",
      );

      if (moduleScript) {
        const classicScript = `<script>${moduleScript[1].replaceAll(
          "import.meta.url",
          "document.baseURI",
        )}</script>`;
        fileSafeHtml = fileSafeHtml
          .replace(moduleScriptPattern, "")
          .replace("</body>", `${classicScript}\n</body>`);
      }

      writeFileSync(outputFile, fileSafeHtml, "utf8");
    },
  };
}

export default defineConfig({
  plugins: [
    react(),
    tailwindcss(),
    viteSingleFile(),
    fileProtocolCompatibleBuild(),
  ],
  base: "./",
});

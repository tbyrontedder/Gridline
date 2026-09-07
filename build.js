import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
const root = path.dirname(fileURLToPath(import.meta.url));
const names = ['engine', 'renderer', 'sample', 'io', 'app'];
let js = '"use strict";\n(() => {\nconst __modules = Object.create(null);\n';
for (const name of names) {
  let source = fs.readFileSync(path.join(root, 'src', name + '.js'), 'utf8');
  const exported = [...source.matchAll(/export\s+(?:(?:async\s+)?function\*?|class|const|let|var)\s+(\w+)/g)].map(m => m[1]);
  source = source.replace(/import\s+\{([^}]+)\}\s+from\s+['"]\.\/([^'"]+)\.js['"];?/g, (_, identifiers, dependency) => `const {${identifiers}} = __modules[${JSON.stringify(dependency)}];`);
  source = source.replace(/\bexport\s+(?=(?:async\s+)?(?:function|class|const|let|var))/g, '');
  js += `\n// ===== ${name}.js =====\n__modules[${JSON.stringify(name)}] = (() => {\n${source}\nreturn { ${exported.join(', ')} };\n})();\n`;
}
js += '\n})();\n';
const css = fs.readFileSync(path.join(root, 'styles.css'), 'utf8');
let html = fs.readFileSync(path.join(root, 'index.html'), 'utf8');
html = html.replace('<link rel="stylesheet" href="styles.css">', () => `<style>\n${css}\n</style>`).replace('<script type="module" src="src/app.js"></script>', () => `<script>\n${js.replaceAll('</script', '<\\/script')}\n</script>`);
fs.mkdirSync(path.join(root, 'dist'), { recursive: true });
fs.writeFileSync(path.join(root, 'dist', 'index.html'), html);
fs.writeFileSync(path.join(root, 'dist', 'gridline.js'), js);
console.log(`Built dist/index.html — ${Buffer.byteLength(html).toLocaleString()} bytes. No runtime dependencies.`);

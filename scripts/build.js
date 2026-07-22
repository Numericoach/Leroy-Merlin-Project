const fs = require('fs');
const path = require('path');
const ts = require('typescript');

const srcDir = path.join(__dirname, '..', 'src');
const files = fs.readdirSync(srcDir);

files.forEach(file => {
  if (file.endsWith('.ts')) {
    const tsPath = path.join(srcDir, file);
    const jsPath = path.join(srcDir, file.replace(/\.ts$/, '.js'));
    const tsCode = fs.readFileSync(tsPath, 'utf8');
    const result = ts.transpileModule(tsCode, {
      compilerOptions: {
        target: ts.ScriptTarget.ES2020,
        module: ts.ModuleKind.None,
        removeComments: false
      }
    });

    let code = result.outputText;
    code = code.replace(/Object\.defineProperty\(exports,\s*"__esModule",\s*\{\s*value:\s*true\s*\}\);?/g, '');
    code = code.replace(/exports\.[a-zA-Z0-9_]+\s*=\s*/g, '');
    code = code.replace(/^export\s+/gm, '');

    fs.writeFileSync(jsPath, code);
    console.log(`Transpiled ${file} -> ${path.basename(jsPath)}`);
  }
});

const fs = require('fs');
const path = require('path');
const ts = require('typescript');

const srcDir = path.join(__dirname, '..', 'src');
const distDir = path.join(__dirname, '..', 'dist');

if (!fs.existsSync(distDir)) {
  fs.mkdirSync(distDir, { recursive: true });
}

const oldFiles = fs.readdirSync(distDir);
oldFiles.forEach(file => {
  const filePath = path.join(distDir, file);
  if (fs.statSync(filePath).isFile()) {
    fs.unlinkSync(filePath);
  }
});

const files = fs.readdirSync(srcDir);

files.forEach(file => {
  const srcPath = path.join(srcDir, file);
  const stat = fs.statSync(srcPath);

  if (stat.isFile()) {
    if (file.endsWith('.ts')) {
      const distJsPath = path.join(distDir, file.replace(/\.ts$/, '.js'));
      const tsCode = fs.readFileSync(srcPath, 'utf8');
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

      fs.writeFileSync(distJsPath, code);
      console.log(`Compiled src/${file} -> dist/${path.basename(distJsPath)}`);
    } else {
      const distPath = path.join(distDir, file);
      fs.copyFileSync(srcPath, distPath);
      console.log(`Copied src/${file} -> dist/${file}`);
    }
  }
});

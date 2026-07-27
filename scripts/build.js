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

function getAllFiles(dirPath, arrayOfFiles) {
  files = fs.readdirSync(dirPath);
  arrayOfFiles = arrayOfFiles || [];
  files.forEach(function(file) {
    if (fs.statSync(path.join(dirPath, file)).isDirectory()) {
      arrayOfFiles = getAllFiles(path.join(dirPath, file), arrayOfFiles);
    } else {
      arrayOfFiles.push(path.join(dirPath, file));
    }
  });
  return arrayOfFiles;
}

const allFiles = getAllFiles(srcDir);

allFiles.forEach(srcPath => {
  const fileBasename = path.basename(srcPath);
  if (fileBasename.endsWith('.ts')) {
    const distJsPath = path.join(distDir, fileBasename.replace(/\.ts$/, '.js'));
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
    code = code.replace(/const\s+[a-zA-Z0-9_]+\s*=\s*require\(.*?\);?/g, '');
    code = code.replace(/var\s+[a-zA-Z0-9_]+\s*=\s*require\(.*?\);?/g, '');
    code = code.replace(/let\s+[a-zA-Z0-9_]+\s*=\s*require\(.*?\);?/g, '');

    fs.writeFileSync(distJsPath, code);
    console.log(`Compiled ${srcPath} -> dist/${fileBasename.replace(/\.ts$/, '.js')}`);
  } else {
    const distPath = path.join(distDir, fileBasename);
    fs.copyFileSync(srcPath, distPath);
    console.log(`Copied ${srcPath} -> dist/${fileBasename}`);
  }
});

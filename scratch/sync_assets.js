const fs = require('fs');
const path = require('path');

const srcDir = path.resolve(__dirname, '../mobile-web');
const destDir = path.resolve(__dirname, '../android/app/src/main/assets');

function copyRecursiveSync(src, dest) {
  const exists = fs.existsSync(src);
  const stats = exists && fs.statSync(src);
  const isDirectory = exists && stats.isDirectory();
  if (isDirectory) {
    if (!fs.existsSync(dest)) {
      fs.mkdirSync(dest, { recursive: true });
    }
    fs.readdirSync(src).forEach((childItemName) => {
      copyRecursiveSync(path.join(src, childItemName), path.join(dest, childItemName));
    });
  } else {
    // For text files, ensure clean content copy
    const ext = path.extname(src).toLowerCase();
    const textExts = ['.html', '.css', '.js', '.json'];
    if (textExts.includes(ext)) {
      let content = fs.readFileSync(src, 'utf8');
      content = content.replace(/\r\n/g, '\n');
      fs.writeFileSync(src, content, 'utf8');
      fs.writeFileSync(dest, content, 'utf8');
    } else {
      fs.copyFileSync(src, dest);
    }
    console.log(`Synced: ${path.relative(srcDir, src)} -> ${path.relative(destDir, dest)}`);
  }
}

console.log('Synchronizing mobile-web assets to android/app/src/main/assets ...');
copyRecursiveSync(srcDir, destDir);
console.log('Synchronization complete!');

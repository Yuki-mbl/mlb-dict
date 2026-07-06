// 親リポジトリの index.html と images/ を www/ にコピーする
// 使い方: npm run copyweb   （または npm run sync でコピー＋cap sync）
const fs = require('fs');
const path = require('path');

const root = path.resolve(__dirname, '..');   // リポジトリ直下
const www = path.join(__dirname, 'www');
const imgSrc = path.join(root, 'images');
const imgDst = path.join(www, 'images');

fs.mkdirSync(imgDst, { recursive: true });

// index.html
fs.copyFileSync(path.join(root, 'index.html'), path.join(www, 'index.html'));

// images（未使用の BSO.jpeg は除外）
for (const f of fs.readdirSync(imgSrc)) {
  if (f.toLowerCase() === 'bso.jpeg') continue;
  fs.copyFileSync(path.join(imgSrc, f), path.join(imgDst, f));
}

console.log('Copied index.html and images/ into www/');

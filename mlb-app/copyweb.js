// 親リポジトリの index.html と images/ を www/ にコピーする
// 使い方: npm run copyweb   （または npm run sync でコピー＋cap sync）
const fs = require('fs');
const path = require('path');

const root = path.resolve(__dirname, '..');   // リポジトリ直下
const www = path.join(__dirname, 'www');
const imgSrc = path.join(root, 'images');
const imgDst = path.join(www, 'images');

fs.mkdirSync(imgDst, { recursive: true });

// index.html / privacy.html / terms.html
for (const f of ['index.html', 'privacy.html', 'terms.html']) {
  fs.copyFileSync(path.join(root, f), path.join(www, f));
}

// images（未使用の BSO.jpeg は除外）
for (const f of fs.readdirSync(imgSrc)) {
  if (f.toLowerCase() === 'bso.jpeg') continue;
  fs.copyFileSync(path.join(imgSrc, f), path.join(imgDst, f));
}

// sounds（打球音・歓声）
const sndSrc = path.join(root, 'sounds');
if (fs.existsSync(sndSrc)) {
  const sndDst = path.join(www, 'sounds');
  fs.mkdirSync(sndDst, { recursive: true });
  for (const f of fs.readdirSync(sndSrc)) {
    fs.copyFileSync(path.join(sndSrc, f), path.join(sndDst, f));
  }
}

console.log('Copied index.html, images/ and sounds/ into www/');

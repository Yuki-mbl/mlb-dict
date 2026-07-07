# TestFlight 配布手順（Mac で作業）

このフォルダは Web アプリ（`../index.html` ＋ `../images/`）を **Capacitor** で
iOS ネイティブアプリに包むためのプロジェクトです。Windows 側で以下は準備済み：

- `www/`（index.html と images を同梱＝オフライン動作）
- `capacitor.config.json`（appId: `jp.fantamstick.mlbdict` / アプリ名: MLB野球用語辞典）
- `resources/icon.png`（1024×1024）, `resources/splash.png`（2732×2732）
- `package.json` / `copyweb.js`（Web更新の取り込み用）

---

## 事前準備（Mac）
- Xcode（App Store から）＋ 一度起動してライセンス同意
- Node.js（18以上推奨）
- CocoaPods: `sudo gem install cocoapods`
- Apple Developer Program 登録済みのApple ID

## セットアップ（Macのターミナルで、この mlb-app フォルダ内）
```bash
npm install
npm run copyweb          # ../index.html と ../images を www/ にコピー
npx cap add ios          # ios/ ネイティブプロジェクト生成
npm run icons            # アイコン・スプラッシュを生成して反映
npx cap sync ios
npx cap open ios         # Xcode が開く
```
※ `www/` は git 管理外です。`npm run copyweb` で毎回リポジトリ直下の
最新 index.html / images から生成されます。

## Xcode での設定
1. 左の「App」ターゲット → **Signing & Capabilities**
   - **Team** に自分の Apple Developer アカウントを選択
   - **Bundle Identifier** が `jp.fantamstick.mlbdict` になっているか確認
     （App Store Connect で作るアプリと一致させる）
2. **General → Display Name** が「MLB野球用語辞典」か確認
3. 実機（iPhone）をつないで一度 Run し、動作確認（辞書・クイズ・今日の10問）

## App Store Connect
1. https://appstoreconnect.apple.com → マイApp → ＋ → 新規App
   - プラットフォーム: iOS / 名前 / 言語: 日本語 / Bundle ID: 上と同じ / SKU: 任意
2. まだ審査提出は不要（TestFlight配布だけならOK）

## ビルドをアップロード
1. Xcode 上部のデバイス選択で **Any iOS Device (arm64)** を選ぶ
2. メニュー **Product → Archive**
3. Organizer が開く → **Distribute App → App Store Connect → Upload**
4. 数分後、App Store Connect の **TestFlight** タブにビルドが現れる
   （「処理中」→ 完了まで数分〜数十分）

## テスターに配布
- **内部テスター**（自分・チーム最大100人）: 審査なしで即配布
  - TestFlight → 内部テスター → ユーザー追加 → ビルドを有効化
- **外部テスター**: 「Beta App Review」に提出（通常1日程度）→ 承認後に配布リンク
- テスターはスマホに **TestFlight アプリ** を入れて招待を受ける

---

## Web の内容を更新したら
1. リポジトリ直下で `python generate_app.py`（index.html 再生成）
2. この mlb-app フォルダで:
   ```bash
   npm run sync        # www/ に最新をコピー＋cap sync
   ```
3. Xcode でバージョン/ビルド番号を上げて再 Archive → Upload

## メモ
- 起動画面は **ネイティブスプラッシュ**（`resources/splash.png`＝Fantamstickロゴ／白背景）
  のみ。アプリ内の表紙スプラッシュは廃止し、起動後は直接ホーム画面が開きます。
  `npm run icons` でロゴがネイティブに反映されます。背景色は白（capacitor.config.json）。
- 得点・お気に入り等は端末内（localStorage）に保存されます。
- 審査ガイドライン 4.2 対策として Web を同梱しオフライン動作にしています。

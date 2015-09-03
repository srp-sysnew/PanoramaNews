# PanoramaNews
本プログラムは「panorama新聞」の[Monaca](https://ja.monaca.io/)用プロジェクトの一部(www以下)を抜粋したものです。  
「panorama新聞」は[株式会社セラフ](http://www.srp.co.jp/)にて作成されました。  
Windows Office365 API を利用し、Exchange Online メールを記事形式で表示するiPadアプリです。

# Requirement
* [Monaca](https://ja.monaca.io/)。ただし Cordova が利用できれば、その限りでないです。
* Windows AzureAD への アプリケーション(NativeClient)登録。

# Windows AzureAD
Office365 API を OAuthで利用しているので、あらかじめアプリケーションの設定が必要になります。  

1. AzureAD より アプリケーション(NativeClient)を新規追加。

2. リダイレクトURLは "http://localhost/callback" を設定。

3. アプリケーション追加で "Office 365 Exchange Online" のアクセス許可を追加しておく。

4. "クライアント ID" を控えておく。

# Usage
基本的な使い方の流れです。

1. [Monaca](https://ja.monaca.io/)にて「最小限のテンプレート」をベースに新規プロジェクト作成。

2. 「Cordovaプラグインの管理」を開く。  
a. Cordovaバージョン 4.1.0 を選択。  
b.「InAppBrowser」「Splashscreen」「MonacaPlugin」が含まれていることを確認。

3. 「JS/CSSコンポーネントの追加と削除」を開く。  
Cordova(PhoneGap)Loader」 (1.0.0)、「Monaca Core Utility」(2.0.4)、「jQuery(Monaca Version)」(2.0.3)が含まれていることを確認。

4. 「PanoramaNews/www」以下一式を www 以下にコピー。

5. js/office365api.js の "clientId:'12345678-1234-1234-1234-123456789012'"を適切な"クライアント ID"へ変更。

6. 実行！

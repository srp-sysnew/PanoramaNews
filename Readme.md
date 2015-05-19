# PanoramaNews
本プログラムは「パノラマ新聞」の[Monaca](https://ja.monaca.io/)用プロジェクトの一部(www以下)を抜粋したものです。

# Requirement
* [Monaca](https://ja.monaca.io/)。ただし Crodova が利用できれば、その限りでないです。
* Windows Azure AD への アプリケーション(NativeClient)登録。

# Usage
1.[Monaca](https://ja.monaca.io/)にて「最小限のテンプレート」をベースに新規プロジェクト作成。

2.「Cordovaプラグインの管理」を開く。  
a. Cordovaバージョン 4.1.0 を選択。  
b.「InAppBrowser」「Splashscreen」「MonacaPlugin」が含まれていることを確認。

3.「JS/CSSコンポーネントの追加と削除」を開く。  
Cordova(PhoneGap)Loader」 (1.0.0)、「Monaca Core Utility」(2.0.4)、「jQuery(Monaca Version)」(2.0.3)が含まれていることを確認。

4.「PanoramaNews/www」以下一式を www 以下にコピー。

5.js/office365api.js の "clientId:'12345678-1234-1234-1234-123456789012'"を適切なものに変更。

6.実行！



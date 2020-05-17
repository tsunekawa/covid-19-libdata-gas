# COVID-19全国図書館調査スクリプト集

saveMLAKが行う[covid-19の影響による図書館の動向調査](https://savemlak.jp/wiki/covid-19-survey)の実施にあたり作成したGoogle Apps Scriptコード集です。

## 構成

このリポジトリは、以下の複数のスクリプトを含みます。

### ワークシート自動分割・結合スクリプト ( ./worksheet 以下 )

調査シートを都道府県別に分割して効率的な分担作業を支援するとともに、調査終了時に分割したシートを統合するためのスクリプトです。

### 参加登録フォーム用スクリプト ( ./registration 以下 )

調査への参加申請フォームに登録した方のGoogleアカウントに、調査シートへの編集権限を自動付与するためのスクリプトです。

## 使用にあたっての要件

本スクリプトの実行には Googleアカウントが必要です。

また、スクリプトの開発、デプロイには以下のミドルウェア、ライブラリが必要です。

* [Node.js](https://nodejs.org/ja/) ver.12.16.12
* [clasp](https://codelabs.developers.google.com/codelabs/clasp/#0) ver.2.3.0

## 開発者
* 常川真央 (Tsunekawa Mao)

## ライセンス
MITライセンスです。
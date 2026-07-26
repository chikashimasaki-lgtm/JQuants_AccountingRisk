# JQuants_AccountingRisk 開発ガイド

## フォント標準化

### 出力形式別フォント指定

- **HTML/Web出力**: **Noto Sans JP**
- **PowerPoint**: **Meiryou UI**
- **Googleドキュメント/スライド/シート**: **Noto Sans JP**

### 実装方法

#### PowerPointの場合
- スライドマスターで「Meiryou UI」をデフォルトフォントに設定
- 全テキストボックスに適用

#### HTMLファイルの場合
```html
<head>
  <style>
    @import url('https://fonts.googleapis.com/css2?family=Noto+Sans+JP:wght@400;500;600;700&display=swap');
    body { font-family: "Noto Sans JP", sans-serif; }
  </style>
</head>
```

#### CSS
```css
/* 全要素にNoto Sans JPを適用 */
* { font-family: "Noto Sans JP", sans-serif; }
```

### 注記
- Google Fontsからインポートして使用
- 複数のウェイト（400, 500, 600, 700）に対応
- フォールバックとしてsans-serifを指定

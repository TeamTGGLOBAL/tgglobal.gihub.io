# tgglobal coding guidelines
## 目次
1. [命名規則](#命名規則)  
1.1 [変数名・定数名](#変数名定数名)  
1.1 [配列や集合の変数名](#配列や集合の変数名)  
1.1 [関数名](#関数名)  
1.1 [グローバルスコープの定数、プロパティストアのキー](#グローバルスコープの定数プロパティストアのキー)  
1. [変数・定数の宣言](#変数名定数名)  
1.1 [const→letの順に優先する](#constletの順に優先する)  
1.1 [スコープはできる限り小さく](#スコープはできる限り小さく)  
1. [プロパティストア](#プロパティストア)  
1. [コードのフォルム](#コードのフォルム)  
1.1 [ステートメント](#ステートメント)  
1.1 [ネストとインデント](#ネストとインデント)  
1.1 [縦に揃える](#縦に揃える)  
1. [コメント](#コメント)  
1.1 [コードを見てわかるコメントは不要](#コードを見てわかるコメントは不要)  
1.1 [ドキュメンテーションコメント](#ドキュメンテーションコメント)  
1.1 [「//」によるコメントを優先](#によるコメントを優先)  
1. [比較演算子](#比較演算子)  
1. [マジックナンバーは禁止](#マジックナンバーは禁止)  
1. [ユーザーの干渉・省略](#ユーザーの干渉省略)  
1.1 [テーブルの走査](#ユーザーの干渉省略)  
1.1 [アクティブシートとシートの保護](#アクティブシートとシートの保護)  
1. [構造データの使用](#構造データの使用)  
1. [スクリプトの高速化](#スクリプトの高速化)  
1.1 [配列の利用](#配列の利用)  
1.1 [シートにレコードを追加する](#シートにレコードを追加する)  
1. [グローバル領域](#グローバル領域)  
1. [クラス](#クラス)  
1. [まとめ](#まとめ)  


# **命名規則**

変数、定数、関数などにつける名前を、**識別子**といいます。

識別子は一定のルールを守っていれば、自由に命名できますが、**スクリプトの読みやすさや開発効率**を考えて以下の点に留意します。

- 中身や役割がわかるような意味のある名前をつける
- 英単語を使い、日本語は使わない

ある程度長い名前でも alt + / で単語補完が使えるので、意味がわからないよりは長い方がマシです。

### **変数名・定数名**

**ブロックスコープ、ローカルスコープの変数名・定数名はキャメル記法**で書きます。

変数の内容が明確かつ宣言から近い距離で完結するのであれば

- url
- body
- message
- thread
- number
- sheet

などの単一単語での変数名も問題ありません。

また、コンテンツアシストでメソッドの引数などで使われている変数名を使用するのは、とても効果的です。

スプレッドシートを代入する変数、定数で以下を使うのは一般的です。

- ss

データ型に合わせて、以下接頭語で始める方法もありますが、それら接頭語を付与せずにデータ型を想定できる単語を使用するのが理想です。

- str：String
- date：Date
- rng：Range
- sheet：sh

状況に応じて以下接頭語も使うことができます。

- row：行
- col：列
- min：最小
- max：最大
- avg：平均

以下については広く一般に使われており、短くてシンプルな名称でも可読性には影響を与えませんので、使っても問題なしです。

- i, j, k：カウント変数
- obj：一時的なオブジェクト変数

**GASは変数とそのデータ型については寛容**で、数値型の値を入れていた変数に、文字列型の値を上書きするといった処理も可能です。

トラブルの元になるので、**その変数がどの型なのかわかるよう**な名称を心がけるとよいでしょう。

### **配列や集合の変数名**

**配列や集合の場合、複数形の変数名**にします。

- values
- sheets
- messages
- threads
- events

### **関数名**

関数名も**アルファベットのキャメル記法**で書きます。

また「～かどうか」を返すBoolean型のFunctionプロシージャに関しては

- is～
- has～
- can～

を使うとわかりやすいですし、if文の条件式にそのまま突っ込めたりしてかっこいいです。

クラス名は、通常の関数名と区別する為に、先頭を大文字で始めるのが一般的です。

### **グローバルスコープの定数、プロパティストアのキー**

グローバルスコープの定数、プロパティストアーのキーは**アルファベットのスネーク記法**で書きます。

```
const TAX_RATE = 1.08;
const USER_ID = 'hogehoge';
```

# **変数・定数の宣言**

### **const→letの順に優先する**

**再代入が必要ないものであれば、constキーワードによる定数**を優先的に使用します。

**再代入が必要であればletによる変数**を使用します。varを使う必要はほぼありません。

### **スコープはできる限り小さく**

原則、**変数・定数はなるべく狭いスコープで**使用するようにしましょう。

優先はブロックスコープ→ローカルスコープ→グローバルスコープの順です。

それにより可読性と安全性が高まります。理由もなくグローバル領域には記述しないように…。

ただし、グローバルで使用する定数の場合は、以下のプロパティストアの利用も考慮しつつ、特定のスクリプトファイルのグローバル領域の先頭にまとめて宣言をしておくのもよいでしょう。後で値を変更するときに、わかりやすい位置にあると便利です。

# **プロパティストア**

GASでは**プロパティストア**に、プロジェクトやドキュメントに紐付く形でデータを格納しておくことができます。

![https://tonari-it.com/wp-content/uploads/010-property-store-680x255.png](https://tonari-it.com/wp-content/uploads/010-property-store-680x255.png)

ファイルのIDや外部と接続するために必要な情報を**スクリプトから分離して安全に管理**することができます。

プロパティストアにはいくつか種類がありますが

- プロジェクトに紐づくデータ：スクリプトプロパティ
- ドキュメントに紐づくデータ：ドキュメントプロパティ
- 実行ユーザーに紐づくデータ：ユーザープロパティ

という使い分けをします。

**[Google Apps Scriptで実行したユーザーごとのスプレッドシートを新規作成する**ユーザーごとのスプレッドシートを作ってそのIDを管理する必要があるなら、ユーザープロパティを使うと便利です。今回は、GASでユーザーごとのスプレッドシートを作成してユーザープロパティで管理する方法です。tonari-it.com2017.12.22](https://tonari-it.com/gas-user-property/)

![https://tonari-it.com/wp-content/uploads/store.jpg](https://tonari-it.com/wp-content/uploads/store.jpg)

[https://www.google.com/s2/favicons?domain=tonari-it.com](https://www.google.com/s2/favicons?domain=tonari-it.com)

# **コードのフォルム**

### **ステートメント**

GASでは自動でステートメントの末尾が判別されますが、判別が常に正しいとは限らないので、**ステートメントの終わりにはセミコロン「;」をつけます**。

また、**特別な理由がない限り1行に1ステートメント**を記述します。

### **ネストとインデント**

ネストであれば必ずその深さの分のインデントを加えてください。理想としては深さは3つまでがいいですね。

```
functionmyFunction1() {
for (let i = 1; i <= 10; i++) {
if (i % 3 === 0 || i % 5 === 0) {
continue;
    }
    console.log(`iの値：${i}`);
  }
}

```

インデントをそろえる場合は、Shift + Tab キー活用すると、自動でそろえてくれます。

**複数の行を選択しても、GASが判断して整形してくれる**ので、大変便利です。

### **縦に揃える**

一行が長いときには、縦に揃えることを意識すると可読性が高まります。

長いメッセージを生成するときや

```
let body = "";
body += "[info]n";
body += values[row][0] + "n[hr]";
body += values[row][1] + "n(";
body += values[row][2] + ")[/info]";

```

オブジェクトや配列のリテラルを記述するときなどに有効です。

```
functionmyFunction() {
const person = {
    name : 'Bob',
    age : 25,
    sex : 'male',
    job : 'Engineer'
  };
  console.log(person.name);
}

```

状況によって、使い分けてください。

# **コメント**

### **コードを見てわかるコメントは不要**

一般的にコードを見ればわかるような内容についてのコメントは不要です。

例えば以下のようなパターン。

```
if (n > 50){ //nが50より大きい場合
  console.log('50より大きい');
}elseif (n < 50){ //nが50より小さい場合
  console.log('50より小さい');
}

```

コメント入れなくてもわかることはコメントしなくてOK！

### **ドキュメンテーションコメント**

関数を作成する際に、**ドキュメンテーションコメント**を設定し、その役割や引数、戻り値を記載するようにしましょう。

また、ドキュメンテーションコメントを入れることで、スプレッドシートで利用するカスタム関数使用時に、補完の候補としたり、詳細情報を表示したりできるので便利です。

```
/**
* 指定した金額の税込価格を返すカスタム関数
*
* @param {number} 金額
* @return {number} 金額
* @Customfunction
*/
functionZEIKOMI(price) {
const TAX_RATE = 0.1;
return price * (1 + TAX_RATE);
}

```

![https://tonari-it.com/wp-content/uploads/5466c15980b0a4b19d5e4e88ebd88b06.png](https://tonari-it.com/wp-content/uploads/5466c15980b0a4b19d5e4e88ebd88b06.png)

### **「//」によるコメントを優先**

Ctrl + / で、現在カーソルのある行や**複数行の選択範囲をコメントイン、コメントアウトできる**ので、「/* ～ */」によるコメントよりも「//」を優先して使用します。

# **比較演算子**

演算子「==」と「!=」は左辺と右辺のデータ型が異なっていても、データ型を変換した上で比較をします。

一方で、演算子「===」と「!==」は**左辺と右辺のデータ型が異なっていることを厳密に判断して比較**します。

```
functionmyFunction() {
  console.log(5 == '5'); //true
  console.log(5 === '5'); //false

  console.log(5 != '5'); //false
  console.log(5 !== '5'); //true
}

```

したがって、データ型も含めて厳密に比較をする**「===」と「!==」を優先して使用**するようにしましょう。

# **マジックナンバーは禁止**

**マジックナンバーは使用禁止**です。メンテナンス性を著しく犠牲にします。マジックナンバー、ダメ！

マジックナンバーになりやすいポイントとしては

- 行数、列数
- 係数
- セルのアドレス
- 配列やオブジェクトの要素数
- 引数
- ファイル名、パス名
- パスワード
- URLやメールアドレス

などがあります。

# **ユーザーの干渉・省略**

### **テーブルの走査**

テーブルの走査をする場合には、getDataRangeメソッド、getValuesメソッドを用いて**データを二次元配列として取得**して操作することを優先します。

```
functionmyFunction() {
const values = SpreadsheetApp.getActiveSheet().getDataRange().getValues();
for(let i = 1; i < values.length; i++){
    //処理
  }
}

```

**[Google Apps Scriptのスプレッドシート読み書きを格段に高速化をする方法**Google Apps Scriptでスプレッドシートの操作をしていて実行速度が遅い！と感じたことがあると思います。今回はスプレッドシートを操作する場合に処理速度を格段に速くする方法をお伝えします。tonari-it.com2016.04.05](https://tonari-it.com/gas-spreadsheet-speedup/)

![https://tonari-it.com/wp-content/uploads/speed-2.jpg](https://tonari-it.com/wp-content/uploads/speed-2.jpg)

[https://www.google.com/s2/favicons?domain=tonari-it.com](https://www.google.com/s2/favicons?domain=tonari-it.com)

**最終行**はgetLastRowメソッドで取得することができますが、配列での処理を優先しましょう。

```
functionmyFunction() {
const sheet = SpreadsheetApp.getActiveSheet();
const lastRow = sheet.getLastRow();
for(let i = 2; i <= lastRow; i++) {
    //処理
  }
}

```

**[【初心者向けGAS】for文を使ったスプレッドシートの繰り返しの超基本**初心者向けGoogle Apps Script超入門、GASプログラミングの基本を学んでいきます。今回は、for文を使った繰り返しの超基本。カウント変数、初期化式、条件式、増加式の意味と使い方です。tonari-it.com2018.03.20](https://tonari-it.com/gas-for/)

![https://tonari-it.com/wp-content/uploads/repeat-1.jpg](https://tonari-it.com/wp-content/uploads/repeat-1.jpg)

[https://www.google.com/s2/favicons?domain=tonari-it.com](https://www.google.com/s2/favicons?domain=tonari-it.com)

### **アクティブシートとシートの保護**

シートが一つであれば、**getActiveSheetメソッド**でシートを指定するようにします。

ユーザーによるシート名の変更にも対応でき、高速化の面でも有利です。

一つのシートで運用できるよう、工夫しましょう。

また、複数シートが存在する場合は、**getSheetByNameメソッド**などを使用する必要があります。

この際、ユーザーがシート名を変更しないように運用する必要があります。

絶対に保護したい場合には、対象のシートを**非表示にしたり、シートを保護**することで、干渉を受けないようにすることができます。

**[【初心者向けGAS】スプレッドシートのシートを取得する２つの方法**初心者向けのGoogle Apps Script入門シリーズとして、GASプログラミングの基礎をお伝えしています。今回は、スプレッドシートからシートを取得する２つの方法をお伝えします。tonari-it.com2018.03.18](https://tonari-it.com/gas-spreadsheet-get-sheet/)

![https://tonari-it.com/wp-content/uploads/sheets-2.jpg](https://tonari-it.com/wp-content/uploads/sheets-2.jpg)

[https://www.google.com/s2/favicons?domain=tonari-it.com](https://www.google.com/s2/favicons?domain=tonari-it.com)

# **構造データの使用**

スクリプトで主に操作をするスプレッドシートのデータは、特に理由がない限り**テーブル形式**になるように心がけましょう。

- A1セルから構成する
- 空白行・空白列を設けない
- セルの結合を使わない
- 見出しは1行で構成する
- 同じ種類のデータはシートを分けない

つまり、ピボットテーブルを作成できるデータの整理の仕方が望ましいです。

データを構造データで持つかどうかが、それを操作するプログラムの作りやすさに大きく影響します。

A1セルからデータが存在する最終行と最終列までの範囲を取得する**getDataRangeメソッド**や、シートに行を追加する**appendRowメソッド**を使うことができます。

# **スクリプトの高速化**

GASではスクリプトの**実行時間に6分という厳しい制限**が設けられています。

APIのアクセス回数を減らして、処理速度を少しでもあげましょう。

### **配列の利用**

スプレッドシートのデータやGmailのメッセージをはじめ、GASで取り扱うデータは、**直接操作をするよりも、配列に格納して処理する方が、格段に処理速度を向上**させることができます。

**[Google Apps Scriptのスプレッドシート読み書きを格段に高速化をする方法**Google Apps Scriptでスプレッドシートの操作をしていて実行速度が遅い！と感じたことがあると思います。今回はスプレッドシートを操作する場合に処理速度を格段に速くする方法をお伝えします。tonari-it.com2016.04.05](https://tonari-it.com/gas-spreadsheet-speedup/)

![https://tonari-it.com/wp-content/uploads/speed-2.jpg](https://tonari-it.com/wp-content/uploads/speed-2.jpg)

[https://www.google.com/s2/favicons?domain=tonari-it.com](https://www.google.com/s2/favicons?domain=tonari-it.com)

スプレッドシートで複数のセルを取り扱う場合は**getValuesメソッド**、**setValuesメソッド**を使用して、配列を利用します。

```
functionmyFunction() {
const sheet = SpreadsheetApp.getActiveSheet();
const values = sheet.getDataRange().getValues();

  //配列valuesに対する処理

  sheet.getRange(1, 1, values.length, values[0].length).setValues(values);
}

```

これはかなり重要なテクニックなので、必ずマスターしたいところです。

### **シートにレコードを追加する**

シートにレコードを追加する場合は、**appendRowメソッドを使うことで配列から直接レコードを追加**することができます。

```
functionmyFunction(){
const sheet = SpreadsheetApp.getActiveSheet();
  sheet.appendRow(配列);
}

```

**[Google Apps Scriptで配列を使ってスプレッドシートにデータ行を追加する方法**Google Apps Scriptでは配列操作が非常に重要です。実行速度が6分を超えてエラーとしないテクニックとして、スプレッドシートへのレコード追加を配列へのpushメソッドで処理する方法をお伝えします。tonari-it.com2017.02.27](https://tonari-it.com/gas-array-push-append/)

![https://tonari-it.com/wp-content/uploads/push-2.jpg](https://tonari-it.com/wp-content/uploads/push-2.jpg)

[https://www.google.com/s2/favicons?domain=tonari-it.com](https://www.google.com/s2/favicons?domain=tonari-it.com)

範囲の指定が不要で、Sheetオブジェクトに直接実行できます。

# **グローバル領域**

どの関数にも属さない領域を**グローバル領域**といいます。

![https://tonari-it.com/wp-content/uploads/020-global.png](https://tonari-it.com/wp-content/uploads/020-global.png)

グローバル領域へのステートメント記述は、**実行順がわかりづらく**なります。

特に理由がない限りは記述しないようにしつつ、記述する場合も特定のgsファイルの一番上にまとめて記述するほうがいいでしょう。

また、グローバル領域で宣言した（**グローバル変数**）は全ての関数からアクセスすることができ、変更を加えることもできます。

望まない影響を受ける可能性があるので、極力使わないようにしましょう。

# **クラス**

# **まとめ**
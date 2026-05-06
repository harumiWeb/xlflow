# jp_weather_info

`jp_weather_info` は、xlflow で管理する Excel VBA サンプルです。  
日本の天気予報を取得して、最初のワークシートに見やすく整形して表示します。

## できること

- Tsukumijima Weather API から天気予報を取得
- 既定の都市コード `130010`（東京）の予報を表示
- 3日分の予報、天気アイコン、降水確率、最低/最高気温を表示
- xlflow からは `--diagnostic` で失敗箇所を確認可能

## 実行方法

1. このフォルダで `xlflow run --diagnostic` を実行して `build/Book.xlsm` を生成または更新します。
2. `build/Book.xlsm` を開きます。
3. `Main.Run` を実行します。
4. 必要に応じて `Ui.RunFromButton` をボタンに割り当てて使えます。

CLI から再実行する場合も、このフォルダで `xlflow run --diagnostic` を使います。

## 変更しやすい場所

- `src/modules/WeatherApi.bas`
  - `DEFAULT_CITY_CODE_VALUE` を変えると、表示する都市を切り替えられます。
- `src/modules/WeatherSheet.bas`
  - シートのレイアウトや表示項目を調整できます。

## 補足

- API キーは不要です。
- 表示内容はネットワーク接続に依存します。

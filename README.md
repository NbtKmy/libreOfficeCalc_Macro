# Macro for　librarian 
この2つのマクロでZoteroからひっぱってきたCSVデータを書籍注文用の表形式に自動的に変換します。  

プロセスとしては  
1. Zoteroの中で複本チェック（[Zotero Swissbib UZH-AOI Locations](https://github.com/NbtKmy/zotero-swissbib-bb-locations)）と市販されている書籍の値段をまとめて調べる（[zotero-price-checker](https://github.com/NbtKmy/zotero-price-checker)）。
1. そうしてできたデータをまとめてCSVとしてとりだす。(BOMなし、UTF-8)
1. そのCSVデータをLibreOffice Calcで開く。2つのマクロはあらかじめ入れておく。
1. そのあとまず、`ZoteroCSVtoOrderTable.bas`で必要な情報を整頓して、`ShimashimaFormat.bas`で見栄えを整える感じ。

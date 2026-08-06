# _archive

淘汰但**不真刪**的檔案放這裡。

## 目前內容：20 個 `*_excel.csv`

這批是 `50_資料庫/` 與 `80_References/` 中同名 CSV 的**重複檔**。
逐檔比對過，去掉 BOM 後**內容完全相同**，差別只是 `_excel` 版多了一個 UTF-8 BOM
（讓 Excel 直接開啟時中文不亂碼）。

**正本一律以無 `_excel` 後綴的檔案為準。**

如果你之後還是需要用 Excel 直接雙擊開檔，不必留兩份——用這行從正本現產：

```bash
printf '\xef\xbb\xbf' > out_excel.csv && cat 50_資料庫/portfolio.csv >> out_excel.csv
```

## 規則

- 這層的檔案**不會被 AI 當成有效資料源**，查數字不要讀這裡
- 確定不再需要時才手動刪除

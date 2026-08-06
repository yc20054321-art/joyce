# _private

私人內容。**AI 不要讀取這層、不要寫進對話、不要出現在 commit 訊息裡。**

## 目前內容：20 張私人照片（約 20MB）

春酒、共識營、人物合照等，與投資資料無關，原本散在 repo 根目錄。

## ⚠️ 這些照片目前仍在 git 追蹤中

已經加進 `.gitignore`，所以**之後新放進來的檔案不會被提交**；
但這 20 張是既有的追蹤檔案，gitignore 對它們無效——它們仍會跟著 repo 被分享出去。

想讓它們停止被追蹤（檔案保留在本機，但從 repo 移除）：

```bash
git rm -r --cached _private
git commit -m "Stop tracking private photos"
```

注意：這只影響**之後**的版本。已經推上去的歷史 commit 裡仍留有這些照片，
要徹底清除得改寫 git 歷史（`git filter-repo` 之類），那是另一件事。

## 分享前的紅線

整包 vault 不要直接給別人或別人的 AI 讀。除了這層照片，
`50_資料庫/friend_*.csv` 也含**他人的完整持股與金額**，同屬 private。

# 像素地城 UI 更新

將 pixel-dungeon-ui.zip 解壓縮至原本的 my-game 專案根目錄，合併 assets 資料夾。請先保留原本 index.html 備份。壓縮檔是修改檔，不是完整遊戲；其餘圖片、Excel 與程式仍使用原專案內容。

## 內容
- index.html：共用角色資訊列、戰鬥分區與紀錄區排版。
- assets/css/ui-theme.css：共用介面樣式。
- assets/css/pixel-dungeon.css：黑鐵古銅卡框、陣營標籤、血魔條、地城背景與導覽樣式。
- assets/images/ui/dungeon-hall.png：新增地城背景。

保留原遊戲角色、武器與七個功能分頁。沒有加入參考圖中尚未實作的信箱、背包快捷鍵或設定頁。

驗證：HTML 內嵌 JavaScript 語法、關鍵元素 ID 唯一性、戰鬥實際運行、瀏覽器錯誤紀錄及預覽畫面邊界。尚未完成所有關卡、解鎖頁面及手動技能流程的完整回歸測試。

## 背景生成紀錄
使用內建 image_gen 工具生成，未使用 CLI。

Prompt:
Create a standalone game background asset, landscape 3:2. Dark fantasy detailed pixel art dungeon throne hall, symmetrical frontal view, black stone gothic pillars on left and right, hanging iron chains and faded burgundy banners, small amber wall torches, circular carved stone arena floor in foreground, distant central stairs. Crisp deliberate pixel clusters in classic high quality 16-bit RPG style. Very dark charcoal and brown palette, restrained bronze and orange torch highlights. Central floor empty and readable. No characters, monsters, weapons, text, letters, UI, frames, cards or logos. This will sit behind an HTML battle interface, not a mockup.

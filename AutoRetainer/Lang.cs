using AutoRetainerAPI.Configuration;
using Dalamud.Utility;
using ECommons.ExcelServices.Sheets;
using Lumina.Excel.Sheets;
using System.Collections.ObjectModel;
using System.Text.RegularExpressions;
using System.Windows.Forms.VisualStyles;

namespace AutoRetainer;

internal static class Lang
{
    private static readonly Dictionary<string, string> Tw = new(StringComparer.Ordinal)
    {
        ["General"] = "一般",
        ["Delays"] = "延遲",
        ["Operation"] = "運作",
        ["Settings"] = "設定",
        ["Behavior"] = "行為",
        ["Integrations"] = "整合",
        ["Utility"] = "工具",
        ["Import/Export"] = "匯入/匯出",
        ["User Interface"] = "介面",
        ["Retainers"] = "僱員",
        ["Deployables"] = "飛空艇/潛水艇",
        ["Troubleshooting"] = "疑難排解",
        ["Statistics"] = "統計",
        ["Record Venture Statistics"] = "記錄僱員探索統計",
        ["About"] = "關於",
        ["Disclaimer"] = "使用提醒",
        ["Common Settings"] = "共通設定",
        ["General Settings"] = "一般設定",
        ["Inventory Cleanup/General Settings"] = "背包清理/一般設定",
        ["Inventory Cleanup/Unconditional Sell List"] = "背包清理/無條件出售清單",
        ["Inventory Cleanup/Quick Venture Sell List"] = "背包清理/快速探索出售清單",
        ["Inventory Cleanup/Discard List"] = "背包清理/丟棄清單",
        ["Inventory Cleanup/Protection List"] = "背包清理/保護清單",
        ["Inventory Cleanup/Fast Addition and Removal"] = "背包清理/快速新增與移除",
        ["Inventory Cleanup/Character Configuration"] = "背包清理/角色設定",
        ["Grand Company Delivery/General Settings"] = "軍票繳納/一般設定",
        ["Grand Company Delivery/Exchange Lists"] = "軍票繳納/兌換清單",
        ["Grand Company Delivery/Character Configuration"] = "軍票繳納/角色設定",
        ["Game startup"] = "遊戲啟動",
        ["Inventory warnings"] = "庫存警告",
        ["Teleportation"] = "傳送",
        ["Bailout Module"] = "斷線恢復模組",
        ["Mass configuration change"] = "批量設定變更",
        ["Plans"] = "計畫",
        ["Alert Settings"] = "提醒設定",
        ["Registration, component and plan automation"] = "登錄、零件與計畫自動化",
        ["Export character and submarine list to CSV"] = "匯出角色與潛水艇清單到 CSV",
        ["General Character Specific Settings"] = "角色個別設定",
        ["Teleport overrides"] = "傳送覆寫",
        ["Character Data Expunge/Reset"] = "角色資料清除/重置",
        ["Character sorting in Retainer tab"] = "僱員分頁角色排序",
        ["Character sorting in Deployables tab"] = "飛空艇/潛水艇分頁角色排序",
        ["Server Time"] = "伺服器時間",
        ["Lock window position and size"] = "鎖定視窗位置與大小",
        ["Open settings window"] = "開啟設定視窗",
        ["AutoRetainer Configuration"] = "AutoRetainer 設定",
        ["Filter..."] = "篩選...",
        ["Keybinds"] = "快捷鍵",
        ["Access summoning bell/workshop panel keybinds"] = "傳喚鈴/工房面板快捷鍵",
        ["Temporarily prevents AutoRetainer from being automatically enabled when using a Summoning Bell/Workshop Panel"] = "使用傳喚鈴/工房面板時，暫時阻止 AutoRetainer 自動啟用",
        ["Temporarily set the Collect Operation mode, preventing ventures from being assigned for the current cycle/Temporarily set Deployables mode to Finalize only"] = "暫時切換為只收取模式，避免本輪重新派出僱員；飛空艇/潛水艇則暫時只結算",
        ["Quick Retainer Action"] = "僱員快速操作",
        ["Sell Item"] = "出售物品",
        ["Entrust Item"] = "託管物品",
        ["Retrieve Item"] = "取回物品",
        ["Put up For Sale"] = "上架販售",
        ["None"] = "無",
        ["+ right click"] = "+ 右鍵",
        ["Now press new key..."] = "現在按下新的按鍵...",
        ["Auto-detect new key"] = "自動偵測新按鍵",
        ["Select key manually:"] = "手動選擇按鍵：",
        ["Exclusions and Order"] = "排除與順序",
        ["FPS Limiter"] = "FPS 限制",
        ["Region Lock"] = "區域鎖定",
        ["Contingency"] = "應急處理",
        ["Inventory Management"] = "背包管理",
        ["Entrust Manager"] = "託管管理",
        ["Inventory Cleanup"] = "背包清理",
        ["Unconditional Sell List"] = "無條件出售清單",
        ["Quick Venture Sell List"] = "快速探索出售清單",
        ["Discard List"] = "丟棄清單",
        ["Protection List"] = "保護清單",
        ["Fast Addition and Removal"] = "快速新增與移除",
        ["Character Configuration"] = "角色設定",
        ["Grand Company Delivery"] = "軍票繳納",
        ["Exchange Lists"] = "兌換清單",
        ["Login Overlay"] = "登入浮層",
        ["Display Login Overlay"] = "顯示登入浮層",
        ["Login overlay scale multiplier"] = "登入浮層縮放倍率",
        ["Login overlay button padding"] = "登入浮層按鈕間距",
        ["Display hidden characters when searching"] = "搜尋時顯示隱藏角色",
        ["Number of columns"] = "欄數",
        ["Overlay height, %"] = "浮層高度，%",
        ["Miscellaneous"] = "其他",
        ["Experiments"] = "實驗功能",
        ["Night Mode"] = "夜間模式",
        ["Notifications"] = "通知",
        ["Advanced"] = "進階",
        ["Log"] = "記錄",
        ["Expert Settings"] = "專家設定",
        ["Character Synchronization"] = "角色同步",
        ["Selected"] = "已選擇",

        ["Time Desynchronization Compensation"] = "時間誤差補償",
        ["Additional amount of seconds that will be subtracted from venture ending time to help mitigate possible issues of time desynchronization between the game and your PC."] = "會從探索結束時間額外扣除的秒數，用來降低遊戲與電腦時間不同步造成的問題。",
        ["Additional Interaction Delay, frames"] = "額外交互延遲，影格",
        ["The lower this value is the faster plugin will use actions. When dealing with low FPS or high latency you may want to increase this value. If you want the plugin to operate faster you may decrease it."] = "數值越低，插件動作越快。若 FPS 低或延遲高，建議調高；若想更快可調低。",
        ["Extra Logging"] = "詳細記錄",
        ["This option enables excessive logging for debugging purposes. It will spam your log and cause performance issues while enabled. This option will disable itself upon plugin reload or game restart."] = "啟用除錯用詳細記錄。會大量寫入 log 並可能影響效能；重新載入插件或重開遊戲後會自動關閉。",
        ["Assign + Reassign"] = "派出 + 重新派出",
        ["Automatically assigns enabled retainers to a Quick Venture if they have none already in progress and reassigns current venture."] = "若已啟用的僱員沒有進行中的探索，會自動派出快速探索，並重新派出目前探索。",
        ["Collect"] = "只收取",
        ["Only collect venture rewards from the retainer, and will not reassign them.\nHold CTRL when interacting with the Summoning Bell to apply this mode temporarily."] = "只收取僱員探索獎勵，不會重新派出。\n與傳喚鈴互動時按住 CTRL 可暫時套用此模式。",
        ["Reassign"] = "重新派出",
        ["Only reassign ventures that retainers are undertaking."] = "只重新派出僱員目前正在進行的探索。",
        ["RetainerSense"] = "僱員鈴感應",
        ["AutoRetainer will automatically enable itself when the player is within interaction range of a Summoning Bell. You must remain stationary or the activation will be cancelled."] = "當角色在傳喚鈴可互動範圍內時，自動啟用 AutoRetainer。啟動時必須保持不動，否則會取消。",
        ["Activation Time"] = "啟動時間",

        ["Anonymise Retainers"] = "匿名化僱員",
        ["Retainer names will be redacted from general UI elements. They will not be hidden in debug menus and plugin logs however. While this option is on, character and retainer numbers are not guaranteed to be equal in different sections of a plugin (for example, retainer 1 in retainers view is not guaranteed to be the same retainer as in statistics view)."] = "僱員名稱會在一般介面元素中隱藏。不過除錯選單與插件記錄不會隱藏。啟用此選項時，不同頁面中的角色/僱員編號不保證相同，例如僱員頁面的「僱員 1」不一定等於統計頁面的同一名僱員。",
        ["Display Quick Menu in Retainer UI"] = "在僱員介面顯示快速選單",
        ["Display Extended Retainer Info"] = "顯示額外僱員資訊",
        ["Displays retainer item level/gathering/perception and the name of their current venture in the main UI."] = "在主介面顯示僱員裝等/獲得力/鑑別力與目前探索名稱。",
        ["Do not close AutoRetainer windows on ESC key press"] = "按 ESC 時不要關閉 AutoRetainer 視窗",
        ["Display only most significant icon in status bar"] = "狀態列只顯示最重要的圖示",
        ["Status bar icon size"] = "狀態列圖示大小",
        ["Open AutoRetainer window on game start"] = "遊戲啟動時打開 AutoRetainer 視窗",
        ["Enable title screen button (requires plugin restart)"] = "啟用標題畫面按鈕（需要重啟插件）",
        ["Hide character search"] = "隱藏角色搜尋",
        ["Don't flash background of characters that are complete"] = "已完成角色不要閃爍背景",
        ["Do not warn about second game instance running from same directory"] = "不要警告同目錄啟動第二個遊戲實例",
        ["This will automatically skip AutoRetainer's loading on second instance of the game and you will have no way of loading it until you disable this option in primary instance"] = "這會讓第二個遊戲實例自動跳過載入 AutoRetainer；在主要實例關閉此選項前，第二個實例無法再載入它。",
        ["Enable"] = "啟用",
        ["This is purely visual order and does not affects character processing in any way."] = "這只影響顯示順序，不會影響角色處理流程。",
        ["Add Entries..."] = "新增項目...",
        ["Data Center"] = "資料中心",
        ["DataCenter"] = "資料中心",
        ["Inventory Slots"] = "背包空格",
        ["Inventory_Slots"] = "背包空格",
        ["Repair Kits"] = "修理材料",
        ["Repair_Kits"] = "修理材料",
        ["Ventures"] = "探險幣",
        ["Ceruleum"] = "青燐水",
        ["World"] = "伺服器",
        ["Name"] = "名稱",
        ["Region JP"] = "區域 JP",
        ["Region NA"] = "區域 NA",
        ["Region EU"] = "區域 EU",
        ["Region OC"] = "區域 OC",
        ["Region_JP"] = "區域 JP",
        ["Region_NA"] = "區域 NA",
        ["Region_EU"] = "區域 EU",
        ["Region_OC"] = "區域 OC",

        ["Wait on login screen"] = "在登入畫面等待",
        ["Disable Multi Mode on Manual Login"] = "手動登入時停用多角色模式",
        ["Do not reset Preferred Character on Manual Login"] = "手動登入時不要重置偏好角色",
        ["Allow entering shared houses"] = "允許進入共享房屋",
        ["Attempt to enter house on login even when Multi Mode is disabled"] = "即使多角色模式關閉，登入後也嘗試進入房屋",
        ["Do not teleport or enter house for retainers when already next to bell"] = "已在傳喚鈴旁時，不為僱員傳送或進屋",
        ["Enable Multi Mode on Game Boot"] = "遊戲啟動時啟用多角色模式",
        ["Enable Multi Mode on Plugin Startup"] = "插件啟動時啟用多角色模式",
        ["Delay, seconds"] = "延遲，秒",
        ["Auto-login on Game Boot"] = "遊戲啟動時自動登入",
        ["Disabled"] = "停用",
        ["Hide Armoury Chest Items"] = "隱藏兵裝庫物品",
        ["Hide Gear Set Items"] = "隱藏套裝配置物品",
        ["Show All Items"] = "顯示所有物品",
        ["Last logged in character"] = "上次登入的角色",
        ["Delay"] = "延遲",
        ["Set appropriate delay to let plugins fully load before logging in and to allow yourself some time to cancel login if needed"] = "設定適當延遲，讓插件先載入完成，也保留取消登入的時間。",
        ["Retainer list: remaining inventory slots warning"] = "僱員清單：剩餘背包格警告",
        ["Retainer list: remaining ventures warning"] = "僱員清單：剩餘探險幣警告",
        ["Deployables list: remaining inventory slots warning"] = "飛空艇/潛水艇清單：剩餘背包格警告",
        ["Deployables list: remaining fuel warning"] = "飛空艇/潛水艇清單：剩餘燃料警告",
        ["Deployables list: remaining repair kit warning"] = "飛空艇/潛水艇清單：剩餘修理材料警告",
        ["Lifestream plugin is required"] = "需要 Lifestream 插件",
        ["You must register houses in Lifestream plugin for every character you want this option to work or enable Simple Teleport."] = "若要使用此選項，必須在 Lifestream 為每個角色登錄房屋，或啟用簡易傳送。",
        ["You can customize these settings per character in character configuration menu."] = "可在角色設定選單中為每個角色個別調整。",
        ["For current character teleport options are customized."] = "目前角色已使用個別傳送設定。",
        ["Enabled"] = "已啟用",
        ["Teleport for retainers..."] = "為僱員傳送...",
        ["...to private house"] = "...到個人房屋",
        ["...to shared estate"] = "...到共享房屋",
        ["...to free company house"] = "...到部隊房屋",
        ["...to apartment"] = "...到公寓",
        ["If all above are disabled or fail, will be teleported to inn."] = "如果上述項目都停用或失敗，將傳送到旅館。",
        ["Teleport to free company house for deployables"] = "為飛空艇/潛水艇傳送到部隊房屋",
        ["Enable Simple Teleport"] = "啟用簡易傳送",
        ["Auto-close and retry logging in on connection errors"] = "連線錯誤時自動關閉並重試登入",
        ["Upon disconnecting, AutoRetainer will attempt to log back in. If the session has expired, no login attempt will be made."] = "斷線後 AutoRetainer 會嘗試重新登入。若登入工作階段已過期，則不會嘗試登入。",

        ["Resend vessels when accessing the Voyage Control Panel"] = "開啟航行控制面板時重新派出船隻",
        ["Finalize all vessels before resending them"] = "重新派出前先結算所有船隻",
        ["Hide Airships from Deployables UI"] = "在飛空艇/潛水艇介面隱藏飛空艇",
        ["Less than possible vessels enabled"] = "啟用的船隻少於可用數量",
        ["Enabled vessel isn't deployed"] = "已啟用船隻尚未派出",
        ["Export only characters enabled for multi mode (otherwise - all)"] = "只匯出已啟用多角色模式的角色（否則匯出全部）",
        ["Export only enabled submarines (otherwise - all)"] = "只匯出已啟用的潛水艇（否則匯出全部）",
        ["Export"] = "匯出",
        ["Deselect All"] = "取消全選",
        ["Select All"] = "全選",
        ["Add vessels by level to selection"] = "依等級加入船隻到選取",
        ["Set behavior"] = "設定行為",
        ["Set unlock mode"] = "設定解鎖模式",
        ["Set unlock plan"] = "設定解鎖計畫",
        ["Set point plan"] = "設定航點計畫",
        ["Enable selected submersibles"] = "啟用已選潛水艇",
        ["Disable selected submersibles"] = "停用已選潛水艇",
        ["Enable deployables multi mode for owners of selected submersibles"] = "為已選潛水艇的持有角色啟用多角色飛空艇/潛水艇模式",
        ["Disable deployables multi mode for owners of selected submersibles"] = "為已選潛水艇的持有角色停用多角色飛空艇/潛水艇模式",
        ["Enable automatic sub registration"] = "啟用自動潛水艇登錄",
        ["Enable automatic components and plan change"] = "啟用自動零件與計畫變更",
        ["Different setup for first Submersible"] = "第一艘潛水艇使用不同設定",
        ["Delete"] = "刪除",
        ["Add"] = "新增",

        ["Select retainers:"] = "選擇僱員：",
        ["Character search"] = "角色搜尋",
        ["By level:"] = "依等級：",
        ["Actions:"] = "動作：",
        ["None selected"] = "尚未選擇",
        ["Add retainers by level to selection"] = "依等級加入僱員到選取",
        ["Enable planner with venture plan"] = "使用探索計畫啟用規劃器",
        ["Set entrust plan"] = "設定託管計畫",
        ["Remove entrust plan from selected retainers"] = "從已選僱員移除託管計畫",
        ["Disable venture planner for selected retainers"] = "停用已選僱員的探索規劃器",
        ["Enable selected retainers"] = "啟用已選僱員",
        ["Disable selected retainers"] = "停用已選僱員",
        ["Enable retainer multi mode for owners of selected retainers"] = "為已選僱員的持有角色啟用多角色僱員模式",
        ["Disable retainer multi mode for owners of selected retainers"] = "為已選僱員的持有角色停用多角色僱員模式",
        ["New plan"] = "新增計畫",
        ["Reset"] = "重置",
        ["Set this plan as default"] = "設為預設計畫",
        ["ALL submersibles"] = "所有潛水艇",
        ["Current character's submersibles"] = "目前角色的潛水艇",
        ["No submersibles"] = "不套用潛水艇",
        ["Copy plan settings"] = "複製計畫設定",
        ["Paste plan settings"] = "貼上計畫設定",
        ["Help"] = "說明",
        ["Open editor"] = "開啟編輯器",
        ["Clear plan"] = "清除計畫",
        ["Open Submarine Unlock Plan Editor"] = "開啟潛水艇解鎖計畫編輯器",
        ["Submersible Voyage Unlockable Planner"] = "潛水艇航行解鎖規劃器",
        ["No or unknown plan selected"] = "未選擇計畫或計畫未知",
        ["No or unknown plan is selected"] = "未選擇計畫或計畫未知",
        ["This plan is set as default."] = "此計畫已設為預設。",
        ["Apply this plan to:"] = "套用此計畫到：",
        ["Delete this plan"] = "刪除此計畫",
        ["Unlock submarine slots. Current slots:"] = "解鎖潛水艇欄位。目前欄位：",
        ["Enforce Spam one destination mode in Deep sea site."] = "在深海站點強制使用重複單一目的地模式。",
        ["Set this plan as enforced."] = "將此計畫設為強制執行。",
        ["Any point selected for unlock in this map will be executed by every single eligible submarine until everything is actually unlocked"] = "此地圖中選為解鎖的航點，會由每一艘符合條件的潛水艇持續執行，直到全部實際解鎖為止。",
        ["Zone"] = "區域",
        ["Map"] = "地圖",
        ["Unlocked by"] = "由此解鎖",
        ["Display current point exploration order"] = "顯示目前航點探索順序",
        ["Open Voyage Route Planner"] = "開啟航線規劃器",
        ["Open Voyage Unlockable Planner"] = "開啟解鎖航線規劃器",
        ["Access submarine list to retrieve data."] = "請開啟潛水艇清單以取得資料。",
        ["Unlocking slots is always prioritized over unlocking routes."] = "解鎖欄位永遠優先於解鎖航線。",
        ["This plan is not used by any submersibles."] = "此計畫未被任何潛水艇使用。",

        ["Stay in retainer menu if there are retainers to finish ventures within 5 minutes or less"] = "若有僱員會在 5 分鐘內完成探索，留在僱員選單",
        ["Action on accessing retainer bell if no ventures available:"] = "開啟傳喚鈴且沒有可處理探索時：",
        ["Action on accessing retainer bell if any ventures available:"] = "開啟傳喚鈴且有可處理探索時：",
        ["Task completion behavior after accessing bell:"] = "開啟傳喚鈴後的任務完成行為：",
        ["Task completion behavior after manual enabling:"] = "手動啟用後的任務完成行為：",
        ["Task completion behavior during plugin operation:"] = "插件運作期間的任務完成行為：",
        ["\"Close retainer list and disable plugin\" option for 3 previous settings is enforced during MultiMode operation."] = "多角色模式運作時，上面 3 個設定會強制使用「關閉僱員清單並停用插件」。",
        ["This option is enforced during MultiMode operation."] = "多角色模式運作時會強制套用此選項。",
        ["Auto-disable plugin when closing retainer list"] = "關閉僱員清單時自動停用插件",
        ["Only applies when you exit menu by yourself. Otherwise, settings above apply."] = "只在你自行離開選單時套用；其他情況使用上方設定。",
        ["Do not show plugin status icons"] = "不要顯示插件狀態圖示",
        ["Display multi mode type selector"] = "顯示多角色模式類型選擇器",
        ["Display deployables checkbox in workshop"] = "在工房顯示飛空艇/潛水艇勾選框",
        ["Enable bailout module"] = "啟用斷線恢復模組",
        ["Timeout before AutoRetainer will attempt to unstuck, seconds"] = "AutoRetainer 嘗試解除卡住前的等待時間，秒",
        ["Disable sorting and collapsing/expanding"] = "停用排序與收合/展開",
        ["Show MultiMode checkbox on plugin UI bar"] = "在插件主列顯示多角色模式勾選框",
        ["Retainer menu delay, seconds"] = "僱員選單延遲，秒",
        ["Allow venture timer to display negative values"] = "允許探索計時器顯示負值",
        ["Do not error check venture planner"] = "不要檢查探索規劃器錯誤",
        ["Upon activating Multi Mode, attempt to enter nearby house"] = "啟用多角色模式時，嘗試進入附近房屋",
        ["Enable Manual relogs character postprocess"] = "允許手動重新登入後處理",
        ["Allow manual command invocation while AutoRetainer locked in postprocess. "] = "AutoRetainer 鎖在後處理時，允許手動呼叫指令。",
        ["Market Cooldown Overlay"] = "市場冷卻浮層",
        ["Artisan integration"] = "Artisan 整合",
        ["Automatically enables AutoRetainer while Artisan is Pauses Artisan operation when ventures are ready to be collected and a retainer bell is within range. Once ventures have been dealt with Artisan will be enabled and resume whatever it was doing."] = "僱員探索可收取且傳喚鈴在範圍內時，會暫停 Artisan 並自動啟用 AutoRetainer。僱員處理完成後會重新啟用 Artisan，讓它恢復原本工作。",
        ["Use server time instead of PC time"] = "使用伺服器時間而不是電腦時間",
        ["Cleanup ghost retainers"] = "清理幽靈僱員",
        ["Export without character data"] = "匯出但不含角色資料",
        ["Import and merge with character data"] = "匯入並合併角色資料",
        ["Do nothing"] = "不執行動作",
        ["Enable AutoRetainer"] = "啟用 AutoRetainer",
        ["Disable AutoRetainer"] = "停用 AutoRetainer",
        ["Pause AutoRetainer"] = "暫停 AutoRetainer",
        ["Close retainer list and disable plugin"] = "關閉僱員清單並停用插件",
        ["Close retainer list and keep plugin enabled"] = "關閉僱員清單並保持插件啟用",
        ["Stay in retainer list and disable plugin"] = "留在僱員清單並停用插件",
        ["Stay in retainer list and keep plugin enabled"] = "留在僱員清單並保持插件啟用",

        ["Pick max amount of destinations"] = "選擇最多目的地",
        ["Spam one destination"] = "重複同一目的地",
        ["Include one unlock destination while levelling"] = "練等時包含一個解鎖目的地",
        ["Preferred Character"] = "偏好角色",
        ["Clear Free company data"] = "清除部隊資料",
        ["Wait For Voyage Completion"] = "等待航行完成",
        ["Options marked with this marker will use values from global configuration"] = "有此標記的選項會使用全域設定值",
        ["Withdraw/Deposit Gil"] = "提領/存入金幣",
        ["Withdraw"] = "提領",
        ["Deposit"] = "存入",
        ["Fake ready"] = "測試：設為完成",
        ["Fake unready"] = "測試：設為未完成",
        ["Enable planner"] = "啟用規劃器",
        ["Quick Exploration"] = "快速探索",
        ["Reload"] = "重新載入",
        ["Show HQ and non-HQ together"] = "HQ 與非 HQ 合併顯示",
        ["Only display character total"] = "只顯示角色總計",
        ["Update"] = "更新",
        ["Update every 30 hours"] = "每 30 小時更新",
        ["Show only wallet FC"] = "只顯示錢包部隊",
        ["Tray notification upon handin completion (requires NotificationMaster)"] = "繳納完成時跳出通知（需要 NotificationMaster）",
        ["Automatic Grand Company Expert Delivery"] = "自動軍票稀有品繳納",
        ["Performance"] = "效能",
        ["Remove minimized FPS restrictions while plugin is operating"] = "插件運作時解除最小化 FPS 限制",
        ["- Also remove general FPS restriction"] = "- 同時解除一般 FPS 限制",
        ["- Also pause ChillFrames plugin"] = "- 同時暫停 ChillFrames 插件",
        ["Raise FFXIV process priority while plugin is operating"] = "插件運作時提高 FFXIV 程序優先權",
        ["May result other programs slowdown"] = "可能導致其他程式變慢",
        ["Close this window without loading AutoRetainer"] = "關閉此視窗且不載入 AutoRetainer",
        ["Learn how to properly run 2 or more game instances"] = "了解如何正確執行兩個以上遊戲實例",
        ["I agree that I may lose all AutoRetainer data"] = "我了解可能會失去所有 AutoRetainer 資料",
        ["Load AutoRetainer"] = "載入 AutoRetainer",
        ["Hide this overlay"] = "隱藏此浮層",
        ["Cancel"] = "取消",
        ["Paused"] = "已暫停",
        ["Multi"] = "多角色",
        ["Night"] = "夜間",
        ["Reset counters"] = "重置計數",
        ["Plugin operation is suppressed by other plugin."] = "插件操作目前被其他插件暫停。",
        ["Multi Mode"] = "多角色模式",
        ["Multi Mode/Retainers"] = "多角色模式/僱員",
        ["Multi Mode/Common Settings"] = "多角色模式/共通設定",
        ["Multi Mode/Deployables"] = "多角色模式/飛空艇/潛水艇",
        ["Multi Mode/FPS Limiter"] = "多角色模式/FPS 限制",
        ["Multi Mode/Exclusions and Order"] = "多角色模式/排除與順序",
        ["Multi Mode/Region Lock"] = "多角色模式/區域鎖定",
        ["Multi Mode/Contingency"] = "多角色模式/應急處理",
        ["Multi Mode - Retainers"] = "多角色模式 - 僱員",
        ["Multi Mode - Deployables"] = "多角色模式 - 飛空艇/潛水艇",
        ["Wait For Venture Completion"] = "等待僱員探索完成",
        ["AutoRetainer will wait for all retainers to return before cycling to the next character in multi mode operation."] = "多角色模式中，AutoRetainer 會等待所有僱員完成後才切到下一個角色。",
        ["Advance Relog Threshold"] = "提前換角門檻",
        ["Minimum inventory slots to continue operation"] = "繼續運作所需最少背包空格",
        ["Synchronise Retainers (one time)"] = "同步僱員時間（一次性）",
        ["AutoRetainer will wait until all enabled retainers have completed their ventures. After that this setting will be disabled automatically and all characters will be processed."] = "AutoRetainer 會等待所有已啟用僱員完成探索；完成後此設定會自動關閉，並開始處理全部角色。",
        ["Enforce Full Character Rotation"] = "強制完整角色輪替",
        ["Recommended for users with > 15 characters, forces multi mode to make sure ventures are processed on all characters in order before returning to the beginning of the cycle."] = "建議角色超過 15 個時啟用。會強制多角色模式依序處理所有角色後，才回到循環開頭。",
        ["Order characters by venture completion time"] = "依探索完成時間排序角色",
        ["Characters that have completed ventures longer time ago will be checked first"] = "較早完成探索的角色會優先檢查。",
        ["Order characters by retainer level and cap"] = "依僱員等級與等級上限排序角色",
        ["Characters with retainers that can be levelled up will be done first; then, characters with retainers at max level; and then characters with retainers less than max level and level capped."] = "會先處理仍可升級僱員的角色，再處理滿級僱員，最後處理未滿級但受角色等級限制的僱員。",
        ["If no character is available for ventures, you will be logged off until any character is available again. Title screen movie will be disabled while this option and MultiMode are enabled."] = "若沒有任何角色可處理僱員探索，會登出並停在登入畫面，直到有角色可處理。此選項與 MultiMode 啟用時會停用標題畫面影片。",
        ["Upon relogging via AutoRetainer's UI or command, disable Multi Mode."] = "透過 AutoRetainer 介面或指令重新登入時，停用多角色模式。",
        ["Upon relogging via AutoRetainer's UI or command, do not reset preferred character."] = "透過 AutoRetainer 介面或指令重新登入時，不重置偏好角色。",
        ["Allows teleporting to houses without registering them in Lifestream. Note: the Lifestream plugin is still required for teleportation to work.\n\nWarning: This option is less reliable than registering your houses in Lifestream. Use it only if necessary."] = "允許不在 Lifestream 登錄房屋也能傳送到房屋。注意：傳送功能仍需要 Lifestream 插件。\n\n警告：此選項比在 Lifestream 登錄房屋更不穩定，只有必要時才使用。",
        ["Inventory Cleanup Plan Selection"] = "背包清理計畫選擇",
        ["Auto-open venture coffers"] = "自動開啟探索寶箱",
        ["Multi Mode only. Before logging out, all coffers will be opened unless your inventory space is too low."] = "僅限多角色模式。登出前會開啟所有寶箱，除非背包空間太低。",
        ["Multi Mode Expert Delivery"] = "多角色軍票繳納",
        ["Enable Multi Mode Expert Delivery"] = "啟用多角色軍票繳納",
        ["Maximum to open at once"] = "一次最多開啟數量",
        ["Enable selling items to retainer"] = "啟用將物品交給僱員出售",
        ["When AutoRetainer checks resents retainers to ventures, items will be sold according to Inventory Cleanup plan."] = "AutoRetainer 檢查並重新派出僱員探索時，會依照背包清理計畫出售物品。",
        ["Enable selling items to housing NPC"] = "啟用賣給房屋 NPC",
        ["When AutoRetainer enters a house, items will be sold according to the Inventory Cleanup plan. A housing vendor that supports item selling must be placed near the house entrance (not the workshop entrance)—you should be able to interact with the NPC immediately after entering."] = "AutoRetainer 進入房屋時，會依照背包清理計畫出售物品。支援出售物品的房屋商人必須放在房屋入口附近（不是工房入口）；進屋後應該能立刻互動。",
        ["Ignore NPC if retainer is available"] = "若僱員可用則忽略 NPC",
        ["Sell now"] = "立即出售",
        ["Auto-desynth items"] = "自動分解物品",
        ["Armory chest: "] = "兵裝庫：",
        ["Desynthese"] = "分解",
        ["Skip"] = "跳過",
        ["Enable context menu integration"] = "啟用右鍵選單整合",
        ["Allow selling/discarding items from Armory Chest"] = "允許從兵裝庫出售/丟棄物品",
        ["Demo mode"] = "演示模式",
        ["Do not sell/discard items, instead print in chat what would be sold"] = "不要實際出售/丟棄物品，改為在聊天窗輸出原本會處理的內容。",
        ["These items will always be sold, regardless of their source, as long as their stack count does not exceeds specified amount that you can specify below. Additionally, only these items will ever be sold to an NPC."] = "這些物品無論來源為何都會出售，只要堆疊數量不超過下方指定的上限。另外，只有這份清單內的物品會被賣給 NPC。",
        ["Maximum stack size to be sold"] = "可出售的最大堆疊數量",
        ["Ignore stack setting for this item"] = "此物品忽略堆疊數量設定",
        ["These items, when obtained from Quick Venture will be sold unless they have stacked with the same item."] = "這些物品若由快速探索取得會被出售，除非它們已與同物品堆疊在一起。",
        ["These items will always be discarded, regardless of their source, as long as their stack count does not exceeds specified amount that you can specify below. Discards occur very frequently, before and after each action that may alter inventory. Discard is always prioritized, even if same item is present in sell or desynthesis list, it will be discarded. Protected items won't be discarded. "] = "這些物品無論來源為何都會被丟棄，只要堆疊數量不超過下方指定的上限。丟棄會很頻繁地在可能改變背包的動作前後執行。丟棄永遠優先，即使同物品也在出售或分解清單中仍會丟棄；保護清單內的物品不會丟棄。",
        ["Maximum stack size to be discarded"] = "可丟棄的最大堆疊數量",
        ["AutoRetainer won't sell, desynthese, discard or hand in to Grand Company these items, even if they are included in any other processing lists."] = "即使這些物品也在其他處理清單中，AutoRetainer 也不會出售、分解、丟棄或繳納給軍隊。",
        ["Here you can assign preconfigured inventory cleanup lists to your registered characters."] = "這裡可以把預先設定好的背包清理清單指派給已登錄角色。",
        ["Here you can assign preconfigured exchange lists to your registered characters, as well as select delivery mode."] = "這裡可以把預先設定好的兌換清單指派給已登錄角色，並選擇繳納模式。",
        ["Search..."] = "搜尋...",
        ["~Character"] = "~角色",
        ["Plan"] = "計畫",
        ["Delivery mode"] = "繳納模式",
        ["Default Plan"] = "預設計畫",
        ["While this text is visible, hover over items while holding:"] = "此文字顯示時，按住下列按鍵並把滑鼠移到物品上：",
        ["Shift - add to Quick Venture Sell List"] = "Shift - 加入快速探索出售清單",
        ["* Items that already in Unconditional Sell List or Discard List WILL NOT BE ADDED to Quick Venture Sell List"] = "* 已在無條件出售清單或丟棄清單的物品不會加入快速探索出售清單",
        ["Ctrl - add to Unconditional Sell List"] = "Ctrl - 加入無條件出售清單",
        ["* Items that already in other lists WILL BE MOVED to Unconditional Sell List"] = "* 已在其他清單的物品會移到無條件出售清單",
        ["Tab - add to Discard List"] = "Tab - 加入丟棄清單",
        ["* Items that already in other lists WILL BE MOVED to Discard List"] = "* 已在其他清單的物品會移到丟棄清單",
        ["Alt - delete from any list"] = "Alt - 從任何清單移除",
        ["\nItems that are protected are unaffected by these actions"] = "\n受保護的物品不受這些操作影響",
        ["Added {0} to Quick Venture Sell List"] = "已將 {0} 加入快速探索出售清單",
        ["Added {0} to Unconditional Sell List"] = "已將 {0} 加入無條件出售清單",
        ["Added {0} to Discard List"] = "已將 {0} 加入丟棄清單",
        ["Removed {0} from Quick Venture Sell List"] = "已將 {0} 從快速探索出售清單移除",
        ["Removed {0} from Unconditional Sell List"] = "已將 {0} 從無條件出售清單移除",
        ["Removed {0} from Discard List"] = "已將 {0} 從丟棄清單移除",
        ["Removed {0} from Desynthesis List"] = "已將 {0} 從分解清單移除",
        ["Enable Expert Delivery continuation"] = "啟用軍票繳納續行",
        ["When Expert Delivery Continuation is enabled:\n- The plugin will automatically spend available Grand Company Seals to purchase items from the configured Exchange List.\n- If the Exchange List is empty, only Ventures will be purchased.\n- Make sure that \"Delivery Mode\" is not set to \"Disabled\" in \"Character Configuration\" section\n\nAfter seals have been spent:\n- Expert Delivery will resume automatically.\n- The process will repeat until there are no eligible items left to deliver or no seals remaining."] = "啟用軍票繳納續行時：\n- 插件會自動花費可用軍票，依照設定的兌換清單購買物品。\n- 如果兌換清單是空的，只會購買探險幣。\n- 請確認「角色設定」中的「繳納模式」不是「停用」。\n\n軍票花費後：\n- 軍票繳納會自動繼續。\n- 流程會重複直到沒有可繳納物品，或軍票不足為止。",
        ["When enabled:\n- Characters with teleportation enabled will automatically deliver items for expert delivery and buy items according to exchange plan, if their rank is sufficient, during multi mode."] = "啟用時：\n- 多角色模式中，已啟用傳送且階級足夠的角色會自動繳納軍票物品，並依照兌換計畫購買物品。",
        ["Only when workstation is not locked"] = "僅在工作站未鎖定時",
        ["Inventory slots remaining to trigger delivery, less or equal"] = "剩餘背包格小於等於此值時觸發繳納",
        ["Only primary inventory is accounted for, not armory"] = "只計算主要背包，不包含兵裝庫",
        ["Trigger on venture exhaustion"] = "探險幣不足時觸發",
        ["This may cause situation where you will just go to GC exchange every login. Make sure you have a purchase plan to buy enough ventures set. "] = "這可能導致每次登入都前往軍隊兌換。請確認購買計畫有設定購買足夠探險幣。",
        ["Ventures remaining to trigger delivery, less or equal"] = "剩餘探險幣小於等於此值時觸發繳納",
        ["Use Priority seal allowance, if possible"] = "可用時使用軍票優先支給券",
        ["Use Free Company seal buff, if possible"] = "可用時使用部隊軍票加成",
        ["Select the items to be purchased automatically during Grand Company Expert Delivery operations.\nPurchase Logic:\n- The system will attempt to purchase the first available item from the list.\n- Purchases will continue until the quantity of that item in your inventory reaches the specified target amount.\nIf no listed items are available for purchase, or they cannot fit into your inventory:\n- The system will purchase Ventures instead.\n- Venture purchases will continue until your Venture count reaches 65,000.\nOnce the Venture cap is reached and no other purchases are possible:\n- Any excess Grand Company Seals will be discarded."] = "選擇軍票繳納流程中要自動購買的物品。\n購買邏輯：\n- 系統會嘗試購買清單中第一個可購買的物品。\n- 會持續購買，直到背包中該物品數量達到指定目標。\n若清單物品都無法購買，或無法放入背包：\n- 系統會改買探險幣。\n- 會持續購買直到探險幣達到 65,000。\n探險幣達上限且沒有其他可購買項目時：\n- 多餘軍票會被丟棄。",
        ["Add new plan"] = "新增計畫",
        ["Copy"] = "複製",
        ["Paste"] = "貼上",
        ["Make this plan default. Current default plan will be overwritten. Hold CTRL and click."] = "將此計畫設為預設。會覆蓋目前預設計畫。按住 CTRL 再點擊。",
        ["Delete this plan. Hold CTRL and click."] = "刪除此計畫。按住 CTRL 再點擊。",
        ["Used by current character"] = "目前角色正在使用",
        ["Not used by current character"] = "目前角色未使用",
        ["Unassign"] = "取消指派",
        ["Assign"] = "指派",
        ["Used by {0} characters in total"] = "總共 {0} 個角色使用",
        ["Not used by any characters"] = "沒有任何角色使用",
        ["Default exchange plan can not be renamed"] = "預設兌換計畫不能重新命名",
        ["Exchange plan name"] = "兌換計畫名稱",
        ["Seals to keep"] = "保留軍票",
        ["This amount of seals will be kept after purchase list is executed. However, this value will be capped to be no more than 20000 seals less than maximum possible, according to character's rank. "] = "執行購買清單後會保留這個數量的軍票。不過依角色軍階，此值最多只能低於軍票上限 20000。",
        ["Finish by purchasing items"] = "最後購買物品",
        ["If selected, after final exchange items will be purchased, otherwise - purchase will not be made until seals are capped again."] = "勾選後，最後一次兌換後會購買物品；未勾選則等軍票再次滿額才購買。",
        ["Add Items"] = "新增物品",
        ["All Categories"] = "全部分類",
        ["Category"] = "分類",
        ["Fill weapons and armor purchases optimally for extra FC points"] = "最佳化購買武器與防具以取得額外部隊點數",
        ["Select this option to fill in your plan with all purchaseable weapons and gear items. By doing so, weapons and items will be purchased and handed right back to the Grand Company, maximizing amount of generated Free Company points. All these items will be placed at the end of the list and only purchased if nothing else is available."] = "選擇此項會把所有可購買的武器與裝備加入計畫。這些物品會被購買後立刻繳回軍隊，以最大化產生的部隊點數。它們會放在清單末端，只在沒有其他可購買項目時購買。",
        ["Add all missing items"] = "加入所有缺少物品",
        ["Reset quantities to 0"] = "數量重設為 0",
        ["Remove 0-quantity items"] = "移除數量為 0 的物品",
        ["Clear the list (Hold CTRL and click)"] = "清空清單（按住 CTRL 再點擊）",
        ["Actions"] = "動作",
        ["Only Selected"] = "只顯示已選",
        ["GC"] = "軍隊",
        ["Lv"] = "等級",
        ["Price"] = "價格",
        ["Keep"] = "保留",
        ["One-Time"] = "單次",
        ["Move to the top"] = "移到最上方",
        [" (unavailable)"] = "（不可用）",
        ["Unique"] = "唯一",
        ["Select amount of items to keep in your inventory"] = "選擇背包中要保留的物品數量",
        ["Select amount of items to purchase once. Whenever purchase is made on any character using this plan, an amount will be subtracted from this value. Once it reaches 0, it will back to \"Keep\" amount."] = "選擇單次要購買的數量。任何使用此計畫的角色購買時，都會從此數值扣除；歸零後會回到「保留」數量。",
        ["Duplicate this listing."] = "複製此項目。",
        ["Deletes item from the list if there are multiple copies of it or sets it's amount to 0 if there is only one copy"] = "若清單中有多份相同物品會刪除此項；若只有一份則把數量設為 0。",
        ["FPS Limiter is only active when Multi Mode is enabled"] = "FPS 限制只會在多角色模式啟用時生效",
        ["Allow extra low FPS limiter values"] = "允許更低的 FPS 限制數值",
        ["No support is provided if you enable this and run into ANY errors in Multi Mode"] = "啟用後若多角色模式發生任何錯誤，原作者不提供支援。",
        ["Here you can sort your characters. This will affect order in which they will be processed by Multi Mode as well as how they will appear in plugin interface and login overlay."] = "這裡可以排序角色。排序會影響多角色模式處理順序，也會影響插件介面與登入浮層的顯示順序。",
        ["Add current account"] = "新增目前帳號",
        ["Ignore other characters that have not been enabled in MultiMode"] = "忽略未啟用 MultiMode 的其他角色",
        ["If game is inactive: (requires NotificationMaster to be installed and enabled)"] = "遊戲不在前景時：（需要安裝並啟用 NotificationMaster）",
        ["Do not notify if AutoRetainer is enabled or MultiMode is running"] = "AutoRetainer 已啟用或 MultiMode 運作中時不要通知",
        ["You may setup account whitelist. In the event you will log in using non-whitelisted account, AutoRetainer will not record any characters, retainers, or submarines."] = "可設定帳號白名單。若登入未列入白名單的帳號，AutoRetainer 不會記錄任何角色、僱員或潛水艇。",
        ["Current whitelist status: Disabled. To enable, add some account to it."] = "目前白名單狀態：停用。加入至少一個帳號即可啟用。",
        ["Current whitelist status: Enabled. To disable, remove all accounts from it."] = "目前白名單狀態：啟用。移除所有帳號即可停用。",
        ["Skip item sell/trade confirmation while plugin is active"] = "插件啟用時略過物品出售/交易確認",
        ["Unoptimal submersible configuration alerts:"] = "潛水艇配置不佳提醒：",
        ["Ctrl+click to delete"] = "Ctrl + 點擊刪除",
        ["Rank:"] = "階級：",
        ["Configurations:"] = "配置：",
        ["NOT"] = "非",
        ["Build (1)"] = "零件配置 (1)",
        ["Build (2)"] = "零件配置 (2)",
        ["Build (3)"] = "零件配置 (3)",
        ["Build (4)"] = "零件配置 (4)",
        ["Level (1)"] = "等級 (1)",
        ["Level (2)"] = "等級 (2)",
        ["Level (3)"] = "等級 (3)",
        ["Level (4)"] = "等級 (4)",
        ["Route (1)"] = "航線 (1)",
        ["Route (2)"] = "航線 (2)",
        ["Route (3)"] = "航線 (3)",
        ["Route (4)"] = "航線 (4)",
        ["Save as..."] = "另存為...",
        ["Comma-separated values"] = "逗號分隔值",
        ["Select submersibles:"] = "選擇潛水艇：",
        ["Search"] = "搜尋",
        ["Unlock plan: "] = "解鎖計畫：",
        ["not selected"] = "未選擇",
        ["Point plan: "] = "航點計畫：",
        ["Ranges:"] = "範圍：",
        ["Level range:"] = "等級範圍：",
        ["Hull:"] = "艇體：",
        ["Stern:"] = "船尾：",
        ["Bow:"] = "船首：",
        ["Bridge:"] = "艦橋：",
        ["Behavior:"] = "行為：",
        ["Plan:"] = "計畫：",
        ["Non selected"] = "未選擇",
        ["Mode:"] = "模式：",
        ["First Sub Behavior:"] = "第一艘潛水艇行為：",
        ["First Sub Plan:"] = "第一艘潛水艇計畫：",
        ["First Sub Mode:"] = "第一艘潛水艇模式：",
        ["Copy to Clipboard"] = "複製到剪貼簿",
        ["Merge with Clipboard"] = "與剪貼簿合併",
        ["Hold CTRL and click"] = "按住 CTRL 並點擊",
        ["Mass addition/removal"] = "批量新增/移除",
        ["Select Categories"] = "選擇分類",
        ["All"] = "全部",
        ["+Main/offhand"] = "+主手/副手",
        ["+Armor"] = "+防具",
        ["Filter by name"] = "依名稱篩選",
        ["Select rarity"] = "選擇稀有度",
        ["Any rarity"] = "任意稀有度",
        ["Minimum item level"] = "最低物品等級",
        ["Maximum item level"] = "最高物品等級",
        ["Tradeable"] = "可交易",
        ["Click to add this single item to list immediately"] = "點擊後立即將此單一物品加入清單",
        ["Right click to remove this single item from list immediately"] = "右鍵點擊後立即從清單移除此單一物品",
        ["Add these items to list"] = "將這些物品加入清單",
        ["Remove these items to list"] = "從清單移除這些物品",
        ["Import discard entries from Discard Helper"] = "從 Discard Helper 匯入丟棄項目",
        ["If you're using Discard Helper plugin, you may import entries from it using this button. They will be merged with your existing entries. Hold CTRL and click."] = "若你使用 Discard Helper 插件，可用此按鈕匯入項目。匯入內容會與既有項目合併。按住 CTRL 並點擊。",
        ["Import blacklisted entries from Discard Helper"] = "從 Discard Helper 匯入黑名單項目",
        ["~Item"] = "~物品",
        ["Space - add to Desynthesis List"] = "Space - 加入分解清單",
        ["* Items that already in other lists WILL BE MOVED to Desynthesis List"] = "* 已在其他清單的物品會移到分解清單",
        ["For"] = "持續",
        ["hours..."] = "小時...",
        ["Remove all locks"] = "移除全部鎖定",
        ["Halt all plugin operation"] = "停止所有插件操作",
        ["Exclude deployable from operation"] = "從操作中排除此飛空艇/潛水艇",
        ["Exclude captain from multi mode rotation"] = "從多角色輪替中排除此船長角色",
        ["Here you can apply various fallback actions to perform in the case of some common failure states or potential operation errors."] = "這裡可設定發生常見失敗狀態或潛在操作錯誤時要執行的備援動作。",
        ["Ceruleum Tanks Expended"] = "青燐水箱耗盡",
        ["Applies selected fallback action in the case of insufficient Ceruleum Tanks to deploy vessel on a new voyage."] = "青燐水箱不足、無法派出新航行時，套用選定的備援動作。",
        ["Unable to Repair Deployable"] = "無法修理飛空艇/潛水艇",
        ["Applies selected fallback action in the case of insufficient Magitek Repair Materials to repair a vessel."] = "魔導機械修理材料不足、無法修理船隻時，套用選定的備援動作。",
        ["Inventory at Capacity"] = "背包已滿",
        ["Applies selected fallback action in the case of the captain's inventory having insufficient space to receive voyage rewards."] = "船長背包空間不足、無法領取航行獎勵時，套用選定的備援動作。",
        ["Critical Operation Failure"] = "嚴重操作失敗",
        ["Applies selected fallback action in the case of any unknown or miscellaneous error."] = "發生未知或其他雜項錯誤時，套用選定的備援動作。",
        ["Jailed by the GM"] = "被 GM 關入監獄",
        ["Terminate the game"] = "結束遊戲",
        ["Applies selected fallback action in the case if you got jailed by the GM while plugin is running. Good luck!"] = "插件運作時若被 GM 關入監獄，套用選定的備援動作。祝你好運。",
        ["Character Order"] = "角色順序",
        ["Character"] = "角色",
        ["Toggles"] = "開關",
        ["Deletion"] = "刪除",
        ["Enable retainers"] = "啟用僱員",
        ["Enable deployables"] = "啟用飛空艇/潛水艇",
        ["Display on login overlay"] = "顯示在登入浮層",
        ["Count gil on this character towards total"] = "將此角色金幣計入總額",
        ["Reset FC data and deployable data for this character. It will regenerate once you log in and access workshop panel."] = "重置此角色的部隊與飛空艇/潛水艇資料。重新登入並開啟工房面板後會再產生。",
        ["Hold CTRL and click to delete stored character data. It will be recreated once you relog back."] = "按住 CTRL 並點擊以刪除已儲存角色資料。重新登入後會重新建立。",
        ["Hold CTRL and click to delete stored character data and prevent it from being ever created again, effectively excluding it from being processed by AutoRetainer entirely in any ways."] = "按住 CTRL 並點擊以刪除角色資料，並阻止之後再次建立，等同完全排除此角色，不讓 AutoRetainer 以任何方式處理。",
        ["Excluded Characters"] = "已排除角色",
        ["Target frame rate when idling"] = "閒置時目標 FPS",
        ["Target frame rate when operating"] = "運作時目標 FPS",
        ["Release FPS lock when game is active"] = "遊戲在前景時解除 FPS 鎖定",
        ["Limiter active only when shutdown timer is set"] = "只有設定關機倒數時啟用限制器",
        ["Wait even when already logged in"] = "即使已登入也等待",
        ["The number of seconds AutoRetainer should log in early before submarines on this character are ready to be resent."] = "此角色潛水艇可重新派出前，AutoRetainer 應提前登入的秒數。",
        ["Retainer venture processing cutoff, minutes"] = "僱員探索處理截止時間，分鐘",
        ["If set to a value greater than 0, AutoRetainer will stop processing any retainers this number of minutes before any character is scheduled to redeploy submarines, taking all previous settings into account."] = "若設定大於 0，AutoRetainer 會依前面設定計算，在任何角色預計重新派出潛水艇前這麼多分鐘停止處理僱員。",
        ["Sell items from Unconditional sell list right after deployment (requires retainers)"] = "派出後立即出售無條件出售清單物品（需要僱員）",
        ["Periodically check FC chest for gil upon entering workshop"] = "進入工房時定期檢查部隊箱金幣",
        ["Periodically checks the Free Company chest when entering the Workshop to keep the gil counter up to date."] = "進入工房時定期檢查部隊箱，讓金幣統計保持最新。",
        ["Check frequency, hours"] = "檢查頻率，小時",
        ["Reset cooldowns"] = "重置冷卻",
        ["Shutdown the game after all deployables have been processed"] = "全部飛空艇/潛水艇處理完後關閉遊戲",
        ["Don't shutdown if there are deployables that return within this amount of hours"] = "若有飛空艇/潛水艇會在此小時數內返航，則不要關閉遊戲",
        ["Can NOT shutdown"] = "無法關閉遊戲",
        ["Display overlay notification if one of retainers has completed a venture"] = "若任一僱員完成探索，顯示浮層通知",
        ["Do not display overlay in duty or combat"] = "副本或戰鬥中不要顯示浮層",
        ["Include other characters"] = "包含其他角色",
        ["Display notification in game chat"] = "在遊戲聊天顯示通知",
        ["Send desktop notification on retainers available"] = "僱員可處理時送出桌面通知",
        ["Flash taskbar"] = "閃爍工作列",
        ["Exclude from list"] = "從清單排除",
        ["Delete listed characters from AutoRetainer"] = "從 AutoRetainer 刪除列出的角色",
        ["Prune deleted characters in a single click."] = "一鍵清理已刪除角色。",
        ["To continue, you need to install JustBackup plugin."] = "若要繼續，需要安裝 JustBackup 插件。",
        ["Open Plugin Installer"] = "開啟插件安裝器",
        ["Open character list now"] = "立即開啟角色清單",
        ["3. Make sure you are logged with the correct account and copy entire page's content by pressing CTRL+A then CTRL+C"] = "3. 確認你已登入正確帳號，並按 CTRL+A 再 CTRL+C 複製整個頁面內容。",
        ["4. Once finished, click the following button:"] = "4. 完成後，點擊下方按鈕：",
        ["Prepare Character Cleanup"] = "準備角色清理",
        ["Update Character List"] = "更新角色清單",
        ["Did not read any characters"] = "未讀取到任何角色",
        ["Expert items"] = "專家項目",
    };

    internal static string T(string text)
    {
        if(string.IsNullOrEmpty(text)) return text;
        var id = "";
        var visible = text;
        var triple = visible.IndexOf("###", StringComparison.Ordinal);
        if(triple >= 0)
        {
            id = visible[triple..];
            visible = visible[..triple];
        }
        else
        {
            var doubleHash = visible.IndexOf("##", StringComparison.Ordinal);
            if(doubleHash >= 0)
            {
                id = visible[doubleHash..];
                visible = visible[..doubleHash];
            }
        }
        if(Tw.TryGetValue(visible, out var translated)) return translated + id;
        var trimmed = visible.Trim();
        if(trimmed != visible && Tw.TryGetValue(trimmed, out translated)) return translated + id;
        if(visible.Contains('/'))
        {
            var translatedPath = string.Join("/", visible.Split('/').Select(part => Tw.TryGetValue(part, out var partTranslated) ? partTranslated : part));
            if(translatedPath != visible) return translatedPath + id;
        }
        return text;
    }

    internal const string CharPlant = "";
    internal const string CharLevel = "";
    internal const string CharItemLevel = "";
    internal const string CharDice = "";
    internal const string CharDeny = "";
    internal const string CharQuestion = "";
    internal const string CharLevelSync = "";
    internal const string CharP = "";
    internal const string StrDCV = "";

    internal const string IconRefresh = "\uf2f9";
    internal const string IconMultiMode = "\uf021";
    internal const string IconDuplicate = "\uf24d";
    internal const string IconGil = "\uf51e";
    internal const string IconPlanner = "\uf0ae";
    internal const string IconSettings = "\uf013";
    internal const string IconWarning = "\uf071";

    internal const string IconAnchor = "\uf13d";
    internal const string IconLevelup = "\ue098";
    internal const string IconResend = "\ue4bb";
    internal const string IconUnlock = "\uf13e";
    internal const string IconRepeat = "\uf363";
    internal const string IconPath = "\uf55b";
    internal const string IconFire = "\uf06d";

    internal static string LogOutAndExitGame => Svc.Data.GetExcelSheet<Addon>().GetRow(116).Text.GetText(true).Cleanup();

    internal static readonly ReadOnlyDictionary<UnlockMode, string> UnlockModeNames = new(new Dictionary<UnlockMode, string>()
    {
        { UnlockMode.MultiSelect, "Pick max amount of destinations" },
        { UnlockMode.SpamOne, "Spam one destination" },
        { UnlockMode.WhileLevelling, "Include one unlock destination while levelling" },
    });

    internal static readonly ReadOnlyDictionary<GCDeliveryType, string> GCDeliveryTypeNames = new(new Dictionary<GCDeliveryType, string>()
    {
        { GCDeliveryType.Disabled, "停用" },
        { GCDeliveryType.Hide_Armoury_Chest_Items, "隱藏兵裝庫物品" },
        { GCDeliveryType.Hide_Gear_Set_Items, "隱藏套裝配置物品" },
        { GCDeliveryType.Show_All_Items, "顯示所有物品" },
    });

    internal static readonly ReadOnlyDictionary<PlanCompleteBehavior, string> PlanCompleteBehaviorNames = new(new Dictionary<PlanCompleteBehavior, string>()
    {
        { PlanCompleteBehavior.Restart_plan, "重新從頭執行" },
        { PlanCompleteBehavior.Assign_Quick_Venture, "改派快速探索" },
        { PlanCompleteBehavior.Do_nothing, "不做任何事" },
        { PlanCompleteBehavior.Repeat_last_venture, "重複最後一項" },
    });

    // Backport of upstream #164: retain TW literals if the regional sheet is missing.
    private static string[] ReadBellText(uint[] rows, string[] fallback)
    {
        try
        {
            var sheet = Svc.Data.GetExcelSheet<QuestDialogueText>(name: "custom/000/CmnDefRetainerCall_00010");
            return Helpers.RetainerCompatibility.MenuCandidates(
                rows.Select(row => sheet?.GetRowOrDefault(row)?.Value.GetText()).Concat(fallback));
        }
        catch(Exception ex)
        {
            PluginLog.Warning($"[AutoRetainer TW] 僱員選單文字讀取失敗，改用既有翻譯：{ex.Message}");
            return Helpers.RetainerCompatibility.MenuCandidates(fallback);
        }
    }

    internal static readonly (string Normal, string GameFont) Digits = ("0123456789", "");

    private static string[] cachedFieldExplorationNames;
    internal static string[] FieldExplorationNames => cachedFieldExplorationNames ??= ReadBellText([196, 198, 200, 202], FieldExplorationNamesFallback);
    private static readonly string[] FieldExplorationNamesFallback =
    [
        "Field Exploration.",
        "Highland Exploration.",
        "Woodland Exploration.",
        "Waterside Exploration.",
        "探索依頼：平地　　（必要ベンチャースクリップ：2枚）",
        "探索依頼：山岳　　（必要ベンチャースクリップ：2枚）",
        "探索依頼：森林　　（必要ベンチャースクリップ：2枚）",
        "探索依頼：水辺　　（必要ベンチャースクリップ：2枚）",
        "Felderkundung (2 Wertmarken)",
        "Hochlanderkundung (2 Wertmarken)",
        "Forsterkundung (2 Wertmarken)",
        "Gewässererkundung (2 Wertmarken)",
        "Exploration en plaine (2 jetons)",
        "Exploration en montagne (2 jetons)",
        "Exploration en forêt (2 jetons)",
        "Exploration en rivage (2 jetons)",
        "平地探索委托（需要2枚探险币）",
        "山岳探索委托（需要2枚探险币）",
        "森林探索委托（需要2枚探险币）",
        "水岸探索委托（需要2枚探险币）",
        "平地探索委託（需要2枚探險幣）",
        "山岳探索委託（需要2枚探險幣）",
        "森林探索委託（需要2枚探險幣）",
        "水岸探索委託（需要2枚探險幣）",
        "탐색수행: 평지 (필요한 집사 급료: 2개)",
        "탐색수행: 산악 (필요한 집사 급료: 2개)",
        "탐색수행: 삼림 (필요한 집사 급료: 2개)",
        "탐색수행: 물가 (필요한 집사 급료: 2개)",
    ];

    private static string[] cachedHuntingVentureNames;
    internal static string[] HuntingVentureNames => cachedHuntingVentureNames ??= ReadBellText([195, 197, 199, 201], HuntingVentureNamesFallback);
    private static readonly string[] HuntingVentureNamesFallback =
    [
        "Hunting.",
        "Mining.",
        "Botany.",
        "Fishing.",
        "調達依頼：渉猟　　（必要ベンチャースクリップ：1枚）",
        "調達依頼：採掘　　（必要ベンチャースクリップ：1枚）",
        "調達依頼：園芸　　（必要ベンチャースクリップ：1枚）",
        "調達依頼：漁猟　　（必要ベンチャースクリップ：1枚）",
        "Beutezug (1 Wertmarke)",
        "Mineraliensuche (1 Wertmarke)",
        "Ernteausflug (1 Wertmarke)",
        "Fischzug (1 Wertmarke)",
        "Travail de chasse (1 jeton)",
        "Travail de mineur (1 jeton)",
        "Travail de botaniste (1 jeton)",
        "Travail de pêche (1 jeton)",
        "狩猎筹集委托（需要1枚探险币）",
        "采矿筹集委托（需要1枚探险币）",
        "采伐筹集委托（需要1枚探险币）",
        "捕鱼筹集委托（需要1枚探险币）",
        "狩獵籌集委託（需要1枚探險幣）",
        "採礦籌集委託（需要1枚探險幣）",
        "採伐籌集委託（需要1枚探險幣）",
        "捕魚籌集委託（需要1枚探險幣）",
        "조달수행: 사냥 (필요한 집사 급료: 1개)",
        "조달수행: 광부 (필요한 집사 급료: 1개)",
        "조달수행: 원예가 (필요한 집사 급료: 1개)",
        "조달수행: 어부 (필요한 집사 급료: 1개)",
    ];

    private static string[] cachedQuickExploration;
    internal static string[] QuickExploration => cachedQuickExploration ??= ReadBellText([402], QuickExplorationFallback);
    private static readonly string[] QuickExplorationFallback =
    [
        "Quick Exploration.",
        "ほりだしもの依頼　（必要ベンチャースクリップ：2枚）",
        "Schneller Streifzug (2 Wertmarken)",
        "Tâche improvisée (2 jetons)",
        "自由探索委托（需要2枚探险币）",
        "自由探索委託（需要2枚探險幣）",
        "발굴수행 (필요한 집사 급료: 2개)",
        "自由尋寶委託（需要2枚探險幣）",
    ];

    internal static readonly string[] Entrance =
    [
        "ハウスへ入る",
        "进入房屋",
        "進入房屋",
        "Eingang",
        "Entrée",
        "Entrance",
        "주택으로 들어가기",
    ];

    internal static string ApartmentEntrance => Svc.Data.GetExcelSheet<EObjName>().GetRow(2007402).Singular.ToString();

    internal static readonly string[] ConfirmHouseEntrance =
    [
        "「ハウス」へ入りますか？",
        "要进入这间房屋吗？",
        "要進入這間房屋嗎？",
        "Das Gebäude betreten?",
        "Entrer dans la maison ?",
        "Enter the estate hall?",
        "'주택'으로 들어가시겠습니까?",
    ];

    private static string[] cachedRetainerAskCategoryText;
    internal static string[] RetainerAskCategoryText => cachedRetainerAskCategoryText ??= ReadBellText([194], RetainerAskCategoryTextFallback);
    private static readonly string[] RetainerAskCategoryTextFallback =
    [
        "依頼するリテイナーベンチャーを選んでください",
        "请选择要委托的探险",
        "請選擇要委託的探險",
        "Wähle eine Unternehmung, auf die du den Gehilfen schicken möchtest.",
        "Choisissez un type de tâche :",
        "Select a category.",
        "집사 수행의 종류를 선택하십시오.",
    ];

    internal static string[] BellName => [Svc.Data.GetExcelSheet<EObjName>().GetRow(2000401).Singular.GetText(), "リテイナーベル"];

    //0	TEXT_HOUFIXMANSIONENTRANCE_00359_HOUSINGAREA_MENU_ENTER_MYROOM	Go to your apartment
    //0	TEXT_HOUFIXMANSIONENTRANCE_00359_HOUSINGAREA_MENU_ENTER_MYROOM	自分の部屋に移動する
    //0	TEXT_HOUFIXMANSIONENTRANCE_00359_HOUSINGAREA_MENU_ENTER_MYROOM	Die eigene Wohnung betreten
    //0	TEXT_HOUFIXMANSIONENTRANCE_00359_HOUSINGAREA_MENU_ENTER_MYROOM	Aller dans votre appartement

    internal static readonly string[] GoToYourApartment =
    [
        "Go to your apartment",
        "自分の部屋に移動する",
        "移动到自己的房间",
        "移動到自己的房間",
        "Die eigene Wohnung betreten",
        "Aller dans votre appartement",
        "자신의 방으로 이동",
    ];

    internal static readonly string[] SkipCutsceneStr =
    [
        "Skip cutscene?",
        "要跳过这段过场动画吗？",
        "要跳過這段過場動畫嗎？",
        "Videosequenz überspringen?",
        "Passer la scène cinématique ?",
        "このカットシーンをスキップしますか？",
        "영상을 건너뛰시겠습니까?",
    ];
    //11	TEXT_CMNDEFHOUSINGPERSONALROOMENTRANCE_00178_GOTO_WORKSHOP	Move to the company workshop
    //11	TEXT_CMNDEFHOUSINGPERSONALROOMENTRANCE_00178_GOTO_WORKSHOP	地下工房に移動する
    //11	TEXT_CMNDEFHOUSINGPERSONALROOMENTRANCE_00178_GOTO_WORKSHOP	Die Ge<SoftHyphen/>sell<SoftHyphen/>schaftswerkstätte betreten
    //11	TEXT_CMNDEFHOUSINGPERSONALROOMENTRANCE_00178_GOTO_WORKSHOP	Aller dans l'atelier de compagnie
    internal static readonly string[] EnterWorkshop = ["Move to the company workshop", "地下工房に移動する", "移动到部队工房", "移動到部隊工房", "Die Gesellschaftswerkstätte betreten", "Aller dans l'atelier de compagnie", "지하공방으로 이동", Svc.Data.GetExcelSheet<QuestDialogueText>(name: "custom/001/CmnDefHousingPersonalRoomEntrance_00178").GetRow(11).Value.GetText()];

    internal static readonly string[] AirshipManagement = ["Airship Management", "飛空艇の管理", "管理飞空艇", "管理飛空艇", "Luftschiff verwalten", "Contrôle aérien", "비공정 관리"];
    internal static readonly string[] SubmarineManagement = ["Submersible Management", "潜水艦の管理", "管理潜水艇", "管理潛水艇", "Tauchboot verwalten", "Contrôle sous-marin", "잠수함 관리"];
    internal static readonly string[] CancelVoyage = ["Cancel", "キャンセル", "取消", "Abbrechen", "Annuler", "취소"];
    internal static readonly string[] NothingVoyage = ["Nothing.", "やめる", "取消", "Nichts", "Annuler", "그만두기"];
    internal static readonly string[] DeployOnSubaquaticVoyage = ["Deploy submersible on subaquatic voyage", "ボイジャー出港", "出发", "出發", "Auf Erkundung gehen", "Expédier le sous-marin", "탐사 출항"];
    internal static readonly string[] ViewPrevVoyageLog = ["View previous voyage log", "前回のボイジャー報告", "上次的远航报告", "上次的遠航報告", "Bericht der letzten Erkundung", "Consulter le journal de la précédente expédition", "이전 탐사 보고서"];
    internal static readonly string[] VoyageQuitEntry = ["Quit", "やめる", "取消", "Beenden", "Annuler", "그만두기"];
    internal static readonly string[] ChangeSubmersibleComponents = ["Change submersible components", "パーツの変更", "Bauteile austauschen", "Changer les éléments", "부품 변경", "更换配件", "更換配件"];
    internal static readonly string[] RegisterSub = ["Outfit and register a submersible.", "潜水艦の新規登録", "Registrierung eines neuen Tauchboots", "Enregistrement d'un sous-marin", "새 잠수함 등록", "登记新的潜水艇", "登記新的潛水艇"];

    internal static readonly string[] PanelAirship = ["Select an airship.", "飛空艇を選択してください。", "请选择飞空艇。", "請選擇飛空艇。", "Wähle ein Luftschiff.", "Choisissez un aéronef.", "비공정을 선택하십시오."];
    internal static readonly string[] PanelSubmersible = ["Select a submersible.", "潜水艦を選択してください。", "请选择潜水艇。", "請選擇潛水艇。", "Wähle ein Tauchboot.", "Choisissez un sous-marin.", "잠수함을 선택하십시오."];

    //2004353	entrance to additional chambers	0	entrances to additional chambers	0	1	1	0	0
    internal static string[] AdditionalChambersEntrance =>
    [
        Svc.Data.GetExcelSheet<EObjName>().GetRow(2004353).Singular.GetText(),
        Regex.Replace(Svc.Data.GetExcelSheet<EObjName>().GetRow(2004353).Singular.GetText(), @"\[.*?\]", "")
    ];

    //2005274	voyage control panel	0	voyage control panels	0	0	1	0	0
    internal static string PanelName => Svc.Data.GetExcelSheet<EObjName>().GetRow(2005274).Singular.GetText();

    //4160	60	9	0	False	Unable to retrieve extracted items. Insufficient inventory/crystal inventory space.
    internal static string VoyageInventoryError => Svc.Data.GetExcelSheet<LogMessage>().GetRow(4160).Text.ToDalamudString().GetText();

    internal static string[] UnableToVisitWorld = ["Unable to execute command. Character is currently visiting the", "他のデータセンター", "无法进行该操作，其他玩家正在操作该潜水艇。", "無法進行該操作，其他玩家正在操作該潛水艇。", "Der Vorgang kann nicht ausgeführt werden, da der Charakter gerade das Datenzentrum", "Impossible d'exécuter cette commande. Le personnage se trouve dans un autre centre de traitement de données", "다른 데이터 센터"];

    //4169	60	9	0	False	Unable to repair vessel component without the required <SheetEn(Item,3,IntegerParameter(1),1,1)/>.
    //4272	60	9	0	False Unable to repair vessel.Insufficient<SheetEn(Item,3,IntegerParameter(1),3,1)/>.
    //4169	60	9	0	False	修理に必要な<Sheet(Item,IntegerParameter(1),0)/>を持っていません。
    //4272	60	9	0	False	修理に必要な<Sheet(Item,IntegerParameter(1),0)/>が足りません。
    //4169	60	9	0	False	未持有修理所必需的<Sheet(Item,IntegerParameter(1),0)/>。
    //4272	60	9	0	False	沒有修理所必需的<Sheet(Item,IntegerParameter(1),0)/>。
    //4272	60	9	0	False	Du hast nicht genug <SheetDe(Item,5,IntegerParameter(1),2,4,1)/> für die Reparatur.
    //4169	60	9	0	False	Für die Reparatur ist <SheetDe(Item,1,IntegerParameter(1),1,1,1)/> erforderlich.
    //4169	60	9	0	False	Réparation impossible. Vous n'avez pas <SheetFr(Item,2,IntegerParameter(1),1,1)/> nécessaire.
    //4272	60	9	0	False	Vous n'avez pas <SheetFr(Item,2,IntegerParameter(1),1,1)/> nécessaire à la réparation.

    internal static readonly string[] UnableToRepairVessel = ["修理に必要な", "修理所必需的", "Unable to repair vessel", "Du hast nicht genug", "Für die Reparatur ist", "Réparation impossible. Vous n'avez pas", "nécessaire à la réparation", "수리에 필요한"];

    //11	TEXT_HOUFIXCOMPANYSUBMARINE_00447_SUBMARINE_CMD_REPAIR_PARTS	パーツの修理
    //11	TEXT_HOUFIXCOMPANYSUBMARINE_00447_SUBMARINE_CMD_REPAIR_PARTS	Bauteile reparieren
    //11	TEXT_HOUFIXCOMPANYSUBMARINE_00447_SUBMARINE_CMD_REPAIR_PARTS	Réparer des éléments
    //11	TEXT_HOUFIXCOMPANYSUBMARINE_00447_SUBMARINE_CMD_REPAIR_PARTS	修理配件
    //10	TEXT_CMNDEFCOMPANYCOMMANDERBOARD_00258_AIRSHIP_CMD_REPAIR_PARTS	パーツの修理
    //10	TEXT_CMNDEFCOMPANYCOMMANDERBOARD_00258_AIRSHIP_CMD_REPAIR_PARTS	Bauteile reparieren
    //10	TEXT_CMNDEFCOMPANYCOMMANDERBOARD_00258_AIRSHIP_CMD_REPAIR_PARTS	Réparer des éléments
    //10	TEXT_CMNDEFCOMPANYCOMMANDERBOARD_00258_AIRSHIP_CMD_REPAIR_PARTS	修理配件

    internal static readonly string[] WorkshopRepair =
    [
        "Repair submersible components",
        "Repair airship components",
        "パーツの修理",
        "Bauteile reparieren",
        "Réparer des éléments",
        "パーツの修理",
        "Bauteile reparieren",
        "Réparer des éléments",
        "修理配件",
        "부품 수리",
    ];

    //Use <If(Equal(IntegerParameter(4),1))>your last <SheetEn(Item,3,IntegerParameter(2),1,1)/><Else/><Value>IntegerParameter(3)</Value> of your <Value>IntegerParameter(4)</Value> <SheetEn(Item,3,IntegerParameter(2),2,1)/></If> to repair your vessel's <SheetEn(Item,3,IntegerParameter(1),1,1)/>?
    //6587	<If(Equal(IntegerParameter(3),1))><Clickable(<SheetDe(Item,2,IntegerParameter(2),1,4,1)/>)/><Else/><Value>IntegerParameter(3)</Value> <SheetDe(Item,5,IntegerParameter(2),2,4,1)/></If> (Besitz: <Value>IntegerParameter(4)</Value>) benutzen, um <SheetDe(Item,2,IntegerParameter(1),1,4,1)/> zu reparieren?
    //6587	Utiliser <If(Equal(IntegerParameter(3),1))><SheetFr(Item,1,IntegerParameter(2),1,1)/><Else/><Value>IntegerParameter(3)</Value> <SheetFr(Item,12,IntegerParameter(2),2,1)/></If> pour réparer <SheetFr(Item,2,IntegerParameter(1),1,1)/> de votre appareil<Indent/>? (<Value>IntegerParameter(4)</Value> possédé<If(LessThanOrEqualTo(IntegerParameter(4),1))><Else/>s</If>)
    /*6587	下記のアイテムを修理しますか？
    <Sheet(Item,IntegerParameter(1),0)/>
    消費:<Sheet(Item,IntegerParameter(2),0)/>×<Value>IntegerParameter(3)</Value>(所持数 <Value>IntegerParameter(4)</Value>)
    */

    internal static readonly string[] WorkshopRepairConfirm =
        [
            "repair",
            "下記のアイテムを修理しますか",
            "reparieren",
            "réparer",
            "要修理下列部件吗",
            "要修理下列部件嗎",
            "要修理下列元件嗎",
            "수리하시겠습니까?",
            "要修理下列組件嗎",
        ];

    // Use the components selected and <If(Equal(IntegerParameter(1),1))>the following item<Else/><Value>IntegerParameter(1)</Value> of the following items</If> to outfit and register your submersible?
    /* 6886 Das Tauchboot mit den gewählten Bauteilen registrieren?
     Verbraucht <Value>IntegerParameter(1)</Value> <If(Equal(IntegerParameter(1),1))>Exemplar<Else/>Exemplare</If> des folgenden Gegenstands:
    */
    // 6886 Utiliser les éléments choisis et <If(Equal(IntegerParameter(1),1))>l'objet suivant<Else/><Value>IntegerParameter(1)</Value> des objets suivants</If> pour équiper et enregistrer le sous-marin<Indent/>?
    /*選択したパーツアイテムと以下のアイテムを
       <Value>IntegerParameter(1)</Value>枚消費して潜水艦を登録します。
       よろしいですか？
    */

    internal static readonly string[] WorkshopRegisterConfirm =
    [
            "to outfit and register your submersible",
            "枚消費して潜水艦を登録します",
            "Das Tauchboot mit den gewählten Bauteilen registrieren",
            "pour équiper et enregistrer le sous-marin",
            "잠수함을 등록하시겠습니까",
            "张下列道具登记新的潜水艇吗",
            "張下列道具登記新的潛水艇嗎",
            //"",
            //""     (Addonsheet - 6886)
    ];

    //Your retainer will be unable to process item buyback requests once recalled. Are you sure you wish to proceed?
    //215	TEXT_CMNDEFRETAINERCALL_00010_ASK_RETURN_WITH_BUYBACK	Wenn du deinen Gehilfen wegschickst, kannst du die von ihm verkauften Gegenstände nicht mehr zurückkaufen. Möchtest du trotzdem fortfahren?
    //215	TEXT_CMNDEFRETAINERCALL_00010_ASK_RETURN_WITH_BUYBACK	Renvoyer le servant effacera la liste de rachat. Confirmer<Indent/>?

    internal static string[] WillBeUnableToProcessBuyback => field ??= [
        Svc.Data.GetExcelSheet<QuestDialogueText>(name:"custom/000/CmnDefRetainerCall_00010").GetRow(215).Value.GetText(),
        ];

    //3290	<Sheet(Item,IntegerParameter(1),0)/>×<Value>IntegerParameter(2)</Value>を、<Format(IntegerParameter(3),FF022C)/>枚の軍票と交換します。
    //よろしいですか？
    //3290	<Format(IntegerParameter(3),FF022E)/> Staatstaler gegen <If(Equal(IntegerParameter(2),1))><SheetDe(Item,1,IntegerParameter(1),1,4,1)/><Else/><Format(IntegerParameter(2),FF022E)/> <SheetDe(Item,5,IntegerParameter(1),2,4,1)/></If> eintauschen?
    //3290	Acheter <Value>IntegerParameter(2)</Value> <SheetFr(Item,12,IntegerParameter(1),IntegerParameter(2),1)/> pour <Format(IntegerParameter(3),FF05021D0103)/> sceau<If(LessThanOrEqualTo(IntegerParameter(3),1))><Else/>x</If><Indent/>?

    internal static readonly string[] GCSealExchangeConfirm = ["Exchange", "よろしいですか？", "Staatstaler gegen", "Acheter", "要交换吗", "교환하시겠습니까", "要交換嗎"];

    internal static readonly string[] DiscardItem = ["Discard", "を捨てます。", "wegwerfen", "Jeter", "确定要舍弃", "確定要捨棄"];
}

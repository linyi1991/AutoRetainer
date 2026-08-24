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
        ["About"] = "關於",
        ["Disclaimer"] = "使用提醒",
        ["Common Settings"] = "共通設定",
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
        ["Enable"] = "啟用",
        ["This is purely visual order and does not affects character processing in any way."] = "這只影響顯示順序，不會影響角色處理流程。",

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
        ["Artisan integration"] = "Artisan 整合",
        ["Use server time instead of PC time"] = "使用伺服器時間而不是電腦時間",
        ["Export without character data"] = "匯出但不含角色資料",
        ["Import and merge with character data"] = "匯入並合併角色資料",

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
        ["FPS Limiter is only active when Multi Mode is enabled"] = "FPS 限制只會在多角色模式啟用時生效",
        ["Allow extra low FPS limiter values"] = "允許更低的 FPS 限制數值",
        ["No support is provided if you enable this and run into ANY errors in Multi Mode"] = "啟用後若多角色模式發生任何錯誤，原作者不提供支援。",
        ["Here you can sort your characters. This will affect order in which they will be processed by Multi Mode as well as how they will appear in plugin interface and login overlay."] = "這裡可以排序角色。排序會影響多角色模式處理順序，也會影響插件介面與登入浮層的顯示順序。",
        ["Temporarily prevents AutoRetainer from being automatically enabled when using a Summoning Bell/Workshop Panel"] = "使用傳喚鈴/工房面板時，暫時阻止 AutoRetainer 自動啟用",
        ["Add current account"] = "新增目前帳號",
        ["Ignore other characters that have not been enabled in MultiMode"] = "忽略未啟用 MultiMode 的其他角色",
        ["If game is inactive: (requires NotificationMaster to be installed and enabled)"] = "遊戲不在前景時：（需要安裝並啟用 NotificationMaster）",
        ["Do not notify if AutoRetainer is enabled or MultiMode is running"] = "AutoRetainer 已啟用或 MultiMode 運作中時不要通知",
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
        return Tw.TryGetValue(visible, out var translated) ? translated + id : text;
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

    internal static readonly (string Normal, string GameFont) Digits = ("0123456789", "");

    internal static readonly string[] FieldExplorationNames =
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

    internal static readonly string[] HuntingVentureNames =
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

    internal static readonly string[] QuickExploration =
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

    internal static readonly string[] RetainerAskCategoryText =
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

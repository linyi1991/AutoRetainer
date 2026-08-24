using AutoRetainer.Modules.Voyage;
using Dalamud.Game;
using ECommons.GameHelpers;
using ECommons.Reflection;

namespace AutoRetainer.UI.MainWindow;
public static unsafe class TroubleshootingUI
{
    private static readonly Config EmptyConfig = new();
    public static void Draw()
    {
        ImGuiEx.TextWrapped("這個分頁會檢查常見設定問題，方便你先自行排除。");

        if(!Player.Available)
        {
            ImGuiEx.TextWrapped($"尚未登入角色時無法進行疑難排解。");
            return;
        }

        if(Data == null)
        {
            ImGuiEx.TextWrapped($"目前角色沒有資料。請開啟傳喚鈴、飛空艇/潛水艇面板，或登出一次以建立資料。");
            return;
        }

        if(!Svc.ClientState.ClientLanguage.EqualsAny(ClientLanguage.Japanese, ClientLanguage.German, ClientLanguage.French, ClientLanguage.English))
        {
            Error($"偵測到地區代理版本客戶端。AutoRetainer 原作者未針對這類 FFXIV 客戶端測試，部分或全部功能可能無法正常運作。另請留意，ottercorp 的中文 Dalamud fork 可能會收集你的電腦、角色、已使用插件與 Dalamud 設定等遙測資料。");
        }

        if(C.DontLogout)
        {
            Error("DontLogout 除錯選項已啟用");
        }

        foreach(var x in C.OfflineData)
        {
            if(x.WorkshopEnabled)
            {
                var a = x.OfflineSubmarineData.Select(x => x.Name);
                if(a.Count() > a.Distinct().Count())
                {
                    Error($"角色 {Censor.Character(x.Name, x.World)} 有重複的潛水艇名稱。潛水艇名稱必須唯一。");
                }
            }
        }

        if((C.GlobalTeleportOptions.Enabled || C.OfflineData.Any(x => x.TeleportOptionsOverride.Enabled == true)) && !Svc.PluginInterface.InstalledPlugins.Any(x => x.InternalName == "Lifestream" && x.IsLoaded))
        {
            Error("已啟用傳送，但 Lifestream 插件未安裝或未載入。AutoRetainer 無法在此設定下正常運作。請停用傳送，或安裝並載入 Lifestream。");
        }

        foreach(var x in C.SubmarineUnlockPlans)
        {
            if(x.EnforcePlan)
            {
                Info($"潛水艇解鎖計畫 {x.Name.NullWhenEmpty() ?? x.GUID} 已設為強制執行。只要還有可解鎖內容，就會覆寫個別潛水艇設定。");
            }
        }

        foreach(var x in C.SubmarineUnlockPlans)
        {
            if(x.EnforceDSSSinglePoint)
            {
                Info($"潛水艇解鎖計畫 {x.Name.NullWhenEmpty() ?? x.GUID} 設定為在深海站點只派往單一航點，因此會忽略手動設定的解鎖行為。");
            }
        }

        try
        {
            if(DalamudReflector.IsOnStaging())
            {
                Error($"偵測到非 release 的 Dalamud 分支，可能造成問題。若可行，請輸入 /xlbranch 開啟分支切換器，改成 \"release\" 後重啟遊戲。");
            }
        }
        catch(Exception e)
        {
        }

        if(Player.Available)
        {
            if(Player.CurrentWorld != Player.HomeWorld)
            {
                Error("你目前正在拜訪其他伺服器。AutoRetainer 要繼續處理此角色前，必須先回到原本伺服器。");
            }
            if(C.Blacklist.Any(x => x.CID == Player.CID))
            {
                Error("目前角色已被完整排除在 AutoRetainer 外，不會被任何流程處理。請到設定 - 排除與順序中修改。");
            }
            if(Data.ExcludeRetainer)
            {
                Error("目前角色已從僱員清單排除。請到設定 - 排除與順序中修改。");
            }
            if(Data.ExcludeWorkshop)
            {
                Error("目前角色已從飛空艇/潛水艇清單排除。請到設定 - 排除與順序中修改。");
            }
        }

        {
            var list = C.OfflineData.Where(x => x.GetAreTeleportSettingsOverriden());
            if(list.Any())
            {
                Info("部分角色使用了個別傳送設定。滑鼠移上來可查看清單。", list.Select(x => $"{x.Name}@{x.World}").Print("\n"));
            }
        }

        if(C.NoTeleportHetWhenNextToBell)
        {
            Warning("角色已在傳喚鈴旁時，會停用傳送或進入房屋/公寓。請留意房屋拆除倒數。");
        }



        if(C.AllowSimpleTeleport)
        {
            Warning("簡易傳送已啟用。它比在 Lifestream 登錄房屋更不穩定。若遇到傳送問題，建議停用此選項，並改在 Lifestream 登錄房屋。");
        }

        if(!C.EnableEntrustManager && C.AdditionalData.Any(x => x.Value.EntrustPlan != Guid.Empty))
        {
            Warning($"託管管理器目前全域停用，但部分僱員仍指定了託管計畫。這些託管計畫只會在手動操作時處理。");
        }

        if(C.ExtraDebug)
        {
            Info("詳細記錄選項已啟用，會大量寫入 log。只有在收集除錯資訊時才建議使用。");
        }

        if(C.UnsyncCompensation > -5)
        {
            Warning("時間誤差補償設定過高（>-5），可能造成問題。");
        }

        if(UIUtils.GetFPSFromMSPT(C.TargetMSPTIdle) < 10)
        {
            Warning("閒置時目標 FPS 設定過低（<10），可能造成問題。");
        }

        if(UIUtils.GetFPSFromMSPT(C.TargetMSPTRunning) < 20)
        {
            Warning("運作時目標 FPS 設定過低（<20），可能造成問題。");
        }

        if(Data?.GetIMSettings().AllowSellFromArmory == true)
        {
            Info("已允許從兵裝庫出售物品。請務必把零式裝備與絕武加入保護清單。");
        }

        {
            var list = C.OfflineData.Where(x => !x.ExcludeRetainer && !x.Enabled && x.RetainerData.Count > 0);
            if(list.Any())
            {
                Warning($"部分有僱員的角色尚未啟用僱員多角色模式。滑鼠移上來可查看清單。", list.Print("\n"));
            }
        }
        {
            var list = C.OfflineData.Where(x => !x.ExcludeRetainer && x.Enabled && x.RetainerData.Count > 0 && C.SelectedRetainers.TryGetValue(x.CID, out var rd) && !x.RetainerData.All(r => rd.Contains(r.Name)));
            if(list.Any())
            {
                Warning($"部分角色並非所有僱員都已啟用處理。滑鼠移上來可查看清單。", list.Print("\n"));
            }
        }
        {
            var list = C.OfflineData.Where(x => !x.ExcludeWorkshop && !x.WorkshopEnabled && (x.OfflineSubmarineData.Count + x.OfflineAirshipData.Count) > 0);
            if(list.Any())
            {
                Warning($"部分已登錄飛空艇/潛水艇的角色尚未啟用飛空艇/潛水艇多角色模式。滑鼠移上來可查看清單。", list.Print("\n"));
            }
        }

        {
            var list = C.OfflineData.Where(x => !x.ExcludeWorkshop && x.WorkshopEnabled && x.GetEnabledVesselsData(Internal.VoyageType.Airship).Count + x.GetEnabledVesselsData(Internal.VoyageType.Submersible).Count < Math.Min(x.OfflineAirshipData.Count + x.OfflineSubmarineData.Count, 4));
            if(list.Any())
            {
                Warning($"部分角色並非所有飛空艇/潛水艇都已啟用處理。滑鼠移上來可查看清單。", list.Print("\n"));
            }
        }

        if(C.MultiModeType != AutoRetainerAPI.Configuration.MultiModeType.Everything)
        {
            Warning($"目前 MultiMode 類型設定為 {C.MultiModeType}。這會限制 AutoRetainer 執行的功能。");
        }

        if(C.OfflineData.Any(x => x.MultiWaitForAllDeployables))
        {
            Info("部分角色已啟用 \"Wait For All Pending Deployables\"。代表這些角色會等待所有飛空艇/潛水艇返回後才處理。滑鼠移上來可查看完整角色清單。", C.OfflineData.Where(x => x.MultiWaitForAllDeployables).Select(x => $"{x.Name}@{x.World}").Print("\n"));
        }

        if(C.MultiModeWorkshopConfiguration.MultiWaitForAll)
        {
            Info("全域選項 \"Wait For Venture Completion\" 已啟用。代表所有角色都會等待飛空艇/潛水艇返回後才處理，即使該角色的個別選項已停用也一樣。");
        }

        if(C.MultiModeWorkshopConfiguration.WaitForAllLoggedIn)
        {
            Info("飛空艇/潛水艇已啟用 \"Wait even when already logged in\"。代表即使你已登入該角色，也會等所有飛空艇/潛水艇完成後才處理。");
        }

        if(C.DisableRetainerVesselReturn > 0)
        {
            if(C.DisableRetainerVesselReturn > 10)
            {
                Warning("\"Retainer venture processing cutoff\" 設定值異常偏高。當飛空艇/潛水艇即將可用時，重新派出僱員可能會明顯延遲。");
            }
            else
            {
                Info("\"Retainer venture processing cutoff\" 已啟用。當飛空艇/潛水艇即將可用時，重新派出僱員可能會延遲。");
            }
        }

        if(C.MultiModeRetainerConfiguration.MultiWaitForAll)
        {
            Info("\"Wait For Venture Completion\" 已啟用。代表 AutoRetainer 會等待該角色所有僱員探索完成後，才登入並處理。");
        }

        if(C.MultiModeRetainerConfiguration.WaitForAllLoggedIn)
        {
            Info("僱員已啟用 \"Wait even when already logged in\"。代表即使你已登入該角色，也會等所有僱員探索完成後才處理。");
        }

        {
            var manualList = new List<string>();
            var deletedList = new List<string>();
            foreach(var x in C.OfflineData)
            {
                foreach(var ret in x.RetainerData)
                {
                    var planId = Utils.GetAdditionalData(x.CID, ret.Name).EntrustPlan;
                    var plan = C.EntrustPlans.FirstOrDefault(s => s.Guid == planId);
                    if(plan != null && plan.ManualPlan) manualList.Add($"{Censor.Character(x.Name)} - {Censor.Retainer(ret.Name)}");
                    if(plan == null && planId != Guid.Empty) deletedList.Add($"{Censor.Character(x.Name)} - {Censor.Retainer(ret.Name)}");
                }
            }
            if(manualList.Count > 0)
            {
                Info("部分僱員使用手動託管計畫。這些計畫不會在重新派出探索後自動處理，只會在 overlay 按鈕手動點擊時執行。滑鼠移上來可查看清單。", manualList.Print("\n"));
            }
            if(deletedList.Count > 0)
            {
                Warning("部分僱員指定的託管計畫已被刪除。使用已刪除託管計畫的僱員不會託管任何物品。滑鼠移上來可查看清單。", deletedList.Print("\n"));
            }
        }

        if(C.No2ndInstanceNotify)
        {
            Info("你已啟用 \"Do not warn about second game instance running from same directory\"。使用相同 Dalamud 目錄啟動第二個遊戲實例時，AutoRetainer 會自動跳過載入。");
        }

        if(Svc.PluginInterface.InstalledPlugins.Any(x => x.InternalName == "SimpleTweaksPlugin" && x.IsLoaded))
        {
            Info("偵測到 Simple Tweaks 插件。任何與僱員或潛水艇相關的 tweak 都可能影響 AutoRetainer。請確認相關設定不會干擾 AutoRetainer。");
        }

        if(Svc.PluginInterface.InstalledPlugins.Any(x => x.InternalName == "PandorasBox" && x.IsLoaded))
        {
            Info("偵測到 Pandora's Box 插件。AutoRetainer 啟用時若自動使用動作，可能影響 AutoRetainer。請確認 Pandora's Box 不會在 AutoRetainer 運作時自動使用動作。");
        }

        if(Svc.PluginInterface.InstalledPlugins.Any(x => x.InternalName == "Automaton" && x.IsLoaded))
        {
            Info("偵測到 Automaton 插件。AutoRetainer 啟用時若自動使用動作或自動輸入數字，可能影響 AutoRetainer。請確認 Automaton 不會在 AutoRetainer 運作時自動操作。");
        }

        if(Svc.PluginInterface.InstalledPlugins.Any(x => x.InternalName == "RotationSolver" && x.IsLoaded))
        {
            Info("偵測到 RotationSolver 插件。AutoRetainer 啟用時若自動使用動作，可能影響 AutoRetainer。請確認 RotationSolver 不會在 AutoRetainer 運作時自動使用動作。");
        }

        if(Svc.PluginInterface.InstalledPlugins.Any(x => x.InternalName.StartsWith("BossMod") && x.IsLoaded))
        {
            Info("偵測到 BossMod 插件。AutoRetainer 啟用時若自動使用動作，可能影響 AutoRetainer。請確認 BossMod 不會在 AutoRetainer 運作時自動使用動作。");
        }

        ImGui.Separator();
        ImGuiEx.TextWrapped("專家設定會改變原作者預期的行為。若遇到問題，請先確認不是專家設定錯誤造成。");
        CheckExpertSetting("Action on accessing retainer bell if no ventures available", nameof(C.OpenBellBehaviorNoVentures));
        CheckExpertSetting("Action on accessing retainer bell if any ventures available", nameof(C.OpenBellBehaviorWithVentures));
        CheckExpertSetting("Task completion behavior after accessing bell", nameof(C.TaskCompletedBehaviorAccess));
        CheckExpertSetting("Task completion behavior after manual enabling", nameof(C.TaskCompletedBehaviorManual));
        CheckExpertSetting("Stay in retainer menu if there are retainers to finish ventures within 5 minutes or less", nameof(C.Stay5));
        CheckExpertSetting("Auto-disable plugin when closing retainer list", nameof(C.AutoDisable));
        CheckExpertSetting("Do not show plugin status icons", nameof(C.HideOverlayIcons));
        CheckExpertSetting("Display multi mode type selector", nameof(C.DisplayMMType));
        CheckExpertSetting("Display deployables checkbox in workshop", nameof(C.ShowDeployables));
        CheckExpertSetting("Enable bailout module", nameof(C.EnableBailout));
        CheckExpertSetting("Timeout before AutoRetainer will attempt to unstuck, seconds", nameof(C.BailoutTimeout));
        CheckExpertSetting("Disable sorting and collapsing/expanding", nameof(C.NoCurrentCharaOnTop));
        CheckExpertSetting("Show MultiMode checkbox on plugin UI bar", nameof(C.MultiModeUIBar));
        CheckExpertSetting("Retainer menu delay, seconds", nameof(C.RetainerMenuDelay));
        CheckExpertSetting("Do not error check venture planner", nameof(C.NoErrorCheckPlanner2));
        CheckExpertSetting("Upon activating Multi Mode, attempt to enter nearby house", nameof(C.MultiHETOnEnable));
        CheckExpertSetting("Artisan integration", nameof(C.ArtisanIntegration));
        CheckExpertSetting("Use server time instead of PC time", nameof(C.UseServerTime));
    }

    private static void Error(string message, string tooltip = null)
    {
        ImGui.PushFont(UiBuilder.IconFont);
        ImGuiEx.Text(EColor.RedBright, "\uf057");
        ImGui.PopFont();
        if(tooltip != null) ImGuiEx.Tooltip(tooltip);
        ImGui.SameLine();
        ImGuiEx.TextWrapped(EColor.RedBright, Lang.T(message));
        if(tooltip != null) ImGuiEx.Tooltip(tooltip);
    }

    private static void Warning(string message, string tooltip = null)
    {
        ImGui.PushFont(UiBuilder.IconFont);
        ImGuiEx.Text(EColor.OrangeBright, "\uf071");
        ImGui.PopFont();
        if(tooltip != null) ImGuiEx.Tooltip(tooltip);
        ImGui.SameLine();
        ImGuiEx.TextWrapped(EColor.OrangeBright, Lang.T(message));
        if(tooltip != null) ImGuiEx.Tooltip(tooltip);
    }

    private static void Info(string message, string tooltip = null)
    {
        ImGui.PushFont(UiBuilder.IconFont);
        ImGuiEx.Text(EColor.YellowBright, "\uf05a");
        ImGui.PopFont();
        if(tooltip != null) ImGuiEx.Tooltip(tooltip);
        ImGui.SameLine();
        ImGuiEx.TextWrapped(EColor.YellowBright, Lang.T(message));
        if(tooltip != null) ImGuiEx.Tooltip(tooltip);
    }

    private static void CheckExpertSetting(string setting, string nameOfSetting)
    {
        var original = EmptyConfig.GetFoP(nameOfSetting);
        var current = C.GetFoP(nameOfSetting);
        if(!original.Equals(current))
        {
            Info($"專家設定「{Lang.T(setting)}」與預設值不同", $"預設值為「{original}」，目前為「{current}」。");
        }
    }
}

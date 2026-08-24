using AutoRetainer.Modules.Voyage;
using AutoRetainer.UI.MainWindow.MultiModeTab;
using AutoRetainerAPI;
using AutoRetainerAPI.Configuration;
using Dalamud.Interface.Components;
using ECommons.Configuration;
using ECommons.Funding;
using NightmareUI;

namespace AutoRetainer.UI.MainWindow;

internal unsafe class AutoRetainerWindow : Window
{
    private TitleBarButton LockButton;

    public AutoRetainerWindow() : base($"")
    {
        PatreonBanner.IsOfficialPlugin = () => true;
        LockButton = new()
        {
            Click = OnLockButtonClick,
            Icon = C.PinWindow ? FontAwesomeIcon.Lock : FontAwesomeIcon.LockOpen,
            IconOffset = new(3, 2),
            ShowTooltip = () => ImGui.SetTooltip(Lang.T("Lock window position and size")),
        };
        SizeConstraints = new()
        {
            MinimumSize = new(250, 100),
            MaximumSize = new(9999, 9999)
        };
        P.WindowSystem.AddWindow(this);
        AllowPinning = false;
        TitleBarButtons.Add(new()
        {
            Click = (m) => { if(m == ImGuiMouseButton.Left) S.NeoWindow.IsOpen = true; },
            Icon = FontAwesomeIcon.Cog,
            IconOffset = new(2, 2),
            ShowTooltip = () => ImGui.SetTooltip(Lang.T("Open settings window")),
        });
        TitleBarButtons.Add(LockButton);
    }

    private Action<string> SomeAction;

    private void OnLockButtonClick(ImGuiMouseButton m)
    {
        SomeAction += (s) => { };
        SomeAction -= (s) => { };
        if(m == ImGuiMouseButton.Left)
        {
            C.PinWindow = !C.PinWindow;
            LockButton.Icon = C.PinWindow ? FontAwesomeIcon.Lock : FontAwesomeIcon.LockOpen;
        }
    }

    public override void PreDraw()
    {
        var prefix = SchedulerMain.PluginEnabled ? $" [{SchedulerMain.Reason}]" : "";
        var tokenRem = TimeSpan.FromMilliseconds(Utils.GetRemainingSessionMiliSeconds());
        WindowName = $"{P.Name} {P.GetType().Assembly.GetName().Version}{prefix} | {FormatToken(tokenRem)}###AutoRetainer";
        if(C.PinWindow)
        {
            ImGuiHelpers.SetNextWindowPosRelativeMainViewport(C.WindowPos);
            ImGui.SetNextWindowSize(C.WindowSize);
        }
    }

    private string FormatToken(TimeSpan time)
    {
        if(time.TotalMilliseconds > 0)
        {
            if(time.Days > 0)
            {
                return $"工作階段將於 {time.Days} 天後到期" + (time.Hours > 0 ? $" {time.Hours} 小時" : "");
            }
            else
            {
                if(time.Hours > 0)
                {
                    return $"工作階段將於 {time.Hours} 小時後到期";
                }
                else
                {
                    return "工作階段將於 1 小時內到期";
                }
            }
        }
        else
        {
            return "工作階段已過期";
        }
    }
    public override void Draw()
    {
        //ImGuiEx.Text(GradientColor.Get(EColor.RedBright, EColor.YellowBright), "This version MUST NOT BE RUNNING UNATTENDED.");
        try
        {
            if(!C.AcceptedDisclamer)
            {
                new NuiBuilder()
                    .Section("Disclaimer")
                    .TextWrapped(ImGuiColors.DalamudYellow, "請注意：嚴禁將 AutoRetainer 用於 RMT 或其他商業用途。")
                    .TextWrapped(ImGuiColors.DalamudRed, "為避免帳號風險，使用 AutoRetainer 時請遵守以下規則：")
                    .TextWrapped("1. 不要在遊戲聊天中承認或提及你正在使用 AutoRetainer。")
                    .TextWrapped("2. 不要長時間無人看管地使用 AutoRetainer。")
                    .TextWrapped("3. 每日遊玩加 AutoRetainer 使用時間請勿超過 16 小時；僱員/潛水艇檢查流程之間應保留休息間隔。")
                    .TextWrapped("4. 若其他玩家以交易或聊天方式測試你，不要回覆，請直接加入黑名單。")
                    .TextWrapped("5. 若被 GM 詢問，請一律表示所有操作皆為手動，且不要承認使用插件。")
                    .TextWrapped("不遵守這些準則可能使帳號承擔風險。")
                    .TextWrapped(GradientColor.Get(ImGuiColors.DalamudYellow, ImGuiColors.DalamudRed), "不得將 AutoRetainer 用於現金交易或其他商業用途。若用於這些目的，將不提供任何支援。")
                    .Widget(() =>
                    {
                        if(ImGuiEx.IconButtonWithText(FontAwesomeIcon.Check, "同意並繼續"))
                        {
                            C.AcceptedDisclamer = true;
                            EzConfig.Save();
                        }
                    })
                    .Draw();
                return;
            }
            var e = SchedulerMain.PluginEnabledInternal;
            var disabled = MultiMode.Active && !ImGui.GetIO().KeyCtrl;

            if(disabled)
            {
                ImGui.BeginDisabled();
            }
            if(ImGui.Checkbox($"啟用 {P.Name}", ref e))
            {
                P.WasEnabled = false;
                if(e)
                {
                    SchedulerMain.EnablePlugin(PluginEnableReason.Auto);
                }
                else
                {
                    SchedulerMain.DisablePlugin();
                }
            }
            if(C.ShowDeployables && (VoyageUtils.Workshops.Contains(Svc.ClientState.TerritoryType) || VoyageScheduler.Enabled))
            {
                ImGui.SameLine();
                ImGui.Checkbox($"飛空艇/潛水艇", ref VoyageScheduler.Enabled);
            }
            if(disabled)
            {
                ImGui.EndDisabled();
                ImGuiComponents.HelpMarker($"此選項由多角色模式控制。按住 CTRL 可暫時覆寫。");
            }

            if(P.WasEnabled)
            {
                ImGui.SameLine();
                ImGuiEx.Text(GradientColor.Get(ImGuiColors.DalamudGrey, ImGuiColors.DalamudGrey3, 500), $"已暫停");
            }

            ImGui.SameLine();
            if(ImGui.Checkbox("多角色", ref MultiMode.Enabled))
            {
                MultiMode.OnMultiModeEnabled();
            }
            Utils.DrawLifestreamAvailabilityIndicator();
            if(C.ShowNightMode)
            {
                ImGui.SameLine();
                if(ImGui.Checkbox("夜間", ref C.NightMode))
                {
                    MultiMode.BailoutNightMode();
                }
            }
            if(C.DisplayMMType)
            {
                ImGui.SameLine();
                ImGuiEx.SetNextItemWidthScaled(100f);
                ImGuiEx.EnumCombo("##mode", ref C.MultiModeType);
            }
            if(C.CharEqualize && MultiMode.Enabled)
            {
                ImGui.SameLine();
                if(ImGui.Button("重置計數"))
                {
                    MultiMode.CharaCnt.Clear();
                }
            }

            Svc.PluginInterface.GetIpcProvider<object>(ApiConsts.OnMainControlsDraw).SendMessage();

            if(IPC.Suppressed)
            {
                ImGuiEx.Text(ImGuiColors.DalamudRed, $"插件操作目前被其他插件暫停。");
                ImGui.SameLine();
                if(ImGui.SmallButton("取消"))
                {
                    IPC.Suppressed = false;
                }
            }

            if(P.TaskManager.IsBusy)
            {
                ImGui.SameLine();
                if(ImGui.Button($"中止 {P.TaskManager.NumQueuedTasks} 個任務"))
                {
                    P.TaskManager.Abort();
                }
            }

            PatreonBanner.DrawRight();
            ImGuiEx.EzTabBar("tabbar", PatreonBanner.Text,
                            ("僱員", MultiModeUI.Draw, null, true),
                            ("飛空艇/潛水艇", WorkshopUI.Draw, null, true),
                            ("疑難排解", TroubleshootingUI.Draw, null, true),
                            ("統計", DrawStats, null, true),
                            ("關於", CustomAboutTab.Draw, null, true)
                            );
            if(!C.PinWindow)
            {
                C.WindowPos = ImGui.GetWindowPos();
                C.WindowSize = ImGui.GetWindowSize();
            }
        }
        catch(Exception e)
        {
            ImGuiEx.TextWrapped(e.ToStringFull());
        }
    }

    private void DrawStats()
    {
        NuiTools.ButtonTabs([[C.RecordStats ? new("Ventures", S.VentureStats.DrawVentures) : null, new("Gil", S.GilDisplay.Draw), new("FC Data", S.FCData.Draw)]]);
    }

    public override void OnClose()
    {
        EzConfig.Save();
        S.VentureStats.Data.Clear();
        MultiModeUI.JustRelogged = false;
    }

    public override void OnOpen()
    {
        MultiModeUI.JustRelogged = true;
    }
}

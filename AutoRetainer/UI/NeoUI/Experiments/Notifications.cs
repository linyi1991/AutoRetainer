namespace AutoRetainer.UI.NeoUI.Experiments;
public class Notifications : ExperimentUIEntry
{
    public override void Draw()
    {
        ImGui.Checkbox($"若任一僱員完成探索，顯示浮層通知", ref C.NotifyEnableOverlay);
        ImGui.Checkbox($"副本或戰鬥中不要顯示浮層", ref C.NotifyCombatDutyNoDisplay);
        ImGui.Checkbox($"包含其他角色", ref C.NotifyIncludeAllChara);
        ImGui.Checkbox($"忽略未啟用 MultiMode 的其他角色", ref C.NotifyIgnoreNoMultiMode);
        ImGui.Checkbox($"在遊戲聊天顯示通知", ref C.NotifyDisplayInChatX);
        ImGuiEx.Text($"遊戲不在前景時：（需要安裝並啟用 NotificationMaster）");
        ImGui.Checkbox($"僱員可處理時送出桌面通知", ref C.NotifyDeskopToast);
        ImGui.Checkbox($"閃爍工作列", ref C.NotifyFlashTaskbar);
        ImGui.Checkbox($"AutoRetainer 已啟用或 MultiMode 運作中時不要通知", ref C.NotifyNoToastWhenRunning);
    }
}

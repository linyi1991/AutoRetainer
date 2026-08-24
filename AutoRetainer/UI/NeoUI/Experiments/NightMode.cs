namespace AutoRetainer.UI.NeoUI.Experiments;

internal class NightMode : ExperimentUIEntry
{
    public override string Name => "Night Mode";
    public override void Draw()
    {
        ImGuiEx.TextWrapped($"夜間模式：\n" +
                $"- 會強制啟用登入畫面等待選項\n" +
                $"- 會強制套用內建 FPS 限制\n" +
                $"- 遊戲不在前景且等待時，會限制到 0.2 FPS\n" +
                $"- 看起來可能像遊戲卡住，重新切回遊戲後請給它最多 5 秒喚醒\n" +
                $"- 預設夜間模式只處理飛空艇/潛水艇\n" +
                $"- 關閉夜間模式後，斷線恢復管理器會啟動並重新登入遊戲。");
        if(ImGui.Checkbox("啟用夜間模式", ref C.NightMode)) MultiMode.BailoutNightMode();
        ImGui.Checkbox("顯示夜間模式勾選框", ref C.ShowNightMode);
        ImGui.Checkbox("夜間模式處理僱員", ref C.NightModeRetainers);
        ImGui.Checkbox("夜間模式處理飛空艇/潛水艇", ref C.NightModeDeployables);
        ImGui.Checkbox("保留夜間模式狀態", ref C.NightModePersistent);
        ImGui.Checkbox("關機指令改為啟用夜間模式，不直接關閉遊戲", ref C.ShutdownMakesNightMode);
    }
}

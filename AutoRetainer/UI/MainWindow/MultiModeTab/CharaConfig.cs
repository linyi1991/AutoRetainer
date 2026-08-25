using AutoRetainerAPI.Configuration;
using Dalamud.Interface.Components;
using PunishLib.ImGuiMethods;

namespace AutoRetainer.UI.MainWindow.MultiModeTab;
public class CharaConfig
{
    public static void Draw(OfflineCharacterData data, bool isRetainer)
    {
        ImGui.PushID(data.CID.ToString());
        SharedUI.DrawMultiModeHeader(data);
        var b = new NuiBuilder()

        .Section("General Character Specific Settings")
        .Widget(() =>
        {
            SharedUI.DrawServiceAccSelector(data);
            SharedUI.DrawPreferredCharacterUI(data);
        });
        if(isRetainer)
        {
            b = b.Section("Retainers").Widget(() =>
            {
                ImGuiEx.Text($"自動軍票稀有品繳納：");
                if(!AutoGCHandin.Operation)
                {
                    ImGuiEx.SetNextItemWidthScaled(200f);
                    ImGuiEx.EnumCombo("##gcHandin", ref data.GCDeliveryType, Lang.GCDeliveryTypeNames);
                }
                else
                {
                    ImGuiEx.Text($"目前無法變更");
                }
            });
        }
        else
        {
            b = b.Section("Deployables").Widget(() =>
            {
                ImGui.Checkbox($"Wait For Voyage Completion", ref data.MultiWaitForAllDeployables);
            ImGuiComponents.HelpMarker("""此設定類似全域選項，但只套用到個別角色。啟用後，AutoRetainer 會等所有飛空艇/潛水艇返航後才登入此角色。若你因其他原因已登入此角色，仍會重新派出已完成的潛水艇，除非全域設定「即使已登入也等待」也已啟用。""");
            });
        }
        b = b.Section("Teleport overrides", data.GetAreTeleportSettingsOverriden() ? ImGui.GetStyle().Colors[(int)ImGuiCol.FrameBg] with { X = 1f } : null, true)
        .Widget(() =>
        {
            ImGuiEx.Text($"你可以為每個角色覆寫傳送設定。");
            bool? demo = null;
            ImGuiEx.Checkbox("Options marked with this marker will use values from global configuration", ref demo);
            ImGuiEx.Checkbox("Enabled", ref data.TeleportOptionsOverride.Enabled);
            ImGui.Indent();
            ImGuiEx.Checkbox("Teleport for retainers...", ref data.TeleportOptionsOverride.Retainers);
            ImGui.Indent();
            ImGuiEx.Checkbox("...to private house", ref data.TeleportOptionsOverride.RetainersPrivate);
            ImGuiEx.Checkbox("...到共享房屋", ref data.TeleportOptionsOverride.RetainersShared);
            ImGuiEx.Checkbox("...to free company house", ref data.TeleportOptionsOverride.RetainersFC);
            ImGuiEx.Checkbox("...to apartment", ref data.TeleportOptionsOverride.RetainersApartment);
            ImGui.Text("If all above are disabled or fail, will be teleported to inn.");
            ImGui.Unindent();
            ImGuiEx.Checkbox("Teleport to free company house for deployables", ref data.TeleportOptionsOverride.Deployables);
            ImGui.Unindent(); 
        }).Draw();
        SharedUI.DrawExcludeReset(data);
        ImGui.PopID();
    }
}

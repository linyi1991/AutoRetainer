global using OverlayTextData = (System.Numerics.Vector2 Curpos, (bool Warning, string Text)[] Texts);
using AutoRetainerAPI.Configuration;
using ECommons.GameHelpers;
using ECommons.Interop;
using Lumina.Excel.Sheets;

namespace AutoRetainer.UI;

internal static class UIUtils
{
    public static void DrawSortableEnumList<T>(string id, List<T> list) where T : struct, Enum
    {
        ref var dragDrop = ref Ref<ImGuiEx.RealtimeDragDrop<T>>.Get($"dsel{id}", () => new($"dsel{id}", x => x.ToString()));
        ImGui.PushID(id);
        if(ImGui.BeginCombo("##addNew", Lang.T("Add Entries..."), ImGuiComboFlags.HeightLarge))
        {
            foreach(var x in Enum.GetValues<T>())
            {
                if(!list.Contains(x))
                {
                    if(ImGui.Selectable(Lang.T(x.ToStringEx()), false, ImGuiSelectableFlags.DontClosePopups))
                    {
                        list.Add(x);
                    }
                }
            }
            ImGui.EndCombo();
        }
        dragDrop.Begin();
        for(var i = 0; i < list.Count; i++)
        {
            var x = list[i];
            ImGui.PushID(x.ToString());
            dragDrop.DrawButtonDummy(x, list, i);
            ImGui.SameLine();
            if(ImGuiEx.IconButton(FontAwesomeIcon.Trash))
            {
                new TickScheduler(() => list.Remove(x));
            }
            ImGui.SameLine();
            ImGuiEx.Text(Lang.T(x.ToStringEx()));
            ImGui.PopID();
        }
        dragDrop.End();
        ImGui.PopID();
    }

    public static string ToStringEx<T>(this T obj) where T : Enum
    {
        return obj.ToString().Replace('_', ' ');
    }

    public static bool PushColIfPreferredCurrent(this OfflineCharacterData data)
    {
        var normalColor = Player.CID == data.CID ? EColor.CyanBright : ImGui.GetStyle().Colors[(int)ImGuiCol.Text];
        if(data.Preferred)
        {
            ImGui.PushStyleColor(ImGuiCol.Text, GradientColor.Get(normalColor, ImGuiColors.ParsedGreen));
            return true;
        }
        else
        {
            ImGui.PushStyleColor(ImGuiCol.Text, normalColor);
            return true;
        }
    }

    public static void DrawSearch()
    {
        if(C.OfflineData.Any(x => x.WorkshopEnabled || x.Enabled))
        {
            if(!Utils.IsLifestreamInstalled())
            {
                Utils.DrawLifestreamWarning("Multi Mode");
            }
        }
        if(!C.NoCharaSearch)
        {
            ImGuiEx.SetNextItemFullWidth();
            ImGui.InputTextWithHint("##search", "搜尋角色...", ref Ref<string>.Get("SearchChara"), 50);
        }
    }

    public static void DrawDCV(this OfflineCharacterData data)
    {
        if(data.WorldOverride != null)
        {
            ImGui.PushFont(UiBuilder.IconFont);
            ImGuiEx.Text("\uf0ac");
            ImGui.PopFont();
            if(ImGuiEx.HoveredAndClicked("正在拜訪其他資料中心。右鍵可清除此狀態。", ImGuiMouseButton.Right))
            {
                data.WorldOverride = null;
            }
            ImGui.SameLine();
        }
    }

    public static void DrawTeleportIcons(ulong cid)
    {
        var offlineData = C.OfflineData.FirstOrDefault(x => x.CID == cid);
        if(offlineData == null) return;
        var data = S.LifestreamIPC.GetHousePathData(cid);
        if(offlineData.GetAllowFcTeleportForSubs() || offlineData.GetAllowFcTeleportForRetainers())
        {
            string error = null;
            if(data.FC == null)
            {
                error = "部隊房屋尚未在 Lifestream 登錄";
            }
            else if(data.FC.PathToEntrance.Count == 0)
            {
                error = "部隊房屋已在 Lifestream 登錄，但尚未設定入口路徑";
            }
            ImGui.PushFont(UiBuilder.IconFont);
            ImGuiEx.Text(error == null ? null : ImGuiColors.DalamudGrey3, "\uf1ad");
            ImGui.PopFont();
            ImGuiEx.Tooltip(error ?? $"部隊房屋已在 Lifestream 登錄並設定入口路徑。重新派出飛空艇/潛水艇時會傳送到部隊房屋；若有啟用，也會為重新派出僱員傳送到部隊房屋。\n地址：{Svc.Data.GetExcelSheet<Aetheryte>().GetRowOrDefault((uint)data.FC.ResidentialDistrict)?.Territory.Value.PlaceNameRegion.Value.Name}，第 {data.FC.Ward + 1} 區，第 {data.FC.Plot + 1} 號");
            ImGui.SameLine(0, 3);
        }
        if(offlineData.GetAllowPrivateTeleportForRetainers())
        {
            string error = null;
            if(data.Private == null)
            {
                error = "個人房屋尚未在 Lifestream 登錄。";
            }
            else if(data.Private.PathToEntrance.Count == 0)
            {
                error = "個人房屋已在 Lifestream 登錄，但尚未設定入口路徑。";
            }
            ImGui.PushFont(UiBuilder.IconFont);
            ImGuiEx.Text(error == null ? null : ImGuiColors.DalamudGrey3, "\ue1b0");
            ImGui.PopFont();
            ImGuiEx.Tooltip(error ?? $"個人房屋已在 Lifestream 登錄並設定入口路徑。重新派出僱員時會傳送到個人房屋。\n地址：{Svc.Data.GetExcelSheet<Aetheryte>().GetRowOrDefault((uint)data.Private.ResidentialDistrict)?.Territory.Value.PlaceNameRegion.Value.Name}，第 {data.Private.Ward + 1} 區，第 {data.Private.Plot + 1} 號");
            ImGui.SameLine(0, 3);
        }
        if(offlineData.GetAllowSharedTeleportForRetainers())
        {
            string error = null;
            var black = false;
            var sharedData = S.LifestreamIPC.GetSharedHousePathData();
            if(Player.CID == offlineData.CID && Player.IsInHomeWorld)
            {
                if(sharedData == null)
                {
                    error = "共享房屋尚未在 Lifestream 登錄。";
                }
                else if(sharedData.PathToEntrance.Count == 0)
                {
                    error = "共享房屋已在 Lifestream 登錄，但尚未設定入口路徑。";
                }
            }
            else
            {
                error = "只有玩家已登入時才能顯示共享房屋資訊";
                black = true;
            }
            ImGui.PushFont(UiBuilder.IconFont);
            ImGuiEx.Text(error == null ? null : black?ImGuiColors.DalamudGrey2:ImGuiColors.DalamudGrey3, black ? "\ue4fe" : "\uf004");
            ImGui.PopFont();
            ImGuiEx.Tooltip(error ?? $"共享房屋已在 Lifestream 登錄並設定入口路徑。重新派出僱員時會傳送到共享房屋。\n地址：{Svc.Data.GetExcelSheet<Aetheryte>().GetRowOrDefault((uint)sharedData.ResidentialDistrict)?.Territory.Value.PlaceNameRegion.Value.Name}，第 {sharedData.Ward + 1} 區，第 {sharedData.Plot + 1} 號");
            ImGui.SameLine(0, 3);
        }
    }

    public static void DrawOverlayTexts(List<OverlayTextData> overlayTexts, ref float statusTextWidth)
    {
        if(overlayTexts.Count > 0)
        {
            var maxSizes = new float[overlayTexts[0].Texts.Length];
            for(var i = 0; i < maxSizes.Length; i++)
            {
                maxSizes[i] = overlayTexts.Select(x => ImGui.CalcTextSize(x.Texts[i].Text).X).Max();
            }
            foreach(var x in overlayTexts)
            {
                var cur = ImGui.GetCursorPos();
                for(var i = x.Texts.Length - 1; i >= 0; i--)
                {
                    var width = maxSizes[i..].Sum() + (maxSizes[i..].Length - 1) * ImGui.CalcTextSize("      ").X;
                    ImGui.SetCursorPos(new(x.Curpos.X - width, x.Curpos.Y));
                    if(statusTextWidth < width) statusTextWidth = width;
                    ImGuiEx.Text(x.Texts[i].Warning ? ImGuiColors.DalamudOrange : null, x.Texts[i].Text);
                }
                ImGui.SetCursorPos(cur);
            }
        }
    }

    public static float CollapsingHeaderSpacingsWidth => ImGui.GetStyle().FramePadding.X * 2f + ImGui.GetStyle().ItemSpacing.X * 2 + ImGui.CalcTextSize("▲...").X;

    public static string GetCutCharaString(this OfflineCharacterData data, float statusTextWidth)
    {
        var chstr = Censor.Character(data.Name, data.World);
        var mod = false;
        while(ImGui.CalcTextSize(chstr).X > ImGui.GetContentRegionAvail().X - statusTextWidth - UIUtils.CollapsingHeaderSpacingsWidth && chstr.Length > 5)
        {
            mod = true;
            chstr = chstr[0..^1];
        }
        if(mod) chstr += "...";
        return chstr;
    }

    internal static void SliderIntFrameTimeAsFPS(string name, ref int frameTime, int min = 1)
    {
        var fps = 60;
        if(frameTime != 0)
        {
            fps = GetFPSFromMSPT(frameTime);
        }
        ImGuiEx.SliderInt(name, ref fps, min, 60, fps == 60 ? "無限制" : null, ImGuiSliderFlags.AlwaysClamp);
        frameTime = fps == 60 ? 0 : (int)(1000f / fps);
    }

    public static int GetFPSFromMSPT(int frameTime)
    {
        return frameTime == 0 ? 60 : (int)(1000f / frameTime);
    }

    internal static void QRA(string text, ref LimitedKeys key)
    {
        if(DrawKeybind(text, ref key))
        {
            P.quickSellItems.Toggle();
        }
        ImGui.SameLine();
        ImGuiEx.Text(Lang.T("+ right click"));
    }

    private static string KeyInputActive = null;
    internal static bool DrawKeybind(string text, ref LimitedKeys key)
    {
        var ret = false;
        ImGui.PushID(text);
        ImGuiEx.Text($"{text}:");
        ImGui.Dummy(new(20, 1));
        ImGui.SameLine();
        ImGuiEx.SetNextItemWidthScaled(200f);
        if(ImGui.BeginCombo("##inputKey", key == LimitedKeys.None ? Lang.T("None") : $"{key}", ImGuiComboFlags.HeightLarge))
        {
            if(text == KeyInputActive)
            {
                ImGuiEx.Text(ImGuiColors.DalamudYellow, Lang.T("Now press new key..."));
                foreach(var x in Enum.GetValues<LimitedKeys>())
                {
                    if(IsKeyPressed(x))
                    {
                        KeyInputActive = null;
                        key = x;
                        ret = true;
                        break;
                    }
                }
            }
            else
            {
                if(ImGui.Selectable(Lang.T("Auto-detect new key"), false, ImGuiSelectableFlags.DontClosePopups))
                {
                    KeyInputActive = text;
                }
                ImGuiEx.Text(Lang.T("Select key manually:"));
                ImGuiEx.SetNextItemFullWidth();
                ImGuiEx.EnumCombo("##selkeyman", ref key);
            }
            ImGui.EndCombo();
        }
        else
        {
            if(text == KeyInputActive)
            {
                KeyInputActive = null;
            }
        }
        if(key != LimitedKeys.None)
        {
            ImGui.SameLine();
            if(ImGuiEx.IconButton(FontAwesomeIcon.Trash))
            {
                key = LimitedKeys.None;
                ret = true;
            }
        }
        ImGui.PopID();
        return ret;
    }
}

using AutoRetainer.Modules.Voyage;
using AutoRetainer.Modules.Voyage.VoyageCalculator;
using AutoRetainerAPI.Configuration;
using ECommons.GameHelpers;
using ECommons.Throttlers;
using FFXIVClientStructs.FFXIV.Client.Game;
using Lumina.Excel.Sheets;
using Newtonsoft.Json;

namespace AutoRetainer.UI.Windows;

internal unsafe class SubmarineUnlockPlanUI : Window
{
    internal string SelectedPlanGuid = Guid.Empty.ToString();
    internal string SelectedPlanName => VoyageUtils.GetSubmarineUnlockPlanByGuid(SelectedPlanGuid)?.Name ?? Lang.T("No or unknown plan selected");
    internal SubmarineUnlockPlan SelectedPlan => VoyageUtils.GetSubmarineUnlockPlanByGuid(SelectedPlanGuid);

    public SubmarineUnlockPlanUI() : base(Lang.T("Submersible Voyage Unlockable Planner"))
    {
        P.WindowSystem.AddWindow(this);
    }

    internal Dictionary<uint, bool> RouteUnlockedCache = [];
    internal Dictionary<uint, bool> RouteExploredCache = [];
    internal int NumUnlockedSubs = 0;

    public static readonly string DrawButtonText = Lang.T("Open Submarine Unlock Plan Editor");
    public static void DrawButton()
    {
        if(ImGuiEx.IconButtonWithText(FontAwesomeIcon.LockOpen, DrawButtonText))
        {
            P.SubmarineUnlockPlanUI.IsOpen = true;
        }
    }

    internal bool IsMapUnlocked(uint map, bool bypassCache = false)
    {
        if(!IsSubDataAvail()) return false;
        var throttle = $"Voyage.MapUnlockedCheck.{map}";
        if(!bypassCache && RouteUnlockedCache.TryGetValue(map, out var val) && !EzThrottler.Check(throttle))
        {
            return val;
        }
        else
        {
            EzThrottler.Throttle(throttle, 2500, true);
            RouteUnlockedCache[map] = HousingManager.IsSubmarineExplorationUnlocked((byte)map);
            return RouteUnlockedCache[map];
        }
    }

    internal bool IsMapExplored(uint map, bool bypassCache = false)
    {
        if(!IsSubDataAvail()) return false;
        var throttle = $"Voyage.MapExploredCheck.{map}";
        if(!bypassCache && RouteExploredCache.TryGetValue(map, out var val) && !EzThrottler.Check(throttle))
        {
            return val;
        }
        else
        {
            EzThrottler.Throttle(throttle, 2500, true);
            RouteExploredCache[map] = HousingManager.IsSubmarineExplorationExplored((byte)map);
            return RouteExploredCache[map];
        }
    }

    internal int? GetNumUnlockedSubs()
    {
        if(!IsSubDataAvail()) return null;
        NumUnlockedSubs = 1 + Unlocks.PointToUnlockPoint.Where(x => x.Value.Sub).Where(x => IsMapExplored(x.Key)).Count();
        return NumUnlockedSubs;
    }

    internal bool IsSubDataAvail()
    {
        if(HousingManager.Instance()->WorkshopTerritory == null) return false;
        if(HousingManager.Instance()->WorkshopTerritory->Submersible.Data.Length == 0) return false;
        if(HousingManager.Instance()->WorkshopTerritory->Submersible.Data[0].Name[0] == 0) return false;
        return true;
    }

    internal int GetAmountOfOtherPlanUsers(string guid)
    {
        var i = 0;
        C.OfflineData.Where(x => x.CID != Player.CID).Each(x => i += x.AdditionalSubmarineData.Count(a => a.Value.SelectedUnlockPlan == guid));
        return i;
    }

    public override void Draw()
    {
        C.SubmarineUnlockPlans.RemoveAll(x => x.Delete);
        ImGuiEx.InputWithRightButtonsArea("SUPSelector", () =>
        {
            if(ImGui.BeginCombo("##supsel", SelectedPlanName, ImGuiComboFlags.HeightLarge))
            {
                foreach(var x in C.SubmarineUnlockPlans)
                {
                    if(ImGui.Selectable(x.Name + $"##{x.GUID}"))
                    {
                        SelectedPlanGuid = x.GUID;
                    }
                }
                ImGui.EndCombo();
            }
        }, () =>
        {
            if(ImGui.Button(Lang.T("New plan")))
            {
                var x = new SubmarineUnlockPlan();
                x.Name = $"Plan {x.GUID}";
                C.SubmarineUnlockPlans.Add(x);
                SelectedPlanGuid = x.GUID;
            }
        });
        ImGui.Separator();
        if(SelectedPlan == null)
        {
            ImGuiEx.Text(Lang.T("No or unknown plan is selected"));
        }
        else
        {
            if(Data != null)
            {
                var users = GetAmountOfOtherPlanUsers(SelectedPlanGuid);
                var my = Data.AdditionalSubmarineData.Where(x => x.Value.SelectedUnlockPlan == SelectedPlanGuid);
                if(users == 0)
                {
                    if(!my.Any())
                    {
                        ImGuiEx.TextWrapped(Lang.T("This plan is not used by any submersibles."));
                    }
                    else
                    {
                        ImGuiEx.TextWrapped($"此計畫使用於 {my.Select(X => X.Key).Print()}。");
                    }
                }
                else
                {
                    if(!my.Any())
                    {
                        ImGuiEx.TextWrapped($"此計畫使用於其他角色的 {users} 艘潛水艇。");
                    }
                    else
                    {
                        ImGuiEx.TextWrapped($"此計畫使用於 {my.Select(X => X.Key).Print()}，以及其他角色的另外 {users} 艘潛水艇。");
                    }
                }
            }
            if(C.DefaultSubmarineUnlockPlan == SelectedPlanGuid)
            {
                ImGuiEx.Text(Lang.T("This plan is set as default."));
                ImGui.SameLine();
                if(ImGui.SmallButton(Lang.T("Reset"))) C.DefaultSubmarineUnlockPlan = "";
            }
            else
            {
                if(ImGui.SmallButton(Lang.T("Set this plan as default"))) C.DefaultSubmarineUnlockPlan = SelectedPlanGuid;
            }
            ImGuiEx.TextV("名稱：");
            ImGui.SameLine();
            ImGuiEx.SetNextItemFullWidth();
            ImGui.InputText($"##planname", ref SelectedPlan.Name, 100);
            ImGuiEx.LineCentered($"planbuttons", () =>
            {
                ImGuiEx.TextV(Lang.T("Apply this plan to:"));
                ImGui.SameLine();
                if(ImGui.Button(Lang.T("ALL submersibles")))
                {
                    C.OfflineData.Each(x => x.AdditionalSubmarineData.Each(s => s.Value.SelectedUnlockPlan = SelectedPlanGuid));
                }
                ImGui.SameLine();
                if(ImGui.Button(Lang.T("Current character's submersibles")))
                {
                    Data.AdditionalSubmarineData.Each(s => s.Value.SelectedUnlockPlan = SelectedPlanGuid);
                }
                ImGui.SameLine();
                if(ImGui.Button(Lang.T("No submersibles")))
                {
                    C.OfflineData.Each(x => x.AdditionalSubmarineData.Where(s => s.Value.SelectedUnlockPlan == SelectedPlanGuid).Each(s => s.Value.SelectedUnlockPlan = Guid.Empty.ToString()));
                }
            });
            ImGuiEx.LineCentered($"planbuttons2", () =>
            {
                if(ImGui.Button(Lang.T("Copy plan settings")))
                {
                    Copy(JsonConvert.SerializeObject(SelectedPlan));
                }
                ImGui.SameLine();
                if(ImGui.Button(Lang.T("Paste plan settings")))
                {
                    try
                    {
                        var unlockPlan = JsonConvert.DeserializeObject<SubmarineUnlockPlan>(Paste());
                        if(!unlockPlan.IsModified())
                        {
                            Notify.Error("無法匯入剪貼簿內容。請確認是否為正確的計畫資料。");
                        }
                        else
                        {
                            SelectedPlan.CopyFrom(unlockPlan);
                        }
                    }
                    catch(Exception ex)
                    {
                        DuoLog.Error($"Could not import plan: {ex.Message}");
                        ex.Log();
                    }
                }
                ImGui.SameLine();
                if(ImGuiEx.ButtonCtrl(Lang.T("Delete this plan")))
                {
                    SelectedPlan.Delete = true;
                }
                ImGui.SameLine();
                if(ImGui.Button(Lang.T("Help")))
                {
                    Svc.Chat.Print($"這裡列出所有可解鎖航點。插件需要選擇解鎖目標時，會從此清單選第一個可用目的地。注意：不能只指定最終航點，路徑上的所有目的地都需要選取。");
                }
            });
            if(ImGui.BeginChild("Plan"))
            {
                if(!IsSubDataAvail())
                {
                    ImGuiEx.TextWrapped(Lang.T("Access submarine list to retrieve data."));
                }
                ImGui.Checkbox($"解鎖潛水艇欄位。目前欄位：{GetNumUnlockedSubs()?.ToString() ?? "未知"}/4", ref SelectedPlan.UnlockSubs);
                ImGuiEx.TextWrapped(Lang.T("Unlocking slots is always prioritized over unlocking routes."));
                ImGui.Checkbox(Lang.T("Enforce Spam one destination mode in Deep sea site."), ref SelectedPlan.EnforceDSSSinglePoint);
                ImGui.Checkbox(Lang.T("Set this plan as enforced."), ref SelectedPlan.EnforcePlan);
                ImGuiEx.HelpMarker(Lang.T("Any point selected for unlock in this map will be executed by every single eligible submarine until everything is actually unlocked"));
                if(ImGui.BeginTable("##planTable", 3, ImGuiTableFlags.SizingFixedFit | ImGuiTableFlags.Borders | ImGuiTableFlags.RowBg))
                {
                    ImGui.TableSetupColumn(Lang.T("Zone"), ImGuiTableColumnFlags.WidthStretch);
                    ImGui.TableSetupColumn(Lang.T("Map"));
                    ImGui.TableSetupColumn(Lang.T("Unlocked by"));
                    ImGui.TableHeadersRow();
                    foreach(var x in Unlocks.PointToUnlockPoint)
                    {
                        if(x.Value.Point < 9000)
                        {
                            ImGui.PushID($"{x.Key}");
                            ImGui.TableNextRow();
                            ImGui.TableNextColumn();
                            var data = Svc.Data.GetExcelSheet<SubmarineExploration>().GetRowOrDefault(x.Key);
                            if(data != null)
                            {
                                try
                                {
                                    var col = IsMapUnlocked(x.Key);
                                    ImGuiEx.CollectionCheckbox($"{data?.FancyDestination()}", x.Key, SelectedPlan.ExcludedRoutes, true);
                                    if(col) ImGui.PushStyleColor(ImGuiCol.Text, ImGuiColors.ParsedGreen);
                                    if(col) ImGui.PopStyleColor();
                                    ImGui.TableNextColumn();
                                    ImGuiEx.TextV($"{data?.Map.ValueNullable?.Name}");
                                    ImGui.TableNextColumn();
                                    var notEnabled = !SelectedPlan.ExcludedRoutes.Contains(x.Key) && SelectedPlan.ExcludedRoutes.Contains(x.Value.Point);
                                    ImGuiEx.TextV(notEnabled ? ImGuiColors.DalamudRed : null, $"{Svc.Data.GetExcelSheet<SubmarineExploration>().GetRowOrDefault(x.Value.Point)?.FancyDestination()}");
                                }
                                catch(Exception e)
                                {
                                    e.Log();
                                }
                            }
                            ImGui.PopID();
                        }
                    }
                    ImGui.EndTable();
                }
                if(ImGui.CollapsingHeader(Lang.T("Display current point exploration order")))
                {
                    ImGuiEx.Text(SelectedPlan.GetPrioritizedPointList().Select(x => $"{Svc.Data.GetExcelSheet<SubmarineExploration>().GetRowOrDefault(x.point)?.Destination} ({x.justification})").Join("\n"));
                }
            }
            ImGui.EndChild();
        }
    }
}

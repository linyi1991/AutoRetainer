using AutoRetainerAPI.Configuration;
using Dalamud.Game;
using ECommons.Configuration;
using ECommons.ExcelServices;
using ECommons.GameHelpers;
using FFXIVClientStructs.FFXIV.Client.Game;
using Lumina.Excel.Sheets;
using System.Globalization;
using System.Text.RegularExpressions;

namespace AutoRetainer.UI.Windows;

public sealed class VenturePlanner : Window
{
    private OfflineRetainerData SelectedRetainer = null;
    private OfflineCharacterData SelectedCharacter = null;
    private string search = "";
    private int minLevel = 1;
    private int maxLevel = Player.MaxLevel;
    private Dictionary<uint, (string l, string r, bool avail)> Cache = [];
    private string importText = "";
    private string importPlanName = "匯入探索計畫";
    private bool importReplaceCurrentPlan = true;
    private bool importEnablePlannerOnApply = true;
    private bool importRecipeRecursive = true;
    private bool importSubtractActiveVentures = true;
    private List<VentureImportEntry> importPreview = [];
    private List<VentureRetainerAssignment> importAssignments = [];
    private List<string> importAssignmentWarnings = [];
    private string importError = "";

    public VenturePlanner() : base("探索規劃器")
    {
        P.WindowSystem.AddWindow(this);
    }

    internal void Open(OfflineCharacterData characterData, OfflineRetainerData selectedRetainer)
    {
        SelectedCharacter = characterData;
        SelectedRetainer = selectedRetainer;
        IsOpen = true;
    }

    public override void Draw()
    {
        ImGuiEx.SetNextItemFullWidth();
        if(ImGui.BeginCombo("##selectRet", SelectedCharacter != null && SelectedRetainer != null ? $"{Censor.Character(SelectedCharacter.Name, SelectedCharacter.World)} - {Censor.Retainer(SelectedRetainer.Name)} - {SelectedRetainer.Level} {ExcelJobHelper.GetJobNameById(SelectedRetainer.Job)}" : "選擇僱員...", ImGuiComboFlags.HeightLarge))
        {
            foreach(var x in C.OfflineData.OrderBy(x => !C.NoCurrentCharaOnTop && x.CID == Player.CID ? 0 : 1))
            {
                foreach(var r in x.RetainerData)
                {
                    if(ImGui.Selectable($"{Censor.Character(x.Name, x.World)} - {Censor.Retainer(r.Name)} - Lv{r.Level} {ExcelJobHelper.GetJobNameById(r.Job)}"))
                    {
                        SelectedRetainer = r;
                        SelectedCharacter = x;
                    }
                }
            }
            ImGui.EndCombo();
        }

        if(SelectedRetainer != null && SelectedCharacter != null)
        {
            var adata = Utils.GetAdditionalData(SelectedCharacter.CID, SelectedRetainer.Name);
            /*ImGuiEx.TextV("Share venture plan with:");
            ImGui.SameLine();
            ImGuiEx.SetNextItemFullWidth();
            var n = "No shared plan";
            if (adata.LinkedVenturePlan != "")
            {
                var linkedAdata = C.AdditionalData.GetOrDefault(adata.LinkedVenturePlan);
                if (linkedAdata != null)
                {
                    var linkedOCD = Utils.GetOfflineCharacterDataFromAdditionalRetainerDataKey(adata.LinkedVenturePlan);
                    var linkedORD = Utils.GetOfflineRetainerDataFromAdditionalRetainerDataKey(adata.LinkedVenturePlan);
                    n = $"{linkedOCD.Name}@{linkedOCD.World} - {linkedORD.Name}";
                }
            }
            if (ImGui.BeginCombo("##selectLinked", n))
            {
                if(ImGui.Selectable("Remove sharing"))
                {
                    adata.LinkedVenturePlan = "";
                }
                foreach (var x in C.OfflineData.OrderBy(x => !C.NoCurrentCharaOnTop && x.CID == Player.CID ? 0 : 1))
                {
                    foreach (var r in x.RetainerData)
                    {
                        if (ImGui.Selectable($"{Censor.Character(x.Name, x.World)} - {Censor.Retainer(r.Name)} - Lv{r.Level} {ExcelJobHelper.GetJobNameById(r.Job)}"))
                        {
                            adata.LinkedVenturePlan = Utils.GetAdditionalDataKey(x.CID, r.Name);
                        }
                    }
                }
                ImGui.EndCombo();
            }*/
            if(adata.LinkedVenturePlan == "")
            {
                var ww = ImGui.GetContentRegionAvail().X;
                ImGui.Columns(2);
                ImGui.SetColumnWidth(0, ww / 2);

                {
                    int? toRem = null;

                    for(var i = 0; i < adata.VenturePlan.List.Count; i++)
                    {
                        var v = adata.VenturePlan.List[i];
                        ImGui.PushID(v.GUID);
                        {
                            var d = i == 0;
                            if(d) ImGui.BeginDisabled();
                            if(ImGuiEx.IconButton(FontAwesomeIcon.ArrowUp))
                            {
                                Safe(() => (adata.VenturePlan.List[i], adata.VenturePlan.List[i - 1]) = (adata.VenturePlan.List[i - 1], adata.VenturePlan.List[i]));
                            }
                            if(d) ImGui.EndDisabled();
                        }
                        ImGui.SameLine();
                        {
                            var d = i == adata.VenturePlan.List.Count - 1;
                            if(d) ImGui.BeginDisabled();
                            if(ImGuiEx.IconButton(FontAwesomeIcon.ArrowDown))
                            {
                                Safe(() => (adata.VenturePlan.List[i], adata.VenturePlan.List[i + 1]) = (adata.VenturePlan.List[i + 1], adata.VenturePlan.List[i]));
                            }
                            if(d) ImGui.EndDisabled();
                        }
                        ImGui.SameLine();
                        ImGuiEx.SetNextItemWidthScaled(100f);
                        ImGui.InputInt("##cnt", ref v.Num.ValidateRange(1, 9999), 1, 1);
                        ImGui.SameLine();
                        if(ImGuiEx.IconButton(FontAwesomeIcon.Trash))
                        {
                            toRem = i;
                        }
                        ImGui.SameLine();
                        ImGuiEx.Text($"{VentureUtils.GetFancyVentureName(v.ID, SelectedCharacter, SelectedRetainer, out _)}");

                        ImGui.PopID();
                    }

                    if(toRem != null)
                    {
                        adata.VenturePlan.List.RemoveAt(toRem.Value);
                        adata.VenturePlanIndex = 0;
                    }
                }


                ImGui.NextColumn();


                if(ImGui.Checkbox("啟用規劃器", ref adata.EnablePlanner))
                {
                    if(adata.EnablePlanner)
                    {
                        adata.VenturePlanIndex = 0;
                    }
                }

                if(C.SavedPlans.Count > 0)
                {
                    ImGuiEx.SetNextItemFullWidth();
                    if(ImGui.BeginCombo("##load", "載入已儲存計畫...", ImGuiComboFlags.HeightLarge))
                    {
                        int? toRem = null;
                        for(var i = 0; i < C.SavedPlans.Count; i++)
                        {
                            var p = C.SavedPlans[i];
                            ImGui.PushID(p.GUID);
                            if(ImGui.Selectable(p.Name))
                            {
                                adata.VenturePlan = p.JSONClone();
                                adata.VenturePlanIndex = 0;
                            }
                            if(ImGui.IsItemClicked(ImGuiMouseButton.Right))
                            {
                                ImGui.OpenPopup($"Context");
                            }
                            if(ImGui.BeginPopup($"Context"))
                            {
                                if(ImGui.Selectable("刪除計畫"))
                                {
                                    toRem = i;
                                }
                                ImGui.EndPopup();
                            }
                            ImGui.PopID();
                        }
                        if(toRem != null)
                        {
                            C.SavedPlans.RemoveAt(toRem.Value);
                        }
                        ImGui.EndCombo();
                    }
                    //ImGui.Separator();
                }


                if(adata.VenturePlan.List.Count > 0)
                {
                    //ImGui.Separator();
                    ImGuiEx.TextV("計畫完成後：");
                    ImGui.SameLine();
                    ImGuiEx.SetNextItemFullWidth();
                    ImGuiEx.EnumCombo("##cBeh", ref adata.VenturePlan.PlanCompleteBehavior, Lang.PlanCompleteBehaviorNames);
                    //ImGui.Separator();
                    var overwrite = C.SavedPlans.Any(x => x.Name == adata.VenturePlan.Name);
                    ImGuiEx.InputWithRightButtonsArea("SavePlan", delegate
                    {
                        if(overwrite) ImGui.PushStyleColor(ImGuiCol.Text, ImGuiColors.DalamudYellow);
                        ImGui.InputTextWithHint("##name", "輸入計畫名稱...", ref adata.VenturePlan.Name, 50);
                        if(overwrite) ImGui.PopStyleColor();
                    }, delegate
                    {
                        if(ImGuiEx.IconButton(FontAwesomeIcon.Save))
                        {
                            if(overwrite)
                            {
                                C.SavedPlans.RemoveAll(x => x.Name == adata.VenturePlan.Name);
                            }
                            C.SavedPlans.Add(adata.VenturePlan.JSONClone());
                            Notify.Success($"已儲存計畫：{adata.VenturePlan.Name}");
                        }
                        ImGuiEx.Tooltip(overwrite ? "覆蓋既有探索計畫" : "儲存探索計畫");
                    });

                    if(ImGui.Button("清空目前僱員計畫"))
                    {
                        ClearRetainerPlan(adata);
                        Notify.Success("已清空目前僱員的探索計畫。");
                        EzConfig.Save();
                    }
                    ImGui.SameLine();
                    if(ImGui.Button("完整刪除此計畫"))
                    {
                        DeletePlanEverywhere(adata.VenturePlan.Name);
                    }
                }

                DrawTeamcraftImporter(adata);

                ImGuiEx.SetNextItemFullWidth();
                if(ImGui.BeginCombo("##addVenture", "新增探索...", ImGuiComboFlags.HeightLarge))
                {
                    ImGuiEx.SetNextItemFullWidth();
                    ImGui.InputTextWithHint("##search", "篩選...", ref search, 100);
                    ImGuiEx.TextV($"等級範圍：");
                    ImGui.SameLine();
                    ImGuiEx.SetNextItemWidthScaled(50f);
                    ImGui.DragInt("##minL", ref minLevel, 1, 1, Player.MaxLevel);
                    ImGui.SameLine();
                    ImGuiEx.Text($"-");
                    ImGui.SameLine();
                    ImGuiEx.SetNextItemWidthScaled(50f);
                    ImGui.DragInt("##maxL", ref maxLevel, 1, 1, Player.MaxLevel);
                    ImGuiEx.TextV($"不可用探索：");
                    ImGui.SameLine();
                    ImGuiEx.SetNextItemFullWidth();
                    ImGuiEx.EnumCombo("##unavail", ref C.UnavailableVentureDisplay);
                    if(ImGui.BeginChild("##ventureCh", new(ImGui.GetContentRegionAvail().X, ImGuiHelpers.MainViewport.Size.Y / 3)))
                    {
                        if(ImGui.CollapsingHeader(VentureUtils.GetHuntingVentureName(SelectedRetainer.Job)))
                        {
                            foreach(var item in VentureUtils.GetHunts(SelectedRetainer.Job).Where(x => search.IsNullOrEmpty() || x.GetVentureName().Contains(search, StringComparison.OrdinalIgnoreCase)).Where(x => x.RetainerLevel >= minLevel && x.RetainerLevel <= maxLevel))
                            {
                                var l = "";
                                var r = "";
                                bool Avail;
                                if(Cache.TryGetValue(item.RowId, out var result))
                                {
                                    l = result.l;
                                    r = result.r;
                                    Avail = result.avail;
                                }
                                else
                                {
                                    item.GetFancyVentureName(SelectedCharacter, SelectedRetainer, out Avail, out l, out r);
                                    Cache[item.RowId] = (l, r, Avail);
                                }
                                if(Avail || C.UnavailableVentureDisplay != UnavailableVentureDisplay.Hide)
                                {
                                    var d = !Avail && C.UnavailableVentureDisplay != UnavailableVentureDisplay.Allow_selection;
                                    if(d) ImGui.BeginDisabled();
                                    var cur = ImGui.GetCursorPos();
                                    ImGui.SetCursorPosX(ImGui.GetCursorPosX() + ImGui.GetContentRegionAvail().X - ImGui.CalcTextSize(r).X);
                                    ImGuiEx.Text(r);
                                    ImGui.SetCursorPos(cur);
                                    if(ImGui.Selectable(l, adata.VenturePlan.List.Any(x => x.ID == item.RowId), ImGuiSelectableFlags.DontClosePopups))
                                    {
                                        adata.VenturePlan.List.Add(new(item));
                                        adata.VenturePlanIndex = 0;
                                    }
                                    if(d) ImGui.EndDisabled();
                                }
                            }
                        }
                        if(ImGui.CollapsingHeader(VentureUtils.GetFieldExVentureName(SelectedRetainer.Job)))
                        {
                            foreach(var item in VentureUtils.GetFieldExplorations(SelectedRetainer.Job).Where(x => search.IsNullOrEmpty() || x.GetVentureName().Contains(search, StringComparison.OrdinalIgnoreCase)).Where(x => x.RetainerLevel >= minLevel && x.RetainerLevel <= maxLevel))
                            {
                                var name = item.GetFancyVentureName(SelectedCharacter, SelectedRetainer, out var Avail);
                                var d = !Avail && C.UnavailableVentureDisplay != UnavailableVentureDisplay.Allow_selection;
                                if(d) ImGui.BeginDisabled();
                                if(ImGui.Selectable(name, adata.VenturePlan.List.Any(x => x.ID == item.RowId), ImGuiSelectableFlags.DontClosePopups))
                                {
                                    adata.VenturePlan.List.Add(new(item));
                                    adata.VenturePlanIndex = 0;
                                }
                                if(d) ImGui.EndDisabled();
                            }
                        }
                        ImGui.PushStyleVar(ImGuiStyleVar.ButtonTextAlign, Vector2.Zero);
                        if(ImGui.Button($"{Lang.CharDice}    快速探索", ImGuiHelpers.GetButtonSize("A") with { X = ImGui.GetContentRegionAvail().X }))
                        {
                            adata.VenturePlan.List.Add(new(VentureUtils.QuickExplorationID));
                            adata.VenturePlanIndex = 0;
                        }
                        ImGui.PopStyleVar();
                        ImGui.EndChild();
                    }
                    ImGui.EndCombo();
                }
                else
                {
                    Cache.Clear();
                }

                if(adata.EnablePlanner && adata.VenturePlan.ListUnwrapped.Count > 0)
                {
                    var pct = adata.VenturePlanIndex / (float)adata.VenturePlan.ListUnwrapped.Count;

                    if(ImGuiEx.IconButton(Lang.IconRefresh))
                    {
                        adata.VenturePlanIndex = 0;
                    }
                    ImGui.SameLine();
                    ImGuiEx.Tooltip("取消此計畫剩餘進度，從頭開始");
                    ImGui.ProgressBar(pct, new Vector2(ImGui.GetContentRegionAvail().X, ImGuiHelpers.GetButtonSize("X").Y));
                }

                if(C.Verbose)
                {
                    if(ImGui.CollapsingHeader("Debug"))
                    {
                        ImGuiEx.InputUint("Index", ref adata.VenturePlanIndex);
                    }
                }

                ImGui.Columns(1);
            }
            else
            {
                ImGuiEx.TextWrapped($"此僱員的探索計畫正在與其他僱員共用。");
            }
        }

    }

    private void DrawTeamcraftImporter(AdditionalRetainerData adata)
    {
        ImGui.Separator();
        if(!ImGui.CollapsingHeader("從成品或材料清單匯入探索計畫")) return;

        ImGuiEx.TextWrapped("可輸入成品或材料清單。成品模式會先依配方拆成材料，再篩出僱員能執行的 1 小時指定探索；材料模式則直接把每行當成材料處理。支援「物品名稱 x數量」、「數量 物品名稱」、「物品名稱: 數量」等格式。");
        ImGuiEx.SetNextItemFullWidth();
        if(ImGui.InputTextMultiline("##ventureImportText", ref importText, 12000, new Vector2(ImGui.GetContentRegionAvail().X, 150)))
        {
            importPreview.Clear();
            importAssignments.Clear();
            importAssignmentWarnings.Clear();
            importError = "";
        }

        ImGuiEx.SetNextItemFullWidth();
        ImGui.InputTextWithHint("##ventureImportPlanName", "計畫名稱", ref importPlanName, 50);
        ImGui.Checkbox("成品分析時遞迴拆半成品", ref importRecipeRecursive);
        ImGui.Checkbox("分析時扣除目前角色背包與水晶數量", ref C.VentureImportSubtractInventory);
        ImGui.Checkbox("分析時扣除已派出探索的預估數量", ref importSubtractActiveVentures);
        ImGui.Checkbox("套用時取代目前計畫", ref importReplaceCurrentPlan);
        ImGui.SameLine();
        ImGui.Checkbox("套用後啟用規劃器", ref importEnablePlannerOnApply);

        ImGui.Separator();
        DrawTrackedGoalList();

        if(ImGui.Button("分析成品材料"))
        {
            importPreview = BuildImportPreview(analyzeRecipes: true);
            importAssignments.Clear();
        }
        ImGui.SameLine();
        if(ImGui.Button("分析材料清單"))
        {
            importPreview = BuildImportPreview(analyzeRecipes: false);
            importAssignments.Clear();
            importAssignmentWarnings.Clear();
        }
        ImGui.SameLine();
        if(ImGui.Button("分析成品並分配"))
        {
            BuildAssignmentPreview(analyzeRecipes: true);
        }
        ImGui.SameLine();
        if(ImGui.Button("分析材料並分配"))
        {
            BuildAssignmentPreview(analyzeRecipes: false);
        }
        ImGui.SameLine();
        if(ImGui.Button("清空"))
        {
            importText = "";
            importPreview.Clear();
            importAssignments.Clear();
            importAssignmentWarnings.Clear();
            importError = "";
        }

        if(!importError.IsNullOrEmpty())
        {
            ImGuiEx.TextWrapped(ImGuiColors.DalamudRed, importError);
        }

        DrawAssignmentPreview();

        if(importPreview.Count == 0) return;

        var matched = importPreview.Where(x => x.Task != null).ToList();
        var unmatched = importPreview.Count - matched.Count;
        ImGuiEx.TextWrapped($"已解析 {importPreview.Count} 筆，找到 {matched.Count} 筆可由目前僱員探索取得；未對應 {unmatched} 筆。");
        if(matched.Count == 0)
        {
            ImGuiEx.TextWrapped(ImGuiColors.DalamudYellow, "目前沒有可套用到這位僱員的探索項目。若要讓 9 位僱員分攤下一輪，請按「分析成品並分配」或「分析材料並分配」，產生全僱員分配預覽後再套用。");
        }

        if(ImGui.BeginTable("##ventureImportPreview", 6, ImGuiTableFlags.Borders | ImGuiTableFlags.RowBg | ImGuiTableFlags.SizingStretchProp))
        {
            ImGui.TableSetupColumn("材料");
            ImGui.TableSetupColumn("需求");
            ImGui.TableSetupColumn("探索");
            ImGui.TableSetupColumn("每次");
            ImGui.TableSetupColumn("次數");
            ImGui.TableSetupColumn("狀態");
            ImGui.TableHeadersRow();

            foreach(var entry in importPreview)
            {
                ImGui.TableNextRow();
                ImGui.TableNextColumn();
                ImGuiEx.TextWrapped(entry.ItemName);
                ImGui.TableNextColumn();
                ImGuiEx.Text($"{entry.Amount}");
                ImGui.TableNextColumn();
                ImGuiEx.TextWrapped(entry.Task?.GetVentureName() ?? "-");
                ImGui.TableNextColumn();
                ImGuiEx.Text(entry.Yield > 0 ? $"{entry.Yield}" : "-");
                ImGui.TableNextColumn();
                ImGuiEx.Text(entry.Runs > 0 ? $"{entry.Runs}" : "-");
                ImGui.TableNextColumn();
                var color = entry.Task == null || !entry.Available ? ImGuiColors.DalamudYellow : ImGuiColors.ParsedGreen;
                ImGuiEx.TextWrapped(color, entry.Status);
            }

            ImGui.EndTable();
        }

        var disableApply = matched.Count == 0;
        if(disableApply) ImGui.BeginDisabled();
        if(ImGui.Button("套用到目前僱員"))
        {
            ApplyImportToPlan(adata, saveOnly: false);
        }
        ImGui.SameLine();
        if(ImGui.Button("儲存成共用計畫"))
        {
            ApplyImportToPlan(adata, saveOnly: true);
        }
        if(disableApply) ImGui.EndDisabled();
    }

    private void DrawTrackedGoalList()
    {
        EnsureTrackedGoalListMigrated();
        ImGuiEx.TextWrapped("追蹤目標會保存多份需求清單。之後可選一份，用最新角色背包、水晶與派出中探索重新計算，再覆蓋下一輪僱員計畫。");

        var selected = GetSelectedTrackedGoal();
        ImGuiEx.SetNextItemFullWidth();
        if(ImGui.BeginCombo("##ventureTrackedGoalList", selected?.Name ?? "選擇追蹤目標...", ImGuiComboFlags.HeightLarge))
        {
            foreach(var goal in C.VentureTrackedGoals.OrderByDescending(x => x.UpdatedAt))
            {
                var label = $"{goal.Name}（{GetImportLineCount(goal.Text)} 行）";
                if(ImGui.Selectable(label, goal.Guid == C.VentureSelectedTrackedGoalGuid))
                {
                    C.VentureSelectedTrackedGoalGuid = goal.Guid;
                    EzConfig.Save();
                }
            }
            ImGui.EndCombo();
        }

        if(ImGui.Button("新增追蹤目標"))
        {
            SaveCurrentInputAsTrackedGoal(forceNew: true);
        }
        ImGui.SameLine();
        var disableUpdate = selected == null;
        if(disableUpdate) ImGui.BeginDisabled();
        if(ImGui.Button("更新選取目標"))
        {
            SaveCurrentInputAsTrackedGoal(forceNew: false);
        }
        ImGui.SameLine();
        if(ImGui.Button("載入選取"))
        {
            LoadTrackedGoal(selected);
        }
        ImGui.SameLine();
        if(ImGui.Button("重算選取並套用"))
        {
            RebuildTrackedGoalAndApply(selected);
        }
        ImGui.SameLine();
        if(ImGui.Button("刪除選取"))
        {
            DeleteTrackedGoal(selected);
        }
        ImGui.SameLine();
        if(ImGui.Button("完整刪除選取"))
        {
            DeleteTrackedGoalEverywhere(selected);
        }
        if(disableUpdate) ImGui.EndDisabled();

        if(selected != null)
        {
            ImGuiEx.TextWrapped(ImGuiColors.ParsedGreen, $"目前選取：{selected.Name}，{GetImportLineCount(selected.Text)} 行。");
        }
    }

    private void EnsureTrackedGoalListMigrated()
    {
        C.VentureTrackedGoals ??= [];
        if(!C.VentureTrackedImportText.IsNullOrEmpty() && C.VentureTrackedGoals.Count == 0)
        {
            var goal = new VentureTrackedGoal()
            {
                Name = "舊追蹤目標",
                Text = C.VentureTrackedImportText,
                AnalyzeRecipes = C.VentureTrackedAnalyzeRecipes,
                UpdatedAt = DateTime.Now
            };
            C.VentureTrackedGoals.Add(goal);
            C.VentureSelectedTrackedGoalGuid = goal.Guid;
            EzConfig.Save();
        }

        if(!C.VentureSelectedTrackedGoalGuid.IsNullOrEmpty() && C.VentureTrackedGoals.All(x => x.Guid != C.VentureSelectedTrackedGoalGuid))
        {
            C.VentureSelectedTrackedGoalGuid = "";
            EzConfig.Save();
        }
    }

    private VentureTrackedGoal GetSelectedTrackedGoal()
    {
        EnsureTrackedGoalListMigrated();
        return C.VentureTrackedGoals.FirstOrDefault(x => x.Guid == C.VentureSelectedTrackedGoalGuid);
    }

    private void SaveCurrentInputAsTrackedGoal(bool forceNew)
    {
        if(importText.IsNullOrEmpty())
        {
            Notify.Warning("目前輸入是空的，無法儲存追蹤目標。");
            return;
        }

        var goal = forceNew ? null : GetSelectedTrackedGoal();
        if(goal == null)
        {
            goal = new VentureTrackedGoal();
            C.VentureTrackedGoals.Add(goal);
        }

        goal.Name = GetTrackedGoalName();
        goal.Text = importText;
        goal.AnalyzeRecipes = true;
        goal.UpdatedAt = DateTime.Now;
        C.VentureSelectedTrackedGoalGuid = goal.Guid;
        C.VentureTrackedImportText = goal.Text;
        C.VentureTrackedAnalyzeRecipes = goal.AnalyzeRecipes;
        EzConfig.Save();
        Notify.Success(forceNew ? $"已新增追蹤目標：{goal.Name}" : $"已更新追蹤目標：{goal.Name}");
    }

    private string GetTrackedGoalName()
    {
        if(!importPlanName.IsNullOrEmpty()) return importPlanName;
        var firstLine = importText.Replace("\r", "\n").Split('\n', StringSplitOptions.RemoveEmptyEntries).FirstOrDefault()?.Trim();
        if(!firstLine.IsNullOrEmpty()) return firstLine.Length > 40 ? firstLine[..40] : firstLine;
        return "追蹤目標";
    }

    private void LoadTrackedGoal(VentureTrackedGoal goal)
    {
        if(goal == null) return;
        importText = goal.Text;
        importPlanName = goal.Name;
        importPreview.Clear();
        importAssignments.Clear();
        importAssignmentWarnings.Clear();
        importError = "";
        C.VentureSelectedTrackedGoalGuid = goal.Guid;
        EzConfig.Save();
    }

    private void DeleteTrackedGoal(VentureTrackedGoal goal)
    {
        if(goal == null) return;
        C.VentureTrackedGoals.Remove(goal);
        C.VentureSelectedTrackedGoalGuid = C.VentureTrackedGoals.OrderByDescending(x => x.UpdatedAt).FirstOrDefault()?.Guid ?? "";
        EzConfig.Save();
        Notify.Success($"已刪除追蹤目標：{goal.Name}");
    }

    private void DeleteTrackedGoalEverywhere(VentureTrackedGoal goal)
    {
        if(goal == null) return;
        var name = goal.Name;
        C.VentureTrackedGoals.Remove(goal);
        C.VentureSelectedTrackedGoalGuid = C.VentureTrackedGoals.OrderByDescending(x => x.UpdatedAt).FirstOrDefault()?.Guid ?? "";
        var cleared = DeletePlanEverywhere(name, notify: false);
        Notify.Success($"已完整刪除：{name}，並清除 {cleared} 位僱員的同名探索計畫。");
        EzConfig.Save();
    }

    private static void ClearRetainerPlan(AdditionalRetainerData adata)
    {
        adata.EnablePlanner = false;
        adata.LinkedVenturePlan = "";
        adata.VenturePlan = new()
        {
            PlanCompleteBehavior = PlanCompleteBehavior.Do_nothing
        };
        adata.VenturePlanIndex = 0;
    }

    private static int DeletePlanEverywhere(string planName, bool notify = true)
    {
        if(planName.IsNullOrEmpty())
        {
            Notify.Warning("目前計畫沒有名稱，無法完整刪除同名計畫。");
            return 0;
        }

        var removedSavedPlans = C.SavedPlans.RemoveAll(x => x.Name == planName);
        var clearedRetainers = 0;
        foreach(var adata in C.AdditionalData.Values)
        {
            if(adata.VenturePlan.Name != planName) continue;
            ClearRetainerPlan(adata);
            clearedRetainers++;
        }

        if(C.VentureTrackedGoals != null)
        {
            C.VentureTrackedGoals.RemoveAll(x => x.Name == planName);
            if(!C.VentureSelectedTrackedGoalGuid.IsNullOrEmpty() && C.VentureTrackedGoals.All(x => x.Guid != C.VentureSelectedTrackedGoalGuid))
            {
                C.VentureSelectedTrackedGoalGuid = C.VentureTrackedGoals.OrderByDescending(x => x.UpdatedAt).FirstOrDefault()?.Guid ?? "";
            }
        }

        EzConfig.Save();
        if(notify)
        {
            Notify.Success($"已完整刪除：{planName}，移除 {removedSavedPlans} 份共用計畫，清除 {clearedRetainers} 位僱員。");
        }
        return clearedRetainers;
    }

    private static int GetImportLineCount(string text)
    {
        return text.Replace("\r", "\n").Split('\n', StringSplitOptions.RemoveEmptyEntries).Length;
    }

    private void DrawAssignmentPreview()
    {
        foreach(var warning in importAssignmentWarnings)
        {
            ImGuiEx.TextWrapped(ImGuiColors.DalamudYellow, warning);
        }

        if(importAssignments.Count == 0) return;

        var totalRuns = importAssignments.Sum(x => x.TotalRuns);
        ImGui.Separator();
        ImGuiEx.TextWrapped($"全僱員分配預覽：{importAssignments.Count} 位僱員，合計 {totalRuns} 趟。");

        if(ImGui.BeginTable("##ventureImportAssignments", 5, ImGuiTableFlags.Borders | ImGuiTableFlags.RowBg | ImGuiTableFlags.SizingStretchProp))
        {
            ImGui.TableSetupColumn("角色");
            ImGui.TableSetupColumn("僱員");
            ImGui.TableSetupColumn("職業");
            ImGui.TableSetupColumn("分配內容");
            ImGui.TableSetupColumn("趟數");
            ImGui.TableHeadersRow();

            foreach(var assignment in importAssignments.OrderBy(x => x.Character.NameWithWorld).ThenBy(x => x.Retainer.Name))
            {
                ImGui.TableNextRow();
                ImGui.TableNextColumn();
                ImGuiEx.TextWrapped(assignment.Character.NameWithWorldCensored);
                ImGui.TableNextColumn();
                ImGuiEx.TextWrapped(Censor.Retainer(assignment.Retainer.Name));
                ImGui.TableNextColumn();
                ImGuiEx.TextWrapped($"{Lang.CharLevel}{assignment.Retainer.Level} {ExcelJobHelper.GetJobNameById(assignment.Retainer.Job)}");
                ImGui.TableNextColumn();
                ImGuiEx.TextWrapped(string.Join("\n", assignment.Entries.Select(x => $"{x.ItemName}：{x.Runs} 趟，約 {x.Provided} 個")));
                ImGui.TableNextColumn();
                ImGuiEx.Text($"{assignment.TotalRuns}");
            }

            ImGui.EndTable();
        }

        if(ImGui.Button("套用到所有可用僱員"))
        {
            ApplyAssignmentsToRetainers();
        }
    }

    private List<VentureImportEntry> BuildImportPreview(bool analyzeRecipes)
    {
        importError = "";
        importAssignmentWarnings.Clear();
        List<VentureImportEntry> result = [];
        if(SelectedCharacter == null || SelectedRetainer == null)
        {
            importError = "請先選擇僱員。";
            return result;
        }

        var requestedItems = ApplyDeficitCredits(GetRequestedMaterials(analyzeRecipes));
        if(requestedItems.Count == 0)
        {
            importError = analyzeRecipes ? "沒有解析到成品材料。請確認每行有成品名稱與數量。" : "沒有解析到材料。請確認每行有材料名稱與數量。";
            return result;
        }

        var tasks = VentureUtils.GetHunts(SelectedRetainer.Job)
            .Where(x => x.GetVentureItemId() != 0)
            .GroupBy(x => x.GetVentureItemId())
            .ToDictionary(x => x.Key, x => x.OrderByDescending(t => t.RetainerLevel).First());

        var namedTasks = VentureUtils.GetHunts(SelectedRetainer.Job)
            .Where(x => x.GetVentureItemId() != 0)
            .GroupBy(x => NormalizeItemName(x.GetVentureName()))
            .ToDictionary(x => x.Key, x => x.OrderByDescending(t => t.RetainerLevel).First(), StringComparer.OrdinalIgnoreCase);

        foreach(var item in requestedItems)
        {
            var itemId = item.ItemId == 0 ? FindItemIdByName(item.Name) : item.ItemId;
            var normalized = NormalizeItemName(item.Name);
            if((itemId != 0 && tasks.TryGetValue(itemId, out var task)) || namedTasks.TryGetValue(normalized, out task))
            {
                var parts = task.GetFancyVentureNameParts(SelectedCharacter, SelectedRetainer, out var available);
                var yield = Math.Max(parts.Yield, 0);
                var runs = yield > 0 ? (int)Math.Ceiling(item.Amount / (double)yield) : 1;
                result.Add(new()
                {
                    ItemName = item.Name,
                    Amount = item.Amount,
                    Task = task,
                    Yield = yield,
                    Runs = Math.Max(runs, 1),
                    Available = available,
                    Status = available ? "可用" : "已對應，但目前僱員等級/能力/解鎖狀態可能不足"
                });
            }
            else
            {
                result.Add(new()
                {
                    ItemName = item.Name,
                    Amount = item.Amount,
                    Status = GetCurrentRetainerFailureReason(item)
                });
            }
        }

        return result;
    }

    private void BuildAssignmentPreview(bool analyzeRecipes)
    {
        importPreview.Clear();
        importAssignments.Clear();
        importAssignmentWarnings.Clear();
        importError = "";

        var requestedItems = ApplyDeficitCredits(GetRequestedMaterials(analyzeRecipes));
        if(requestedItems.Count == 0)
        {
            importError = analyzeRecipes ? "沒有解析到成品材料。請確認每行有成品名稱與數量。" : "沒有解析到材料。請確認每行有材料名稱與數量。";
            return;
        }

        Dictionary<string, VentureRetainerAssignment> assignments = [];

        foreach(var item in requestedItems)
        {
            var candidates = FindAssignmentCandidates(item);
            if(candidates.Count == 0)
            {
                importAssignmentWarnings.Add($"未分配：{item.Name} x{item.Amount}，{GetAssignmentFailureReason(item)}");
                continue;
            }

            var remaining = item.Amount;
            var guard = 0;
            while(remaining > 0 && guard++ < 20000)
            {
                var candidate = candidates
                    .OrderBy(x => GetOrCreateAssignment(assignments, x).TotalRuns)
                    .ThenByDescending(x => x.Yield)
                    .First();
                var assignment = GetOrCreateAssignment(assignments, candidate);
                assignment.Add(candidate.Task, item.Name, 1, candidate.Yield);
                remaining -= candidate.Yield;
            }

            if(remaining > 0)
            {
                importAssignmentWarnings.Add($"部分未分配：{item.Name} 還剩約 {remaining} 個。");
            }
        }

        importAssignments = assignments.Values.Where(x => x.Entries.Count > 0).ToList();
        if(importAssignments.Count == 0 && importAssignmentWarnings.Count == 0)
        {
            importError = "沒有產生任何分配結果。";
        }
    }

    private List<VentureAssignmentCandidate> FindAssignmentCandidates(VentureRequestedItem item)
    {
        var itemId = item.ItemId == 0 ? FindItemIdByName(item.Name) : item.ItemId;
        var normalized = NormalizeItemName(item.Name);
        List<VentureAssignmentCandidate> ret = [];

        foreach(var character in C.OfflineData)
        {
            foreach(var retainer in character.RetainerData)
            {
                var task = VentureUtils.GetHunts(retainer.Job)
                    .Cast<RetainerTask?>()
                    .FirstOrDefault(x =>
                    {
                        if(x == null) return false;
                        if(itemId != 0 && x.Value.GetVentureItemId() == itemId) return true;
                        return NormalizeItemName(x.Value.GetVentureName() ?? "") == normalized;
                    });
                if(task == null) continue;

                var parts = task.Value.GetFancyVentureNameParts(character, retainer, out var available);
                if(!available || parts.Yield <= 0) continue;

                ret.Add(new()
                {
                    Character = character,
                    Retainer = retainer,
                    Task = task.Value,
                    Yield = parts.Yield
                });
            }
        }

        return ret;
    }

    private string GetCurrentRetainerFailureReason(VentureRequestedItem item)
    {
        if(SelectedCharacter == null || SelectedRetainer == null) return "請先選擇僱員";

        var itemId = item.ItemId == 0 ? FindItemIdByName(item.Name) : item.ItemId;
        var normalized = NormalizeItemName(item.Name);
        var task = VentureUtils.GetHunts(SelectedRetainer.Job)
            .Cast<RetainerTask?>()
            .FirstOrDefault(x => IsMatchingVentureTask(x, itemId, normalized));

        if(task == null)
        {
            return "目前僱員職業沒有對應的 1 小時指定探索；可改按「分析成品並分配」讓其他職業僱員承接";
        }

        var parts = task.Value.GetFancyVentureNameParts(SelectedCharacter, SelectedRetainer, out var available);
        if(parts.Yield <= 0) return "已找到探索，但無法估算每次取得量";
        return available ? "可用" : "已找到探索，但目前僱員等級、能力或採集解鎖狀態不足";
    }

    private string GetAssignmentFailureReason(VentureRequestedItem item)
    {
        var itemId = item.ItemId == 0 ? FindItemIdByName(item.Name) : item.ItemId;
        var normalized = NormalizeItemName(item.Name);
        var anyTask = false;
        var unavailable = 0;
        var noYield = 0;
        HashSet<string> jobs = [];

        foreach(var character in C.OfflineData)
        {
            foreach(var retainer in character.RetainerData)
            {
                var task = VentureUtils.GetHunts(retainer.Job)
                    .Cast<RetainerTask?>()
                    .FirstOrDefault(x => IsMatchingVentureTask(x, itemId, normalized));
                if(task == null) continue;

                anyTask = true;
                jobs.Add(ExcelJobHelper.GetJobNameById(retainer.Job));
                var parts = task.Value.GetFancyVentureNameParts(character, retainer, out var available);
                if(parts.Yield <= 0)
                {
                    noYield++;
                    continue;
                }
                if(!available) unavailable++;
            }
        }

        if(anyTask)
        {
            var jobText = jobs.Count == 0 ? "" : $"（職業：{string.Join("、", jobs.Order())}）";
            if(unavailable > 0) return $"有找到對應探索{jobText}，但僱員等級、能力或採集解鎖狀態不足。";
            if(noYield > 0) return $"有找到對應探索{jobText}，但無法估算每次取得量。";
            return $"有找到對應探索{jobText}，但沒有符合目前條件的可用僱員。";
        }

        if(ExistsAnyVentureTaskForItem(itemId, normalized))
        {
            return "遊戲資料有對應探索，但目前記錄的僱員職業不符合。";
        }

        return "遊戲資料中沒有對應的 1 小時指定探索，可能需要手動採集、購買，或它不是僱員可指定取得的材料。";
    }

    private static bool ExistsAnyVentureTaskForItem(uint itemId, string normalizedName)
    {
        try
        {
            return Svc.Data.GetExcelSheet<RetainerTask>()
                .Where(x => x.MaxTimemin == 60)
                .Cast<RetainerTask?>()
                .Any(x => IsMatchingVentureTask(x, itemId, normalizedName));
        }
        catch
        {
            return false;
        }
    }

    private static bool IsMatchingVentureTask(RetainerTask? task, uint itemId, string normalizedName)
    {
        if(task == null) return false;
        if(task.Value.GetVentureItemId() == 0) return false;
        if(itemId != 0 && task.Value.GetVentureItemId() == itemId) return true;
        return NormalizeItemName(task.Value.GetVentureName() ?? "") == normalizedName;
    }

    private List<VentureRequestedItem> GetRequestedMaterials(bool analyzeRecipes)
    {
        var requestedItems = ParseImportLines(importText);
        if(!analyzeRecipes || requestedItems.Count == 0) return requestedItems;

        Dictionary<uint, VentureRequestedItem> materialsById = [];
        Dictionary<string, VentureRequestedItem> materialsByName = new(StringComparer.OrdinalIgnoreCase);
        foreach(var item in requestedItems)
        {
            try
            {
                ExpandRecipeMaterials(item.Name, item.Amount, materialsById, materialsByName, new HashSet<uint>(), 0);
            }
            catch(Exception e)
            {
                importAssignmentWarnings.Add($"成品分析失敗：{item.Name} x{item.Amount}，已略過。{e.GetType().Name}");
            }
        }
        return materialsById.Values.Concat(materialsByName.Values).ToList();
    }

    private void ExpandRecipeMaterials(string itemName, int amountNeeded, Dictionary<uint, VentureRequestedItem> materialsById, Dictionary<string, VentureRequestedItem> materialsByName, HashSet<uint> seenRecipes, int depth)
    {
        var recipe = FindRecipeByResultName(itemName);
        if(recipe == null)
        {
            importAssignmentWarnings.Add($"未找到成品配方：{itemName} x{amountNeeded}，已改以材料處理。");
            AddRequestedMaterial(materialsById, materialsByName, itemName, amountNeeded, FindItemIdByName(itemName));
            return;
        }

        var recipeValue = recipe.Value;
        if(!seenRecipes.Add(recipeValue.RowId))
        {
            importAssignmentWarnings.Add($"配方循環略過：{GetItemName(recipeValue.ItemResult)}。");
            return;
        }

        var resultAmount = Math.Max(recipeValue.AmountResult, (byte)1);
        var craftCount = (int)Math.Ceiling(amountNeeded / (double)resultAmount);
        for(var i = 0; i < Math.Min(recipeValue.Ingredient.Count, recipeValue.AmountIngredient.Count); i++)
        {
            Lumina.Excel.RowRef<Item> ingredient;
            byte ingredientAmount;
            try
            {
                ingredient = recipeValue.Ingredient[i];
                ingredientAmount = recipeValue.AmountIngredient[i];
            }
            catch
            {
                continue;
            }

            if(ingredient.RowId == 0) continue;
            if(ingredientAmount == 0) continue;

            var ingredientName = GetItemName(ingredient);
            if(ingredientName.IsNullOrEmpty()) continue;

            var total = ingredientAmount * craftCount;
            if(importRecipeRecursive && depth < 8 && FindRecipeByResultName(ingredientName) != null)
            {
                ExpandRecipeMaterials(ingredientName, total, materialsById, materialsByName, new HashSet<uint>(seenRecipes), depth + 1);
            }
            else
            {
                AddRequestedMaterial(materialsById, materialsByName, ingredientName, total, ingredient.RowId);
            }
        }
    }

    private static void AddRequestedMaterial(Dictionary<uint, VentureRequestedItem> materialsById, Dictionary<string, VentureRequestedItem> materialsByName, string itemName, int amount, uint itemId)
    {
        if(itemId != 0)
        {
            if(materialsById.TryGetValue(itemId, out var existing)) existing.Amount += amount;
            else materialsById[itemId] = new(itemName, amount, itemId);
            return;
        }

        if(materialsByName.TryGetValue(itemName, out var named)) named.Amount += amount;
        else materialsByName[itemName] = new(itemName, amount, 0);
    }

    private static Recipe? FindRecipeByResultName(string itemName)
    {
        var normalized = NormalizeItemName(itemName);
        foreach(var language in GetRecipeSearchLanguages())
        {
            Lumina.Excel.ExcelSheet<Recipe> sheet;
            try
            {
                sheet = Svc.Data.GetExcelSheet<Recipe>(language);
            }
            catch
            {
                continue;
            }
            if(sheet == null) continue;
            var match = sheet
                .Where(x =>
                {
                    try
                    {
                        return x.ItemResult.RowId != 0 && x.AmountResult > 0 && NormalizeItemName(GetItemName(x.ItemResult)) == normalized;
                    }
                    catch
                    {
                        return false;
                    }
                })
                .OrderByDescending(x =>
                {
                    try
                    {
                        return x.RecipeLevelTable.RowId;
                    }
                    catch
                    {
                        return 0u;
                    }
                })
                .ThenBy(x => x.RowId)
                .Cast<Recipe?>()
                .FirstOrDefault();
            if(match != null) return match;
        }
        return null;
    }

    private static IEnumerable<ClientLanguage> GetRecipeSearchLanguages()
    {
        List<ClientLanguage> languages =
        [
            Svc.ClientState.ClientLanguage,
            Svc.Data.Language,
            ClientLanguage.English,
            ClientLanguage.Japanese,
            ClientLanguage.German,
            ClientLanguage.French,
            (ClientLanguage)4,
            (ClientLanguage)5,
        ];
        return languages.Distinct();
    }

    private List<VentureRequestedItem> ApplyDeficitCredits(List<VentureRequestedItem> requestedItems)
    {
        return ApplyActiveVentureCredits(ApplyInventoryCredits(requestedItems));
    }

    private List<VentureRequestedItem> ApplyInventoryCredits(List<VentureRequestedItem> requestedItems)
    {
        if(!C.VentureImportSubtractInventory || requestedItems.Count == 0) return requestedItems;

        List<VentureRequestedItem> result = [];
        foreach(var item in requestedItems)
        {
            var itemId = item.ItemId == 0 ? FindItemIdByName(item.Name) : item.ItemId;
            var amount = item.Amount;
            if(itemId != 0)
            {
                var owned = GetCurrentPlayerItemCount(itemId);
                if(owned > 0)
                {
                    var deducted = Math.Min(amount, owned);
                    amount -= deducted;
                    importAssignmentWarnings.Add($"已扣除目前持有：{item.Name} -{deducted}（角色背包/水晶）。");
                }
            }
            if(amount > 0) result.Add(new(item.Name, amount, itemId));
        }

        return result;
    }

    private static int GetCurrentPlayerItemCount(uint itemId)
    {
        try
        {
            return Utils.GetItemCount(Utils.PlayerInvetoriesWithCrystals, itemId);
        }
        catch
        {
            return 0;
        }
    }

    private List<VentureRequestedItem> ApplyActiveVentureCredits(List<VentureRequestedItem> requestedItems)
    {
        if(!importSubtractActiveVentures || requestedItems.Count == 0) return requestedItems;

        var credits = GetActiveVentureCredits();
        if(credits.Count == 0) return requestedItems;

        List<VentureRequestedItem> result = [];
        foreach(var item in requestedItems)
        {
            var itemId = item.ItemId == 0 ? FindItemIdByName(item.Name) : item.ItemId;
            var amount = item.Amount;
            if(itemId != 0 && credits.TryGetValue(itemId, out var credit) && credit.Amount > 0)
            {
                var deducted = Math.Min(amount, credit.Amount);
                amount -= deducted;
                credit.Amount -= deducted;
                importAssignmentWarnings.Add($"已扣除派出中預估：{item.Name} -{deducted}（{credit.Label}）。");
            }
            if(amount > 0) result.Add(new(item.Name, amount, itemId));
        }

        return result;
    }

    private static Dictionary<uint, ActiveVentureCredit> GetActiveVentureCredits()
    {
        Dictionary<uint, ActiveVentureCredit> ret = [];
        foreach(var character in C.OfflineData)
        {
            foreach(var retainer in character.RetainerData)
            {
                if(!retainer.HasVenture || retainer.VentureID == 0) continue;
                if(retainer.GetVentureSecondsRemaining() <= 0) continue;

                var task = VentureUtils.GetVentureById(retainer.VentureID);
                var itemId = task.GetVentureItemId();
                if(itemId == 0) continue;

                var parts = task.GetFancyVentureNameParts(character, retainer, out _);
                if(parts.Yield <= 0) continue;

                if(!ret.TryGetValue(itemId, out var credit))
                {
                    credit = new()
                    {
                        Label = task.GetVentureName() ?? $"#{itemId}"
                    };
                    ret[itemId] = credit;
                }
                credit.Amount += parts.Yield;
            }
        }
        return ret;
    }

    private static uint FindItemIdByName(string itemName)
    {
        var normalized = NormalizeItemName(itemName);
        if(normalized.IsNullOrEmpty()) return 0;

        foreach(var language in GetRecipeSearchLanguages())
        {
            Lumina.Excel.ExcelSheet<Item> sheet;
            try
            {
                sheet = Svc.Data.GetExcelSheet<Item>(language);
            }
            catch
            {
                continue;
            }
            if(sheet == null) continue;

            foreach(var item in sheet)
            {
                try
                {
                    if(item.RowId != 0 && NormalizeItemName(item.Name.ToString()) == normalized) return item.RowId;
                }
                catch
                {
                    //
                }
            }
        }

        return 0;
    }

    private static string GetItemName(Lumina.Excel.RowRef<Item> item)
    {
        try
        {
            return item.RowId == 0 ? "" : item.Value.Name.ToString();
        }
        catch
        {
            return "";
        }
    }

    private static VentureRetainerAssignment GetOrCreateAssignment(Dictionary<string, VentureRetainerAssignment> assignments, VentureAssignmentCandidate candidate)
    {
        var key = $"{candidate.Character.Identity}/{candidate.Retainer.Identity}";
        if(assignments.TryGetValue(key, out var assignment)) return assignment;
        assignment = new()
        {
            Character = candidate.Character,
            Retainer = candidate.Retainer
        };
        assignments[key] = assignment;
        return assignment;
    }

    private void RebuildTrackedGoalAndApply(VentureTrackedGoal goal)
    {
        if(goal == null || goal.Text.IsNullOrEmpty())
        {
            Notify.Warning("沒有已儲存的追蹤目標。");
            return;
        }

        importText = goal.Text;
        importPlanName = goal.Name;
        C.VentureSelectedTrackedGoalGuid = goal.Guid;
        BuildAssignmentPreview(goal.AnalyzeRecipes);
        if(importAssignments.Count == 0)
        {
            Notify.Warning(importError.IsNullOrEmpty() ? "追蹤目標沒有可套用的分配結果。" : importError);
            return;
        }

        ApplyAssignmentsToRetainers();
    }

    private void ApplyAssignmentsToRetainers()
    {
        if(importAssignments.Count == 0)
        {
            Notify.Warning("沒有可套用的分配結果。");
            return;
        }

        var planName = importPlanName.IsNullOrEmpty() ? "匯入探索計畫" : importPlanName;
        foreach(var assignment in importAssignments)
        {
            var adata = Utils.GetAdditionalData(assignment.Character.CID, assignment.Retainer.Name);
            var plan = new VenturePlan()
            {
                Name = planName,
                PlanCompleteBehavior = PlanCompleteBehavior.Do_nothing
            };
            foreach(var entry in assignment.Entries)
            {
                plan.List.Add(new(entry.Task, entry.Runs));
            }

            if(importReplaceCurrentPlan)
            {
                adata.VenturePlan = plan;
            }
            else
            {
                if(adata.VenturePlan.Name.IsNullOrEmpty()) adata.VenturePlan.Name = planName;
                foreach(var entry in plan.List)
                {
                    adata.VenturePlan.List.Add(entry);
                }
                adata.VenturePlan.PlanCompleteBehavior = plan.PlanCompleteBehavior;
            }

            adata.VenturePlanIndex = 0;
            if(importEnablePlannerOnApply) adata.EnablePlanner = true;
        }

        Notify.Success($"已套用分配到 {importAssignments.Count} 位僱員。");
        EzConfig.Save();
    }

    private void ApplyImportToPlan(AdditionalRetainerData adata, bool saveOnly)
    {
        var matched = importPreview.Where(x => x.Task != null && x.Runs > 0).ToList();
        if(matched.Count == 0)
        {
            Notify.Warning("沒有可套用的探索項目。");
            return;
        }

        var plan = new VenturePlan()
        {
            Name = importPlanName.IsNullOrEmpty() ? "匯入探索計畫" : importPlanName,
            PlanCompleteBehavior = PlanCompleteBehavior.Do_nothing
        };
        foreach(var entry in matched)
        {
            plan.List.Add(new(entry.Task.Value, entry.Runs));
        }

        if(saveOnly)
        {
            C.SavedPlans.RemoveAll(x => x.Name == plan.Name);
            C.SavedPlans.Add(plan.JSONClone());
            Notify.Success($"已儲存共用計畫：{plan.Name}");
            return;
        }

        if(importReplaceCurrentPlan)
        {
            adata.VenturePlan = plan;
        }
        else
        {
            adata.VenturePlan.Name = plan.Name;
            foreach(var entry in plan.List)
            {
                adata.VenturePlan.List.Add(entry);
            }
            adata.VenturePlan.PlanCompleteBehavior = plan.PlanCompleteBehavior;
        }
        adata.VenturePlanIndex = 0;
        Notify.Success($"已套用 {matched.Count} 筆探索到目前僱員。");
    }

    private static List<VentureRequestedItem> ParseImportLines(string text)
    {
        Dictionary<string, VentureRequestedItem> result = new(StringComparer.OrdinalIgnoreCase);
        foreach(var rawLine in text.Replace("\r", "\n").Split('\n', StringSplitOptions.RemoveEmptyEntries))
        {
            var line = CleanupImportLine(rawLine);
            if(line.IsNullOrEmpty()) continue;
            var parsed = TryParseImportLine(line);
            if(parsed == null) continue;
            var name = parsed.Value.Name;
            var amount = parsed.Value.Amount;
            if(name.IsNullOrEmpty() || amount <= 0) continue;
            var itemId = FindItemIdByName(name);
            var key = itemId == 0 ? name : $"#{itemId}";
            if(result.TryGetValue(key, out var existing)) existing.Amount += amount;
            else result[key] = new(name, amount, itemId);
        }
        return result.Values.ToList();
    }

    private static (string Name, int Amount)? TryParseImportLine(string line)
    {
        var patterns = new[]
        {
            @"^(?<name>.+?)\s*[xX×＊*]\s*(?<amount>[0-9][0-9,]*)$",
            @"^(?<amount>[0-9][0-9,]*)\s*[xX×＊*]?\s*(?<name>.+)$",
            @"^(?<name>.+?)\s*[:：]\s*(?<amount>[0-9][0-9,]*)$",
            @"^(?<name>.+?)\s+(?<amount>[0-9][0-9,]*)$",
        };
        foreach(var pattern in patterns)
        {
            var match = Regex.Match(line, pattern);
            if(!match.Success) continue;
            var name = CleanupItemName(match.Groups["name"].Value);
            if(!int.TryParse(match.Groups["amount"].Value.Replace(",", ""), NumberStyles.Integer, CultureInfo.InvariantCulture, out var amount)) continue;
            return (name, amount);
        }
        return null;
    }

    private static string CleanupImportLine(string line)
    {
        line = Regex.Replace(line.Trim(), @"^[\-\*\u2022\[\]\(\)\s]+", "");
        line = Regex.Replace(line, @"\s+", " ");
        line = Regex.Replace(line, @"\s*\([0-9,]+\s*/\s*[0-9,]+\)\s*", " ");
        return line.Trim();
    }

    private static string CleanupItemName(string name)
    {
        name = Regex.Replace(name.Trim(), @"^[\-\*\u2022\[\]\(\)\s]+", "");
        name = Regex.Replace(name, @"\s*\([HQNQhq nq]+\)\s*", " ");
        name = name.Replace("", "").Replace("", "");
        name = Regex.Replace(name, @"\s+", " ");
        return name.Trim();
    }

    private static string NormalizeItemName(string name)
    {
        return CleanupItemName(name).Replace(" ", "").ToLowerInvariant();
    }

    private sealed class VentureAssignmentCandidate
    {
        public OfflineCharacterData Character;
        public OfflineRetainerData Retainer;
        public RetainerTask Task;
        public int Yield;
    }

    private sealed class ActiveVentureCredit
    {
        public int Amount;
        public string Label = "";
    }

    private sealed class VentureRequestedItem
    {
        public string Name;
        public int Amount;
        public uint ItemId;

        public VentureRequestedItem(string name, int amount, uint itemId)
        {
            Name = name;
            Amount = amount;
            ItemId = itemId;
        }
    }

    private sealed class VentureRetainerAssignment
    {
        public OfflineCharacterData Character;
        public OfflineRetainerData Retainer;
        public List<VentureAssignedEntry> Entries = [];
        public int TotalRuns => Entries.Sum(x => x.Runs);

        public void Add(RetainerTask task, string itemName, int runs, int yield)
        {
            var entry = Entries.FirstOrDefault(x => x.Task.RowId == task.RowId);
            if(entry == null)
            {
                entry = new()
                {
                    Task = task,
                    ItemName = itemName,
                    Yield = yield
                };
                Entries.Add(entry);
            }
            entry.Runs += runs;
        }
    }

    private sealed class VentureAssignedEntry
    {
        public RetainerTask Task;
        public string ItemName = "";
        public int Runs;
        public int Yield;
        public int Provided => Runs * Yield;
    }

    private sealed class VentureImportEntry
    {
        public string ItemName = "";
        public int Amount;
        public RetainerTask? Task;
        public int Yield;
        public int Runs;
        public bool Available;
        public string Status = "";
    }
}

using AutoRetainerAPI.Configuration;
using Dalamud.Plugin.Ipc.Exceptions;
using ECommons.Throttlers;

namespace AutoRetainer.Scheduler;

internal static class Artisan
{
    internal static bool IsListRunning => Svc.PluginInterface.GetIpcSubscriber<bool>("Artisan.IsListRunning").InvokeFunc();
    internal static bool IsListPaused => Svc.PluginInterface.GetIpcSubscriber<bool>("Artisan.IsListPaused").InvokeFunc();
    internal static bool GetStopRequest => Svc.PluginInterface.GetIpcSubscriber<bool>("Artisan.GetStopRequest").InvokeFunc();
    internal static bool GetEnduranceStatus => Svc.PluginInterface.GetIpcSubscriber<bool>("Artisan.GetEnduranceStatus").InvokeFunc();
    internal static void SetEnduranceStatus(bool b)
    {
        Svc.PluginInterface.GetIpcSubscriber<bool, object>("Artisan.SetEnduranceStatus").InvokeAction(b);
    }

    internal static void SetListPause(bool b)
    {
        Svc.PluginInterface.GetIpcSubscriber<bool, object>("Artisan.SetListPause").InvokeAction(b);
    }

    internal static void SetStopRequest(bool b)
    {
        Svc.PluginInterface.GetIpcSubscriber<bool, object>("Artisan.SetStopRequest").InvokeAction(b);
    }

    internal static bool WasPaused = false;

    internal static void ArtisanTick()
    {
        if(C.ArtisanIntegration)
        {
            if(IsCurrentlyOperating() && MultiMode.EnsureCharacterValidity(true))
            {
                try
                {
                    var bell = Utils.GetReachableRetainerBell(true);
                    if(AnyRetainersAvailable() && bell != null)
                    {
                        PauseArtisanForRetainers();

                        if(!SchedulerMain.PluginEnabled || SchedulerMain.Reason != PluginEnableReason.Artisan)
                        {
                            SchedulerMain.EnablePlugin(PluginEnableReason.Artisan);
                            DebugLog($"Enabling AutoRetainer because of Artisan integration");
                        }
                    }
                }
                catch(IpcNotReadyError) { }
                catch(Exception ex)
                {
                    {
                        ex.Log();
                    }
                }
            }
            if(!AnyRetainersAvailable() && WasPaused)
            {
                if(IsOccupied())
                {
                    EzThrottler.Throttle("ArtisanCanReenableOccupied", 2500, true);
                }
                if(EzThrottler.Check("ArtisanCanReenableOccupied"))
                {
                    ResumeArtisanAfterRetainers();
                }
            }
        }
    }

    private static void PauseArtisanForRetainers()
    {
        try
        {
            if(!WasPaused)
            {
                WasPaused = true;
                DebugLog("Pausing Artisan because retainers are ready");
            }

            if(!GetStopRequest)
            {
                SetStopRequest(true);
            }

            if(IsListRunning && !IsListPaused)
            {
                SetListPause(true);
            }
        }
        catch(IpcNotReadyError) { }
        catch(Exception ex)
        {
            ex.LogWarning();
        }
    }

    private static void ResumeArtisanAfterRetainers()
    {
        try
        {
            WasPaused = false;
            DebugLog("Resuming Artisan after retainer processing");

            if(GetStopRequest)
            {
                SetStopRequest(false);
            }

            if(IsListPaused)
            {
                SetListPause(false);
            }
        }
        catch(IpcNotReadyError) { }
        catch(Exception ex)
        {
            ex.LogWarning();
        }
    }

    internal static bool AnyRetainersAvailable()
    {
        if(C.OfflineData.TryGetFirst(x => x.CID == Svc.ClientState.LocalContentId, out var data))
        {
            return data.GetEnabledRetainers().Any(z => z.GetVentureSecondsRemaining() <= C.UnsyncCompensation);
        }
        return false;
    }

    internal static bool IsCurrentlyOperating()
    {
        try
        {
            return IsListRunning || GetEnduranceStatus;
        }
        catch(IpcNotReadyError) { }
        catch(Exception ex)
        {
            ex.LogWarning();
        }
        return false;
    }
}

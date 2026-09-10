using System;
using System.Collections.Generic;
using System.Linq;

namespace AutoRetainer.Helpers;

internal static class RetainerCompatibility
{
    internal static string[] MenuCandidates(IEnumerable<string> values) =>
        values.Where(x => !string.IsNullOrWhiteSpace(x)).Distinct(StringComparer.Ordinal).ToArray();

    internal static bool CanRecoverMenu(bool retainerActive, bool voyageActive, bool manualWindow, bool suppressed) =>
        !manualWindow && !suppressed && (retainerActive || voyageActive);
}

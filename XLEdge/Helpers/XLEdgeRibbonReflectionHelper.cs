using System;
using System.Reflection;
using XLEdge.Utilities;

namespace XLEdge.Helpers
{
    internal static class XLEdgeRibbonReflectionHelper
    {
        /// <summary>
        /// Safely attempts to get a ribbon control instance from the AddinModule.
        /// Tries public property, public field, then (as a last resort) non-public field.
        /// Returns null if not found or access is not allowed.
        /// </summary>
        public static object GetRibbonControl(object addinModuleInstance, string controlName)
        {
            if (addinModuleInstance == null || string.IsNullOrWhiteSpace(controlName))
                return null;

            var t = addinModuleInstance.GetType();

            // 1) Try public property
            try
            {
                var prop = t.GetProperty(controlName, BindingFlags.Instance | BindingFlags.Public);
                if (prop != null && prop.CanRead)
                {
                    try
                    {
                        return prop.GetValue(addinModuleInstance, null);
                    }
                    catch (TargetInvocationException)
                    {
                        // Deliberately silent: routine fallthrough to the public/non-public field
                        // lookups below - the final "not found" outcome is logged once at the end.
                    }
                    catch (Exception)
                    {
                        // Deliberately silent: same routine fallthrough as above.
                    }
                }
            }
            catch (Exception)
            {
                // Deliberately silent: property lookup itself failed, falls through to field lookups.
            }

            // 2) Try public field
            try
            {
                var field = t.GetField(controlName, BindingFlags.Instance | BindingFlags.Public);
                if (field != null)
                {
                    try
                    {
                        return field.GetValue(addinModuleInstance);
                    }
                    catch (Exception)
                    {
                        // Deliberately silent: routine fallthrough to the non-public field lookup below.
                    }
                }
            }
            catch (Exception)
            {
                // Deliberately silent: public-field lookup itself failed, falls through to non-public field.
            }

            // 3) LAST RESORT: try non-public field (fragile — log and return)
            try
            {
                var nonPublicField = t.GetField(controlName, BindingFlags.Instance | BindingFlags.NonPublic);
                if (nonPublicField != null)
                {
                    try
                    {
                        var val = nonPublicField.GetValue(addinModuleInstance);
                        // optional: only return if not null and type looks like a ribbon control
                        if (val != null)
                            return val;
                    }
                    catch (FieldAccessException ex)
                    {
                        // Not allowed to access non-public member in this environment - worth knowing
                        // about since it means this control can never be resolved via reflection here.
                        LogUtility.LogWarn($"{nameof(GetRibbonControl)}: not allowed to access non-public field '{controlName}' - {ex.Message}");
                        return null;
                    }
                    catch (Exception ex)
                    {
                        // Safe to ignore: swallow and return null, caller treats this as "control not found".
                        LogUtility.LogDebug($"{nameof(GetRibbonControl)}: failed reading non-public field '{controlName}' - {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogDebug($"{nameof(GetRibbonControl)}: non-public-field lookup for '{controlName}' failed - {ex.Message}");
            }

            return null;
        }
    }
}

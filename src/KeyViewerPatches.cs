using HarmonyLib;

namespace KeyViewer
{
    /// <summary>
    /// Patches for the Key Viewer tweak.
    /// </summary>
    internal static class KeyViewerPatches
    {

        [HarmonyPatch(typeof(scrController), "CountValidKeysPressed")]
        private static class CountValidKeysPressedPatch
        {
            [HarmonyBefore("KeyViewer.KeyViewer")]
            public static bool Prefix(ref int __result) {
                // Stop player inputs while we're editing the keys
                if (Main.Settings.IsListening) {
                    __result = 0;
                    return false;
                }

                return true;
            }
        }
        //[HarmonyPatch(typeof(scnEditor), "Play")]
        //private static class Patch1
        //{
        //    public static void Postfix() => KeyViewerTweak.ResetRainbow();
        //}
        //[HarmonyPatch(typeof(scrPlanet), "Rewind")]
        //private static class Patch2
        //{
        //    public static void Postfix() => KeyViewerTweak.ResetRainbow();
        //}
        [HarmonyPatch(typeof(scrPlanet), "LoadPlanetColor")]
        private static class Patch
        {
            [HarmonyPostfix]
            public static void Reset() => KeyViewerTweak.ResetRainbow();
        }
        //[HarmonyPatch(typeof(scrUIController), "FadeFromBlack")]
        //private static class UIPatch1
        //{
        //    public static void Postfix() => KeyViewerTweak.ResetRainbow();
        //}
        //[HarmonyPatch(typeof(scrUIController), "FadeToBlack")]
        //private static class UIPatch2
        //{
        //    public static void Postfix() => KeyViewerTweak.ResetRainbow();
        //}
        //[HarmonyPatch(typeof(scrUIController), "WipeFromBlack")]
        //private static class UIPatch3
        //{
        //    public static void Postfix() => KeyViewerTweak.ResetRainbow();
        //}
        //[HarmonyPatch(typeof(scrUIController), "WipeToBlack")]
        //private static class UIPatch4
        //{
        //    public static void Postfix() => KeyViewerTweak.ResetRainbow();
        //}
    }
}

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
    }
}

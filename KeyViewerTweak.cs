using System;
using System.Collections.Generic;
using UnityEngine;
using static UnityModManagerNet.UnityModManager;

namespace KeyViewer
{
    public class KeyViewerTweak
    {
        /// <summary>
        /// Keys that should not be listened to.
        /// </summary>
        public static readonly ISet<KeyCode> SKIPPED_KEYS = new HashSet<KeyCode>() {
            KeyCode.Mouse0,
            KeyCode.Mouse1,
            KeyCode.Mouse2,
            KeyCode.Mouse3,
            KeyCode.Mouse4,
            KeyCode.Mouse5,
            KeyCode.Mouse6,
            KeyCode.Escape,
        };

        private static KeyViewerProfile CurrentProfile { get => Main.Settings.CurrentProfile; }

        private static Dictionary<KeyCode, bool> keyState;
        public static KeyViewer keyViewer;

        /// <inheritdoc/>
        public static void OnUpdate(ModEntry modEntry, float deltaTime) {
            KeyViewer.UpdateKPSTotal();
            UpdateRegisteredKeys();
            UpdateKeyState();
        }
        static bool Ignore(KeyCode code) => Main.Settings.IgnoreSkippedKeys ? false : SKIPPED_KEYS.Contains(code);
        private static void UpdateRegisteredKeys() {
            if (!Main.Settings.IsListening) {
                return;
            }
            bool changed = false;
            foreach (KeyCode code in Enum.GetValues(typeof(KeyCode))) {
                // Skip key if not pressed or should be skipped
                if (!Input.GetKeyDown(code) || Ignore(code)) {
                    continue;
                }

                // Register/unregister the key
                if (CurrentProfile.ActiveKeys.Contains(code)) {
                    CurrentProfile.ActiveKeys.Remove(code);
                    changed = true;
                } else {
                    CurrentProfile.ActiveKeys.Add(code);
                    changed = true;
                }
            }
            if (changed) {
                KeyViewer.UpdateKeys();
            }
        }

        private static void UpdateKeyState() {
            UpdateViewerVisibility();
            foreach (KeyCode code in CurrentProfile.ActiveKeys) {
                keyState[code] = Input.GetKey(code);
                if (Input.GetKeyDown(code))
                    Calculate.Press();
            }
            keyViewer.UpdateState(keyState);
        }

        /// <inheritdoc/>
        public static void OnHideGUI(ModEntry modEntry) {
            Main.Settings.IsListening = false;
        }
        public static KeyViewerSettings Settings => Main.Settings;
        public static bool changedk;
        public static bool changedt;
        public static int f;
        public static bool DrawVector(ref Vector2 vec2, string label)
        {
            Vector2 cache = vec2;
            bool changed = false;
            GUILayout.BeginHorizontal();
            GUILayout.Label(label);
            GUILayout.EndHorizontal();
            vec2.x = GUILayout.HorizontalSlider(vec2.x, 0, 1920);
            vec2.y = GUILayout.HorizontalSlider(vec2.y, 0, 1080);
            changed = vec2 != cache;
            return changed;
        }
        /// <inheritdoc/>
        public static void OnSettingsGUI(ModEntry modEntry) {
#if DEBUG
            Settings.DisplayKPS = GUILayout.Toggle(Settings.DisplayKPS, "Display KPS");
            Settings.DisplayTotal = GUILayout.Toggle(Settings.DisplayTotal, "Display Total");
            if (Settings.DisplayTotal != changedt)
            {
                changedt = Settings.DisplayTotal;
                KeyViewer.UpdateLayout();
            }
            if (Settings.DisplayKPS != changedk)
            {
                changedk = Settings.DisplayKPS;
                KeyViewer.UpdateLayout();
            }
            UI.DrawIntField(ref f, "Test");
            KeyViewer.keyCounts[KeyCode.F] = f;
#endif
            Settings.SaveKeyCounts = GUILayout.Toggle(Settings.SaveKeyCounts, Main.Settings.Language == LanguageEnum.KOREAN ? Lang.SAVE_KEYCOUNTS : Lang.SAVE_KEYCOUNTS_ENG);
            GUILayout.BeginHorizontal();
            if(UI.DrawIntField(ref Settings.UpdateRate, "Update Rate")) 
                Calculate.ChangeUpdateRate(Settings.UpdateRate);
            GUILayout.FlexibleSpace();
            GUILayout.EndHorizontal();
            DrawProfileSettingsGUI();
            GUILayout.Space(12f);
            MoreGUILayout.HorizontalLine(1f, 400f);
            GUILayout.Space(8f);
            DrawKeyRegisterSettingsGUI();
            GUILayout.Space(8f);
            DrawKeyViewerSettingsGUI();
        }

        private static void DrawProfileSettingsGUI() {
            GUILayout.Space(4f);

            // New, Duplicate, Delete buttons
            GUILayout.BeginHorizontal();
            if (GUILayout.Button(
                TweakStrings.Get(TranslationKeys.KeyViewer.NEW))) {
                Main.Settings.Profiles.Add(new KeyViewerProfile());
                Main.Settings.ProfileIndex = Main.Settings.Profiles.Count - 1;
                Main.Settings.CurrentProfile.Name += "Profile " + Main.Settings.Profiles.Count;
                KeyViewer.Profile = Main.Settings.CurrentProfile;
            }
            if (GUILayout.Button(
                TweakStrings.Get(TranslationKeys.KeyViewer.DUPLICATE))) {
                Main.Settings.Profiles.Add(Main.Settings.CurrentProfile.Copy());
                Main.Settings.ProfileIndex = Main.Settings.Profiles.Count - 1;
                Main.Settings.CurrentProfile.Name += " Copy";
                KeyViewer.Profile = Main.Settings.CurrentProfile;
            }
            if (Main.Settings.Profiles.Count > 1
                && GUILayout.Button(
                    TweakStrings.Get(TranslationKeys.KeyViewer.DELETE))) {
                Main.Settings.Profiles.RemoveAt(Main.Settings.ProfileIndex);
                Main.Settings.ProfileIndex =
                    Math.Min(Main.Settings.ProfileIndex, Main.Settings.Profiles.Count - 1);
                KeyViewer.Profile = Main.Settings.CurrentProfile;
            }
            GUILayout.FlexibleSpace();
            GUILayout.EndHorizontal();

            GUILayout.Space(4f);

            // Profile name
            Main.Settings.CurrentProfile.Name =
                MoreGUILayout.NamedTextField(
                    TweakStrings.Get(TranslationKeys.KeyViewer.PROFILE_NAME),
                    Main.Settings.CurrentProfile.Name,
                    400f);

            // Profile list
            GUILayout.Label(TweakStrings.Get(TranslationKeys.KeyViewer.PROFILES));
            int selected = Main.Settings.ProfileIndex;
            if (MoreGUILayout.ToggleList(Main.Settings.Profiles, ref selected, p => p.Name)) {
                Main.Settings.ProfileIndex = selected;
                KeyViewer.Profile = Main.Settings.CurrentProfile;
            }
        }

        private static void DrawKeyRegisterSettingsGUI() {
            GUILayout.BeginHorizontal();
            Main.Settings.IgnoreSkippedKeys = GUILayout.Toggle(Main.Settings.IgnoreSkippedKeys, Main.Settings.Language == LanguageEnum.KOREAN ? Lang.IGNORE_SKIPPED_KEYS : Lang.IGNORE_SKIPPED_KEYS_ENG);
            GUILayout.EndHorizontal();
            // List of registered keys
            GUILayout.Label(TweakStrings.Get(TranslationKeys.KeyViewer.REGISTERED_KEYS));
            GUILayout.BeginHorizontal();
            GUILayout.Space(20f);
            GUILayout.BeginVertical();
            GUILayout.Space(8f);
            GUILayout.EndVertical();
            foreach (KeyCode code in CurrentProfile.ActiveKeys) {
                GUILayout.Label(code.ToString());
                GUILayout.Space(8f);
            }
            GUILayout.FlexibleSpace();
            GUILayout.EndHorizontal();
            GUILayout.Space(12f);

            // Record keys toggle
            GUILayout.BeginHorizontal();
            if (Main.Settings.IsListening) {
                if (GUILayout.Button(TweakStrings.Get(TranslationKeys.KeyViewer.DONE))) {
                    Main.Settings.IsListening = false;
                }
                GUILayout.Label(TweakStrings.Get(TranslationKeys.KeyViewer.PRESS_KEY_REGISTER));
            } else {
                if (GUILayout.Button(TweakStrings.Get(TranslationKeys.KeyViewer.CHANGE_KEYS))) {
                    Main.Settings.IsListening = true;
                }
                if (GUILayout.Button("Mouse0"))
                {
                    if (!CurrentProfile.ActiveKeys.Contains(KeyCode.Mouse0))
                        CurrentProfile.ActiveKeys.Add(KeyCode.Mouse0);
                    else CurrentProfile.ActiveKeys.Remove(KeyCode.Mouse0);
                    KeyViewer.UpdateKeys();
                }
                if (GUILayout.Button("Mouse1"))
                {
                    if (!CurrentProfile.ActiveKeys.Contains(KeyCode.Mouse1))
                        CurrentProfile.ActiveKeys.Add(KeyCode.Mouse1);
                    else CurrentProfile.ActiveKeys.Remove(KeyCode.Mouse1);
                    KeyViewer.UpdateKeys();
                }
            }
            GUILayout.FlexibleSpace();
            GUILayout.EndHorizontal();
        }

        private static void DrawKeyViewerSettingsGUI() {
            MoreGUILayout.BeginIndent();

            // Show only in gameplay toggle
            CurrentProfile.ViewerOnlyGameplay =
                GUILayout.Toggle(
                    CurrentProfile.ViewerOnlyGameplay,
                    TweakStrings.Get(TranslationKeys.KeyViewer.VIEWER_ONLY_GAMEPLAY));

            // Animate keys toggle
            CurrentProfile.AnimateKeys =
                GUILayout.Toggle(
                    CurrentProfile.AnimateKeys,
                    TweakStrings.Get(TranslationKeys.KeyViewer.ANIMATE_KEYS));

            // Key press total toggle
            bool newShowTotal =
                GUILayout.Toggle(
                    CurrentProfile.ShowKeyPressTotal,
                    TweakStrings.Get(TranslationKeys.KeyViewer.SHOW_KEY_PRESS_TOTAL));
            if (newShowTotal != CurrentProfile.ShowKeyPressTotal) {
                CurrentProfile.ShowKeyPressTotal = newShowTotal;
                KeyViewer.UpdateLayout();
            }

            // Size slider
            float newSize =
                MoreGUILayout.NamedSlider(
                    TweakStrings.Get(TranslationKeys.KeyViewer.KEY_VIEWER_SIZE),
                    CurrentProfile.KeyViewerSize,
                    10f,
                    200f,
                    300f,
                    roundNearest: 1f);
            if (newSize != CurrentProfile.KeyViewerSize) {
                CurrentProfile.KeyViewerSize = newSize;
                KeyViewer.UpdateLayout();
            }

            // X position slider
            float newX =
                MoreGUILayout.NamedSlider(
                    TweakStrings.Get(TranslationKeys.KeyViewer.KEY_VIEWER_X_POS),
                    CurrentProfile.KeyViewerXPos,
                    0f,
                    1f,
                    300f,
                    roundNearest: 0.01f,
                    valueFormat: "{0:0.##}");
            if (newX != CurrentProfile.KeyViewerXPos) {
                CurrentProfile.KeyViewerXPos = newX;
                KeyViewer.UpdateLayout();
            }

            // Y position slider
            float newY =
                MoreGUILayout.NamedSlider(
                    TweakStrings.Get(TranslationKeys.KeyViewer.KEY_VIEWER_Y_POS),
                    CurrentProfile.KeyViewerYPos,
                    0f,
                    1f,
                    300f,
                    roundNearest: 0.01f,
                    valueFormat: "{0:0.##}");
            if (newY != CurrentProfile.KeyViewerYPos) {
                CurrentProfile.KeyViewerYPos = newY;
                KeyViewer.UpdateLayout();
            }

            GUILayout.Space(8f);

            Color newPressed, newReleased;
            string newPressedHex, newReleasedHex;

            // Outline color header
            GUILayout.BeginHorizontal();
            GUILayout.Label(
                TweakStrings.Get(TranslationKeys.KeyViewer.PRESSED_OUTLINE_COLOR),
                GUILayout.Width(200f));
            GUILayout.FlexibleSpace();
            GUILayout.Space(8f);
            GUILayout.Label(
                TweakStrings.Get(TranslationKeys.KeyViewer.RELEASED_OUTLINE_COLOR),
                GUILayout.Width(200f));
            GUILayout.FlexibleSpace();
            GUILayout.Space(20f);
            GUILayout.EndHorizontal();
            MoreGUILayout.BeginIndent();

            // Outline color RGBA sliders
            (newPressed, newReleased) =
                MoreGUILayout.ColorRgbaSlidersPair(
                    CurrentProfile.PressedOutlineColor, CurrentProfile.ReleasedOutlineColor);
            if (newPressed != CurrentProfile.PressedOutlineColor) {
                CurrentProfile.PressedOutlineColor = newPressed;
                KeyViewer.UpdateLayout();
            }
            if (newReleased != CurrentProfile.ReleasedOutlineColor) {
                CurrentProfile.ReleasedOutlineColor = newReleased;
                KeyViewer.UpdateLayout();
            }

            // Outline color hex
            (newPressedHex, newReleasedHex) =
                MoreGUILayout.NamedTextFieldPair(
                    "Hex:",
                    "Hex:",
                    CurrentProfile.PressedOutlineColorHex,
                    CurrentProfile.ReleasedOutlineColorHex,
                    100f,
                    40f);
            if (newPressedHex != CurrentProfile.PressedOutlineColorHex
                && ColorUtility.TryParseHtmlString($"#{newPressedHex}", out newPressed)) {
                CurrentProfile.PressedOutlineColor = newPressed;
                KeyViewer.UpdateLayout();
            }
            if (newReleasedHex != CurrentProfile.ReleasedOutlineColorHex
                && ColorUtility.TryParseHtmlString($"#{newReleasedHex}", out newReleased)) {
                CurrentProfile.ReleasedOutlineColor = newReleased;
                KeyViewer.UpdateLayout();
            }
            CurrentProfile.PressedOutlineColorHex = newPressedHex;
            CurrentProfile.ReleasedOutlineColorHex = newReleasedHex;

            MoreGUILayout.EndIndent();

            GUILayout.Space(8f);

            // Background color header
            GUILayout.BeginHorizontal();
            GUILayout.Label(
                TweakStrings.Get(TranslationKeys.KeyViewer.PRESSED_BACKGROUND_COLOR),
                GUILayout.Width(200f));
            GUILayout.FlexibleSpace();
            GUILayout.Space(8f);
            GUILayout.Label(
                TweakStrings.Get(TranslationKeys.KeyViewer.RELEASED_BACKGROUND_COLOR),
                GUILayout.Width(200f));
            GUILayout.FlexibleSpace();
            GUILayout.Space(20f);
            GUILayout.EndHorizontal();
            MoreGUILayout.BeginIndent();

            // Background color RGBA sliders
            (newPressed, newReleased) =
                MoreGUILayout.ColorRgbaSlidersPair(
                    CurrentProfile.PressedBackgroundColor, CurrentProfile.ReleasedBackgroundColor);
            if (newPressed != CurrentProfile.PressedBackgroundColor) {
                CurrentProfile.PressedBackgroundColor = newPressed;
                KeyViewer.UpdateLayout();
            }
            if (newReleased != CurrentProfile.ReleasedBackgroundColor) {
                CurrentProfile.ReleasedBackgroundColor = newReleased;
                KeyViewer.UpdateLayout();
            }

            // Background color hex
            (newPressedHex, newReleasedHex) =
                MoreGUILayout.NamedTextFieldPair(
                    "Hex:",
                    "Hex:",
                    CurrentProfile.PressedBackgroundColorHex,
                    CurrentProfile.ReleasedBackgroundColorHex,
                    100f,
                    40f);
            if (newPressedHex != CurrentProfile.PressedBackgroundColorHex
                && ColorUtility.TryParseHtmlString($"#{newPressedHex}", out newPressed)) {
                CurrentProfile.PressedBackgroundColor = newPressed;
                KeyViewer.UpdateLayout();
            }
            if (newReleasedHex != CurrentProfile.ReleasedBackgroundColorHex
                && ColorUtility.TryParseHtmlString($"#{newReleasedHex}", out newReleased)) {
                CurrentProfile.ReleasedBackgroundColor = newReleased;
                KeyViewer.UpdateLayout();
            }
            CurrentProfile.PressedBackgroundColorHex = newPressedHex;
            CurrentProfile.ReleasedBackgroundColorHex = newReleasedHex;

            MoreGUILayout.EndIndent();

            GUILayout.Space(8f);

            // Text color header
            GUILayout.BeginHorizontal();
            GUILayout.Label(
                TweakStrings.Get(TranslationKeys.KeyViewer.PRESSED_TEXT_COLOR),
                GUILayout.Width(200f));
            GUILayout.FlexibleSpace();
            GUILayout.Space(8f);
            GUILayout.Label(
                TweakStrings.Get(TranslationKeys.KeyViewer.RELEASED_TEXT_COLOR),
                GUILayout.Width(200f));
            GUILayout.FlexibleSpace();
            GUILayout.Space(20f);
            GUILayout.EndHorizontal();
            MoreGUILayout.BeginIndent();

            // Text color RGBA sliders
            (newPressed, newReleased) =
                MoreGUILayout.ColorRgbaSlidersPair(
                    CurrentProfile.PressedTextColor, CurrentProfile.ReleasedTextColor);
            if (newPressed != CurrentProfile.PressedTextColor) {
                CurrentProfile.PressedTextColor = newPressed;
                KeyViewer.UpdateLayout();
            }
            if (newReleased != CurrentProfile.ReleasedTextColor) {
                CurrentProfile.ReleasedTextColor = newReleased;
                KeyViewer.UpdateLayout();
            }

            // Text color hex
            (newPressedHex, newReleasedHex) =
                MoreGUILayout.NamedTextFieldPair(
                    "Hex:",
                    "Hex:",
                    CurrentProfile.PressedTextColorHex,
                    CurrentProfile.ReleasedTextColorHex,
                    100f,
                    40f);
            if (newPressedHex != CurrentProfile.PressedTextColorHex
                && ColorUtility.TryParseHtmlString($"#{newPressedHex}", out newPressed)) {
                CurrentProfile.PressedTextColor = newPressed;
                KeyViewer.UpdateLayout();
            }
            if (newReleasedHex != CurrentProfile.ReleasedTextColorHex
                && ColorUtility.TryParseHtmlString($"#{newReleasedHex}", out newReleased)) {
                CurrentProfile.ReleasedTextColor = newReleased;
                KeyViewer.UpdateLayout();
            }
            CurrentProfile.PressedTextColorHex = newPressedHex;
            CurrentProfile.ReleasedTextColorHex = newReleasedHex;

            MoreGUILayout.EndIndent();

            MoreGUILayout.EndIndent();
        }

        /// <inheritdoc/>
        public static void OnEnable() {
            if (Main.Settings.Profiles.Count == 0) {
                Main.Settings.Profiles.Add(new KeyViewerProfile() { Name = "Default Profile" });
            }
            if (Main.Settings.ProfileIndex < 0 || Main.Settings.ProfileIndex >= Main.Settings.Profiles.Count) {
                Main.Settings.ProfileIndex = 0;
            }
            Calculate.Set();
            MigrateOldSettings();

            GameObject keyViewerObj = new GameObject();
            GameObject.DontDestroyOnLoad(keyViewerObj);
            keyViewer = keyViewerObj.AddComponent<KeyViewer>();
            KeyViewer.Profile = CurrentProfile;

            UpdateViewerVisibility();
            keyState = new Dictionary<KeyCode, bool>() { {KeyCode.None, false }, { KeyCode.Joystick1Button0, false } };
        }

        /// <summary>
        /// Migrates old KeyLimiter settings to a KeyViewer profile if there are
        /// settings to migrate.
        /// TODO: Delete this after a few releases.
        /// </summary>
        private static void MigrateOldSettings() {
            // Check if there are settings to migrate
            if (Main.Settings.PressedBackgroundColor == Color.black
                && Main.Settings.ReleasedBackgroundColor == Color.black
                && Main.Settings.PressedOutlineColor == Color.black
                && Main.Settings.ReleasedOutlineColor == Color.black
                && Main.Settings.PressedTextColor == Color.black
                && Main.Settings.ReleasedTextColor == Color.black) {
                return;
            }

            // Copy into new profile
            KeyViewerProfile profile = new KeyViewerProfile {
                Name = "Old Profile",
                ActiveKeys = new List<KeyCode>(Main.Settings.ActiveKeys),
                ViewerOnlyGameplay = Main.Settings.ViewerOnlyGameplay,
                AnimateKeys = Main.Settings.AnimateKeys,
                KeyViewerSize = Main.Settings.KeyViewerSize,
                KeyViewerXPos = Main.Settings.KeyViewerXPos,
                KeyViewerYPos = Main.Settings.KeyViewerYPos,
                PressedOutlineColor = Main.Settings.PressedOutlineColor,
                ReleasedOutlineColor = Main.Settings.ReleasedOutlineColor,
                PressedBackgroundColor = Main.Settings.PressedBackgroundColor,
                ReleasedBackgroundColor = Main.Settings.ReleasedBackgroundColor,
                PressedTextColor = Main.Settings.PressedTextColor,
                ReleasedTextColor = Main.Settings.ReleasedTextColor,
            };

            // Set current to migrated profile
            Main.Settings.Profiles.Insert(0, profile);
            Main.Settings.ProfileIndex = 0;

            // Clear old settings
            Main.Settings.PressedBackgroundColor = Color.black;
            Main.Settings.ReleasedBackgroundColor = Color.black;
            Main.Settings.PressedOutlineColor = Color.black;
            Main.Settings.ReleasedOutlineColor = Color.black;
            Main.Settings.PressedTextColor = Color.black;
            Main.Settings.ReleasedTextColor = Color.black;
        }

        /// <inheritdoc/>
        public static void OnDisable() {
            GameObject.Destroy(keyViewer.gameObject);
            Calculate.UnSet();
        }

        private static void UpdateViewerVisibility() {
            bool showViewer = true;
            if (CurrentProfile.ViewerOnlyGameplay
                && scrController.instance
                && scrConductor.instance) {
                bool playing = !scrController.instance.paused && scrConductor.instance.isGameWorld;
                showViewer &= playing;
            }
            if (showViewer != keyViewer.gameObject.activeSelf) {
                keyViewer.gameObject.SetActive(showViewer);
            }
        }
    }
}

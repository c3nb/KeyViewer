using System.Collections.Generic;
using DG.Tweening;
using UnityEngine;
using UnityEngine.UI;

namespace KeyViewer
{
    /// <summary>
    /// A key viewer that shows if a list of given keys are currently being
    /// pressed.
    /// </summary>
    public class KeyViewer : MonoBehaviour
    {
        private static readonly Dictionary<KeyCode, string> KEY_TO_STRING =
            new Dictionary<KeyCode, string>() {
                { KeyCode.Alpha0, "0" },
                { KeyCode.Alpha1, "1" },
                { KeyCode.Alpha2, "2" },
                { KeyCode.Alpha3, "3" },
                { KeyCode.Alpha4, "4" },
                { KeyCode.Alpha5, "5" },
                { KeyCode.Alpha6, "6" },
                { KeyCode.Alpha7, "7" },
                { KeyCode.Alpha8, "8" },
                { KeyCode.Alpha9, "9" },
                { KeyCode.Keypad0, "0" },
                { KeyCode.Keypad1, "1" },
                { KeyCode.Keypad2, "2" },
                { KeyCode.Keypad3, "3" },
                { KeyCode.Keypad4, "4" },
                { KeyCode.Keypad5, "5" },
                { KeyCode.Keypad6, "6" },
                { KeyCode.Keypad7, "7" },
                { KeyCode.Keypad8, "8" },
                { KeyCode.Keypad9, "9" },
                { KeyCode.KeypadPlus, "+" },
                { KeyCode.KeypadMinus, "-" },
                { KeyCode.KeypadMultiply, "*" },
                { KeyCode.KeypadDivide, "/" },
                { KeyCode.KeypadEnter, "↵" },
                { KeyCode.KeypadEquals, "=" },
                { KeyCode.KeypadPeriod, "." },
                { KeyCode.Return, "↵" },
                { KeyCode.None, " " },
                { KeyCode.Tab, "⇥" },
                { KeyCode.Backslash, "\\" },
                { KeyCode.Slash, "/" },
                { KeyCode.Minus, "-" },
                { KeyCode.Equals, "=" },
                { KeyCode.LeftBracket, "[" },
                { KeyCode.RightBracket, "]" },
                { KeyCode.Semicolon, ";" },
                { KeyCode.Comma, "," },
                { KeyCode.Period, "." },
                { KeyCode.Quote, "'" },
                { KeyCode.UpArrow, "↑" },
                { KeyCode.DownArrow, "↓" },
                { KeyCode.LeftArrow, "←" },
                { KeyCode.RightArrow, "→" },
                { KeyCode.Space, "␣" },
                { KeyCode.BackQuote, "`" },
                { KeyCode.LeftShift, "L⇧" },
                { KeyCode.RightShift, "R⇧" },
                { KeyCode.LeftControl, "LCtrl" },
                { KeyCode.RightControl, "RCtrl" },
                { KeyCode.LeftAlt, "LAlt" },
                { KeyCode.RightAlt, "AAlt" },
                { KeyCode.Delete, "Del" },
                { KeyCode.PageDown, "Pg↓" },
                { KeyCode.PageUp, "Pg↑" },
                { KeyCode.Insert, "Ins" },
                { KeyCode.Mouse0, "M0" },
                { KeyCode.Mouse1, "M1" },
                { KeyCode.Mouse2, "M2" },
                { KeyCode.Mouse3, "M3" },
                { KeyCode.Mouse4, "M4" },
                { KeyCode.Mouse5, "M5" },
                { KeyCode.Mouse6, "M6" },
            };

        private const float EASE_DURATION = 0.1f;
        private const float SHRINK_FACTOR = 0.9f;

        internal static Dictionary<KeyCode, int> keyCounts = new Dictionary<KeyCode, int>();

        private static GameObject keysObject;
        private static Dictionary<KeyCode, Image> keyBgImages;
        private static Dictionary<KeyCode, Image> keyOutlineImages;
        private static Dictionary<KeyCode, Text> keyTexts;
        private static Dictionary<KeyCode, Text> keyCountTexts;
        private static Dictionary<KeyCode, bool> keyPrevStates;
        private static RectTransform keysRectTransform;

        private static KeyViewerProfile _profile = new KeyViewerProfile();

        /// <summary>
        /// The current profile that this key viewer is using.
        /// </summary>
        public static KeyViewerProfile Profile {
            get => _profile;
            set {
                _profile = value;
                UpdateKeys();
            }
        }

        /// <summary>
        /// Unity's Awake lifecycle event handler. Creates the key viewer.
        /// </summary>
        protected void Awake() {
            Canvas mainCanvas = gameObject.AddComponent<Canvas>();
            mainCanvas.renderMode = RenderMode.ScreenSpaceOverlay;
            mainCanvas.sortingOrder = 10001;
            CanvasScaler scaler = gameObject.AddComponent<CanvasScaler>();
            scaler.referenceResolution = new Vector2(1920, 1080);

            UpdateKeys();
        }

        /// <summary>
        /// Updates what keys are displayed on the key viewer.
        /// </summary>
        public static void UpdateKeys() {
            if (keysObject) {
                Destroy(keysObject);
            }

            keysObject = new GameObject();
            keysObject.transform.SetParent(KeyViewerTweak.keyViewer.transform);
            keysObject.AddComponent<Canvas>();
            keysRectTransform = keysObject.GetComponent<RectTransform>();

            keyBgImages = new Dictionary<KeyCode, Image>();
            keyOutlineImages = new Dictionary<KeyCode, Image>();
            keyTexts = new Dictionary<KeyCode, Text>();
            keyCountTexts = new Dictionary<KeyCode, Text>();
            keyPrevStates = new Dictionary<KeyCode, bool>();
            foreach (KeyCode code in Profile.ActiveKeys) {
                KeyViewerTweak.keyViewer.GenerateKey(code);
            }
            if (Main.Settings.DisplayKPS)
            {
                GameObject keyBgObj = new GameObject();
                keyBgObj.transform.SetParent(keysObject.transform);
                Image bgImage = keyBgObj.AddComponent<Image>();
                bgImage.sprite = Asset.KeyBackgroundSprite;
                bgImage.color = Profile.ReleasedBackgroundColor;
                keyBgImages[KeyCode.None] = bgImage;
                bgImage.type = Image.Type.Sliced;

                GameObject keyOutlineObj = new GameObject();
                keyOutlineObj.transform.SetParent(keysObject.transform);
                Image outlineImage = keyOutlineObj.AddComponent<Image>();
                outlineImage.sprite = Asset.KeyOutlineSprite;
                outlineImage.color = Profile.ReleasedOutlineColor;
                keyOutlineImages[KeyCode.None] = outlineImage;
                outlineImage.type = Image.Type.Sliced;

                GameObject keyTextObj = new GameObject();
                keyTextObj.transform.SetParent(keysObject.transform);
                Text text = keyTextObj.AddComponent<Text>();
                text.font = RDString.GetFontDataForLanguage(SystemLanguage.English).font;
                text.color = Profile.ReleasedTextColor;
                text.alignment = TextAnchor.UpperCenter;
                text.text = "KPS";
                keyTexts[KeyCode.None] = text;

                GameObject keyCountTextObj = new GameObject();
                keyCountTextObj.transform.SetParent(keysObject.transform);
                Text countText = keyCountTextObj.AddComponent<Text>();
                countText.resizeTextForBestFit = true;
                countText.resizeTextMinSize = 0;
                countText.resizeTextMaxSize = 40;
                countText.font = RDString.GetFontDataForLanguage(SystemLanguage.English).font;
                countText.color = Profile.ReleasedTextColor;
                countText.alignment = TextAnchor.LowerCenter;
                countText.text = Calculate.kps + "";
                keyCountTexts[KeyCode.None] = countText;

                keyPrevStates[KeyCode.None] = false;
            }

            if (Main.Settings.DisplayTotal)
            {
                GameObject keyBgObj = new GameObject();
                keyBgObj.transform.SetParent(keysObject.transform);
                Image bgImage = keyBgObj.AddComponent<Image>();
                bgImage.sprite = Asset.KeyBackgroundSprite;
                bgImage.color = Profile.ReleasedBackgroundColor;
                keyBgImages[KeyCode.Joystick1Button0] = bgImage;
                bgImage.type = Image.Type.Sliced;

                GameObject keyOutlineObj = new GameObject();
                keyOutlineObj.transform.SetParent(keysObject.transform);
                Image outlineImage = keyOutlineObj.AddComponent<Image>();
                outlineImage.sprite = Asset.KeyOutlineSprite;
                outlineImage.color = Profile.ReleasedOutlineColor;
                keyOutlineImages[KeyCode.Joystick1Button0] = outlineImage;
                outlineImage.type = Image.Type.Sliced;

                GameObject keyTextObj = new GameObject();
                keyTextObj.transform.SetParent(keysObject.transform);
                Text text = keyTextObj.AddComponent<Text>();
                text.font = RDString.GetFontDataForLanguage(SystemLanguage.English).font;
                text.color = Profile.ReleasedTextColor;
                text.alignment = TextAnchor.UpperCenter;
                text.text = "Total";
                keyTexts[KeyCode.Joystick1Button0] = text;

                GameObject keyCountTextObj = new GameObject();
                keyCountTextObj.transform.SetParent(keysObject.transform);
                Text countText = keyCountTextObj.AddComponent<Text>();
                countText.resizeTextForBestFit = true;
                countText.resizeTextMinSize = 0;
                countText.resizeTextMaxSize = 40;
                countText.font = RDString.GetFontDataForLanguage(SystemLanguage.English).font;
                countText.color = Profile.ReleasedTextColor;
                countText.alignment = TextAnchor.LowerCenter;
                countText.text = Settings.Total + "";
                keyCountTexts[KeyCode.Joystick1Button0] = countText;

                keyPrevStates[KeyCode.Joystick1Button0] = false;
            }
            UpdateLayout();
        }
        public static void UpdateKPSTotal()
        {
            keyCountTexts[KeyCode.None].text = Calculate.kps + "";
            keyCountTexts[KeyCode.Joystick1Button0].text = Settings.Total + "";
        }
        public void GenerateKey(KeyCode code)
        {
            if (!keyCounts.ContainsKey(code))
            {
                keyCounts[code] = 0;
            }

            GameObject keyBgObj = new GameObject();
            keyBgObj.transform.SetParent(keysObject.transform);
            Image bgImage = keyBgObj.AddComponent<Image>();
            bgImage.sprite = Asset.KeyBackgroundSprite;
            bgImage.color = Profile.ReleasedBackgroundColor;
            keyBgImages[code] = bgImage;
            bgImage.type = Image.Type.Sliced;

            GameObject keyOutlineObj = new GameObject();
            keyOutlineObj.transform.SetParent(keysObject.transform);
            Image outlineImage = keyOutlineObj.AddComponent<Image>();
            outlineImage.sprite = Asset.KeyOutlineSprite;
            outlineImage.color = Profile.ReleasedOutlineColor;
            keyOutlineImages[code] = outlineImage;
            outlineImage.type = Image.Type.Sliced;

            GameObject keyTextObj = new GameObject();
            keyTextObj.transform.SetParent(keysObject.transform);
            Text text = keyTextObj.AddComponent<Text>();
            text.font = RDString.GetFontDataForLanguage(SystemLanguage.English).font;
            text.color = Profile.ReleasedTextColor;
            text.alignment = TextAnchor.UpperCenter;
            if (!KEY_TO_STRING.TryGetValue(code, out string codeString))
            {
                codeString = code.ToString();
            }
            text.text = codeString;
            keyTexts[code] = text;
            ContentSizeFitter csf = keyTextObj.AddComponent<ContentSizeFitter>();
            csf.horizontalFit = ContentSizeFitter.FitMode.Unconstrained;
            csf.verticalFit = ContentSizeFitter.FitMode.Unconstrained;

            GameObject keyCountTextObj = new GameObject();
            keyCountTextObj.transform.SetParent(keysObject.transform);
            Text countText = keyCountTextObj.AddComponent<Text>();
            countText.resizeTextForBestFit = true;
            countText.resizeTextMinSize = 0;
            countText.resizeTextMaxSize = 50;
            countText.lineSpacing = 0;
            countText.font = RDString.GetFontDataForLanguage(SystemLanguage.English).font;
            countText.color = Profile.ReleasedTextColor;
            countText.alignment = TextAnchor.LowerCenter;
            countText.text = keyCounts[code] + "";
            keyCountTexts[code] = countText;
            ContentSizeFitter csf2 = keyCountTextObj.AddComponent<ContentSizeFitter>();
            csf2.horizontalFit = ContentSizeFitter.FitMode.Unconstrained;
            csf2.verticalFit = ContentSizeFitter.FitMode.Unconstrained;

            keyPrevStates[code] = false;
        }
        internal struct Layout
        {
            public static implicit operator Layout((Vector2 IP, Vector2 IS, Vector2 TP, Vector2 TS, Vector2 CTP, Vector2 CTS) IPISTPTSCTPCTS) => new Layout(IPISTPTSCTPCTS.IP, IPISTPTSCTPCTS.IS, IPISTPTSCTPCTS.TP, IPISTPTSCTPCTS.TS, IPISTPTSCTPCTS.CTP, IPISTPTSCTPCTS.CTS);
            public Layout(Vector2 IP, Vector2 IS, Vector2 TP, Vector2 TS, Vector2 CTP, Vector2 CTS)
            {
                ImagePos = IP;
                ImageSize = IS;
                TextPos = TP;
                TextSize = TS;
                CountTextPos = CTP;
                CountTextSize = CTS;
            }
            public Vector2 ImagePos;
            public Vector2 ImageSize;
            public Vector2 TextPos;
            public Vector2 TextSize;
            public Vector2 CountTextPos;
            public Vector2 CountTextSize;
        }
        internal static KeyViewerSettings Settings => Main.Settings;
        /// <summary>
        /// Updates the position, size, and color of the displayed keys.
        /// </summary>
        public static void UpdateLayout() {
            int count = keyOutlineImages.Keys.Count;
            float keyWidth = 100;
            float keyHeight = Profile.ShowKeyPressTotal ? 150 : 100;
            float spacing = 10;
            float width = count * keyWidth + (count - 1) * spacing;
            Vector2 pos = new Vector2(Profile.KeyViewerXPos, Profile.KeyViewerYPos);


            keysRectTransform.anchorMin = pos;
            keysRectTransform.anchorMax = pos;
            keysRectTransform.pivot = pos;
            keysRectTransform.sizeDelta = new Vector2(width, keyHeight);
            keysRectTransform.anchoredPosition = Vector2.zero;
            keysRectTransform.localScale = new Vector3(1, 1, 1) * Profile.KeyViewerSize / 100f;
            float x = 0;
            foreach (KeyCode code in keyOutlineImages.Keys)
            {
                DOTween.Kill(TweenIdForKeyCode(code));
                Image bgImage = keyBgImages[code];
                bgImage.rectTransform.anchorMin = Vector2.zero;
                bgImage.rectTransform.anchorMax = Vector2.zero;
                bgImage.rectTransform.pivot = new Vector2(0.5f, 0.5f);
                bgImage.rectTransform.sizeDelta = new Vector2(keyWidth, keyHeight);
                bgImage.rectTransform.anchoredPosition =
                    new Vector2(x + keyWidth / 2, keyHeight / 2);

                Image outlineImage = keyOutlineImages[code];
                outlineImage.rectTransform.anchorMin = Vector2.zero;
                outlineImage.rectTransform.anchorMax = Vector2.zero;
                outlineImage.rectTransform.pivot = new Vector2(0.5f, 0.5f);
                outlineImage.rectTransform.sizeDelta = new Vector2(keyWidth, keyHeight);
                outlineImage.rectTransform.anchoredPosition =
                    new Vector2(x + keyWidth / 2, keyHeight / 2);

                float heightOffset = Profile.ShowKeyPressTotal ? 0 : keyHeight / 20f;
                Text text = keyTexts[code];
                text.rectTransform.anchorMin = Vector2.zero;
                text.rectTransform.anchorMax = Vector2.zero;
                text.rectTransform.pivot = new Vector2(0.5f, 0.5f);
                text.rectTransform.sizeDelta = new Vector2(keyWidth, keyHeight * 1.03f);
                text.rectTransform.anchoredPosition =
                    new Vector2(
                        x + keyWidth / 2,
                        keyHeight / 2 + heightOffset);
                text.fontSize = Mathf.RoundToInt(keyWidth * 3 / 4);
                text.alignment =
                    Profile.ShowKeyPressTotal ? TextAnchor.UpperCenter : TextAnchor.MiddleCenter;

                Text countText = keyCountTexts[code];
                countText.rectTransform.anchorMin = Vector2.zero;
                countText.rectTransform.anchorMax = Vector2.zero;
                countText.rectTransform.pivot = new Vector2(0.5f, 0.5f);
                countText.rectTransform.sizeDelta = new Vector2(keyWidth, keyHeight * 0.8f);
                countText.rectTransform.anchoredPosition =
                    new Vector2(
                        x + keyWidth / 2,
                        keyHeight / 2);
                countText.fontSize = Mathf.RoundToInt(keyWidth / 2);
                countText.gameObject.SetActive(Profile.ShowKeyPressTotal);

                // Press/release state
                Vector3 scale = new Vector3(1, 1, 1);
                if (keyPrevStates[code])
                {
                    bgImage.color = Profile.PressedBackgroundColor;
                    outlineImage.color = Profile.PressedOutlineColor;
                    text.color = Profile.PressedTextColor;
                    countText.color = Profile.PressedTextColor;
                    scale *= SHRINK_FACTOR;
                }
                else
                {
                    bgImage.color = Profile.ReleasedBackgroundColor;
                    outlineImage.color = Profile.ReleasedOutlineColor;
                    text.color = Profile.ReleasedTextColor;
                    countText.color = Profile.ReleasedTextColor;
                }
                bgImage.rectTransform.localScale = scale;
                outlineImage.rectTransform.localScale = scale;
                text.rectTransform.localScale = scale;
                countText.rectTransform.localScale = scale;

                x += keyWidth + spacing;
            }
            float KTH = Profile.ShowKeyPressTotal ? 75 : 62.5f;
            float KkeyWidth = (Settings.IsAloneKPS ? (x - keyWidth - 20) : (x - keyWidth - 20) / 2) - (Settings.DisplayTotal ? 55 : 0);//((count - (Settings.IsAloneKPS ? 1 : 2)) * 100) + (x / spacing) + ((Settings.ActiveKeys.Count - (Settings.IsAloneKPS ? 1 : 2)) * 10);
            float KkeyHeight = Profile.ShowKeyPressTotal ? 75 : 50;
            float TkeyWidth = (Settings.IsAloneTotal ? (x - keyWidth - 20) : (x - keyWidth - 20) / 2) - (Settings.DisplayKPS ? 55 : 0);//((count - (Settings.IsAloneKPS ? 1 : 2)) * 100) + (x / spacing) + ((Settings.ActiveKeys.Count - (Settings.IsAloneKPS ? 1 : 2)) * 10);
            if (Main.Settings.DisplayKPS)
            {
                DOTween.Kill(TweenIdForKeyCode(KeyCode.None));
                Image bgImage = keyBgImages[KeyCode.None];
                bgImage.rectTransform.anchorMin = Vector2.zero;
                bgImage.rectTransform.anchorMax = Vector2.zero;
                bgImage.rectTransform.pivot = new Vector2(0.5f, 0.5f);
                bgImage.rectTransform.sizeDelta = new Vector2(KkeyWidth, 75);
                bgImage.rectTransform.anchoredPosition =
                    new Vector2(KkeyWidth / 2, KkeyHeight / 2 - KTH);

                Image outlineImage = keyOutlineImages[KeyCode.None];
                outlineImage.rectTransform.anchorMin = Vector2.zero;
                outlineImage.rectTransform.anchorMax = Vector2.zero;
                outlineImage.rectTransform.pivot = new Vector2(0.5f, 0.5f);
                outlineImage.rectTransform.sizeDelta = new Vector2(KkeyWidth, 75);
                outlineImage.rectTransform.anchoredPosition =
                    new Vector2(KkeyWidth / 2, KkeyHeight / 2 - KTH);

                float heightOffset = Profile.ShowKeyPressTotal ? 0 : KTH / 20f;
                Text text = keyTexts[KeyCode.None];
                text.rectTransform.anchorMin = Vector2.zero;
                text.rectTransform.anchorMax = Vector2.zero;
                text.rectTransform.pivot = new Vector2(0.5f, 0.5f);
                text.rectTransform.sizeDelta = new Vector2(KkeyWidth, KTH * 1.03f);
                text.rectTransform.anchoredPosition =
                    new Vector2(
                        KkeyWidth / 2,
                        KTH / 2 + heightOffset - KTH + 12);
                text.fontSize = 50;//Mathf.RoundToInt(KkeyWidth * 3 / 4);
                text.alignment =
                    Profile.ShowKeyPressTotal ? TextAnchor.UpperCenter : TextAnchor.MiddleCenter;

                Text countText = keyCountTexts[KeyCode.None];
                countText.rectTransform.anchorMin = Vector2.zero;
                countText.rectTransform.anchorMax = Vector2.zero;
                countText.rectTransform.pivot = new Vector2(0.5f, 0.5f);
                countText.rectTransform.sizeDelta = new Vector2(KkeyWidth, KTH * 0.8f);
                countText.rectTransform.anchoredPosition =
                    new Vector2(
                        KkeyWidth / 2,
                        Profile.ShowKeyPressTotal ? KTH / 2 + heightOffset - KTH - 2 : KTH / 2 + heightOffset - KTH - 15.5f);
                countText.fontSize = 40;//Mathf.RoundToInt(KkeyWidth / 2);
                countText.gameObject.SetActive(Settings.DisplayKPS);

                // Press/release state
                Vector3 scale = new Vector3(1, 1, 1);
                if (keyPrevStates[KeyCode.None])
                {
                    bgImage.color = Profile.PressedBackgroundColor;
                    outlineImage.color = Profile.PressedOutlineColor;
                    text.color = Profile.PressedTextColor;
                    countText.color = Profile.PressedTextColor;
                    scale *= SHRINK_FACTOR;
                }
                else
                {
                    bgImage.color = Profile.ReleasedBackgroundColor;
                    outlineImage.color = Profile.ReleasedOutlineColor;
                    text.color = Profile.ReleasedTextColor;
                    countText.color = Profile.ReleasedTextColor;
                }
                bgImage.rectTransform.localScale = scale;
                outlineImage.rectTransform.localScale = scale;
                text.rectTransform.localScale = scale;
                countText.rectTransform.localScale = scale;
            }
            if (Main.Settings.DisplayTotal)
            {
                DOTween.Kill(TweenIdForKeyCode(KeyCode.Joystick1Button0));
                Image bgImage = keyBgImages[KeyCode.Joystick1Button0];
                bgImage.rectTransform.anchorMin = Vector2.zero;
                bgImage.rectTransform.anchorMax = Vector2.zero;
                bgImage.rectTransform.pivot = new Vector2(0.5f, 0.5f);
                bgImage.rectTransform.sizeDelta = new Vector2(TkeyWidth, 75);
                bgImage.rectTransform.anchoredPosition =
                    new Vector2(TkeyWidth * 1.5f, KkeyHeight / 2 - KTH);

                Image outlineImage = keyOutlineImages[KeyCode.Joystick1Button0];
                outlineImage.rectTransform.anchorMin = Vector2.zero;
                outlineImage.rectTransform.anchorMax = Vector2.zero;
                outlineImage.rectTransform.pivot = new Vector2(0.5f, 0.5f);
                outlineImage.rectTransform.sizeDelta = new Vector2(TkeyWidth, 75);
                outlineImage.rectTransform.anchoredPosition =
                    new Vector2(TkeyWidth * 1.5f, KkeyHeight / 2 - KTH);

                float heightOffset = Profile.ShowKeyPressTotal ? 0 : KTH / 20f;
                Text text = keyTexts[KeyCode.Joystick1Button0];
                text.rectTransform.anchorMin = Vector2.zero;
                text.rectTransform.anchorMax = Vector2.zero;
                text.rectTransform.pivot = new Vector2(0.5f, 0.5f);
                text.rectTransform.sizeDelta = new Vector2(TkeyWidth, KTH * 1.03f);
                text.rectTransform.anchoredPosition =
                    new Vector2(
                        TkeyWidth * 1.5f,
                        KTH / 2 + heightOffset - KTH + 12);
                text.fontSize = 50;//Mathf.RoundToInt(keyWidth * 3 / 4);
                text.alignment =
                    Profile.ShowKeyPressTotal ? TextAnchor.UpperCenter : TextAnchor.MiddleCenter;

                Text countText = keyCountTexts[KeyCode.Joystick1Button0];
                countText.rectTransform.anchorMin = Vector2.zero;
                countText.rectTransform.anchorMax = Vector2.zero;
                countText.rectTransform.pivot = new Vector2(0.5f, 0.5f);
                countText.rectTransform.sizeDelta = new Vector2(TkeyWidth, KTH * 0.8f);
                countText.rectTransform.anchoredPosition =
                    new Vector2(
                        TkeyWidth * 1.5f,
                        Profile.ShowKeyPressTotal ? KTH / 2 + heightOffset - KTH - 2 : KTH / 2 + heightOffset - KTH - 15.5f);
                countText.fontSize = 40;//Mathf.RoundToInt(keyWidth / 2);
                countText.gameObject.SetActive(Settings.DisplayTotal);

                // Press/release state
                Vector3 scale = new Vector3(1, 1, 1);
                if (keyPrevStates[KeyCode.Joystick1Button0])
                {
                    bgImage.color = Profile.PressedBackgroundColor;
                    outlineImage.color = Profile.PressedOutlineColor;
                    text.color = Profile.PressedTextColor;
                    countText.color = Profile.PressedTextColor;
                    scale *= SHRINK_FACTOR;
                }
                else
                {
                    bgImage.color = Profile.ReleasedBackgroundColor;
                    outlineImage.color = Profile.ReleasedOutlineColor;
                    text.color = Profile.ReleasedTextColor;
                    countText.color = Profile.ReleasedTextColor;
                }
                bgImage.rectTransform.localScale = scale;
                outlineImage.rectTransform.localScale = scale;
                text.rectTransform.localScale = scale;
                countText.rectTransform.localScale = scale;
            }
        }
        /// <summary>
        /// Updates the current state of the keys.
        /// </summary>
        /// <param name="state">The current state of the keys.</param>
        public void UpdateState(Dictionary<KeyCode, bool> state) {
            foreach (KeyCode code in keyOutlineImages.Keys) {
                // Only change if the state changed
                if (state[code] == keyPrevStates[code] || code == KeyCode.None || code == KeyCode.Joystick1Button0) {
                    continue;
                }
                keyPrevStates[code] = state[code];

                string id = TweenIdForKeyCode(code);
                if (DOTween.IsTweening(id)) {
                    DOTween.Kill(id);
                }

                Image bgImage = keyBgImages[code];
                Image outlineImage = keyOutlineImages[code];
                Text text = keyTexts[code];
                Text countText = keyCountTexts[code];

                // Handle key count increment
                if (state[code]) {
                    keyCounts[code]++;
                    countText.text = keyCounts[code] + "";
                }

                // Calculate the new color/size
                Color bgColor, outlineColor, textColor;
                Vector3 scale = new Vector3(1, 1, 1);
                if (state[code]) {
                    bgColor = Profile.PressedBackgroundColor;
                    outlineColor = Profile.PressedOutlineColor;
                    textColor = Profile.PressedTextColor;
                    if (Profile.AnimateKeys) {
                        scale *= SHRINK_FACTOR;
                    }
                } else {
                    bgColor = Profile.ReleasedBackgroundColor;
                    outlineColor = Profile.ReleasedOutlineColor;
                    textColor = Profile.ReleasedTextColor;
                }

                // Apply the new color/size
                bgImage.color = bgColor;
                outlineImage.color = outlineColor;
                text.color = textColor;
                countText.color = textColor;
                if (Profile.AnimateKeys) {
                    bgImage.rectTransform.DOScale(scale, EASE_DURATION)
                        .SetId(id)
                        .SetEase(Ease.OutExpo)
                        .SetUpdate(true)
                        .OnKill(() => bgImage.rectTransform.localScale = scale);
                    outlineImage.rectTransform.DOScale(scale, EASE_DURATION)
                        .SetId(id)
                        .SetEase(Ease.OutExpo)
                        .SetUpdate(true)
                        .OnKill(() => outlineImage.rectTransform.localScale = scale);
                    text.rectTransform.DOScale(scale, EASE_DURATION)
                        .SetId(id)
                        .SetEase(Ease.OutExpo)
                        .SetUpdate(true)
                        .OnKill(() => text.rectTransform.localScale = scale);
                    countText.rectTransform.DOScale(scale, EASE_DURATION)
                        .SetId(id)
                        .SetEase(Ease.OutExpo)
                        .SetUpdate(true)
                        .OnKill(() => countText.rectTransform.localScale = scale);
                } else {
                    bgImage.rectTransform.localScale = scale;
                    outlineImage.rectTransform.localScale = scale;
                    text.rectTransform.localScale = scale;
                    countText.rectTransform.localScale = scale;
                }
            }
        }

        private static string TweenIdForKeyCode(KeyCode code) {
            return $"KeyViewer.{code}";
        }
    }
}

using System.Collections.Generic;
using DG.Tweening;
using System;
using System.Linq;
using UnityEngine.UI;
using UnityEngine;
using UnityEngine.EventSystems;

namespace KeyViewer
{
    /// <summary>
    /// A key viewer that shows if a list of given keys are currently being
    /// pressed.
    /// </summary>
    public class KeyViewer : MonoBehaviour
    {
        public static readonly Dictionary<KeyCode, string> KEY_TO_STRING =
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

        internal const float EASE_DURATION = 0.1f;
        internal const float SHRINK_FACTOR = 0.9f;

        internal static Dictionary<KeyCode, int> keyCounts = new Dictionary<KeyCode, int>();

        internal static GameObject keysObject;
        internal static Dictionary<KeyCode, Image> keyBgImages;
        internal static Dictionary<KeyCode, Image> keyOutlineImages;
        internal static Dictionary<KeyCode, Text> keyTexts;
        internal static Dictionary<KeyCode, Text> keyCountTexts;
        internal static Dictionary<KeyCode, bool> keyPrevStates;
        internal static RectTransform keysRectTransform;

        internal static KeyViewerProfile _profile = new KeyViewerProfile();

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
            scaler.referenceResolution = new Vector2(Screen.width, Screen.height);

            UpdateKeys();
        }

        /// <summary>
        /// Updates what keys are displayed on the key viewer.
        /// </summary>
        public static void UpdateKeys() {
            if (keysObject) {
                Destroy(keysObject);
                foreach (KeyValuePair<KeyCode, Key> kvp in Key.Keys) kvp.Value.Destroy();
                Key.Keys.Clear();
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

            foreach (KeyCode code in Profile.ActiveKeys)
            {
                Key k = new(code, new KeySetting());
                k.Create();
            }
            //if (Main.Settings.DisplayKPS)
            //{
            //    GameObject keyBgObj = new GameObject();
            //    keyBgObj.transform.SetParent(keysObject.transform);
            //    Image bgImage = keyBgObj.AddComponent<Image>();
            //    bgImage.sprite = Asset.KeyBackgroundSprite;
            //    bgImage.color = Profile.ReleasedBackgroundColor;
            //    keyBgImages[KeyCode.None] = bgImage;
            //    bgImage.type = Image.Type.Sliced;

            //    GameObject keyOutlineObj = new GameObject();
            //    keyOutlineObj.transform.SetParent(keysObject.transform);
            //    Image outlineImage = keyOutlineObj.AddComponent<Image>();
            //    outlineImage.sprite = Asset.KeyOutlineSprite;
            //    outlineImage.color = Profile.ReleasedOutlineColor;
            //    keyOutlineImages[KeyCode.None] = outlineImage;
            //    outlineImage.type = Image.Type.Sliced;

            //    GameObject keyTextObj = new GameObject();
            //    keyTextObj.transform.SetParent(keysObject.transform);
            //    Text text = keyTextObj.AddComponent<Text>();
            //    text.font = RDString.GetFontDataForLanguage(SystemLanguage.English).font;
            //    text.color = Profile.ReleasedTextColor;
            //    text.alignment = TextAnchor.UpperCenter;
            //    text.text = "KPS";
            //    keyTexts[KeyCode.None] = text;

            //    GameObject keyCountTextObj = new GameObject();
            //    keyCountTextObj.transform.SetParent(keysObject.transform);
            //    Text countText = keyCountTextObj.AddComponent<Text>();
            //    countText.resizeTextForBestFit = true;
            //    countText.resizeTextMinSize = 0;
            //    countText.resizeTextMaxSize = 40;
            //    countText.font = RDString.GetFontDataForLanguage(SystemLanguage.English).font;
            //    countText.color = Profile.ReleasedTextColor;
            //    countText.alignment = TextAnchor.LowerCenter;
            //    countText.text = Calculate.kps + "";
            //    keyCountTexts[KeyCode.None] = countText;

            //    keyPrevStates[KeyCode.None] = false;
            //}

            //if (Main.Settings.DisplayTotal)
            //{
            //    GameObject keyBgObj = new GameObject();
            //    keyBgObj.transform.SetParent(keysObject.transform);
            //    Image bgImage = keyBgObj.AddComponent<Image>();
            //    bgImage.sprite = Asset.KeyBackgroundSprite;
            //    bgImage.color = Profile.ReleasedBackgroundColor;
            //    keyBgImages[KeyCode.Joystick1Button0] = bgImage;
            //    bgImage.type = Image.Type.Sliced;

            //    GameObject keyOutlineObj = new GameObject();
            //    keyOutlineObj.transform.SetParent(keysObject.transform);
            //    Image outlineImage = keyOutlineObj.AddComponent<Image>();
            //    outlineImage.sprite = Asset.KeyOutlineSprite;
            //    outlineImage.color = Profile.ReleasedOutlineColor;
            //    keyOutlineImages[KeyCode.Joystick1Button0] = outlineImage;
            //    outlineImage.type = Image.Type.Sliced;

            //    GameObject keyTextObj = new GameObject();
            //    keyTextObj.transform.SetParent(keysObject.transform);
            //    Text text = keyTextObj.AddComponent<Text>();
            //    text.font = RDString.GetFontDataForLanguage(SystemLanguage.English).font;
            //    text.color = Profile.ReleasedTextColor;
            //    text.alignment = TextAnchor.UpperCenter;
            //    text.text = "Total";
            //    keyTexts[KeyCode.Joystick1Button0] = text;

            //    GameObject keyCountTextObj = new GameObject();
            //    keyCountTextObj.transform.SetParent(keysObject.transform);
            //    Text countText = keyCountTextObj.AddComponent<Text>();
            //    countText.resizeTextForBestFit = true;
            //    countText.resizeTextMinSize = 0;
            //    countText.resizeTextMaxSize = 40;
            //    countText.font = RDString.GetFontDataForLanguage(SystemLanguage.English).font;
            //    countText.color = Profile.ReleasedTextColor;
            //    countText.alignment = TextAnchor.LowerCenter;
            //    countText.text = Settings.Total + "";
            //    keyCountTexts[KeyCode.Joystick1Button0] = countText;

            //    keyPrevStates[KeyCode.Joystick1Button0] = false;
            //}
            UpdateLayout();
        }
        public static void UpdateKPSTotal()
        {
            if (keyCountTexts.TryGetValue(KeyCode.None, out var k)) k.text = Calculate.kps + "";
            if (keyCountTexts.TryGetValue(KeyCode.Joystick1Button0, out var t)) t.text = Settings.Total + "";
        }
        //public void GenerateKey(KeyCode code)
        //{
        //    if (!keyCounts.ContainsKey(code))
        //    {
        //        keyCounts[code] = 0;
        //    }

        //    GameObject keyBgObj = new GameObject(code.ToString());
        //    keyBgObj.transform.SetParent(keysObject.transform);
        //    Image bgImage = keyBgObj.AddComponent<Image>();
        //    bgImage.sprite = Asset.KeyBackgroundSprite;
        //    bgImage.color = Profile.ReleasedBackgroundColor;
        //    keyBgImages[code] = bgImage;
        //    bgImage.type = Image.Type.Sliced;
        //    keyBgObj.AddComponent<DragHandler>();

        //    GameObject keyOutlineObj = new GameObject(code.ToString());
        //    keyOutlineObj.transform.SetParent(keysObject.transform);
        //    Image outlineImage = keyOutlineObj.AddComponent<Image>();
        //    outlineImage.sprite = Asset.KeyOutlineSprite;
        //    outlineImage.color = Profile.ReleasedOutlineColor;
        //    keyOutlineImages[code] = outlineImage;
        //    outlineImage.type = Image.Type.Sliced;
        //    keyOutlineObj.AddComponent<DragHandler>();

        //    GameObject keyTextObj = new GameObject(code.ToString());
        //    keyTextObj.transform.SetParent(keysObject.transform);
        //    Text text = keyTextObj.AddComponent<Text>();
        //    text.font = RDString.GetFontDataForLanguage(SystemLanguage.English).font;
        //    text.color = Profile.ReleasedTextColor;
        //    text.alignment = TextAnchor.UpperCenter;
        //    if (!KEY_TO_STRING.TryGetValue(code, out string codeString))
        //    {
        //        codeString = code.ToString();
        //    }
        //    text.text = codeString;
        //    keyTexts[code] = text;
        //    ContentSizeFitter csf = keyTextObj.AddComponent<ContentSizeFitter>();
        //    csf.horizontalFit = ContentSizeFitter.FitMode.Unconstrained;
        //    csf.verticalFit = ContentSizeFitter.FitMode.Unconstrained;
        //    keyTextObj.AddComponent<DragHandler>();

        //    GameObject keyCountTextObj = new GameObject(code.ToString());
        //    keyCountTextObj.transform.SetParent(keysObject.transform);
        //    Text countText = keyCountTextObj.AddComponent<Text>();
        //    countText.resizeTextForBestFit = true;
        //    countText.resizeTextMinSize = 0;
        //    countText.resizeTextMaxSize = 50;
        //    countText.lineSpacing = 0;
        //    countText.font = RDString.GetFontDataForLanguage(SystemLanguage.English).font;
        //    countText.color = Profile.ReleasedTextColor;
        //    countText.alignment = TextAnchor.LowerCenter;
        //    countText.text = keyCounts[code] + "";
        //    keyCountTexts[code] = countText;
        //    ContentSizeFitter csf2 = keyCountTextObj.AddComponent<ContentSizeFitter>();
        //    csf2.horizontalFit = ContentSizeFitter.FitMode.Unconstrained;
        //    csf2.verticalFit = ContentSizeFitter.FitMode.Unconstrained;
        //    keyCountTextObj.AddComponent<DragHandler>();

        //    keyPrevStates[code] = false;
        //}
        internal static KeyViewerSettings Settings => Main.Settings;
        internal static bool setpos = false;
        internal static Dictionary<KeyCode, PoSize> keyposize = new Dictionary<KeyCode, PoSize>();
        public static float Length = 0;
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
                if (Key.Keys.TryGetValue(code, out Key key))
                {
                    //if (code == KeyCode.None)
                    //{
                    //    key.countText.fontSize = 50;
                    //    key.text.fontSize = 75;
                    //}
                    //if (code == KeyCode.Joystick1Button0)
                    //{
                    //    key.countText.fontSize = 50;
                    //    key.text.fontSize = 55;
                    //}
                    PoSize ps = (key.ks.ps.Pos, key.ks.ps.Size);
                    keyposize[code] = ps;
                    key.UpdateLayout(ref x);
                    continue;
                }
                else
                {
                    Key k = new Key(code, new KeySetting());
                    //if (code == KeyCode.None)
                    //{
                    //    k.countText.fontSize = 50;
                    //    k.text.fontSize = 75;
                    //}
                    //if (code == KeyCode.Joystick1Button0)
                    //{
                    //    k.countText.fontSize = 50;
                    //    k.text.fontSize = 55;
                    //}
                    PoSize ps = (k.ks.ps.Pos, k.ks.ps.Size);
                    keyposize[code] = ps;
                    k.UpdateLayout(ref x);
                }
            }

            Length = x - 10;
            return;
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

        internal static string TweenIdForKeyCode(KeyCode code) {
            return $"KeyViewer.{code}";
        }
    }
}

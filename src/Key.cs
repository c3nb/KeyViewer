using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DG.Tweening;
using SA.GoogleDoc;
using UnityEngine;
using UnityEngine.UI;
using UnityEngine.EventSystems;
using UnityModManagerNet;

namespace KeyViewer
{
    [Serializable]
    public class KeySetting
    {
        public KeySetting()
        {
            TextResize = false;
            CTextResize = true;
            TextF = 75;
            CTextF = 50;
            TextF = 75;
            CTextf = 0;
            ps = PoSize.Zero;
            TextPos = Point.Zero;
            CTextPos = Point.Zero;
        }

        public void SetBar()
        {
            _TextResize = TextResize;
            _CTextResize = CTextResize;
            _TextF = TextF;
            _Textf = Textf;
            _CTextF = CTextF;
            _CTextf = CTextf;
            _ps = ps;
            _TextPos = TextPos;
            _CTextPos = CTextPos;

            TextResize = false;
            CTextResize = true;
            TextF = 50;
            Textf = 0;
            CTextF = 40;
            CTextf = 0;
            ps.Size.x = ((Key.Keys.Count - Key.KT - Key.Space) / 2 - 1) * 110;
            ps.Size.y = -80;
            TextPos.y = 11.5f;
            CTextPos.y = -9f;
            KeyViewer.UpdateLayout();
        }

        public void Undo()
        {
            TextResize = _TextResize;
            CTextResize = _CTextResize;
            TextF = _TextF;
            Textf = _Textf;
            CTextF = _CTextF;
            CTextf = _CTextf;
            ps = _ps;
            TextPos = _TextPos;
            CTextPos = _CTextPos;
            KeyViewer.UpdateLayout();
        }
        public bool TextResize;
        public bool CTextResize;
        public int TextF;
        public int Textf;
        public int CTextF;
        public int CTextf;
        public PoSize ps;
        public Point TextPos;
        public Point CTextPos;

        private bool _TextResize;
        private bool _CTextResize;
        private int _TextF;
        private int _Textf;
        private int _CTextF;
        private int _CTextf;
        private PoSize _ps;
        private Point _TextPos;
        private Point _CTextPos;
    }

    public class Key
    {
        public static Dictionary<KeyCode, Key> Keys = new();
        public static Dictionary<KeyCode, KeySetting> KeySettings = new();
        public static Dictionary<KeyCode, Key> Groups = new();
        public static float KTLength = 0;
        public static int KT = 0;
        public KeyCode code;
        public GameObject keyBgObj;
        public GameObject keyOutlineObj;
        public GameObject keyCountTextObj;
        public GameObject keyTextObj;
        public Image bgImage;
        public Image outlineImage;
        public Text text;
        public Text countText;
        public string p;
        public string s;
        public KeySetting ks;
        public KeySetting cache;
        public bool group = false;
        public static int Space = 0;
#pragma warning disable IDE0051
        public Key(KeyCode Code, KeySetting ks)
        {
            code = Code;
            if (!Keys.ContainsKey(code)) Keys.Add(code, this);
            if (!KeySettings.ContainsKey(code)) KeySettings.Add(code, ks);
            if (!KeyViewer.keyCounts.ContainsKey(code)) KeyViewer.keyCounts[code] = 0;
            this.ks = KeySettings[code];
            Main.Logger.Log($"{code} Creating...");
            p = $"{(code == KeyCode.None ? "KPS" : (code == KeyCode.Joystick1Button0 ? "Total" : code))} Position";
            s = $"{(code == KeyCode.None ? "KPS" : (code == KeyCode.Joystick1Button0 ? "Total" : code))} Size";
            if (code == KeyCode.None || code == KeyCode.Joystick1Button0)
            {
                KTLength += 110;
                KT++;
            }
            if (code == KeyCode.Space)
            {
                Space++;
            }
        }

        void SetResize(Text text = null, Text countText = null)
        {
            if (text)
            {
                text.resizeTextForBestFit = ks.TextResize;
                text.resizeTextMinSize = ks.Textf;
                text.resizeTextMaxSize = ks.TextF;
                if (!text.resizeTextForBestFit) text.fontSize = ks.TextF;
                return;
            }
            else if (countText)
            {
                countText.resizeTextForBestFit = ks.CTextResize;
                countText.resizeTextMinSize = ks.CTextf;
                countText.resizeTextMaxSize = ks.CTextF;
                if (!countText.resizeTextForBestFit) countText.fontSize = ks.CTextF;
                return;
            }
            else throw new NullReferenceException("Text and CountText are Null");
        }
        public void Create()
        {
            keyBgObj = new GameObject();
            keyBgObj.transform.SetParent(KeyViewer.keysObject.transform);
            bgImage = keyBgObj.AddComponent<Image>();
            bgImage.sprite = Asset.KeyBackgroundSprite;
            bgImage.color = KeyViewer.Profile.ReleasedBackgroundColor;
            KeyViewer.keyBgImages[code] = bgImage;
            bgImage.type = Image.Type.Sliced;

            keyOutlineObj = new GameObject();
            keyOutlineObj.transform.SetParent(KeyViewer.keysObject.transform);
            outlineImage = keyOutlineObj.AddComponent<Image>();
            outlineImage.sprite = Asset.KeyOutlineSprite;
            outlineImage.color = KeyViewer.Profile.ReleasedOutlineColor;
            KeyViewer.keyOutlineImages[code] = outlineImage;
            outlineImage.type = Image.Type.Sliced;

            keyTextObj = new GameObject();
            keyTextObj.transform.SetParent(KeyViewer.keysObject.transform);
            text = keyTextObj.AddComponent<Text>();

            SetResize(text: text);



            text.font = RDString.GetFontDataForLanguage(SystemLanguage.English).font;
            text.color = KeyViewer.Profile.ReleasedTextColor;
            text.alignment = TextAnchor.UpperCenter;
            if (!KeyViewer.KEY_TO_STRING.TryGetValue(code, out string codeString))
            {
                codeString = code.ToString();
            }
            if (code == KeyCode.None) text.text = "KPS";
            else if (code == KeyCode.Joystick1Button0)
            {
                text.text = "Total";
                text.fontSize = 65;
            }
            else text.text = codeString;
            KeyViewer.keyTexts[code] = text;

            keyCountTextObj = new GameObject();
            keyCountTextObj.transform.SetParent(KeyViewer.keysObject.transform);
            countText = keyCountTextObj.AddComponent<Text>();
            if (code == KeyCode.None) countText.text = Calculate.kps + "";
            else if (code == KeyCode.Joystick1Button0) countText.text = KeyViewer.Settings.Total + "";
            else countText.text = KeyViewer.keyCounts[code] + "";

            SetResize(countText: countText);
            countText.font = RDString.GetFontDataForLanguage(SystemLanguage.English).font;
            countText.color = KeyViewer.Profile.ReleasedTextColor;
            countText.alignment = TextAnchor.LowerCenter;

            KeyViewer.keyCountTexts[code] = countText;

            KeyViewer.keyPrevStates[code] = false;

            return;
        }

        public void Destroy()
        {
            UnityEngine.Object.Destroy(keyBgObj);
            UnityEngine.Object.Destroy(keyOutlineObj);
            UnityEngine.Object.Destroy(keyCountTextObj);
            UnityEngine.Object.Destroy(keyTextObj);
        }
        public void UpdateLayout(ref float x)
        {
            float keyWidth = 100;
            float keyHeight = KeyViewer.Profile.ShowKeyPressTotal ? 150 : 100;
            float spacing = 10;
            DOTween.Kill(KeyViewer.TweenIdForKeyCode(code));

            bgImage = KeyViewer.keyBgImages[code];
            bgImage.rectTransform.anchorMin = Vector2.zero;
            bgImage.rectTransform.anchorMax = Vector2.zero;
            bgImage.rectTransform.pivot = new Vector2(0.5f, 0.5f);
            bgImage.rectTransform.sizeDelta = new Vector2(keyWidth + ks.ps.Size.x, keyHeight + ks.ps.Size.y);
            bgImage.rectTransform.anchoredPosition =
                new Vector2(x + keyWidth / 2 + ks.ps.Pos.x, keyHeight / 2 + ks.ps.Pos.y);

            outlineImage = KeyViewer.keyOutlineImages[code];
            outlineImage.rectTransform.anchorMin = Vector2.zero;
            outlineImage.rectTransform.anchorMax = Vector2.zero;
            outlineImage.rectTransform.pivot = new Vector2(0.5f, 0.5f);
            outlineImage.rectTransform.sizeDelta = new Vector2(keyWidth + ks.ps.Size.x, keyHeight + ks.ps.Size.y);
            outlineImage.rectTransform.anchoredPosition =
                new Vector2(x + keyWidth / 2 + ks.ps.Pos.x, keyHeight / 2 + ks.ps.Pos.y);

            float heightOffset = KeyViewer.Profile.ShowKeyPressTotal ? 0 : keyHeight / 20f;
            text = KeyViewer.keyTexts[code];
            text.rectTransform.anchorMin = Vector2.zero;
            text.rectTransform.anchorMax = Vector2.zero;
            text.rectTransform.pivot = new Vector2(0.5f, 0.5f);
            text.rectTransform.sizeDelta = new Vector2(keyWidth + ks.ps.Size.x, keyHeight * 1.03f + ks.ps.Size.y);
            text.rectTransform.anchoredPosition =
                new Vector2(
                    x + keyWidth / 2 + ks.ps.Pos.x + ks.TextPos.x,
                    keyHeight / 2 + heightOffset + ks.ps.Pos.y + ks.TextPos.y);
            if (code == KeyCode.Joystick1Button0) text.fontSize = 65;
            else text.fontSize = Mathf.RoundToInt(keyWidth * 3 / 4);
            text.alignment =
                KeyViewer.Profile.ShowKeyPressTotal ? TextAnchor.UpperCenter : TextAnchor.MiddleCenter;
            SetResize(text);

            countText = KeyViewer.keyCountTexts[code];
            countText.rectTransform.anchorMin = Vector2.zero;
            countText.rectTransform.anchorMax = Vector2.zero;
            countText.rectTransform.pivot = new Vector2(0.5f, 0.5f);
            countText.rectTransform.sizeDelta = new Vector2(keyWidth + ks.ps.Size.x, keyHeight * 0.8f + ks.ps.Size.y);
            countText.rectTransform.anchoredPosition =
                new Vector2(
                    x + keyWidth / 2 + ks.ps.Pos.x + ks.CTextPos.x,
                    keyHeight / 2 + ks.ps.Pos.y + ks.CTextPos.y);
            SetResize(countText: countText);
            //if (!countText.resizeTextForBestFit)
            //{
            //    countText.fontSize = Mathf.RoundToInt(keyWidth / 2);
            //}
            countText.gameObject.SetActive(KeyViewer.Profile.ShowKeyPressTotal);

            // Press/release state
            Vector3 scale = new Vector3(1, 1, 1);
            if (KeyViewer.keyPrevStates[code])
            {
                bgImage.color = KeyViewer.Profile.PressedBackgroundColor;
                outlineImage.color = KeyViewer.Profile.PressedOutlineColor;
                text.color = KeyViewer.Profile.PressedTextColor;
                countText.color = KeyViewer.Profile.PressedTextColor;
                scale *= KeyViewer.Settings.sf;
            }
            else
            {
                bgImage.color = KeyViewer.Profile.ReleasedBackgroundColor;
                outlineImage.color = KeyViewer.Profile.ReleasedOutlineColor;
                text.color = KeyViewer.Profile.ReleasedTextColor;
                countText.color = KeyViewer.Profile.ReleasedTextColor;
            }
            bgImage.rectTransform.localScale = scale;
            outlineImage.rectTransform.localScale = scale;
            text.rectTransform.localScale = scale;
            countText.rectTransform.localScale = scale;
            x += keyWidth + spacing;
        }
        public void UI()
        {
            GUILayout.BeginHorizontal();
            if (Groups.ContainsKey(code))
            {
                if (GUILayout.Button("Ungroup"))
                {
                    Groups.Remove(code);
                    KeyViewerTweak.groupkeys.Remove($"{(code == KeyCode.None ? "KPS" : code == KeyCode.Joystick1Button0 ? "Total" : code)}");
                }
            }
            else
            {
                if (GUILayout.Button("Group"))
                {
                    Groups.Add(code, this);
                    KeyViewerTweak.groupkeys.Add($"{(code == KeyCode.None ? "KPS" : code == KeyCode.Joystick1Button0 ? "Total" : code)}");
                }
            }
            GUILayout.FlexibleSpace();
            GUILayout.EndHorizontal();
            KeyViewerTweak.HoFlGUI("", () =>
            {
                if (GUILayout.Button("Set Bar")) ks.SetBar();
                if (GUILayout.Button("Undo")) ks.Undo();
            });
            if (DrawFontSetting()) KeyViewer.UpdateLayout();
            if (DrawPos(p)) KeyViewer.UpdateLayout();
            if (DrawSize(s)) KeyViewer.UpdateLayout();
            GUILayout.Space(20);
        }

        public bool DrawPos(string label, Action<Point> CallBack = null)
        {
            Point cache = ks.ps.Pos;
            GUILayout.BeginHorizontal();
            GUILayout.Label(label);
            GUILayout.EndHorizontal();
            GUILayout.BeginHorizontal();
            GUILayout.Label("(");
            float.TryParse(GUILayout.TextField(ks.ps.Pos.x.ToString("F2")), out ks.ps.Pos.x);
            GUILayout.Label(", ");
            float.TryParse(GUILayout.TextField(ks.ps.Pos.y.ToString("F2")), out ks.ps.Pos.y);
            GUILayout.Label(")");
            GUILayout.FlexibleSpace();
            GUILayout.EndHorizontal();
            GUILayout.BeginHorizontal();
            ks.ps.Pos.x = GUILayout.HorizontalSlider(ks.ps.Pos.x, -Screen.width, Screen.width);
            GUILayout.EndHorizontal();

            GUILayout.BeginHorizontal();
            ks.ps.Pos.y = GUILayout.HorizontalSlider(ks.ps.Pos.y, -Screen.height, Screen.height);
            GUILayout.EndHorizontal();
            CallBack?.Invoke(cache);
            return ks.ps.Pos != cache;
        }
        public bool DrawSize(string label, Action<Point> CallBack = null)
        {
            Point cache = ks.ps.Size;
            GUILayout.BeginHorizontal();
            GUILayout.Label(label);
            GUILayout.EndHorizontal();
            GUILayout.BeginHorizontal();
            GUILayout.Label("(");
            float.TryParse(GUILayout.TextField(ks.ps.Size.x.ToString("F2")), out ks.ps.Size.x);
            GUILayout.Label(", ");
            float.TryParse(GUILayout.TextField(ks.ps.Size.y.ToString("F2")), out ks.ps.Size.y);
            GUILayout.Label(")");
            GUILayout.FlexibleSpace();
            GUILayout.EndHorizontal();
            GUILayout.BeginHorizontal();
            ks.ps.Size.x = GUILayout.HorizontalSlider(ks.ps.Size.x, -100, Screen.width);
            GUILayout.EndHorizontal();

            GUILayout.BeginHorizontal();
            ks.ps.Size.y = GUILayout.HorizontalSlider(ks.ps.Size.y, -100, Screen.height);
            GUILayout.EndHorizontal();
            CallBack?.Invoke(cache);
            return ks.ps.Size != cache;
        }

        public bool cts;
        public bool ts;
        public bool DrawFontSetting() => KeyCodeText() || CountText();

        public bool KeyCodeText()
        {
            if (ts = GUILayout.Toggle(ts, "KeyCode Text"))
            {
                if (KeyViewerTweak.DrawSlider(ref ks.TextPos.x, "X")) return true;
                if (KeyViewerTweak.DrawSlider(ref ks.TextPos.y, "Y")) return true;
                if (!(ks.TextResize = GUILayout.Toggle(ks.TextResize, "Resize Font")))
                {
                    bool changed = false;
                    int cache = ks.TextF;
                    GUILayout.BeginHorizontal();
                    GUILayout.Space(10);
                    GUILayout.Label("Font Size");
                    int.TryParse(GUILayout.TextField(ks.TextF.ToString(), GUI.skin.textField), out cache);
                    GUILayout.FlexibleSpace();
                    GUILayout.EndHorizontal();
                    changed = cache != ks.TextF;
                    ks.TextF = cache;
                    return changed;
                }
                else
                {
                    bool changed = false;
                    int cache = ks.TextF;
                    int cache2 = ks.Textf;
                    GUILayout.BeginHorizontal();
                    GUILayout.Space(10);
                    GUILayout.Label("Min Font Size");
                    int.TryParse(GUILayout.TextField(ks.Textf.ToString(), GUI.skin.textField), out cache2);
                    GUILayout.Label("Max Font Size");
                    int.TryParse(GUILayout.TextField(ks.TextF.ToString(), GUI.skin.textField), out cache);
                    GUILayout.FlexibleSpace();
                    GUILayout.EndHorizontal();
                    changed = cache != ks.TextF || cache2 != ks.Textf;
                    ks.TextF = cache;
                    return changed;
                }
            }

            return false;
        }

        public bool CountText()
        {
            if (cts = GUILayout.Toggle(cts, "KeyCount Text"))
            {
                if (KeyViewerTweak.DrawSlider(ref ks.CTextPos.x, "X")) return true;
                if (KeyViewerTweak.DrawSlider(ref ks.CTextPos.y, "Y")) return true;
                if (!(ks.CTextResize = GUILayout.Toggle(ks.CTextResize, "Resize Font")))
                {
                    bool changed = false;
                    int cache = ks.CTextF;
                    GUILayout.BeginHorizontal();
                    GUILayout.Space(10);
                    GUILayout.Label("Font Size");
                    int.TryParse(GUILayout.TextField(ks.CTextF.ToString(), GUI.skin.textField), out cache);
                    GUILayout.FlexibleSpace();
                    GUILayout.EndHorizontal();
                    changed = cache != ks.CTextF;
                    ks.CTextF = cache;
                    return changed;
                }
                else
                {
                    bool changed = false;
                    int cache = ks.CTextF;
                    int cache2 = ks.CTextf;
                    GUILayout.BeginHorizontal();
                    GUILayout.Space(10);
                    GUILayout.Label("Min Font Size");
                    int.TryParse(GUILayout.TextField(ks.CTextf.ToString(), GUI.skin.textField), out cache2);
                    GUILayout.Label("Max Font Size");
                    int.TryParse(GUILayout.TextField(ks.CTextF.ToString(), GUI.skin.textField), out cache);
                    GUILayout.FlexibleSpace();
                    GUILayout.EndHorizontal();
                    changed = cache != ks.CTextF || cache2 != ks.CTextf;
                    ks.CTextF = cache;
                    return changed;
                }
            }

            return false;
        }
    }
}

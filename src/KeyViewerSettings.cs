using System.Collections.Generic;
using System.Xml.Serialization;
using DG.Tweening;
using UnityEngine;
using UnityModManagerNet;

namespace KeyViewer
{
    /// <summary>
    /// Settings for the Key Viewer tweak.
    /// </summary>
    public class KeyViewerSettings : UnityModManager.ModSettings
    {
        public bool OptimizeEase = true;
        public Ease ease = Ease.InElastic;
        public bool RainbowKT = false;
        public RainbowType type = RainbowType.Press;
        public bool Rainbow = false;
        public float rainbowspd = 0.5f;
        public float ed = 0.1f;
        public float sf = 0.9f;
        public bool IgnoreSkippedKeys = false;
        [XmlIgnore] public int Total = 0;
        public float danwi = 0;
        public int UpdateRate = 20;
        public bool SaveKeyCounts = false;
        public override void Save(UnityModManager.ModEntry modEntry) => Save(this, modEntry);
        public LanguageEnum Language { get; set; } = LanguageEnum.ENGLISH;
        /// <summary>
        /// The keys that are counted as input.
        /// </summary>
        public List<KeyCode> ActiveKeys { get; set; } = new List<KeyCode>();

        /// <summary>
        /// Old setting for showing the key viewer.
        /// </summary>
        public bool ShowKeyViewer { get; set; }

        /// <summary>
        /// Old setting for showing the key viewer only in gameplay.
        /// </summary>
        public bool ViewerOnlyGameplay { get; set; }

        /// <summary>
        /// Old setting for animating key presses.
        /// </summary>
        public bool AnimateKeys { get; set; } = true;

        /// <summary>
        /// Old setting for the key viewer size.
        /// </summary>
        public float KeyViewerSize { get; set; } = 100f;

        /// <summary>
        /// Old setting for the horizontal position of the key viewer.
        /// </summary>
        public float KeyViewerXPos { get; set; } = 0.89f;

        /// <summary>
        /// Old setting for the vertical position of the key viewer.
        /// </summary>
        public float KeyViewerYPos { get; set; } = 0.03f;

        /// <summary>
        /// Limit key on custom level selection scene.
        /// </summary>
        public bool LimitKeyOnCLS { get; set; } = true;

        /// <summary>
        /// Limit key on main screen.
        /// </summary>
        public bool LimitKeyOnMainScreen { get; set; } = true;

        private Color _pressedOutlineColor;

        /// <summary>
        /// Old setting for the outline color of pressed keys.
        /// </summary>
        public Color PressedOutlineColor
        {
            get => _pressedOutlineColor;
            set
            {
                _pressedOutlineColor = value;
                PressedOutlineColorHex = ColorUtility.ToHtmlStringRGBA(value);
            }
        }

        private Color _releasedOutlineColor;

        /// <summary>
        /// Old setting for the outline color of released keys.
        /// </summary>
        public Color ReleasedOutlineColor
        {
            get => _releasedOutlineColor;
            set
            {
                _releasedOutlineColor = value;
                ReleasedOutlineColorHex = ColorUtility.ToHtmlStringRGBA(value);
            }
        }

        private Color _pressedBackgroundColor;

        /// <summary>
        /// Old setting for the background/fill color of pressed keys.
        /// </summary>
        public Color PressedBackgroundColor
        {
            get => _pressedBackgroundColor;
            set
            {
                _pressedBackgroundColor = value;
                PressedBackgroundColorHex = ColorUtility.ToHtmlStringRGBA(value);
            }
        }

        private Color _releasedBackgroundColor;

        /// <summary>
        /// Old setting for the background/fill color of released keys.
        /// </summary>
        public Color ReleasedBackgroundColor
        {
            get => _releasedBackgroundColor;
            set
            {
                _releasedBackgroundColor = value;
                ReleasedBackgroundColorHex = ColorUtility.ToHtmlStringRGBA(value);
            }
        }

        private Color _pressedTextColor;

        /// <summary>
        /// Old setting for the text color of pressed keys.
        /// </summary>
        public Color PressedTextColor
        {
            get => _pressedTextColor;
            set
            {
                _pressedTextColor = value;
                PressedTextColorHex = ColorUtility.ToHtmlStringRGBA(value);
            }
        }

        private Color _releasedTextColor;

        /// <summary>
        /// Old setting for the text color of released keys.
        /// </summary>
        public Color ReleasedTextColor
        {
            get => _releasedTextColor;
            set
            {
                _releasedTextColor = value;
                ReleasedTextColorHex = ColorUtility.ToHtmlStringRGBA(value);
            }
        }

        /// <summary>
        /// Old setting for the hex code for the pressed outline color.
        /// </summary>
        [XmlIgnore]
        public string PressedOutlineColorHex { get; set; }

        /// <summary>
        /// Old setting for the hex code for the released outline color.
        /// </summary>
        [XmlIgnore]
        public string ReleasedOutlineColorHex { get; set; }

        /// <summary>
        /// Old setting for the hex code for the pressed background/fill color.
        /// </summary>
        [XmlIgnore]
        public string PressedBackgroundColorHex { get; set; }

        /// <summary>
        /// Old setting for the hex code for the released background/fill color.
        /// </summary>
        [XmlIgnore]
        public string ReleasedBackgroundColorHex { get; set; }

        /// <summary>
        /// Old setting for the hex code for the pressed text color.
        /// </summary>
        [XmlIgnore]
        public string PressedTextColorHex { get; set; }

        /// <summary>
        /// Old setting for the hex code for the released text color.
        /// </summary>
        [XmlIgnore]
        public string ReleasedTextColorHex { get; set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="KeyLimiterSettings"/>
        /// class with some default values set.
        /// </summary>
        public KeyViewerSettings()
        {
            // Settings no one would ever use
            PressedOutlineColor = Color.black;
            ReleasedOutlineColor = Color.black;
            PressedBackgroundColor = Color.black;
            ReleasedBackgroundColor = Color.black;
            PressedTextColor = Color.black;
            ReleasedTextColor = Color.black;
            Profiles = new List<KeyViewerProfile>();
            ProfileIndex = 0;
        }
        /// <summary>
        /// A list of profiles that the user has saved.
        /// </summary>
        public List<KeyViewerProfile> Profiles { get; set; }

        /// <summary>
        /// The index of the current profile being used.
        /// </summary>
        public int ProfileIndex { get; set; }

        /// <summary>
        /// A reference to the current profile being used.
        /// </summary>
        [XmlIgnore]
        public KeyViewerProfile CurrentProfile { get => Profiles[ProfileIndex]; }

        /// <summary>
        /// Whether the tweak is listening for keystrokes to change what keys
        /// are active.
        /// </summary>
        [XmlIgnore]
        public bool IsListening { get; set; }
    }
}

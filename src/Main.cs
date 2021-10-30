using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using UnityEngine;
using System.IO;
using System.Reflection;
using LiteDB;
using UnityModManagerNet;
using System.Globalization;
using System.Threading;
using System.Diagnostics;
using System.Management.Instrumentation;
using static UnityModManagerNet.UnityModManager;
using HarmonyLib;
using System.Windows.Forms;
using 불 = System.Boolean;
using 정수 = System.Int32;
using 어 = System.Collections.IEnumerable;
using 아잇 = System.Collections.IEnumerator;
using System.Runtime.CompilerServices;
using SD;

namespace KeyViewer
{
    public class KeyList : IEnumerable<Key>, IEnumerator<Key>
    {
        private List<Key> keys = new();
        private int pos = -1;

        public Key this[int i]
        {
            get => keys[i];
            set => keys[i] = value;
        }

        public Key Remove(Key k)
        {
            Key cache = k;
            keys.Remove(k);
            return cache;
        }

        public Key Add(Key key)
        {
            keys.Add(key);
            return key;
        }
        public void Reverse() => keys.Reverse();
        public void RemoveAll(Predicate<Key> predicate) => keys.RemoveAll(predicate);
        public void RemoveRange(int index, int count) => keys.RemoveRange(index, count);
        public Key RemoveAt(int index)
        {
            Key cache = keys[index];
            keys.RemoveAt(index);
            return cache;
        }
        아잇 어.GetEnumerator() => GetEnumerator();
        public IEnumerator<Key> GetEnumerator() => this;
        public bool MoveNext()
        {
            if (pos < keys.Count)
            {
                ++pos;
                return true;
            }

            return false;
        }
        public Key Current => keys[pos];
        object 아잇.Current => Current;
        public void Reset() => pos = -1;
        public void Dispose() => GC.SuppressFinalize(this);
    }
    public class AtomicInteger : IEquatable<AtomicInteger>, IEquatable<int>, IFormattable, IConvertible
    {
        public static implicit operator AtomicInteger(int value) => new AtomicInteger(value);
        private int value;

        public void Set(int value) => Value = value;
        /// <summary>
        /// Creates a new <see cref="AtomicInteger"/> with the default inital value, <c>0</c>.
        /// </summary>
        public AtomicInteger()
            : this(0)
        { }

        /// <summary>
        /// Creates a new <see cref="AtomicInteger"/> with the given initial <paramref name="value"/>.
        /// </summary>
        /// <param name="value">The initial value.</param>
        public AtomicInteger(int value)
        {
            this.value = value;
        }

        /// <summary>
        /// Gets or sets the current value. Note that these operations can be done
        /// implicitly by setting the <see cref="AtomicInteger"/> to an <see cref="int"/>.
        /// <code>
        /// AtomicInteger aint = new AtomicInteger(4);
        /// int x = aint;
        /// </code>
        /// </summary>
        /// <remarks>
        /// Properties are inherently not atomic. Operators such as ++ and -- should not
        /// be used on <see cref="Value"/> because they perform both a separate get and a set operation. Instead,
        /// use the atomic methods such as <see cref="IncrementAndGet()"/>, <see cref="GetAndIncrement()"/>
        /// <see cref="DecrementAndGet()"/> and <see cref="GetAndDecrement()"/> to ensure atomicity.
        /// </remarks>
        public int Value // Port Note: This is a replacement for Get() and Set()
        {
            get => Interlocked.CompareExchange(ref value, 0, 0);
            set => Interlocked.Exchange(ref this.value, value);
        }

        /// <summary>
        /// Atomically sets to the given value and returns the old value.
        /// </summary>
        /// <param name="newValue">The new value.</param>
        /// <returns>The previous value.</returns>
        public int GetAndSet(int newValue)
        {
            return Interlocked.Exchange(ref value, newValue);
        }

        /// <summary>
        /// Atomically sets the value to the given updated value
        /// if the current value equals the expected value.
        /// </summary>
        /// <param name="expect">The expected value (the comparand).</param>
        /// <param name="update">The new value that will be set if the current value equals the expected value.</param>
        /// <returns><c>true</c> if successful. A <c>false</c> return value indicates that the actual value
        /// was not equal to the expected value.</returns>
        public bool CompareAndSet(int expect, int update)
        {
            int rc = Interlocked.CompareExchange(ref value, update, expect);
            return rc == expect;
        }

        /// <summary>
        /// Atomically increments by one the current value.
        /// </summary>
        /// <returns>The previous value, before the increment.</returns>
        public int GetAndIncrement()
        {
            return Interlocked.Increment(ref value) - 1;
        }

        /// <summary>
        /// Atomically decrements by one the current value.
        /// </summary>
        /// <returns>The previous value, before the decrement.</returns>
        public int GetAndDecrement()
        {
            return Interlocked.Decrement(ref value) + 1;
        }

        /// <summary>
        /// Atomically adds the given value to the current value.
        /// </summary>
        /// <param name="value">The value to add.</param>
        /// <returns>The previous value, before the addition.</returns>
        public int GetAndAdd(int value)
        {
            return Interlocked.Add(ref this.value, value) - value;
        }

        /// <summary>
        /// Atomically increments by one the current value.
        /// </summary>
        /// <returns>The updated value.</returns>
        public int IncrementAndGet()
        {
            return Interlocked.Increment(ref value);
        }

        /// <summary>
        /// Atomically decrements by one the current value.
        /// </summary>
        /// <returns>The updated value.</returns>
        public int DecrementAndGet()
        {
            return Interlocked.Decrement(ref value);
        }

        /// <summary>
        /// Atomically adds the given value to the current value.
        /// </summary>
        /// <param name="value">The value to add.</param>
        /// <returns>The updated value.</returns>
        public int AddAndGet(int value)
        {
            return Interlocked.Add(ref this.value, value);
        }

        /// <summary>
        /// Determines whether the specified <see cref="AtomicInteger"/> is equal to the current <see cref="AtomicInteger"/>.
        /// </summary>
        /// <param name="other">The <see cref="AtomicInteger"/> to compare with the current <see cref="AtomicInteger"/>.</param>
        /// <returns><c>true</c> if <paramref name="other"/> is equal to the current <see cref="AtomicInteger"/>; otherwise, <c>false</c>.</returns>
        public bool Equals(AtomicInteger? other)
        {
            if (other is null)
                return false;
            return Value == other.Value;
        }

        /// <summary>
        /// Determines whether the specified <see cref="int"/> is equal to the current <see cref="AtomicInteger"/>.
        /// </summary>
        /// <param name="other">The <see cref="int"/> to compare with the current <see cref="AtomicInteger"/>.</param>
        /// <returns><c>true</c> if <paramref name="other"/> is equal to the current <see cref="AtomicInteger"/>; otherwise, <c>false</c>.</returns>
        public bool Equals(int other)
        {
            return Value == other;
        }

        /// <summary>
        /// Determines whether the specified <see cref="object"/> is equal to the current <see cref="AtomicInteger"/>.
        /// <para/>
        /// If <paramref name="other"/> is a <see cref="AtomicInteger"/>, the comparison is not done atomically.
        /// </summary>
        /// <param name="other">The <see cref="object"/> to compare with the current <see cref="AtomicInteger"/>.</param>
        /// <returns><c>true</c> if <paramref name="other"/> is equal to the current <see cref="AtomicInteger"/>; otherwise, <c>false</c>.</returns>
        public override bool Equals(object? other)
        {
            if (other is AtomicInteger ai)
                return Equals(ai);
            if (other is int i)
                return Equals(i);
            return false;
        }

        /// <summary>
        /// Returns the hash code for this instance.
        /// </summary>
        /// <returns>A 32-bit signed integer hash code.</returns>
        public override int GetHashCode()
        {
            return Value.GetHashCode();
        }
        /// <summary>
        /// Returns the <see cref="TypeCode"/> for value type <see cref="int"/>.
        /// </summary>
        /// <returns></returns>
        public TypeCode GetTypeCode() => TypeCode.Int32;

        bool IConvertible.ToBoolean(IFormatProvider? provider) => Convert.ToBoolean(Value);

        byte IConvertible.ToByte(IFormatProvider? provider) => Convert.ToByte(Value);

        char IConvertible.ToChar(IFormatProvider? provider) => Convert.ToChar(Value);

        DateTime IConvertible.ToDateTime(IFormatProvider? provider) => Convert.ToDateTime(Value);

        decimal IConvertible.ToDecimal(IFormatProvider? provider) => Convert.ToDecimal(Value);

        double IConvertible.ToDouble(IFormatProvider? provider) => Convert.ToDouble(Value);

        short IConvertible.ToInt16(IFormatProvider? provider) => Convert.ToInt16(Value);

        int IConvertible.ToInt32(IFormatProvider? provider) => Value;

        long IConvertible.ToInt64(IFormatProvider? provider) => Convert.ToInt64(Value);

        sbyte IConvertible.ToSByte(IFormatProvider? provider) => Convert.ToSByte(Value);

        float IConvertible.ToSingle(IFormatProvider? provider) => Convert.ToSingle(Value);

        object IConvertible.ToType(Type conversionType, IFormatProvider? provider) => ((IConvertible)Value).ToType(conversionType, provider);

        ushort IConvertible.ToUInt16(IFormatProvider? provider) => Convert.ToUInt16(Value);

        uint IConvertible.ToUInt32(IFormatProvider? provider) => Convert.ToUInt32(Value);

        ulong IConvertible.ToUInt64(IFormatProvider? provider) => Convert.ToUInt64(Value);

        string IFormattable.ToString(string format, IFormatProvider formatProvider)
        {
            return Value.ToString(format, CultureInfo.InvariantCulture);
        }

        TypeCode IConvertible.GetTypeCode()
        {
            return TypeCode.Int32;
        }

        string IConvertible.ToString(IFormatProvider provider)
        {
            return Main.ToString(Value, provider);
        }

        #region Operator Overrides

        /// <summary>
        /// Implicitly converts an <see cref="AtomicInteger"/> to an <see cref="int"/>.
        /// </summary>
        /// <param name="AtomicInteger">The <see cref="AtomicInteger"/> to convert.</param>
        public static implicit operator int(AtomicInteger AtomicInteger)
        {
            return AtomicInteger.Value;
        }

        /// <summary>
        /// Compares <paramref name="a1"/> and <paramref name="a2"/> for value equality.
        /// </summary>
        /// <param name="a1">The first number.</param>
        /// <param name="a2">The second number.</param>
        /// <returns><c>true</c> if the given numbers are equal; otherwise, <c>false</c>.</returns>
        public static bool operator ==(AtomicInteger a1, AtomicInteger a2)
        {
            return a1.Value == a2.Value;
        }

        /// <summary>
        /// Compares <paramref name="a1"/> and <paramref name="a2"/> for value inequality.
        /// </summary>
        /// <param name="a1">The first number.</param>
        /// <param name="a2">The second number.</param>
        /// <returns><c>true</c> if the given numbers are not equal; otherwise, <c>false</c>.</returns>
        public static bool operator !=(AtomicInteger a1, AtomicInteger a2)
        {
            return !(a1 == a2);
        }

        /// <summary>
        /// Compares <paramref name="a1"/> and <paramref name="a2"/> for value equality.
        /// </summary>
        /// <param name="a1">The first number.</param>
        /// <param name="a2">The second number.</param>
        /// <returns><c>true</c> if the given numbers are equal; otherwise, <c>false</c>.</returns>
        public static bool operator ==(AtomicInteger a1, int a2)
        {
            return a1.Value == a2;
        }

        /// <summary>
        /// Compares <paramref name="a1"/> and <paramref name="a2"/> for value inequality.
        /// </summary>
        /// <param name="a1">The first number.</param>
        /// <param name="a2">The second number.</param>
        /// <returns><c>true</c> if the given numbers are not equal; otherwise, <c>false</c>.</returns>
        public static bool operator !=(AtomicInteger a1, int a2)
        {
            return !(a1 == a2);
        }

        /// <summary>
        /// Compares <paramref name="a1"/> and <paramref name="a2"/> for value equality.
        /// </summary>
        /// <param name="a1">The first number.</param>
        /// <param name="a2">The second number.</param>
        /// <returns><c>true</c> if the given numbers are equal; otherwise, <c>false</c>.</returns>
        public static bool operator ==(int a1, AtomicInteger a2)
        {
            return a1 == a2.Value;
        }

        /// <summary>
        /// Compares <paramref name="a1"/> and <paramref name="a2"/> for value inequality.
        /// </summary>
        /// <param name="a1">The first number.</param>
        /// <param name="a2">The second number.</param>
        /// <returns><c>true</c> if the given numbers are not equal; otherwise, <c>false</c>.</returns>
        public static bool operator !=(int a1, AtomicInteger a2)
        {
            return !(a1 == a2);
        }

        /// <summary>
        /// Compares <paramref name="a1"/> and <paramref name="a2"/> for value equality.
        /// </summary>
        /// <param name="a1">The first number.</param>
        /// <param name="a2">The second number.</param>
        /// <returns><c>true</c> if the given numbers are equal; otherwise, <c>false</c>.</returns>
        public static bool operator ==(AtomicInteger a1, int? a2)
        {
            return a1.Value == a2.GetValueOrDefault();
        }

        /// <summary>
        /// Compares <paramref name="a1"/> and <paramref name="a2"/> for value inequality.
        /// </summary>
        /// <param name="a1">The first number.</param>
        /// <param name="a2">The second number.</param>
        /// <returns><c>true</c> if the given numbers are not equal; otherwise, <c>false</c>.</returns>
        public static bool operator !=(AtomicInteger a1, int? a2)
        {
            return !(a1 == a2);
        }

        /// <summary>
        /// Compares <paramref name="a1"/> and <paramref name="a2"/> for value equality.
        /// </summary>
        /// <param name="a1">The first number.</param>
        /// <param name="a2">The second number.</param>
        /// <returns><c>true</c> if the given numbers are equal; otherwise, <c>false</c>.</returns>
        public static bool operator ==(int? a1, AtomicInteger a2)
        {
            return a1.GetValueOrDefault() == a2.Value;
        }

        /// <summary>
        /// Compares <paramref name="a1"/> and <paramref name="a2"/> for value inequality.
        /// </summary>
        /// <param name="a1">The first number.</param>
        /// <param name="a2">The second number.</param>
        /// <returns><c>true</c> if the given numbers are not equal; otherwise, <c>false</c>.</returns>
        public static bool operator !=(int? a1, AtomicInteger a2)
        {
            return !(a1 == a2);
        }

        #endregion Operator Overrides
    }
    internal static class Calculate
    {
        internal static int kps;
        internal static AtomicInteger tmp = new AtomicInteger(0);
        private static LinkedList<int> timepoints = new LinkedList<int>();
        private static 정수 max;
        private static 정수 prev;
        private static long n = 0;
        public static double avg;
        internal static Thread t;
        internal static void Set()
        {
            ChangeUpdateRate(Main.Settings.UpdateRate);
            t = new Thread(Run);
            t.Start();
        }
        internal static void Refresh()
        {
           
        }
        internal static void UnSet()
        {
            t.Abort();
        }
        internal static void ChangeUpdateRate(int newRate)
        {
            n *= (long)((double)Main.Settings.UpdateRate / (double)newRate);
            tmp.Set(0);
            timepoints.Clear();
            Main.Settings.UpdateRate = newRate;
            Refresh();
        }
        internal static void Press()
        {
            tmp.IncrementAndGet();
        }
        internal static void Reset()
        {
            n = 0;
            avg = 0;
            max = 0;
            tmp.Set(0);
        }
        internal static Stopwatch s = new Stopwatch();
        private static void Run()
        {
            s.Start();
            while (Main.IsEnabled)
            {
                if (s.ElapsedMilliseconds >= Main.Settings.UpdateRate)
                {
                    int currentTmp = tmp.GetAndSet(0);
                    int totaltmp = currentTmp;
                    foreach (int i in timepoints)
                    {
                        totaltmp += i;
                    }
                    if (totaltmp > max)
                    {
                        max = totaltmp;
                    }
                    if (totaltmp != 0)
                    {
                        avg = (avg * n + totaltmp) / (n + 1.0D);
                        n++;
                        Main.Settings.Total += currentTmp;
                    }
                    prev = totaltmp;
                    timepoints.AddFirst(currentTmp);
                    if (timepoints.Count >= 1000 / Main.Settings.UpdateRate)
                    {
                        timepoints.RemoveLast();
                    }
                    kps = totaltmp;
                    s.Restart();
                }
            }
        }
    }

    /// <summary>
    /// The POCO for storing translation data.
    /// </summary>
    public class TweakString
    {
        /// <summary>
        /// The translation key for this tweak string.
        /// </summary>
        public string Key { get; set; }

        /// <summary>
        /// The content of the tweak string.
        /// </summary>
        public string Content { get; set; }

        /// <summary>
        /// The language of the tweak string.
        /// </summary>
        public LanguageEnum Language { get; set; }
    }
    /// <summary>
    /// Extensions for <see cref="LanguageEnum"/>.
    /// </summary>
    public static class LanguageEnumExtensions
    {
        /// <summary>
        /// Returns <c>true</c> for languages that have hard to read symbols on
        /// smaller fonts (e.g. Korean, Chinese, etc.).
        /// </summary>
        /// <param name="language">The language to check.</param>
        /// <returns><c>true</c> if the language contains symbols.</returns>
        public static bool IsSymbolLanguage(this LanguageEnum language)
        {
            return language == LanguageEnum.KOREAN
                || language == LanguageEnum.CHINESE_SIMPLIFIED;
        }
    }
    /// <summary>
    /// Supported languages in AdofaiTweaks.
    /// </summary>
    public enum LanguageEnum
    {
        ENGLISH = 0,
        KOREAN = 1,
        SPANISH = 2,
        POLISH = 3,
        FRENCH = 4,
        VIETNAMESE = 5,
        CHINESE_SIMPLIFIED = 6,
    }
    /// <summary>
    /// Wrapper around LiteDB to read the tweak strings from the database. When
    /// accessing any string for a given language, all strings for that language
    /// are read and cached.
    /// </summary>
    internal class TweakStringsDb
    {
        private static TweakStringsDb _instance;

        /// <summary>
        /// Singleton instance for <see cref="TweakStringsDb"/>.
        /// </summary>
        public static TweakStringsDb Instance
        {
            get
            {
                if (_instance == null)
                {
                    _instance = new TweakStringsDb();
                }
                return _instance;
            }
        }

        private readonly Dictionary<LanguageEnum, Dictionary<string, TweakString>> cache;

        private void LoadFromDb(LanguageEnum language)
        {
            using (var db = new LiteDatabase(Assembly.GetExecutingAssembly().GetManifestResourceStream("KeyViewer.KVLang.db")))
            {
                var collection = db.GetCollection<TweakString>();
                var results = collection.Query()
                    .Where(ts => ts.Language == language)
                    .ToEnumerable();
                var dict = new Dictionary<string, TweakString>();
                foreach (TweakString tweakString in results)
                {
                    dict[tweakString.Key] = tweakString;
                }
                cache[language] = dict;
            }
        }

        /// <summary>
        /// Gets the tweak string for the given language and key. If no
        /// translation is found, then the English translation is returned.
        /// </summary>
        /// <param name="language">
        /// The language that the tweak string will be translated as.
        /// </param>
        /// <param name="key">The key for the tweak string.</param>
        /// <returns>The translated tweak string.</returns>
        public string GetForLanguage(LanguageEnum language, string key)
        {
            if (!cache.ContainsKey(language))
            {
                LoadFromDb(language);
            }
            Dictionary<string, TweakString> dict = cache[language];
            if (!dict.ContainsKey(key))
            {
                return $"no such key {key}";
            }
            if (string.IsNullOrEmpty(dict[key].Content))
            {
                return cache[LanguageEnum.ENGLISH][key].Content;
            }
            return dict[key].Content;
        }

        private TweakStringsDb()
        {
            cache = new Dictionary<LanguageEnum, Dictionary<string, TweakString>>();
            LoadFromDb(LanguageEnum.ENGLISH);
        }
    }
    /// <summary>
    /// Public-facing API for retrieving the translated tweak string from the
    /// database.
    /// </summary>
    public static class TweakStrings
    {
        /// <summary>
        /// <para>
        /// Gets the tweak string for the specified key with optional arguments
        /// to insert into the string. The arguments will be inserted via
        /// standard <see cref="string.Format"/> rules.
        /// </para>
        /// <para>
        /// The string will be translated based on the user's current settings.
        /// Use the keys from <see cref="TranslationKeys"/> to ensure the
        /// correct strings are retrieved.
        /// </para>
        /// </summary>
        /// <param name="key">The key of the tweak string.</param>
        /// <param name="args">The arguments to insert into the string.</param>
        /// <returns>The translated string.</returns>
        public static string Get(string key, params object[] args)
        {
            return GetForLanguage(key, Main.Settings.Language, args);
        }

        /// <summary>
        /// <para>
        /// Gets the tweak string for the specified key and language with
        /// optional arguments to insert into the string. The arguments will be
        /// inserted via standard <see cref="string.Format"/> rules.
        /// </para>
        /// <para>
        /// Use the keys from <see cref="TranslationKeys"/> to ensure the
        /// correct strings are retrieved.
        /// </para>
        /// </summary>
        /// <param name="key">The key of the tweak string.</param>
        /// <param name="language">
        /// The language to use for the translation.
        /// </param>
        /// <param name="args">The arguments to insert into the string.</param>
        /// <returns>The translated string.</returns>
        public static string GetForLanguage(
            string key, LanguageEnum language, params object[] args)
        {
            string format = TweakStringsDb.Instance.GetForLanguage(language, key);
            if (format == null)
            {
                return $"no such key {key}";
            }
            if (args.Length == 0)
            {
                return format;
            }
            return string.Format(format, args);
        }
    }
    public class TK
    {
        public class G
        {
            public const string TEST_KEY = "GLOBAL_TEST_KEY";

            public const string GLOBAL_LANGUAGE = "GLOBAL_GLOBAL_LANGUAGE";

            public const string LANGUAGE_NAME = "GLOBAL_LANGUAGE_NAME";
        }

        public class KV
        {
            public const string SAVE_KEYCOUNTS = "KEY_VIEWER_SAVE_KEYCOUNTS";

            public const string IGNORE_SKIPPED_KEYS = "KEY_VIEWER_IGNORE_SKIPPED_KEYS";

            public const string NAME = "KEY_VIEWER_NAME";

            public const string DESCRIPTION = "KEY_VIEWER_DESCRIPTION";

            public const string REGISTERED_KEYS = "KEY_VIEWER_REGISTERED_KEYS";

            public const string DONE = "KEY_VIEWER_DONE";

            public const string PRESS_KEY_REGISTER = "KEY_VIEWER_PRESS_KEY_REGISTER";

            public const string CHANGE_KEYS = "KEY_VIEWER_CHANGE_KEYS";

            public const string VIEWER_ONLY_GAMEPLAY = "KEY_VIEWER_VIEWER_ONLY_GAMEPLAY";

            public const string ANIMATE_KEYS = "KEY_VIEWER_ANIMATE_KEYS";

            public const string KEY_VIEWER_SIZE = "KEY_VIEWER_KEY_VIEWER_SIZE";

            public const string KEY_VIEWER_X_POS = "KEY_VIEWER_KEY_VIEWER_X_POS";

            public const string KEY_VIEWER_Y_POS = "KEY_VIEWER_KEY_VIEWER_Y_POS";

            public const string PRESSED_OUTLINE_COLOR = "KEY_VIEWER_PRESSED_OUTLINE_COLOR";

            public const string RELEASED_OUTLINE_COLOR = "KEY_VIEWER_RELEASED_OUTLINE_COLOR";

            public const string PRESSED_BACKGROUND_COLOR = "KEY_VIEWER_PRESSED_BACKGROUND_COLOR";

            public const string RELEASED_BACKGROUND_COLOR = "KEY_VIEWER_RELEASED_BACKGROUND_COLOR";

            public const string PRESSED_TEXT_COLOR = "KEY_VIEWER_PRESSED_TEXT_COLOR";

            public const string RELEASED_TEXT_COLOR = "KEY_VIEWER_RELEASED_TEXT_COLOR";

            public const string PROFILES = "KEY_VIEWER_PROFILES";

            public const string PROFILE_NAME = "KEY_VIEWER_PROFILE_NAME";

            public const string NEW = "KEY_VIEWER_NEW";

            public const string DUPLICATE = "KEY_VIEWER_DUPLICATE";

            public const string DELETE = "KEY_VIEWER_DELETE";

            public const string SHOW_KEY_PRESS_TOTAL = "KEY_VIEWER_SHOW_KEY_PRESS_TOTAL";
        }
    }
    public static class Asset
    {
        /// <summary>
        /// The normal font to use for text for symbol languages.
        /// </summary>
        public static Font SymbolLangNormalFont { get; private set; }

        /// <summary>
        /// The bold font to use for Korean text.
        /// </summary>
        public static Font KoreanBoldFont { get; private set; }
        /// <summary>
        /// The sprite for the key's outline in the key viewer.
        /// </summary>
        public static Sprite KeyOutlineSprite { get; private set; }

        /// <summary>
        /// The sprite for the key's background fill in the key viewer.
        /// </summary>
        public static Sprite KeyBackgroundSprite { get; private set; }

        private static readonly AssetBundle assets;
        static Asset()
        {
            assets = Utils.LoadBundle();
            SymbolLangNormalFont = assets.LoadAsset<Font>("Assets/NanumGothic-Regular.ttf");
            KoreanBoldFont = assets.LoadAsset<Font>("Assets/NanumGothic-Bold.ttf");
            KeyOutlineSprite = assets.LoadAsset<Sprite>("Assets/KeyOutline.png");
            KeyBackgroundSprite = assets.LoadAsset<Sprite>("Assets/KeyBackground.png");
        }
    }
    public static class Utils
    {
        private static byte[] ReadFully(this Stream input)
        {
            using (var ms = new MemoryStream())
            {
                byte[] buffer = new byte[81920];
                int read;
                while ((read = input.Read(buffer, 0, buffer.Length)) != 0)
                    ms.Write(buffer, 0, read);
                return ms.ToArray();
            }
        }
        public static AssetBundle LoadBundle()
        {
            return AssetBundle.LoadFromMemory(Assembly.GetExecutingAssembly().GetManifestResourceStream("KeyViewer.key.assets").ReadFully());
        }

    }
    public static class Main
    {
        [ModuleInitializer]
        public static void Initialize()
        {
            return;
        }

        public static 불 ImplementsGenericInterface(this Type target, Type interfaceType)
        {
            if (target is null)
                throw new ArgumentNullException(nameof(target));
            if (interfaceType is null)
                return false;

#if FEATURE_TYPEEXTENSIONS_GETTYPEINFO
            return target.GetTypeInfo().IsGenericType && target.GetGenericTypeDefinition().GetInterfaces().Any(
                x => x.GetTypeInfo().IsGenericType && interfaceType.IsAssignableFrom(x.GetGenericTypeDefinition())
#else
            return target.IsGenericType && target.GetGenericTypeDefinition().GetInterfaces().Any(
                x => x.IsGenericType && interfaceType.IsAssignableFrom(x.GetGenericTypeDefinition())
#endif
            );
        }
        /// <summary>
        /// This is a helper method that assists with recursively building
        /// a string of the current collection and all nested collections, plus the ability
        /// to specify culture for formatting of nested numbers and dates. Note that
        /// this overload will change the culture of the current thread.
        /// </summary>
        public static string ToString(object? obj, IFormatProvider? provider)
        {
            if (obj is null) return "null";
            Type t = obj.GetType();
            if (
#if FEATURE_TYPEEXTENSIONS_GETTYPEINFO
                t.GetTypeInfo().IsGenericType
#else
                t.IsGenericType
#endif
                && (t.ImplementsGenericInterface(typeof(ICollection<>)))
                || t.ImplementsGenericInterface(typeof(IDictionary<,>)))
            {
                return ToStringImpl(obj, t, provider);
            }

            return obj.ToString()!;
        }

        public static string ToStringImpl(object? obj, Type type, IFormatProvider? provider)
        {
            dynamic? genericType = Convert.ChangeType(obj, type);
            return ToString(genericType, provider);
        }
        public static bool IsEnabled { get; private set; }
        public static KeyViewerSettings Settings { get; private set; } = new KeyViewerSettings();
        public static ModEntry.ModLogger Logger { get; private set; }
        public static ModEntry modEntry { get; private set; }
        public static Harmony harmony { get; private set; }
        public static bool Load(ModEntry modEntry)
        {
            Settings = ModSettings.Load<KeyViewerSettings>(modEntry);
            Main.modEntry = modEntry;
            Logger = modEntry.Logger;
            modEntry.OnToggle = OnToggle;
            modEntry.OnGUI = OnGUI;
            modEntry.OnHideGUI = KeyViewerTweak.OnHideGUI;
            modEntry.OnUpdate = KeyViewerTweak.OnUpdate;
            modEntry.OnSaveGUI = (e) =>
            {
                if (Settings.SaveKeyCounts)
                {
                    KeyViewer.keyCounts.SerializeJson(Path.Combine("Mods", "KeyViewer", "KeyCounts.kc"));
                }
                Key.KeySettings.SerializeJson(Path.Combine("Mods", "KeyViewer", "KeySettings.ks"));
                //KeyViewer.keyposize.SerializeJson(Path.Combine("Mods", "KeyViewer", "Keys.kv"));
                Settings.Save(modEntry);
            };
            return true;
        }
        public static void OnGUI(ModEntry entry)
        {
            // Set some default GUI settings for better layouts
            if (Settings.Language.IsSymbolLanguage())
            {
                GUI.skin.button.font = Asset.SymbolLangNormalFont;
                GUI.skin.label.font = Asset.SymbolLangNormalFont;
                GUI.skin.textArea.font = Asset.SymbolLangNormalFont;
                GUI.skin.textField.font = Asset.SymbolLangNormalFont;
                GUI.skin.toggle.font = Asset.SymbolLangNormalFont;
                GUI.skin.button.fontSize = 15;
                GUI.skin.label.fontSize = 15;
                GUI.skin.textArea.fontSize = 15;
                GUI.skin.textField.fontSize = 15;
                GUI.skin.toggle.fontSize = 15;
            }
            GUI.skin.toggle = new GUIStyle(GUI.skin.toggle)
            {
                margin = new RectOffset(0, 4, 6, 6),
            };
            GUI.skin.label.wordWrap = false;

            GUILayout.Space(4);

            // Language chooser
            GUILayout.BeginHorizontal();
            GUILayout.Space(4);
            GUILayout.Label(
                TweakStrings.Get(TK.G.GLOBAL_LANGUAGE),
                new GUIStyle(GUI.skin.label)
                {
                    fontStyle = Settings.Language.IsSymbolLanguage()
                        ? FontStyle.Normal
                        : FontStyle.Bold,
                    font = Settings.Language.IsSymbolLanguage()
                        ? Asset.KoreanBoldFont
                        : null,
                });
            foreach (LanguageEnum language in Enum.GetValues(typeof(LanguageEnum)))
            {
                string langString =
                    TweakStrings.GetForLanguage(TK.G.LANGUAGE_NAME, language);

                // Set special styles for selected and Korean language
                GUIStyle style = new GUIStyle(GUI.skin.button);
                if (language == Settings.Language)
                {
                    if (language.IsSymbolLanguage())
                    {
                        style.font = Asset.KoreanBoldFont;
                        style.fontSize = 15;
                    }
                    else
                    {
                        style.fontStyle = FontStyle.Bold;
                    }
                }
                else if (language.IsSymbolLanguage())
                {
                    style.font = Asset.SymbolLangNormalFont;
                    style.fontSize = 15;
                }

                bool click = GUILayout.Button(langString, style);
                if (click)
                {
                    Settings.Language = language;
                }
            }
            GUILayout.FlexibleSpace();
            GUILayout.EndHorizontal();

            // Show each tweak's GUI
            GUILayout.Space(4);
            KeyViewerTweak.OnSettingsGUI(entry);
            GUI.skin.button.font = null;
            GUI.skin.label.font = null;
            GUI.skin.textArea.font = null;
            GUI.skin.textField.font = null;
            GUI.skin.toggle.font = null;
            GUI.skin.button.fontSize = 0;
            GUI.skin.label.fontSize = 0;
            GUI.skin.textArea.fontSize = 0;
            GUI.skin.textField.fontSize = 0;
            GUI.skin.toggle.fontSize = 0;
        }
        public static bool OnToggle(ModEntry modEntry, bool value)
        {
            IsEnabled = value;
            if (IsEnabled)
            {
                if (File.Exists(Path.Combine("Mods", "KeyViewer", "KeySettings.ks")))
                {
                    Key.KeySettings = Path.Combine("Mods", "KeyViewer", "KeySettings.ks")
                        .DeserializeJson<Dictionary<KeyCode, KeySetting>>();
                }
                if (Settings.SaveKeyCounts && File.Exists(Path.Combine("Mods", "KeyViewer", "KeyCounts.kc")))
                {
                    KeyViewer.keyCounts = Path.Combine("Mods", "KeyViewer", "KeyCounts.kc")
                        .DeserializeJson<Dictionary<KeyCode, int>>();
                    foreach (int i in KeyViewer.keyCounts.Values)
                    {
                        Settings.Total += i;
                    }
                }
                else Settings.Total = 0;
                //if (File.Exists(Path.Combine("Mods", "KeyViewer", "Keys.kv")))
                //{
                //    KeyViewer.keyposize = Path.Combine("Mods", "KeyViewer", "Keys.kv").DeserializeJson<Dictionary<KeyCode, PoSize>>();
                //}
                harmony = new Harmony(modEntry.Info.Id);
                harmony.PatchAll(Assembly.GetExecutingAssembly());
                KeyViewerTweak.OnEnable();
            }
            else
            {
                Key.KeySettings.SerializeJson(Path.Combine("Mods", "KeyViewer", "KeySettings.ks"));
                if (Settings.SaveKeyCounts)
                {
                    KeyViewer.keyCounts.SerializeJson(Path.Combine("Mods", "KeyViewer", "KeyCounts.kc"));
                }
                harmony.UnpatchAll(harmony.Id);
                harmony = null;
                KeyViewerTweak.OnDisable();
            }
            return true;
        }
    }
}

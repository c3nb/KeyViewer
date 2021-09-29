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
using static UnityModManagerNet.UnityModManager;
using HarmonyLib;
using System.Windows.Forms;
using 불 = System.Boolean;
using 정수 = System.Int32;
using 아잇 = System.Collections.IEnumerable;
using 어 = System.Collections.IEnumerator;
using System.Runtime.CompilerServices;
using SD;

namespace KeyViewer
{
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
        internal static 정수 count;
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
            count++;
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
    internal class Lang
    {
        internal const string IGNORE_SKIPPED_KEYS = "스킵된 키 무시";
        internal const string IGNORE_SKIPPED_KEYS_ENG = "Ignore Skipped Keys";
        internal const string SAVE_KEYCOUNTS = "누른 횟수 저장";
        internal const string SAVE_KEYCOUNTS_ENG = "Save KeyCounts";
    }
    public class TranslationKeys
    {
        // Token: 0x02000003 RID: 3
        public class Global
        {
            // Token: 0x04000001 RID: 1
            public static string TEST_KEY = "GLOBAL_TEST_KEY";

            // Token: 0x04000002 RID: 2
            public static string GLOBAL_LANGUAGE = "GLOBAL_GLOBAL_LANGUAGE";

            // Token: 0x04000003 RID: 3
            public static string LANGUAGE_NAME = "GLOBAL_LANGUAGE_NAME";
        }

        // Token: 0x02000004 RID: 4
        public class DisableEffects
        {
            // Token: 0x04000004 RID: 4
            public static string NAME = "DISABLE_EFFECTS_NAME";

            // Token: 0x04000005 RID: 5
            public static string DESCRIPTION = "DISABLE_EFFECTS_DESCRIPTION";

            // Token: 0x04000006 RID: 6
            public static string FILTER = "DISABLE_EFFECTS_FILTER";

            // Token: 0x04000007 RID: 7
            public static string BLOOM = "DISABLE_EFFECTS_BLOOM";

            // Token: 0x04000008 RID: 8
            public static string FLASH = "DISABLE_EFFECTS_FLASH";

            // Token: 0x04000009 RID: 9
            public static string HALL_OF_MIRRORS = "DISABLE_EFFECTS_HALL_OF_MIRRORS";

            // Token: 0x0400000A RID: 10
            public static string SCREEN_SHAKE = "DISABLE_EFFECTS_SCREEN_SHAKE";

            // Token: 0x0400000B RID: 11
            public static string TILE_MOVEMENT_MAX = "DISABLE_EFFECTS_TILE_MOVEMENT_MAX";

            // Token: 0x0400000C RID: 12
            public static string TILE_MOVEMENT_UNLIMITED = "DISABLE_EFFECTS_TILE_MOVEMENT_UNLIMITED";

            // Token: 0x0400000D RID: 13
            public static string EXCLUDE_FILTER_LIST = "DISABLE_EFFECTS_EXCLUDE_FILTER_LIST";

            // Token: 0x0400000E RID: 14
            public static string EXCLUDE_FILTER_START_LIST = "DISABLE_EFFECTS_EXCLUDE_FILTER_START_LIST";
        }

        // Token: 0x02000005 RID: 5
        public class HideUiElements
        {
            // Token: 0x0400000F RID: 15
            public static string NAME = "HIDE_UI_ELEMENTS_NAME";

            // Token: 0x04000010 RID: 16
            public static string DESCRIPTION = "HIDE_UI_ELEMENTS_DESCRIPTION";

            // Token: 0x04000011 RID: 17
            public static string EVERYTHING = "HIDE_UI_ELEMENTS_EVERYTHING";

            // Token: 0x04000012 RID: 18
            public static string JUDGE_TEXT = "HIDE_UI_ELEMENTS_JUDGE_TEXT";

            // Token: 0x04000013 RID: 19
            public static string MISSES = "HIDE_UI_ELEMENTS_MISSES";

            // Token: 0x04000014 RID: 20
            public static string SONG_TITLE = "HIDE_UI_ELEMENTS_SONG_TITLE";

            // Token: 0x04000015 RID: 21
            public static string AUTO = "HIDE_UI_ELEMENTS_AUTO";

            // Token: 0x04000016 RID: 22
            public static string BETA_BUILD = "HIDE_UI_ELEMENTS_BETA_BUILD";
        }

        // Token: 0x02000006 RID: 6
        public class JudgmentVisuals
        {
            // Token: 0x04000017 RID: 23
            public static string NAME = "JUDGMENT_VISUALS_NAME";

            // Token: 0x04000018 RID: 24
            public static string DESCRIPTION = "JUDGMENT_VISUALS_DESCRIPTION";

            // Token: 0x04000019 RID: 25
            public static string SHOW_HIT_ERROR_METER = "JUDGMENT_VISUALS_SHOW_HIT_ERROR_METER";

            // Token: 0x0400001A RID: 26
            public static string ERROR_METER_SCALE = "JUDGMENT_VISUALS_ERROR_METER_SCALE";

            // Token: 0x0400001B RID: 27
            public static string ERROR_METER_TICK_LIFE = "JUDGMENT_VISUALS_ERROR_METER_TICK_LIFE";

            // Token: 0x0400001C RID: 28
            public static string ERROR_METER_TICK_SECONDS = "JUDGMENT_VISUALS_ERROR_METER_TICK_SECONDS";

            // Token: 0x0400001D RID: 29
            public static string ERROR_METER_SENSITIVITY = "JUDGMENT_VISUALS_ERROR_METER_SENSITIVITY";

            // Token: 0x0400001E RID: 30
            public static string ERROR_METER_MORE_STABLE = "JUDGMENT_VISUALS_ERROR_METER_MORE_STABLE";

            // Token: 0x0400001F RID: 31
            public static string ERROR_METER_LESS_STABLE = "JUDGMENT_VISUALS_ERROR_METER_LESS_STABLE";

            // Token: 0x04000020 RID: 32
            public static string HIDE_PERFECTS = "JUDGMENT_VISUALS_HIDE_PERFECTS";

            // Token: 0x04000021 RID: 33
            public static string ERROR_METER_X_POS = "JUDGMENT_VISUALS_ERROR_METER_X_POS";

            // Token: 0x04000022 RID: 34
            public static string ERROR_METER_Y_POS = "JUDGMENT_VISUALS_ERROR_METER_Y_POS";
        }

        // Token: 0x02000007 RID: 7
        public class KeyViewer
        {
            // Token: 0x04000023 RID: 35
            public static string NAME = "KEY_VIEWER_NAME";

            // Token: 0x04000024 RID: 36
            public static string DESCRIPTION = "KEY_VIEWER_DESCRIPTION";

            // Token: 0x04000025 RID: 37
            public static string REGISTERED_KEYS = "KEY_VIEWER_REGISTERED_KEYS";

            // Token: 0x04000026 RID: 38
            public static string DONE = "KEY_VIEWER_DONE";

            // Token: 0x04000027 RID: 39
            public static string PRESS_KEY_REGISTER = "KEY_VIEWER_PRESS_KEY_REGISTER";

            // Token: 0x04000028 RID: 40
            public static string CHANGE_KEYS = "KEY_VIEWER_CHANGE_KEYS";

            // Token: 0x04000029 RID: 41
            public static string VIEWER_ONLY_GAMEPLAY = "KEY_VIEWER_VIEWER_ONLY_GAMEPLAY";

            // Token: 0x0400002A RID: 42
            public static string ANIMATE_KEYS = "KEY_VIEWER_ANIMATE_KEYS";

            // Token: 0x0400002B RID: 43
            public static string KEY_VIEWER_SIZE = "KEY_VIEWER_KEY_VIEWER_SIZE";

            // Token: 0x0400002C RID: 44
            public static string KEY_VIEWER_X_POS = "KEY_VIEWER_KEY_VIEWER_X_POS";

            // Token: 0x0400002D RID: 45
            public static string KEY_VIEWER_Y_POS = "KEY_VIEWER_KEY_VIEWER_Y_POS";

            // Token: 0x0400002E RID: 46
            public static string PRESSED_OUTLINE_COLOR = "KEY_VIEWER_PRESSED_OUTLINE_COLOR";

            // Token: 0x0400002F RID: 47
            public static string RELEASED_OUTLINE_COLOR = "KEY_VIEWER_RELEASED_OUTLINE_COLOR";

            // Token: 0x04000030 RID: 48
            public static string PRESSED_BACKGROUND_COLOR = "KEY_VIEWER_PRESSED_BACKGROUND_COLOR";

            // Token: 0x04000031 RID: 49
            public static string RELEASED_BACKGROUND_COLOR = "KEY_VIEWER_RELEASED_BACKGROUND_COLOR";

            // Token: 0x04000032 RID: 50
            public static string PRESSED_TEXT_COLOR = "KEY_VIEWER_PRESSED_TEXT_COLOR";

            // Token: 0x04000033 RID: 51
            public static string RELEASED_TEXT_COLOR = "KEY_VIEWER_RELEASED_TEXT_COLOR";

            // Token: 0x04000034 RID: 52
            public static string PROFILES = "KEY_VIEWER_PROFILES";

            // Token: 0x04000035 RID: 53
            public static string PROFILE_NAME = "KEY_VIEWER_PROFILE_NAME";

            // Token: 0x04000036 RID: 54
            public static string NEW = "KEY_VIEWER_NEW";

            // Token: 0x04000037 RID: 55
            public static string DUPLICATE = "KEY_VIEWER_DUPLICATE";

            // Token: 0x04000038 RID: 56
            public static string DELETE = "KEY_VIEWER_DELETE";

            // Token: 0x04000039 RID: 57
            public static string SHOW_KEY_PRESS_TOTAL = "KEY_VIEWER_SHOW_KEY_PRESS_TOTAL";
        }

        // Token: 0x02000008 RID: 8
        public class KeyLimiter
        {
            // Token: 0x0400003A RID: 58
            public static string NAME = "KEY_LIMITER_NAME";

            // Token: 0x0400003B RID: 59
            public static string DESCRIPTION = "KEY_LIMITER_DESCRIPTION";

            // Token: 0x0400003C RID: 60
            public static string REGISTERED_KEYS = "KEY_LIMITER_REGISTERED_KEYS";

            // Token: 0x0400003D RID: 61
            public static string DONE = "KEY_LIMITER_DONE";

            // Token: 0x0400003E RID: 62
            public static string PRESS_KEY_REGISTER = "KEY_LIMITER_PRESS_KEY_REGISTER";

            // Token: 0x0400003F RID: 63
            public static string CHANGE_KEYS = "KEY_LIMITER_CHANGE_KEYS";

            // Token: 0x04000040 RID: 64
            public static string LIMIT_CLS = "KEY_LIMITER_LIMIT_CLS";

            // Token: 0x04000041 RID: 65
            public static string LIMIT_MAIN_MENU = "KEY_LIMITER_LIMIT_MAIN_MENU";
        }

        // Token: 0x02000009 RID: 9
        public class PlanetOpacity
        {
            // Token: 0x04000042 RID: 66
            public static string NAME = "PLANET_OPACITY_NAME";

            // Token: 0x04000043 RID: 67
            public static string DESCRIPTION = "PLANET_OPACITY_DESCRIPTION";

            // Token: 0x04000044 RID: 68
            public static string PLANET_ONE = "PLANET_OPACITY_PLANET_ONE";

            // Token: 0x04000045 RID: 69
            public static string PLANET_TWO = "PLANET_OPACITY_PLANET_TWO";
        }

        // Token: 0x0200000A RID: 10
        public class Miscellaneous
        {
            // Token: 0x04000046 RID: 70
            public static string NAME = "MISCELLANEOUS_NAME";

            // Token: 0x04000047 RID: 71
            public static string DESCRIPTION = "MISCELLANEOUS_DESCRIPTION";

            // Token: 0x04000048 RID: 72
            public static string GLITCH_FLIP = "MISCELLANEOUS_GLITCH_FLIP";

            // Token: 0x04000049 RID: 73
            public static string EDITOR_ZOOM = "MISCELLANEOUS_EDITOR_ZOOM";

            // Token: 0x0400004A RID: 74
            public static string SET_HITSOUND_VOLUME = "MISCELLANEOUS_SET_HITSOUND_VOLUME";

            // Token: 0x0400004B RID: 75
            public static string CURRENT_HITSOUND_VOLUME = "MISCELLANEOUS_CURRENT_HITSOUND_VOLUME";

            // Token: 0x0400004C RID: 76
            public static string HITSOUND_FORCE_VOLUME = "MISCELLANEOUS_HITSOUND_FORCE_VOLUME";

            // Token: 0x0400004D RID: 77
            public static string HITSOUND_IGNORE_STARTING_VALUE = "MISCELLANEOUS_HITSOUND_IGNORE_STARTING_VALUE";
        }

        // Token: 0x0200000B RID: 11
        public class PlanetColor
        {
            // Token: 0x0400004E RID: 78
            public static string NAME = "PLANET_COLOR_NAME";

            // Token: 0x0400004F RID: 79
            public static string DESCRIPTION = "PLANET_COLOR_DESCRIPTION";

            // Token: 0x04000050 RID: 80
            public static string PLANET_ONE = "PLANET_COLOR_PLANET_ONE";

            // Token: 0x04000051 RID: 81
            public static string PLANET_TWO = "PLANET_COLOR_PLANET_TWO";

            // Token: 0x04000052 RID: 82
            public static string BODY = "PLANET_COLOR_BODY";

            // Token: 0x04000053 RID: 83
            public static string TAIL = "PLANET_COLOR_TAIL";
        }

        // Token: 0x0200000C RID: 12
        public class RestrictJudgments
        {
            // Token: 0x04000054 RID: 84
            public static string NAME = "RESTRICT_JUDGMENTS_NAME";

            // Token: 0x04000055 RID: 85
            public static string DESCRIPTION = "RESTRICT_JUDGMENTS_DESCRIPTION";

            // Token: 0x04000056 RID: 86
            public static string RESTRICT_HEADER = "RESTRICT_JUDGMENTS_RESTRICT_HEADER";

            // Token: 0x04000057 RID: 87
            public static string RESTRICT = "RESTRICT_JUDGMENTS_RESTRICT";

            // Token: 0x04000058 RID: 88
            public static string CUSTOM_HEADER = "RESTRICT_JUDGMENTS_CUSTOM_HEADER";

            // Token: 0x04000059 RID: 89
            public static string RESTRICT_ACTION = "RESTRICT_JUDGMENTS_RESTRICT_ACTION";

            // Token: 0x0400005A RID: 90
            public static string CUSTOM_DEATH = "RESTRICT_JUDGMENTS_CUSTOM_DEATH";
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
            string dbPath = Path.Combine("Mods", "KeyViewer", "TweakStrings.db");
            using (var db = new LiteDatabase(Assembly.GetExecutingAssembly().GetManifestResourceStream("KeyViewer.TweakStrings.db")))
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
        /// The sprite for the bottom arrow on the hit error meter.
        /// </summary>
        public static Sprite HandSprite { get; private set; }

        /// <summary>
        /// The sprite for the colored part of the hit error meter.
        /// </summary>
        public static Sprite MeterSprite { get; private set; }

        /// <summary>
        /// The sprite for the colored ticks on the hit error meter.
        /// </summary>
        public static Sprite TickSprite { get; private set; }

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
            HandSprite = assets.LoadAsset<Sprite>("Assets/Hand.png");
            MeterSprite = assets.LoadAsset<Sprite>("Assets/Meter.png");
            TickSprite = assets.LoadAsset<Sprite>("Assets/Tick.png");
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
            return AssetBundle.LoadFromMemory(Assembly.GetExecutingAssembly().GetManifestResourceStream("KeyViewer.KeyViewer.assets").ReadFully());
        }

    }
    public static class Main
    {
        [ModuleInitializer]
        public static void Initialize()
        {
            AppDomain.CurrentDomain.AssemblyResolve += (sender, args) =>
            {
                using (Stream s = Assembly.GetExecutingAssembly().GetManifestResourceStream("KeyViewer.LiteDB.dll"))
                {
                    byte[] buffer = new byte[s.Length];
                    s.Read(buffer, 0, buffer.Length);
                    return Assembly.Load(buffer);
                }
            };
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
            KeyViewerTweak.changedk = !Settings.DisplayKPS;
            KeyViewerTweak.changedt = !Settings.DisplayTotal;
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
                TweakStrings.Get(TranslationKeys.Global.GLOBAL_LANGUAGE),
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
                    TweakStrings.GetForLanguage(TranslationKeys.Global.LANGUAGE_NAME, language);

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
                if (Settings.SaveKeyCounts && File.Exists(Path.Combine("Mods", "KeyViewer", "KeyCounts.kc")))
                {
                    KeyViewer.keyCounts = Path.Combine("Mods", "KeyViewer", "KeyCounts.kc").DeserializeJson<Dictionary<KeyCode, int>>();
                }
                else Settings.Total = 0;
                harmony = new Harmony(modEntry.Info.Id);
                harmony.PatchAll(Assembly.GetExecutingAssembly());
                KeyViewerTweak.OnEnable();
            }
            else
            {
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

using System.Collections.Generic;
using UnityEngine;
using UnityEngine.UI;
using UnityEngine.EventSystems;
using DG.Tweening;
using System;

namespace KeyViewer
{
    [Obsolete("려ck")]
    public enum KeyType
    {
        Key,
        KPS,
        Total,
        Fuck,
        려차
    }
    public struct Point : IEquatable<Point>
    {
        public bool Equals(Point point) => this == point;
        public override bool Equals(object obj) => this == (Point)obj;
        public override int GetHashCode() => base.GetHashCode();
        public static implicit operator Point(int[] point) => new Point(point);
        public static implicit operator int[](Point point) => new int[] { (int)point.x, (int)point.y };
        public static implicit operator Point(float[] point) => new Point(point);
        public static implicit operator float[](Point point) => new float[] { point.x, point.y };
        public static bool operator !=(Point lhs, Point rhs) => !(lhs == rhs);
        public static bool operator ==(Point lhs, Point rhs) => lhs.x == rhs.x && lhs.y == rhs.y;
        public static implicit operator Point(int value) => new int[] {value, value};
        public static implicit operator Point(float value) => new float[] { value, value };
        public static bool operator !(Point p) => p.x == 0 && p.y == 0;
        public static Point operator +(Point lhs, Point rhs)
        {
            float x = lhs.x + rhs.x;
            float y = lhs.y + rhs.y;
            return new Point(new float[] { x, y });
        }
        public static Point operator -(Point lhs, Point rhs)
        {
            float x = lhs.x - rhs.x;
            float y = lhs.y - rhs.y;
            return new Point(new float[] { x, y });
        }
        public Point(int[] point)
        {
            x = point[0];
            y = point[1];
        }
        public Point(float[] point)
        {
            x = point[0];
            y = point[1];
        }
        public float x;
        public float y;
        public static readonly Point Zero = new Point(new float[] {0, 0});
    }
    public struct PoSize
    {
        public PoSize(Point[] points)
        {
            Pos = points[0];
            Size = points[1];
        }
        public Point Pos;
        public Point Size;

        public static implicit operator PoSize((Point Pos, Point Size) PoSize) =>
            new PoSize(new Point[] {PoSize.Pos, PoSize.Size});

        public static readonly PoSize Zero = new PoSize(new Point[]{Point.Zero, Point.Zero});
    }
    [Obsolete("Fu차")]
    public class KeyManager
    {
        public const string 려차 = "Fuck";
    }
}

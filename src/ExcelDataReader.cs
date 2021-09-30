using System;
using ExcelDataReader.Core;
using System.IO;
using System.Text;
using ExcelDataReader.Core.BinaryFormat;
using ExcelDataReader.Core.CsvFormat;
using System.Collections.Generic;
using System.Data;
using ExcelDataReader.Core.OpenXmlFormat;
using ExcelDataReader.Core.CompoundFormat;
using ExcelDataReader.Core.OfficeCrypto;
using ExcelDataReader.Exceptions;
using ExcelDataReader.Misc;
using ExcelDataReader.Core.NumberFormat;
using System.Globalization;
using System.Text.RegularExpressions;
using System.IO.Compression;
using System.Xml;
using System.Runtime.Serialization;
using ExcelDataReader.Log.Logger;
using System.Collections.ObjectModel;
using System.Security.Cryptography;
using ExcelDataReader.Log;
using ExcelDataReader.Core.OpenXmlFormat.Records;
using System.Collections;
using ExcelDataReader.Core.OpenXmlFormat.BinaryFormat;
using ExcelDataReader.Core.OpenXmlFormat.XmlFormat;

namespace ExcelDataReader
{
    /// <summary>
    /// Formula error
    /// </summary>
    public enum CellError : byte
    {
        /// <summary>
        /// #NULL!
        /// </summary>
        NULL = 0x00,

        /// <summary>
        /// #DIV/0!
        /// </summary>
        DIV0 = 0x07,

        /// <summary>
        /// #VALUE!
        /// </summary>
        VALUE = 0x0F,

        /// <summary>
        /// #REF!
        /// </summary>
        REF = 0x17,

        /// <summary>
        /// #NAME?
        /// </summary>
        NAME = 0x1D,

        /// <summary>
        /// #NUM!
        /// </summary>
        NUM = 0x24,

        /// <summary>
        /// #N/A
        /// </summary>
        NA = 0x2A,

        /// <summary>
        /// #GETTING_DATA
        /// </summary>
        GETTING_DATA = 0x2B,
    }
}


namespace ExcelDataReader
{
    /// <summary>
    /// A range for cells using 0 index positions. 
    /// </summary>
    public sealed class CellRange
    {
        internal CellRange(string range)
        {
            var fromTo = range.Split(':');
            if (fromTo.Length == 2)
            {
                ReferenceHelper.ParseReference(fromTo[0], out int column, out int row);
                
                // 0 indexed vs 1 indexed
                FromColumn = column - 1;
                FromRow = row - 1;

                ReferenceHelper.ParseReference(fromTo[1], out column, out row);

                // 0 indexed vs 1 indexed
                ToColumn = column - 1;
                ToRow = row - 1;
            }
        }

        internal CellRange(int fromColumn, int fromRow, int toColumn, int toRow)
        {
            FromColumn = fromColumn;
            FromRow = fromRow;
            ToColumn = toColumn;
            ToRow = toRow;
        }

        /// <summary>
        /// Gets the column the range starts in
        /// </summary>
        public int FromColumn { get; }

        /// <summary>
        /// Gets the row the range starts in
        /// </summary>
        public int FromRow { get; }

        /// <summary>
        /// Gets the column the range ends in
        /// </summary>
        public int ToColumn { get; }

        /// <summary>
        /// Gets the row the range ends in
        /// </summary>
        public int ToRow { get; }

        /// <inheritsdoc/>
        public override string ToString() => $"{FromRow}, {ToRow}, {FromColumn}, {ToColumn}";
    }
}

namespace ExcelDataReader
{
    /// <summary>
    /// Horizontal alignment. 
    /// </summary>
    public enum HorizontalAlignment
    {
        /// <summary>
        /// General.
        /// </summary>
        General,

        /// <summary>
        /// Left.
        /// </summary>
        Left,

        /// <summary>
        /// Centered.
        /// </summary>
        Centered,

        /// <summary>
        /// Right.
        /// </summary>
        Right,

        /// <summary>
        /// Filled.
        /// </summary>
        Filled,

        /// <summary>
        /// Justified.
        /// </summary>
        Justified,

        /// <summary>
        /// Centered across selection.
        /// </summary>
        CenteredAcrossSelection,

        /// <summary>
        /// Distributed.
        /// </summary>
        Distributed,
    }

    /// <summary>
    /// Holds style information for a cell.
    /// </summary>
    public class CellStyle
    {
        /// <summary>
        /// Gets the font index.
        /// </summary>
        public int FontIndex { get; internal set; }

        /// <summary>
        /// Gets the number format index.
        /// </summary>
        public int NumberFormatIndex { get; internal set; }

        /// <summary>
        /// Gets the indent level.
        /// </summary>
        public int IndentLevel { get; internal set; }

        /// <summary>
        /// Gets the horizontal alignment.
        /// </summary>
        public HorizontalAlignment HorizontalAlignment { get; internal set; }

        /// <summary>
        /// Gets a value indicating whether the cell is hidden.
        /// </summary>
        public bool Hidden { get; internal set; }

        /// <summary>
        /// Gets a value indicating whether the cell is locked.
        /// </summary>
        public bool Locked { get; internal set; }
    }
}
namespace ExcelDataReader
{
    internal static class Errors
    {
        public const string ErrorStreamWorkbookNotFound = "Neither stream 'Workbook' nor 'Book' was found in file.";
        public const string ErrorWorkbookIsNotStream = "Workbook directory entry is not a Stream.";
        public const string ErrorWorkbookGlobalsInvalidData = "Error reading Workbook Globals - Stream has invalid data.";
        public const string ErrorFatBadSector = "Error reading as FAT table : There's no such sector in FAT.";
        public const string ErrorFatRead = "Error reading stream from FAT area.";
        public const string ErrorEndOfFile = "The excel file may be corrupt or truncated. We've read past the end of the file.";
        public const string ErrorCyclicSectorChain = "Cyclic sector chain in compound document.";
        public const string ErrorHeaderSignature = "Invalid file signature.";
        public const string ErrorHeaderOrder = "Invalid byte order specified in header.";
        public const string ErrorBiffRecordSize = "Buffer size is less than minimum BIFF record size.";
        public const string ErrorBiffIlegalBefore = "BIFF Stream error: Moving before stream start.";
        public const string ErrorBiffIlegalAfter = "BIFF Stream error: Moving after stream end.";

        public const string ErrorDirectoryEntryArray = "Directory Entry error: Array is too small.";
        public const string ErrorCompoundNoOpenXml = "Detected compound document, but not a valid OpenXml file.";
        public const string ErrorZipNoOpenXml = "Detected ZIP file, but not a valid OpenXml file.";
        public const string ErrorInvalidPassword = "Invalid password.";
    }
}


namespace ExcelDataReader
{
    /// <summary>
    /// ExcelDataReader Class
    /// </summary>
    internal class ExcelBinaryReader : ExcelDataReader<XlsWorkbook, XlsWorksheet>
    {
        public ExcelBinaryReader(Stream stream, string password, Encoding fallbackEncoding)
        {
            Workbook = new XlsWorkbook(stream, password, fallbackEncoding);

            // By default, the data reader is positioned on the first result.
            Reset();
        }

        public override void Close()
        {
            base.Close();
            Workbook?.Stream?.Dispose();
            Workbook = null;
        }
    }
}


namespace ExcelDataReader
{
    internal class ExcelCsvReader : ExcelDataReader<CsvWorkbook, CsvWorksheet>
    {
        public ExcelCsvReader(Stream stream, Encoding fallbackEncoding, char[] autodetectSeparators, int analyzeInitialCsvRows)
        {
            Workbook = new CsvWorkbook(stream, fallbackEncoding, autodetectSeparators, analyzeInitialCsvRows);

            // By default, the data reader is positioned on the first result.
            Reset();
        }

        public override void Close()
        {
            base.Close();
            Workbook?.Stream?.Dispose();
            Workbook = null;
        }
    }
}


namespace ExcelDataReader
{
    /// <summary>
    /// A generic implementation of the IExcelDataReader interface using IWorkbook/IWorksheet to enumerate data.
    /// </summary>
    /// <typeparam name="TWorkbook">A type implementing IWorkbook</typeparam>
    /// <typeparam name="TWorksheet">A type implementing IWorksheet</typeparam>
    internal abstract class ExcelDataReader<TWorkbook, TWorksheet> : IExcelDataReader
        where TWorkbook : IWorkbook<TWorksheet>
        where TWorksheet : IWorksheet
    {
        private IEnumerator<TWorksheet> _worksheetIterator;
        private IEnumerator<Row> _rowIterator;
        private IEnumerator<TWorksheet> _cachedWorksheetIterator;
        private List<TWorksheet> _cachedWorksheets;

        ~ExcelDataReader()
        {
            Dispose(false);
        }

        public string Name => _worksheetIterator?.Current?.Name;

        public string CodeName => _worksheetIterator?.Current?.CodeName;

        public string VisibleState => _worksheetIterator?.Current?.VisibleState;

        public HeaderFooter HeaderFooter => _worksheetIterator?.Current?.HeaderFooter;

        // We shouldn't expose the internal array here. 
        public CellRange[] MergeCells => _worksheetIterator?.Current?.MergeCells;

        public int Depth { get; private set; }

        public int ResultsCount => Workbook?.ResultsCount ?? -1;

        public bool IsClosed { get; private set; }

        public int FieldCount => _worksheetIterator?.Current?.FieldCount ?? 0;

        public int RowCount => _worksheetIterator?.Current?.RowCount ?? 0;

        public int RecordsAffected => throw new NotSupportedException();

        public double RowHeight => _rowIterator?.Current.Height ?? 0;

        protected TWorkbook Workbook { get; set; }

        protected Cell[] RowCells { get; set; }

        public object this[int i] => GetValue(i);

        public object this[string name] => throw new NotSupportedException();

        public bool GetBoolean(int i) => (bool)GetValue(i);

        public byte GetByte(int i) => (byte)GetValue(i);

        public long GetBytes(int i, long fieldOffset, byte[] buffer, int bufferoffset, int length)
            => throw new NotSupportedException();

        public char GetChar(int i) => (char)GetValue(i);

        public long GetChars(int i, long fieldoffset, char[] buffer, int bufferoffset, int length)
             => throw new NotSupportedException();

        public IDataReader GetData(int i) => throw new NotSupportedException();

        public string GetDataTypeName(int i) => throw new NotSupportedException();

        public DateTime GetDateTime(int i) => (DateTime)GetValue(i);

        public decimal GetDecimal(int i) => (decimal)GetValue(i);

        public double GetDouble(int i) => (double)GetValue(i);

        public Type GetFieldType(int i) => GetValue(i)?.GetType();

        public float GetFloat(int i) => (float)GetValue(i);

        public Guid GetGuid(int i) => (Guid)GetValue(i);

        public short GetInt16(int i) => (short)GetValue(i);

        public int GetInt32(int i) => (int)GetValue(i);

        public long GetInt64(int i) => (long)GetValue(i);

        public string GetName(int i) => throw new NotSupportedException();

        public int GetOrdinal(string name) => throw new NotSupportedException();

        /// <inheritdoc />
        public DataTable GetSchemaTable() => throw new NotSupportedException();

        public string GetString(int i) => (string)GetValue(i);

        public object GetValue(int i)
        {
            if (RowCells == null)
                throw new InvalidOperationException("No data exists for the row/column.");
            
            return RowCells[i]?.Value;
        }

        public int GetValues(object[] values) => throw new NotSupportedException();
               
        public bool IsDBNull(int i) => GetValue(i) == null;

        public string GetNumberFormatString(int i)
        {
            if (RowCells == null)
                throw new InvalidOperationException("No data exists for the row/column.");
            if (RowCells[i] == null)
                return null;
            if (RowCells[i].EffectiveStyle == null)
                return null;
            return Workbook.GetNumberFormatString(RowCells[i].EffectiveStyle.NumberFormatIndex)?.FormatString;
        }

        public int GetNumberFormatIndex(int i)
        {
            if (RowCells == null)
                throw new InvalidOperationException("No data exists for the row/column.");
            if (RowCells[i] == null)
                return -1;
            if (RowCells[i].EffectiveStyle == null)
                return -1;
            return RowCells[i].EffectiveStyle.NumberFormatIndex;
        }

        public double GetColumnWidth(int i)
        {
            if (i >= FieldCount)
            {
                throw new ArgumentException($"Column at index {i} does not exist.", nameof(i));
            }

            var columnWidths = _worksheetIterator?.Current?.ColumnWidths ?? null;
            double? retWidth = null;
            if (columnWidths != null)
            {
                foreach (var columnWidth in columnWidths)
                {
                    if (i >= columnWidth.Minimum && i <= columnWidth.Maximum)
                    {
                        retWidth = columnWidth.Hidden ? 0 : columnWidth.Width;
                        break;
                    }
                }
            }

            const double DefaultColumnWidth = 8.43D;

            return retWidth ?? DefaultColumnWidth;
        }

        public CellStyle GetCellStyle(int i)
        {
            if (RowCells == null)
                throw new InvalidOperationException("No data exists for the row/column.");

            var result = new CellStyle();
            if (RowCells[i] == null)
            {
                return result;
            }

            var effectiveStyle = RowCells[i].EffectiveStyle;
            if (effectiveStyle == null)
            {
                return result;
            }

            result.FontIndex = effectiveStyle.FontIndex;
            result.NumberFormatIndex = effectiveStyle.NumberFormatIndex;
            result.IndentLevel = effectiveStyle.IndentLevel;
            result.HorizontalAlignment = effectiveStyle.HorizontalAlignment;
            result.Hidden = effectiveStyle.Hidden;
            result.Locked = effectiveStyle.Locked;
            return result;
        }

        public CellError? GetCellError(int i)
        {
            if (RowCells == null)
                throw new InvalidOperationException("No data exists for the row/column.");
            
            return RowCells[i]?.Error;
        }

        /// <inheritdoc />
        public void Reset()
        {
            _worksheetIterator?.Dispose();
            _rowIterator?.Dispose();

            _worksheetIterator = null;
            _rowIterator = null;

            ResetSheetData();

            if (Workbook != null)
            {
                _worksheetIterator = ReadWorksheetsWithCache().GetEnumerator(); // Workbook.ReadWorksheets().GetEnumerator();
                if (!_worksheetIterator.MoveNext())
                {
                    _worksheetIterator.Dispose();
                    _worksheetIterator = null;
                    return;
                }

                _rowIterator = _worksheetIterator.Current.ReadRows().GetEnumerator();
            }
        }

        public virtual void Close()
        {
            if (IsClosed)
                return;

            _worksheetIterator?.Dispose();
            _rowIterator?.Dispose();

            _worksheetIterator = null;
            _rowIterator = null;
            RowCells = null;
            IsClosed = true;
        }

        public bool NextResult()
        {
            if (_worksheetIterator == null)
            {
                return false;
            }

            ResetSheetData();

            _rowIterator?.Dispose();
            _rowIterator = null;

            if (!_worksheetIterator.MoveNext())
            {
                _worksheetIterator.Dispose();
                _worksheetIterator = null;
                return false;
            }

            _rowIterator = _worksheetIterator.Current.ReadRows().GetEnumerator();
            return true;
        }

        public bool Read()
        {
            if (_worksheetIterator == null || _rowIterator == null)
            {
                return false;
            }

            if (!_rowIterator.MoveNext())
            {
                _rowIterator.Dispose();
                _rowIterator = null;
                return false;
            }

            ReadCurrentRow();

            Depth++;
            return true;
        }

        public void Dispose()
        {
            Dispose(true);

            GC.SuppressFinalize(this);
        }

        private void Dispose(bool disposing)
        {
            if (disposing)
                Close();
        }

        private IEnumerable<TWorksheet> ReadWorksheetsWithCache()
        {
            // Iterate TWorkbook.ReadWorksheets() only once and cache the 
            // worksheet instances, which are expensive to create. 
            if (_cachedWorksheets != null)
            {
                foreach (var worksheet in _cachedWorksheets)
                {
                    yield return worksheet;
                }

                if (_cachedWorksheetIterator == null)
                {
                    yield break;
                }
            }
            else
            {
                _cachedWorksheets = new List<TWorksheet>();
            }

            if (_cachedWorksheetIterator == null)
            {
                _cachedWorksheetIterator = Workbook.ReadWorksheets().GetEnumerator();
            }

            while (_cachedWorksheetIterator.MoveNext())
            {
                _cachedWorksheets.Add(_cachedWorksheetIterator.Current);
                yield return _cachedWorksheetIterator.Current;
            }

            _cachedWorksheetIterator.Dispose();
            _cachedWorksheetIterator = null;
        }

        private void ResetSheetData()
        {
            Depth = -1;
            RowCells = null;
        }

        private void ReadCurrentRow()
        {
            if (RowCells == null)
            {
                RowCells = new Cell[FieldCount];
            }

            Array.Clear(RowCells, 0, RowCells.Length);

            foreach (var cell in _rowIterator.Current.Cells)
            {
                if (cell.ColumnIndex < RowCells.Length)
                {
                    RowCells[cell.ColumnIndex] = cell;
                }
            }
        }
    }
}


namespace ExcelDataReader
{
    internal class ExcelOpenXmlReader : ExcelDataReader<XlsxWorkbook, XlsxWorksheet>
    {
        public ExcelOpenXmlReader(Stream stream)
        {
            Document = new ZipWorker(stream);
            Workbook = new XlsxWorkbook(Document);

            // By default, the data reader is positioned on the first result.
            Reset();
        }

        private ZipWorker Document { get; set; }

        public override void Close()
        {
            base.Close();

            Document?.Dispose();
            Workbook = null;
            Document = null;
        }
    }
}


namespace ExcelDataReader
{
    /// <summary>
    /// Configuration options for an instance of ExcelDataReader.
    /// </summary>
    public class ExcelReaderConfiguration
    {
        /// <summary>
        /// Gets or sets a value indicating the encoding to use when the input XLS lacks a CodePage record, 
        /// or when the input CSV lacks a BOM and does not parse as UTF8. Default: cp1252. (XLS BIFF2-5 and CSV only)
        /// </summary>
        public Encoding FallbackEncoding { get; set; } = Encoding.GetEncoding(1252);

        /// <summary>
        /// Gets or sets the password used to open password protected workbooks.
        /// </summary>
        public string Password { get; set; }

        /// <summary>
        /// Gets or sets an array of CSV separator candidates. The reader autodetects which best fits the input data. Default: , ; TAB | # (CSV only)
        /// </summary>
        public char[] AutodetectSeparators { get; set; } = new char[] { ',', ';', '\t', '|', '#' };

        /// <summary>
        /// Gets or sets a value indicating whether to leave the stream open after the IExcelDataReader object is disposed. Default: false
        /// </summary>
        public bool LeaveOpen { get; set; }

        /// <summary>
        /// Gets or sets a value indicating the number of rows to analyze for encoding, separator and field count in a CSV.
        /// When set, this option causes the IExcelDataReader.RowCount property to throw an exception.
        /// Default: 0 - analyzes the entire file (CSV only, has no effect on other formats)
        /// </summary>
        public int AnalyzeInitialCsvRows { get; set; }
    }
}


namespace ExcelDataReader
{
    /// <summary>
    /// The ExcelReader Factory
    /// </summary>
    public static class ExcelReaderFactory
    {
        private const string DirectoryEntryWorkbook = "Workbook";
        private const string DirectoryEntryBook = "Book";
        private const string DirectoryEntryEncryptedPackage = "EncryptedPackage";
        private const string DirectoryEntryEncryptionInfo = "EncryptionInfo";

        /// <summary>
        /// Creates an instance of <see cref="ExcelBinaryReader"/> or <see cref="ExcelOpenXmlReader"/>
        /// </summary>
        /// <param name="fileStream">The file stream.</param>
        /// <param name="configuration">The configuration object.</param>
        /// <returns>The excel data reader.</returns>
        public static IExcelDataReader CreateReader(Stream fileStream, ExcelReaderConfiguration configuration = null)
        {
            if (configuration == null)
            {
                configuration = new ExcelReaderConfiguration();
            }

            if (configuration.LeaveOpen)
            {
                fileStream = new LeaveOpenStream(fileStream);
            }

            var probe = new byte[8];
            fileStream.Seek(0, SeekOrigin.Begin);
            fileStream.Read(probe, 0, probe.Length);
            fileStream.Seek(0, SeekOrigin.Begin);

            if (CompoundDocument.IsCompoundDocument(probe))
            {
                // Can be BIFF5-8 or password protected OpenXml
                var document = new CompoundDocument(fileStream);
                if (TryGetWorkbook(fileStream, document, out var stream))
                {
                    return new ExcelBinaryReader(stream, configuration.Password, configuration.FallbackEncoding);
                }

                if (TryGetEncryptedPackage(fileStream, document, configuration.Password, out stream))
                {
                    return new ExcelOpenXmlReader(stream);
                }

                throw new ExcelReaderException(Errors.ErrorStreamWorkbookNotFound);
            }

            if (XlsWorkbook.IsRawBiffStream(probe))
            {
                return new ExcelBinaryReader(fileStream, configuration.Password, configuration.FallbackEncoding);
            }

            if (probe[0] == 0x50 && probe[1] == 0x4B)
            {
                // zip files start with 'PK'
                return new ExcelOpenXmlReader(fileStream);
            }

            throw new HeaderException(Errors.ErrorHeaderSignature);
        }

        /// <summary>
        /// Creates an instance of <see cref="ExcelBinaryReader"/>
        /// </summary>
        /// <param name="fileStream">The file stream.</param>
        /// <param name="configuration">The configuration object.</param>
        /// <returns>The excel data reader.</returns>
        public static IExcelDataReader CreateBinaryReader(Stream fileStream, ExcelReaderConfiguration configuration = null)
        {
            if (configuration == null)
            {
                configuration = new ExcelReaderConfiguration();
            }

            if (configuration.LeaveOpen)
            {
                fileStream = new LeaveOpenStream(fileStream);
            }

            var probe = new byte[8];
            fileStream.Seek(0, SeekOrigin.Begin);
            fileStream.Read(probe, 0, probe.Length);
            fileStream.Seek(0, SeekOrigin.Begin);

            if (CompoundDocument.IsCompoundDocument(probe))
            {
                var document = new CompoundDocument(fileStream);
                if (TryGetWorkbook(fileStream, document, out var stream))
                {
                    return new ExcelBinaryReader(stream, configuration.Password, configuration.FallbackEncoding);
                }
                else
                {
                    throw new ExcelReaderException(Errors.ErrorStreamWorkbookNotFound);
                }
            }
            else if (XlsWorkbook.IsRawBiffStream(probe))
            {
                return new ExcelBinaryReader(fileStream, configuration.Password, configuration.FallbackEncoding);
            }
            else
            {
                throw new HeaderException(Errors.ErrorHeaderSignature);
            }
        }

        /// <summary>
        /// Creates an instance of <see cref="ExcelOpenXmlReader"/>
        /// </summary>
        /// <param name="fileStream">The file stream.</param>
        /// <param name="configuration">The reader configuration -or- <see langword="null"/> to use the default configuration.</param>
        /// <returns>The excel data reader.</returns>
        public static IExcelDataReader CreateOpenXmlReader(Stream fileStream, ExcelReaderConfiguration configuration = null)
        {
            if (configuration == null)
            {
                configuration = new ExcelReaderConfiguration();
            }

            if (configuration.LeaveOpen)
            {
                fileStream = new LeaveOpenStream(fileStream);
            }

            var probe = new byte[8];
            fileStream.Seek(0, SeekOrigin.Begin);
            fileStream.Read(probe, 0, probe.Length);
            fileStream.Seek(0, SeekOrigin.Begin);

            // Probe for password protected compound document or zip file
            if (CompoundDocument.IsCompoundDocument(probe))
            {
                var document = new CompoundDocument(fileStream);
                if (TryGetEncryptedPackage(fileStream, document, configuration.Password, out var stream))
                {
                    return new ExcelOpenXmlReader(stream);
                }

                throw new ExcelReaderException(Errors.ErrorCompoundNoOpenXml);
            }

            if (probe[0] == 0x50 && probe[1] == 0x4B)
            {
                // Zip files start with 'PK'
                return new ExcelOpenXmlReader(fileStream);
            }

            throw new HeaderException(Errors.ErrorHeaderSignature);
        }

        /// <summary>
        /// Creates an instance of ExcelCsvReader
        /// </summary>
        /// <param name="fileStream">The file stream.</param>
        /// <param name="configuration">The reader configuration -or- <see langword="null"/> to use the default configuration.</param>
        /// <returns>The excel data reader.</returns>
        public static IExcelDataReader CreateCsvReader(Stream fileStream, ExcelReaderConfiguration configuration = null)
        {
            if (configuration == null)
            {
                configuration = new ExcelReaderConfiguration();
            }

            if (configuration.LeaveOpen)
            {
                fileStream = new LeaveOpenStream(fileStream);
            }

            return new ExcelCsvReader(fileStream, configuration.FallbackEncoding, configuration.AutodetectSeparators, configuration.AnalyzeInitialCsvRows);
        }

        private static bool TryGetWorkbook(Stream fileStream, CompoundDocument document, out Stream stream)
        {
            var workbookEntry = document.FindEntry(DirectoryEntryWorkbook, DirectoryEntryBook);
            if (workbookEntry != null)
            {
                if (workbookEntry.EntryType != STGTY.STGTY_STREAM)
                {
                    throw new ExcelReaderException(Errors.ErrorWorkbookIsNotStream);
                }

                stream = new CompoundStream(document, fileStream, workbookEntry.StreamFirstSector, (int)workbookEntry.StreamSize, workbookEntry.IsEntryMiniStream, false);
                return true;
            }

            stream = null;
            return false;
        }

        private static bool TryGetEncryptedPackage(Stream fileStream, CompoundDocument document, string password, out Stream stream)
        {
            var encryptedPackage = document.FindEntry(DirectoryEntryEncryptedPackage);
            var encryptionInfo = document.FindEntry(DirectoryEntryEncryptionInfo);

            if (encryptedPackage == null || encryptionInfo == null)
            {
                stream = null;
                return false;
            }

            var infoBytes = document.ReadStream(fileStream, encryptionInfo.StreamFirstSector, (int)encryptionInfo.StreamSize, encryptionInfo.IsEntryMiniStream);
            var encryption = EncryptionInfo.Create(infoBytes);

            if (encryption.VerifyPassword("VelvetSweatshop"))
            {
                // Magic password used for write-protected workbooks
                password = "VelvetSweatshop";
            }
            else if (password == null || !encryption.VerifyPassword(password))
            {
                throw new InvalidPasswordException(Errors.ErrorInvalidPassword);
            }

            var secretKey = encryption.GenerateSecretKey(password);
            var packageStream = new CompoundStream(document, fileStream, encryptedPackage.StreamFirstSector, (int)encryptedPackage.StreamSize, encryptedPackage.IsEntryMiniStream, false);

            stream = encryption.CreateEncryptedPackageStream(packageStream, secretKey);
            return true;
        }
    }
}


#nullable enable

namespace ExcelDataReader
{
    /// <summary>
    /// Header and footer text. 
    /// </summary>
    public sealed class HeaderFooter
    {
        internal HeaderFooter(bool hasDifferentFirst, bool hasDifferentOddEven)
        {
            HasDifferentFirst = hasDifferentFirst;
            HasDifferentOddEven = hasDifferentOddEven;
        }

        internal HeaderFooter(string? footer, string? header)
            : this(false, false)
        {
            OddHeader = header;
            OddFooter = footer;
        }

        /// <summary>
        /// Gets a value indicating whether the header and footer are different on the first page. 
        /// </summary>
        public bool HasDifferentFirst { get; }

        /// <summary>
        /// Gets a value indicating whether the header and footer are different on odd and even pages.
        /// </summary>
        public bool HasDifferentOddEven { get; }

        /// <summary>
        /// Gets the header used for the first page if <see cref="HasDifferentFirst"/> is <see langword="true"/>.
        /// </summary>
        public string? FirstHeader { get; internal set; }

        /// <summary>
        /// Gets the footer used for the first page if <see cref="HasDifferentFirst"/> is <see langword="true"/>.
        /// </summary>
        public string? FirstFooter { get; internal set; }

        /// <summary>
        /// Gets the header used for odd pages -or- all pages if <see cref="HasDifferentOddEven"/> is <see langword="false"/>. 
        /// </summary>
        public string? OddHeader { get; internal set; }

        /// <summary>
        /// Gets the footer used for odd pages -or- all pages if <see cref="HasDifferentOddEven"/> is <see langword="false"/>. 
        /// </summary>
        public string? OddFooter { get; internal set; }

        /// <summary>
        /// Gets the header used for even pages if <see cref="HasDifferentOddEven"/> is <see langword="true"/>. 
        /// </summary>
        public string? EvenHeader { get; internal set; }

        /// <summary>
        /// Gets the footer used for even pages if <see cref="HasDifferentOddEven"/> is <see langword="true"/>. 
        /// </summary>
        public string? EvenFooter { get; internal set; }
    }
}

namespace ExcelDataReader
{
    /// <summary>
    /// The ExcelDataReader interface
    /// </summary>
    public interface IExcelDataReader : IDataReader
    {
        /// <summary>
        /// Gets the sheet name.
        /// </summary>
        string Name { get; }

        /// <summary>
        /// Gets the sheet VBA code name.
        /// </summary>
        string CodeName { get; }

        /// <summary>
        /// Gets the sheet visible state.
        /// </summary>
        string VisibleState { get; }

        /// <summary>
        /// Gets the sheet header and footer -or- <see langword="null"/> if none set.
        /// </summary>
        HeaderFooter HeaderFooter { get; }

        /// <summary>
        /// Gets the list of merged cell ranges.
        /// </summary>
        CellRange[] MergeCells { get; }

        /// <summary>
        /// Gets the number of results (workbooks).
        /// </summary>
        int ResultsCount { get; }

        /// <summary>
        /// Gets the number of rows in the current result.
        /// </summary>
        int RowCount { get; }

        /// <summary>
        /// Gets the height of the current row in points.
        /// </summary>
        double RowHeight { get; }

        /// <summary>
        /// Seeks to the first result.
        /// </summary>
        void Reset();

        /// <summary>
        /// Gets the number format for the specified field -or- <see langword="null"/> if there is no value.
        /// </summary>
        /// <param name="i">The index of the field to find.</param>
        /// <returns>The number format string of the specified field.</returns>
        string GetNumberFormatString(int i);

        /// <summary>
        /// Gets the number format index for the specified field -or- -1 if there is no value.
        /// </summary>
        /// <param name="i">The index of the field to find.</param>
        /// <returns>The number format index of the specified field.</returns>
        int GetNumberFormatIndex(int i);

        /// <summary>
        /// Gets the width the specified column.
        /// </summary>
        /// <param name="i">The index of the column to find.</param>
        /// <returns>The width of the specified column.</returns>
        double GetColumnWidth(int i);

        /// <summary>
        /// Gets the cell style.
        /// </summary>
        /// <param name="i">The index of the column to find.</param>
        /// <returns>The cell style.</returns>
        CellStyle GetCellStyle(int i);

        /// <summary>
        /// Gets the cell error.
        /// </summary>
        /// <param name="i">The index of the column to find.</param>
        /// <returns>The cell error, or null if no error.</returns>
        CellError? GetCellError(int i);
    }
}

namespace ExcelDataReader
{
    /// <summary>
    /// ExcelDataReader DataSet extensions
    /// </summary>
    public static class ExcelDataReaderExtensions
    {
        /// <summary>
        /// Converts all sheets to a DataSet
        /// </summary>
        /// <param name="self">The IExcelDataReader instance</param>
        /// <param name="configuration">An optional configuration object to modify the behavior of the conversion</param>
        /// <returns>A dataset with all workbook contents</returns>
        public static DataSet AsDataSet(this IExcelDataReader self, ExcelDataSetConfiguration configuration = null)
        {
            if (configuration == null)
            {
                configuration = new ExcelDataSetConfiguration();
            }

            self.Reset();

            var tableIndex = -1;
            var result = new DataSet();
            do
            {
                tableIndex++;
                if (configuration.FilterSheet != null && !configuration.FilterSheet(self, tableIndex))
                {
                    continue;
                }

                var tableConfiguration = configuration.ConfigureDataTable != null
                    ? configuration.ConfigureDataTable(self)
                    : null;

                if (tableConfiguration == null)
                {
                    tableConfiguration = new ExcelDataTableConfiguration();
                }

                var table = AsDataTable(self, tableConfiguration);
                result.Tables.Add(table);
            }
            while (self.NextResult());

            result.AcceptChanges();

            if (configuration.UseColumnDataType)
            {
                FixDataTypes(result);
            }

            self.Reset();

            return result;
        }

        private static string GetUniqueColumnName(DataTable table, string name)
        {
            var columnName = name;
            var i = 1;
            while (table.Columns[columnName] != null)
            {
                columnName = string.Format("{0}_{1}", name, i);
                i++;
            }

            return columnName;
        }

        private static DataTable AsDataTable(IExcelDataReader self, ExcelDataTableConfiguration configuration)
        {
            var result = new DataTable { TableName = self.Name };
            result.ExtendedProperties.Add("visiblestate", self.VisibleState);
            var first = true;
            var emptyRows = 0;
            var columnIndices = new List<int>();
            while (self.Read())
            {
                if (first)
                {
                    if (configuration.UseHeaderRow && configuration.ReadHeaderRow != null)
                    {
                        configuration.ReadHeaderRow(self);
                    }

                    for (var i = 0; i < self.FieldCount; i++)
                    {
                        if (configuration.FilterColumn != null && !configuration.FilterColumn(self, i))
                        {
                            continue;
                        }

                        var name = configuration.UseHeaderRow
                            ? Convert.ToString(self.GetValue(i))
                            : null;

                        if (string.IsNullOrEmpty(name))
                        {
                            name = configuration.EmptyColumnNamePrefix + i;
                        }

                        // if a column already exists with the name append _i to the duplicates
                        var columnName = GetUniqueColumnName(result, name);
                        var column = new DataColumn(columnName, typeof(object)) { Caption = name };
                        result.Columns.Add(column);
                        columnIndices.Add(i);
                    }

                    result.BeginLoadData();
                    first = false;

                    if (configuration.UseHeaderRow)
                    {
                        continue;
                    }
                }

                if (configuration.FilterRow != null && !configuration.FilterRow(self))
                {
                    continue;
                }

                if (IsEmptyRow(self, configuration))
                {
                    emptyRows++;
                    continue;
                }

                for (var i = 0; i < emptyRows; i++)
                {
                    result.Rows.Add(result.NewRow());
                }

                emptyRows = 0;

                var row = result.NewRow();

                for (var i = 0; i < columnIndices.Count; i++)
                {
                    var columnIndex = columnIndices[i];

                    var value = self.GetValue(columnIndex);
                    if (configuration.TransformValue != null)
                    {
                        var transformedValue = configuration.TransformValue(self, i, value);
                        if (transformedValue != null)
                            value = transformedValue;
                    }

                    row[i] = value;
                }

                result.Rows.Add(row);
            }

            result.EndLoadData();
            return result;
        }

        private static bool IsEmptyRow(IExcelDataReader reader, ExcelDataTableConfiguration configuration)
        {
            for (var i = 0; i < reader.FieldCount; i++)
            {
                var value = reader.GetValue(i);
                if (configuration.TransformValue != null)
                {
                    var transformedValue = configuration.TransformValue(reader, i, value);
                    if (transformedValue != null)
                        value = transformedValue;
                }

                if (value != null)
                    return false;
            }

            return true;
        }

        private static void FixDataTypes(DataSet dataset)
        {
            var tables = new List<DataTable>(dataset.Tables.Count);
            bool convert = false;
            foreach (DataTable table in dataset.Tables)
            {
                if (table.Rows.Count == 0)
                {
                    tables.Add(table);
                    continue;
                }

                DataTable newTable = null;
                for (int i = 0; i < table.Columns.Count; i++)
                {
                    Type type = null;
                    foreach (DataRow row in table.Rows)
                    {
                        if (row.IsNull(i))
                            continue;
                        var curType = row[i].GetType();
                        if (curType != type)
                        {
                            if (type == null)
                            {
                                type = curType;
                            }
                            else
                            {
                                type = null;
                                break;
                            }
                        }
                    }

                    if (type == null)
                        continue;
                    convert = true;
                    if (newTable == null)
                        newTable = table.Clone();
                    newTable.Columns[i].DataType = type;
                }

                if (newTable != null)
                {
                    newTable.BeginLoadData();
                    foreach (DataRow row in table.Rows)
                    {
                        newTable.ImportRow(row);
                    }

                    newTable.EndLoadData();
                    tables.Add(newTable);
                }
                else
                {
                    tables.Add(table);
                }
            }

            if (convert)
            {
                dataset.Tables.Clear();
                dataset.Tables.AddRange(tables.ToArray());
            }
        }
    }
}


namespace ExcelDataReader
{
    /// <summary>
    /// Processing configuration options and callbacks for IExcelDataReader.AsDataSet().
    /// </summary>
    public class ExcelDataSetConfiguration
    {
        /// <summary>
        /// Gets or sets a value indicating whether to set the DataColumn.DataType property in a second pass.
        /// </summary>
        public bool UseColumnDataType { get; set; } = true;

        /// <summary>
        /// Gets or sets a callback to obtain configuration options for a DataTable. 
        /// </summary>
        public Func<IExcelDataReader, ExcelDataTableConfiguration> ConfigureDataTable { get; set; }

        /// <summary>
        /// Gets or sets a callback to determine whether to include the current sheet in the DataSet. Called once per sheet before ConfigureDataTable.
        /// </summary>
        public Func<IExcelDataReader, int, bool> FilterSheet { get; set; }
    }
}


namespace ExcelDataReader
{
    /// <summary>
    /// Processing configuration options and callbacks for AsDataTable().
    /// </summary>
    public class ExcelDataTableConfiguration
    {
        /// <summary>
        /// Gets or sets a value indicating the prefix of generated column names.
        /// </summary>
        public string EmptyColumnNamePrefix { get; set; } = "Column";

        /// <summary>
        /// Gets or sets a value indicating whether to use a row from the data as column names.
        /// </summary>
        public bool UseHeaderRow { get; set; } = false;

        /// <summary>
        /// Gets or sets a callback to determine which row is the header row. Only called when UseHeaderRow = true.
        /// </summary>
        public Action<IExcelDataReader> ReadHeaderRow { get; set; }

        /// <summary>
        /// Gets or sets a callback to determine whether to include the current row in the DataTable.
        /// </summary>
        public Func<IExcelDataReader, bool> FilterRow { get; set; }

        /// <summary>
        /// Gets or sets a callback to determine whether to include the specific column in the DataTable. Called once per column after reading the headers.
        /// </summary>
        public Func<IExcelDataReader, int, bool> FilterColumn { get; set; }

        /// <summary>
        /// Gets or sets a callback to determine whether to transform the cell value.
        /// </summary>
        public Func<IExcelDataReader, int, object, object> TransformValue { get; set; }
    }
}


namespace ExcelDataReader.Core
{
    internal static class BuiltinNumberFormat
    {
        private static Dictionary<int, NumberFormatString> Formats { get; } = new Dictionary<int, NumberFormatString>()
        {
            { 0, new NumberFormatString("General") },
            { 1, new NumberFormatString("0") },
            { 2, new NumberFormatString("0.00") },
            { 3, new NumberFormatString("#,##0") },
            { 4, new NumberFormatString("#,##0.00") },
            { 5, new NumberFormatString("\"$\"#,##0_);(\"$\"#,##0)") },
            { 6, new NumberFormatString("\"$\"#,##0_);[Red](\"$\"#,##0)") },
            { 7, new NumberFormatString("\"$\"#,##0.00_);(\"$\"#,##0.00)") },
            { 8, new NumberFormatString("\"$\"#,##0.00_);[Red](\"$\"#,##0.00)") },
            { 9, new NumberFormatString("0%") },
            { 10, new NumberFormatString("0.00%") },
            { 11, new NumberFormatString("0.00E+00") },
            { 12, new NumberFormatString("# ?/?") },
            { 13, new NumberFormatString("# ??/??") },
            { 14, new NumberFormatString("d/m/yyyy") },
            { 15, new NumberFormatString("d-mmm-yy") },
            { 16, new NumberFormatString("d-mmm") },
            { 17, new NumberFormatString("mmm-yy") },
            { 18, new NumberFormatString("h:mm AM/PM") },
            { 19, new NumberFormatString("h:mm:ss AM/PM") },
            { 20, new NumberFormatString("h:mm") },
            { 21, new NumberFormatString("h:mm:ss") },
            { 22, new NumberFormatString("m/d/yy h:mm") },

            // 23..36 international/unused
            { 37, new NumberFormatString("#,##0_);(#,##0)") },
            { 38, new NumberFormatString("#,##0_);[Red](#,##0)") },
            { 39, new NumberFormatString("#,##0.00_);(#,##0.00)") },
            { 40, new NumberFormatString("#,##0.00_);[Red](#,##0.00)") },
            { 41, new NumberFormatString("_(\"$\"* #,##0_);_(\"$\"* (#,##0);_(\"$\"* \"-\"_);_(@_)") },
            { 42, new NumberFormatString("_(* #,##0_);_(* (#,##0);_(* \"-\"_);_(@_)") },
            { 43, new NumberFormatString("_(\"$\"* #,##0.00_);_(\"$\"* (#,##0.00);_(\"$\"* \"-\"??_);_(@_)") },
            { 44, new NumberFormatString("_(* #,##0.00_);_(* (#,##0.00);_(* \"-\"??_);_(@_)") },
            { 45, new NumberFormatString("mm:ss") },
            { 46, new NumberFormatString("[h]:mm:ss") },
            { 47, new NumberFormatString("mm:ss.0") },
            { 48, new NumberFormatString("##0.0E+0") },
            { 49, new NumberFormatString("@") },
        };

        public static NumberFormatString GetBuiltinNumberFormat(int numFmtId)
        {
            if (Formats.TryGetValue(numFmtId, out var result))
                return result;

            return null;
        }
    }
}


namespace ExcelDataReader.Core
{
    internal class Cell
    {
        public Cell(int columnIndex, object value, ExtendedFormat effectiveStyle, CellError? error)
        {
            ColumnIndex = columnIndex;
            Value = value;
            EffectiveStyle = effectiveStyle;
            Error = error;
        }

        /// <summary>
        /// Gets the zero-based column index.
        /// </summary>
        public int ColumnIndex { get; }

        /// <summary>
        /// Gets the effective style on the cell. The effective style is determined from
        /// the Cell XF, with optional overrides from a Cell Style XF.
        /// </summary>
        public ExtendedFormat EffectiveStyle { get; }

        public object Value { get; }

        public CellError? Error { get; }
    }
}

#nullable enable

namespace ExcelDataReader.Core
{
    internal class Column
    {
        public Column(int minimum, int maximum, bool hidden, double? width)
        {
            Minimum = minimum;
            Maximum = maximum;
            Hidden = hidden;
            Width = width;
        }

        public int Minimum { get; }

        public int Maximum { get; }

        public bool Hidden { get; }

        public double? Width { get; }
    }
}


namespace ExcelDataReader.Core
{
    /// <summary>
    /// Common handling of extended formats (XF) and mappings between file-based and global number format indices.
    /// </summary>
    internal class CommonWorkbook
    {
        /// <summary>
        /// Gets the dictionary of global number format strings. Always includes the built-in formats at their
        /// corresponding indices and any additional formats specified in the workbook file.
        /// </summary>
        public Dictionary<int, NumberFormatString> Formats { get; } = new Dictionary<int, NumberFormatString>();

        /// <summary>
        /// Gets the Cell XFs
        /// </summary>
        public List<ExtendedFormat> ExtendedFormats { get; } = new List<ExtendedFormat>();

        /// <summary>
        /// Gets the Cell Style XFs
        /// </summary>
        public List<ExtendedFormat> CellStyleExtendedFormats { get; } = new List<ExtendedFormat>();

        private NumberFormatString GeneralNumberFormat { get; } = new NumberFormatString("General");

        public ExtendedFormat GetEffectiveCellStyle(int xfIndex, int numberFormatFromCell)
        {
            if (xfIndex >= 0 && xfIndex < ExtendedFormats.Count)
            {
                return ExtendedFormats[xfIndex];
            }

            return new ExtendedFormat()
            {
                NumberFormatIndex = numberFormatFromCell,
            };
        }

        /// <summary>
        /// Registers a number format string in the workbook's Formats dictionary.
        /// </summary>
        public void AddNumberFormat(int formatIndexInFile, string formatString)
        {
            if (!Formats.ContainsKey(formatIndexInFile))
                Formats.Add(formatIndexInFile, new NumberFormatString(formatString));
        }

        public NumberFormatString GetNumberFormatString(int numberFormatIndex)
        {
            if (Formats.TryGetValue(numberFormatIndex, out var numberFormat))
            {
                return numberFormat;
            }

            numberFormat = BuiltinNumberFormat.GetBuiltinNumberFormat(numberFormatIndex);
            if (numberFormat != null)
            {
                return numberFormat;
            }

            // Fall back to "General" if the number format index is invalid
            return GeneralNumberFormat;
        }
    }
}


namespace ExcelDataReader.Core
{
    internal class EncodingHelper
    {
        public static Encoding GetEncoding(ushort codePage)
        {
            var encoding = (Encoding)null;
            switch (codePage)
            {
                case 037: encoding = Encoding.GetEncoding("IBM037"); break;
                case 437: encoding = Encoding.GetEncoding("IBM437"); break;
                case 500: encoding = Encoding.GetEncoding("IBM500"); break;
                case 708: encoding = Encoding.GetEncoding("ASMO-708"); break;
                case 709: encoding = Encoding.GetEncoding(string.Empty); break;
                case 710: encoding = Encoding.GetEncoding(string.Empty); break;
                case 720: encoding = Encoding.GetEncoding("DOS-720"); break;
                case 737: encoding = Encoding.GetEncoding("ibm737"); break;
                case 775: encoding = Encoding.GetEncoding("ibm775"); break;
                case 850: encoding = Encoding.GetEncoding("ibm850"); break;
                case 852: encoding = Encoding.GetEncoding("ibm852"); break;
                case 855: encoding = Encoding.GetEncoding("IBM855"); break;
                case 857: encoding = Encoding.GetEncoding("ibm857"); break;
                case 858: encoding = Encoding.GetEncoding("IBM00858"); break;
                case 860: encoding = Encoding.GetEncoding("IBM860"); break;
                case 861: encoding = Encoding.GetEncoding("ibm861"); break;
                case 862: encoding = Encoding.GetEncoding("DOS-862"); break;
                case 863: encoding = Encoding.GetEncoding("IBM863"); break;
                case 864: encoding = Encoding.GetEncoding("IBM864"); break;
                case 865: encoding = Encoding.GetEncoding("IBM865"); break;
                case 866: encoding = Encoding.GetEncoding("cp866"); break;
                case 869: encoding = Encoding.GetEncoding("ibm869"); break;
                case 870: encoding = Encoding.GetEncoding("IBM870"); break;
                case 874: encoding = Encoding.GetEncoding("windows-874"); break;
                case 875: encoding = Encoding.GetEncoding("cp875"); break;
                case 932: encoding = Encoding.GetEncoding("shift_jis"); break;
                case 936: encoding = Encoding.GetEncoding("gb2312"); break;
                case 949: encoding = Encoding.GetEncoding("ks_c_5601-1987"); break;
                case 950: encoding = Encoding.GetEncoding("big5"); break;
                case 1026: encoding = Encoding.GetEncoding("IBM1026"); break;
                case 1047: encoding = Encoding.GetEncoding("IBM01047"); break;
                case 1140: encoding = Encoding.GetEncoding("IBM01140"); break;
                case 1141: encoding = Encoding.GetEncoding("IBM01141"); break;
                case 1142: encoding = Encoding.GetEncoding("IBM01142"); break;
                case 1143: encoding = Encoding.GetEncoding("IBM01143"); break;
                case 1144: encoding = Encoding.GetEncoding("IBM01144"); break;
                case 1145: encoding = Encoding.GetEncoding("IBM01145"); break;
                case 1146: encoding = Encoding.GetEncoding("IBM01146"); break;
                case 1147: encoding = Encoding.GetEncoding("IBM01147"); break;
                case 1148: encoding = Encoding.GetEncoding("IBM01148"); break;
                case 1149: encoding = Encoding.GetEncoding("IBM01149"); break;
                case 1200: encoding = Encoding.GetEncoding("utf-16"); break;
                case 1201: encoding = Encoding.GetEncoding("unicodeFFFE"); break;
                case 1250: encoding = Encoding.GetEncoding("windows-1250"); break;
                case 1251: encoding = Encoding.GetEncoding("windows-1251"); break;
                case 1252: encoding = Encoding.GetEncoding("windows-1252"); break;
                case 1253: encoding = Encoding.GetEncoding("windows-1253"); break;
                case 1254: encoding = Encoding.GetEncoding("windows-1254"); break;
                case 1255: encoding = Encoding.GetEncoding("windows-1255"); break;
                case 1256: encoding = Encoding.GetEncoding("windows-1256"); break;
                case 1257: encoding = Encoding.GetEncoding("windows-1257"); break;
                case 1258: encoding = Encoding.GetEncoding("windows-1258"); break;
                case 1361: encoding = Encoding.GetEncoding("Johab"); break;
                case 10000: encoding = Encoding.GetEncoding("macintosh"); break;
                case 10001: encoding = Encoding.GetEncoding("x-mac-japanese"); break;
                case 10002: encoding = Encoding.GetEncoding("x-mac-chinesetrad"); break;
                case 10003: encoding = Encoding.GetEncoding("x-mac-korean"); break;
                case 10004: encoding = Encoding.GetEncoding("x-mac-arabic"); break;
                case 10005: encoding = Encoding.GetEncoding("x-mac-hebrew"); break;
                case 10006: encoding = Encoding.GetEncoding("x-mac-greek"); break;
                case 10007: encoding = Encoding.GetEncoding("x-mac-cyrillic"); break;
                case 10008: encoding = Encoding.GetEncoding("x-mac-chinesesimp"); break;
                case 10010: encoding = Encoding.GetEncoding("x-mac-romanian"); break;
                case 10017: encoding = Encoding.GetEncoding("x-mac-ukrainian"); break;
                case 10021: encoding = Encoding.GetEncoding("x-mac-thai"); break;
                case 10029: encoding = Encoding.GetEncoding("x-mac-ce"); break;
                case 10079: encoding = Encoding.GetEncoding("x-mac-icelandic"); break;
                case 10081: encoding = Encoding.GetEncoding("x-mac-turkish"); break;
                case 10082: encoding = Encoding.GetEncoding("x-mac-croatian"); break;
                case 12000: encoding = Encoding.GetEncoding("utf-32"); break;
                case 12001: encoding = Encoding.GetEncoding("utf-32BE"); break;
                case 20000: encoding = Encoding.GetEncoding("x-Chinese_CNS"); break;
                case 20001: encoding = Encoding.GetEncoding("x-cp20001"); break;
                case 20002: encoding = Encoding.GetEncoding("x_Chinese-Eten"); break;
                case 20003: encoding = Encoding.GetEncoding("x-cp20003"); break;
                case 20004: encoding = Encoding.GetEncoding("x-cp20004"); break;
                case 20005: encoding = Encoding.GetEncoding("x-cp20005"); break;
                case 20105: encoding = Encoding.GetEncoding("x-IA5"); break;
                case 20106: encoding = Encoding.GetEncoding("x-IA5-German"); break;
                case 20107: encoding = Encoding.GetEncoding("x-IA5-Swedish"); break;
                case 20108: encoding = Encoding.GetEncoding("x-IA5-Norwegian"); break;
                case 20127: encoding = Encoding.GetEncoding("us-ascii"); break;
                case 20261: encoding = Encoding.GetEncoding("x-cp20261"); break;
                case 20269: encoding = Encoding.GetEncoding("x-cp20269"); break;
                case 20273: encoding = Encoding.GetEncoding("IBM273"); break;
                case 20277: encoding = Encoding.GetEncoding("IBM277"); break;
                case 20278: encoding = Encoding.GetEncoding("IBM278"); break;
                case 20280: encoding = Encoding.GetEncoding("IBM280"); break;
                case 20284: encoding = Encoding.GetEncoding("IBM284"); break;
                case 20285: encoding = Encoding.GetEncoding("IBM285"); break;
                case 20290: encoding = Encoding.GetEncoding("IBM290"); break;
                case 20297: encoding = Encoding.GetEncoding("IBM297"); break;
                case 20420: encoding = Encoding.GetEncoding("IBM420"); break;
                case 20423: encoding = Encoding.GetEncoding("IBM423"); break;
                case 20424: encoding = Encoding.GetEncoding("IBM424"); break;
                case 20833: encoding = Encoding.GetEncoding("x-EBCDIC-KoreanExtended"); break;
                case 20838: encoding = Encoding.GetEncoding("IBM-Thai"); break;
                case 20866: encoding = Encoding.GetEncoding("koi8-r"); break;
                case 20871: encoding = Encoding.GetEncoding("IBM871"); break;
                case 20880: encoding = Encoding.GetEncoding("IBM880"); break;
                case 20905: encoding = Encoding.GetEncoding("IBM905"); break;
                case 20924: encoding = Encoding.GetEncoding("IBM00924"); break;
                case 20932: encoding = Encoding.GetEncoding("EUC-JP"); break;
                case 20936: encoding = Encoding.GetEncoding("x-cp20936"); break;
                case 20949: encoding = Encoding.GetEncoding("x-cp20949"); break;
                case 21025: encoding = Encoding.GetEncoding("cp1025"); break;
                case 21027: encoding = Encoding.GetEncoding(string.Empty); break;
                case 21866: encoding = Encoding.GetEncoding("koi8-u"); break;
                case 28591: encoding = Encoding.GetEncoding("iso-8859-1"); break;
                case 28592: encoding = Encoding.GetEncoding("iso-8859-2"); break;
                case 28593: encoding = Encoding.GetEncoding("iso-8859-3"); break;
                case 28594: encoding = Encoding.GetEncoding("iso-8859-4"); break;
                case 28595: encoding = Encoding.GetEncoding("iso-8859-5"); break;
                case 28596: encoding = Encoding.GetEncoding("iso-8859-6"); break;
                case 28597: encoding = Encoding.GetEncoding("iso-8859-7"); break;
                case 28598: encoding = Encoding.GetEncoding("iso-8859-8"); break;
                case 28599: encoding = Encoding.GetEncoding("iso-8859-9"); break;
                case 28603: encoding = Encoding.GetEncoding("iso-8859-13"); break;
                case 28605: encoding = Encoding.GetEncoding("iso-8859-15"); break;
                case 29001: encoding = Encoding.GetEncoding("x-Europa"); break;
                case 32768: encoding = Encoding.GetEncoding("macintosh"); break;
                case 32769: encoding = Encoding.GetEncoding("windows-1252"); break;
                case 38598: encoding = Encoding.GetEncoding("iso-8859-8-i"); break;
                case 50220: encoding = Encoding.GetEncoding("iso-2022-jp"); break;
                case 50221: encoding = Encoding.GetEncoding("csISO2022JP"); break;
                case 50222: encoding = Encoding.GetEncoding("iso-2022-jp"); break;
                case 50225: encoding = Encoding.GetEncoding("iso-2022-kr"); break;
                case 50227: encoding = Encoding.GetEncoding("x-cp50227"); break;
                case 50229: encoding = Encoding.GetEncoding(string.Empty); break;
                case 50930: encoding = Encoding.GetEncoding(string.Empty); break;
                case 50931: encoding = Encoding.GetEncoding(string.Empty); break;
                case 50933: encoding = Encoding.GetEncoding(string.Empty); break;
                case 50935: encoding = Encoding.GetEncoding(string.Empty); break;
                case 50936: encoding = Encoding.GetEncoding(string.Empty); break;
                case 50937: encoding = Encoding.GetEncoding(string.Empty); break;
                case 50939: encoding = Encoding.GetEncoding(string.Empty); break;
                case 51932: encoding = Encoding.GetEncoding("euc-jp"); break;
                case 51936: encoding = Encoding.GetEncoding("EUC-CN"); break;
                case 51949: encoding = Encoding.GetEncoding("euc-kr"); break;
                case 51950: encoding = Encoding.GetEncoding(string.Empty); break;
                case 52936: encoding = Encoding.GetEncoding("hz-gb-2312"); break;
                case 54936: encoding = Encoding.GetEncoding("GB18030"); break;
                case 57002: encoding = Encoding.GetEncoding("x-iscii-de"); break;
                case 57003: encoding = Encoding.GetEncoding("x-iscii-be"); break;
                case 57004: encoding = Encoding.GetEncoding("x-iscii-ta"); break;
                case 57005: encoding = Encoding.GetEncoding("x-iscii-te"); break;
                case 57006: encoding = Encoding.GetEncoding("x-iscii-as"); break;
                case 57007: encoding = Encoding.GetEncoding("x-iscii-or"); break;
                case 57008: encoding = Encoding.GetEncoding("x-iscii-ka"); break;
                case 57009: encoding = Encoding.GetEncoding("x-iscii-ma"); break;
                case 57010: encoding = Encoding.GetEncoding("x-iscii-gu"); break;
                case 57011: encoding = Encoding.GetEncoding("x-iscii-pa"); break;
                case 65000: encoding = Encoding.GetEncoding("utf-7"); break;
                case 65001: encoding = Encoding.GetEncoding("utf-8"); break;
            }

            return encoding;
        }
    }
}
namespace ExcelDataReader.Core
{
    internal class ExtendedFormat
    {
        /// <summary>
        /// Gets or sets the index to the parent Cell Style CF record with overrides for this XF. Only used with Cell XFs.
        /// 0xFFF means no override
        /// </summary>
        public int ParentCellStyleXf { get; set; }

        public int FontIndex { get; set; }

        public int NumberFormatIndex { get; set; }

        public bool Locked { get; set; }

        public bool Hidden { get; set; }

        public int IndentLevel { get; set; }

        public HorizontalAlignment HorizontalAlignment { get; set; }
    }
}


namespace ExcelDataReader.Core
{
    /// <summary>
    /// Helpers class
    /// </summary>
    internal static class Helpers
    {
        private static readonly Regex EscapeRegex = new Regex("_x([0-9A-F]{4,4})_");

        /// <summary>
        /// Determines whether the encoding is single byte or not.
        /// </summary>
        /// <param name="encoding">The encoding.</param>
        /// <returns>
        ///     <see langword="true"/> if the specified encoding is single byte; otherwise, <see langword="false"/>.
        /// </returns>
        public static bool IsSingleByteEncoding(Encoding encoding)
        {
            return encoding.GetByteCount(new[] { 'a' }) == 1;
        }

        public static string ConvertEscapeChars(string input)
        {
            return EscapeRegex.Replace(input, m => ((char)uint.Parse(m.Groups[1].Value, NumberStyles.HexNumber)).ToString());
        }

        /// <summary>
        /// Convert a double from Excel to an OA DateTime double. 
        /// The returned value is normalized to the '1900' date mode and adjusted for the 1900 leap year bug.
        /// </summary>
        public static double AdjustOADateTime(double value, bool date1904)
        {
            if (!date1904)
            {
                // Workaround for 1900 leap year bug in Excel
                if (value >= 0.0 && value < 60.0)
                {
                    return value + 1;
                }
            }
            else
            {
                return value + 1462.0;
            }

            return value;
        }

        public static bool IsValidOADateTime(double value)
        {
            return value > DateTimeHelper.OADateMinAsDouble && value < DateTimeHelper.OADateMaxAsDouble;
        }

        public static object ConvertFromOATime(double value, bool date1904)
        {
            var dateValue = AdjustOADateTime(value, date1904);
            if (IsValidOADateTime(dateValue))
                return DateTimeHelper.FromOADate(dateValue);
            return value;
        }

        public static object ConvertFromOATime(int value, bool date1904)
        {
            var dateValue = AdjustOADateTime(value, date1904);
            if (IsValidOADateTime(dateValue))
                return DateTimeHelper.FromOADate(dateValue);
            return value;
        }
    }
}

namespace ExcelDataReader.Core
{
    /// <summary>
    /// The common workbook interface between the binary and OpenXml formats
    /// </summary>
    /// <typeparam name="TWorksheet">A type implementing IWorksheet</typeparam>
    internal interface IWorkbook<TWorksheet>
        where TWorksheet : IWorksheet
    {
        int ResultsCount { get; }

        IEnumerable<TWorksheet> ReadWorksheets();

        NumberFormatString GetNumberFormatString(int index);
    }
}


namespace ExcelDataReader.Core
{
    /// <summary>
    /// The common worksheet interface between the binary and OpenXml formats
    /// </summary>
    internal interface IWorksheet
    {
        string Name { get; }

        string CodeName { get; }

        string VisibleState { get; }

        HeaderFooter HeaderFooter { get; }

        int FieldCount { get; }

        int RowCount { get; }

        CellRange[] MergeCells { get; }

        Column[] ColumnWidths { get; }

        IEnumerable<Row> ReadRows();
    }
}


namespace ExcelDataReader.Core
{
    internal static class ReferenceHelper
    {
        /// <summary>
        /// Logic for the Excel dimensions. Ex: A15
        /// </summary>
        /// <param name="value">The value.</param>
        /// <param name="column">The column, 1-based.</param>
        /// <param name="row">The row, 1-based.</param>
        public static bool ParseReference(string value, out int column, out int row)
        {
            column = 0;
            var position = 0;
            const int offset = 'A' - 1;

            if (value != null)
            {
                while (position < value.Length)
                {
                    var c = value[position];
                    if (c >= 'A' && c <= 'Z')
                    {
                        position++;
                        column *= 26;
                        column += c - offset;
                        continue;
                    }

                    if (char.IsDigit(c))
                        break;

                    position = 0;
                    break;
                }
            }

            if (position == 0)
            {
                column = 0;
                row = 0;
                return false;
            }

            if (!int.TryParse(value.Substring(position), NumberStyles.None, CultureInfo.InvariantCulture, out row))
            {
                return false;
            }

            return true;
        }
    }
}


#nullable enable

namespace ExcelDataReader.Core
{
    internal class Row
    {
        public Row(int rowIndex, double height, List<Cell> cells) 
        {
            RowIndex = rowIndex;
            Height = height;
            Cells = cells;
        }

        /// <summary>
        /// Gets the zero-based row index.
        /// </summary>
        public int RowIndex { get; }

        /// <summary>
        /// Gets the height of this row in points. Zero if hidden or collapsed.
        /// </summary>
        public double Height { get; }

        /// <summary>
        /// Gets the cells in this row.
        /// </summary>
        public List<Cell> Cells { get; }
    }
}


namespace ExcelDataReader.Core
{
    public static class StringHelper
    {
        public static bool IsSingleByteEncoding(this Encoding encoding)
        {
            return encoding.GetByteCount(new char[] { 'a' }) == 1;
        }
    }
}

namespace ExcelDataReader.Core
{
    internal static class XmlReaderHelper
    {
        public static bool ReadFirstContent(XmlReader xmlReader)
        {
            if (xmlReader.IsEmptyElement)
            {
                xmlReader.Read();
                return false;
            }

            xmlReader.MoveToContent();
            xmlReader.Read();
            return true;
        }

        public static bool SkipContent(XmlReader xmlReader)
        {
            if (xmlReader.NodeType == XmlNodeType.EndElement)
            {
                xmlReader.Read();
                return false;
            }

            xmlReader.Skip();
            return true;
        }
    }
}


namespace ExcelDataReader.Exceptions
{
    /// <summary>
    /// Thrown when there is a problem parsing the Compound Document container format used by XLS and password-protected XLSX.
    /// </summary>
    public class CompoundDocumentException : ExcelReaderException
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="CompoundDocumentException"/> class.
        /// </summary>
        /// <param name="message">The error message</param>
        public CompoundDocumentException(string message)
            : base(message)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="CompoundDocumentException"/> class.
        /// </summary>
        /// <param name="message">The error message</param>
        /// <param name="inner">The inner exception</param>
        public CompoundDocumentException(string message, Exception inner)
            : base(message, inner)
        {
        }
    }
}

#if NET20 || NET45 || NETSTANDARD2_0
#endif

namespace ExcelDataReader.Exceptions
{
    /// <summary>
    /// Base class for exceptions thrown by ExcelDataReader
    /// </summary>
#if NET20 || NET45 || NETSTANDARD2_0
    [Serializable]
#endif
    public class ExcelReaderException : Exception
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelReaderException"/> class.
        /// </summary>
        public ExcelReaderException()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelReaderException"/> class.
        /// </summary>
        /// <param name="message">The error message</param>
        public ExcelReaderException(string message)
            : base(message)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelReaderException"/> class.
        /// </summary>
        /// <param name="message">The error message</param>
        /// <param name="inner">The inner exception</param>
        public ExcelReaderException(string message, Exception inner)
            : base(message, inner)
        {
        }

#if NET20 || NET45 || NETSTANDARD2_0
        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelReaderException"/> class.
        /// </summary>
        /// <param name="info">The serialization info</param>
        /// <param name="context">The streaming context</param>
        protected ExcelReaderException(SerializationInfo info, StreamingContext context)
            : base(info, context)
        {
        }
#endif
    }
}

#if NET20 || NET45 || NETSTANDARD2_0
#endif

namespace ExcelDataReader.Exceptions
{
    /// <summary>
    /// Thrown when ExcelDataReader cannot parse the header
    /// </summary>
#if NET20 || NET45 || NETSTANDARD2_0
    [Serializable]
#endif
    public class HeaderException : ExcelReaderException
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="HeaderException"/> class.
        /// </summary>
        public HeaderException()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="HeaderException"/> class.
        /// </summary>
        /// <param name="message">The error message</param>
        public HeaderException(string message)
            : base(message)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="HeaderException"/> class.
        /// </summary>
        /// <param name="message">The error message</param>
        /// <param name="inner">The inner exception</param>
        public HeaderException(string message, Exception inner)
            : base(message, inner)
        {
        }

#if NET20 || NET45 || NETSTANDARD2_0
        /// <summary>
        /// Initializes a new instance of the <see cref="HeaderException"/> class.
        /// </summary>
        /// <param name="info">The serialization info</param>
        /// <param name="context">The streaming context</param>
        protected HeaderException(SerializationInfo info, StreamingContext context)
            : base(info, context)
        {
        }
#endif
    }
}

namespace ExcelDataReader.Exceptions
{
    /// <summary>
    /// Thrown when ExcelDataReader cannot open a password protected document because the password
    /// </summary>
    public class InvalidPasswordException : ExcelReaderException
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="InvalidPasswordException"/> class.
        /// </summary>
        /// <param name="message">The error message</param>
        public InvalidPasswordException(string message)
            : base(message)
        {
        }
    }
}


namespace ExcelDataReader.Log
{
    /// <summary>
    /// Custom interface for logging messages
    /// </summary>
    public interface ILog
    {
        /// <summary>
        /// Debug level of the specified message. The other method is preferred since the execution is deferred.
        /// </summary>
        /// <param name="message">The message.</param>
        /// <param name="formatting">The formatting.</param>
        void Debug(string message, params object[] formatting);

        /// <summary>
        /// Info level of the specified message. The other method is preferred since the execution is deferred.
        /// </summary>
        /// <param name="message">The message.</param>
        /// <param name="formatting">The formatting.</param>
        void Info(string message, params object[] formatting);

        /// <summary>
        /// Warn level of the specified message. The other method is preferred since the execution is deferred.
        /// </summary>
        /// <param name="message">The message.</param>
        /// <param name="formatting">The formatting.</param>
        void Warn(string message, params object[] formatting);

        /// <summary>
        /// Error level of the specified message. The other method is preferred since the execution is deferred.
        /// </summary>
        /// <param name="message">The message.</param>
        /// <param name="formatting">The formatting.</param>
        void Error(string message, params object[] formatting);

        /// <summary>
        /// Fatal level of the specified message. The other method is preferred since the execution is deferred.
        /// </summary>
        /// <param name="message">The message.</param>
        /// <param name="formatting">The formatting.</param>
        void Fatal(string message, params object[] formatting);
    }

    /// <summary>
    /// Factory interface for loggers.
    /// </summary>
    public interface ILogFactory
    {
        /// <summary>
        /// Create a logger for the specified type.
        /// </summary>
        /// <param name="loggingType">The type to create a logger for.</param>
        /// <returns>The logger instance.</returns>
        ILog Create(Type loggingType);
    }
}

namespace ExcelDataReader.Log
{
    /// <summary>
    /// logger type initialization
    /// </summary>
    public static class Log
    {
        private static readonly object LockObject = new object();

        private static Type logType = typeof(NullLogFactory);
        private static ILogFactory factoryInstance;

        /// <summary>
        /// Sets up logging to be with a certain type
        /// </summary>
        /// <typeparam name="T">The type of ILog for the application to use</typeparam>
        public static void InitializeWith<T>() 
            where T : ILogFactory, new()
        {
            lock (LockObject)
            {
                logType = typeof(T);
                factoryInstance = null;
            }
        }

        /// <summary>
        /// Initializes a new instance of a logger for an object.
        /// This should be done only once per object name.
        /// </summary>
        /// <param name="loggingType">The type to get a logger for.</param>
        /// <returns>ILog instance for an object if log type has been intialized; otherwise a null logger.</returns>
        public static ILog GetLoggerFor(Type loggingType)
        {
            var factory = factoryInstance;
            if (factory == null)
            {
                lock (LockObject)
                {
                    if (factory == null)
                    {
                        factory = factoryInstance = (ILogFactory)Activator.CreateInstance(logType);
                    }
                }
            }

            return factory.Create(loggingType);
        }
    }
}

namespace ExcelDataReader.Log
{
    /// <summary>
    /// 2.0 version of LogExtensions, not as awesome as Extension methods
    /// </summary>
    public static class LogManager
    {
        /// <summary>
        /// Gets the logger for a type.
        /// </summary>
        /// <typeparam name="T">The type to fetch a logger for.</typeparam>
        /// <param name="type">The type to get the logger for.</param>
        /// <returns>Instance of a logger for the object.</returns>
        /// <remarks>This method is thread safe.</remarks>
        public static ILog Log<T>(T type)
        {
            return ExcelDataReader.Log.Log.GetLoggerFor(typeof(T));
        }
    }
}


namespace ExcelDataReader.Misc
{
    internal static class DateTimeHelper
    {
        // All OA dates must be greater than (not >=) OADateMinAsDouble
        public const double OADateMinAsDouble = -657435.0;

        // All OA dates must be less than (not <=) OADateMaxAsDouble
        public const double OADateMaxAsDouble = 2958466.0;

        // From DateTime class to enable OADate in PCL
        // Number of 100ns ticks per time unit
        private const long TicksPerMillisecond = 10000;
        private const long TicksPerSecond = TicksPerMillisecond * 1000;
        private const long TicksPerMinute = TicksPerSecond * 60;
        private const long TicksPerHour = TicksPerMinute * 60;
        private const long TicksPerDay = TicksPerHour * 24;

        // Number of milliseconds per time unit
        private const int MillisPerSecond = 1000;
        private const int MillisPerMinute = MillisPerSecond * 60;
        private const int MillisPerHour = MillisPerMinute * 60;
        private const int MillisPerDay = MillisPerHour * 24;

        // Number of days in a non-leap year
        private const int DaysPerYear = 365;

        // Number of days in 4 years
        private const int DaysPer4Years = DaysPerYear * 4 + 1;

        // Number of days in 100 years
        private const int DaysPer100Years = DaysPer4Years * 25 - 1;

        // Number of days in 400 years
        private const int DaysPer400Years = DaysPer100Years * 4 + 1;

        // Number of days from 1/1/0001 to 12/30/1899
        private const int DaysTo1899 = DaysPer400Years * 4 + DaysPer100Years * 3 - 367;

        // Number of days from 1/1/0001 to 12/31/9999
        private const int DaysTo10000 = DaysPer400Years * 25 - 366;

        private const long MaxMillis = (long)DaysTo10000 * MillisPerDay;

        private const long DoubleDateOffset = DaysTo1899 * TicksPerDay;

        public static DateTime FromOADate(double d)
        {
            return new DateTime(DoubleDateToTicks(d), DateTimeKind.Unspecified);
        }

        // duplicated from DateTime
        internal static long DoubleDateToTicks(double value)
        {
            if (value >= OADateMaxAsDouble || value <= OADateMinAsDouble)
                throw new ArgumentException("Invalid OA Date");
            long millis = (long)(value * MillisPerDay + (value >= 0 ? 0.5 : -0.5));

            // The interesting thing here is when you have a value like 12.5 it all positive 12 days and 12 hours from 01/01/1899
            // However if you a value of -12.25 it is minus 12 days but still positive 6 hours, almost as though you meant -11.75 all negative
            // This line below fixes up the millis in the negative case
            if (millis < 0)
            {
                millis -= millis % MillisPerDay * 2;
            }

            millis += DoubleDateOffset / TicksPerMillisecond;

            if (millis < 0 || millis >= MaxMillis)
                throw new ArgumentException("OA Date out of range");
            return millis * TicksPerMillisecond;
        }
    }
}


namespace ExcelDataReader.Misc
{
    internal class LeaveOpenStream : Stream
    {
        public LeaveOpenStream(Stream baseStream)
        {
            BaseStream = baseStream;
        }

        public Stream BaseStream { get; }

        public override bool CanRead => BaseStream.CanRead;

        public override bool CanSeek => BaseStream.CanSeek;

        public override bool CanWrite => BaseStream.CanWrite;

        public override long Length => BaseStream.Length;

        public override long Position { get => BaseStream.Position; set => BaseStream.Position = value; }

        public override void Flush() => BaseStream.Flush();

        public override int Read(byte[] buffer, int offset, int count) => BaseStream.Read(buffer, offset, count);

        public override long Seek(long offset, SeekOrigin origin) => BaseStream.Seek(offset, origin);

        public override void SetLength(long value) => BaseStream.SetLength(value);

        public override void Write(byte[] buffer, int offset, int count) => BaseStream.Write(buffer, offset, count);
    }
}

#if NET20
using ICSharpCode.SharpZipLib.Zip;

namespace ExcelDataReader.Core
{
    internal sealed class ZipArchive : IDisposable
    {
        private readonly ZipFile _handle;

        public ZipArchive(Stream stream)
        {
            if (stream.CanSeek)
            {
                _handle = new ZipFile(stream);
            }
            else
            {
                // Password protected xlsx using "Standard Encryption" come as a non-seekable CryptoStream.
                // Must wrap in a MemoryStream to load
                var memoryStream = ReadToMemoryStream(stream);
                _handle = new ZipFile(memoryStream);
            }

            var entries = new List<ZipArchiveEntry>();
            foreach (ZipEntry entry in _handle)
                entries.Add(new ZipArchiveEntry(_handle, entry));
            Entries = new ReadOnlyCollection<ZipArchiveEntry>(entries);
        }

        public ReadOnlyCollection<ZipArchiveEntry> Entries { get; }

        public void Dispose()
        {
            (_handle as IDisposable)?.Dispose();
        }

        private static MemoryStream ReadToMemoryStream(Stream input)
        {
            byte[] buffer = new byte[16 * 1024];
            int read;
            var ms = new MemoryStream();
            while ((read = input.Read(buffer, 0, buffer.Length)) > 0)
            {
                ms.Write(buffer, 0, read);
            }

            ms.Position = 0;
            return ms;
        }
    }
}
#endif

#if NET20

using ICSharpCode.SharpZipLib.Zip;

namespace ExcelDataReader.Core
{
    internal sealed class ZipArchiveEntry
    {
        private readonly ZipFile _handle;
        private readonly ICSharpCode.SharpZipLib.Zip.ZipEntry _entry;

        internal ZipArchiveEntry(ZipFile handle, ICSharpCode.SharpZipLib.Zip.ZipEntry entry)
        {
            _handle = handle;
            _entry = entry;
        }

        public string FullName => _entry.Name;

        public Stream Open()
        {
            return _handle.GetInputStream(_entry);
        }
    }
}
#endif
// ReSharper disable InconsistentNaming
namespace ExcelDataReader.Core.BinaryFormat
{
    internal enum BIFFTYPE : ushort
    {
        WorkbookGlobals = 0x0005,
        VBModule = 0x0006,
        Worksheet = 0x0010,
        Chart = 0x0020,
#pragma warning disable SA1300 // Element must begin with upper-case letter
        v4MacroSheet = 0x0040,
        v4WorkbookGlobals = 0x0100
#pragma warning restore SA1300 // Element must begin with upper-case letter
    }

    internal enum BIFFRECORDTYPE : ushort
    {
        INTERFACEHDR = 0x00E1,
        MMS = 0x00C1,
        MERGECELLS = 0x00E5, // Record containing list of merged cell ranges
        INTERFACEEND = 0x00E2,
        WRITEACCESS = 0x005C,
        CODEPAGE = 0x0042,
        DSF = 0x0161,
        TABID = 0x013D,
        FNGROUPCOUNT = 0x009C,
        FILEPASS = 0x002F,
        WINDOWPROTECT = 0x0019,
        PROTECT = 0x0012,
        PASSWORD = 0x0013,
        PROT4REV = 0x01AF,
        PROT4REVPASSWORD = 0x01BC,
        WINDOW1 = 0x003D,
        BACKUP = 0x0040,
        HIDEOBJ = 0x008D,
        RECORD1904 = 0x0022,
        REFRESHALL = 0x01B7,
        BOOKBOOL = 0x00DA,

        FONT = 0x0031, // Font record, BIFF2, 5 and later

        FONT_V34 = 0x0231, // Font record, BIFF3, 4

        FORMAT = 0x041E, // Format record, BIFF4 and later

        FORMAT_V23 = 0x001E, // Format record, BIFF2, 3

        XF = 0x00E0, // Extended format record, BIFF5 and later

        XF_V4 = 0x0443, // Extended format record, BIFF4

        XF_V3 = 0x0243, // Extended format record, BIFF3

        XF_V2 = 0x0043, // Extended format record, BIFF2

        IXFE = 0x0044, // Index to XF, BIFF2

        STYLE = 0x0293,
        BOUNDSHEET = 0x0085,
        COUNTRY = 0x008C,
        SST = 0x00FC, // Global string storage (for BIFF8)

        CONTINUE = 0x003C,
        EXTSST = 0x00FF,
        BOF = 0x0809, // BOF Id for BIFF5 and later

        BOF_V2 = 0x0009, // BOF Id for BIFF2

        BOF_V3 = 0x0209, // BOF Id for BIFF3

        BOF_V4 = 0x0409, // BOF Id for BIFF4

        EOF = 0x000A, // End of block started with BOF

        CALCCOUNT = 0x000C,
        CALCMODE = 0x000D,
        PRECISION = 0x000E,
        REFMODE = 0x000F,
        DELTA = 0x0010,
        ITERATION = 0x0011,
        SAVERECALC = 0x005F,
        PRINTHEADERS = 0x002A,
        PRINTGRIDLINES = 0x002B,
        GUTS = 0x0080,
        WSBOOL = 0x0081,
        GRIDSET = 0x0082,
        DEFAULTROWHEIGHT_V2 = 0x0025,
        DEFAULTROWHEIGHT = 0x0225,
        HEADER = 0x0014,
        FOOTER = 0x0015,
        HCENTER = 0x0083,
        VCENTER = 0x0084,
        PRINTSETUP = 0x00A1,
        DFAULTCOLWIDTH = 0x0055,
        DIMENSIONS = 0x0200, // Size of area used for data
        DIMENSIONS_V2 = 0x0000, // BIFF2

        ROW_V2 = 0x0008, // Row record
        ROW = 0x0208, // Row record

        WINDOW2 = 0x023E,
        SELECTION = 0x001D,
        INDEX = 0x020B, // Index record, unsure about signature

        DBCELL = 0x00D7, // DBCell record, unsure about signature

        BLANK = 0x0201, // Empty cell

        BLANK_OLD = 0x0001, // Empty cell, old format

        MULBLANK = 0x00BE, // Equivalent of up to 256 blank cells

        INTEGER = 0x0202, // Integer cell (0..65535)

        INTEGER_OLD = 0x0002, // Integer cell (0..65535), old format

        NUMBER = 0x0203, // Numeric cell

        NUMBER_OLD = 0x0003, // Numeric cell, old format

        LABEL = 0x0204, // String cell (up to 255 symbols)

        LABEL_OLD = 0x0004, // String cell (up to 255 symbols), old format

        LABELSST = 0x00FD, // String cell with value from SST (for BIFF8)

        FORMULA = 0x0006, // Formula cell, BIFF2, BIFF5-8

        FORMULA_V3 = 0x0206, // Formula cell, BIFF3

        FORMULA_V4 = 0x0406, // Formula cell, BIFF4

        BOOLERR = 0x0205, // Boolean or error cell

        BOOLERR_OLD = 0x0005, // Boolean or error cell, old format

        ARRAY = 0x0221, // Range of cells for multi-cell formula

        RK = 0x027E, // RK-format numeric cell

        MULRK = 0x00BD, // Equivalent of up to 256 RK cells

        RSTRING = 0x00D6, // Rich-formatted string cell

        SHAREDFMLA = 0x04BC, // One more formula optimization element

        SHAREDFMLA_OLD = 0x00BC, // One more formula optimization element, old format

        STRING = 0x0207, // And one more, for string formula results

        STRING_OLD = 0x0007, // Old string formula results

        CF = 0x01B1,
        CODENAME = 0x01BA,
        CONDFMT = 0x01B0,
        DCONBIN = 0x01B5,
        DV = 0x01BE,
        DVAL = 0x01B2,
        HLINK = 0x01B8,
        MSODRAWINGGROUP = 0x00EB,
        MSODRAWING = 0x00EC,
        MSODRAWINGSELECTION = 0x00ED,
        PARAMQRY = 0x00DC,
        QSI = 0x01AD,
        SUPBOOK = 0x01AE,
        SXDB = 0x00C6,
        SXDBEX = 0x0122,
        SXFDBTYPE = 0x01BB,
        SXRULE = 0x00F0,
        SXEX = 0x00F1,
        SXFILT = 0x00F2,
        SXNAME = 0x00F6,
        SXSELECT = 0x00F7,
        SXPAIR = 0x00F8,
        SXFMLA = 0x00F9,
        SXFORMAT = 0x00FB,
        SXFORMULA = 0x0103,
        SXVDEX = 0x0100,
        TXO = 0x01B6,
        USERBVIEW = 0x01A9,
        USERSVIEWBEGIN = 0x01AA,
        USERSVIEWEND = 0x01AB,
        USESELFS = 0x0160,
        XL5MODIFY = 0x0162,
        OBJ = 0x005D,
        NOTE = 0x001C,
        SXEXT = 0x00DC,
        VERTICALPAGEBREAKS = 0x001A,
        XCT = 0x0059,

        /// <summary>
        /// If present the Calculate Message was in the status bar when Excel saved the file.
        /// This occurs if the sheet changed, the Manual calculation option was on, and the Recalculate Before Save option was off.
        /// </summary>
        UNCALCED = 0x005E,
        QUICKTIP = 0x0800,
        COLINFO = 0x007D
    }
}


namespace ExcelDataReader.Core.BinaryFormat
{
    internal interface IXlsString
    {
        /// <summary>
        /// Gets the string value. Encoding is only used with BIFF2-5 byte strings.
        /// </summary>
        string GetValue(Encoding encoding);
    }
}
namespace ExcelDataReader.Core.BinaryFormat
{
    /// <summary>
    /// Represents blank cell
    /// Base class for all cell types
    /// </summary>
    internal class XlsBiffBlankCell : XlsBiffRecord
    {
        internal XlsBiffBlankCell(byte[] bytes)
            : base(bytes)
        {
        }

        /// <summary>
        /// Gets the zero-based index of row containing this cell.
        /// </summary>
        public ushort RowIndex => ReadUInt16(0x0);

        /// <summary>
        /// Gets the zero-based index of column containing this cell.
        /// </summary>
        public ushort ColumnIndex => ReadUInt16(0x2);

        /// <summary>
        /// Gets the extended format used for this cell. If BIFF2 and this value is 63, this record was preceded by an IXFE record containing the actual XFormat >= 63.
        /// </summary>
        public ushort XFormat => IsBiff2Cell ? (ushort)(ReadByte(0x4) & 0x3F) : ReadUInt16(0x4);

        /// <summary>
        /// Gets the number format used for this cell. Only used in BIFF2 without XF records. Used by Excel 2.0/2.1 instead of XF/IXFE records.
        /// </summary>
        public ushort Format => IsBiff2Cell ? (ushort)(ReadByte(0x5) & 0x3F) : (ushort)0;

        /// <summary>
        /// Gets a value indicating whether the cell's record identifier is BIFF2-specific. 
        /// The shared binary layout of BIFF2 cells are different from BIFF3+.
        /// </summary>
        public bool IsBiff2Cell
        {
            get
            {
                switch (Id)
                {
                    case BIFFRECORDTYPE.NUMBER_OLD:
                    case BIFFRECORDTYPE.INTEGER_OLD:
                    case BIFFRECORDTYPE.LABEL_OLD:
                    case BIFFRECORDTYPE.BLANK_OLD:
                    case BIFFRECORDTYPE.BOOLERR_OLD:
                        return true;
                    default:
                        return false;
                }
            }
        }
    }
}
namespace ExcelDataReader.Core.BinaryFormat
{
    /// <summary>
    /// Represents BIFF BOF record
    /// </summary>
    internal class XlsBiffBOF : XlsBiffRecord
    {
        internal XlsBiffBOF(byte[] bytes)
            : base(bytes)
        {
        }

        /// <summary>
        /// Gets the version.
        /// </summary>
        public ushort Version => ReadUInt16(0x0);

        /// <summary>
        /// Gets the type of the BIFF block
        /// </summary>
        public BIFFTYPE Type => (BIFFTYPE)ReadUInt16(0x2);

        /// <summary>
        /// Gets the creation Id.
        /// </summary>
        /// <remarks>Not used before BIFF5</remarks>
        public ushort CreationId
        {
            get
            {
                if (RecordSize < 6)
                    return 0;
                return ReadUInt16(0x4);
            }
        }

        /// <summary>
        /// Gets the creation year.
        /// </summary>
        /// <remarks>Not used before BIFF5</remarks>
        public ushort CreationYear
        {
            get
            {
                if (RecordSize < 8)
                    return 0;
                return ReadUInt16(0x6);
            }
        }

        /// <summary>
        /// Gets the file history flag.
        /// </summary>
        /// <remarks>Not used before BIFF8</remarks>
        public uint HistoryFlag
        {
            get
            {
                if (RecordSize < 12)
                    return 0;
                return ReadUInt32(0x8);
            }
        }

        /// <summary>
        /// Gets the minimum Excel version to open this file.
        /// </summary>
        /// <remarks>Not used before BIFF8</remarks>
        public uint MinVersionToOpen
        {
            get
            {
                if (RecordSize < 16)
                    return 0;
                return ReadUInt32(0xC);
            }
        }
    }
}

namespace ExcelDataReader.Core.BinaryFormat
{
    /// <summary>
    /// Represents Sheet record in Workbook Globals
    /// </summary>
    internal class XlsBiffBoundSheet : XlsBiffRecord
    {
        private readonly IXlsString _sheetName;

        internal XlsBiffBoundSheet(byte[] bytes, int biffVersion)
            : base(bytes)
        {
            StartOffset = ReadUInt32(0x0);
            Type = (SheetType)ReadByte(0x5);
            VisibleState = (SheetVisibility)ReadByte(0x4);

            if (biffVersion == 8)
            {
                _sheetName = new XlsShortUnicodeString(bytes, ContentOffset + 6);
            }
            else if (biffVersion == 5)
            {
                _sheetName = new XlsShortByteString(bytes, ContentOffset + 6);
            }
            else 
            {
                throw new ArgumentException("Unexpected BIFF version " + biffVersion, nameof(biffVersion));
            }
        }

        internal XlsBiffBoundSheet(uint startOffset, SheetType type, SheetVisibility visibleState, string name)
            : base(new byte[32])
        {
            StartOffset = startOffset;
            Type = type;
            VisibleState = visibleState;
            _sheetName = new XlsInternalString(name);
        }

        public enum SheetType : byte
        {
            Worksheet = 0x0,
            MacroSheet = 0x1,
            Chart = 0x2,

            // ReSharper disable once InconsistentNaming
            VBModule = 0x6
        }

        public enum SheetVisibility : byte
        {
            Visible = 0x0,
            Hidden = 0x1,
            VeryHidden = 0x2
        }

        /// <summary>
        /// Gets the worksheet data start offset.
        /// </summary>
        public uint StartOffset { get; }

        /// <summary>
        /// Gets the worksheet type.
        /// </summary>
        public SheetType Type { get; }

        /// <summary>
        /// Gets the visibility of the worksheet.
        /// </summary>
        public SheetVisibility VisibleState { get; }

        /// <summary>
        /// Gets the name of the worksheet.
        /// </summary>
        public string GetSheetName(Encoding encoding)
        {
            return _sheetName.GetValue(encoding);
        }
    }
}

namespace ExcelDataReader.Core.BinaryFormat
{
    internal class XlsBiffCodeName : XlsBiffRecord
    {
        private readonly IXlsString _xlsString;

        internal XlsBiffCodeName(byte[] bytes)
            : base(bytes)
        {
            // BIFF8 only
            _xlsString = new XlsUnicodeString(bytes, ContentOffset);
        }

        public string GetValue(Encoding encoding)
        {
            return _xlsString.GetValue(encoding);
        }
    }
}


namespace ExcelDataReader.Core.BinaryFormat
{
    internal class XlsBiffColInfo : XlsBiffRecord
    {
        public XlsBiffColInfo(byte[] bytes)
            : base(bytes)
        {
            var colFirst = ReadUInt16(0x0);
            var colLast = ReadUInt16(0x2);
            var colDx = ReadUInt16(0x4);
            var flags = (ColInfoSettings)ReadUInt16(0x8);
            var userSet = (flags & ColInfoSettings.UserSet) != 0;
            var hidden = (flags & ColInfoSettings.Hidden) != 0;

            Value = new Column(colFirst, colLast, hidden, userSet ? (double?)colDx / 256.0 : null);
        }

        [Flags]
        private enum ColInfoSettings
        {
            Hidden = 0b01,
            UserSet = 0b10,
        }

        public Column Value { get; }
    }
}

namespace ExcelDataReader.Core.BinaryFormat
{
    /// <summary>
    /// Represents additional space for very large records
    /// </summary>
    internal class XlsBiffContinue : XlsBiffRecord
    {
        internal XlsBiffContinue(byte[] bytes)
            : base(bytes)
        {
        }
    }
}


namespace ExcelDataReader.Core.BinaryFormat
{
    internal class XlsBiffDefaultRowHeight : XlsBiffRecord
    {
        public XlsBiffDefaultRowHeight(byte[] bytes, int biffVersion)
            : base(bytes)
        {
            if (biffVersion == 2)
            {
                RowHeight = ReadUInt16(0x0) & 0x7FFF;
            }
            else
            {
                var flags = (DefaultRowHeightFlags)ReadUInt16(0x0);
                RowHeight = (flags & DefaultRowHeightFlags.DyZero) == 0 ? ReadUInt16(0x2) : 0;
                
                // UnhiddenRowHeight => (Flags & DefaultRowHeightFlags.DyZero) != 0 ? ReadInt16(0x2) : 0;
            }
        }

        internal enum DefaultRowHeightFlags : ushort
        {
            Unsynced = 1,
            DyZero = 2,
            ExAsc = 4,
            ExDsc = 8
        }

        /// <summary>
        /// Gets the row height in twips
        /// </summary>
        public int RowHeight { get; }
    }
}

namespace ExcelDataReader.Core.BinaryFormat
{
    /// <summary>
    /// Represents Dimensions of worksheet
    /// </summary>
    internal class XlsBiffDimensions : XlsBiffRecord
    {
        internal XlsBiffDimensions(byte[] bytes, int biffVersion)
            : base(bytes)
        {
            if (biffVersion < 8)
            {
                FirstRow = ReadUInt16(0x0);
                LastRow = ReadUInt16(0x2);
                FirstColumn = ReadUInt16(0x4);
                LastColumn = ReadUInt16(0x6);
            }
            else
            {
                FirstRow = ReadUInt32(0x0); // TODO: [MS-XLS] RwLongU
                LastRow = ReadUInt32(0x4);
                FirstColumn = ReadUInt16(0x8); // TODO: [MS-XLS] ColU
                LastColumn = ReadUInt16(0xA);
            }
        }

        /// <summary>
        /// Gets the index of first row.
        /// </summary>
        public uint FirstRow { get; }

        /// <summary>
        /// Gets the index of last row + 1.
        /// </summary>
        public uint LastRow { get; }

        /// <summary>
        /// Gets the index of first column.
        /// </summary>
        public ushort FirstColumn { get; }

        /// <summary>
        /// Gets the index of last column + 1.
        /// </summary>
        public ushort LastColumn { get; }
    }
}
namespace ExcelDataReader.Core.BinaryFormat
{
    /// <summary>
    /// Represents BIFF EOF resord
    /// </summary>
    internal class XlsBiffEof : XlsBiffRecord
    {
        internal XlsBiffEof(byte[] bytes)
            : base(bytes)
        {
        }
    }
}


namespace ExcelDataReader.Core.BinaryFormat
{
    /// <summary>
    /// Represents FILEPASS record containing XOR obfuscation details or a an EncryptionInfo structure
    /// </summary>
    internal class XlsBiffFilePass : XlsBiffRecord
    {
        internal XlsBiffFilePass(byte[] bytes, int biffVersion)
            : base(bytes)
        {
            if (biffVersion >= 2 && biffVersion <= 5)
            {
                // Cipher = EncryptionType.XOR;
                var encryptionKey = ReadUInt16(0);
                var hashValue = ReadUInt16(2);
                EncryptionInfo = EncryptionInfo.Create(encryptionKey, hashValue);
            }
            else
            {
                ushort type = ReadUInt16(0);

                if (type == 0)
                {
                    var encryptionKey = ReadUInt16(2);
                    var hashValue = ReadUInt16(4);
                    EncryptionInfo = EncryptionInfo.Create(encryptionKey, hashValue);
                }
                else if (type == 1)
                {
                    var encryptionInfo = new byte[bytes.Length - 6]; // 6 = 4 + 2 = biffVersion header + filepass enryptiontype
                    Array.Copy(bytes, 6, encryptionInfo, 0, bytes.Length - 6);
                    EncryptionInfo = EncryptionInfo.Create(encryptionInfo);
                }
                else
                {
                    throw new NotSupportedException("Unknown encryption type: " + type);
                }
            }
        }

        public EncryptionInfo EncryptionInfo { get; }
    }
}

namespace ExcelDataReader.Core.BinaryFormat
{
    /// <summary>
    /// The font with index 4 is omitted in all BIFF versions. This means the first four fonts have zero-based indexes, and the fifth font and all following fonts are referenced with one-based indexes.
    /// </summary>
    internal class XlsBiffFont : XlsBiffRecord
    {
        private readonly IXlsString _fontName;

        internal XlsBiffFont(byte[] bytes, int biffVersion)
            : base(bytes)
        {
            if (Id == BIFFRECORDTYPE.FONT_V34)
            {
                _fontName = new XlsShortByteString(bytes, ContentOffset + 6);
            }
            else if (Id == BIFFRECORDTYPE.FONT && biffVersion == 2)
            {
                _fontName = new XlsShortByteString(bytes, ContentOffset + 4);
            }
            else if (Id == BIFFRECORDTYPE.FONT && biffVersion == 5)
            {
                _fontName = new XlsShortByteString(bytes, ContentOffset + 14);
            }
            else if (Id == BIFFRECORDTYPE.FONT && biffVersion == 8)
            {
                _fontName = new XlsShortUnicodeString(bytes, ContentOffset + 14);
            }
            else
            {
                _fontName = new XlsInternalString(string.Empty);
            }

            if (Id == BIFFRECORDTYPE.FONT && biffVersion >= 5)
            {
                // Encodings were mapped by correlating this:
                // https://docs.microsoft.com/en-us/windows/desktop/intl/code-page-identifiers
                // with the FONT record character set table here:
                // https://www.openoffice.org/sc/excelfileformat.pdf
                var byteStringCharacterSet = ReadByte(12);
                switch (byteStringCharacterSet)
                {
                    case 0: // ANSI Latin
                    case 1: // System default
                        ByteStringEncoding = EncodingHelper.GetEncoding(1252);
                        break;
                    case 77: // Apple roman
                        ByteStringEncoding = EncodingHelper.GetEncoding(10000);
                        break;
                    case 128: // ANSI Japanese Shift-JIS
                        ByteStringEncoding = EncodingHelper.GetEncoding(932);
                        break;
                    case 129: // ANSI Korean (Hangul)
                        ByteStringEncoding = EncodingHelper.GetEncoding(949);
                        break;
                    case 130: // ANSI Korean (Johab)
                        ByteStringEncoding = EncodingHelper.GetEncoding(1361);
                        break;
                    case 134: // ANSI Chinese Simplified GBK
                        ByteStringEncoding = EncodingHelper.GetEncoding(936);
                        break;
                    case 136: // ANSI Chinese Traditional BIG5
                        ByteStringEncoding = EncodingHelper.GetEncoding(950);
                        break;
                    case 161: // ANSI Greek
                        ByteStringEncoding = EncodingHelper.GetEncoding(1253);
                        break;
                    case 162: // ANSI Turkish
                        ByteStringEncoding = EncodingHelper.GetEncoding(1254);
                        break;
                    case 163: // ANSI Vietnamese
                        ByteStringEncoding = EncodingHelper.GetEncoding(1258);
                        break;
                    case 177: // ANSI Hebrew
                        ByteStringEncoding = EncodingHelper.GetEncoding(1255);
                        break;
                    case 178: // ANSI Arabic
                        ByteStringEncoding = EncodingHelper.GetEncoding(1256);
                        break;
                    case 186: // ANSI Baltic
                        ByteStringEncoding = EncodingHelper.GetEncoding(1257);
                        break;
                    case 204: // ANSI Cyrillic
                        ByteStringEncoding = EncodingHelper.GetEncoding(1251);
                        break;
                    case 222: // ANSI Thai
                        ByteStringEncoding = EncodingHelper.GetEncoding(874);
                        break;
                    case 238: // ANSI Latin II
                        ByteStringEncoding = EncodingHelper.GetEncoding(1250);
                        break;
                    case 255: // OEM Latin
                        ByteStringEncoding = EncodingHelper.GetEncoding(850);
                        break;
                }
            }
        }

        public Encoding ByteStringEncoding { get; }

        public string GetFontName(Encoding encoding) => _fontName.GetValue(encoding);
    }
}


namespace ExcelDataReader.Core.BinaryFormat
{
    /// <summary>
    /// Represents a string value of format
    /// </summary>
    internal class XlsBiffFormatString : XlsBiffRecord
    {
        private readonly IXlsString _xlsString;

        internal XlsBiffFormatString(byte[] bytes, int biffVersion)
            : base(bytes)
        {
            if (Id == BIFFRECORDTYPE.FORMAT_V23)
            {
                // BIFF2-3
                _xlsString = new XlsShortByteString(bytes, ContentOffset);
            }
            else if (biffVersion >= 2 && biffVersion <= 5)
            {
                // BIFF4-5, or if there is a newer format record in a BIFF2-3 stream
                _xlsString = new XlsShortByteString(bytes, ContentOffset + 2);
            }
            else if (biffVersion == 8)
            {
                // BIFF8
                _xlsString = new XlsUnicodeString(bytes, ContentOffset + 2);
            }
            else
            {
                throw new ArgumentException("Unexpected BIFF version " + biffVersion, nameof(biffVersion));
            }
        }

        public ushort Index
        {
            get
            {
                switch (Id)
                {
                    case BIFFRECORDTYPE.FORMAT_V23:
                        throw new NotSupportedException("Index is not available for BIFF2 and BIFF3 FORMAT records.");
                    default:
                        return ReadUInt16(0);
                }
            }
        }

        /// <summary>
        /// Gets the string value.
        /// </summary>
        public string GetValue(Encoding encoding)
        {
            return _xlsString.GetValue(encoding);
        }
    }
}

namespace ExcelDataReader.Core.BinaryFormat
{
    /// <summary>
    /// Represents a cell containing formula
    /// </summary>
    internal class XlsBiffFormulaCell : XlsBiffBlankCell
    {
        // private FormulaFlags _flags;
        private readonly int _biffVersion;
        private bool _booleanValue;
        private CellError _errorValue;
        private double _xNumValue;
        private FormulaValueType _formulaType;
        private bool _initialized;

        internal XlsBiffFormulaCell(byte[] bytes, int biffVersion)
            : base(bytes)
        {
            _biffVersion = biffVersion;
        }

        [Flags]
        public enum FormulaFlags : ushort
        {
            AlwaysCalc = 0x0001,
            CalcOnLoad = 0x0002,
            SharedFormulaGroup = 0x0008
        }

        public enum FormulaValueType
        {
            Unknown,

            /// <summary>
            /// Indicates that a string value is stored in a String record that immediately follows this record. See[MS - XLS] 2.5.133 FormulaValue.
            /// </summary>
            String,

            /// <summary>
            /// Indecates that the formula value is an empty string.
            /// </summary>
            EmptyString,

            /// <summary>
            /// Indicates that the <see cref="BooleanValue"/> property is valid.
            /// </summary>
            Boolean,

            /// <summary>
            /// Indicates that the <see cref="ErrorValue"/> property is valid.
            /// </summary>
            Error,

            /// <summary>
            /// Indicates that the <see cref="XNumValue"/> property is valid.
            /// </summary>
            Number
        }

        /// <summary>
        /// Gets the formula value type.
        /// </summary>
        public FormulaValueType FormulaType
        {
            get
            {
                LazyInit();
                return _formulaType;
            }
        }

        public bool BooleanValue
        {
            get
            {
                LazyInit();
                return _booleanValue;
            }
        }

        public CellError ErrorValue
        {
            get
            {
                LazyInit();
                return _errorValue;
            }
        }

        public double XNumValue
        {
            get
            {
                LazyInit();
                return _xNumValue;
            }
        }

        /*
        public FormulaFlags Flags
        {
            get
            {
                LazyInit();
                return _flags;
            }
        }
        */

        private void LazyInit()
        {
            if (_initialized)
                return;
            _initialized = true;

            if (_biffVersion == 2)
            {
                // _flags = (FormulaFlags)ReadUInt16(0xF);
                _xNumValue = ReadDouble(0x7);
                _formulaType = FormulaValueType.Number;
            }
            else
            {
                // _flags = (FormulaFlags)ReadUInt16(0xE);
                var formulaValueExprO = ReadUInt16(0xC);
                if (formulaValueExprO != 0xFFFF)
                {
                    _formulaType = FormulaValueType.Number;
                    _xNumValue = ReadDouble(0x6);
                }
                else
                {
                    // var formulaLength = ReadByte(0xF);
                    var formulaValueByte1 = ReadByte(0x6);
                    var formulaValueByte3 = ReadByte(0x8);
                    switch (formulaValueByte1)
                    {
                        case 0x00:
                            _formulaType = FormulaValueType.String;
                            break;
                        case 0x01:
                            _formulaType = FormulaValueType.Boolean;
                            _booleanValue = formulaValueByte3 != 0;
                            break;
                        case 0x02:
                            _formulaType = FormulaValueType.Error;
                            _errorValue = (CellError)formulaValueByte3;
                            break;
                        case 0x03:
                            _formulaType = FormulaValueType.EmptyString;
                            break;
                        default:
                            _formulaType = FormulaValueType.Unknown;
                            break;
                    }
                }
            }
        }
    }
}

namespace ExcelDataReader.Core.BinaryFormat
{
    /// <summary>
    /// Represents a string value of formula
    /// </summary>
    internal class XlsBiffFormulaString : XlsBiffRecord
    {
        private readonly IXlsString _xlsString;

        internal XlsBiffFormulaString(byte[] bytes, int biffVersion)
            : base(bytes)
        {
            if (biffVersion == 2)
            {
                // BIFF2
                _xlsString = new XlsShortByteString(bytes, ContentOffset);
            }
            else if (biffVersion >= 3 && biffVersion <= 5)
            {
                // BIFF3-5
                _xlsString = new XlsByteString(bytes, ContentOffset);
            }
            else if (biffVersion == 8)
            {
                // BIFF8
                _xlsString = new XlsUnicodeString(bytes, ContentOffset);
            }
            else
            {
                throw new ArgumentException("Unexpected BIFF version " + biffVersion, nameof(biffVersion));
            }
        }

        /// <summary>
        /// Gets the string value.
        /// </summary>
        public string GetValue(Encoding encoding)
        {
            return _xlsString.GetValue(encoding);
        }
    }
}

namespace ExcelDataReader.Core.BinaryFormat
{
    /// <summary>
    /// Represents a string value of a header or footer.
    /// </summary>
    internal sealed class XlsBiffHeaderFooterString : XlsBiffRecord
    {
        private readonly IXlsString _xlsString;

        internal XlsBiffHeaderFooterString(byte[] bytes, int biffVersion)
            : base(bytes)
        {
            if (biffVersion < 8)
                _xlsString = new XlsShortByteString(bytes, ContentOffset);
            else if (biffVersion == 8)
                _xlsString = new XlsUnicodeString(bytes, ContentOffset);
            else
                throw new ArgumentException("Unexpected BIFF version " + biffVersion, nameof(biffVersion));
        }

        /// <summary>
        /// Gets the string value.
        /// </summary>
        public string GetValue(Encoding encoding)
        {
            return _xlsString.GetValue(encoding);
        }
    }
}
namespace ExcelDataReader.Core.BinaryFormat
{
    /// <summary>
    /// Represents a constant integer number in range 0..65535
    /// </summary>
    internal class XlsBiffIntegerCell : XlsBiffBlankCell
    {
        internal XlsBiffIntegerCell(byte[] bytes)
            : base(bytes)
        {
        }

        /// <summary>
        /// Gets the cell value.
        /// </summary>
        public int Value => Id == BIFFRECORDTYPE.INTEGER_OLD ? ReadUInt16(0x7) : ReadUInt16(0x6);
    }
}
namespace ExcelDataReader.Core.BinaryFormat
{
    /// <summary>
    /// Represents InterfaceHdr record in Wokrbook Globals
    /// </summary>
    internal class XlsBiffInterfaceHdr : XlsBiffRecord
    {
        internal XlsBiffInterfaceHdr(byte[] bytes)
            : base(bytes)
        {
        }

        /// <summary>
        /// Gets the CodePage for Interface Header
        /// </summary>
        public ushort CodePage => ReadUInt16(0x0);
    }
}


namespace ExcelDataReader.Core.BinaryFormat
{
    /// <summary>
    /// [MS-XLS] 2.4.148 Label
    /// Represents a string
    /// </summary>
    internal class XlsBiffLabelCell : XlsBiffBlankCell
    {
        private readonly IXlsString _xlsString;

        internal XlsBiffLabelCell(byte[] bytes, int biffVersion)
            : base(bytes)
        {
            if (Id == BIFFRECORDTYPE.LABEL_OLD)
            {
                // BIFF2
                _xlsString = new XlsShortByteString(bytes, ContentOffset + 7);
            }
            else if (biffVersion >= 2 && biffVersion <= 5)
            {
                // BIFF3-5, or if there is a newer label record present in a BIFF2 stream
                _xlsString = new XlsByteString(bytes, ContentOffset + 6);
            }
            else if (biffVersion == 8)
            {
                // BIFF8
                _xlsString = new XlsUnicodeString(bytes, ContentOffset + 6);
            }
            else
            {
                throw new ArgumentException("Unexpected BIFF version " + biffVersion, nameof(biffVersion));
            }
        }

        /// <summary>
        /// Gets the cell value.
        /// </summary>
        public string GetValue(Encoding encoding)
        {
            return _xlsString.GetValue(encoding);
        }
    }
}
namespace ExcelDataReader.Core.BinaryFormat
{
    /// <summary>
    /// Represents a string stored in SST
    /// </summary>
    internal class XlsBiffLabelSSTCell : XlsBiffBlankCell
    {
        internal XlsBiffLabelSSTCell(byte[] bytes)
            : base(bytes)
        {
        }

        /// <summary>
        /// Gets the index of string in Shared String Table
        /// </summary>
        public uint SSTIndex => ReadUInt32(0x6);
    }
}

namespace ExcelDataReader.Core.BinaryFormat
{
    /// <summary>
    /// [MS-XLS] 2.4.168 MergeCells
    ///  If the count of the merged cells in the document is greater than 1026, the file will contain multiple adjacent MergeCells records.
    /// </summary>
    internal class XlsBiffMergeCells : XlsBiffRecord
    {
        public XlsBiffMergeCells(byte[] bytes)
            : base(bytes)
        {
            var count = ReadUInt16(0);

            MergeCells = new List<CellRange>();
            for (int i = 0; i < count; i++)
            {
                var fromRow = ReadInt16(2 + i * 8 + 0);
                var toRow = ReadInt16(2 + i * 8 + 2);
                var fromCol = ReadInt16(2 + i * 8 + 4);
                var toCol = ReadInt16(2 + i * 8 + 6);

                CellRange mergeCell = new CellRange(fromCol, fromRow, toCol, toRow);
                MergeCells.Add(mergeCell);
            }
        }

        public List<CellRange> MergeCells { get; }
    }
}

namespace ExcelDataReader.Core.BinaryFormat
{
    /// <summary>
    /// Represents MSO Drawing record
    /// </summary>
    internal class XlsBiffMSODrawing : XlsBiffRecord
    {
        internal XlsBiffMSODrawing(byte[] bytes)
            : base(bytes)
        {
        }
    }
}

namespace ExcelDataReader.Core.BinaryFormat
{
    /// <summary>
    /// Represents multiple Blank cell
    /// </summary>
    internal class XlsBiffMulBlankCell : XlsBiffBlankCell
    {
        internal XlsBiffMulBlankCell(byte[] bytes)
            : base(bytes)
        {
        }

        /// <summary>
        /// Gets the zero-based index of last described column
        /// </summary>
        public ushort LastColumnIndex => ReadUInt16(RecordSize - 2);

        /// <summary>
        /// Returns format forspecified column, column must be between ColumnIndex and LastColumnIndex
        /// </summary>
        /// <param name="columnIdx">Index of column</param>
        /// <returns>Format</returns>
        public ushort GetXF(ushort columnIdx)
        {
            int ofs = 4 + 6 * (columnIdx - ColumnIndex);
            if (ofs > RecordSize - 2)
                return 0;
            return ReadUInt16(ofs);
        }
    }
}
namespace ExcelDataReader.Core.BinaryFormat
{
    /// <summary>
    /// Represents multiple RK number cells
    /// </summary>
    internal class XlsBiffMulRKCell : XlsBiffBlankCell
    {
        internal XlsBiffMulRKCell(byte[] bytes)
            : base(bytes)
        {
        }

        /// <summary>
        /// Gets the zero-based index of last described column
        /// </summary>
        public ushort LastColumnIndex => ReadUInt16(RecordSize - 2);

        /// <summary>
        /// Returns format for specified column
        /// </summary>
        /// <param name="columnIdx">Index of column, must be between ColumnIndex and LastColumnIndex</param>
        /// <returns>The format.</returns>
        public ushort GetXF(ushort columnIdx)
        {
            int ofs = 4 + 6 * (columnIdx - ColumnIndex);
            if (ofs > RecordSize - 2)
                return 0;
            return ReadUInt16(ofs);
        }

        /// <summary>
        /// Gets the value for specified column
        /// </summary>
        /// <param name="columnIdx">Index of column, must be between ColumnIndex and LastColumnIndex</param>
        /// <returns>The value.</returns>
        public double GetValue(ushort columnIdx)
        {
            int ofs = 6 + 6 * (columnIdx - ColumnIndex);
            if (ofs > RecordSize)
                return 0;
            return XlsBiffRKCell.NumFromRK(ReadUInt32(ofs));
        }
    }
}
namespace ExcelDataReader.Core.BinaryFormat
{
    /// <summary>
    /// Represents a floating-point number 
    /// </summary>
    internal class XlsBiffNumberCell : XlsBiffBlankCell
    {
        internal XlsBiffNumberCell(byte[] bytes)
            : base(bytes)
        {
        }

        /// <summary>
        /// Gets the value of this cell
        /// </summary>
        public double Value => Id == BIFFRECORDTYPE.NUMBER_OLD ? ReadDouble(0x7) : ReadDouble(0x6);
    }
}
namespace ExcelDataReader.Core.BinaryFormat
{
    /// <summary>
    /// For now QuickTip will do nothing, it seems to have a different
    /// </summary>
    internal class XlsBiffQuickTip : XlsBiffRecord
    {
        internal XlsBiffQuickTip(byte[] bytes)
            : base(bytes)
        {
        }
    }
}

namespace ExcelDataReader.Core.BinaryFormat
{
    /// <summary>
    /// Represents basic BIFF record
    /// Base class for all BIFF record types
    /// </summary>
    internal class XlsBiffRecord
    {
        protected const int ContentOffset = 4;
        
        public XlsBiffRecord(byte[] bytes)
        {
            if (bytes.Length < 4)
                throw new ArgumentException(Errors.ErrorBiffRecordSize);
            Bytes = bytes;
        }
        
        /// <summary>
        /// Gets the type Id of this entry
        /// </summary>
        public BIFFRECORDTYPE Id => (BIFFRECORDTYPE)BitConverter.ToUInt16(Bytes, 0);

        /// <summary>
        /// Gets the data size of this entry
        /// </summary>
        public ushort RecordSize => BitConverter.ToUInt16(Bytes, 2);

        /// <summary>
        /// Gets the whole size of structure
        /// </summary>
        public int Size => ContentOffset + RecordSize;
        
        internal byte[] Bytes { get; }

        public byte ReadByte(int offset)
        {
            return Buffer.GetByte(Bytes, ContentOffset + offset);
        }

        public ushort ReadUInt16(int offset)
        {
            return BitConverter.ToUInt16(Bytes, ContentOffset + offset);
        }

        public uint ReadUInt32(int offset)
        {
            return BitConverter.ToUInt32(Bytes, ContentOffset + offset);
        }

        public ulong ReadUInt64(int offset)
        {
            return BitConverter.ToUInt64(Bytes, ContentOffset + offset);
        }

        public short ReadInt16(int offset)
        {
            return BitConverter.ToInt16(Bytes, ContentOffset + offset);
        }

        public int ReadInt32(int offset)
        {
            return BitConverter.ToInt32(Bytes, ContentOffset + offset);
        }

        public long ReadInt64(int offset)
        {
            return BitConverter.ToInt64(Bytes, ContentOffset + offset);
        }

        public byte[] ReadArray(int offset, int size)
        {
            byte[] tmp = new byte[size];
            Buffer.BlockCopy(Bytes, ContentOffset + offset, tmp, 0, size);
            return tmp;
        }

        public float ReadFloat(int offset)
        {
            return BitConverter.ToSingle(Bytes, ContentOffset + offset);
        }

        public double ReadDouble(int offset)
        {
            return BitConverter.ToDouble(Bytes, ContentOffset + offset);
        }
    }
}


namespace ExcelDataReader.Core.BinaryFormat
{
    /// <summary>
    /// Represents an RK number cell
    /// </summary>
    internal class XlsBiffRKCell : XlsBiffBlankCell
    {
        internal XlsBiffRKCell(byte[] bytes)
            : base(bytes)
        {
        }

        /// <summary>
        /// Gets the value of this cell
        /// </summary>
        public double Value => NumFromRK(ReadUInt32(0x6));

        /// <summary>
        /// Decodes RK-encoded number
        /// </summary>
        /// <param name="rk">Encoded number</param>
        /// <returns>The number.</returns>
        public static double NumFromRK(uint rk)
        {
            double num;
            if ((rk & 0x2) == 0x2)
            {
                num = (int)(rk >> 2 | ((rk & 0x80000000) == 0 ? 0 : 0xC0000000));
            }
            else
            {
                // hi words of IEEE num
                num = BitConverter.Int64BitsToDouble((long)(rk & 0xfffffffc) << 32);
            }

            if ((rk & 0x1) == 0x1)
                num /= 100; // divide by 100

            return num;
        }
    }
}
namespace ExcelDataReader.Core.BinaryFormat
{
    /// <summary>
    /// Represents row record in table
    /// </summary>
    internal class XlsBiffRow : XlsBiffRecord
    {
        internal XlsBiffRow(byte[] bytes)
            : base(bytes)
        {
            if (Id == BIFFRECORDTYPE.ROW_V2)
            {
                RowIndex = ReadUInt16(0x0);
                FirstDefinedColumn = ReadUInt16(0x2);
                LastDefinedColumn = ReadUInt16(0x4);
                var heightBits = ReadUInt16(0x6);

                UseDefaultRowHeight = (heightBits & 0x8000) != 0;
                RowHeight = heightBits & 0x7FFFF;

                UseXFormat = ReadByte(0xA) != 0;
                if (UseXFormat)
                    XFormat = ReadUInt16(0x10);
            }
            else
            {
                RowIndex = ReadUInt16(0x0);
                FirstDefinedColumn = ReadUInt16(0x2);
                LastDefinedColumn = ReadUInt16(0x4);

                var heightBits = ReadUInt16(0x6);
                UseDefaultRowHeight = (heightBits & 0x8000) != 0;
                RowHeight = heightBits & 0x7FFFF;

                var flags = (RowHeightFlags)ReadUInt16(0xC);
                RowHeight = (flags & RowHeightFlags.DyZero) == 0 ? RowHeight : 0;

                UseXFormat = (flags & RowHeightFlags.GhostDirty) != 0;
                XFormat = (ushort)(ReadUInt16(0xE) & 0xFFF);
            }
        }

        internal enum RowHeightFlags : ushort
        {
            OutlineLevelMask = 3,
            Collapsed = 16,
            DyZero = 32,
            Unsynced = 64,
            GhostDirty = 128
        }

        /// <summary>
        /// Gets the zero-based index of row described
        /// </summary>
        public ushort RowIndex { get; }

        /// <summary>
        /// Gets the index of first defined column
        /// </summary>
        public ushort FirstDefinedColumn { get; }

        /// <summary>
        /// Gets the index of last defined column
        /// </summary>
        public ushort LastDefinedColumn { get; }

        /// <summary>
        /// Gets a value indicating whether to use the default row height instead of the RowHeight property
        /// </summary>
        public bool UseDefaultRowHeight { get; }

        /// <summary>
        /// Gets the row height in twips.
        /// </summary>
        public int RowHeight { get; }

        /// <summary>
        /// Gets a value indicating whether the XFormat property is used
        /// </summary>
        public bool UseXFormat { get; }

        /// <summary>
        /// Gets the default format for this row
        /// </summary>
        public ushort XFormat { get; }
    }
}
namespace ExcelDataReader.Core.BinaryFormat
{
    /// <summary>
    /// Represents record with the only two-bytes value
    /// </summary>
    internal class XlsBiffSimpleValueRecord : XlsBiffRecord
    {
        internal XlsBiffSimpleValueRecord(byte[] bytes)
            : base(bytes)
        {
        }

        /// <summary>
        /// Gets the value
        /// </summary>
        public ushort Value => ReadUInt16(0x0);
    }
}


namespace ExcelDataReader.Core.BinaryFormat
{
    /// <summary>
    /// Represents a Shared String Table in BIFF8 format
    /// </summary>
    internal class XlsBiffSST : XlsBiffRecord
    {
        private readonly List<IXlsString> _strings;
        private readonly XlsSSTReader _reader = new XlsSSTReader();

        internal XlsBiffSST(byte[] bytes)
            : base(bytes)
        {
            _strings = new List<IXlsString>();
            ReadSstStrings();
        }

        /// <summary>
        /// Gets the number of strings in SST
        /// </summary>
        public uint Count => ReadUInt32(0x0);

        /// <summary>
        /// Gets the count of unique strings in SST
        /// </summary>
        public uint UniqueCount => ReadUInt32(0x4);

        /// <summary>
        /// Parses strings out of a Continue record
        /// </summary>
        public void ReadContinueStrings(XlsBiffContinue sstContinue)
        {
            if (_strings.Count == UniqueCount)
            {
                return;
            }

            foreach (var str in _reader.ReadStringsFromContinue(sstContinue))
            {
                _strings.Add(str);

                if (_strings.Count == UniqueCount)
                {
                    break;
                }
            }
        }

        public void Flush()
        {
            var str = _reader.Flush();
            if (str != null)
            {
                _strings.Add(str);
            }
        }

        /// <summary>
        /// Returns string at specified index
        /// </summary>
        /// <param name="sstIndex">Index of string to get</param>
        /// <param name="encoding">Workbook encoding</param>
        /// <returns>string value if it was found, empty string otherwise</returns>
        public string GetString(uint sstIndex, Encoding encoding)
        {
            if (sstIndex < _strings.Count)
                return _strings[(int)sstIndex].GetValue(encoding);

            return null; // #VALUE error
        }

        /// <summary>
        /// Parses strings out of this SST record
        /// </summary>
        private void ReadSstStrings()
        {
            if (_strings.Count == UniqueCount)
            {
                return;
            }

            foreach (var str in _reader.ReadStringsFromSST(this))
            {
                _strings.Add(str);

                if (_strings.Count == UniqueCount)
                {
                    break;
                }
            }
        }
    }
}

namespace ExcelDataReader.Core.BinaryFormat
{
    /// <summary>
    /// Represents a BIFF stream
    /// </summary>
    internal class XlsBiffStream : IDisposable
    {
        public XlsBiffStream(Stream baseStream, int offset = 0, int explicitVersion = 0, string password = null, byte[] secretKey = null, EncryptionInfo encryption = null)
        {
            BaseStream = baseStream;
            Position = offset;

            var bof = Read() as XlsBiffBOF;
            if (bof != null)
            { 
                BiffVersion = explicitVersion == 0 ? GetBiffVersion(bof) : explicitVersion;
                BiffType = bof.Type;
            }

            CipherBlock = -1;
            if (secretKey != null)
            {
                SecretKey = secretKey;
                Encryption = encryption;
                Cipher = Encryption.CreateCipher();
            }
            else
            {
                var filePass = Read() as XlsBiffFilePass;
                if (filePass == null)
                    filePass = Read() as XlsBiffFilePass;

                if (filePass != null)
                {
                    Encryption = filePass.EncryptionInfo;

                    if (Encryption.VerifyPassword("VelvetSweatshop"))
                    {
                        // Magic password used for write-protected workbooks
                        password = "VelvetSweatshop";
                    }
                    else if (password == null || !Encryption.VerifyPassword(password))
                    {
                        throw new InvalidPasswordException(Errors.ErrorInvalidPassword);
                    }

                    SecretKey = Encryption.GenerateSecretKey(password);
                    Cipher = Encryption.CreateCipher();
                }
            }

            Position = offset;
        }

        public int BiffVersion { get; }

        public BIFFTYPE BiffType { get; }

        /// <summary>
        /// Gets the size of BIFF stream in bytes
        /// </summary>
        public int Size => (int)BaseStream.Length;

        /// <summary>
        /// Gets or sets the current position in BIFF stream
        /// </summary>
        public int Position { get => (int)BaseStream.Position; set => Seek(value, SeekOrigin.Begin); }

        public Stream BaseStream { get; }

        public byte[] SecretKey { get; }

        public EncryptionInfo Encryption { get; }

        public SymmetricAlgorithm Cipher { get; }

        /// <summary>
        /// Gets or sets the ICryptoTransform instance used to decrypt the current block
        /// </summary>
        public ICryptoTransform CipherTransform { get; set; }

        /// <summary>
        /// Gets or sets the current block number being decrypted with CipherTransform
        /// </summary>
        public int CipherBlock { get; set; }

        /// <summary>
        /// Sets stream pointer to the specified offset
        /// </summary>
        /// <param name="offset">Offset value</param>
        /// <param name="origin">Offset origin</param>
        public void Seek(int offset, SeekOrigin origin)
        {
            BaseStream.Seek(offset, origin);

            if (Position < 0)
                throw new ArgumentOutOfRangeException(string.Format("{0} On offset={1}", Errors.ErrorBiffIlegalBefore, offset));
            if (Position > Size)
                throw new ArgumentOutOfRangeException(string.Format("{0} On offset={1}", Errors.ErrorBiffIlegalAfter, offset));

            if (SecretKey != null)
            { 
                CreateBlockDecryptor(offset / 1024);
                AlignBlockDecryptor(offset % 1024);
            }
        }

        /// <summary>
        /// Reads record under cursor and advances cursor position to next record
        /// </summary>
        /// <returns>The record -or- null.</returns>
        public XlsBiffRecord Read()
        {
            // Minimum record size is 4
            if ((uint)Position + 4 >= Size)
                return null;

            var record = GetRecord(BaseStream);

            if (Position > Size)
            {
                record = null;
            }

            return record;
        }

        /// <summary>
        /// Returns record at specified offset
        /// </summary>
        /// <param name="stream">The stream</param>
        /// <returns>The record -or- null.</returns>
        public XlsBiffRecord GetRecord(Stream stream)
        {
            var recordOffset = (int)stream.Position;
            var header = new byte[4];
            stream.Read(header, 0, 4);

            // Does this work on a big endian system?
            var id = (BIFFRECORDTYPE)BitConverter.ToUInt16(header, 0);
            int recordSize = BitConverter.ToUInt16(header, 2);

            var bytes = new byte[4 + recordSize];
            Array.Copy(header, bytes, 4);
            stream.Read(bytes, 4, recordSize);
            
            if (SecretKey != null)
                DecryptRecord(recordOffset, id, bytes);

            int biffVersion = BiffVersion;

            switch (id)
            {
                case BIFFRECORDTYPE.BOF_V2:
                case BIFFRECORDTYPE.BOF_V3:
                case BIFFRECORDTYPE.BOF_V4:
                case BIFFRECORDTYPE.BOF:
                    return new XlsBiffBOF(bytes);
                case BIFFRECORDTYPE.EOF:
                    return new XlsBiffEof(bytes);
                case BIFFRECORDTYPE.INTERFACEHDR:
                    return new XlsBiffInterfaceHdr(bytes);

                case BIFFRECORDTYPE.SST:
                    return new XlsBiffSST(bytes);

                case BIFFRECORDTYPE.DEFAULTROWHEIGHT_V2:
                case BIFFRECORDTYPE.DEFAULTROWHEIGHT:
                    return new XlsBiffDefaultRowHeight(bytes, biffVersion);
                case BIFFRECORDTYPE.ROW_V2:
                case BIFFRECORDTYPE.ROW:
                    return new XlsBiffRow(bytes);

                case BIFFRECORDTYPE.BOOLERR:
                case BIFFRECORDTYPE.BOOLERR_OLD:
                case BIFFRECORDTYPE.BLANK:
                case BIFFRECORDTYPE.BLANK_OLD:
                    return new XlsBiffBlankCell(bytes);
                case BIFFRECORDTYPE.MULBLANK:
                    return new XlsBiffMulBlankCell(bytes);
                case BIFFRECORDTYPE.LABEL_OLD:
                case BIFFRECORDTYPE.LABEL:
                case BIFFRECORDTYPE.RSTRING:
                    return new XlsBiffLabelCell(bytes, biffVersion);
                case BIFFRECORDTYPE.LABELSST:
                    return new XlsBiffLabelSSTCell(bytes);
                case BIFFRECORDTYPE.INTEGER:
                case BIFFRECORDTYPE.INTEGER_OLD:
                    return new XlsBiffIntegerCell(bytes);
                case BIFFRECORDTYPE.NUMBER:
                case BIFFRECORDTYPE.NUMBER_OLD:
                    return new XlsBiffNumberCell(bytes);
                case BIFFRECORDTYPE.RK:
                    return new XlsBiffRKCell(bytes);
                case BIFFRECORDTYPE.MULRK:
                    return new XlsBiffMulRKCell(bytes);
                case BIFFRECORDTYPE.FORMULA:
                case BIFFRECORDTYPE.FORMULA_V3:
                case BIFFRECORDTYPE.FORMULA_V4:
                    return new XlsBiffFormulaCell(bytes, biffVersion);
                case BIFFRECORDTYPE.FORMAT_V23:
                case BIFFRECORDTYPE.FORMAT:
                    return new XlsBiffFormatString(bytes, biffVersion);
                case BIFFRECORDTYPE.STRING:
                case BIFFRECORDTYPE.STRING_OLD:
                    return new XlsBiffFormulaString(bytes, biffVersion);
                case BIFFRECORDTYPE.CONTINUE:
                    return new XlsBiffContinue(bytes);
                case BIFFRECORDTYPE.DIMENSIONS:
                case BIFFRECORDTYPE.DIMENSIONS_V2:
                    return new XlsBiffDimensions(bytes, biffVersion);
                case BIFFRECORDTYPE.BOUNDSHEET:
                    return new XlsBiffBoundSheet(bytes, biffVersion);
                case BIFFRECORDTYPE.WINDOW1:
                    return new XlsBiffWindow1(bytes);
                case BIFFRECORDTYPE.CODEPAGE:
                    return new XlsBiffSimpleValueRecord(bytes);
                case BIFFRECORDTYPE.FNGROUPCOUNT:
                    return new XlsBiffSimpleValueRecord(bytes);
                case BIFFRECORDTYPE.RECORD1904:
                    return new XlsBiffSimpleValueRecord(bytes);
                case BIFFRECORDTYPE.BOOKBOOL:
                    return new XlsBiffSimpleValueRecord(bytes);
                case BIFFRECORDTYPE.BACKUP:
                    return new XlsBiffSimpleValueRecord(bytes);
                case BIFFRECORDTYPE.HIDEOBJ:
                    return new XlsBiffSimpleValueRecord(bytes);
                case BIFFRECORDTYPE.USESELFS:
                    return new XlsBiffSimpleValueRecord(bytes);
                case BIFFRECORDTYPE.UNCALCED:
                    return new XlsBiffUncalced(bytes);
                case BIFFRECORDTYPE.QUICKTIP:
                    return new XlsBiffQuickTip(bytes);
                case BIFFRECORDTYPE.MSODRAWING:
                    return new XlsBiffMSODrawing(bytes);
                case BIFFRECORDTYPE.FILEPASS:
                    return new XlsBiffFilePass(bytes, biffVersion);
                case BIFFRECORDTYPE.HEADER:
                case BIFFRECORDTYPE.FOOTER:
                    return new XlsBiffHeaderFooterString(bytes, biffVersion);
                case BIFFRECORDTYPE.CODENAME:
                    return new XlsBiffCodeName(bytes);
                case BIFFRECORDTYPE.XF:
                case BIFFRECORDTYPE.XF_V2:
                case BIFFRECORDTYPE.XF_V3:
                case BIFFRECORDTYPE.XF_V4:
                    return new XlsBiffXF(bytes, biffVersion);
                case BIFFRECORDTYPE.FONT:
                    return new XlsBiffFont(bytes, biffVersion);
                case BIFFRECORDTYPE.MERGECELLS:
                    return new XlsBiffMergeCells(bytes);
                case BIFFRECORDTYPE.COLINFO:
                    return new XlsBiffColInfo(bytes);
                default:
                    return new XlsBiffRecord(bytes);
            }
        }

        public void Dispose()
        {
            CipherTransform?.Dispose();
            ((IDisposable)Cipher)?.Dispose();
        }

        private int GetBiffVersion(XlsBiffBOF bof)
        {
            switch (bof.Id)
            {
                case BIFFRECORDTYPE.BOF_V2:
                    return 2;
                case BIFFRECORDTYPE.BOF_V3:
                    return 3;
                case BIFFRECORDTYPE.BOF_V4:
                    return 4;
                case BIFFRECORDTYPE.BOF:
                    if (bof.Version == 0x200)
                        return 2;
                    else if (bof.Version == 0x300)
                        return 3;
                    else if (bof.Version == 0x400)
                        return 4;
                    else if (bof.Version == 0x500 || bof.Version == 0)
                        return 5;
                    if (bof.Version == 0x600)
                        return 8;
                    break;
            }

            return 0;
        }

        /// <summary>
        /// Create an ICryptoTransform instance to decrypt a 1024-byte block
        /// </summary>
        private void CreateBlockDecryptor(int blockNumber)
        {
            CipherTransform?.Dispose();

            var blockKey = Encryption.GenerateBlockKey(blockNumber, SecretKey);
            CipherTransform = Cipher.CreateDecryptor(blockKey, null);
            CipherBlock = blockNumber;
        }

        /// <summary>
        /// Decrypt some dummy bytes to align the decryptor with the position in the current 1024-byte block
        /// </summary>
        private void AlignBlockDecryptor(int blockOffset)
        {
            var bytes = new byte[blockOffset];
            CryptoHelpers.DecryptBytes(CipherTransform, bytes);
        }

        private void DecryptRecord(int startPosition, BIFFRECORDTYPE id, byte[] bytes)
        {
            // Decrypt the last read record, find it's start offset relative to the current stream position
            int startDecrypt = 4;
            int recordSize = bytes.Length;
            switch (id)
            {
                case BIFFRECORDTYPE.BOF:
                case BIFFRECORDTYPE.FILEPASS:
                case BIFFRECORDTYPE.INTERFACEHDR:
                    startDecrypt = recordSize;
                    break;
                case BIFFRECORDTYPE.BOUNDSHEET:
                    startDecrypt += 4; // For some reason the sheet offset is not encrypted
                    break;
            }

            var position = 0;
            while (position < recordSize)
            {
                var offset = startPosition + position;
                int blockNumber = offset / 1024;
                var blockOffset = offset % 1024;

                if (blockNumber != CipherBlock)
                {
                    CreateBlockDecryptor(blockNumber);
                }

                if (Encryption.IsXor)
                {
                    // Bypass everything and hook into the XorTransform instance to set the XorArrayIndex pr record.
                    // This is a hack to use the XorTransform otherwise transparently to the other encryption methods.
                    var xorTransform = (XorManaged.XorTransform)CipherTransform;
                    xorTransform.XorArrayIndex = offset + recordSize - 4;
                }

                // Decrypt at most up to the next 1024 byte boundary
                var chunkSize = (int)Math.Min(recordSize - position, 1024 - blockOffset);
                var block = new byte[chunkSize];

                Array.Copy(bytes, position, block, 0, chunkSize);

                var decryptedblock = CryptoHelpers.DecryptBytes(CipherTransform, block);
                for (var i = 0; i < decryptedblock.Length; i++)
                {
                    if (position >= startDecrypt)
                        bytes[position] = decryptedblock[i];
                    position++;
                }
            }
        }
    }
}

namespace ExcelDataReader.Core.BinaryFormat
{
    /// <summary>
    /// If present the Calculate Message was in the status bar when Excel saved the file.
    /// This occurs if the sheet changed, the Manual calculation option was on, and the Recalculate Before Save option was off.    
    /// </summary>
    internal class XlsBiffUncalced : XlsBiffRecord
    {
        internal XlsBiffUncalced(byte[] bytes)
            : base(bytes)
        {
        }
    }
}

namespace ExcelDataReader.Core.BinaryFormat
{
    /// <summary>
    /// Represents Workbook's global window description
    /// </summary>
    internal class XlsBiffWindow1 : XlsBiffRecord
    {
        internal XlsBiffWindow1(byte[] bytes)
            : base(bytes)
        {
        }

        [Flags]
        public enum Window1Flags : ushort
        {
            Hidden = 0x1,
            Minimized = 0x2,
            
            // (Reserved) = 0x4,
            HScrollVisible = 0x8,
            VScrollVisible = 0x10,
            WorkbookTabs = 0x20
            
            // (Other bits are reserved)
        }

        /// <summary>
        /// Gets the X position of a window
        /// </summary>
        public ushort Left => ReadUInt16(0x0);

        /// <summary>
        /// Gets the Y position of a window
        /// </summary>
        public ushort Top => ReadUInt16(0x2);

        /// <summary>
        /// Gets the width of the window
        /// </summary>
        public ushort Width => ReadUInt16(0x4);

        /// <summary>
        /// Gets the height of the window
        /// </summary>
        public ushort Height => ReadUInt16(0x6);

        /// <summary>
        /// Gets the window flags
        /// </summary>
        public Window1Flags Flags => (Window1Flags)ReadUInt16(0x8);

        /// <summary>
        /// Gets the active workbook tab (zero-based)
        /// </summary>
        public ushort ActiveTab => ReadUInt16(0xA);

        /// <summary>
        /// Gets the first visible workbook tab (zero-based)
        /// </summary>
        public ushort FirstVisibleTab => ReadUInt16(0xC);

        /// <summary>
        /// Gets the number of selected workbook tabs
        /// </summary>
        public ushort SelectedTabCount => ReadUInt16(0xE);

        /// <summary>
        /// Gets the workbook tab width to horizontal scrollbar width
        /// </summary>
        public ushort TabRatio => ReadUInt16(0x10);
    }
}

namespace ExcelDataReader.Core.BinaryFormat
{
    internal class XlsBiffXF : XlsBiffRecord
    {
        internal XlsBiffXF(byte[] bytes, int biffVersion)
            : base(bytes)
        {
            switch (Id)
            {
                case BIFFRECORDTYPE.XF_V2:
                    Font = ReadByte(0);
                    Format = ReadByte(2) & 0x3F;
                    IsLocked = (ReadByte(2) & 0x40) != 0;
                    IsHidden = (ReadByte(2) & 0x80) != 0;
                    HorizontalAlignment = (HorizontalAlignment)(ReadByte(3) & 0x07);
                    ParentCellStyleXf = 0xfff;                    
                    break;
                case BIFFRECORDTYPE.XF_V3:
                    Font = ReadByte(0);
                    Format = ReadByte(1);
                    IsLocked = (ReadByte(2) & 1) != 0;
                    IsHidden = (ReadByte(2) & 2) != 0;
                    IsCellStyleXf = (ReadByte(2) & 4) != 0;
                    ParentCellStyleXf = ReadUInt16(4) >> 4;
                    HorizontalAlignment = (HorizontalAlignment)(ReadByte(4) & 0x07);
                    break;
                case BIFFRECORDTYPE.XF_V4:
                    Font = ReadByte(0);
                    Format = ReadByte(1);
                    IsLocked = (ReadByte(2) & 1) != 0;
                    IsHidden = (ReadByte(2) & 2) != 0;
                    IsCellStyleXf = (ReadByte(2) & 4) != 0;
                    ParentCellStyleXf = ReadUInt16(2) >> 4;
                    HorizontalAlignment = (HorizontalAlignment)(ReadByte(4) & 0x07);
                    break;
                default:
                    Font = ReadUInt16(0);
                    Format = ReadUInt16(2);
                    IsLocked = (ReadByte(4) & 1) != 0;
                    IsHidden = (ReadByte(4) & 2) != 0;
                    IsCellStyleXf = (ReadByte(4) & 4) != 0;
                    ParentCellStyleXf = ReadUInt16(4) >> 4;
                    HorizontalAlignment = (HorizontalAlignment)(ReadByte(6) & 0x07);
                    if (biffVersion == 8)
                    {
                        IndentLevel = ReadByte(8) & 0x0F;
                    }

                    break;
            }

            // Paren 0xfff = do not inherit any cell style XF
            if (ParentCellStyleXf == 0xfff)
            {
                ParentCellStyleXf = -1;
            }

            // The font with index 4 is omitted in all BIFF versions. This means the first four
            // fonts have zero-based indexes, and the fifth font and all following fonts are 
            // referenced with one-based indexes.
            if (Font > 4)
            {
                Font--;
            }
        }

        public int Font { get; }
        
        public int Format { get; }

        public int ParentCellStyleXf { get; }

        public bool IsCellStyleXf { get; }

        public bool IsLocked { get; }

        public bool IsHidden { get; }
        
        public int IndentLevel { get; }

        public HorizontalAlignment HorizontalAlignment { get; }
    }
}


namespace ExcelDataReader.Core.BinaryFormat
{
    /// <summary>
    /// Word-sized string, stored as single bytes with encoding from CodePage record. Used in BIFF2-5 
    /// </summary>
    internal class XlsByteString : IXlsString
    {
        private readonly byte[] _bytes;
        private readonly uint _offset;

        public XlsByteString(byte[] bytes, uint offset)
        {
            _bytes = bytes;
            _offset = offset;
        }
        
        /// <summary>
        /// Gets the number of characters in the string.
        /// </summary>
        public ushort CharacterCount => BitConverter.ToUInt16(_bytes, (int)_offset);

        /// <summary>
        /// Gets the value.
        /// </summary>
        public string GetValue(Encoding encoding)
        {
            var stringBytes = ReadArray(0x2, CharacterCount * (Helpers.IsSingleByteEncoding(encoding) ? 1 : 2));
            return encoding.GetString(stringBytes, 0, stringBytes.Length);
        }

        public byte[] ReadArray(int offset, int size)
        {
            byte[] tmp = new byte[size];
            Buffer.BlockCopy(_bytes, (int)(_offset + offset), tmp, 0, size);
            return tmp;
        }
    }
}

namespace ExcelDataReader.Core.BinaryFormat
{
    /// <summary>
    /// Plain string without backing storage. Used internally
    /// </summary>
    internal class XlsInternalString : IXlsString
    {
        private readonly string stringValue;

        public XlsInternalString(string value)
        {
            stringValue = value;
        }

        public string GetValue(Encoding encoding)
        {
            return stringValue;
        }
    }
}


namespace ExcelDataReader.Core.BinaryFormat
{
    /// <summary>
    /// Byte sized string, stored as bytes, with encoding from CodePage record. Used in BIFF2-5 .
    /// </summary>
    internal class XlsShortByteString : IXlsString
    {
        private readonly byte[] _bytes;
        private readonly uint _offset;

        public XlsShortByteString(byte[] bytes, uint offset)
        {
            _bytes = bytes;
            _offset = offset;
        }

        public ushort CharacterCount => _bytes[_offset];

        public string GetValue(Encoding encoding)
        {
            // Supposedly this is never multibyte, but technically could be
            if (!Helpers.IsSingleByteEncoding(encoding))
            {
                return encoding.GetString(_bytes, (int)_offset + 1, CharacterCount * 2);
            }

            return encoding.GetString(_bytes, (int)_offset + 1, CharacterCount);
        }
    }
}


namespace ExcelDataReader.Core.BinaryFormat
{
    /// <summary>
    /// [MS-XLS] 2.5.240 ShortXLUnicodeString
    /// Byte-sized string, stored as single or multibyte unicode characters.
    /// </summary>
    internal class XlsShortUnicodeString : IXlsString
    {
        private readonly byte[] _bytes;
        private readonly uint _offset;

        public XlsShortUnicodeString(byte[] bytes, uint offset)
        {
            _bytes = bytes;
            _offset = offset;
        }

        public ushort CharacterCount => _bytes[_offset];

        /// <summary>
        /// Gets a value indicating whether the string is a multibyte string or not.
        /// </summary>
        public bool IsMultiByte => (_bytes[_offset + 1] & 0x01) != 0;

        public string GetValue(Encoding encoding)
        {
            if (CharacterCount == 0)
            {
                return string.Empty;
            }

            if (IsMultiByte)
            {
                return Encoding.Unicode.GetString(_bytes, (int)_offset + 2, CharacterCount * 2);
            }

            byte[] bytes = new byte[CharacterCount * 2];
            for (int i = 0; i < CharacterCount; i++)
            {
                bytes[i * 2] = _bytes[_offset + 2 + i];
            }

            return Encoding.Unicode.GetString(bytes);
        }
    }
}


namespace ExcelDataReader.Core.BinaryFormat
{
    /// <summary>
    /// Helper class for parsing the BIFF8 Shared String Table (SST)
    /// </summary>
    internal class XlsSSTReader
    {
        private enum SstState
        {
            StartStringHeader,
            StringHeader,
            StringData,
            StringTail,
        }

        private XlsBiffRecord CurrentRecord { get; set; }

        /// <summary>
        /// Gets or sets the offset into the current record's byte content. May point at the end when the current record has been parsed entirely.
        /// </summary>
        private int CurrentRecordOffset { get; set; }

        private SstState CurrentState { get; set; } = SstState.StartStringHeader;

        private XlsSSTStringHeader CurrentHeader { get; set; }

        private int CurrentRemainingCharacters { get; set; }

        private byte[] CurrentResult { get; set; }

        private int CurrentResultOffset { get; set; }

        private int CurrentHeaderBytes { get; set; }

        private int CurrentTailBytes { get; set; }

        private bool CurrentIsMultiByte { get; set; } = false;

        public IEnumerable<IXlsString> ReadStringsFromSST(XlsBiffSST sst)
        {
            CurrentRecord = sst;
            CurrentRecordOffset = 4 + 8;

            while (true)
            {
                if (!TryReadString(out var result))
                {
                    yield break;
                }

                yield return result;
            }
        }

        public IEnumerable<IXlsString> ReadStringsFromContinue(XlsBiffContinue sstContinue)
        {
            CurrentRecord = sstContinue;
            CurrentRecordOffset = 4; // +4 skips BIFF header

            if (sstContinue.Size - CurrentRecordOffset == 0)
            {
                yield break;
            }

            if (CurrentState == SstState.StringData)
            {
                byte b = ReadByte();
                CurrentIsMultiByte = b != 0;
            }

            while (true)
            {
                if (!TryReadString(out var result))
                {
                    yield break;
                }

                yield return result;
            }
        }

        public IXlsString Flush()
        {
            if (CurrentState == SstState.StringTail)
            {
                return new XlsUnicodeString(CurrentResult, 0);
            }

            return null;
        }

        private bool TryReadString(out IXlsString result)
        {
            if (CurrentState == SstState.StartStringHeader)
            {
                if (CurrentRecord.Size - CurrentRecordOffset == 0)
                {
                    result = null;
                    return false;
                }

                CurrentHeader = new XlsSSTStringHeader(CurrentRecord.Bytes, CurrentRecordOffset);
                CurrentIsMultiByte = CurrentHeader.IsMultiByte;
                CurrentHeaderBytes = (int)CurrentHeader.HeadSize;
                CurrentRemainingCharacters = (int)CurrentHeader.CharacterCount;

                const int XlsUnicodeStringHeaderSize = 3;

                CurrentResult = new byte[XlsUnicodeStringHeaderSize + CurrentRemainingCharacters * 2];
                CurrentResult[0] = (byte)(CurrentRemainingCharacters & 0x00FF);
                CurrentResult[1] = (byte)((CurrentRemainingCharacters & 0xFF00) >> 8);
                CurrentResult[2] = 1; // IsMultiByte = true

                CurrentResultOffset = XlsUnicodeStringHeaderSize;

                CurrentState = SstState.StringHeader;
            }

            if (CurrentState == SstState.StringHeader)
            {
                if (!Advance(CurrentHeaderBytes, out int advanceBytes))
                {
                    CurrentHeaderBytes -= advanceBytes;
                    result = null;
                    return false;
                }

                CurrentState = SstState.StringData;

                if (CurrentRecord.Size - CurrentRecordOffset == 0)
                {
                    // End of buffer before string data. Return false in StringData state to ensure reading the multibyte flag in the next record
                    result = null;
                    return false;
                }
            }

            if (CurrentState == SstState.StringData)
            {
                var bytesPerCharacter = CurrentIsMultiByte ? 2 : 1;
                var maxRecordCharacters = (CurrentRecord.Size - CurrentRecordOffset) / bytesPerCharacter;
                var readCharacters = Math.Min(maxRecordCharacters, CurrentRemainingCharacters);

                ReadUnicodeBytes(CurrentResult, CurrentResultOffset, readCharacters, CurrentIsMultiByte);

                CurrentResultOffset += readCharacters * 2; // The result is always multibyte
                CurrentRemainingCharacters -= readCharacters;

                if (CurrentIsMultiByte && CurrentRecord.Size - CurrentRecordOffset == 1)
                {
                    // Skip leftover byte at the end of a multibyte Continue record
                    ReadByte();
                }

                if (CurrentRemainingCharacters > 0 && CurrentRecord.Size - CurrentRecordOffset == 0)
                {
                    result = null;
                    return false;
                }

                CurrentState = SstState.StringTail;
                CurrentTailBytes = (int)CurrentHeader.TailSize;
            }

            if (CurrentState == SstState.StringTail)
            {
                // Skip formatting runs and phonetic/extended data. Can also span
                // multiple Continue records
                if (!Advance(CurrentTailBytes, out var advanceBytes))
                {
                    result = null;
                    CurrentTailBytes -= advanceBytes;
                    return false;
                }

                CurrentState = SstState.StartStringHeader;
                result = new XlsUnicodeString(CurrentResult, 0);
                return true;
            }

            throw new InvalidOperationException("Unexpected state in SST reader");
        }

        private void ReadUnicodeBytes(byte[] dest, int offset, int characterCount, bool isMultiByte)
        {
            if (isMultiByte)
            {
                Array.Copy(CurrentRecord.Bytes, CurrentRecordOffset, dest, offset, characterCount * 2);
                CurrentRecordOffset += characterCount * 2;
            }
            else
            {
                for (int i = 0; i < characterCount; i++)
                {
                    dest[offset + i * 2] = CurrentRecord.Bytes[CurrentRecordOffset + i];
                    dest[offset + i * 2 + 1] = 0;
                }

                CurrentRecordOffset += characterCount;
            }
        }

        private byte ReadByte()
        {
            if (CurrentRecordOffset >= CurrentRecord.Size)
            {
                throw new InvalidOperationException("SST read position out of range");
            }

            var result = CurrentRecord.Bytes[CurrentRecordOffset];
            CurrentRecordOffset++;
            return result;
        }

        private bool Advance(int bytes, out int advanceBytes)
        {
            advanceBytes = Math.Min(CurrentRecord.Size - CurrentRecordOffset, bytes);
            CurrentRecordOffset += advanceBytes;
            return bytes == advanceBytes;
        }
    }
}


namespace ExcelDataReader.Core.BinaryFormat
{
    /// <summary>
    /// [MS-XLS] 2.5.293 XLUnicodeRichExtendedString
    /// Word-sized formatted string in SST, stored as single or multibyte unicode characters potentially spanning multiple Continue records.
    /// </summary>
    internal class XlsSSTStringHeader
    {
        private readonly byte[] _bytes;
        private readonly int _offset;

        public XlsSSTStringHeader(byte[] bytes, int offset)
        {
            _bytes = bytes;
            _offset = offset;
        }

        [Flags]
        public enum FormattedUnicodeStringFlags : byte
        {
            MultiByte = 0x01,
            HasExtendedString = 0x04,
            HasFormatting = 0x08,
        }

        /// <summary>
        /// Gets the number of characters in the string.
        /// </summary>
        public ushort CharacterCount => BitConverter.ToUInt16(_bytes, _offset);

        /// <summary>
        /// Gets the flags.
        /// </summary>
        public FormattedUnicodeStringFlags Flags => (FormattedUnicodeStringFlags)Buffer.GetByte(_bytes, _offset + 2);

        /// <summary>
        /// Gets a value indicating whether the string has an extended record. 
        /// </summary>
        public bool HasExtString => (Flags & FormattedUnicodeStringFlags.HasExtendedString) == FormattedUnicodeStringFlags.HasExtendedString;

        /// <summary>
        /// Gets a value indicating whether the string has a formatting record.
        /// </summary>
        public bool HasFormatting => (Flags & FormattedUnicodeStringFlags.HasFormatting) == FormattedUnicodeStringFlags.HasFormatting;

        /// <summary>
        /// Gets a value indicating whether the string is a multibyte string or not.
        /// </summary>
        public bool IsMultiByte => (Flags & FormattedUnicodeStringFlags.MultiByte) == FormattedUnicodeStringFlags.MultiByte;

        /// <summary>
        /// Gets the number of formats used for formatting (0 if string has no formatting)
        /// </summary>
        public ushort FormatCount => HasFormatting ? BitConverter.ToUInt16(_bytes, (int)_offset + 3) : (ushort)0;

        /// <summary>
        /// Gets the size of extended string in bytes, 0 if there is no one
        /// </summary>
        public uint ExtendedStringSize => HasExtString ? (uint)BitConverter.ToUInt32(_bytes, (int)_offset + (HasFormatting ? 5 : 3)) : 0;

        /// <summary>
        /// Gets the head (before string data) size in bytes
        /// </summary>
        public uint HeadSize => (uint)(HasFormatting ? 2 : 0) + (uint)(HasExtString ? 4 : 0) + 3;

        /// <summary>
        /// Gets the tail (after string data) size in bytes
        /// </summary>
        public uint TailSize => (uint)(HasFormatting ? 4 * FormatCount : 0) + (HasExtString ? ExtendedStringSize : 0);
    }
}

namespace ExcelDataReader.Core.BinaryFormat
{
    /// <summary>
    /// [MS-XLS] 2.5.294 XLUnicodeString
    /// Word-sized string, stored as single or multibyte unicode characters.
    /// </summary>
    internal class XlsUnicodeString : IXlsString
    {
        private readonly byte[] _bytes;
        private readonly uint _offset;

        public XlsUnicodeString(byte[] bytes, uint offset)
        {
            _bytes = bytes;
            _offset = offset;
        }

        public ushort CharacterCount => BitConverter.ToUInt16(_bytes, (int)_offset);

        /// <summary>
        /// Gets a value indicating whether the string is a multibyte string or not.
        /// </summary>
        public bool IsMultiByte => (_bytes[_offset + 2] & 0x01) != 0;

        public string GetValue(Encoding encoding)
        {
            if (CharacterCount == 0)
            {
                return string.Empty;
            }

            if (IsMultiByte)
            {
                return Encoding.Unicode.GetString(_bytes, (int)_offset + 3, CharacterCount * 2);
            }

            byte[] bytes = new byte[CharacterCount * 2];
            for (int i = 0; i < CharacterCount; i++)
            {
                bytes[i * 2] = _bytes[_offset + 3 + i];
            }

            return Encoding.Unicode.GetString(bytes);
        }
    }
}


namespace ExcelDataReader.Core.BinaryFormat
{
    /// <summary>
    /// Represents Globals section of workbook
    /// </summary>
    internal class XlsWorkbook : CommonWorkbook, IWorkbook<XlsWorksheet>
    {
        internal XlsWorkbook(Stream stream, string password, Encoding fallbackEncoding)
        {
            Stream = stream;

            using (var biffStream = new XlsBiffStream(stream, 0, 0, password))
            {
                if (biffStream.BiffVersion == 0)
                    throw new ExcelReaderException(Errors.ErrorWorkbookGlobalsInvalidData);

                BiffVersion = biffStream.BiffVersion;
                SecretKey = biffStream.SecretKey;
                Encryption = biffStream.Encryption;
                Encoding = biffStream.BiffVersion == 8 ? Encoding.Unicode : fallbackEncoding;

                if (biffStream.BiffType == BIFFTYPE.WorkbookGlobals)
                {
                    ReadWorkbookGlobals(biffStream);
                }
                else if (biffStream.BiffType == BIFFTYPE.Worksheet)
                {
                    // set up 'virtual' bound sheet pointing at this
                    Sheets.Add(new XlsBiffBoundSheet(0, XlsBiffBoundSheet.SheetType.Worksheet, XlsBiffBoundSheet.SheetVisibility.Visible, "Sheet"));
                }
                else
                {
                    throw new ExcelReaderException(Errors.ErrorWorkbookGlobalsInvalidData);
                }
            }
        }

        public Stream Stream { get; }

        public int BiffVersion { get; }

        public byte[] SecretKey { get; }

        public EncryptionInfo Encryption { get; }

        public Encoding Encoding { get; private set; }

        public XlsBiffInterfaceHdr InterfaceHdr { get; set; }

        public XlsBiffRecord Mms { get; set; }

        public XlsBiffRecord WriteAccess { get; set; }

        public XlsBiffSimpleValueRecord CodePage { get; set; }

        public XlsBiffRecord Dsf { get; set; }

        public XlsBiffRecord Country { get; set; }

        public XlsBiffSimpleValueRecord Backup { get; set; }

        public List<XlsBiffFont> Fonts { get; } = new List<XlsBiffFont>();

        public List<XlsBiffBoundSheet> Sheets { get; } = new List<XlsBiffBoundSheet>();

        /// <summary>
        /// Gets or sets the Shared String Table of workbook
        /// </summary>
        public XlsBiffSST SST { get; set; }

        public XlsBiffRecord ExtSST { get; set; }

        public bool IsDate1904 { get; private set; }

        public int ResultsCount => Sheets?.Count ?? -1;

        public static bool IsRawBiffStream(byte[] bytes)
        {
            if (bytes.Length < 8)
            {
                throw new ArgumentException("Needs at least 8 bytes to probe", nameof(bytes));
            }

            ushort rid = BitConverter.ToUInt16(bytes, 0);
            ushort size = BitConverter.ToUInt16(bytes, 2);
            ushort bofVersion = BitConverter.ToUInt16(bytes, 4);
            ushort type = BitConverter.ToUInt16(bytes, 6);

            switch (rid)
            {
                case 0x0009: // BIFF2
                    if (size != 4)
                        return false;
                    if (type != 0x10 && type != 0x20 && type != 0x40)
                        return false;
                    return true;
                case 0x0209: // BIFF3
                case 0x0409: // BIFF4
                    if (size != 6)
                        return false;
                    if (type != 0x10 && type != 0x20 && type != 0x40 && type != 0x0100)
                        return false;
                    /* removed this additional check to keep the probe at 8 bytes
                    ushort notUsed = BitConverter.ToUInt16(bytes, 8);
                    if (notUsed != 0x00)
                        return false;*/
                    return true;
                case 0x0809: // BIFF5/BIFF8
                    if (size < 4)
                        return false;
                    if (bofVersion != 0 && bofVersion != 0x0200 && bofVersion != 0x0300 && bofVersion != 0x0400 && bofVersion != 0x0500 && bofVersion != 0x600)
                        return false;
                    if (type != 0x5 && type != 0x6 && type != 0x10 && type != 0x20 && type != 0x40 && type != 0x0100)
                        return false;
                    /* removed this additional check to keep the probe at 8 bytes
                    ushort identifier = BitConverter.ToUInt16(bytes, 10);
                    if (identifier == 0)
                        return false;*/
                    return true;
            }

            return false;
        }

        public IEnumerable<XlsWorksheet> ReadWorksheets()
        {
            for (var i = 0; i < Sheets.Count; ++i)
            {
                yield return new XlsWorksheet(this, Sheets[i], Stream);
            }
        }

        internal void AddXf(XlsBiffXF xf)
        {
            var extendedFormat = new ExtendedFormat()
            {
                FontIndex = xf.Font,
                NumberFormatIndex = xf.Format,
                Locked = xf.IsLocked,
                Hidden = xf.IsHidden,
                HorizontalAlignment = xf.HorizontalAlignment,
                IndentLevel = xf.IndentLevel,
                ParentCellStyleXf = xf.ParentCellStyleXf,
            };

            // The workbook holds two kinds of XF records: Cell XFs, and Cell Style XFs.
            // In the binary XLS format, both kinds of XF records are saved in a single list,
            // whereas the XLSX format has two separate lists - like the CommonWorkbook internals.
            // The Cell XFs hold indexes into the Cell Style XF list, so adding the XF in both lists 
            // here to keep the indexes the same.
            ExtendedFormats.Add(extendedFormat);
            CellStyleExtendedFormats.Add(extendedFormat);
        }

        private void ReadWorkbookGlobals(XlsBiffStream biffStream)
        {
            var formats = new Dictionary<int, XlsBiffFormatString>();

            XlsBiffRecord rec;
            while ((rec = biffStream.Read()) != null && !(rec is XlsBiffEof))
            {
                switch (rec)
                {
                    case XlsBiffInterfaceHdr hdr:
                        InterfaceHdr = hdr;
                        break;
                    case XlsBiffBoundSheet sheet:
                        if (sheet.Type != XlsBiffBoundSheet.SheetType.Worksheet)
                            break;
                        Sheets.Add(sheet);
                        break;
                    case XlsBiffSimpleValueRecord codePage when rec.Id == BIFFRECORDTYPE.CODEPAGE:
                        // [MS-XLS 2.4.52 CodePage] An unsigned integer that specifies the workbook�s code page.The value MUST be one
                        // of the code page values specified in [CODEPG] or the special value 1200, which means that the
                        // workbook is Unicode.
                        CodePage = codePage;
                        Encoding = EncodingHelper.GetEncoding(CodePage.Value);
                        break;
                    case XlsBiffSimpleValueRecord is1904 when rec.Id == BIFFRECORDTYPE.RECORD1904:
                        IsDate1904 = is1904.Value == 1;
                        break;
                    case XlsBiffFont font:
                        Fonts.Add(font);
                        break;
                    case XlsBiffFormatString format23 when rec.Id == BIFFRECORDTYPE.FORMAT_V23:
                        formats.Add((ushort)formats.Count, format23);
                        break;
                    case XlsBiffFormatString fmt when rec.Id == BIFFRECORDTYPE.FORMAT:
                        var index = fmt.Index;
                        if (!formats.ContainsKey(index))
                            formats.Add(index, fmt);
                        break;
                    case XlsBiffXF xf:
                        AddXf(xf);
                        break;
                    case XlsBiffSST sst:
                        SST = sst;
                        break;
                    case XlsBiffContinue sstContinue:
                        if (SST != null)
                        {
                            SST.ReadContinueStrings(sstContinue);
                        }

                        break;
                    case XlsBiffRecord _ when rec.Id == BIFFRECORDTYPE.MMS:
                        Mms = rec;
                        break;
                    case XlsBiffRecord _ when rec.Id == BIFFRECORDTYPE.COUNTRY:
                        Country = rec;
                        break;
                    case XlsBiffRecord _ when rec.Id == BIFFRECORDTYPE.EXTSST:
                        ExtSST = rec;
                        break;

                    // case BIFFRECORDTYPE.PROTECT:
                    // case BIFFRECORDTYPE.PROT4REVPASSWORD:
                        // IsProtected
                        // break;
                    // case BIFFRECORDTYPE.PASSWORD:
                    default:
                        break;
                }
            }

            if (SST != null)
            {
                SST.Flush();
            }

            foreach (var format in formats)
            {
                // We don't decode the value until here in-case there are format records before the 
                // codepage record. 
                Formats.Add(format.Key, new NumberFormatString(format.Value.GetValue(Encoding)));
            }
        }
    }
}


namespace ExcelDataReader.Core.BinaryFormat
{
    /// <summary>
    /// Represents Worksheet section in workbook
    /// </summary>
    internal class XlsWorksheet : IWorksheet
    {
        public XlsWorksheet(XlsWorkbook workbook, XlsBiffBoundSheet refSheet, Stream stream)
        {
            Workbook = workbook;
            Stream = stream;

            IsDate1904 = workbook.IsDate1904;
            Encoding = workbook.Encoding;
            RowOffsetMap = new Dictionary<int, XlsRowOffset>();
            DefaultRowHeight = 255; // 12.75 points

            Name = refSheet.GetSheetName(workbook.Encoding);
            DataOffset = refSheet.StartOffset;

            VisibleState = refSheet.VisibleState switch
            {
                XlsBiffBoundSheet.SheetVisibility.Hidden => "hidden",
                XlsBiffBoundSheet.SheetVisibility.VeryHidden => "veryhidden",
                _ => "visible",
            };
            ReadWorksheetGlobals();
        }

        /// <summary>
        /// Gets the worksheet name
        /// </summary>
        public string Name { get; }

        public string CodeName { get; private set; }

        /// <summary>
        /// Gets the visibility of worksheet
        /// </summary>
        public string VisibleState { get; }

        public HeaderFooter HeaderFooter { get; private set; }

        public CellRange[] MergeCells { get; private set; }

        public Column[] ColumnWidths { get; private set; }

        /// <summary>
        /// Gets the worksheet data offset.
        /// </summary>
        public uint DataOffset { get; }

        public Stream Stream { get; }

        public Encoding Encoding { get; private set; }

        public double DefaultRowHeight { get; set; }

        public Dictionary<int, XlsRowOffset> RowOffsetMap { get; }

        /*
            TODO: populate these in ReadWorksheetGlobals() if needed
                public XlsBiffSimpleValueRecord CalcMode { get; set; }

                public XlsBiffSimpleValueRecord CalcCount { get; set; }

                public XlsBiffSimpleValueRecord RefMode { get; set; }

                public XlsBiffSimpleValueRecord Iteration { get; set; }

                public XlsBiffRecord Delta { get; set; }

                public XlsBiffRecord Window { get; set; }
        */

        public int FieldCount { get; private set; }

        public int RowCount { get; private set; }

        public bool IsDate1904 { get; private set; }

        public XlsWorkbook Workbook { get; }

        public IEnumerable<Row> ReadRows()
        {
            var rowIndex = 0;
            using (var biffStream = new XlsBiffStream(Stream, (int)DataOffset, Workbook.BiffVersion, null, Workbook.SecretKey, Workbook.Encryption))
            {
                foreach (var rowBlock in ReadWorksheetRows(biffStream))
                {
                    for (; rowIndex < rowBlock.RowIndex; ++rowIndex)
                    {
                        yield return new Row(rowIndex, DefaultRowHeight / 20.0, new List<Cell>());
                    }

                    rowIndex++;
                    yield return rowBlock;
                }
            }
        }

        /// <summary>
        /// Find how many rows to read at a time and their offset in the file.
        /// If rows are stored sequentially in the file, returns a block size of up to 32 rows.
        /// If rows are stored non-sequentially, the block size may extend up to the entire worksheet stream
        /// </summary>
        private void GetBlockSize(int startRow, out int blockRowCount, out int minOffset, out int maxOffset)
        {
            minOffset = int.MaxValue;
            maxOffset = int.MinValue;

            var i = 0;
            blockRowCount = Math.Min(32, RowCount - startRow);

            while (i < blockRowCount)
            {
                if (RowOffsetMap.TryGetValue(startRow + i, out var rowOffset))
                {
                    minOffset = Math.Min(rowOffset.MinCellOffset, minOffset);
                    maxOffset = Math.Max(rowOffset.MaxCellOffset, maxOffset);

                    if (rowOffset.MaxOverlapRowIndex != int.MinValue)
                    {
                        var maxOverlapRowCount = rowOffset.MaxOverlapRowIndex + 1;
                        blockRowCount = Math.Max(blockRowCount, maxOverlapRowCount - startRow);
                    }
                }

                i++;
            }
        }

        private IEnumerable<Row> ReadWorksheetRows(XlsBiffStream biffStream)
        {
            var rowIndex = 0;

            while (rowIndex < RowCount)
            {
                GetBlockSize(rowIndex, out var blockRowCount, out var minOffset, out var maxOffset);

                var block = ReadNextBlock(biffStream, rowIndex, blockRowCount, minOffset, maxOffset);

                for (var i = 0; i < blockRowCount; ++i)
                {
                    if (block.Rows.TryGetValue(rowIndex + i, out var row))
                    {
                        yield return row;
                    }
                }

                rowIndex += blockRowCount;
            }
        }

        private XlsRowBlock ReadNextBlock(XlsBiffStream biffStream, int startRow, int rows, int minOffset, int maxOffset)
        {
            var result = new XlsRowBlock { Rows = new Dictionary<int, Row>() };

            // Ensure rows with physical records are initialized with height
            for (var i = 0; i < rows; i++)
            {
                if (RowOffsetMap.TryGetValue(startRow + i, out _))
                {
                    EnsureRow(result, startRow + i);
                }
            }

            if (minOffset == int.MaxValue)
            {
                return result;
            }

            biffStream.Position = minOffset;

            XlsBiffRecord rec;
            XlsBiffRecord ixfe = null;
            while (biffStream.Position <= maxOffset && (rec = biffStream.Read()) != null)
            {
                if (rec.Id == BIFFRECORDTYPE.IXFE)
                {
                    // BIFF2: If cell.xformat == 63, this contains the actual XF index >= 63
                    ixfe = rec;
                }

                if (rec is XlsBiffBlankCell cell)
                {
                    var currentRow = EnsureRow(result, cell.RowIndex);

                    if (cell.Id == BIFFRECORDTYPE.MULRK)
                    {
                        var cellValues = ReadMultiCell(cell);
                        currentRow.Cells.AddRange(cellValues);
                    }
                    else
                    {
                        var xfIndex = GetXfIndexForCell(cell, ixfe);
                        var cellValue = ReadSingleCell(biffStream, cell, xfIndex);
                        currentRow.Cells.Add(cellValue);
                    }

                    ixfe = null;
                }
            }

            return result;
        }

        private Row EnsureRow(XlsRowBlock result, int rowIndex)
        {
            if (!result.Rows.TryGetValue(rowIndex, out var currentRow))
            {
                var height = DefaultRowHeight / 20.0;
                if (RowOffsetMap.TryGetValue(rowIndex, out var rowOffset) && rowOffset.Record != null)
                {
                    height = (rowOffset.Record.UseDefaultRowHeight ? DefaultRowHeight : rowOffset.Record.RowHeight) / 20.0;
                }

                currentRow = new Row(rowIndex, height, new List<Cell>());

                result.Rows.Add(rowIndex, currentRow);
            }

            return currentRow;
        }

        private IEnumerable<Cell> ReadMultiCell(XlsBiffBlankCell cell)
        {
            LogManager.Log(this).Debug("ReadMultiCell {0}", cell.Id);

            switch (cell.Id)
            {
                case BIFFRECORDTYPE.MULRK:

                    XlsBiffMulRKCell rkCell = (XlsBiffMulRKCell)cell;
                    ushort lastColumnIndex = rkCell.LastColumnIndex;
                    for (ushort j = cell.ColumnIndex; j <= lastColumnIndex; j++)
                    {
                        var xfIndex = rkCell.GetXF(j);
                        var effectiveStyle = Workbook.GetEffectiveCellStyle(xfIndex, cell.Format);

                        var value = TryConvertOADateTime(rkCell.GetValue(j), effectiveStyle.NumberFormatIndex);
                        LogManager.Log(this).Debug("CELL[{0}] = {1}", j, value);
                        yield return new Cell(j, value, effectiveStyle, null);
                    }

                    break;
            }
        }

        /// <summary>
        /// Reads additional records if needed: a string record might follow a formula result
        /// </summary>
        private Cell ReadSingleCell(XlsBiffStream biffStream, XlsBiffBlankCell cell, int xfIndex)
        {
            LogManager.Log(this).Debug("ReadSingleCell {0}", cell.Id);

            var effectiveStyle = Workbook.GetEffectiveCellStyle(xfIndex, cell.Format);
            var numberFormatIndex = effectiveStyle.NumberFormatIndex;

            object value = null;
            CellError? error = null;
            switch (cell.Id)
            {
                case BIFFRECORDTYPE.BOOLERR:
                    if (cell.ReadByte(7) == 0)
                        value = cell.ReadByte(6) != 0;
                    else
                        error = (CellError)cell.ReadByte(6);
                    break;
                case BIFFRECORDTYPE.BOOLERR_OLD:
                    if (cell.ReadByte(8) == 0)
                        value = cell.ReadByte(7) != 0;
                    else
                        error = (CellError)cell.ReadByte(7);
                    break;
                case BIFFRECORDTYPE.INTEGER:
                case BIFFRECORDTYPE.INTEGER_OLD:
                    value = TryConvertOADateTime(((XlsBiffIntegerCell)cell).Value, numberFormatIndex);
                    break;
                case BIFFRECORDTYPE.NUMBER:
                case BIFFRECORDTYPE.NUMBER_OLD:
                    value = TryConvertOADateTime(((XlsBiffNumberCell)cell).Value, numberFormatIndex);
                    break;
                case BIFFRECORDTYPE.LABEL:
                case BIFFRECORDTYPE.LABEL_OLD:
                case BIFFRECORDTYPE.RSTRING:
                    value = GetLabelString((XlsBiffLabelCell)cell, effectiveStyle);
                    break;
                case BIFFRECORDTYPE.LABELSST:
                    value = Workbook.SST.GetString(((XlsBiffLabelSSTCell)cell).SSTIndex, Encoding);
                    break;
                case BIFFRECORDTYPE.RK:
                    value = TryConvertOADateTime(((XlsBiffRKCell)cell).Value, numberFormatIndex);
                    break;
                case BIFFRECORDTYPE.BLANK:
                case BIFFRECORDTYPE.BLANK_OLD:
                case BIFFRECORDTYPE.MULBLANK:
                    // Skip blank cells
                    break;
                case BIFFRECORDTYPE.FORMULA:
                case BIFFRECORDTYPE.FORMULA_V3:
                case BIFFRECORDTYPE.FORMULA_V4:
                    value = TryGetFormulaValue(biffStream, (XlsBiffFormulaCell)cell, effectiveStyle, out error);
                    break;
            }

            return new Cell(cell.ColumnIndex, value, effectiveStyle, error);
        }

        private string GetLabelString(XlsBiffLabelCell cell, ExtendedFormat effectiveStyle)
        {
            // 1. Use encoding from font's character set (BIFF5-8)
            // 2. If not specified, use encoding from CODEPAGE BIFF record
            // 3. If not specified, use configured fallback encoding
            // Encoding is only used on BIFF2-5 byte strings. BIFF8 uses XlsUnicodeString which ignores the encoding.
            var labelEncoding = GetFont(effectiveStyle.FontIndex)?.ByteStringEncoding ?? Encoding;
            return cell.GetValue(labelEncoding);
        }

        private XlsBiffFont GetFont(int fontIndex)
        {
            if (fontIndex < 0 || fontIndex >= Workbook.Fonts.Count)
            {
                return null;
            }

            return Workbook.Fonts[fontIndex];
        }

        private object TryGetFormulaValue(XlsBiffStream biffStream, XlsBiffFormulaCell formulaCell, ExtendedFormat effectiveStyle, out CellError? error)
        {
            error = null;
            switch (formulaCell.FormulaType)
            {
                case XlsBiffFormulaCell.FormulaValueType.Boolean: return formulaCell.BooleanValue;
                case XlsBiffFormulaCell.FormulaValueType.Error:
                    error = (CellError)formulaCell.ErrorValue;
                    return null;
                case XlsBiffFormulaCell.FormulaValueType.EmptyString: return string.Empty;
                case XlsBiffFormulaCell.FormulaValueType.Number: return TryConvertOADateTime(formulaCell.XNumValue, effectiveStyle.NumberFormatIndex);
                case XlsBiffFormulaCell.FormulaValueType.String: return TryGetFormulaString(biffStream, effectiveStyle);

                // Bad data or new formula value type
                default: return null;
            }
        }

        private string TryGetFormulaString(XlsBiffStream biffStream, ExtendedFormat effectiveStyle)
        {
            var rec = biffStream.Read();
            if (rec != null && rec.Id == BIFFRECORDTYPE.SHAREDFMLA)
            {
                rec = biffStream.Read();
            }

            if (rec != null && rec.Id == BIFFRECORDTYPE.STRING)
            {
                var stringRecord = (XlsBiffFormulaString)rec;
                var formulaEncoding = GetFont(effectiveStyle.FontIndex)?.ByteStringEncoding ?? Encoding; // Workbook.GetFontEncodingFromXF(xFormat) ?? Encoding;
                return stringRecord.GetValue(formulaEncoding);
            }

            // Bad data - could not find a string following the formula
            return null;
        }

        private object TryConvertOADateTime(double value, int numberFormatIndex)
        {
            var format = Workbook.GetNumberFormatString(numberFormatIndex);
            if (format != null)
            {
                if (format.IsDateTimeFormat)
                    return Helpers.ConvertFromOATime(value, IsDate1904);
                if (format.IsTimeSpanFormat)
                    return TimeSpan.FromDays(value);
            }

            return value;
        }

        private object TryConvertOADateTime(int value, int numberFormatIndex)
        {
            var format = Workbook.GetNumberFormatString(numberFormatIndex);
            if (format != null)
            {
                if (format.IsDateTimeFormat)
                    return Helpers.ConvertFromOATime(value, IsDate1904);
                if (format.IsTimeSpanFormat)
                    return TimeSpan.FromDays(value);
            }

            return value;
        }

        /// <summary>
        /// Returns an index into Workbook.ExtendedFormats for the given cell and preceding ixfe record.
        /// </summary>
        private int GetXfIndexForCell(XlsBiffBlankCell cell, XlsBiffRecord ixfe)
        {
            if (Workbook.BiffVersion == 2)
            {
                if (cell.XFormat == 63 && ixfe != null)
                {
                    var xFormat = ixfe.ReadUInt16(0);
                    return xFormat;
                }
                else if (cell.XFormat > 63)
                {
                    // Invalid XF ref on cell in BIFF2 stream
                    return -1;
                }
                else if (cell.XFormat < Workbook.ExtendedFormats.Count)
                {
                    return cell.XFormat;
                }
                else
                {
                    // Either the file has no XFs, or XF was out of range
                    return -1;
                }
            }

            return cell.XFormat;
        }

        private void ReadWorksheetGlobals()
        {
            using (var biffStream = new XlsBiffStream(Stream, (int)DataOffset, Workbook.BiffVersion, null, Workbook.SecretKey, Workbook.Encryption))
            {
                // Check the expected BOF record was found in the BIFF stream
                if (biffStream.BiffVersion == 0 || biffStream.BiffType != BIFFTYPE.Worksheet)
                    return;

                XlsBiffHeaderFooterString header = null;
                XlsBiffHeaderFooterString footer = null;
                
                var ixfeOffset = -1;

                int maxCellColumn = 0;
                int maxRowCount = 0; // number of rows with cell records
                int maxRowCountFromRowRecord = 0; // number of rows with row records

                var mergeCells = new List<CellRange>();
                var biffFormats = new Dictionary<ushort, XlsBiffFormatString>();
                var recordOffset = biffStream.Position;
                var rec = biffStream.Read();
                var columnWidths = new List<Column>();

                while (rec != null && !(rec is XlsBiffEof))
                {
                    switch (rec)
                    {
                        case XlsBiffDimensions dims:
                            FieldCount = dims.LastColumn;
                            RowCount = (int)dims.LastRow;
                            break;
                        case XlsBiffDefaultRowHeight defaultRowHeightRecord:
                            DefaultRowHeight = defaultRowHeightRecord.RowHeight;
                            break;
                        case XlsBiffSimpleValueRecord is1904 when rec.Id == BIFFRECORDTYPE.RECORD1904:
                            IsDate1904 = is1904.Value == 1;
                            break;
                        case XlsBiffXF xf when rec.Id == BIFFRECORDTYPE.XF_V2 || rec.Id == BIFFRECORDTYPE.XF_V3 || rec.Id == BIFFRECORDTYPE.XF_V4:
                            // NOTE: XF records should only occur in raw BIFF2-4 single worksheet documents without the workbook stream, or globally in the workbook stream.
                            // It is undefined behavior if multiple worksheets in a workbook declare XF records.
                            Workbook.AddXf(xf);
                            break;
                        case XlsBiffMergeCells mc:
                            mergeCells.AddRange(mc.MergeCells);
                            break;
                        case XlsBiffColInfo colInfo:
                            columnWidths.Add(colInfo.Value);
                            break;
                        case XlsBiffFormatString fmt when rec.Id == BIFFRECORDTYPE.FORMAT:
                            if (Workbook.BiffVersion >= 5)
                            {
                                // fmt.Index exists on BIFF5+ only
                                biffFormats.Add(fmt.Index, fmt);
                            }
                            else
                            {
                                biffFormats.Add((ushort)biffFormats.Count, fmt);
                            }

                            break;

                        case XlsBiffFormatString fmt23 when rec.Id == BIFFRECORDTYPE.FORMAT_V23:
                            biffFormats.Add((ushort)biffFormats.Count, fmt23);
                            break;
                        case XlsBiffSimpleValueRecord codePage when rec.Id == BIFFRECORDTYPE.CODEPAGE:
                            Encoding = EncodingHelper.GetEncoding(codePage.Value);
                            break;
                        case XlsBiffHeaderFooterString h when rec.Id == BIFFRECORDTYPE.HEADER && rec.RecordSize > 0:
                            header = h;
                            break;
                        case XlsBiffHeaderFooterString f when rec.Id == BIFFRECORDTYPE.FOOTER && rec.RecordSize > 0:
                            footer = f;
                            break;
                        case XlsBiffCodeName codeName:
                            CodeName = codeName.GetValue(Encoding);
                            break;
                        case XlsBiffRow row:
                            SetMinMaxRow(row.RowIndex, row);

                            // Count rows by row records without affecting the overlap in OffsetMap
                            maxRowCountFromRowRecord = Math.Max(maxRowCountFromRowRecord, row.RowIndex + 1);
                            break;
                        case XlsBiffBlankCell cell:
                            maxCellColumn = Math.Max(maxCellColumn, cell.ColumnIndex + 1);
                            maxRowCount = Math.Max(maxRowCount, cell.RowIndex + 1);
                            if (ixfeOffset != -1)
                            {
                                SetMinMaxRowOffset(cell.RowIndex, ixfeOffset, maxRowCount - 1);
                                ixfeOffset = -1;
                            }

                            SetMinMaxRowOffset(cell.RowIndex, recordOffset, maxRowCount - 1);
                            break;
                        case XlsBiffRecord ixfe when rec.Id == BIFFRECORDTYPE.IXFE:
                            ixfeOffset = recordOffset; 
                            break;
                    }

                    recordOffset = biffStream.Position;
                    rec = biffStream.Read();

                    // Stop if we find the start out a new substream. Not always that files have the required EOF before a substream BOF.
                    if (rec is XlsBiffBOF)
                        break;
                }

                if (header != null || footer != null)
                {
                    HeaderFooter = new HeaderFooter(footer?.GetValue(Encoding), header?.GetValue(Encoding));
                }

                foreach (var biffFormat in biffFormats)
                {
                    Workbook.AddNumberFormat(biffFormat.Key, biffFormat.Value.GetValue(Encoding));
                }

                if (mergeCells.Count > 0)
                    MergeCells = mergeCells.ToArray();

                if (FieldCount < maxCellColumn)
                    FieldCount = maxCellColumn;

                maxRowCount = Math.Max(maxRowCount, maxRowCountFromRowRecord);
                if (RowCount < maxRowCount)
                    RowCount = maxRowCount;

                if (columnWidths.Count > 0)
                {
                    ColumnWidths = columnWidths.ToArray();
                }
            }
        }

        private void SetMinMaxRow(int rowIndex, XlsBiffRow row)
        {
            if (!RowOffsetMap.TryGetValue(rowIndex, out var rowOffset))
            {
                rowOffset = new XlsRowOffset();
                rowOffset.MinCellOffset = int.MaxValue;
                rowOffset.MaxCellOffset = int.MinValue;
                rowOffset.MaxOverlapRowIndex = int.MinValue;
                RowOffsetMap.Add(rowIndex, rowOffset);
            }

            rowOffset.Record = row;
        }

        private void SetMinMaxRowOffset(int rowIndex, int recordOffset, int maxOverlapRow)
        {
            if (!RowOffsetMap.TryGetValue(rowIndex, out var rowOffset))
            {
                rowOffset = new XlsRowOffset();
                rowOffset.MinCellOffset = int.MaxValue;
                rowOffset.MaxCellOffset = int.MinValue;
                rowOffset.MaxOverlapRowIndex = int.MinValue;
                RowOffsetMap.Add(rowIndex, rowOffset);
            }

            rowOffset.MinCellOffset = Math.Min(recordOffset, rowOffset.MinCellOffset);
            rowOffset.MaxCellOffset = Math.Max(recordOffset, rowOffset.MaxCellOffset);
            rowOffset.MaxOverlapRowIndex = Math.Max(maxOverlapRow, rowOffset.MaxOverlapRowIndex);
        }

        internal class XlsRowBlock
        {
            public Dictionary<int, Row> Rows { get; set; }
        }

        internal class XlsRowOffset
        {
            public XlsBiffRow Record { get; set; }

            public int MinCellOffset { get; set; }

            public int MaxCellOffset { get; set; }

            public int MaxOverlapRowIndex { get; set; }
        }
    }
}


namespace ExcelDataReader.Core.CompoundFormat
{
    /// <summary>
    /// Represents single Root Directory record
    /// </summary>
    internal class CompoundDirectoryEntry
    {
        /// <summary>
        /// Gets or sets the name of directory entry
        /// </summary>
        public string EntryName { get; set; }

        /// <summary>
        /// Gets or sets the entry type
        /// </summary>
        public STGTY EntryType { get; set; }

        /// <summary>
        /// Gets or sets the entry "color" in directory tree
        /// </summary>
        public DECOLOR EntryColor { get; set; }

        /// <summary>
        /// Gets or sets the SID of left sibling
        /// </summary>
        /// <remarks>0xFFFFFFFF if there's no one</remarks>
        public uint LeftSiblingSid { get; set; }

        /// <summary>
        /// Gets or sets the SID of right sibling
        /// </summary>
        /// <remarks>0xFFFFFFFF if there's no one</remarks>
        public uint RightSiblingSid { get; set; }

        /// <summary>
        /// Gets or sets the SID of first child (if EntryType is STGTY_STORAGE)
        /// </summary>
        /// <remarks>0xFFFFFFFF if there's no one</remarks>
        public uint ChildSid { get; set; }

        /// <summary>
        /// Gets or sets the CLSID of container (if EntryType is STGTY_STORAGE)
        /// </summary>
        public Guid ClassId { get; set; }

        /// <summary>
        /// Gets or sets the user flags of container (if EntryType is STGTY_STORAGE)
        /// </summary>
        public uint UserFlags { get; set; }

        /// <summary>
        /// Gets or sets the creation time of entry
        /// </summary>
        public DateTime CreationTime { get; set; }

        /// <summary>
        /// Gets or sets the last modification time of entry
        /// </summary>
        public DateTime LastWriteTime { get; set; }

        /// <summary>
        /// Gets or sets the first sector of data stream (if EntryType is STGTY_STREAM)
        /// </summary>
        /// <remarks>if EntryType is STGTY_ROOT, this can be first sector of MiniStream</remarks>
        public uint StreamFirstSector { get; set; }

        /// <summary>
        /// Gets or sets the size of data stream (if EntryType is STGTY_STREAM)
        /// </summary>
        /// <remarks>if EntryType is STGTY_ROOT, this can be size of MiniStream</remarks>
        public uint StreamSize { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether this entry relats to a ministream
        /// </summary>
        public bool IsEntryMiniStream { get; set; }

        /// <summary>
        /// Gets or sets the prop type. Reserved, must be 0.
        /// </summary>
        public uint PropType { get; set; }
    }
}


namespace ExcelDataReader.Core.CompoundFormat
{
    internal class CompoundDocument
    {
        public CompoundDocument(Stream stream)
        {
            var reader = new BinaryReader(stream);

            Header = ReadHeader(reader);

            if (!Header.IsSignatureValid)
                throw new HeaderException(Errors.ErrorHeaderSignature);
            if (Header.ByteOrder != 0xFFFE && Header.ByteOrder != 0xFFFF) // Some broken xls files uses 0xFFFF
                throw new HeaderException(Errors.ErrorHeaderOrder);

            var difSectorChain = ReadDifSectorChain(reader);
            SectorTable = ReadSectorTable(reader, difSectorChain);

            var miniChain = GetSectorChain(Header.MiniFatFirstSector, SectorTable);
            MiniSectorTable = ReadSectorTable(reader, miniChain);

            var directoryChain = GetSectorChain(Header.RootDirectoryEntryStart, SectorTable);
            var bytes = ReadStream(stream, directoryChain, directoryChain.Count * Header.SectorSize);
            ReadDirectoryEntries(bytes);
        }

        internal CompoundHeader Header { get; }

        internal List<uint> SectorTable { get; }

        internal List<uint> MiniSectorTable { get; }

        internal CompoundDirectoryEntry RootEntry { get; set; }

        internal List<CompoundDirectoryEntry> Entries { get; set; }

        // NOTE: DateTime.MaxValue.ToFileTime() fails on Unity in timezones with DST and +~6h offset, like Sidney Australia
        private static long SafeFileTimeMaxDate { get; } = DateTime.MaxValue.ToFileTimeUtc();

        internal static bool IsCompoundDocument(byte[] probe)
        {
            return BitConverter.ToUInt64(probe, 0) == 0xE11AB1A1E011CFD0;
        }

        internal CompoundDirectoryEntry FindEntry(params string[] entryNames)
        {
            foreach (var e in Entries)
            {
                foreach (var entryName in entryNames)
                {
                    if (string.Equals(e.EntryName, entryName, StringComparison.CurrentCultureIgnoreCase))
                        return e;
                }
            }

            return null;
        }

        internal long GetMiniSectorOffset(uint sector)
        {
            return Header.MiniSectorSize * sector;
        }

        internal long GetSectorOffset(uint sector)
        {
            return 512 + Header.SectorSize * sector;
        }

        internal List<uint> GetSectorChain(uint sector, List<uint> sectorTable)
        {
            List<uint> chain = new List<uint>();
            while (sector != (uint)FATMARKERS.FAT_EndOfChain)
            {
                chain.Add(sector);
                sector = GetNextSector(sector, sectorTable);

                if (chain.Contains(sector))
                {
                    throw new CompoundDocumentException(Errors.ErrorCyclicSectorChain);
                }
            }

            TrimSectorChain(chain, FATMARKERS.FAT_FreeSpace);
            return chain;
        }

        /// <summary>
        /// Reads bytes from a regular or mini stream.
        /// </summary>
        internal byte[] ReadStream(Stream stream, uint baseSector, int length, bool isMini)
        {
            using (var cfb = new CompoundStream(this, stream, baseSector, length, isMini, true))
            {
                var bytes = new byte[length];
                cfb.Read(bytes, 0, length);
                return bytes;
            }
        }

        internal byte[] ReadStream(Stream stream, List<uint> sectors, int length)
        {
            using (var cfb = new CompoundStream(this, stream, sectors, length, true))
            {
                var bytes = new byte[length];
                cfb.Read(bytes, 0, length);
                return bytes;
            }
        }

        private void ReadDirectoryEntries(byte[] bytes)
        {
            try
            {
                Entries = new List<CompoundDirectoryEntry>();
                using (var stream = new MemoryStream(bytes))
                {
                    using (var reader = new BinaryReader(stream))
                    {
                        RootEntry = ReadDirectoryEntry(reader);
                        Entries.Add(RootEntry);

                        while (stream.Position < stream.Length)
                        {
                            var entry = ReadDirectoryEntry(reader);
                            Entries.Add(entry);
                        }
                    }
                }
            }
            catch (EndOfStreamException ex)
            {
                throw new CompoundDocumentException(Errors.ErrorEndOfFile, ex);
            }
        }

        private CompoundDirectoryEntry ReadDirectoryEntry(BinaryReader reader)
        {
            var result = new CompoundDirectoryEntry();
            var name = reader.ReadBytes(64);
            var nameLength = reader.ReadUInt16();

            if (nameLength > 0)
            {
                nameLength = Math.Min((ushort)64, nameLength);
                result.EntryName = Encoding.Unicode.GetString(name, 0, nameLength).TrimEnd('\0');
            }

            result.EntryType = (STGTY)reader.ReadByte();
            result.EntryColor = (DECOLOR)reader.ReadByte();
            result.LeftSiblingSid = reader.ReadUInt32();
            result.RightSiblingSid = reader.ReadUInt32();
            result.ChildSid = reader.ReadUInt32();
            result.ClassId = new Guid(reader.ReadBytes(16));
            result.UserFlags = reader.ReadUInt32();
            result.CreationTime = ReadFileTime(reader);
            result.LastWriteTime = ReadFileTime(reader);
            result.StreamFirstSector = reader.ReadUInt32();
            result.StreamSize = reader.ReadUInt32();
            result.PropType = reader.ReadUInt32();
            result.IsEntryMiniStream = result.StreamSize < Header.MiniStreamCutoff;
            return result;
        }

        private DateTime ReadFileTime(BinaryReader reader)
        {
            var d = reader.ReadInt64();
            if (d < 0 || d > SafeFileTimeMaxDate)
            {
                d = 0;
            }

            return DateTime.FromFileTime(d);
        }

        private CompoundHeader ReadHeader(BinaryReader reader)
        {
            var result = new CompoundHeader();
            result.Signature = reader.ReadUInt64();
            result.ClassId = new Guid(reader.ReadBytes(16));
            result.Version = reader.ReadUInt16();
            result.DllVersion = reader.ReadUInt16();
            result.ByteOrder = reader.ReadUInt16();
            result.SectorSizeInPot = reader.ReadUInt16();
            result.MiniSectorSizeInPot = reader.ReadUInt16();
            reader.ReadBytes(6); // skip 6 unused bytes
            result.DirectorySectorCount = reader.ReadInt32();
            result.FatSectorCount = reader.ReadInt32();
            result.RootDirectoryEntryStart = reader.ReadUInt32();
            result.TransactionSignature = reader.ReadUInt32();
            result.MiniStreamCutoff = reader.ReadUInt32();
            result.MiniFatFirstSector = reader.ReadUInt32();
            result.MiniFatSectorCount = reader.ReadInt32();
            result.DifFirstSector = reader.ReadUInt32();
            result.DifSectorCount = reader.ReadInt32();

            var chain = new List<uint>();
            for (int i = 0; i < 109; ++i)
            {
                chain.Add(reader.ReadUInt32());
            }

            result.First109DifSectorChain = chain;

            return result;
        }

        /// <summary>
        /// The header contains the first 109 DIF entries. If there are any more, read from a separate stream.
        /// </summary>
        private List<uint> ReadDifSectorChain(BinaryReader reader)
        {
            // Read the DIF chain sectors directly, can't use ReadStream yet because it depends on the DIF chain
            var difSectorChain = new List<uint>(Header.First109DifSectorChain);
            if (Header.DifFirstSector != (uint)FATMARKERS.FAT_EndOfChain)
            {
                try
                {
                    var difSector = Header.DifFirstSector;
                    for (var i = 0; i < Header.DifSectorCount; ++i)
                    {
                        var difContent = ReadSectorAsUInt32s(reader, difSector);
                        difSectorChain.AddRange(difContent.GetRange(0, difContent.Count - 1));

                        // The DIFAT sectors are linked together by the "Next DIFAT Sector Location" in each DIFAT sector:
                        difSector = difContent[difContent.Count - 1];
                    }
                }
                catch (EndOfStreamException ex)
                {
                    throw new CompoundDocumentException(Errors.ErrorEndOfFile, ex);
                }
            }

            TrimSectorChain(difSectorChain, FATMARKERS.FAT_FreeSpace);

            // A special value of ENDOFCHAIN (0xFFFFFFFE) is stored in the "Next DIFAT Sector Location" field of the
            // last DIFAT sector, or in the header when no DIFAT sectors are needed.
            TrimSectorChain(difSectorChain, FATMARKERS.FAT_EndOfChain);

            return difSectorChain;
        }

        private List<uint> ReadSectorTable(BinaryReader reader, List<uint> chain)
        {
            var sectorTable = new List<uint>(Header.SectorSize / 4 * chain.Count);
            try
            {
                foreach (var sector in chain)
                {
                    var result = ReadSectorAsUInt32s(reader, sector);
                    sectorTable.AddRange(result);
                }
            }
            catch (EndOfStreamException ex)
            {
                throw new CompoundDocumentException(Errors.ErrorEndOfFile, ex);
            }

            TrimSectorChain(sectorTable, FATMARKERS.FAT_FreeSpace);

            return sectorTable;
        }

        private List<uint> ReadSectorAsUInt32s(BinaryReader reader, uint sector)
        {
            var result = new List<uint>(Header.SectorSize / 4);
            reader.BaseStream.Seek(GetSectorOffset(sector), SeekOrigin.Begin);
            for (var i = 0; i < Header.SectorSize / 4; ++i)
            {
                var value = reader.ReadUInt32();
                result.Add(value);
            }

            return result;
        }

        private void TrimSectorChain(List<uint> chain, FATMARKERS marker)
        {
            while (chain.Count > 0 && chain[chain.Count - 1] == (uint)marker)
            {
                chain.RemoveAt(chain.Count - 1);
            }
        }

        private uint GetNextSector(uint sector, List<uint> sectorTable)
        {
            if (sector < sectorTable.Count)
            {
                return sectorTable[(int)sector];
            }

            return (uint)FATMARKERS.FAT_EndOfChain;
        }
    }
}

namespace ExcelDataReader.Core.CompoundFormat
{
    internal enum STGTY : byte
    {
        STGTY_INVALID = 0,
        STGTY_STORAGE = 1,
        STGTY_STREAM = 2,
        STGTY_LOCKBYTES = 3,
        STGTY_PROPERTY = 4,
        STGTY_ROOT = 5
    }

    internal enum DECOLOR : byte
    {
        DE_RED = 0,
        DE_BLACK = 1
    }

    internal enum FATMARKERS : uint
    {
        FAT_EndOfChain = 0xFFFFFFFE,
        FAT_FreeSpace = 0xFFFFFFFF,
        FAT_FatSector = 0xFFFFFFFD,
        FAT_DifSector = 0xFFFFFFFC
    }
}


namespace ExcelDataReader.Core.CompoundFormat
{
    /// <summary>
    /// Represents Excel file header
    /// </summary>
    internal class CompoundHeader
    {
        /// <summary>
        /// Gets or sets the file signature
        /// </summary>
        public ulong Signature { get; set; }

        /// <summary>
        /// Gets a value indicating whether the signature is valid. 
        /// </summary>
        public bool IsSignatureValid => Signature == 0xE11AB1A1E011CFD0;

        /// <summary>
        /// Gets or sets the class id. Typically filled with zeroes
        /// </summary>
        public Guid ClassId { get; set; }

        /// <summary>
        /// Gets or sets the version. Must be 0x003E
        /// </summary>
        public ushort Version { get; set; }

        /// <summary>
        /// Gets or sets the dll version. Must be 0x0003
        /// </summary>
        public ushort DllVersion { get; set; }

        /// <summary>
        /// Gets or sets the byte order. Must be 0xFFFE
        /// </summary>
        public ushort ByteOrder { get; set; }

        /// <summary>
        /// Gets or sets the sector size in Pot
        /// </summary>
        public int SectorSizeInPot { get; set; }

        /// <summary>
        /// Gets the sector size. Typically 512
        /// </summary>
        public int SectorSize => 1 << SectorSizeInPot;

        /// <summary>
        /// Gets or sets the mini sector size in Pot
        /// </summary>
        public int MiniSectorSizeInPot { get; set; }

        /// <summary>
        /// Gets the mini sector size. Typically 64
        /// </summary>
        public int MiniSectorSize => 1 << MiniSectorSizeInPot;

        /// <summary>
        /// Gets or sets the number of directory sectors. If Major Version is 3, the Number of 
        /// Directory Sectors MUST be zero. This field is not supported for version 3 compound files
        /// </summary>
        public int DirectorySectorCount { get; set; }

        /// <summary>
        /// Gets or sets the number of FAT sectors
        /// </summary>
        public int FatSectorCount { get; set; }

        /// <summary>
        /// Gets or sets the number of first Root Directory Entry (Property Set Storage, FAT Directory) sector
        /// </summary>
        public uint RootDirectoryEntryStart { get; set; }

        /// <summary>
        /// Gets or sets the transaction signature, 0 for Excel
        /// </summary>
        public uint TransactionSignature { get; set; }

        /// <summary>
        /// Gets or sets the maximum size for small stream, typically 4096 bytes
        /// </summary>
        public uint MiniStreamCutoff { get; set; }

        /// <summary>
        /// Gets or sets the first sector of Mini FAT, FAT_EndOfChain if there's no one
        /// </summary>
        public uint MiniFatFirstSector { get; set; }

        /// <summary>
        /// Gets or sets the number of sectors in Mini FAT, 0 if there's no one
        /// </summary>
        public int MiniFatSectorCount { get; set; }

        /// <summary>
        /// Gets or sets the first sector of DIF, FAT_EndOfChain if there's no one
        /// </summary>
        public uint DifFirstSector { get; set; }

        /// <summary>
        /// Gets or sets the number of sectors in DIF, 0 if there's no one
        /// </summary>
        public int DifSectorCount { get; set; }

        /// <summary>
        /// Gets or sets the first 109 locations in the DIF sector chain
        /// </summary>
        public List<uint> First109DifSectorChain { get; set; }
    }
}


namespace ExcelDataReader.Core.CompoundFormat
{
    internal class CompoundStream : Stream
    {
        public CompoundStream(CompoundDocument document, Stream baseStream, List<uint> sectorChain, int length, bool leaveOpen)
        {
            Document = document;
            BaseStream = baseStream;
            IsMini = false;
            LeaveOpen = leaveOpen;
            Length = length;
            SectorChain = sectorChain;
            ReadSector();
        }

        public CompoundStream(CompoundDocument document, Stream baseStream, uint baseSector, int length, bool isMini, bool leaveOpen)
        {
            Document = document;
            BaseStream = baseStream;
            IsMini = isMini;
            Length = length;
            LeaveOpen = leaveOpen;

            if (IsMini)
            {
                SectorChain = Document.GetSectorChain(baseSector, Document.MiniSectorTable);
                RootSectorChain = Document.GetSectorChain(Document.RootEntry.StreamFirstSector, Document.SectorTable);
            }
            else
            {
                SectorChain = Document.GetSectorChain(baseSector, Document.SectorTable);
            }

            ReadSector();
        }

        public List<uint> SectorChain { get; }

        public List<uint> RootSectorChain { get; }

        public override bool CanRead => true;

        public override bool CanSeek => false;

        public override bool CanWrite => false;

        public override long Length { get; }

        public override long Position { get => Offset - SectorBytes.Length + SectorOffset; set => Seek(value, SeekOrigin.Begin); }

        private Stream BaseStream { get; set; }

        private CompoundDocument Document { get; }

        private bool IsMini { get; }

        private bool LeaveOpen { get; }

        private int SectorChainOffset { get; set; }

        private int Offset { get; set; }

        private int SectorOffset { get; set; }

        private byte[] SectorBytes { get; set; }

        public override void Flush()
        {
        }

        public override int Read(byte[] buffer, int offset, int count)
        {
            int index = 0;
            while (index < count && Position < Length)
            {
                if (SectorOffset == SectorBytes.Length)
                {
                    ReadSector();
                    SectorOffset = 0;
                }

                var chunkSize = Math.Min(count - index, SectorBytes.Length - SectorOffset);
                Array.Copy(SectorBytes, SectorOffset, buffer, offset + index, chunkSize);
                index += chunkSize;
                SectorOffset += chunkSize;
            }

            return index;
        }

        public override long Seek(long offset, SeekOrigin origin)
        {
            var sectorSize = IsMini ? Document.Header.MiniSectorSize : Document.Header.SectorSize;
            switch (origin)
            {
                case SeekOrigin.Begin:
                    SectorChainOffset = (int)(offset / sectorSize);
                    Offset = SectorChainOffset * sectorSize;
                    SectorOffset = (int)(offset % sectorSize);
                    if (Offset < Length)
                        ReadSector();
                    return Position;
                case SeekOrigin.Current:
                    return Seek(Position + offset, SeekOrigin.Begin);
                case SeekOrigin.End:
                    return Seek(Length + offset, SeekOrigin.Begin);
                default:
                    return Offset;
            }
        }

        public override void SetLength(long value)
        {
            throw new NotImplementedException();
        }

        public override void Write(byte[] buffer, int offset, int count)
        {
            throw new NotImplementedException();
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing && !LeaveOpen)
            {
                BaseStream?.Dispose();
                BaseStream = null;
            }

            base.Dispose(disposing);
        }

        private void ReadSector()
        {
            if (IsMini)
            {
                ReadMiniSector();
            }
            else
            {
                ReadRegularSector();
            }
        }

        private void ReadMiniSector()
        {
            var sector = SectorChain[SectorChainOffset];
            var miniStreamOffset = (int)Document.GetMiniSectorOffset(sector);

            var rootSectorIndex = miniStreamOffset / Document.Header.SectorSize;
            if (rootSectorIndex >= RootSectorChain.Count)
            {
                throw new CompoundDocumentException(Errors.ErrorEndOfFile);
            }

            var rootSector = RootSectorChain[rootSectorIndex];
            var rootOffset = miniStreamOffset % Document.Header.SectorSize;

            BaseStream.Seek(Document.GetSectorOffset(rootSector) + rootOffset, SeekOrigin.Begin);

            var chunkSize = (int)Math.Min(Length - Offset, Document.Header.MiniSectorSize);
            SectorBytes = new byte[chunkSize];
            if (BaseStream.Read(SectorBytes, 0, chunkSize) < chunkSize)
            {
                throw new CompoundDocumentException(Errors.ErrorEndOfFile);
            }

            Offset += chunkSize;
            SectorChainOffset++;
        }

        private void ReadRegularSector()
        {
            var sector = SectorChain[SectorChainOffset];
            BaseStream.Seek(Document.GetSectorOffset(sector), SeekOrigin.Begin);

            var chunkSize = (int)Math.Min(Length - Offset, Document.Header.SectorSize);
            SectorBytes = new byte[chunkSize];
            if (BaseStream.Read(SectorBytes, 0, chunkSize) < chunkSize)
            {
                throw new CompoundDocumentException(Errors.ErrorEndOfFile);
            }

            Offset += chunkSize;
            SectorChainOffset++;
        }
    }
}


namespace ExcelDataReader.Core.CsvFormat
{
    internal static class CsvAnalyzer
    {
        /// <summary>
        /// Reads completely through a CSV stream to determine encoding, separator, field count and row count. 
        /// Uses fallbackEncoding if there is no BOM. Throws DecoderFallbackException if there are invalid characters in the stream.
        /// Returns the separator whose average field count is closest to its max field count.
        /// </summary>
        public static void Analyze(Stream stream, char[] separators, Encoding fallbackEncoding, int analyzeInitialCsvRows, out int fieldCount, out char autodetectSeparator, out Encoding autodetectEncoding, out int bomLength, out int rowCount)
        {
            var bufferSize = 1024;
            var probeSize = 16;
            var buffer = new byte[bufferSize];
            var bytesRead = stream.Read(buffer, 0, probeSize);

            autodetectEncoding = GetEncodingFromBom(buffer, out bomLength);
            if (autodetectEncoding == null)
            {
                autodetectEncoding = fallbackEncoding;
            }

            if (separators == null || separators.Length == 0)
            {
                separators = new char[] { '\0' };
            }

            var separatorInfos = new SeparatorInfo[separators.Length];
            for (var i = 0; i < separators.Length; i++)
            {
                separatorInfos[i] = new SeparatorInfo();
                separatorInfos[i].Buffer = new CsvParser(separators[i], autodetectEncoding);
            }

            AnalyzeCsvRows(stream, buffer, bytesRead, bomLength, analyzeInitialCsvRows, separators, separatorInfos);

            FlushSeparatorsBuffers(separators, separatorInfos);

            SeparatorInfo bestSeparatorInfo = separatorInfos[0];
            char bestSeparator = separators[0];
            double bestDistance = double.MaxValue;

            for (var i = 0; i < separators.Length; i++)
            {
                var separator = separators[i];
                var separatorInfo = separatorInfos[i];

                // Row has one column if there are no separators, there must be at least one separator to count
                if (separatorInfo.RowCount == 0 || separatorInfo.MaxFieldCount <= 1)
                {
                    continue;
                }

                var average = separatorInfo.SumFieldCount / (double)separatorInfo.RowCount;
                var dist = separatorInfo.MaxFieldCount - average;

                if (dist < bestDistance)
                {
                    bestDistance = dist;
                    bestSeparator = separator;
                    bestSeparatorInfo = separatorInfo;
                }
            }

            autodetectSeparator = bestSeparator;
            fieldCount = bestSeparatorInfo.MaxFieldCount;
            rowCount = analyzeInitialCsvRows == 0 ? bestSeparatorInfo.RowCount : -1;
        }

        private static void AnalyzeCsvRows(Stream inputStream, byte[] buffer, int initialBytesRead, int bomLength, int analyzeInitialCsvRows, char[] separators, SeparatorInfo[] separatorInfos)
        {
            ParseSeparatorsBuffer(buffer, bomLength, initialBytesRead - bomLength, separators, separatorInfos);

            if (IsMinNumberOfRowAnalyzed(analyzeInitialCsvRows, separatorInfos))
            {
                return;
            }

            while (inputStream.Position < inputStream.Length)
            {
                var bytesRead = inputStream.Read(buffer, 0, buffer.Length);
                ParseSeparatorsBuffer(buffer, 0, bytesRead, separators, separatorInfos);
                if (IsMinNumberOfRowAnalyzed(analyzeInitialCsvRows, separatorInfos))
                {
                    return;
                }
            }
        }

        private static bool IsMinNumberOfRowAnalyzed(
            int analyzeInitialCsvRows,
            SeparatorInfo[] separatorInfos)
        {
            if (analyzeInitialCsvRows > 0)
            {
                foreach (var separatorInfo in separatorInfos)
                {
                    if (separatorInfo.RowCount >= analyzeInitialCsvRows)
                    {
                        return true;
                    }
                }
            }

            return false;
        }

        private static void ParseSeparatorsBuffer(byte[] bytes, int offset, int count, char[] separators, SeparatorInfo[] separatorInfos)
        {
            for (var i = 0; i < separators.Length; i++)
            {
                var separator = separators[i];
                SeparatorInfo separatorInfo = separatorInfos[i];

                separatorInfo.Buffer.ParseBuffer(bytes, offset, count, out var rows);

                foreach (var row in rows)
                {
                    separatorInfo.MaxFieldCount = Math.Max(separatorInfo.MaxFieldCount, row.Count);
                    separatorInfo.SumFieldCount += row.Count;
                    separatorInfo.RowCount++;
                }
            }
        }

        private static void FlushSeparatorsBuffers(char[] separators, SeparatorInfo[] separatorInfos)
        {
            for (var i = 0; i < separators.Length; i++)
            {
                var separator = separators[i];
                SeparatorInfo separatorInfo = separatorInfos[i];

                separatorInfo.Buffer.Flush(out var rows);

                foreach (var row in rows)
                {
                    separatorInfo.MaxFieldCount = Math.Max(separatorInfo.MaxFieldCount, row.Count);
                    separatorInfo.SumFieldCount += row.Count;
                    separatorInfo.RowCount++;
                }
            }
        }

        private static Encoding GetEncodingFromBom(byte[] bom, out int bomLength)
        {
            var encodings = new Encoding[]
            {
                Encoding.Unicode, Encoding.BigEndianUnicode, Encoding.UTF8
            };

            foreach (var encoding in encodings)
            {
                if (IsEncodingPreamble(bom, encoding, out int length))
                {
                    bomLength = length;
                    return encoding;
                }
            }

            bomLength = 0;
            return null;
        }

        private static bool IsEncodingPreamble(byte[] bom, Encoding encoding, out int bomLength)
        {
            bomLength = 0;
            var preabmle = encoding.GetPreamble();
            if (preabmle.Length > bom.Length)
                return false;
            var i = 0;
            for (; i < preabmle.Length; i++)
            {
                if (preabmle[i] != bom[i])
                    return false;
            }

            bomLength = i;
            return true;
        }

        private class SeparatorInfo
        {
            public int MaxFieldCount { get; set; }

            public int SumFieldCount { get; set; }

            public int RowCount { get; set; }

            public CsvParser Buffer { get; set; }
        }
    }
}


namespace ExcelDataReader.Core.CsvFormat
{
    /// <summary>
    /// Low level, reentrant CSV parser. Call ParseBuffer() in a loop, and finally Flush() to empty the internal buffers.
    /// </summary>
    internal class CsvParser
    {
        public CsvParser(char separator, Encoding encoding)
        {
            Separator = separator;
            QuoteChar = '"';

            Decoder = encoding.GetDecoder();
            Decoder.Fallback = new DecoderExceptionFallback();

            var bufferSize = 1024;
            CharBuffer = new char[bufferSize];

            State = CsvState.PreValue;
        }

        private enum CsvState
        {
            PreValue,
            Value,
            QuotedValue,
            QuotedValueQuote,
            Separator,
            Linebreak,
            EndOfFile,
        }

        private CsvState State { get; set; }

        private char QuoteChar { get; }

        private int TrailingWhitespaceCount { get; set; }

        private Decoder Decoder { get; }

        private bool HasCarriageReturn { get; set; }

        private char Separator { get; }

        private char[] CharBuffer { get; set; }

        private StringBuilder ValueResult { get; set; } = new StringBuilder();

        private List<string> RowResult { get; set; } = new List<string>();

        private List<List<string>> RowsResult { get; set; } = new List<List<string>>();

        public void ParseBuffer(byte[] bytes, int offset, int count, out List<List<string>> rows)
        {
            while (count > 0)
            {
                Decoder.Convert(bytes, offset, count, CharBuffer, 0, CharBuffer.Length, false, out var bytesUsed, out var charsUsed, out var completed);

                offset += bytesUsed;
                count -= bytesUsed;

                for (var i = 0; i < charsUsed; i++)
                {
                    ParseChar(CharBuffer[i], 1);
                }
            }

            rows = RowsResult;
            RowsResult = new List<List<string>>();
        }

        public void Flush(out List<List<string>> rows)
        {
            if (ValueResult.Length > 0 || RowResult.Count > 0)
            {
                AddValueToRow();
                AddRowToResult();
            }

            rows = RowsResult;
            RowsResult = new List<List<string>>();
        }

        private void ParseChar(char c, int bytesUsed)
        {
            var parsed = false;
            while (!parsed)
            {
                switch (State)
                {
                    case CsvState.PreValue:
                        parsed = ReadPreValue(c, bytesUsed);
                        break;
                    case CsvState.Value:
                        parsed = ReadValue(c, bytesUsed);
                        break;
                    case CsvState.QuotedValue:
                        parsed = ReadQuotedValue(c, bytesUsed);
                        break;
                    case CsvState.QuotedValueQuote:
                        parsed = ReadQuotedValueQuote(c, bytesUsed);
                        break;
                    case CsvState.Separator:
                        parsed = ReadSeparator(c, bytesUsed);
                        break;
                    case CsvState.Linebreak:
                        parsed = ReadLinebreak(c, bytesUsed);
                        break;
                    default:
                        throw new InvalidOperationException("Unhandled parser state: " + State);
                }
            }
        }

        private bool ReadPreValue(char c, int bytesUsed)
        {
            if (IsWhitespace(c))
            {
                return true;
            }
            else if (c == QuoteChar)
            {
                State = CsvState.QuotedValue;
                return true;
            }
            else if (c == Separator)
            {
                State = CsvState.Separator;
                return false;
            }
            else if (c == '\r' || c == '\n' || (c == '\0' && bytesUsed == 0))
            {
                State = CsvState.Linebreak;
                return false;
            }
            else if (c == '\0' && bytesUsed == 0)
            {
                State = CsvState.EndOfFile;
                return false;
            }
            else
            {
                State = CsvState.Value;
                return false;
            }
        }

        private bool ReadValue(char c, int bytesUsed)
        {
            if (c == Separator)
            {
                State = CsvState.Separator;
                return false;
            }
            else if (c == '\r' || c == '\n')
            {
                State = CsvState.Linebreak;
                return false;
            }
            else if (c == '\0' && bytesUsed == 0)
            {
                State = CsvState.EndOfFile;
                return false;
            }
            else
            {
                if (IsWhitespace(c))
                {
                    TrailingWhitespaceCount++;
                }
                else
                {
                    TrailingWhitespaceCount = 0;
                }

                ValueResult.Append(c);
                return true;
            }
        }

        private bool ReadQuotedValue(char c, int bytesUsed)
        {
            if (c == QuoteChar)
            {
                State = CsvState.QuotedValueQuote;
                return true;
            }
            else
            {
                ValueResult.Append(c);
                return true;
            }
        }

        private bool ReadQuotedValueQuote(char c, int bytesUsed)
        {
            if (c == QuoteChar)
            {
                // Is escaped quote
                ValueResult.Append(c);
                State = CsvState.QuotedValue;
                return true;
            }
            else
            {
                // End of quote, read remainder of field as a regular value until separator
                State = CsvState.Value;
                return false;
            }
        }

        private bool ReadSeparator(char c, int bytesUsed)
        {
            AddValueToRow();
            State = CsvState.PreValue;
            return true;
        }

        private bool ReadLinebreak(char c, int bytesUsed)
        {
            if (HasCarriageReturn)
            {
                HasCarriageReturn = false;
                AddValueToRow();
                AddRowToResult();
                State = CsvState.PreValue;
                return c == '\n';
            }
            else if (c == '\r')
            {
                HasCarriageReturn = true;
                return true;
            }
            else
            {
                AddValueToRow();
                AddRowToResult();
                State = CsvState.PreValue;
                return true;
            }
        }

        private void AddValueToRow()
        {
            RowResult.Add(ValueResult.ToString(0, ValueResult.Length - TrailingWhitespaceCount)); 
            ValueResult = new StringBuilder();
            TrailingWhitespaceCount = 0;
        }

        private void AddRowToResult()
        {
            RowsResult.Add(RowResult);
            RowResult = new List<string>();
        }

        private bool IsWhitespace(char c)
        {
            if (c == ' ')
            {
                return true;
            }

            if (Separator != '\t' && c == '\t')
            {
                return true;
            }

            return false;
        }
    }
}


namespace ExcelDataReader.Core.CsvFormat
{
    internal class CsvWorkbook : IWorkbook<CsvWorksheet>
    {
        public CsvWorkbook(Stream stream, Encoding encoding, char[] autodetectSeparators, int analyzeInitialCsvRows)
        {
            Stream = stream;
            Encoding = encoding;
            AutodetectSeparators = autodetectSeparators;
            AnalyzeInitialCsvRows = analyzeInitialCsvRows;
        }

        public int ResultsCount => 1;

        public Stream Stream { get; }

        public Encoding Encoding { get; }

        public char[] AutodetectSeparators { get; }

        public int AnalyzeInitialCsvRows { get; }

        public IEnumerable<CsvWorksheet> ReadWorksheets()
        {
            yield return new CsvWorksheet(Stream, Encoding, AutodetectSeparators, AnalyzeInitialCsvRows);
        }

        public NumberFormatString GetNumberFormatString(int index)
        {
            return null;
        }
    }
}


namespace ExcelDataReader.Core.CsvFormat
{
    internal class CsvWorksheet : IWorksheet
    {
        public CsvWorksheet(Stream stream, Encoding fallbackEncoding, char[] autodetectSeparators, int analyzeInitialCsvRows)
        {
            Stream = stream;
            Stream.Seek(0, SeekOrigin.Begin);
            try
            {
                // Try as UTF-8 first, or use BOM if present
                CsvAnalyzer.Analyze(Stream, autodetectSeparators, Encoding.UTF8, analyzeInitialCsvRows, out var fieldCount, out var separator, out var encoding, out var bomLength, out var rowCount);
                FieldCount = fieldCount;
                AnalyzedRowCount = rowCount;
                AnalyzedPartial = analyzeInitialCsvRows > 0;
                Encoding = encoding;
                Separator = separator;
                BomLength = bomLength;
            }
            catch (DecoderFallbackException)
            {
                // If cannot parse as UTF-8, try fallback encoding
                Stream.Seek(0, SeekOrigin.Begin);

                CsvAnalyzer.Analyze(Stream, autodetectSeparators, fallbackEncoding, analyzeInitialCsvRows, out var fieldCount, out var separator, out var encoding, out var bomLength, out var rowCount);
                FieldCount = fieldCount;
                AnalyzedRowCount = rowCount;
                AnalyzedPartial = analyzeInitialCsvRows > 0;
                Encoding = encoding;
                Separator = separator;
                BomLength = bomLength;
            }
        }

        public string Name => string.Empty;

        public string CodeName => null;

        public string VisibleState => null;

        public HeaderFooter HeaderFooter => null;

        public CellRange[] MergeCells => null;

        public int FieldCount { get; }

        public int RowCount
        {
            get
            {
                if (AnalyzedPartial)
                {
                    throw new InvalidOperationException("Cannot use RowCount with AnalyzeInitialCsvRows > 0");
                }

                return AnalyzedRowCount;
            }
        }

        public Stream Stream { get; }

        public Encoding Encoding { get; }

        public char Separator { get; }

        public Column[] ColumnWidths => null;

        private int BomLength { get; set; }

        private bool AnalyzedPartial { get; }

        private int AnalyzedRowCount { get; }

        public IEnumerable<Row> ReadRows()
        {
            var bufferSize = 1024;
            var buffer = new byte[bufferSize];
            var rowIndex = 0;
            var csv = new CsvParser(Separator, Encoding);
            var skipBomBytes = BomLength;

            Stream.Seek(0, SeekOrigin.Begin);
            while (Stream.Position < Stream.Length)
            {
                var bytesRead = Stream.Read(buffer, 0, bufferSize);
                csv.ParseBuffer(buffer, skipBomBytes, bytesRead - skipBomBytes, out var bufferRows);

                skipBomBytes = 0; // Only skip bom on first iteration

                foreach (var row in GetReaderRows(rowIndex, bufferRows))
                {
                    yield return row;
                }

                rowIndex += bufferRows.Count;
            }

            csv.Flush(out var flushRows);
            foreach (var row in GetReaderRows(rowIndex, flushRows))
            {
                yield return row;
            }
        }

        private IEnumerable<Row> GetReaderRows(int rowIndex, List<List<string>> rows)
        {
            foreach (var row in rows)
            {
                var cells = new List<Cell>(row.Count);
                for (var index = 0; index < row.Count; index++)
                {
                    object value = row[index];
                    cells.Add(new Cell(index, value, new ExtendedFormat(), null));
                }

                yield return new Row(rowIndex, 12.75 /* 255 twips */, cells);

                rowIndex++;
            }
        }
    }
}

namespace ExcelDataReader.Core.NumberFormat
{
    internal class Color
    {
        public string Value { get; set; }
    }
}

namespace ExcelDataReader.Core.NumberFormat
{
    internal class Condition
    {
        public string Operator { get; set; }

        public double Value { get; set; }
    }
}


namespace ExcelDataReader.Core.NumberFormat
{
    internal class DecimalSection
    {
        public bool ThousandSeparator { get; set; }

        public double ThousandDivisor { get; set; }

        public double PercentMultiplier { get; set; }

        public List<string> BeforeDecimal { get; set; }

        public bool DecimalSeparator { get; set; }

        public List<string> AfterDecimal { get; set; }

        public static bool TryParse(List<string> tokens, out DecimalSection format)
        {
            if (Parser.ParseNumberTokens(tokens, 0, out var beforeDecimal, out var decimalSeparator, out var afterDecimal) == tokens.Count)
            {
                bool thousandSeparator;
                var divisor = GetTrailingCommasDivisor(tokens, out thousandSeparator);
                var multiplier = GetPercentMultiplier(tokens);

                format = new DecimalSection()
                {
                    BeforeDecimal = beforeDecimal,
                    DecimalSeparator = decimalSeparator,
                    AfterDecimal = afterDecimal,
                    PercentMultiplier = multiplier,
                    ThousandDivisor = divisor,
                    ThousandSeparator = thousandSeparator
                };

                return true;
            }

            format = null;
            return false;
        }

        private static double GetPercentMultiplier(List<string> tokens)
        {
            // If there is a percentage literal in the part list, multiply the result by 100
            foreach (var token in tokens)
            {
                if (token == "%")
                    return 100;
            }

            return 1;
        }

        private static double GetTrailingCommasDivisor(List<string> tokens, out bool thousandSeparator)
        {
            // This parses all comma literals in the part list:
            // Each comma after the last digit placeholder divides the result by 1000.
            // If there are any other commas, display the result with thousand separators.
            bool hasLastPlaceholder = false;
            var divisor = 1.0;

            for (var j = 0; j < tokens.Count; j++)
            {
                var tokenIndex = tokens.Count - 1 - j;
                var token = tokens[tokenIndex];

                if (!hasLastPlaceholder)
                {
                    if (Token.IsPlaceholder(token))
                    {
                        // Each trailing comma multiplies the divisor by 1000
                        for (var k = tokenIndex + 1; k < tokens.Count; k++)
                        {
                            token = tokens[k];
                            if (token == ",")
                                divisor *= 1000.0;
                            else
                                break;
                        }

                        // Continue scanning backwards from the last digit placeholder, 
                        // but now look for a thousand separator comma
                        hasLastPlaceholder = true;
                    }
                }
                else
                {
                    if (token == ",")
                    {
                        thousandSeparator = true;
                        return divisor;
                    }
                }
            }

            thousandSeparator = false;
            return divisor;
        }
    }
}

namespace ExcelDataReader.Core.NumberFormat
{
    internal class ExponentialSection
    {
        public List<string> BeforeDecimal { get; set; }

        public bool DecimalSeparator { get; set; }

        public List<string> AfterDecimal { get; set; }

        public string ExponentialToken { get; set; }

        public List<string> Power { get; set; }

        public static bool TryParse(List<string> tokens, out ExponentialSection format)
        {
            format = null;

            string exponentialToken;

            int partCount = Parser.ParseNumberTokens(tokens, 0, out var beforeDecimal, out var decimalSeparator, out var afterDecimal);

            if (partCount == 0)
                return false;

            int position = partCount;
            if (position < tokens.Count && Token.IsExponent(tokens[position]))
            {
                exponentialToken = tokens[position];
                position++;
            }
            else
            {
                return false;
            }

            format = new ExponentialSection()
            {
                BeforeDecimal = beforeDecimal,
                DecimalSeparator = decimalSeparator,
                AfterDecimal = afterDecimal,
                ExponentialToken = exponentialToken,
                Power = tokens.GetRange(position, tokens.Count - position)
            };

            return true;
        }
    }
}

namespace ExcelDataReader.Core.NumberFormat
{
    internal class FractionSection
    {
        public List<string> IntegerPart { get; set; }

        public List<string> Numerator { get; set; }

        public List<string> DenominatorPrefix { get; set; }

        public List<string> Denominator { get; set; }

        public int DenominatorConstant { get; set; }

        public List<string> DenominatorSuffix { get; set; }

        public List<string> FractionSuffix { get; set; }

        public static bool TryParse(List<string> tokens, out FractionSection format)
        {
            List<string> numeratorParts = null;
            List<string> denominatorParts = null;

            for (var i = 0; i < tokens.Count; i++)
            {
                var part = tokens[i];
                if (part == "/")
                {
                    numeratorParts = tokens.GetRange(0, i);
                    i++;
                    denominatorParts = tokens.GetRange(i, tokens.Count - i);
                    break;
                }
            }

            if (numeratorParts == null)
            {
                format = null;
                return false;
            }

            GetNumerator(numeratorParts, out var integerPart, out var numeratorPart);

            if (!TryGetDenominator(denominatorParts, out var denominatorPrefix, out var denominatorPart, out var denominatorConstant, out var denominatorSuffix, out var fractionSuffix))
            {
                format = null;
                return false;
            }

            format = new FractionSection()
            {
                IntegerPart = integerPart,
                Numerator = numeratorPart,
                DenominatorPrefix = denominatorPrefix,
                Denominator = denominatorPart,
                DenominatorConstant = denominatorConstant,
                DenominatorSuffix = denominatorSuffix,
                FractionSuffix = fractionSuffix
            };

            return true;
        }

        private static void GetNumerator(List<string> tokens, out List<string> integerPart, out List<string> numeratorPart)
        {
            var hasPlaceholder = false;
            var hasSpace = false;
            var hasIntegerPart = false;
            var numeratorIndex = -1;
            var index = tokens.Count - 1;
            while (index >= 0)
            {
                var token = tokens[index];
                if (Token.IsPlaceholder(token))
                {
                    hasPlaceholder = true;

                    if (hasSpace)
                    {
                        hasIntegerPart = true;
                        break;
                    }
                }
                else
                {
                    if (hasPlaceholder && !hasSpace)
                    {
                        // First time we get here marks the end of the integer part
                        hasSpace = true;
                        numeratorIndex = index + 1;
                    }
                }

                index--;
            }

            if (hasIntegerPart)
            {
                integerPart = tokens.GetRange(0, numeratorIndex);
                numeratorPart = tokens.GetRange(numeratorIndex, tokens.Count - numeratorIndex);
            }
            else
            {
                integerPart = null;
                numeratorPart = tokens;
            }
        }

        private static bool TryGetDenominator(List<string> tokens, out List<string> denominatorPrefix, out List<string> denominatorPart, out int denominatorConstant, out List<string> denominatorSuffix, out List<string> fractionSuffix)
        {
            var index = 0;
            var hasPlaceholder = false;
            var hasConstant = false;

            var constant = new StringBuilder();

            // Read literals until the first number placeholder or digit
            while (index < tokens.Count)
            {
                var token = tokens[index];
                if (Token.IsPlaceholder(token))
                {
                    hasPlaceholder = true;
                    break;
                }
                else if (Token.IsDigit19(token))
                {
                    hasConstant = true;
                    break;
                }

                index++;
            }

            if (!hasPlaceholder && !hasConstant)
            {
                denominatorPrefix = null;
                denominatorPart = null;
                denominatorConstant = 0;
                denominatorSuffix = null;
                fractionSuffix = null;
                return false;
            }

            // The denominator starts here, keep the index
            var denominatorIndex = index;

            // Read placeholders or digits in sequence
            while (index < tokens.Count)
            {
                var token = tokens[index];
                if (hasPlaceholder && Token.IsPlaceholder(token))
                {
                    // OK
                }
                else
                if (hasConstant && Token.IsDigit09(token))
                {
                    constant.Append(token);
                }
                else
                {
                    break;
                }

                index++;
            }

            // 'index' is now at the first token after the denominator placeholders.
            // The remaining, if anything, is to be treated in one or two parts:
            // Any ultimately terminating literals are considered the "Fraction suffix".
            // Anything between the denominator and the fraction suffix is the "Denominator suffix".
            // Placeholders in the denominator suffix are treated as insignificant zeros.

            // Scan backwards to determine the fraction suffix
            int fractionSuffixIndex = tokens.Count;
            while (fractionSuffixIndex > index)
            {
                var token = tokens[fractionSuffixIndex - 1];
                if (Token.IsPlaceholder(token))
                {
                    break;
                }

                fractionSuffixIndex--;
            }

            // Finally extract the detected token ranges
            if (denominatorIndex > 0)
                denominatorPrefix = tokens.GetRange(0, denominatorIndex);
            else
                denominatorPrefix = null;

            if (hasConstant)
                denominatorConstant = int.Parse(constant.ToString());
            else
                denominatorConstant = 0;

            denominatorPart = tokens.GetRange(denominatorIndex, index - denominatorIndex);

            if (index < fractionSuffixIndex)
                denominatorSuffix = tokens.GetRange(index, fractionSuffixIndex - index);
            else
                denominatorSuffix = null;

            if (fractionSuffixIndex < tokens.Count)
                fractionSuffix = tokens.GetRange(fractionSuffixIndex, tokens.Count - fractionSuffixIndex);
            else
                fractionSuffix = null;

            return true;
        }
    }
}

namespace ExcelDataReader.Core.NumberFormat
{
    /// <summary>
    /// Parse ECMA-376 number format strings from Excel and other spreadsheet softwares.
    /// </summary>
    public class NumberFormatString
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="NumberFormatString"/> class.
        /// </summary>
        /// <param name="formatString">The number format string.</param>
        public NumberFormatString(string formatString)
        {
            var tokenizer = new Tokenizer(formatString);
            var sections = new List<Section>();
            var isValid = true;
            while (true)
            {
                var section = Parser.ParseSection(tokenizer, out var syntaxError);

                if (syntaxError)
                    isValid = false;

                if (section == null)
                    break;

                sections.Add(section);
            }

            IsValid = isValid;
            FormatString = formatString;

            if (isValid)
            {
                Sections = sections;
                IsDateTimeFormat = GetFirstSection(SectionType.Date) != null;
                IsTimeSpanFormat = GetFirstSection(SectionType.Duration) != null;
            }
            else
            {
                Sections = new List<Section>();
            }
        }

        /// <summary>
        /// Gets a value indicating whether the number format string is valid.
        /// </summary>
        public bool IsValid { get; }

        /// <summary>
        /// Gets the number format string.
        /// </summary>
        public string FormatString { get; }

        /// <summary>
        /// Gets a value indicating whether the format represents a DateTime
        /// </summary>
        public bool IsDateTimeFormat { get; }

        /// <summary>
        /// Gets a value indicating whether the format represents a TimeSpan
        /// </summary>
        public bool IsTimeSpanFormat { get; }

#if NET20
        internal IList<Section> Sections { get; }
#else
        internal IReadOnlyList<Section> Sections { get; }
#endif

        private Section GetFirstSection(SectionType type)
        {
            foreach (var section in Sections)
            {
                if (section.Type == type)
                {
                    return section;
                }
            }

            return null;
        }
    }
}


namespace ExcelDataReader.Core.NumberFormat
{
    internal static class Parser
    {
        public static Section ParseSection(Tokenizer reader, out bool syntaxError)
        {
            bool hasDateParts = false;
            bool hasDurationParts = false;
            bool hasGeneralPart = false;
            bool hasTextPart = false;
            Condition condition = null;
            Color color = null;
            string token;
            List<string> tokens = new List<string>();

            syntaxError = false;
            while ((token = ReadToken(reader, out syntaxError)) != null)
            {
                if (token == ";")
                    break;

                if (Token.IsDatePart(token))
                {
                    hasDateParts |= true;
                    hasDurationParts |= Token.IsDurationPart(token);
                    tokens.Add(token);
                }
                else if (Token.IsGeneral(token))
                {
                    hasGeneralPart |= true;
                    tokens.Add(token);
                }
                else if (token == "@")
                {
                    hasTextPart |= true;
                    tokens.Add(token);
                }
                else if (token.StartsWith("["))
                {
                    // Does not add to tokens. Absolute/elapsed time tokens
                    // also start with '[', but handled as date part above
                    var expression = token.Substring(1, token.Length - 2);
                    if (TryParseCondition(expression, out var parseCondition))
                        condition = parseCondition;
                    else if (TryParseColor(expression, out var parseColor))
                        color = parseColor;
                }
                else
                {
                    tokens.Add(token);
                }
            }

            if (syntaxError || tokens.Count == 0)
            {
                return null;
            }

            if (
                (hasDateParts && (hasGeneralPart || hasTextPart)) ||
                (hasGeneralPart && (hasDateParts || hasTextPart)) ||
                (hasTextPart && (hasGeneralPart || hasDateParts)))
            {
                // Cannot mix date, general and/or text parts
                syntaxError = true;
                return null;
            }

            SectionType type;
            FractionSection fraction = null;
            ExponentialSection exponential = null;
            DecimalSection number = null;
            List<string> generalTextDateDuration = null;

            if (hasDateParts)
            {
                if (hasDurationParts)
                {
                    type = SectionType.Duration;
                    generalTextDateDuration = tokens;
                }
                else
                {
                    type = SectionType.Date;
                    ParseDate(tokens, out generalTextDateDuration);
                }
            }
            else if (hasGeneralPart)
            {
                type = SectionType.General;
                generalTextDateDuration = tokens;
            }
            else if (hasTextPart)
            {
                type = SectionType.Text;
                generalTextDateDuration = tokens;
            }
            else if (FractionSection.TryParse(tokens, out fraction))
            {
                type = SectionType.Fraction;
            }
            else if (ExponentialSection.TryParse(tokens, out exponential))
            {
                type = SectionType.Exponential;
            }
            else if (DecimalSection.TryParse(tokens, out number))
            {
                type = SectionType.Number;
            }
            else
            {
                // Unable to parse format string
                syntaxError = true;
                return null;
            }

            return new Section()
            {
                Type = type,
                Color = color,
                Condition = condition,
                Fraction = fraction,
                Exponential = exponential,
                Number = number,
                GeneralTextDateDurationParts = generalTextDateDuration
            };
        }

        /// <summary>
        /// Parses as many placeholders and literals needed to format a number with optional decimals. 
        /// Returns number of tokens parsed, or 0 if the tokens didn't form a number.
        /// </summary>
        internal static int ParseNumberTokens(List<string> tokens, int startPosition, out List<string> beforeDecimal, out bool decimalSeparator, out List<string> afterDecimal)
        {
            beforeDecimal = null;
            afterDecimal = null;
            decimalSeparator = false;

            List<string> remainder = new List<string>();
            var index = 0;
            for (index = 0; index < tokens.Count; ++index)
            {
                var token = tokens[index];
                if (token == "." && beforeDecimal == null)
                {
                    decimalSeparator = true;
                    beforeDecimal = tokens.GetRange(0, index); // TODO: why not remainder? has only valid tokens...

                    remainder = new List<string>();
                }
                else if (Token.IsNumberLiteral(token))
                {
                    remainder.Add(token);
                }
                else if (token.StartsWith("["))
                {
                    // ignore
                }
                else
                {
                    break;
                }
            }

            if (remainder.Count > 0)
            {
                if (beforeDecimal != null)
                {
                    afterDecimal = remainder;
                }
                else
                {
                    beforeDecimal = remainder;
                }
            }
            
            return index;
        }

        private static void ParseDate(List<string> tokens, out List<string> result)
        {
            // if tokens form .0 through .000.., combine to single subsecond token
            result = new List<string>();
            for (var i = 0; i < tokens.Count; i++)
            {
                var token = tokens[i];
                if (token == ".")
                {
                    var zeros = 0;
                    while (i + 1 < tokens.Count && tokens[i + 1] == "0")
                    {
                        i++;
                        zeros++;
                    }

                    if (zeros > 0)
                        result.Add("." + new string('0', zeros));
                    else
                        result.Add(".");
                }
                else
                {
                    result.Add(token);
                }
            }
        }

        private static string ReadToken(Tokenizer reader, out bool syntaxError)
        {
            var offset = reader.Position;
            if (
                ReadLiteral(reader) ||
                reader.ReadEnclosed('[', ']') ||

                // Symbols
                reader.ReadOneOf("#?,!&%+-$€£0123456789{}():;/.@ ") ||
                reader.ReadString("e+", true) ||
                reader.ReadString("e-", true) ||
                reader.ReadString("General", true) ||

                // Date
                reader.ReadString("am/pm", true) ||
                reader.ReadString("a/p", true) ||
                reader.ReadOneOrMore('y') ||
                reader.ReadOneOrMore('Y') ||
                reader.ReadOneOrMore('m') ||
                reader.ReadOneOrMore('M') ||
                reader.ReadOneOrMore('d') ||
                reader.ReadOneOrMore('D') ||
                reader.ReadOneOrMore('h') ||
                reader.ReadOneOrMore('H') ||
                reader.ReadOneOrMore('s') ||
                reader.ReadOneOrMore('S') ||
                reader.ReadOneOrMore('g') ||
                reader.ReadOneOrMore('G'))
            {
                syntaxError = false;
                var length = reader.Position - offset;
                return reader.Substring(offset, length);
            }

            syntaxError = reader.Position < reader.Length;
            return null;
        }

        private static bool ReadLiteral(Tokenizer reader)
        {
            if (reader.Peek() == '\\' || reader.Peek() == '*' || reader.Peek() == '_')
            {
                reader.Advance(2);
                return true;
            }
            else if (reader.ReadEnclosed('"', '"'))
            {
                return true;
            }

            return false;
        }

        private static bool TryParseCondition(string token, out Condition result)
        {
            var tokenizer = new Tokenizer(token);

            if (tokenizer.ReadString("<=") ||
                tokenizer.ReadString("<>") ||
                tokenizer.ReadString("<") ||
                tokenizer.ReadString(">=") ||
                tokenizer.ReadString(">") ||
                tokenizer.ReadString("="))
            {
                var conditionPosition = tokenizer.Position;
                var op = tokenizer.Substring(0, conditionPosition);

                if (ReadConditionValue(tokenizer))
                {
                    var valueString = tokenizer.Substring(conditionPosition, tokenizer.Position - conditionPosition);

                    result = new Condition()
                    {
                        Operator = op,
                        Value = double.Parse(valueString, CultureInfo.InvariantCulture)
                    };
                    return true;
                }
            }

            result = null;
            return false;
        }

        private static bool ReadConditionValue(Tokenizer tokenizer)
        {
            // NFPartCondNum = [ASCII-HYPHEN-MINUS] NFPartIntNum [INTL-CHAR-DECIMAL-SEP NFPartIntNum] [NFPartExponential NFPartIntNum]
            tokenizer.ReadString("-");
            while (tokenizer.ReadOneOf("0123456789"))
            {
            }

            if (tokenizer.ReadString("."))
            {
                while (tokenizer.ReadOneOf("0123456789"))
                {
                }
            }

            if (tokenizer.ReadString("e+", true) || tokenizer.ReadString("e-", true))
            {
                if (tokenizer.ReadOneOf("0123456789"))
                {
                    while (tokenizer.ReadOneOf("0123456789"))
                    {
                    }
                }
                else
                {
                    return false;
                }
            }

            return true;
        }

        private static bool TryParseColor(string token, out Color color)
        {
            // TODO: Color1..59
            var tokenizer = new Tokenizer(token);
            if (
                tokenizer.ReadString("black", true) ||
                tokenizer.ReadString("blue", true) ||
                tokenizer.ReadString("cyan", true) ||
                tokenizer.ReadString("green", true) ||
                tokenizer.ReadString("magenta", true) ||
                tokenizer.ReadString("red", true) ||
                tokenizer.ReadString("white", true) ||
                tokenizer.ReadString("yellow", true))
            {
                color = new Color()
                {
                    Value = tokenizer.Substring(0, tokenizer.Position)
                };
                return true;
            }

            color = null;
            return false;
        }
    }
}


namespace ExcelDataReader.Core.NumberFormat
{
    internal class Section
    {
        public SectionType Type { get; set; }

        public Color Color { get; set; }

        public Condition Condition { get; set; }

        public ExponentialSection Exponential { get; set; }

        public FractionSection Fraction { get; set; }

        public DecimalSection Number { get; set; }

        public List<string> GeneralTextDateDurationParts { get; set; }
    }
}
namespace ExcelDataReader.Core.NumberFormat
{
    internal enum SectionType
    {
        General,
        Number,
        Fraction,
        Exponential,
        Date,
        Duration,
        Text,
    }
}

namespace ExcelDataReader.Core.NumberFormat
{
    internal static class Token
    {
        public static bool IsExponent(string token)
        {
            return
                (string.Compare(token, "e+", StringComparison.OrdinalIgnoreCase) == 0) ||
                (string.Compare(token, "e-", StringComparison.OrdinalIgnoreCase) == 0);
        }

        public static bool IsLiteral(string token)
        {
            return
                token.StartsWith("_") ||
                token.StartsWith("\\") ||
                token.StartsWith("\"") ||
                token.StartsWith("*") ||
                token == "," ||
                token == "!" ||
                token == "&" ||
                token == "%" ||
                token == "+" ||
                token == "-" ||
                token == "$" ||
                token == "€" ||
                token == "£" ||
                token == "1" ||
                token == "2" ||
                token == "3" ||
                token == "4" ||
                token == "5" ||
                token == "6" ||
                token == "7" ||
                token == "8" ||
                token == "9" ||
                token == "{" ||
                token == "}" ||
                token == "(" ||
                token == ")" ||
                token == " ";
        }

        public static bool IsNumberLiteral(string token)
        {
            return
                IsPlaceholder(token) ||
                IsLiteral(token) ||
                token == ".";
        }

        public static bool IsPlaceholder(string token)
        {
            return token == "0" || token == "#" || token == "?";
        }

        public static bool IsGeneral(string token)
        {
            return string.Compare(token, "general", StringComparison.OrdinalIgnoreCase) == 0;
        }

        public static bool IsDatePart(string token)
        {
            return
                token.StartsWith("y", StringComparison.OrdinalIgnoreCase) ||
                token.StartsWith("m", StringComparison.OrdinalIgnoreCase) ||
                token.StartsWith("d", StringComparison.OrdinalIgnoreCase) ||
                token.StartsWith("s", StringComparison.OrdinalIgnoreCase) ||
                token.StartsWith("h", StringComparison.OrdinalIgnoreCase) ||
                (token.StartsWith("g", StringComparison.OrdinalIgnoreCase) && !IsGeneral(token)) ||
                string.Compare(token, "am/pm", StringComparison.OrdinalIgnoreCase) == 0 ||
                string.Compare(token, "a/p", StringComparison.OrdinalIgnoreCase) == 0 ||
                IsDurationPart(token);
        }

        public static bool IsDurationPart(string token)
        {
            return
                token.StartsWith("[h", StringComparison.OrdinalIgnoreCase) ||
                token.StartsWith("[m", StringComparison.OrdinalIgnoreCase) ||
                token.StartsWith("[s", StringComparison.OrdinalIgnoreCase);
        }

        public static bool IsDigit09(string token)
        {
            return token == "0" || IsDigit19(token);
        }

        public static bool IsDigit19(string token)
        {
            switch (token)
            {
                case "1":
                case "2":
                case "3":
                case "4":
                case "5":
                case "6":
                case "7":
                case "8":
                case "9":
                    return true;
                default:
                    return false;
            }
        }
    }
}


namespace ExcelDataReader.Core.NumberFormat
{
    internal class Tokenizer
    {
        private string formatString;
        private int formatStringPosition = 0;

        public Tokenizer(string fmt)
        {
            formatString = fmt;
        }

        public int Position => formatStringPosition;

        public int Length => formatString.Length;

        public string Substring(int startIndex, int length)
        {
            return formatString.Substring(startIndex, length);
        }

        public int Peek(int offset = 0)
        {
            if (formatStringPosition + offset >= formatString.Length)
                return -1;
            return formatString[formatStringPosition + offset];
        }

        public int PeekUntil(int startOffset, int until)
        {
            int offset = startOffset;
            while (true)
            {
                var c = Peek(offset++);
                if (c == -1)
                    break;
                if (c == until)
                    return offset - startOffset;
            }

            return 0;
        }

        public bool PeekOneOf(int offset, string s)
        {
            foreach (var c in s)
            {
                if (Peek(offset) == c)
                {
                    return true;
                }
            }

            return false;
        }

        public void Advance(int characters = 1)
        {
            formatStringPosition = Math.Min(formatStringPosition + characters, formatString.Length);
        }

        public bool ReadOneOrMore(int c)
        {
            if (Peek() != c)
                return false;

            while (Peek() == c)
                Advance();

            return true;
        }

        public bool ReadOneOf(string s)
        {
            if (PeekOneOf(0, s))
            {
                Advance();
                return true;
            }

            return false;
        }

        public bool ReadString(string s, bool ignoreCase = false)
        {
            if (formatStringPosition + s.Length > formatString.Length)
                return false;

            for (var i = 0; i < s.Length; i++)
            {
                var c1 = s[i];
                var c2 = (char)Peek(i);
                if (ignoreCase)
                {
                    if (char.ToLower(c1) != char.ToLower(c2))
                        return false;
                }
                else
                {
                    if (c1 != c2)
                        return false;
                }
            }

            Advance(s.Length);
            return true;
        }

        public bool ReadEnclosed(char open, char close)
        {
            if (Peek() == open)
            {
                int length = PeekUntil(1, close);
                if (length > 0)
                {
                    Advance(1 + length);
                    return true;
                }
            }

            return false;
        }
    }
}


namespace ExcelDataReader.Core.OfficeCrypto
{
    /// <summary>
    /// A seekable stream for reading an EncryptedPackage blob using OpenXml Agile Encryption. 
    /// </summary>
    internal class AgileEncryptedPackageStream : Stream
    {
        private const int SegmentLength = 4096;

        public AgileEncryptedPackageStream(Stream stream, byte[] key, byte[] iv, EncryptionInfo encryption)
        {
            Stream = stream;
            Key = key;
            IV = iv;
            Encryption = encryption;

            Stream.Read(SegmentBytes, 0, 8);
            DecryptedLength = BitConverter.ToInt32(SegmentBytes, 0);
            ReadSegment();
        }

        public override bool CanRead => true;

        public override bool CanSeek => true;

        public override bool CanWrite => false;

        public override long Length => DecryptedLength;

        public override long Position { get => Offset - SegmentLength + SegmentOffset; set => Seek(value, SeekOrigin.Begin); }

        private Stream Stream { get; set; }

        private byte[] Key { get; }

        private byte[] IV { get; }

        private HashIdentifier HashAlgorithm { get; }

        private EncryptionInfo Encryption { get; }

        private int Offset { get; set; }

        private byte[] SegmentBytes { get; set; } = new byte[SegmentLength];

        private int SegmentOffset { get; set; }

        private int SegmentIndex { get; set; }

        private int DecryptedLength { get; set; }

        public override void Flush()
        {
        }

        public override int Read(byte[] buffer, int offset, int count)
        {
            if (Position >= Length)
            {
                throw new InvalidOperationException("Tried to read past the end of the encrypted stream");
            }

            int index = 0;
            while (index < count)
            {
                if (SegmentOffset == SegmentBytes.Length)
                {
                    ReadSegment();
                    SegmentOffset = 0;
                }

                var chunkSize = Math.Min(count - index, SegmentBytes.Length - SegmentOffset);
                Array.Copy(SegmentBytes, SegmentOffset, buffer, offset + index, chunkSize);
                index += chunkSize;
                SegmentOffset += chunkSize;
            }

            return index;
        }

        public override long Seek(long offset, SeekOrigin origin)
        {
            switch (origin)
            {
                case SeekOrigin.Begin:
                    SegmentIndex = (int)(offset / SegmentLength);
                    Offset = SegmentIndex * SegmentLength;
                    SegmentOffset = (int)(offset % SegmentLength);
                    if (Offset < Length)
                        ReadSegment();
                    return Position;
                case SeekOrigin.Current:
                    return Seek(Position + offset, SeekOrigin.Begin);
                case SeekOrigin.End:
                    return Seek(Length + offset, SeekOrigin.Begin);
                default:
                    return Offset;
            }
        }

        public override void SetLength(long value)
        {
            throw new NotImplementedException();
        }

        public override void Write(byte[] buffer, int offset, int count)
        {
            throw new NotImplementedException();
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                Stream?.Dispose();
                Stream = null;
            }

            base.Dispose(disposing);
        }

        private void ReadSegment()
        {
            var salt = Encryption.GenerateBlockKey(SegmentIndex, IV);
            
            // NOTE: +8 skips EncryptedPackage header
            Stream.Seek(8 + Offset, SeekOrigin.Begin);
            Stream.Read(SegmentBytes, 0, SegmentLength);

            using (var cipher = Encryption.CreateCipher())
            {
                SegmentBytes = CryptoHelpers.DecryptBytes(cipher, SegmentBytes, Key, salt);
            }

            SegmentIndex++;
            Offset += SegmentLength;
        }
    }
}


namespace ExcelDataReader.Core.OfficeCrypto
{
    /// <summary>
    /// Represents "Agile Encryption" used in XLSX (Office 2010 and newer)
    /// </summary>
    internal class AgileEncryption : EncryptionInfo
    {
        private const string NEncryption = "encryption";
        private const string NKeyData = "keyData";
        private const string NKeyEncryptors = "keyEncryptors";
        private const string NKeyEncryptor = "keyEncryptor";
        private const string NEncryptedKey = "encryptedKey";
        private const string NsEncryption = "http://schemas.microsoft.com/office/2006/encryption";
        private const string NsPassword = "http://schemas.microsoft.com/office/2006/keyEncryptor/password";

        public AgileEncryption(byte[] bytes)
        {
            using (var stream = new MemoryStream(bytes, 8, bytes.Length - 8))
            {
                using (var xmlReader = XmlReader.Create(stream))
                {
                    ReadXmlEncryptionInfoStream(xmlReader);
                }
            }
        }

        public CipherIdentifier CipherAlgorithm { get; set; }

        public CipherMode CipherChaining { get; set; }

        public HashIdentifier HashAlgorithm { get; set; }

        public int KeyBits { get; set; }

        public int BlockSize { get; set; }

        public int HashSize { get; set; }

        public byte[] SaltValue { get; set; }

        public byte[] PasswordSaltValue { get; set; }

        public CipherIdentifier PasswordCipherAlgorithm { get; set; }

        public CipherMode PasswordCipherChaining { get; set; }

        public HashIdentifier PasswordHashAlgorithm { get; set; }

        public byte[] PasswordEncryptedKeyValue { get; set; }

        public byte[] PasswordEncryptedVerifierHashInput { get; set; }

        public byte[] PasswordEncryptedVerifierHashValue { get; set; }

        public int PasswordSpinCount { get; set; }

        public int PasswordKeyBits { get; set; }

        public int PasswordBlockSize { get; set; }

        public override bool IsXor => false;

        public override SymmetricAlgorithm CreateCipher()
        {
            return CryptoHelpers.CreateCipher(CipherAlgorithm, KeyBits, BlockSize * 8, CipherChaining);
        }

        public override byte[] GenerateSecretKey(string password)
        {
            using (var cipher = CryptoHelpers.CreateCipher(PasswordCipherAlgorithm, PasswordKeyBits, PasswordBlockSize * 8, PasswordCipherChaining))
            {
                return GenerateSecretKey(password, PasswordSaltValue, PasswordHashAlgorithm, PasswordEncryptedKeyValue, PasswordSpinCount, PasswordKeyBits, cipher);
            }
        }

        public override byte[] GenerateBlockKey(int blockNumber, byte[] secretKey)
        {
            var salt = CryptoHelpers.HashBytes(CryptoHelpers.Combine(secretKey, BitConverter.GetBytes(blockNumber)), HashAlgorithm);
            Array.Resize(ref salt, BlockSize);
            return salt;
        }

        public override Stream CreateEncryptedPackageStream(Stream stream, byte[] secretKey)
        {
            return new AgileEncryptedPackageStream(stream, secretKey, SaltValue, this);
        }

        public override bool VerifyPassword(string password)
        {
            byte[] secretKey;
            using (var hashAlgorithm = CryptoHelpers.Create(PasswordHashAlgorithm))
                secretKey = HashPassword(password, PasswordSaltValue, hashAlgorithm, PasswordSpinCount);

            var inputBlockKey = CryptoHelpers.HashBytes(
                CryptoHelpers.Combine(secretKey, new byte[] { 0xfe, 0xa7, 0xd2, 0x76, 0x3b, 0x4b, 0x9e, 0x79 }),
                PasswordHashAlgorithm);
            Array.Resize(ref inputBlockKey, PasswordKeyBits / 8);

            var valueBlockKey = CryptoHelpers.HashBytes(
                CryptoHelpers.Combine(secretKey, new byte[] { 0xd7, 0xaa, 0x0f, 0x6d, 0x30, 0x61, 0x34, 0x4e }),
                PasswordHashAlgorithm);
            Array.Resize(ref valueBlockKey, PasswordKeyBits / 8);

            using (var cipher = CryptoHelpers.CreateCipher(PasswordCipherAlgorithm, PasswordKeyBits, PasswordBlockSize * 8, PasswordCipherChaining))
            {
                var decryptedVerifier = CryptoHelpers.DecryptBytes(cipher, PasswordEncryptedVerifierHashInput, inputBlockKey, PasswordSaltValue);
                var decryptedVerifierHash = CryptoHelpers.DecryptBytes(cipher, PasswordEncryptedVerifierHashValue, valueBlockKey, PasswordSaltValue);

                var verifierHash = CryptoHelpers.HashBytes(decryptedVerifier, PasswordHashAlgorithm);
                for (var i = 0; i < Math.Min(decryptedVerifierHash.Length, verifierHash.Length); ++i)
                {
                    if (decryptedVerifierHash[i] != verifierHash[i])
                        return false;
                }

                return true;
            }
        }

        private static byte[] GenerateSecretKey(string password, byte[] saltValue, HashIdentifier hashIdentifier, byte[] encryptedKeyValue, int spinCount, int keyBits, SymmetricAlgorithm cipher)
        {
            var block3 = new byte[] { 0x14, 0x6e, 0x0b, 0xe7, 0xab, 0xac, 0xd0, 0xd6 };

            byte[] hash;
            using (var hashAlgorithm = CryptoHelpers.Create(hashIdentifier))
            {
                hash = HashPassword(password, saltValue, hashAlgorithm, spinCount);

                hash = CryptoHelpers.HashBytes(CryptoHelpers.Combine(hash, block3), hashIdentifier);
            }

            // Truncate or pad with 0x36
            var hashSize = hash.Length;
            Array.Resize(ref hash, keyBits / 8);
            for (var i = hashSize; i < keyBits / 8; i++)
            {
                hash[i] = 0x36;
            }

            // NOTE: the stored salt is padded to a multiple of the block size which affects AES-192
            var decryptedKeyValue = CryptoHelpers.DecryptBytes(cipher, encryptedKeyValue, hash, saltValue);
            Array.Resize(ref decryptedKeyValue, keyBits / 8);
            return decryptedKeyValue;
        }

        private static byte[] HashPassword(string password, byte[] saltValue, HashAlgorithm hashAlgorithm, int spinCount)
        {
            var h = hashAlgorithm.ComputeHash(CryptoHelpers.Combine(saltValue, System.Text.Encoding.Unicode.GetBytes(password)));

            for (var i = 0; i < spinCount; i++)
            {
                h = hashAlgorithm.ComputeHash(CryptoHelpers.Combine(BitConverter.GetBytes(i), h));
            }

            return h;
        }

        private HashIdentifier ParseHash(string value)
        {
            return (HashIdentifier)Enum.Parse(typeof(HashIdentifier), value);
        }

        private CipherIdentifier ParseCipher(string value, int blockBits)
        {
            if (value == "AES")
            {
                return CipherIdentifier.AES;
            }
            else if (value == "DES")
            {
                return CipherIdentifier.DES;
            }
            else if (value == "3DES")
            {
                return CipherIdentifier.DES3;
            }
            else if (value == "RC2")
            {
                return CipherIdentifier.RC2;
            }

            throw new ArgumentException(nameof(value), "Unknown encryption: " + value);
        }

        private CipherMode ParseCipherMode(string value)
        {
            if (value == "ChainingModeCBC")
                return CipherMode.CBC;
#if NET20 || NET45 || NETSTANDARD2_0
            else if (value == "ChainingModeCFB")
                return CipherMode.CFB;
#endif
            throw new ArgumentException("Invalid CipherMode " + value);
        }

        private void ReadXmlEncryptionInfoStream(XmlReader xmlReader)
        {
            if (!xmlReader.IsStartElement(NEncryption, NsEncryption))
            {
                return;
            }

            if (!XmlReaderHelper.ReadFirstContent(xmlReader))
            {
                return;
            }

            while (!xmlReader.EOF)
            {
                if (xmlReader.IsStartElement(NKeyData, NsEncryption))
                {
                    // <keyData saltSize="16" blockSize="16" keyBits="256" hashSize="64" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA512" saltValue="zYmgeIEW4PVmYPiNJItVCQ=="/>
                    // <dataIntegrity encryptedHmacKey="v11xCwbBfQ6Wq03h2M6Nh5Z9fwNnFQwEzu8vmBDps55kd+HfLDzrnuKzuQq4tlpxW0nX99VWh+n2X6ukU6v9FQ==" encryptedHmacValue="SvDwFQR4dNsXOzNstFWHqSHpAUWHQvAr63IhxlxhlQEAczDPIwCWD32aIEFipY7NOlW+LvYPaKC8zO1otxit2g=="/>
                    // <keyEncryptors><keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password"><p:encryptedKey spinCount="100000" saltSize="16" blockSize="16" keyBits="256" hashSize="64" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA512" saltValue="n37HW2mNfJuGwVxTeBY1LA==" encryptedVerifierHashInput="2Y2Oo+QDyMdo327gZUcejA==" encryptedVerifierHashValue="PmkCD5y5cHqMQqbgACUgxLRgISYZL6+jj3K0PSrFDWlEG+fjzFevIee1FubgdpY2P22IIM6W7C/bXE0ayAo8yg==" encryptedKeyValue="qzkvVPIBy2Bk/w2/fp+hhpq5sPReA8aUu414/Xh7494="/></keyEncryptor></keyEncryptors>
                    int saltSize, blockSize, keyBits, hashSize;
                    var cipherAlgorithm = xmlReader.GetAttribute("cipherAlgorithm");
                    var cipherChaining = xmlReader.GetAttribute("cipherChaining");
                    var hashAlgorithm = xmlReader.GetAttribute("hashAlgorithm");
                    var saltValue = xmlReader.GetAttribute("saltValue");

                    int.TryParse(xmlReader.GetAttribute("saltSize"), out saltSize);
                    int.TryParse(xmlReader.GetAttribute("blockSize"), out blockSize);
                    int.TryParse(xmlReader.GetAttribute("keyBits"), out keyBits);
                    int.TryParse(xmlReader.GetAttribute("hashSize"), out hashSize);

                    SaltValue = Convert.FromBase64String(saltValue);
                    HashSize = hashSize; // given in bytes, also given implicitly by SHA512
                    KeyBits = keyBits;
                    BlockSize = blockSize;
                    CipherAlgorithm = ParseCipher(cipherAlgorithm, blockSize * 8);
                    CipherChaining = ParseCipherMode(cipherChaining);
                    HashAlgorithm = ParseHash(hashAlgorithm);
                    xmlReader.Skip();
                }
                else if (xmlReader.IsStartElement(NKeyEncryptors, NsEncryption))
                {
                    ReadKeyEncryptors(xmlReader);
                }
                else if (!XmlReaderHelper.SkipContent(xmlReader))
                {
                    break;
                }
            }
        }

        private void ReadKeyEncryptors(XmlReader xmlReader)
        {
            if (!XmlReaderHelper.ReadFirstContent(xmlReader))
            {
                return;
            }

            while (!xmlReader.EOF)
            {
                if (xmlReader.IsStartElement(NKeyEncryptor, NsEncryption))
                {
                    // <keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
                    ReadKeyEncryptor(xmlReader);
                }
                else if (!XmlReaderHelper.SkipContent(xmlReader))
                {
                    break;
                }
            }
        }

        private void ReadKeyEncryptor(XmlReader xmlReader)
        {
            if (!XmlReaderHelper.ReadFirstContent(xmlReader))
            {
                return;
            }

            while (!xmlReader.EOF)
            {
                if (xmlReader.IsStartElement(NEncryptedKey, NsPassword))
                {
                    // <p:encryptedKey spinCount="100000" saltSize="16" blockSize="16" keyBits="256" hashSize="64" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA512" saltValue="n37HW2mNfJuGwVxTeBY1LA==" encryptedVerifierHashInput="2Y2Oo+QDyMdo327gZUcejA==" encryptedVerifierHashValue="PmkCD5y5cHqMQqbgACUgxLRgISYZL6+jj3K0PSrFDWlEG+fjzFevIee1FubgdpY2P22IIM6W7C/bXE0ayAo8yg==" encryptedKeyValue="qzkvVPIBy2Bk/w2/fp+hhpq5sPReA8aUu414/Xh7494="/></keyEncryptor></keyEncryptors>
                    int spinCount, saltSize, blockSize, keyBits, hashSize;
                    var cipherAlgorithm = xmlReader.GetAttribute("cipherAlgorithm");
                    var cipherChaining = xmlReader.GetAttribute("cipherChaining");
                    var hashAlgorithm = xmlReader.GetAttribute("hashAlgorithm");
                    var saltValue = xmlReader.GetAttribute("saltValue");
                    var encryptedVerifierHashInput = xmlReader.GetAttribute("encryptedVerifierHashInput");
                    var encryptedVerifierHashValue = xmlReader.GetAttribute("encryptedVerifierHashValue");
                    var encryptedKeyValue = xmlReader.GetAttribute("encryptedKeyValue");

                    int.TryParse(xmlReader.GetAttribute("spinCount"), out spinCount);
                    int.TryParse(xmlReader.GetAttribute("saltSize"), out saltSize);
                    int.TryParse(xmlReader.GetAttribute("blockSize"), out blockSize);
                    int.TryParse(xmlReader.GetAttribute("keyBits"), out keyBits);
                    int.TryParse(xmlReader.GetAttribute("hashSize"), out hashSize);

                    PasswordSaltValue = Convert.FromBase64String(saltValue);
                    PasswordCipherAlgorithm = ParseCipher(cipherAlgorithm, blockSize * 8);
                    PasswordCipherChaining = ParseCipherMode(cipherChaining);
                    PasswordHashAlgorithm = ParseHash(hashAlgorithm);
                    PasswordEncryptedKeyValue = Convert.FromBase64String(encryptedKeyValue);
                    PasswordEncryptedVerifierHashInput = Convert.FromBase64String(encryptedVerifierHashInput);
                    PasswordEncryptedVerifierHashValue = Convert.FromBase64String(encryptedVerifierHashValue);
                    PasswordSpinCount = spinCount;
                    PasswordKeyBits = keyBits;
                    PasswordBlockSize = blockSize;

                    xmlReader.Skip();
                }
                else if (!XmlReaderHelper.SkipContent(xmlReader))
                {
                    break;
                }
            }
        }
    }
}


namespace ExcelDataReader.Core.OfficeCrypto
{
    internal static class CryptoHelpers
    {
        public static HashAlgorithm Create(HashIdentifier hashAlgorithm) 
        {
            switch (hashAlgorithm)
            {
                case HashIdentifier.SHA512:
                    return SHA512.Create();
                case HashIdentifier.SHA384:
                    return SHA384.Create();
                case HashIdentifier.SHA256:
                    return SHA256.Create();
                case HashIdentifier.SHA1:
                    return SHA1.Create();
                case HashIdentifier.MD5:
                    return MD5.Create();
                default:
                    throw new InvalidOperationException("Unsupported hash algorithm");
            }
        }

        public static byte[] HashBytes(byte[] bytes, HashIdentifier hashAlgorithm)
        {
            using (HashAlgorithm hash = Create(hashAlgorithm))
            {
                return hash.ComputeHash(bytes);
            }
        }

        public static byte[] Combine(params byte[][] arrays)
        {
            var length = 0;
            for (var i = 0; i < arrays.Length; i++)
                length += arrays[i].Length;

            byte[] ret = new byte[length];
            int offset = 0;
            foreach (byte[] data in arrays)
            {
                Buffer.BlockCopy(data, 0, ret, offset, data.Length);
                offset += data.Length;
            }

            return ret;
        }

        public static SymmetricAlgorithm CreateCipher(CipherIdentifier identifier, int keySize, int blockSize, CipherMode mode)
        {
            switch (identifier)
            {
                case CipherIdentifier.RC4:
                    return new RC4Managed();
                case CipherIdentifier.DES3:
                    return InitCipher(TripleDES.Create(), keySize, blockSize, mode);
#if NET20 || NET45 || NETSTANDARD2_0
                case CipherIdentifier.RC2:
                    return InitCipher(RC2.Create(), keySize, blockSize, mode);
                case CipherIdentifier.DES:
                    return InitCipher(DES.Create(), keySize, blockSize, mode);
                case CipherIdentifier.AES:
                    return InitCipher(new RijndaelManaged(), keySize, blockSize, mode);
#else
                case CipherIdentifier.AES:
                    return InitCipher(Aes.Create(), keySize, blockSize, mode);
#endif
            }

            throw new InvalidOperationException("Unsupported encryption method: " + identifier.ToString());
        }

        public static SymmetricAlgorithm InitCipher(SymmetricAlgorithm cipher, int keySize, int blockSize, CipherMode mode)
        {
            cipher.KeySize = keySize;
            cipher.BlockSize = blockSize;
            cipher.Mode = mode;
            cipher.Padding = PaddingMode.Zeros;
            return cipher;
        }

        public static byte[] DecryptBytes(SymmetricAlgorithm algo, byte[] bytes, byte[] key, byte[] iv)
        {
            using (var decryptor = algo.CreateDecryptor(key, iv))
            {
                return DecryptBytes(decryptor, bytes);
            }
        }

        public static byte[] DecryptBytes(ICryptoTransform transform, byte[] bytes)
        {
            var length = bytes.Length;
            using (MemoryStream msDecrypt = new MemoryStream(bytes, 0, length))
            {
                using (CryptoStream csDecrypt = new CryptoStream(msDecrypt, transform, CryptoStreamMode.Read))
                {
                    var result = new byte[length];
                    csDecrypt.Read(result, 0, length);
                    return result;
                }
            }
        }
    }
}


namespace ExcelDataReader.Core.OfficeCrypto
{
    /// <summary>
    /// Base class for the various encryption schemes used by Excel
    /// </summary>
    internal abstract class EncryptionInfo
    {
        /// <summary>
        /// Gets a value indicating whether XOR obfuscation is used.
        /// When true, the ICryptoTransform can be cast to XorTransform and
        /// handle the special case where XorArrayIndex must be manipulated
        /// per record.
        /// </summary>
        public abstract bool IsXor { get; }

        public static EncryptionInfo Create(ushort xorEncryptionKey, ushort xorHashValue)
        {
            return new XorEncryption()
            {
                EncryptionKey = xorEncryptionKey,
                HashValue = xorHashValue
            };
        }

        public static EncryptionInfo Create(byte[] bytes)
        {
            // TODO Does this work on a big endian system?
            var versionMajor = BitConverter.ToUInt16(bytes, 0);
            var versionMinor = BitConverter.ToUInt16(bytes, 2);

            if (versionMajor == 1 && versionMinor == 1)
            {
                return new RC4Encryption(bytes);
            }
            else if ((versionMajor == 2 || versionMajor == 3 || versionMajor == 4) && versionMinor == 2)
            {
                // 2.3.4.5 \EncryptionInfo Stream (Standard Encryption)
                return new StandardEncryption(bytes);
            }
            else if ((versionMajor == 3 || versionMajor == 4) && versionMinor == 3)
            {
                // 2.3.4.6 \EncryptionInfo Stream (Extensible Encryption)
                throw new InvalidOperationException("Extensible Encryption not supported");
            }
            else if (versionMajor == 4 && versionMinor == 4)
            {
                // 2.3.4.10 \EncryptionInfo Stream (Agile Encryption)
                return new AgileEncryption(bytes);
            }
            else
            {
                throw new InvalidOperationException("Unsupported EncryptionInfo version " + versionMajor + "." + versionMinor);
            }
        }

        public abstract byte[] GenerateSecretKey(string password);

        public abstract byte[] GenerateBlockKey(int blockNumber, byte[] secretKey);

        public abstract Stream CreateEncryptedPackageStream(Stream stream, byte[] secretKey);

        public abstract bool VerifyPassword(string password);

        public abstract SymmetricAlgorithm CreateCipher();
    }
}

namespace ExcelDataReader.Core.OfficeCrypto
{
    internal enum CipherIdentifier
    {
        None,
        RC2,
        DES,
        DES3,
        AES,
        RC4
    }

    internal enum HashIdentifier
    {
        None,
        MD5,
        SHA1,
        SHA256,
        SHA384,
        SHA512,
    }
}


namespace ExcelDataReader.Core.OfficeCrypto
{
    /// <summary>
    /// Represents the binary RC4+MD5 encryption header used in XLS.
    /// </summary>
    internal class RC4Encryption : EncryptionInfo
    {
        public RC4Encryption(byte[] bytes)
        {
            Salt = new byte[16];
            EncryptedVerifier = new byte[16];
            EncryptedVerifierHash = new byte[16];
            Array.Copy(bytes, 4, Salt, 0, 16);
            Array.Copy(bytes, 4 + 16, EncryptedVerifier, 0, 16);
            Array.Copy(bytes, 4 + 32, EncryptedVerifierHash, 0, 16);
        }

        public byte[] Salt { get; }

        public byte[] EncryptedVerifier { get; }

        public byte[] EncryptedVerifierHash { get; }

        public override bool IsXor => false;

        public static byte[] GenerateSecretKey(string password, byte[] salt)
        {
            if (password.Length > 16)
                password = password.Substring(0, 16);
            var h = CryptoHelpers.HashBytes(System.Text.Encoding.Unicode.GetBytes(password), HashIdentifier.MD5);
            Array.Resize(ref h, 5);

            // Combine h + salt 16 times:
            h = CryptoHelpers.Combine(h, salt, h, salt, h, salt, h, salt, h, salt, h, salt, h, salt, h, salt, h, salt, h, salt, h, salt, h, salt, h, salt, h, salt, h, salt, h, salt);
            h = CryptoHelpers.HashBytes(h, HashIdentifier.MD5);
            Array.Resize(ref h, 5);
            return h;
        }

        public override SymmetricAlgorithm CreateCipher()
        {
            return CryptoHelpers.CreateCipher(CipherIdentifier.RC4, 0, 0, 0);
        }

        public override Stream CreateEncryptedPackageStream(Stream stream, byte[] secretKey)
        {
            throw new NotImplementedException();
        }

        public override byte[] GenerateBlockKey(int blockNumber, byte[] secretKey)
        {
            var salt = CryptoHelpers.Combine(secretKey, BitConverter.GetBytes(blockNumber));
            return CryptoHelpers.HashBytes(salt, HashIdentifier.MD5);
        }

        public override byte[] GenerateSecretKey(string password)
        {
            return GenerateSecretKey(password, Salt);
        }

        public override bool VerifyPassword(string password)
        {
            // 2.3.6.4 Password Verification
            var secretKey = GenerateSecretKey(password);
            var blockKey = GenerateBlockKey(0, secretKey);

            using (var cipher = CryptoHelpers.CreateCipher(CipherIdentifier.RC4, 0, 0, 0))
            {
                using (var transform = cipher.CreateDecryptor(blockKey, null))
                {
                    var decryptedVerifier = CryptoHelpers.DecryptBytes(transform, EncryptedVerifier);
                    var decryptedVerifierHash = CryptoHelpers.DecryptBytes(transform, EncryptedVerifierHash);

                    var verifierHash = CryptoHelpers.HashBytes(decryptedVerifier, HashIdentifier.MD5);
                    for (var i = 0; i < 16; ++i)
                    {
                        if (decryptedVerifierHash[i] != verifierHash[i])
                            return false;
                    }

                    return true;
                }
            }
        }
    }
}


namespace ExcelDataReader.Core.OfficeCrypto
{
    /// <summary>
    /// Minimal RC4 decryption compatible with System.Security.Cryptography.SymmetricAlgorithm.
    /// </summary>
    internal class RC4Managed : SymmetricAlgorithm
    {
        public override ICryptoTransform CreateDecryptor(byte[] rgbKey, byte[] rgbIV)
        {
            return new RC4Transform(rgbKey);
        }

        public override ICryptoTransform CreateEncryptor(byte[] rgbKey, byte[] rgbIV)
        {
            throw new NotImplementedException();
        }

        public override void GenerateIV()
        {
            throw new NotImplementedException();
        }

        public override void GenerateKey()
        {
            throw new NotImplementedException();
        }

        internal class RC4Transform : ICryptoTransform
        {
            private readonly byte[] _s = new byte[256];

            private int _index1;

            private int _index2;

            public RC4Transform(byte[] key)
            {
                Key = key;
                for (int i = 0; i < _s.Length; i++)
                {
                    _s[i] = (byte)i;
                }

                for (int i = 0, j = 0; i < 256; i++)
                {
                    j = (j + key[i % key.Length] + _s[i]) & 255;

                    Swap(_s, i, j);
                }
            }

            public int InputBlockSize => 1024;

            public int OutputBlockSize => 1024;

            public bool CanTransformMultipleBlocks => false;

            public bool CanReuseTransform => false;

            public byte[] Key { get; }

            public void Dispose()
            {
            }

            public int TransformBlock(byte[] inputBuffer, int inputOffset, int inputCount, byte[] outputBuffer, int outputOffset)
            {
                for (var i = 0; i < inputCount; i++)
                {
                    byte mask = Output();
                    outputBuffer[outputOffset + i] = (byte)(inputBuffer[inputOffset + i] ^ mask);
                }

                return inputCount;
            }

            public byte[] TransformFinalBlock(byte[] inputBuffer, int inputOffset, int inputCount)
            {
                var result = new byte[inputCount];
                TransformBlock(inputBuffer, inputOffset, inputCount, result, 0);
                return result;
            }

            private static void Swap(byte[] s, int i, int j)
            {
                byte c = s[i];

                s[i] = s[j];
                s[j] = c;
            }

            private byte Output()
            {
                _index1 = (_index1 + 1) & 255;
                _index2 = (_index2 + _s[_index1]) & 255;

                Swap(_s, _index1, _index2);

                return _s[(_s[_index1] + _s[_index2]) & 255];
            }
        }
    }
}


namespace ExcelDataReader.Core.OfficeCrypto
{
    internal class StandardEncryptedPackageStream : Stream
    {
        public StandardEncryptedPackageStream(Stream underlyingStream, byte[] secretKey, StandardEncryption encryption)
        {
            Cipher = CryptoHelpers.CreateCipher(encryption.CipherAlgorithm, encryption.KeySize, encryption.BlockSize, CipherMode.ECB);
            Decryptor = Cipher.CreateDecryptor(secretKey, encryption.SaltValue);

            var header = new byte[8];
            underlyingStream.Read(header, 0, 8);
            DecryptedLength = BitConverter.ToInt32(header, 0);

            // Wrap CryptoStream to override the length and dispose the cipher and transform 
            // Zip readers scan backwards from the end for the central zip directory, and could fail if its too far away
            // CryptoStream is forward-only, so assume the zip readers read everything to memory
            BaseStream = new CryptoStream(underlyingStream, Decryptor, CryptoStreamMode.Read);
        }

        public override bool CanRead => BaseStream.CanRead;

        public override bool CanSeek => BaseStream.CanSeek;

        public override bool CanWrite => BaseStream.CanWrite;

        public override long Length => DecryptedLength;

        public override long Position
        {
            get => BaseStream.Position;
            set => BaseStream.Position = value;
        }

        private Stream BaseStream { get; set; }

        private SymmetricAlgorithm Cipher { get; set; }

        private ICryptoTransform Decryptor { get; set; }

        private long DecryptedLength { get; }

        public override void Flush()
        {
            BaseStream.Flush();
        }

        public override int Read(byte[] buffer, int offset, int count)
        {
            return BaseStream.Read(buffer, offset, count);
        }

        public override long Seek(long offset, SeekOrigin origin)
        {
            return BaseStream.Seek(offset, origin);
        }

        public override void SetLength(long value)
        {
            BaseStream.SetLength(value);
        }

        public override void Write(byte[] buffer, int offset, int count)
        {
            BaseStream.Write(buffer, offset, count);
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                Decryptor?.Dispose();
                Decryptor = null;

                ((IDisposable)Cipher)?.Dispose();
                Cipher = null;

                BaseStream?.Dispose();
                BaseStream = null;
            }

            base.Dispose(disposing);
        }
    }
}


namespace ExcelDataReader.Core.OfficeCrypto
{
    /// <summary>
    /// Represents the binary "Standard Encryption" header used in XLS and XLSX.
    /// XLS uses RC4+SHA1. XLSX uses AES+SHA1.
    /// </summary>
    internal class StandardEncryption : EncryptionInfo
    {
        private const int AesBlockSize = 128;
        private const int RC4BlockSize = 8;

        public StandardEncryption(byte[] bytes)
        {
            Flags = (EncryptionHeaderFlags)BitConverter.ToUInt32(bytes, 4);

            var headerSize = BitConverter.ToInt32(bytes, 8);

            // Using ProviderType and KeySize instead
            var cipher = (StandardCipher)BitConverter.ToUInt32(bytes, 20);

            var hashAlgorithm = (StandardHash)BitConverter.ToUInt32(bytes, 24);

            if ((Flags & EncryptionHeaderFlags.External) == 0)
            {
                switch (hashAlgorithm)
                {
                    case StandardHash.Default:
                    case StandardHash.SHA1:
                        HashAlgorithm = HashIdentifier.SHA1;
                        break;
                }
            }

            // ECMA-376: 0x00000080 (AES-128), 0x000000C0 (AES-192), or 0x00000100 (AES-256).
            // RC4: 0x00000028 – 0x00000080 (inclusive), 8-bits increments
            KeySize = BitConverter.ToInt32(bytes, 28);

            // Don't use this; is implementation-specific
            var providerType = (StandardProvider)BitConverter.ToUInt32(bytes, 32);

            // skip two reserved dwords
            CSPName = System.Text.Encoding.Unicode.GetString(bytes, 44, headerSize - 44 + 12); // +12 because we start counting from the offset after HeaderSize

            var saltSize = BitConverter.ToInt32(bytes, 12 + headerSize);

            SaltValue = new byte[saltSize];
            Array.Copy(bytes, 12 + headerSize + 4, SaltValue, 0, saltSize);

            Verifier = new byte[16];
            Array.Copy(bytes, 12 + headerSize + 4 + saltSize, Verifier, 0, 16);

            // An unsigned integer that specifies the number of bytes needed to
            // contain the hash of the data used to generate the EncryptedVerifier field.
            VerifierHashBytesNeeded = BitConverter.ToInt32(bytes, 12 + headerSize + 4 + saltSize + 16);

            // If the encryption algorithm is RC4, the length MUST be 20 bytes. If the encryption algorithm is AES, the length MUST be 32 bytes
            var verifierHashSize = ((Flags & EncryptionHeaderFlags.AES) != 0) ? 32 : 20;

            if (cipher == StandardCipher.RC4)
            {
                BlockSize = RC4BlockSize;
                verifierHashSize = 20;
            }
            else if (cipher == StandardCipher.AES128 || cipher == StandardCipher.AES192 || cipher == StandardCipher.AES256)
            {
                BlockSize = AesBlockSize;
                verifierHashSize = 32;
            }

            VerifierHash = new byte[verifierHashSize];
            Array.Copy(bytes, 12 + headerSize + 4 + saltSize + 16 + 4, VerifierHash, 0, verifierHashSize);

            if ((Flags & EncryptionHeaderFlags.External) == 0)
            {
                switch (cipher)
                {
                    case StandardCipher.Default:
                        if ((Flags & EncryptionHeaderFlags.AES) != 0)
                        {
                            CipherAlgorithm = CipherIdentifier.AES;
                        }
                        else if ((Flags & EncryptionHeaderFlags.CryptoAPI) != 0)
                        {
                            CipherAlgorithm = CipherIdentifier.RC4;
                        }

                        break;
                    case StandardCipher.AES128:
                    case StandardCipher.AES192:
                    case StandardCipher.AES256:
                        CipherAlgorithm = CipherIdentifier.AES;
                        break;

                    case StandardCipher.RC4:
                        CipherAlgorithm = CipherIdentifier.RC4;
                        break;
                }
            }
        }

        private enum StandardProvider
        {
            Default = 0x00000000,
            RC4 = 0x00000001,
            AES = 0x00000018,
        }

        private enum StandardCipher
        {
            Default = 0x00000000,
            AES128 = 0x0000660E,
            AES192 = 0x0000660F,
            AES256 = 0x00006610,
            RC4 = 0x00006801
        }

        private enum StandardHash
        {
            Default = 0x00000000,
            SHA1 = 0x00008004,
        }

        private enum EncryptionHeaderFlags : uint
        {
            CryptoAPI = 0x00000004,
            DocProps = 0x00000008,
            External = 0x00000010,
            AES = 0x00000020,
        }

        public CipherIdentifier CipherAlgorithm { get; set; }

        public HashIdentifier HashAlgorithm { get; set; }

        public int BlockSize { get; set; }

        public int KeySize { get; set; }

        public string CSPName { get; set; }

        public byte[] SaltValue { get; set; }

        public byte[] Verifier { get; set; }

        public byte[] VerifierHash { get; set; }

        public int VerifierHashBytesNeeded { get; set; }

        public override bool IsXor => false;

        private EncryptionHeaderFlags Flags { get; set; }

        public override SymmetricAlgorithm CreateCipher()
        {
            return CryptoHelpers.CreateCipher(CipherAlgorithm, KeySize, BlockSize, CipherMode.ECB);
        }

        public override Stream CreateEncryptedPackageStream(Stream stream, byte[] secretKey)
        {
            return new StandardEncryptedPackageStream(stream, secretKey, this);
        }

        public override byte[] GenerateBlockKey(int blockNumber, byte[] secretKey)
        {
            if ((Flags & EncryptionHeaderFlags.AES) != 0)
            {
                /*var salt = CryptoHelpers.Combine(secretKey, BitConverter.GetBytes(blockNumber));
                salt = CryptoHelpers.HashBytes(salt, HashAlgorithm);
                Array.Resize(ref salt, (int)KeySize / 8);
                return salt;*/
                throw new Exception("Block key for ECMA-376 Standard Encryption not implemented");
            }
            else if ((Flags & EncryptionHeaderFlags.CryptoAPI) != 0)
            {
                var salt = CryptoHelpers.Combine(secretKey, BitConverter.GetBytes(blockNumber));
                salt = CryptoHelpers.HashBytes(salt, HashAlgorithm);
                Array.Resize(ref salt, (int)KeySize / 8);
                if (KeySize == 40)
                {
                    // 2.3.5.2: If keyLength is exactly 40 bits, the encryption key MUST be composed of the first 40 bits of Hfinal and 88 bits set to zero, creating a 128-bit key.
                    salt = CryptoHelpers.Combine(salt, new byte[] { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 });
                }

                return salt;
            }
            else
            {
                throw new InvalidOperationException("Unknown encryption type");
            }
        }

        public override byte[] GenerateSecretKey(string password)
        {
            if ((Flags & EncryptionHeaderFlags.AES) != 0)
            {
                // 2.3.4.7 ECMA-376 Document Encryption Key Generation (Standard Encryption)
                return GenerateEcma376SecretKey(password, SaltValue, HashAlgorithm, (int)KeySize, VerifierHashBytesNeeded);
            }
            else if ((Flags & EncryptionHeaderFlags.CryptoAPI) != 0)
            {
                // 2.3.5.2 RC4 CryptoAPI Encryption Key Generation
                return GenerateCryptoApiSecretKey(password, SaltValue, HashAlgorithm, (int)KeySize);
            }
            else
            {
                throw new InvalidOperationException("Unknown encryption type");
            }
        }

        public override bool VerifyPassword(string password)
        {
            // 2.3.4.9 Password Verification (Standard Encryption)
            // 2.3.5.6 Password Verification
            var secretKey = GenerateSecretKey(password);

            var blockKey = ((Flags & EncryptionHeaderFlags.AES) != 0) ? secretKey : GenerateBlockKey(0, secretKey);

            using (var cipher = CryptoHelpers.CreateCipher(CipherAlgorithm, KeySize, BlockSize, CipherMode.ECB))
            {
                using (var transform = cipher.CreateDecryptor(blockKey, SaltValue))
                {
                    var decryptedVerifier = CryptoHelpers.DecryptBytes(transform, Verifier);
                    var decryptedVerifierHash = CryptoHelpers.DecryptBytes(transform, VerifierHash);

                    var verifierHash = CryptoHelpers.HashBytes(decryptedVerifier, HashAlgorithm);
                    for (var i = 0; i < 16; ++i)
                    {
                        if (decryptedVerifierHash[i] != verifierHash[i])
                            return false;
                    }

                    return true;
                }
            }
        }

        /// <summary>
        /// 2.3.5.2 RC4 CryptoAPI Encryption Key Generation
        /// </summary>
        private static byte[] GenerateCryptoApiSecretKey(string password, byte[] saltValue, HashIdentifier hashAlgorithm, int keySize)
        {
            return CryptoHelpers.HashBytes(CryptoHelpers.Combine(saltValue, System.Text.Encoding.Unicode.GetBytes(password)), hashAlgorithm);
        }

        /// <summary>
        /// 2.3.4.7 ECMA-376 Document Encryption Key Generation (Standard Encryption)
        /// </summary>
        private static byte[] GenerateEcma376SecretKey(string password, byte[] saltValue, HashIdentifier hashIdentifier, int keySize, int verifierHashSize)
        {
            byte[] hash;
            using (var hashAlgorithm = CryptoHelpers.Create(hashIdentifier))
            {
                hash = hashAlgorithm.ComputeHash(CryptoHelpers.Combine(saltValue, System.Text.Encoding.Unicode.GetBytes(password)));
                for (int i = 0; i < 50000; i++)
                {
                    hash = hashAlgorithm.ComputeHash(CryptoHelpers.Combine(BitConverter.GetBytes(i), hash));
                }

                hash = hashAlgorithm.ComputeHash(CryptoHelpers.Combine(hash, BitConverter.GetBytes(0)));

                // The algorithm in this 'DeriveKey' function is the bit that's not clear from the documentation
                hash = DeriveKey(hash, hashAlgorithm, keySize, verifierHashSize);
            }

            Array.Resize(ref hash, keySize / 8);

            return hash;
        }

        private static byte[] DeriveKey(byte[] hashValue, HashAlgorithm hashAlgorithm, int keySize, int verifierHashSize)
        {
            // And one more hash to derive the key
            byte[] derivedKey = new byte[64];

            // This is step 4a in 2.3.4.7 of MS_OFFCRYPT version 1.0
            // and is required even though the notes say it should be 
            // used only when the encryption algorithm key > hash length.
            for (int i = 0; i < derivedKey.Length; i++)
                derivedKey[i] = (byte)(i < hashValue.Length ? 0x36 ^ hashValue[i] : 0x36);

            byte[] x1 = hashAlgorithm.ComputeHash(derivedKey);

            if (verifierHashSize > keySize / 8)
                return x1;

            for (int i = 0; i < derivedKey.Length; i++)
                derivedKey[i] = (byte)(i < hashValue.Length ? 0x5C ^ hashValue[i] : 0x5C);

            byte[] x2 = hashAlgorithm.ComputeHash(derivedKey);
            return CryptoHelpers.Combine(x1, x2);
        }
    }
}

namespace ExcelDataReader.Core.OfficeCrypto
{
    /// <summary>
    /// Represents "XOR Deobfucation Method 1" used in XLS.
    /// </summary>
    internal class XorEncryption : EncryptionInfo
    {
        public ushort EncryptionKey { get; set; }

        public ushort HashValue { get; set; }

        public override bool IsXor => true;

        public override SymmetricAlgorithm CreateCipher()
        {
            return new XorManaged();
        }

        public override Stream CreateEncryptedPackageStream(Stream stream, byte[] secretKey)
        {
            throw new NotImplementedException();
        }

        public override byte[] GenerateBlockKey(int blockNumber, byte[] secretKey)
        {
            return secretKey;
        }

        public override byte[] GenerateSecretKey(string password)
        {
            var passwordBytes = System.Text.Encoding.ASCII.GetBytes(password.Substring(0, Math.Min(password.Length, 15)));
            return XorManaged.CreateXorArray_Method1(passwordBytes);
        }

        public override bool VerifyPassword(string password)
        {
            var passwordBytes = System.Text.Encoding.ASCII.GetBytes(password.Substring(0, Math.Min(password.Length, 15)));
            var verifier = XorManaged.CreatePasswordVerifier_Method1(passwordBytes);
            return verifier == HashValue;
        }
    }
}


namespace ExcelDataReader.Core.OfficeCrypto
{
    /// <summary>
    /// Minimal Office "XOR Deobfuscation Method 1" implementation compatible
    /// with System.Security.Cryptography.SymmetricAlgorithm.
    /// </summary>
    internal class XorManaged : SymmetricAlgorithm
    {
        private static byte[] padArray = new byte[]
        {
            0xBB, 0xFF, 0xFF, 0xBA, 0xFF, 0xFF, 0xB9, 0x80,
            0x00, 0xBE, 0x0F, 0x00, 0xBF, 0x0F, 0x00
        };

        private static ushort[] initialCode = new ushort[]
        {
            0xE1F0, 0x1D0F, 0xCC9C, 0x84C0, 0x110C,
            0x0E10, 0xF1CE, 0x313E, 0x1872, 0xE139,
            0xD40F, 0x84F9, 0x280C, 0xA96A, 0x4EC3
        };

        private static ushort[] xorMatrix = new ushort[]
        {
            0xAEFC, 0x4DD9, 0x9BB2, 0x2745, 0x4E8A, 0x9D14, 0x2A09,
            0x7B61, 0xF6C2, 0xFDA5, 0xEB6B, 0xC6F7, 0x9DCF, 0x2BBF,
            0x4563, 0x8AC6, 0x05AD, 0x0B5A, 0x16B4, 0x2D68, 0x5AD0,
            0x0375, 0x06EA, 0x0DD4, 0x1BA8, 0x3750, 0x6EA0, 0xDD40,
            0xD849, 0xA0B3, 0x5147, 0xA28E, 0x553D, 0xAA7A, 0x44D5,
            0x6F45, 0xDE8A, 0xAD35, 0x4A4B, 0x9496, 0x390D, 0x721A,
            0xEB23, 0xC667, 0x9CEF, 0x29FF, 0x53FE, 0xA7FC, 0x5FD9,
            0x47D3, 0x8FA6, 0x0F6D, 0x1EDA, 0x3DB4, 0x7B68, 0xF6D0,
            0xB861, 0x60E3, 0xC1C6, 0x93AD, 0x377B, 0x6EF6, 0xDDEC,
            0x45A0, 0x8B40, 0x06A1, 0x0D42, 0x1A84, 0x3508, 0x6A10,
            0xAA51, 0x4483, 0x8906, 0x022D, 0x045A, 0x08B4, 0x1168,
            0x76B4, 0xED68, 0xCAF1, 0x85C3, 0x1BA7, 0x374E, 0x6E9C,
            0x3730, 0x6E60, 0xDCC0, 0xA9A1, 0x4363, 0x86C6, 0x1DAD,
            0x3331, 0x6662, 0xCCC4, 0x89A9, 0x0373, 0x06E6, 0x0DCC,
            0x1021, 0x2042, 0x4084, 0x8108, 0x1231, 0x2462, 0x48C4
        };

        public XorManaged()
        {
        }

        public override ICryptoTransform CreateDecryptor(byte[] rgbKey, byte[] rgbIV)
        {
            return new XorTransform(rgbKey, 0);
        }

        public override ICryptoTransform CreateEncryptor(byte[] rgbKey, byte[] rgbIV)
        {
            throw new NotImplementedException();
        }

        public override void GenerateIV()
        {
            throw new NotImplementedException();
        }

        public override void GenerateKey()
        {
            throw new NotImplementedException();
        }

        internal static ushort CreatePasswordVerifier_Method1(byte[] passwordBytes)
        {
            var passwordArray = CryptoHelpers.Combine(new byte[] { (byte)passwordBytes.Length }, passwordBytes);
            ushort verifier = 0x0000;
            for (var i = 0; i < passwordArray.Length; ++i)
            {
                var passwordByte = passwordArray[passwordArray.Length - 1 - i];
                ushort intermediate1 = (ushort)(((verifier & 0x4000) == 0) ? 0 : 1);
                ushort intermediate2 = (ushort)(verifier * 2);
                intermediate2 &= 0x7FFF;
                ushort intermediate3 = (ushort)(intermediate1 | intermediate2);
                verifier = (ushort)(intermediate3 ^ passwordByte);
            }

            return (ushort)(verifier ^ 0xCE4B);
        }

        internal static ushort CreateXorKey_Method1(byte[] passwordBytes)
        {
            ushort xorKey = initialCode[passwordBytes.Length - 1];
            var currentElement = 0x68;

            for (var i = 0; i < passwordBytes.Length; ++i)
            {
                var c = passwordBytes[passwordBytes.Length - 1 - i];
                for (var j = 0; j < 7; ++j)
                {
                    if ((c & 0x40) != 0)
                    {
                        xorKey ^= xorMatrix[currentElement];
                    }

                    c *= 2;
                    currentElement--;
                }
            }

            return xorKey;
        }

        /// <summary>
        /// Generates a 16 byte obfuscation array based on the POI/LibreOffice implementations
        /// </summary>
        internal static byte[] CreateXorArray_Method1(byte[] passwordBytes)
        {
            var index = passwordBytes.Length;
            var obfuscationArray = new byte[16];
            Array.Copy(passwordBytes, 0, obfuscationArray, 0, passwordBytes.Length);
            Array.Copy(padArray, 0, obfuscationArray, passwordBytes.Length, padArray.Length - passwordBytes.Length + 1);

            var xorKey = CreateXorKey_Method1(passwordBytes);
            byte[] baseKeyLE = new byte[] { (byte)(xorKey & 0xFF), (byte)((xorKey >> 8) & 0xFF) };
            int nRotateSize = 2;
            for (int i = 0; i < obfuscationArray.Length; i++)
            {
                obfuscationArray[i] ^= baseKeyLE[i & 1];
                obfuscationArray[i] = RotateLeft(obfuscationArray[i], nRotateSize);
            }

            return obfuscationArray;
        }

        private static byte RotateLeft(byte b, int shift)
        {
            return (byte)(((b << shift) | (b >> (8 - shift))) & 0xFF);
        }

        internal class XorTransform : ICryptoTransform
        {
            public XorTransform(byte[] key, int xorArrayIndex)
            {
                XorArray = key;
                XorArrayIndex = xorArrayIndex;
            }

            public int InputBlockSize => 1024;

            public int OutputBlockSize => 1024;

            public bool CanTransformMultipleBlocks => false;

            public bool CanReuseTransform => false;

            /// <summary>
            /// Gets or sets the obfuscation array index. BIFF obfuscation uses a different XorArrayIndex per record.
            /// </summary>
            public int XorArrayIndex { get; set; }

            private byte[] XorArray { get; }

            public void Dispose()
            {
            }

            public int TransformBlock(byte[] inputBuffer, int inputOffset, int inputCount, byte[] outputBuffer, int outputOffset)
            {
                for (var i = 0; i < inputCount; ++i)
                {
                    var value = inputBuffer[inputOffset + i];
                    value = RotateLeft(value, 3);
                    value ^= XorArray[XorArrayIndex % 16];
                    outputBuffer[outputOffset + i] = value;
                    XorArrayIndex++;
                }

                return inputCount;
            }

            public byte[] TransformFinalBlock(byte[] inputBuffer, int inputOffset, int inputCount)
            {
                var result = new byte[inputCount];
                TransformBlock(inputBuffer, inputOffset, inputCount, result, 0);
                return result;
            }
        }
    }
}



#nullable enable

namespace ExcelDataReader.Core.OpenXmlFormat
{
    internal abstract class RecordReader : IDisposable
    {
        ~RecordReader()
        {
            Dispose(false);
        }

        /// <inheritdoc />
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        public abstract Record? Read();

        protected virtual void Dispose(bool disposing) 
        {
        }
    }
}


namespace ExcelDataReader.Core.OpenXmlFormat
{
    /// <summary>
    /// Shared string table
    /// </summary>
    internal class XlsxSST : List<string>
    {
    }
}



namespace ExcelDataReader.Core.OpenXmlFormat
{
    internal class XlsxWorkbook : CommonWorkbook, IWorkbook<XlsxWorksheet>
    {
        private const string NsRelationship = "http://schemas.openxmlformats.org/package/2006/relationships";

        private const string ElementRelationship = "Relationship";
        private const string ElementRelationships = "Relationships";
        private const string AttributeId = "Id";
        private const string AttributeTarget = "Target";

        private readonly ZipWorker _zipWorker;

        public XlsxWorkbook(ZipWorker zipWorker)
        {
            _zipWorker = zipWorker;

            ReadWorkbook();
            ReadWorkbookRels();
            ReadSharedStrings();
            ReadStyles();
        }

        public List<SheetRecord> Sheets { get; } = new List<SheetRecord>();

        public XlsxSST SST { get; } = new XlsxSST();

        public bool IsDate1904 { get; private set; }

        public int ResultsCount => Sheets?.Count ?? -1;

        public IEnumerable<XlsxWorksheet> ReadWorksheets()
        {
            foreach (var sheet in Sheets)
            {
                yield return new XlsxWorksheet(_zipWorker, this, sheet);
            }
        }

        private void ReadWorkbook()
        {
            using var reader = _zipWorker.GetWorkbookReader();

            Record record;
            while ((record = reader.Read()) != null)
            {
                switch (record)
                {
                    case WorkbookPrRecord pr:
                        IsDate1904 = pr.Date1904;
                        break;
                    case SheetRecord sheet:
                        Sheets.Add(sheet);
                        break;
                }
            }
        }

        private void ReadWorkbookRels()
        {
            using var stream = _zipWorker.GetWorkbookRelsStream();
            if (stream == null)
            {
                return;
            }

            using XmlReader reader = XmlReader.Create(stream);
            ReadWorkbookRels(reader);
        }

        private void ReadWorkbookRels(XmlReader reader)
        {
            if (!reader.IsStartElement(ElementRelationships, NsRelationship))
            {
                return;
            }

            if (!XmlReaderHelper.ReadFirstContent(reader))
            {
                return;
            }

            while (!reader.EOF)
            {
                if (reader.IsStartElement(ElementRelationship, NsRelationship))
                {
                    string rid = reader.GetAttribute(AttributeId);
                    foreach (var sheet in Sheets)
                    {
                        if (sheet.Rid == rid)
                        {
                            sheet.Path = reader.GetAttribute(AttributeTarget);
                            break;
                        }
                    }

                    reader.Skip();
                }
                else if (!XmlReaderHelper.SkipContent(reader))
                {
                    break;
                }
            }
        }

        private void ReadSharedStrings()
        {
            using var reader = _zipWorker.GetSharedStringsReader();
            if (reader == null)
                return;

            Record record;
            while ((record = reader.Read()) != null)
            {
                switch (record)
                {
                    case SharedStringRecord pr:
                        SST.Add(pr.Value);
                        break;
                }
            }
        }

        private void ReadStyles()
        {
            using var reader = _zipWorker.GetStylesReader();
            if (reader == null)
                return;

            Record record;
            while ((record = reader.Read()) != null)
            {
                switch (record)
                {
                    case ExtendedFormatRecord xf:
                        ExtendedFormats.Add(xf.ExtendedFormat);
                        break;
                    case CellStyleExtendedFormatRecord csxf:
                        CellStyleExtendedFormats.Add(csxf.ExtendedFormat);
                        break;
                    case NumberFormatRecord nf:
                        AddNumberFormat(nf.FormatIndexInFile, nf.FormatString);
                        break;
                }
            }
        }
    }
}


namespace ExcelDataReader.Core.OpenXmlFormat
{
    internal class XlsxWorksheet : IWorksheet
    {
        public XlsxWorksheet(ZipWorker document, XlsxWorkbook workbook, SheetRecord refSheet)
        {
            Document = document;
            Workbook = workbook;

            Name = refSheet.Name;
            Id = refSheet.Id;
            Rid = refSheet.Rid;
            VisibleState = refSheet.VisibleState;
            Path = refSheet.Path;
            DefaultRowHeight = 15;

            if (string.IsNullOrEmpty(Path))
                return;

            using var sheetStream = Document.GetWorksheetReader(Path);
            if (sheetStream == null)
                return;

            int rowIndexMaximum = int.MinValue;
            int columnIndexMaximum = int.MinValue;

            List<Column> columnWidths = new List<Column>();
            List<CellRange> cellRanges = new List<CellRange>();

            bool inSheetData = false;

            Record record;
            while ((record = sheetStream.Read()) != null)
            {
                switch (record)
                {
                    case SheetDataBeginRecord _:
                        inSheetData = true;
                        break;
                    case SheetDataEndRecord _:
                        inSheetData = false;
                        break;
                    case RowHeaderRecord row when inSheetData:
                        rowIndexMaximum = Math.Max(rowIndexMaximum, row.RowIndex);
                        break;
                    case CellRecord cell when inSheetData:
                        columnIndexMaximum = Math.Max(columnIndexMaximum, cell.ColumnIndex);
                        break;
                    case ColumnRecord column:
                        columnWidths.Add(column.Column);
                        break;
                    case SheetFormatPrRecord sheetFormatProperties:
                        if (sheetFormatProperties.DefaultRowHeight != null)
                            DefaultRowHeight = sheetFormatProperties.DefaultRowHeight.Value;
                        break;
                    case SheetPrRecord sheetProperties:
                        CodeName = sheetProperties.CodeName;
                        break;
                    case MergeCellRecord mergeCell:
                        cellRanges.Add(mergeCell.Range);
                        break;
                    case HeaderFooterRecord headerFooter:
                        HeaderFooter = headerFooter.HeaderFooter;
                        break;
                }
            }

            ColumnWidths = columnWidths.ToArray();
            MergeCells = cellRanges.ToArray();

            if (rowIndexMaximum != int.MinValue && columnIndexMaximum != int.MinValue)
            {
                FieldCount = columnIndexMaximum + 1;
                RowCount = rowIndexMaximum + 1;
            }
        }

        public int FieldCount { get; }

        public int RowCount { get; }

        public string Name { get; }

        public string CodeName { get; }

        public string VisibleState { get; }

        public HeaderFooter HeaderFooter { get; }

        public double DefaultRowHeight { get; }

        public uint Id { get; }

        public string Rid { get; set; }

        public string Path { get; set; }

        public CellRange[] MergeCells { get; }

        public Column[] ColumnWidths { get; }

        private ZipWorker Document { get; }

        private XlsxWorkbook Workbook { get; }

        public IEnumerable<Row> ReadRows()
        {
            if (string.IsNullOrEmpty(Path))
                yield break;

            using var sheetStream = Document.GetWorksheetReader(Path);
            if (sheetStream == null)
                yield break;

            var rowIndex = 0;
            List<Cell> cells = null;
            double height = 0;

            bool inSheetData = false;
            Record record;
            while ((record = sheetStream.Read()) != null)
            {
                switch (record)
                {
                    case SheetDataBeginRecord _:
                        inSheetData = true;
                        break;
                    case SheetDataEndRecord _:
                        inSheetData = false;
                        break;
                    case RowHeaderRecord row when inSheetData:
                        int currentRowIndex = row.RowIndex;

                        if (cells != null && rowIndex != currentRowIndex)
                        {
                            yield return new Row(rowIndex++, height, cells);
                            cells = null;
                        }

                        if (cells == null)
                        {
                            height = row.Hidden ? 0 : row.Height ?? DefaultRowHeight;
                            cells = new List<Cell>();
                        }

                        for (; rowIndex < currentRowIndex; rowIndex++)
                        {
                            yield return new Row(rowIndex, DefaultRowHeight, new List<Cell>());
                        }

                        break;
                    case CellRecord cell when inSheetData:
                        // TODO What if we get a cell without a row?
                        var extendedFormat = Workbook.GetEffectiveCellStyle(cell.XfIndex, 0);
                        cells.Add(new Cell(cell.ColumnIndex, ConvertCellValue(cell.Value, extendedFormat.NumberFormatIndex), extendedFormat, cell.Error));
                        break;
                }
            }

            if (cells != null)
                yield return new Row(rowIndex, height, cells);
        }

        private object ConvertCellValue(object value, int numberFormatIndex)
        {
            switch (value)
            {
                case int sstIndex:
                    if (sstIndex >= 0 && sstIndex < Workbook.SST.Count)
                    {
                        return Helpers.ConvertEscapeChars(Workbook.SST[sstIndex]);
                    }

                    return null;

                case double number:
                    var format = Workbook.GetNumberFormatString(numberFormatIndex);
                    if (format != null)
                    {
                        if (format.IsDateTimeFormat)
                            return Helpers.ConvertFromOATime(number, Workbook.IsDate1904);
                        if (format.IsTimeSpanFormat)
                            return TimeSpan.FromDays(number);
                    }

                    return number;
                default:
                    return value;
            }
        }
    }
}

#if !NET20
#endif

namespace ExcelDataReader.Core.OpenXmlFormat
{
    internal partial class ZipWorker : IDisposable
    {
        private const string FileSharedStrings = "xl/sharedStrings.{0}";
        private const string FileStyles = "xl/styles.{0}";
        private const string FileWorkbook = "xl/workbook.{0}";
        private const string FileRels = "xl/_rels/workbook.{0}.rels";

        private const string Format = "xml";
        private const string BinFormat = "bin";

        private static readonly XmlReaderSettings XmlSettings = new XmlReaderSettings 
        {
            IgnoreComments = true, 
            IgnoreWhitespace = true,
#if !NETSTANDARD1_3
            XmlResolver = null,
#endif
        };

        private readonly Dictionary<string, ZipArchiveEntry> _entries;
        private bool _disposed;
        private Stream _zipStream;
        private ZipArchive _zipFile;

        /// <summary>
        /// Initializes a new instance of the <see cref="ZipWorker"/> class. 
        /// </summary>
        /// <param name="fileStream">The zip file stream.</param>
        public ZipWorker(Stream fileStream)
        {
            _zipStream = fileStream ?? throw new ArgumentNullException(nameof(fileStream));
            _zipFile = new ZipArchive(fileStream);
            _entries = new Dictionary<string, ZipArchiveEntry>(StringComparer.OrdinalIgnoreCase);
            foreach (var entry in _zipFile.Entries)
            {
                _entries.Add(entry.FullName.Replace('\\', '/'), entry);
            }
        }

        /// <summary>
        /// Gets the shared strings reader.
        /// </summary>
        public RecordReader GetSharedStringsReader()
        {
            var entry = FindEntry(string.Format(FileSharedStrings, Format));
            if (entry != null)
                return new XmlSharedStringsReader(XmlReader.Create(entry.Open(), XmlSettings));

            entry = FindEntry(string.Format(FileSharedStrings, BinFormat));
            if (entry != null)
                return new BiffSharedStringsReader(entry.Open());

            return null;
        }

        /// <summary>
        /// Gets the styles reader.
        /// </summary>
        public RecordReader GetStylesReader()
        {
            var entry = FindEntry(string.Format(FileStyles, Format));
            if (entry != null)
                return new XmlStylesReader(XmlReader.Create(entry.Open(), XmlSettings));

            entry = FindEntry(string.Format(FileStyles, BinFormat));
            if (entry != null)
                return new BiffStylesReader(entry.Open());

            return null;
        }

        /// <summary>
        /// Gets the workbook reader.
        /// </summary>
        public RecordReader GetWorkbookReader()
        {
            var entry = FindEntry(string.Format(FileWorkbook, Format));
            if (entry != null)
                return new XmlWorkbookReader(XmlReader.Create(entry.Open(), XmlSettings));

            entry = FindEntry(string.Format(FileWorkbook, BinFormat));
            if (entry != null)
                return new BiffWorkbookReader(entry.Open());

            throw new Exceptions.HeaderException(Errors.ErrorZipNoOpenXml);
        }

        public RecordReader GetWorksheetReader(string sheetPath)
        {
            // its possible sheetPath starts with /xl. in this case trim the /
            // see the test "Issue_11522_OpenXml"
            if (sheetPath.StartsWith("/xl/", StringComparison.OrdinalIgnoreCase))
                sheetPath = sheetPath.Substring(1);
            else
                sheetPath = "xl/" + sheetPath;

            var zipEntry = FindEntry(sheetPath);
            if (zipEntry != null)
            {
                return Path.GetExtension(sheetPath) switch
                {
                    ".xml" => new XmlWorksheetReader(XmlReader.Create(zipEntry.Open(), XmlSettings)),
                    ".bin" => new BiffWorksheetReader(zipEntry.Open()),
                    _ => null,
                };
            }

            return null;
        }

        /// <summary>
        /// Gets the workbook rels stream.
        /// </summary>
        /// <returns>The rels stream.</returns>
        public Stream GetWorkbookRelsStream()
        {
            var zipEntry = FindEntry(string.Format(FileRels, Format));
            if (zipEntry != null)
                return zipEntry.Open();

            zipEntry = FindEntry(string.Format(FileRels, BinFormat));
            if (zipEntry != null)
                return zipEntry.Open();

            return null;
        }

        private ZipArchiveEntry FindEntry(string name)
        {
            if (_entries.TryGetValue(name, out var entry))
                return entry;
            return null;
        }
    }

    internal partial class ZipWorker
    {
        ~ZipWorker()
        {
            Dispose(false);
        }

        public void Dispose()
        {
            Dispose(true);

            GC.SuppressFinalize(this);
        }

        private void Dispose(bool disposing)
        {
            // Check to see if Dispose has already been called.
            if (!_disposed)
            {
                if (disposing)
                {
                    if (_zipFile != null)
                    {
                        _zipFile.Dispose();
                        _zipFile = null;
                    }

                    if (_zipStream != null)
                    {
                        _zipStream.Dispose();
                        _zipStream = null;
                    }
                }

                _disposed = true;
            }
        }
    }
}


#nullable enable

namespace ExcelDataReader.Core.OpenXmlFormat.BinaryFormat
{
    internal abstract class BiffReader : RecordReader
    {
        private readonly byte[] _buffer = new byte[128];

        public BiffReader(Stream stream)
        {
            Stream = stream ?? throw new ArgumentNullException(nameof(stream));
        }

        protected Stream Stream { get; }

        public override Record? Read()
        {
            if (!TryReadVariableValue(out var recordId) ||
                !TryReadVariableValue(out var recordLength))
                return null;

            byte[] buffer = recordLength < _buffer.Length ? _buffer : new byte[recordLength];
            if (Stream.Read(buffer, 0, (int)recordLength) != recordLength)
                return null;

            return ReadOverride(buffer, recordId, recordLength);
        }

        protected static uint GetDWord(byte[] buffer, uint offset)
        {
            uint result = (uint)buffer[offset + 3] << 24;
            result += (uint)buffer[offset + 2] << 16;
            result += (uint)buffer[offset + 1] << 8;
            result += buffer[offset];
            return result;
        }

        protected static int GetInt32(byte[] buffer, uint offset)
        {
            int result = buffer[offset + 3] << 24;
            result += buffer[offset + 2] << 16;
            result += buffer[offset + 1] << 8;
            result += buffer[offset];
            return result;
        }

        protected static ushort GetWord(byte[] buffer, uint offset)
        {
            ushort result = (ushort)(buffer[offset + 1] << 8);
            result += buffer[offset];
            return result;
        }

        protected static string GetString(byte[] buffer, uint offset, uint length)
        {
            StringBuilder sb = new StringBuilder((int)length);
            for (uint i = offset; i < offset + 2 * length; i += 2)
                sb.Append((char)GetWord(buffer, i));
            return sb.ToString();
        }

        protected static string? GetNullableString(byte[] buffer, ref uint offset)
        {
            var length = GetDWord(buffer, offset);
            offset += 4;
            if (length == uint.MaxValue)
                return null;
            StringBuilder sb = new StringBuilder((int)length);
            uint end = offset + length * 2;
            for (; offset < end; offset += 2)
                sb.Append((char)GetWord(buffer, offset));
            return sb.ToString();
        }

        protected static double GetRkNumber(byte[] buffer, uint offset)
        {
            double result;

            byte flags = buffer[offset];

            if ((flags & 0x02) != 0)
            {
                result = GetInt32(buffer, offset) >> 2;
            }
            else
            {
                result = BitConverter.Int64BitsToDouble((GetDWord(buffer, offset) & -4) << 32);
            }

            if ((flags & 0x01) != 0)
            {
                result /= 100;
            }

            return result;
        }

        protected static double GetDouble(byte[] buffer, uint offset)
        {
            uint num = GetDWord(buffer, offset);
            uint num2 = GetDWord(buffer, offset + 4);
            long num3 = ((long)num2 << 32) | num;
            return BitConverter.Int64BitsToDouble(num3);
        }

        protected abstract Record ReadOverride(byte[] buffer, uint recordId, uint recordLength);

        /// <inheritdoc />
        protected override void Dispose(bool disposing)
        {
            if (disposing)
                Stream.Dispose();
        }

        private bool TryReadVariableValue(out uint value)
        {
            value = 0;

            if (Stream.Read(_buffer, 0, 1) == 0)
                return false;

            byte b1 = _buffer[0];
            value = (uint)(b1 & 0x7F);

            if ((b1 & 0x80) == 0)
                return true;

            if (Stream.Read(_buffer, 0, 1) == 0)
                return false;
            byte b2 = _buffer[0];
            value = ((uint)(b2 & 0x7F) << 7) | value;

            if ((b2 & 0x80) == 0)
                return true;

            if (Stream.Read(_buffer, 0, 1) == 0)
                return false;
            byte b3 = _buffer[0];
            value = ((uint)(b3 & 0x7F) << 14) | value;

            if ((b3 & 0x80) == 0)
                return true;

            if (Stream.Read(_buffer, 0, 1) == 0)
                return false;
            byte b4 = _buffer[0];
            value = ((uint)(b4 & 0x7F) << 21) | value;

            return true;
        }        
    }
}


namespace ExcelDataReader.Core.OpenXmlFormat.BinaryFormat
{
    internal sealed class BiffSharedStringsReader : BiffReader
    {
        private const int StringItem = 0x13;

        public BiffSharedStringsReader(Stream stream) 
            : base(stream)
        {
        }

        protected override Record ReadOverride(byte[] buffer, uint recordId, uint recordLength)
        {
            switch (recordId) 
            {
                case StringItem:
                    // Must be between 0 and 255 characters
                    uint length = GetDWord(buffer, 1);
                    string value = GetString(buffer, 1 + 4, length);
                    return new SharedStringRecord(value);
            }

            return Record.Default;
        }
    }
}


namespace ExcelDataReader.Core.OpenXmlFormat.BinaryFormat
{
    internal sealed class BiffStylesReader : BiffReader
    {
        private const int Xf = 0x2f;

        private const int CellXfStart = 0x269;
        private const int CellXfEnd = 0x26a;

        private const int CellStyleXfStart = 0x272;
        private const int CellStyleXfEnd = 0x273;

        private const int NumberFormatStart = 0x267;
        private const int NumberFormat = 0x2c;
        private const int NumberFormatEnd = 0x268;

        private bool _inCellXf;
        private bool _inCellStyleXf;
        private bool _inNumberFormat;

        public BiffStylesReader(Stream stream)
            : base(stream)
        {
        }

        protected override Record ReadOverride(byte[] buffer, uint recordId, uint recordLength)
        {
            switch (recordId)
            {
                case CellXfStart:
                    _inCellXf = true;
                    break;
                case CellXfEnd:
                    _inCellXf = false;
                    break;
                case CellStyleXfStart:
                    _inCellStyleXf = true;
                    break;
                case CellStyleXfEnd:
                    _inCellStyleXf = false;
                    break;
                case NumberFormatStart:
                    _inNumberFormat = true;
                    break;
                case NumberFormatEnd:
                    _inNumberFormat = false;
                    break;

                case Xf when _inCellXf:
                case Xf when _inCellStyleXf:
                    {
                        var flags = buffer[14];
                        var extendedFormat = new ExtendedFormat()
                        {
                            ParentCellStyleXf = GetWord(buffer, 0),
                            NumberFormatIndex = GetWord(buffer, 2),
                            FontIndex = GetWord(buffer, 4),
                            IndentLevel = (int)(uint)buffer[11],
                            HorizontalAlignment = (HorizontalAlignment)(buffer[12] & 0b111),
                            Locked = (buffer[13] & 0x10000) != 0,
                            Hidden = (buffer[13] & 0x100000) != 0,
                        };

                        if (_inCellXf)
                            return new ExtendedFormatRecord(extendedFormat);
                        return new CellStyleExtendedFormatRecord(extendedFormat);
                    }

                case NumberFormat when _inNumberFormat:
                    {
                        // Must be between 1 and 255 characters
                        int format = GetWord(buffer, 0);
                        uint length = GetDWord(buffer, 2);
                        string formatString = GetString(buffer, 2 + 4, length);

                        return new NumberFormatRecord(format, formatString);
                    }
            }

            return Record.Default;
        }
    }
}


#nullable enable

namespace ExcelDataReader.Core.OpenXmlFormat.BinaryFormat
{
    internal sealed class BiffWorkbookReader : BiffReader
    {
        private const int WorkbookPr = 0x99;
        private const int Sheet = 0x9C;

        public BiffWorkbookReader(Stream stream)
            : base(stream)
        {
        }

        private enum SheetVisibility : byte
        {
            Visible = 0x0,
            Hidden = 0x1,
            VeryHidden = 0x2
        }

        protected override Record ReadOverride(byte[] buffer, uint recordId, uint recordLength)
        {
            switch (recordId)
            {
                case WorkbookPr:
                    return new WorkbookPrRecord((buffer[0] & 0x01) == 1);
                case Sheet: // BrtBundleSh
                    var state = (SheetVisibility)GetDWord(buffer, 0) switch
                    {
                        SheetVisibility.Hidden => "hidden",
                        SheetVisibility.VeryHidden => "veryhidden",
                        _ => "visible"
                    };

                    uint id = GetDWord(buffer, 4);

                    uint offset = 8;
                    string? rid = GetNullableString(buffer, ref offset);

                    // Must be between 1 and 31 characters
                    uint nameLength = GetDWord(buffer, offset);
                    string name = GetString(buffer, offset + 4, nameLength);

                    return new SheetRecord(name, id, rid, state);
                default:
                    return Record.Default;
            }
        }
    }
}


namespace ExcelDataReader.Core.OpenXmlFormat.BinaryFormat
{
    internal sealed class BiffWorksheetReader : BiffReader
    {
        private const uint Row = 0x00; 
        private const uint Blank = 0x01;
        private const uint Number = 0x02;
        private const uint BoolError = 0x03; 
        private const uint Bool = 0x04; 
        private const uint Float = 0x05;
        private const uint String = 0x06;
        private const uint SharedString = 0x07;
        private const uint FormulaString = 0x08;
        private const uint FormulaNumber = 0x09;
        private const uint FormulaBool = 0x0a;
        private const uint FormulaError = 0x0b;

        // private const uint WorksheetBegin = 0x81;
        // private const uint WorksheetEnd = 0x82;
        private const uint SheetDataBegin = 0x91;
        private const uint SheetDataEnd = 0x92;
        private const uint SheetPr = 0x93;
        private const uint SheetFormatPr = 0x1E5;

        // private const uint ColumnsBegin = 0x186;
        private const uint Column = 0x3C; // column info

        // private const uint ColumnsEnd = 0x187;
        private const uint HeaderFooter = 0x1DF;

        // private const uint MergeCellsBegin = 177;
        // private const uint MergeCellsEnd = 178;
        private const uint MergeCell = 176;

        public BiffWorksheetReader(Stream stream) 
            : base(stream)
        {
        }

        protected override Record ReadOverride(byte[] buffer, uint recordId, uint recordLength)
        {
            switch (recordId) 
            {
                case SheetDataBegin:
                    return new SheetDataBeginRecord();
                case SheetDataEnd:
                    return new SheetDataEndRecord();
                case SheetPr: // BrtWsProp
                    {
                        // Must be between 0 and 31 characters
                        uint length = GetDWord(buffer, 19);

                        // To behave the same as when reading an xml based file. 
                        // GetAttribute returns null both if the attribute is missing
                        // or if it is empty.
                        string codeName = length == 0 ? null : GetString(buffer, 19 + 4, length);
                        return new SheetPrRecord(codeName);
                    }

                case SheetFormatPr: // BrtWsFmtInfo 
                    {
                        // TODO Default column width
                        var unsynced = (buffer[8] & 0b1000) != 0;
                        uint? defaultHeight = null;
                        if (unsynced)
                            defaultHeight = GetWord(buffer, 6);
                        return new SheetFormatPrRecord(defaultHeight);
                    }

                case Column: // BrtColInfo 
                    {
                        int minimum = GetInt32(buffer, 0);
                        int maximum = GetInt32(buffer, 4);
                        byte flags = buffer[16];
                        bool hidden = (flags & 0b1) != 0;
                        bool unsynced = (flags & 0b10) != 0;

                        double? width = null;
                        if (unsynced)
                            width = GetDWord(buffer, 8) / 256.0;
                        return new ColumnRecord(new Column(minimum, maximum, hidden, width));
                    }

                case HeaderFooter: // BrtBeginHeaderFooter 
                    {
                        var flags = buffer[0];
                        bool differentOddEven = (flags & 1) != 0;
                        bool differentFirst = (flags & 0b10) != 0;
                        uint offset = 2;
                        var header = GetNullableString(buffer, ref offset);
                        var footer = GetNullableString(buffer, ref offset);
                        var headerEven = GetNullableString(buffer, ref offset);
                        var footerEven = GetNullableString(buffer, ref offset);
                        var headerFirst = GetNullableString(buffer, ref offset);
                        var footerFirst = GetNullableString(buffer, ref offset);
                        return new HeaderFooterRecord(new HeaderFooter(differentFirst, differentOddEven) 
                        {
                            FirstHeader = headerFirst,
                            FirstFooter = footerFirst,
                            OddHeader = header,
                            OddFooter = footer,
                            EvenHeader = headerEven,
                            EvenFooter = footerEven,
                        });
                    }

                case MergeCell:
                    int fromRow = GetInt32(buffer, 0);
                    int toRow = GetInt32(buffer, 4);
                    int fromColumn = GetInt32(buffer, 8);
                    int toColumn = GetInt32(buffer, 12);
                    return new MergeCellRecord(new CellRange(fromColumn, fromRow, toColumn, toRow));
                case Row: // BrtRowHdr 
                    {
                        int rowIndex = GetInt32(buffer, 0);
                        byte flags = buffer[11];
                        bool hidden = (flags & 0b10000) != 0;
                        bool unsynced = (flags & 0b100000) != 0;

                        double? height = null;
                        if (unsynced)
                            height = GetWord(buffer, 8) / 20.0; // Where does 20 come from?

                        // TODO: Default format ?
                        return new RowHeaderRecord(rowIndex, hidden, height);
                    }

                case Blank:
                    return ReadCell(null);
                case BoolError:
                case FormulaError:
                    return ReadCell(null, (CellError)buffer[8]);
                case Number:
                    return ReadCell(GetRkNumber(buffer, 8));
                case Bool:
                case FormulaBool:
                    return ReadCell(buffer[8] == 1);
                case FormulaNumber:
                case Float:
                    return ReadCell(GetDouble(buffer, 8));
                case String:
                case FormulaString:
                    {
                        // Must be less than 32768 characters
                        var length = GetDWord(buffer, 8);
                        return ReadCell(GetString(buffer, 8 + 4, length));
                    }

                case SharedString:
                    return ReadCell((int)GetDWord(buffer, 8));
                default:
                    return Record.Default;
            }

            CellRecord ReadCell(object value, CellError? errorValue = null) 
            {
                int column = (int)GetDWord(buffer, 0);
                uint xfIndex = GetDWord(buffer, 4) & 0xffffff;

                return new CellRecord(column, (int)xfIndex, value, errorValue);
            }
        }
    }
}

namespace ExcelDataReader.Core.OpenXmlFormat.Records
{
    internal sealed class CellRecord : Record
    {
        public CellRecord(int columnIndex, int xfIndex, object value, CellError? error)
        {
            ColumnIndex = columnIndex;
            XfIndex = xfIndex;
            Value = value;
            Error = error;
        }

        public int ColumnIndex { get; }

        public int XfIndex { get; }

        public object Value { get; }

        public CellError? Error { get; }
    }
}


namespace ExcelDataReader.Core.OpenXmlFormat.Records
{
    internal sealed class CellStyleExtendedFormatRecord : Record
    {
        public CellStyleExtendedFormatRecord(ExtendedFormat extendedFormat)
        {
            ExtendedFormat = extendedFormat;
        }

        public ExtendedFormat ExtendedFormat { get; }
    }
}


namespace ExcelDataReader.Core.OpenXmlFormat.Records
{
    internal sealed class ColumnRecord : Record
    {
        public ColumnRecord(Column column)
        {
            Column = column;
        }

        public Column Column { get; }
    }
}


namespace ExcelDataReader.Core.OpenXmlFormat.Records
{
    internal sealed class ExtendedFormatRecord : Record
    {
        public ExtendedFormatRecord(ExtendedFormat extendedFormat) 
        {
            ExtendedFormat = extendedFormat;
        }

        public ExtendedFormat ExtendedFormat { get; }
    }
}


namespace ExcelDataReader.Core.OpenXmlFormat.Records
{
    internal sealed class HeaderFooterRecord : Record
    {
        public HeaderFooterRecord(HeaderFooter headerFooter) 
        {
            HeaderFooter = headerFooter;
        }

        public HeaderFooter HeaderFooter { get; }
    }
}


namespace ExcelDataReader.Core.OpenXmlFormat.Records
{
    internal sealed class MergeCellRecord : Record
    {
        public MergeCellRecord(CellRange range) 
        {
            Range = range;
        }

        public CellRange Range { get; }
    }
}


namespace ExcelDataReader.Core.OpenXmlFormat.Records
{
    internal sealed class NumberFormatRecord : Record
    {
        public NumberFormatRecord(int formatIndexInFile, string formatString) 
        {
            FormatIndexInFile = formatIndexInFile;
            FormatString = formatString;
        }

        public int FormatIndexInFile { get; }

        public string FormatString { get; }
    }
}


namespace ExcelDataReader.Core.OpenXmlFormat.Records
{
    internal abstract class Record
    {
        internal static Record Default { get; } = new DefaultRecord();

        private sealed class DefaultRecord : Record
        {
        }
    }
}


namespace ExcelDataReader.Core.OpenXmlFormat.Records
{
    internal sealed class RowHeaderRecord : Record
    {
        public RowHeaderRecord(int rowIndex, bool hidden, double? height) 
        {
            RowIndex = rowIndex;
            Hidden = hidden;
            Height = height;
        }

        public int RowIndex { get; }

        public bool Hidden { get; }

        public double? Height { get; }
    }
}


namespace ExcelDataReader.Core.OpenXmlFormat.Records
{
    internal sealed class SharedStringRecord : Record
    {
        public SharedStringRecord(string value) 
        {
            Value = value;
        }

        public string Value { get; }
    }
}


namespace ExcelDataReader.Core.OpenXmlFormat.Records
{
    internal sealed class SheetDataBeginRecord : Record 
    {
    }
}


namespace ExcelDataReader.Core.OpenXmlFormat.Records
{
    internal sealed class SheetDataEndRecord : Record 
    {
    }
}


namespace ExcelDataReader.Core.OpenXmlFormat.Records
{
    internal sealed class SheetFormatPrRecord : Record
    {
        public SheetFormatPrRecord(double? defaultRowHeight)
        {
            DefaultRowHeight = defaultRowHeight;
        }

        public double? DefaultRowHeight { get; }
    }
}


namespace ExcelDataReader.Core.OpenXmlFormat.Records
{
    internal sealed class SheetPrRecord : Record
    {
        public SheetPrRecord(string codeName)
        {
            CodeName = codeName;
        }

        public string CodeName { get; }
    }
}


#nullable enable

namespace ExcelDataReader.Core.OpenXmlFormat.Records
{
    internal sealed class SheetRecord : Record
    {
        public SheetRecord(string name, uint id, string? rid, string visibleState)
        {
            Name = name;
            Id = id;
            Rid = rid;
            VisibleState = string.IsNullOrEmpty(visibleState) ? "visible" : visibleState.ToLower();
        }

        public string Name { get; }

        public string VisibleState { get; }

        public uint Id { get; }

        public string? Rid { get; set; }

        public string? Path { get; set; }
    }
}


namespace ExcelDataReader.Core.OpenXmlFormat.Records
{
    internal sealed class WorkbookPrRecord : Record
    {
        public WorkbookPrRecord(bool date1904)
        {
            Date1904 = date1904;
        }

        public bool Date1904 { get; }
    }
}

namespace ExcelDataReader.Core.OpenXmlFormat.XmlFormat
{
    internal static class StringHelper
    {
        private const string NsSpreadsheetMl = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

        private const string ElementT = "t";
        private const string ElementR = "r";

        public static string ReadStringItem(XmlReader reader)
        {
            string result = string.Empty;
            if (!XmlReaderHelper.ReadFirstContent(reader))
            {
                return result;
            }

            while (!reader.EOF)
            {
                if (reader.IsStartElement(ElementT, NsSpreadsheetMl))
                {
                    // There are multiple <t> in a <si>. Concatenate <t> within an <si>.
                    result += reader.ReadElementContentAsString();
                }
                else if (reader.IsStartElement(ElementR, NsSpreadsheetMl))
                {
                    result += ReadRichTextRun(reader);
                }
                else if (!XmlReaderHelper.SkipContent(reader))
                {
                    break;
                }
            }

            return result;
        }

        private static string ReadRichTextRun(XmlReader reader)
        {
            string result = string.Empty;
            if (!XmlReaderHelper.ReadFirstContent(reader))
            {
                return result;
            }

            while (!reader.EOF)
            {
                if (reader.IsStartElement(ElementT, NsSpreadsheetMl))
                {
                    result += reader.ReadElementContentAsString();
                }
                else if (!XmlReaderHelper.SkipContent(reader))
                {
                    break;
                }
            }

            return result;
        }
    }
}


namespace ExcelDataReader.Core.OpenXmlFormat.XmlFormat
{
    internal abstract class XmlRecordReader : RecordReader
    {
        private IEnumerator<Record> _enumerator;

        public XmlRecordReader(XmlReader reader)
        {
            Reader = reader;
        }

        protected XmlReader Reader { get; }

        public override Record Read()
        {
            if (_enumerator == null)
                _enumerator = ReadOverride().GetEnumerator();
            if (_enumerator.MoveNext())
                return _enumerator.Current;
            return null;
        }

        protected abstract IEnumerable<Record> ReadOverride();

        /// <inheritdoc />
        protected override void Dispose(bool disposing)
        {
            _enumerator?.Dispose();
#if NET20
            if (disposing)
                Reader.Close();
#else
            if (disposing)
                Reader.Dispose();
#endif
        }
    }
}


namespace ExcelDataReader.Core.OpenXmlFormat.XmlFormat
{
    internal sealed class XmlSharedStringsReader : XmlRecordReader
    {
        private const string NsSpreadsheetMl = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        private const string ElementSst = "sst";
        private const string ElementStringItem = "si";

        public XmlSharedStringsReader(XmlReader reader)
            : base(reader)
        {
        }

        protected override IEnumerable<Record> ReadOverride()
        {
            if (!Reader.IsStartElement(ElementSst, NsSpreadsheetMl))
            {
                yield break;
            }

            if (!XmlReaderHelper.ReadFirstContent(Reader))
            {
                yield break;
            }

            while (!Reader.EOF)
            {
                if (Reader.IsStartElement(ElementStringItem, NsSpreadsheetMl))
                {
                    var value = StringHelper.ReadStringItem(Reader);
                    yield return new SharedStringRecord(value);
                }
                else if (!XmlReaderHelper.SkipContent(Reader))
                {
                    break;
                }
            }
        }
    }
}


namespace ExcelDataReader.Core.OpenXmlFormat.XmlFormat
{
    internal sealed class XmlStylesReader : XmlRecordReader
    {
        private const string NsSpreadsheetMl = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

        private const string ElementStyleSheet = "styleSheet";

        private const string ANumFmtId = "numFmtId";

        private const string ElementCellCrossReference = "cellXfs";
        private const string ElementCellStyleCrossReference = "cellStyleXfs";
        private const string NXF = "xf";
        private const string AXFId = "xfId";
        private const string AApplyNumberFormat = "applyNumberFormat";
        private const string AApplyAlignment = "applyAlignment";
        private const string AApplyProtection = "applyProtection";

        private const string ElementNumberFormats = "numFmts";
        private const string NNumFmt = "numFmt";
        private const string AFormatCode = "formatCode";

        private const string NAlignment = "alignment";
        private const string AIndent = "indent";
        private const string AHorizontal = "horizontal";

        private const string NProtection = "protection";
        private const string AHidden = "hidden";
        private const string ALocked = "locked";

        public XmlStylesReader(XmlReader reader) 
            : base(reader)
        {
        }

        protected override IEnumerable<Record> ReadOverride()
        {
            if (!Reader.IsStartElement(ElementStyleSheet, NsSpreadsheetMl))
            {
                yield break;
            }

            if (!XmlReaderHelper.ReadFirstContent(Reader))
            {
                yield break;
            }

            while (!Reader.EOF)
            {
                if (Reader.IsStartElement(ElementCellCrossReference, NsSpreadsheetMl))
                {
                    foreach (var xf in ReadCellXfs())
                        yield return new ExtendedFormatRecord(xf);
                }
                else if (Reader.IsStartElement(ElementCellStyleCrossReference, NsSpreadsheetMl))
                {
                    foreach (var xf in ReadCellXfs())
                        yield return new CellStyleExtendedFormatRecord(xf);
                }
                else if (Reader.IsStartElement(ElementNumberFormats, NsSpreadsheetMl))
                {
                    if (!XmlReaderHelper.ReadFirstContent(Reader))
                    {
                        continue;
                    }

                    while (!Reader.EOF)
                    {
                        if (Reader.IsStartElement(NNumFmt, NsSpreadsheetMl))
                        {
                            int.TryParse(Reader.GetAttribute(ANumFmtId), out var numFmtId);
                            var formatCode = Reader.GetAttribute(AFormatCode);

                            yield return new NumberFormatRecord(numFmtId, formatCode);
                            Reader.Skip();
                        }
                        else if (!XmlReaderHelper.SkipContent(Reader))
                        {
                            break;
                        }
                    }
                }
                else if (!XmlReaderHelper.SkipContent(Reader))
                {
                    break;
                }
            }
        }

        private IEnumerable<ExtendedFormat> ReadCellXfs()
        {
            if (!XmlReaderHelper.ReadFirstContent(Reader))
            {
                yield break;
            }

            while (!Reader.EOF)
            {
                if (Reader.IsStartElement(NXF, NsSpreadsheetMl))
                {
                    int.TryParse(Reader.GetAttribute(AXFId), out var xfId);
                    int.TryParse(Reader.GetAttribute(ANumFmtId), out var numFmtId);
                    var applyNumberFormat = Reader.GetAttribute(AApplyNumberFormat) == "1";
                    var applyAlignment = Reader.GetAttribute(AApplyAlignment) == "1";
                    var applyProtection = Reader.GetAttribute(AApplyProtection) == "1";
                    ReadAlignment(Reader, out int indentLevel, out HorizontalAlignment horizontalAlignment, out var hidden, out var locked);

                    yield return new ExtendedFormat()
                    {
                        FontIndex = -1,
                        ParentCellStyleXf = xfId,
                        NumberFormatIndex = numFmtId,
                        HorizontalAlignment = horizontalAlignment,
                        IndentLevel = indentLevel,
                        Hidden = hidden,
                        Locked = locked,
                    };

                    // reader.Skip();
                }
                else if (!XmlReaderHelper.SkipContent(Reader))
                {
                    break;
                }
            }
        }

        private void ReadAlignment(XmlReader reader, out int indentLevel, out HorizontalAlignment horizontalAlignment, out bool hidden, out bool locked)
        {
            indentLevel = 0;
            horizontalAlignment = HorizontalAlignment.General;
            hidden = false;
            locked = false;

            if (!XmlReaderHelper.ReadFirstContent(reader))
            {
                return;
            }

            while (!reader.EOF)
            {
                if (reader.IsStartElement(NAlignment, NsSpreadsheetMl))
                {
                    int.TryParse(reader.GetAttribute(AIndent), out indentLevel);
                    try
                    {
                        horizontalAlignment = (HorizontalAlignment)Enum.Parse(typeof(HorizontalAlignment), reader.GetAttribute(AHorizontal), true);
                    }
                    catch (ArgumentException)
                    {
                    }
                    catch (OverflowException)
                    {
                    }

                    reader.Skip();
                }
                else if (reader.IsStartElement(NProtection, NsSpreadsheetMl))
                {
                    locked = reader.GetAttribute(ALocked) == "1";
                    hidden = reader.GetAttribute(AHidden) == "1";
                    reader.Skip();
                }
                else if (!XmlReaderHelper.SkipContent(reader))
                {
                    break;
                }
            }
        }
    }
}



namespace ExcelDataReader.Core.OpenXmlFormat.XmlFormat
{
    internal sealed class XmlWorkbookReader : XmlRecordReader
    {
        private const string NsSpreadsheetMl = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        private const string NsDocumentRelationship = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

        private const string ElementWorkbook = "workbook";
        private const string ElementWorkbookProperties = "workbookPr";
        private const string ElementSheets = "sheets";
        private const string ElementSheet = "sheet";

        private const string AttributeSheetId = "sheetId";
        private const string AttributeVisibleState = "state";
        private const string AttributeName = "name";
        private const string AttributeRelationshipId = "id";

        public XmlWorkbookReader(XmlReader reader)
            : base(reader)
        {
        }

        protected override IEnumerable<Record> ReadOverride()
        {
            if (!Reader.IsStartElement(ElementWorkbook, NsSpreadsheetMl))
            {
                yield break;
            }

            if (!XmlReaderHelper.ReadFirstContent(Reader))
            {
                yield break;
            }

            while (!Reader.EOF)
            {
                if (Reader.IsStartElement(ElementWorkbookProperties, NsSpreadsheetMl))
                {
                    // Workbook VBA CodeName: reader.GetAttribute("codeName");
                    bool date1904 = Reader.GetAttribute("date1904") == "1";
                    yield return new WorkbookPrRecord(date1904);
                    Reader.Skip();
                }
                else if (Reader.IsStartElement(ElementSheets, NsSpreadsheetMl))
                {
                    if (!XmlReaderHelper.ReadFirstContent(Reader))
                    {
                        continue;
                    }

                    while (!Reader.EOF)
                    {
                        if (Reader.IsStartElement(ElementSheet, NsSpreadsheetMl))
                        {
                            yield return new SheetRecord(
                                Reader.GetAttribute(AttributeName),
                                uint.Parse(Reader.GetAttribute(AttributeSheetId)),
                                Reader.GetAttribute(AttributeRelationshipId, NsDocumentRelationship),
                                Reader.GetAttribute(AttributeVisibleState));
                            Reader.Skip();
                        }
                        else if (!XmlReaderHelper.SkipContent(Reader))
                        {
                            break;
                        }
                    }
                }
                else if (!XmlReaderHelper.SkipContent(Reader))
                {
                    yield break;
                }
            }
        }
    }
}


namespace ExcelDataReader.Core.OpenXmlFormat.XmlFormat
{
    internal sealed class XmlWorksheetReader : XmlRecordReader
    {
        private const string NsSpreadsheetMl = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

        private const string NWorksheet = "worksheet";
        private const string NSheetData = "sheetData";
        private const string NRow = "row";
        private const string ARef = "ref";
        private const string AR = "r";
        private const string NV = "v";
        private const string NIs = "is";
        private const string AT = "t";
        private const string AS = "s";

        private const string NC = "c"; // cell
        private const string NInlineStr = "inlineStr";
        private const string NStr = "str";

        private const string NMergeCells = "mergeCells";

        private const string NSheetProperties = "sheetPr";
        private const string NSheetFormatProperties = "sheetFormatPr";
        private const string ADefaultRowHeight = "defaultRowHeight";

        private const string NHeaderFooter = "headerFooter";
        private const string ADifferentFirst = "differentFirst";
        private const string ADifferentOddEven = "differentOddEven";
        private const string NFirstHeader = "firstHeader";
        private const string NFirstFooter = "firstFooter";
        private const string NOddHeader = "oddHeader";
        private const string NOddFooter = "oddFooter";
        private const string NEvenHeader = "evenHeader";
        private const string NEvenFooter = "evenFooter";

        private const string NCols = "cols";
        private const string NCol = "col";
        private const string AMin = "min";
        private const string AMax = "max";
        private const string AHidden = "hidden";
        private const string AWidth = "width";
        private const string ACustomWidth = "customWidth";

        private const string NMergeCell = "mergeCell";

        private const string ACustomHeight = "customHeight";
        private const string AHt = "ht";

        public XmlWorksheetReader(XmlReader reader) 
            : base(reader)
        {
        }

        protected override IEnumerable<Record> ReadOverride()
        {
            if (!Reader.IsStartElement(NWorksheet, NsSpreadsheetMl))
            {
                yield break;
            }

            if (!XmlReaderHelper.ReadFirstContent(Reader))
            {
                yield break;
            }

            while (!Reader.EOF)
            {
                if (Reader.IsStartElement(NSheetData, NsSpreadsheetMl))
                {
                    yield return new SheetDataBeginRecord();
                    if (!XmlReaderHelper.ReadFirstContent(Reader))
                    {
                        yield return new SheetDataEndRecord();
                        continue;
                    }

                    int rowIndex = -1;
                    while (!Reader.EOF)
                    {
                        if (Reader.IsStartElement(NRow, NsSpreadsheetMl))
                        {
                            if (int.TryParse(Reader.GetAttribute(AR), out int arValue))
                                rowIndex = arValue - 1; // The row attribute is 1-based
                            else
                                rowIndex++;

                            int.TryParse(Reader.GetAttribute(AHidden), out int hidden);
                            int.TryParse(Reader.GetAttribute(ACustomHeight), out int customHeight);
                            double? height;
                            if (customHeight != 0 && double.TryParse(Reader.GetAttribute(AHt), NumberStyles.Any, CultureInfo.InvariantCulture, out var ahtValue))
                                height = ahtValue;
                            else
                                height = null;

                            yield return new RowHeaderRecord(rowIndex, hidden != 0, height);

                            if (!XmlReaderHelper.ReadFirstContent(Reader))
                            {
                                continue;
                            }

                            int nextColumnIndex = 0;
                            while (!Reader.EOF)
                            {
                                if (Reader.IsStartElement(NC, NsSpreadsheetMl))
                                {
                                    var cell = ReadCell(nextColumnIndex);
                                    nextColumnIndex = cell.ColumnIndex + 1;
                                    yield return cell;
                                }
                                else if (!XmlReaderHelper.SkipContent(Reader))
                                {
                                    break;
                                }
                            }
                        }
                        else if (!XmlReaderHelper.SkipContent(Reader))
                        {
                            break;
                        }
                    }

                    yield return new SheetDataEndRecord();
                }
                else if (Reader.IsStartElement(NMergeCells, NsSpreadsheetMl))
                {
                    if (!XmlReaderHelper.ReadFirstContent(Reader))
                    {
                        continue;
                    }

                    while (!Reader.EOF)
                    {
                        if (Reader.IsStartElement(NMergeCell, NsSpreadsheetMl))
                        {
                            var cellRefs = Reader.GetAttribute(ARef);
                            yield return new MergeCellRecord(new CellRange(cellRefs));

                            Reader.Skip();
                        }
                        else if (!XmlReaderHelper.SkipContent(Reader))
                        {
                            break;
                        }
                    }
                }
                else if (Reader.IsStartElement(NHeaderFooter, NsSpreadsheetMl))
                {
                    var result = ReadHeaderFooter();
                    if (result != null)
                        yield return new HeaderFooterRecord(result);
                }
                else if (Reader.IsStartElement(NCols, NsSpreadsheetMl))
                {
                    if (!XmlReaderHelper.ReadFirstContent(Reader))
                    {
                        continue;
                    }

                    while (!Reader.EOF)
                    {
                        if (Reader.IsStartElement(NCol, NsSpreadsheetMl))
                        {
                            var min = Reader.GetAttribute(AMin);
                            var max = Reader.GetAttribute(AMax);
                            var width = Reader.GetAttribute(AWidth);
                            var customWidth = Reader.GetAttribute(ACustomWidth);
                            var hidden = Reader.GetAttribute(AHidden);

                            var maxVal = int.Parse(max);
                            var minVal = int.Parse(min);
                            var widthVal = double.Parse(width, CultureInfo.InvariantCulture);

                            // Note: column indexes need to be converted to be zero-indexed
                            yield return new ColumnRecord(new Column(minVal - 1, maxVal - 1, hidden == "1", customWidth == "1" ? (double?)widthVal : null));

                            Reader.Skip();
                        }
                        else if (!XmlReaderHelper.SkipContent(Reader))
                        {
                            break;
                        }
                    }
                }
                else if (Reader.IsStartElement(NSheetProperties, NsSpreadsheetMl))
                {
                    var codeName = Reader.GetAttribute("codeName");
                    yield return new SheetPrRecord(codeName);

                    Reader.Skip();
                }
                else if (Reader.IsStartElement(NSheetFormatProperties, NsSpreadsheetMl))
                {
                    if (double.TryParse(Reader.GetAttribute(ADefaultRowHeight), NumberStyles.Any, CultureInfo.InvariantCulture, out var defaultRowHeight))
                        yield return new SheetFormatPrRecord(defaultRowHeight);

                    Reader.Skip();
                }
                else if (!XmlReaderHelper.SkipContent(Reader))
                {
                    break;
                }
            }
        }

        private HeaderFooter ReadHeaderFooter()
        {
            var differentFirst = Reader.GetAttribute(ADifferentFirst) == "1";
            var differentOddEven = Reader.GetAttribute(ADifferentOddEven) == "1";

            if (!XmlReaderHelper.ReadFirstContent(Reader))
            {
                return null;
            }

            var headerFooter = new HeaderFooter(differentFirst, differentOddEven);

            while (!Reader.EOF)
            {
                if (Reader.IsStartElement(NOddHeader, NsSpreadsheetMl))
                {
                    headerFooter.OddHeader = Reader.ReadElementContentAsString();
                }
                else if (Reader.IsStartElement(NOddFooter, NsSpreadsheetMl))
                {
                    headerFooter.OddFooter = Reader.ReadElementContentAsString();
                }
                else if (Reader.IsStartElement(NEvenHeader, NsSpreadsheetMl))
                {
                    headerFooter.EvenHeader = Reader.ReadElementContentAsString();
                }
                else if (Reader.IsStartElement(NEvenFooter, NsSpreadsheetMl))
                {
                    headerFooter.EvenFooter = Reader.ReadElementContentAsString();
                }
                else if (Reader.IsStartElement(NFirstHeader, NsSpreadsheetMl))
                {
                    headerFooter.FirstHeader = Reader.ReadElementContentAsString();
                }
                else if (Reader.IsStartElement(NFirstFooter, NsSpreadsheetMl))
                {
                    headerFooter.FirstFooter = Reader.ReadElementContentAsString();
                }
                else if (!XmlReaderHelper.SkipContent(Reader))
                {
                    break;
                }
            }

            return headerFooter;
        }

        private CellRecord ReadCell(int nextColumnIndex)
        {
            int columnIndex;
            int xfIndex = -1;

            var aS = Reader.GetAttribute(AS);
            var aT = Reader.GetAttribute(AT);
            var aR = Reader.GetAttribute(AR);

            if (ReferenceHelper.ParseReference(aR, out int referenceColumn, out _))
                columnIndex = referenceColumn - 1; // ParseReference is 1-based
            else
                columnIndex = nextColumnIndex;

            if (aS != null)
            {
                if (int.TryParse(aS, NumberStyles.Any, CultureInfo.InvariantCulture, out var styleIndex))
                {
                    xfIndex = styleIndex;
                }
            }

            if (!XmlReaderHelper.ReadFirstContent(Reader))
            {
                return new CellRecord(columnIndex, xfIndex, null, null);
            }

            object value = null;
            CellError? error = null;
            while (!Reader.EOF)
            {
                if (Reader.IsStartElement(NV, NsSpreadsheetMl))
                {
                    string rawValue = Reader.ReadElementContentAsString();
                    if (!string.IsNullOrEmpty(rawValue))
                        ConvertCellValue(rawValue, aT, out value, out error);
                }
                else if (Reader.IsStartElement(NIs, NsSpreadsheetMl))
                {
                    string rawValue = StringHelper.ReadStringItem(Reader);
                    if (!string.IsNullOrEmpty(rawValue))
                        ConvertCellValue(rawValue, aT, out value, out error);
                }
                else if (!XmlReaderHelper.SkipContent(Reader))
                {
                    break;
                }
            }

            return new CellRecord(columnIndex, xfIndex, value, error);
        }

        private void ConvertCellValue(string rawValue, string aT, out object value, out CellError? error)
        {
            const NumberStyles style = NumberStyles.Any;
            var invariantCulture = CultureInfo.InvariantCulture;

            error = null;
            switch (aT)
            {
                case AS: //// if string
                    if (int.TryParse(rawValue, style, invariantCulture, out var sstIndex))
                    {
                        // TODO: Can we get here when the sstIndex is not a valid index in the SST list?
                        value = sstIndex;
                        return;
                    }

                    value = rawValue;
                    return;
                case NInlineStr: //// if string inline
                case NStr: //// if cached formula string
                    value = Helpers.ConvertEscapeChars(rawValue);
                    return;
                case "b": //// boolean
                    value = rawValue == "1";
                    return;
                case "d": //// ISO 8601 date
                    if (DateTime.TryParseExact(rawValue, "yyyy-MM-dd", invariantCulture, DateTimeStyles.AllowLeadingWhite | DateTimeStyles.AllowTrailingWhite, out var date))
                    {
                        value = date;
                        return;
                    }

                    value = rawValue;
                    return;
                case "e": //// error
                    error = ConvertError(rawValue);
                    value = null;
                    return;
                default:
                    if (double.TryParse(rawValue, style, invariantCulture, out double number))
                    {
                        value = number;
                        return;
                    }

                    value = rawValue;
                    return;
            }
        }

        private CellError? ConvertError(string e)
        {
            // 2.5.97.2 BErr
            switch (e)
            {
                case "#NULL!":
                    return CellError.NULL;
                case "#DIV/0!":
                    return CellError.DIV0;
                case "#VALUE!":
                    return CellError.VALUE;
                case "#REF!":
                    return CellError.REF;
                case "#NAME?":
                    return CellError.NAME;
                case "#NUM!":
                    return CellError.NUM;
                case "#N/A":
                    return CellError.NA;
                case "#GETTING_DATA":
                    return CellError.GETTING_DATA;
                default:
                    return null;
            }
        }
    }
}


namespace ExcelDataReader.Log.Logger
{
    /// <summary>
    /// The default logger until one is set.
    /// </summary>
    public struct NullLogFactory : ILogFactory, ILog
    {
        /// <inheritdoc />
        public void Debug(string message, params object[] formatting)
        {
        }

        /// <inheritdoc />
        public void Info(string message, params object[] formatting)
        {
        }

        /// <inheritdoc />
        public void Warn(string message, params object[] formatting)
        {
        }

        /// <inheritdoc />
        public void Error(string message, params object[] formatting)
        {
        }

        /// <inheritdoc />
        public void Fatal(string message, params object[] formatting)
        {
        }

        /// <inheritdoc />
        public ILog Create(Type loggingType)
        {
            return this;
        }
    }
}

#if NET20
// ReSharper disable once CheckNamespace
namespace System.Runtime.CompilerServices
{
    [AttributeUsage(AttributeTargets.Assembly | AttributeTargets.Class | AttributeTargets.Method, AllowMultiple = false, Inherited = false)]
    internal class ExtensionAttribute : Attribute
    {
    }
}
#endif

#if NET20
namespace ExcelDataReader
{
    /// <summary>
    /// Encapsulates a method that has one parameter and returns a value of the type specified by the TResult parameter.
    /// </summary>
    /// <typeparam name="T1">The type of the parameter of the method that this delegate encapsulates.</typeparam>
    /// <typeparam name="TResult">The type of the return value of the method that this delegate encapsulates.</typeparam>
    /// <param name="arg1">The parameter of the method that this delegate encapsulates.</param>
    /// <returns>The return value of the method that this delegate encapsulates.</returns>
    public delegate TResult Func<T1, TResult>(T1 arg1);

    /// <summary>
    /// Encapsulates a method that has two parameters and returns a value of the type specified by the TResult parameter.
    /// </summary>
    /// <typeparam name="T1">The type of the first parameter of the method that this delegate encapsulates.</typeparam>
    /// <typeparam name="T2">The type of the second parameter of the method that this delegate encapsulates.</typeparam>
    /// <typeparam name="TResult">The type of the return value of the method that this delegate encapsulates</typeparam>
    /// <param name="arg1">The first parameter of the method that this delegate encapsulates.</param>
    /// <param name="arg2">The second parameter of the method that this delegate encapsulates.</param>
    /// <returns>The return value of the method that this delegate encapsulates.</returns>
    public delegate TResult Func<T1, T2, TResult>(T1 arg1, T2 arg2);

    /// <summary>
    /// Encapsulates a method that has two parameters and returns a value of the type specified by the TResult parameter.
    /// </summary>
    /// <typeparam name="T1">The type of the first parameter of the method that this delegate encapsulates.</typeparam>
    /// <typeparam name="T2">The type of the second parameter of the method that this delegate encapsulates.</typeparam>
    /// <typeparam name="T3">The type of the third parameter of the method that this delegate encapsulates.</typeparam>
    /// <typeparam name="TResult">The type of the return value of the method that this delegate encapsulates</typeparam>
    /// <param name="arg1">The first parameter of the method that this delegate encapsulates.</param>
    /// <param name="arg2">The second parameter of the method that this delegate encapsulates.</param>
    /// <param name="arg3">The third parameter of the method that this delegate encapsulates.</param>
    /// <returns>The return value of the method that this delegate encapsulates.</returns>
    public delegate TResult Func<T1, T2, T3, TResult>(T1 arg1, T2 arg2, T3 arg3);
}
#endif

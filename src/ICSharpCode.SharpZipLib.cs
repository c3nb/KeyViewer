using System;
using System.IO;
using System.Runtime.Serialization;
using System.Runtime.CompilerServices;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
using System.Linq;
using System.Security.Cryptography;
using System.Threading;
using System.Threading.Tasks;
using ICSharpCode.SharpZipLib.Zip;
using ICSharpCode.SharpZipLib.Checksum;
using ICSharpCode.SharpZipLib.Core;
using ICSharpCode.SharpZipLib.Encryption;
using ICSharpCode.SharpZipLib.Zip.Compression;
using ICSharpCode.SharpZipLib.Zip.Compression.Streams;
using System.Numerics;
using static ICSharpCode.SharpZipLib.Zip.Compression.Deflater;
using static ICSharpCode.SharpZipLib.Zip.ZipEntryFactory;
using System.Runtime.InteropServices;
using System.Collections;
namespace ICSharpCode.SharpZipLib.BZip2
{
	/// <summary>
	/// An example class to demonstrate compression and decompression of BZip2 streams.
	/// </summary>
	public static class BZip2
	{
		/// <summary>
		/// Decompress the <paramref name="inStream">input</paramref> writing
		/// uncompressed data to the <paramref name="outStream">output stream</paramref>
		/// </summary>
		/// <param name="inStream">The readable stream containing data to decompress.</param>
		/// <param name="outStream">The output stream to receive the decompressed data.</param>
		/// <param name="isStreamOwner">Both streams are closed on completion if true.</param>
		public static void Decompress(Stream inStream, Stream outStream, bool isStreamOwner)
		{
			if (inStream == null)
				throw new ArgumentNullException(nameof(inStream));

			if (outStream == null)
				throw new ArgumentNullException(nameof(outStream));

			try
			{
				using (BZip2InputStream bzipInput = new BZip2InputStream(inStream))
				{
					bzipInput.IsStreamOwner = isStreamOwner;
					Core.StreamUtils.Copy(bzipInput, outStream, new byte[4096]);
				}
			}
			finally
			{
				if (isStreamOwner)
				{
					// inStream is closed by the BZip2InputStream if stream owner
					outStream.Dispose();
				}
			}
		}

		/// <summary>
		/// Compress the <paramref name="inStream">input stream</paramref> sending
		/// result data to <paramref name="outStream">output stream</paramref>
		/// </summary>
		/// <param name="inStream">The readable stream to compress.</param>
		/// <param name="outStream">The output stream to receive the compressed data.</param>
		/// <param name="isStreamOwner">Both streams are closed on completion if true.</param>
		/// <param name="level">Block size acts as compression level (1 to 9) with 1 giving
		/// the lowest compression and 9 the highest.</param>
		public static void Compress(Stream inStream, Stream outStream, bool isStreamOwner, int level)
		{
			if (inStream == null)
				throw new ArgumentNullException(nameof(inStream));

			if (outStream == null)
				throw new ArgumentNullException(nameof(outStream));

			try
			{
				using (BZip2OutputStream bzipOutput = new BZip2OutputStream(outStream, level))
				{
					bzipOutput.IsStreamOwner = isStreamOwner;
					Core.StreamUtils.Copy(inStream, bzipOutput, new byte[4096]);
				}
			}
			finally
			{
				if (isStreamOwner)
				{
					// outStream is closed by the BZip2OutputStream if stream owner
					inStream.Dispose();
				}
			}
		}
	}
}

namespace ICSharpCode.SharpZipLib.BZip2
{
	/// <summary>
	/// Defines internal values for both compression and decompression
	/// </summary>
	internal static class BZip2Constants
	{
		/// <summary>
		/// Random numbers used to randomise repetitive blocks
		/// </summary>
		public readonly static int[] RandomNumbers = {
			619, 720, 127, 481, 931, 816, 813, 233, 566, 247,
			985, 724, 205, 454, 863, 491, 741, 242, 949, 214,
			733, 859, 335, 708, 621, 574,  73, 654, 730, 472,
			419, 436, 278, 496, 867, 210, 399, 680, 480,  51,
			878, 465, 811, 169, 869, 675, 611, 697, 867, 561,
			862, 687, 507, 283, 482, 129, 807, 591, 733, 623,
			150, 238,  59, 379, 684, 877, 625, 169, 643, 105,
			170, 607, 520, 932, 727, 476, 693, 425, 174, 647,
			 73, 122, 335, 530, 442, 853, 695, 249, 445, 515,
			909, 545, 703, 919, 874, 474, 882, 500, 594, 612,
			641, 801, 220, 162, 819, 984, 589, 513, 495, 799,
			161, 604, 958, 533, 221, 400, 386, 867, 600, 782,
			382, 596, 414, 171, 516, 375, 682, 485, 911, 276,
			 98, 553, 163, 354, 666, 933, 424, 341, 533, 870,
			227, 730, 475, 186, 263, 647, 537, 686, 600, 224,
			469,  68, 770, 919, 190, 373, 294, 822, 808, 206,
			184, 943, 795, 384, 383, 461, 404, 758, 839, 887,
			715,  67, 618, 276, 204, 918, 873, 777, 604, 560,
			951, 160, 578, 722,  79, 804,  96, 409, 713, 940,
			652, 934, 970, 447, 318, 353, 859, 672, 112, 785,
			645, 863, 803, 350, 139,  93, 354,  99, 820, 908,
			609, 772, 154, 274, 580, 184,  79, 626, 630, 742,
			653, 282, 762, 623, 680,  81, 927, 626, 789, 125,
			411, 521, 938, 300, 821,  78, 343, 175, 128, 250,
			170, 774, 972, 275, 999, 639, 495,  78, 352, 126,
			857, 956, 358, 619, 580, 124, 737, 594, 701, 612,
			669, 112, 134, 694, 363, 992, 809, 743, 168, 974,
			944, 375, 748,  52, 600, 747, 642, 182, 862,  81,
			344, 805, 988, 739, 511, 655, 814, 334, 249, 515,
			897, 955, 664, 981, 649, 113, 974, 459, 893, 228,
			433, 837, 553, 268, 926, 240, 102, 654, 459,  51,
			686, 754, 806, 760, 493, 403, 415, 394, 687, 700,
			946, 670, 656, 610, 738, 392, 760, 799, 887, 653,
			978, 321, 576, 617, 626, 502, 894, 679, 243, 440,
			680, 879, 194, 572, 640, 724, 926,  56, 204, 700,
			707, 151, 457, 449, 797, 195, 791, 558, 945, 679,
			297,  59,  87, 824, 713, 663, 412, 693, 342, 606,
			134, 108, 571, 364, 631, 212, 174, 643, 304, 329,
			343,  97, 430, 751, 497, 314, 983, 374, 822, 928,
			140, 206,  73, 263, 980, 736, 876, 478, 430, 305,
			170, 514, 364, 692, 829,  82, 855, 953, 676, 246,
			369, 970, 294, 750, 807, 827, 150, 790, 288, 923,
			804, 378, 215, 828, 592, 281, 565, 555, 710,  82,
			896, 831, 547, 261, 524, 462, 293, 465, 502,  56,
			661, 821, 976, 991, 658, 869, 905, 758, 745, 193,
			768, 550, 608, 933, 378, 286, 215, 979, 792, 961,
			 61, 688, 793, 644, 986, 403, 106, 366, 905, 644,
			372, 567, 466, 434, 645, 210, 389, 550, 919, 135,
			780, 773, 635, 389, 707, 100, 626, 958, 165, 504,
			920, 176, 193, 713, 857, 265, 203,  50, 668, 108,
			645, 990, 626, 197, 510, 357, 358, 850, 858, 364,
			936, 638
		};

		/// <summary>
		/// When multiplied by compression parameter (1-9) gives the block size for compression
		/// 9 gives the best compression but uses the most memory.
		/// </summary>
		public const int BaseBlockSize = 100000;

		/// <summary>
		/// Backend constant
		/// </summary>
		public const int MaximumAlphaSize = 258;

		/// <summary>
		/// Backend constant
		/// </summary>
		public const int MaximumCodeLength = 23;

		/// <summary>
		/// Backend constant
		/// </summary>
		public const int RunA = 0;

		/// <summary>
		/// Backend constant
		/// </summary>
		public const int RunB = 1;

		/// <summary>
		/// Backend constant
		/// </summary>
		public const int GroupCount = 6;

		/// <summary>
		/// Backend constant
		/// </summary>
		public const int GroupSize = 50;

		/// <summary>
		/// Backend constant
		/// </summary>
		public const int NumberOfIterations = 4;

		/// <summary>
		/// Backend constant
		/// </summary>
		public const int MaximumSelectors = (2 + (900000 / GroupSize));

		/// <summary>
		/// Backend constant
		/// </summary>
		public const int OvershootBytes = 20;
	}
}


namespace ICSharpCode.SharpZipLib.BZip2
{
	/// <summary>
	/// BZip2Exception represents exceptions specific to BZip2 classes and code.
	/// </summary>
	[Serializable]
	public class BZip2Exception : SharpZipBaseException
	{
		/// <summary>
		/// Initialise a new instance of <see cref="BZip2Exception" />.
		/// </summary>
		public BZip2Exception()
		{
		}

		/// <summary>
		/// Initialise a new instance of <see cref="BZip2Exception" /> with its message string.
		/// </summary>
		/// <param name="message">A <see cref="string"/> that describes the error.</param>
		public BZip2Exception(string message)
			: base(message)
		{
		}

		/// <summary>
		/// Initialise a new instance of <see cref="BZip2Exception" />.
		/// </summary>
		/// <param name="message">A <see cref="string"/> that describes the error.</param>
		/// <param name="innerException">The <see cref="Exception"/> that caused this exception.</param>
		public BZip2Exception(string message, Exception innerException)
			: base(message, innerException)
		{
		}

		/// <summary>
		/// Initializes a new instance of the BZip2Exception class with serialized data.
		/// </summary>
		/// <param name="info">
		/// The System.Runtime.Serialization.SerializationInfo that holds the serialized
		/// object data about the exception being thrown.
		/// </param>
		/// <param name="context">
		/// The System.Runtime.Serialization.StreamingContext that contains contextual information
		/// about the source or destination.
		/// </param>
		protected BZip2Exception(SerializationInfo info, StreamingContext context)
			: base(info, context)
		{
		}
	}
}

#if NETSTANDARD2_1_OR_GREATER || NETCOREAPP3_1_OR_GREATER
	#define VECTORIZE_MEMORY_MOVE
#endif


namespace ICSharpCode.SharpZipLib.BZip2
{
	/// <summary>
	/// An input stream that decompresses files in the BZip2 format
	/// </summary>
	public class BZip2InputStream : Stream
	{
		#region Constants

		private const int START_BLOCK_STATE = 1;
		private const int RAND_PART_A_STATE = 2;
		private const int RAND_PART_B_STATE = 3;
		private const int RAND_PART_C_STATE = 4;
		private const int NO_RAND_PART_A_STATE = 5;
		private const int NO_RAND_PART_B_STATE = 6;
		private const int NO_RAND_PART_C_STATE = 7;

#if VECTORIZE_MEMORY_MOVE
		private static readonly int VectorSize = System.Numerics.Vector<byte>.Count;
#endif // VECTORIZE_MEMORY_MOVE

#endregion Constants

		#region Instance Fields

		/*--
		index of the last char in the block, so
		the block size == last + 1.
		--*/
		private int last;

		/*--
		index in zptr[] of original string after sorting.
		--*/
		private int origPtr;

		/*--
		always: in the range 0 .. 9.
		The current block size is 100000 * this number.
		--*/
		private int blockSize100k;

		private bool blockRandomised;

		private int bsBuff;
		private int bsLive;
		private IChecksum mCrc = new BZip2Crc();

		private bool[] inUse = new bool[256];
		private int nInUse;

		private byte[] seqToUnseq = new byte[256];
		private byte[] unseqToSeq = new byte[256];

		private byte[] selector = new byte[BZip2Constants.MaximumSelectors];
		private byte[] selectorMtf = new byte[BZip2Constants.MaximumSelectors];

		private int[] tt;
		private byte[] ll8;

		/*--
		freq table collected to save a pass over the data
		during decompression.
		--*/
		private int[] unzftab = new int[256];

		private int[][] limit = new int[BZip2Constants.GroupCount][];
		private int[][] baseArray = new int[BZip2Constants.GroupCount][];
		private int[][] perm = new int[BZip2Constants.GroupCount][];
		private int[] minLens = new int[BZip2Constants.GroupCount];

		private readonly Stream baseStream;
		private bool streamEnd;

		private int currentChar = -1;

		private int currentState = START_BLOCK_STATE;

		private int storedBlockCRC, storedCombinedCRC;
		private int computedBlockCRC;
		private uint computedCombinedCRC;

		private int count, chPrev, ch2;
		private int tPos;
		private int rNToGo;
		private int rTPos;
		private int i2, j2;
		private byte z;

		#endregion Instance Fields

		/// <summary>
		/// Construct instance for reading from stream
		/// </summary>
		/// <param name="stream">Data source</param>
		public BZip2InputStream(Stream stream)
		{
			if (stream == null)
				throw new ArgumentNullException(nameof(stream));
			// init arrays
			for (int i = 0; i < BZip2Constants.GroupCount; ++i)
			{
				limit[i] = new int[BZip2Constants.MaximumAlphaSize];
				baseArray[i] = new int[BZip2Constants.MaximumAlphaSize];
				perm[i] = new int[BZip2Constants.MaximumAlphaSize];
			}

			baseStream = stream;
			bsLive = 0;
			bsBuff = 0;
			Initialize();
			InitBlock();
			SetupBlock();
		}

		/// <summary>
		/// Get/set flag indicating ownership of underlying stream.
		/// When the flag is true <see cref="Stream.Dispose()" /> will close the underlying stream also.
		/// </summary>
		public bool IsStreamOwner { get; set; } = true;

		#region Stream Overrides

		/// <summary>
		/// Gets a value indicating if the stream supports reading
		/// </summary>
		public override bool CanRead
		{
			get
			{
				return baseStream.CanRead;
			}
		}

		/// <summary>
		/// Gets a value indicating whether the current stream supports seeking.
		/// </summary>
		public override bool CanSeek
		{
			get
			{
				return false;
			}
		}

		/// <summary>
		/// Gets a value indicating whether the current stream supports writing.
		/// This property always returns false
		/// </summary>
		public override bool CanWrite
		{
			get
			{
				return false;
			}
		}

		/// <summary>
		/// Gets the length in bytes of the stream.
		/// </summary>
		public override long Length
		{
			get
			{
				return baseStream.Length;
			}
		}

		/// <summary>
		/// Gets the current position of the stream.
		/// Setting the position is not supported and will throw a NotSupportException.
		/// </summary>
		/// <exception cref="NotSupportedException">Any attempt to set the position.</exception>
		public override long Position
		{
			get
			{
				return baseStream.Position;
			}
			set
			{
				throw new NotSupportedException("BZip2InputStream position cannot be set");
			}
		}

		/// <summary>
		/// Flushes the stream.
		/// </summary>
		public override void Flush()
		{
			baseStream.Flush();
		}

		/// <summary>
		/// Set the streams position.  This operation is not supported and will throw a NotSupportedException
		/// </summary>
		/// <param name="offset">A byte offset relative to the <paramref name="origin"/> parameter.</param>
		/// <param name="origin">A value of type <see cref="SeekOrigin"/> indicating the reference point used to obtain the new position.</param>
		/// <returns>The new position of the stream.</returns>
		/// <exception cref="NotSupportedException">Any access</exception>
		public override long Seek(long offset, SeekOrigin origin)
		{
			throw new NotSupportedException("BZip2InputStream Seek not supported");
		}

		/// <summary>
		/// Sets the length of this stream to the given value.
		/// This operation is not supported and will throw a NotSupportedExceptionortedException
		/// </summary>
		/// <param name="value">The new length for the stream.</param>
		/// <exception cref="NotSupportedException">Any access</exception>
		public override void SetLength(long value)
		{
			throw new NotSupportedException("BZip2InputStream SetLength not supported");
		}

		/// <summary>
		/// Writes a block of bytes to this stream using data from a buffer.
		/// This operation is not supported and will throw a NotSupportedException
		/// </summary>
		/// <param name="buffer">The buffer to source data from.</param>
		/// <param name="offset">The offset to start obtaining data from.</param>
		/// <param name="count">The number of bytes of data to write.</param>
		/// <exception cref="NotSupportedException">Any access</exception>
		public override void Write(byte[] buffer, int offset, int count)
		{
			throw new NotSupportedException("BZip2InputStream Write not supported");
		}

		/// <summary>
		/// Writes a byte to the current position in the file stream.
		/// This operation is not supported and will throw a NotSupportedException
		/// </summary>
		/// <param name="value">The value to write.</param>
		/// <exception cref="NotSupportedException">Any access</exception>
		public override void WriteByte(byte value)
		{
			throw new NotSupportedException("BZip2InputStream WriteByte not supported");
		}

		/// <summary>
		/// Read a sequence of bytes and advances the read position by one byte.
		/// </summary>
		/// <param name="buffer">Array of bytes to store values in</param>
		/// <param name="offset">Offset in array to begin storing data</param>
		/// <param name="count">The maximum number of bytes to read</param>
		/// <returns>The total number of bytes read into the buffer. This might be less
		/// than the number of bytes requested if that number of bytes are not
		/// currently available or zero if the end of the stream is reached.
		/// </returns>
		public override int Read(byte[] buffer, int offset, int count)
		{
			if (buffer == null)
			{
				throw new ArgumentNullException(nameof(buffer));
			}

			for (int i = 0; i < count; ++i)
			{
				int rb = ReadByte();
				if (rb == -1)
				{
					return i;
				}
				buffer[offset + i] = (byte)rb;
			}
			return count;
		}

		/// <summary>
		/// Closes the stream, releasing any associated resources.
		/// </summary>
		protected override void Dispose(bool disposing)
		{
			if (disposing && IsStreamOwner)
			{
				baseStream.Dispose();
			}
		}

		/// <summary>
		/// Read a byte from stream advancing position
		/// </summary>
		/// <returns>byte read or -1 on end of stream</returns>
		public override int ReadByte()
		{
			if (streamEnd)
			{
				return -1; // ok
			}

			int retChar = currentChar;
			switch (currentState)
			{
				case RAND_PART_B_STATE:
					SetupRandPartB();
					break;

				case RAND_PART_C_STATE:
					SetupRandPartC();
					break;

				case NO_RAND_PART_B_STATE:
					SetupNoRandPartB();
					break;

				case NO_RAND_PART_C_STATE:
					SetupNoRandPartC();
					break;

				case START_BLOCK_STATE:
				case NO_RAND_PART_A_STATE:
				case RAND_PART_A_STATE:
					break;
			}
			return retChar;
		}

		#endregion Stream Overrides

		private void MakeMaps()
		{
			nInUse = 0;
			for (int i = 0; i < 256; ++i)
			{
				if (inUse[i])
				{
					seqToUnseq[nInUse] = (byte)i;
					unseqToSeq[i] = (byte)nInUse;
					nInUse++;
				}
			}
		}

		private void Initialize()
		{
			char magic1 = BsGetUChar();
			char magic2 = BsGetUChar();

			char magic3 = BsGetUChar();
			char magic4 = BsGetUChar();

			if (magic1 != 'B' || magic2 != 'Z' || magic3 != 'h' || magic4 < '1' || magic4 > '9')
			{
				streamEnd = true;
				return;
			}

			SetDecompressStructureSizes(magic4 - '0');
			computedCombinedCRC = 0;
		}

		private void InitBlock()
		{
			char magic1 = BsGetUChar();
			char magic2 = BsGetUChar();
			char magic3 = BsGetUChar();
			char magic4 = BsGetUChar();
			char magic5 = BsGetUChar();
			char magic6 = BsGetUChar();

			if (magic1 == 0x17 && magic2 == 0x72 && magic3 == 0x45 && magic4 == 0x38 && magic5 == 0x50 && magic6 == 0x90)
			{
				Complete();
				return;
			}

			if (magic1 != 0x31 || magic2 != 0x41 || magic3 != 0x59 || magic4 != 0x26 || magic5 != 0x53 || magic6 != 0x59)
			{
				BadBlockHeader();
				streamEnd = true;
				return;
			}

			storedBlockCRC = BsGetInt32();

			blockRandomised = (BsR(1) == 1);

			GetAndMoveToFrontDecode();

			mCrc.Reset();
			currentState = START_BLOCK_STATE;
		}

		private void EndBlock()
		{
			computedBlockCRC = (int)mCrc.Value;

			// -- A bad CRC is considered a fatal error. --
			if (storedBlockCRC != computedBlockCRC)
			{
				CrcError();
			}

			// 1528150659
			computedCombinedCRC = ((computedCombinedCRC << 1) & 0xFFFFFFFF) | (computedCombinedCRC >> 31);
			computedCombinedCRC = computedCombinedCRC ^ (uint)computedBlockCRC;
		}

		private void Complete()
		{
			storedCombinedCRC = BsGetInt32();
			if (storedCombinedCRC != (int)computedCombinedCRC)
			{
				CrcError();
			}

			streamEnd = true;
		}

		private void FillBuffer()
		{
			int thech = 0;

			try
			{
				thech = baseStream.ReadByte();
			}
			catch (Exception)
			{
				CompressedStreamEOF();
			}

			if (thech == -1)
			{
				CompressedStreamEOF();
			}

			bsBuff = (bsBuff << 8) | (thech & 0xFF);
			bsLive += 8;
		}

		private int BsR(int n)
		{
			while (bsLive < n)
			{
				FillBuffer();
			}

			int v = (bsBuff >> (bsLive - n)) & ((1 << n) - 1);
			bsLive -= n;
			return v;
		}

		private char BsGetUChar()
		{
			return (char)BsR(8);
		}

		private int BsGetIntVS(int numBits)
		{
			return BsR(numBits);
		}

		private int BsGetInt32()
		{
			int result = BsR(8);
			result = (result << 8) | BsR(8);
			result = (result << 8) | BsR(8);
			result = (result << 8) | BsR(8);
			return result;
		}

		private void RecvDecodingTables()
		{
			char[][] len = new char[BZip2Constants.GroupCount][];
			for (int i = 0; i < BZip2Constants.GroupCount; ++i)
			{
				len[i] = new char[BZip2Constants.MaximumAlphaSize];
			}

			bool[] inUse16 = new bool[16];

			//--- Receive the mapping table ---
			for (int i = 0; i < 16; i++)
			{
				inUse16[i] = (BsR(1) == 1);
			}

			for (int i = 0; i < 16; i++)
			{
				if (inUse16[i])
				{
					for (int j = 0; j < 16; j++)
					{
						inUse[i * 16 + j] = (BsR(1) == 1);
					}
				}
				else
				{
					for (int j = 0; j < 16; j++)
					{
						inUse[i * 16 + j] = false;
					}
				}
			}

			MakeMaps();
			int alphaSize = nInUse + 2;

			//--- Now the selectors ---
			int nGroups = BsR(3);
			int nSelectors = BsR(15);

			for (int i = 0; i < nSelectors; i++)
			{
				int j = 0;
				while (BsR(1) == 1)
				{
					j++;
				}
				selectorMtf[i] = (byte)j;
			}

			//--- Undo the MTF values for the selectors. ---
			byte[] pos = new byte[BZip2Constants.GroupCount];
			for (int v = 0; v < nGroups; v++)
			{
				pos[v] = (byte)v;
			}

			for (int i = 0; i < nSelectors; i++)
			{
				int v = selectorMtf[i];
				byte tmp = pos[v];
				while (v > 0)
				{
					pos[v] = pos[v - 1];
					v--;
				}
				pos[0] = tmp;
				selector[i] = tmp;
			}

			//--- Now the coding tables ---
			for (int t = 0; t < nGroups; t++)
			{
				int curr = BsR(5);
				for (int i = 0; i < alphaSize; i++)
				{
					while (BsR(1) == 1)
					{
						if (BsR(1) == 0)
						{
							curr++;
						}
						else
						{
							curr--;
						}
					}
					len[t][i] = (char)curr;
				}
			}

			//--- Create the Huffman decoding tables ---
			for (int t = 0; t < nGroups; t++)
			{
				int minLen = 32;
				int maxLen = 0;
				for (int i = 0; i < alphaSize; i++)
				{
					maxLen = Math.Max(maxLen, len[t][i]);
					minLen = Math.Min(minLen, len[t][i]);
				}
				HbCreateDecodeTables(limit[t], baseArray[t], perm[t], len[t], minLen, maxLen, alphaSize);
				minLens[t] = minLen;
			}
		}

		private void GetAndMoveToFrontDecode()
		{
			byte[] yy = new byte[256];
			int nextSym;

			int limitLast = BZip2Constants.BaseBlockSize * blockSize100k;
			origPtr = BsGetIntVS(24);

			RecvDecodingTables();
			int EOB = nInUse + 1;
			int groupNo = -1;
			int groupPos = 0;

			/*--
			Setting up the unzftab entries here is not strictly
			necessary, but it does save having to do it later
			in a separate pass, and so saves a block's worth of
			cache misses.
			--*/
			for (int i = 0; i <= 255; i++)
			{
				unzftab[i] = 0;
			}

			for (int i = 0; i <= 255; i++)
			{
				yy[i] = (byte)i;
			}

			last = -1;

			if (groupPos == 0)
			{
				groupNo++;
				groupPos = BZip2Constants.GroupSize;
			}

			groupPos--;
			int zt = selector[groupNo];
			int zn = minLens[zt];
			int zvec = BsR(zn);
			int zj;

			while (zvec > limit[zt][zn])
			{
				if (zn > 20)
				{ // the longest code
					throw new BZip2Exception("Bzip data error");
				}
				zn++;
				while (bsLive < 1)
				{
					FillBuffer();
				}
				zj = (bsBuff >> (bsLive - 1)) & 1;
				bsLive--;
				zvec = (zvec << 1) | zj;
			}
			if (zvec - baseArray[zt][zn] < 0 || zvec - baseArray[zt][zn] >= BZip2Constants.MaximumAlphaSize)
			{
				throw new BZip2Exception("Bzip data error");
			}
			nextSym = perm[zt][zvec - baseArray[zt][zn]];

			while (true)
			{
				if (nextSym == EOB)
				{
					break;
				}

				if (nextSym == BZip2Constants.RunA || nextSym == BZip2Constants.RunB)
				{
					int s = -1;
					int n = 1;
					do
					{
						if (nextSym == BZip2Constants.RunA)
						{
							s += (0 + 1) * n;
						}
						else if (nextSym == BZip2Constants.RunB)
						{
							s += (1 + 1) * n;
						}

						n <<= 1;

						if (groupPos == 0)
						{
							groupNo++;
							groupPos = BZip2Constants.GroupSize;
						}

						groupPos--;

						zt = selector[groupNo];
						zn = minLens[zt];
						zvec = BsR(zn);

						while (zvec > limit[zt][zn])
						{
							zn++;
							while (bsLive < 1)
							{
								FillBuffer();
							}
							zj = (bsBuff >> (bsLive - 1)) & 1;
							bsLive--;
							zvec = (zvec << 1) | zj;
						}
						nextSym = perm[zt][zvec - baseArray[zt][zn]];
					} while (nextSym == BZip2Constants.RunA || nextSym == BZip2Constants.RunB);

					s++;
					byte ch = seqToUnseq[yy[0]];
					unzftab[ch] += s;

					while (s > 0)
					{
						last++;
						ll8[last] = ch;
						s--;
					}

					if (last >= limitLast)
					{
						BlockOverrun();
					}
					continue;
				}
				else
				{
					last++;
					if (last >= limitLast)
					{
						BlockOverrun();
					}

					byte tmp = yy[nextSym - 1];
					unzftab[seqToUnseq[tmp]]++;
					ll8[last] = seqToUnseq[tmp];

					var j = nextSym - 1;

#if VECTORIZE_MEMORY_MOVE
					// This is vectorized memory move. Going from the back, we're taking chunks of array
					// and write them at the new location shifted by one. Since chunks are VectorSize long,
					// at the end we have to move "tail" (or head actually) of the array using a plain loop.
					// If System.Numerics.Vector API is not available, the plain loop is used to do the whole copying.

					while(j >= VectorSize)
					{
						var arrayPart = new System.Numerics.Vector<byte>(yy, j - VectorSize);
						arrayPart.CopyTo(yy, j - VectorSize + 1);
						j -= VectorSize;
					}
#endif // VECTORIZE_MEMORY_MOVE

					while(j > 0)
					{
						yy[j] = yy[--j];
					}

					yy[0] = tmp;

					if (groupPos == 0)
					{
						groupNo++;
						groupPos = BZip2Constants.GroupSize;
					}

					groupPos--;
					zt = selector[groupNo];
					zn = minLens[zt];
					zvec = BsR(zn);
					while (zvec > limit[zt][zn])
					{
						zn++;
						while (bsLive < 1)
						{
							FillBuffer();
						}
						zj = (bsBuff >> (bsLive - 1)) & 1;
						bsLive--;
						zvec = (zvec << 1) | zj;
					}
					nextSym = perm[zt][zvec - baseArray[zt][zn]];
					continue;
				}
			}
		}

		private void SetupBlock()
		{
			int[] cftab = new int[257];

			cftab[0] = 0;
			Array.Copy(unzftab, 0, cftab, 1, 256);

			for (int i = 1; i <= 256; i++)
			{
				cftab[i] += cftab[i - 1];
			}

			for (int i = 0; i <= last; i++)
			{
				byte ch = ll8[i];
				tt[cftab[ch]] = i;
				cftab[ch]++;
			}

			cftab = null;

			tPos = tt[origPtr];

			count = 0;
			i2 = 0;
			ch2 = 256;   /*-- not a char and not EOF --*/

			if (blockRandomised)
			{
				rNToGo = 0;
				rTPos = 0;
				SetupRandPartA();
			}
			else
			{
				SetupNoRandPartA();
			}
		}

		private void SetupRandPartA()
		{
			if (i2 <= last)
			{
				chPrev = ch2;
				ch2 = ll8[tPos];
				tPos = tt[tPos];
				if (rNToGo == 0)
				{
					rNToGo = BZip2Constants.RandomNumbers[rTPos];
					rTPos++;
					if (rTPos == 512)
					{
						rTPos = 0;
					}
				}
				rNToGo--;
				ch2 ^= (int)((rNToGo == 1) ? 1 : 0);
				i2++;

				currentChar = ch2;
				currentState = RAND_PART_B_STATE;
				mCrc.Update(ch2);
			}
			else
			{
				EndBlock();
				InitBlock();
				SetupBlock();
			}
		}

		private void SetupNoRandPartA()
		{
			if (i2 <= last)
			{
				chPrev = ch2;
				ch2 = ll8[tPos];
				tPos = tt[tPos];
				i2++;

				currentChar = ch2;
				currentState = NO_RAND_PART_B_STATE;
				mCrc.Update(ch2);
			}
			else
			{
				EndBlock();
				InitBlock();
				SetupBlock();
			}
		}

		private void SetupRandPartB()
		{
			if (ch2 != chPrev)
			{
				currentState = RAND_PART_A_STATE;
				count = 1;
				SetupRandPartA();
			}
			else
			{
				count++;
				if (count >= 4)
				{
					z = ll8[tPos];
					tPos = tt[tPos];
					if (rNToGo == 0)
					{
						rNToGo = BZip2Constants.RandomNumbers[rTPos];
						rTPos++;
						if (rTPos == 512)
						{
							rTPos = 0;
						}
					}
					rNToGo--;
					z ^= (byte)((rNToGo == 1) ? 1 : 0);
					j2 = 0;
					currentState = RAND_PART_C_STATE;
					SetupRandPartC();
				}
				else
				{
					currentState = RAND_PART_A_STATE;
					SetupRandPartA();
				}
			}
		}

		private void SetupRandPartC()
		{
			if (j2 < (int)z)
			{
				currentChar = ch2;
				mCrc.Update(ch2);
				j2++;
			}
			else
			{
				currentState = RAND_PART_A_STATE;
				i2++;
				count = 0;
				SetupRandPartA();
			}
		}

		private void SetupNoRandPartB()
		{
			if (ch2 != chPrev)
			{
				currentState = NO_RAND_PART_A_STATE;
				count = 1;
				SetupNoRandPartA();
			}
			else
			{
				count++;
				if (count >= 4)
				{
					z = ll8[tPos];
					tPos = tt[tPos];
					currentState = NO_RAND_PART_C_STATE;
					j2 = 0;
					SetupNoRandPartC();
				}
				else
				{
					currentState = NO_RAND_PART_A_STATE;
					SetupNoRandPartA();
				}
			}
		}

		private void SetupNoRandPartC()
		{
			if (j2 < (int)z)
			{
				currentChar = ch2;
				mCrc.Update(ch2);
				j2++;
			}
			else
			{
				currentState = NO_RAND_PART_A_STATE;
				i2++;
				count = 0;
				SetupNoRandPartA();
			}
		}

		private void SetDecompressStructureSizes(int newSize100k)
		{
			if (!(0 <= newSize100k && newSize100k <= 9 && 0 <= blockSize100k && blockSize100k <= 9))
			{
				throw new BZip2Exception("Invalid block size");
			}

			blockSize100k = newSize100k;

			if (newSize100k == 0)
			{
				return;
			}

			int n = BZip2Constants.BaseBlockSize * newSize100k;
			ll8 = new byte[n];
			tt = new int[n];
		}

		private static void CompressedStreamEOF()
		{
			throw new EndOfStreamException("BZip2 input stream end of compressed stream");
		}

		private static void BlockOverrun()
		{
			throw new BZip2Exception("BZip2 input stream block overrun");
		}

		private static void BadBlockHeader()
		{
			throw new BZip2Exception("BZip2 input stream bad block header");
		}

		private static void CrcError()
		{
			throw new BZip2Exception("BZip2 input stream crc error");
		}

		private static void HbCreateDecodeTables(int[] limit, int[] baseArray, int[] perm, char[] length, int minLen, int maxLen, int alphaSize)
		{
			int pp = 0;

			for (int i = minLen; i <= maxLen; ++i)
			{
				for (int j = 0; j < alphaSize; ++j)
				{
					if (length[j] == i)
					{
						perm[pp] = j;
						++pp;
					}
				}
			}

			for (int i = 0; i < BZip2Constants.MaximumCodeLength; i++)
			{
				baseArray[i] = 0;
			}

			for (int i = 0; i < alphaSize; i++)
			{
				++baseArray[length[i] + 1];
			}

			for (int i = 1; i < BZip2Constants.MaximumCodeLength; i++)
			{
				baseArray[i] += baseArray[i - 1];
			}

			for (int i = 0; i < BZip2Constants.MaximumCodeLength; i++)
			{
				limit[i] = 0;
			}

			int vec = 0;

			for (int i = minLen; i <= maxLen; i++)
			{
				vec += (baseArray[i + 1] - baseArray[i]);
				limit[i] = vec - 1;
				vec <<= 1;
			}

			for (int i = minLen + 1; i <= maxLen; i++)
			{
				baseArray[i] = ((limit[i - 1] + 1) << 1) - baseArray[i];
			}
		}
	}
}


namespace ICSharpCode.SharpZipLib.BZip2
{
	/// <summary>
	/// An output stream that compresses into the BZip2 format
	/// including file header chars into another stream.
	/// </summary>
	public class BZip2OutputStream : Stream
	{
		#region Constants

		private const int SETMASK = (1 << 21);
		private const int CLEARMASK = (~SETMASK);
		private const int GREATER_ICOST = 15;
		private const int LESSER_ICOST = 0;
		private const int SMALL_THRESH = 20;
		private const int DEPTH_THRESH = 10;

		/*--
		If you are ever unlucky/improbable enough
		to get a stack overflow whilst sorting,
		increase the following constant and try
		again.  In practice I have never seen the
		stack go above 27 elems, so the following
		limit seems very generous.
		--*/
		private const int QSORT_STACK_SIZE = 1000;

		/*--
		Knuth's increments seem to work better
		than Incerpi-Sedgewick here.  Possibly
		because the number of elems to sort is
		usually small, typically <= 20.
		--*/

		private readonly int[] increments = {
												  1, 4, 13, 40, 121, 364, 1093, 3280,
												  9841, 29524, 88573, 265720,
												  797161, 2391484
											  };

		#endregion Constants

		#region Instance Fields

		/*--
		index of the last char in the block, so
		the block size == last + 1.
		--*/
		private int last;

		/*--
		index in zptr[] of original string after sorting.
		--*/
		private int origPtr;

		/*--
		always: in the range 0 .. 9.
		The current block size is 100000 * this number.
		--*/
		private int blockSize100k;

		private bool blockRandomised;

		private int bytesOut;
		private int bsBuff;
		private int bsLive;
		private IChecksum mCrc = new BZip2Crc();

		private bool[] inUse = new bool[256];
		private int nInUse;

		private char[] seqToUnseq = new char[256];
		private char[] unseqToSeq = new char[256];

		private char[] selector = new char[BZip2Constants.MaximumSelectors];
		private char[] selectorMtf = new char[BZip2Constants.MaximumSelectors];

		private byte[] block;
		private int[] quadrant;
		private int[] zptr;
		private short[] szptr;
		private int[] ftab;

		private int nMTF;

		private int[] mtfFreq = new int[BZip2Constants.MaximumAlphaSize];

		/*
		* Used when sorting.  If too many long comparisons
		* happen, we stop sorting, randomise the block
		* slightly, and try again.
		*/
		private int workFactor;
		private int workDone;
		private int workLimit;
		private bool firstAttempt;
		private int nBlocksRandomised;

		private int currentChar = -1;
		private int runLength;
		private uint blockCRC, combinedCRC;
		private int allowableBlockSize;
		private readonly Stream baseStream;
		private bool disposed_;

		#endregion Instance Fields

		/// <summary>
		/// Construct a default output stream with maximum block size
		/// </summary>
		/// <param name="stream">The stream to write BZip data onto.</param>
		public BZip2OutputStream(Stream stream) : this(stream, 9)
		{
		}

		/// <summary>
		/// Initialise a new instance of the <see cref="BZip2OutputStream"></see>
		/// for the specified stream, using the given blocksize.
		/// </summary>
		/// <param name="stream">The stream to write compressed data to.</param>
		/// <param name="blockSize">The block size to use.</param>
		/// <remarks>
		/// Valid block sizes are in the range 1..9, with 1 giving
		/// the lowest compression and 9 the highest.
		/// </remarks>
		public BZip2OutputStream(Stream stream, int blockSize)
		{
			if (stream == null)
				throw new ArgumentNullException(nameof(stream));

			baseStream = stream;
			bsLive = 0;
			bsBuff = 0;
			bytesOut = 0;

			workFactor = 50;
			if (blockSize > 9)
			{
				blockSize = 9;
			}

			if (blockSize < 1)
			{
				blockSize = 1;
			}
			blockSize100k = blockSize;
			AllocateCompressStructures();
			Initialize();
			InitBlock();
		}

		/// <summary>
		/// Ensures that resources are freed and other cleanup operations
		/// are performed when the garbage collector reclaims the BZip2OutputStream.
		/// </summary>
		~BZip2OutputStream()
		{
			Dispose(false);
		}

		/// <summary>
		/// Gets or sets a flag indicating ownership of underlying stream.
		/// When the flag is true <see cref="Stream.Dispose()" /> will close the underlying stream also.
		/// </summary>
		/// <remarks>The default value is true.</remarks>
		public bool IsStreamOwner { get; set; } = true;

		/// <summary>
		/// Gets a value indicating whether the current stream supports reading
		/// </summary>
		public override bool CanRead
		{
			get
			{
				return false;
			}
		}

		/// <summary>
		/// Gets a value indicating whether the current stream supports seeking
		/// </summary>
		public override bool CanSeek
		{
			get
			{
				return false;
			}
		}

		/// <summary>
		/// Gets a value indicating whether the current stream supports writing
		/// </summary>
		public override bool CanWrite
		{
			get
			{
				return baseStream.CanWrite;
			}
		}

		/// <summary>
		/// Gets the length in bytes of the stream
		/// </summary>
		public override long Length
		{
			get
			{
				return baseStream.Length;
			}
		}

		/// <summary>
		/// Gets or sets the current position of this stream.
		/// </summary>
		public override long Position
		{
			get
			{
				return baseStream.Position;
			}
			set
			{
				throw new NotSupportedException("BZip2OutputStream position cannot be set");
			}
		}

		/// <summary>
		/// Sets the current position of this stream to the given value.
		/// </summary>
		/// <param name="offset">The point relative to the offset from which to being seeking.</param>
		/// <param name="origin">The reference point from which to begin seeking.</param>
		/// <returns>The new position in the stream.</returns>
		public override long Seek(long offset, SeekOrigin origin)
		{
			throw new NotSupportedException("BZip2OutputStream Seek not supported");
		}

		/// <summary>
		/// Sets the length of this stream to the given value.
		/// </summary>
		/// <param name="value">The new stream length.</param>
		public override void SetLength(long value)
		{
			throw new NotSupportedException("BZip2OutputStream SetLength not supported");
		}

		/// <summary>
		/// Read a byte from the stream advancing the position.
		/// </summary>
		/// <returns>The byte read cast to an int; -1 if end of stream.</returns>
		public override int ReadByte()
		{
			throw new NotSupportedException("BZip2OutputStream ReadByte not supported");
		}

		/// <summary>
		/// Read a block of bytes
		/// </summary>
		/// <param name="buffer">The buffer to read into.</param>
		/// <param name="offset">The offset in the buffer to start storing data at.</param>
		/// <param name="count">The maximum number of bytes to read.</param>
		/// <returns>The total number of bytes read. This might be less than the number of bytes
		/// requested if that number of bytes are not currently available, or zero
		/// if the end of the stream is reached.</returns>
		public override int Read(byte[] buffer, int offset, int count)
		{
			throw new NotSupportedException("BZip2OutputStream Read not supported");
		}

		/// <summary>
		/// Write a block of bytes to the stream
		/// </summary>
		/// <param name="buffer">The buffer containing data to write.</param>
		/// <param name="offset">The offset of the first byte to write.</param>
		/// <param name="count">The number of bytes to write.</param>
		public override void Write(byte[] buffer, int offset, int count)
		{
			if (buffer == null)
			{
				throw new ArgumentNullException(nameof(buffer));
			}

			if (offset < 0)
			{
				throw new ArgumentOutOfRangeException(nameof(offset));
			}

			if (count < 0)
			{
				throw new ArgumentOutOfRangeException(nameof(count));
			}

			if (buffer.Length - offset < count)
			{
				throw new ArgumentException("Offset/count out of range");
			}

			for (int i = 0; i < count; ++i)
			{
				WriteByte(buffer[offset + i]);
			}
		}

		/// <summary>
		/// Write a byte to the stream.
		/// </summary>
		/// <param name="value">The byte to write to the stream.</param>
		public override void WriteByte(byte value)
		{
			int b = (256 + value) % 256;
			if (currentChar != -1)
			{
				if (currentChar == b)
				{
					runLength++;
					if (runLength > 254)
					{
						WriteRun();
						currentChar = -1;
						runLength = 0;
					}
				}
				else
				{
					WriteRun();
					runLength = 1;
					currentChar = b;
				}
			}
			else
			{
				currentChar = b;
				runLength++;
			}
		}

		private void MakeMaps()
		{
			nInUse = 0;
			for (int i = 0; i < 256; i++)
			{
				if (inUse[i])
				{
					seqToUnseq[nInUse] = (char)i;
					unseqToSeq[i] = (char)nInUse;
					nInUse++;
				}
			}
		}

		/// <summary>
		/// Get the number of bytes written to output.
		/// </summary>
		private void WriteRun()
		{
			if (last < allowableBlockSize)
			{
				inUse[currentChar] = true;
				for (int i = 0; i < runLength; i++)
				{
					mCrc.Update(currentChar);
				}

				switch (runLength)
				{
					case 1:
						last++;
						block[last + 1] = (byte)currentChar;
						break;

					case 2:
						last++;
						block[last + 1] = (byte)currentChar;
						last++;
						block[last + 1] = (byte)currentChar;
						break;

					case 3:
						last++;
						block[last + 1] = (byte)currentChar;
						last++;
						block[last + 1] = (byte)currentChar;
						last++;
						block[last + 1] = (byte)currentChar;
						break;

					default:
						inUse[runLength - 4] = true;
						last++;
						block[last + 1] = (byte)currentChar;
						last++;
						block[last + 1] = (byte)currentChar;
						last++;
						block[last + 1] = (byte)currentChar;
						last++;
						block[last + 1] = (byte)currentChar;
						last++;
						block[last + 1] = (byte)(runLength - 4);
						break;
				}
			}
			else
			{
				EndBlock();
				InitBlock();
				WriteRun();
			}
		}

		/// <summary>
		/// Get the number of bytes written to the output.
		/// </summary>
		public int BytesWritten
		{
			get { return bytesOut; }
		}

		/// <summary>
		/// Releases the unmanaged resources used by the <see cref="BZip2OutputStream"/> and optionally releases the managed resources.
		/// </summary>
		/// <param name="disposing">true to release both managed and unmanaged resources; false to release only unmanaged resources.</param>
		override protected void Dispose(bool disposing)
		{
			try
			{
				try
				{
					base.Dispose(disposing);
					if (!disposed_)
					{
						disposed_ = true;

						if (runLength > 0)
						{
							WriteRun();
						}

						currentChar = -1;
						EndBlock();
						EndCompression();
						Flush();
					}
				}
				finally
				{
					if (disposing)
					{
						if (IsStreamOwner)
						{
							baseStream.Dispose();
						}
					}
				}
			}
			catch
			{
			}
		}

		/// <summary>
		/// Flush output buffers
		/// </summary>
		public override void Flush()
		{
			baseStream.Flush();
		}

		private void Initialize()
		{
			bytesOut = 0;
			nBlocksRandomised = 0;

			/*--- Write header `magic' bytes indicating file-format == huffmanised,
			followed by a digit indicating blockSize100k.
			---*/

			BsPutUChar('B');
			BsPutUChar('Z');

			BsPutUChar('h');
			BsPutUChar('0' + blockSize100k);

			combinedCRC = 0;
		}

		private void InitBlock()
		{
			mCrc.Reset();
			last = -1;

			for (int i = 0; i < 256; i++)
			{
				inUse[i] = false;
			}

			/*--- 20 is just a paranoia constant ---*/
			allowableBlockSize = BZip2Constants.BaseBlockSize * blockSize100k - 20;
		}

		private void EndBlock()
		{
			if (last < 0)
			{       // dont do anything for empty files, (makes empty files compatible with original Bzip)
				return;
			}

			blockCRC = unchecked((uint)mCrc.Value);
			combinedCRC = (combinedCRC << 1) | (combinedCRC >> 31);
			combinedCRC ^= blockCRC;

			/*-- sort the block and establish position of original string --*/
			DoReversibleTransformation();

			/*--
			A 6-byte block header, the value chosen arbitrarily
			as 0x314159265359 :-).  A 32 bit value does not really
			give a strong enough guarantee that the value will not
			appear by chance in the compressed datastream.  Worst-case
			probability of this event, for a 900k block, is about
			2.0e-3 for 32 bits, 1.0e-5 for 40 bits and 4.0e-8 for 48 bits.
			For a compressed file of size 100Gb -- about 100000 blocks --
			only a 48-bit marker will do.  NB: normal compression/
			decompression do *not* rely on these statistical properties.
			They are only important when trying to recover blocks from
			damaged files.
			--*/
			BsPutUChar(0x31);
			BsPutUChar(0x41);
			BsPutUChar(0x59);
			BsPutUChar(0x26);
			BsPutUChar(0x53);
			BsPutUChar(0x59);

			/*-- Now the block's CRC, so it is in a known place. --*/
			unchecked
			{
				BsPutint((int)blockCRC);
			}

			/*-- Now a single bit indicating randomisation. --*/
			if (blockRandomised)
			{
				BsW(1, 1);
				nBlocksRandomised++;
			}
			else
			{
				BsW(1, 0);
			}

			/*-- Finally, block's contents proper. --*/
			MoveToFrontCodeAndSend();
		}

		private void EndCompression()
		{
			/*--
			Now another magic 48-bit number, 0x177245385090, to
			indicate the end of the last block.  (sqrt(pi), if
			you want to know.  I did want to use e, but it contains
			too much repetition -- 27 18 28 18 28 46 -- for me
			to feel statistically comfortable.  Call me paranoid.)
			--*/
			BsPutUChar(0x17);
			BsPutUChar(0x72);
			BsPutUChar(0x45);
			BsPutUChar(0x38);
			BsPutUChar(0x50);
			BsPutUChar(0x90);

			unchecked
			{
				BsPutint((int)combinedCRC);
			}

			BsFinishedWithStream();
		}

		private void BsFinishedWithStream()
		{
			while (bsLive > 0)
			{
				int ch = (bsBuff >> 24);
				baseStream.WriteByte((byte)ch); // write 8-bit
				bsBuff <<= 8;
				bsLive -= 8;
				bytesOut++;
			}
		}

		private void BsW(int n, int v)
		{
			while (bsLive >= 8)
			{
				int ch = (bsBuff >> 24);
				unchecked { baseStream.WriteByte((byte)ch); } // write 8-bit
				bsBuff <<= 8;
				bsLive -= 8;
				++bytesOut;
			}
			bsBuff |= (v << (32 - bsLive - n));
			bsLive += n;
		}

		private void BsPutUChar(int c)
		{
			BsW(8, c);
		}

		private void BsPutint(int u)
		{
			BsW(8, (u >> 24) & 0xFF);
			BsW(8, (u >> 16) & 0xFF);
			BsW(8, (u >> 8) & 0xFF);
			BsW(8, u & 0xFF);
		}

		private void BsPutIntVS(int numBits, int c)
		{
			BsW(numBits, c);
		}

		private void SendMTFValues()
		{
			char[][] len = new char[BZip2Constants.GroupCount][];
			for (int i = 0; i < BZip2Constants.GroupCount; ++i)
			{
				len[i] = new char[BZip2Constants.MaximumAlphaSize];
			}

			int gs, ge, totc, bt, bc, iter;
			int nSelectors = 0, alphaSize, minLen, maxLen, selCtr;
			int nGroups;

			alphaSize = nInUse + 2;
			for (int t = 0; t < BZip2Constants.GroupCount; t++)
			{
				for (int v = 0; v < alphaSize; v++)
				{
					len[t][v] = (char)GREATER_ICOST;
				}
			}

			/*--- Decide how many coding tables to use ---*/
			if (nMTF <= 0)
			{
				Panic();
			}

			if (nMTF < 200)
			{
				nGroups = 2;
			}
			else if (nMTF < 600)
			{
				nGroups = 3;
			}
			else if (nMTF < 1200)
			{
				nGroups = 4;
			}
			else if (nMTF < 2400)
			{
				nGroups = 5;
			}
			else
			{
				nGroups = 6;
			}

			/*--- Generate an initial set of coding tables ---*/
			int nPart = nGroups;
			int remF = nMTF;
			gs = 0;
			while (nPart > 0)
			{
				int tFreq = remF / nPart;
				int aFreq = 0;
				ge = gs - 1;
				while (aFreq < tFreq && ge < alphaSize - 1)
				{
					ge++;
					aFreq += mtfFreq[ge];
				}

				if (ge > gs && nPart != nGroups && nPart != 1 && ((nGroups - nPart) % 2 == 1))
				{
					aFreq -= mtfFreq[ge];
					ge--;
				}

				for (int v = 0; v < alphaSize; v++)
				{
					if (v >= gs && v <= ge)
					{
						len[nPart - 1][v] = (char)LESSER_ICOST;
					}
					else
					{
						len[nPart - 1][v] = (char)GREATER_ICOST;
					}
				}

				nPart--;
				gs = ge + 1;
				remF -= aFreq;
			}

			int[][] rfreq = new int[BZip2Constants.GroupCount][];
			for (int i = 0; i < BZip2Constants.GroupCount; ++i)
			{
				rfreq[i] = new int[BZip2Constants.MaximumAlphaSize];
			}

			int[] fave = new int[BZip2Constants.GroupCount];
			short[] cost = new short[BZip2Constants.GroupCount];
			/*---
			Iterate up to N_ITERS times to improve the tables.
			---*/
			for (iter = 0; iter < BZip2Constants.NumberOfIterations; ++iter)
			{
				for (int t = 0; t < nGroups; ++t)
				{
					fave[t] = 0;
				}

				for (int t = 0; t < nGroups; ++t)
				{
					for (int v = 0; v < alphaSize; ++v)
					{
						rfreq[t][v] = 0;
					}
				}

				nSelectors = 0;
				totc = 0;
				gs = 0;
				while (true)
				{
					/*--- Set group start & end marks. --*/
					if (gs >= nMTF)
					{
						break;
					}
					ge = gs + BZip2Constants.GroupSize - 1;
					if (ge >= nMTF)
					{
						ge = nMTF - 1;
					}

					/*--
					Calculate the cost of this group as coded
					by each of the coding tables.
					--*/
					for (int t = 0; t < nGroups; t++)
					{
						cost[t] = 0;
					}

					if (nGroups == 6)
					{
						short cost0, cost1, cost2, cost3, cost4, cost5;
						cost0 = cost1 = cost2 = cost3 = cost4 = cost5 = 0;
						for (int i = gs; i <= ge; ++i)
						{
							short icv = szptr[i];
							cost0 += (short)len[0][icv];
							cost1 += (short)len[1][icv];
							cost2 += (short)len[2][icv];
							cost3 += (short)len[3][icv];
							cost4 += (short)len[4][icv];
							cost5 += (short)len[5][icv];
						}
						cost[0] = cost0;
						cost[1] = cost1;
						cost[2] = cost2;
						cost[3] = cost3;
						cost[4] = cost4;
						cost[5] = cost5;
					}
					else
					{
						for (int i = gs; i <= ge; ++i)
						{
							short icv = szptr[i];
							for (int t = 0; t < nGroups; t++)
							{
								cost[t] += (short)len[t][icv];
							}
						}
					}

					/*--
					Find the coding table which is best for this group,
					and record its identity in the selector table.
					--*/
					bc = 999999999;
					bt = -1;
					for (int t = 0; t < nGroups; ++t)
					{
						if (cost[t] < bc)
						{
							bc = cost[t];
							bt = t;
						}
					}
					totc += bc;
					fave[bt]++;
					selector[nSelectors] = (char)bt;
					nSelectors++;

					/*--
					Increment the symbol frequencies for the selected table.
					--*/
					for (int i = gs; i <= ge; ++i)
					{
						++rfreq[bt][szptr[i]];
					}

					gs = ge + 1;
				}

				/*--
				Recompute the tables based on the accumulated frequencies.
				--*/
				for (int t = 0; t < nGroups; ++t)
				{
					HbMakeCodeLengths(len[t], rfreq[t], alphaSize, 20);
				}
			}

			rfreq = null;
			fave = null;
			cost = null;

			if (!(nGroups < 8))
			{
				Panic();
			}

			if (!(nSelectors < 32768 && nSelectors <= (2 + (900000 / BZip2Constants.GroupSize))))
			{
				Panic();
			}

			/*--- Compute MTF values for the selectors. ---*/
			char[] pos = new char[BZip2Constants.GroupCount];
			char ll_i, tmp2, tmp;

			for (int i = 0; i < nGroups; i++)
			{
				pos[i] = (char)i;
			}

			for (int i = 0; i < nSelectors; i++)
			{
				ll_i = selector[i];
				int j = 0;
				tmp = pos[j];
				while (ll_i != tmp)
				{
					j++;
					tmp2 = tmp;
					tmp = pos[j];
					pos[j] = tmp2;
				}
				pos[0] = tmp;
				selectorMtf[i] = (char)j;
			}

			int[][] code = new int[BZip2Constants.GroupCount][];

			for (int i = 0; i < BZip2Constants.GroupCount; ++i)
			{
				code[i] = new int[BZip2Constants.MaximumAlphaSize];
			}

			/*--- Assign actual codes for the tables. --*/
			for (int t = 0; t < nGroups; t++)
			{
				minLen = 32;
				maxLen = 0;
				for (int i = 0; i < alphaSize; i++)
				{
					if (len[t][i] > maxLen)
					{
						maxLen = len[t][i];
					}
					if (len[t][i] < minLen)
					{
						minLen = len[t][i];
					}
				}
				if (maxLen > 20)
				{
					Panic();
				}
				if (minLen < 1)
				{
					Panic();
				}
				HbAssignCodes(code[t], len[t], minLen, maxLen, alphaSize);
			}

			/*--- Transmit the mapping table. ---*/
			bool[] inUse16 = new bool[16];
			for (int i = 0; i < 16; ++i)
			{
				inUse16[i] = false;
				for (int j = 0; j < 16; ++j)
				{
					if (inUse[i * 16 + j])
					{
						inUse16[i] = true;
					}
				}
			}

			for (int i = 0; i < 16; ++i)
			{
				if (inUse16[i])
				{
					BsW(1, 1);
				}
				else
				{
					BsW(1, 0);
				}
			}

			for (int i = 0; i < 16; ++i)
			{
				if (inUse16[i])
				{
					for (int j = 0; j < 16; ++j)
					{
						if (inUse[i * 16 + j])
						{
							BsW(1, 1);
						}
						else
						{
							BsW(1, 0);
						}
					}
				}
			}

			/*--- Now the selectors. ---*/
			BsW(3, nGroups);
			BsW(15, nSelectors);
			for (int i = 0; i < nSelectors; ++i)
			{
				for (int j = 0; j < selectorMtf[i]; ++j)
				{
					BsW(1, 1);
				}
				BsW(1, 0);
			}

			/*--- Now the coding tables. ---*/
			for (int t = 0; t < nGroups; ++t)
			{
				int curr = len[t][0];
				BsW(5, curr);
				for (int i = 0; i < alphaSize; ++i)
				{
					while (curr < len[t][i])
					{
						BsW(2, 2);
						curr++; /* 10 */
					}
					while (curr > len[t][i])
					{
						BsW(2, 3);
						curr--; /* 11 */
					}
					BsW(1, 0);
				}
			}

			/*--- And finally, the block data proper ---*/
			selCtr = 0;
			gs = 0;
			while (true)
			{
				if (gs >= nMTF)
				{
					break;
				}
				ge = gs + BZip2Constants.GroupSize - 1;
				if (ge >= nMTF)
				{
					ge = nMTF - 1;
				}

				for (int i = gs; i <= ge; i++)
				{
					BsW(len[selector[selCtr]][szptr[i]], code[selector[selCtr]][szptr[i]]);
				}

				gs = ge + 1;
				++selCtr;
			}
			if (!(selCtr == nSelectors))
			{
				Panic();
			}
		}

		private void MoveToFrontCodeAndSend()
		{
			BsPutIntVS(24, origPtr);
			GenerateMTFValues();
			SendMTFValues();
		}

		private void SimpleSort(int lo, int hi, int d)
		{
			int i, j, h, bigN, hp;
			int v;

			bigN = hi - lo + 1;
			if (bigN < 2)
			{
				return;
			}

			hp = 0;
			while (increments[hp] < bigN)
			{
				hp++;
			}
			hp--;

			for (; hp >= 0; hp--)
			{
				h = increments[hp];

				i = lo + h;
				while (true)
				{
					/*-- copy 1 --*/
					if (i > hi)
						break;
					v = zptr[i];
					j = i;
					while (FullGtU(zptr[j - h] + d, v + d))
					{
						zptr[j] = zptr[j - h];
						j = j - h;
						if (j <= (lo + h - 1))
							break;
					}
					zptr[j] = v;
					i++;

					/*-- copy 2 --*/
					if (i > hi)
					{
						break;
					}
					v = zptr[i];
					j = i;
					while (FullGtU(zptr[j - h] + d, v + d))
					{
						zptr[j] = zptr[j - h];
						j = j - h;
						if (j <= (lo + h - 1))
						{
							break;
						}
					}
					zptr[j] = v;
					i++;

					/*-- copy 3 --*/
					if (i > hi)
					{
						break;
					}
					v = zptr[i];
					j = i;
					while (FullGtU(zptr[j - h] + d, v + d))
					{
						zptr[j] = zptr[j - h];
						j = j - h;
						if (j <= (lo + h - 1))
						{
							break;
						}
					}
					zptr[j] = v;
					i++;

					if (workDone > workLimit && firstAttempt)
					{
						return;
					}
				}
			}
		}

		private void Vswap(int p1, int p2, int n)
		{
			int temp = 0;
			while (n > 0)
			{
				temp = zptr[p1];
				zptr[p1] = zptr[p2];
				zptr[p2] = temp;
				p1++;
				p2++;
				n--;
			}
		}

		private void QSort3(int loSt, int hiSt, int dSt)
		{
			int unLo, unHi, ltLo, gtHi, med, n, m;
			int lo, hi, d;

			StackElement[] stack = new StackElement[QSORT_STACK_SIZE];

			int sp = 0;

			stack[sp].ll = loSt;
			stack[sp].hh = hiSt;
			stack[sp].dd = dSt;
			sp++;

			while (sp > 0)
			{
				if (sp >= QSORT_STACK_SIZE)
				{
					Panic();
				}

				sp--;
				lo = stack[sp].ll;
				hi = stack[sp].hh;
				d = stack[sp].dd;

				if (hi - lo < SMALL_THRESH || d > DEPTH_THRESH)
				{
					SimpleSort(lo, hi, d);
					if (workDone > workLimit && firstAttempt)
					{
						return;
					}
					continue;
				}

				med = Med3(block[zptr[lo] + d + 1],
						   block[zptr[hi] + d + 1],
						   block[zptr[(lo + hi) >> 1] + d + 1]);

				unLo = ltLo = lo;
				unHi = gtHi = hi;

				while (true)
				{
					while (true)
					{
						if (unLo > unHi)
						{
							break;
						}
						n = ((int)block[zptr[unLo] + d + 1]) - med;
						if (n == 0)
						{
							int temp = zptr[unLo];
							zptr[unLo] = zptr[ltLo];
							zptr[ltLo] = temp;
							ltLo++;
							unLo++;
							continue;
						}
						if (n > 0)
						{
							break;
						}
						unLo++;
					}

					while (true)
					{
						if (unLo > unHi)
						{
							break;
						}
						n = ((int)block[zptr[unHi] + d + 1]) - med;
						if (n == 0)
						{
							int temp = zptr[unHi];
							zptr[unHi] = zptr[gtHi];
							zptr[gtHi] = temp;
							gtHi--;
							unHi--;
							continue;
						}
						if (n < 0)
						{
							break;
						}
						unHi--;
					}

					if (unLo > unHi)
					{
						break;
					}

					{
						int temp = zptr[unLo];
						zptr[unLo] = zptr[unHi];
						zptr[unHi] = temp;
						unLo++;
						unHi--;
					}
				}

				if (gtHi < ltLo)
				{
					stack[sp].ll = lo;
					stack[sp].hh = hi;
					stack[sp].dd = d + 1;
					sp++;
					continue;
				}

				n = ((ltLo - lo) < (unLo - ltLo)) ? (ltLo - lo) : (unLo - ltLo);
				Vswap(lo, unLo - n, n);
				m = ((hi - gtHi) < (gtHi - unHi)) ? (hi - gtHi) : (gtHi - unHi);
				Vswap(unLo, hi - m + 1, m);

				n = lo + unLo - ltLo - 1;
				m = hi - (gtHi - unHi) + 1;

				stack[sp].ll = lo;
				stack[sp].hh = n;
				stack[sp].dd = d;
				sp++;

				stack[sp].ll = n + 1;
				stack[sp].hh = m - 1;
				stack[sp].dd = d + 1;
				sp++;

				stack[sp].ll = m;
				stack[sp].hh = hi;
				stack[sp].dd = d;
				sp++;
			}
		}

		private void MainSort()
		{
			int i, j, ss, sb;
			int[] runningOrder = new int[256];
			int[] copy = new int[256];
			bool[] bigDone = new bool[256];
			int c1, c2;
			int numQSorted;

			/*--
			In the various block-sized structures, live data runs
			from 0 to last+NUM_OVERSHOOT_BYTES inclusive.  First,
			set up the overshoot area for block.
			--*/

			//   if (verbosity >= 4) fprintf ( stderr, "        sort initialise ...\n" );
			for (i = 0; i < BZip2Constants.OvershootBytes; i++)
			{
				block[last + i + 2] = block[(i % (last + 1)) + 1];
			}
			for (i = 0; i <= last + BZip2Constants.OvershootBytes; i++)
			{
				quadrant[i] = 0;
			}

			block[0] = (byte)(block[last + 1]);

			if (last < 4000)
			{
				/*--
				Use simpleSort(), since the full sorting mechanism
				has quite a large constant overhead.
				--*/
				for (i = 0; i <= last; i++)
				{
					zptr[i] = i;
				}
				firstAttempt = false;
				workDone = workLimit = 0;
				SimpleSort(0, last, 0);
			}
			else
			{
				numQSorted = 0;
				for (i = 0; i <= 255; i++)
				{
					bigDone[i] = false;
				}
				for (i = 0; i <= 65536; i++)
				{
					ftab[i] = 0;
				}

				c1 = block[0];
				for (i = 0; i <= last; i++)
				{
					c2 = block[i + 1];
					ftab[(c1 << 8) + c2]++;
					c1 = c2;
				}

				for (i = 1; i <= 65536; i++)
				{
					ftab[i] += ftab[i - 1];
				}

				c1 = block[1];
				for (i = 0; i < last; i++)
				{
					c2 = block[i + 2];
					j = (c1 << 8) + c2;
					c1 = c2;
					ftab[j]--;
					zptr[ftab[j]] = i;
				}

				j = ((block[last + 1]) << 8) + (block[1]);
				ftab[j]--;
				zptr[ftab[j]] = last;

				/*--
				Now ftab contains the first loc of every small bucket.
				Calculate the running order, from smallest to largest
				big bucket.
				--*/

				for (i = 0; i <= 255; i++)
				{
					runningOrder[i] = i;
				}

				int vv;
				int h = 1;
				do
				{
					h = 3 * h + 1;
				} while (h <= 256);
				do
				{
					h = h / 3;
					for (i = h; i <= 255; i++)
					{
						vv = runningOrder[i];
						j = i;
						while ((ftab[((runningOrder[j - h]) + 1) << 8] - ftab[(runningOrder[j - h]) << 8]) > (ftab[((vv) + 1) << 8] - ftab[(vv) << 8]))
						{
							runningOrder[j] = runningOrder[j - h];
							j = j - h;
							if (j <= (h - 1))
							{
								break;
							}
						}
						runningOrder[j] = vv;
					}
				} while (h != 1);

				/*--
				The main sorting loop.
				--*/
				for (i = 0; i <= 255; i++)
				{
					/*--
					Process big buckets, starting with the least full.
					--*/
					ss = runningOrder[i];

					/*--
					Complete the big bucket [ss] by quicksorting
					any unsorted small buckets [ss, j].  Hopefully
					previous pointer-scanning phases have already
					completed many of the small buckets [ss, j], so
					we don't have to sort them at all.
					--*/
					for (j = 0; j <= 255; j++)
					{
						sb = (ss << 8) + j;
						if (!((ftab[sb] & SETMASK) == SETMASK))
						{
							int lo = ftab[sb] & CLEARMASK;
							int hi = (ftab[sb + 1] & CLEARMASK) - 1;
							if (hi > lo)
							{
								QSort3(lo, hi, 2);
								numQSorted += (hi - lo + 1);
								if (workDone > workLimit && firstAttempt)
								{
									return;
								}
							}
							ftab[sb] |= SETMASK;
						}
					}

					/*--
					The ss big bucket is now done.  Record this fact,
					and update the quadrant descriptors.  Remember to
					update quadrants in the overshoot area too, if
					necessary.  The "if (i < 255)" test merely skips
					this updating for the last bucket processed, since
					updating for the last bucket is pointless.
					--*/
					bigDone[ss] = true;

					if (i < 255)
					{
						int bbStart = ftab[ss << 8] & CLEARMASK;
						int bbSize = (ftab[(ss + 1) << 8] & CLEARMASK) - bbStart;
						int shifts = 0;

						while ((bbSize >> shifts) > 65534)
						{
							shifts++;
						}

						for (j = 0; j < bbSize; j++)
						{
							int a2update = zptr[bbStart + j];
							int qVal = (j >> shifts);
							quadrant[a2update] = qVal;
							if (a2update < BZip2Constants.OvershootBytes)
							{
								quadrant[a2update + last + 1] = qVal;
							}
						}

						if (!(((bbSize - 1) >> shifts) <= 65535))
						{
							Panic();
						}
					}

					/*--
					Now scan this big bucket so as to synthesise the
					sorted order for small buckets [t, ss] for all t != ss.
					--*/
					for (j = 0; j <= 255; j++)
					{
						copy[j] = ftab[(j << 8) + ss] & CLEARMASK;
					}

					for (j = ftab[ss << 8] & CLEARMASK; j < (ftab[(ss + 1) << 8] & CLEARMASK); j++)
					{
						c1 = block[zptr[j]];
						if (!bigDone[c1])
						{
							zptr[copy[c1]] = zptr[j] == 0 ? last : zptr[j] - 1;
							copy[c1]++;
						}
					}

					for (j = 0; j <= 255; j++)
					{
						ftab[(j << 8) + ss] |= SETMASK;
					}
				}
			}
		}

		private void RandomiseBlock()
		{
			int i;
			int rNToGo = 0;
			int rTPos = 0;
			for (i = 0; i < 256; i++)
			{
				inUse[i] = false;
			}

			for (i = 0; i <= last; i++)
			{
				if (rNToGo == 0)
				{
					rNToGo = (int)BZip2Constants.RandomNumbers[rTPos];
					rTPos++;
					if (rTPos == 512)
					{
						rTPos = 0;
					}
				}
				rNToGo--;
				block[i + 1] ^= (byte)((rNToGo == 1) ? 1 : 0);
				// handle 16 bit signed numbers
				block[i + 1] &= 0xFF;

				inUse[block[i + 1]] = true;
			}
		}

		private void DoReversibleTransformation()
		{
			workLimit = workFactor * last;
			workDone = 0;
			blockRandomised = false;
			firstAttempt = true;

			MainSort();

			if (workDone > workLimit && firstAttempt)
			{
				RandomiseBlock();
				workLimit = workDone = 0;
				blockRandomised = true;
				firstAttempt = false;
				MainSort();
			}

			origPtr = -1;
			for (int i = 0; i <= last; i++)
			{
				if (zptr[i] == 0)
				{
					origPtr = i;
					break;
				}
			}

			if (origPtr == -1)
			{
				Panic();
			}
		}

		private bool FullGtU(int i1, int i2)
		{
			int k;
			byte c1, c2;
			int s1, s2;

			c1 = block[i1 + 1];
			c2 = block[i2 + 1];
			if (c1 != c2)
			{
				return c1 > c2;
			}
			i1++;
			i2++;

			c1 = block[i1 + 1];
			c2 = block[i2 + 1];
			if (c1 != c2)
			{
				return c1 > c2;
			}
			i1++;
			i2++;

			c1 = block[i1 + 1];
			c2 = block[i2 + 1];
			if (c1 != c2)
			{
				return c1 > c2;
			}
			i1++;
			i2++;

			c1 = block[i1 + 1];
			c2 = block[i2 + 1];
			if (c1 != c2)
			{
				return c1 > c2;
			}
			i1++;
			i2++;

			c1 = block[i1 + 1];
			c2 = block[i2 + 1];
			if (c1 != c2)
			{
				return c1 > c2;
			}
			i1++;
			i2++;

			c1 = block[i1 + 1];
			c2 = block[i2 + 1];
			if (c1 != c2)
			{
				return c1 > c2;
			}
			i1++;
			i2++;

			k = last + 1;

			do
			{
				c1 = block[i1 + 1];
				c2 = block[i2 + 1];
				if (c1 != c2)
				{
					return c1 > c2;
				}
				s1 = quadrant[i1];
				s2 = quadrant[i2];
				if (s1 != s2)
				{
					return s1 > s2;
				}
				i1++;
				i2++;

				c1 = block[i1 + 1];
				c2 = block[i2 + 1];
				if (c1 != c2)
				{
					return c1 > c2;
				}
				s1 = quadrant[i1];
				s2 = quadrant[i2];
				if (s1 != s2)
				{
					return s1 > s2;
				}
				i1++;
				i2++;

				c1 = block[i1 + 1];
				c2 = block[i2 + 1];
				if (c1 != c2)
				{
					return c1 > c2;
				}
				s1 = quadrant[i1];
				s2 = quadrant[i2];
				if (s1 != s2)
				{
					return s1 > s2;
				}
				i1++;
				i2++;

				c1 = block[i1 + 1];
				c2 = block[i2 + 1];
				if (c1 != c2)
				{
					return c1 > c2;
				}
				s1 = quadrant[i1];
				s2 = quadrant[i2];
				if (s1 != s2)
				{
					return s1 > s2;
				}
				i1++;
				i2++;

				if (i1 > last)
				{
					i1 -= last;
					i1--;
				}
				if (i2 > last)
				{
					i2 -= last;
					i2--;
				}

				k -= 4;
				++workDone;
			} while (k >= 0);

			return false;
		}

		private void AllocateCompressStructures()
		{
			int n = BZip2Constants.BaseBlockSize * blockSize100k;
			block = new byte[(n + 1 + BZip2Constants.OvershootBytes)];
			quadrant = new int[(n + BZip2Constants.OvershootBytes)];
			zptr = new int[n];
			ftab = new int[65537];

			if (block == null || quadrant == null || zptr == null || ftab == null)
			{
				//		int totalDraw = (n + 1 + NUM_OVERSHOOT_BYTES) + (n + NUM_OVERSHOOT_BYTES) + n + 65537;
				//		compressOutOfMemory ( totalDraw, n );
			}

			/*
			The back end needs a place to store the MTF values
			whilst it calculates the coding tables.  We could
			put them in the zptr array.  However, these values
			will fit in a short, so we overlay szptr at the
			start of zptr, in the hope of reducing the number
			of cache misses induced by the multiple traversals
			of the MTF values when calculating coding tables.
			Seems to improve compression speed by about 1%.
			*/
			//	szptr = zptr;

			szptr = new short[2 * n];
		}

		private void GenerateMTFValues()
		{
			char[] yy = new char[256];
			int i, j;
			char tmp;
			char tmp2;
			int zPend;
			int wr;
			int EOB;

			MakeMaps();
			EOB = nInUse + 1;

			for (i = 0; i <= EOB; i++)
			{
				mtfFreq[i] = 0;
			}

			wr = 0;
			zPend = 0;
			for (i = 0; i < nInUse; i++)
			{
				yy[i] = (char)i;
			}

			for (i = 0; i <= last; i++)
			{
				char ll_i;

				ll_i = unseqToSeq[block[zptr[i]]];

				j = 0;
				tmp = yy[j];
				while (ll_i != tmp)
				{
					j++;
					tmp2 = tmp;
					tmp = yy[j];
					yy[j] = tmp2;
				}
				yy[0] = tmp;

				if (j == 0)
				{
					zPend++;
				}
				else
				{
					if (zPend > 0)
					{
						zPend--;
						while (true)
						{
							switch (zPend % 2)
							{
								case 0:
									szptr[wr] = (short)BZip2Constants.RunA;
									wr++;
									mtfFreq[BZip2Constants.RunA]++;
									break;

								case 1:
									szptr[wr] = (short)BZip2Constants.RunB;
									wr++;
									mtfFreq[BZip2Constants.RunB]++;
									break;
							}
							if (zPend < 2)
							{
								break;
							}
							zPend = (zPend - 2) / 2;
						}
						zPend = 0;
					}
					szptr[wr] = (short)(j + 1);
					wr++;
					mtfFreq[j + 1]++;
				}
			}

			if (zPend > 0)
			{
				zPend--;
				while (true)
				{
					switch (zPend % 2)
					{
						case 0:
							szptr[wr] = (short)BZip2Constants.RunA;
							wr++;
							mtfFreq[BZip2Constants.RunA]++;
							break;

						case 1:
							szptr[wr] = (short)BZip2Constants.RunB;
							wr++;
							mtfFreq[BZip2Constants.RunB]++;
							break;
					}
					if (zPend < 2)
					{
						break;
					}
					zPend = (zPend - 2) / 2;
				}
			}

			szptr[wr] = (short)EOB;
			wr++;
			mtfFreq[EOB]++;

			nMTF = wr;
		}

		private static void Panic()
		{
			throw new BZip2Exception("BZip2 output stream panic");
		}

		private static void HbMakeCodeLengths(char[] len, int[] freq, int alphaSize, int maxLen)
		{
			/*--
			Nodes and heap entries run from 1.  Entry 0
			for both the heap and nodes is a sentinel.
			--*/
			int nNodes, nHeap, n1, n2, j, k;
			bool tooLong;

			int[] heap = new int[BZip2Constants.MaximumAlphaSize + 2];
			int[] weight = new int[BZip2Constants.MaximumAlphaSize * 2];
			int[] parent = new int[BZip2Constants.MaximumAlphaSize * 2];

			for (int i = 0; i < alphaSize; ++i)
			{
				weight[i + 1] = (freq[i] == 0 ? 1 : freq[i]) << 8;
			}

			while (true)
			{
				nNodes = alphaSize;
				nHeap = 0;

				heap[0] = 0;
				weight[0] = 0;
				parent[0] = -2;

				for (int i = 1; i <= alphaSize; ++i)
				{
					parent[i] = -1;
					nHeap++;
					heap[nHeap] = i;
					int zz = nHeap;
					int tmp = heap[zz];
					while (weight[tmp] < weight[heap[zz >> 1]])
					{
						heap[zz] = heap[zz >> 1];
						zz >>= 1;
					}
					heap[zz] = tmp;
				}
				if (!(nHeap < (BZip2Constants.MaximumAlphaSize + 2)))
				{
					Panic();
				}

				while (nHeap > 1)
				{
					n1 = heap[1];
					heap[1] = heap[nHeap];
					nHeap--;
					int zz = 1;
					int yy = 0;
					int tmp = heap[zz];
					while (true)
					{
						yy = zz << 1;
						if (yy > nHeap)
						{
							break;
						}
						if (yy < nHeap && weight[heap[yy + 1]] < weight[heap[yy]])
						{
							yy++;
						}
						if (weight[tmp] < weight[heap[yy]])
						{
							break;
						}

						heap[zz] = heap[yy];
						zz = yy;
					}
					heap[zz] = tmp;
					n2 = heap[1];
					heap[1] = heap[nHeap];
					nHeap--;

					zz = 1;
					yy = 0;
					tmp = heap[zz];
					while (true)
					{
						yy = zz << 1;
						if (yy > nHeap)
						{
							break;
						}
						if (yy < nHeap && weight[heap[yy + 1]] < weight[heap[yy]])
						{
							yy++;
						}
						if (weight[tmp] < weight[heap[yy]])
						{
							break;
						}
						heap[zz] = heap[yy];
						zz = yy;
					}
					heap[zz] = tmp;
					nNodes++;
					parent[n1] = parent[n2] = nNodes;

					weight[nNodes] = (int)((weight[n1] & 0xffffff00) + (weight[n2] & 0xffffff00)) |
						(int)(1 + (((weight[n1] & 0x000000ff) > (weight[n2] & 0x000000ff)) ? (weight[n1] & 0x000000ff) : (weight[n2] & 0x000000ff)));

					parent[nNodes] = -1;
					nHeap++;
					heap[nHeap] = nNodes;

					zz = nHeap;
					tmp = heap[zz];
					while (weight[tmp] < weight[heap[zz >> 1]])
					{
						heap[zz] = heap[zz >> 1];
						zz >>= 1;
					}
					heap[zz] = tmp;
				}
				if (!(nNodes < (BZip2Constants.MaximumAlphaSize * 2)))
				{
					Panic();
				}

				tooLong = false;
				for (int i = 1; i <= alphaSize; ++i)
				{
					j = 0;
					k = i;
					while (parent[k] >= 0)
					{
						k = parent[k];
						j++;
					}
					len[i - 1] = (char)j;
					tooLong |= j > maxLen;
				}

				if (!tooLong)
				{
					break;
				}

				for (int i = 1; i < alphaSize; ++i)
				{
					j = weight[i] >> 8;
					j = 1 + (j / 2);
					weight[i] = j << 8;
				}
			}
		}

		private static void HbAssignCodes(int[] code, char[] length, int minLen, int maxLen, int alphaSize)
		{
			int vec = 0;
			for (int n = minLen; n <= maxLen; ++n)
			{
				for (int i = 0; i < alphaSize; ++i)
				{
					if (length[i] == n)
					{
						code[i] = vec;
						++vec;
					}
				}
				vec <<= 1;
			}
		}

		private static byte Med3(byte a, byte b, byte c)
		{
			byte t;
			if (a > b)
			{
				t = a;
				a = b;
				b = t;
			}
			if (b > c)
			{
				t = b;
				b = c;
				c = t;
			}
			if (a > b)
			{
				b = a;
			}
			return b;
		}

		private struct StackElement
		{
			public int ll;
			public int hh;
			public int dd;
		}
	}
}


namespace ICSharpCode.SharpZipLib.Checksum
{
	/// <summary>
	/// Computes Adler32 checksum for a stream of data. An Adler32
	/// checksum is not as reliable as a CRC32 checksum, but a lot faster to
	/// compute.
	///
	/// The specification for Adler32 may be found in RFC 1950.
	/// ZLIB Compressed Data Format Specification version 3.3)
	///
	///
	/// From that document:
	///
	///      "ADLER32 (Adler-32 checksum)
	///       This contains a checksum value of the uncompressed data
	///       (excluding any dictionary data) computed according to Adler-32
	///       algorithm. This algorithm is a 32-bit extension and improvement
	///       of the Fletcher algorithm, used in the ITU-T X.224 / ISO 8073
	///       standard.
	///
	///       Adler-32 is composed of two sums accumulated per byte: s1 is
	///       the sum of all bytes, s2 is the sum of all s1 values. Both sums
	///       are done modulo 65521. s1 is initialized to 1, s2 to zero.  The
	///       Adler-32 checksum is stored as s2*65536 + s1 in most-
	///       significant-byte first (network) order."
	///
	///  "8.2. The Adler-32 algorithm
	///
	///    The Adler-32 algorithm is much faster than the CRC32 algorithm yet
	///    still provides an extremely low probability of undetected errors.
	///
	///    The modulo on unsigned long accumulators can be delayed for 5552
	///    bytes, so the modulo operation time is negligible.  If the bytes
	///    are a, b, c, the second sum is 3a + 2b + c + 3, and so is position
	///    and order sensitive, unlike the first sum, which is just a
	///    checksum.  That 65521 is prime is important to avoid a possible
	///    large class of two-byte errors that leave the check unchanged.
	///    (The Fletcher checksum uses 255, which is not prime and which also
	///    makes the Fletcher check insensitive to single byte changes 0 -
	///    255.)
	///
	///    The sum s1 is initialized to 1 instead of zero to make the length
	///    of the sequence part of s2, so that the length does not have to be
	///    checked separately. (Any sequence of zeroes has a Fletcher
	///    checksum of zero.)"
	/// </summary>
	/// <see cref="ICSharpCode.SharpZipLib.Zip.Compression.Streams.InflaterInputStream"/>
	/// <see cref="ICSharpCode.SharpZipLib.Zip.Compression.Streams.DeflaterOutputStream"/>
	public sealed class Adler32 : IChecksum
	{
		#region Instance Fields

		/// <summary>
		/// largest prime smaller than 65536
		/// </summary>
		private static readonly uint BASE = 65521;

		/// <summary>
		/// The CRC data checksum so far.
		/// </summary>
		private uint checkValue;

		#endregion Instance Fields

		/// <summary>
		/// Initialise a default instance of <see cref="Adler32"></see>
		/// </summary>
		public Adler32()
		{
			Reset();
		}

		/// <summary>
		/// Resets the Adler32 data checksum as if no update was ever called.
		/// </summary>
		public void Reset()
		{
			checkValue = 1;
		}

		/// <summary>
		/// Returns the Adler32 data checksum computed so far.
		/// </summary>
		public long Value
		{
			get
			{
				return checkValue;
			}
		}

		/// <summary>
		/// Updates the checksum with the byte b.
		/// </summary>
		/// <param name="bval">
		/// The data value to add. The high byte of the int is ignored.
		/// </param>
		public void Update(int bval)
		{
			// We could make a length 1 byte array and call update again, but I
			// would rather not have that overhead
			uint s1 = checkValue & 0xFFFF;
			uint s2 = checkValue >> 16;

			s1 = (s1 + ((uint)bval & 0xFF)) % BASE;
			s2 = (s1 + s2) % BASE;

			checkValue = (s2 << 16) + s1;
		}

		/// <summary>
		/// Updates the Adler32 data checksum with the bytes taken from
		/// a block of data.
		/// </summary>
		/// <param name="buffer">Contains the data to update the checksum with.</param>
		public void Update(byte[] buffer)
		{
			if (buffer == null)
			{
				throw new ArgumentNullException(nameof(buffer));
			}

			Update(new ArraySegment<byte>(buffer, 0, buffer.Length));
		}

		/// <summary>
		/// Update Adler32 data checksum based on a portion of a block of data
		/// </summary>
		/// <param name = "segment">
		/// The chunk of data to add
		/// </param>
		public void Update(ArraySegment<byte> segment)
		{
			//(By Per Bothner)
			uint s1 = checkValue & 0xFFFF;
			uint s2 = checkValue >> 16;
			var count = segment.Count;
			var offset = segment.Offset;
			while (count > 0)
			{
				// We can defer the modulo operation:
				// s1 maximally grows from 65521 to 65521 + 255 * 3800
				// s2 maximally grows by 3800 * median(s1) = 2090079800 < 2^31
				int n = 3800;
				if (n > count)
				{
					n = count;
				}
				count -= n;
				while (--n >= 0)
				{
					s1 = s1 + (uint)(segment.Array[offset++] & 0xff);
					s2 = s2 + s1;
				}
				s1 %= BASE;
				s2 %= BASE;
			}
			checkValue = (s2 << 16) | s1;
		}
	}
}


namespace ICSharpCode.SharpZipLib.Checksum
{
	/// <summary>
	/// CRC-32 with unreversed data and reversed output
	/// </summary>
	/// <remarks>
	/// Generate a table for a byte-wise 32-bit CRC calculation on the polynomial:
	/// x^32+x^26+x^23+x^22+x^16+x^12+x^11+x^10+x^8+x^7+x^5+x^4+x^2+x^1+x^0.
	///
	/// Polynomials over GF(2) are represented in binary, one bit per coefficient,
	/// with the lowest powers in the most significant bit.  Then adding polynomials
	/// is just exclusive-or, and multiplying a polynomial by x is a right shift by
	/// one.  If we call the above polynomial p, and represent a byte as the
	/// polynomial q, also with the lowest power in the most significant bit (so the
	/// byte 0xb1 is the polynomial x^7+x^3+x+1), then the CRC is (q*x^32) mod p,
	/// where a mod b means the remainder after dividing a by b.
	///
	/// This calculation is done using the shift-register method of multiplying and
	/// taking the remainder.  The register is initialized to zero, and for each
	/// incoming bit, x^32 is added mod p to the register if the bit is a one (where
	/// x^32 mod p is p+x^32 = x^26+...+1), and the register is multiplied mod p by
	/// x (which is shifting right by one and adding x^32 mod p if the bit shifted
	/// out is a one).  We start with the highest power (least significant bit) of
	/// q and repeat for all eight bits of q.
	///
	/// This implementation uses sixteen lookup tables stored in one linear array
	/// to implement the slicing-by-16 algorithm, a variant of the slicing-by-8
	/// algorithm described in this Intel white paper:
	///
	/// https://web.archive.org/web/20120722193753/http://download.intel.com/technology/comms/perfnet/download/slicing-by-8.pdf
	///
	/// The first lookup table is simply the CRC of all possible eight bit values.
	/// Each successive lookup table is derived from the original table generated
	/// by Sarwate's algorithm. Slicing a 16-bit input and XORing the outputs
	/// together will produce the same output as a byte-by-byte CRC loop with
	/// fewer arithmetic and bit manipulation operations, at the cost of increased
	/// memory consumed by the lookup tables. (Slicing-by-16 requires a 16KB table,
	/// which is still small enough to fit in most processors' L1 cache.)
	/// </remarks>
	public sealed class BZip2Crc : IChecksum
	{
		#region Instance Fields

		private const uint crcInit = 0xFFFFFFFF;
		//const uint crcXor = 0x00000000;

		private static readonly uint[] crcTable = CrcUtilities.GenerateSlicingLookupTable(0x04C11DB7, isReversed: false);

		/// <summary>
		/// The CRC data checksum so far.
		/// </summary>
		private uint checkValue;

		#endregion Instance Fields

		/// <summary>
		/// Initialise a default instance of <see cref="BZip2Crc"></see>
		/// </summary>
		public BZip2Crc()
		{
			Reset();
		}

		/// <summary>
		/// Resets the CRC data checksum as if no update was ever called.
		/// </summary>
		public void Reset()
		{
			checkValue = crcInit;
		}

		/// <summary>
		/// Returns the CRC data checksum computed so far.
		/// </summary>
		/// <remarks>Reversed Out = true</remarks>
		public long Value
		{
			get
			{
				// Technically, the output should be:
				//return (long)(~checkValue ^ crcXor);
				// but x ^ 0 = x, so there is no point in adding
				// the XOR operation
				return (long)(~checkValue);
			}
		}

		/// <summary>
		/// Updates the checksum with the int bval.
		/// </summary>
		/// <param name = "bval">
		/// the byte is taken as the lower 8 bits of bval
		/// </param>
		/// <remarks>Reversed Data = false</remarks>
		[MethodImpl(MethodImplOptions.AggressiveInlining)]
		public void Update(int bval)
		{
			checkValue = unchecked(crcTable[(byte)(((checkValue >> 24) & 0xFF) ^ bval)] ^ (checkValue << 8));
		}

		/// <summary>
		/// Updates the CRC data checksum with the bytes taken from
		/// a block of data.
		/// </summary>
		/// <param name="buffer">Contains the data to update the CRC with.</param>
		public void Update(byte[] buffer)
		{
			if (buffer == null)
			{
				throw new ArgumentNullException(nameof(buffer));
			}

			Update(buffer, 0, buffer.Length);
		}

		/// <summary>
		/// Update CRC data checksum based on a portion of a block of data
		/// </summary>
		/// <param name = "segment">
		/// The chunk of data to add
		/// </param>
		public void Update(ArraySegment<byte> segment)
		{
			Update(segment.Array, segment.Offset, segment.Count);
		}

		/// <summary>
		/// </summary>
		/// <param name="data">The array containing the data to add</param>
		/// <param name="offset">Range start for <paramref name="data"/> (inclusive)</param>
		/// <param name="count">The number of bytes to checksum starting from <paramref name="offset"/></param>
		private void Update(byte[] data, int offset, int count)
		{
			int remainder = count % CrcUtilities.SlicingDegree;
			int end = offset + count - remainder;

			while (offset != end)
			{
				checkValue = CrcUtilities.UpdateDataForNormalPoly(data, offset, crcTable, checkValue);
				offset += CrcUtilities.SlicingDegree;
			}

			if (remainder != 0)
			{
				SlowUpdateLoop(data, offset, end + remainder);
			}
		}

		/// <summary>
		/// A non-inlined function for updating data that doesn't fit in a 16-byte
		/// block. We don't expect to enter this function most of the time, and when
		/// we do we're not here for long, so disabling inlining here improves
		/// performance overall.
		/// </summary>
		/// <param name="data">The array containing the data to add</param>
		/// <param name="offset">Range start for <paramref name="data"/> (inclusive)</param>
		/// <param name="end">Range end for <paramref name="data"/> (exclusive)</param>
		[MethodImpl(MethodImplOptions.NoInlining)]
		private void SlowUpdateLoop(byte[] data, int offset, int end)
		{
			while (offset != end)
			{
				Update(data[offset++]);
			}
		}
	}
}


namespace ICSharpCode.SharpZipLib.Checksum
{
	/// <summary>
	/// CRC-32 with reversed data and unreversed output
	/// </summary>
	/// <remarks>
	/// Generate a table for a byte-wise 32-bit CRC calculation on the polynomial:
	/// x^32+x^26+x^23+x^22+x^16+x^12+x^11+x^10+x^8+x^7+x^5+x^4+x^2+x^1+x^0.
	///
	/// Polynomials over GF(2) are represented in binary, one bit per coefficient,
	/// with the lowest powers in the most significant bit.  Then adding polynomials
	/// is just exclusive-or, and multiplying a polynomial by x is a right shift by
	/// one.  If we call the above polynomial p, and represent a byte as the
	/// polynomial q, also with the lowest power in the most significant bit (so the
	/// byte 0xb1 is the polynomial x^7+x^3+x+1), then the CRC is (q*x^32) mod p,
	/// where a mod b means the remainder after dividing a by b.
	///
	/// This calculation is done using the shift-register method of multiplying and
	/// taking the remainder.  The register is initialized to zero, and for each
	/// incoming bit, x^32 is added mod p to the register if the bit is a one (where
	/// x^32 mod p is p+x^32 = x^26+...+1), and the register is multiplied mod p by
	/// x (which is shifting right by one and adding x^32 mod p if the bit shifted
	/// out is a one).  We start with the highest power (least significant bit) of
	/// q and repeat for all eight bits of q.
	///
	/// This implementation uses sixteen lookup tables stored in one linear array
	/// to implement the slicing-by-16 algorithm, a variant of the slicing-by-8
	/// algorithm described in this Intel white paper:
	///
	/// https://web.archive.org/web/20120722193753/http://download.intel.com/technology/comms/perfnet/download/slicing-by-8.pdf
	///
	/// The first lookup table is simply the CRC of all possible eight bit values.
	/// Each successive lookup table is derived from the original table generated
	/// by Sarwate's algorithm. Slicing a 16-bit input and XORing the outputs
	/// together will produce the same output as a byte-by-byte CRC loop with
	/// fewer arithmetic and bit manipulation operations, at the cost of increased
	/// memory consumed by the lookup tables. (Slicing-by-16 requires a 16KB table,
	/// which is still small enough to fit in most processors' L1 cache.)
	/// </remarks>
	public sealed class Crc32 : IChecksum
	{
		#region Instance Fields

		private static readonly uint crcInit = 0xFFFFFFFF;
		private static readonly uint crcXor = 0xFFFFFFFF;

		private static readonly uint[] crcTable = CrcUtilities.GenerateSlicingLookupTable(0xEDB88320, isReversed: true);

		/// <summary>
		/// The CRC data checksum so far.
		/// </summary>
		private uint checkValue;

		#endregion Instance Fields

		[MethodImpl(MethodImplOptions.AggressiveInlining)]
		internal static uint ComputeCrc32(uint oldCrc, byte bval)
		{
			return (uint)(Crc32.crcTable[(oldCrc ^ bval) & 0xFF] ^ (oldCrc >> 8));
		}

		/// <summary>
		/// Initialise a default instance of <see cref="Crc32"></see>
		/// </summary>
		public Crc32()
		{
			Reset();
		}

		/// <summary>
		/// Resets the CRC data checksum as if no update was ever called.
		/// </summary>
		public void Reset()
		{
			checkValue = crcInit;
		}

		/// <summary>
		/// Returns the CRC data checksum computed so far.
		/// </summary>
		/// <remarks>Reversed Out = false</remarks>
		public long Value
		{
			get
			{
				return (long)(checkValue ^ crcXor);
			}
		}

		/// <summary>
		/// Updates the checksum with the int bval.
		/// </summary>
		/// <param name = "bval">
		/// the byte is taken as the lower 8 bits of bval
		/// </param>
		/// <remarks>Reversed Data = true</remarks>
		[MethodImpl(MethodImplOptions.AggressiveInlining)]
		public void Update(int bval)
		{
			checkValue = unchecked(crcTable[(checkValue ^ bval) & 0xFF] ^ (checkValue >> 8));
		}

		/// <summary>
		/// Updates the CRC data checksum with the bytes taken from
		/// a block of data.
		/// </summary>
		/// <param name="buffer">Contains the data to update the CRC with.</param>
		public void Update(byte[] buffer)
		{
			if (buffer == null)
			{
				throw new ArgumentNullException(nameof(buffer));
			}

			Update(buffer, 0, buffer.Length);
		}

		/// <summary>
		/// Update CRC data checksum based on a portion of a block of data
		/// </summary>
		/// <param name = "segment">
		/// The chunk of data to add
		/// </param>
		public void Update(ArraySegment<byte> segment)
		{
			Update(segment.Array, segment.Offset, segment.Count);
		}

		/// <summary>
		/// </summary>
		/// <param name="data">The array containing the data to add</param>
		/// <param name="offset">Range start for <paramref name="data"/> (inclusive)</param>
		/// <param name="count">The number of bytes to checksum starting from <paramref name="offset"/></param>
		private void Update(byte[] data, int offset, int count)
		{
			int remainder = count % CrcUtilities.SlicingDegree;
			int end = offset + count - remainder;

			while (offset != end)
			{
				checkValue = CrcUtilities.UpdateDataForReversedPoly(data, offset, crcTable, checkValue);
				offset += CrcUtilities.SlicingDegree;
			}

			if (remainder != 0)
			{
				SlowUpdateLoop(data, offset, end + remainder);
			}
		}

		/// <summary>
		/// A non-inlined function for updating data that doesn't fit in a 16-byte
		/// block. We don't expect to enter this function most of the time, and when
		/// we do we're not here for long, so disabling inlining here improves
		/// performance overall.
		/// </summary>
		/// <param name="data">The array containing the data to add</param>
		/// <param name="offset">Range start for <paramref name="data"/> (inclusive)</param>
		/// <param name="end">Range end for <paramref name="data"/> (exclusive)</param>
		[MethodImpl(MethodImplOptions.NoInlining)]
		private void SlowUpdateLoop(byte[] data, int offset, int end)
		{
			while (offset != end)
			{
				Update(data[offset++]);
			}
		}
	}
}


namespace ICSharpCode.SharpZipLib.Checksum
{
	internal static class CrcUtilities
	{
		/// <summary>
		/// The number of slicing lookup tables to generate.
		/// </summary>
		internal const int SlicingDegree = 16;

		/// <summary>
		/// Generates multiple CRC lookup tables for a given polynomial, stored
		/// in a linear array of uints. The first block (i.e. the first 256
		/// elements) is the same as the byte-by-byte CRC lookup table. 
		/// </summary>
		/// <param name="polynomial">The generating CRC polynomial</param>
		/// <param name="isReversed">Whether the polynomial is in reversed bit order</param>
		/// <returns>A linear array of 256 * <see cref="SlicingDegree"/> elements</returns>
		/// <remarks>
		/// This table could also be generated as a rectangular array, but the
		/// JIT compiler generates slower code than if we use a linear array.
		/// Known issue, see: https://github.com/dotnet/runtime/issues/30275
		/// </remarks>
		internal static uint[] GenerateSlicingLookupTable(uint polynomial, bool isReversed)
		{
			var table = new uint[256 * SlicingDegree];
			uint one = isReversed ? 1 : (1U << 31);

			for (int i = 0; i < 256; i++)
			{
				uint res = (uint)(isReversed ? i : i << 24);
				for (int j = 0; j < SlicingDegree; j++)
				{
					for (int k = 0; k < 8; k++)
					{
						if (isReversed)
						{
							res = (res & one) == 1 ? polynomial ^ (res >> 1) : res >> 1;
						}
						else
						{
							res = (res & one) != 0 ? polynomial ^ (res << 1) : res << 1;
						}
					}

					table[(256 * j) + i] = res;
				}
			}

			return table;
		}

		/// <summary>
		/// Mixes the first four bytes of input with <paramref name="checkValue"/>
		/// using normal ordering before calling <see cref="UpdateDataCommon"/>.
		/// </summary>
		/// <param name="input">Array of data to checksum</param>
		/// <param name="offset">Offset to start reading <paramref name="input"/> from</param>
		/// <param name="crcTable">The table to use for slicing-by-16 lookup</param>
		/// <param name="checkValue">Checksum state before this update call</param>
		/// <returns>A new unfinalized checksum value</returns>
		/// <seealso cref="UpdateDataForReversedPoly"/>
		/// <remarks>
		/// Assumes input[offset]..input[offset + 15] are valid array indexes.
		/// For performance reasons, this must be checked by the caller.
		/// </remarks>
		[MethodImpl(MethodImplOptions.AggressiveInlining)]
		internal static uint UpdateDataForNormalPoly(byte[] input, int offset, uint[] crcTable, uint checkValue)
		{
			byte x1 = (byte)((byte)(checkValue >> 24) ^ input[offset]);
			byte x2 = (byte)((byte)(checkValue >> 16) ^ input[offset + 1]);
			byte x3 = (byte)((byte)(checkValue >> 8) ^ input[offset + 2]);
			byte x4 = (byte)((byte)checkValue ^ input[offset + 3]);

			return UpdateDataCommon(input, offset, crcTable, x1, x2, x3, x4);
		}

		/// <summary>
		/// Mixes the first four bytes of input with <paramref name="checkValue"/>
		/// using reflected ordering before calling <see cref="UpdateDataCommon"/>.
		/// </summary>
		/// <param name="input">Array of data to checksum</param>
		/// <param name="offset">Offset to start reading <paramref name="input"/> from</param>
		/// <param name="crcTable">The table to use for slicing-by-16 lookup</param>
		/// <param name="checkValue">Checksum state before this update call</param>
		/// <returns>A new unfinalized checksum value</returns>
		/// <seealso cref="UpdateDataForNormalPoly"/>
		/// <remarks>
		/// Assumes input[offset]..input[offset + 15] are valid array indexes.
		/// For performance reasons, this must be checked by the caller.
		/// </remarks>
		[MethodImpl(MethodImplOptions.AggressiveInlining)]
		internal static uint UpdateDataForReversedPoly(byte[] input, int offset, uint[] crcTable, uint checkValue)
		{
			byte x1 = (byte)((byte)checkValue ^ input[offset]);
			byte x2 = (byte)((byte)(checkValue >>= 8) ^ input[offset + 1]);
			byte x3 = (byte)((byte)(checkValue >>= 8) ^ input[offset + 2]);
			byte x4 = (byte)((byte)(checkValue >>= 8) ^ input[offset + 3]);

			return UpdateDataCommon(input, offset, crcTable, x1, x2, x3, x4);
		}

		/// <summary>
		/// </summary>
		/// <param name="input">Array of data to checksum</param>
		/// <param name="offset">Offset to start reading <paramref name="input"/> from</param>
		/// <param name="crcTable">The table to use for slicing-by-16 lookup</param>
		/// <param name="x1">First byte of input after mixing with the old CRC</param>
		/// <param name="x2">Second byte of input after mixing with the old CRC</param>
		/// <param name="x3">Third byte of input after mixing with the old CRC</param>
		/// <param name="x4">Fourth byte of input after mixing with the old CRC</param>
		/// <returns>A new unfinalized checksum value</returns>
		/// <remarks>
		/// <para>
		/// Even though the first four bytes of input are fed in as arguments,
		/// <paramref name="offset"/> should be the same value passed to this
		/// function's caller (either <see cref="UpdateDataForNormalPoly"/> or
		/// <see cref="UpdateDataForReversedPoly"/>). This method will get inlined
		/// into both functions, so using the same offset produces faster code.
		/// </para>
		/// <para>
		/// Because most processors running C# have some kind of instruction-level
		/// parallelism, the order of XOR operations can affect performance. This
		/// ordering assumes that the assembly code generated by the just-in-time
		/// compiler will emit a bunch of arithmetic operations for checking array
		/// bounds. Then it opportunistically XORs a1 and a2 to keep the processor
		/// busy while those other parts of the pipeline handle the range check
		/// calculations.
		/// </para>
		/// </remarks>
		[MethodImpl(MethodImplOptions.AggressiveInlining)]
		private static uint UpdateDataCommon(byte[] input, int offset, uint[] crcTable, byte x1, byte x2, byte x3, byte x4)
		{
			uint result;
			uint a1 = crcTable[x1 + 3840] ^ crcTable[x2 + 3584];
			uint a2 = crcTable[x3 + 3328] ^ crcTable[x4 + 3072];

			result = crcTable[input[offset + 4] + 2816];
			result ^= crcTable[input[offset + 5] + 2560];
			a1 ^= crcTable[input[offset + 9] + 1536];
			result ^= crcTable[input[offset + 6] + 2304];
			result ^= crcTable[input[offset + 7] + 2048];
			result ^= crcTable[input[offset + 8] + 1792];
			a2 ^= crcTable[input[offset + 13] + 512];
			result ^= crcTable[input[offset + 10] + 1280];
			result ^= crcTable[input[offset + 11] + 1024];
			result ^= crcTable[input[offset + 12] + 768];
			result ^= a1;
			result ^= crcTable[input[offset + 14] + 256];
			result ^= crcTable[input[offset + 15]];
			result ^= a2;

			return result;
		}
	}
}


namespace ICSharpCode.SharpZipLib.Checksum
{
	/// <summary>
	/// Interface to compute a data checksum used by checked input/output streams.
	/// A data checksum can be updated by one byte or with a byte array. After each
	/// update the value of the current checksum can be returned by calling
	/// <code>getValue</code>. The complete checksum object can also be reset
	/// so it can be used again with new data.
	/// </summary>
	public interface IChecksum
	{
		/// <summary>
		/// Resets the data checksum as if no update was ever called.
		/// </summary>
		void Reset();

		/// <summary>
		/// Returns the data checksum computed so far.
		/// </summary>
		long Value
		{
			get;
		}

		/// <summary>
		/// Adds one byte to the data checksum.
		/// </summary>
		/// <param name = "bval">
		/// the data value to add. The high byte of the int is ignored.
		/// </param>
		void Update(int bval);

		/// <summary>
		/// Updates the data checksum with the bytes taken from the array.
		/// </summary>
		/// <param name="buffer">
		/// buffer an array of bytes
		/// </param>
		void Update(byte[] buffer);

		/// <summary>
		/// Adds the byte array to the data checksum.
		/// </summary>
		/// <param name = "segment">
		/// The chunk of data to add
		/// </param>
		void Update(ArraySegment<byte> segment);
	}
}


namespace ICSharpCode.SharpZipLib.Core
{
	internal static class Empty
	{
#if NET45
		internal static class EmptyArray<T>
		{
			public static readonly T[] Value = new T[0];
		}
		public static T[] Array<T>() => EmptyArray<T>.Value;
#else
		public static T[] Array<T>() => System.Array.Empty<T>();
#endif
	}
}


namespace ICSharpCode.SharpZipLib.Core
{
	#region EventArgs

	/// <summary>
	/// Event arguments for scanning.
	/// </summary>
	public class ScanEventArgs : EventArgs
	{
		#region Constructors

		/// <summary>
		/// Initialise a new instance of <see cref="ScanEventArgs"/>
		/// </summary>
		/// <param name="name">The file or directory name.</param>
		public ScanEventArgs(string name)
		{
			name_ = name;
		}

		#endregion Constructors

		/// <summary>
		/// The file or directory name for this event.
		/// </summary>
		public string Name
		{
			get { return name_; }
		}

		/// <summary>
		/// Get set a value indicating if scanning should continue or not.
		/// </summary>
		public bool ContinueRunning
		{
			get { return continueRunning_; }
			set { continueRunning_ = value; }
		}

		#region Instance Fields

		private string name_;
		private bool continueRunning_ = true;

		#endregion Instance Fields
	}

	/// <summary>
	/// Event arguments during processing of a single file or directory.
	/// </summary>
	public class ProgressEventArgs : EventArgs
	{
		#region Constructors

		/// <summary>
		/// Initialise a new instance of <see cref="ScanEventArgs"/>
		/// </summary>
		/// <param name="name">The file or directory name if known.</param>
		/// <param name="processed">The number of bytes processed so far</param>
		/// <param name="target">The total number of bytes to process, 0 if not known</param>
		public ProgressEventArgs(string name, long processed, long target)
		{
			name_ = name;
			processed_ = processed;
			target_ = target;
		}

		#endregion Constructors

		/// <summary>
		/// The name for this event if known.
		/// </summary>
		public string Name
		{
			get { return name_; }
		}

		/// <summary>
		/// Get set a value indicating whether scanning should continue or not.
		/// </summary>
		public bool ContinueRunning
		{
			get { return continueRunning_; }
			set { continueRunning_ = value; }
		}

		/// <summary>
		/// Get a percentage representing how much of the <see cref="Target"></see> has been processed
		/// </summary>
		/// <value>0.0 to 100.0 percent; 0 if target is not known.</value>
		public float PercentComplete
		{
			get
			{
				float result;
				if (target_ <= 0)
				{
					result = 0;
				}
				else
				{
					result = ((float)processed_ / (float)target_) * 100.0f;
				}
				return result;
			}
		}

		/// <summary>
		/// The number of bytes processed so far
		/// </summary>
		public long Processed
		{
			get { return processed_; }
		}

		/// <summary>
		/// The number of bytes to process.
		/// </summary>
		/// <remarks>Target may be 0 or negative if the value isnt known.</remarks>
		public long Target
		{
			get { return target_; }
		}

		#region Instance Fields

		private string name_;
		private long processed_;
		private long target_;
		private bool continueRunning_ = true;

		#endregion Instance Fields
	}

	/// <summary>
	/// Event arguments for directories.
	/// </summary>
	public class DirectoryEventArgs : ScanEventArgs
	{
		#region Constructors

		/// <summary>
		/// Initialize an instance of <see cref="DirectoryEventArgs"></see>.
		/// </summary>
		/// <param name="name">The name for this directory.</param>
		/// <param name="hasMatchingFiles">Flag value indicating if any matching files are contained in this directory.</param>
		public DirectoryEventArgs(string name, bool hasMatchingFiles)
			: base(name)
		{
			hasMatchingFiles_ = hasMatchingFiles;
		}

		#endregion Constructors

		/// <summary>
		/// Get a value indicating if the directory contains any matching files or not.
		/// </summary>
		public bool HasMatchingFiles
		{
			get { return hasMatchingFiles_; }
		}

		private readonly

		#region Instance Fields

		bool hasMatchingFiles_;

		#endregion Instance Fields
	}

	/// <summary>
	/// Arguments passed when scan failures are detected.
	/// </summary>
	public class ScanFailureEventArgs : EventArgs
	{
		#region Constructors

		/// <summary>
		/// Initialise a new instance of <see cref="ScanFailureEventArgs"></see>
		/// </summary>
		/// <param name="name">The name to apply.</param>
		/// <param name="e">The exception to use.</param>
		public ScanFailureEventArgs(string name, Exception e)
		{
			name_ = name;
			exception_ = e;
			continueRunning_ = true;
		}

		#endregion Constructors

		/// <summary>
		/// The applicable name.
		/// </summary>
		public string Name
		{
			get { return name_; }
		}

		/// <summary>
		/// The applicable exception.
		/// </summary>
		public Exception Exception
		{
			get { return exception_; }
		}

		/// <summary>
		/// Get / set a value indicating whether scanning should continue.
		/// </summary>
		public bool ContinueRunning
		{
			get { return continueRunning_; }
			set { continueRunning_ = value; }
		}

		#region Instance Fields

		private string name_;
		private Exception exception_;
		private bool continueRunning_;

		#endregion Instance Fields
	}

	#endregion EventArgs

	#region Delegates

	/// <summary>
	/// Delegate invoked before starting to process a file.
	/// </summary>
	/// <param name="sender">The source of the event</param>
	/// <param name="e">The event arguments.</param>
	public delegate void ProcessFileHandler(object sender, ScanEventArgs e);

	/// <summary>
	/// Delegate invoked during processing of a file or directory
	/// </summary>
	/// <param name="sender">The source of the event</param>
	/// <param name="e">The event arguments.</param>
	public delegate void ProgressHandler(object sender, ProgressEventArgs e);

	/// <summary>
	/// Delegate invoked when a file has been completely processed.
	/// </summary>
	/// <param name="sender">The source of the event</param>
	/// <param name="e">The event arguments.</param>
	public delegate void CompletedFileHandler(object sender, ScanEventArgs e);

	/// <summary>
	/// Delegate invoked when a directory failure is detected.
	/// </summary>
	/// <param name="sender">The source of the event</param>
	/// <param name="e">The event arguments.</param>
	public delegate void DirectoryFailureHandler(object sender, ScanFailureEventArgs e);

	/// <summary>
	/// Delegate invoked when a file failure is detected.
	/// </summary>
	/// <param name="sender">The source of the event</param>
	/// <param name="e">The event arguments.</param>
	public delegate void FileFailureHandler(object sender, ScanFailureEventArgs e);

	#endregion Delegates

	/// <summary>
	/// FileSystemScanner provides facilities scanning of files and directories.
	/// </summary>
	public class FileSystemScanner
	{
		#region Constructors

		/// <summary>
		/// Initialise a new instance of <see cref="FileSystemScanner"></see>
		/// </summary>
		/// <param name="filter">The <see cref="PathFilter">file filter</see> to apply when scanning.</param>
		public FileSystemScanner(string filter)
		{
			fileFilter_ = new PathFilter(filter);
		}

		/// <summary>
		/// Initialise a new instance of <see cref="FileSystemScanner"></see>
		/// </summary>
		/// <param name="fileFilter">The <see cref="PathFilter">file filter</see> to apply.</param>
		/// <param name="directoryFilter">The <see cref="PathFilter"> directory filter</see> to apply.</param>
		public FileSystemScanner(string fileFilter, string directoryFilter)
		{
			fileFilter_ = new PathFilter(fileFilter);
			directoryFilter_ = new PathFilter(directoryFilter);
		}

		/// <summary>
		/// Initialise a new instance of <see cref="FileSystemScanner"></see>
		/// </summary>
		/// <param name="fileFilter">The file <see cref="IScanFilter">filter</see> to apply.</param>
		public FileSystemScanner(IScanFilter fileFilter)
		{
			fileFilter_ = fileFilter;
		}

		/// <summary>
		/// Initialise a new instance of <see cref="FileSystemScanner"></see>
		/// </summary>
		/// <param name="fileFilter">The file <see cref="IScanFilter">filter</see>  to apply.</param>
		/// <param name="directoryFilter">The directory <see cref="IScanFilter">filter</see>  to apply.</param>
		public FileSystemScanner(IScanFilter fileFilter, IScanFilter directoryFilter)
		{
			fileFilter_ = fileFilter;
			directoryFilter_ = directoryFilter;
		}

		#endregion Constructors

		#region Delegates

		/// <summary>
		/// Delegate to invoke when a directory is processed.
		/// </summary>
		public event EventHandler<DirectoryEventArgs> ProcessDirectory;

		/// <summary>
		/// Delegate to invoke when a file is processed.
		/// </summary>
		public ProcessFileHandler ProcessFile;

		/// <summary>
		/// Delegate to invoke when processing for a file has finished.
		/// </summary>
		public CompletedFileHandler CompletedFile;

		/// <summary>
		/// Delegate to invoke when a directory failure is detected.
		/// </summary>
		public DirectoryFailureHandler DirectoryFailure;

		/// <summary>
		/// Delegate to invoke when a file failure is detected.
		/// </summary>
		public FileFailureHandler FileFailure;

		#endregion Delegates

		/// <summary>
		/// Raise the DirectoryFailure event.
		/// </summary>
		/// <param name="directory">The directory name.</param>
		/// <param name="e">The exception detected.</param>
		private bool OnDirectoryFailure(string directory, Exception e)
		{
			DirectoryFailureHandler handler = DirectoryFailure;
			bool result = (handler != null);
			if (result)
			{
				var args = new ScanFailureEventArgs(directory, e);
				handler(this, args);
				alive_ = args.ContinueRunning;
			}
			return result;
		}

		/// <summary>
		/// Raise the FileFailure event.
		/// </summary>
		/// <param name="file">The file name.</param>
		/// <param name="e">The exception detected.</param>
		private bool OnFileFailure(string file, Exception e)
		{
			FileFailureHandler handler = FileFailure;

			bool result = (handler != null);

			if (result)
			{
				var args = new ScanFailureEventArgs(file, e);
				FileFailure(this, args);
				alive_ = args.ContinueRunning;
			}
			return result;
		}

		/// <summary>
		/// Raise the ProcessFile event.
		/// </summary>
		/// <param name="file">The file name.</param>
		private void OnProcessFile(string file)
		{
			ProcessFileHandler handler = ProcessFile;

			if (handler != null)
			{
				var args = new ScanEventArgs(file);
				handler(this, args);
				alive_ = args.ContinueRunning;
			}
		}

		/// <summary>
		/// Raise the complete file event
		/// </summary>
		/// <param name="file">The file name</param>
		private void OnCompleteFile(string file)
		{
			CompletedFileHandler handler = CompletedFile;

			if (handler != null)
			{
				var args = new ScanEventArgs(file);
				handler(this, args);
				alive_ = args.ContinueRunning;
			}
		}

		/// <summary>
		/// Raise the ProcessDirectory event.
		/// </summary>
		/// <param name="directory">The directory name.</param>
		/// <param name="hasMatchingFiles">Flag indicating if the directory has matching files.</param>
		private void OnProcessDirectory(string directory, bool hasMatchingFiles)
		{
			EventHandler<DirectoryEventArgs> handler = ProcessDirectory;

			if (handler != null)
			{
				var args = new DirectoryEventArgs(directory, hasMatchingFiles);
				handler(this, args);
				alive_ = args.ContinueRunning;
			}
		}

		/// <summary>
		/// Scan a directory.
		/// </summary>
		/// <param name="directory">The base directory to scan.</param>
		/// <param name="recurse">True to recurse subdirectories, false to scan a single directory.</param>
		public void Scan(string directory, bool recurse)
		{
			alive_ = true;
			ScanDir(directory, recurse);
		}

		private void ScanDir(string directory, bool recurse)
		{
			try
			{
				string[] names = System.IO.Directory.GetFiles(directory);
				bool hasMatch = false;
				for (int fileIndex = 0; fileIndex < names.Length; ++fileIndex)
				{
					if (!fileFilter_.IsMatch(names[fileIndex]))
					{
						names[fileIndex] = null;
					}
					else
					{
						hasMatch = true;
					}
				}

				OnProcessDirectory(directory, hasMatch);

				if (alive_ && hasMatch)
				{
					foreach (string fileName in names)
					{
						try
						{
							if (fileName != null)
							{
								OnProcessFile(fileName);
								if (!alive_)
								{
									break;
								}
							}
						}
						catch (Exception e)
						{
							if (!OnFileFailure(fileName, e))
							{
								throw;
							}
						}
					}
				}
			}
			catch (Exception e)
			{
				if (!OnDirectoryFailure(directory, e))
				{
					throw;
				}
			}

			if (alive_ && recurse)
			{
				try
				{
					string[] names = System.IO.Directory.GetDirectories(directory);
					foreach (string fulldir in names)
					{
						if ((directoryFilter_ == null) || (directoryFilter_.IsMatch(fulldir)))
						{
							ScanDir(fulldir, true);
							if (!alive_)
							{
								break;
							}
						}
					}
				}
				catch (Exception e)
				{
					if (!OnDirectoryFailure(directory, e))
					{
						throw;
					}
				}
			}
		}

		#region Instance Fields

		/// <summary>
		/// The file filter currently in use.
		/// </summary>
		private IScanFilter fileFilter_;

		/// <summary>
		/// The directory filter currently in use.
		/// </summary>
		private IScanFilter directoryFilter_;

		/// <summary>
		/// Flag indicating if scanning should continue running.
		/// </summary>
		private bool alive_;

		#endregion Instance Fields
	}
}

namespace ICSharpCode.SharpZipLib.Core
{
	/// <summary>
	/// INameTransform defines how file system names are transformed for use with archives, or vice versa.
	/// </summary>
	public interface INameTransform
	{
		/// <summary>
		/// Given a file name determine the transformed value.
		/// </summary>
		/// <param name="name">The name to transform.</param>
		/// <returns>The transformed file name.</returns>
		string TransformFile(string name);

		/// <summary>
		/// Given a directory name determine the transformed value.
		/// </summary>
		/// <param name="name">The name to transform.</param>
		/// <returns>The transformed directory name</returns>
		string TransformDirectory(string name);
	}
}


namespace ICSharpCode.SharpZipLib.Core
{
	/// <summary>
	/// InvalidNameException is thrown for invalid names such as directory traversal paths and names with invalid characters
	/// </summary>
	[Serializable]
	public class InvalidNameException : SharpZipBaseException
	{
		/// <summary>
		/// Initializes a new instance of the InvalidNameException class with a default error message.
		/// </summary>
		public InvalidNameException() : base("An invalid name was specified")
		{
		}

		/// <summary>
		/// Initializes a new instance of the InvalidNameException class with a specified error message.
		/// </summary>
		/// <param name="message">A message describing the exception.</param>
		public InvalidNameException(string message) : base(message)
		{
		}

		/// <summary>
		/// Initializes a new instance of the InvalidNameException class with a specified
		/// error message and a reference to the inner exception that is the cause of this exception.
		/// </summary>
		/// <param name="message">A message describing the exception.</param>
		/// <param name="innerException">The inner exception</param>
		public InvalidNameException(string message, Exception innerException) : base(message, innerException)
		{
		}

		/// <summary>
		/// Initializes a new instance of the InvalidNameException class with serialized data.
		/// </summary>
		/// <param name="info">
		/// The System.Runtime.Serialization.SerializationInfo that holds the serialized
		/// object data about the exception being thrown.
		/// </param>
		/// <param name="context">
		/// The System.Runtime.Serialization.StreamingContext that contains contextual information
		/// about the source or destination.
		/// </param>
		protected InvalidNameException(SerializationInfo info, StreamingContext context)
			: base(info, context)
		{
		}
	}
}

namespace ICSharpCode.SharpZipLib.Core
{
	/// <summary>
	/// Scanning filters support filtering of names.
	/// </summary>
	public interface IScanFilter
	{
		/// <summary>
		/// Test a name to see if it 'matches' the filter.
		/// </summary>
		/// <param name="name">The name to test.</param>
		/// <returns>Returns true if the name matches the filter, false if it does not match.</returns>
		bool IsMatch(string name);
	}
}


namespace ICSharpCode.SharpZipLib.Core
{
	/// <summary>
	/// NameFilter is a string matching class which allows for both positive and negative
	/// matching.
	/// A filter is a sequence of independant <see cref="Regex">regular expressions</see> separated by semi-colons ';'.
	/// To include a semi-colon it may be quoted as in \;. Each expression can be prefixed by a plus '+' sign or
	/// a minus '-' sign to denote the expression is intended to include or exclude names.
	/// If neither a plus or minus sign is found include is the default.
	/// A given name is tested for inclusion before checking exclusions.  Only names matching an include spec
	/// and not matching an exclude spec are deemed to match the filter.
	/// An empty filter matches any name.
	/// </summary>
	/// <example>The following expression includes all name ending in '.dat' with the exception of 'dummy.dat'
	/// "+\.dat$;-^dummy\.dat$"
	/// </example>
	public class NameFilter : IScanFilter
	{
		#region Constructors

		/// <summary>
		/// Construct an instance based on the filter expression passed
		/// </summary>
		/// <param name="filter">The filter expression.</param>
		public NameFilter(string filter)
		{
			filter_ = filter;
			inclusions_ = new List<Regex>();
			exclusions_ = new List<Regex>();
			Compile();
		}

		#endregion Constructors

		/// <summary>
		/// Test a string to see if it is a valid regular expression.
		/// </summary>
		/// <param name="expression">The expression to test.</param>
		/// <returns>True if expression is a valid <see cref="System.Text.RegularExpressions.Regex"/> false otherwise.</returns>
		public static bool IsValidExpression(string expression)
		{
			bool result = true;
			try
			{
				var exp = new Regex(expression, RegexOptions.IgnoreCase | RegexOptions.Singleline);
			}
			catch (ArgumentException)
			{
				result = false;
			}
			return result;
		}

		/// <summary>
		/// Test an expression to see if it is valid as a filter.
		/// </summary>
		/// <param name="toTest">The filter expression to test.</param>
		/// <returns>True if the expression is valid, false otherwise.</returns>
		public static bool IsValidFilterExpression(string toTest)
		{
			bool result = true;

			try
			{
				if (toTest != null)
				{
					string[] items = SplitQuoted(toTest);
					for (int i = 0; i < items.Length; ++i)
					{
						if ((items[i] != null) && (items[i].Length > 0))
						{
							string toCompile;

							if (items[i][0] == '+')
							{
								toCompile = items[i].Substring(1, items[i].Length - 1);
							}
							else if (items[i][0] == '-')
							{
								toCompile = items[i].Substring(1, items[i].Length - 1);
							}
							else
							{
								toCompile = items[i];
							}

							var testRegex = new Regex(toCompile, RegexOptions.IgnoreCase | RegexOptions.Singleline);
						}
					}
				}
			}
			catch (ArgumentException)
			{
				result = false;
			}

			return result;
		}

		/// <summary>
		/// Split a string into its component pieces
		/// </summary>
		/// <param name="original">The original string</param>
		/// <returns>Returns an array of <see cref="System.String"/> values containing the individual filter elements.</returns>
		public static string[] SplitQuoted(string original)
		{
			char escape = '\\';
			char[] separators = { ';' };

			var result = new List<string>();

			if (!string.IsNullOrEmpty(original))
			{
				int endIndex = -1;
				var b = new StringBuilder();

				while (endIndex < original.Length)
				{
					endIndex += 1;
					if (endIndex >= original.Length)
					{
						result.Add(b.ToString());
					}
					else if (original[endIndex] == escape)
					{
						endIndex += 1;
						if (endIndex >= original.Length)
						{
							throw new ArgumentException("Missing terminating escape character", nameof(original));
						}
						// include escape if this is not an escaped separator
						if (Array.IndexOf(separators, original[endIndex]) < 0)
							b.Append(escape);

						b.Append(original[endIndex]);
					}
					else
					{
						if (Array.IndexOf(separators, original[endIndex]) >= 0)
						{
							result.Add(b.ToString());
							b.Length = 0;
						}
						else
						{
							b.Append(original[endIndex]);
						}
					}
				}
			}

			return result.ToArray();
		}

		/// <summary>
		/// Convert this filter to its string equivalent.
		/// </summary>
		/// <returns>The string equivalent for this filter.</returns>
		public override string ToString()
		{
			return filter_;
		}

		/// <summary>
		/// Test a value to see if it is included by the filter.
		/// </summary>
		/// <param name="name">The value to test.</param>
		/// <returns>True if the value is included, false otherwise.</returns>
		public bool IsIncluded(string name)
		{
			bool result = false;
			if (inclusions_.Count == 0)
			{
				result = true;
			}
			else
			{
				foreach (Regex r in inclusions_)
				{
					if (r.IsMatch(name))
					{
						result = true;
						break;
					}
				}
			}
			return result;
		}

		/// <summary>
		/// Test a value to see if it is excluded by the filter.
		/// </summary>
		/// <param name="name">The value to test.</param>
		/// <returns>True if the value is excluded, false otherwise.</returns>
		public bool IsExcluded(string name)
		{
			bool result = false;
			foreach (Regex r in exclusions_)
			{
				if (r.IsMatch(name))
				{
					result = true;
					break;
				}
			}
			return result;
		}

		#region IScanFilter Members

		/// <summary>
		/// Test a value to see if it matches the filter.
		/// </summary>
		/// <param name="name">The value to test.</param>
		/// <returns>True if the value matches, false otherwise.</returns>
		public bool IsMatch(string name)
		{
			return (IsIncluded(name) && !IsExcluded(name));
		}

		#endregion IScanFilter Members

		/// <summary>
		/// Compile this filter.
		/// </summary>
		private void Compile()
		{
			// TODO: Check to see if combining RE's makes it faster/smaller.
			// simple scheme would be to have one RE for inclusion and one for exclusion.
			if (filter_ == null)
			{
				return;
			}

			string[] items = SplitQuoted(filter_);
			for (int i = 0; i < items.Length; ++i)
			{
				if ((items[i] != null) && (items[i].Length > 0))
				{
					bool include = (items[i][0] != '-');
					string toCompile;

					if (items[i][0] == '+')
					{
						toCompile = items[i].Substring(1, items[i].Length - 1);
					}
					else if (items[i][0] == '-')
					{
						toCompile = items[i].Substring(1, items[i].Length - 1);
					}
					else
					{
						toCompile = items[i];
					}

					// NOTE: Regular expressions can fail to compile here for a number of reasons that cause an exception
					// these are left unhandled here as the caller is responsible for ensuring all is valid.
					// several functions IsValidFilterExpression and IsValidExpression are provided for such checking
					if (include)
					{
						inclusions_.Add(new Regex(toCompile, RegexOptions.IgnoreCase | RegexOptions.Compiled | RegexOptions.Singleline));
					}
					else
					{
						exclusions_.Add(new Regex(toCompile, RegexOptions.IgnoreCase | RegexOptions.Compiled | RegexOptions.Singleline));
					}
				}
			}
		}

		#region Instance Fields

		private string filter_;
		private List<Regex> inclusions_;
		private List<Regex> exclusions_;

		#endregion Instance Fields
	}
}


namespace ICSharpCode.SharpZipLib.Core
{
	/// <summary>
	/// PathFilter filters directories and files using a form of <see cref="System.Text.RegularExpressions.Regex">regular expressions</see>
	/// by full path name.
	/// See <see cref="NameFilter">NameFilter</see> for more detail on filtering.
	/// </summary>
	public class PathFilter : IScanFilter
	{
		#region Constructors

		/// <summary>
		/// Initialise a new instance of <see cref="PathFilter"></see>.
		/// </summary>
		/// <param name="filter">The <see cref="NameFilter">filter</see> expression to apply.</param>
		public PathFilter(string filter)
		{
			nameFilter_ = new NameFilter(filter);
		}

		#endregion Constructors

		#region IScanFilter Members

		/// <summary>
		/// Test a name to see if it matches the filter.
		/// </summary>
		/// <param name="name">The name to test.</param>
		/// <returns>True if the name matches, false otherwise.</returns>
		/// <remarks><see cref="Path.GetFullPath(string)"/> is used to get the full path before matching.</remarks>
		public virtual bool IsMatch(string name)
		{
			bool result = false;

			if (name != null)
			{
				string cooked = (name.Length > 0) ? Path.GetFullPath(name) : "";
				result = nameFilter_.IsMatch(cooked);
			}
			return result;
		}

		private readonly

		#endregion IScanFilter Members

		#region Instance Fields

		NameFilter nameFilter_;

		#endregion Instance Fields
	}

	/// <summary>
	/// ExtendedPathFilter filters based on name, file size, and the last write time of the file.
	/// </summary>
	/// <remarks>Provides an example of how to customise filtering.</remarks>
	public class ExtendedPathFilter : PathFilter
	{
		#region Constructors

		/// <summary>
		/// Initialise a new instance of ExtendedPathFilter.
		/// </summary>
		/// <param name="filter">The filter to apply.</param>
		/// <param name="minSize">The minimum file size to include.</param>
		/// <param name="maxSize">The maximum file size to include.</param>
		public ExtendedPathFilter(string filter,
			long minSize, long maxSize)
			: base(filter)
		{
			MinSize = minSize;
			MaxSize = maxSize;
		}

		/// <summary>
		/// Initialise a new instance of ExtendedPathFilter.
		/// </summary>
		/// <param name="filter">The filter to apply.</param>
		/// <param name="minDate">The minimum <see cref="DateTime"/> to include.</param>
		/// <param name="maxDate">The maximum <see cref="DateTime"/> to include.</param>
		public ExtendedPathFilter(string filter,
			DateTime minDate, DateTime maxDate)
			: base(filter)
		{
			MinDate = minDate;
			MaxDate = maxDate;
		}

		/// <summary>
		/// Initialise a new instance of ExtendedPathFilter.
		/// </summary>
		/// <param name="filter">The filter to apply.</param>
		/// <param name="minSize">The minimum file size to include.</param>
		/// <param name="maxSize">The maximum file size to include.</param>
		/// <param name="minDate">The minimum <see cref="DateTime"/> to include.</param>
		/// <param name="maxDate">The maximum <see cref="DateTime"/> to include.</param>
		public ExtendedPathFilter(string filter,
			long minSize, long maxSize,
			DateTime minDate, DateTime maxDate)
			: base(filter)
		{
			MinSize = minSize;
			MaxSize = maxSize;
			MinDate = minDate;
			MaxDate = maxDate;
		}

		#endregion Constructors

		#region IScanFilter Members

		/// <summary>
		/// Test a filename to see if it matches the filter.
		/// </summary>
		/// <param name="name">The filename to test.</param>
		/// <returns>True if the filter matches, false otherwise.</returns>
		/// <exception cref="System.IO.FileNotFoundException">The <see paramref="fileName"/> doesnt exist</exception>
		public override bool IsMatch(string name)
		{
			bool result = base.IsMatch(name);

			if (result)
			{
				var fileInfo = new FileInfo(name);
				result =
					(MinSize <= fileInfo.Length) &&
					(MaxSize >= fileInfo.Length) &&
					(MinDate <= fileInfo.LastWriteTime) &&
					(MaxDate >= fileInfo.LastWriteTime)
					;
			}
			return result;
		}

		#endregion IScanFilter Members

		#region Properties

		/// <summary>
		/// Get/set the minimum size/length for a file that will match this filter.
		/// </summary>
		/// <remarks>The default value is zero.</remarks>
		/// <exception cref="ArgumentOutOfRangeException">value is less than zero; greater than <see cref="MaxSize"/></exception>
		public long MinSize
		{
			get { return minSize_; }
			set
			{
				if ((value < 0) || (maxSize_ < value))
				{
					throw new ArgumentOutOfRangeException(nameof(value));
				}

				minSize_ = value;
			}
		}

		/// <summary>
		/// Get/set the maximum size/length for a file that will match this filter.
		/// </summary>
		/// <remarks>The default value is <see cref="System.Int64.MaxValue"/></remarks>
		/// <exception cref="ArgumentOutOfRangeException">value is less than zero or less than <see cref="MinSize"/></exception>
		public long MaxSize
		{
			get { return maxSize_; }
			set
			{
				if ((value < 0) || (minSize_ > value))
				{
					throw new ArgumentOutOfRangeException(nameof(value));
				}

				maxSize_ = value;
			}
		}

		/// <summary>
		/// Get/set the minimum <see cref="DateTime"/> value that will match for this filter.
		/// </summary>
		/// <remarks>Files with a LastWrite time less than this value are excluded by the filter.</remarks>
		public DateTime MinDate
		{
			get
			{
				return minDate_;
			}

			set
			{
				if (value > maxDate_)
				{
					throw new ArgumentOutOfRangeException(nameof(value), "Exceeds MaxDate");
				}

				minDate_ = value;
			}
		}

		/// <summary>
		/// Get/set the maximum <see cref="DateTime"/> value that will match for this filter.
		/// </summary>
		/// <remarks>Files with a LastWrite time greater than this value are excluded by the filter.</remarks>
		public DateTime MaxDate
		{
			get
			{
				return maxDate_;
			}

			set
			{
				if (minDate_ > value)
				{
					throw new ArgumentOutOfRangeException(nameof(value), "Exceeds MinDate");
				}

				maxDate_ = value;
			}
		}

		#endregion Properties

		#region Instance Fields

		private long minSize_;
		private long maxSize_ = long.MaxValue;
		private DateTime minDate_ = DateTime.MinValue;
		private DateTime maxDate_ = DateTime.MaxValue;

		#endregion Instance Fields
	}

	/// <summary>
	/// NameAndSizeFilter filters based on name and file size.
	/// </summary>
	/// <remarks>A sample showing how filters might be extended.</remarks>
	[Obsolete("Use ExtendedPathFilter instead")]
	public class NameAndSizeFilter : PathFilter
	{
		/// <summary>
		/// Initialise a new instance of NameAndSizeFilter.
		/// </summary>
		/// <param name="filter">The filter to apply.</param>
		/// <param name="minSize">The minimum file size to include.</param>
		/// <param name="maxSize">The maximum file size to include.</param>
		public NameAndSizeFilter(string filter, long minSize, long maxSize)
			: base(filter)
		{
			MinSize = minSize;
			MaxSize = maxSize;
		}

		/// <summary>
		/// Test a filename to see if it matches the filter.
		/// </summary>
		/// <param name="name">The filename to test.</param>
		/// <returns>True if the filter matches, false otherwise.</returns>
		public override bool IsMatch(string name)
		{
			bool result = base.IsMatch(name);

			if (result)
			{
				var fileInfo = new FileInfo(name);
				long length = fileInfo.Length;
				result =
					(MinSize <= length) &&
					(MaxSize >= length);
			}
			return result;
		}

		/// <summary>
		/// Get/set the minimum size for a file that will match this filter.
		/// </summary>
		public long MinSize
		{
			get { return minSize_; }
			set
			{
				if ((value < 0) || (maxSize_ < value))
				{
					throw new ArgumentOutOfRangeException(nameof(value));
				}

				minSize_ = value;
			}
		}

		/// <summary>
		/// Get/set the maximum size for a file that will match this filter.
		/// </summary>
		public long MaxSize
		{
			get { return maxSize_; }
			set
			{
				if ((value < 0) || (minSize_ > value))
				{
					throw new ArgumentOutOfRangeException(nameof(value));
				}

				maxSize_ = value;
			}
		}

		#region Instance Fields

		private long minSize_;
		private long maxSize_ = long.MaxValue;

		#endregion Instance Fields
	}
}


namespace ICSharpCode.SharpZipLib.Core
{
	/// <summary>
	/// PathUtils provides simple utilities for handling paths.
	/// </summary>
	public static class PathUtils
	{
		/// <summary>
		/// Remove any path root present in the path
		/// </summary>
		/// <param name="path">A <see cref="string"/> containing path information.</param>
		/// <returns>The path with the root removed if it was present; path otherwise.</returns>
		public static string DropPathRoot(string path)
		{
			var invalidChars = Path.GetInvalidPathChars();
			// If the first character after the root is a ':', .NET < 4.6.2 throws
			var cleanRootSep = path.Length >= 3 && path[1] == ':' && path[2] == ':';
			
			// Replace any invalid path characters with '_' to prevent Path.GetPathRoot from throwing.
			// Only pass the first 258 (should be 260, but that still throws for some reason) characters
			// as .NET < 4.6.2 throws on longer paths
			var cleanPath = new string(path.Take(258)
				.Select( (c, i) => invalidChars.Contains(c) || (i == 2 && cleanRootSep) ? '_' : c).ToArray());

			var stripLength = Path.GetPathRoot(cleanPath).Length;
			while (path.Length > stripLength && (path[stripLength] == '/' || path[stripLength] == '\\')) stripLength++;
			return path.Substring(stripLength);
		}

		/// <summary>
		/// Returns a random file name in the users temporary directory, or in directory of <paramref name="original"/> if specified
		/// </summary>
		/// <param name="original">If specified, used as the base file name for the temporary file</param>
		/// <returns>Returns a temporary file name</returns>
		public static string GetTempFileName(string original = null)
		{
			string fileName;
			var tempPath = Path.GetTempPath();

			do
			{
				fileName = original == null
					? Path.Combine(tempPath, Path.GetRandomFileName())
					: $"{original}.{Path.GetRandomFileName()}";
			} while (File.Exists(fileName));

			return fileName;
		}
	}
}


namespace ICSharpCode.SharpZipLib.Core
{
	/// <summary>
	/// Provides simple <see cref="Stream"/>" utilities.
	/// </summary>
	public sealed class StreamUtils
	{
		/// <summary>
		/// Read from a <see cref="Stream"/> ensuring all the required data is read.
		/// </summary>
		/// <param name="stream">The stream to read.</param>
		/// <param name="buffer">The buffer to fill.</param>
		/// <seealso cref="ReadFully(Stream,byte[],int,int)"/>
		static public void ReadFully(Stream stream, byte[] buffer)
		{
			ReadFully(stream, buffer, 0, buffer.Length);
		}

		/// <summary>
		/// Read from a <see cref="Stream"/>" ensuring all the required data is read.
		/// </summary>
		/// <param name="stream">The stream to read data from.</param>
		/// <param name="buffer">The buffer to store data in.</param>
		/// <param name="offset">The offset at which to begin storing data.</param>
		/// <param name="count">The number of bytes of data to store.</param>
		/// <exception cref="ArgumentNullException">Required parameter is null</exception>
		/// <exception cref="ArgumentOutOfRangeException"><paramref name="offset"/> and or <paramref name="count"/> are invalid.</exception>
		/// <exception cref="EndOfStreamException">End of stream is encountered before all the data has been read.</exception>
		static public void ReadFully(Stream stream, byte[] buffer, int offset, int count)
		{
			if (stream == null)
			{
				throw new ArgumentNullException(nameof(stream));
			}

			if (buffer == null)
			{
				throw new ArgumentNullException(nameof(buffer));
			}

			// Offset can equal length when buffer and count are 0.
			if ((offset < 0) || (offset > buffer.Length))
			{
				throw new ArgumentOutOfRangeException(nameof(offset));
			}

			if ((count < 0) || (offset + count > buffer.Length))
			{
				throw new ArgumentOutOfRangeException(nameof(count));
			}

			while (count > 0)
			{
				int readCount = stream.Read(buffer, offset, count);
				if (readCount <= 0)
				{
					throw new EndOfStreamException();
				}
				offset += readCount;
				count -= readCount;
			}
		}

		/// <summary>
		/// Read as much data as possible from a <see cref="Stream"/>", up to the requested number of bytes
		/// </summary>
		/// <param name="stream">The stream to read data from.</param>
		/// <param name="buffer">The buffer to store data in.</param>
		/// <param name="offset">The offset at which to begin storing data.</param>
		/// <param name="count">The number of bytes of data to store.</param>
		/// <exception cref="ArgumentNullException">Required parameter is null</exception>
		/// <exception cref="ArgumentOutOfRangeException"><paramref name="offset"/> and or <paramref name="count"/> are invalid.</exception>
		static public int ReadRequestedBytes(Stream stream, byte[] buffer, int offset, int count)
		{
			if (stream == null)
			{
				throw new ArgumentNullException(nameof(stream));
			}

			if (buffer == null)
			{
				throw new ArgumentNullException(nameof(buffer));
			}

			// Offset can equal length when buffer and count are 0.
			if ((offset < 0) || (offset > buffer.Length))
			{
				throw new ArgumentOutOfRangeException(nameof(offset));
			}

			if ((count < 0) || (offset + count > buffer.Length))
			{
				throw new ArgumentOutOfRangeException(nameof(count));
			}

			int totalReadCount = 0;
			while (count > 0)
			{
				int readCount = stream.Read(buffer, offset, count);
				if (readCount <= 0)
				{
					break;
				}
				offset += readCount;
				count -= readCount;
				totalReadCount += readCount;
			}

			return totalReadCount;
		}

		/// <summary>
		/// Copy the contents of one <see cref="Stream"/> to another.
		/// </summary>
		/// <param name="source">The stream to source data from.</param>
		/// <param name="destination">The stream to write data to.</param>
		/// <param name="buffer">The buffer to use during copying.</param>
		static public void Copy(Stream source, Stream destination, byte[] buffer)
		{
			if (source == null)
			{
				throw new ArgumentNullException(nameof(source));
			}

			if (destination == null)
			{
				throw new ArgumentNullException(nameof(destination));
			}

			if (buffer == null)
			{
				throw new ArgumentNullException(nameof(buffer));
			}

			// Ensure a reasonable size of buffer is used without being prohibitive.
			if (buffer.Length < 128)
			{
				throw new ArgumentException("Buffer is too small", nameof(buffer));
			}

			bool copying = true;

			while (copying)
			{
				int bytesRead = source.Read(buffer, 0, buffer.Length);
				if (bytesRead > 0)
				{
					destination.Write(buffer, 0, bytesRead);
				}
				else
				{
					destination.Flush();
					copying = false;
				}
			}
		}

		/// <summary>
		/// Copy the contents of one <see cref="Stream"/> to another.
		/// </summary>
		/// <param name="source">The stream to source data from.</param>
		/// <param name="destination">The stream to write data to.</param>
		/// <param name="buffer">The buffer to use during copying.</param>
		/// <param name="progressHandler">The <see cref="ProgressHandler">progress handler delegate</see> to use.</param>
		/// <param name="updateInterval">The minimum <see cref="TimeSpan"/> between progress updates.</param>
		/// <param name="sender">The source for this event.</param>
		/// <param name="name">The name to use with the event.</param>
		/// <remarks>This form is specialised for use within #Zip to support events during archive operations.</remarks>
		static public void Copy(Stream source, Stream destination,
			byte[] buffer, ProgressHandler progressHandler, TimeSpan updateInterval, object sender, string name)
		{
			Copy(source, destination, buffer, progressHandler, updateInterval, sender, name, -1);
		}

		/// <summary>
		/// Copy the contents of one <see cref="Stream"/> to another.
		/// </summary>
		/// <param name="source">The stream to source data from.</param>
		/// <param name="destination">The stream to write data to.</param>
		/// <param name="buffer">The buffer to use during copying.</param>
		/// <param name="progressHandler">The <see cref="ProgressHandler">progress handler delegate</see> to use.</param>
		/// <param name="updateInterval">The minimum <see cref="TimeSpan"/> between progress updates.</param>
		/// <param name="sender">The source for this event.</param>
		/// <param name="name">The name to use with the event.</param>
		/// <param name="fixedTarget">A predetermined fixed target value to use with progress updates.
		/// If the value is negative the target is calculated by looking at the stream.</param>
		/// <remarks>This form is specialised for use within #Zip to support events during archive operations.</remarks>
		static public void Copy(Stream source, Stream destination,
			byte[] buffer,
			ProgressHandler progressHandler, TimeSpan updateInterval,
			object sender, string name, long fixedTarget)
		{
			if (source == null)
			{
				throw new ArgumentNullException(nameof(source));
			}

			if (destination == null)
			{
				throw new ArgumentNullException(nameof(destination));
			}

			if (buffer == null)
			{
				throw new ArgumentNullException(nameof(buffer));
			}

			// Ensure a reasonable size of buffer is used without being prohibitive.
			if (buffer.Length < 128)
			{
				throw new ArgumentException("Buffer is too small", nameof(buffer));
			}

			if (progressHandler == null)
			{
				throw new ArgumentNullException(nameof(progressHandler));
			}

			bool copying = true;

			DateTime marker = DateTime.Now;
			long processed = 0;
			long target = 0;

			if (fixedTarget >= 0)
			{
				target = fixedTarget;
			}
			else if (source.CanSeek)
			{
				target = source.Length - source.Position;
			}

			// Always fire 0% progress..
			var args = new ProgressEventArgs(name, processed, target);
			progressHandler(sender, args);

			bool progressFired = true;

			while (copying)
			{
				int bytesRead = source.Read(buffer, 0, buffer.Length);
				if (bytesRead > 0)
				{
					processed += bytesRead;
					progressFired = false;
					destination.Write(buffer, 0, bytesRead);
				}
				else
				{
					destination.Flush();
					copying = false;
				}

				if (DateTime.Now - marker > updateInterval)
				{
					progressFired = true;
					marker = DateTime.Now;
					args = new ProgressEventArgs(name, processed, target);
					progressHandler(sender, args);

					copying = args.ContinueRunning;
				}
			}

			if (!progressFired)
			{
				args = new ProgressEventArgs(name, processed, target);
				progressHandler(sender, args);
			}
		}

		/// <summary>
		/// Initialise an instance of <see cref="StreamUtils"></see>
		/// </summary>
		private StreamUtils()
		{
			// Do nothing.
		}
	}
}


namespace ICSharpCode.SharpZipLib.Encryption
{
	/// <summary>
	/// PkzipClassic embodies the classic or original encryption facilities used in Pkzip archives.
	/// While it has been superceded by more recent and more powerful algorithms, its still in use and
	/// is viable for preventing casual snooping
	/// </summary>
	public abstract class PkzipClassic : SymmetricAlgorithm
	{
		/// <summary>
		/// Generates new encryption keys based on given seed
		/// </summary>
		/// <param name="seed">The seed value to initialise keys with.</param>
		/// <returns>A new key value.</returns>
		static public byte[] GenerateKeys(byte[] seed)
		{
			if (seed == null)
			{
				throw new ArgumentNullException(nameof(seed));
			}

			if (seed.Length == 0)
			{
				throw new ArgumentException("Length is zero", nameof(seed));
			}

			uint[] newKeys = {
				0x12345678,
				0x23456789,
				0x34567890
			 };

			for (int i = 0; i < seed.Length; ++i)
			{
				newKeys[0] = Crc32.ComputeCrc32(newKeys[0], seed[i]);
				newKeys[1] = newKeys[1] + (byte)newKeys[0];
				newKeys[1] = newKeys[1] * 134775813 + 1;
				newKeys[2] = Crc32.ComputeCrc32(newKeys[2], (byte)(newKeys[1] >> 24));
			}

			byte[] result = new byte[12];
			result[0] = (byte)(newKeys[0] & 0xff);
			result[1] = (byte)((newKeys[0] >> 8) & 0xff);
			result[2] = (byte)((newKeys[0] >> 16) & 0xff);
			result[3] = (byte)((newKeys[0] >> 24) & 0xff);
			result[4] = (byte)(newKeys[1] & 0xff);
			result[5] = (byte)((newKeys[1] >> 8) & 0xff);
			result[6] = (byte)((newKeys[1] >> 16) & 0xff);
			result[7] = (byte)((newKeys[1] >> 24) & 0xff);
			result[8] = (byte)(newKeys[2] & 0xff);
			result[9] = (byte)((newKeys[2] >> 8) & 0xff);
			result[10] = (byte)((newKeys[2] >> 16) & 0xff);
			result[11] = (byte)((newKeys[2] >> 24) & 0xff);
			return result;
		}
	}

	/// <summary>
	/// PkzipClassicCryptoBase provides the low level facilities for encryption
	/// and decryption using the PkzipClassic algorithm.
	/// </summary>
	internal class PkzipClassicCryptoBase
	{
		/// <summary>
		/// Transform a single byte
		/// </summary>
		/// <returns>
		/// The transformed value
		/// </returns>
		protected byte TransformByte()
		{
			uint temp = ((keys[2] & 0xFFFF) | 2);
			return (byte)((temp * (temp ^ 1)) >> 8);
		}

		/// <summary>
		/// Set the key schedule for encryption/decryption.
		/// </summary>
		/// <param name="keyData">The data use to set the keys from.</param>
		protected void SetKeys(byte[] keyData)
		{
			if (keyData == null)
			{
				throw new ArgumentNullException(nameof(keyData));
			}

			if (keyData.Length != 12)
			{
				throw new InvalidOperationException("Key length is not valid");
			}

			keys = new uint[3];
			keys[0] = (uint)((keyData[3] << 24) | (keyData[2] << 16) | (keyData[1] << 8) | keyData[0]);
			keys[1] = (uint)((keyData[7] << 24) | (keyData[6] << 16) | (keyData[5] << 8) | keyData[4]);
			keys[2] = (uint)((keyData[11] << 24) | (keyData[10] << 16) | (keyData[9] << 8) | keyData[8]);
		}

		/// <summary>
		/// Update encryption keys
		/// </summary>
		protected void UpdateKeys(byte ch)
		{
			keys[0] = Crc32.ComputeCrc32(keys[0], ch);
			keys[1] = keys[1] + (byte)keys[0];
			keys[1] = keys[1] * 134775813 + 1;
			keys[2] = Crc32.ComputeCrc32(keys[2], (byte)(keys[1] >> 24));
		}

		/// <summary>
		/// Reset the internal state.
		/// </summary>
		protected void Reset()
		{
			keys[0] = 0;
			keys[1] = 0;
			keys[2] = 0;
		}

		#region Instance Fields

		private uint[] keys;

		#endregion Instance Fields
	}

	/// <summary>
	/// PkzipClassic CryptoTransform for encryption.
	/// </summary>
	internal class PkzipClassicEncryptCryptoTransform : PkzipClassicCryptoBase, ICryptoTransform
	{
		/// <summary>
		/// Initialise a new instance of <see cref="PkzipClassicEncryptCryptoTransform"></see>
		/// </summary>
		/// <param name="keyBlock">The key block to use.</param>
		internal PkzipClassicEncryptCryptoTransform(byte[] keyBlock)
		{
			SetKeys(keyBlock);
		}

		#region ICryptoTransform Members

		/// <summary>
		/// Transforms the specified region of the specified byte array.
		/// </summary>
		/// <param name="inputBuffer">The input for which to compute the transform.</param>
		/// <param name="inputOffset">The offset into the byte array from which to begin using data.</param>
		/// <param name="inputCount">The number of bytes in the byte array to use as data.</param>
		/// <returns>The computed transform.</returns>
		public byte[] TransformFinalBlock(byte[] inputBuffer, int inputOffset, int inputCount)
		{
			byte[] result = new byte[inputCount];
			TransformBlock(inputBuffer, inputOffset, inputCount, result, 0);
			return result;
		}

		/// <summary>
		/// Transforms the specified region of the input byte array and copies
		/// the resulting transform to the specified region of the output byte array.
		/// </summary>
		/// <param name="inputBuffer">The input for which to compute the transform.</param>
		/// <param name="inputOffset">The offset into the input byte array from which to begin using data.</param>
		/// <param name="inputCount">The number of bytes in the input byte array to use as data.</param>
		/// <param name="outputBuffer">The output to which to write the transform.</param>
		/// <param name="outputOffset">The offset into the output byte array from which to begin writing data.</param>
		/// <returns>The number of bytes written.</returns>
		public int TransformBlock(byte[] inputBuffer, int inputOffset, int inputCount, byte[] outputBuffer, int outputOffset)
		{
			for (int i = inputOffset; i < inputOffset + inputCount; ++i)
			{
				byte oldbyte = inputBuffer[i];
				outputBuffer[outputOffset++] = (byte)(inputBuffer[i] ^ TransformByte());
				UpdateKeys(oldbyte);
			}
			return inputCount;
		}

		/// <summary>
		/// Gets a value indicating whether the current transform can be reused.
		/// </summary>
		public bool CanReuseTransform
		{
			get
			{
				return true;
			}
		}

		/// <summary>
		/// Gets the size of the input data blocks in bytes.
		/// </summary>
		public int InputBlockSize
		{
			get
			{
				return 1;
			}
		}

		/// <summary>
		/// Gets the size of the output data blocks in bytes.
		/// </summary>
		public int OutputBlockSize
		{
			get
			{
				return 1;
			}
		}

		/// <summary>
		/// Gets a value indicating whether multiple blocks can be transformed.
		/// </summary>
		public bool CanTransformMultipleBlocks
		{
			get
			{
				return true;
			}
		}

		#endregion ICryptoTransform Members

		#region IDisposable Members

		/// <summary>
		/// Cleanup internal state.
		/// </summary>
		public void Dispose()
		{
			Reset();
		}

		#endregion IDisposable Members
	}

	/// <summary>
	/// PkzipClassic CryptoTransform for decryption.
	/// </summary>
	internal class PkzipClassicDecryptCryptoTransform : PkzipClassicCryptoBase, ICryptoTransform
	{
		/// <summary>
		/// Initialise a new instance of <see cref="PkzipClassicDecryptCryptoTransform"></see>.
		/// </summary>
		/// <param name="keyBlock">The key block to decrypt with.</param>
		internal PkzipClassicDecryptCryptoTransform(byte[] keyBlock)
		{
			SetKeys(keyBlock);
		}

		#region ICryptoTransform Members

		/// <summary>
		/// Transforms the specified region of the specified byte array.
		/// </summary>
		/// <param name="inputBuffer">The input for which to compute the transform.</param>
		/// <param name="inputOffset">The offset into the byte array from which to begin using data.</param>
		/// <param name="inputCount">The number of bytes in the byte array to use as data.</param>
		/// <returns>The computed transform.</returns>
		public byte[] TransformFinalBlock(byte[] inputBuffer, int inputOffset, int inputCount)
		{
			byte[] result = new byte[inputCount];
			TransformBlock(inputBuffer, inputOffset, inputCount, result, 0);
			return result;
		}

		/// <summary>
		/// Transforms the specified region of the input byte array and copies
		/// the resulting transform to the specified region of the output byte array.
		/// </summary>
		/// <param name="inputBuffer">The input for which to compute the transform.</param>
		/// <param name="inputOffset">The offset into the input byte array from which to begin using data.</param>
		/// <param name="inputCount">The number of bytes in the input byte array to use as data.</param>
		/// <param name="outputBuffer">The output to which to write the transform.</param>
		/// <param name="outputOffset">The offset into the output byte array from which to begin writing data.</param>
		/// <returns>The number of bytes written.</returns>
		public int TransformBlock(byte[] inputBuffer, int inputOffset, int inputCount, byte[] outputBuffer, int outputOffset)
		{
			for (int i = inputOffset; i < inputOffset + inputCount; ++i)
			{
				var newByte = (byte)(inputBuffer[i] ^ TransformByte());
				outputBuffer[outputOffset++] = newByte;
				UpdateKeys(newByte);
			}
			return inputCount;
		}

		/// <summary>
		/// Gets a value indicating whether the current transform can be reused.
		/// </summary>
		public bool CanReuseTransform
		{
			get
			{
				return true;
			}
		}

		/// <summary>
		/// Gets the size of the input data blocks in bytes.
		/// </summary>
		public int InputBlockSize
		{
			get
			{
				return 1;
			}
		}

		/// <summary>
		/// Gets the size of the output data blocks in bytes.
		/// </summary>
		public int OutputBlockSize
		{
			get
			{
				return 1;
			}
		}

		/// <summary>
		/// Gets a value indicating whether multiple blocks can be transformed.
		/// </summary>
		public bool CanTransformMultipleBlocks
		{
			get
			{
				return true;
			}
		}

		#endregion ICryptoTransform Members

		#region IDisposable Members

		/// <summary>
		/// Cleanup internal state.
		/// </summary>
		public void Dispose()
		{
			Reset();
		}

		#endregion IDisposable Members
	}

	/// <summary>
	/// Defines a wrapper object to access the Pkzip algorithm.
	/// This class cannot be inherited.
	/// </summary>
	public sealed class PkzipClassicManaged : PkzipClassic
	{
		/// <summary>
		/// Get / set the applicable block size in bits.
		/// </summary>
		/// <remarks>The only valid block size is 8.</remarks>
		public override int BlockSize
		{
			get
			{
				return 8;
			}

			set
			{
				if (value != 8)
				{
					throw new CryptographicException("Block size is invalid");
				}
			}
		}

		/// <summary>
		/// Get an array of legal <see cref="KeySizes">key sizes.</see>
		/// </summary>
		public override KeySizes[] LegalKeySizes
		{
			get
			{
				KeySizes[] keySizes = new KeySizes[1];
				keySizes[0] = new KeySizes(12 * 8, 12 * 8, 0);
				return keySizes;
			}
		}

		/// <summary>
		/// Generate an initial vector.
		/// </summary>
		public override void GenerateIV()
		{
			// Do nothing.
		}

		/// <summary>
		/// Get an array of legal <see cref="KeySizes">block sizes</see>.
		/// </summary>
		public override KeySizes[] LegalBlockSizes
		{
			get
			{
				KeySizes[] keySizes = new KeySizes[1];
				keySizes[0] = new KeySizes(1 * 8, 1 * 8, 0);
				return keySizes;
			}
		}

		/// <summary>
		/// Get / set the key value applicable.
		/// </summary>
		public override byte[] Key
		{
			get
			{
				if (key_ == null)
				{
					GenerateKey();
				}

				return (byte[])key_.Clone();
			}

			set
			{
				if (value == null)
				{
					throw new ArgumentNullException(nameof(value));
				}

				if (value.Length != 12)
				{
					throw new CryptographicException("Key size is illegal");
				}

				key_ = (byte[])value.Clone();
			}
		}

		/// <summary>
		/// Generate a new random key.
		/// </summary>
		public override void GenerateKey()
		{
			key_ = new byte[12];
			using (var rng = new RNGCryptoServiceProvider())
			{
				rng.GetBytes(key_);
			}
		}

		/// <summary>
		/// Create an encryptor.
		/// </summary>
		/// <param name="rgbKey">The key to use for this encryptor.</param>
		/// <param name="rgbIV">Initialisation vector for the new encryptor.</param>
		/// <returns>Returns a new PkzipClassic encryptor</returns>
		public override ICryptoTransform CreateEncryptor(
			byte[] rgbKey,
			byte[] rgbIV)
		{
			key_ = rgbKey;
			return new PkzipClassicEncryptCryptoTransform(Key);
		}

		/// <summary>
		/// Create a decryptor.
		/// </summary>
		/// <param name="rgbKey">Keys to use for this new decryptor.</param>
		/// <param name="rgbIV">Initialisation vector for the new decryptor.</param>
		/// <returns>Returns a new decryptor.</returns>
		public override ICryptoTransform CreateDecryptor(
			byte[] rgbKey,
			byte[] rgbIV)
		{
			key_ = rgbKey;
			return new PkzipClassicDecryptCryptoTransform(Key);
		}

		#region Instance Fields

		private byte[] key_;

		#endregion Instance Fields
	}
}

namespace ICSharpCode.SharpZipLib.Encryption
{
	/// <summary>
	/// Encrypts and decrypts AES ZIP
	/// </summary>
	/// <remarks>
	/// Based on information from http://www.winzip.com/aes_info.htm
	/// and http://www.gladman.me.uk/cryptography_technology/fileencrypt/
	/// </remarks>
	internal class ZipAESStream : CryptoStream
	{
		/// <summary>
		/// Constructor
		/// </summary>
		/// <param name="stream">The stream on which to perform the cryptographic transformation.</param>
		/// <param name="transform">Instance of ZipAESTransform</param>
		/// <param name="mode">Read or Write</param>
		public ZipAESStream(Stream stream, ZipAESTransform transform, CryptoStreamMode mode)
			: base(stream, transform, mode)
		{
			_stream = stream;
			_transform = transform;
			_slideBuffer = new byte[1024];

			// mode:
			//  CryptoStreamMode.Read means we read from "stream" and pass decrypted to our Read() method.
			//  Write bypasses this stream and uses the Transform directly.
			if (mode != CryptoStreamMode.Read)
			{
				throw new Exception("ZipAESStream only for read");
			}
		}

		// The final n bytes of the AES stream contain the Auth Code.
		private const int AUTH_CODE_LENGTH = 10;

		// Blocksize is always 16 here, even for AES-256 which has transform.InputBlockSize of 32.
		private const int CRYPTO_BLOCK_SIZE = 16;

		// total length of block + auth code
		private const int BLOCK_AND_AUTH = CRYPTO_BLOCK_SIZE + AUTH_CODE_LENGTH;

		private Stream _stream;
		private ZipAESTransform _transform;
		private byte[] _slideBuffer;
		private int _slideBufStartPos;
		private int _slideBufFreePos;

		// Buffer block transforms to enable partial reads
		private byte[] _transformBuffer = null;// new byte[CRYPTO_BLOCK_SIZE];
		private int _transformBufferFreePos;
		private int _transformBufferStartPos;

		// Do we have some buffered data available?
		private bool HasBufferedData =>_transformBuffer != null && _transformBufferStartPos < _transformBufferFreePos;

		/// <summary>
		/// Reads a sequence of bytes from the current CryptoStream into buffer,
		/// and advances the position within the stream by the number of bytes read.
		/// </summary>
		public override int Read(byte[] buffer, int offset, int count)
		{
			// Nothing to do
			if (count == 0)
				return 0;

			// If we have buffered data, read that first
			int nBytes = 0;
			if (HasBufferedData)
			{
				nBytes = ReadBufferedData(buffer, offset, count);

				// Read all requested data from the buffer
				if (nBytes == count)
					return nBytes;

				offset += nBytes;
				count -= nBytes;
			}

			// Read more data from the input, if available
			if (_slideBuffer != null)
				nBytes += ReadAndTransform(buffer, offset, count);

			return nBytes;
		}

		/// <inheritdoc/>
		public override Task<int> ReadAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken)
		{
			var readCount = Read(buffer, offset, count);
			return Task.FromResult(readCount);
		}

		// Read data from the underlying stream and decrypt it
		private int ReadAndTransform(byte[] buffer, int offset, int count)
		{
			int nBytes = 0;
			while (nBytes < count)
			{
				int bytesLeftToRead = count - nBytes;

				// Calculate buffer quantities vs read-ahead size, and check for sufficient free space
				int byteCount = _slideBufFreePos - _slideBufStartPos;

				// Need to handle final block and Auth Code specially, but don't know total data length.
				// Maintain a read-ahead equal to the length of (crypto block + Auth Code).
				// When that runs out we can detect these final sections.
				int lengthToRead = BLOCK_AND_AUTH - byteCount;
				if (_slideBuffer.Length - _slideBufFreePos < lengthToRead)
				{
					// Shift the data to the beginning of the buffer
					int iTo = 0;
					for (int iFrom = _slideBufStartPos; iFrom < _slideBufFreePos; iFrom++, iTo++)
					{
						_slideBuffer[iTo] = _slideBuffer[iFrom];
					}
					_slideBufFreePos -= _slideBufStartPos;      // Note the -=
					_slideBufStartPos = 0;
				}
				int obtained = StreamUtils.ReadRequestedBytes(_stream, _slideBuffer, _slideBufFreePos, lengthToRead);
				_slideBufFreePos += obtained;

				// Recalculate how much data we now have
				byteCount = _slideBufFreePos - _slideBufStartPos;
				if (byteCount >= BLOCK_AND_AUTH)
				{
					var read = TransformAndBufferBlock(buffer, offset, bytesLeftToRead, CRYPTO_BLOCK_SIZE);
					nBytes += read;
					offset += read;
				}
				else
				{
					// Last round.
					if (byteCount > AUTH_CODE_LENGTH)
					{
						// At least one byte of data plus auth code
						int finalBlock = byteCount - AUTH_CODE_LENGTH;
						nBytes += TransformAndBufferBlock(buffer, offset, bytesLeftToRead, finalBlock);
					}
					else if (byteCount < AUTH_CODE_LENGTH)
						throw new ZipException("Internal error missed auth code"); // Coding bug
																				// Final block done. Check Auth code.
					byte[] calcAuthCode = _transform.GetAuthCode();
					for (int i = 0; i < AUTH_CODE_LENGTH; i++)
					{
						if (calcAuthCode[i] != _slideBuffer[_slideBufStartPos + i])
						{
							throw new ZipException("AES Authentication Code does not match. This is a super-CRC check on the data in the file after compression and encryption. \r\n"
								+ "The file may be damaged.");
						}
					}

					// don't need this any more, so use it as a 'complete' flag
					_slideBuffer = null;

					break;  // Reached the auth code
				}
			}
			return nBytes;
		}

		// read some buffered data
		private int ReadBufferedData(byte[] buffer, int offset, int count)
		{
			int copyCount = Math.Min(count, _transformBufferFreePos - _transformBufferStartPos);

			Array.Copy(_transformBuffer, _transformBufferStartPos, buffer, offset, copyCount);
			_transformBufferStartPos += copyCount;

			return copyCount;
		}

		// Perform the crypto transform, and buffer the data if less than one block has been requested.
		private int TransformAndBufferBlock(byte[] buffer, int offset, int count, int blockSize)
		{
			// If the requested data is greater than one block, transform it directly into the output
			// If it's smaller, do it into a temporary buffer and copy the requested part
			bool bufferRequired = (blockSize > count);

			if (bufferRequired && _transformBuffer == null)
				_transformBuffer = new byte[CRYPTO_BLOCK_SIZE];

			var targetBuffer = bufferRequired ? _transformBuffer : buffer;
			var targetOffset = bufferRequired ? 0 : offset;

			// Transform the data
			_transform.TransformBlock(_slideBuffer,
									  _slideBufStartPos,
									  blockSize,
									  targetBuffer,
									  targetOffset);

			_slideBufStartPos += blockSize;

			if (!bufferRequired)
			{
				return blockSize;
			}
			else
			{
				Array.Copy(_transformBuffer, 0, buffer, offset, count);
				_transformBufferStartPos = count;
				_transformBufferFreePos = blockSize;

				return count;
			}
		}

		/// <summary>
		/// Writes a sequence of bytes to the current stream and advances the current position within this stream by the number of bytes written.
		/// </summary>
		/// <param name="buffer">An array of bytes. This method copies count bytes from buffer to the current stream. </param>
		/// <param name="offset">The byte offset in buffer at which to begin copying bytes to the current stream. </param>
		/// <param name="count">The number of bytes to be written to the current stream. </param>
		public override void Write(byte[] buffer, int offset, int count)
		{
			// ZipAESStream is used for reading but not for writing. Writing uses the ZipAESTransform directly.
			throw new NotImplementedException();
		}
	}
}
namespace ICSharpCode.SharpZipLib.Encryption
{
	/// <summary>
	/// Transforms stream using AES in CTR mode
	/// </summary>
	internal class ZipAESTransform : ICryptoTransform
	{
#if NET45
		class IncrementalHash : HMACSHA1
		{
			bool _finalised;
			public IncrementalHash(byte[] key) : base(key) { }
			public static IncrementalHash CreateHMAC(string n, byte[] key) => new IncrementalHash(key);
			public void AppendData(byte[] buffer, int offset, int count) => TransformBlock(buffer, offset, count, buffer, offset);
			public byte[] GetHashAndReset()
			{
				if (!_finalised)
				{
					byte[] dummy = new byte[0];
					TransformFinalBlock(dummy, 0, 0);
					_finalised = true;
				}
				return Hash;
			}
		}

		static class HashAlgorithmName
		{
			public static string SHA1 = null;
		}
#endif

		private const int PWD_VER_LENGTH = 2;

		// WinZip use iteration count of 1000 for PBKDF2 key generation
		private const int KEY_ROUNDS = 1000;

		// For 128-bit AES (16 bytes) the encryption is implemented as expected.
		// For 256-bit AES (32 bytes) WinZip do full 256 bit AES of the nonce to create the encryption
		// block but use only the first 16 bytes of it, and discard the second half.
		private const int ENCRYPT_BLOCK = 16;

		private int _blockSize;
		private readonly ICryptoTransform _encryptor;
		private readonly byte[] _counterNonce;
		private byte[] _encryptBuffer;
		private int _encrPos;
		private byte[] _pwdVerifier;
		private IncrementalHash _hmacsha1;
		private byte[] _authCode = null;

		private bool _writeMode;

		/// <summary>
		/// Constructor.
		/// </summary>
		/// <param name="key">Password string</param>
		/// <param name="saltBytes">Random bytes, length depends on encryption strength.
		/// 128 bits = 8 bytes, 192 bits = 12 bytes, 256 bits = 16 bytes.</param>
		/// <param name="blockSize">The encryption strength, in bytes eg 16 for 128 bits.</param>
		/// <param name="writeMode">True when creating a zip, false when reading. For the AuthCode.</param>
		///
		public ZipAESTransform(string key, byte[] saltBytes, int blockSize, bool writeMode)
		{
			if (blockSize != 16 && blockSize != 32) // 24 valid for AES but not supported by Winzip
				throw new Exception("Invalid blocksize " + blockSize + ". Must be 16 or 32.");
			if (saltBytes.Length != blockSize / 2)
				throw new Exception("Invalid salt len. Must be " + blockSize / 2 + " for blocksize " + blockSize);
			// initialise the encryption buffer and buffer pos
			_blockSize = blockSize;
			_encryptBuffer = new byte[_blockSize];
			_encrPos = ENCRYPT_BLOCK;

			// Performs the equivalent of derive_key in Dr Brian Gladman's pwd2key.c
			var pdb = new Rfc2898DeriveBytes(key, saltBytes, KEY_ROUNDS);
			var rm = Aes.Create();
			rm.Mode = CipherMode.ECB;           // No feedback from cipher for CTR mode
			_counterNonce = new byte[_blockSize];
			byte[] key1bytes = pdb.GetBytes(_blockSize);
			byte[] key2bytes = pdb.GetBytes(_blockSize);

			// Use empty IV for AES
			_encryptor = rm.CreateEncryptor(key1bytes, new byte[16]);
			_pwdVerifier = pdb.GetBytes(PWD_VER_LENGTH);
			//
			_hmacsha1 = IncrementalHash.CreateHMAC(HashAlgorithmName.SHA1, key2bytes);
			_writeMode = writeMode;
		}

		/// <summary>
		/// Implement the ICryptoTransform method.
		/// </summary>
		public int TransformBlock(byte[] inputBuffer, int inputOffset, int inputCount, byte[] outputBuffer, int outputOffset)
		{
			// Pass the data stream to the hash algorithm for generating the Auth Code.
			// This does not change the inputBuffer. Do this before decryption for read mode.
			if (!_writeMode)
			{
				_hmacsha1.AppendData(inputBuffer, inputOffset, inputCount);
			}
			// Encrypt with AES in CTR mode. Regards to Dr Brian Gladman for this.
			int ix = 0;
			while (ix < inputCount)
			{
				if (_encrPos == ENCRYPT_BLOCK)
				{
					/* increment encryption nonce   */
					int j = 0;
					while (++_counterNonce[j] == 0)
					{
						++j;
					}
					/* encrypt the nonce to form next xor buffer    */
					_encryptor.TransformBlock(_counterNonce, 0, _blockSize, _encryptBuffer, 0);
					_encrPos = 0;
				}
				outputBuffer[ix + outputOffset] = (byte)(inputBuffer[ix + inputOffset] ^ _encryptBuffer[_encrPos++]);
				//
				ix++;
			}
			if (_writeMode)
			{
				// This does not change the buffer.
				_hmacsha1.AppendData(outputBuffer, outputOffset, inputCount);
			}
			return inputCount;
		}

		/// <summary>
		/// Returns the 2 byte password verifier
		/// </summary>
		public byte[] PwdVerifier
		{
			get
			{
				return _pwdVerifier;
			}
		}

		/// <summary>
		/// Returns the 10 byte AUTH CODE to be checked or appended immediately following the AES data stream.
		/// </summary>
		public byte[] GetAuthCode()
		{
			if (_authCode == null)
			{
				_authCode = _hmacsha1.GetHashAndReset();
			}
			return _authCode;
		}

		#region ICryptoTransform Members

		/// <summary>
		/// Not implemented.
		/// </summary>
		public byte[] TransformFinalBlock(byte[] inputBuffer, int inputOffset, int inputCount)
		{
			if(inputCount > 0)
			{
				throw new NotImplementedException("TransformFinalBlock is not implemented and inputCount is greater than 0");
			}
			return Empty.Array<byte>();
		}

		/// <summary>
		/// Gets the size of the input data blocks in bytes.
		/// </summary>
		public int InputBlockSize
		{
			get
			{
				return _blockSize;
			}
		}

		/// <summary>
		/// Gets the size of the output data blocks in bytes.
		/// </summary>
		public int OutputBlockSize
		{
			get
			{
				return _blockSize;
			}
		}

		/// <summary>
		/// Gets a value indicating whether multiple blocks can be transformed.
		/// </summary>
		public bool CanTransformMultipleBlocks
		{
			get
			{
				return true;
			}
		}

		/// <summary>
		/// Gets a value indicating whether the current transform can be reused.
		/// </summary>
		public bool CanReuseTransform
		{
			get
			{
				return true;
			}
		}

		/// <summary>
		/// Cleanup internal state.
		/// </summary>
		public void Dispose()
		{
			_encryptor.Dispose();
		}

		#endregion ICryptoTransform Members
	}
}


namespace ICSharpCode.SharpZipLib.GZip
{

	/// <summary>
	/// An example class to demonstrate compression and decompression of GZip streams.
	/// </summary>
	public static class GZip
	{
		/// <summary>
		/// Decompress the <paramref name="inStream">input</paramref> writing
		/// uncompressed data to the <paramref name="outStream">output stream</paramref>
		/// </summary>
		/// <param name="inStream">The readable stream containing data to decompress.</param>
		/// <param name="outStream">The output stream to receive the decompressed data.</param>
		/// <param name="isStreamOwner">Both streams are closed on completion if true.</param>
		/// <exception cref="ArgumentNullException">Input or output stream is null</exception>
		public static void Decompress(Stream inStream, Stream outStream, bool isStreamOwner)
		{
			if (inStream == null)
				throw new ArgumentNullException(nameof(inStream), "Input stream is null");

			if (outStream == null)
				throw new ArgumentNullException(nameof(outStream), "Output stream is null");

			try
			{
				using (GZipInputStream gzipInput = new GZipInputStream(inStream))
				{
					gzipInput.IsStreamOwner = isStreamOwner;
					Core.StreamUtils.Copy(gzipInput, outStream, new byte[4096]);
				}
			}
			finally
			{
				if (isStreamOwner)
				{
					// inStream is closed by the GZipInputStream if stream owner
					outStream.Dispose();
				}
			}
		}

		/// <summary>
		/// Compress the <paramref name="inStream">input stream</paramref> sending
		/// result data to <paramref name="outStream">output stream</paramref>
		/// </summary>
		/// <param name="inStream">The readable stream to compress.</param>
		/// <param name="outStream">The output stream to receive the compressed data.</param>
		/// <param name="isStreamOwner">Both streams are closed on completion if true.</param>
		/// <param name="bufferSize">Deflate buffer size, minimum 512</param>
		/// <param name="level">Deflate compression level, 0-9</param>
		/// <exception cref="ArgumentNullException">Input or output stream is null</exception>
		/// <exception cref="ArgumentOutOfRangeException">Buffer Size is smaller than 512</exception>
		/// <exception cref="ArgumentOutOfRangeException">Compression level outside 0-9</exception>
		public static void Compress(Stream inStream, Stream outStream, bool isStreamOwner, int bufferSize = 512, int level = 6)
		{
			if (inStream == null)
				throw new ArgumentNullException(nameof(inStream), "Input stream is null");

			if (outStream == null)
				throw new ArgumentNullException(nameof(outStream), "Output stream is null");

			if (bufferSize < 512)
				throw new ArgumentOutOfRangeException(nameof(bufferSize), "Deflate buffer size must be >= 512");

			if (level < NO_COMPRESSION || level > BEST_COMPRESSION)
				throw new ArgumentOutOfRangeException(nameof(level), "Compression level must be 0-9");

			try
			{
				using (GZipOutputStream gzipOutput = new GZipOutputStream(outStream, bufferSize))
				{
					gzipOutput.SetLevel(level);
					gzipOutput.IsStreamOwner = isStreamOwner;
					Core.StreamUtils.Copy(inStream, gzipOutput, new byte[bufferSize]);
				}
			}
			finally
			{
				if (isStreamOwner)
				{
					// outStream is closed by the GZipOutputStream if stream owner
					inStream.Dispose();
				}
			}
		}
	}
}


namespace ICSharpCode.SharpZipLib.GZip
{
	/// <summary>
	/// This class contains constants used for gzip.
	/// </summary>
	[System.Diagnostics.CodeAnalysis.SuppressMessage("Naming", "CA1707:Identifiers should not contain underscores", Justification = "kept for backwards compatibility")]
	sealed public class GZipConstants
	{
		/// <summary>
		/// First GZip identification byte
		/// </summary>
		public const byte ID1 = 0x1F;

		/// <summary>
		/// Second GZip identification byte
		/// </summary>
		public const byte ID2 = 0x8B;

		/// <summary>
		/// Deflate compression method
		/// </summary>
		public const byte CompressionMethodDeflate = 0x8;

		/// <summary>
		/// Get the GZip specified encoding (CP-1252 if supported, otherwise ASCII)
		/// </summary>
		public static Encoding Encoding
		{
			get
			{
				try
				{
					return Encoding.GetEncoding(1252);
				}
				catch
				{
					return Encoding.ASCII;
				}
			}
		}

	}

	/// <summary>
	/// GZip header flags
	/// </summary>
	[Flags]
	public enum GZipFlags: byte
	{
		/// <summary>
		/// Text flag hinting that the file is in ASCII
		/// </summary>
		FTEXT = 0x1 << 0,

		/// <summary>
		/// CRC flag indicating that a CRC16 preceeds the data
		/// </summary>
		FHCRC = 0x1 << 1,

		/// <summary>
		/// Extra flag indicating that extra fields are present
		/// </summary>
		FEXTRA = 0x1 << 2,

		/// <summary>
		/// Filename flag indicating that the original filename is present
		/// </summary>
		FNAME = 0x1 << 3,

		/// <summary>
		/// Flag bit mask indicating that a comment is present
		/// </summary>
		FCOMMENT = 0x1 << 4,
	}
}


namespace ICSharpCode.SharpZipLib.GZip
{
	/// <summary>
	/// GZipException represents exceptions specific to GZip classes and code.
	/// </summary>
	[Serializable]
	public class GZipException : SharpZipBaseException
	{
		/// <summary>
		/// Initialise a new instance of <see cref="GZipException" />.
		/// </summary>
		public GZipException()
		{
		}

		/// <summary>
		/// Initialise a new instance of <see cref="GZipException" /> with its message string.
		/// </summary>
		/// <param name="message">A <see cref="string"/> that describes the error.</param>
		public GZipException(string message)
			: base(message)
		{
		}

		/// <summary>
		/// Initialise a new instance of <see cref="GZipException" />.
		/// </summary>
		/// <param name="message">A <see cref="string"/> that describes the error.</param>
		/// <param name="innerException">The <see cref="Exception"/> that caused this exception.</param>
		public GZipException(string message, Exception innerException)
			: base(message, innerException)
		{
		}

		/// <summary>
		/// Initializes a new instance of the GZipException class with serialized data.
		/// </summary>
		/// <param name="info">
		/// The System.Runtime.Serialization.SerializationInfo that holds the serialized
		/// object data about the exception being thrown.
		/// </param>
		/// <param name="context">
		/// The System.Runtime.Serialization.StreamingContext that contains contextual information
		/// about the source or destination.
		/// </param>
		protected GZipException(SerializationInfo info, StreamingContext context)
			: base(info, context)
		{
		}
	}
}


namespace ICSharpCode.SharpZipLib.GZip
{
	/// <summary>
	/// This filter stream is used to decompress a "GZIP" format stream.
	/// The "GZIP" format is described baseInputStream RFC 1952.
	///
	/// author of the original java version : John Leuner
	/// </summary>
	/// <example> This sample shows how to unzip a gzipped file
	/// <code>
	///
	/// using ICSharpCode.SharpZipLib.Core;
	/// using ICSharpCode.SharpZipLib.GZip;
	///
	/// class MainClass
	/// {
	/// 	public static void Main(string[] args)
	/// 	{
	///			using (Stream inStream = new GZipInputStream(File.OpenRead(args[0])))
	///			using (FileStream outStream = File.Create(Path.GetFileNameWithoutExtension(args[0]))) {
	///				byte[] buffer = new byte[4096];
	///				StreamUtils.Copy(inStream, outStream, buffer);
	/// 		}
	/// 	}
	/// }
	/// </code>
	/// </example>
	public class GZipInputStream : InflaterInputStream
	{
		#region Instance Fields

		/// <summary>
		/// CRC-32 value for uncompressed data
		/// </summary>
		protected Crc32 crc;

		/// <summary>
		/// Flag to indicate if we've read the GZIP header yet for the current member (block of compressed data).
		/// This is tracked per-block as the file is parsed.
		/// </summary>
		private bool readGZIPHeader;

		/// <summary>
		/// Flag to indicate if at least one block in a stream with concatenated blocks was read successfully.
		/// This allows us to exit gracefully if downstream data is not in gzip format.
		/// </summary>
		private bool completedLastBlock;

		private string fileName;

		#endregion Instance Fields

		#region Constructors

		/// <summary>
		/// Creates a GZipInputStream with the default buffer size
		/// </summary>
		/// <param name="baseInputStream">
		/// The stream to read compressed data from (baseInputStream GZIP format)
		/// </param>
		public GZipInputStream(Stream baseInputStream)
			: this(baseInputStream, 4096)
		{
		}

		/// <summary>
		/// Creates a GZIPInputStream with the specified buffer size
		/// </summary>
		/// <param name="baseInputStream">
		/// The stream to read compressed data from (baseInputStream GZIP format)
		/// </param>
		/// <param name="size">
		/// Size of the buffer to use
		/// </param>
		public GZipInputStream(Stream baseInputStream, int size)
			: base(baseInputStream, new Inflater(true), size)
		{
		}

		#endregion Constructors

		#region Stream overrides

		/// <summary>
		/// Reads uncompressed data into an array of bytes
		/// </summary>
		/// <param name="buffer">
		/// The buffer to read uncompressed data into
		/// </param>
		/// <param name="offset">
		/// The offset indicating where the data should be placed
		/// </param>
		/// <param name="count">
		/// The number of uncompressed bytes to be read
		/// </param>
		/// <returns>Returns the number of bytes actually read.</returns>
		public override int Read(byte[] buffer, int offset, int count)
		{
			// A GZIP file can contain multiple blocks of compressed data, although this is quite rare.
			// A compressed block could potentially be empty, so we need to loop until we reach EOF or
			// we find data.
			while (true)
			{
				// If we haven't read the header for this block, read it
				if (!readGZIPHeader)
				{
					// Try to read header. If there is no header (0 bytes available), this is EOF. If there is
					// an incomplete header, this will throw an exception.
					try
					{
						if (!ReadHeader())
						{
							return 0;
						}
					}
					catch (Exception ex) when (completedLastBlock && (ex is GZipException || ex is EndOfStreamException))
					{
						// if we completed the last block (i.e. we're in a stream that has multiple blocks concatenated
						// we want to return gracefully from any header parsing exceptions since sometimes there may
						// be trailing garbage on a stream
						return 0;
					}
				}

				// Try to read compressed data
				int bytesRead = base.Read(buffer, offset, count);
				if (bytesRead > 0)
				{
					crc.Update(new ArraySegment<byte>(buffer, offset, bytesRead));
				}

				// If this is the end of stream, read the footer
				if (inf.IsFinished)
				{
					ReadFooter();
				}

				// Attempting to read 0 bytes will never yield any bytesRead, so we return instead of looping forever
				if (bytesRead > 0 || count == 0)
				{
					return bytesRead;
				}
			}
		}

		/// <summary>
		/// Retrieves the filename header field for the block last read
		/// </summary>
		/// <returns></returns>
		public string GetFilename()
		{
			return fileName;
		}

		#endregion Stream overrides

		#region Support routines

		private bool ReadHeader()
		{
			// Initialize CRC for this block
			crc = new Crc32();

			// Make sure there is data in file. We can't rely on ReadLeByte() to fill the buffer, as this could be EOF,
			// which is fine, but ReadLeByte() throws an exception if it doesn't find data, so we do this part ourselves.
			if (inputBuffer.Available <= 0)
			{
				inputBuffer.Fill();
				if (inputBuffer.Available <= 0)
				{
					// No header, EOF.
					return false;
				}
			}

			var headCRC = new Crc32();

			// 1. Check the two magic bytes

			var magic = inputBuffer.ReadLeByte();
			headCRC.Update(magic);
			if (magic != GZipConstants.ID1)
			{
				throw new GZipException("Error GZIP header, first magic byte doesn't match");
			}

			magic = inputBuffer.ReadLeByte();
			if (magic != GZipConstants.ID2)
			{
				throw new GZipException("Error GZIP header,  second magic byte doesn't match");
			}
			headCRC.Update(magic);

			// 2. Check the compression type (must be 8)
			var compressionType = inputBuffer.ReadLeByte();

			if (compressionType != GZipConstants.CompressionMethodDeflate)
			{
				throw new GZipException("Error GZIP header, data not in deflate format");
			}
			headCRC.Update(compressionType);

			// 3. Check the flags
			var flagsByte = inputBuffer.ReadLeByte();

			headCRC.Update(flagsByte);

			// 3.1 Check the reserved bits are zero

			if ((flagsByte & 0xE0) != 0)
			{
				throw new GZipException("Reserved flag bits in GZIP header != 0");
			}

			var flags = (GZipFlags)flagsByte;

			// 4.-6. Skip the modification time, extra flags, and OS type
			for (int i = 0; i < 6; i++)
			{
				headCRC.Update(inputBuffer.ReadLeByte());
			}

			// 7. Read extra field
			if (flags.HasFlag(GZipFlags.FEXTRA))
			{
				// XLEN is total length of extra subfields, we will skip them all
				var len1 = inputBuffer.ReadLeByte();
				var len2 = inputBuffer.ReadLeByte();

				headCRC.Update(len1);
				headCRC.Update(len2);

				int extraLen = (len2 << 8) | len1;      // gzip is LSB first
				for (int i = 0; i < extraLen; i++)
				{
					headCRC.Update(inputBuffer.ReadLeByte());
				}
			}

			// 8. Read file name
			if (flags.HasFlag(GZipFlags.FNAME))
			{
				var fname = new byte[1024];
				var fnamePos = 0;
				int readByte;
				while ((readByte = inputBuffer.ReadLeByte()) > 0)
				{
					if (fnamePos < 1024)
					{
						fname[fnamePos++] = (byte)readByte;
					}
					headCRC.Update(readByte);
				}

				headCRC.Update(readByte);

				fileName = GZipConstants.Encoding.GetString(fname, 0, fnamePos);
			}
			else
			{
				fileName = null;
			}

			// 9. Read comment
			if (flags.HasFlag(GZipFlags.FCOMMENT))
			{
				int readByte;
				while ((readByte = inputBuffer.ReadLeByte()) > 0)
				{
					headCRC.Update(readByte);
				}

				headCRC.Update(readByte);
			}

			// 10. Read header CRC
			if (flags.HasFlag(GZipFlags.FHCRC))
			{
				int tempByte;
				int crcval = inputBuffer.ReadLeByte();
				if (crcval < 0)
				{
					throw new EndOfStreamException("EOS reading GZIP header");
				}

				tempByte = inputBuffer.ReadLeByte();
				if (tempByte < 0)
				{
					throw new EndOfStreamException("EOS reading GZIP header");
				}

				crcval = (crcval << 8) | tempByte;
				if (crcval != ((int)headCRC.Value & 0xffff))
				{
					throw new GZipException("Header CRC value mismatch");
				}
			}

			readGZIPHeader = true;
			return true;
		}

		private void ReadFooter()
		{
			byte[] footer = new byte[8];

			// End of stream; reclaim all bytes from inf, read the final byte count, and reset the inflator
			long bytesRead = inf.TotalOut & 0xffffffff;
			inputBuffer.Available += inf.RemainingInput;
			inf.Reset();

			// Read footer from inputBuffer
			int needed = 8;
			while (needed > 0)
			{
				int count = inputBuffer.ReadClearTextBuffer(footer, 8 - needed, needed);
				if (count <= 0)
				{
					throw new EndOfStreamException("EOS reading GZIP footer");
				}
				needed -= count; // Jewel Jan 16
			}

			// Calculate CRC
			int crcval = (footer[0] & 0xff) | ((footer[1] & 0xff) << 8) | ((footer[2] & 0xff) << 16) | (footer[3] << 24);
			if (crcval != (int)crc.Value)
			{
				throw new GZipException("GZIP crc sum mismatch, theirs \"" + crcval + "\" and ours \"" + (int)crc.Value);
			}

			// NOTE The total here is the original total modulo 2 ^ 32.
			uint total =
				(uint)((uint)footer[4] & 0xff) |
				(uint)(((uint)footer[5] & 0xff) << 8) |
				(uint)(((uint)footer[6] & 0xff) << 16) |
				(uint)((uint)footer[7] << 24);

			if (bytesRead != total)
			{
				throw new GZipException("Number of bytes mismatch in footer");
			}

			// Mark header read as false so if another header exists, we'll continue reading through the file
			readGZIPHeader = false;

			// Indicate that we succeeded on at least one block so we can exit gracefully if there is trailing garbage downstream
			completedLastBlock = true;
		}

		#endregion Support routines
	}
}

namespace ICSharpCode.SharpZipLib.GZip
{
	/// <summary>
	/// This filter stream is used to compress a stream into a "GZIP" stream.
	/// The "GZIP" format is described in RFC 1952.
	///
	/// author of the original java version : John Leuner
	/// </summary>
	/// <example> This sample shows how to gzip a file
	/// <code>
	///
	/// using ICSharpCode.SharpZipLib.GZip;
	/// using ICSharpCode.SharpZipLib.Core;
	///
	/// class MainClass
	/// {
	/// 	public static void Main(string[] args)
	/// 	{
	/// 			using (Stream s = new GZipOutputStream(File.Create(args[0] + ".gz")))
	/// 			using (FileStream fs = File.OpenRead(args[0])) {
	/// 				byte[] writeData = new byte[4096];
	/// 				Streamutils.Copy(s, fs, writeData);
	/// 			}
	/// 		}
	/// 	}
	/// }
	/// </code>
	/// </example>
	public class GZipOutputStream : DeflaterOutputStream
	{
		private enum OutputState
		{
			Header,
			Footer,
			Finished,
			Closed,
		};

		#region Instance Fields

		/// <summary>
		/// CRC-32 value for uncompressed data
		/// </summary>
		protected Crc32 crc = new Crc32();

		private OutputState state_ = OutputState.Header;

		private string fileName;

		private GZipFlags flags = 0;

		#endregion Instance Fields

		#region Constructors

		/// <summary>
		/// Creates a GzipOutputStream with the default buffer size
		/// </summary>
		/// <param name="baseOutputStream">
		/// The stream to read data (to be compressed) from
		/// </param>
		public GZipOutputStream(Stream baseOutputStream)
			: this(baseOutputStream, 4096)
		{
		}

		/// <summary>
		/// Creates a GZipOutputStream with the specified buffer size
		/// </summary>
		/// <param name="baseOutputStream">
		/// The stream to read data (to be compressed) from
		/// </param>
		/// <param name="size">
		/// Size of the buffer to use
		/// </param>
		public GZipOutputStream(Stream baseOutputStream, int size) : base(baseOutputStream, new Deflater(Deflater.DEFAULT_COMPRESSION, true), size)
		{
		}

		#endregion Constructors

		#region Public API

		/// <summary>
		/// Sets the active compression level (0-9).  The new level will be activated
		/// immediately.
		/// </summary>
		/// <param name="level">The compression level to set.</param>
		/// <exception cref="ArgumentOutOfRangeException">
		/// Level specified is not supported.
		/// </exception>
		/// <see cref="Deflater"/>
		public void SetLevel(int level)
		{
			if (level < Deflater.NO_COMPRESSION || level > Deflater.BEST_COMPRESSION)
				throw new ArgumentOutOfRangeException(nameof(level), "Compression level must be 0-9");

			deflater_.SetLevel(level);
		}

		/// <summary>
		/// Get the current compression level.
		/// </summary>
		/// <returns>The current compression level.</returns>
		public int GetLevel()
		{
			return deflater_.GetLevel();
		}

		/// <summary>
		/// Original filename
		/// </summary>
		public string FileName
		{
			get => fileName;
			set
			{
				fileName = CleanFilename(value);
				if (string.IsNullOrEmpty(fileName))
				{
					flags &= ~GZipFlags.FNAME;
				}
				else
				{
					flags |= GZipFlags.FNAME;
				}
			}
		}

		#endregion Public API

		#region Stream overrides

		/// <summary>
		/// Write given buffer to output updating crc
		/// </summary>
		/// <param name="buffer">Buffer to write</param>
		/// <param name="offset">Offset of first byte in buf to write</param>
		/// <param name="count">Number of bytes to write</param>
		public override void Write(byte[] buffer, int offset, int count)
		{
			if (state_ == OutputState.Header)
			{
				WriteHeader();
			}

			if (state_ != OutputState.Footer)
			{
				throw new InvalidOperationException("Write not permitted in current state");
			}

			crc.Update(new ArraySegment<byte>(buffer, offset, count));
			base.Write(buffer, offset, count);
		}

		/// <summary>
		/// Writes remaining compressed output data to the output stream
		/// and closes it.
		/// </summary>
		protected override void Dispose(bool disposing)
		{
			try
			{
				Finish();
			}
			finally
			{
				if (state_ != OutputState.Closed)
				{
					state_ = OutputState.Closed;
					if (IsStreamOwner)
					{
						baseOutputStream_.Dispose();
					}
				}
			}
		}

		/// <summary>
		/// Flushes the stream by ensuring the header is written, and then calling <see cref="DeflaterOutputStream.Flush">Flush</see>
		/// on the deflater.
		/// </summary>
		public override void Flush()
		{
			if (state_ == OutputState.Header)
			{
				WriteHeader();
			}

			base.Flush();
		}

		#endregion Stream overrides

		#region DeflaterOutputStream overrides

		/// <summary>
		/// Finish compression and write any footer information required to stream
		/// </summary>
		public override void Finish()
		{
			// If no data has been written a header should be added.
			if (state_ == OutputState.Header)
			{
				WriteHeader();
			}

			if (state_ == OutputState.Footer)
			{
				state_ = OutputState.Finished;
				base.Finish();

				var totalin = (uint)(deflater_.TotalIn & 0xffffffff);
				var crcval = (uint)(crc.Value & 0xffffffff);

				byte[] gzipFooter;

				unchecked
				{
					gzipFooter = new byte[] {
					(byte) crcval, (byte) (crcval >> 8),
					(byte) (crcval >> 16), (byte) (crcval >> 24),

					(byte) totalin, (byte) (totalin >> 8),
					(byte) (totalin >> 16), (byte) (totalin >> 24)
				};
				}

				baseOutputStream_.Write(gzipFooter, 0, gzipFooter.Length);
			}
		}

		#endregion DeflaterOutputStream overrides

		#region Support Routines

		private static string CleanFilename(string path)
			=> path.Substring(path.LastIndexOf('/') + 1);

		private void WriteHeader()
		{
			if (state_ == OutputState.Header)
			{
				state_ = OutputState.Footer;

				var mod_time = (int)((DateTime.Now.Ticks - new DateTime(1970, 1, 1).Ticks) / 10000000L);  // Ticks give back 100ns intervals
				byte[] gzipHeader = {
					// The two magic bytes
					GZipConstants.ID1, 
					GZipConstants.ID2,

					// The compression type
					GZipConstants.CompressionMethodDeflate,

					// The flags (not set)
					(byte)flags,

					// The modification time
					(byte) mod_time, (byte) (mod_time >> 8),
					(byte) (mod_time >> 16), (byte) (mod_time >> 24),

					// The extra flags
					0,

					// The OS type (unknown)
					255
				};

				baseOutputStream_.Write(gzipHeader, 0, gzipHeader.Length);

				if (flags.HasFlag(GZipFlags.FNAME))
				{
					var fname = GZipConstants.Encoding.GetBytes(fileName);
					baseOutputStream_.Write(fname, 0, fname.Length);

					// End filename string with a \0
					baseOutputStream_.Write(new byte[] { 0 }, 0, 1);
				}
			}
		}

		#endregion Support Routines
	}
}

namespace ICSharpCode.SharpZipLib.Lzw
{
	/// <summary>
	/// This class contains constants used for LZW
	/// </summary>
	[System.Diagnostics.CodeAnalysis.SuppressMessage("Naming", "CA1707:Identifiers should not contain underscores", Justification = "kept for backwards compatibility")]
	sealed public class LzwConstants
	{
		/// <summary>
		/// Magic number found at start of LZW header: 0x1f 0x9d
		/// </summary>
		public const int MAGIC = 0x1f9d;

		/// <summary>
		/// Maximum number of bits per code
		/// </summary>
		public const int MAX_BITS = 16;

		/* 3rd header byte:
         * bit 0..4 Number of compression bits
         * bit 5    Extended header
         * bit 6    Free
         * bit 7    Block mode
         */

		/// <summary>
		/// Mask for 'number of compression bits'
		/// </summary>
		public const int BIT_MASK = 0x1f;

		/// <summary>
		/// Indicates the presence of a fourth header byte
		/// </summary>
		public const int EXTENDED_MASK = 0x20;

		//public const int FREE_MASK      = 0x40;

		/// <summary>
		/// Reserved bits
		/// </summary>
		public const int RESERVED_MASK = 0x60;

		/// <summary>
		/// Block compression: if table is full and compression rate is dropping,
		/// clear the dictionary.
		/// </summary>
		public const int BLOCK_MODE_MASK = 0x80;

		/// <summary>
		/// LZW file header size (in bytes)
		/// </summary>
		public const int HDR_SIZE = 3;

		/// <summary>
		/// Initial number of bits per code
		/// </summary>
		public const int INIT_BITS = 9;

		private LzwConstants()
		{
		}
	}
}


namespace ICSharpCode.SharpZipLib.Lzw
{
	/// <summary>
	/// LzwException represents exceptions specific to LZW classes and code.
	/// </summary>
	[Serializable]
	public class LzwException : SharpZipBaseException
	{
		/// <summary>
		/// Initialise a new instance of <see cref="LzwException" />.
		/// </summary>
		public LzwException()
		{
		}

		/// <summary>
		/// Initialise a new instance of <see cref="LzwException" /> with its message string.
		/// </summary>
		/// <param name="message">A <see cref="string"/> that describes the error.</param>
		public LzwException(string message)
			: base(message)
		{
		}

		/// <summary>
		/// Initialise a new instance of <see cref="LzwException" />.
		/// </summary>
		/// <param name="message">A <see cref="string"/> that describes the error.</param>
		/// <param name="innerException">The <see cref="Exception"/> that caused this exception.</param>
		public LzwException(string message, Exception innerException)
			: base(message, innerException)
		{
		}

		/// <summary>
		/// Initializes a new instance of the LzwException class with serialized data.
		/// </summary>
		/// <param name="info">
		/// The System.Runtime.Serialization.SerializationInfo that holds the serialized
		/// object data about the exception being thrown.
		/// </param>
		/// <param name="context">
		/// The System.Runtime.Serialization.StreamingContext that contains contextual information
		/// about the source or destination.
		/// </param>
		protected LzwException(SerializationInfo info, StreamingContext context)
			: base(info, context)
		{
		}
	}
}


namespace ICSharpCode.SharpZipLib.Lzw
{
	/// <summary>
	/// This filter stream is used to decompress a LZW format stream.
	/// Specifically, a stream that uses the LZC compression method.
	/// This file format is usually associated with the .Z file extension.
	///
	/// See http://en.wikipedia.org/wiki/Compress
	/// See http://wiki.wxwidgets.org/Development:_Z_File_Format
	///
	/// The file header consists of 3 (or optionally 4) bytes. The first two bytes
	/// contain the magic marker "0x1f 0x9d", followed by a byte of flags.
	///
	/// Based on Java code by Ronald Tschalar, which in turn was based on the unlzw.c
	/// code in the gzip package.
	/// </summary>
	/// <example> This sample shows how to unzip a compressed file
	/// <code>
	///
	/// using ICSharpCode.SharpZipLib.Core;
	/// using ICSharpCode.SharpZipLib.LZW;
	///
	/// class MainClass
	/// {
	/// 	public static void Main(string[] args)
	/// 	{
	///			using (Stream inStream = new LzwInputStream(File.OpenRead(args[0])))
	///			using (FileStream outStream = File.Create(Path.GetFileNameWithoutExtension(args[0]))) {
	///				byte[] buffer = new byte[4096];
	///				StreamUtils.Copy(inStream, outStream, buffer);
	///                         // OR
	///                         inStream.Read(buffer, 0, buffer.Length);
	///                         // now do something with the buffer
	/// 		}
	/// 	}
	/// }
	/// </code>
	/// </example>
	public class LzwInputStream : Stream
	{
		/// <summary>
		/// Gets or sets a flag indicating ownership of underlying stream.
		/// When the flag is true <see cref="Stream.Dispose()" /> will close the underlying stream also.
		/// </summary>
		/// <remarks>The default value is true.</remarks>
		public bool IsStreamOwner { get; set; } = true;

		/// <summary>
		/// Creates a LzwInputStream
		/// </summary>
		/// <param name="baseInputStream">
		/// The stream to read compressed data from (baseInputStream LZW format)
		/// </param>
		public LzwInputStream(Stream baseInputStream)
		{
			this.baseInputStream = baseInputStream;
		}

		/// <summary>
		/// See <see cref="System.IO.Stream.ReadByte"/>
		/// </summary>
		/// <returns></returns>
		public override int ReadByte()
		{
			int b = Read(one, 0, 1);
			if (b == 1)
				return (one[0] & 0xff);
			return -1;
		}

		/// <summary>
		/// Reads decompressed data into the provided buffer byte array
		/// </summary>
		/// <param name ="buffer">
		/// The array to read and decompress data into
		/// </param>
		/// <param name ="offset">
		/// The offset indicating where the data should be placed
		/// </param>
		/// <param name ="count">
		/// The number of bytes to decompress
		/// </param>
		/// <returns>The number of bytes read. Zero signals the end of stream</returns>
		public override int Read(byte[] buffer, int offset, int count)
		{
			if (!headerParsed)
				ParseHeader();

			if (eof)
				return 0;

			int start = offset;

			/* Using local copies of various variables speeds things up by as
	         * much as 30% in Java! Performance not tested in C#.
	         */
			int[] lTabPrefix = tabPrefix;
			byte[] lTabSuffix = tabSuffix;
			byte[] lStack = stack;
			int lNBits = nBits;
			int lMaxCode = maxCode;
			int lMaxMaxCode = maxMaxCode;
			int lBitMask = bitMask;
			int lOldCode = oldCode;
			byte lFinChar = finChar;
			int lStackP = stackP;
			int lFreeEnt = freeEnt;
			byte[] lData = data;
			int lBitPos = bitPos;

			// empty stack if stuff still left
			int sSize = lStack.Length - lStackP;
			if (sSize > 0)
			{
				int num = (sSize >= count) ? count : sSize;
				Array.Copy(lStack, lStackP, buffer, offset, num);
				offset += num;
				count -= num;
				lStackP += num;
			}

			if (count == 0)
			{
				stackP = lStackP;
				return offset - start;
			}

		// loop, filling local buffer until enough data has been decompressed
		MainLoop:
			do
			{
				if (end < EXTRA)
				{
					Fill();
				}

				int bitIn = (got > 0) ? (end - end % lNBits) << 3 :
										(end << 3) - (lNBits - 1);

				while (lBitPos < bitIn)
				{
					#region A

					// handle 1-byte reads correctly
					if (count == 0)
					{
						nBits = lNBits;
						maxCode = lMaxCode;
						maxMaxCode = lMaxMaxCode;
						bitMask = lBitMask;
						oldCode = lOldCode;
						finChar = lFinChar;
						stackP = lStackP;
						freeEnt = lFreeEnt;
						bitPos = lBitPos;

						return offset - start;
					}

					// check for code-width expansion
					if (lFreeEnt > lMaxCode)
					{
						int nBytes = lNBits << 3;
						lBitPos = (lBitPos - 1) +
						nBytes - (lBitPos - 1 + nBytes) % nBytes;

						lNBits++;
						lMaxCode = (lNBits == maxBits) ? lMaxMaxCode :
														(1 << lNBits) - 1;

						lBitMask = (1 << lNBits) - 1;
						lBitPos = ResetBuf(lBitPos);
						goto MainLoop;
					}

					#endregion A

					#region B

					// read next code
					int pos = lBitPos >> 3;
					int code = (((lData[pos] & 0xFF) |
						((lData[pos + 1] & 0xFF) << 8) |
						((lData[pos + 2] & 0xFF) << 16)) >>
						(lBitPos & 0x7)) & lBitMask;

					lBitPos += lNBits;

					// handle first iteration
					if (lOldCode == -1)
					{
						if (code >= 256)
							throw new LzwException("corrupt input: " + code + " > 255");

						lFinChar = (byte)(lOldCode = code);
						buffer[offset++] = lFinChar;
						count--;
						continue;
					}

					// handle CLEAR code
					if (code == TBL_CLEAR && blockMode)
					{
						Array.Copy(zeros, 0, lTabPrefix, 0, zeros.Length);
						lFreeEnt = TBL_FIRST - 1;

						int nBytes = lNBits << 3;
						lBitPos = (lBitPos - 1) + nBytes - (lBitPos - 1 + nBytes) % nBytes;
						lNBits = LzwConstants.INIT_BITS;
						lMaxCode = (1 << lNBits) - 1;
						lBitMask = lMaxCode;

						// Code tables reset

						lBitPos = ResetBuf(lBitPos);
						goto MainLoop;
					}

					#endregion B

					#region C

					// setup
					int inCode = code;
					lStackP = lStack.Length;

					// Handle KwK case
					if (code >= lFreeEnt)
					{
						if (code > lFreeEnt)
						{
							throw new LzwException("corrupt input: code=" + code +
								", freeEnt=" + lFreeEnt);
						}

						lStack[--lStackP] = lFinChar;
						code = lOldCode;
					}

					// Generate output characters in reverse order
					while (code >= 256)
					{
						lStack[--lStackP] = lTabSuffix[code];
						code = lTabPrefix[code];
					}

					lFinChar = lTabSuffix[code];
					buffer[offset++] = lFinChar;
					count--;

					// And put them out in forward order
					sSize = lStack.Length - lStackP;
					int num = (sSize >= count) ? count : sSize;
					Array.Copy(lStack, lStackP, buffer, offset, num);
					offset += num;
					count -= num;
					lStackP += num;

					#endregion C

					#region D

					// generate new entry in table
					if (lFreeEnt < lMaxMaxCode)
					{
						lTabPrefix[lFreeEnt] = lOldCode;
						lTabSuffix[lFreeEnt] = lFinChar;
						lFreeEnt++;
					}

					// Remember previous code
					lOldCode = inCode;

					// if output buffer full, then return
					if (count == 0)
					{
						nBits = lNBits;
						maxCode = lMaxCode;
						bitMask = lBitMask;
						oldCode = lOldCode;
						finChar = lFinChar;
						stackP = lStackP;
						freeEnt = lFreeEnt;
						bitPos = lBitPos;

						return offset - start;
					}

					#endregion D
				}   // while

				lBitPos = ResetBuf(lBitPos);
			} while (got > 0);  // do..while

			nBits = lNBits;
			maxCode = lMaxCode;
			bitMask = lBitMask;
			oldCode = lOldCode;
			finChar = lFinChar;
			stackP = lStackP;
			freeEnt = lFreeEnt;
			bitPos = lBitPos;

			eof = true;
			return offset - start;
		}

		/// <summary>
		/// Moves the unread data in the buffer to the beginning and resets
		/// the pointers.
		/// </summary>
		/// <param name="bitPosition"></param>
		/// <returns></returns>
		private int ResetBuf(int bitPosition)
		{
			int pos = bitPosition >> 3;
			Array.Copy(data, pos, data, 0, end - pos);
			end -= pos;
			return 0;
		}

		private void Fill()
		{
			got = baseInputStream.Read(data, end, data.Length - 1 - end);
			if (got > 0)
			{
				end += got;
			}
		}

		private void ParseHeader()
		{
			headerParsed = true;

			byte[] hdr = new byte[LzwConstants.HDR_SIZE];

			int result = baseInputStream.Read(hdr, 0, hdr.Length);

			// Check the magic marker
			if (result < 0)
				throw new LzwException("Failed to read LZW header");

			if (hdr[0] != (LzwConstants.MAGIC >> 8) || hdr[1] != (LzwConstants.MAGIC & 0xff))
			{
				throw new LzwException(String.Format(
					"Wrong LZW header. Magic bytes don't match. 0x{0:x2} 0x{1:x2}",
					hdr[0], hdr[1]));
			}

			// Check the 3rd header byte
			blockMode = (hdr[2] & LzwConstants.BLOCK_MODE_MASK) > 0;
			maxBits = hdr[2] & LzwConstants.BIT_MASK;

			if (maxBits > LzwConstants.MAX_BITS)
			{
				throw new LzwException("Stream compressed with " + maxBits +
					" bits, but decompression can only handle " +
					LzwConstants.MAX_BITS + " bits.");
			}

			if ((hdr[2] & LzwConstants.RESERVED_MASK) > 0)
			{
				throw new LzwException("Unsupported bits set in the header.");
			}

			// Initialize variables
			maxMaxCode = 1 << maxBits;
			nBits = LzwConstants.INIT_BITS;
			maxCode = (1 << nBits) - 1;
			bitMask = maxCode;
			oldCode = -1;
			finChar = 0;
			freeEnt = blockMode ? TBL_FIRST : 256;

			tabPrefix = new int[1 << maxBits];
			tabSuffix = new byte[1 << maxBits];
			stack = new byte[1 << maxBits];
			stackP = stack.Length;

			for (int idx = 255; idx >= 0; idx--)
				tabSuffix[idx] = (byte)idx;
		}

		#region Stream Overrides

		/// <summary>
		/// Gets a value indicating whether the current stream supports reading
		/// </summary>
		public override bool CanRead
		{
			get
			{
				return baseInputStream.CanRead;
			}
		}

		/// <summary>
		/// Gets a value of false indicating seeking is not supported for this stream.
		/// </summary>
		public override bool CanSeek
		{
			get
			{
				return false;
			}
		}

		/// <summary>
		/// Gets a value of false indicating that this stream is not writeable.
		/// </summary>
		public override bool CanWrite
		{
			get
			{
				return false;
			}
		}

		/// <summary>
		/// A value representing the length of the stream in bytes.
		/// </summary>
		public override long Length
		{
			get
			{
				return got;
			}
		}

		/// <summary>
		/// The current position within the stream.
		/// Throws a NotSupportedException when attempting to set the position
		/// </summary>
		/// <exception cref="NotSupportedException">Attempting to set the position</exception>
		public override long Position
		{
			get
			{
				return baseInputStream.Position;
			}
			set
			{
				throw new NotSupportedException("InflaterInputStream Position not supported");
			}
		}

		/// <summary>
		/// Flushes the baseInputStream
		/// </summary>
		public override void Flush()
		{
			baseInputStream.Flush();
		}

		/// <summary>
		/// Sets the position within the current stream
		/// Always throws a NotSupportedException
		/// </summary>
		/// <param name="offset">The relative offset to seek to.</param>
		/// <param name="origin">The <see cref="SeekOrigin"/> defining where to seek from.</param>
		/// <returns>The new position in the stream.</returns>
		/// <exception cref="NotSupportedException">Any access</exception>
		public override long Seek(long offset, SeekOrigin origin)
		{
			throw new NotSupportedException("Seek not supported");
		}

		/// <summary>
		/// Set the length of the current stream
		/// Always throws a NotSupportedException
		/// </summary>
		/// <param name="value">The new length value for the stream.</param>
		/// <exception cref="NotSupportedException">Any access</exception>
		public override void SetLength(long value)
		{
			throw new NotSupportedException("InflaterInputStream SetLength not supported");
		}

		/// <summary>
		/// Writes a sequence of bytes to stream and advances the current position
		/// This method always throws a NotSupportedException
		/// </summary>
		/// <param name="buffer">The buffer containing data to write.</param>
		/// <param name="offset">The offset of the first byte to write.</param>
		/// <param name="count">The number of bytes to write.</param>
		/// <exception cref="NotSupportedException">Any access</exception>
		public override void Write(byte[] buffer, int offset, int count)
		{
			throw new NotSupportedException("InflaterInputStream Write not supported");
		}

		/// <summary>
		/// Writes one byte to the current stream and advances the current position
		/// Always throws a NotSupportedException
		/// </summary>
		/// <param name="value">The byte to write.</param>
		/// <exception cref="NotSupportedException">Any access</exception>
		public override void WriteByte(byte value)
		{
			throw new NotSupportedException("InflaterInputStream WriteByte not supported");
		}

		/// <summary>
		/// Closes the input stream.  When <see cref="IsStreamOwner"></see>
		/// is true the underlying stream is also closed.
		/// </summary>
		protected override void Dispose(bool disposing)
		{
			if (!isClosed)
			{
				isClosed = true;
				if (IsStreamOwner)
				{
					baseInputStream.Dispose();
				}
			}
		}

		#endregion Stream Overrides

		#region Instance Fields

		private Stream baseInputStream;

		/// <summary>
		/// Flag indicating wether this instance has been closed or not.
		/// </summary>
		private bool isClosed;

		private readonly byte[] one = new byte[1];
		private bool headerParsed;

		// string table stuff
		private const int TBL_CLEAR = 0x100;

		private const int TBL_FIRST = TBL_CLEAR + 1;

		private int[] tabPrefix;
		private byte[] tabSuffix;
		private readonly int[] zeros = new int[256];
		private byte[] stack;

		// various state
		private bool blockMode;

		private int nBits;
		private int maxBits;
		private int maxMaxCode;
		private int maxCode;
		private int bitMask;
		private int oldCode;
		private byte finChar;
		private int stackP;
		private int freeEnt;

		// input buffer
		private readonly byte[] data = new byte[1024 * 8];

		private int bitPos;
		private int end;
		private int got;
		private bool eof;
		private const int EXTRA = 64;

		#endregion Instance Fields
	}
}


namespace ICSharpCode.SharpZipLib.Tar
{
	/// <summary>
	/// This exception is used to indicate that there is a problem
	/// with a TAR archive header.
	/// </summary>
	[Serializable]
	public class InvalidHeaderException : TarException
	{
		/// <summary>
		/// Initialise a new instance of the InvalidHeaderException class.
		/// </summary>
		public InvalidHeaderException()
		{
		}

		/// <summary>
		/// Initialises a new instance of the InvalidHeaderException class with a specified message.
		/// </summary>
		/// <param name="message">Message describing the exception cause.</param>
		public InvalidHeaderException(string message)
			: base(message)
		{
		}

		/// <summary>
		/// Initialise a new instance of InvalidHeaderException
		/// </summary>
		/// <param name="message">Message describing the problem.</param>
		/// <param name="exception">The exception that is the cause of the current exception.</param>
		public InvalidHeaderException(string message, Exception exception)
			: base(message, exception)
		{
		}

		/// <summary>
		/// Initializes a new instance of the InvalidHeaderException class with serialized data.
		/// </summary>
		/// <param name="info">
		/// The System.Runtime.Serialization.SerializationInfo that holds the serialized
		/// object data about the exception being thrown.
		/// </param>
		/// <param name="context">
		/// The System.Runtime.Serialization.StreamingContext that contains contextual information
		/// about the source or destination.
		/// </param>
		protected InvalidHeaderException(SerializationInfo info, StreamingContext context)
			: base(info, context)
		{
		}
	}
}


namespace ICSharpCode.SharpZipLib.Tar
{
	/// <summary>
	/// Used to advise clients of 'events' while processing archives
	/// </summary>
	public delegate void ProgressMessageHandler(TarArchive archive, TarEntry entry, string message);

	/// <summary>
	/// The TarArchive class implements the concept of a
	/// 'Tape Archive'. A tar archive is a series of entries, each of
	/// which represents a file system object. Each entry in
	/// the archive consists of a header block followed by 0 or more data blocks.
	/// Directory entries consist only of the header block, and are followed by entries
	/// for the directory's contents. File entries consist of a
	/// header followed by the number of blocks needed to
	/// contain the file's contents. All entries are written on
	/// block boundaries. Blocks are 512 bytes long.
	///
	/// TarArchives are instantiated in either read or write mode,
	/// based upon whether they are instantiated with an InputStream
	/// or an OutputStream. Once instantiated TarArchives read/write
	/// mode can not be changed.
	///
	/// There is currently no support for random access to tar archives.
	/// However, it seems that subclassing TarArchive, and using the
	/// TarBuffer.CurrentRecord and TarBuffer.CurrentBlock
	/// properties, this would be rather trivial.
	/// </summary>
	public class TarArchive : IDisposable
	{
		/// <summary>
		/// Client hook allowing detailed information to be reported during processing
		/// </summary>
		public event ProgressMessageHandler ProgressMessageEvent;

		/// <summary>
		/// Raises the ProgressMessage event
		/// </summary>
		/// <param name="entry">The <see cref="TarEntry">TarEntry</see> for this event</param>
		/// <param name="message">message for this event.  Null is no message</param>
		protected virtual void OnProgressMessageEvent(TarEntry entry, string message)
		{
			ProgressMessageHandler handler = ProgressMessageEvent;
			if (handler != null)
			{
				handler(this, entry, message);
			}
		}

		#region Constructors

		/// <summary>
		/// Constructor for a default <see cref="TarArchive"/>.
		/// </summary>
		protected TarArchive()
		{
		}

		/// <summary>
		/// Initialise a TarArchive for input.
		/// </summary>
		/// <param name="stream">The <see cref="TarInputStream"/> to use for input.</param>
		protected TarArchive(TarInputStream stream)
		{
			if (stream == null)
			{
				throw new ArgumentNullException(nameof(stream));
			}

			tarIn = stream;
		}

		/// <summary>
		/// Initialise a TarArchive for output.
		/// </summary>
		/// <param name="stream">The <see cref="TarOutputStream"/> to use for output.</param>
		protected TarArchive(TarOutputStream stream)
		{
			if (stream == null)
			{
				throw new ArgumentNullException(nameof(stream));
			}

			tarOut = stream;
		}

		#endregion Constructors

		#region Static factory methods

		/// <summary>
		/// The InputStream based constructors create a TarArchive for the
		/// purposes of extracting or listing a tar archive. Thus, use
		/// these constructors when you wish to extract files from or list
		/// the contents of an existing tar archive.
		/// </summary>
		/// <param name="inputStream">The stream to retrieve archive data from.</param>
		/// <returns>Returns a new <see cref="TarArchive"/> suitable for reading from.</returns>
		[Obsolete("No Encoding for Name field is specified, any non-ASCII bytes will be discarded")]
		public static TarArchive CreateInputTarArchive(Stream inputStream)
		{
			return CreateInputTarArchive(inputStream, null);
		}

		/// <summary>
		/// The InputStream based constructors create a TarArchive for the
		/// purposes of extracting or listing a tar archive. Thus, use
		/// these constructors when you wish to extract files from or list
		/// the contents of an existing tar archive.
		/// </summary>
		/// <param name="inputStream">The stream to retrieve archive data from.</param>
		/// <param name="nameEncoding">The <see cref="Encoding"/> used for the Name fields, or null for ASCII only</param>
		/// <returns>Returns a new <see cref="TarArchive"/> suitable for reading from.</returns>
		public static TarArchive CreateInputTarArchive(Stream inputStream, Encoding nameEncoding)
		{
			if (inputStream == null)
			{
				throw new ArgumentNullException(nameof(inputStream));
			}

			var tarStream = inputStream as TarInputStream;

			TarArchive result;
			if (tarStream != null)
			{
				result = new TarArchive(tarStream);
			}
			else
			{
				result = CreateInputTarArchive(inputStream, TarBuffer.DefaultBlockFactor, nameEncoding);
			}
			return result;
		}

		/// <summary>
		/// Create TarArchive for reading setting block factor
		/// </summary>
		/// <param name="inputStream">A stream containing the tar archive contents</param>
		/// <param name="blockFactor">The blocking factor to apply</param>
		/// <returns>Returns a <see cref="TarArchive"/> suitable for reading.</returns>
		[Obsolete("No Encoding for Name field is specified, any non-ASCII bytes will be discarded")]
		public static TarArchive CreateInputTarArchive(Stream inputStream, int blockFactor)
		{
			return CreateInputTarArchive(inputStream, blockFactor, null);
		}

		/// <summary>
		/// Create TarArchive for reading setting block factor
		/// </summary>
		/// <param name="inputStream">A stream containing the tar archive contents</param>
		/// <param name="blockFactor">The blocking factor to apply</param>
		/// <param name="nameEncoding">The <see cref="Encoding"/> used for the Name fields, or null for ASCII only</param>
		/// <returns>Returns a <see cref="TarArchive"/> suitable for reading.</returns>
		public static TarArchive CreateInputTarArchive(Stream inputStream, int blockFactor, Encoding nameEncoding)
		{
			if (inputStream == null)
			{
				throw new ArgumentNullException(nameof(inputStream));
			}

			if (inputStream is TarInputStream)
			{
				throw new ArgumentException("TarInputStream not valid");
			}

			return new TarArchive(new TarInputStream(inputStream, blockFactor, nameEncoding));
		}
		/// <summary>
		/// Create a TarArchive for writing to, using the default blocking factor
		/// </summary>
		/// <param name="outputStream">The <see cref="Stream"/> to write to</param>
		/// <param name="nameEncoding">The <see cref="Encoding"/> used for the Name fields, or null for ASCII only</param>
		/// <returns>Returns a <see cref="TarArchive"/> suitable for writing.</returns>
		public static TarArchive CreateOutputTarArchive(Stream outputStream, Encoding nameEncoding)
		{
			if (outputStream == null)
			{
				throw new ArgumentNullException(nameof(outputStream));
			}

			var tarStream = outputStream as TarOutputStream;

			TarArchive result;
			if (tarStream != null)
			{
				result = new TarArchive(tarStream);
			}
			else
			{
				result = CreateOutputTarArchive(outputStream, TarBuffer.DefaultBlockFactor, nameEncoding);
			}
			return result;
		}
		/// <summary>
		/// Create a TarArchive for writing to, using the default blocking factor
		/// </summary>
		/// <param name="outputStream">The <see cref="Stream"/> to write to</param>
		/// <returns>Returns a <see cref="TarArchive"/> suitable for writing.</returns>
		public static TarArchive CreateOutputTarArchive(Stream outputStream)
		{
			return CreateOutputTarArchive(outputStream, null);
		}

		/// <summary>
		/// Create a <see cref="TarArchive">tar archive</see> for writing.
		/// </summary>
		/// <param name="outputStream">The stream to write to</param>
		/// <param name="blockFactor">The blocking factor to use for buffering.</param>
		/// <returns>Returns a <see cref="TarArchive"/> suitable for writing.</returns>
		public static TarArchive CreateOutputTarArchive(Stream outputStream, int blockFactor)
		{
			return CreateOutputTarArchive(outputStream, blockFactor, null);
		}
		/// <summary>
		/// Create a <see cref="TarArchive">tar archive</see> for writing.
		/// </summary>
		/// <param name="outputStream">The stream to write to</param>
		/// <param name="blockFactor">The blocking factor to use for buffering.</param>
		/// <param name="nameEncoding">The <see cref="Encoding"/> used for the Name fields, or null for ASCII only</param>
		/// <returns>Returns a <see cref="TarArchive"/> suitable for writing.</returns>
		public static TarArchive CreateOutputTarArchive(Stream outputStream, int blockFactor, Encoding nameEncoding)
		{
			if (outputStream == null)
			{
				throw new ArgumentNullException(nameof(outputStream));
			}

			if (outputStream is TarOutputStream)
			{
				throw new ArgumentException("TarOutputStream is not valid");
			}

			return new TarArchive(new TarOutputStream(outputStream, blockFactor, nameEncoding));
		}

		#endregion Static factory methods

		/// <summary>
		/// Set the flag that determines whether existing files are
		/// kept, or overwritten during extraction.
		/// </summary>
		/// <param name="keepExistingFiles">
		/// If true, do not overwrite existing files.
		/// </param>
		public void SetKeepOldFiles(bool keepExistingFiles)
		{
			if (isDisposed)
			{
				throw new ObjectDisposedException("TarArchive");
			}

			keepOldFiles = keepExistingFiles;
		}

		/// <summary>
		/// Get/set the ascii file translation flag. If ascii file translation
		/// is true, then the file is checked to see if it a binary file or not.
		/// If the flag is true and the test indicates it is ascii text
		/// file, it will be translated. The translation converts the local
		/// operating system's concept of line ends into the UNIX line end,
		/// '\n', which is the defacto standard for a TAR archive. This makes
		/// text files compatible with UNIX.
		/// </summary>
		public bool AsciiTranslate
		{
			get
			{
				if (isDisposed)
				{
					throw new ObjectDisposedException("TarArchive");
				}

				return asciiTranslate;
			}

			set
			{
				if (isDisposed)
				{
					throw new ObjectDisposedException("TarArchive");
				}

				asciiTranslate = value;
			}
		}

		/// <summary>
		/// Set the ascii file translation flag.
		/// </summary>
		/// <param name= "translateAsciiFiles">
		/// If true, translate ascii text files.
		/// </param>
		[Obsolete("Use the AsciiTranslate property")]
		public void SetAsciiTranslation(bool translateAsciiFiles)
		{
			if (isDisposed)
			{
				throw new ObjectDisposedException("TarArchive");
			}

			asciiTranslate = translateAsciiFiles;
		}

		/// <summary>
		/// PathPrefix is added to entry names as they are written if the value is not null.
		/// A slash character is appended after PathPrefix
		/// </summary>
		public string PathPrefix
		{
			get
			{
				if (isDisposed)
				{
					throw new ObjectDisposedException("TarArchive");
				}

				return pathPrefix;
			}

			set
			{
				if (isDisposed)
				{
					throw new ObjectDisposedException("TarArchive");
				}

				pathPrefix = value;
			}
		}

		/// <summary>
		/// RootPath is removed from entry names if it is found at the
		/// beginning of the name.
		/// </summary>
		public string RootPath
		{
			get
			{
				if (isDisposed)
				{
					throw new ObjectDisposedException("TarArchive");
				}

				return rootPath;
			}

			set
			{
				if (isDisposed)
				{
					throw new ObjectDisposedException("TarArchive");
				}
				// Convert to forward slashes for matching. Trim trailing / for correct final path
				rootPath = value.Replace('\\', '/').TrimEnd('/');
			}
		}

		/// <summary>
		/// Set user and group information that will be used to fill in the
		/// tar archive's entry headers. This information is based on that available
		/// for the linux operating system, which is not always available on other
		/// operating systems.  TarArchive allows the programmer to specify values
		/// to be used in their place.
		/// <see cref="ApplyUserInfoOverrides"/> is set to true by this call.
		/// </summary>
		/// <param name="userId">
		/// The user id to use in the headers.
		/// </param>
		/// <param name="userName">
		/// The user name to use in the headers.
		/// </param>
		/// <param name="groupId">
		/// The group id to use in the headers.
		/// </param>
		/// <param name="groupName">
		/// The group name to use in the headers.
		/// </param>
		public void SetUserInfo(int userId, string userName, int groupId, string groupName)
		{
			if (isDisposed)
			{
				throw new ObjectDisposedException("TarArchive");
			}

			this.userId = userId;
			this.userName = userName;
			this.groupId = groupId;
			this.groupName = groupName;
			applyUserInfoOverrides = true;
		}

		/// <summary>
		/// Get or set a value indicating if overrides defined by <see cref="SetUserInfo">SetUserInfo</see> should be applied.
		/// </summary>
		/// <remarks>If overrides are not applied then the values as set in each header will be used.</remarks>
		public bool ApplyUserInfoOverrides
		{
			get
			{
				if (isDisposed)
				{
					throw new ObjectDisposedException("TarArchive");
				}

				return applyUserInfoOverrides;
			}

			set
			{
				if (isDisposed)
				{
					throw new ObjectDisposedException("TarArchive");
				}

				applyUserInfoOverrides = value;
			}
		}

		/// <summary>
		/// Get the archive user id.
		/// See <see cref="ApplyUserInfoOverrides">ApplyUserInfoOverrides</see> for detail
		/// on how to allow setting values on a per entry basis.
		/// </summary>
		/// <returns>
		/// The current user id.
		/// </returns>
		public int UserId
		{
			get
			{
				if (isDisposed)
				{
					throw new ObjectDisposedException("TarArchive");
				}

				return userId;
			}
		}

		/// <summary>
		/// Get the archive user name.
		/// See <see cref="ApplyUserInfoOverrides">ApplyUserInfoOverrides</see> for detail
		/// on how to allow setting values on a per entry basis.
		/// </summary>
		/// <returns>
		/// The current user name.
		/// </returns>
		public string UserName
		{
			get
			{
				if (isDisposed)
				{
					throw new ObjectDisposedException("TarArchive");
				}

				return userName;
			}
		}

		/// <summary>
		/// Get the archive group id.
		/// See <see cref="ApplyUserInfoOverrides">ApplyUserInfoOverrides</see> for detail
		/// on how to allow setting values on a per entry basis.
		/// </summary>
		/// <returns>
		/// The current group id.
		/// </returns>
		public int GroupId
		{
			get
			{
				if (isDisposed)
				{
					throw new ObjectDisposedException("TarArchive");
				}

				return groupId;
			}
		}

		/// <summary>
		/// Get the archive group name.
		/// See <see cref="ApplyUserInfoOverrides">ApplyUserInfoOverrides</see> for detail
		/// on how to allow setting values on a per entry basis.
		/// </summary>
		/// <returns>
		/// The current group name.
		/// </returns>
		public string GroupName
		{
			get
			{
				if (isDisposed)
				{
					throw new ObjectDisposedException("TarArchive");
				}

				return groupName;
			}
		}

		/// <summary>
		/// Get the archive's record size. Tar archives are composed of
		/// a series of RECORDS each containing a number of BLOCKS.
		/// This allowed tar archives to match the IO characteristics of
		/// the physical device being used. Archives are expected
		/// to be properly "blocked".
		/// </summary>
		/// <returns>
		/// The record size this archive is using.
		/// </returns>
		public int RecordSize
		{
			get
			{
				if (isDisposed)
				{
					throw new ObjectDisposedException("TarArchive");
				}

				if (tarIn != null)
				{
					return tarIn.RecordSize;
				}
				else if (tarOut != null)
				{
					return tarOut.RecordSize;
				}
				return TarBuffer.DefaultRecordSize;
			}
		}

		/// <summary>
		/// Sets the IsStreamOwner property on the underlying stream.
		/// Set this to false to prevent the Close of the TarArchive from closing the stream.
		/// </summary>
		public bool IsStreamOwner
		{
			set
			{
				if (tarIn != null)
				{
					tarIn.IsStreamOwner = value;
				}
				else
				{
					tarOut.IsStreamOwner = value;
				}
			}
		}

		/// <summary>
		/// Close the archive.
		/// </summary>
		[Obsolete("Use Close instead")]
		public void CloseArchive()
		{
			Close();
		}

		/// <summary>
		/// Perform the "list" command for the archive contents.
		///
		/// NOTE That this method uses the <see cref="ProgressMessageEvent"> progress event</see> to actually list
		/// the contents. If the progress display event is not set, nothing will be listed!
		/// </summary>
		public void ListContents()
		{
			if (isDisposed)
			{
				throw new ObjectDisposedException("TarArchive");
			}

			while (true)
			{
				TarEntry entry = tarIn.GetNextEntry();

				if (entry == null)
				{
					break;
				}
				OnProgressMessageEvent(entry, null);
			}
		}

		/// <summary>
		/// Perform the "extract" command and extract the contents of the archive.
		/// </summary>
		/// <param name="destinationDirectory">
		/// The destination directory into which to extract.
		/// </param>
		public void ExtractContents(string destinationDirectory) 
			=> ExtractContents(destinationDirectory, false);

		/// <summary>
		/// Perform the "extract" command and extract the contents of the archive.
		/// </summary>
		/// <param name="destinationDirectory">
		/// The destination directory into which to extract.
		/// </param>
		/// <param name="allowParentTraversal">Allow parent directory traversal in file paths (e.g. ../file)</param>
		public void ExtractContents(string destinationDirectory, bool allowParentTraversal)
		{
			if (isDisposed)
			{
				throw new ObjectDisposedException("TarArchive");
			}

			var fullDistDir = Path.GetFullPath(destinationDirectory).TrimEnd('/', '\\');

			while (true)
			{
				TarEntry entry = tarIn.GetNextEntry();

				if (entry == null)
				{
					break;
				}

				if (entry.TarHeader.TypeFlag == TarHeader.LF_LINK || entry.TarHeader.TypeFlag == TarHeader.LF_SYMLINK)
					continue;

				ExtractEntry(fullDistDir, entry, allowParentTraversal);
			}
		}

		/// <summary>
		/// Extract an entry from the archive. This method assumes that the
		/// tarIn stream has been properly set with a call to GetNextEntry().
		/// </summary>
		/// <param name="destDir">
		/// The destination directory into which to extract.
		/// </param>
		/// <param name="entry">
		/// The TarEntry returned by tarIn.GetNextEntry().
		/// </param>
		/// <param name="allowParentTraversal">Allow parent directory traversal in file paths (e.g. ../file)</param>
		private void ExtractEntry(string destDir, TarEntry entry, bool allowParentTraversal)
		{
			OnProgressMessageEvent(entry, null);

			string name = entry.Name;

			if (Path.IsPathRooted(name))
			{
				// NOTE:
				// for UNC names...  \\machine\share\zoom\beet.txt gives \zoom\beet.txt
				name = name.Substring(Path.GetPathRoot(name).Length);
			}

			name = name.Replace('/', Path.DirectorySeparatorChar);

			string destFile = Path.Combine(destDir, name);
			var destFileDir = Path.GetDirectoryName(Path.GetFullPath(destFile)) ?? "";

			if (!allowParentTraversal && !destFileDir.StartsWith(destDir, StringComparison.InvariantCultureIgnoreCase))
			{
				throw new InvalidNameException("Parent traversal in paths is not allowed");
			}

			if (entry.IsDirectory)
			{
				EnsureDirectoryExists(destFile);
			}
			else
			{
				string parentDirectory = Path.GetDirectoryName(destFile);
				EnsureDirectoryExists(parentDirectory);

				bool process = true;
				var fileInfo = new FileInfo(destFile);
				if (fileInfo.Exists)
				{
					if (keepOldFiles)
					{
						OnProgressMessageEvent(entry, "Destination file already exists");
						process = false;
					}
					else if ((fileInfo.Attributes & FileAttributes.ReadOnly) != 0)
					{
						OnProgressMessageEvent(entry, "Destination file already exists, and is read-only");
						process = false;
					}
				}

				if (process)
				{
					using (var outputStream = File.Create(destFile))
					{
						if (this.asciiTranslate)
						{
							// May need to translate the file.
							ExtractAndTranslateEntry(destFile, outputStream);
						}
						else
						{
							// If translation is disabled, just copy the entry across directly.
							tarIn.CopyEntryContents(outputStream);
						}
					}
				}
			}
		}

		// Extract a TAR entry, and perform an ASCII translation if required.
		private void ExtractAndTranslateEntry(string destFile, Stream outputStream)
		{
			bool asciiTrans = !IsBinary(destFile);

			if (asciiTrans)
			{
				using (var outw = new StreamWriter(outputStream, new UTF8Encoding(false), 1024, true))
				{
					byte[] rdbuf = new byte[32 * 1024];

					while (true)
					{
						int numRead = tarIn.Read(rdbuf, 0, rdbuf.Length);

						if (numRead <= 0)
						{
							break;
						}

						for (int off = 0, b = 0; b < numRead; ++b)
						{
							if (rdbuf[b] == 10)
							{
								string s = Encoding.ASCII.GetString(rdbuf, off, (b - off));
								outw.WriteLine(s);
								off = b + 1;
							}
						}
					}
				}
			}
			else
			{
				// No translation required.
				tarIn.CopyEntryContents(outputStream);
			}
		}

		/// <summary>
		/// Write an entry to the archive. This method will call the putNextEntry
		/// and then write the contents of the entry, and finally call closeEntry()
		/// for entries that are files. For directories, it will call putNextEntry(),
		/// and then, if the recurse flag is true, process each entry that is a
		/// child of the directory.
		/// </summary>
		/// <param name="sourceEntry">
		/// The TarEntry representing the entry to write to the archive.
		/// </param>
		/// <param name="recurse">
		/// If true, process the children of directory entries.
		/// </param>
		public void WriteEntry(TarEntry sourceEntry, bool recurse)
		{
			if (sourceEntry == null)
			{
				throw new ArgumentNullException(nameof(sourceEntry));
			}

			if (isDisposed)
			{
				throw new ObjectDisposedException("TarArchive");
			}

			try
			{
				if (recurse)
				{
					TarHeader.SetValueDefaults(sourceEntry.UserId, sourceEntry.UserName,
											   sourceEntry.GroupId, sourceEntry.GroupName);
				}
				WriteEntryCore(sourceEntry, recurse);
			}
			finally
			{
				if (recurse)
				{
					TarHeader.RestoreSetValues();
				}
			}
		}

		/// <summary>
		/// Write an entry to the archive. This method will call the putNextEntry
		/// and then write the contents of the entry, and finally call closeEntry()
		/// for entries that are files. For directories, it will call putNextEntry(),
		/// and then, if the recurse flag is true, process each entry that is a
		/// child of the directory.
		/// </summary>
		/// <param name="sourceEntry">
		/// The TarEntry representing the entry to write to the archive.
		/// </param>
		/// <param name="recurse">
		/// If true, process the children of directory entries.
		/// </param>
		private void WriteEntryCore(TarEntry sourceEntry, bool recurse)
		{
			string tempFileName = null;
			string entryFilename = sourceEntry.File;

			var entry = (TarEntry)sourceEntry.Clone();

			if (applyUserInfoOverrides)
			{
				entry.GroupId = groupId;
				entry.GroupName = groupName;
				entry.UserId = userId;
				entry.UserName = userName;
			}

			OnProgressMessageEvent(entry, null);

			if (asciiTranslate && !entry.IsDirectory)
			{
				if (!IsBinary(entryFilename))
				{
					tempFileName = PathUtils.GetTempFileName();

					using (StreamReader inStream = File.OpenText(entryFilename))
					{
						using (Stream outStream = File.Create(tempFileName))
						{
							while (true)
							{
								string line = inStream.ReadLine();
								if (line == null)
								{
									break;
								}
								byte[] data = Encoding.ASCII.GetBytes(line);
								outStream.Write(data, 0, data.Length);
								outStream.WriteByte((byte)'\n');
							}

							outStream.Flush();
						}
					}

					entry.Size = new FileInfo(tempFileName).Length;
					entryFilename = tempFileName;
				}
			}

			string newName = null;

			if (!String.IsNullOrEmpty(rootPath))
			{
				if (entry.Name.StartsWith(rootPath, StringComparison.OrdinalIgnoreCase))
				{
					newName = entry.Name.Substring(rootPath.Length + 1);
				}
			}

			if (pathPrefix != null)
			{
				newName = (newName == null) ? pathPrefix + "/" + entry.Name : pathPrefix + "/" + newName;
			}

			if (newName != null)
			{
				entry.Name = newName;
			}

			tarOut.PutNextEntry(entry);

			if (entry.IsDirectory)
			{
				if (recurse)
				{
					TarEntry[] list = entry.GetDirectoryEntries();
					for (int i = 0; i < list.Length; ++i)
					{
						WriteEntryCore(list[i], recurse);
					}
				}
			}
			else
			{
				using (Stream inputStream = File.OpenRead(entryFilename))
				{
					byte[] localBuffer = new byte[32 * 1024];
					while (true)
					{
						int numRead = inputStream.Read(localBuffer, 0, localBuffer.Length);

						if (numRead <= 0)
						{
							break;
						}

						tarOut.Write(localBuffer, 0, numRead);
					}
				}

				if (!string.IsNullOrEmpty(tempFileName))
				{
					File.Delete(tempFileName);
				}

				tarOut.CloseEntry();
			}
		}

		/// <summary>
		/// Performs application-defined tasks associated with freeing, releasing, or resetting unmanaged resources.
		/// </summary>
		public void Dispose()
		{
			Dispose(true);
			GC.SuppressFinalize(this);
		}

		/// <summary>
		/// Releases the unmanaged resources used by the FileStream and optionally releases the managed resources.
		/// </summary>
		/// <param name="disposing">true to release both managed and unmanaged resources;
		/// false to release only unmanaged resources.</param>
		protected virtual void Dispose(bool disposing)
		{
			if (!isDisposed)
			{
				isDisposed = true;
				if (disposing)
				{
					if (tarOut != null)
					{
						tarOut.Flush();
						tarOut.Dispose();
					}

					if (tarIn != null)
					{
						tarIn.Dispose();
					}
				}
			}
		}

		/// <summary>
		/// Closes the archive and releases any associated resources.
		/// </summary>
		public virtual void Close()
		{
			Dispose(true);
		}

		/// <summary>
		/// Ensures that resources are freed and other cleanup operations are performed
		/// when the garbage collector reclaims the <see cref="TarArchive"/>.
		/// </summary>
		~TarArchive()
		{
			Dispose(false);
		}

		private static void EnsureDirectoryExists(string directoryName)
		{
			if (!Directory.Exists(directoryName))
			{
				try
				{
					Directory.CreateDirectory(directoryName);
				}
				catch (Exception e)
				{
					throw new TarException("Exception creating directory '" + directoryName + "', " + e.Message, e);
				}
			}
		}

		// TODO: TarArchive - Is there a better way to test for a text file?
		// It no longer reads entire files into memory but is still a weak test!
		// This assumes that byte values 0-7, 14-31 or 255 are binary
		// and that all non text files contain one of these values
		private static bool IsBinary(string filename)
		{
			using (FileStream fs = File.OpenRead(filename))
			{
				int sampleSize = Math.Min(4096, (int)fs.Length);
				byte[] content = new byte[sampleSize];

				int bytesRead = fs.Read(content, 0, sampleSize);

				for (int i = 0; i < bytesRead; ++i)
				{
					byte b = content[i];
					if ((b < 8) || ((b > 13) && (b < 32)) || (b == 255))
					{
						return true;
					}
				}
			}
			return false;
		}

		#region Instance Fields

		private bool keepOldFiles;
		private bool asciiTranslate;

		private int userId;
		private string userName = string.Empty;
		private int groupId;
		private string groupName = string.Empty;

		private string rootPath;
		private string pathPrefix;

		private bool applyUserInfoOverrides;

		private TarInputStream tarIn;
		private TarOutputStream tarOut;
		private bool isDisposed;

		#endregion Instance Fields
	}
}


namespace ICSharpCode.SharpZipLib.Tar
{
	/// <summary>
	/// The TarBuffer class implements the tar archive concept
	/// of a buffered input stream. This concept goes back to the
	/// days of blocked tape drives and special io devices. In the
	/// C# universe, the only real function that this class
	/// performs is to ensure that files have the correct "record"
	/// size, or other tars will complain.
	/// <p>
	/// You should never have a need to access this class directly.
	/// TarBuffers are created by Tar IO Streams.
	/// </p>
	/// </summary>
	public class TarBuffer
	{
		/* A quote from GNU tar man file on blocking and records
		   A `tar' archive file contains a series of blocks.  Each block
		contains `BLOCKSIZE' bytes.  Although this format may be thought of as
		being on magnetic tape, other media are often used.

		   Each file archived is represented by a header block which describes
		the file, followed by zero or more blocks which give the contents of
		the file.  At the end of the archive file there may be a block filled
		with binary zeros as an end-of-file marker.  A reasonable system should
		write a block of zeros at the end, but must not assume that such a
		block exists when reading an archive.

		   The blocks may be "blocked" for physical I/O operations.  Each
		record of N blocks is written with a single 'write ()'
		operation.  On magnetic tapes, the result of such a write is a single
		record.  When writing an archive, the last record of blocks should be
		written at the full size, with blocks after the zero block containing
		all zeros.  When reading an archive, a reasonable system should
		properly handle an archive whose last record is shorter than the rest,
		or which contains garbage records after a zero block.
		*/

		#region Constants

		/// <summary>
		/// The size of a block in a tar archive in bytes.
		/// </summary>
		/// <remarks>This is 512 bytes.</remarks>
		public const int BlockSize = 512;

		/// <summary>
		/// The number of blocks in a default record.
		/// </summary>
		/// <remarks>
		/// The default value is 20 blocks per record.
		/// </remarks>
		public const int DefaultBlockFactor = 20;

		/// <summary>
		/// The size in bytes of a default record.
		/// </summary>
		/// <remarks>
		/// The default size is 10KB.
		/// </remarks>
		public const int DefaultRecordSize = BlockSize * DefaultBlockFactor;

		#endregion Constants

		/// <summary>
		/// Get the record size for this buffer
		/// </summary>
		/// <value>The record size in bytes.
		/// This is equal to the <see cref="BlockFactor"/> multiplied by the <see cref="BlockSize"/></value>
		public int RecordSize
		{
			get
			{
				return recordSize;
			}
		}

		/// <summary>
		/// Get the TAR Buffer's record size.
		/// </summary>
		/// <returns>The record size in bytes.
		/// This is equal to the <see cref="BlockFactor"/> multiplied by the <see cref="BlockSize"/></returns>
		[Obsolete("Use RecordSize property instead")]
		public int GetRecordSize()
		{
			return recordSize;
		}

		/// <summary>
		/// Get the Blocking factor for the buffer
		/// </summary>
		/// <value>This is the number of blocks in each record.</value>
		public int BlockFactor
		{
			get
			{
				return blockFactor;
			}
		}

		/// <summary>
		/// Get the TAR Buffer's block factor
		/// </summary>
		/// <returns>The block factor; the number of blocks per record.</returns>
		[Obsolete("Use BlockFactor property instead")]
		public int GetBlockFactor()
		{
			return blockFactor;
		}

		/// <summary>
		/// Construct a default TarBuffer
		/// </summary>
		protected TarBuffer()
		{
		}

		/// <summary>
		/// Create TarBuffer for reading with default BlockFactor
		/// </summary>
		/// <param name="inputStream">Stream to buffer</param>
		/// <returns>A new <see cref="TarBuffer"/> suitable for input.</returns>
		public static TarBuffer CreateInputTarBuffer(Stream inputStream)
		{
			if (inputStream == null)
			{
				throw new ArgumentNullException(nameof(inputStream));
			}

			return CreateInputTarBuffer(inputStream, DefaultBlockFactor);
		}

		/// <summary>
		/// Construct TarBuffer for reading inputStream setting BlockFactor
		/// </summary>
		/// <param name="inputStream">Stream to buffer</param>
		/// <param name="blockFactor">Blocking factor to apply</param>
		/// <returns>A new <see cref="TarBuffer"/> suitable for input.</returns>
		public static TarBuffer CreateInputTarBuffer(Stream inputStream, int blockFactor)
		{
			if (inputStream == null)
			{
				throw new ArgumentNullException(nameof(inputStream));
			}

			if (blockFactor <= 0)
			{
				throw new ArgumentOutOfRangeException(nameof(blockFactor), "Factor cannot be negative");
			}

			var tarBuffer = new TarBuffer();
			tarBuffer.inputStream = inputStream;
			tarBuffer.outputStream = null;
			tarBuffer.Initialize(blockFactor);

			return tarBuffer;
		}

		/// <summary>
		/// Construct TarBuffer for writing with default BlockFactor
		/// </summary>
		/// <param name="outputStream">output stream for buffer</param>
		/// <returns>A new <see cref="TarBuffer"/> suitable for output.</returns>
		public static TarBuffer CreateOutputTarBuffer(Stream outputStream)
		{
			if (outputStream == null)
			{
				throw new ArgumentNullException(nameof(outputStream));
			}

			return CreateOutputTarBuffer(outputStream, DefaultBlockFactor);
		}

		/// <summary>
		/// Construct TarBuffer for writing Tar output to streams.
		/// </summary>
		/// <param name="outputStream">Output stream to write to.</param>
		/// <param name="blockFactor">Blocking factor to apply</param>
		/// <returns>A new <see cref="TarBuffer"/> suitable for output.</returns>
		public static TarBuffer CreateOutputTarBuffer(Stream outputStream, int blockFactor)
		{
			if (outputStream == null)
			{
				throw new ArgumentNullException(nameof(outputStream));
			}

			if (blockFactor <= 0)
			{
				throw new ArgumentOutOfRangeException(nameof(blockFactor), "Factor cannot be negative");
			}

			var tarBuffer = new TarBuffer();
			tarBuffer.inputStream = null;
			tarBuffer.outputStream = outputStream;
			tarBuffer.Initialize(blockFactor);

			return tarBuffer;
		}

		/// <summary>
		/// Initialization common to all constructors.
		/// </summary>
		private void Initialize(int archiveBlockFactor)
		{
			blockFactor = archiveBlockFactor;
			recordSize = archiveBlockFactor * BlockSize;
			recordBuffer = new byte[RecordSize];

			if (inputStream != null)
			{
				currentRecordIndex = -1;
				currentBlockIndex = BlockFactor;
			}
			else
			{
				currentRecordIndex = 0;
				currentBlockIndex = 0;
			}
		}

		/// <summary>
		/// Determine if an archive block indicates End of Archive. End of
		/// archive is indicated by a block that consists entirely of null bytes.
		/// All remaining blocks for the record should also be null's
		/// However some older tars only do a couple of null blocks (Old GNU tar for one)
		/// and also partial records
		/// </summary>
		/// <param name = "block">The data block to check.</param>
		/// <returns>Returns true if the block is an EOF block; false otherwise.</returns>
		[Obsolete("Use IsEndOfArchiveBlock instead")]
		public bool IsEOFBlock(byte[] block)
		{
			if (block == null)
			{
				throw new ArgumentNullException(nameof(block));
			}

			if (block.Length != BlockSize)
			{
				throw new ArgumentException("block length is invalid");
			}

			for (int i = 0; i < BlockSize; ++i)
			{
				if (block[i] != 0)
				{
					return false;
				}
			}

			return true;
		}

		/// <summary>
		/// Determine if an archive block indicates the End of an Archive has been reached.
		/// End of archive is indicated by a block that consists entirely of null bytes.
		/// All remaining blocks for the record should also be null's
		/// However some older tars only do a couple of null blocks (Old GNU tar for one)
		/// and also partial records
		/// </summary>
		/// <param name = "block">The data block to check.</param>
		/// <returns>Returns true if the block is an EOF block; false otherwise.</returns>
		public static bool IsEndOfArchiveBlock(byte[] block)
		{
			if (block == null)
			{
				throw new ArgumentNullException(nameof(block));
			}

			if (block.Length != BlockSize)
			{
				throw new ArgumentException("block length is invalid");
			}

			for (int i = 0; i < BlockSize; ++i)
			{
				if (block[i] != 0)
				{
					return false;
				}
			}

			return true;
		}

		/// <summary>
		/// Skip over a block on the input stream.
		/// </summary>
		public void SkipBlock()
		{
			if (inputStream == null)
			{
				throw new TarException("no input stream defined");
			}

			if (currentBlockIndex >= BlockFactor)
			{
				if (!ReadRecord())
				{
					throw new TarException("Failed to read a record");
				}
			}

			currentBlockIndex++;
		}

		/// <summary>
		/// Read a block from the input stream.
		/// </summary>
		/// <returns>
		/// The block of data read.
		/// </returns>
		public byte[] ReadBlock()
		{
			if (inputStream == null)
			{
				throw new TarException("TarBuffer.ReadBlock - no input stream defined");
			}

			if (currentBlockIndex >= BlockFactor)
			{
				if (!ReadRecord())
				{
					throw new TarException("Failed to read a record");
				}
			}

			byte[] result = new byte[BlockSize];

			Array.Copy(recordBuffer, (currentBlockIndex * BlockSize), result, 0, BlockSize);
			currentBlockIndex++;
			return result;
		}

		/// <summary>
		/// Read a record from data stream.
		/// </summary>
		/// <returns>
		/// false if End-Of-File, else true.
		/// </returns>
		private bool ReadRecord()
		{
			if (inputStream == null)
			{
				throw new TarException("no input stream defined");
			}

			currentBlockIndex = 0;

			int offset = 0;
			int bytesNeeded = RecordSize;

			while (bytesNeeded > 0)
			{
				long numBytes = inputStream.Read(recordBuffer, offset, bytesNeeded);

				//
				// NOTE
				// We have found EOF, and the record is not full!
				//
				// This is a broken archive. It does not follow the standard
				// blocking algorithm. However, because we are generous, and
				// it requires little effort, we will simply ignore the error
				// and continue as if the entire record were read. This does
				// not appear to break anything upstream. We used to return
				// false in this case.
				//
				// Thanks to 'Yohann.Roussel@alcatel.fr' for this fix.
				//
				if (numBytes <= 0)
				{
					break;
				}

				offset += (int)numBytes;
				bytesNeeded -= (int)numBytes;
			}

			currentRecordIndex++;
			return true;
		}

		/// <summary>
		/// Get the current block number, within the current record, zero based.
		/// </summary>
		/// <remarks>Block numbers are zero based values</remarks>
		/// <seealso cref="RecordSize"/>
		public int CurrentBlock
		{
			get { return currentBlockIndex; }
		}

		/// <summary>
		/// Gets or sets a flag indicating ownership of underlying stream.
		/// When the flag is true <see cref="Close" /> will close the underlying stream also.
		/// </summary>
		/// <remarks>The default value is true.</remarks>
		public bool IsStreamOwner { get; set; } = true;

		/// <summary>
		/// Get the current block number, within the current record, zero based.
		/// </summary>
		/// <returns>
		/// The current zero based block number.
		/// </returns>
		/// <remarks>
		/// The absolute block number = (<see cref="GetCurrentRecordNum">record number</see> * <see cref="BlockFactor">block factor</see>) + <see cref="GetCurrentBlockNum">block number</see>.
		/// </remarks>
		[Obsolete("Use CurrentBlock property instead")]
		public int GetCurrentBlockNum()
		{
			return currentBlockIndex;
		}

		/// <summary>
		/// Get the current record number.
		/// </summary>
		/// <returns>
		/// The current zero based record number.
		/// </returns>
		public int CurrentRecord
		{
			get { return currentRecordIndex; }
		}

		/// <summary>
		/// Get the current record number.
		/// </summary>
		/// <returns>
		/// The current zero based record number.
		/// </returns>
		[Obsolete("Use CurrentRecord property instead")]
		public int GetCurrentRecordNum()
		{
			return currentRecordIndex;
		}

		/// <summary>
		/// Write a block of data to the archive.
		/// </summary>
		/// <param name="block">
		/// The data to write to the archive.
		/// </param>
		public void WriteBlock(byte[] block)
		{
			if (block == null)
			{
				throw new ArgumentNullException(nameof(block));
			}

			if (outputStream == null)
			{
				throw new TarException("TarBuffer.WriteBlock - no output stream defined");
			}

			if (block.Length != BlockSize)
			{
				string errorText = string.Format("TarBuffer.WriteBlock - block to write has length '{0}' which is not the block size of '{1}'",
					block.Length, BlockSize);
				throw new TarException(errorText);
			}

			if (currentBlockIndex >= BlockFactor)
			{
				WriteRecord();
			}

			Array.Copy(block, 0, recordBuffer, (currentBlockIndex * BlockSize), BlockSize);
			currentBlockIndex++;
		}

		/// <summary>
		/// Write an archive record to the archive, where the record may be
		/// inside of a larger array buffer. The buffer must be "offset plus
		/// record size" long.
		/// </summary>
		/// <param name="buffer">
		/// The buffer containing the record data to write.
		/// </param>
		/// <param name="offset">
		/// The offset of the record data within buffer.
		/// </param>
		public void WriteBlock(byte[] buffer, int offset)
		{
			if (buffer == null)
			{
				throw new ArgumentNullException(nameof(buffer));
			}

			if (outputStream == null)
			{
				throw new TarException("TarBuffer.WriteBlock - no output stream defined");
			}

			if ((offset < 0) || (offset >= buffer.Length))
			{
				throw new ArgumentOutOfRangeException(nameof(offset));
			}

			if ((offset + BlockSize) > buffer.Length)
			{
				string errorText = string.Format("TarBuffer.WriteBlock - record has length '{0}' with offset '{1}' which is less than the record size of '{2}'",
					buffer.Length, offset, recordSize);
				throw new TarException(errorText);
			}

			if (currentBlockIndex >= BlockFactor)
			{
				WriteRecord();
			}

			Array.Copy(buffer, offset, recordBuffer, (currentBlockIndex * BlockSize), BlockSize);

			currentBlockIndex++;
		}

		/// <summary>
		/// Write a TarBuffer record to the archive.
		/// </summary>
		private void WriteRecord()
		{
			if (outputStream == null)
			{
				throw new TarException("TarBuffer.WriteRecord no output stream defined");
			}

			outputStream.Write(recordBuffer, 0, RecordSize);
			outputStream.Flush();

			currentBlockIndex = 0;
			currentRecordIndex++;
		}

		/// <summary>
		/// WriteFinalRecord writes the current record buffer to output any unwritten data is present.
		/// </summary>
		/// <remarks>Any trailing bytes are set to zero which is by definition correct behaviour
		/// for the end of a tar stream.</remarks>
		private void WriteFinalRecord()
		{
			if (outputStream == null)
			{
				throw new TarException("TarBuffer.WriteFinalRecord no output stream defined");
			}

			if (currentBlockIndex > 0)
			{
				int dataBytes = currentBlockIndex * BlockSize;
				Array.Clear(recordBuffer, dataBytes, RecordSize - dataBytes);
				WriteRecord();
			}

			outputStream.Flush();
		}

		/// <summary>
		/// Close the TarBuffer. If this is an output buffer, also flush the
		/// current block before closing.
		/// </summary>
		public void Close()
		{
			if (outputStream != null)
			{
				WriteFinalRecord();

				if (IsStreamOwner)
				{
					outputStream.Dispose();
				}
				outputStream = null;
			}
			else if (inputStream != null)
			{
				if (IsStreamOwner)
				{
					inputStream.Dispose();
				}
				inputStream = null;
			}
		}

		#region Instance Fields

		private Stream inputStream;
		private Stream outputStream;

		private byte[] recordBuffer;
		private int currentBlockIndex;
		private int currentRecordIndex;

		private int recordSize = DefaultRecordSize;
		private int blockFactor = DefaultBlockFactor;

		#endregion Instance Fields
	}
}


namespace ICSharpCode.SharpZipLib.Tar
{
	/// <summary>
	/// This class represents an entry in a Tar archive. It consists
	/// of the entry's header, as well as the entry's File. Entries
	/// can be instantiated in one of three ways, depending on how
	/// they are to be used.
	/// <p>
	/// TarEntries that are created from the header bytes read from
	/// an archive are instantiated with the TarEntry( byte[] )
	/// constructor. These entries will be used when extracting from
	/// or listing the contents of an archive. These entries have their
	/// header filled in using the header bytes. They also set the File
	/// to null, since they reference an archive entry not a file.</p>
	/// <p>
	/// TarEntries that are created from files that are to be written
	/// into an archive are instantiated with the CreateEntryFromFile(string)
	/// pseudo constructor. These entries have their header filled in using
	/// the File's information. They also keep a reference to the File
	/// for convenience when writing entries.</p>
	/// <p>
	/// Finally, TarEntries can be constructed from nothing but a name.
	/// This allows the programmer to construct the entry by hand, for
	/// instance when only an InputStream is available for writing to
	/// the archive, and the header information is constructed from
	/// other information. In this case the header fields are set to
	/// defaults and the File is set to null.</p>
	/// <see cref="TarHeader"/>
	/// </summary>
	public class TarEntry
	{
		#region Constructors

		/// <summary>
		/// Initialise a default instance of <see cref="TarEntry"/>.
		/// </summary>
		private TarEntry()
		{
			header = new TarHeader();
		}

		/// <summary>
		/// Construct an entry from an archive's header bytes. File is set
		/// to null.
		/// </summary>
		/// <param name = "headerBuffer">
		/// The header bytes from a tar archive entry.
		/// </param>
		[Obsolete("No Encoding for Name field is specified, any non-ASCII bytes will be discarded")]
		public TarEntry(byte[] headerBuffer) : this(headerBuffer, null)
		{
		}

		/// <summary>
		/// Construct an entry from an archive's header bytes. File is set
		/// to null.
		/// </summary>
		/// <param name = "headerBuffer">
		/// The header bytes from a tar archive entry.
		/// </param>
		/// <param name = "nameEncoding">
		/// The <see cref="Encoding"/> used for the Name fields, or null for ASCII only
		/// </param>
		public TarEntry(byte[] headerBuffer, Encoding nameEncoding)
		{
			header = new TarHeader();
			header.ParseBuffer(headerBuffer, nameEncoding);
		}

		/// <summary>
		/// Construct a TarEntry using the <paramref name="header">header</paramref> provided
		/// </summary>
		/// <param name="header">Header details for entry</param>
		public TarEntry(TarHeader header)
		{
			if (header == null)
			{
				throw new ArgumentNullException(nameof(header));
			}

			this.header = (TarHeader)header.Clone();
		}

		#endregion Constructors

		#region ICloneable Members

		/// <summary>
		/// Clone this tar entry.
		/// </summary>
		/// <returns>Returns a clone of this entry.</returns>
		public object Clone()
		{
			var entry = new TarEntry();
			entry.file = file;
			entry.header = (TarHeader)header.Clone();
			entry.Name = Name;
			return entry;
		}

		#endregion ICloneable Members

		/// <summary>
		/// Construct an entry with only a <paramref name="name">name</paramref>.
		/// This allows the programmer to construct the entry's header "by hand".
		/// </summary>
		/// <param name="name">The name to use for the entry</param>
		/// <returns>Returns the newly created <see cref="TarEntry"/></returns>
		public static TarEntry CreateTarEntry(string name)
		{
			var entry = new TarEntry();
			TarEntry.NameTarHeader(entry.header, name);
			return entry;
		}

		/// <summary>
		/// Construct an entry for a file. File is set to file, and the
		/// header is constructed from information from the file.
		/// </summary>
		/// <param name = "fileName">The file name that the entry represents.</param>
		/// <returns>Returns the newly created <see cref="TarEntry"/></returns>
		public static TarEntry CreateEntryFromFile(string fileName)
		{
			var entry = new TarEntry();
			entry.GetFileTarHeader(entry.header, fileName);
			return entry;
		}

		/// <summary>
		/// Determine if the two entries are equal. Equality is determined
		/// by the header names being equal.
		/// </summary>
		/// <param name="obj">The <see cref="Object"/> to compare with the current Object.</param>
		/// <returns>
		/// True if the entries are equal; false if not.
		/// </returns>
		public override bool Equals(object obj)
		{
			var localEntry = obj as TarEntry;

			if (localEntry != null)
			{
				return Name.Equals(localEntry.Name);
			}
			return false;
		}

		/// <summary>
		/// Derive a Hash value for the current <see cref="Object"/>
		/// </summary>
		/// <returns>A Hash code for the current <see cref="Object"/></returns>
		public override int GetHashCode()
		{
			return Name.GetHashCode();
		}

		/// <summary>
		/// Determine if the given entry is a descendant of this entry.
		/// Descendancy is determined by the name of the descendant
		/// starting with this entry's name.
		/// </summary>
		/// <param name = "toTest">
		/// Entry to be checked as a descendent of this.
		/// </param>
		/// <returns>
		/// True if entry is a descendant of this.
		/// </returns>
		public bool IsDescendent(TarEntry toTest)
		{
			if (toTest == null)
			{
				throw new ArgumentNullException(nameof(toTest));
			}

			return toTest.Name.StartsWith(Name, StringComparison.Ordinal);
		}

		/// <summary>
		/// Get this entry's header.
		/// </summary>
		/// <returns>
		/// This entry's TarHeader.
		/// </returns>
		public TarHeader TarHeader
		{
			get
			{
				return header;
			}
		}

		/// <summary>
		/// Get/Set this entry's name.
		/// </summary>
		public string Name
		{
			get
			{
				return header.Name;
			}
			set
			{
				header.Name = value;
			}
		}

		/// <summary>
		/// Get/set this entry's user id.
		/// </summary>
		public int UserId
		{
			get
			{
				return header.UserId;
			}
			set
			{
				header.UserId = value;
			}
		}

		/// <summary>
		/// Get/set this entry's group id.
		/// </summary>
		public int GroupId
		{
			get
			{
				return header.GroupId;
			}
			set
			{
				header.GroupId = value;
			}
		}

		/// <summary>
		/// Get/set this entry's user name.
		/// </summary>
		public string UserName
		{
			get
			{
				return header.UserName;
			}
			set
			{
				header.UserName = value;
			}
		}

		/// <summary>
		/// Get/set this entry's group name.
		/// </summary>
		public string GroupName
		{
			get
			{
				return header.GroupName;
			}
			set
			{
				header.GroupName = value;
			}
		}

		/// <summary>
		/// Convenience method to set this entry's group and user ids.
		/// </summary>
		/// <param name="userId">
		/// This entry's new user id.
		/// </param>
		/// <param name="groupId">
		/// This entry's new group id.
		/// </param>
		public void SetIds(int userId, int groupId)
		{
			UserId = userId;
			GroupId = groupId;
		}

		/// <summary>
		/// Convenience method to set this entry's group and user names.
		/// </summary>
		/// <param name="userName">
		/// This entry's new user name.
		/// </param>
		/// <param name="groupName">
		/// This entry's new group name.
		/// </param>
		public void SetNames(string userName, string groupName)
		{
			UserName = userName;
			GroupName = groupName;
		}

		/// <summary>
		/// Get/Set the modification time for this entry
		/// </summary>
		public DateTime ModTime
		{
			get
			{
				return header.ModTime;
			}
			set
			{
				header.ModTime = value;
			}
		}

		/// <summary>
		/// Get this entry's file.
		/// </summary>
		/// <returns>
		/// This entry's file.
		/// </returns>
		public string File
		{
			get
			{
				return file;
			}
		}

		/// <summary>
		/// Get/set this entry's recorded file size.
		/// </summary>
		public long Size
		{
			get
			{
				return header.Size;
			}
			set
			{
				header.Size = value;
			}
		}

		/// <summary>
		/// Return true if this entry represents a directory, false otherwise
		/// </summary>
		/// <returns>
		/// True if this entry is a directory.
		/// </returns>
		public bool IsDirectory
		{
			get
			{
				if (file != null)
				{
					return Directory.Exists(file);
				}

				if (header != null)
				{
					if ((header.TypeFlag == TarHeader.LF_DIR) || Name.EndsWith("/", StringComparison.Ordinal))
					{
						return true;
					}
				}
				return false;
			}
		}

		/// <summary>
		/// Fill in a TarHeader with information from a File.
		/// </summary>
		/// <param name="header">
		/// The TarHeader to fill in.
		/// </param>
		/// <param name="file">
		/// The file from which to get the header information.
		/// </param>
		public void GetFileTarHeader(TarHeader header, string file)
		{
			if (header == null)
			{
				throw new ArgumentNullException(nameof(header));
			}

			if (file == null)
			{
				throw new ArgumentNullException(nameof(file));
			}

			this.file = file;

			// bugfix from torhovl from #D forum:
			string name = file;

			// 23-Jan-2004 GnuTar allows device names in path where the name is not local to the current directory
			if (name.IndexOf(Directory.GetCurrentDirectory(), StringComparison.Ordinal) == 0)
			{
				name = name.Substring(Directory.GetCurrentDirectory().Length);
			}

			/*
						if (Path.DirectorySeparatorChar == '\\')
						{
							// check if the OS is Windows
							// Strip off drive letters!
							if (name.Length > 2)
							{
								char ch1 = name[0];
								char ch2 = name[1];

								if (ch2 == ':' && Char.IsLetter(ch1))
								{
									name = name.Substring(2);
								}
							}
						}
			*/

			name = name.Replace(Path.DirectorySeparatorChar, '/');

			// No absolute pathnames
			// Windows (and Posix?) paths can start with UNC style "\\NetworkDrive\",
			// so we loop on starting /'s.
			while (name.StartsWith("/", StringComparison.Ordinal))
			{
				name = name.Substring(1);
			}

			header.LinkName = String.Empty;
			header.Name = name;

			if (Directory.Exists(file))
			{
				header.Mode = 1003; // Magic number for security access for a UNIX filesystem
				header.TypeFlag = TarHeader.LF_DIR;
				if ((header.Name.Length == 0) || header.Name[header.Name.Length - 1] != '/')
				{
					header.Name = header.Name + "/";
				}

				header.Size = 0;
			}
			else
			{
				header.Mode = 33216; // Magic number for security access for a UNIX filesystem
				header.TypeFlag = TarHeader.LF_NORMAL;
				header.Size = new FileInfo(file.Replace('/', Path.DirectorySeparatorChar)).Length;
			}

			header.ModTime = System.IO.File.GetLastWriteTime(file.Replace('/', Path.DirectorySeparatorChar)).ToUniversalTime();
			header.DevMajor = 0;
			header.DevMinor = 0;
		}

		/// <summary>
		/// Get entries for all files present in this entries directory.
		/// If this entry doesnt represent a directory zero entries are returned.
		/// </summary>
		/// <returns>
		/// An array of TarEntry's for this entry's children.
		/// </returns>
		public TarEntry[] GetDirectoryEntries()
		{
			if ((file == null) || !Directory.Exists(file))
			{
				return Empty.Array<TarEntry>();
			}

			string[] list = Directory.GetFileSystemEntries(file);
			TarEntry[] result = new TarEntry[list.Length];

			for (int i = 0; i < list.Length; ++i)
			{
				result[i] = TarEntry.CreateEntryFromFile(list[i]);
			}

			return result;
		}

		/// <summary>
		/// Write an entry's header information to a header buffer.
		/// </summary>
		/// <param name = "outBuffer">
		/// The tar entry header buffer to fill in.
		/// </param>
		[Obsolete("No Encoding for Name field is specified, any non-ASCII bytes will be discarded")]
		public void WriteEntryHeader(byte[] outBuffer)
		{
			WriteEntryHeader(outBuffer, null);
		}

		/// <summary>
		/// Write an entry's header information to a header buffer.
		/// </summary>
		/// <param name = "outBuffer">
		/// The tar entry header buffer to fill in.
		/// </param>
		/// <param name = "nameEncoding">
		/// The <see cref="Encoding"/> used for the Name fields, or null for ASCII only
		/// </param>
		public void WriteEntryHeader(byte[] outBuffer, Encoding nameEncoding)
		{
			header.WriteHeader(outBuffer, nameEncoding);
		}

		/// <summary>
		/// Convenience method that will modify an entry's name directly
		/// in place in an entry header buffer byte array.
		/// </summary>
		/// <param name="buffer">
		/// The buffer containing the entry header to modify.
		/// </param>
		/// <param name="newName">
		/// The new name to place into the header buffer.
		/// </param>
		[Obsolete("No Encoding for Name field is specified, any non-ASCII bytes will be discarded")]
		static public void AdjustEntryName(byte[] buffer, string newName)
		{
			AdjustEntryName(buffer, newName, null);
		}

		/// <summary>
		/// Convenience method that will modify an entry's name directly
		/// in place in an entry header buffer byte array.
		/// </summary>
		/// <param name="buffer">
		/// The buffer containing the entry header to modify.
		/// </param>
		/// <param name="newName">
		/// The new name to place into the header buffer.
		/// </param>
		/// <param name="nameEncoding">
		/// The <see cref="Encoding"/> used for the Name fields, or null for ASCII only
		/// </param>
		static public void AdjustEntryName(byte[] buffer, string newName, Encoding nameEncoding)
		{
			TarHeader.GetNameBytes(newName, buffer, 0, TarHeader.NAMELEN, nameEncoding);
		}

		/// <summary>
		/// Fill in a TarHeader given only the entry's name.
		/// </summary>
		/// <param name="header">
		/// The TarHeader to fill in.
		/// </param>
		/// <param name="name">
		/// The tar entry name.
		/// </param>
		static public void NameTarHeader(TarHeader header, string name)
		{
			if (header == null)
			{
				throw new ArgumentNullException(nameof(header));
			}

			if (name == null)
			{
				throw new ArgumentNullException(nameof(name));
			}

			bool isDir = name.EndsWith("/", StringComparison.Ordinal);

			header.Name = name;
			header.Mode = isDir ? 1003 : 33216;
			header.UserId = 0;
			header.GroupId = 0;
			header.Size = 0;

			header.ModTime = DateTime.UtcNow;

			header.TypeFlag = isDir ? TarHeader.LF_DIR : TarHeader.LF_NORMAL;

			header.LinkName = String.Empty;
			header.UserName = String.Empty;
			header.GroupName = String.Empty;

			header.DevMajor = 0;
			header.DevMinor = 0;
		}

		#region Instance Fields

		/// <summary>
		/// The name of the file this entry represents or null if the entry is not based on a file.
		/// </summary>
		private string file;

		/// <summary>
		/// The entry's header information.
		/// </summary>
		private TarHeader header;

		#endregion Instance Fields
	}
}


namespace ICSharpCode.SharpZipLib.Tar
{
	/// <summary>
	/// TarException represents exceptions specific to Tar classes and code.
	/// </summary>
	[Serializable]
	public class TarException : SharpZipBaseException
	{
		/// <summary>
		/// Initialise a new instance of <see cref="TarException" />.
		/// </summary>
		public TarException()
		{
		}

		/// <summary>
		/// Initialise a new instance of <see cref="TarException" /> with its message string.
		/// </summary>
		/// <param name="message">A <see cref="string"/> that describes the error.</param>
		public TarException(string message)
			: base(message)
		{
		}

		/// <summary>
		/// Initialise a new instance of <see cref="TarException" />.
		/// </summary>
		/// <param name="message">A <see cref="string"/> that describes the error.</param>
		/// <param name="innerException">The <see cref="Exception"/> that caused this exception.</param>
		public TarException(string message, Exception innerException)
			: base(message, innerException)
		{
		}

		/// <summary>
		/// Initializes a new instance of the TarException class with serialized data.
		/// </summary>
		/// <param name="info">
		/// The System.Runtime.Serialization.SerializationInfo that holds the serialized
		/// object data about the exception being thrown.
		/// </param>
		/// <param name="context">
		/// The System.Runtime.Serialization.StreamingContext that contains contextual information
		/// about the source or destination.
		/// </param>
		protected TarException(SerializationInfo info, StreamingContext context)
			: base(info, context)
		{
		}
	}
}


namespace ICSharpCode.SharpZipLib.Tar
{
	/// <summary>
	/// Reads the extended header of a Tar stream
	/// </summary>
	public class TarExtendedHeaderReader
	{
		private const byte LENGTH = 0;
		private const byte KEY = 1;
		private const byte VALUE = 2;
		private const byte END = 3;

		private readonly Dictionary<string, string> headers = new Dictionary<string, string>();

		private string[] headerParts = new string[3];

		private int bbIndex;
		private byte[] byteBuffer;
		private char[] charBuffer;

		private readonly StringBuilder sb = new StringBuilder();
		private readonly Decoder decoder = Encoding.UTF8.GetDecoder();

		private int state = LENGTH;

		private static readonly byte[] StateNext = new[] { (byte)' ', (byte)'=', (byte)'\n' };

		/// <summary>
		/// Creates a new <see cref="TarExtendedHeaderReader"/>.
		/// </summary>
		public TarExtendedHeaderReader()
		{
			ResetBuffers();
		}

		/// <summary>
		/// Read <paramref name="length"/> bytes from <paramref name="buffer"/>
		/// </summary>
		/// <param name="buffer"></param>
		/// <param name="length"></param>
		public void Read(byte[] buffer, int length)
		{
			for (int i = 0; i < length; i++)
			{
				byte next = buffer[i];

				if (next == StateNext[state])
				{
					Flush();
					headerParts[state] = sb.ToString();
					sb.Clear();

					if (++state == END)
					{
						headers.Add(headerParts[KEY], headerParts[VALUE]);
						headerParts = new string[3];
						state = LENGTH;
					}
				}
				else
				{
					byteBuffer[bbIndex++] = next;
					if (bbIndex == 4)
						Flush();
				}
			}
		}

		private void Flush()
		{
			decoder.Convert(byteBuffer, 0, bbIndex, charBuffer, 0, 4, false, out int bytesUsed, out int charsUsed, out bool completed);

			sb.Append(charBuffer, 0, charsUsed);
			ResetBuffers();
		}

		private void ResetBuffers()
		{
			charBuffer = new char[4];
			byteBuffer = new byte[4];
			bbIndex = 0;
		}

		/// <summary>
		/// Returns the parsed headers as key-value strings
		/// </summary>
		public Dictionary<string, string> Headers
		{
			get
			{
				// TODO: Check for invalid state? -NM 2018-07-01
				return headers;
			}
		}
	}
}


namespace ICSharpCode.SharpZipLib.Tar
{
	/// <summary>
	/// This class encapsulates the Tar Entry Header used in Tar Archives.
	/// The class also holds a number of tar constants, used mostly in headers.
	/// </summary>
	/// <remarks>
	///    The tar format and its POSIX successor PAX have a long history which makes for compatability
	///    issues when creating and reading files.
	///
	///    This is further complicated by a large number of programs with variations on formats
	///    One common issue is the handling of names longer than 100 characters.
	///    GNU style long names are currently supported.
	///
	/// This is the ustar (Posix 1003.1) header.
	///
	/// struct header
	/// {
	/// 	char t_name[100];          //   0 Filename
	/// 	char t_mode[8];            // 100 Permissions
	/// 	char t_uid[8];             // 108 Numerical User ID
	/// 	char t_gid[8];             // 116 Numerical Group ID
	/// 	char t_size[12];           // 124 Filesize
	/// 	char t_mtime[12];          // 136 st_mtime
	/// 	char t_chksum[8];          // 148 Checksum
	/// 	char t_typeflag;           // 156 Type of File
	/// 	char t_linkname[100];      // 157 Target of Links
	/// 	char t_magic[6];           // 257 "ustar" or other...
	/// 	char t_version[2];         // 263 Version fixed to 00
	/// 	char t_uname[32];          // 265 User Name
	/// 	char t_gname[32];          // 297 Group Name
	/// 	char t_devmajor[8];        // 329 Major for devices
	/// 	char t_devminor[8];        // 337 Minor for devices
	/// 	char t_prefix[155];        // 345 Prefix for t_name
	/// 	char t_mfill[12];          // 500 Filler up to 512
	/// };
	/// </remarks>
	public class TarHeader
	{
		#region Constants

		/// <summary>
		/// The length of the name field in a header buffer.
		/// </summary>
		public const int NAMELEN = 100;

		/// <summary>
		/// The length of the mode field in a header buffer.
		/// </summary>
		public const int MODELEN = 8;

		/// <summary>
		/// The length of the user id field in a header buffer.
		/// </summary>
		public const int UIDLEN = 8;

		/// <summary>
		/// The length of the group id field in a header buffer.
		/// </summary>
		public const int GIDLEN = 8;

		/// <summary>
		/// The length of the checksum field in a header buffer.
		/// </summary>
		public const int CHKSUMLEN = 8;

		/// <summary>
		/// Offset of checksum in a header buffer.
		/// </summary>
		public const int CHKSUMOFS = 148;

		/// <summary>
		/// The length of the size field in a header buffer.
		/// </summary>
		public const int SIZELEN = 12;

		/// <summary>
		/// The length of the magic field in a header buffer.
		/// </summary>
		public const int MAGICLEN = 6;

		/// <summary>
		/// The length of the version field in a header buffer.
		/// </summary>
		public const int VERSIONLEN = 2;

		/// <summary>
		/// The length of the modification time field in a header buffer.
		/// </summary>
		public const int MODTIMELEN = 12;

		/// <summary>
		/// The length of the user name field in a header buffer.
		/// </summary>
		public const int UNAMELEN = 32;

		/// <summary>
		/// The length of the group name field in a header buffer.
		/// </summary>
		public const int GNAMELEN = 32;

		/// <summary>
		/// The length of the devices field in a header buffer.
		/// </summary>
		public const int DEVLEN = 8;

		/// <summary>
		/// The length of the name prefix field in a header buffer.
		/// </summary>
		public const int PREFIXLEN = 155;

		//
		// LF_ constants represent the "type" of an entry
		//

		/// <summary>
		///  The "old way" of indicating a normal file.
		/// </summary>
		public const byte LF_OLDNORM = 0;

		/// <summary>
		/// Normal file type.
		/// </summary>
		public const byte LF_NORMAL = (byte)'0';

		/// <summary>
		/// Link file type.
		/// </summary>
		public const byte LF_LINK = (byte)'1';

		/// <summary>
		/// Symbolic link file type.
		/// </summary>
		public const byte LF_SYMLINK = (byte)'2';

		/// <summary>
		/// Character device file type.
		/// </summary>
		public const byte LF_CHR = (byte)'3';

		/// <summary>
		/// Block device file type.
		/// </summary>
		public const byte LF_BLK = (byte)'4';

		/// <summary>
		/// Directory file type.
		/// </summary>
		public const byte LF_DIR = (byte)'5';

		/// <summary>
		/// FIFO (pipe) file type.
		/// </summary>
		public const byte LF_FIFO = (byte)'6';

		/// <summary>
		/// Contiguous file type.
		/// </summary>
		public const byte LF_CONTIG = (byte)'7';

		/// <summary>
		/// Posix.1 2001 global extended header
		/// </summary>
		public const byte LF_GHDR = (byte)'g';

		/// <summary>
		/// Posix.1 2001 extended header
		/// </summary>
		public const byte LF_XHDR = (byte)'x';

		// POSIX allows for upper case ascii type as extensions

		/// <summary>
		/// Solaris access control list file type
		/// </summary>
		public const byte LF_ACL = (byte)'A';

		/// <summary>
		/// GNU dir dump file type
		/// This is a dir entry that contains the names of files that were in the
		/// dir at the time the dump was made
		/// </summary>
		public const byte LF_GNU_DUMPDIR = (byte)'D';

		/// <summary>
		/// Solaris Extended Attribute File
		/// </summary>
		public const byte LF_EXTATTR = (byte)'E';

		/// <summary>
		/// Inode (metadata only) no file content
		/// </summary>
		public const byte LF_META = (byte)'I';

		/// <summary>
		/// Identifies the next file on the tape as having a long link name
		/// </summary>
		public const byte LF_GNU_LONGLINK = (byte)'K';

		/// <summary>
		/// Identifies the next file on the tape as having a long name
		/// </summary>
		public const byte LF_GNU_LONGNAME = (byte)'L';

		/// <summary>
		/// Continuation of a file that began on another volume
		/// </summary>
		public const byte LF_GNU_MULTIVOL = (byte)'M';

		/// <summary>
		/// For storing filenames that dont fit in the main header (old GNU)
		/// </summary>
		public const byte LF_GNU_NAMES = (byte)'N';

		/// <summary>
		/// GNU Sparse file
		/// </summary>
		public const byte LF_GNU_SPARSE = (byte)'S';

		/// <summary>
		/// GNU Tape/volume header ignore on extraction
		/// </summary>
		public const byte LF_GNU_VOLHDR = (byte)'V';

		/// <summary>
		/// The magic tag representing a POSIX tar archive.  (would be written with a trailing NULL)
		/// </summary>
		public const string TMAGIC = "ustar";

		/// <summary>
		/// The magic tag representing an old GNU tar archive where version is included in magic and overwrites it
		/// </summary>
		public const string GNU_TMAGIC = "ustar  ";

		private const long timeConversionFactor = 10000000L;           // 1 tick == 100 nanoseconds
		private static readonly DateTime dateTime1970 = new DateTime(1970, 1, 1, 0, 0, 0, 0);

		#endregion Constants

		#region Constructors

		/// <summary>
		/// Initialise a default TarHeader instance
		/// </summary>
		public TarHeader()
		{
			Magic = TMAGIC;
			Version = " ";

			Name = "";
			LinkName = "";

			UserId = defaultUserId;
			GroupId = defaultGroupId;
			UserName = defaultUser;
			GroupName = defaultGroupName;
			Size = 0;
		}

		#endregion Constructors

		#region Properties

		/// <summary>
		/// Get/set the name for this tar entry.
		/// </summary>
		/// <exception cref="ArgumentNullException">Thrown when attempting to set the property to null.</exception>
		public string Name
		{
			get { return name; }
			set
			{
				if (value == null)
				{
					throw new ArgumentNullException(nameof(value));
				}
				name = value;
			}
		}

		/// <summary>
		/// Get the name of this entry.
		/// </summary>
		/// <returns>The entry's name.</returns>
		[Obsolete("Use the Name property instead", true)]
		public string GetName()
		{
			return name;
		}

		/// <summary>
		/// Get/set the entry's Unix style permission mode.
		/// </summary>
		public int Mode
		{
			get { return mode; }
			set { mode = value; }
		}

		/// <summary>
		/// The entry's user id.
		/// </summary>
		/// <remarks>
		/// This is only directly relevant to unix systems.
		/// The default is zero.
		/// </remarks>
		public int UserId
		{
			get { return userId; }
			set { userId = value; }
		}

		/// <summary>
		/// Get/set the entry's group id.
		/// </summary>
		/// <remarks>
		/// This is only directly relevant to linux/unix systems.
		/// The default value is zero.
		/// </remarks>
		public int GroupId
		{
			get { return groupId; }
			set { groupId = value; }
		}

		/// <summary>
		/// Get/set the entry's size.
		/// </summary>
		/// <exception cref="ArgumentOutOfRangeException">Thrown when setting the size to less than zero.</exception>
		public long Size
		{
			get { return size; }
			set
			{
				if (value < 0)
				{
					throw new ArgumentOutOfRangeException(nameof(value), "Cannot be less than zero");
				}
				size = value;
			}
		}

		/// <summary>
		/// Get/set the entry's modification time.
		/// </summary>
		/// <remarks>
		/// The modification time is only accurate to within a second.
		/// </remarks>
		/// <exception cref="ArgumentOutOfRangeException">Thrown when setting the date time to less than 1/1/1970.</exception>
		public DateTime ModTime
		{
			get { return modTime; }
			set
			{
				if (value < dateTime1970)
				{
					throw new ArgumentOutOfRangeException(nameof(value), "ModTime cannot be before Jan 1st 1970");
				}
				modTime = new DateTime(value.Year, value.Month, value.Day, value.Hour, value.Minute, value.Second);
			}
		}

		/// <summary>
		/// Get the entry's checksum.  This is only valid/updated after writing or reading an entry.
		/// </summary>
		public int Checksum
		{
			get { return checksum; }
		}

		/// <summary>
		/// Get value of true if the header checksum is valid, false otherwise.
		/// </summary>
		public bool IsChecksumValid
		{
			get { return isChecksumValid; }
		}

		/// <summary>
		/// Get/set the entry's type flag.
		/// </summary>
		public byte TypeFlag
		{
			get { return typeFlag; }
			set { typeFlag = value; }
		}

		/// <summary>
		/// The entry's link name.
		/// </summary>
		/// <exception cref="ArgumentNullException">Thrown when attempting to set LinkName to null.</exception>
		public string LinkName
		{
			get { return linkName; }
			set
			{
				if (value == null)
				{
					throw new ArgumentNullException(nameof(value));
				}
				linkName = value;
			}
		}

		/// <summary>
		/// Get/set the entry's magic tag.
		/// </summary>
		/// <exception cref="ArgumentNullException">Thrown when attempting to set Magic to null.</exception>
		public string Magic
		{
			get { return magic; }
			set
			{
				if (value == null)
				{
					throw new ArgumentNullException(nameof(value));
				}
				magic = value;
			}
		}

		/// <summary>
		/// The entry's version.
		/// </summary>
		/// <exception cref="ArgumentNullException">Thrown when attempting to set Version to null.</exception>
		public string Version
		{
			get
			{
				return version;
			}

			set
			{
				if (value == null)
				{
					throw new ArgumentNullException(nameof(value));
				}
				version = value;
			}
		}

		/// <summary>
		/// The entry's user name.
		/// </summary>
		public string UserName
		{
			get { return userName; }
			set
			{
				if (value != null)
				{
					userName = value.Substring(0, Math.Min(UNAMELEN, value.Length));
				}
				else
				{
					string currentUser = "user";
					if (currentUser.Length > UNAMELEN)
					{
						currentUser = currentUser.Substring(0, UNAMELEN);
					}
					userName = currentUser;
				}
			}
		}

		/// <summary>
		/// Get/set the entry's group name.
		/// </summary>
		/// <remarks>
		/// This is only directly relevant to unix systems.
		/// </remarks>
		public string GroupName
		{
			get { return groupName; }
			set
			{
				if (value == null)
				{
					groupName = "None";
				}
				else
				{
					groupName = value;
				}
			}
		}

		/// <summary>
		/// Get/set the entry's major device number.
		/// </summary>
		public int DevMajor
		{
			get { return devMajor; }
			set { devMajor = value; }
		}

		/// <summary>
		/// Get/set the entry's minor device number.
		/// </summary>
		public int DevMinor
		{
			get { return devMinor; }
			set { devMinor = value; }
		}

		#endregion Properties

		#region ICloneable Members

		/// <summary>
		/// Create a new <see cref="TarHeader"/> that is a copy of the current instance.
		/// </summary>
		/// <returns>A new <see cref="Object"/> that is a copy of the current instance.</returns>
		public object Clone()
		{
			return this.MemberwiseClone();
		}

		#endregion ICloneable Members

		/// <summary>
		/// Parse TarHeader information from a header buffer.
		/// </summary>
		/// <param name = "header">
		/// The tar entry header buffer to get information from.
		/// </param>
		/// <param name = "nameEncoding">
		/// The <see cref="Encoding"/> used for the Name field, or null for ASCII only
		/// </param>
		public void ParseBuffer(byte[] header, Encoding nameEncoding)
		{
			if (header == null)
			{
				throw new ArgumentNullException(nameof(header));
			}

			int offset = 0;

			name = ParseName(header, offset, NAMELEN, nameEncoding).ToString();
			offset += NAMELEN;

			mode = (int)ParseOctal(header, offset, MODELEN);
			offset += MODELEN;

			UserId = (int)ParseOctal(header, offset, UIDLEN);
			offset += UIDLEN;

			GroupId = (int)ParseOctal(header, offset, GIDLEN);
			offset += GIDLEN;

			Size = ParseBinaryOrOctal(header, offset, SIZELEN);
			offset += SIZELEN;

			ModTime = GetDateTimeFromCTime(ParseOctal(header, offset, MODTIMELEN));
			offset += MODTIMELEN;

			checksum = (int)ParseOctal(header, offset, CHKSUMLEN);
			offset += CHKSUMLEN;

			TypeFlag = header[offset++];

			LinkName = ParseName(header, offset, NAMELEN, nameEncoding).ToString();
			offset += NAMELEN;

			Magic = ParseName(header, offset, MAGICLEN, nameEncoding).ToString();
			offset += MAGICLEN;

			if (Magic == "ustar")
			{
				Version = ParseName(header, offset, VERSIONLEN, nameEncoding).ToString();
				offset += VERSIONLEN;

				UserName = ParseName(header, offset, UNAMELEN, nameEncoding).ToString();
				offset += UNAMELEN;

				GroupName = ParseName(header, offset, GNAMELEN, nameEncoding).ToString();
				offset += GNAMELEN;

				DevMajor = (int)ParseOctal(header, offset, DEVLEN);
				offset += DEVLEN;

				DevMinor = (int)ParseOctal(header, offset, DEVLEN);
				offset += DEVLEN;

				string prefix = ParseName(header, offset, PREFIXLEN, nameEncoding).ToString();
				if (!string.IsNullOrEmpty(prefix)) Name = prefix + '/' + Name;
			}

			isChecksumValid = Checksum == TarHeader.MakeCheckSum(header);
		}

		/// <summary>
		/// Parse TarHeader information from a header buffer.
		/// </summary>
		/// <param name = "header">
		/// The tar entry header buffer to get information from.
		/// </param>
		[Obsolete("No Encoding for Name field is specified, any non-ASCII bytes will be discarded")]
		public void ParseBuffer(byte[] header)
		{
			ParseBuffer(header, null);
		}

		/// <summary>
		/// 'Write' header information to buffer provided, updating the <see cref="Checksum">check sum</see>.
		/// </summary>
		/// <param name="outBuffer">output buffer for header information</param>
		[Obsolete("No Encoding for Name field is specified, any non-ASCII bytes will be discarded")]
		public void WriteHeader(byte[] outBuffer)
		{
			WriteHeader(outBuffer, null);
		}

		/// <summary>
		/// 'Write' header information to buffer provided, updating the <see cref="Checksum">check sum</see>.
		/// </summary>
		/// <param name="outBuffer">output buffer for header information</param>
		/// <param name="nameEncoding">The <see cref="Encoding"/> used for the Name field, or null for ASCII only</param>
		public void WriteHeader(byte[] outBuffer, Encoding nameEncoding)
		{
			if (outBuffer == null)
			{
				throw new ArgumentNullException(nameof(outBuffer));
			}

			int offset = 0;

			offset = GetNameBytes(Name, outBuffer, offset, NAMELEN, nameEncoding);
			offset = GetOctalBytes(mode, outBuffer, offset, MODELEN);
			offset = GetOctalBytes(UserId, outBuffer, offset, UIDLEN);
			offset = GetOctalBytes(GroupId, outBuffer, offset, GIDLEN);

			offset = GetBinaryOrOctalBytes(Size, outBuffer, offset, SIZELEN);
			offset = GetOctalBytes(GetCTime(ModTime), outBuffer, offset, MODTIMELEN);

			int csOffset = offset;
			for (int c = 0; c < CHKSUMLEN; ++c)
			{
				outBuffer[offset++] = (byte)' ';
			}

			outBuffer[offset++] = TypeFlag;

			offset = GetNameBytes(LinkName, outBuffer, offset, NAMELEN, nameEncoding);
			offset = GetAsciiBytes(Magic, 0, outBuffer, offset, MAGICLEN, nameEncoding);
			offset = GetNameBytes(Version, outBuffer, offset, VERSIONLEN, nameEncoding);
			offset = GetNameBytes(UserName, outBuffer, offset, UNAMELEN, nameEncoding);
			offset = GetNameBytes(GroupName, outBuffer, offset, GNAMELEN, nameEncoding);

			if ((TypeFlag == LF_CHR) || (TypeFlag == LF_BLK))
			{
				offset = GetOctalBytes(DevMajor, outBuffer, offset, DEVLEN);
				offset = GetOctalBytes(DevMinor, outBuffer, offset, DEVLEN);
			}

			for (; offset < outBuffer.Length;)
			{
				outBuffer[offset++] = 0;
			}

			checksum = ComputeCheckSum(outBuffer);

			GetCheckSumOctalBytes(checksum, outBuffer, csOffset, CHKSUMLEN);
			isChecksumValid = true;
		}

		/// <summary>
		/// Get a hash code for the current object.
		/// </summary>
		/// <returns>A hash code for the current object.</returns>
		public override int GetHashCode()
		{
			return Name.GetHashCode();
		}

		/// <summary>
		/// Determines if this instance is equal to the specified object.
		/// </summary>
		/// <param name="obj">The object to compare with.</param>
		/// <returns>true if the objects are equal, false otherwise.</returns>
		public override bool Equals(object obj)
		{
			var localHeader = obj as TarHeader;

			bool result;
			if (localHeader != null)
			{
				result = (name == localHeader.name)
					&& (mode == localHeader.mode)
					&& (UserId == localHeader.UserId)
					&& (GroupId == localHeader.GroupId)
					&& (Size == localHeader.Size)
					&& (ModTime == localHeader.ModTime)
					&& (Checksum == localHeader.Checksum)
					&& (TypeFlag == localHeader.TypeFlag)
					&& (LinkName == localHeader.LinkName)
					&& (Magic == localHeader.Magic)
					&& (Version == localHeader.Version)
					&& (UserName == localHeader.UserName)
					&& (GroupName == localHeader.GroupName)
					&& (DevMajor == localHeader.DevMajor)
					&& (DevMinor == localHeader.DevMinor);
			}
			else
			{
				result = false;
			}
			return result;
		}

		/// <summary>
		/// Set defaults for values used when constructing a TarHeader instance.
		/// </summary>
		/// <param name="userId">Value to apply as a default for userId.</param>
		/// <param name="userName">Value to apply as a default for userName.</param>
		/// <param name="groupId">Value to apply as a default for groupId.</param>
		/// <param name="groupName">Value to apply as a default for groupName.</param>
		static internal void SetValueDefaults(int userId, string userName, int groupId, string groupName)
		{
			defaultUserId = userIdAsSet = userId;
			defaultUser = userNameAsSet = userName;
			defaultGroupId = groupIdAsSet = groupId;
			defaultGroupName = groupNameAsSet = groupName;
		}

		static internal void RestoreSetValues()
		{
			defaultUserId = userIdAsSet;
			defaultUser = userNameAsSet;
			defaultGroupId = groupIdAsSet;
			defaultGroupName = groupNameAsSet;
		}

		// Return value that may be stored in octal or binary. Length must exceed 8.
		//
		static private long ParseBinaryOrOctal(byte[] header, int offset, int length)
		{
			if (header[offset] >= 0x80)
			{
				// File sizes over 8GB are stored in 8 right-justified bytes of binary indicated by setting the high-order bit of the leftmost byte of a numeric field.
				long result = 0;
				for (int pos = length - 8; pos < length; pos++)
				{
					result = result << 8 | header[offset + pos];
				}
				return result;
			}
			return ParseOctal(header, offset, length);
		}

		/// <summary>
		/// Parse an octal string from a header buffer.
		/// </summary>
		/// <param name = "header">The header buffer from which to parse.</param>
		/// <param name = "offset">The offset into the buffer from which to parse.</param>
		/// <param name = "length">The number of header bytes to parse.</param>
		/// <returns>The long equivalent of the octal string.</returns>
		static public long ParseOctal(byte[] header, int offset, int length)
		{
			if (header == null)
			{
				throw new ArgumentNullException(nameof(header));
			}

			long result = 0;
			bool stillPadding = true;

			int end = offset + length;
			for (int i = offset; i < end; ++i)
			{
				if (header[i] == 0)
				{
					break;
				}

				if (header[i] == (byte)' ' || header[i] == '0')
				{
					if (stillPadding)
					{
						continue;
					}

					if (header[i] == (byte)' ')
					{
						break;
					}
				}

				stillPadding = false;

				result = (result << 3) + (header[i] - '0');
			}

			return result;
		}

		/// <summary>
		/// Parse a name from a header buffer.
		/// </summary>
		/// <param name="header">
		/// The header buffer from which to parse.
		/// </param>
		/// <param name="offset">
		/// The offset into the buffer from which to parse.
		/// </param>
		/// <param name="length">
		/// The number of header bytes to parse.
		/// </param>
		/// <returns>
		/// The name parsed.
		/// </returns>
		[Obsolete("No Encoding for Name field is specified, any non-ASCII bytes will be discarded")]
		static public StringBuilder ParseName(byte[] header, int offset, int length)
		{
			return ParseName(header, offset, length, null);
		}

		/// <summary>
		/// Parse a name from a header buffer.
		/// </summary>
		/// <param name="header">
		/// The header buffer from which to parse.
		/// </param>
		/// <param name="offset">
		/// The offset into the buffer from which to parse.
		/// </param>
		/// <param name="length">
		/// The number of header bytes to parse.
		/// </param>
		/// <param name="encoding">
		/// name encoding, or null for ASCII only
		/// </param>
		/// <returns>
		/// The name parsed.
		/// </returns>
		static public StringBuilder ParseName(byte[] header, int offset, int length, Encoding encoding)
		{
			if (header == null)
			{
				throw new ArgumentNullException(nameof(header));
			}

			if (offset < 0)
			{
				throw new ArgumentOutOfRangeException(nameof(offset), "Cannot be less than zero");
			}

			if (length < 0)
			{
				throw new ArgumentOutOfRangeException(nameof(length), "Cannot be less than zero");
			}

			if (offset + length > header.Length)
			{
				throw new ArgumentException("Exceeds header size", nameof(length));
			}

			var result = new StringBuilder(length);

			int count = 0;
			if(encoding == null)
			{
				for (int i = offset; i < offset + length; ++i)
				{
					if (header[i] == 0)
					{
						break;
					}
					result.Append((char)header[i]);
				}
			}
			else
			{
				for(int i = offset; i < offset + length; ++i, ++count)
				{
					if(header[i] == 0)
					{
						break;
					}
				}
				result.Append(encoding.GetString(header, offset, count));
			}

			return result;
		}

		/// <summary>
		/// Add <paramref name="name">name</paramref> to the buffer as a collection of bytes
		/// </summary>
		/// <param name="name">The name to add</param>
		/// <param name="nameOffset">The offset of the first character</param>
		/// <param name="buffer">The buffer to add to</param>
		/// <param name="bufferOffset">The index of the first byte to add</param>
		/// <param name="length">The number of characters/bytes to add</param>
		/// <returns>The next free index in the <paramref name="buffer"/></returns>
		public static int GetNameBytes(StringBuilder name, int nameOffset, byte[] buffer, int bufferOffset, int length)
		{
			return GetNameBytes(name.ToString(), nameOffset, buffer, bufferOffset, length, null);
		}

		/// <summary>
		/// Add <paramref name="name">name</paramref> to the buffer as a collection of bytes
		/// </summary>
		/// <param name="name">The name to add</param>
		/// <param name="nameOffset">The offset of the first character</param>
		/// <param name="buffer">The buffer to add to</param>
		/// <param name="bufferOffset">The index of the first byte to add</param>
		/// <param name="length">The number of characters/bytes to add</param>
		/// <returns>The next free index in the <paramref name="buffer"/></returns>
		public static int GetNameBytes(string name, int nameOffset, byte[] buffer, int bufferOffset, int length)
		{
			return GetNameBytes(name, nameOffset, buffer, bufferOffset, length, null);
		}

		/// <summary>
		/// Add <paramref name="name">name</paramref> to the buffer as a collection of bytes
		/// </summary>
		/// <param name="name">The name to add</param>
		/// <param name="nameOffset">The offset of the first character</param>
		/// <param name="buffer">The buffer to add to</param>
		/// <param name="bufferOffset">The index of the first byte to add</param>
		/// <param name="length">The number of characters/bytes to add</param>
		/// <param name="encoding">name encoding, or null for ASCII only</param>
		/// <returns>The next free index in the <paramref name="buffer"/></returns>
		public static int GetNameBytes(string name, int nameOffset, byte[] buffer, int bufferOffset, int length, Encoding encoding)
		{
			if (name == null)
			{
				throw new ArgumentNullException(nameof(name));
			}

			if (buffer == null)
			{
				throw new ArgumentNullException(nameof(buffer));
			}

			int i;
			if(encoding != null)
			{
				var nameArray = name.ToCharArray(nameOffset, Math.Min(name.Length - nameOffset, length));
				var bytes = encoding.GetBytes(nameArray, 0, nameArray.Length);
				i = Math.Min(bytes.Length, length);
				Array.Copy(bytes, 0, buffer, bufferOffset, i);
			}
			else
			{
				for (i = 0; i < length && nameOffset + i < name.Length; ++i)
				{
					buffer[bufferOffset + i] = (byte)name[nameOffset + i];
				}
			}

			for (; i < length; ++i)
			{
				buffer[bufferOffset + i] = 0;
			}
			return bufferOffset + length;
		}
		/// <summary>
		/// Add an entry name to the buffer
		/// </summary>
		/// <param name="name">
		/// The name to add
		/// </param>
		/// <param name="buffer">
		/// The buffer to add to
		/// </param>
		/// <param name="offset">
		/// The offset into the buffer from which to start adding
		/// </param>
		/// <param name="length">
		/// The number of header bytes to add
		/// </param>
		/// <returns>
		/// The index of the next free byte in the buffer
		/// </returns>
		/// TODO: what should be default behavior?(omit upper byte or UTF8?)
		[Obsolete("No Encoding for Name field is specified, any non-ASCII bytes will be discarded")]
		public static int GetNameBytes(StringBuilder name, byte[] buffer, int offset, int length)
		{
			return GetNameBytes(name, buffer, offset, length, null);
		}

		/// <summary>
		/// Add an entry name to the buffer
		/// </summary>
		/// <param name="name">
		/// The name to add
		/// </param>
		/// <param name="buffer">
		/// The buffer to add to
		/// </param>
		/// <param name="offset">
		/// The offset into the buffer from which to start adding
		/// </param>
		/// <param name="length">
		/// The number of header bytes to add
		/// </param>
		/// <param name="encoding">
		/// </param>
		/// <returns>
		/// The index of the next free byte in the buffer
		/// </returns>
		public static int GetNameBytes(StringBuilder name, byte[] buffer, int offset, int length, Encoding encoding)
		{
			if (name == null)
			{
				throw new ArgumentNullException(nameof(name));
			}

			if (buffer == null)
			{
				throw new ArgumentNullException(nameof(buffer));
			}

			return GetNameBytes(name.ToString(), 0, buffer, offset, length, encoding);
		}

		/// <summary>
		/// Add an entry name to the buffer
		/// </summary>
		/// <param name="name">The name to add</param>
		/// <param name="buffer">The buffer to add to</param>
		/// <param name="offset">The offset into the buffer from which to start adding</param>
		/// <param name="length">The number of header bytes to add</param>
		/// <returns>The index of the next free byte in the buffer</returns>
		/// TODO: what should be default behavior?(omit upper byte or UTF8?)
		[Obsolete("No Encoding for Name field is specified, any non-ASCII bytes will be discarded")]
		public static int GetNameBytes(string name, byte[] buffer, int offset, int length)
		{
			return GetNameBytes(name, buffer, offset, length, null);
		}

		/// <summary>
		/// Add an entry name to the buffer
		/// </summary>
		/// <param name="name">The name to add</param>
		/// <param name="buffer">The buffer to add to</param>
		/// <param name="offset">The offset into the buffer from which to start adding</param>
		/// <param name="length">The number of header bytes to add</param>
		/// <param name="encoding"></param>
		/// <returns>The index of the next free byte in the buffer</returns>
		public static int GetNameBytes(string name, byte[] buffer, int offset, int length, Encoding encoding)
		{
			if (name == null)
			{
				throw new ArgumentNullException(nameof(name));
			}

			if (buffer == null)
			{
				throw new ArgumentNullException(nameof(buffer));
			}

			return GetNameBytes(name, 0, buffer, offset, length, encoding);
		}
		/// <summary>
		/// Add a string to a buffer as a collection of ascii bytes.
		/// </summary>
		/// <param name="toAdd">The string to add</param>
		/// <param name="nameOffset">The offset of the first character to add.</param>
		/// <param name="buffer">The buffer to add to.</param>
		/// <param name="bufferOffset">The offset to start adding at.</param>
		/// <param name="length">The number of ascii characters to add.</param>
		/// <returns>The next free index in the buffer.</returns>
		[Obsolete("No Encoding for Name field is specified, any non-ASCII bytes will be discarded")]
		public static int GetAsciiBytes(string toAdd, int nameOffset, byte[] buffer, int bufferOffset, int length)
		{
			return GetAsciiBytes(toAdd, nameOffset, buffer, bufferOffset, length, null);
		}

		/// <summary>
		/// Add a string to a buffer as a collection of ascii bytes.
		/// </summary>
		/// <param name="toAdd">The string to add</param>
		/// <param name="nameOffset">The offset of the first character to add.</param>
		/// <param name="buffer">The buffer to add to.</param>
		/// <param name="bufferOffset">The offset to start adding at.</param>
		/// <param name="length">The number of ascii characters to add.</param>
		/// <param name="encoding">String encoding, or null for ASCII only</param>
		/// <returns>The next free index in the buffer.</returns>
		public static int GetAsciiBytes(string toAdd, int nameOffset, byte[] buffer, int bufferOffset, int length, Encoding encoding)
		{
			if (toAdd == null)
			{
				throw new ArgumentNullException(nameof(toAdd));
			}

			if (buffer == null)
			{
				throw new ArgumentNullException(nameof(buffer));
			}

			int i;
			if(encoding == null)
			{
				for (i = 0; i < length && nameOffset + i < toAdd.Length; ++i)
				{
					buffer[bufferOffset + i] = (byte)toAdd[nameOffset + i];
				}
			}
			else
			{
				// It can be more sufficient if using unsafe code or Span(ToCharArray can be omitted)
				var chars = toAdd.ToCharArray();
				var bytes = encoding.GetBytes(chars, nameOffset, Math.Min(toAdd.Length - nameOffset, length));
				i = Math.Min(bytes.Length, length);
				Array.Copy(bytes, 0, buffer, bufferOffset, i);
			}
			// If length is beyond the toAdd string length (which is OK by the prev loop condition), eg if a field has fixed length and the string is shorter, make sure all of the extra chars are written as NULLs, so that the reader func would ignore them and get back the original string
			for (; i < length; ++i)
				buffer[bufferOffset + i] = 0;
			return bufferOffset + length;
		}

		/// <summary>
		/// Put an octal representation of a value into a buffer
		/// </summary>
		/// <param name = "value">
		/// the value to be converted to octal
		/// </param>
		/// <param name = "buffer">
		/// buffer to store the octal string
		/// </param>
		/// <param name = "offset">
		/// The offset into the buffer where the value starts
		/// </param>
		/// <param name = "length">
		/// The length of the octal string to create
		/// </param>
		/// <returns>
		/// The offset of the character next byte after the octal string
		/// </returns>
		public static int GetOctalBytes(long value, byte[] buffer, int offset, int length)
		{
			if (buffer == null)
			{
				throw new ArgumentNullException(nameof(buffer));
			}

			int localIndex = length - 1;

			// Either a space or null is valid here.  We use NULL as per GNUTar
			buffer[offset + localIndex] = 0;
			--localIndex;

			if (value > 0)
			{
				for (long v = value; (localIndex >= 0) && (v > 0); --localIndex)
				{
					buffer[offset + localIndex] = (byte)((byte)'0' + (byte)(v & 7));
					v >>= 3;
				}
			}

			for (; localIndex >= 0; --localIndex)
			{
				buffer[offset + localIndex] = (byte)'0';
			}

			return offset + length;
		}

		/// <summary>
		/// Put an octal or binary representation of a value into a buffer
		/// </summary>
		/// <param name = "value">Value to be convert to octal</param>
		/// <param name = "buffer">The buffer to update</param>
		/// <param name = "offset">The offset into the buffer to store the value</param>
		/// <param name = "length">The length of the octal string. Must be 12.</param>
		/// <returns>Index of next byte</returns>
		private static int GetBinaryOrOctalBytes(long value, byte[] buffer, int offset, int length)
		{
			if (value > 0x1FFFFFFFF)
			{  // Octal 77777777777 (11 digits)
			   // Put value as binary, right-justified into the buffer. Set high order bit of left-most byte.
				for (int pos = length - 1; pos > 0; pos--)
				{
					buffer[offset + pos] = (byte)value;
					value = value >> 8;
				}
				buffer[offset] = 0x80;
				return offset + length;
			}
			return GetOctalBytes(value, buffer, offset, length);
		}

		/// <summary>
		/// Add the checksum integer to header buffer.
		/// </summary>
		/// <param name = "value"></param>
		/// <param name = "buffer">The header buffer to set the checksum for</param>
		/// <param name = "offset">The offset into the buffer for the checksum</param>
		/// <param name = "length">The number of header bytes to update.
		/// It's formatted differently from the other fields: it has 6 digits, a
		/// null, then a space -- rather than digits, a space, then a null.
		/// The final space is already there, from checksumming
		/// </param>
		/// <returns>The modified buffer offset</returns>
		private static void GetCheckSumOctalBytes(long value, byte[] buffer, int offset, int length)
		{
			GetOctalBytes(value, buffer, offset, length - 1);
		}

		/// <summary>
		/// Compute the checksum for a tar entry header.
		/// The checksum field must be all spaces prior to this happening
		/// </summary>
		/// <param name = "buffer">The tar entry's header buffer.</param>
		/// <returns>The computed checksum.</returns>
		private static int ComputeCheckSum(byte[] buffer)
		{
			int sum = 0;
			for (int i = 0; i < buffer.Length; ++i)
			{
				sum += buffer[i];
			}
			return sum;
		}

		/// <summary>
		/// Make a checksum for a tar entry ignoring the checksum contents.
		/// </summary>
		/// <param name = "buffer">The tar entry's header buffer.</param>
		/// <returns>The checksum for the buffer</returns>
		private static int MakeCheckSum(byte[] buffer)
		{
			int sum = 0;
			for (int i = 0; i < CHKSUMOFS; ++i)
			{
				sum += buffer[i];
			}

			for (int i = 0; i < CHKSUMLEN; ++i)
			{
				sum += (byte)' ';
			}

			for (int i = CHKSUMOFS + CHKSUMLEN; i < buffer.Length; ++i)
			{
				sum += buffer[i];
			}
			return sum;
		}

		private static int GetCTime(DateTime dateTime)
		{
			return unchecked((int)((dateTime.Ticks - dateTime1970.Ticks) / timeConversionFactor));
		}

		private static DateTime GetDateTimeFromCTime(long ticks)
		{
			DateTime result;

			try
			{
				result = new DateTime(dateTime1970.Ticks + ticks * timeConversionFactor);
			}
			catch (ArgumentOutOfRangeException)
			{
				result = dateTime1970;
			}
			return result;
		}

		#region Instance Fields

		private string name;
		private int mode;
		private int userId;
		private int groupId;
		private long size;
		private DateTime modTime;
		private int checksum;
		private bool isChecksumValid;
		private byte typeFlag;
		private string linkName;
		private string magic;
		private string version;
		private string userName;
		private string groupName;
		private int devMajor;
		private int devMinor;

		#endregion Instance Fields

		#region Class Fields

		// Values used during recursive operations.
		static internal int userIdAsSet;

		static internal int groupIdAsSet;
		static internal string userNameAsSet;
		static internal string groupNameAsSet = "None";

		static internal int defaultUserId;
		static internal int defaultGroupId;
		static internal string defaultGroupName = "None";
		static internal string defaultUser;

		#endregion Class Fields
	}
}


namespace ICSharpCode.SharpZipLib.Tar
{
	/// <summary>
	/// The TarInputStream reads a UNIX tar archive as an InputStream.
	/// methods are provided to position at each successive entry in
	/// the archive, and the read each entry as a normal input stream
	/// using read().
	/// </summary>
	public class TarInputStream : Stream
	{
		#region Constructors

		/// <summary>
		/// Construct a TarInputStream with default block factor
		/// </summary>
		/// <param name="inputStream">stream to source data from</param>
		[Obsolete("No Encoding for Name field is specified, any non-ASCII bytes will be discarded")]
		public TarInputStream(Stream inputStream)
			: this(inputStream, TarBuffer.DefaultBlockFactor, null)
		{
		}
		/// <summary>
		/// Construct a TarInputStream with default block factor
		/// </summary>
		/// <param name="inputStream">stream to source data from</param>
		/// <param name="nameEncoding">The <see cref="Encoding"/> used for the Name fields, or null for ASCII only</param>
		public TarInputStream(Stream inputStream, Encoding nameEncoding)
			: this(inputStream, TarBuffer.DefaultBlockFactor, nameEncoding)
		{
		}

		/// <summary>
		/// Construct a TarInputStream with user specified block factor
		/// </summary>
		/// <param name="inputStream">stream to source data from</param>
		/// <param name="blockFactor">block factor to apply to archive</param>
		[Obsolete("No Encoding for Name field is specified, any non-ASCII bytes will be discarded")]
		public TarInputStream(Stream inputStream, int blockFactor)
		{
			this.inputStream = inputStream;
			tarBuffer = TarBuffer.CreateInputTarBuffer(inputStream, blockFactor);
			encoding = null;
		}

		/// <summary>
		/// Construct a TarInputStream with user specified block factor
		/// </summary>
		/// <param name="inputStream">stream to source data from</param>
		/// <param name="blockFactor">block factor to apply to archive</param>
		/// <param name="nameEncoding">The <see cref="Encoding"/> used for the Name fields, or null for ASCII only</param>
		public TarInputStream(Stream inputStream, int blockFactor, Encoding nameEncoding)
		{
			this.inputStream = inputStream;
			tarBuffer = TarBuffer.CreateInputTarBuffer(inputStream, blockFactor);
			encoding = nameEncoding;
		}

		#endregion Constructors

		/// <summary>
		/// Gets or sets a flag indicating ownership of underlying stream.
		/// When the flag is true <see cref="Stream.Dispose()" /> will close the underlying stream also.
		/// </summary>
		/// <remarks>The default value is true.</remarks>
		public bool IsStreamOwner
		{
			get { return tarBuffer.IsStreamOwner; }
			set { tarBuffer.IsStreamOwner = value; }
		}

		#region Stream Overrides

		/// <summary>
		/// Gets a value indicating whether the current stream supports reading
		/// </summary>
		public override bool CanRead
		{
			get
			{
				return inputStream.CanRead;
			}
		}

		/// <summary>
		/// Gets a value indicating whether the current stream supports seeking
		/// This property always returns false.
		/// </summary>
		public override bool CanSeek
		{
			get
			{
				return false;
			}
		}

		/// <summary>
		/// Gets a value indicating if the stream supports writing.
		/// This property always returns false.
		/// </summary>
		public override bool CanWrite
		{
			get
			{
				return false;
			}
		}

		/// <summary>
		/// The length in bytes of the stream
		/// </summary>
		public override long Length
		{
			get
			{
				return inputStream.Length;
			}
		}

		/// <summary>
		/// Gets or sets the position within the stream.
		/// Setting the Position is not supported and throws a NotSupportedExceptionNotSupportedException
		/// </summary>
		/// <exception cref="NotSupportedException">Any attempt to set position</exception>
		public override long Position
		{
			get
			{
				return inputStream.Position;
			}
			set
			{
				throw new NotSupportedException("TarInputStream Seek not supported");
			}
		}

		/// <summary>
		/// Flushes the baseInputStream
		/// </summary>
		public override void Flush()
		{
			inputStream.Flush();
		}

		/// <summary>
		/// Set the streams position.  This operation is not supported and will throw a NotSupportedException
		/// </summary>
		/// <param name="offset">The offset relative to the origin to seek to.</param>
		/// <param name="origin">The <see cref="SeekOrigin"/> to start seeking from.</param>
		/// <returns>The new position in the stream.</returns>
		/// <exception cref="NotSupportedException">Any access</exception>
		public override long Seek(long offset, SeekOrigin origin)
		{
			throw new NotSupportedException("TarInputStream Seek not supported");
		}

		/// <summary>
		/// Sets the length of the stream
		/// This operation is not supported and will throw a NotSupportedException
		/// </summary>
		/// <param name="value">The new stream length.</param>
		/// <exception cref="NotSupportedException">Any access</exception>
		public override void SetLength(long value)
		{
			throw new NotSupportedException("TarInputStream SetLength not supported");
		}

		/// <summary>
		/// Writes a block of bytes to this stream using data from a buffer.
		/// This operation is not supported and will throw a NotSupportedException
		/// </summary>
		/// <param name="buffer">The buffer containing bytes to write.</param>
		/// <param name="offset">The offset in the buffer of the frist byte to write.</param>
		/// <param name="count">The number of bytes to write.</param>
		/// <exception cref="NotSupportedException">Any access</exception>
		public override void Write(byte[] buffer, int offset, int count)
		{
			throw new NotSupportedException("TarInputStream Write not supported");
		}

		/// <summary>
		/// Writes a byte to the current position in the file stream.
		/// This operation is not supported and will throw a NotSupportedException
		/// </summary>
		/// <param name="value">The byte value to write.</param>
		/// <exception cref="NotSupportedException">Any access</exception>
		public override void WriteByte(byte value)
		{
			throw new NotSupportedException("TarInputStream WriteByte not supported");
		}

		/// <summary>
		/// Reads a byte from the current tar archive entry.
		/// </summary>
		/// <returns>A byte cast to an int; -1 if the at the end of the stream.</returns>
		public override int ReadByte()
		{
			byte[] oneByteBuffer = new byte[1];
			int num = Read(oneByteBuffer, 0, 1);
			if (num <= 0)
			{
				// return -1 to indicate that no byte was read.
				return -1;
			}
			return oneByteBuffer[0];
		}

		/// <summary>
		/// Reads bytes from the current tar archive entry.
		///
		/// This method is aware of the boundaries of the current
		/// entry in the archive and will deal with them appropriately
		/// </summary>
		/// <param name="buffer">
		/// The buffer into which to place bytes read.
		/// </param>
		/// <param name="offset">
		/// The offset at which to place bytes read.
		/// </param>
		/// <param name="count">
		/// The number of bytes to read.
		/// </param>
		/// <returns>
		/// The number of bytes read, or 0 at end of stream/EOF.
		/// </returns>
		public override int Read(byte[] buffer, int offset, int count)
		{
			if (buffer == null)
			{
				throw new ArgumentNullException(nameof(buffer));
			}

			int totalRead = 0;

			if (entryOffset >= entrySize)
			{
				return 0;
			}

			long numToRead = count;

			if ((numToRead + entryOffset) > entrySize)
			{
				numToRead = entrySize - entryOffset;
			}

			if (readBuffer != null)
			{
				int sz = (numToRead > readBuffer.Length) ? readBuffer.Length : (int)numToRead;

				Array.Copy(readBuffer, 0, buffer, offset, sz);

				if (sz >= readBuffer.Length)
				{
					readBuffer = null;
				}
				else
				{
					int newLen = readBuffer.Length - sz;
					byte[] newBuf = new byte[newLen];
					Array.Copy(readBuffer, sz, newBuf, 0, newLen);
					readBuffer = newBuf;
				}

				totalRead += sz;
				numToRead -= sz;
				offset += sz;
			}

			while (numToRead > 0)
			{
				byte[] rec = tarBuffer.ReadBlock();
				if (rec == null)
				{
					// Unexpected EOF!
					throw new TarException("unexpected EOF with " + numToRead + " bytes unread");
				}

				var sz = (int)numToRead;
				int recLen = rec.Length;

				if (recLen > sz)
				{
					Array.Copy(rec, 0, buffer, offset, sz);
					readBuffer = new byte[recLen - sz];
					Array.Copy(rec, sz, readBuffer, 0, recLen - sz);
				}
				else
				{
					sz = recLen;
					Array.Copy(rec, 0, buffer, offset, recLen);
				}

				totalRead += sz;
				numToRead -= sz;
				offset += sz;
			}

			entryOffset += totalRead;

			return totalRead;
		}

		/// <summary>
		/// Closes this stream. Calls the TarBuffer's close() method.
		/// The underlying stream is closed by the TarBuffer.
		/// </summary>
		protected override void Dispose(bool disposing)
		{
			if (disposing)
			{
				tarBuffer.Close();
			}
		}

		#endregion Stream Overrides

		/// <summary>
		/// Set the entry factory for this instance.
		/// </summary>
		/// <param name="factory">The factory for creating new entries</param>
		public void SetEntryFactory(IEntryFactory factory)
		{
			entryFactory = factory;
		}

		/// <summary>
		/// Get the record size being used by this stream's TarBuffer.
		/// </summary>
		public int RecordSize
		{
			get { return tarBuffer.RecordSize; }
		}

		/// <summary>
		/// Get the record size being used by this stream's TarBuffer.
		/// </summary>
		/// <returns>
		/// TarBuffer record size.
		/// </returns>
		[Obsolete("Use RecordSize property instead")]
		public int GetRecordSize()
		{
			return tarBuffer.RecordSize;
		}

		/// <summary>
		/// Get the available data that can be read from the current
		/// entry in the archive. This does not indicate how much data
		/// is left in the entire archive, only in the current entry.
		/// This value is determined from the entry's size header field
		/// and the amount of data already read from the current entry.
		/// </summary>
		/// <returns>
		/// The number of available bytes for the current entry.
		/// </returns>
		public long Available
		{
			get
			{
				return entrySize - entryOffset;
			}
		}

		/// <summary>
		/// Skip bytes in the input buffer. This skips bytes in the
		/// current entry's data, not the entire archive, and will
		/// stop at the end of the current entry's data if the number
		/// to skip extends beyond that point.
		/// </summary>
		/// <param name="skipCount">
		/// The number of bytes to skip.
		/// </param>
		public void Skip(long skipCount)
		{
			// TODO: REVIEW efficiency of TarInputStream.Skip
			// This is horribly inefficient, but it ensures that we
			// properly skip over bytes via the TarBuffer...
			//
			byte[] skipBuf = new byte[8 * 1024];

			for (long num = skipCount; num > 0;)
			{
				int toRead = num > skipBuf.Length ? skipBuf.Length : (int)num;
				int numRead = Read(skipBuf, 0, toRead);

				if (numRead == -1)
				{
					break;
				}

				num -= numRead;
			}
		}

		/// <summary>
		/// Return a value of true if marking is supported; false otherwise.
		/// </summary>
		/// <remarks>Currently marking is not supported, the return value is always false.</remarks>
		public bool IsMarkSupported
		{
			get
			{
				return false;
			}
		}

		/// <summary>
		/// Since we do not support marking just yet, we do nothing.
		/// </summary>
		/// <param name ="markLimit">
		/// The limit to mark.
		/// </param>
		public void Mark(int markLimit)
		{
		}

		/// <summary>
		/// Since we do not support marking just yet, we do nothing.
		/// </summary>
		public void Reset()
		{
		}

		/// <summary>
		/// Get the next entry in this tar archive. This will skip
		/// over any remaining data in the current entry, if there
		/// is one, and place the input stream at the header of the
		/// next entry, and read the header and instantiate a new
		/// TarEntry from the header bytes and return that entry.
		/// If there are no more entries in the archive, null will
		/// be returned to indicate that the end of the archive has
		/// been reached.
		/// </summary>
		/// <returns>
		/// The next TarEntry in the archive, or null.
		/// </returns>
		public TarEntry GetNextEntry()
		{
			if (hasHitEOF)
			{
				return null;
			}

			if (currentEntry != null)
			{
				SkipToNextEntry();
			}

			byte[] headerBuf = tarBuffer.ReadBlock();

			if (headerBuf == null)
			{
				hasHitEOF = true;
			}
			else if (TarBuffer.IsEndOfArchiveBlock(headerBuf))
			{
				hasHitEOF = true;

				// Read the second zero-filled block
				tarBuffer.ReadBlock();
			}
			else
			{
				hasHitEOF = false;
			}

			if (hasHitEOF)
			{
				currentEntry = null;
			}
			else
			{
				try
				{
					var header = new TarHeader();
					header.ParseBuffer(headerBuf, encoding);
					if (!header.IsChecksumValid)
					{
						throw new TarException("Header checksum is invalid");
					}
					this.entryOffset = 0;
					this.entrySize = header.Size;

					StringBuilder longName = null;

					if (header.TypeFlag == TarHeader.LF_GNU_LONGNAME)
					{
						byte[] nameBuffer = new byte[TarBuffer.BlockSize];
						long numToRead = this.entrySize;

						longName = new StringBuilder();

						while (numToRead > 0)
						{
							int numRead = this.Read(nameBuffer, 0, (numToRead > nameBuffer.Length ? nameBuffer.Length : (int)numToRead));

							if (numRead == -1)
							{
								throw new InvalidHeaderException("Failed to read long name entry");
							}

							longName.Append(TarHeader.ParseName(nameBuffer, 0, numRead, encoding).ToString());
							numToRead -= numRead;
						}

						SkipToNextEntry();
						headerBuf = this.tarBuffer.ReadBlock();
					}
					else if (header.TypeFlag == TarHeader.LF_GHDR)
					{  // POSIX global extended header
					   // Ignore things we dont understand completely for now
						SkipToNextEntry();
						headerBuf = this.tarBuffer.ReadBlock();
					}
					else if (header.TypeFlag == TarHeader.LF_XHDR)
					{  // POSIX extended header
						byte[] nameBuffer = new byte[TarBuffer.BlockSize];
						long numToRead = this.entrySize;

						var xhr = new TarExtendedHeaderReader();

						while (numToRead > 0)
						{
							int numRead = this.Read(nameBuffer, 0, (numToRead > nameBuffer.Length ? nameBuffer.Length : (int)numToRead));

							if (numRead == -1)
							{
								throw new InvalidHeaderException("Failed to read long name entry");
							}

							xhr.Read(nameBuffer, numRead);
							numToRead -= numRead;
						}

						if (xhr.Headers.TryGetValue("path", out string name))
						{
							longName = new StringBuilder(name);
						}

						SkipToNextEntry();
						headerBuf = this.tarBuffer.ReadBlock();
					}
					else if (header.TypeFlag == TarHeader.LF_GNU_VOLHDR)
					{
						// TODO: could show volume name when verbose
						SkipToNextEntry();
						headerBuf = this.tarBuffer.ReadBlock();
					}
					else if (header.TypeFlag != TarHeader.LF_NORMAL &&
							 header.TypeFlag != TarHeader.LF_OLDNORM &&
							 header.TypeFlag != TarHeader.LF_LINK &&
							 header.TypeFlag != TarHeader.LF_SYMLINK &&
							 header.TypeFlag != TarHeader.LF_DIR)
					{
						// Ignore things we dont understand completely for now
						SkipToNextEntry();
						headerBuf = tarBuffer.ReadBlock();
					}

					if (entryFactory == null)
					{
						currentEntry = new TarEntry(headerBuf, encoding);
						if (longName != null)
						{
							currentEntry.Name = longName.ToString();
						}
					}
					else
					{
						currentEntry = entryFactory.CreateEntry(headerBuf);
					}

					// Magic was checked here for 'ustar' but there are multiple valid possibilities
					// so this is not done anymore.

					entryOffset = 0;

					// TODO: Review How do we resolve this discrepancy?!
					entrySize = this.currentEntry.Size;
				}
				catch (InvalidHeaderException ex)
				{
					entrySize = 0;
					entryOffset = 0;
					currentEntry = null;
					string errorText = string.Format("Bad header in record {0} block {1} {2}",
						tarBuffer.CurrentRecord, tarBuffer.CurrentBlock, ex.Message);
					throw new InvalidHeaderException(errorText);
				}
			}
			return currentEntry;
		}

		/// <summary>
		/// Copies the contents of the current tar archive entry directly into
		/// an output stream.
		/// </summary>
		/// <param name="outputStream">
		/// The OutputStream into which to write the entry's data.
		/// </param>
		public void CopyEntryContents(Stream outputStream)
		{
			byte[] tempBuffer = new byte[32 * 1024];

			while (true)
			{
				int numRead = Read(tempBuffer, 0, tempBuffer.Length);
				if (numRead <= 0)
				{
					break;
				}
				outputStream.Write(tempBuffer, 0, numRead);
			}
		}

		private void SkipToNextEntry()
		{
			long numToSkip = entrySize - entryOffset;

			if (numToSkip > 0)
			{
				Skip(numToSkip);
			}

			readBuffer = null;
		}

		/// <summary>
		/// This interface is provided, along with the method <see cref="SetEntryFactory"/>, to allow
		/// the programmer to have their own <see cref="TarEntry"/> subclass instantiated for the
		/// entries return from <see cref="GetNextEntry"/>.
		/// </summary>
		public interface IEntryFactory
		{
			// This interface does not considering name encoding.
			// How this interface should be?
			/// <summary>
			/// Create an entry based on name alone
			/// </summary>
			/// <param name="name">
			/// Name of the new EntryPointNotFoundException to create
			/// </param>
			/// <returns>created TarEntry or descendant class</returns>
			TarEntry CreateEntry(string name);

			/// <summary>
			/// Create an instance based on an actual file
			/// </summary>
			/// <param name="fileName">
			/// Name of file to represent in the entry
			/// </param>
			/// <returns>
			/// Created TarEntry or descendant class
			/// </returns>
			TarEntry CreateEntryFromFile(string fileName);

			/// <summary>
			/// Create a tar entry based on the header information passed
			/// </summary>
			/// <param name="headerBuffer">
			/// Buffer containing header information to create an entry from.
			/// </param>
			/// <returns>
			/// Created TarEntry or descendant class
			/// </returns>
			TarEntry CreateEntry(byte[] headerBuffer);
		}

		/// <summary>
		/// Standard entry factory class creating instances of the class TarEntry
		/// </summary>
		public class EntryFactoryAdapter : IEntryFactory
		{
			Encoding nameEncoding;
			/// <summary>
			/// Construct standard entry factory class with ASCII name encoding
			/// </summary>
			[Obsolete("No Encoding for Name field is specified, any non-ASCII bytes will be discarded")]
			public EntryFactoryAdapter()
			{
			}
			/// <summary>
			/// Construct standard entry factory with name encoding
			/// </summary>
			/// <param name="nameEncoding">The <see cref="Encoding"/> used for the Name fields, or null for ASCII only</param>
			public EntryFactoryAdapter(Encoding nameEncoding)
			{
				this.nameEncoding = nameEncoding;
			}
			/// <summary>
			/// Create a <see cref="TarEntry"/> based on named
			/// </summary>
			/// <param name="name">The name to use for the entry</param>
			/// <returns>A new <see cref="TarEntry"/></returns>
			public TarEntry CreateEntry(string name)
			{
				return TarEntry.CreateTarEntry(name);
			}

			/// <summary>
			/// Create a tar entry with details obtained from <paramref name="fileName">file</paramref>
			/// </summary>
			/// <param name="fileName">The name of the file to retrieve details from.</param>
			/// <returns>A new <see cref="TarEntry"/></returns>
			public TarEntry CreateEntryFromFile(string fileName)
			{
				return TarEntry.CreateEntryFromFile(fileName);
			}

			/// <summary>
			/// Create an entry based on details in <paramref name="headerBuffer">header</paramref>
			/// </summary>
			/// <param name="headerBuffer">The buffer containing entry details.</param>
			/// <returns>A new <see cref="TarEntry"/></returns>
			public TarEntry CreateEntry(byte[] headerBuffer)
			{
				return new TarEntry(headerBuffer, nameEncoding);
			}
		}

		#region Instance Fields

		/// <summary>
		/// Flag set when last block has been read
		/// </summary>
		protected bool hasHitEOF;

		/// <summary>
		/// Size of this entry as recorded in header
		/// </summary>
		protected long entrySize;

		/// <summary>
		/// Number of bytes read for this entry so far
		/// </summary>
		protected long entryOffset;

		/// <summary>
		/// Buffer used with calls to <code>Read()</code>
		/// </summary>
		protected byte[] readBuffer;

		/// <summary>
		/// Working buffer
		/// </summary>
		protected TarBuffer tarBuffer;

		/// <summary>
		/// Current entry being read
		/// </summary>
		private TarEntry currentEntry;

		/// <summary>
		/// Factory used to create TarEntry or descendant class instance
		/// </summary>
		protected IEntryFactory entryFactory;

		/// <summary>
		/// Stream used as the source of input data.
		/// </summary>
		private readonly Stream inputStream;

		private readonly Encoding encoding;

		#endregion Instance Fields
	}
}


namespace ICSharpCode.SharpZipLib.Tar
{
	/// <summary>
	/// The TarOutputStream writes a UNIX tar archive as an OutputStream.
	/// Methods are provided to put entries, and then write their contents
	/// by writing to this stream using write().
	/// </summary>
	/// public
	public class TarOutputStream : Stream
	{
		#region Constructors

		/// <summary>
		/// Construct TarOutputStream using default block factor
		/// </summary>
		/// <param name="outputStream">stream to write to</param>
		[Obsolete("No Encoding for Name field is specified, any non-ASCII bytes will be discarded")]
		public TarOutputStream(Stream outputStream)
			: this(outputStream, TarBuffer.DefaultBlockFactor)
		{
		}

		/// <summary>
		/// Construct TarOutputStream using default block factor
		/// </summary>
		/// <param name="outputStream">stream to write to</param>
		/// <param name="nameEncoding">The <see cref="Encoding"/> used for the Name fields, or null for ASCII only</param>
		public TarOutputStream(Stream outputStream, Encoding nameEncoding)
			: this(outputStream, TarBuffer.DefaultBlockFactor, nameEncoding)
		{
		}

		/// <summary>
		/// Construct TarOutputStream with user specified block factor
		/// </summary>
		/// <param name="outputStream">stream to write to</param>
		/// <param name="blockFactor">blocking factor</param>
		[Obsolete("No Encoding for Name field is specified, any non-ASCII bytes will be discarded")]
		public TarOutputStream(Stream outputStream, int blockFactor)
		{
			if (outputStream == null)
			{
				throw new ArgumentNullException(nameof(outputStream));
			}

			this.outputStream = outputStream;
			buffer = TarBuffer.CreateOutputTarBuffer(outputStream, blockFactor);

			assemblyBuffer = new byte[TarBuffer.BlockSize];
			blockBuffer = new byte[TarBuffer.BlockSize];
		}

		/// <summary>
		/// Construct TarOutputStream with user specified block factor
		/// </summary>
		/// <param name="outputStream">stream to write to</param>
		/// <param name="blockFactor">blocking factor</param>
		/// <param name="nameEncoding">The <see cref="Encoding"/> used for the Name fields, or null for ASCII only</param>
		public TarOutputStream(Stream outputStream, int blockFactor, Encoding nameEncoding)
		{
			if (outputStream == null)
			{
				throw new ArgumentNullException(nameof(outputStream));
			}

			this.outputStream = outputStream;
			buffer = TarBuffer.CreateOutputTarBuffer(outputStream, blockFactor);

			assemblyBuffer = new byte[TarBuffer.BlockSize];
			blockBuffer = new byte[TarBuffer.BlockSize];

			this.nameEncoding = nameEncoding;
		}

		#endregion Constructors

		/// <summary>
		/// Gets or sets a flag indicating ownership of underlying stream.
		/// When the flag is true <see cref="Stream.Dispose()" /> will close the underlying stream also.
		/// </summary>
		/// <remarks>The default value is true.</remarks>
		public bool IsStreamOwner
		{
			get { return buffer.IsStreamOwner; }
			set { buffer.IsStreamOwner = value; }
		}

		/// <summary>
		/// true if the stream supports reading; otherwise, false.
		/// </summary>
		public override bool CanRead
		{
			get
			{
				return outputStream.CanRead;
			}
		}

		/// <summary>
		/// true if the stream supports seeking; otherwise, false.
		/// </summary>
		public override bool CanSeek
		{
			get
			{
				return outputStream.CanSeek;
			}
		}

		/// <summary>
		/// true if stream supports writing; otherwise, false.
		/// </summary>
		public override bool CanWrite
		{
			get
			{
				return outputStream.CanWrite;
			}
		}

		/// <summary>
		/// length of stream in bytes
		/// </summary>
		public override long Length
		{
			get
			{
				return outputStream.Length;
			}
		}

		/// <summary>
		/// gets or sets the position within the current stream.
		/// </summary>
		public override long Position
		{
			get
			{
				return outputStream.Position;
			}
			set
			{
				outputStream.Position = value;
			}
		}

		/// <summary>
		/// set the position within the current stream
		/// </summary>
		/// <param name="offset">The offset relative to the <paramref name="origin"/> to seek to</param>
		/// <param name="origin">The <see cref="SeekOrigin"/> to seek from.</param>
		/// <returns>The new position in the stream.</returns>
		public override long Seek(long offset, SeekOrigin origin)
		{
			return outputStream.Seek(offset, origin);
		}

		/// <summary>
		/// Set the length of the current stream
		/// </summary>
		/// <param name="value">The new stream length.</param>
		public override void SetLength(long value)
		{
			outputStream.SetLength(value);
		}

		/// <summary>
		/// Read a byte from the stream and advance the position within the stream
		/// by one byte or returns -1 if at the end of the stream.
		/// </summary>
		/// <returns>The byte value or -1 if at end of stream</returns>
		public override int ReadByte()
		{
			return outputStream.ReadByte();
		}

		/// <summary>
		/// read bytes from the current stream and advance the position within the
		/// stream by the number of bytes read.
		/// </summary>
		/// <param name="buffer">The buffer to store read bytes in.</param>
		/// <param name="offset">The index into the buffer to being storing bytes at.</param>
		/// <param name="count">The desired number of bytes to read.</param>
		/// <returns>The total number of bytes read, or zero if at the end of the stream.
		/// The number of bytes may be less than the <paramref name="count">count</paramref>
		/// requested if data is not available.</returns>
		public override int Read(byte[] buffer, int offset, int count)
		{
			return outputStream.Read(buffer, offset, count);
		}

		/// <summary>
		/// All buffered data is written to destination
		/// </summary>
		public override void Flush()
		{
			outputStream.Flush();
		}

		/// <summary>
		/// Ends the TAR archive without closing the underlying OutputStream.
		/// The result is that the EOF block of nulls is written.
		/// </summary>
		public void Finish()
		{
			if (IsEntryOpen)
			{
				CloseEntry();
			}
			WriteEofBlock();
		}

		/// <summary>
		/// Ends the TAR archive and closes the underlying OutputStream.
		/// </summary>
		/// <remarks>This means that Finish() is called followed by calling the
		/// TarBuffer's Close().</remarks>
		protected override void Dispose(bool disposing)
		{
			if (!isClosed)
			{
				isClosed = true;
				Finish();
				buffer.Close();
			}
		}

		/// <summary>
		/// Get the record size being used by this stream's TarBuffer.
		/// </summary>
		public int RecordSize
		{
			get { return buffer.RecordSize; }
		}

		/// <summary>
		/// Get the record size being used by this stream's TarBuffer.
		/// </summary>
		/// <returns>
		/// The TarBuffer record size.
		/// </returns>
		[Obsolete("Use RecordSize property instead")]
		public int GetRecordSize()
		{
			return buffer.RecordSize;
		}

		/// <summary>
		/// Get a value indicating whether an entry is open, requiring more data to be written.
		/// </summary>
		private bool IsEntryOpen
		{
			get { return (currBytes < currSize); }
		}

		/// <summary>
		/// Put an entry on the output stream. This writes the entry's
		/// header and positions the output stream for writing
		/// the contents of the entry. Once this method is called, the
		/// stream is ready for calls to write() to write the entry's
		/// contents. Once the contents are written, closeEntry()
		/// <B>MUST</B> be called to ensure that all buffered data
		/// is completely written to the output stream.
		/// </summary>
		/// <param name="entry">
		/// The TarEntry to be written to the archive.
		/// </param>
		public void PutNextEntry(TarEntry entry)
		{
			if (entry == null)
			{
				throw new ArgumentNullException(nameof(entry));
			}

			var namelen = nameEncoding != null ? nameEncoding.GetByteCount(entry.TarHeader.Name) : entry.TarHeader.Name.Length;

			if (namelen > TarHeader.NAMELEN)
			{
				var longHeader = new TarHeader();
				longHeader.TypeFlag = TarHeader.LF_GNU_LONGNAME;
				longHeader.Name = longHeader.Name + "././@LongLink";
				longHeader.Mode = 420;//644 by default
				longHeader.UserId = entry.UserId;
				longHeader.GroupId = entry.GroupId;
				longHeader.GroupName = entry.GroupName;
				longHeader.UserName = entry.UserName;
				longHeader.LinkName = "";
				longHeader.Size = namelen + 1;  // Plus one to avoid dropping last char

				longHeader.WriteHeader(blockBuffer, nameEncoding);
				buffer.WriteBlock(blockBuffer);  // Add special long filename header block

				int nameCharIndex = 0;

				while (nameCharIndex < namelen + 1 /* we've allocated one for the null char, now we must make sure it gets written out */)
				{
					Array.Clear(blockBuffer, 0, blockBuffer.Length);
					TarHeader.GetAsciiBytes(entry.TarHeader.Name, nameCharIndex, this.blockBuffer, 0, TarBuffer.BlockSize, nameEncoding); // This func handles OK the extra char out of string length
					nameCharIndex += TarBuffer.BlockSize;
					buffer.WriteBlock(blockBuffer);
				}
			}

			entry.WriteEntryHeader(blockBuffer, nameEncoding);
			buffer.WriteBlock(blockBuffer);

			currBytes = 0;

			currSize = entry.IsDirectory ? 0 : entry.Size;
		}

		/// <summary>
		/// Close an entry. This method MUST be called for all file
		/// entries that contain data. The reason is that we must
		/// buffer data written to the stream in order to satisfy
		/// the buffer's block based writes. Thus, there may be
		/// data fragments still being assembled that must be written
		/// to the output stream before this entry is closed and the
		/// next entry written.
		/// </summary>
		public void CloseEntry()
		{
			if (assemblyBufferLength > 0)
			{
				Array.Clear(assemblyBuffer, assemblyBufferLength, assemblyBuffer.Length - assemblyBufferLength);

				buffer.WriteBlock(assemblyBuffer);

				currBytes += assemblyBufferLength;
				assemblyBufferLength = 0;
			}

			if (currBytes < currSize)
			{
				string errorText = string.Format(
					"Entry closed at '{0}' before the '{1}' bytes specified in the header were written",
					currBytes, currSize);
				throw new TarException(errorText);
			}
		}

		/// <summary>
		/// Writes a byte to the current tar archive entry.
		/// This method simply calls Write(byte[], int, int).
		/// </summary>
		/// <param name="value">
		/// The byte to be written.
		/// </param>
		public override void WriteByte(byte value)
		{
			Write(new byte[] { value }, 0, 1);
		}

		/// <summary>
		/// Writes bytes to the current tar archive entry. This method
		/// is aware of the current entry and will throw an exception if
		/// you attempt to write bytes past the length specified for the
		/// current entry. The method is also (painfully) aware of the
		/// record buffering required by TarBuffer, and manages buffers
		/// that are not a multiple of recordsize in length, including
		/// assembling records from small buffers.
		/// </summary>
		/// <param name = "buffer">
		/// The buffer to write to the archive.
		/// </param>
		/// <param name = "offset">
		/// The offset in the buffer from which to get bytes.
		/// </param>
		/// <param name = "count">
		/// The number of bytes to write.
		/// </param>
		public override void Write(byte[] buffer, int offset, int count)
		{
			if (buffer == null)
			{
				throw new ArgumentNullException(nameof(buffer));
			}

			if (offset < 0)
			{
				throw new ArgumentOutOfRangeException(nameof(offset), "Cannot be negative");
			}

			if (buffer.Length - offset < count)
			{
				throw new ArgumentException("offset and count combination is invalid");
			}

			if (count < 0)
			{
				throw new ArgumentOutOfRangeException(nameof(count), "Cannot be negative");
			}

			if ((currBytes + count) > currSize)
			{
				string errorText = string.Format("request to write '{0}' bytes exceeds size in header of '{1}' bytes",
					count, this.currSize);
				throw new ArgumentOutOfRangeException(nameof(count), errorText);
			}

			//
			// We have to deal with assembly!!!
			// The programmer can be writing little 32 byte chunks for all
			// we know, and we must assemble complete blocks for writing.
			// TODO  REVIEW Maybe this should be in TarBuffer? Could that help to
			//        eliminate some of the buffer copying.
			//
			if (assemblyBufferLength > 0)
			{
				if ((assemblyBufferLength + count) >= blockBuffer.Length)
				{
					int aLen = blockBuffer.Length - assemblyBufferLength;

					Array.Copy(assemblyBuffer, 0, blockBuffer, 0, assemblyBufferLength);
					Array.Copy(buffer, offset, blockBuffer, assemblyBufferLength, aLen);

					this.buffer.WriteBlock(blockBuffer);

					currBytes += blockBuffer.Length;

					offset += aLen;
					count -= aLen;

					assemblyBufferLength = 0;
				}
				else
				{
					Array.Copy(buffer, offset, assemblyBuffer, assemblyBufferLength, count);
					offset += count;
					assemblyBufferLength += count;
					count -= count;
				}
			}

			//
			// When we get here we have EITHER:
			//   o An empty "assembly" buffer.
			//   o No bytes to write (count == 0)
			//
			while (count > 0)
			{
				if (count < blockBuffer.Length)
				{
					Array.Copy(buffer, offset, assemblyBuffer, assemblyBufferLength, count);
					assemblyBufferLength += count;
					break;
				}

				this.buffer.WriteBlock(buffer, offset);

				int bufferLength = blockBuffer.Length;
				currBytes += bufferLength;
				count -= bufferLength;
				offset += bufferLength;
			}
		}

		/// <summary>
		/// Write an EOF (end of archive) block to the tar archive.
		/// The	end of the archive is indicated	by two blocks consisting entirely of zero bytes.
		/// </summary>
		private void WriteEofBlock()
		{
			Array.Clear(blockBuffer, 0, blockBuffer.Length);
			buffer.WriteBlock(blockBuffer);
			buffer.WriteBlock(blockBuffer);
		}

		#region Instance Fields

		/// <summary>
		/// bytes written for this entry so far
		/// </summary>
		private long currBytes;

		/// <summary>
		/// current 'Assembly' buffer length
		/// </summary>
		private int assemblyBufferLength;

		/// <summary>
		/// Flag indicating whether this instance has been closed or not.
		/// </summary>
		private bool isClosed;

		/// <summary>
		/// Size for the current entry
		/// </summary>
		protected long currSize;

		/// <summary>
		/// single block working buffer
		/// </summary>
		protected byte[] blockBuffer;

		/// <summary>
		/// 'Assembly' buffer used to assemble data before writing
		/// </summary>
		protected byte[] assemblyBuffer;

		/// <summary>
		/// TarBuffer used to provide correct blocking factor
		/// </summary>
		protected TarBuffer buffer;

		/// <summary>
		/// the destination stream for the archive contents
		/// </summary>
		protected Stream outputStream;

		/// <summary>
		/// name encoding
		/// </summary>
		protected Encoding nameEncoding;

		#endregion Instance Fields
	}
}

namespace ICSharpCode.SharpZipLib.Zip
{
	/// <summary>
	/// FastZipEvents supports all events applicable to <see cref="FastZip">FastZip</see> operations.
	/// </summary>
	public class FastZipEvents
	{
		/// <summary>
		/// Delegate to invoke when processing directories.
		/// </summary>
		public event EventHandler<DirectoryEventArgs> ProcessDirectory;

		/// <summary>
		/// Delegate to invoke when processing files.
		/// </summary>
		public ProcessFileHandler ProcessFile;

		/// <summary>
		/// Delegate to invoke during processing of files.
		/// </summary>
		public ProgressHandler Progress;

		/// <summary>
		/// Delegate to invoke when processing for a file has been completed.
		/// </summary>
		public CompletedFileHandler CompletedFile;

		/// <summary>
		/// Delegate to invoke when processing directory failures.
		/// </summary>
		public DirectoryFailureHandler DirectoryFailure;

		/// <summary>
		/// Delegate to invoke when processing file failures.
		/// </summary>
		public FileFailureHandler FileFailure;

		/// <summary>
		/// Raise the <see cref="DirectoryFailure">directory failure</see> event.
		/// </summary>
		/// <param name="directory">The directory causing the failure.</param>
		/// <param name="e">The exception for this event.</param>
		/// <returns>A boolean indicating if execution should continue or not.</returns>
		public bool OnDirectoryFailure(string directory, Exception e)
		{
			bool result = false;
			DirectoryFailureHandler handler = DirectoryFailure;

			if (handler != null)
			{
				var args = new ScanFailureEventArgs(directory, e);
				handler(this, args);
				result = args.ContinueRunning;
			}
			return result;
		}

		/// <summary>
		/// Fires the <see cref="FileFailure"> file failure handler delegate</see>.
		/// </summary>
		/// <param name="file">The file causing the failure.</param>
		/// <param name="e">The exception for this failure.</param>
		/// <returns>A boolean indicating if execution should continue or not.</returns>
		public bool OnFileFailure(string file, Exception e)
		{
			FileFailureHandler handler = FileFailure;
			bool result = (handler != null);

			if (result)
			{
				var args = new ScanFailureEventArgs(file, e);
				handler(this, args);
				result = args.ContinueRunning;
			}
			return result;
		}

		/// <summary>
		/// Fires the <see cref="ProcessFile">ProcessFile delegate</see>.
		/// </summary>
		/// <param name="file">The file being processed.</param>
		/// <returns>A boolean indicating if execution should continue or not.</returns>
		public bool OnProcessFile(string file)
		{
			bool result = true;
			ProcessFileHandler handler = ProcessFile;

			if (handler != null)
			{
				var args = new ScanEventArgs(file);
				handler(this, args);
				result = args.ContinueRunning;
			}
			return result;
		}

		/// <summary>
		/// Fires the <see cref="CompletedFile"/> delegate
		/// </summary>
		/// <param name="file">The file whose processing has been completed.</param>
		/// <returns>A boolean indicating if execution should continue or not.</returns>
		public bool OnCompletedFile(string file)
		{
			bool result = true;
			CompletedFileHandler handler = CompletedFile;
			if (handler != null)
			{
				var args = new ScanEventArgs(file);
				handler(this, args);
				result = args.ContinueRunning;
			}
			return result;
		}

		/// <summary>
		/// Fires the <see cref="ProcessDirectory">process directory</see> delegate.
		/// </summary>
		/// <param name="directory">The directory being processed.</param>
		/// <param name="hasMatchingFiles">Flag indicating if the directory has matching files as determined by the current filter.</param>
		/// <returns>A <see cref="bool"/> of true if the operation should continue; false otherwise.</returns>
		public bool OnProcessDirectory(string directory, bool hasMatchingFiles)
		{
			bool result = true;
			EventHandler<DirectoryEventArgs> handler = ProcessDirectory;
			if (handler != null)
			{
				var args = new DirectoryEventArgs(directory, hasMatchingFiles);
				handler(this, args);
				result = args.ContinueRunning;
			}
			return result;
		}

		/// <summary>
		/// The minimum timespan between <see cref="Progress"/> events.
		/// </summary>
		/// <value>The minimum period of time between <see cref="Progress"/> events.</value>
		/// <seealso cref="Progress"/>
		/// <remarks>The default interval is three seconds.</remarks>
		public TimeSpan ProgressInterval
		{
			get { return progressInterval_; }
			set { progressInterval_ = value; }
		}

		#region Instance Fields

		private TimeSpan progressInterval_ = TimeSpan.FromSeconds(3);

		#endregion Instance Fields
	}

	/// <summary>
	/// FastZip provides facilities for creating and extracting zip files.
	/// </summary>
	public class FastZip
	{
		#region Enumerations

		/// <summary>
		/// Defines the desired handling when overwriting files during extraction.
		/// </summary>
		public enum Overwrite
		{
			/// <summary>
			/// Prompt the user to confirm overwriting
			/// </summary>
			Prompt,

			/// <summary>
			/// Never overwrite files.
			/// </summary>
			Never,

			/// <summary>
			/// Always overwrite files.
			/// </summary>
			Always
		}

		#endregion Enumerations

		#region Constructors

		/// <summary>
		/// Initialise a default instance of <see cref="FastZip"/>.
		/// </summary>
		public FastZip()
		{
		}

		/// <summary>
		/// Initialise a new instance of <see cref="FastZip"/> using the specified <see cref="TimeSetting"/>
		/// </summary>
		/// <param name="timeSetting">The <see cref="TimeSetting">time setting</see> to use when creating or extracting <see cref="ZipEntry">Zip entries</see>.</param>
		/// <remarks>Using <see cref="TimeSetting.LastAccessTime">TimeSetting.LastAccessTime</see><see cref="TimeSetting.LastAccessTimeUtc">[Utc]</see> when
		/// creating an archive will set the file time to the moment of reading.
		/// </remarks>
		public FastZip(TimeSetting timeSetting)
		{
			entryFactory_ = new ZipEntryFactory(timeSetting);
			restoreDateTimeOnExtract_ = true;
		}

		/// <summary>
		/// Initialise a new instance of <see cref="FastZip"/> using the specified <see cref="DateTime"/>
		/// </summary>
		/// <param name="time">The time to set all <see cref="ZipEntry.DateTime"/> values for created or extracted <see cref="ZipEntry">Zip Entries</see>.</param>
		public FastZip(DateTime time)
		{
			entryFactory_ = new ZipEntryFactory(time);
			restoreDateTimeOnExtract_ = true;
		}

		/// <summary>
		/// Initialise a new instance of <see cref="FastZip"/>
		/// </summary>
		/// <param name="events">The <see cref="FastZipEvents">events</see> to use during operations.</param>
		public FastZip(FastZipEvents events)
		{
			events_ = events;
		}

		#endregion Constructors

		#region Properties

		/// <summary>
		/// Get/set a value indicating whether empty directories should be created.
		/// </summary>
		public bool CreateEmptyDirectories
		{
			get { return createEmptyDirectories_; }
			set { createEmptyDirectories_ = value; }
		}

		/// <summary>
		/// Get / set the password value.
		/// </summary>
		public string Password
		{
			get { return password_; }
			set { password_ = value; }
		}

		/// <summary>
		/// Get / set the method of encrypting entries.
		/// </summary>
		/// <remarks>
		/// Only applies when <see cref="Password"/> is set.
		/// Defaults to ZipCrypto for backwards compatibility purposes.
		/// </remarks>
		public ZipEncryptionMethod EntryEncryptionMethod { get; set; } = ZipEncryptionMethod.ZipCrypto;

		/// <summary>
		/// Get or set the <see cref="INameTransform"></see> active when creating Zip files.
		/// </summary>
		/// <seealso cref="EntryFactory"></seealso>
		public INameTransform NameTransform
		{
			get { return entryFactory_.NameTransform; }
			set
			{
				entryFactory_.NameTransform = value;
			}
		}

		/// <summary>
		/// Get or set the <see cref="IEntryFactory"></see> active when creating Zip files.
		/// </summary>
		public IEntryFactory EntryFactory
		{
			get { return entryFactory_; }
			set
			{
				if (value == null)
				{
					entryFactory_ = new ZipEntryFactory();
				}
				else
				{
					entryFactory_ = value;
				}
			}
		}

		/// <summary>
		/// Gets or sets the setting for <see cref="UseZip64">Zip64 handling when writing.</see>
		/// </summary>
		/// <remarks>
		/// The default value is dynamic which is not backwards compatible with old
		/// programs and can cause problems with XP's built in compression which cant
		/// read Zip64 archives. However it does avoid the situation were a large file
		/// is added and cannot be completed correctly.
		/// NOTE: Setting the size for entries before they are added is the best solution!
		/// By default the EntryFactory used by FastZip will set the file size.
		/// </remarks>
		public UseZip64 UseZip64
		{
			get { return useZip64_; }
			set { useZip64_ = value; }
		}

		/// <summary>
		/// Get/set a value indicating whether file dates and times should
		/// be restored when extracting files from an archive.
		/// </summary>
		/// <remarks>The default value is false.</remarks>
		public bool RestoreDateTimeOnExtract
		{
			get
			{
				return restoreDateTimeOnExtract_;
			}
			set
			{
				restoreDateTimeOnExtract_ = value;
			}
		}

		/// <summary>
		/// Get/set a value indicating whether file attributes should
		/// be restored during extract operations
		/// </summary>
		public bool RestoreAttributesOnExtract
		{
			get { return restoreAttributesOnExtract_; }
			set { restoreAttributesOnExtract_ = value; }
		}

		/// <summary>
		/// Get/set the Compression Level that will be used
		/// when creating the zip
		/// </summary>
		public Deflater.CompressionLevel CompressionLevel
		{
			get { return compressionLevel_; }
			set { compressionLevel_ = value; }
		}

		#endregion Properties

		#region Delegates

		/// <summary>
		/// Delegate called when confirming overwriting of files.
		/// </summary>
		public delegate bool ConfirmOverwriteDelegate(string fileName);

		#endregion Delegates

		#region CreateZip

		/// <summary>
		/// Create a zip file.
		/// </summary>
		/// <param name="zipFileName">The name of the zip file to create.</param>
		/// <param name="sourceDirectory">The directory to source files from.</param>
		/// <param name="recurse">True to recurse directories, false for no recursion.</param>
		/// <param name="fileFilter">The <see cref="PathFilter">file filter</see> to apply.</param>
		/// <param name="directoryFilter">The <see cref="PathFilter">directory filter</see> to apply.</param>
		public void CreateZip(string zipFileName, string sourceDirectory,
			bool recurse, string fileFilter, string directoryFilter)
		{
			CreateZip(File.Create(zipFileName), sourceDirectory, recurse, fileFilter, directoryFilter);
		}

		/// <summary>
		/// Create a zip file/archive.
		/// </summary>
		/// <param name="zipFileName">The name of the zip file to create.</param>
		/// <param name="sourceDirectory">The directory to obtain files and directories from.</param>
		/// <param name="recurse">True to recurse directories, false for no recursion.</param>
		/// <param name="fileFilter">The file filter to apply.</param>
		public void CreateZip(string zipFileName, string sourceDirectory, bool recurse, string fileFilter)
		{
			CreateZip(File.Create(zipFileName), sourceDirectory, recurse, fileFilter, null);
		}

		/// <summary>
		/// Create a zip archive sending output to the <paramref name="outputStream"/> passed.
		/// </summary>
		/// <param name="outputStream">The stream to write archive data to.</param>
		/// <param name="sourceDirectory">The directory to source files from.</param>
		/// <param name="recurse">True to recurse directories, false for no recursion.</param>
		/// <param name="fileFilter">The <see cref="PathFilter">file filter</see> to apply.</param>
		/// <param name="directoryFilter">The <see cref="PathFilter">directory filter</see> to apply.</param>
		/// <remarks>The <paramref name="outputStream"/> is closed after creation.</remarks>
		public void CreateZip(Stream outputStream, string sourceDirectory, bool recurse, string fileFilter, string directoryFilter)
		{
			CreateZip(outputStream, sourceDirectory, recurse, fileFilter, directoryFilter, false);
		}

		/// <summary>
		/// Create a zip archive sending output to the <paramref name="outputStream"/> passed.
		/// </summary>
		/// <param name="outputStream">The stream to write archive data to.</param>
		/// <param name="sourceDirectory">The directory to source files from.</param>
		/// <param name="recurse">True to recurse directories, false for no recursion.</param>
		/// <param name="fileFilter">The <see cref="PathFilter">file filter</see> to apply.</param>
		/// <param name="directoryFilter">The <see cref="PathFilter">directory filter</see> to apply.</param>
		/// <param name="leaveOpen">true to leave <paramref name="outputStream"/> open after the zip has been created, false to dispose it.</param>
		public void CreateZip(Stream outputStream, string sourceDirectory, bool recurse, string fileFilter, string directoryFilter, bool leaveOpen)
		{
			var scanner = new FileSystemScanner(fileFilter, directoryFilter);
			CreateZip(outputStream, sourceDirectory, recurse, scanner, leaveOpen);
		}

		/// <summary>
		/// Create a zip file.
		/// </summary>
		/// <param name="zipFileName">The name of the zip file to create.</param>
		/// <param name="sourceDirectory">The directory to source files from.</param>
		/// <param name="recurse">True to recurse directories, false for no recursion.</param>
		/// <param name="fileFilter">The <see cref="IScanFilter">file filter</see> to apply.</param>
		/// <param name="directoryFilter">The <see cref="IScanFilter">directory filter</see> to apply.</param>
		public void CreateZip(string zipFileName, string sourceDirectory,
			bool recurse, IScanFilter fileFilter, IScanFilter directoryFilter)
		{
			CreateZip(File.Create(zipFileName), sourceDirectory, recurse, fileFilter, directoryFilter, false);
		}

		/// <summary>
		/// Create a zip archive sending output to the <paramref name="outputStream"/> passed.
		/// </summary>
		/// <param name="outputStream">The stream to write archive data to.</param>
		/// <param name="sourceDirectory">The directory to source files from.</param>
		/// <param name="recurse">True to recurse directories, false for no recursion.</param>
		/// <param name="fileFilter">The <see cref="IScanFilter">file filter</see> to apply.</param>
		/// <param name="directoryFilter">The <see cref="IScanFilter">directory filter</see> to apply.</param>
		/// <param name="leaveOpen">true to leave <paramref name="outputStream"/> open after the zip has been created, false to dispose it.</param>
		public void CreateZip(Stream outputStream, string sourceDirectory, bool recurse, IScanFilter fileFilter, IScanFilter directoryFilter, bool leaveOpen = false)
		{
			var scanner = new FileSystemScanner(fileFilter, directoryFilter);
			CreateZip(outputStream, sourceDirectory, recurse, scanner, leaveOpen);
		}

		/// <summary>
		/// Create a zip archive sending output to the <paramref name="outputStream"/> passed.
		/// </summary>
		/// <param name="outputStream">The stream to write archive data to.</param>
		/// <param name="sourceDirectory">The directory to source files from.</param>
		/// <param name="recurse">True to recurse directories, false for no recursion.</param>
		/// <param name="scanner">For performing the actual file system scan</param>
		/// <param name="leaveOpen">true to leave <paramref name="outputStream"/> open after the zip has been created, false to dispose it.</param>
		/// <remarks>The <paramref name="outputStream"/> is closed after creation.</remarks>
		private void CreateZip(Stream outputStream, string sourceDirectory, bool recurse, FileSystemScanner scanner, bool leaveOpen)
		{
			NameTransform = new ZipNameTransform(sourceDirectory);
			sourceDirectory_ = sourceDirectory;

			using (outputStream_ = new ZipOutputStream(outputStream))
			{
				outputStream_.SetLevel((int)CompressionLevel);
				outputStream_.IsStreamOwner = !leaveOpen;
				outputStream_.NameTransform = null; // all required transforms handled by us

				if (false == string.IsNullOrEmpty(password_) && EntryEncryptionMethod != ZipEncryptionMethod.None)
				{
					outputStream_.Password = password_;
				}

				outputStream_.UseZip64 = UseZip64;
				scanner.ProcessFile += ProcessFile;
				if (this.CreateEmptyDirectories)
				{
					scanner.ProcessDirectory += ProcessDirectory;
				}

				if (events_ != null)
				{
					if (events_.FileFailure != null)
					{
						scanner.FileFailure += events_.FileFailure;
					}

					if (events_.DirectoryFailure != null)
					{
						scanner.DirectoryFailure += events_.DirectoryFailure;
					}
				}

				scanner.Scan(sourceDirectory, recurse);
			}
		}

		#endregion CreateZip

		#region ExtractZip

		/// <summary>
		/// Extract the contents of a zip file.
		/// </summary>
		/// <param name="zipFileName">The zip file to extract from.</param>
		/// <param name="targetDirectory">The directory to save extracted information in.</param>
		/// <param name="fileFilter">A filter to apply to files.</param>
		public void ExtractZip(string zipFileName, string targetDirectory, string fileFilter)
		{
			ExtractZip(zipFileName, targetDirectory, Overwrite.Always, null, fileFilter, null, restoreDateTimeOnExtract_);
		}

		/// <summary>
		/// Extract the contents of a zip file.
		/// </summary>
		/// <param name="zipFileName">The zip file to extract from.</param>
		/// <param name="targetDirectory">The directory to save extracted information in.</param>
		/// <param name="overwrite">The style of <see cref="Overwrite">overwriting</see> to apply.</param>
		/// <param name="confirmDelegate">A delegate to invoke when confirming overwriting.</param>
		/// <param name="fileFilter">A filter to apply to files.</param>
		/// <param name="directoryFilter">A filter to apply to directories.</param>
		/// <param name="restoreDateTime">Flag indicating whether to restore the date and time for extracted files.</param>
		/// <param name="allowParentTraversal">Allow parent directory traversal in file paths (e.g. ../file)</param>
		public void ExtractZip(string zipFileName, string targetDirectory,
							   Overwrite overwrite, ConfirmOverwriteDelegate confirmDelegate,
							   string fileFilter, string directoryFilter, bool restoreDateTime, bool allowParentTraversal = false)
		{
			Stream inputStream = File.Open(zipFileName, FileMode.Open, FileAccess.Read, FileShare.Read);
			ExtractZip(inputStream, targetDirectory, overwrite, confirmDelegate, fileFilter, directoryFilter, restoreDateTime, true, allowParentTraversal);
		}

		/// <summary>
		/// Extract the contents of a zip file held in a stream.
		/// </summary>
		/// <param name="inputStream">The seekable input stream containing the zip to extract from.</param>
		/// <param name="targetDirectory">The directory to save extracted information in.</param>
		/// <param name="overwrite">The style of <see cref="Overwrite">overwriting</see> to apply.</param>
		/// <param name="confirmDelegate">A delegate to invoke when confirming overwriting.</param>
		/// <param name="fileFilter">A filter to apply to files.</param>
		/// <param name="directoryFilter">A filter to apply to directories.</param>
		/// <param name="restoreDateTime">Flag indicating whether to restore the date and time for extracted files.</param>
		/// <param name="isStreamOwner">Flag indicating whether the inputStream will be closed by this method.</param>
		/// <param name="allowParentTraversal">Allow parent directory traversal in file paths (e.g. ../file)</param>
		public void ExtractZip(Stream inputStream, string targetDirectory,
					   Overwrite overwrite, ConfirmOverwriteDelegate confirmDelegate,
					   string fileFilter, string directoryFilter, bool restoreDateTime,
					   bool isStreamOwner, bool allowParentTraversal = false)
		{
			if ((overwrite == Overwrite.Prompt) && (confirmDelegate == null))
			{
				throw new ArgumentNullException(nameof(confirmDelegate));
			}

			continueRunning_ = true;
			overwrite_ = overwrite;
			confirmDelegate_ = confirmDelegate;
			extractNameTransform_ = new WindowsNameTransform(targetDirectory, allowParentTraversal);

			fileFilter_ = new NameFilter(fileFilter);
			directoryFilter_ = new NameFilter(directoryFilter);
			restoreDateTimeOnExtract_ = restoreDateTime;

			using (zipFile_ = new ZipFile(inputStream, !isStreamOwner))
			{
				if (password_ != null)
				{
					zipFile_.Password = password_;
				}

				System.Collections.IEnumerator enumerator = zipFile_.GetEnumerator();
				while (continueRunning_ && enumerator.MoveNext())
				{
					var entry = (ZipEntry)enumerator.Current;
					if (entry.IsFile)
					{
						// TODO Path.GetDirectory can fail here on invalid characters.
						if (directoryFilter_.IsMatch(Path.GetDirectoryName(entry.Name)) && fileFilter_.IsMatch(entry.Name))
						{
							ExtractEntry(entry);
						}
					}
					else if (entry.IsDirectory)
					{
						if (directoryFilter_.IsMatch(entry.Name) && CreateEmptyDirectories)
						{
							ExtractEntry(entry);
						}
					}
					else
					{
						// Do nothing for volume labels etc...
					}
				}
			}
		}

		#endregion ExtractZip

		#region Internal Processing

		private void ProcessDirectory(object sender, DirectoryEventArgs e)
		{
			if (!e.HasMatchingFiles && CreateEmptyDirectories)
			{
				if (events_ != null)
				{
					events_.OnProcessDirectory(e.Name, e.HasMatchingFiles);
				}

				if (e.ContinueRunning)
				{
					if (e.Name != sourceDirectory_)
					{
						ZipEntry entry = entryFactory_.MakeDirectoryEntry(e.Name);
						outputStream_.PutNextEntry(entry);
					}
				}
			}
		}

		private void ProcessFile(object sender, ScanEventArgs e)
		{
			if ((events_ != null) && (events_.ProcessFile != null))
			{
				events_.ProcessFile(sender, e);
			}

			if (e.ContinueRunning)
			{
				try
				{
					// The open below is equivalent to OpenRead which guarantees that if opened the
					// file will not be changed by subsequent openers, but precludes opening in some cases
					// were it could succeed. ie the open may fail as its already open for writing and the share mode should reflect that.
					using (FileStream stream = File.Open(e.Name, FileMode.Open, FileAccess.Read, FileShare.Read))
					{
						ZipEntry entry = entryFactory_.MakeFileEntry(e.Name);

						// Set up AES encryption for the entry if required.
						ConfigureEntryEncryption(entry);

						outputStream_.PutNextEntry(entry);
						AddFileContents(e.Name, stream);
					}
				}
				catch (Exception ex)
				{
					if (events_ != null)
					{
						continueRunning_ = events_.OnFileFailure(e.Name, ex);
					}
					else
					{
						continueRunning_ = false;
						throw;
					}
				}
			}
		}

		// Set up the encryption method to use for the specific entry.
		private void ConfigureEntryEncryption(ZipEntry entry)
		{
			// Only alter the entries options if AES isn't already enabled for it
			// (it might have been set up by the entry factory, and if so we let that take precedence)
			if (!string.IsNullOrEmpty(Password) && entry.AESEncryptionStrength == 0)
			{
				switch (EntryEncryptionMethod)
				{
					case ZipEncryptionMethod.AES128:
						entry.AESKeySize = 128;
						break;

					case ZipEncryptionMethod.AES256:
						entry.AESKeySize = 256;
						break;
				}
			}
		}

		private void AddFileContents(string name, Stream stream)
		{
			if (stream == null)
			{
				throw new ArgumentNullException(nameof(stream));
			}

			if (buffer_ == null)
			{
				buffer_ = new byte[4096];
			}

			if ((events_ != null) && (events_.Progress != null))
			{
				StreamUtils.Copy(stream, outputStream_, buffer_,
					events_.Progress, events_.ProgressInterval, this, name);
			}
			else
			{
				StreamUtils.Copy(stream, outputStream_, buffer_);
			}

			if (events_ != null)
			{
				continueRunning_ = events_.OnCompletedFile(name);
			}
		}

		private void ExtractFileEntry(ZipEntry entry, string targetName)
		{
			bool proceed = true;
			if (overwrite_ != Overwrite.Always)
			{
				if (File.Exists(targetName))
				{
					if ((overwrite_ == Overwrite.Prompt) && (confirmDelegate_ != null))
					{
						proceed = confirmDelegate_(targetName);
					}
					else
					{
						proceed = false;
					}
				}
			}

			if (proceed)
			{
				if (events_ != null)
				{
					continueRunning_ = events_.OnProcessFile(entry.Name);
				}

				if (continueRunning_)
				{
					try
					{
						using (FileStream outputStream = File.Create(targetName))
						{
							if (buffer_ == null)
							{
								buffer_ = new byte[4096];
							}

							using (var inputStream = zipFile_.GetInputStream(entry))
							{
								if ((events_ != null) && (events_.Progress != null))
								{
									StreamUtils.Copy(inputStream, outputStream, buffer_,
										events_.Progress, events_.ProgressInterval, this, entry.Name, entry.Size);
								}
								else
								{
									StreamUtils.Copy(inputStream, outputStream, buffer_);
								}
							}

							if (events_ != null)
							{
								continueRunning_ = events_.OnCompletedFile(entry.Name);
							}
						}

						if (restoreDateTimeOnExtract_)
						{
							switch (entryFactory_.Setting)
							{
								case TimeSetting.CreateTime:
									File.SetCreationTime(targetName, entry.DateTime);
									break;

								case TimeSetting.CreateTimeUtc:
									File.SetCreationTimeUtc(targetName, entry.DateTime);
									break;

								case TimeSetting.LastAccessTime:
									File.SetLastAccessTime(targetName, entry.DateTime);
									break;

								case TimeSetting.LastAccessTimeUtc:
									File.SetLastAccessTimeUtc(targetName, entry.DateTime);
									break;

								case TimeSetting.LastWriteTime:
									File.SetLastWriteTime(targetName, entry.DateTime);
									break;

								case TimeSetting.LastWriteTimeUtc:
									File.SetLastWriteTimeUtc(targetName, entry.DateTime);
									break;

								case TimeSetting.Fixed:
									File.SetLastWriteTime(targetName, entryFactory_.FixedDateTime);
									break;

								default:
									throw new ZipException("Unhandled time setting in ExtractFileEntry");
							}
						}

						if (RestoreAttributesOnExtract && entry.IsDOSEntry && (entry.ExternalFileAttributes != -1))
						{
							var fileAttributes = (FileAttributes)entry.ExternalFileAttributes;
							// TODO: FastZip - Setting of other file attributes on extraction is a little trickier.
							fileAttributes &= (FileAttributes.Archive | FileAttributes.Normal | FileAttributes.ReadOnly | FileAttributes.Hidden);
							File.SetAttributes(targetName, fileAttributes);
						}
					}
					catch (Exception ex)
					{
						if (events_ != null)
						{
							continueRunning_ = events_.OnFileFailure(targetName, ex);
						}
						else
						{
							continueRunning_ = false;
							throw;
						}
					}
				}
			}
		}

		private void ExtractEntry(ZipEntry entry)
		{
			bool doExtraction = entry.IsCompressionMethodSupported();
			string targetName = entry.Name;

			if (doExtraction)
			{
				if (entry.IsFile)
				{
					targetName = extractNameTransform_.TransformFile(targetName);
				}
				else if (entry.IsDirectory)
				{
					targetName = extractNameTransform_.TransformDirectory(targetName);
				}

				doExtraction = !(string.IsNullOrEmpty(targetName));
			}

			// TODO: Fire delegate/throw exception were compression method not supported, or name is invalid?

			string dirName = string.Empty;

			if (doExtraction)
			{
				if (entry.IsDirectory)
				{
					dirName = targetName;
				}
				else
				{
					dirName = Path.GetDirectoryName(Path.GetFullPath(targetName));
				}
			}

			if (doExtraction && !Directory.Exists(dirName))
			{
				if (!entry.IsDirectory || CreateEmptyDirectories)
				{
					try
					{
						continueRunning_ = events_?.OnProcessDirectory(dirName, true) ?? true;
						if (continueRunning_)
						{
							Directory.CreateDirectory(dirName);
							if (entry.IsDirectory && restoreDateTimeOnExtract_)
							{
								switch (entryFactory_.Setting)
								{
									case TimeSetting.CreateTime:
										Directory.SetCreationTime(dirName, entry.DateTime);
										break;

									case TimeSetting.CreateTimeUtc:
										Directory.SetCreationTimeUtc(dirName, entry.DateTime);
										break;

									case TimeSetting.LastAccessTime:
										Directory.SetLastAccessTime(dirName, entry.DateTime);
										break;

									case TimeSetting.LastAccessTimeUtc:
										Directory.SetLastAccessTimeUtc(dirName, entry.DateTime);
										break;

									case TimeSetting.LastWriteTime:
										Directory.SetLastWriteTime(dirName, entry.DateTime);
										break;

									case TimeSetting.LastWriteTimeUtc:
										Directory.SetLastWriteTimeUtc(dirName, entry.DateTime);
										break;

									case TimeSetting.Fixed:
										Directory.SetLastWriteTime(dirName, entryFactory_.FixedDateTime);
										break;

									default:
										throw new ZipException("Unhandled time setting in ExtractEntry");
								}
							}
						}
						else
						{
							doExtraction = false;
						}
					}
					catch (Exception ex)
					{
						doExtraction = false;
						if (events_ != null)
						{
							if (entry.IsDirectory)
							{
								continueRunning_ = events_.OnDirectoryFailure(targetName, ex);
							}
							else
							{
								continueRunning_ = events_.OnFileFailure(targetName, ex);
							}
						}
						else
						{
							continueRunning_ = false;
							throw;
						}
					}
				}
			}

			if (doExtraction && entry.IsFile)
			{
				ExtractFileEntry(entry, targetName);
			}
		}

		private static int MakeExternalAttributes(FileInfo info)
		{
			return (int)info.Attributes;
		}

		private static bool NameIsValid(string name)
		{
			return !string.IsNullOrEmpty(name) &&
				(name.IndexOfAny(Path.GetInvalidPathChars()) < 0);
		}

		#endregion Internal Processing

		#region Instance Fields

		private bool continueRunning_;
		private byte[] buffer_;
		private ZipOutputStream outputStream_;
		private ZipFile zipFile_;
		private string sourceDirectory_;
		private NameFilter fileFilter_;
		private NameFilter directoryFilter_;
		private Overwrite overwrite_;
		private ConfirmOverwriteDelegate confirmDelegate_;

		private bool restoreDateTimeOnExtract_;
		private bool restoreAttributesOnExtract_;
		private bool createEmptyDirectories_;
		private FastZipEvents events_;
		private IEntryFactory entryFactory_ = new ZipEntryFactory();
		private INameTransform extractNameTransform_;
		private UseZip64 useZip64_ = UseZip64.Dynamic;
		private CompressionLevel compressionLevel_ = CompressionLevel.DEFAULT_COMPRESSION;

		private string password_;

		#endregion Instance Fields
	}
}


namespace ICSharpCode.SharpZipLib.Zip
{
	/// <summary>
	/// Defines factory methods for creating new <see cref="ZipEntry"></see> values.
	/// </summary>
	public interface IEntryFactory
	{
		/// <summary>
		/// Create a <see cref="ZipEntry"/> for a file given its name
		/// </summary>
		/// <param name="fileName">The name of the file to create an entry for.</param>
		/// <returns>Returns a <see cref="ZipEntry">file entry</see> based on the <paramref name="fileName"/> passed.</returns>
		ZipEntry MakeFileEntry(string fileName);

		/// <summary>
		/// Create a <see cref="ZipEntry"/> for a file given its name
		/// </summary>
		/// <param name="fileName">The name of the file to create an entry for.</param>
		/// <param name="useFileSystem">If true get details from the file system if the file exists.</param>
		/// <returns>Returns a <see cref="ZipEntry">file entry</see> based on the <paramref name="fileName"/> passed.</returns>
		ZipEntry MakeFileEntry(string fileName, bool useFileSystem);

		/// <summary>
		/// Create a <see cref="ZipEntry"/> for a file given its actual name and optional override name
		/// </summary>
		/// <param name="fileName">The name of the file to create an entry for.</param>
		/// <param name="entryName">An alternative name to be used for the new entry. Null if not applicable.</param>
		/// <param name="useFileSystem">If true get details from the file system if the file exists.</param>
		/// <returns>Returns a <see cref="ZipEntry">file entry</see> based on the <paramref name="fileName"/> passed.</returns>
		ZipEntry MakeFileEntry(string fileName, string entryName, bool useFileSystem);

		/// <summary>
		/// Create a <see cref="ZipEntry"/> for a directory given its name
		/// </summary>
		/// <param name="directoryName">The name of the directory to create an entry for.</param>
		/// <returns>Returns a <see cref="ZipEntry">directory entry</see> based on the <paramref name="directoryName"/> passed.</returns>
		ZipEntry MakeDirectoryEntry(string directoryName);

		/// <summary>
		/// Create a <see cref="ZipEntry"/> for a directory given its name
		/// </summary>
		/// <param name="directoryName">The name of the directory to create an entry for.</param>
		/// <param name="useFileSystem">If true get details from the file system for this directory if it exists.</param>
		/// <returns>Returns a <see cref="ZipEntry">directory entry</see> based on the <paramref name="directoryName"/> passed.</returns>
		ZipEntry MakeDirectoryEntry(string directoryName, bool useFileSystem);

		/// <summary>
		/// Get/set the <see cref="INameTransform"></see> applicable.
		/// </summary>
		INameTransform NameTransform { get; set; }

		/// <summary>
		/// Get the <see cref="TimeSetting"/> in use.
		/// </summary>
		TimeSetting Setting { get; }

		/// <summary>
		/// Get the <see cref="DateTime"/> value to use when <see cref="Setting"/> is set to <see cref="TimeSetting.Fixed"/>,
		/// or if not specified, the value of <see cref="DateTime.Now"/> when the class was the initialized
		/// </summary>
		DateTime FixedDateTime { get; }
	}
}

namespace ICSharpCode.SharpZipLib.Zip
{
	/// <summary>
	/// WindowsNameTransform transforms <see cref="ZipFile"/> names to windows compatible ones.
	/// </summary>
	public class WindowsNameTransform : INameTransform
	{
		/// <summary>
		///  The maximum windows path name permitted.
		/// </summary>
		/// <remarks>This may not valid for all windows systems - CE?, etc but I cant find the equivalent in the CLR.</remarks>
		private const int MaxPath = 260;

		private string _baseDirectory;
		private bool _trimIncomingPaths;
		private char _replacementChar = '_';
		private bool _allowParentTraversal;

		/// <summary>
		/// In this case we need Windows' invalid path characters.
		/// Path.GetInvalidPathChars() only returns a subset invalid on all platforms.
		/// </summary>
		private static readonly char[] InvalidEntryChars = new char[] {
			'"', '<', '>', '|', '\0', '\u0001', '\u0002', '\u0003', '\u0004', '\u0005',
			'\u0006', '\a', '\b', '\t', '\n', '\v', '\f', '\r', '\u000e', '\u000f',
			'\u0010', '\u0011', '\u0012', '\u0013', '\u0014', '\u0015', '\u0016',
			'\u0017', '\u0018', '\u0019', '\u001a', '\u001b', '\u001c', '\u001d',
			'\u001e', '\u001f',
			// extra characters for masks, etc.
			'*', '?', ':'
		};

		/// <summary>
		/// Initialises a new instance of <see cref="WindowsNameTransform"/>
		/// </summary>
		/// <param name="baseDirectory"></param>
		/// <param name="allowParentTraversal">Allow parent directory traversal in file paths (e.g. ../file)</param>
		public WindowsNameTransform(string baseDirectory, bool allowParentTraversal = false)
		{
			BaseDirectory = baseDirectory ?? throw new ArgumentNullException(nameof(baseDirectory), "Directory name is invalid");
			AllowParentTraversal = allowParentTraversal;
		}

		/// <summary>
		/// Initialise a default instance of <see cref="WindowsNameTransform"/>
		/// </summary>
		public WindowsNameTransform()
		{
			// Do nothing.
		}

		/// <summary>
		/// Gets or sets a value containing the target directory to prefix values with.
		/// </summary>
		public string BaseDirectory
		{
			get { return _baseDirectory; }
			set
			{
				if (value == null)
				{
					throw new ArgumentNullException(nameof(value));
				}

				_baseDirectory = Path.GetFullPath(value);
			}
		}

		/// <summary>
		/// Allow parent directory traversal in file paths (e.g. ../file)
		/// </summary>
		public bool AllowParentTraversal
		{
			get => _allowParentTraversal;
			set => _allowParentTraversal = value;
		}

		/// <summary>
		/// Gets or sets a value indicating whether paths on incoming values should be removed.
		/// </summary>
		public bool TrimIncomingPaths
		{
			get { return _trimIncomingPaths; }
			set { _trimIncomingPaths = value; }
		}

		/// <summary>
		/// Transform a Zip directory name to a windows directory name.
		/// </summary>
		/// <param name="name">The directory name to transform.</param>
		/// <returns>The transformed name.</returns>
		public string TransformDirectory(string name)
		{
			name = TransformFile(name);
			if (name.Length > 0)
			{
				while (name.EndsWith(Path.DirectorySeparatorChar.ToString(), StringComparison.Ordinal))
				{
					name = name.Remove(name.Length - 1, 1);
				}
			}
			else
			{
				throw new InvalidNameException("Cannot have an empty directory name");
			}
			return name;
		}

		/// <summary>
		/// Transform a Zip format file name to a windows style one.
		/// </summary>
		/// <param name="name">The file name to transform.</param>
		/// <returns>The transformed name.</returns>
		public string TransformFile(string name)
		{
			if (name != null)
			{
				name = MakeValidName(name, _replacementChar);

				if (_trimIncomingPaths)
				{
					name = Path.GetFileName(name);
				}

				// This may exceed windows length restrictions.
				// Combine will throw a PathTooLongException in that case.
				if (_baseDirectory != null)
				{
					name = Path.Combine(_baseDirectory, name);

					// Ensure base directory ends with directory separator ('/' or '\' depending on OS)
					var pathBase = Path.GetFullPath(_baseDirectory);
					if (pathBase[pathBase.Length - 1] != Path.DirectorySeparatorChar)
					{
						pathBase += Path.DirectorySeparatorChar;
					}

					if (!_allowParentTraversal && !Path.GetFullPath(name).StartsWith(pathBase, StringComparison.InvariantCultureIgnoreCase))
					{
						throw new InvalidNameException("Parent traversal in paths is not allowed");
					}
				}
			}
			else
			{
				name = string.Empty;
			}
			return name;
		}

		/// <summary>
		/// Test a name to see if it is a valid name for a windows filename as extracted from a Zip archive.
		/// </summary>
		/// <param name="name">The name to test.</param>
		/// <returns>Returns true if the name is a valid zip name; false otherwise.</returns>
		/// <remarks>The filename isnt a true windows path in some fundamental ways like no absolute paths, no rooted paths etc.</remarks>
		public static bool IsValidName(string name)
		{
			bool result =
				(name != null) &&
				(name.Length <= MaxPath) &&
				(string.Compare(name, MakeValidName(name, '_'), StringComparison.Ordinal) == 0)
				;

			return result;
		}

		/// <summary>
		/// Force a name to be valid by replacing invalid characters with a fixed value
		/// </summary>
		/// <param name="name">The name to make valid</param>
		/// <param name="replacement">The replacement character to use for any invalid characters.</param>
		/// <returns>Returns a valid name</returns>
		public static string MakeValidName(string name, char replacement)
		{
			if (name == null)
			{
				throw new ArgumentNullException(nameof(name));
			}

			name = PathUtils.DropPathRoot(name.Replace("/", Path.DirectorySeparatorChar.ToString()));

			// Drop any leading slashes.
			while ((name.Length > 0) && (name[0] == Path.DirectorySeparatorChar))
			{
				name = name.Remove(0, 1);
			}

			// Drop any trailing slashes.
			while ((name.Length > 0) && (name[name.Length - 1] == Path.DirectorySeparatorChar))
			{
				name = name.Remove(name.Length - 1, 1);
			}

			// Convert consecutive \\ characters to \
			int index = name.IndexOf(string.Format("{0}{0}", Path.DirectorySeparatorChar), StringComparison.Ordinal);
			while (index >= 0)
			{
				name = name.Remove(index, 1);
				index = name.IndexOf(string.Format("{0}{0}", Path.DirectorySeparatorChar), StringComparison.Ordinal);
			}

			// Convert any invalid characters using the replacement one.
			index = name.IndexOfAny(InvalidEntryChars);
			if (index >= 0)
			{
				var builder = new StringBuilder(name);

				while (index >= 0)
				{
					builder[index] = replacement;

					if (index >= name.Length)
					{
						index = -1;
					}
					else
					{
						index = name.IndexOfAny(InvalidEntryChars, index + 1);
					}
				}
				name = builder.ToString();
			}

			// Check for names greater than MaxPath characters.
			// TODO: Were is CLR version of MaxPath defined?  Can't find it in Environment.
			if (name.Length > MaxPath)
			{
				throw new PathTooLongException();
			}

			return name;
		}

		/// <summary>
		/// Gets or set the character to replace invalid characters during transformations.
		/// </summary>
		public char Replacement
		{
			get { return _replacementChar; }
			set
			{
				for (int i = 0; i < InvalidEntryChars.Length; ++i)
				{
					if (InvalidEntryChars[i] == value)
					{
						throw new ArgumentException("invalid path character");
					}
				}

				if ((value == Path.DirectorySeparatorChar) || (value == Path.AltDirectorySeparatorChar))
				{
					throw new ArgumentException("invalid replacement character");
				}

				_replacementChar = value;
			}
		}
	}
}


namespace ICSharpCode.SharpZipLib.Zip
{
	#region Enumerations

	/// <summary>
	/// Determines how entries are tested to see if they should use Zip64 extensions or not.
	/// </summary>
	public enum UseZip64
	{
		/// <summary>
		/// Zip64 will not be forced on entries during processing.
		/// </summary>
		/// <remarks>An entry can have this overridden if required <see cref="ZipEntry.ForceZip64"></see></remarks>
		Off,

		/// <summary>
		/// Zip64 should always be used.
		/// </summary>
		On,

		/// <summary>
		/// #ZipLib will determine use based on entry values when added to archive.
		/// </summary>
		Dynamic,
	}

	/// <summary>
	/// The kind of compression used for an entry in an archive
	/// </summary>
	public enum CompressionMethod
	{
		/// <summary>
		/// A direct copy of the file contents is held in the archive
		/// </summary>
		Stored = 0,

		/// <summary>
		/// Common Zip compression method using a sliding dictionary
		/// of up to 32KB and secondary compression from Huffman/Shannon-Fano trees
		/// </summary>
		Deflated = 8,

		/// <summary>
		/// An extension to deflate with a 64KB window. Not supported by #Zip currently
		/// </summary>
		Deflate64 = 9,

		/// <summary>
		/// BZip2 compression. Not supported by #Zip.
		/// </summary>
		BZip2 = 12,

		/// <summary>
		/// LZMA compression. Not supported by #Zip.
		/// </summary>
		LZMA = 14,

		/// <summary>
		/// PPMd compression. Not supported by #Zip.
		/// </summary>
		PPMd = 98,

		/// <summary>
		/// WinZip special for AES encryption, Now supported by #Zip.
		/// </summary>
		WinZipAES = 99,
	}

	/// <summary>
	/// Identifies the encryption algorithm used for an entry
	/// </summary>
	public enum EncryptionAlgorithm
	{
		/// <summary>
		/// No encryption has been used.
		/// </summary>
		None = 0,

		/// <summary>
		/// Encrypted using PKZIP 2.0 or 'classic' encryption.
		/// </summary>
		PkzipClassic = 1,

		/// <summary>
		/// DES encryption has been used.
		/// </summary>
		Des = 0x6601,

		/// <summary>
		/// RC2 encryption has been used for encryption.
		/// </summary>
		RC2 = 0x6602,

		/// <summary>
		/// Triple DES encryption with 168 bit keys has been used for this entry.
		/// </summary>
		TripleDes168 = 0x6603,

		/// <summary>
		/// Triple DES with 112 bit keys has been used for this entry.
		/// </summary>
		TripleDes112 = 0x6609,

		/// <summary>
		/// AES 128 has been used for encryption.
		/// </summary>
		Aes128 = 0x660e,

		/// <summary>
		/// AES 192 has been used for encryption.
		/// </summary>
		Aes192 = 0x660f,

		/// <summary>
		/// AES 256 has been used for encryption.
		/// </summary>
		Aes256 = 0x6610,

		/// <summary>
		/// RC2 corrected has been used for encryption.
		/// </summary>
		RC2Corrected = 0x6702,

		/// <summary>
		/// Blowfish has been used for encryption.
		/// </summary>
		Blowfish = 0x6720,

		/// <summary>
		/// Twofish has been used for encryption.
		/// </summary>
		Twofish = 0x6721,

		/// <summary>
		/// RC4 has been used for encryption.
		/// </summary>
		RC4 = 0x6801,

		/// <summary>
		/// An unknown algorithm has been used for encryption.
		/// </summary>
		Unknown = 0xffff
	}

	/// <summary>
	/// Defines the contents of the general bit flags field for an archive entry.
	/// </summary>
	[Flags]
	public enum GeneralBitFlags
	{
		/// <summary>
		/// Bit 0 if set indicates that the file is encrypted
		/// </summary>
		Encrypted = 0x0001,

		/// <summary>
		/// Bits 1 and 2 - Two bits defining the compression method (only for Method 6 Imploding and 8,9 Deflating)
		/// </summary>
		Method = 0x0006,

		/// <summary>
		/// Bit 3 if set indicates a trailing data descriptor is appended to the entry data
		/// </summary>
		Descriptor = 0x0008,

		/// <summary>
		/// Bit 4 is reserved for use with method 8 for enhanced deflation
		/// </summary>
		ReservedPKware4 = 0x0010,

		/// <summary>
		/// Bit 5 if set indicates the file contains Pkzip compressed patched data.
		/// Requires version 2.7 or greater.
		/// </summary>
		Patched = 0x0020,

		/// <summary>
		/// Bit 6 if set indicates strong encryption has been used for this entry.
		/// </summary>
		StrongEncryption = 0x0040,

		/// <summary>
		/// Bit 7 is currently unused
		/// </summary>
		Unused7 = 0x0080,

		/// <summary>
		/// Bit 8 is currently unused
		/// </summary>
		Unused8 = 0x0100,

		/// <summary>
		/// Bit 9 is currently unused
		/// </summary>
		Unused9 = 0x0200,

		/// <summary>
		/// Bit 10 is currently unused
		/// </summary>
		Unused10 = 0x0400,

		/// <summary>
		/// Bit 11 if set indicates the filename and
		/// comment fields for this file must be encoded using UTF-8.
		/// </summary>
		UnicodeText = 0x0800,

		/// <summary>
		/// Bit 12 is documented as being reserved by PKware for enhanced compression.
		/// </summary>
		EnhancedCompress = 0x1000,

		/// <summary>
		/// Bit 13 if set indicates that values in the local header are masked to hide
		/// their actual values, and the central directory is encrypted.
		/// </summary>
		/// <remarks>
		/// Used when encrypting the central directory contents.
		/// </remarks>
		HeaderMasked = 0x2000,

		/// <summary>
		/// Bit 14 is documented as being reserved for use by PKware
		/// </summary>
		ReservedPkware14 = 0x4000,

		/// <summary>
		/// Bit 15 is documented as being reserved for use by PKware
		/// </summary>
		ReservedPkware15 = 0x8000
	}

	#endregion Enumerations

	/// <summary>
	/// This class contains constants used for Zip format files
	/// </summary>
	[System.Diagnostics.CodeAnalysis.SuppressMessage("Naming", "CA1707:Identifiers should not contain underscores", Justification = "kept for backwards compatibility")]
	public static class ZipConstants
	{
		#region Versions

		/// <summary>
		/// The version made by field for entries in the central header when created by this library
		/// </summary>
		/// <remarks>
		/// This is also the Zip version for the library when comparing against the version required to extract
		/// for an entry.  See <see cref="ZipEntry.CanDecompress"/>.
		/// </remarks>
		public const int VersionMadeBy = 51; // was 45 before AES

		/// <summary>
		/// The version made by field for entries in the central header when created by this library
		/// </summary>
		/// <remarks>
		/// This is also the Zip version for the library when comparing against the version required to extract
		/// for an entry.  See <see cref="ZipInputStream.CanDecompressEntry">ZipInputStream.CanDecompressEntry</see>.
		/// </remarks>
		[Obsolete("Use VersionMadeBy instead")]
		public const int VERSION_MADE_BY = 51;

		/// <summary>
		/// The minimum version required to support strong encryption
		/// </summary>
		public const int VersionStrongEncryption = 50;

		/// <summary>
		/// The minimum version required to support strong encryption
		/// </summary>
		[Obsolete("Use VersionStrongEncryption instead")]
		public const int VERSION_STRONG_ENCRYPTION = 50;

		/// <summary>
		/// Version indicating AES encryption
		/// </summary>
		public const int VERSION_AES = 51;

		/// <summary>
		/// The version required for Zip64 extensions (4.5 or higher)
		/// </summary>
		public const int VersionZip64 = 45;

		/// <summary>
		/// The version required for BZip2 compression (4.6 or higher)
		/// </summary>
		public const int VersionBZip2 = 46;

		#endregion Versions

		#region Header Sizes

		/// <summary>
		/// Size of local entry header (excluding variable length fields at end)
		/// </summary>
		public const int LocalHeaderBaseSize = 30;

		/// <summary>
		/// Size of local entry header (excluding variable length fields at end)
		/// </summary>
		[Obsolete("Use LocalHeaderBaseSize instead")]
		public const int LOCHDR = 30;

		/// <summary>
		/// Size of Zip64 data descriptor
		/// </summary>
		public const int Zip64DataDescriptorSize = 24;

		/// <summary>
		/// Size of data descriptor
		/// </summary>
		public const int DataDescriptorSize = 16;

		/// <summary>
		/// Size of data descriptor
		/// </summary>
		[Obsolete("Use DataDescriptorSize instead")]
		public const int EXTHDR = 16;

		/// <summary>
		/// Size of central header entry (excluding variable fields)
		/// </summary>
		public const int CentralHeaderBaseSize = 46;

		/// <summary>
		/// Size of central header entry
		/// </summary>
		[Obsolete("Use CentralHeaderBaseSize instead")]
		public const int CENHDR = 46;

		/// <summary>
		/// Size of end of central record (excluding variable fields)
		/// </summary>
		public const int EndOfCentralRecordBaseSize = 22;

		/// <summary>
		/// Size of end of central record (excluding variable fields)
		/// </summary>
		[Obsolete("Use EndOfCentralRecordBaseSize instead")]
		public const int ENDHDR = 22;

		/// <summary>
		/// Size of 'classic' cryptographic header stored before any entry data
		/// </summary>
		public const int CryptoHeaderSize = 12;

		/// <summary>
		/// Size of cryptographic header stored before entry data
		/// </summary>
		[Obsolete("Use CryptoHeaderSize instead")]
		public const int CRYPTO_HEADER_SIZE = 12;

		/// <summary>
		/// The size of the Zip64 central directory locator.
		/// </summary>
		public const int Zip64EndOfCentralDirectoryLocatorSize = 20;

		#endregion Header Sizes

		#region Header Signatures

		/// <summary>
		/// Signature for local entry header
		/// </summary>
		public const int LocalHeaderSignature = 'P' | ('K' << 8) | (3 << 16) | (4 << 24);

		/// <summary>
		/// Signature for local entry header
		/// </summary>
		[Obsolete("Use LocalHeaderSignature instead")]
		public const int LOCSIG = 'P' | ('K' << 8) | (3 << 16) | (4 << 24);

		/// <summary>
		/// Signature for spanning entry
		/// </summary>
		public const int SpanningSignature = 'P' | ('K' << 8) | (7 << 16) | (8 << 24);

		/// <summary>
		/// Signature for spanning entry
		/// </summary>
		[Obsolete("Use SpanningSignature instead")]
		public const int SPANNINGSIG = 'P' | ('K' << 8) | (7 << 16) | (8 << 24);

		/// <summary>
		/// Signature for temporary spanning entry
		/// </summary>
		public const int SpanningTempSignature = 'P' | ('K' << 8) | ('0' << 16) | ('0' << 24);

		/// <summary>
		/// Signature for temporary spanning entry
		/// </summary>
		[Obsolete("Use SpanningTempSignature instead")]
		public const int SPANTEMPSIG = 'P' | ('K' << 8) | ('0' << 16) | ('0' << 24);

		/// <summary>
		/// Signature for data descriptor
		/// </summary>
		/// <remarks>
		/// This is only used where the length, Crc, or compressed size isnt known when the
		/// entry is created and the output stream doesnt support seeking.
		/// The local entry cannot be 'patched' with the correct values in this case
		/// so the values are recorded after the data prefixed by this header, as well as in the central directory.
		/// </remarks>
		public const int DataDescriptorSignature = 'P' | ('K' << 8) | (7 << 16) | (8 << 24);

		/// <summary>
		/// Signature for data descriptor
		/// </summary>
		/// <remarks>
		/// This is only used where the length, Crc, or compressed size isnt known when the
		/// entry is created and the output stream doesnt support seeking.
		/// The local entry cannot be 'patched' with the correct values in this case
		/// so the values are recorded after the data prefixed by this header, as well as in the central directory.
		/// </remarks>
		[Obsolete("Use DataDescriptorSignature instead")]
		public const int EXTSIG = 'P' | ('K' << 8) | (7 << 16) | (8 << 24);

		/// <summary>
		/// Signature for central header
		/// </summary>
		[Obsolete("Use CentralHeaderSignature instead")]
		public const int CENSIG = 'P' | ('K' << 8) | (1 << 16) | (2 << 24);

		/// <summary>
		/// Signature for central header
		/// </summary>
		public const int CentralHeaderSignature = 'P' | ('K' << 8) | (1 << 16) | (2 << 24);

		/// <summary>
		/// Signature for Zip64 central file header
		/// </summary>
		public const int Zip64CentralFileHeaderSignature = 'P' | ('K' << 8) | (6 << 16) | (6 << 24);

		/// <summary>
		/// Signature for Zip64 central file header
		/// </summary>
		[Obsolete("Use Zip64CentralFileHeaderSignature instead")]
		public const int CENSIG64 = 'P' | ('K' << 8) | (6 << 16) | (6 << 24);

		/// <summary>
		/// Signature for Zip64 central directory locator
		/// </summary>
		public const int Zip64CentralDirLocatorSignature = 'P' | ('K' << 8) | (6 << 16) | (7 << 24);

		/// <summary>
		/// Signature for archive extra data signature (were headers are encrypted).
		/// </summary>
		public const int ArchiveExtraDataSignature = 'P' | ('K' << 8) | (6 << 16) | (7 << 24);

		/// <summary>
		/// Central header digital signature
		/// </summary>
		public const int CentralHeaderDigitalSignature = 'P' | ('K' << 8) | (5 << 16) | (5 << 24);

		/// <summary>
		/// Central header digital signature
		/// </summary>
		[Obsolete("Use CentralHeaderDigitalSignaure instead")]
		public const int CENDIGITALSIG = 'P' | ('K' << 8) | (5 << 16) | (5 << 24);

		/// <summary>
		/// End of central directory record signature
		/// </summary>
		public const int EndOfCentralDirectorySignature = 'P' | ('K' << 8) | (5 << 16) | (6 << 24);

		/// <summary>
		/// End of central directory record signature
		/// </summary>
		[Obsolete("Use EndOfCentralDirectorySignature instead")]
		public const int ENDSIG = 'P' | ('K' << 8) | (5 << 16) | (6 << 24);

		#endregion Header Signatures

		/// <summary>
		/// Default encoding used for string conversion.  0 gives the default system OEM code page.
		/// Using the default code page isnt the full solution necessarily
		/// there are many variable factors, codepage 850 is often a good choice for
		/// European users, however be careful about compatability.
		/// </summary>
		[Obsolete("Use ZipStrings instead")]
		public static int DefaultCodePage
		{
			get => ZipStrings.CodePage;
			set => ZipStrings.CodePage = value;
		}

		/// <summary> Deprecated wrapper for <see cref="ZipStrings.ConvertToString(byte[], int)"/></summary>
		[Obsolete("Use ZipStrings.ConvertToString instead")]
		public static string ConvertToString(byte[] data, int count)
			=> ZipStrings.ConvertToString(data, count);

		/// <summary> Deprecated wrapper for <see cref="ZipStrings.ConvertToString(byte[])"/></summary>
		[Obsolete("Use ZipStrings.ConvertToString instead")]
		public static string ConvertToString(byte[] data)
			=> ZipStrings.ConvertToString(data);

		/// <summary> Deprecated wrapper for <see cref="ZipStrings.ConvertToStringExt(int, byte[], int)"/></summary>
		[Obsolete("Use ZipStrings.ConvertToStringExt instead")]
		public static string ConvertToStringExt(int flags, byte[] data, int count)
			=> ZipStrings.ConvertToStringExt(flags, data, count);

		/// <summary> Deprecated wrapper for <see cref="ZipStrings.ConvertToStringExt(int, byte[])"/></summary>
		[Obsolete("Use ZipStrings.ConvertToStringExt instead")]
		public static string ConvertToStringExt(int flags, byte[] data)
			=> ZipStrings.ConvertToStringExt(flags, data);

		/// <summary> Deprecated wrapper for <see cref="ZipStrings.ConvertToArray(string)"/></summary>
		[Obsolete("Use ZipStrings.ConvertToArray instead")]
		public static byte[] ConvertToArray(string str)
			=> ZipStrings.ConvertToArray(str);

		/// <summary> Deprecated wrapper for <see cref="ZipStrings.ConvertToArray(int, string)"/></summary>
		[Obsolete("Use ZipStrings.ConvertToArray instead")]
		public static byte[] ConvertToArray(int flags, string str)
			=> ZipStrings.ConvertToArray(flags, str);
	}
}

namespace ICSharpCode.SharpZipLib.Zip
{
	/// <summary>
	/// The method of encrypting entries when creating zip archives.
	/// </summary>
	public enum ZipEncryptionMethod
	{
		/// <summary>
		/// No encryption will be used.
		/// </summary>
		None,

		/// <summary>
		/// Encrypt entries with ZipCrypto.
		/// </summary>
		ZipCrypto,

		/// <summary>
		/// Encrypt entries with AES 128.
		/// </summary>
		AES128,

		/// <summary>
		/// Encrypt entries with AES 256.
		/// </summary>
		AES256
	}
}


namespace ICSharpCode.SharpZipLib.Zip
{
	/// <summary>
	/// Defines known values for the <see cref="HostSystemID"/> property.
	/// </summary>
	public enum HostSystemID
	{
		/// <summary>
		/// Host system = MSDOS
		/// </summary>
		Msdos = 0,

		/// <summary>
		/// Host system = Amiga
		/// </summary>
		Amiga = 1,

		/// <summary>
		/// Host system = Open VMS
		/// </summary>
		OpenVms = 2,

		/// <summary>
		/// Host system = Unix
		/// </summary>
		Unix = 3,

		/// <summary>
		/// Host system = VMCms
		/// </summary>
		VMCms = 4,

		/// <summary>
		/// Host system = Atari ST
		/// </summary>
		AtariST = 5,

		/// <summary>
		/// Host system = OS2
		/// </summary>
		OS2 = 6,

		/// <summary>
		/// Host system = Macintosh
		/// </summary>
		Macintosh = 7,

		/// <summary>
		/// Host system = ZSystem
		/// </summary>
		ZSystem = 8,

		/// <summary>
		/// Host system = Cpm
		/// </summary>
		Cpm = 9,

		/// <summary>
		/// Host system = Windows NT
		/// </summary>
		WindowsNT = 10,

		/// <summary>
		/// Host system = MVS
		/// </summary>
		MVS = 11,

		/// <summary>
		/// Host system = VSE
		/// </summary>
		Vse = 12,

		/// <summary>
		/// Host system = Acorn RISC
		/// </summary>
		AcornRisc = 13,

		/// <summary>
		/// Host system = VFAT
		/// </summary>
		Vfat = 14,

		/// <summary>
		/// Host system = Alternate MVS
		/// </summary>
		AlternateMvs = 15,

		/// <summary>
		/// Host system = BEOS
		/// </summary>
		BeOS = 16,

		/// <summary>
		/// Host system = Tandem
		/// </summary>
		Tandem = 17,

		/// <summary>
		/// Host system = OS400
		/// </summary>
		OS400 = 18,

		/// <summary>
		/// Host system = OSX
		/// </summary>
		OSX = 19,

		/// <summary>
		/// Host system = WinZIP AES
		/// </summary>
		WinZipAES = 99,
	}

	/// <summary>
	/// This class represents an entry in a zip archive.  This can be a file
	/// or a directory
	/// ZipFile and ZipInputStream will give you instances of this class as
	/// information about the members in an archive.  ZipOutputStream
	/// uses an instance of this class when creating an entry in a Zip file.
	/// <br/>
	/// <br/>Author of the original java version : Jochen Hoenicke
	/// </summary>
	public class ZipEntry
	{
		[Flags]
		private enum Known : byte
		{
			None = 0,
			Size = 0x01,
			CompressedSize = 0x02,
			Crc = 0x04,
			Time = 0x08,
			ExternalAttributes = 0x10,
		}

		#region Constructors

		/// <summary>
		/// Creates a zip entry with the given name.
		/// </summary>
		/// <param name="name">
		/// The name for this entry. Can include directory components.
		/// The convention for names is 'unix' style paths with relative names only.
		/// There are with no device names and path elements are separated by '/' characters.
		/// </param>
		/// <exception cref="ArgumentNullException">
		/// The name passed is null
		/// </exception>
		public ZipEntry(string name)
			: this(name, 0, ZipConstants.VersionMadeBy, CompressionMethod.Deflated)
		{
		}

		/// <summary>
		/// Creates a zip entry with the given name and version required to extract
		/// </summary>
		/// <param name="name">
		/// The name for this entry. Can include directory components.
		/// The convention for names is 'unix'  style paths with no device names and
		/// path elements separated by '/' characters.  This is not enforced see <see cref="CleanName(string)">CleanName</see>
		/// on how to ensure names are valid if this is desired.
		/// </param>
		/// <param name="versionRequiredToExtract">
		/// The minimum 'feature version' required this entry
		/// </param>
		/// <exception cref="ArgumentNullException">
		/// The name passed is null
		/// </exception>
		internal ZipEntry(string name, int versionRequiredToExtract)
			: this(name, versionRequiredToExtract, ZipConstants.VersionMadeBy,
			CompressionMethod.Deflated)
		{
		}

		/// <summary>
		/// Initializes an entry with the given name and made by information
		/// </summary>
		/// <param name="name">Name for this entry</param>
		/// <param name="madeByInfo">Version and HostSystem Information</param>
		/// <param name="versionRequiredToExtract">Minimum required zip feature version required to extract this entry</param>
		/// <param name="method">Compression method for this entry.</param>
		/// <exception cref="ArgumentNullException">
		/// The name passed is null
		/// </exception>
		/// <exception cref="ArgumentOutOfRangeException">
		/// versionRequiredToExtract should be 0 (auto-calculate) or > 10
		/// </exception>
		/// <remarks>
		/// This constructor is used by the ZipFile class when reading from the central header
		/// It is not generally useful, use the constructor specifying the name only.
		/// </remarks>
		internal ZipEntry(string name, int versionRequiredToExtract, int madeByInfo,
			CompressionMethod method)
		{
			if (name == null)
			{
				throw new ArgumentNullException(nameof(name));
			}

			if (name.Length > 0xffff)
			{
				throw new ArgumentException("Name is too long", nameof(name));
			}

			if ((versionRequiredToExtract != 0) && (versionRequiredToExtract < 10))
			{
				throw new ArgumentOutOfRangeException(nameof(versionRequiredToExtract));
			}

			this.DateTime = DateTime.Now;
			this.name = name;
			this.versionMadeBy = (ushort)madeByInfo;
			this.versionToExtract = (ushort)versionRequiredToExtract;
			this.method = method;

			IsUnicodeText = ZipStrings.UseUnicode;
		}

		/// <summary>
		/// Creates a deep copy of the given zip entry.
		/// </summary>
		/// <param name="entry">
		/// The entry to copy.
		/// </param>
		[Obsolete("Use Clone instead")]
		public ZipEntry(ZipEntry entry)
		{
			if (entry == null)
			{
				throw new ArgumentNullException(nameof(entry));
			}

			known = entry.known;
			name = entry.name;
			size = entry.size;
			compressedSize = entry.compressedSize;
			crc = entry.crc;
			dateTime = entry.DateTime;
			method = entry.method;
			comment = entry.comment;
			versionToExtract = entry.versionToExtract;
			versionMadeBy = entry.versionMadeBy;
			externalFileAttributes = entry.externalFileAttributes;
			flags = entry.flags;

			zipFileIndex = entry.zipFileIndex;
			offset = entry.offset;

			forceZip64_ = entry.forceZip64_;

			if (entry.extra != null)
			{
				extra = new byte[entry.extra.Length];
				Array.Copy(entry.extra, 0, extra, 0, entry.extra.Length);
			}
		}

		#endregion Constructors

		/// <summary>
		/// Get a value indicating whether the entry has a CRC value available.
		/// </summary>
		public bool HasCrc => (known & Known.Crc) != 0;

		/// <summary>
		/// Get/Set flag indicating if entry is encrypted.
		/// A simple helper routine to aid interpretation of <see cref="Flags">flags</see>
		/// </summary>
		/// <remarks>This is an assistant that interprets the <see cref="Flags">flags</see> property.</remarks>
		public bool IsCrypted
		{
			get => this.HasFlag(GeneralBitFlags.Encrypted);
			set => this.SetFlag(GeneralBitFlags.Encrypted, value);
		}

		/// <summary>
		/// Get / set a flag indicating whether entry name and comment text are
		/// encoded in <a href="http://www.unicode.org">unicode UTF8</a>.
		/// </summary>
		/// <remarks>This is an assistant that interprets the <see cref="Flags">flags</see> property.</remarks>
		public bool IsUnicodeText
		{
			get => this.HasFlag(GeneralBitFlags.UnicodeText);
			set => this.SetFlag(GeneralBitFlags.UnicodeText, value);
		}

		/// <summary>
		/// Value used during password checking for PKZIP 2.0 / 'classic' encryption.
		/// </summary>
		internal byte CryptoCheckValue
		{
			get => cryptoCheckValue_;
			set => cryptoCheckValue_ = value;
		}

		/// <summary>
		/// Get/Set general purpose bit flag for entry
		/// </summary>
		/// <remarks>
		/// General purpose bit flag<br/>
		/// <br/>
		/// Bit 0: If set, indicates the file is encrypted<br/>
		/// Bit 1-2 Only used for compression type 6 Imploding, and 8, 9 deflating<br/>
		/// Imploding:<br/>
		/// Bit 1 if set indicates an 8K sliding dictionary was used.  If clear a 4k dictionary was used<br/>
		/// Bit 2 if set indicates 3 Shannon-Fanno trees were used to encode the sliding dictionary, 2 otherwise<br/>
		/// <br/>
		/// Deflating:<br/>
		///   Bit 2    Bit 1<br/>
		///     0        0       Normal compression was used<br/>
		///     0        1       Maximum compression was used<br/>
		///     1        0       Fast compression was used<br/>
		///     1        1       Super fast compression was used<br/>
		/// <br/>
		/// Bit 3: If set, the fields crc-32, compressed size
		/// and uncompressed size are were not able to be written during zip file creation
		/// The correct values are held in a data descriptor immediately following the compressed data. <br/>
		/// Bit 4: Reserved for use by PKZIP for enhanced deflating<br/>
		/// Bit 5: If set indicates the file contains compressed patch data<br/>
		/// Bit 6: If set indicates strong encryption was used.<br/>
		/// Bit 7-10: Unused or reserved<br/>
		/// Bit 11: If set the name and comments for this entry are in <a href="http://www.unicode.org">unicode</a>.<br/>
		/// Bit 12-15: Unused or reserved<br/>
		/// </remarks>
		/// <seealso cref="IsUnicodeText"></seealso>
		/// <seealso cref="IsCrypted"></seealso>
		public int Flags
		{
			get => flags;
			set => flags = value;
		}

		/// <summary>
		/// Get/Set index of this entry in Zip file
		/// </summary>
		/// <remarks>This is only valid when the entry is part of a <see cref="ZipFile"></see></remarks>
		public long ZipFileIndex
		{
			get => zipFileIndex;
			set => zipFileIndex = value;
		}

		/// <summary>
		/// Get/set offset for use in central header
		/// </summary>
		public long Offset
		{
			get => offset;
			set => offset = value;
		}

		/// <summary>
		/// Get/Set external file attributes as an integer.
		/// The values of this are operating system dependent see
		/// <see cref="HostSystem">HostSystem</see> for details
		/// </summary>
		public int ExternalFileAttributes
		{
			get => (known & Known.ExternalAttributes) == 0 ? -1 : externalFileAttributes;

			set
			{
				externalFileAttributes = value;
				known |= Known.ExternalAttributes;
			}
		}

		/// <summary>
		/// Get the version made by for this entry or zero if unknown.
		/// The value / 10 indicates the major version number, and
		/// the value mod 10 is the minor version number
		/// </summary>
		public int VersionMadeBy => versionMadeBy & 0xff;

		/// <summary>
		/// Get a value indicating this entry is for a DOS/Windows system.
		/// </summary>
		public bool IsDOSEntry
			=> (HostSystem == (int)HostSystemID.Msdos) 
			|| (HostSystem == (int)HostSystemID.WindowsNT);

		/// <summary>
		/// Test the external attributes for this <see cref="ZipEntry"/> to
		/// see if the external attributes are Dos based (including WINNT and variants)
		/// and match the values
		/// </summary>
		/// <param name="attributes">The attributes to test.</param>
		/// <returns>Returns true if the external attributes are known to be DOS/Windows
		/// based and have the same attributes set as the value passed.</returns>
		private bool HasDosAttributes(int attributes)
		{
			bool result = false;
			if ((known & Known.ExternalAttributes) != 0)
			{
				result |= (((HostSystem == (int)HostSystemID.Msdos) ||
					(HostSystem == (int)HostSystemID.WindowsNT)) &&
					(ExternalFileAttributes & attributes) == attributes);
			}
			return result;
		}

		/// <summary>
		/// Gets the compatibility information for the <see cref="ExternalFileAttributes">external file attribute</see>
		/// If the external file attributes are compatible with MS-DOS and can be read
		/// by PKZIP for DOS version 2.04g then this value will be zero.  Otherwise the value
		/// will be non-zero and identify the host system on which the attributes are compatible.
		/// </summary>
		///
		/// <remarks>
		/// The values for this as defined in the Zip File format and by others are shown below.  The values are somewhat
		/// misleading in some cases as they are not all used as shown.  You should consult the relevant documentation
		/// to obtain up to date and correct information.  The modified appnote by the infozip group is
		/// particularly helpful as it documents a lot of peculiarities.  The document is however a little dated.
		/// <list type="table">
		/// <item>0 - MS-DOS and OS/2 (FAT / VFAT / FAT32 file systems)</item>
		/// <item>1 - Amiga</item>
		/// <item>2 - OpenVMS</item>
		/// <item>3 - Unix</item>
		/// <item>4 - VM/CMS</item>
		/// <item>5 - Atari ST</item>
		/// <item>6 - OS/2 HPFS</item>
		/// <item>7 - Macintosh</item>
		/// <item>8 - Z-System</item>
		/// <item>9 - CP/M</item>
		/// <item>10 - Windows NTFS</item>
		/// <item>11 - MVS (OS/390 - Z/OS)</item>
		/// <item>12 - VSE</item>
		/// <item>13 - Acorn Risc</item>
		/// <item>14 - VFAT</item>
		/// <item>15 - Alternate MVS</item>
		/// <item>16 - BeOS</item>
		/// <item>17 - Tandem</item>
		/// <item>18 - OS/400</item>
		/// <item>19 - OS/X (Darwin)</item>
		/// <item>99 - WinZip AES</item>
		/// <item>remainder - unused</item>
		/// </list>
		/// </remarks>
		public int HostSystem
		{
			get => (versionMadeBy >> 8) & 0xff;

			set
			{
				versionMadeBy &= 0x00ff;
				versionMadeBy |= (ushort)((value & 0xff) << 8);
			}
		}

		/// <summary>
		/// Get minimum Zip feature version required to extract this entry
		/// </summary>
		/// <remarks>
		/// Minimum features are defined as:<br/>
		/// 1.0 - Default value<br/>
		/// 1.1 - File is a volume label<br/>
		/// 2.0 - File is a folder/directory<br/>
		/// 2.0 - File is compressed using Deflate compression<br/>
		/// 2.0 - File is encrypted using traditional encryption<br/>
		/// 2.1 - File is compressed using Deflate64<br/>
		/// 2.5 - File is compressed using PKWARE DCL Implode<br/>
		/// 2.7 - File is a patch data set<br/>
		/// 4.5 - File uses Zip64 format extensions<br/>
		/// 4.6 - File is compressed using BZIP2 compression<br/>
		/// 5.0 - File is encrypted using DES<br/>
		/// 5.0 - File is encrypted using 3DES<br/>
		/// 5.0 - File is encrypted using original RC2 encryption<br/>
		/// 5.0 - File is encrypted using RC4 encryption<br/>
		/// 5.1 - File is encrypted using AES encryption<br/>
		/// 5.1 - File is encrypted using corrected RC2 encryption<br/>
		/// 5.1 - File is encrypted using corrected RC2-64 encryption<br/>
		/// 6.1 - File is encrypted using non-OAEP key wrapping<br/>
		/// 6.2 - Central directory encryption (not confirmed yet)<br/>
		/// 6.3 - File is compressed using LZMA<br/>
		/// 6.3 - File is compressed using PPMD+<br/>
		/// 6.3 - File is encrypted using Blowfish<br/>
		/// 6.3 - File is encrypted using Twofish<br/>
		/// </remarks>
		/// <seealso cref="CanDecompress"></seealso>
		public int Version
		{
			get
			{
				// Return recorded version if known.
				if (versionToExtract != 0)
					// Only lower order byte. High order is O/S file system.
					return versionToExtract & 0x00ff;

				if (AESKeySize > 0)
					// Ver 5.1 = AES
					return ZipConstants.VERSION_AES;

				if (CompressionMethod.BZip2 == method)
					return ZipConstants.VersionBZip2;

				if (CentralHeaderRequiresZip64)
					return ZipConstants.VersionZip64;

				if (CompressionMethod.Deflated == method || IsDirectory || IsCrypted)
					return 20;
				
				if (HasDosAttributes(0x08))
					return 11;
				
				return 10;
			}
		}

		/// <summary>
		/// Get a value indicating whether this entry can be decompressed by the library.
		/// </summary>
		/// <remarks>This is based on the <see cref="Version"></see> and
		/// whether the <see cref="IsCompressionMethodSupported()">compression method</see> is supported.</remarks>
		public bool CanDecompress 
			=> Version <= ZipConstants.VersionMadeBy 
			&& (Version == 10 || Version == 11 || Version == 20 || Version == 45 || Version == 46 || Version == 51) 
			&& IsCompressionMethodSupported();

		/// <summary>
		/// Force this entry to be recorded using Zip64 extensions.
		/// </summary>
		public void ForceZip64() => forceZip64_ = true;

		/// <summary>
		/// Get a value indicating whether Zip64 extensions were forced.
		/// </summary>
		/// <returns>A <see cref="bool"/> value of true if Zip64 extensions have been forced on; false if not.</returns>
		public bool IsZip64Forced() => forceZip64_;

		/// <summary>
		/// Gets a value indicating if the entry requires Zip64 extensions
		/// to store the full entry values.
		/// </summary>
		/// <value>A <see cref="bool"/> value of true if a local header requires Zip64 extensions; false if not.</value>
		public bool LocalHeaderRequiresZip64
		{
			get
			{
				bool result = forceZip64_;

				if (!result)
				{
					ulong trueCompressedSize = compressedSize;

					if ((versionToExtract == 0) && IsCrypted)
					{
						trueCompressedSize += (ulong)this.EncryptionOverheadSize;
					}

					// TODO: A better estimation of the true limit based on compression overhead should be used
					// to determine when an entry should use Zip64.
					result =
						((this.size >= uint.MaxValue) || (trueCompressedSize >= uint.MaxValue)) &&
						((versionToExtract == 0) || (versionToExtract >= ZipConstants.VersionZip64));
				}

				return result;
			}
		}

		/// <summary>
		/// Get a value indicating whether the central directory entry requires Zip64 extensions to be stored.
		/// </summary>
		public bool CentralHeaderRequiresZip64 
			=> LocalHeaderRequiresZip64 || (offset >= uint.MaxValue);

		/// <summary>
		/// Get/Set DosTime value.
		/// </summary>
		/// <remarks>
		/// The MS-DOS date format can only represent dates between 1/1/1980 and 12/31/2107.
		/// </remarks>
		public long DosTime
		{
			get
			{
				if ((known & Known.Time) == 0)
				{
					return 0;
				}

				var year = (uint)DateTime.Year;
				var month = (uint)DateTime.Month;
				var day = (uint)DateTime.Day;
				var hour = (uint)DateTime.Hour;
				var minute = (uint)DateTime.Minute;
				var second = (uint)DateTime.Second;

				if (year < 1980)
				{
					year = 1980;
					month = 1;
					day = 1;
					hour = 0;
					minute = 0;
					second = 0;
				}
				else if (year > 2107)
				{
					year = 2107;
					month = 12;
					day = 31;
					hour = 23;
					minute = 59;
					second = 59;
				}

				return ((year - 1980) & 0x7f) << 25 |
				       (month << 21) |
				       (day << 16) |
				       (hour << 11) |
				       (minute << 5) |
				       (second >> 1);
			}

			set
			{
				unchecked
				{
					var dosTime = (uint)value;
					uint sec = Math.Min(59, 2 * (dosTime & 0x1f));
					uint min = Math.Min(59, (dosTime >> 5) & 0x3f);
					uint hrs = Math.Min(23, (dosTime >> 11) & 0x1f);
					uint mon = Math.Max(1, Math.Min(12, ((uint)(value >> 21) & 0xf)));
					uint year = ((dosTime >> 25) & 0x7f) + 1980;
					int day = Math.Max(1, Math.Min(DateTime.DaysInMonth((int)year, (int)mon), (int)((value >> 16) & 0x1f)));
					DateTime = new DateTime((int)year, (int)mon, day, (int)hrs, (int)min, (int)sec, DateTimeKind.Unspecified);
				}
			}
		}

		/// <summary>
		/// Gets/Sets the time of last modification of the entry.
		/// </summary>
		/// <remarks>
		/// The <see cref="DosTime"></see> property is updated to match this as far as possible.
		/// </remarks>
		public DateTime DateTime
		{
			get => dateTime;

			set
			{
				dateTime = value;
				known |= Known.Time;
			}
		}

		/// <summary>
		/// Returns the entry name.
		/// </summary>
		/// <remarks>
		/// The unix naming convention is followed.
		/// Path components in the entry should always separated by forward slashes ('/').
		/// Dos device names like C: should also be removed.
		/// See the <see cref="ZipNameTransform"/> class, or <see cref="CleanName(string)"/>
		///</remarks>
		public string Name
		{
			get => name;
			internal set => name = value;
		}

		/// <summary>
		/// Gets/Sets the size of the uncompressed data.
		/// </summary>
		/// <returns>
		/// The size or -1 if unknown.
		/// </returns>
		/// <remarks>Setting the size before adding an entry to an archive can help
		/// avoid compatibility problems with some archivers which don't understand Zip64 extensions.</remarks>
		public long Size
		{
			get => (known & Known.Size) != 0 ? (long)size : -1L;
			set
			{
				size = (ulong)value;
				known |= Known.Size;
			}
		}

		/// <summary>
		/// Gets/Sets the size of the compressed data.
		/// </summary>
		/// <returns>
		/// The compressed entry size or -1 if unknown.
		/// </returns>
		public long CompressedSize
		{
			get => (known & Known.CompressedSize) != 0 ? (long)compressedSize : -1L;
			set
			{
				compressedSize = (ulong)value;
				known |= Known.CompressedSize;
			}
		}

		/// <summary>
		/// Gets/Sets the crc of the uncompressed data.
		/// </summary>
		/// <exception cref="System.ArgumentOutOfRangeException">
		/// Crc is not in the range 0..0xffffffffL
		/// </exception>
		/// <returns>
		/// The crc value or -1 if unknown.
		/// </returns>
		public long Crc
		{
			get => (known & Known.Crc) != 0 ? crc & 0xffffffffL : -1L;
			set
			{
				if ((crc & 0xffffffff00000000L) != 0)
				{
					throw new ArgumentOutOfRangeException(nameof(value));
				}
				this.crc = (uint)value;
				this.known |= Known.Crc;
			}
		}

		/// <summary>
		/// Gets/Sets the compression method.
		/// </summary>
		/// <returns>
		/// The compression method for this entry
		/// </returns>
		public CompressionMethod CompressionMethod
		{
			get => method;
			set => method = value;
		}

		/// <summary>
		/// Gets the compression method for outputting to the local or central header.
		/// Returns same value as CompressionMethod except when AES encrypting, which
		/// places 99 in the method and places the real method in the extra data.
		/// </summary>
		internal CompressionMethod CompressionMethodForHeader 
			=> (AESKeySize > 0) ? CompressionMethod.WinZipAES : method;

		/// <summary>
		/// Gets/Sets the extra data.
		/// </summary>
		/// <exception cref="System.ArgumentOutOfRangeException">
		/// Extra data is longer than 64KB (0xffff) bytes.
		/// </exception>
		/// <returns>
		/// Extra data or null if not set.
		/// </returns>
		public byte[] ExtraData
		{
			// TODO: This is slightly safer but less efficient.  Think about whether it should change.
			//				return (byte[]) extra.Clone();
			get => extra;

			set
			{
				if (value == null)
				{
					extra = null;
				}
				else
				{
					if (value.Length > 0xffff)
					{
						throw new System.ArgumentOutOfRangeException(nameof(value));
					}

					extra = new byte[value.Length];
					Array.Copy(value, 0, extra, 0, value.Length);
				}
			}
		}

		/// <summary>
		/// For AES encrypted files returns or sets the number of bits of encryption (128, 192 or 256).
		/// When setting, only 0 (off), 128 or 256 is supported.
		/// </summary>
		public int AESKeySize
		{
			get
			{
				// the strength (1 or 3) is in the entry header
				switch (_aesEncryptionStrength)
				{
					case 0:
						return 0;   // Not AES
					case 1:
						return 128;

					case 2:
						return 192; // Not used by WinZip
					case 3:
						return 256;

					default:
						throw new ZipException("Invalid AESEncryptionStrength " + _aesEncryptionStrength);
				}
			}
			set
			{
				switch (value)
				{
					case 0:
						_aesEncryptionStrength = 0;
						break;

					case 128:
						_aesEncryptionStrength = 1;
						break;

					case 256:
						_aesEncryptionStrength = 3;
						break;

					default:
						throw new ZipException("AESKeySize must be 0, 128 or 256: " + value);
				}
			}
		}

		/// <summary>
		/// AES Encryption strength for storage in extra data in entry header.
		/// 1 is 128 bit, 2 is 192 bit, 3 is 256 bit.
		/// </summary>
		internal byte AESEncryptionStrength => (byte)_aesEncryptionStrength;

		/// <summary>
		/// Returns the length of the salt, in bytes
		/// </summary>
		/// Key size -> Salt length: 128 bits = 8 bytes, 192 bits = 12 bytes, 256 bits = 16 bytes.
		internal int AESSaltLen => AESKeySize / 16;

		/// <summary>
		/// Number of extra bytes required to hold the AES Header fields (Salt, Pwd verify, AuthCode)
		/// </summary>
		/// File format:
		/// Bytes	 |	Content
		/// ---------+---------------------------
		/// Variable |	Salt value
		/// 2		 |	Password verification value
		/// Variable |	Encrypted file data
		/// 10		 |	Authentication code
		internal int AESOverheadSize => 12 + AESSaltLen;

		/// <summary>
		/// Number of extra bytes required to hold the encryption header fields.
		/// </summary>
		internal int EncryptionOverheadSize =>
			!IsCrypted
				// Entry is not encrypted - no overhead
				? 0
				: _aesEncryptionStrength == 0
					// Entry is encrypted using ZipCrypto
					? ZipConstants.CryptoHeaderSize
					// Entry is encrypted using AES
					: AESOverheadSize;

		/// <summary>
		/// Process extra data fields updating the entry based on the contents.
		/// </summary>
		/// <param name="localHeader">True if the extra data fields should be handled
		/// for a local header, rather than for a central header.
		/// </param>
		internal void ProcessExtraData(bool localHeader)
		{
			var extraData = new ZipExtraData(this.extra);

			if (extraData.Find(0x0001))
			{
				// Version required to extract is ignored here as some archivers dont set it correctly
				// in theory it should be version 45 or higher

				// The recorded size will change but remember that this is zip64.
				forceZip64_ = true;

				if (extraData.ValueLength < 4)
				{
					throw new ZipException("Extra data extended Zip64 information length is invalid");
				}

				// (localHeader ||) was deleted, because actually there is no specific difference with reading sizes between local header & central directory
				// https://pkware.cachefly.net/webdocs/casestudies/APPNOTE.TXT
				// ...
				// 4.4  Explanation of fields
				// ...
				//	4.4.8 compressed size: (4 bytes)
				//	4.4.9 uncompressed size: (4 bytes)
				//
				//		The size of the file compressed (4.4.8) and uncompressed,
				//		(4.4.9) respectively.  When a decryption header is present it
				//		will be placed in front of the file data and the value of the
				//		compressed file size will include the bytes of the decryption
				//		header.  If bit 3 of the general purpose bit flag is set,
				//		these fields are set to zero in the local header and the
				//		correct values are put in the data descriptor and
				//		in the central directory.  If an archive is in ZIP64 format
				//		and the value in this field is 0xFFFFFFFF, the size will be
				//		in the corresponding 8 byte ZIP64 extended information
				//		extra field.  When encrypting the central directory, if the
				//		local header is not in ZIP64 format and general purpose bit
				//		flag 13 is set indicating masking, the value stored for the
				//		uncompressed size in the Local Header will be zero.
				//
				// Otherwise there is problem with minizip implementation
				if (size == uint.MaxValue)
				{
					size = (ulong)extraData.ReadLong();
				}

				if (compressedSize == uint.MaxValue)
				{
					compressedSize = (ulong)extraData.ReadLong();
				}

				if (!localHeader && (offset == uint.MaxValue))
				{
					offset = extraData.ReadLong();
				}

				// Disk number on which file starts is ignored
			}
			else
			{
				if (
					((versionToExtract & 0xff) >= ZipConstants.VersionZip64) &&
					((size == uint.MaxValue) || (compressedSize == uint.MaxValue))
				)
				{
					throw new ZipException("Zip64 Extended information required but is missing.");
				}
			}

			DateTime = GetDateTime(extraData) ?? DateTime;
			if (method == CompressionMethod.WinZipAES)
			{
				ProcessAESExtraData(extraData);
			}
		}

		private static DateTime? GetDateTime(ZipExtraData extraData)
		{
			// Check for NT timestamp
			// NOTE: Disable by default to match behavior of InfoZIP
#if RESPECT_NT_TIMESTAMP
			NTTaggedData ntData = extraData.GetData<NTTaggedData>();
			if (ntData != null)
				return ntData.LastModificationTime;
#endif

			// Check for Unix timestamp
			ExtendedUnixData unixData = extraData.GetData<ExtendedUnixData>();
			if (unixData != null && unixData.Include.HasFlag(ExtendedUnixData.Flags.ModificationTime))
				return unixData.ModificationTime;

			return null;
		}

		// For AES the method in the entry is 99, and the real compression method is in the extradata
		private void ProcessAESExtraData(ZipExtraData extraData)
		{
			if (extraData.Find(0x9901))
			{
				// Set version for Zipfile.CreateAndInitDecryptionStream
				versionToExtract = ZipConstants.VERSION_AES;            // Ver 5.1 = AES see "Version" getter

				//
				// Unpack AES extra data field see http://www.winzip.com/aes_info.htm
				int length = extraData.ValueLength;         // Data size currently 7
				if (length < 7)
					throw new ZipException("AES Extra Data Length " + length + " invalid.");
				int ver = extraData.ReadShort();            // Version number (1=AE-1 2=AE-2)
				int vendorId = extraData.ReadShort();       // 2-character vendor ID 0x4541 = "AE"
				int encrStrength = extraData.ReadByte();    // encryption strength 1 = 128 2 = 192 3 = 256
				int actualCompress = extraData.ReadShort(); // The actual compression method used to compress the file
				_aesVer = ver;
				_aesEncryptionStrength = encrStrength;
				method = (CompressionMethod)actualCompress;
			}
			else
				throw new ZipException("AES Extra Data missing");
		}

		/// <summary>
		/// Gets/Sets the entry comment.
		/// </summary>
		/// <exception cref="ArgumentOutOfRangeException">
		/// If comment is longer than 0xffff.
		/// </exception>
		/// <returns>
		/// The comment or null if not set.
		/// </returns>
		/// <remarks>
		/// A comment is only available for entries when read via the <see cref="ZipFile"/> class.
		/// The <see cref="ZipInputStream"/> class doesn't have the comment data available.
		/// </remarks>
		public string Comment
		{
			get => comment;
			set
			{
				// This test is strictly incorrect as the length is in characters
				// while the storage limit is in bytes.
				// While the test is partially correct in that a comment of this length or greater
				// is definitely invalid, shorter comments may also have an invalid length
				// where there are multi-byte characters
				// The full test is not possible here however as the code page to apply conversions with
				// isn't available.
				if ((value != null) && (value.Length > 0xffff))
				{
					throw new ArgumentOutOfRangeException(nameof(value), "cannot exceed 65535");
				}

				comment = value;
			}
		}

		/// <summary>
		/// Gets a value indicating if the entry is a directory.
		/// however.
		/// </summary>
		/// <remarks>
		/// A directory is determined by an entry name with a trailing slash '/'.
		/// The external file attributes can also indicate an entry is for a directory.
		/// Currently only dos/windows attributes are tested in this manner.
		/// The trailing slash convention should always be followed.
		/// </remarks>
		public bool IsDirectory 
			=> name.Length > 0 
			&& (name[name.Length - 1] == '/' || name[name.Length - 1] == '\\') || HasDosAttributes(16);

		/// <summary>
		/// Get a value of true if the entry appears to be a file; false otherwise
		/// </summary>
		/// <remarks>
		/// This only takes account of DOS/Windows attributes.  Other operating systems are ignored.
		/// For linux and others the result may be incorrect.
		/// </remarks>
		public bool IsFile => !IsDirectory && !HasDosAttributes(8);

		/// <summary>
		/// Test entry to see if data can be extracted.
		/// </summary>
		/// <returns>Returns true if data can be extracted for this entry; false otherwise.</returns>
		public bool IsCompressionMethodSupported() => IsCompressionMethodSupported(CompressionMethod);

		#region ICloneable Members

		/// <summary>
		/// Creates a copy of this zip entry.
		/// </summary>
		/// <returns>An <see cref="Object"/> that is a copy of the current instance.</returns>
		public object Clone()
		{
			var result = (ZipEntry)this.MemberwiseClone();

			// Ensure extra data is unique if it exists.
			if (extra != null)
			{
				result.extra = new byte[extra.Length];
				Array.Copy(extra, 0, result.extra, 0, extra.Length);
			}

			return result;
		}

		#endregion ICloneable Members

		/// <summary>
		/// Gets a string representation of this ZipEntry.
		/// </summary>
		/// <returns>A readable textual representation of this <see cref="ZipEntry"/></returns>
		public override string ToString() => name;

		/// <summary>
		/// Test a <see cref="CompressionMethod">compression method</see> to see if this library
		/// supports extracting data compressed with that method
		/// </summary>
		/// <param name="method">The compression method to test.</param>
		/// <returns>Returns true if the compression method is supported; false otherwise</returns>
		public static bool IsCompressionMethodSupported(CompressionMethod method) 
			=> method == CompressionMethod.Deflated
			|| method == CompressionMethod.Stored
			|| method == CompressionMethod.BZip2;

		/// <summary>
		/// Cleans a name making it conform to Zip file conventions.
		/// Devices names ('c:\') and UNC share names ('\\server\share') are removed
		/// and back slashes ('\') are converted to forward slashes ('/').
		/// Names are made relative by trimming leading slashes which is compatible
		/// with the ZIP naming convention.
		/// </summary>
		/// <param name="name">The name to clean</param>
		/// <returns>The 'cleaned' name.</returns>
		/// <remarks>
		/// The <seealso cref="ZipNameTransform">Zip name transform</seealso> class is more flexible.
		/// </remarks>
		public static string CleanName(string name)
		{
			if (name == null)
			{
				return string.Empty;
			}

			if (Path.IsPathRooted(name))
			{
				// NOTE:
				// for UNC names...  \\machine\share\zoom\beet.txt gives \zoom\beet.txt
				name = name.Substring(Path.GetPathRoot(name).Length);
			}

			name = name.Replace(@"\", "/");

			while ((name.Length > 0) && (name[0] == '/'))
			{
				name = name.Remove(0, 1);
			}
			return name;
		}

		#region Instance Fields

		private Known known;
		private int externalFileAttributes = -1;     // contains external attributes (O/S dependant)

		private ushort versionMadeBy;                   // Contains host system and version information
														// only relevant for central header entries

		private string name;
		private ulong size;
		private ulong compressedSize;
		private ushort versionToExtract;                // Version required to extract (library handles <= 2.0)
		private uint crc;
		private DateTime dateTime;

		private CompressionMethod method = CompressionMethod.Deflated;
		private byte[] extra;
		private string comment;

		private int flags;                             // general purpose bit flags

		private long zipFileIndex = -1;                // used by ZipFile
		private long offset;                           // used by ZipFile and ZipOutputStream

		private bool forceZip64_;
		private byte cryptoCheckValue_;
		private int _aesVer;                            // Version number (2 = AE-2 ?). Assigned but not used.
		private int _aesEncryptionStrength;             // Encryption strength 1 = 128 2 = 192 3 = 256

		#endregion Instance Fields
	}
}


namespace ICSharpCode.SharpZipLib.Zip
{
	/// <summary>
	/// General ZipEntry helper extensions
	/// </summary>
	public static class ZipEntryExtensions
	{
		/// <summary>
		/// Efficiently check if a <see cref="GeneralBitFlags">flag</see> is set without enum un-/boxing
		/// </summary>
		/// <param name="entry"></param>
		/// <param name="flag"></param>
		/// <returns>Returns whether the flag was set</returns>
		public static bool HasFlag(this ZipEntry entry, GeneralBitFlags flag)
			=> (entry.Flags & (int) flag) != 0;

		/// <summary>
		/// Efficiently set a <see cref="GeneralBitFlags">flag</see> without enum un-/boxing
		/// </summary>
		/// <param name="entry"></param>
		/// <param name="flag"></param>
		/// <param name="enabled">Whether the passed flag should be set (1) or cleared (0)</param>
		public static void SetFlag(this ZipEntry entry, GeneralBitFlags flag, bool enabled = true)
			=> entry.Flags = enabled 
				? entry.Flags | (int) flag 
				: entry.Flags & ~(int) flag;
	}
}


namespace ICSharpCode.SharpZipLib.Zip
{
	/// <summary>
	/// Basic implementation of <see cref="IEntryFactory"></see>
	/// </summary>
	public class ZipEntryFactory : IEntryFactory
	{
		#region Enumerations

		/// <summary>
		/// Defines the possible values to be used for the <see cref="ZipEntry.DateTime"/>.
		/// </summary>
		public enum TimeSetting
		{
			/// <summary>
			/// Use the recorded LastWriteTime value for the file.
			/// </summary>
			LastWriteTime,

			/// <summary>
			/// Use the recorded LastWriteTimeUtc value for the file
			/// </summary>
			LastWriteTimeUtc,

			/// <summary>
			/// Use the recorded CreateTime value for the file.
			/// </summary>
			CreateTime,

			/// <summary>
			/// Use the recorded CreateTimeUtc value for the file.
			/// </summary>
			CreateTimeUtc,

			/// <summary>
			/// Use the recorded LastAccessTime value for the file.
			/// </summary>
			LastAccessTime,

			/// <summary>
			/// Use the recorded LastAccessTimeUtc value for the file.
			/// </summary>
			LastAccessTimeUtc,

			/// <summary>
			/// Use a fixed value.
			/// </summary>
			/// <remarks>The actual <see cref="DateTime"/> value used can be
			/// specified via the <see cref="ZipEntryFactory(DateTime)"/> constructor or
			/// using the <see cref="ZipEntryFactory(TimeSetting)"/> with the setting set
			/// to <see cref="TimeSetting.Fixed"/> which will use the <see cref="DateTime"/> when this class was constructed.
			/// The <see cref="FixedDateTime"/> property can also be used to set this value.</remarks>
			Fixed,
		}

		#endregion Enumerations

		#region Constructors

		/// <summary>
		/// Initialise a new instance of the <see cref="ZipEntryFactory"/> class.
		/// </summary>
		/// <remarks>A default <see cref="INameTransform"/>, and the LastWriteTime for files is used.</remarks>
		public ZipEntryFactory()
		{
			nameTransform_ = new ZipNameTransform();
			isUnicodeText_ = ZipStrings.UseUnicode;
		}

		/// <summary>
		/// Initialise a new instance of <see cref="ZipEntryFactory"/> using the specified <see cref="TimeSetting"/>
		/// </summary>
		/// <param name="timeSetting">The <see cref="TimeSetting">time setting</see> to use when creating <see cref="ZipEntry">Zip entries</see>.</param>
		public ZipEntryFactory(TimeSetting timeSetting) : this()
		{
			timeSetting_ = timeSetting;
		}

		/// <summary>
		/// Initialise a new instance of <see cref="ZipEntryFactory"/> using the specified <see cref="DateTime"/>
		/// </summary>
		/// <param name="time">The time to set all <see cref="ZipEntry.DateTime"/> values to.</param>
		public ZipEntryFactory(DateTime time) : this()
		{
			timeSetting_ = TimeSetting.Fixed;
			FixedDateTime = time;
		}

		#endregion Constructors

		#region Properties

		/// <summary>
		/// Get / set the <see cref="INameTransform"/> to be used when creating new <see cref="ZipEntry"/> values.
		/// </summary>
		/// <remarks>
		/// Setting this property to null will cause a default <see cref="ZipNameTransform">name transform</see> to be used.
		/// </remarks>
		public INameTransform NameTransform
		{
			get { return nameTransform_; }
			set
			{
				if (value == null)
				{
					nameTransform_ = new ZipNameTransform();
				}
				else
				{
					nameTransform_ = value;
				}
			}
		}

		/// <summary>
		/// Get / set the <see cref="TimeSetting"/> in use.
		/// </summary>
		public TimeSetting Setting
		{
			get { return timeSetting_; }
			set { timeSetting_ = value; }
		}

		/// <summary>
		/// Get / set the <see cref="DateTime"/> value to use when <see cref="Setting"/> is set to <see cref="TimeSetting.Fixed"/>
		/// </summary>
		public DateTime FixedDateTime
		{
			get { return fixedDateTime_; }
			set
			{
				if (value.Year < 1970)
				{
					throw new ArgumentException("Value is too old to be valid", nameof(value));
				}
				fixedDateTime_ = value;
			}
		}

		/// <summary>
		/// A bitmask defining the attributes to be retrieved from the actual file.
		/// </summary>
		/// <remarks>The default is to get all possible attributes from the actual file.</remarks>
		public int GetAttributes
		{
			get { return getAttributes_; }
			set { getAttributes_ = value; }
		}

		/// <summary>
		/// A bitmask defining which attributes are to be set on.
		/// </summary>
		/// <remarks>By default no attributes are set on.</remarks>
		public int SetAttributes
		{
			get { return setAttributes_; }
			set { setAttributes_ = value; }
		}

		/// <summary>
		/// Get set a value indicating whether unidoce text should be set on.
		/// </summary>
		public bool IsUnicodeText
		{
			get { return isUnicodeText_; }
			set { isUnicodeText_ = value; }
		}

		#endregion Properties

		#region IEntryFactory Members

		/// <summary>
		/// Make a new <see cref="ZipEntry"/> for a file.
		/// </summary>
		/// <param name="fileName">The name of the file to create a new entry for.</param>
		/// <returns>Returns a new <see cref="ZipEntry"/> based on the <paramref name="fileName"/>.</returns>
		public ZipEntry MakeFileEntry(string fileName)
		{
			return MakeFileEntry(fileName, null, true);
		}

		/// <summary>
		/// Make a new <see cref="ZipEntry"/> for a file.
		/// </summary>
		/// <param name="fileName">The name of the file to create a new entry for.</param>
		/// <param name="useFileSystem">If true entry detail is retrieved from the file system if the file exists.</param>
		/// <returns>Returns a new <see cref="ZipEntry"/> based on the <paramref name="fileName"/>.</returns>
		public ZipEntry MakeFileEntry(string fileName, bool useFileSystem)
		{
			return MakeFileEntry(fileName, null, useFileSystem);
		}

		/// <summary>
		/// Make a new <see cref="ZipEntry"/> from a name.
		/// </summary>
		/// <param name="fileName">The name of the file to create a new entry for.</param>
		/// <param name="entryName">An alternative name to be used for the new entry. Null if not applicable.</param>
		/// <param name="useFileSystem">If true entry detail is retrieved from the file system if the file exists.</param>
		/// <returns>Returns a new <see cref="ZipEntry"/> based on the <paramref name="fileName"/>.</returns>
		public ZipEntry MakeFileEntry(string fileName, string entryName, bool useFileSystem)
		{
			var result = new ZipEntry(nameTransform_.TransformFile(!string.IsNullOrEmpty(entryName) ? entryName : fileName));
			result.IsUnicodeText = isUnicodeText_;

			int externalAttributes = 0;
			bool useAttributes = (setAttributes_ != 0);

			FileInfo fi = null;
			if (useFileSystem)
			{
				fi = new FileInfo(fileName);
			}

			if ((fi != null) && fi.Exists)
			{
				switch (timeSetting_)
				{
					case TimeSetting.CreateTime:
						result.DateTime = fi.CreationTime;
						break;

					case TimeSetting.CreateTimeUtc:
						result.DateTime = fi.CreationTimeUtc;
						break;

					case TimeSetting.LastAccessTime:
						result.DateTime = fi.LastAccessTime;
						break;

					case TimeSetting.LastAccessTimeUtc:
						result.DateTime = fi.LastAccessTimeUtc;
						break;

					case TimeSetting.LastWriteTime:
						result.DateTime = fi.LastWriteTime;
						break;

					case TimeSetting.LastWriteTimeUtc:
						result.DateTime = fi.LastWriteTimeUtc;
						break;

					case TimeSetting.Fixed:
						result.DateTime = fixedDateTime_;
						break;

					default:
						throw new ZipException("Unhandled time setting in MakeFileEntry");
				}

				result.Size = fi.Length;

				useAttributes = true;
				externalAttributes = ((int)fi.Attributes & getAttributes_);
			}
			else
			{
				if (timeSetting_ == TimeSetting.Fixed)
				{
					result.DateTime = fixedDateTime_;
				}
			}

			if (useAttributes)
			{
				externalAttributes |= setAttributes_;
				result.ExternalFileAttributes = externalAttributes;
			}

			return result;
		}

		/// <summary>
		/// Make a new <see cref="ZipEntry"></see> for a directory.
		/// </summary>
		/// <param name="directoryName">The raw untransformed name for the new directory</param>
		/// <returns>Returns a new <see cref="ZipEntry"></see> representing a directory.</returns>
		public ZipEntry MakeDirectoryEntry(string directoryName)
		{
			return MakeDirectoryEntry(directoryName, true);
		}

		/// <summary>
		/// Make a new <see cref="ZipEntry"></see> for a directory.
		/// </summary>
		/// <param name="directoryName">The raw untransformed name for the new directory</param>
		/// <param name="useFileSystem">If true entry detail is retrieved from the file system if the file exists.</param>
		/// <returns>Returns a new <see cref="ZipEntry"></see> representing a directory.</returns>
		public ZipEntry MakeDirectoryEntry(string directoryName, bool useFileSystem)
		{
			var result = new ZipEntry(nameTransform_.TransformDirectory(directoryName));
			result.IsUnicodeText = isUnicodeText_;
			result.Size = 0;

			int externalAttributes = 0;

			DirectoryInfo di = null;

			if (useFileSystem)
			{
				di = new DirectoryInfo(directoryName);
			}

			if ((di != null) && di.Exists)
			{
				switch (timeSetting_)
				{
					case TimeSetting.CreateTime:
						result.DateTime = di.CreationTime;
						break;

					case TimeSetting.CreateTimeUtc:
						result.DateTime = di.CreationTimeUtc;
						break;

					case TimeSetting.LastAccessTime:
						result.DateTime = di.LastAccessTime;
						break;

					case TimeSetting.LastAccessTimeUtc:
						result.DateTime = di.LastAccessTimeUtc;
						break;

					case TimeSetting.LastWriteTime:
						result.DateTime = di.LastWriteTime;
						break;

					case TimeSetting.LastWriteTimeUtc:
						result.DateTime = di.LastWriteTimeUtc;
						break;

					case TimeSetting.Fixed:
						result.DateTime = fixedDateTime_;
						break;

					default:
						throw new ZipException("Unhandled time setting in MakeDirectoryEntry");
				}

				externalAttributes = ((int)di.Attributes & getAttributes_);
			}
			else
			{
				if (timeSetting_ == TimeSetting.Fixed)
				{
					result.DateTime = fixedDateTime_;
				}
			}

			// Always set directory attribute on.
			externalAttributes |= (setAttributes_ | 16);
			result.ExternalFileAttributes = externalAttributes;

			return result;
		}

		#endregion IEntryFactory Members

		#region Instance Fields

		private INameTransform nameTransform_;
		private DateTime fixedDateTime_ = DateTime.Now;
		private TimeSetting timeSetting_ = TimeSetting.LastWriteTime;
		private bool isUnicodeText_;

		private int getAttributes_ = -1;
		private int setAttributes_;

		#endregion Instance Fields
	}
}


namespace ICSharpCode.SharpZipLib.Zip
{
	/// <summary>
	/// ZipException represents exceptions specific to Zip classes and code.
	/// </summary>
	[Serializable]
	public class ZipException : SharpZipBaseException
	{
		/// <summary>
		/// Initialise a new instance of <see cref="ZipException" />.
		/// </summary>
		public ZipException()
		{
		}

		/// <summary>
		/// Initialise a new instance of <see cref="ZipException" /> with its message string.
		/// </summary>
		/// <param name="message">A <see cref="string"/> that describes the error.</param>
		public ZipException(string message)
			: base(message)
		{
		}

		/// <summary>
		/// Initialise a new instance of <see cref="ZipException" />.
		/// </summary>
		/// <param name="message">A <see cref="string"/> that describes the error.</param>
		/// <param name="innerException">The <see cref="Exception"/> that caused this exception.</param>
		public ZipException(string message, Exception innerException)
			: base(message, innerException)
		{
		}

		/// <summary>
		/// Initializes a new instance of the ZipException class with serialized data.
		/// </summary>
		/// <param name="info">
		/// The System.Runtime.Serialization.SerializationInfo that holds the serialized
		/// object data about the exception being thrown.
		/// </param>
		/// <param name="context">
		/// The System.Runtime.Serialization.StreamingContext that contains contextual information
		/// about the source or destination.
		/// </param>
		protected ZipException(SerializationInfo info, StreamingContext context)
			: base(info, context)
		{
		}
	}
}


namespace ICSharpCode.SharpZipLib.Zip
{
	// TODO: Sort out whether tagged data is useful and what a good implementation might look like.
	// Its just a sketch of an idea at the moment.

	/// <summary>
	/// ExtraData tagged value interface.
	/// </summary>
	public interface ITaggedData
	{
		/// <summary>
		/// Get the ID for this tagged data value.
		/// </summary>
		short TagID { get; }

		/// <summary>
		/// Set the contents of this instance from the data passed.
		/// </summary>
		/// <param name="data">The data to extract contents from.</param>
		/// <param name="offset">The offset to begin extracting data from.</param>
		/// <param name="count">The number of bytes to extract.</param>
		void SetData(byte[] data, int offset, int count);

		/// <summary>
		/// Get the data representing this instance.
		/// </summary>
		/// <returns>Returns the data for this instance.</returns>
		byte[] GetData();
	}

	/// <summary>
	/// A raw binary tagged value
	/// </summary>
	public class RawTaggedData : ITaggedData
	{
		/// <summary>
		/// Initialise a new instance.
		/// </summary>
		/// <param name="tag">The tag ID.</param>
		public RawTaggedData(short tag)
		{
			_tag = tag;
		}

		#region ITaggedData Members

		/// <summary>
		/// Get the ID for this tagged data value.
		/// </summary>
		public short TagID
		{
			get { return _tag; }
			set { _tag = value; }
		}

		/// <summary>
		/// Set the data from the raw values provided.
		/// </summary>
		/// <param name="data">The raw data to extract values from.</param>
		/// <param name="offset">The index to start extracting values from.</param>
		/// <param name="count">The number of bytes available.</param>
		public void SetData(byte[] data, int offset, int count)
		{
			if (data == null)
			{
				throw new ArgumentNullException(nameof(data));
			}

			_data = new byte[count];
			Array.Copy(data, offset, _data, 0, count);
		}

		/// <summary>
		/// Get the binary data representing this instance.
		/// </summary>
		/// <returns>The raw binary data representing this instance.</returns>
		public byte[] GetData()
		{
			return _data;
		}

		#endregion ITaggedData Members

		/// <summary>
		/// Get /set the binary data representing this instance.
		/// </summary>
		/// <returns>The raw binary data representing this instance.</returns>
		public byte[] Data
		{
			get { return _data; }
			set { _data = value; }
		}

		#region Instance Fields

		/// <summary>
		/// The tag ID for this instance.
		/// </summary>
		private short _tag;

		private byte[] _data;

		#endregion Instance Fields
	}

	/// <summary>
	/// Class representing extended unix date time values.
	/// </summary>
	public class ExtendedUnixData : ITaggedData
	{
		/// <summary>
		/// Flags indicate which values are included in this instance.
		/// </summary>
		[Flags]
		public enum Flags : byte
		{
			/// <summary>
			/// The modification time is included
			/// </summary>
			ModificationTime = 0x01,

			/// <summary>
			/// The access time is included
			/// </summary>
			AccessTime = 0x02,

			/// <summary>
			/// The create time is included.
			/// </summary>
			CreateTime = 0x04,
		}

		#region ITaggedData Members

		/// <summary>
		/// Get the ID
		/// </summary>
		public short TagID
		{
			get { return 0x5455; }
		}

		/// <summary>
		/// Set the data from the raw values provided.
		/// </summary>
		/// <param name="data">The raw data to extract values from.</param>
		/// <param name="index">The index to start extracting values from.</param>
		/// <param name="count">The number of bytes available.</param>
		public void SetData(byte[] data, int index, int count)
		{
			using (MemoryStream ms = new MemoryStream(data, index, count, false))
			using (ZipHelperStream helperStream = new ZipHelperStream(ms))
			{
				// bit 0           if set, modification time is present
				// bit 1           if set, access time is present
				// bit 2           if set, creation time is present

				_flags = (Flags)helperStream.ReadByte();
				if (((_flags & Flags.ModificationTime) != 0))
				{
					int iTime = helperStream.ReadLEInt();

					_modificationTime = new DateTime(1970, 1, 1, 0, 0, 0, 0, DateTimeKind.Utc) +
						new TimeSpan(0, 0, 0, iTime, 0);

					// Central-header version is truncated after modification time
					if (count <= 5) return;
				}

				if ((_flags & Flags.AccessTime) != 0)
				{
					int iTime = helperStream.ReadLEInt();

					_lastAccessTime = new DateTime(1970, 1, 1, 0, 0, 0, 0, DateTimeKind.Utc) +
						new TimeSpan(0, 0, 0, iTime, 0);
				}

				if ((_flags & Flags.CreateTime) != 0)
				{
					int iTime = helperStream.ReadLEInt();

					_createTime = new DateTime(1970, 1, 1, 0, 0, 0, 0, DateTimeKind.Utc) +
						new TimeSpan(0, 0, 0, iTime, 0);
				}
			}
		}

		/// <summary>
		/// Get the binary data representing this instance.
		/// </summary>
		/// <returns>The raw binary data representing this instance.</returns>
		public byte[] GetData()
		{
			using (MemoryStream ms = new MemoryStream())
			using (ZipHelperStream helperStream = new ZipHelperStream(ms))
			{
				helperStream.IsStreamOwner = false;
				helperStream.WriteByte((byte)_flags);     // Flags
				if ((_flags & Flags.ModificationTime) != 0)
				{
					TimeSpan span = _modificationTime - new DateTime(1970, 1, 1, 0, 0, 0, 0, DateTimeKind.Utc);
					var seconds = (int)span.TotalSeconds;
					helperStream.WriteLEInt(seconds);
				}
				if ((_flags & Flags.AccessTime) != 0)
				{
					TimeSpan span = _lastAccessTime - new DateTime(1970, 1, 1, 0, 0, 0, 0, DateTimeKind.Utc);
					var seconds = (int)span.TotalSeconds;
					helperStream.WriteLEInt(seconds);
				}
				if ((_flags & Flags.CreateTime) != 0)
				{
					TimeSpan span = _createTime - new DateTime(1970, 1, 1, 0, 0, 0, 0, DateTimeKind.Utc);
					var seconds = (int)span.TotalSeconds;
					helperStream.WriteLEInt(seconds);
				}
				return ms.ToArray();
			}
		}

		#endregion ITaggedData Members

		/// <summary>
		/// Test a <see cref="DateTime"> value to see if is valid and can be represented here.</see>
		/// </summary>
		/// <param name="value">The <see cref="DateTime">value</see> to test.</param>
		/// <returns>Returns true if the value is valid and can be represented; false if not.</returns>
		/// <remarks>The standard Unix time is a signed integer data type, directly encoding the Unix time number,
		/// which is the number of seconds since 1970-01-01.
		/// Being 32 bits means the values here cover a range of about 136 years.
		/// The minimum representable time is 1901-12-13 20:45:52,
		/// and the maximum representable time is 2038-01-19 03:14:07.
		/// </remarks>
		public static bool IsValidValue(DateTime value)
		{
			return ((value >= new DateTime(1901, 12, 13, 20, 45, 52)) ||
					(value <= new DateTime(2038, 1, 19, 03, 14, 07)));
		}

		/// <summary>
		/// Get /set the Modification Time
		/// </summary>
		/// <exception cref="ArgumentOutOfRangeException"></exception>
		/// <seealso cref="IsValidValue"></seealso>
		public DateTime ModificationTime
		{
			get { return _modificationTime; }
			set
			{
				if (!IsValidValue(value))
				{
					throw new ArgumentOutOfRangeException(nameof(value));
				}

				_flags |= Flags.ModificationTime;
				_modificationTime = value;
			}
		}

		/// <summary>
		/// Get / set the Access Time
		/// </summary>
		/// <exception cref="ArgumentOutOfRangeException"></exception>
		/// <seealso cref="IsValidValue"></seealso>
		public DateTime AccessTime
		{
			get { return _lastAccessTime; }
			set
			{
				if (!IsValidValue(value))
				{
					throw new ArgumentOutOfRangeException(nameof(value));
				}

				_flags |= Flags.AccessTime;
				_lastAccessTime = value;
			}
		}

		/// <summary>
		/// Get / Set the Create Time
		/// </summary>
		/// <exception cref="ArgumentOutOfRangeException"></exception>
		/// <seealso cref="IsValidValue"></seealso>
		public DateTime CreateTime
		{
			get { return _createTime; }
			set
			{
				if (!IsValidValue(value))
				{
					throw new ArgumentOutOfRangeException(nameof(value));
				}

				_flags |= Flags.CreateTime;
				_createTime = value;
			}
		}

		/// <summary>
		/// Get/set the <see cref="Flags">values</see> to include.
		/// </summary>
		public Flags Include
		{
			get { return _flags; }
			set { _flags = value; }
		}

		#region Instance Fields

		private Flags _flags;
		private DateTime _modificationTime = new DateTime(1970, 1, 1);
		private DateTime _lastAccessTime = new DateTime(1970, 1, 1);
		private DateTime _createTime = new DateTime(1970, 1, 1);

		#endregion Instance Fields
	}

	/// <summary>
	/// Class handling NT date time values.
	/// </summary>
	public class NTTaggedData : ITaggedData
	{
		/// <summary>
		/// Get the ID for this tagged data value.
		/// </summary>
		public short TagID
		{
			get { return 10; }
		}

		/// <summary>
		/// Set the data from the raw values provided.
		/// </summary>
		/// <param name="data">The raw data to extract values from.</param>
		/// <param name="index">The index to start extracting values from.</param>
		/// <param name="count">The number of bytes available.</param>
		public void SetData(byte[] data, int index, int count)
		{
			using (MemoryStream ms = new MemoryStream(data, index, count, false))
			using (ZipHelperStream helperStream = new ZipHelperStream(ms))
			{
				helperStream.ReadLEInt(); // Reserved
				while (helperStream.Position < helperStream.Length)
				{
					int ntfsTag = helperStream.ReadLEShort();
					int ntfsLength = helperStream.ReadLEShort();
					if (ntfsTag == 1)
					{
						if (ntfsLength >= 24)
						{
							long lastModificationTicks = helperStream.ReadLELong();
							_lastModificationTime = DateTime.FromFileTimeUtc(lastModificationTicks);

							long lastAccessTicks = helperStream.ReadLELong();
							_lastAccessTime = DateTime.FromFileTimeUtc(lastAccessTicks);

							long createTimeTicks = helperStream.ReadLELong();
							_createTime = DateTime.FromFileTimeUtc(createTimeTicks);
						}
						break;
					}
					else
					{
						// An unknown NTFS tag so simply skip it.
						helperStream.Seek(ntfsLength, SeekOrigin.Current);
					}
				}
			}
		}

		/// <summary>
		/// Get the binary data representing this instance.
		/// </summary>
		/// <returns>The raw binary data representing this instance.</returns>
		public byte[] GetData()
		{
			using (MemoryStream ms = new MemoryStream())
			using (ZipHelperStream helperStream = new ZipHelperStream(ms))
			{
				helperStream.IsStreamOwner = false;
				helperStream.WriteLEInt(0);       // Reserved
				helperStream.WriteLEShort(1);     // Tag
				helperStream.WriteLEShort(24);    // Length = 3 x 8.
				helperStream.WriteLELong(_lastModificationTime.ToFileTimeUtc());
				helperStream.WriteLELong(_lastAccessTime.ToFileTimeUtc());
				helperStream.WriteLELong(_createTime.ToFileTimeUtc());
				return ms.ToArray();
			}
		}

		/// <summary>
		/// Test a <see cref="DateTime"> valuie to see if is valid and can be represented here.</see>
		/// </summary>
		/// <param name="value">The <see cref="DateTime">value</see> to test.</param>
		/// <returns>Returns true if the value is valid and can be represented; false if not.</returns>
		/// <remarks>
		/// NTFS filetimes are 64-bit unsigned integers, stored in Intel
		/// (least significant byte first) byte order. They determine the
		/// number of 1.0E-07 seconds (1/10th microseconds!) past WinNT "epoch",
		/// which is "01-Jan-1601 00:00:00 UTC". 28 May 60056 is the upper limit
		/// </remarks>
		public static bool IsValidValue(DateTime value)
		{
			bool result = true;
			try
			{
				value.ToFileTimeUtc();
			}
			catch
			{
				result = false;
			}
			return result;
		}

		/// <summary>
		/// Get/set the <see cref="DateTime">last modification time</see>.
		/// </summary>
		public DateTime LastModificationTime
		{
			get { return _lastModificationTime; }
			set
			{
				if (!IsValidValue(value))
				{
					throw new ArgumentOutOfRangeException(nameof(value));
				}
				_lastModificationTime = value;
			}
		}

		/// <summary>
		/// Get /set the <see cref="DateTime">create time</see>
		/// </summary>
		public DateTime CreateTime
		{
			get { return _createTime; }
			set
			{
				if (!IsValidValue(value))
				{
					throw new ArgumentOutOfRangeException(nameof(value));
				}
				_createTime = value;
			}
		}

		/// <summary>
		/// Get /set the <see cref="DateTime">last access time</see>.
		/// </summary>
		public DateTime LastAccessTime
		{
			get { return _lastAccessTime; }
			set
			{
				if (!IsValidValue(value))
				{
					throw new ArgumentOutOfRangeException(nameof(value));
				}
				_lastAccessTime = value;
			}
		}

		#region Instance Fields

		private DateTime _lastAccessTime = DateTime.FromFileTimeUtc(0);
		private DateTime _lastModificationTime = DateTime.FromFileTimeUtc(0);
		private DateTime _createTime = DateTime.FromFileTimeUtc(0);

		#endregion Instance Fields
	}

	/// <summary>
	/// A factory that creates <see cref="ITaggedData">tagged data</see> instances.
	/// </summary>
	internal interface ITaggedDataFactory
	{
		/// <summary>
		/// Get data for a specific tag value.
		/// </summary>
		/// <param name="tag">The tag ID to find.</param>
		/// <param name="data">The data to search.</param>
		/// <param name="offset">The offset to begin extracting data from.</param>
		/// <param name="count">The number of bytes to extract.</param>
		/// <returns>The located <see cref="ITaggedData">value found</see>, or null if not found.</returns>
		ITaggedData Create(short tag, byte[] data, int offset, int count);
	}

	///
	/// <summary>
	/// A class to handle the extra data field for Zip entries
	/// </summary>
	/// <remarks>
	/// Extra data contains 0 or more values each prefixed by a header tag and length.
	/// They contain zero or more bytes of actual data.
	/// The data is held internally using a copy on write strategy.  This is more efficient but
	/// means that for extra data created by passing in data can have the values modified by the caller
	/// in some circumstances.
	/// </remarks>
	sealed public class ZipExtraData : IDisposable
	{
		#region Constructors

		/// <summary>
		/// Initialise a default instance.
		/// </summary>
		public ZipExtraData()
		{
			Clear();
		}

		/// <summary>
		/// Initialise with known extra data.
		/// </summary>
		/// <param name="data">The extra data.</param>
		public ZipExtraData(byte[] data)
		{
			if (data == null)
			{
				_data = Empty.Array<byte>();
			}
			else
			{
				_data = data;
			}
		}

		#endregion Constructors

		/// <summary>
		/// Get the raw extra data value
		/// </summary>
		/// <returns>Returns the raw byte[] extra data this instance represents.</returns>
		public byte[] GetEntryData()
		{
			if (Length > ushort.MaxValue)
			{
				throw new ZipException("Data exceeds maximum length");
			}

			return (byte[])_data.Clone();
		}

		/// <summary>
		/// Clear the stored data.
		/// </summary>
		public void Clear()
		{
			if ((_data == null) || (_data.Length != 0))
			{
				_data = Empty.Array<byte>();
			}
		}

		/// <summary>
		/// Gets the current extra data length.
		/// </summary>
		public int Length
		{
			get { return _data.Length; }
		}

		/// <summary>
		/// Get a read-only <see cref="Stream"/> for the associated tag.
		/// </summary>
		/// <param name="tag">The tag to locate data for.</param>
		/// <returns>Returns a <see cref="Stream"/> containing tag data or null if no tag was found.</returns>
		public Stream GetStreamForTag(int tag)
		{
			Stream result = null;
			if (Find(tag))
			{
				result = new MemoryStream(_data, _index, _readValueLength, false);
			}
			return result;
		}

		/// <summary>
		/// Get the <see cref="ITaggedData">tagged data</see> for a tag.
		/// </summary>
		/// <typeparam name="T">The tag to search for.</typeparam>
		/// <returns>Returns a <see cref="ITaggedData">tagged value</see> or null if none found.</returns>
		public T GetData<T>()
			where T : class, ITaggedData, new()
		{
			T result = new T();
			if (Find(result.TagID))
			{
				result.SetData(_data, _readValueStart, _readValueLength);
				return result;
			}
			else return null;
		}

		/// <summary>
		/// Get the length of the last value found by <see cref="Find"/>
		/// </summary>
		/// <remarks>This is only valid if <see cref="Find"/> has previously returned true.</remarks>
		public int ValueLength
		{
			get { return _readValueLength; }
		}

		/// <summary>
		/// Get the index for the current read value.
		/// </summary>
		/// <remarks>This is only valid if <see cref="Find"/> has previously returned true.
		/// Initially the result will be the index of the first byte of actual data.  The value is updated after calls to
		/// <see cref="ReadInt"/>, <see cref="ReadShort"/> and <see cref="ReadLong"/>. </remarks>
		public int CurrentReadIndex
		{
			get { return _index; }
		}

		/// <summary>
		/// Get the number of bytes remaining to be read for the current value;
		/// </summary>
		public int UnreadCount
		{
			get
			{
				if ((_readValueStart > _data.Length) ||
					(_readValueStart < 4))
				{
					throw new ZipException("Find must be called before calling a Read method");
				}

				return _readValueStart + _readValueLength - _index;
			}
		}

		/// <summary>
		/// Find an extra data value
		/// </summary>
		/// <param name="headerID">The identifier for the value to find.</param>
		/// <returns>Returns true if the value was found; false otherwise.</returns>
		public bool Find(int headerID)
		{
			_readValueStart = _data.Length;
			_readValueLength = 0;
			_index = 0;

			int localLength = _readValueStart;
			int localTag = headerID - 1;

			// Trailing bytes that cant make up an entry (as there arent enough
			// bytes for a tag and length) are ignored!
			while ((localTag != headerID) && (_index < _data.Length - 3))
			{
				localTag = ReadShortInternal();
				localLength = ReadShortInternal();
				if (localTag != headerID)
				{
					_index += localLength;
				}
			}

			bool result = (localTag == headerID) && ((_index + localLength) <= _data.Length);

			if (result)
			{
				_readValueStart = _index;
				_readValueLength = localLength;
			}

			return result;
		}

		/// <summary>
		/// Add a new entry to extra data.
		/// </summary>
		/// <param name="taggedData">The <see cref="ITaggedData"/> value to add.</param>
		public void AddEntry(ITaggedData taggedData)
		{
			if (taggedData == null)
			{
				throw new ArgumentNullException(nameof(taggedData));
			}
			AddEntry(taggedData.TagID, taggedData.GetData());
		}

		/// <summary>
		/// Add a new entry to extra data
		/// </summary>
		/// <param name="headerID">The ID for this entry.</param>
		/// <param name="fieldData">The data to add.</param>
		/// <remarks>If the ID already exists its contents are replaced.</remarks>
		public void AddEntry(int headerID, byte[] fieldData)
		{
			if ((headerID > ushort.MaxValue) || (headerID < 0))
			{
				throw new ArgumentOutOfRangeException(nameof(headerID));
			}

			int addLength = (fieldData == null) ? 0 : fieldData.Length;

			if (addLength > ushort.MaxValue)
			{
				throw new ArgumentOutOfRangeException(nameof(fieldData), "exceeds maximum length");
			}

			// Test for new length before adjusting data.
			int newLength = _data.Length + addLength + 4;

			if (Find(headerID))
			{
				newLength -= (ValueLength + 4);
			}

			if (newLength > ushort.MaxValue)
			{
				throw new ZipException("Data exceeds maximum length");
			}

			Delete(headerID);

			byte[] newData = new byte[newLength];
			_data.CopyTo(newData, 0);
			int index = _data.Length;
			_data = newData;
			SetShort(ref index, headerID);
			SetShort(ref index, addLength);
			if (fieldData != null)
			{
				fieldData.CopyTo(newData, index);
			}
		}

		/// <summary>
		/// Start adding a new entry.
		/// </summary>
		/// <remarks>Add data using <see cref="AddData(byte[])"/>, <see cref="AddLeShort"/>, <see cref="AddLeInt"/>, or <see cref="AddLeLong"/>.
		/// The new entry is completed and actually added by calling <see cref="AddNewEntry"/></remarks>
		/// <seealso cref="AddEntry(ITaggedData)"/>
		public void StartNewEntry()
		{
			_newEntry = new MemoryStream();
		}

		/// <summary>
		/// Add entry data added since <see cref="StartNewEntry"/> using the ID passed.
		/// </summary>
		/// <param name="headerID">The identifier to use for this entry.</param>
		public void AddNewEntry(int headerID)
		{
			byte[] newData = _newEntry.ToArray();
			_newEntry = null;
			AddEntry(headerID, newData);
		}

		/// <summary>
		/// Add a byte of data to the pending new entry.
		/// </summary>
		/// <param name="data">The byte to add.</param>
		/// <seealso cref="StartNewEntry"/>
		public void AddData(byte data)
		{
			_newEntry.WriteByte(data);
		}

		/// <summary>
		/// Add data to a pending new entry.
		/// </summary>
		/// <param name="data">The data to add.</param>
		/// <seealso cref="StartNewEntry"/>
		public void AddData(byte[] data)
		{
			if (data == null)
			{
				throw new ArgumentNullException(nameof(data));
			}

			_newEntry.Write(data, 0, data.Length);
		}

		/// <summary>
		/// Add a short value in little endian order to the pending new entry.
		/// </summary>
		/// <param name="toAdd">The data to add.</param>
		/// <seealso cref="StartNewEntry"/>
		public void AddLeShort(int toAdd)
		{
			unchecked
			{
				_newEntry.WriteByte((byte)toAdd);
				_newEntry.WriteByte((byte)(toAdd >> 8));
			}
		}

		/// <summary>
		/// Add an integer value in little endian order to the pending new entry.
		/// </summary>
		/// <param name="toAdd">The data to add.</param>
		/// <seealso cref="StartNewEntry"/>
		public void AddLeInt(int toAdd)
		{
			unchecked
			{
				AddLeShort((short)toAdd);
				AddLeShort((short)(toAdd >> 16));
			}
		}

		/// <summary>
		/// Add a long value in little endian order to the pending new entry.
		/// </summary>
		/// <param name="toAdd">The data to add.</param>
		/// <seealso cref="StartNewEntry"/>
		public void AddLeLong(long toAdd)
		{
			unchecked
			{
				AddLeInt((int)(toAdd & 0xffffffff));
				AddLeInt((int)(toAdd >> 32));
			}
		}

		/// <summary>
		/// Delete an extra data field.
		/// </summary>
		/// <param name="headerID">The identifier of the field to delete.</param>
		/// <returns>Returns true if the field was found and deleted.</returns>
		public bool Delete(int headerID)
		{
			bool result = false;

			if (Find(headerID))
			{
				result = true;
				int trueStart = _readValueStart - 4;

				byte[] newData = new byte[_data.Length - (ValueLength + 4)];
				Array.Copy(_data, 0, newData, 0, trueStart);

				int trueEnd = trueStart + ValueLength + 4;
				Array.Copy(_data, trueEnd, newData, trueStart, _data.Length - trueEnd);
				_data = newData;
			}
			return result;
		}

		#region Reading Support

		/// <summary>
		/// Read a long in little endian form from the last <see cref="Find">found</see> data value
		/// </summary>
		/// <returns>Returns the long value read.</returns>
		public long ReadLong()
		{
			ReadCheck(8);
			return (ReadInt() & 0xffffffff) | (((long)ReadInt()) << 32);
		}

		/// <summary>
		/// Read an integer in little endian form from the last <see cref="Find">found</see> data value.
		/// </summary>
		/// <returns>Returns the integer read.</returns>
		public int ReadInt()
		{
			ReadCheck(4);

			int result = _data[_index] + (_data[_index + 1] << 8) +
				(_data[_index + 2] << 16) + (_data[_index + 3] << 24);
			_index += 4;
			return result;
		}

		/// <summary>
		/// Read a short value in little endian form from the last <see cref="Find">found</see> data value.
		/// </summary>
		/// <returns>Returns the short value read.</returns>
		public int ReadShort()
		{
			ReadCheck(2);
			int result = _data[_index] + (_data[_index + 1] << 8);
			_index += 2;
			return result;
		}

		/// <summary>
		/// Read a byte from an extra data
		/// </summary>
		/// <returns>The byte value read or -1 if the end of data has been reached.</returns>
		public int ReadByte()
		{
			int result = -1;
			if ((_index < _data.Length) && (_readValueStart + _readValueLength > _index))
			{
				result = _data[_index];
				_index += 1;
			}
			return result;
		}

		/// <summary>
		/// Skip data during reading.
		/// </summary>
		/// <param name="amount">The number of bytes to skip.</param>
		public void Skip(int amount)
		{
			ReadCheck(amount);
			_index += amount;
		}

		private void ReadCheck(int length)
		{
			if ((_readValueStart > _data.Length) ||
				(_readValueStart < 4))
			{
				throw new ZipException("Find must be called before calling a Read method");
			}

			if (_index > _readValueStart + _readValueLength - length)
			{
				throw new ZipException("End of extra data");
			}

			if (_index + length < 4)
			{
				throw new ZipException("Cannot read before start of tag");
			}
		}

		/// <summary>
		/// Internal form of <see cref="ReadShort"/> that reads data at any location.
		/// </summary>
		/// <returns>Returns the short value read.</returns>
		private int ReadShortInternal()
		{
			if (_index > _data.Length - 2)
			{
				throw new ZipException("End of extra data");
			}

			int result = _data[_index] + (_data[_index + 1] << 8);
			_index += 2;
			return result;
		}

		private void SetShort(ref int index, int source)
		{
			_data[index] = (byte)source;
			_data[index + 1] = (byte)(source >> 8);
			index += 2;
		}

		#endregion Reading Support

		#region IDisposable Members

		/// <summary>
		/// Dispose of this instance.
		/// </summary>
		public void Dispose()
		{
			if (_newEntry != null)
			{
				_newEntry.Dispose();
			}
		}

		#endregion IDisposable Members

		#region Instance Fields

		private int _index;
		private int _readValueStart;
		private int _readValueLength;

		private MemoryStream _newEntry;
		private byte[] _data;

		#endregion Instance Fields
	}
}


namespace ICSharpCode.SharpZipLib.Zip
{
	#region Keys Required Event Args

	/// <summary>
	/// Arguments used with KeysRequiredEvent
	/// </summary>
	public class KeysRequiredEventArgs : EventArgs
	{
		#region Constructors

		/// <summary>
		/// Initialise a new instance of <see cref="KeysRequiredEventArgs"></see>
		/// </summary>
		/// <param name="name">The name of the file for which keys are required.</param>
		public KeysRequiredEventArgs(string name)
		{
			fileName = name;
		}

		/// <summary>
		/// Initialise a new instance of <see cref="KeysRequiredEventArgs"></see>
		/// </summary>
		/// <param name="name">The name of the file for which keys are required.</param>
		/// <param name="keyValue">The current key value.</param>
		public KeysRequiredEventArgs(string name, byte[] keyValue)
		{
			fileName = name;
			key = keyValue;
		}

		#endregion Constructors

		#region Properties

		/// <summary>
		/// Gets the name of the file for which keys are required.
		/// </summary>
		public string FileName
		{
			get { return fileName; }
		}

		/// <summary>
		/// Gets or sets the key value
		/// </summary>
		public byte[] Key
		{
			get { return key; }
			set { key = value; }
		}

		#endregion Properties

		#region Instance Fields

		private readonly string fileName;
		private byte[] key;

		#endregion Instance Fields
	}

	#endregion Keys Required Event Args

	#region Test Definitions

	/// <summary>
	/// The strategy to apply to testing.
	/// </summary>
	public enum TestStrategy
	{
		/// <summary>
		/// Find the first error only.
		/// </summary>
		FindFirstError,

		/// <summary>
		/// Find all possible errors.
		/// </summary>
		FindAllErrors,
	}

	/// <summary>
	/// The operation in progress reported by a <see cref="ZipTestResultHandler"/> during testing.
	/// </summary>
	/// <seealso cref="ZipFile.TestArchive(bool)">TestArchive</seealso>
	public enum TestOperation
	{
		/// <summary>
		/// Setting up testing.
		/// </summary>
		Initialising,

		/// <summary>
		/// Testing an individual entries header
		/// </summary>
		EntryHeader,

		/// <summary>
		/// Testing an individual entries data
		/// </summary>
		EntryData,

		/// <summary>
		/// Testing an individual entry has completed.
		/// </summary>
		EntryComplete,

		/// <summary>
		/// Running miscellaneous tests
		/// </summary>
		MiscellaneousTests,

		/// <summary>
		/// Testing is complete
		/// </summary>
		Complete,
	}

	/// <summary>
	/// Status returned by <see cref="ZipTestResultHandler"/> during testing.
	/// </summary>
	/// <seealso cref="ZipFile.TestArchive(bool)">TestArchive</seealso>
	public class TestStatus
	{
		#region Constructors

		/// <summary>
		/// Initialise a new instance of <see cref="TestStatus"/>
		/// </summary>
		/// <param name="file">The <see cref="ZipFile"/> this status applies to.</param>
		public TestStatus(ZipFile file)
		{
			file_ = file;
		}

		#endregion Constructors

		#region Properties

		/// <summary>
		/// Get the current <see cref="TestOperation"/> in progress.
		/// </summary>
		public TestOperation Operation
		{
			get { return operation_; }
		}

		/// <summary>
		/// Get the <see cref="ZipFile"/> this status is applicable to.
		/// </summary>
		public ZipFile File
		{
			get { return file_; }
		}

		/// <summary>
		/// Get the current/last entry tested.
		/// </summary>
		public ZipEntry Entry
		{
			get { return entry_; }
		}

		/// <summary>
		/// Get the number of errors detected so far.
		/// </summary>
		public int ErrorCount
		{
			get { return errorCount_; }
		}

		/// <summary>
		/// Get the number of bytes tested so far for the current entry.
		/// </summary>
		public long BytesTested
		{
			get { return bytesTested_; }
		}

		/// <summary>
		/// Get a value indicating whether the last entry test was valid.
		/// </summary>
		public bool EntryValid
		{
			get { return entryValid_; }
		}

		#endregion Properties

		#region Internal API

		internal void AddError()
		{
			errorCount_++;
			entryValid_ = false;
		}

		internal void SetOperation(TestOperation operation)
		{
			operation_ = operation;
		}

		internal void SetEntry(ZipEntry entry)
		{
			entry_ = entry;
			entryValid_ = true;
			bytesTested_ = 0;
		}

		internal void SetBytesTested(long value)
		{
			bytesTested_ = value;
		}

		#endregion Internal API

		#region Instance Fields

		private readonly ZipFile file_;
		private ZipEntry entry_;
		private bool entryValid_;
		private int errorCount_;
		private long bytesTested_;
		private TestOperation operation_;

		#endregion Instance Fields
	}

	/// <summary>
	/// Delegate invoked during <see cref="ZipFile.TestArchive(bool, TestStrategy, ZipTestResultHandler)">testing</see> if supplied indicating current progress and status.
	/// </summary>
	/// <remarks>If the message is non-null an error has occured.  If the message is null
	/// the operation as found in <see cref="TestStatus">status</see> has started.</remarks>
	public delegate void ZipTestResultHandler(TestStatus status, string message);

	#endregion Test Definitions

	#region Update Definitions

	/// <summary>
	/// The possible ways of <see cref="ZipFile.CommitUpdate()">applying updates</see> to an archive.
	/// </summary>
	public enum FileUpdateMode
	{
		/// <summary>
		/// Perform all updates on temporary files ensuring that the original file is saved.
		/// </summary>
		Safe,

		/// <summary>
		/// Update the archive directly, which is faster but less safe.
		/// </summary>
		Direct,
	}

	#endregion Update Definitions

	#region ZipFile Class

	/// <summary>
	/// This class represents a Zip archive.  You can ask for the contained
	/// entries, or get an input stream for a file entry.  The entry is
	/// automatically decompressed.
	///
	/// You can also update the archive adding or deleting entries.
	///
	/// This class is thread safe for input:  You can open input streams for arbitrary
	/// entries in different threads.
	/// <br/>
	/// <br/>Author of the original java version : Jochen Hoenicke
	/// </summary>
	/// <example>
	/// <code>
	///
	/// using ICSharpCode.SharpZipLib.Zip;
	///
	/// class MainClass
	/// {
	/// 	static public void Main(string[] args)
	/// 	{
	/// 		using (ZipFile zFile = new ZipFile(args[0])) {
	/// 			Console.WriteLine("Listing of : " + zFile.Name);
	/// 			Console.WriteLine("");
	/// 			Console.WriteLine("Raw Size    Size      Date     Time     Name");
	/// 			Console.WriteLine("--------  --------  --------  ------  ---------");
	/// 			foreach (ZipEntry e in zFile) {
	/// 				if ( e.IsFile ) {
	/// 					DateTime d = e.DateTime;
	/// 					Console.WriteLine("{0, -10}{1, -10}{2}  {3}   {4}", e.Size, e.CompressedSize,
	/// 						d.ToString("dd-MM-yy"), d.ToString("HH:mm"),
	/// 						e.Name);
	/// 				}
	/// 			}
	/// 		}
	/// 	}
	/// }
	/// </code>
	/// </example>
	public class ZipFile : IEnumerable, IDisposable
	{
		#region KeyHandling

		/// <summary>
		/// Delegate for handling keys/password setting during compression/decompression.
		/// </summary>
		public delegate void KeysRequiredEventHandler(
			object sender,
			KeysRequiredEventArgs e
		);

		/// <summary>
		/// Event handler for handling encryption keys.
		/// </summary>
		public KeysRequiredEventHandler KeysRequired;

		/// <summary>
		/// Handles getting of encryption keys when required.
		/// </summary>
		/// <param name="fileName">The file for which encryption keys are required.</param>
		private void OnKeysRequired(string fileName)
		{
			if (KeysRequired != null)
			{
				var krea = new KeysRequiredEventArgs(fileName, key);
				KeysRequired(this, krea);
				key = krea.Key;
			}
		}

		/// <summary>
		/// Get/set the encryption key value.
		/// </summary>
		private byte[] Key
		{
			get { return key; }
			set { key = value; }
		}

		/// <summary>
		/// Password to be used for encrypting/decrypting files.
		/// </summary>
		/// <remarks>Set to null if no password is required.</remarks>
		public string Password
		{
			set
			{
				if (string.IsNullOrEmpty(value))
				{
					key = null;
				}
				else
				{
					key = PkzipClassic.GenerateKeys(ZipStrings.ConvertToArray(value));
				}

				rawPassword_ = value;
			}
		}

		/// <summary>
		/// Get a value indicating whether encryption keys are currently available.
		/// </summary>
		private bool HaveKeys
		{
			get { return key != null; }
		}

		#endregion KeyHandling

		#region Constructors

		/// <summary>
		/// Opens a Zip file with the given name for reading.
		/// </summary>
		/// <param name="name">The name of the file to open.</param>
		/// <exception cref="ArgumentNullException">The argument supplied is null.</exception>
		/// <exception cref="IOException">
		/// An i/o error occurs
		/// </exception>
		/// <exception cref="ZipException">
		/// The file doesn't contain a valid zip archive.
		/// </exception>
		public ZipFile(string name)
		{
			name_ = name ?? throw new ArgumentNullException(nameof(name));

			baseStream_ = File.Open(name, FileMode.Open, FileAccess.Read, FileShare.Read);
			isStreamOwner = true;

			try
			{
				ReadEntries();
			}
			catch
			{
				DisposeInternal(true);
				throw;
			}
		}

		/// <summary>
		/// Opens a Zip file reading the given <see cref="FileStream"/>.
		/// </summary>
		/// <param name="file">The <see cref="FileStream"/> to read archive data from.</param>
		/// <exception cref="ArgumentNullException">The supplied argument is null.</exception>
		/// <exception cref="IOException">
		/// An i/o error occurs.
		/// </exception>
		/// <exception cref="ZipException">
		/// The file doesn't contain a valid zip archive.
		/// </exception>
		public ZipFile(FileStream file) :
			this(file, false)
		{

		}

		/// <summary>
		/// Opens a Zip file reading the given <see cref="FileStream"/>.
		/// </summary>
		/// <param name="file">The <see cref="FileStream"/> to read archive data from.</param>
		/// <param name="leaveOpen">true to leave the <see cref="FileStream">file</see> open when the ZipFile is disposed, false to dispose of it</param>
		/// <exception cref="ArgumentNullException">The supplied argument is null.</exception>
		/// <exception cref="IOException">
		/// An i/o error occurs.
		/// </exception>
		/// <exception cref="ZipException">
		/// The file doesn't contain a valid zip archive.
		/// </exception>
		public ZipFile(FileStream file, bool leaveOpen)
		{
			if (file == null)
			{
				throw new ArgumentNullException(nameof(file));
			}

			if (!file.CanSeek)
			{
				throw new ArgumentException("Stream is not seekable", nameof(file));
			}

			baseStream_ = file;
			name_ = file.Name;
			isStreamOwner = !leaveOpen;

			try
			{
				ReadEntries();
			}
			catch
			{
				DisposeInternal(true);
				throw;
			}
		}

		/// <summary>
		/// Opens a Zip file reading the given <see cref="Stream"/>.
		/// </summary>
		/// <param name="stream">The <see cref="Stream"/> to read archive data from.</param>
		/// <exception cref="IOException">
		/// An i/o error occurs
		/// </exception>
		/// <exception cref="ZipException">
		/// The stream doesn't contain a valid zip archive.<br/>
		/// </exception>
		/// <exception cref="ArgumentException">
		/// The <see cref="Stream">stream</see> doesnt support seeking.
		/// </exception>
		/// <exception cref="ArgumentNullException">
		/// The <see cref="Stream">stream</see> argument is null.
		/// </exception>
		public ZipFile(Stream stream) :
			this(stream, false)
		{

		}

		/// <summary>
		/// Opens a Zip file reading the given <see cref="Stream"/>.
		/// </summary>
		/// <param name="stream">The <see cref="Stream"/> to read archive data from.</param>
		/// <param name="leaveOpen">true to leave the <see cref="Stream">stream</see> open when the ZipFile is disposed, false to dispose of it</param>
		/// <exception cref="IOException">
		/// An i/o error occurs
		/// </exception>
		/// <exception cref="ZipException">
		/// The stream doesn't contain a valid zip archive.<br/>
		/// </exception>
		/// <exception cref="ArgumentException">
		/// The <see cref="Stream">stream</see> doesnt support seeking.
		/// </exception>
		/// <exception cref="ArgumentNullException">
		/// The <see cref="Stream">stream</see> argument is null.
		/// </exception>
		public ZipFile(Stream stream, bool leaveOpen)
		{
			if (stream == null)
			{
				throw new ArgumentNullException(nameof(stream));
			}

			if (!stream.CanSeek)
			{
				throw new ArgumentException("Stream is not seekable", nameof(stream));
			}

			baseStream_ = stream;
			isStreamOwner = !leaveOpen;

			if (baseStream_.Length > 0)
			{
				try
				{
					ReadEntries();
				}
				catch
				{
					DisposeInternal(true);
					throw;
				}
			}
			else
			{
				entries_ = Empty.Array<ZipEntry>();
				isNewArchive_ = true;
			}
		}

		/// <summary>
		/// Initialises a default <see cref="ZipFile"/> instance with no entries and no file storage.
		/// </summary>
		internal ZipFile()
		{
			entries_ = Empty.Array<ZipEntry>();
			isNewArchive_ = true;
		}

		#endregion Constructors

		#region Destructors and Closing

		/// <summary>
		/// Finalize this instance.
		/// </summary>
		~ZipFile()
		{
			Dispose(false);
		}

		/// <summary>
		/// Closes the ZipFile.  If the stream is <see cref="IsStreamOwner">owned</see> then this also closes the underlying input stream.
		/// Once closed, no further instance methods should be called.
		/// </summary>
		/// <exception cref="System.IO.IOException">
		/// An i/o error occurs.
		/// </exception>
		public void Close()
		{
			DisposeInternal(true);
			GC.SuppressFinalize(this);
		}

		#endregion Destructors and Closing

		#region Creators

		/// <summary>
		/// Create a new <see cref="ZipFile"/> whose data will be stored in a file.
		/// </summary>
		/// <param name="fileName">The name of the archive to create.</param>
		/// <returns>Returns the newly created <see cref="ZipFile"/></returns>
		/// <exception cref="ArgumentNullException"><paramref name="fileName"></paramref> is null</exception>
		public static ZipFile Create(string fileName)
		{
			if (fileName == null)
			{
				throw new ArgumentNullException(nameof(fileName));
			}

			FileStream fs = File.Create(fileName);

			return new ZipFile
			{
				name_ = fileName,
				baseStream_ = fs,
				isStreamOwner = true
			};
		}

		/// <summary>
		/// Create a new <see cref="ZipFile"/> whose data will be stored on a stream.
		/// </summary>
		/// <param name="outStream">The stream providing data storage.</param>
		/// <returns>Returns the newly created <see cref="ZipFile"/></returns>
		/// <exception cref="ArgumentNullException"><paramref name="outStream"> is null</paramref></exception>
		/// <exception cref="ArgumentException"><paramref name="outStream"> doesnt support writing.</paramref></exception>
		public static ZipFile Create(Stream outStream)
		{
			if (outStream == null)
			{
				throw new ArgumentNullException(nameof(outStream));
			}

			if (!outStream.CanWrite)
			{
				throw new ArgumentException("Stream is not writeable", nameof(outStream));
			}

			if (!outStream.CanSeek)
			{
				throw new ArgumentException("Stream is not seekable", nameof(outStream));
			}

			var result = new ZipFile
			{
				baseStream_ = outStream
			};
			return result;
		}

		#endregion Creators

		#region Properties

		/// <summary>
		/// Get/set a flag indicating if the underlying stream is owned by the ZipFile instance.
		/// If the flag is true then the stream will be closed when <see cref="Close">Close</see> is called.
		/// </summary>
		/// <remarks>
		/// The default value is true in all cases.
		/// </remarks>
		public bool IsStreamOwner
		{
			get { return isStreamOwner; }
			set { isStreamOwner = value; }
		}

		/// <summary>
		/// Get a value indicating whether
		/// this archive is embedded in another file or not.
		/// </summary>
		public bool IsEmbeddedArchive
		{
			// Not strictly correct in all circumstances currently
			get { return offsetOfFirstEntry > 0; }
		}

		/// <summary>
		/// Get a value indicating that this archive is a new one.
		/// </summary>
		public bool IsNewArchive
		{
			get { return isNewArchive_; }
		}

		/// <summary>
		/// Gets the comment for the zip file.
		/// </summary>
		public string ZipFileComment
		{
			get { return comment_; }
		}

		/// <summary>
		/// Gets the name of this zip file.
		/// </summary>
		public string Name
		{
			get { return name_; }
		}

		/// <summary>
		/// Gets the number of entries in this zip file.
		/// </summary>
		/// <exception cref="InvalidOperationException">
		/// The Zip file has been closed.
		/// </exception>
		[Obsolete("Use the Count property instead")]
		public int Size
		{
			get
			{
				return entries_.Length;
			}
		}

		/// <summary>
		/// Get the number of entries contained in this <see cref="ZipFile"/>.
		/// </summary>
		public long Count
		{
			get
			{
				return entries_.Length;
			}
		}

		/// <summary>
		/// Indexer property for ZipEntries
		/// </summary>
		[System.Runtime.CompilerServices.IndexerNameAttribute("EntryByIndex")]
		public ZipEntry this[int index]
		{
			get
			{
				return (ZipEntry)entries_[index].Clone();
			}
		}

		#endregion Properties

		#region Input Handling

		/// <summary>
		/// Gets an enumerator for the Zip entries in this Zip file.
		/// </summary>
		/// <returns>Returns an <see cref="IEnumerator"/> for this archive.</returns>
		/// <exception cref="ObjectDisposedException">
		/// The Zip file has been closed.
		/// </exception>
		public IEnumerator GetEnumerator()
		{
			if (isDisposed_)
			{
				throw new ObjectDisposedException("ZipFile");
			}

			return new ZipEntryEnumerator(entries_);
		}

		/// <summary>
		/// Return the index of the entry with a matching name
		/// </summary>
		/// <param name="name">Entry name to find</param>
		/// <param name="ignoreCase">If true the comparison is case insensitive</param>
		/// <returns>The index position of the matching entry or -1 if not found</returns>
		/// <exception cref="ObjectDisposedException">
		/// The Zip file has been closed.
		/// </exception>
		public int FindEntry(string name, bool ignoreCase)
		{
			if (isDisposed_)
			{
				throw new ObjectDisposedException("ZipFile");
			}

			// TODO: This will be slow as the next ice age for huge archives!
			for (int i = 0; i < entries_.Length; i++)
			{
				if (string.Compare(name, entries_[i].Name, ignoreCase ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal) == 0)
				{
					return i;
				}
			}
			return -1;
		}

		/// <summary>
		/// Searches for a zip entry in this archive with the given name.
		/// String comparisons are case insensitive
		/// </summary>
		/// <param name="name">
		/// The name to find. May contain directory components separated by slashes ('/').
		/// </param>
		/// <returns>
		/// A clone of the zip entry, or null if no entry with that name exists.
		/// </returns>
		/// <exception cref="ObjectDisposedException">
		/// The Zip file has been closed.
		/// </exception>
		public ZipEntry GetEntry(string name)
		{
			if (isDisposed_)
			{
				throw new ObjectDisposedException("ZipFile");
			}

			int index = FindEntry(name, true);
			return (index >= 0) ? (ZipEntry)entries_[index].Clone() : null;
		}

		/// <summary>
		/// Gets an input stream for reading the given zip entry data in an uncompressed form.
		/// Normally the <see cref="ZipEntry"/> should be an entry returned by GetEntry().
		/// </summary>
		/// <param name="entry">The <see cref="ZipEntry"/> to obtain a data <see cref="Stream"/> for</param>
		/// <returns>An input <see cref="Stream"/> containing data for this <see cref="ZipEntry"/></returns>
		/// <exception cref="ObjectDisposedException">
		/// The ZipFile has already been closed
		/// </exception>
		/// <exception cref="ICSharpCode.SharpZipLib.Zip.ZipException">
		/// The compression method for the entry is unknown
		/// </exception>
		/// <exception cref="IndexOutOfRangeException">
		/// The entry is not found in the ZipFile
		/// </exception>
		public Stream GetInputStream(ZipEntry entry)
		{
			if (entry == null)
			{
				throw new ArgumentNullException(nameof(entry));
			}

			if (isDisposed_)
			{
				throw new ObjectDisposedException("ZipFile");
			}

			long index = entry.ZipFileIndex;
			if ((index < 0) || (index >= entries_.Length) || (entries_[index].Name != entry.Name))
			{
				index = FindEntry(entry.Name, true);
				if (index < 0)
				{
					throw new ZipException("Entry cannot be found");
				}
			}
			return GetInputStream(index);
		}

		/// <summary>
		/// Creates an input stream reading a zip entry
		/// </summary>
		/// <param name="entryIndex">The index of the entry to obtain an input stream for.</param>
		/// <returns>
		/// An input <see cref="Stream"/> containing data for this <paramref name="entryIndex"/>
		/// </returns>
		/// <exception cref="ObjectDisposedException">
		/// The ZipFile has already been closed
		/// </exception>
		/// <exception cref="ICSharpCode.SharpZipLib.Zip.ZipException">
		/// The compression method for the entry is unknown
		/// </exception>
		/// <exception cref="IndexOutOfRangeException">
		/// The entry is not found in the ZipFile
		/// </exception>
		public Stream GetInputStream(long entryIndex)
		{
			if (isDisposed_)
			{
				throw new ObjectDisposedException("ZipFile");
			}

			long start = LocateEntry(entries_[entryIndex]);
			CompressionMethod method = entries_[entryIndex].CompressionMethod;
			Stream result = new PartialInputStream(this, start, entries_[entryIndex].CompressedSize);

			if (entries_[entryIndex].IsCrypted == true)
			{
				result = CreateAndInitDecryptionStream(result, entries_[entryIndex]);
				if (result == null)
				{
					throw new ZipException("Unable to decrypt this entry");
				}
			}

			switch (method)
			{
				case CompressionMethod.Stored:
					// read as is.
					break;

				case CompressionMethod.Deflated:
					// No need to worry about ownership and closing as underlying stream close does nothing.
					result = new InflaterInputStream(result, new Inflater(true));
					break;

				case CompressionMethod.BZip2:
					result = new BZip2.BZip2InputStream(result);
					break;

				default:
					throw new ZipException("Unsupported compression method " + method);
			}

			return result;
		}

		#endregion Input Handling

		#region Archive Testing

		/// <summary>
		/// Test an archive for integrity/validity
		/// </summary>
		/// <param name="testData">Perform low level data Crc check</param>
		/// <returns>true if all tests pass, false otherwise</returns>
		/// <remarks>Testing will terminate on the first error found.</remarks>
		public bool TestArchive(bool testData)
		{
			return TestArchive(testData, TestStrategy.FindFirstError, null);
		}

		/// <summary>
		/// Test an archive for integrity/validity
		/// </summary>
		/// <param name="testData">Perform low level data Crc check</param>
		/// <param name="strategy">The <see cref="TestStrategy"></see> to apply.</param>
		/// <param name="resultHandler">The <see cref="ZipTestResultHandler"></see> handler to call during testing.</param>
		/// <returns>true if all tests pass, false otherwise</returns>
		/// <exception cref="ObjectDisposedException">The object has already been closed.</exception>
		public bool TestArchive(bool testData, TestStrategy strategy, ZipTestResultHandler resultHandler)
		{
			if (isDisposed_)
			{
				throw new ObjectDisposedException("ZipFile");
			}

			var status = new TestStatus(this);

			resultHandler?.Invoke(status, null);

			HeaderTest test = testData ? (HeaderTest.Header | HeaderTest.Extract) : HeaderTest.Header;

			bool testing = true;

			try
			{
				int entryIndex = 0;

				while (testing && (entryIndex < Count))
				{
					if (resultHandler != null)
					{
						status.SetEntry(this[entryIndex]);
						status.SetOperation(TestOperation.EntryHeader);
						resultHandler(status, null);
					}

					try
					{
						TestLocalHeader(this[entryIndex], test);
					}
					catch (ZipException ex)
					{
						status.AddError();

						resultHandler?.Invoke(status, $"Exception during test - '{ex.Message}'");

						testing &= strategy != TestStrategy.FindFirstError;
					}

					if (testing && testData && this[entryIndex].IsFile)
					{
						// Don't check CRC for AES encrypted archives
						var checkCRC = this[entryIndex].AESKeySize == 0;

						if (resultHandler != null)
						{
							status.SetOperation(TestOperation.EntryData);
							resultHandler(status, null);
						}

						var crc = new Crc32();

						using (Stream entryStream = this.GetInputStream(this[entryIndex]))
						{
							byte[] buffer = new byte[4096];
							long totalBytes = 0;
							int bytesRead;
							while ((bytesRead = entryStream.Read(buffer, 0, buffer.Length)) > 0)
							{
								if (checkCRC)
								{
									crc.Update(new ArraySegment<byte>(buffer, 0, bytesRead));
								}

								if (resultHandler != null)
								{
									totalBytes += bytesRead;
									status.SetBytesTested(totalBytes);
									resultHandler(status, null);
								}
							}
						}

						if (checkCRC && this[entryIndex].Crc != crc.Value)
						{
							status.AddError();

							resultHandler?.Invoke(status, "CRC mismatch");

							testing &= strategy != TestStrategy.FindFirstError;
						}

						if ((this[entryIndex].Flags & (int)GeneralBitFlags.Descriptor) != 0)
						{
							var helper = new ZipHelperStream(baseStream_);
							var data = new DescriptorData();
							helper.ReadDataDescriptor(this[entryIndex].LocalHeaderRequiresZip64, data);

							if (checkCRC && this[entryIndex].Crc != data.Crc)
							{
								status.AddError();
								resultHandler?.Invoke(status, "Descriptor CRC mismatch");
							}

							if (this[entryIndex].CompressedSize != data.CompressedSize)
							{
								status.AddError();
								resultHandler?.Invoke(status, "Descriptor compressed size mismatch");
							}

							if (this[entryIndex].Size != data.Size)
							{
								status.AddError();
								resultHandler?.Invoke(status, "Descriptor size mismatch");
							}
						}
					}

					if (resultHandler != null)
					{
						status.SetOperation(TestOperation.EntryComplete);
						resultHandler(status, null);
					}

					entryIndex += 1;
				}

				if (resultHandler != null)
				{
					status.SetOperation(TestOperation.MiscellaneousTests);
					resultHandler(status, null);
				}

				// TODO: the 'Corrina Johns' test where local headers are missing from
				// the central directory.  They are therefore invisible to many archivers.
			}
			catch (Exception ex)
			{
				status.AddError();

				resultHandler?.Invoke(status, $"Exception during test - '{ex.Message}'");
			}

			if (resultHandler != null)
			{
				status.SetOperation(TestOperation.Complete);
				status.SetEntry(null);
				resultHandler(status, null);
			}

			return (status.ErrorCount == 0);
		}

		[Flags]
		private enum HeaderTest
		{
			Extract = 0x01,     // Check that this header represents an entry whose data can be extracted
			Header = 0x02,     // Check that this header contents are valid
		}

		/// <summary>
		/// Test a local header against that provided from the central directory
		/// </summary>
		/// <param name="entry">
		/// The entry to test against
		/// </param>
		/// <param name="tests">The type of <see cref="HeaderTest">tests</see> to carry out.</param>
		/// <returns>The offset of the entries data in the file</returns>
		private long TestLocalHeader(ZipEntry entry, HeaderTest tests)
		{
			lock (baseStream_)
			{
				bool testHeader = (tests & HeaderTest.Header) != 0;
				bool testData = (tests & HeaderTest.Extract) != 0;

				var entryAbsOffset = offsetOfFirstEntry + entry.Offset;
				
				baseStream_.Seek(entryAbsOffset, SeekOrigin.Begin);
				var signature = (int)ReadLEUint();

				if (signature != ZipConstants.LocalHeaderSignature)
				{
					throw new ZipException(string.Format("Wrong local header signature at 0x{0:x}, expected 0x{1:x8}, actual 0x{2:x8}",
						entryAbsOffset, ZipConstants.LocalHeaderSignature, signature));
				}

				var extractVersion = (short)(ReadLEUshort() & 0x00ff);
				var localFlags = (short)ReadLEUshort();
				var compressionMethod = (short)ReadLEUshort();
				var fileTime = (short)ReadLEUshort();
				var fileDate = (short)ReadLEUshort();
				uint crcValue = ReadLEUint();
				long compressedSize = ReadLEUint();
				long size = ReadLEUint();
				int storedNameLength = ReadLEUshort();
				int extraDataLength = ReadLEUshort();

				byte[] nameData = new byte[storedNameLength];
				StreamUtils.ReadFully(baseStream_, nameData);

				byte[] extraData = new byte[extraDataLength];
				StreamUtils.ReadFully(baseStream_, extraData);

				var localExtraData = new ZipExtraData(extraData);

				// Extra data / zip64 checks
				if (localExtraData.Find(1))
				{
					// 2010-03-04 Forum 10512: removed checks for version >= ZipConstants.VersionZip64
					// and size or compressedSize = MaxValue, due to rogue creators.

					size = localExtraData.ReadLong();
					compressedSize = localExtraData.ReadLong();

					if ((localFlags & (int)GeneralBitFlags.Descriptor) != 0)
					{
						// These may be valid if patched later
						if ((size != -1) && (size != entry.Size))
						{
							throw new ZipException("Size invalid for descriptor");
						}

						if ((compressedSize != -1) && (compressedSize != entry.CompressedSize))
						{
							throw new ZipException("Compressed size invalid for descriptor");
						}
					}
				}
				else
				{
					// No zip64 extra data but entry requires it.
					if ((extractVersion >= ZipConstants.VersionZip64) &&
						(((uint)size == uint.MaxValue) || ((uint)compressedSize == uint.MaxValue)))
					{
						throw new ZipException("Required Zip64 extended information missing");
					}
				}

				if (testData)
				{
					if (entry.IsFile)
					{
						if (!entry.IsCompressionMethodSupported())
						{
							throw new ZipException("Compression method not supported");
						}

						if ((extractVersion > ZipConstants.VersionMadeBy)
							|| ((extractVersion > 20) && (extractVersion < ZipConstants.VersionZip64)))
						{
							throw new ZipException(string.Format("Version required to extract this entry not supported ({0})", extractVersion));
						}

						if ((localFlags & (int)(GeneralBitFlags.Patched | GeneralBitFlags.StrongEncryption | GeneralBitFlags.EnhancedCompress | GeneralBitFlags.HeaderMasked)) != 0)
						{
							throw new ZipException("The library does not support the zip version required to extract this entry");
						}
					}
				}

				if (testHeader)
				{
					if ((extractVersion <= 63) &&   // Ignore later versions as we dont know about them..
						(extractVersion != 10) &&
						(extractVersion != 11) &&
						(extractVersion != 20) &&
						(extractVersion != 21) &&
						(extractVersion != 25) &&
						(extractVersion != 27) &&
						(extractVersion != 45) &&
						(extractVersion != 46) &&
						(extractVersion != 50) &&
						(extractVersion != 51) &&
						(extractVersion != 52) &&
						(extractVersion != 61) &&
						(extractVersion != 62) &&
						(extractVersion != 63)
						)
					{
						throw new ZipException(string.Format("Version required to extract this entry is invalid ({0})", extractVersion));
					}

					// Local entry flags dont have reserved bit set on.
					if ((localFlags & (int)(GeneralBitFlags.ReservedPKware4 | GeneralBitFlags.ReservedPkware14 | GeneralBitFlags.ReservedPkware15)) != 0)
					{
						throw new ZipException("Reserved bit flags cannot be set.");
					}

					// Encryption requires extract version >= 20
					if (((localFlags & (int)GeneralBitFlags.Encrypted) != 0) && (extractVersion < 20))
					{
						throw new ZipException(string.Format("Version required to extract this entry is too low for encryption ({0})", extractVersion));
					}

					// Strong encryption requires encryption flag to be set and extract version >= 50.
					if ((localFlags & (int)GeneralBitFlags.StrongEncryption) != 0)
					{
						if ((localFlags & (int)GeneralBitFlags.Encrypted) == 0)
						{
							throw new ZipException("Strong encryption flag set but encryption flag is not set");
						}

						if (extractVersion < 50)
						{
							throw new ZipException(string.Format("Version required to extract this entry is too low for encryption ({0})", extractVersion));
						}
					}

					// Patched entries require extract version >= 27
					if (((localFlags & (int)GeneralBitFlags.Patched) != 0) && (extractVersion < 27))
					{
						throw new ZipException(string.Format("Patched data requires higher version than ({0})", extractVersion));
					}

					// Central header flags match local entry flags.
					if (localFlags != entry.Flags)
					{
						throw new ZipException("Central header/local header flags mismatch");
					}

					// Central header compression method matches local entry
					if (entry.CompressionMethodForHeader != (CompressionMethod)compressionMethod)
					{
						throw new ZipException("Central header/local header compression method mismatch");
					}

					if (entry.Version != extractVersion)
					{
						throw new ZipException("Extract version mismatch");
					}

					// Strong encryption and extract version match
					if ((localFlags & (int)GeneralBitFlags.StrongEncryption) != 0)
					{
						if (extractVersion < 62)
						{
							throw new ZipException("Strong encryption flag set but version not high enough");
						}
					}

					if ((localFlags & (int)GeneralBitFlags.HeaderMasked) != 0)
					{
						if ((fileTime != 0) || (fileDate != 0))
						{
							throw new ZipException("Header masked set but date/time values non-zero");
						}
					}

					if ((localFlags & (int)GeneralBitFlags.Descriptor) == 0)
					{
						if (crcValue != (uint)entry.Crc)
						{
							throw new ZipException("Central header/local header crc mismatch");
						}
					}

					// Crc valid for empty entry.
					// This will also apply to streamed entries where size isnt known and the header cant be patched
					if ((size == 0) && (compressedSize == 0))
					{
						if (crcValue != 0)
						{
							throw new ZipException("Invalid CRC for empty entry");
						}
					}

					// TODO: make test more correct...  can't compare lengths as was done originally as this can fail for MBCS strings
					// Assuming a code page at this point is not valid?  Best is to store the name length in the ZipEntry probably
					if (entry.Name.Length > storedNameLength)
					{
						throw new ZipException("File name length mismatch");
					}

					// Name data has already been read convert it and compare.
					string localName = ZipStrings.ConvertToStringExt(localFlags, nameData);

					// Central directory and local entry name match
					if (localName != entry.Name)
					{
						throw new ZipException("Central header and local header file name mismatch");
					}

					// Directories have zero actual size but can have compressed size
					if (entry.IsDirectory)
					{
						if (size > 0)
						{
							throw new ZipException("Directory cannot have size");
						}

						// There may be other cases where the compressed size can be greater than this?
						// If so until details are known we will be strict.
						if (entry.IsCrypted)
						{
							if (compressedSize > entry.EncryptionOverheadSize + 2)
							{
								throw new ZipException("Directory compressed size invalid");
							}
						}
						else if (compressedSize > 2)
						{
							// When not compressed the directory size can validly be 2 bytes
							// if the true size wasn't known when data was originally being written.
							// NOTE: Versions of the library 0.85.4 and earlier always added 2 bytes
							throw new ZipException("Directory compressed size invalid");
						}
					}

					if (!ZipNameTransform.IsValidName(localName, true))
					{
						throw new ZipException("Name is invalid");
					}
				}

				// Tests that apply to both data and header.

				// Size can be verified only if it is known in the local header.
				// it will always be known in the central header.
				if (((localFlags & (int)GeneralBitFlags.Descriptor) == 0) ||
					((size > 0 || compressedSize > 0) && entry.Size > 0))
				{
					if ((size != 0)
						&& (size != entry.Size))
					{
						throw new ZipException(
							string.Format("Size mismatch between central header({0}) and local header({1})",
								entry.Size, size));
					}

					if ((compressedSize != 0)
						&& (compressedSize != entry.CompressedSize && compressedSize != 0xFFFFFFFF && compressedSize != -1))
					{
						throw new ZipException(
							string.Format("Compressed size mismatch between central header({0}) and local header({1})",
							entry.CompressedSize, compressedSize));
					}
				}

				int extraLength = storedNameLength + extraDataLength;
				return offsetOfFirstEntry + entry.Offset + ZipConstants.LocalHeaderBaseSize + extraLength;
			}
		}

		#endregion Archive Testing

		#region Updating

		private const int DefaultBufferSize = 4096;

		/// <summary>
		/// The kind of update to apply.
		/// </summary>
		private enum UpdateCommand
		{
			Copy,       // Copy original file contents.
			Modify,     // Change encryption, compression, attributes, name, time etc, of an existing file.
			Add,        // Add a new file to the archive.
		}

		#region Properties

		/// <summary>
		/// Get / set the <see cref="INameTransform"/> to apply to names when updating.
		/// </summary>
		public INameTransform NameTransform
		{
			get
			{
				return updateEntryFactory_.NameTransform;
			}

			set
			{
				updateEntryFactory_.NameTransform = value;
			}
		}

		/// <summary>
		/// Get/set the <see cref="IEntryFactory"/> used to generate <see cref="ZipEntry"/> values
		/// during updates.
		/// </summary>
		public IEntryFactory EntryFactory
		{
			get
			{
				return updateEntryFactory_;
			}

			set
			{
				if (value == null)
				{
					updateEntryFactory_ = new ZipEntryFactory();
				}
				else
				{
					updateEntryFactory_ = value;
				}
			}
		}

		/// <summary>
		/// Get /set the buffer size to be used when updating this zip file.
		/// </summary>
		public int BufferSize
		{
			get { return bufferSize_; }
			set
			{
				if (value < 1024)
				{
					throw new ArgumentOutOfRangeException(nameof(value), "cannot be below 1024");
				}

				if (bufferSize_ != value)
				{
					bufferSize_ = value;
					copyBuffer_ = null;
				}
			}
		}

		/// <summary>
		/// Get a value indicating an update has <see cref="BeginUpdate()">been started</see>.
		/// </summary>
		public bool IsUpdating
		{
			get { return updates_ != null; }
		}

		/// <summary>
		/// Get / set a value indicating how Zip64 Extension usage is determined when adding entries.
		/// </summary>
		public UseZip64 UseZip64
		{
			get { return useZip64_; }
			set { useZip64_ = value; }
		}

		#endregion Properties

		#region Immediate updating

		//		TBD: Direct form of updating
		//
		//		public void Update(IEntryMatcher deleteMatcher)
		//		{
		//		}
		//
		//		public void Update(IScanner addScanner)
		//		{
		//		}

		#endregion Immediate updating

		#region Deferred Updating

		/// <summary>
		/// Begin updating this <see cref="ZipFile"/> archive.
		/// </summary>
		/// <param name="archiveStorage">The <see cref="IArchiveStorage">archive storage</see> for use during the update.</param>
		/// <param name="dataSource">The <see cref="IDynamicDataSource">data source</see> to utilise during updating.</param>
		/// <exception cref="ObjectDisposedException">ZipFile has been closed.</exception>
		/// <exception cref="ArgumentNullException">One of the arguments provided is null</exception>
		/// <exception cref="ObjectDisposedException">ZipFile has been closed.</exception>
		public void BeginUpdate(IArchiveStorage archiveStorage, IDynamicDataSource dataSource)
		{
			if (isDisposed_)
			{
				throw new ObjectDisposedException("ZipFile");
			}

			if (IsEmbeddedArchive)
			{
				throw new ZipException("Cannot update embedded/SFX archives");
			}

			archiveStorage_ = archiveStorage ?? throw new ArgumentNullException(nameof(archiveStorage));
			updateDataSource_ = dataSource ?? throw new ArgumentNullException(nameof(dataSource));

			// NOTE: the baseStream_ may not currently support writing or seeking.

			updateIndex_ = new Dictionary<string, int>();

			updates_ = new List<ZipUpdate>(entries_.Length);
			foreach (ZipEntry entry in entries_)
			{
				int index = updates_.Count;
				updates_.Add(new ZipUpdate(entry));
				updateIndex_.Add(entry.Name, index);
			}

			// We must sort by offset before using offset's calculated sizes
			updates_.Sort(new UpdateComparer());

			int idx = 0;
			foreach (ZipUpdate update in updates_)
			{
				//If last entry, there is no next entry offset to use
				if (idx == updates_.Count - 1)
					break;

				update.OffsetBasedSize = ((ZipUpdate)updates_[idx + 1]).Entry.Offset - update.Entry.Offset;
				idx++;
			}
			updateCount_ = updates_.Count;

			contentsEdited_ = false;
			commentEdited_ = false;
			newComment_ = null;
		}

		/// <summary>
		/// Begin updating to this <see cref="ZipFile"/> archive.
		/// </summary>
		/// <param name="archiveStorage">The storage to use during the update.</param>
		public void BeginUpdate(IArchiveStorage archiveStorage)
		{
			BeginUpdate(archiveStorage, new DynamicDiskDataSource());
		}

		/// <summary>
		/// Begin updating this <see cref="ZipFile"/> archive.
		/// </summary>
		/// <seealso cref="BeginUpdate(IArchiveStorage)"/>
		/// <seealso cref="CommitUpdate"></seealso>
		/// <seealso cref="AbortUpdate"></seealso>
		public void BeginUpdate()
		{
			if (Name == null)
			{
				BeginUpdate(new MemoryArchiveStorage(), new DynamicDiskDataSource());
			}
			else
			{
				BeginUpdate(new DiskArchiveStorage(this), new DynamicDiskDataSource());
			}
		}

		/// <summary>
		/// Commit current updates, updating this archive.
		/// </summary>
		/// <seealso cref="BeginUpdate()"></seealso>
		/// <seealso cref="AbortUpdate"></seealso>
		/// <exception cref="ObjectDisposedException">ZipFile has been closed.</exception>
		public void CommitUpdate()
		{
			if (isDisposed_)
			{
				throw new ObjectDisposedException("ZipFile");
			}

			CheckUpdating();

			try
			{
				updateIndex_.Clear();
				updateIndex_ = null;

				if (contentsEdited_)
				{
					RunUpdates();
				}
				else if (commentEdited_)
				{
					UpdateCommentOnly();
				}
				else
				{
					// Create an empty archive if none existed originally.
					if (entries_.Length == 0)
					{
						byte[] theComment = (newComment_ != null) ? newComment_.RawComment : ZipStrings.ConvertToArray(comment_);
						using (ZipHelperStream zhs = new ZipHelperStream(baseStream_))
						{
							zhs.WriteEndOfCentralDirectory(0, 0, 0, theComment);
						}
					}
				}
			}
			finally
			{
				PostUpdateCleanup();
			}
		}

		/// <summary>
		/// Abort updating leaving the archive unchanged.
		/// </summary>
		/// <seealso cref="BeginUpdate()"></seealso>
		/// <seealso cref="CommitUpdate"></seealso>
		public void AbortUpdate()
		{
			PostUpdateCleanup();
		}

		/// <summary>
		/// Set the file comment to be recorded when the current update is <see cref="CommitUpdate">commited</see>.
		/// </summary>
		/// <param name="comment">The comment to record.</param>
		/// <exception cref="ObjectDisposedException">ZipFile has been closed.</exception>
		public void SetComment(string comment)
		{
			if (isDisposed_)
			{
				throw new ObjectDisposedException("ZipFile");
			}

			CheckUpdating();

			newComment_ = new ZipString(comment);

			if (newComment_.RawLength > 0xffff)
			{
				newComment_ = null;
				throw new ZipException("Comment length exceeds maximum - 65535");
			}

			// We dont take account of the original and current comment appearing to be the same
			// as encoding may be different.
			commentEdited_ = true;
		}

		#endregion Deferred Updating

		#region Adding Entries

		private void AddUpdate(ZipUpdate update)
		{
			contentsEdited_ = true;

			int index = FindExistingUpdate(update.Entry.Name, isEntryName: true);

			if (index >= 0)
			{
				if (updates_[index] == null)
				{
					updateCount_ += 1;
				}

				// Direct replacement is faster than delete and add.
				updates_[index] = update;
			}
			else
			{
				index = updates_.Count;
				updates_.Add(update);
				updateCount_ += 1;
				updateIndex_.Add(update.Entry.Name, index);
			}
		}

		/// <summary>
		/// Add a new entry to the archive.
		/// </summary>
		/// <param name="fileName">The name of the file to add.</param>
		/// <param name="compressionMethod">The compression method to use.</param>
		/// <param name="useUnicodeText">Ensure Unicode text is used for name and comment for this entry.</param>
		/// <exception cref="ArgumentNullException">Argument supplied is null.</exception>
		/// <exception cref="ObjectDisposedException">ZipFile has been closed.</exception>
		/// <exception cref="NotImplementedException">Compression method is not supported for creating entries.</exception>
		public void Add(string fileName, CompressionMethod compressionMethod, bool useUnicodeText)
		{
			if (fileName == null)
			{
				throw new ArgumentNullException(nameof(fileName));
			}

			if (isDisposed_)
			{
				throw new ObjectDisposedException("ZipFile");
			}

			CheckSupportedCompressionMethod(compressionMethod);
			CheckUpdating();
			contentsEdited_ = true;

			ZipEntry entry = EntryFactory.MakeFileEntry(fileName);
			entry.IsUnicodeText = useUnicodeText;
			entry.CompressionMethod = compressionMethod;

			AddUpdate(new ZipUpdate(fileName, entry));
		}

		/// <summary>
		/// Add a new entry to the archive.
		/// </summary>
		/// <param name="fileName">The name of the file to add.</param>
		/// <param name="compressionMethod">The compression method to use.</param>
		/// <exception cref="ArgumentNullException">ZipFile has been closed.</exception>
		/// <exception cref="NotImplementedException">Compression method is not supported for creating entries.</exception>
		public void Add(string fileName, CompressionMethod compressionMethod)
		{
			if (fileName == null)
			{
				throw new ArgumentNullException(nameof(fileName));
			}

			CheckSupportedCompressionMethod(compressionMethod);
			CheckUpdating();
			contentsEdited_ = true;

			ZipEntry entry = EntryFactory.MakeFileEntry(fileName);
			entry.CompressionMethod = compressionMethod;
			AddUpdate(new ZipUpdate(fileName, entry));
		}

		/// <summary>
		/// Add a file to the archive.
		/// </summary>
		/// <param name="fileName">The name of the file to add.</param>
		/// <exception cref="ArgumentNullException">Argument supplied is null.</exception>
		public void Add(string fileName)
		{
			if (fileName == null)
			{
				throw new ArgumentNullException(nameof(fileName));
			}

			CheckUpdating();
			AddUpdate(new ZipUpdate(fileName, EntryFactory.MakeFileEntry(fileName)));
		}

		/// <summary>
		/// Add a file to the archive.
		/// </summary>
		/// <param name="fileName">The name of the file to add.</param>
		/// <param name="entryName">The name to use for the <see cref="ZipEntry"/> on the Zip file created.</param>
		/// <exception cref="ArgumentNullException">Argument supplied is null.</exception>
		public void Add(string fileName, string entryName)
		{
			if (fileName == null)
			{
				throw new ArgumentNullException(nameof(fileName));
			}

			if (entryName == null)
			{
				throw new ArgumentNullException(nameof(entryName));
			}

			CheckUpdating();
			AddUpdate(new ZipUpdate(fileName, EntryFactory.MakeFileEntry(fileName, entryName, true)));
		}

		/// <summary>
		/// Add a file entry with data.
		/// </summary>
		/// <param name="dataSource">The source of the data for this entry.</param>
		/// <param name="entryName">The name to give to the entry.</param>
		public void Add(IStaticDataSource dataSource, string entryName)
		{
			if (dataSource == null)
			{
				throw new ArgumentNullException(nameof(dataSource));
			}

			if (entryName == null)
			{
				throw new ArgumentNullException(nameof(entryName));
			}

			CheckUpdating();
			AddUpdate(new ZipUpdate(dataSource, EntryFactory.MakeFileEntry(entryName, false)));
		}

		/// <summary>
		/// Add a file entry with data.
		/// </summary>
		/// <param name="dataSource">The source of the data for this entry.</param>
		/// <param name="entryName">The name to give to the entry.</param>
		/// <param name="compressionMethod">The compression method to use.</param>
		/// <exception cref="NotImplementedException">Compression method is not supported for creating entries.</exception>
		public void Add(IStaticDataSource dataSource, string entryName, CompressionMethod compressionMethod)
		{
			if (dataSource == null)
			{
				throw new ArgumentNullException(nameof(dataSource));
			}

			if (entryName == null)
			{
				throw new ArgumentNullException(nameof(entryName));
			}

			CheckSupportedCompressionMethod(compressionMethod);
			CheckUpdating();

			ZipEntry entry = EntryFactory.MakeFileEntry(entryName, false);
			entry.CompressionMethod = compressionMethod;

			AddUpdate(new ZipUpdate(dataSource, entry));
		}

		/// <summary>
		/// Add a file entry with data.
		/// </summary>
		/// <param name="dataSource">The source of the data for this entry.</param>
		/// <param name="entryName">The name to give to the entry.</param>
		/// <param name="compressionMethod">The compression method to use.</param>
		/// <param name="useUnicodeText">Ensure Unicode text is used for name and comments for this entry.</param>
		/// <exception cref="NotImplementedException">Compression method is not supported for creating entries.</exception>
		public void Add(IStaticDataSource dataSource, string entryName, CompressionMethod compressionMethod, bool useUnicodeText)
		{
			if (dataSource == null)
			{
				throw new ArgumentNullException(nameof(dataSource));
			}

			if (entryName == null)
			{
				throw new ArgumentNullException(nameof(entryName));
			}

			CheckSupportedCompressionMethod(compressionMethod);
			CheckUpdating();

			ZipEntry entry = EntryFactory.MakeFileEntry(entryName, false);
			entry.IsUnicodeText = useUnicodeText;
			entry.CompressionMethod = compressionMethod;

			AddUpdate(new ZipUpdate(dataSource, entry));
		}

		/// <summary>
		/// Add a <see cref="ZipEntry"/> that contains no data.
		/// </summary>
		/// <param name="entry">The entry to add.</param>
		/// <remarks>This can be used to add directories, volume labels, or empty file entries.</remarks>
		public void Add(ZipEntry entry)
		{
			if (entry == null)
			{
				throw new ArgumentNullException(nameof(entry));
			}

			CheckUpdating();

			if ((entry.Size != 0) || (entry.CompressedSize != 0))
			{
				throw new ZipException("Entry cannot have any data");
			}

			AddUpdate(new ZipUpdate(UpdateCommand.Add, entry));
		}

		/// <summary>
		/// Add a <see cref="ZipEntry"/> with data.
		/// </summary>
		/// <param name="dataSource">The source of the data for this entry.</param>
		/// <param name="entry">The entry to add.</param>
		/// <remarks>This can be used to add file entries with a custom data source.</remarks>
		/// <exception cref="NotSupportedException">
		/// The encryption method specified in <paramref name="entry"/> is unsupported.
		/// </exception>
		/// <exception cref="NotImplementedException">Compression method is not supported for creating entries.</exception>
		public void Add(IStaticDataSource dataSource, ZipEntry entry)
		{
			if (entry == null)
			{
				throw new ArgumentNullException(nameof(entry));
			}

			if (dataSource == null)
			{
				throw new ArgumentNullException(nameof(dataSource));
			}

			// We don't currently support adding entries with AES encryption, so throw
			// up front instead of failing or falling back to ZipCrypto later on
			if (entry.AESKeySize > 0)
			{
				throw new NotSupportedException("Creation of AES encrypted entries is not supported");
			}

			CheckSupportedCompressionMethod(entry.CompressionMethod);
			CheckUpdating();

			AddUpdate(new ZipUpdate(dataSource, entry));
		}

		/// <summary>
		/// Add a directory entry to the archive.
		/// </summary>
		/// <param name="directoryName">The directory to add.</param>
		public void AddDirectory(string directoryName)
		{
			if (directoryName == null)
			{
				throw new ArgumentNullException(nameof(directoryName));
			}

			CheckUpdating();

			ZipEntry dirEntry = EntryFactory.MakeDirectoryEntry(directoryName);
			AddUpdate(new ZipUpdate(UpdateCommand.Add, dirEntry));
		}

		/// <summary>
		/// Check if the specified compression method is supported for adding a new entry.
		/// </summary>
		/// <param name="compressionMethod">The compression method for the new entry.</param>
		private static void CheckSupportedCompressionMethod(CompressionMethod compressionMethod)
		{
			if (compressionMethod != CompressionMethod.Deflated && compressionMethod != CompressionMethod.Stored && compressionMethod != CompressionMethod.BZip2)
			{
				throw new NotImplementedException("Compression method not supported");
			}
		}

		#endregion Adding Entries

		#region Modifying Entries

		/* Modify not yet ready for public consumption.
		   Direct modification of an entry should not overwrite original data before its read.
		   Safe mode is trivial in this sense.
				public void Modify(ZipEntry original, ZipEntry updated)
				{
					if ( original == null ) {
						throw new ArgumentNullException("original");
					}
					if ( updated == null ) {
						throw new ArgumentNullException("updated");
					}
					CheckUpdating();
					contentsEdited_ = true;
					updates_.Add(new ZipUpdate(original, updated));
				}
		*/

		#endregion Modifying Entries

		#region Deleting Entries

		/// <summary>
		/// Delete an entry by name
		/// </summary>
		/// <param name="fileName">The filename to delete</param>
		/// <returns>True if the entry was found and deleted; false otherwise.</returns>
		public bool Delete(string fileName)
		{
			if (fileName == null)
			{
				throw new ArgumentNullException(nameof(fileName));
			}

			CheckUpdating();

			bool result = false;
			int index = FindExistingUpdate(fileName);
			if ((index >= 0) && (updates_[index] != null))
			{
				result = true;
				contentsEdited_ = true;
				updates_[index] = null;
				updateCount_ -= 1;
			}
			else
			{
				throw new ZipException("Cannot find entry to delete");
			}
			return result;
		}

		/// <summary>
		/// Delete a <see cref="ZipEntry"/> from the archive.
		/// </summary>
		/// <param name="entry">The entry to delete.</param>
		public void Delete(ZipEntry entry)
		{
			if (entry == null)
			{
				throw new ArgumentNullException(nameof(entry));
			}

			CheckUpdating();

			int index = FindExistingUpdate(entry);
			if (index >= 0)
			{
				contentsEdited_ = true;
				updates_[index] = null;
				updateCount_ -= 1;
			}
			else
			{
				throw new ZipException("Cannot find entry to delete");
			}
		}

		#endregion Deleting Entries

		#region Update Support

		#region Writing Values/Headers

		private void WriteLEShort(int value)
		{
			baseStream_.WriteByte((byte)(value & 0xff));
			baseStream_.WriteByte((byte)((value >> 8) & 0xff));
		}

		/// <summary>
		/// Write an unsigned short in little endian byte order.
		/// </summary>
		private void WriteLEUshort(ushort value)
		{
			baseStream_.WriteByte((byte)(value & 0xff));
			baseStream_.WriteByte((byte)(value >> 8));
		}

		/// <summary>
		/// Write an int in little endian byte order.
		/// </summary>
		private void WriteLEInt(int value)
		{
			WriteLEShort(value & 0xffff);
			WriteLEShort(value >> 16);
		}

		/// <summary>
		/// Write an unsigned int in little endian byte order.
		/// </summary>
		private void WriteLEUint(uint value)
		{
			WriteLEUshort((ushort)(value & 0xffff));
			WriteLEUshort((ushort)(value >> 16));
		}

		/// <summary>
		/// Write a long in little endian byte order.
		/// </summary>
		private void WriteLeLong(long value)
		{
			WriteLEInt((int)(value & 0xffffffff));
			WriteLEInt((int)(value >> 32));
		}

		private void WriteLEUlong(ulong value)
		{
			WriteLEUint((uint)(value & 0xffffffff));
			WriteLEUint((uint)(value >> 32));
		}

		private void WriteLocalEntryHeader(ZipUpdate update)
		{
			ZipEntry entry = update.OutEntry;

			// TODO: Local offset will require adjusting for multi-disk zip files.
			entry.Offset = baseStream_.Position;

			// TODO: Need to clear any entry flags that dont make sense or throw an exception here.
			if (update.Command != UpdateCommand.Copy)
			{
				if (entry.CompressionMethod == CompressionMethod.Deflated)
				{
					if (entry.Size == 0)
					{
						// No need to compress - no data.
						entry.CompressedSize = entry.Size;
						entry.Crc = 0;
						entry.CompressionMethod = CompressionMethod.Stored;
					}
				}
				else if (entry.CompressionMethod == CompressionMethod.Stored)
				{
					entry.Flags &= ~(int)GeneralBitFlags.Descriptor;
				}

				if (HaveKeys)
				{
					entry.IsCrypted = true;
					if (entry.Crc < 0)
					{
						entry.Flags |= (int)GeneralBitFlags.Descriptor;
					}
				}
				else
				{
					entry.IsCrypted = false;
				}

				switch (useZip64_)
				{
					case UseZip64.Dynamic:
						if (entry.Size < 0)
						{
							entry.ForceZip64();
						}
						break;

					case UseZip64.On:
						entry.ForceZip64();
						break;

					case UseZip64.Off:
						// Do nothing.  The entry itself may be using Zip64 independently.
						break;
				}
			}

			// Write the local file header
			WriteLEInt(ZipConstants.LocalHeaderSignature);

			WriteLEShort(entry.Version);
			WriteLEShort(entry.Flags);

			WriteLEShort((byte)entry.CompressionMethodForHeader);
			WriteLEInt((int)entry.DosTime);

			if (!entry.HasCrc)
			{
				// Note patch address for updating CRC later.
				update.CrcPatchOffset = baseStream_.Position;
				WriteLEInt((int)0);
			}
			else
			{
				WriteLEInt(unchecked((int)entry.Crc));
			}

			if (entry.LocalHeaderRequiresZip64)
			{
				WriteLEInt(-1);
				WriteLEInt(-1);
			}
			else
			{
				if ((entry.CompressedSize < 0) || (entry.Size < 0))
				{
					update.SizePatchOffset = baseStream_.Position;
				}

				WriteLEInt((int)entry.CompressedSize);
				WriteLEInt((int)entry.Size);
			}

			byte[] name = ZipStrings.ConvertToArray(entry.Flags, entry.Name);

			if (name.Length > 0xFFFF)
			{
				throw new ZipException("Entry name too long.");
			}

			var ed = new ZipExtraData(entry.ExtraData);

			if (entry.LocalHeaderRequiresZip64)
			{
				ed.StartNewEntry();

				// Local entry header always includes size and compressed size.
				// NOTE the order of these fields is reversed when compared to the normal headers!
				ed.AddLeLong(entry.Size);
				ed.AddLeLong(entry.CompressedSize);
				ed.AddNewEntry(1);
			}
			else
			{
				ed.Delete(1);
			}

			entry.ExtraData = ed.GetEntryData();

			WriteLEShort(name.Length);
			WriteLEShort(entry.ExtraData.Length);

			if (name.Length > 0)
			{
				baseStream_.Write(name, 0, name.Length);
			}

			if (entry.LocalHeaderRequiresZip64)
			{
				if (!ed.Find(1))
				{
					throw new ZipException("Internal error cannot find extra data");
				}

				update.SizePatchOffset = baseStream_.Position + ed.CurrentReadIndex;
			}

			if (entry.ExtraData.Length > 0)
			{
				baseStream_.Write(entry.ExtraData, 0, entry.ExtraData.Length);
			}
		}

		private int WriteCentralDirectoryHeader(ZipEntry entry)
		{
			if (entry.CompressedSize < 0)
			{
				throw new ZipException("Attempt to write central directory entry with unknown csize");
			}

			if (entry.Size < 0)
			{
				throw new ZipException("Attempt to write central directory entry with unknown size");
			}

			if (entry.Crc < 0)
			{
				throw new ZipException("Attempt to write central directory entry with unknown crc");
			}

			// Write the central file header
			WriteLEInt(ZipConstants.CentralHeaderSignature);

			// Version made by
			WriteLEShort((entry.HostSystem << 8) | entry.VersionMadeBy);

			// Version required to extract
			WriteLEShort(entry.Version);

			WriteLEShort(entry.Flags);

			unchecked
			{
				WriteLEShort((byte)entry.CompressionMethodForHeader);
				WriteLEInt((int)entry.DosTime);
				WriteLEInt((int)entry.Crc);
			}

			bool useExtraCompressedSize = false; //Do we want to store the compressed size in the extra data?
			if ((entry.IsZip64Forced()) || (entry.CompressedSize >= 0xffffffff))
			{
				useExtraCompressedSize = true;
				WriteLEInt(-1);
			}
			else
			{
				WriteLEInt((int)(entry.CompressedSize & 0xffffffff));
			}

			bool useExtraUncompressedSize = false; //Do we want to store the uncompressed size in the extra data?
			if ((entry.IsZip64Forced()) || (entry.Size >= 0xffffffff))
			{
				useExtraUncompressedSize = true;
				WriteLEInt(-1);
			}
			else
			{
				WriteLEInt((int)entry.Size);
			}

			byte[] name = ZipStrings.ConvertToArray(entry.Flags, entry.Name);

			if (name.Length > 0xFFFF)
			{
				throw new ZipException("Entry name is too long.");
			}

			WriteLEShort(name.Length);

			// Central header extra data is different to local header version so regenerate.
			var ed = new ZipExtraData(entry.ExtraData);

			if (entry.CentralHeaderRequiresZip64)
			{
				ed.StartNewEntry();

				if (useExtraUncompressedSize)
				{
					ed.AddLeLong(entry.Size);
				}

				if (useExtraCompressedSize)
				{
					ed.AddLeLong(entry.CompressedSize);
				}

				if (entry.Offset >= 0xffffffff)
				{
					ed.AddLeLong(entry.Offset);
				}

				// Number of disk on which this file starts isnt supported and is never written here.
				ed.AddNewEntry(1);
			}
			else
			{
				// Should have already be done when local header was added.
				ed.Delete(1);
			}

			byte[] centralExtraData = ed.GetEntryData();

			WriteLEShort(centralExtraData.Length);
			WriteLEShort(entry.Comment != null ? entry.Comment.Length : 0);

			WriteLEShort(0);    // disk number
			WriteLEShort(0);    // internal file attributes

			// External file attributes...
			if (entry.ExternalFileAttributes != -1)
			{
				WriteLEInt(entry.ExternalFileAttributes);
			}
			else
			{
				if (entry.IsDirectory)
				{
					WriteLEUint(16);
				}
				else
				{
					WriteLEUint(0);
				}
			}

			if (entry.Offset >= 0xffffffff)
			{
				WriteLEUint(0xffffffff);
			}
			else
			{
				WriteLEUint((uint)(int)entry.Offset);
			}

			if (name.Length > 0)
			{
				baseStream_.Write(name, 0, name.Length);
			}

			if (centralExtraData.Length > 0)
			{
				baseStream_.Write(centralExtraData, 0, centralExtraData.Length);
			}

			byte[] rawComment = (entry.Comment != null) ? Encoding.ASCII.GetBytes(entry.Comment) : Empty.Array<byte>();

			if (rawComment.Length > 0)
			{
				baseStream_.Write(rawComment, 0, rawComment.Length);
			}

			return ZipConstants.CentralHeaderBaseSize + name.Length + centralExtraData.Length + rawComment.Length;
		}

		#endregion Writing Values/Headers

		private void PostUpdateCleanup()
		{
			updateDataSource_ = null;
			updates_ = null;
			updateIndex_ = null;

			if (archiveStorage_ != null)
			{
				archiveStorage_.Dispose();
				archiveStorage_ = null;
			}
		}

		private string GetTransformedFileName(string name)
		{
			INameTransform transform = NameTransform;
			return (transform != null) ?
				transform.TransformFile(name) :
				name;
		}

		private string GetTransformedDirectoryName(string name)
		{
			INameTransform transform = NameTransform;
			return (transform != null) ?
				transform.TransformDirectory(name) :
				name;
		}

		/// <summary>
		/// Get a raw memory buffer.
		/// </summary>
		/// <returns>Returns a raw memory buffer.</returns>
		private byte[] GetBuffer()
		{
			if (copyBuffer_ == null)
			{
				copyBuffer_ = new byte[bufferSize_];
			}
			return copyBuffer_;
		}

		private void CopyDescriptorBytes(ZipUpdate update, Stream dest, Stream source)
		{
			// Don't include the signature size to allow copy without seeking
			var bytesToCopy = GetDescriptorSize(update, false);

			// Don't touch the source stream if no descriptor is present
			if (bytesToCopy == 0) return;

			var buffer = GetBuffer();

			// Copy the first 4 bytes of the descriptor
			source.Read(buffer, 0, sizeof(int));
			dest.Write(buffer, 0, sizeof(int));

			if (BitConverter.ToUInt32(buffer, 0) != ZipConstants.DataDescriptorSignature)
			{
				// The initial bytes wasn't the descriptor, reduce the pending byte count
				bytesToCopy -= buffer.Length;
			}

			while (bytesToCopy > 0)
			{
				int readSize = Math.Min(buffer.Length, bytesToCopy);

				int bytesRead = source.Read(buffer, 0, readSize);
				if (bytesRead > 0)
				{
					dest.Write(buffer, 0, bytesRead);
					bytesToCopy -= bytesRead;
				}
				else
				{
					throw new ZipException("Unxpected end of stream");
				}
			}
		}

		private void CopyBytes(ZipUpdate update, Stream destination, Stream source,
			long bytesToCopy, bool updateCrc)
		{
			if (destination == source)
			{
				throw new InvalidOperationException("Destination and source are the same");
			}

			// NOTE: Compressed size is updated elsewhere.
			var crc = new Crc32();
			byte[] buffer = GetBuffer();

			long targetBytes = bytesToCopy;
			long totalBytesRead = 0;

			int bytesRead;
			do
			{
				int readSize = buffer.Length;

				if (bytesToCopy < readSize)
				{
					readSize = (int)bytesToCopy;
				}

				bytesRead = source.Read(buffer, 0, readSize);
				if (bytesRead > 0)
				{
					if (updateCrc)
					{
						crc.Update(new ArraySegment<byte>(buffer, 0, bytesRead));
					}
					destination.Write(buffer, 0, bytesRead);
					bytesToCopy -= bytesRead;
					totalBytesRead += bytesRead;
				}
			}
			while ((bytesRead > 0) && (bytesToCopy > 0));

			if (totalBytesRead != targetBytes)
			{
				throw new ZipException(string.Format("Failed to copy bytes expected {0} read {1}", targetBytes, totalBytesRead));
			}

			if (updateCrc)
			{
				update.OutEntry.Crc = crc.Value;
			}
		}

		/// <summary>
		/// Get the size of the source descriptor for a <see cref="ZipUpdate"/>.
		/// </summary>
		/// <param name="update">The update to get the size for.</param>
		/// <param name="includingSignature">Whether to include the signature size</param>
		/// <returns>The descriptor size, zero if there isn't one.</returns>
		private static int GetDescriptorSize(ZipUpdate update, bool includingSignature)
		{
			if (!((GeneralBitFlags)update.Entry.Flags).HasFlag(GeneralBitFlags.Descriptor)) 
				return 0;
			
			var descriptorWithSignature = update.Entry.LocalHeaderRequiresZip64 
				? ZipConstants.Zip64DataDescriptorSize 
				: ZipConstants.DataDescriptorSize;

			return includingSignature 
				? descriptorWithSignature 
				: descriptorWithSignature - sizeof(int);
		}

		private void CopyDescriptorBytesDirect(ZipUpdate update, Stream stream, ref long destinationPosition, long sourcePosition)
		{
			var buffer = GetBuffer(); ;

			stream.Position = sourcePosition;
			stream.Read(buffer, 0, sizeof(int));
			var sourceHasSignature = BitConverter.ToUInt32(buffer, 0) == ZipConstants.DataDescriptorSignature;

			var bytesToCopy = GetDescriptorSize(update, sourceHasSignature);

			while (bytesToCopy > 0)
			{
				stream.Position = sourcePosition;

				var bytesRead = stream.Read(buffer, 0, bytesToCopy);
				if (bytesRead > 0)
				{
					stream.Position = destinationPosition;
					stream.Write(buffer, 0, bytesRead);
					bytesToCopy -= bytesRead;
					destinationPosition += bytesRead;
					sourcePosition += bytesRead;
				}
				else
				{
					throw new ZipException("Unexpected end of stream");
				}
			}
		}

		private void CopyEntryDataDirect(ZipUpdate update, Stream stream, bool updateCrc, ref long destinationPosition, ref long sourcePosition)
		{
			long bytesToCopy = update.Entry.CompressedSize;

			// NOTE: Compressed size is updated elsewhere.
			var crc = new Crc32();
			byte[] buffer = GetBuffer();

			long targetBytes = bytesToCopy;
			long totalBytesRead = 0;

			int bytesRead;
			do
			{
				int readSize = buffer.Length;

				if (bytesToCopy < readSize)
				{
					readSize = (int)bytesToCopy;
				}

				stream.Position = sourcePosition;
				bytesRead = stream.Read(buffer, 0, readSize);
				if (bytesRead > 0)
				{
					if (updateCrc)
					{
						crc.Update(new ArraySegment<byte>(buffer, 0, bytesRead));
					}
					stream.Position = destinationPosition;
					stream.Write(buffer, 0, bytesRead);

					destinationPosition += bytesRead;
					sourcePosition += bytesRead;
					bytesToCopy -= bytesRead;
					totalBytesRead += bytesRead;
				}
			}
			while ((bytesRead > 0) && (bytesToCopy > 0));

			if (totalBytesRead != targetBytes)
			{
				throw new ZipException(string.Format("Failed to copy bytes expected {0} read {1}", targetBytes, totalBytesRead));
			}

			if (updateCrc)
			{
				update.OutEntry.Crc = crc.Value;
			}
		}

		private int FindExistingUpdate(ZipEntry entry)
		{
			int result = -1;
			if (updateIndex_.ContainsKey(entry.Name))
			{
				result = (int)updateIndex_[entry.Name];
			}
			/*
						// This is slow like the coming of the next ice age but takes less storage and may be useful
						// for CF?
						for (int index = 0; index < updates_.Count; ++index)
						{
							ZipUpdate zu = ( ZipUpdate )updates_[index];
							if ( (zu.Entry.ZipFileIndex == entry.ZipFileIndex) &&
								(string.Compare(convertedName, zu.Entry.Name, true, CultureInfo.InvariantCulture) == 0) ) {
								result = index;
								break;
							}
						}
			 */
			return result;
		}

		private int FindExistingUpdate(string fileName, bool isEntryName = false)
		{
			int result = -1;

			string convertedName = !isEntryName ? GetTransformedFileName(fileName) : fileName;

			if (updateIndex_.ContainsKey(convertedName))
			{
				result = (int)updateIndex_[convertedName];
			}

			/*
						// This is slow like the coming of the next ice age but takes less storage and may be useful
						// for CF?
						for ( int index = 0; index < updates_.Count; ++index ) {
							if ( string.Compare(convertedName, (( ZipUpdate )updates_[index]).Entry.Name,
								true, CultureInfo.InvariantCulture) == 0 ) {
								result = index;
								break;
							}
						}
			 */

			return result;
		}

		/// <summary>
		/// Get an output stream for the specified <see cref="ZipEntry"/>
		/// </summary>
		/// <param name="entry">The entry to get an output stream for.</param>
		/// <returns>The output stream obtained for the entry.</returns>
		private Stream GetOutputStream(ZipEntry entry)
		{
			Stream result = baseStream_;

			if (entry.IsCrypted == true)
			{
				result = CreateAndInitEncryptionStream(result, entry);
			}

			switch (entry.CompressionMethod)
			{
				case CompressionMethod.Stored:
					if (!entry.IsCrypted)
					{
						// If there is an encryption stream in use, that can be returned directly
						// otherwise, wrap the base stream in an UncompressedStream instead of returning it directly
						result = new UncompressedStream(result);
					}
					break;

				case CompressionMethod.Deflated:
					var dos = new DeflaterOutputStream(result, new Deflater(9, true))
					{
						// If there is an encryption stream in use, then we want that to be disposed when the deflator stream is disposed
						// If not, then we don't want it to dispose the base stream
						IsStreamOwner = entry.IsCrypted
					};
					result = dos;
					break;

				case CompressionMethod.BZip2:
					var bzos = new BZip2.BZip2OutputStream(result)
					{
						// If there is an encryption stream in use, then we want that to be disposed when the BZip2OutputStream stream is disposed
						// If not, then we don't want it to dispose the base stream
						IsStreamOwner = entry.IsCrypted
					};
					result = bzos;
					break;

				default:
					throw new ZipException("Unknown compression method " + entry.CompressionMethod);
			}
			return result;
		}

		private void AddEntry(ZipFile workFile, ZipUpdate update)
		{
			Stream source = null;

			if (update.Entry.IsFile)
			{
				source = update.GetSource();

				if (source == null)
				{
					source = updateDataSource_.GetSource(update.Entry, update.Filename);
				}
			}

			var useCrc = update.Entry.AESKeySize == 0;

			if (source != null)
			{
				using (source)
				{
					long sourceStreamLength = source.Length;
					if (update.OutEntry.Size < 0)
					{
						update.OutEntry.Size = sourceStreamLength;
					}
					else
					{
						// Check for errant entries.
						if (update.OutEntry.Size != sourceStreamLength)
						{
							throw new ZipException("Entry size/stream size mismatch");
						}
					}

					workFile.WriteLocalEntryHeader(update);

					long dataStart = workFile.baseStream_.Position;

					using (Stream output = workFile.GetOutputStream(update.OutEntry))
					{
						CopyBytes(update, output, source, sourceStreamLength, useCrc);
					}

					long dataEnd = workFile.baseStream_.Position;
					update.OutEntry.CompressedSize = dataEnd - dataStart;

					if ((update.OutEntry.Flags & (int)GeneralBitFlags.Descriptor) == (int)GeneralBitFlags.Descriptor)
					{
						var helper = new ZipHelperStream(workFile.baseStream_);
						helper.WriteDataDescriptor(update.OutEntry);
					}
				}
			}
			else
			{
				workFile.WriteLocalEntryHeader(update);
				update.OutEntry.CompressedSize = 0;
			}
		}

		private void ModifyEntry(ZipFile workFile, ZipUpdate update)
		{
			workFile.WriteLocalEntryHeader(update);
			long dataStart = workFile.baseStream_.Position;

			// TODO: This is slow if the changes don't effect the data!!
			if (update.Entry.IsFile && (update.Filename != null))
			{
				using (Stream output = workFile.GetOutputStream(update.OutEntry))
				{
					using (Stream source = this.GetInputStream(update.Entry))
					{
						CopyBytes(update, output, source, source.Length, true);
					}
				}
			}

			long dataEnd = workFile.baseStream_.Position;
			update.Entry.CompressedSize = dataEnd - dataStart;
		}

		private void CopyEntryDirect(ZipFile workFile, ZipUpdate update, ref long destinationPosition)
		{
			bool skipOver = false || update.Entry.Offset == destinationPosition;

			if (!skipOver)
			{
				baseStream_.Position = destinationPosition;
				workFile.WriteLocalEntryHeader(update);
				destinationPosition = baseStream_.Position;
			}

			long sourcePosition = 0;

			const int NameLengthOffset = 26;

			// TODO: Add base for SFX friendly handling
			long entryDataOffset = update.Entry.Offset + NameLengthOffset;

			baseStream_.Seek(entryDataOffset, SeekOrigin.Begin);

			// Clumsy way of handling retrieving the original name and extra data length for now.
			// TODO: Stop re-reading name and data length in CopyEntryDirect.
			
			uint nameLength = ReadLEUshort();
			uint extraLength = ReadLEUshort();

			sourcePosition = baseStream_.Position + nameLength + extraLength;

			if (skipOver)
			{
				if (update.OffsetBasedSize != -1)
				{
					destinationPosition += update.OffsetBasedSize;
				}
				else
				{
					// Skip entry header
					destinationPosition += (sourcePosition - entryDataOffset) + NameLengthOffset;

					// Skip entry compressed data
					destinationPosition += update.Entry.CompressedSize;

					// Seek to end of entry to check for descriptor signature
					baseStream_.Seek(destinationPosition, SeekOrigin.Begin);

					var descriptorHasSignature = ReadLEUint() == ZipConstants.DataDescriptorSignature;

					// Skip descriptor and it's signature (if present)
					destinationPosition += GetDescriptorSize(update, descriptorHasSignature);
				}
			}
			else
			{
				if (update.Entry.CompressedSize > 0)
				{
					CopyEntryDataDirect(update, baseStream_, false, ref destinationPosition, ref sourcePosition);
				}
				CopyDescriptorBytesDirect(update, baseStream_, ref destinationPosition, sourcePosition);
			}
		}

		private void CopyEntry(ZipFile workFile, ZipUpdate update)
		{
			workFile.WriteLocalEntryHeader(update);

			if (update.Entry.CompressedSize > 0)
			{
				const int NameLengthOffset = 26;

				long entryDataOffset = update.Entry.Offset + NameLengthOffset;

				// TODO: This wont work for SFX files!
				baseStream_.Seek(entryDataOffset, SeekOrigin.Begin);

				uint nameLength = ReadLEUshort();
				uint extraLength = ReadLEUshort();

				baseStream_.Seek(nameLength + extraLength, SeekOrigin.Current);

				CopyBytes(update, workFile.baseStream_, baseStream_, update.Entry.CompressedSize, false);
			}
			CopyDescriptorBytes(update, workFile.baseStream_, baseStream_);
		}

		private void Reopen(Stream source)
		{
			isNewArchive_ = false;
			baseStream_ = source ?? throw new ZipException("Failed to reopen archive - no source");
			ReadEntries();
		}

		private void Reopen()
		{
			if (Name == null)
			{
				throw new InvalidOperationException("Name is not known cannot Reopen");
			}

			Reopen(File.Open(Name, FileMode.Open, FileAccess.Read, FileShare.Read));
		}

		private void UpdateCommentOnly()
		{
			long baseLength = baseStream_.Length;

			ZipHelperStream updateFile = null;

			if (archiveStorage_.UpdateMode == FileUpdateMode.Safe)
			{
				Stream copyStream = archiveStorage_.MakeTemporaryCopy(baseStream_);
				updateFile = new ZipHelperStream(copyStream)
				{
					IsStreamOwner = true
				};

				baseStream_.Dispose();
				baseStream_ = null;
			}
			else
			{
				if (archiveStorage_.UpdateMode == FileUpdateMode.Direct)
				{
					// TODO: archiveStorage wasnt originally intended for this use.
					// Need to revisit this to tidy up handling as archive storage currently doesnt
					// handle the original stream well.
					// The problem is when using an existing zip archive with an in memory archive storage.
					// The open stream wont support writing but the memory storage should open the same file not an in memory one.

					// Need to tidy up the archive storage interface and contract basically.
					baseStream_ = archiveStorage_.OpenForDirectUpdate(baseStream_);
					updateFile = new ZipHelperStream(baseStream_);
				}
				else
				{
					baseStream_.Dispose();
					baseStream_ = null;
					updateFile = new ZipHelperStream(Name);
				}
			}

			using (updateFile)
			{
				long locatedCentralDirOffset =
					updateFile.LocateBlockWithSignature(ZipConstants.EndOfCentralDirectorySignature,
														baseLength, ZipConstants.EndOfCentralRecordBaseSize, 0xffff);
				if (locatedCentralDirOffset < 0)
				{
					throw new ZipException("Cannot find central directory");
				}

				const int CentralHeaderCommentSizeOffset = 16;
				updateFile.Position += CentralHeaderCommentSizeOffset;

				byte[] rawComment = newComment_.RawComment;

				updateFile.WriteLEShort(rawComment.Length);
				updateFile.Write(rawComment, 0, rawComment.Length);
				updateFile.SetLength(updateFile.Position);
			}

			if (archiveStorage_.UpdateMode == FileUpdateMode.Safe)
			{
				Reopen(archiveStorage_.ConvertTemporaryToFinal());
			}
			else
			{
				ReadEntries();
			}
		}

		/// <summary>
		/// Class used to sort updates.
		/// </summary>
		private class UpdateComparer : IComparer<ZipUpdate>
		{
			/// <summary>
			/// Compares two objects and returns a value indicating whether one is
			/// less than, equal to or greater than the other.
			/// </summary>
			/// <param name="x">First object to compare</param>
			/// <param name="y">Second object to compare.</param>
			/// <returns>Compare result.</returns>
			public int Compare(ZipUpdate x, ZipUpdate y)
			{
				int result;

				if (x == null)
				{
					if (y == null)
					{
						result = 0;
					}
					else
					{
						result = -1;
					}
				}
				else if (y == null)
				{
					result = 1;
				}
				else
				{
					int xCmdValue = ((x.Command == UpdateCommand.Copy) || (x.Command == UpdateCommand.Modify)) ? 0 : 1;
					int yCmdValue = ((y.Command == UpdateCommand.Copy) || (y.Command == UpdateCommand.Modify)) ? 0 : 1;

					result = xCmdValue - yCmdValue;
					if (result == 0)
					{
						long offsetDiff = x.Entry.Offset - y.Entry.Offset;
						if (offsetDiff < 0)
						{
							result = -1;
						}
						else if (offsetDiff == 0)
						{
							result = 0;
						}
						else
						{
							result = 1;
						}
					}
				}
				return result;
			}
		}

		private void RunUpdates()
		{
			long sizeEntries = 0;
			long endOfStream = 0;
			bool directUpdate = false;
			long destinationPosition = 0; // NOT SFX friendly

			ZipFile workFile;

			if (IsNewArchive)
			{
				workFile = this;
				workFile.baseStream_.Position = 0;
				directUpdate = true;
			}
			else if (archiveStorage_.UpdateMode == FileUpdateMode.Direct)
			{
				workFile = this;
				workFile.baseStream_.Position = 0;
				directUpdate = true;

				// Sort the updates by offset within copies/modifies, then adds.
				// This ensures that data required by copies will not be overwritten.
				updates_.Sort(new UpdateComparer());
			}
			else
			{
				workFile = ZipFile.Create(archiveStorage_.GetTemporaryOutput());
				workFile.UseZip64 = UseZip64;

				if (key != null)
				{
					workFile.key = (byte[])key.Clone();
				}
			}

			try
			{
				foreach (ZipUpdate update in updates_)
				{
					if (update != null)
					{
						switch (update.Command)
						{
							case UpdateCommand.Copy:
								if (directUpdate)
								{
									CopyEntryDirect(workFile, update, ref destinationPosition);
								}
								else
								{
									CopyEntry(workFile, update);
								}
								break;

							case UpdateCommand.Modify:
								// TODO: Direct modifying of an entry will take some legwork.
								ModifyEntry(workFile, update);
								break;

							case UpdateCommand.Add:
								if (!IsNewArchive && directUpdate)
								{
									workFile.baseStream_.Position = destinationPosition;
								}

								AddEntry(workFile, update);

								if (directUpdate)
								{
									destinationPosition = workFile.baseStream_.Position;
								}
								break;
						}
					}
				}

				if (!IsNewArchive && directUpdate)
				{
					workFile.baseStream_.Position = destinationPosition;
				}

				long centralDirOffset = workFile.baseStream_.Position;

				foreach (ZipUpdate update in updates_)
				{
					if (update != null)
					{
						sizeEntries += workFile.WriteCentralDirectoryHeader(update.OutEntry);
					}
				}

				byte[] theComment = (newComment_ != null) ? newComment_.RawComment : ZipStrings.ConvertToArray(comment_);
				using (ZipHelperStream zhs = new ZipHelperStream(workFile.baseStream_))
				{
					zhs.WriteEndOfCentralDirectory(updateCount_, sizeEntries, centralDirOffset, theComment);
				}

				endOfStream = workFile.baseStream_.Position;

				// And now patch entries...
				foreach (ZipUpdate update in updates_)
				{
					if (update != null)
					{
						// If the size of the entry is zero leave the crc as 0 as well.
						// The calculated crc will be all bits on...
						if ((update.CrcPatchOffset > 0) && (update.OutEntry.CompressedSize > 0))
						{
							workFile.baseStream_.Position = update.CrcPatchOffset;
							workFile.WriteLEInt((int)update.OutEntry.Crc);
						}

						if (update.SizePatchOffset > 0)
						{
							workFile.baseStream_.Position = update.SizePatchOffset;
							if (update.OutEntry.LocalHeaderRequiresZip64)
							{
								workFile.WriteLeLong(update.OutEntry.Size);
								workFile.WriteLeLong(update.OutEntry.CompressedSize);
							}
							else
							{
								workFile.WriteLEInt((int)update.OutEntry.CompressedSize);
								workFile.WriteLEInt((int)update.OutEntry.Size);
							}
						}
					}
				}
			}
			catch
			{
				workFile.Close();
				if (!directUpdate && (workFile.Name != null))
				{
					File.Delete(workFile.Name);
				}
				throw;
			}

			if (directUpdate)
			{
				workFile.baseStream_.SetLength(endOfStream);
				workFile.baseStream_.Flush();
				isNewArchive_ = false;
				ReadEntries();
			}
			else
			{
				baseStream_.Dispose();
				Reopen(archiveStorage_.ConvertTemporaryToFinal());
			}
		}

		private void CheckUpdating()
		{
			if (updates_ == null)
			{
				throw new InvalidOperationException("BeginUpdate has not been called");
			}
		}

		#endregion Update Support

		#region ZipUpdate class

		/// <summary>
		/// Represents a pending update to a Zip file.
		/// </summary>
		private class ZipUpdate
		{
			#region Constructors

			public ZipUpdate(string fileName, ZipEntry entry)
			{
				command_ = UpdateCommand.Add;
				entry_ = entry;
				filename_ = fileName;
			}

			[Obsolete]
			public ZipUpdate(string fileName, string entryName, CompressionMethod compressionMethod)
			{
				command_ = UpdateCommand.Add;
				entry_ = new ZipEntry(entryName)
				{
					CompressionMethod = compressionMethod
				};
				filename_ = fileName;
			}

			[Obsolete]
			public ZipUpdate(string fileName, string entryName)
				: this(fileName, entryName, CompressionMethod.Deflated)
			{
				// Do nothing.
			}

			[Obsolete]
			public ZipUpdate(IStaticDataSource dataSource, string entryName, CompressionMethod compressionMethod)
			{
				command_ = UpdateCommand.Add;
				entry_ = new ZipEntry(entryName)
				{
					CompressionMethod = compressionMethod
				};
				dataSource_ = dataSource;
			}

			public ZipUpdate(IStaticDataSource dataSource, ZipEntry entry)
			{
				command_ = UpdateCommand.Add;
				entry_ = entry;
				dataSource_ = dataSource;
			}

			public ZipUpdate(ZipEntry original, ZipEntry updated)
			{
				throw new ZipException("Modify not currently supported");
				/*
					command_ = UpdateCommand.Modify;
					entry_ = ( ZipEntry )original.Clone();
					outEntry_ = ( ZipEntry )updated.Clone();
				*/
			}

			public ZipUpdate(UpdateCommand command, ZipEntry entry)
			{
				command_ = command;
				entry_ = (ZipEntry)entry.Clone();
			}

			/// <summary>
			/// Copy an existing entry.
			/// </summary>
			/// <param name="entry">The existing entry to copy.</param>
			public ZipUpdate(ZipEntry entry)
				: this(UpdateCommand.Copy, entry)
			{
				// Do nothing.
			}

			#endregion Constructors

			/// <summary>
			/// Get the <see cref="ZipEntry"/> for this update.
			/// </summary>
			/// <remarks>This is the source or original entry.</remarks>
			public ZipEntry Entry
			{
				get { return entry_; }
			}

			/// <summary>
			/// Get the <see cref="ZipEntry"/> that will be written to the updated/new file.
			/// </summary>
			public ZipEntry OutEntry
			{
				get
				{
					if (outEntry_ == null)
					{
						outEntry_ = (ZipEntry)entry_.Clone();
					}

					return outEntry_;
				}
			}

			/// <summary>
			/// Get the command for this update.
			/// </summary>
			public UpdateCommand Command
			{
				get { return command_; }
			}

			/// <summary>
			/// Get the filename if any for this update.  Null if none exists.
			/// </summary>
			public string Filename
			{
				get { return filename_; }
			}

			/// <summary>
			/// Get/set the location of the size patch for this update.
			/// </summary>
			public long SizePatchOffset
			{
				get { return sizePatchOffset_; }
				set { sizePatchOffset_ = value; }
			}

			/// <summary>
			/// Get /set the location of the crc patch for this update.
			/// </summary>
			public long CrcPatchOffset
			{
				get { return crcPatchOffset_; }
				set { crcPatchOffset_ = value; }
			}

			/// <summary>
			/// Get/set the size calculated by offset.
			/// Specifically, the difference between this and next entry's starting offset.
			/// </summary>
			public long OffsetBasedSize
			{
				get { return _offsetBasedSize; }
				set { _offsetBasedSize = value; }
			}

			public Stream GetSource()
			{
				Stream result = null;
				if (dataSource_ != null)
				{
					result = dataSource_.GetSource();
				}

				return result;
			}

			#region Instance Fields

			private ZipEntry entry_;
			private ZipEntry outEntry_;
			private readonly UpdateCommand command_;
			private IStaticDataSource dataSource_;
			private readonly string filename_;
			private long sizePatchOffset_ = -1;
			private long crcPatchOffset_ = -1;
			private long _offsetBasedSize = -1;

			#endregion Instance Fields
		}

		#endregion ZipUpdate class

		#endregion Updating

		#region Disposing

		#region IDisposable Members

		void IDisposable.Dispose()
		{
			Close();
		}

		#endregion IDisposable Members

		private void DisposeInternal(bool disposing)
		{
			if (!isDisposed_)
			{
				isDisposed_ = true;
				entries_ = Empty.Array<ZipEntry>();

				if (IsStreamOwner && (baseStream_ != null))
				{
					lock (baseStream_)
					{
						baseStream_.Dispose();
					}
				}

				PostUpdateCleanup();
			}
		}

		/// <summary>
		/// Releases the unmanaged resources used by the this instance and optionally releases the managed resources.
		/// </summary>
		/// <param name="disposing">true to release both managed and unmanaged resources;
		/// false to release only unmanaged resources.</param>
		protected virtual void Dispose(bool disposing)
		{
			DisposeInternal(disposing);
		}

		#endregion Disposing

		#region Internal routines

		#region Reading

		/// <summary>
		/// Read an unsigned short in little endian byte order.
		/// </summary>
		/// <returns>Returns the value read.</returns>
		/// <exception cref="EndOfStreamException">
		/// The stream ends prematurely
		/// </exception>
		private ushort ReadLEUshort()
		{
			int data1 = baseStream_.ReadByte();

			if (data1 < 0)
			{
				throw new EndOfStreamException("End of stream");
			}

			int data2 = baseStream_.ReadByte();

			if (data2 < 0)
			{
				throw new EndOfStreamException("End of stream");
			}

			return unchecked((ushort)((ushort)data1 | (ushort)(data2 << 8)));
		}

		/// <summary>
		/// Read a uint in little endian byte order.
		/// </summary>
		/// <returns>Returns the value read.</returns>
		/// <exception cref="IOException">
		/// An i/o error occurs.
		/// </exception>
		/// <exception cref="System.IO.EndOfStreamException">
		/// The file ends prematurely
		/// </exception>
		private uint ReadLEUint()
		{
			return (uint)(ReadLEUshort() | (ReadLEUshort() << 16));
		}

		private ulong ReadLEUlong()
		{
			return ReadLEUint() | ((ulong)ReadLEUint() << 32);
		}

		#endregion Reading

		// NOTE this returns the offset of the first byte after the signature.
		private long LocateBlockWithSignature(int signature, long endLocation, int minimumBlockSize, int maximumVariableData)
		{
			using (ZipHelperStream les = new ZipHelperStream(baseStream_))
			{
				return les.LocateBlockWithSignature(signature, endLocation, minimumBlockSize, maximumVariableData);
			}
		}

		/// <summary>
		/// Search for and read the central directory of a zip file filling the entries array.
		/// </summary>
		/// <exception cref="System.IO.IOException">
		/// An i/o error occurs.
		/// </exception>
		/// <exception cref="ICSharpCode.SharpZipLib.Zip.ZipException">
		/// The central directory is malformed or cannot be found
		/// </exception>
		private void ReadEntries()
		{
			// Search for the End Of Central Directory.  When a zip comment is
			// present the directory will start earlier
			//
			// The search is limited to 64K which is the maximum size of a trailing comment field to aid speed.
			// This should be compatible with both SFX and ZIP files but has only been tested for Zip files
			// If a SFX file has the Zip data attached as a resource and there are other resources occurring later then
			// this could be invalid.
			// Could also speed this up by reading memory in larger blocks.

			if (baseStream_.CanSeek == false)
			{
				throw new ZipException("ZipFile stream must be seekable");
			}

			long locatedEndOfCentralDir = LocateBlockWithSignature(ZipConstants.EndOfCentralDirectorySignature,
				baseStream_.Length, ZipConstants.EndOfCentralRecordBaseSize, 0xffff);

			if (locatedEndOfCentralDir < 0)
			{
				throw new ZipException("Cannot find central directory");
			}

			// Read end of central directory record
			ushort thisDiskNumber = ReadLEUshort();
			ushort startCentralDirDisk = ReadLEUshort();
			ulong entriesForThisDisk = ReadLEUshort();
			ulong entriesForWholeCentralDir = ReadLEUshort();
			ulong centralDirSize = ReadLEUint();
			long offsetOfCentralDir = ReadLEUint();
			uint commentSize = ReadLEUshort();

			if (commentSize > 0)
			{
				byte[] comment = new byte[commentSize];

				StreamUtils.ReadFully(baseStream_, comment);
				comment_ = ZipStrings.ConvertToString(comment);
			}
			else
			{
				comment_ = string.Empty;
			}

			bool isZip64 = false;
			bool requireZip64 = false;

			// Check if zip64 header information is required.
			if ((thisDiskNumber == 0xffff) ||
				(startCentralDirDisk == 0xffff) ||
				(entriesForThisDisk == 0xffff) ||
				(entriesForWholeCentralDir == 0xffff) ||
				(centralDirSize == 0xffffffff) ||
				(offsetOfCentralDir == 0xffffffff))
			{
				requireZip64 = true;
			}

			// #357 - always check for the existance of the Zip64 central directory.
			// #403 - Take account of the fixed size of the locator when searching.
			//    Subtract from locatedEndOfCentralDir so that the endLocation is the location of EndOfCentralDirectorySignature,
			//    rather than the data following the signature.
			long locatedZip64EndOfCentralDirLocator = LocateBlockWithSignature(
				ZipConstants.Zip64CentralDirLocatorSignature,
				locatedEndOfCentralDir - 4,
				ZipConstants.Zip64EndOfCentralDirectoryLocatorSize,
				0);

			if (locatedZip64EndOfCentralDirLocator < 0)
			{
				if (requireZip64)
				{
					// This is only an error in cases where the Zip64 directory is required.
					throw new ZipException("Cannot find Zip64 locator");
				}
			}
			else
			{
				isZip64 = true;

				// number of the disk with the start of the zip64 end of central directory 4 bytes
				// relative offset of the zip64 end of central directory record 8 bytes
				// total number of disks 4 bytes
				ReadLEUint(); // startDisk64 is not currently used
				ulong offset64 = ReadLEUlong();
				uint totalDisks = ReadLEUint();

				baseStream_.Position = (long)offset64;
				long sig64 = ReadLEUint();

				if (sig64 != ZipConstants.Zip64CentralFileHeaderSignature)
				{
					throw new ZipException(string.Format("Invalid Zip64 Central directory signature at {0:X}", offset64));
				}

				// NOTE: Record size = SizeOfFixedFields + SizeOfVariableData - 12.
				ulong recordSize = ReadLEUlong();
				int versionMadeBy = ReadLEUshort();
				int versionToExtract = ReadLEUshort();
				uint thisDisk = ReadLEUint();
				uint centralDirDisk = ReadLEUint();
				entriesForThisDisk = ReadLEUlong();
				entriesForWholeCentralDir = ReadLEUlong();
				centralDirSize = ReadLEUlong();
				offsetOfCentralDir = (long)ReadLEUlong();

				// NOTE: zip64 extensible data sector (variable size) is ignored.
			}

			entries_ = new ZipEntry[entriesForThisDisk];

			// SFX/embedded support, find the offset of the first entry vis the start of the stream
			// This applies to Zip files that are appended to the end of an SFX stub.
			// Or are appended as a resource to an executable.
			// Zip files created by some archivers have the offsets altered to reflect the true offsets
			// and so dont require any adjustment here...
			// TODO: Difficulty with Zip64 and SFX offset handling needs resolution - maths?
			if (!isZip64 && (offsetOfCentralDir < locatedEndOfCentralDir - (4 + (long)centralDirSize)))
			{
				offsetOfFirstEntry = locatedEndOfCentralDir - (4 + (long)centralDirSize + offsetOfCentralDir);
				if (offsetOfFirstEntry <= 0)
				{
					throw new ZipException("Invalid embedded zip archive");
				}
			}

			baseStream_.Seek(offsetOfFirstEntry + offsetOfCentralDir, SeekOrigin.Begin);

			for (ulong i = 0; i < entriesForThisDisk; i++)
			{
				if (ReadLEUint() != ZipConstants.CentralHeaderSignature)
				{
					throw new ZipException("Wrong Central Directory signature");
				}

				int versionMadeBy = ReadLEUshort();
				int versionToExtract = ReadLEUshort();
				int bitFlags = ReadLEUshort();
				int method = ReadLEUshort();
				uint dostime = ReadLEUint();
				uint crc = ReadLEUint();
				var csize = (long)ReadLEUint();
				var size = (long)ReadLEUint();
				int nameLen = ReadLEUshort();
				int extraLen = ReadLEUshort();
				int commentLen = ReadLEUshort();

				int diskStartNo = ReadLEUshort();  // Not currently used
				int internalAttributes = ReadLEUshort();  // Not currently used

				uint externalAttributes = ReadLEUint();
				long offset = ReadLEUint();

				byte[] buffer = new byte[Math.Max(nameLen, commentLen)];

				StreamUtils.ReadFully(baseStream_, buffer, 0, nameLen);
				string name = ZipStrings.ConvertToStringExt(bitFlags, buffer, nameLen);

				var entry = new ZipEntry(name, versionToExtract, versionMadeBy, (CompressionMethod)method)
				{
					Crc = crc & 0xffffffffL,
					Size = size & 0xffffffffL,
					CompressedSize = csize & 0xffffffffL,
					Flags = bitFlags,
					DosTime = dostime,
					ZipFileIndex = (long)i,
					Offset = offset,
					ExternalFileAttributes = (int)externalAttributes
				};

				if ((bitFlags & 8) == 0)
				{
					entry.CryptoCheckValue = (byte)(crc >> 24);
				}
				else
				{
					entry.CryptoCheckValue = (byte)((dostime >> 8) & 0xff);
				}

				if (extraLen > 0)
				{
					byte[] extra = new byte[extraLen];
					StreamUtils.ReadFully(baseStream_, extra);
					entry.ExtraData = extra;
				}

				entry.ProcessExtraData(false);

				if (commentLen > 0)
				{
					StreamUtils.ReadFully(baseStream_, buffer, 0, commentLen);
					entry.Comment = ZipStrings.ConvertToStringExt(bitFlags, buffer, commentLen);
				}

				entries_[i] = entry;
			}
		}

		/// <summary>
		/// Locate the data for a given entry.
		/// </summary>
		/// <returns>
		/// The start offset of the data.
		/// </returns>
		/// <exception cref="System.IO.EndOfStreamException">
		/// The stream ends prematurely
		/// </exception>
		/// <exception cref="ICSharpCode.SharpZipLib.Zip.ZipException">
		/// The local header signature is invalid, the entry and central header file name lengths are different
		/// or the local and entry compression methods dont match
		/// </exception>
		private long LocateEntry(ZipEntry entry)
		{
			return TestLocalHeader(entry, HeaderTest.Extract);
		}

		private Stream CreateAndInitDecryptionStream(Stream baseStream, ZipEntry entry)
		{
			CryptoStream result = null;

			if (entry.CompressionMethodForHeader == CompressionMethod.WinZipAES)
			{
				if (entry.Version >= ZipConstants.VERSION_AES)
				{
					// Issue #471 - accept an empty string as a password, but reject null.
					OnKeysRequired(entry.Name);
					if (rawPassword_ == null)
					{
						throw new ZipException("No password available for AES encrypted stream");
					}
					int saltLen = entry.AESSaltLen;
					byte[] saltBytes = new byte[saltLen];
					int saltIn = StreamUtils.ReadRequestedBytes(baseStream, saltBytes, 0, saltLen);
					if (saltIn != saltLen)
						throw new ZipException("AES Salt expected " + saltLen + " got " + saltIn);
					//
					byte[] pwdVerifyRead = new byte[2];
					StreamUtils.ReadFully(baseStream, pwdVerifyRead);
					int blockSize = entry.AESKeySize / 8;   // bits to bytes

					var decryptor = new ZipAESTransform(rawPassword_, saltBytes, blockSize, false);
					byte[] pwdVerifyCalc = decryptor.PwdVerifier;
					if (pwdVerifyCalc[0] != pwdVerifyRead[0] || pwdVerifyCalc[1] != pwdVerifyRead[1])
						throw new ZipException("Invalid password for AES");
					result = new ZipAESStream(baseStream, decryptor, CryptoStreamMode.Read);
				}
				else
				{
					throw new ZipException("Decryption method not supported");
				}
			}
			else
			{
				if ((entry.Version < ZipConstants.VersionStrongEncryption)
					|| (entry.Flags & (int)GeneralBitFlags.StrongEncryption) == 0)
				{
					var classicManaged = new PkzipClassicManaged();

					OnKeysRequired(entry.Name);
					if (HaveKeys == false)
					{
						throw new ZipException("No password available for encrypted stream");
					}

					result = new CryptoStream(baseStream, classicManaged.CreateDecryptor(key, null), CryptoStreamMode.Read);
					CheckClassicPassword(result, entry);
				}
				else
				{
					// We don't support PKWare strong encryption
					throw new ZipException("Decryption method not supported");
				}
			}

			return result;
		}

		private Stream CreateAndInitEncryptionStream(Stream baseStream, ZipEntry entry)
		{
			CryptoStream result = null;
			if ((entry.Version < ZipConstants.VersionStrongEncryption)
				|| (entry.Flags & (int)GeneralBitFlags.StrongEncryption) == 0)
			{
				var classicManaged = new PkzipClassicManaged();

				OnKeysRequired(entry.Name);
				if (HaveKeys == false)
				{
					throw new ZipException("No password available for encrypted stream");
				}

				// Closing a CryptoStream will close the base stream as well so wrap it in an UncompressedStream
				// which doesnt do this.
				result = new CryptoStream(new UncompressedStream(baseStream),
					classicManaged.CreateEncryptor(key, null), CryptoStreamMode.Write);

				if ((entry.Crc < 0) || (entry.Flags & 8) != 0)
				{
					WriteEncryptionHeader(result, entry.DosTime << 16);
				}
				else
				{
					WriteEncryptionHeader(result, entry.Crc);
				}
			}
			return result;
		}

		private static void CheckClassicPassword(CryptoStream classicCryptoStream, ZipEntry entry)
		{
			byte[] cryptbuffer = new byte[ZipConstants.CryptoHeaderSize];
			StreamUtils.ReadFully(classicCryptoStream, cryptbuffer);
			if (cryptbuffer[ZipConstants.CryptoHeaderSize - 1] != entry.CryptoCheckValue)
			{
				throw new ZipException("Invalid password");
			}
		}

		private static void WriteEncryptionHeader(Stream stream, long crcValue)
		{
			byte[] cryptBuffer = new byte[ZipConstants.CryptoHeaderSize];
			using (var rng = new RNGCryptoServiceProvider())
			{
				rng.GetBytes(cryptBuffer);
			}
			cryptBuffer[11] = (byte)(crcValue >> 24);
			stream.Write(cryptBuffer, 0, cryptBuffer.Length);
		}

		#endregion Internal routines

		#region Instance Fields

		private bool isDisposed_;
		private string name_;
		private string comment_;
		private string rawPassword_;
		private Stream baseStream_;
		private bool isStreamOwner;
		private long offsetOfFirstEntry;
		private ZipEntry[] entries_;
		private byte[] key;
		private bool isNewArchive_;

		// Default is dynamic which is not backwards compatible and can cause problems
		// with XP's built in compression which cant read Zip64 archives.
		// However it does avoid the situation were a large file is added and cannot be completed correctly.
		// Hint: Set always ZipEntry size before they are added to an archive and this setting isnt needed.
		private UseZip64 useZip64_ = UseZip64.Dynamic;

		#region Zip Update Instance Fields

		private List<ZipUpdate> updates_;
		private long updateCount_; // Count is managed manually as updates_ can contain nulls!
		private Dictionary<string, int> updateIndex_;
		private IArchiveStorage archiveStorage_;
		private IDynamicDataSource updateDataSource_;
		private bool contentsEdited_;
		private int bufferSize_ = DefaultBufferSize;
		private byte[] copyBuffer_;
		private ZipString newComment_;
		private bool commentEdited_;
		private IEntryFactory updateEntryFactory_ = new ZipEntryFactory();

		#endregion Zip Update Instance Fields

		#endregion Instance Fields

		#region Support Classes

		/// <summary>
		/// Represents a string from a <see cref="ZipFile"/> which is stored as an array of bytes.
		/// </summary>
		private class ZipString
		{
			#region Constructors

			/// <summary>
			/// Initialise a <see cref="ZipString"/> with a string.
			/// </summary>
			/// <param name="comment">The textual string form.</param>
			public ZipString(string comment)
			{
				comment_ = comment;
				isSourceString_ = true;
			}

			/// <summary>
			/// Initialise a <see cref="ZipString"/> using a string in its binary 'raw' form.
			/// </summary>
			/// <param name="rawString"></param>
			public ZipString(byte[] rawString)
			{
				rawComment_ = rawString;
			}

			#endregion Constructors

			/// <summary>
			/// Get a value indicating the original source of data for this instance.
			/// True if the source was a string; false if the source was binary data.
			/// </summary>
			public bool IsSourceString
			{
				get { return isSourceString_; }
			}

			/// <summary>
			/// Get the length of the comment when represented as raw bytes.
			/// </summary>
			public int RawLength
			{
				get
				{
					MakeBytesAvailable();
					return rawComment_.Length;
				}
			}

			/// <summary>
			/// Get the comment in its 'raw' form as plain bytes.
			/// </summary>
			public byte[] RawComment
			{
				get
				{
					MakeBytesAvailable();
					return (byte[])rawComment_.Clone();
				}
			}

			/// <summary>
			/// Reset the comment to its initial state.
			/// </summary>
			public void Reset()
			{
				if (isSourceString_)
				{
					rawComment_ = null;
				}
				else
				{
					comment_ = null;
				}
			}

			private void MakeTextAvailable()
			{
				if (comment_ == null)
				{
					comment_ = ZipStrings.ConvertToString(rawComment_);
				}
			}

			private void MakeBytesAvailable()
			{
				if (rawComment_ == null)
				{
					rawComment_ = ZipStrings.ConvertToArray(comment_);
				}
			}

			/// <summary>
			/// Implicit conversion of comment to a string.
			/// </summary>
			/// <param name="zipString">The <see cref="ZipString"/> to convert to a string.</param>
			/// <returns>The textual equivalent for the input value.</returns>
			static public implicit operator string(ZipString zipString)
			{
				zipString.MakeTextAvailable();
				return zipString.comment_;
			}

			#region Instance Fields

			private string comment_;
			private byte[] rawComment_;
			private readonly bool isSourceString_;

			#endregion Instance Fields
		}

		/// <summary>
		/// An <see cref="IEnumerator">enumerator</see> for <see cref="ZipEntry">Zip entries</see>
		/// </summary>
		private class ZipEntryEnumerator : IEnumerator
		{
			#region Constructors

			public ZipEntryEnumerator(ZipEntry[] entries)
			{
				array = entries;
			}

			#endregion Constructors

			#region IEnumerator Members

			public object Current
			{
				get
				{
					return array[index];
				}
			}

			public void Reset()
			{
				index = -1;
			}

			public bool MoveNext()
			{
				return (++index < array.Length);
			}

			#endregion IEnumerator Members

			#region Instance Fields

			private ZipEntry[] array;
			private int index = -1;

			#endregion Instance Fields
		}

		/// <summary>
		/// An <see cref="UncompressedStream"/> is a stream that you can write uncompressed data
		/// to and flush, but cannot read, seek or do anything else to.
		/// </summary>
		private class UncompressedStream : Stream
		{
			#region Constructors

			public UncompressedStream(Stream baseStream)
			{
				baseStream_ = baseStream;
			}

			#endregion Constructors

			/// <summary>
			/// Gets a value indicating whether the current stream supports reading.
			/// </summary>
			public override bool CanRead
			{
				get
				{
					return false;
				}
			}

			/// <summary>
			/// Write any buffered data to underlying storage.
			/// </summary>
			public override void Flush()
			{
				baseStream_.Flush();
			}

			/// <summary>
			/// Gets a value indicating whether the current stream supports writing.
			/// </summary>
			public override bool CanWrite
			{
				get
				{
					return baseStream_.CanWrite;
				}
			}

			/// <summary>
			/// Gets a value indicating whether the current stream supports seeking.
			/// </summary>
			public override bool CanSeek
			{
				get
				{
					return false;
				}
			}

			/// <summary>
			/// Get the length in bytes of the stream.
			/// </summary>
			public override long Length
			{
				get
				{
					return 0;
				}
			}

			/// <summary>
			/// Gets or sets the position within the current stream.
			/// </summary>
			public override long Position
			{
				get
				{
					return baseStream_.Position;
				}
				set
				{
					throw new NotImplementedException();
				}
			}

			/// <summary>
			/// Reads a sequence of bytes from the current stream and advances the position within the stream by the number of bytes read.
			/// </summary>
			/// <param name="buffer">An array of bytes. When this method returns, the buffer contains the specified byte array with the values between offset and (offset + count - 1) replaced by the bytes read from the current source.</param>
			/// <param name="offset">The zero-based byte offset in buffer at which to begin storing the data read from the current stream.</param>
			/// <param name="count">The maximum number of bytes to be read from the current stream.</param>
			/// <returns>
			/// The total number of bytes read into the buffer. This can be less than the number of bytes requested if that many bytes are not currently available, or zero (0) if the end of the stream has been reached.
			/// </returns>
			/// <exception cref="System.ArgumentException">The sum of offset and count is larger than the buffer length. </exception>
			/// <exception cref="System.ObjectDisposedException">Methods were called after the stream was closed. </exception>
			/// <exception cref="System.NotSupportedException">The stream does not support reading. </exception>
			/// <exception cref="System.ArgumentNullException">buffer is null. </exception>
			/// <exception cref="System.IO.IOException">An I/O error occurs. </exception>
			/// <exception cref="System.ArgumentOutOfRangeException">offset or count is negative. </exception>
			public override int Read(byte[] buffer, int offset, int count)
			{
				return 0;
			}

			/// <summary>
			/// Sets the position within the current stream.
			/// </summary>
			/// <param name="offset">A byte offset relative to the origin parameter.</param>
			/// <param name="origin">A value of type <see cref="System.IO.SeekOrigin"></see> indicating the reference point used to obtain the new position.</param>
			/// <returns>
			/// The new position within the current stream.
			/// </returns>
			/// <exception cref="System.IO.IOException">An I/O error occurs. </exception>
			/// <exception cref="System.NotSupportedException">The stream does not support seeking, such as if the stream is constructed from a pipe or console output. </exception>
			/// <exception cref="System.ObjectDisposedException">Methods were called after the stream was closed. </exception>
			public override long Seek(long offset, SeekOrigin origin)
			{
				return 0;
			}

			/// <summary>
			/// Sets the length of the current stream.
			/// </summary>
			/// <param name="value">The desired length of the current stream in bytes.</param>
			/// <exception cref="System.NotSupportedException">The stream does not support both writing and seeking, such as if the stream is constructed from a pipe or console output. </exception>
			/// <exception cref="System.IO.IOException">An I/O error occurs. </exception>
			/// <exception cref="System.ObjectDisposedException">Methods were called after the stream was closed. </exception>
			public override void SetLength(long value)
			{
			}

			/// <summary>
			/// Writes a sequence of bytes to the current stream and advances the current position within this stream by the number of bytes written.
			/// </summary>
			/// <param name="buffer">An array of bytes. This method copies count bytes from buffer to the current stream.</param>
			/// <param name="offset">The zero-based byte offset in buffer at which to begin copying bytes to the current stream.</param>
			/// <param name="count">The number of bytes to be written to the current stream.</param>
			/// <exception cref="System.IO.IOException">An I/O error occurs. </exception>
			/// <exception cref="System.NotSupportedException">The stream does not support writing. </exception>
			/// <exception cref="System.ObjectDisposedException">Methods were called after the stream was closed. </exception>
			/// <exception cref="System.ArgumentNullException">buffer is null. </exception>
			/// <exception cref="System.ArgumentException">The sum of offset and count is greater than the buffer length. </exception>
			/// <exception cref="System.ArgumentOutOfRangeException">offset or count is negative. </exception>
			public override void Write(byte[] buffer, int offset, int count)
			{
				baseStream_.Write(buffer, offset, count);
			}

			private readonly

			#region Instance Fields

			Stream baseStream_;

			#endregion Instance Fields
		}

		/// <summary>
		/// A <see cref="PartialInputStream"/> is an <see cref="InflaterInputStream"/>
		/// whose data is only a part or subsection of a file.
		/// </summary>
		private class PartialInputStream : Stream
		{
			#region Constructors

			/// <summary>
			/// Initialise a new instance of the <see cref="PartialInputStream"/> class.
			/// </summary>
			/// <param name="zipFile">The <see cref="ZipFile"/> containing the underlying stream to use for IO.</param>
			/// <param name="start">The start of the partial data.</param>
			/// <param name="length">The length of the partial data.</param>
			public PartialInputStream(ZipFile zipFile, long start, long length)
			{
				start_ = start;
				length_ = length;

				// Although this is the only time the zipfile is used
				// keeping a reference here prevents premature closure of
				// this zip file and thus the baseStream_.

				// Code like this will cause apparently random failures depending
				// on the size of the files and when garbage is collected.
				//
				// ZipFile z = new ZipFile (stream);
				// Stream reader = z.GetInputStream(0);
				// uses reader here....
				zipFile_ = zipFile;
				baseStream_ = zipFile_.baseStream_;
				readPos_ = start;
				end_ = start + length;
			}

			#endregion Constructors

			/// <summary>
			/// Read a byte from this stream.
			/// </summary>
			/// <returns>Returns the byte read or -1 on end of stream.</returns>
			public override int ReadByte()
			{
				if (readPos_ >= end_)
				{
					// -1 is the correct value at end of stream.
					return -1;
				}

				lock (baseStream_)
				{
					baseStream_.Seek(readPos_++, SeekOrigin.Begin);
					return baseStream_.ReadByte();
				}
			}

			/// <summary>
			/// Reads a sequence of bytes from the current stream and advances the position within the stream by the number of bytes read.
			/// </summary>
			/// <param name="buffer">An array of bytes. When this method returns, the buffer contains the specified byte array with the values between offset and (offset + count - 1) replaced by the bytes read from the current source.</param>
			/// <param name="offset">The zero-based byte offset in buffer at which to begin storing the data read from the current stream.</param>
			/// <param name="count">The maximum number of bytes to be read from the current stream.</param>
			/// <returns>
			/// The total number of bytes read into the buffer. This can be less than the number of bytes requested if that many bytes are not currently available, or zero (0) if the end of the stream has been reached.
			/// </returns>
			/// <exception cref="System.ArgumentException">The sum of offset and count is larger than the buffer length. </exception>
			/// <exception cref="System.ObjectDisposedException">Methods were called after the stream was closed. </exception>
			/// <exception cref="System.NotSupportedException">The stream does not support reading. </exception>
			/// <exception cref="System.ArgumentNullException">buffer is null. </exception>
			/// <exception cref="System.IO.IOException">An I/O error occurs. </exception>
			/// <exception cref="System.ArgumentOutOfRangeException">offset or count is negative. </exception>
			public override int Read(byte[] buffer, int offset, int count)
			{
				lock (baseStream_)
				{
					if (count > end_ - readPos_)
					{
						count = (int)(end_ - readPos_);
						if (count == 0)
						{
							return 0;
						}
					}
					// Protect against Stream implementations that throw away their buffer on every Seek
					// (for example, Mono FileStream)
					if (baseStream_.Position != readPos_)
					{
						baseStream_.Seek(readPos_, SeekOrigin.Begin);
					}
					int readCount = baseStream_.Read(buffer, offset, count);
					if (readCount > 0)
					{
						readPos_ += readCount;
					}
					return readCount;
				}
			}

			/// <summary>
			/// Writes a sequence of bytes to the current stream and advances the current position within this stream by the number of bytes written.
			/// </summary>
			/// <param name="buffer">An array of bytes. This method copies count bytes from buffer to the current stream.</param>
			/// <param name="offset">The zero-based byte offset in buffer at which to begin copying bytes to the current stream.</param>
			/// <param name="count">The number of bytes to be written to the current stream.</param>
			/// <exception cref="System.IO.IOException">An I/O error occurs. </exception>
			/// <exception cref="System.NotSupportedException">The stream does not support writing. </exception>
			/// <exception cref="System.ObjectDisposedException">Methods were called after the stream was closed. </exception>
			/// <exception cref="System.ArgumentNullException">buffer is null. </exception>
			/// <exception cref="System.ArgumentException">The sum of offset and count is greater than the buffer length. </exception>
			/// <exception cref="System.ArgumentOutOfRangeException">offset or count is negative. </exception>
			public override void Write(byte[] buffer, int offset, int count)
			{
				throw new NotSupportedException();
			}

			/// <summary>
			/// When overridden in a derived class, sets the length of the current stream.
			/// </summary>
			/// <param name="value">The desired length of the current stream in bytes.</param>
			/// <exception cref="System.NotSupportedException">The stream does not support both writing and seeking, such as if the stream is constructed from a pipe or console output. </exception>
			/// <exception cref="System.IO.IOException">An I/O error occurs. </exception>
			/// <exception cref="System.ObjectDisposedException">Methods were called after the stream was closed. </exception>
			public override void SetLength(long value)
			{
				throw new NotSupportedException();
			}

			/// <summary>
			/// When overridden in a derived class, sets the position within the current stream.
			/// </summary>
			/// <param name="offset">A byte offset relative to the origin parameter.</param>
			/// <param name="origin">A value of type <see cref="System.IO.SeekOrigin"></see> indicating the reference point used to obtain the new position.</param>
			/// <returns>
			/// The new position within the current stream.
			/// </returns>
			/// <exception cref="System.IO.IOException">An I/O error occurs. </exception>
			/// <exception cref="System.NotSupportedException">The stream does not support seeking, such as if the stream is constructed from a pipe or console output. </exception>
			/// <exception cref="System.ObjectDisposedException">Methods were called after the stream was closed. </exception>
			public override long Seek(long offset, SeekOrigin origin)
			{
				long newPos = readPos_;

				switch (origin)
				{
					case SeekOrigin.Begin:
						newPos = start_ + offset;
						break;

					case SeekOrigin.Current:
						newPos = readPos_ + offset;
						break;

					case SeekOrigin.End:
						newPos = end_ + offset;
						break;
				}

				if (newPos < start_)
				{
					throw new ArgumentException("Negative position is invalid");
				}

				if (newPos > end_)
				{
					throw new IOException("Cannot seek past end");
				}
				readPos_ = newPos;
				return readPos_;
			}

			/// <summary>
			/// Clears all buffers for this stream and causes any buffered data to be written to the underlying device.
			/// </summary>
			/// <exception cref="System.IO.IOException">An I/O error occurs. </exception>
			public override void Flush()
			{
				// Nothing to do.
			}

			/// <summary>
			/// Gets or sets the position within the current stream.
			/// </summary>
			/// <value></value>
			/// <returns>The current position within the stream.</returns>
			/// <exception cref="System.IO.IOException">An I/O error occurs. </exception>
			/// <exception cref="System.NotSupportedException">The stream does not support seeking. </exception>
			/// <exception cref="System.ObjectDisposedException">Methods were called after the stream was closed. </exception>
			public override long Position
			{
				get { return readPos_ - start_; }
				set
				{
					long newPos = start_ + value;

					if (newPos < start_)
					{
						throw new ArgumentException("Negative position is invalid");
					}

					if (newPos > end_)
					{
						throw new InvalidOperationException("Cannot seek past end");
					}
					readPos_ = newPos;
				}
			}

			/// <summary>
			/// Gets the length in bytes of the stream.
			/// </summary>
			/// <value></value>
			/// <returns>A long value representing the length of the stream in bytes.</returns>
			/// <exception cref="System.NotSupportedException">A class derived from Stream does not support seeking. </exception>
			/// <exception cref="System.ObjectDisposedException">Methods were called after the stream was closed. </exception>
			public override long Length
			{
				get { return length_; }
			}

			/// <summary>
			/// Gets a value indicating whether the current stream supports writing.
			/// </summary>
			/// <value>false</value>
			/// <returns>true if the stream supports writing; otherwise, false.</returns>
			public override bool CanWrite
			{
				get { return false; }
			}

			/// <summary>
			/// Gets a value indicating whether the current stream supports seeking.
			/// </summary>
			/// <value>true</value>
			/// <returns>true if the stream supports seeking; otherwise, false.</returns>
			public override bool CanSeek
			{
				get { return true; }
			}

			/// <summary>
			/// Gets a value indicating whether the current stream supports reading.
			/// </summary>
			/// <value>true.</value>
			/// <returns>true if the stream supports reading; otherwise, false.</returns>
			public override bool CanRead
			{
				get { return true; }
			}

			/// <summary>
			/// Gets a value that determines whether the current stream can time out.
			/// </summary>
			/// <value></value>
			/// <returns>A value that determines whether the current stream can time out.</returns>
			public override bool CanTimeout
			{
				get { return baseStream_.CanTimeout; }
			}

			#region Instance Fields

			private ZipFile zipFile_;
			private Stream baseStream_;
			private readonly long start_;
			private readonly long length_;
			private long readPos_;
			private readonly long end_;

			#endregion Instance Fields
		}

		#endregion Support Classes
	}

	#endregion ZipFile Class

	#region DataSources

	/// <summary>
	/// Provides a static way to obtain a source of data for an entry.
	/// </summary>
	public interface IStaticDataSource
	{
		/// <summary>
		/// Get a source of data by creating a new stream.
		/// </summary>
		/// <returns>Returns a <see cref="Stream"/> to use for compression input.</returns>
		/// <remarks>Ideally a new stream is created and opened to achieve this, to avoid locking problems.</remarks>
		Stream GetSource();
	}

	/// <summary>
	/// Represents a source of data that can dynamically provide
	/// multiple <see cref="Stream">data sources</see> based on the parameters passed.
	/// </summary>
	public interface IDynamicDataSource
	{
		/// <summary>
		/// Get a data source.
		/// </summary>
		/// <param name="entry">The <see cref="ZipEntry"/> to get a source for.</param>
		/// <param name="name">The name for data if known.</param>
		/// <returns>Returns a <see cref="Stream"/> to use for compression input.</returns>
		/// <remarks>Ideally a new stream is created and opened to achieve this, to avoid locking problems.</remarks>
		Stream GetSource(ZipEntry entry, string name);
	}

	/// <summary>
	/// Default implementation of a <see cref="IStaticDataSource"/> for use with files stored on disk.
	/// </summary>
	public class StaticDiskDataSource : IStaticDataSource
	{
		/// <summary>
		/// Initialise a new instance of <see cref="StaticDiskDataSource"/>
		/// </summary>
		/// <param name="fileName">The name of the file to obtain data from.</param>
		public StaticDiskDataSource(string fileName)
		{
			fileName_ = fileName;
		}

		#region IDataSource Members

		/// <summary>
		/// Get a <see cref="Stream"/> providing data.
		/// </summary>
		/// <returns>Returns a <see cref="Stream"/> providing data.</returns>
		public Stream GetSource()
		{
			return File.Open(fileName_, FileMode.Open, FileAccess.Read, FileShare.Read);
		}

		private readonly

		#endregion IDataSource Members

		#region Instance Fields

		string fileName_;

		#endregion Instance Fields
	}

	/// <summary>
	/// Default implementation of <see cref="IDynamicDataSource"/> for files stored on disk.
	/// </summary>
	public class DynamicDiskDataSource : IDynamicDataSource
	{
		#region IDataSource Members

		/// <summary>
		/// Get a <see cref="Stream"/> providing data for an entry.
		/// </summary>
		/// <param name="entry">The entry to provide data for.</param>
		/// <param name="name">The file name for data if known.</param>
		/// <returns>Returns a stream providing data; or null if not available</returns>
		public Stream GetSource(ZipEntry entry, string name)
		{
			Stream result = null;

			if (name != null)
			{
				result = File.Open(name, FileMode.Open, FileAccess.Read, FileShare.Read);
			}

			return result;
		}

		#endregion IDataSource Members
	}

	#endregion DataSources

	#region Archive Storage

	/// <summary>
	/// Defines facilities for data storage when updating Zip Archives.
	/// </summary>
	public interface IArchiveStorage
	{
		/// <summary>
		/// Get the <see cref="FileUpdateMode"/> to apply during updates.
		/// </summary>
		FileUpdateMode UpdateMode { get; }

		/// <summary>
		/// Get an empty <see cref="Stream"/> that can be used for temporary output.
		/// </summary>
		/// <returns>Returns a temporary output <see cref="Stream"/></returns>
		/// <seealso cref="ConvertTemporaryToFinal"></seealso>
		Stream GetTemporaryOutput();

		/// <summary>
		/// Convert a temporary output stream to a final stream.
		/// </summary>
		/// <returns>The resulting final <see cref="Stream"/></returns>
		/// <seealso cref="GetTemporaryOutput"/>
		Stream ConvertTemporaryToFinal();

		/// <summary>
		/// Make a temporary copy of the original stream.
		/// </summary>
		/// <param name="stream">The <see cref="Stream"/> to copy.</param>
		/// <returns>Returns a temporary output <see cref="Stream"/> that is a copy of the input.</returns>
		Stream MakeTemporaryCopy(Stream stream);

		/// <summary>
		/// Return a stream suitable for performing direct updates on the original source.
		/// </summary>
		/// <param name="stream">The current stream.</param>
		/// <returns>Returns a stream suitable for direct updating.</returns>
		/// <remarks>This may be the current stream passed.</remarks>
		Stream OpenForDirectUpdate(Stream stream);

		/// <summary>
		/// Dispose of this instance.
		/// </summary>
		void Dispose();
	}

	/// <summary>
	/// An abstract <see cref="IArchiveStorage"/> suitable for extension by inheritance.
	/// </summary>
	abstract public class BaseArchiveStorage : IArchiveStorage
	{
		#region Constructors

		/// <summary>
		/// Initializes a new instance of the <see cref="BaseArchiveStorage"/> class.
		/// </summary>
		/// <param name="updateMode">The update mode.</param>
		protected BaseArchiveStorage(FileUpdateMode updateMode)
		{
			updateMode_ = updateMode;
		}

		#endregion Constructors

		#region IArchiveStorage Members

		/// <summary>
		/// Gets a temporary output <see cref="Stream"/>
		/// </summary>
		/// <returns>Returns the temporary output stream.</returns>
		/// <seealso cref="ConvertTemporaryToFinal"></seealso>
		public abstract Stream GetTemporaryOutput();

		/// <summary>
		/// Converts the temporary <see cref="Stream"/> to its final form.
		/// </summary>
		/// <returns>Returns a <see cref="Stream"/> that can be used to read
		/// the final storage for the archive.</returns>
		/// <seealso cref="GetTemporaryOutput"/>
		public abstract Stream ConvertTemporaryToFinal();

		/// <summary>
		/// Make a temporary copy of a <see cref="Stream"/>.
		/// </summary>
		/// <param name="stream">The <see cref="Stream"/> to make a copy of.</param>
		/// <returns>Returns a temporary output <see cref="Stream"/> that is a copy of the input.</returns>
		public abstract Stream MakeTemporaryCopy(Stream stream);

		/// <summary>
		/// Return a stream suitable for performing direct updates on the original source.
		/// </summary>
		/// <param name="stream">The <see cref="Stream"/> to open for direct update.</param>
		/// <returns>Returns a stream suitable for direct updating.</returns>
		public abstract Stream OpenForDirectUpdate(Stream stream);

		/// <summary>
		/// Disposes this instance.
		/// </summary>
		public abstract void Dispose();

		/// <summary>
		/// Gets the update mode applicable.
		/// </summary>
		/// <value>The update mode.</value>
		public FileUpdateMode UpdateMode
		{
			get
			{
				return updateMode_;
			}
		}

		#endregion IArchiveStorage Members

		#region Instance Fields

		private readonly FileUpdateMode updateMode_;

		#endregion Instance Fields
	}

	/// <summary>
	/// An <see cref="IArchiveStorage"/> implementation suitable for hard disks.
	/// </summary>
	public class DiskArchiveStorage : BaseArchiveStorage
	{
		#region Constructors

		/// <summary>
		/// Initializes a new instance of the <see cref="DiskArchiveStorage"/> class.
		/// </summary>
		/// <param name="file">The file.</param>
		/// <param name="updateMode">The update mode.</param>
		public DiskArchiveStorage(ZipFile file, FileUpdateMode updateMode)
			: base(updateMode)
		{
			if (file.Name == null)
			{
				throw new ZipException("Cant handle non file archives");
			}

			fileName_ = file.Name;
		}

		/// <summary>
		/// Initializes a new instance of the <see cref="DiskArchiveStorage"/> class.
		/// </summary>
		/// <param name="file">The file.</param>
		public DiskArchiveStorage(ZipFile file)
			: this(file, FileUpdateMode.Safe)
		{
		}

		#endregion Constructors

		#region IArchiveStorage Members

		/// <summary>
		/// Gets a temporary output <see cref="Stream"/> for performing updates on.
		/// </summary>
		/// <returns>Returns the temporary output stream.</returns>
		public override Stream GetTemporaryOutput()
		{
			temporaryName_ = PathUtils.GetTempFileName(temporaryName_);
			temporaryStream_ = File.Open(temporaryName_, FileMode.OpenOrCreate, FileAccess.Write, FileShare.None);

			return temporaryStream_;
		}

		/// <summary>
		/// Converts a temporary <see cref="Stream"/> to its final form.
		/// </summary>
		/// <returns>Returns a <see cref="Stream"/> that can be used to read
		/// the final storage for the archive.</returns>
		public override Stream ConvertTemporaryToFinal()
		{
			if (temporaryStream_ == null)
			{
				throw new ZipException("No temporary stream has been created");
			}

			Stream result = null;

			string moveTempName = PathUtils.GetTempFileName(fileName_);
			bool newFileCreated = false;

			try
			{
				temporaryStream_.Dispose();
				File.Move(fileName_, moveTempName);
				File.Move(temporaryName_, fileName_);
				newFileCreated = true;
				File.Delete(moveTempName);

				result = File.Open(fileName_, FileMode.Open, FileAccess.Read, FileShare.Read);
			}
			catch (Exception)
			{
				result = null;

				// Try to roll back changes...
				if (!newFileCreated)
				{
					File.Move(moveTempName, fileName_);
					File.Delete(temporaryName_);
				}

				throw;
			}

			return result;
		}

		/// <summary>
		/// Make a temporary copy of a stream.
		/// </summary>
		/// <param name="stream">The <see cref="Stream"/> to copy.</param>
		/// <returns>Returns a temporary output <see cref="Stream"/> that is a copy of the input.</returns>
		public override Stream MakeTemporaryCopy(Stream stream)
		{
			stream.Dispose();

			temporaryName_ = PathUtils.GetTempFileName(fileName_);
			File.Copy(fileName_, temporaryName_, true);

			temporaryStream_ = new FileStream(temporaryName_,
				FileMode.Open,
				FileAccess.ReadWrite);
			return temporaryStream_;
		}

		/// <summary>
		/// Return a stream suitable for performing direct updates on the original source.
		/// </summary>
		/// <param name="stream">The current stream.</param>
		/// <returns>Returns a stream suitable for direct updating.</returns>
		/// <remarks>If the <paramref name="stream"/> is not null this is used as is.</remarks>
		public override Stream OpenForDirectUpdate(Stream stream)
		{
			Stream result;
			if ((stream == null) || !stream.CanWrite)
			{
				if (stream != null)
				{
					stream.Dispose();
				}

				result = new FileStream(fileName_,
						FileMode.Open,
						FileAccess.ReadWrite);
			}
			else
			{
				result = stream;
			}

			return result;
		}

		/// <summary>
		/// Disposes this instance.
		/// </summary>
		public override void Dispose()
		{
			if (temporaryStream_ != null)
			{
				temporaryStream_.Dispose();
			}
		}

		#endregion IArchiveStorage Members

		#region Instance Fields

		private Stream temporaryStream_;
		private readonly string fileName_;
		private string temporaryName_;

		#endregion Instance Fields
	}

	/// <summary>
	/// An <see cref="IArchiveStorage"/> implementation suitable for in memory streams.
	/// </summary>
	public class MemoryArchiveStorage : BaseArchiveStorage
	{
		#region Constructors

		/// <summary>
		/// Initializes a new instance of the <see cref="MemoryArchiveStorage"/> class.
		/// </summary>
		public MemoryArchiveStorage()
			: base(FileUpdateMode.Direct)
		{
		}

		/// <summary>
		/// Initializes a new instance of the <see cref="MemoryArchiveStorage"/> class.
		/// </summary>
		/// <param name="updateMode">The <see cref="FileUpdateMode"/> to use</param>
		/// <remarks>This constructor is for testing as memory streams dont really require safe mode.</remarks>
		public MemoryArchiveStorage(FileUpdateMode updateMode)
			: base(updateMode)
		{
		}

		#endregion Constructors

		#region Properties

		/// <summary>
		/// Get the stream returned by <see cref="ConvertTemporaryToFinal"/> if this was in fact called.
		/// </summary>
		public MemoryStream FinalStream
		{
			get { return finalStream_; }
		}

		#endregion Properties

		#region IArchiveStorage Members

		/// <summary>
		/// Gets the temporary output <see cref="Stream"/>
		/// </summary>
		/// <returns>Returns the temporary output stream.</returns>
		public override Stream GetTemporaryOutput()
		{
			temporaryStream_ = new MemoryStream();
			return temporaryStream_;
		}

		/// <summary>
		/// Converts the temporary <see cref="Stream"/> to its final form.
		/// </summary>
		/// <returns>Returns a <see cref="Stream"/> that can be used to read
		/// the final storage for the archive.</returns>
		public override Stream ConvertTemporaryToFinal()
		{
			if (temporaryStream_ == null)
			{
				throw new ZipException("No temporary stream has been created");
			}

			finalStream_ = new MemoryStream(temporaryStream_.ToArray());
			return finalStream_;
		}

		/// <summary>
		/// Make a temporary copy of the original stream.
		/// </summary>
		/// <param name="stream">The <see cref="Stream"/> to copy.</param>
		/// <returns>Returns a temporary output <see cref="Stream"/> that is a copy of the input.</returns>
		public override Stream MakeTemporaryCopy(Stream stream)
		{
			temporaryStream_ = new MemoryStream();
			stream.Position = 0;
			StreamUtils.Copy(stream, temporaryStream_, new byte[4096]);
			return temporaryStream_;
		}

		/// <summary>
		/// Return a stream suitable for performing direct updates on the original source.
		/// </summary>
		/// <param name="stream">The original source stream</param>
		/// <returns>Returns a stream suitable for direct updating.</returns>
		/// <remarks>If the <paramref name="stream"/> passed is not null this is used;
		/// otherwise a new <see cref="MemoryStream"/> is returned.</remarks>
		public override Stream OpenForDirectUpdate(Stream stream)
		{
			Stream result;
			if ((stream == null) || !stream.CanWrite)
			{
				result = new MemoryStream();

				if (stream != null)
				{
					stream.Position = 0;
					StreamUtils.Copy(stream, result, new byte[4096]);

					stream.Dispose();
				}
			}
			else
			{
				result = stream;
			}

			return result;
		}

		/// <summary>
		/// Disposes this instance.
		/// </summary>
		public override void Dispose()
		{
			if (temporaryStream_ != null)
			{
				temporaryStream_.Dispose();
			}
		}

		#endregion IArchiveStorage Members

		#region Instance Fields

		private MemoryStream temporaryStream_;
		private MemoryStream finalStream_;

		#endregion Instance Fields
	}

	#endregion Archive Storage
}


namespace ICSharpCode.SharpZipLib.Zip
{
	/// <summary>
	/// Holds data pertinent to a data descriptor.
	/// </summary>
	public class DescriptorData
	{
		/// <summary>
		/// Get /set the compressed size of data.
		/// </summary>
		public long CompressedSize
		{
			get { return compressedSize; }
			set { compressedSize = value; }
		}

		/// <summary>
		/// Get / set the uncompressed size of data
		/// </summary>
		public long Size
		{
			get { return size; }
			set { size = value; }
		}

		/// <summary>
		/// Get /set the crc value.
		/// </summary>
		public long Crc
		{
			get { return crc; }
			set { crc = (value & 0xffffffff); }
		}

		#region Instance Fields

		private long size;
		private long compressedSize;
		private long crc;

		#endregion Instance Fields
	}

	internal class EntryPatchData
	{
		public long SizePatchOffset
		{
			get { return sizePatchOffset_; }
			set { sizePatchOffset_ = value; }
		}

		public long CrcPatchOffset
		{
			get { return crcPatchOffset_; }
			set { crcPatchOffset_ = value; }
		}

		#region Instance Fields

		private long sizePatchOffset_;
		private long crcPatchOffset_;

		#endregion Instance Fields
	}

	/// <summary>
	/// This class assists with writing/reading from Zip files.
	/// </summary>
	internal class ZipHelperStream : Stream
	{
		#region Constructors

		/// <summary>
		/// Initialise an instance of this class.
		/// </summary>
		/// <param name="name">The name of the file to open.</param>
		public ZipHelperStream(string name)
		{
			stream_ = new FileStream(name, FileMode.Open, FileAccess.ReadWrite);
			isOwner_ = true;
		}

		/// <summary>
		/// Initialise a new instance of <see cref="ZipHelperStream"/>.
		/// </summary>
		/// <param name="stream">The stream to use.</param>
		public ZipHelperStream(Stream stream)
		{
			stream_ = stream;
		}

		#endregion Constructors

		/// <summary>
		/// Get / set a value indicating whether the underlying stream is owned or not.
		/// </summary>
		/// <remarks>If the stream is owned it is closed when this instance is closed.</remarks>
		public bool IsStreamOwner
		{
			get { return isOwner_; }
			set { isOwner_ = value; }
		}

		#region Base Stream Methods

		public override bool CanRead
		{
			get { return stream_.CanRead; }
		}

		public override bool CanSeek
		{
			get { return stream_.CanSeek; }
		}

		public override bool CanTimeout
		{
			get { return stream_.CanTimeout; }
		}

		public override long Length
		{
			get { return stream_.Length; }
		}

		public override long Position
		{
			get { return stream_.Position; }
			set { stream_.Position = value; }
		}

		public override bool CanWrite
		{
			get { return stream_.CanWrite; }
		}

		public override void Flush()
		{
			stream_.Flush();
		}

		public override long Seek(long offset, SeekOrigin origin)
		{
			return stream_.Seek(offset, origin);
		}

		public override void SetLength(long value)
		{
			stream_.SetLength(value);
		}

		public override int Read(byte[] buffer, int offset, int count)
		{
			return stream_.Read(buffer, offset, count);
		}

		public override void Write(byte[] buffer, int offset, int count)
		{
			stream_.Write(buffer, offset, count);
		}

		/// <summary>
		/// Close the stream.
		/// </summary>
		/// <remarks>
		/// The underlying stream is closed only if <see cref="IsStreamOwner"/> is true.
		/// </remarks>
		protected override void Dispose(bool disposing)
		{
			Stream toClose = stream_;
			stream_ = null;
			if (isOwner_ && (toClose != null))
			{
				isOwner_ = false;
				toClose.Dispose();
			}
		}

		#endregion Base Stream Methods

		// Write the local file header
		// TODO: ZipHelperStream.WriteLocalHeader is not yet used and needs checking for ZipFile and ZipOuptutStream usage
		private void WriteLocalHeader(ZipEntry entry, EntryPatchData patchData)
		{
			CompressionMethod method = entry.CompressionMethod;
			bool headerInfoAvailable = true; // How to get this?
			bool patchEntryHeader = false;

			WriteLEInt(ZipConstants.LocalHeaderSignature);

			WriteLEShort(entry.Version);
			WriteLEShort(entry.Flags);
			WriteLEShort((byte)method);
			WriteLEInt((int)entry.DosTime);

			if (headerInfoAvailable == true)
			{
				WriteLEInt((int)entry.Crc);
				if (entry.LocalHeaderRequiresZip64)
				{
					WriteLEInt(-1);
					WriteLEInt(-1);
				}
				else
				{
					WriteLEInt(entry.IsCrypted ? (int)entry.CompressedSize + ZipConstants.CryptoHeaderSize : (int)entry.CompressedSize);
					WriteLEInt((int)entry.Size);
				}
			}
			else
			{
				if (patchData != null)
				{
					patchData.CrcPatchOffset = stream_.Position;
				}
				WriteLEInt(0);  // Crc

				if (patchData != null)
				{
					patchData.SizePatchOffset = stream_.Position;
				}

				// For local header both sizes appear in Zip64 Extended Information
				if (entry.LocalHeaderRequiresZip64 && patchEntryHeader)
				{
					WriteLEInt(-1);
					WriteLEInt(-1);
				}
				else
				{
					WriteLEInt(0);  // Compressed size
					WriteLEInt(0);  // Uncompressed size
				}
			}

			byte[] name = ZipStrings.ConvertToArray(entry.Flags, entry.Name);

			if (name.Length > 0xFFFF)
			{
				throw new ZipException("Entry name too long.");
			}

			var ed = new ZipExtraData(entry.ExtraData);

			if (entry.LocalHeaderRequiresZip64 && (headerInfoAvailable || patchEntryHeader))
			{
				ed.StartNewEntry();
				if (headerInfoAvailable)
				{
					ed.AddLeLong(entry.Size);
					ed.AddLeLong(entry.CompressedSize);
				}
				else
				{
					ed.AddLeLong(-1);
					ed.AddLeLong(-1);
				}
				ed.AddNewEntry(1);

				if (!ed.Find(1))
				{
					throw new ZipException("Internal error cant find extra data");
				}

				if (patchData != null)
				{
					patchData.SizePatchOffset = ed.CurrentReadIndex;
				}
			}
			else
			{
				ed.Delete(1);
			}

			byte[] extra = ed.GetEntryData();

			WriteLEShort(name.Length);
			WriteLEShort(extra.Length);

			if (name.Length > 0)
			{
				stream_.Write(name, 0, name.Length);
			}

			if (entry.LocalHeaderRequiresZip64 && patchEntryHeader)
			{
				patchData.SizePatchOffset += stream_.Position;
			}

			if (extra.Length > 0)
			{
				stream_.Write(extra, 0, extra.Length);
			}
		}

		/// <summary>
		/// Locates a block with the desired <paramref name="signature"/>.
		/// </summary>
		/// <param name="signature">The signature to find.</param>
		/// <param name="endLocation">Location, marking the end of block.</param>
		/// <param name="minimumBlockSize">Minimum size of the block.</param>
		/// <param name="maximumVariableData">The maximum variable data.</param>
		/// <returns>Returns the offset of the first byte after the signature; -1 if not found</returns>
		public long LocateBlockWithSignature(int signature, long endLocation, int minimumBlockSize, int maximumVariableData)
		{
			long pos = endLocation - minimumBlockSize;
			if (pos < 0)
			{
				return -1;
			}

			long giveUpMarker = Math.Max(pos - maximumVariableData, 0);

			// TODO: This loop could be optimised for speed.
			do
			{
				if (pos < giveUpMarker)
				{
					return -1;
				}
				Seek(pos--, SeekOrigin.Begin);
			} while (ReadLEInt() != signature);

			return Position;
		}

		/// <summary>
		/// Write Zip64 end of central directory records (File header and locator).
		/// </summary>
		/// <param name="noOfEntries">The number of entries in the central directory.</param>
		/// <param name="sizeEntries">The size of entries in the central directory.</param>
		/// <param name="centralDirOffset">The offset of the central directory.</param>
		public void WriteZip64EndOfCentralDirectory(long noOfEntries, long sizeEntries, long centralDirOffset)
		{
			long centralSignatureOffset = centralDirOffset + sizeEntries;
			WriteLEInt(ZipConstants.Zip64CentralFileHeaderSignature);
			WriteLELong(44);    // Size of this record (total size of remaining fields in header or full size - 12)
			WriteLEShort(ZipConstants.VersionMadeBy);   // Version made by
			WriteLEShort(ZipConstants.VersionZip64);   // Version to extract
			WriteLEInt(0);      // Number of this disk
			WriteLEInt(0);      // number of the disk with the start of the central directory
			WriteLELong(noOfEntries);       // No of entries on this disk
			WriteLELong(noOfEntries);       // Total No of entries in central directory
			WriteLELong(sizeEntries);       // Size of the central directory
			WriteLELong(centralDirOffset);  // offset of start of central directory
											// zip64 extensible data sector not catered for here (variable size)

			// Write the Zip64 end of central directory locator
			WriteLEInt(ZipConstants.Zip64CentralDirLocatorSignature);

			// no of the disk with the start of the zip64 end of central directory
			WriteLEInt(0);

			// relative offset of the zip64 end of central directory record
			WriteLELong(centralSignatureOffset);

			// total number of disks
			WriteLEInt(1);
		}

		/// <summary>
		/// Write the required records to end the central directory.
		/// </summary>
		/// <param name="noOfEntries">The number of entries in the directory.</param>
		/// <param name="sizeEntries">The size of the entries in the directory.</param>
		/// <param name="startOfCentralDirectory">The start of the central directory.</param>
		/// <param name="comment">The archive comment.  (This can be null).</param>
		public void WriteEndOfCentralDirectory(long noOfEntries, long sizeEntries,
			long startOfCentralDirectory, byte[] comment)
		{
			if ((noOfEntries >= 0xffff) ||
				(startOfCentralDirectory >= 0xffffffff) ||
				(sizeEntries >= 0xffffffff))
			{
				WriteZip64EndOfCentralDirectory(noOfEntries, sizeEntries, startOfCentralDirectory);
			}

			WriteLEInt(ZipConstants.EndOfCentralDirectorySignature);

			// TODO: ZipFile Multi disk handling not done
			WriteLEShort(0);                    // number of this disk
			WriteLEShort(0);                    // no of disk with start of central dir

			// Number of entries
			if (noOfEntries >= 0xffff)
			{
				WriteLEUshort(0xffff);  // Zip64 marker
				WriteLEUshort(0xffff);
			}
			else
			{
				WriteLEShort((short)noOfEntries);          // entries in central dir for this disk
				WriteLEShort((short)noOfEntries);          // total entries in central directory
			}

			// Size of the central directory
			if (sizeEntries >= 0xffffffff)
			{
				WriteLEUint(0xffffffff);    // Zip64 marker
			}
			else
			{
				WriteLEInt((int)sizeEntries);
			}

			// offset of start of central directory
			if (startOfCentralDirectory >= 0xffffffff)
			{
				WriteLEUint(0xffffffff);    // Zip64 marker
			}
			else
			{
				WriteLEInt((int)startOfCentralDirectory);
			}

			int commentLength = (comment != null) ? comment.Length : 0;

			if (commentLength > 0xffff)
			{
				throw new ZipException(string.Format("Comment length({0}) is too long can only be 64K", commentLength));
			}

			WriteLEShort(commentLength);

			if (commentLength > 0)
			{
				Write(comment, 0, comment.Length);
			}
		}

		#region LE value reading/writing

		/// <summary>
		/// Read an unsigned short in little endian byte order.
		/// </summary>
		/// <returns>Returns the value read.</returns>
		/// <exception cref="IOException">
		/// An i/o error occurs.
		/// </exception>
		/// <exception cref="EndOfStreamException">
		/// The file ends prematurely
		/// </exception>
		public int ReadLEShort()
		{
			int byteValue1 = stream_.ReadByte();

			if (byteValue1 < 0)
			{
				throw new EndOfStreamException();
			}

			int byteValue2 = stream_.ReadByte();
			if (byteValue2 < 0)
			{
				throw new EndOfStreamException();
			}

			return byteValue1 | (byteValue2 << 8);
		}

		/// <summary>
		/// Read an int in little endian byte order.
		/// </summary>
		/// <returns>Returns the value read.</returns>
		/// <exception cref="IOException">
		/// An i/o error occurs.
		/// </exception>
		/// <exception cref="System.IO.EndOfStreamException">
		/// The file ends prematurely
		/// </exception>
		public int ReadLEInt()
		{
			return ReadLEShort() | (ReadLEShort() << 16);
		}

		/// <summary>
		/// Read a long in little endian byte order.
		/// </summary>
		/// <returns>The value read.</returns>
		public long ReadLELong()
		{
			return (uint)ReadLEInt() | ((long)ReadLEInt() << 32);
		}

		/// <summary>
		/// Write an unsigned short in little endian byte order.
		/// </summary>
		/// <param name="value">The value to write.</param>
		public void WriteLEShort(int value)
		{
			stream_.WriteByte((byte)(value & 0xff));
			stream_.WriteByte((byte)((value >> 8) & 0xff));
		}

		/// <summary>
		/// Write a ushort in little endian byte order.
		/// </summary>
		/// <param name="value">The value to write.</param>
		public void WriteLEUshort(ushort value)
		{
			stream_.WriteByte((byte)(value & 0xff));
			stream_.WriteByte((byte)(value >> 8));
		}

		/// <summary>
		/// Write an int in little endian byte order.
		/// </summary>
		/// <param name="value">The value to write.</param>
		public void WriteLEInt(int value)
		{
			WriteLEShort(value);
			WriteLEShort(value >> 16);
		}

		/// <summary>
		/// Write a uint in little endian byte order.
		/// </summary>
		/// <param name="value">The value to write.</param>
		public void WriteLEUint(uint value)
		{
			WriteLEUshort((ushort)(value & 0xffff));
			WriteLEUshort((ushort)(value >> 16));
		}

		/// <summary>
		/// Write a long in little endian byte order.
		/// </summary>
		/// <param name="value">The value to write.</param>
		public void WriteLELong(long value)
		{
			WriteLEInt((int)value);
			WriteLEInt((int)(value >> 32));
		}

		/// <summary>
		/// Write a ulong in little endian byte order.
		/// </summary>
		/// <param name="value">The value to write.</param>
		public void WriteLEUlong(ulong value)
		{
			WriteLEUint((uint)(value & 0xffffffff));
			WriteLEUint((uint)(value >> 32));
		}

		#endregion LE value reading/writing

		/// <summary>
		/// Write a data descriptor.
		/// </summary>
		/// <param name="entry">The entry to write a descriptor for.</param>
		/// <returns>Returns the number of descriptor bytes written.</returns>
		public int WriteDataDescriptor(ZipEntry entry)
		{
			if (entry == null)
			{
				throw new ArgumentNullException(nameof(entry));
			}

			int result = 0;

			// Add data descriptor if flagged as required
			if ((entry.Flags & (int)GeneralBitFlags.Descriptor) != 0)
			{
				// The signature is not PKZIP originally but is now described as optional
				// in the PKZIP Appnote documenting the format.
				WriteLEInt(ZipConstants.DataDescriptorSignature);
				WriteLEInt(unchecked((int)(entry.Crc)));

				result += 8;

				if (entry.LocalHeaderRequiresZip64)
				{
					WriteLELong(entry.CompressedSize);
					WriteLELong(entry.Size);
					result += 16;
				}
				else
				{
					WriteLEInt((int)entry.CompressedSize);
					WriteLEInt((int)entry.Size);
					result += 8;
				}
			}

			return result;
		}

		/// <summary>
		/// Read data descriptor at the end of compressed data.
		/// </summary>
		/// <param name="zip64">if set to <c>true</c> [zip64].</param>
		/// <param name="data">The data to fill in.</param>
		/// <returns>Returns the number of bytes read in the descriptor.</returns>
		public void ReadDataDescriptor(bool zip64, DescriptorData data)
		{
			int intValue = ReadLEInt();

			// In theory this may not be a descriptor according to PKZIP appnote.
			// In practice its always there.
			if (intValue != ZipConstants.DataDescriptorSignature)
			{
				throw new ZipException("Data descriptor signature not found");
			}

			data.Crc = ReadLEInt();

			if (zip64)
			{
				data.CompressedSize = ReadLELong();
				data.Size = ReadLELong();
			}
			else
			{
				data.CompressedSize = ReadLEInt();
				data.Size = ReadLEInt();
			}
		}

		#region Instance Fields

		private bool isOwner_;
		private Stream stream_;

		#endregion Instance Fields
	}
}


namespace ICSharpCode.SharpZipLib.Zip
{
	/// <summary>
	/// This is an InflaterInputStream that reads the files baseInputStream an zip archive
	/// one after another.  It has a special method to get the zip entry of
	/// the next file.  The zip entry contains information about the file name
	/// size, compressed size, Crc, etc.
	/// It includes support for Stored and Deflated entries.
	/// <br/>
	/// <br/>Author of the original java version : Jochen Hoenicke
	/// </summary>
	///
	/// <example> This sample shows how to read a zip file
	/// <code lang="C#">
	///
	/// using ICSharpCode.SharpZipLib.Zip;
	///
	/// class MainClass
	/// {
	/// 	public static void Main(string[] args)
	/// 	{
	/// 		using ( ZipInputStream s = new ZipInputStream(File.OpenRead(args[0]))) {
	///
	/// 			ZipEntry theEntry;
	/// 			const int size = 2048;
	/// 			byte[] data = new byte[2048];
	///
	/// 			while ((theEntry = s.GetNextEntry()) != null) {
	///                 if ( entry.IsFile ) {
	/// 				    Console.Write("Show contents (y/n) ?");
	/// 				    if (Console.ReadLine() == "y") {
	/// 				    	while (true) {
	/// 				    		size = s.Read(data, 0, data.Length);
	/// 				    		if (size > 0) {
	/// 				    			Console.Write(new ASCIIEncoding().GetString(data, 0, size));
	/// 				    		} else {
	/// 				    			break;
	/// 				    		}
	/// 				    	}
	/// 				    }
	/// 				}
	/// 			}
	/// 		}
	/// 	}
	/// }
	/// </code>
	/// </example>
	public class ZipInputStream : InflaterInputStream
	{
		#region Instance Fields

		/// <summary>
		/// Delegate for reading bytes from a stream.
		/// </summary>
		private delegate int ReadDataHandler(byte[] b, int offset, int length);

		/// <summary>
		/// The current reader this instance.
		/// </summary>
		private ReadDataHandler internalReader;

		private Crc32 crc = new Crc32();
		private ZipEntry entry;

		private long size;
		private CompressionMethod method;
		private int flags;
		private string password;

		#endregion Instance Fields

		#region Constructors

		/// <summary>
		/// Creates a new Zip input stream, for reading a zip archive.
		/// </summary>
		/// <param name="baseInputStream">The underlying <see cref="Stream"/> providing data.</param>
		public ZipInputStream(Stream baseInputStream)
			: base(baseInputStream, new Inflater(true))
		{
			internalReader = new ReadDataHandler(ReadingNotAvailable);
		}

		/// <summary>
		/// Creates a new Zip input stream, for reading a zip archive.
		/// </summary>
		/// <param name="baseInputStream">The underlying <see cref="Stream"/> providing data.</param>
		/// <param name="bufferSize">Size of the buffer.</param>
		public ZipInputStream(Stream baseInputStream, int bufferSize)
			: base(baseInputStream, new Inflater(true), bufferSize)
		{
			internalReader = new ReadDataHandler(ReadingNotAvailable);
		}

		#endregion Constructors

		/// <summary>
		/// Optional password used for encryption when non-null
		/// </summary>
		/// <value>A password for all encrypted <see cref="ZipEntry">entries </see> in this <see cref="ZipInputStream"/></value>
		public string Password
		{
			get
			{
				return password;
			}
			set
			{
				password = value;
			}
		}

		/// <summary>
		/// Gets a value indicating if there is a current entry and it can be decompressed
		/// </summary>
		/// <remarks>
		/// The entry can only be decompressed if the library supports the zip features required to extract it.
		/// See the <see cref="ZipEntry.Version">ZipEntry Version</see> property for more details.
		///
		/// Since <see cref="ZipInputStream"/> uses the local headers for extraction, entries with no compression combined with the
		/// <see cref="GeneralBitFlags.Descriptor"/> flag set, cannot be extracted as the end of the entry data cannot be deduced.
		/// </remarks>
		public bool CanDecompressEntry 
			=> entry != null
			&& IsEntryCompressionMethodSupported(entry)
			&& entry.CanDecompress
			&& (!entry.HasFlag(GeneralBitFlags.Descriptor) || entry.CompressionMethod != CompressionMethod.Stored || entry.IsCrypted);

		/// <summary>
		/// Is the compression method for the specified entry supported?
		/// </summary>
		/// <remarks>
		/// Uses entry.CompressionMethodForHeader so that entries of type WinZipAES will be rejected. 
		/// </remarks>
		/// <param name="entry">the entry to check.</param>
		/// <returns>true if the compression method is supported, false if not.</returns>
		private static bool IsEntryCompressionMethodSupported(ZipEntry entry)
		{
			var entryCompressionMethod = entry.CompressionMethodForHeader;

			return entryCompressionMethod == CompressionMethod.Deflated ||
				   entryCompressionMethod == CompressionMethod.Stored;
		}

		/// <summary>
		/// Advances to the next entry in the archive
		/// </summary>
		/// <returns>
		/// The next <see cref="ZipEntry">entry</see> in the archive or null if there are no more entries.
		/// </returns>
		/// <remarks>
		/// If the previous entry is still open <see cref="CloseEntry">CloseEntry</see> is called.
		/// </remarks>
		/// <exception cref="InvalidOperationException">
		/// Input stream is closed
		/// </exception>
		/// <exception cref="ZipException">
		/// Password is not set, password is invalid, compression method is invalid,
		/// version required to extract is not supported
		/// </exception>
		public ZipEntry GetNextEntry()
		{
			if (crc == null)
			{
				throw new InvalidOperationException("Closed.");
			}

			if (entry != null)
			{
				CloseEntry();
			}

			int header = inputBuffer.ReadLeInt();

			if (header == ZipConstants.CentralHeaderSignature ||
				header == ZipConstants.EndOfCentralDirectorySignature ||
				header == ZipConstants.CentralHeaderDigitalSignature ||
				header == ZipConstants.ArchiveExtraDataSignature ||
				header == ZipConstants.Zip64CentralFileHeaderSignature)
			{
				// No more individual entries exist
				Dispose();
				return null;
			}

			// -jr- 07-Dec-2003 Ignore spanning temporary signatures if found
			// Spanning signature is same as descriptor signature and is untested as yet.
			if ((header == ZipConstants.SpanningTempSignature) || (header == ZipConstants.SpanningSignature))
			{
				header = inputBuffer.ReadLeInt();
			}

			if (header != ZipConstants.LocalHeaderSignature)
			{
				throw new ZipException("Wrong Local header signature: 0x" + String.Format("{0:X}", header));
			}

			var versionRequiredToExtract = (short)inputBuffer.ReadLeShort();

			flags = inputBuffer.ReadLeShort();
			method = (CompressionMethod)inputBuffer.ReadLeShort();
			var dostime = (uint)inputBuffer.ReadLeInt();
			int crc2 = inputBuffer.ReadLeInt();
			csize = inputBuffer.ReadLeInt();
			size = inputBuffer.ReadLeInt();
			int nameLen = inputBuffer.ReadLeShort();
			int extraLen = inputBuffer.ReadLeShort();

			bool isCrypted = (flags & 1) == 1;

			byte[] buffer = new byte[nameLen];
			inputBuffer.ReadRawBuffer(buffer);

			string name = ZipStrings.ConvertToStringExt(flags, buffer);

			entry = new ZipEntry(name, versionRequiredToExtract, ZipConstants.VersionMadeBy, method)
			{
				Flags = flags,
			};

			if ((flags & 8) == 0)
			{
				entry.Crc = crc2 & 0xFFFFFFFFL;
				entry.Size = size & 0xFFFFFFFFL;
				entry.CompressedSize = csize & 0xFFFFFFFFL;

				entry.CryptoCheckValue = (byte)((crc2 >> 24) & 0xff);
			}
			else
			{
				// This allows for GNU, WinZip and possibly other archives, the PKZIP spec
				// says these values are zero under these circumstances.
				if (crc2 != 0)
				{
					entry.Crc = crc2 & 0xFFFFFFFFL;
				}

				if (size != 0)
				{
					entry.Size = size & 0xFFFFFFFFL;
				}

				if (csize != 0)
				{
					entry.CompressedSize = csize & 0xFFFFFFFFL;
				}

				entry.CryptoCheckValue = (byte)((dostime >> 8) & 0xff);
			}

			entry.DosTime = dostime;

			// If local header requires Zip64 is true then the extended header should contain
			// both values.

			// Handle extra data if present.  This can set/alter some fields of the entry.
			if (extraLen > 0)
			{
				byte[] extra = new byte[extraLen];
				inputBuffer.ReadRawBuffer(extra);
				entry.ExtraData = extra;
			}

			entry.ProcessExtraData(true);
			if (entry.CompressedSize >= 0)
			{
				csize = entry.CompressedSize;
			}

			if (entry.Size >= 0)
			{
				size = entry.Size;
			}

			if (method == CompressionMethod.Stored && (!isCrypted && csize != size || (isCrypted && csize - ZipConstants.CryptoHeaderSize != size)))
			{
				throw new ZipException("Stored, but compressed != uncompressed");
			}

			// Determine how to handle reading of data if this is attempted.
			if (IsEntryCompressionMethodSupported(entry))
			{
				internalReader = new ReadDataHandler(InitialRead);
			}
			else
			{
				internalReader = new ReadDataHandler(ReadingNotSupported);
			}

			return entry;
		}

		/// <summary>
		/// Read data descriptor at the end of compressed data.
		/// </summary>
		private void ReadDataDescriptor()
		{
			if (inputBuffer.ReadLeInt() != ZipConstants.DataDescriptorSignature)
			{
				throw new ZipException("Data descriptor signature not found");
			}

			entry.Crc = inputBuffer.ReadLeInt() & 0xFFFFFFFFL;

			if (entry.LocalHeaderRequiresZip64)
			{
				csize = inputBuffer.ReadLeLong();
				size = inputBuffer.ReadLeLong();
			}
			else
			{
				csize = inputBuffer.ReadLeInt();
				size = inputBuffer.ReadLeInt();
			}
			entry.CompressedSize = csize;
			entry.Size = size;
		}

		/// <summary>
		/// Complete cleanup as the final part of closing.
		/// </summary>
		/// <param name="testCrc">True if the crc value should be tested</param>
		private void CompleteCloseEntry(bool testCrc)
		{
			StopDecrypting();

			if ((flags & 8) != 0)
			{
				ReadDataDescriptor();
			}

			size = 0;

			if (testCrc &&
				((crc.Value & 0xFFFFFFFFL) != entry.Crc) && (entry.Crc != -1))
			{
				throw new ZipException("CRC mismatch");
			}

			crc.Reset();

			if (method == CompressionMethod.Deflated)
			{
				inf.Reset();
			}
			entry = null;
		}

		/// <summary>
		/// Closes the current zip entry and moves to the next one.
		/// </summary>
		/// <exception cref="InvalidOperationException">
		/// The stream is closed
		/// </exception>
		/// <exception cref="ZipException">
		/// The Zip stream ends early
		/// </exception>
		public void CloseEntry()
		{
			if (crc == null)
			{
				throw new InvalidOperationException("Closed");
			}

			if (entry == null)
			{
				return;
			}

			if (method == CompressionMethod.Deflated)
			{
				if ((flags & 8) != 0)
				{
					// We don't know how much we must skip, read until end.
					byte[] tmp = new byte[4096];

					// Read will close this entry
					while (Read(tmp, 0, tmp.Length) > 0)
					{
					}
					return;
				}

				csize -= inf.TotalIn;
				inputBuffer.Available += inf.RemainingInput;
			}

			if ((inputBuffer.Available > csize) && (csize >= 0))
			{
				inputBuffer.Available = (int)((long)inputBuffer.Available - csize);
			}
			else
			{
				csize -= inputBuffer.Available;
				inputBuffer.Available = 0;
				while (csize != 0)
				{
					long skipped = Skip(csize);

					if (skipped <= 0)
					{
						throw new ZipException("Zip archive ends early.");
					}

					csize -= skipped;
				}
			}

			CompleteCloseEntry(false);
		}

		/// <summary>
		/// Returns 1 if there is an entry available
		/// Otherwise returns 0.
		/// </summary>
		public override int Available
		{
			get
			{
				return entry != null ? 1 : 0;
			}
		}

		/// <summary>
		/// Returns the current size that can be read from the current entry if available
		/// </summary>
		/// <exception cref="ZipException">Thrown if the entry size is not known.</exception>
		/// <exception cref="InvalidOperationException">Thrown if no entry is currently available.</exception>
		public override long Length
		{
			get
			{
				if (entry != null)
				{
					if (entry.Size >= 0)
					{
						return entry.Size;
					}
					else
					{
						throw new ZipException("Length not available for the current entry");
					}
				}
				else
				{
					throw new InvalidOperationException("No current entry");
				}
			}
		}

		/// <summary>
		/// Reads a byte from the current zip entry.
		/// </summary>
		/// <returns>
		/// The byte or -1 if end of stream is reached.
		/// </returns>
		public override int ReadByte()
		{
			byte[] b = new byte[1];
			if (Read(b, 0, 1) <= 0)
			{
				return -1;
			}
			return b[0] & 0xff;
		}

		/// <summary>
		/// Handle attempts to read by throwing an <see cref="InvalidOperationException"/>.
		/// </summary>
		/// <param name="destination">The destination array to store data in.</param>
		/// <param name="offset">The offset at which data read should be stored.</param>
		/// <param name="count">The maximum number of bytes to read.</param>
		/// <returns>Returns the number of bytes actually read.</returns>
		private int ReadingNotAvailable(byte[] destination, int offset, int count)
		{
			throw new InvalidOperationException("Unable to read from this stream");
		}

		/// <summary>
		/// Handle attempts to read from this entry by throwing an exception
		/// </summary>
		private int ReadingNotSupported(byte[] destination, int offset, int count)
		{
			throw new ZipException("The compression method for this entry is not supported");
		}

		/// <summary>
		/// Handle attempts to read from this entry by throwing an exception
		/// </summary>
		private int StoredDescriptorEntry(byte[] destination, int offset, int count) =>
			throw new StreamUnsupportedException(
				"The combination of Stored compression method and Descriptor flag is not possible to read using ZipInputStream");
		

		/// <summary>
		/// Perform the initial read on an entry which may include
		/// reading encryption headers and setting up inflation.
		/// </summary>
		/// <param name="destination">The destination to fill with data read.</param>
		/// <param name="offset">The offset to start reading at.</param>
		/// <param name="count">The maximum number of bytes to read.</param>
		/// <returns>The actual number of bytes read.</returns>
		private int InitialRead(byte[] destination, int offset, int count)
		{
			var usesDescriptor = (entry.Flags & (int)GeneralBitFlags.Descriptor) != 0;

			// Handle encryption if required.
			if (entry.IsCrypted)
			{
				if (password == null)
				{
					throw new ZipException("No password set.");
				}

				// Generate and set crypto transform...
				var managed = new PkzipClassicManaged();
				byte[] key = PkzipClassic.GenerateKeys(ZipStrings.ConvertToArray(password));

				inputBuffer.CryptoTransform = managed.CreateDecryptor(key, null);

				byte[] cryptbuffer = new byte[ZipConstants.CryptoHeaderSize];
				inputBuffer.ReadClearTextBuffer(cryptbuffer, 0, ZipConstants.CryptoHeaderSize);

				if (cryptbuffer[ZipConstants.CryptoHeaderSize - 1] != entry.CryptoCheckValue)
				{
					throw new ZipException("Invalid password");
				}

				if (csize >= ZipConstants.CryptoHeaderSize)
				{
					csize -= ZipConstants.CryptoHeaderSize;
				}
				else if (!usesDescriptor)
				{
					throw new ZipException($"Entry compressed size {csize} too small for encryption");
				}
			}
			else
			{
				inputBuffer.CryptoTransform = null;
			}

			if (csize > 0 || usesDescriptor)
			{
				if (method == CompressionMethod.Deflated && inputBuffer.Available > 0)
				{
					inputBuffer.SetInflaterInput(inf);
				}

				// It's not possible to know how many bytes to read when using "Stored" compression (unless using encryption)
				if (!entry.IsCrypted && method == CompressionMethod.Stored && usesDescriptor)
				{
					internalReader = StoredDescriptorEntry;
					return StoredDescriptorEntry(destination, offset, count);
				}

				if (!CanDecompressEntry)
				{
					internalReader = ReadingNotSupported;
					return ReadingNotSupported(destination, offset, count);
				}

				internalReader = BodyRead;
				return BodyRead(destination, offset, count);
			}
			

			internalReader = ReadingNotAvailable;
			return 0;
		}

		/// <summary>
		/// Read a block of bytes from the stream.
		/// </summary>
		/// <param name="buffer">The destination for the bytes.</param>
		/// <param name="offset">The index to start storing data.</param>
		/// <param name="count">The number of bytes to attempt to read.</param>
		/// <returns>Returns the number of bytes read.</returns>
		/// <remarks>Zero bytes read means end of stream.</remarks>
		public override int Read(byte[] buffer, int offset, int count)
		{
			if (buffer == null)
			{
				throw new ArgumentNullException(nameof(buffer));
			}

			if (offset < 0)
			{
				throw new ArgumentOutOfRangeException(nameof(offset), "Cannot be negative");
			}

			if (count < 0)
			{
				throw new ArgumentOutOfRangeException(nameof(count), "Cannot be negative");
			}

			if ((buffer.Length - offset) < count)
			{
				throw new ArgumentException("Invalid offset/count combination");
			}

			return internalReader(buffer, offset, count);
		}

		/// <summary>
		/// Reads a block of bytes from the current zip entry.
		/// </summary>
		/// <returns>
		/// The number of bytes read (this may be less than the length requested, even before the end of stream), or 0 on end of stream.
		/// </returns>
		/// <exception cref="IOException">
		/// An i/o error occurred.
		/// </exception>
		/// <exception cref="ZipException">
		/// The deflated stream is corrupted.
		/// </exception>
		/// <exception cref="InvalidOperationException">
		/// The stream is not open.
		/// </exception>
		private int BodyRead(byte[] buffer, int offset, int count)
		{
			if (crc == null)
			{
				throw new InvalidOperationException("Closed");
			}

			if ((entry == null) || (count <= 0))
			{
				return 0;
			}

			if (offset + count > buffer.Length)
			{
				throw new ArgumentException("Offset + count exceeds buffer size");
			}

			bool finished = false;

			switch (method)
			{
				case CompressionMethod.Deflated:
					count = base.Read(buffer, offset, count);
					if (count <= 0)
					{
						if (!inf.IsFinished)
						{
							throw new ZipException("Inflater not finished!");
						}
						inputBuffer.Available = inf.RemainingInput;

						// A csize of -1 is from an unpatched local header
						if ((flags & 8) == 0 &&
							(inf.TotalIn != csize && csize != 0xFFFFFFFF && csize != -1 || inf.TotalOut != size))
						{
							throw new ZipException("Size mismatch: " + csize + ";" + size + " <-> " + inf.TotalIn + ";" + inf.TotalOut);
						}
						inf.Reset();
						finished = true;
					}
					break;

				case CompressionMethod.Stored:
					if ((count > csize) && (csize >= 0))
					{
						count = (int)csize;
					}

					if (count > 0)
					{
						count = inputBuffer.ReadClearTextBuffer(buffer, offset, count);
						if (count > 0)
						{
							csize -= count;
							size -= count;
						}
					}

					if (csize == 0)
					{
						finished = true;
					}
					else
					{
						if (count < 0)
						{
							throw new ZipException("EOF in stored block");
						}
					}
					break;
			}

			if (count > 0)
			{
				crc.Update(new ArraySegment<byte>(buffer, offset, count));
			}

			if (finished)
			{
				CompleteCloseEntry(true);
			}

			return count;
		}

		/// <summary>
		/// Closes the zip input stream
		/// </summary>
		protected override void Dispose(bool disposing)
		{
			internalReader = new ReadDataHandler(ReadingNotAvailable);
			crc = null;
			entry = null;

			base.Dispose(disposing);
		}
	}
}


namespace ICSharpCode.SharpZipLib.Zip
{
	/// <summary>
	/// ZipNameTransform transforms names as per the Zip file naming convention.
	/// </summary>
	/// <remarks>The use of absolute names is supported although its use is not valid
	/// according to Zip naming conventions, and should not be used if maximum compatability is desired.</remarks>
	public class ZipNameTransform : INameTransform
	{
		#region Constructors

		/// <summary>
		/// Initialize a new instance of <see cref="ZipNameTransform"></see>
		/// </summary>
		public ZipNameTransform()
		{
		}

		/// <summary>
		/// Initialize a new instance of <see cref="ZipNameTransform"></see>
		/// </summary>
		/// <param name="trimPrefix">The string to trim from the front of paths if found.</param>
		public ZipNameTransform(string trimPrefix)
		{
			TrimPrefix = trimPrefix;
		}

		#endregion Constructors

		/// <summary>
		/// Static constructor.
		/// </summary>
		static ZipNameTransform()
		{
			char[] invalidPathChars;
			invalidPathChars = Path.GetInvalidPathChars();
			int howMany = invalidPathChars.Length + 2;

			InvalidEntryCharsRelaxed = new char[howMany];
			Array.Copy(invalidPathChars, 0, InvalidEntryCharsRelaxed, 0, invalidPathChars.Length);
			InvalidEntryCharsRelaxed[howMany - 1] = '*';
			InvalidEntryCharsRelaxed[howMany - 2] = '?';

			howMany = invalidPathChars.Length + 4;
			InvalidEntryChars = new char[howMany];
			Array.Copy(invalidPathChars, 0, InvalidEntryChars, 0, invalidPathChars.Length);
			InvalidEntryChars[howMany - 1] = ':';
			InvalidEntryChars[howMany - 2] = '\\';
			InvalidEntryChars[howMany - 3] = '*';
			InvalidEntryChars[howMany - 4] = '?';
		}

		/// <summary>
		/// Transform a windows directory name according to the Zip file naming conventions.
		/// </summary>
		/// <param name="name">The directory name to transform.</param>
		/// <returns>The transformed name.</returns>
		public string TransformDirectory(string name)
		{
			name = TransformFile(name);
			if (name.Length > 0)
			{
				if (!name.EndsWith("/", StringComparison.Ordinal))
				{
					name += "/";
				}
			}
			else
			{
				throw new ZipException("Cannot have an empty directory name");
			}
			return name;
		}

		/// <summary>
		/// Transform a windows file name according to the Zip file naming conventions.
		/// </summary>
		/// <param name="name">The file name to transform.</param>
		/// <returns>The transformed name.</returns>
		public string TransformFile(string name)
		{
			if (name != null)
			{
				string lowerName = name.ToLower();
				if ((trimPrefix_ != null) && (lowerName.IndexOf(trimPrefix_, StringComparison.Ordinal) == 0))
				{
					name = name.Substring(trimPrefix_.Length);
				}

				name = name.Replace(@"\", "/");
				name = PathUtils.DropPathRoot(name);

				// Drop any leading and trailing slashes.
				name = name.Trim('/');

				// Convert consecutive // characters to /
				int index = name.IndexOf("//", StringComparison.Ordinal);
				while (index >= 0)
				{
					name = name.Remove(index, 1);
					index = name.IndexOf("//", StringComparison.Ordinal);
				}

				name = MakeValidName(name, '_');
			}
			else
			{
				name = string.Empty;
			}
			return name;
		}

		/// <summary>
		/// Get/set the path prefix to be trimmed from paths if present.
		/// </summary>
		/// <remarks>The prefix is trimmed before any conversion from
		/// a windows path is done.</remarks>
		public string TrimPrefix
		{
			get { return trimPrefix_; }
			set
			{
				trimPrefix_ = value;
				if (trimPrefix_ != null)
				{
					trimPrefix_ = trimPrefix_.ToLower();
				}
			}
		}

		/// <summary>
		/// Force a name to be valid by replacing invalid characters with a fixed value
		/// </summary>
		/// <param name="name">The name to force valid</param>
		/// <param name="replacement">The replacement character to use.</param>
		/// <returns>Returns a valid name</returns>
		private static string MakeValidName(string name, char replacement)
		{
			int index = name.IndexOfAny(InvalidEntryChars);
			if (index >= 0)
			{
				var builder = new StringBuilder(name);

				while (index >= 0)
				{
					builder[index] = replacement;

					if (index >= name.Length)
					{
						index = -1;
					}
					else
					{
						index = name.IndexOfAny(InvalidEntryChars, index + 1);
					}
				}
				name = builder.ToString();
			}

			if (name.Length > 0xffff)
			{
				throw new PathTooLongException();
			}

			return name;
		}

		/// <summary>
		/// Test a name to see if it is a valid name for a zip entry.
		/// </summary>
		/// <param name="name">The name to test.</param>
		/// <param name="relaxed">If true checking is relaxed about windows file names and absolute paths.</param>
		/// <returns>Returns true if the name is a valid zip name; false otherwise.</returns>
		/// <remarks>Zip path names are actually in Unix format, and should only contain relative paths.
		/// This means that any path stored should not contain a drive or
		/// device letter, or a leading slash.  All slashes should forward slashes '/'.
		/// An empty name is valid for a file where the input comes from standard input.
		/// A null name is not considered valid.
		/// </remarks>
		public static bool IsValidName(string name, bool relaxed)
		{
			bool result = (name != null);

			if (result)
			{
				if (relaxed)
				{
					result = name.IndexOfAny(InvalidEntryCharsRelaxed) < 0;
				}
				else
				{
					result =
						(name.IndexOfAny(InvalidEntryChars) < 0) &&
						(name.IndexOf('/') != 0);
				}
			}

			return result;
		}

		/// <summary>
		/// Test a name to see if it is a valid name for a zip entry.
		/// </summary>
		/// <param name="name">The name to test.</param>
		/// <returns>Returns true if the name is a valid zip name; false otherwise.</returns>
		/// <remarks>Zip path names are actually in unix format,
		/// and should only contain relative paths if a path is present.
		/// This means that the path stored should not contain a drive or
		/// device letter, or a leading slash.  All slashes should forward slashes '/'.
		/// An empty name is valid where the input comes from standard input.
		/// A null name is not considered valid.
		/// </remarks>
		public static bool IsValidName(string name)
		{
			bool result =
				(name != null) &&
				(name.IndexOfAny(InvalidEntryChars) < 0) &&
				(name.IndexOf('/') != 0)
				;
			return result;
		}

		#region Instance Fields

		private string trimPrefix_;

		#endregion Instance Fields

		#region Class Fields

		private static readonly char[] InvalidEntryChars;
		private static readonly char[] InvalidEntryCharsRelaxed;

		#endregion Class Fields
	}

	/// <summary>
	/// An implementation of INameTransform that transforms entry paths as per the Zip file naming convention.
	/// Strips path roots and puts directory separators in the correct format ('/')
	/// </summary>
	public class PathTransformer : INameTransform
	{
		/// <summary>
		/// Initialize a new instance of <see cref="PathTransformer"></see>
		/// </summary>
		public PathTransformer()
		{
		}

		/// <summary>
		/// Transform a windows directory name according to the Zip file naming conventions.
		/// </summary>
		/// <param name="name">The directory name to transform.</param>
		/// <returns>The transformed name.</returns>
		public string TransformDirectory(string name)
		{
			name = TransformFile(name);
			
			if (name.Length > 0)
			{
				if (!name.EndsWith("/", StringComparison.Ordinal))
				{
					name += "/";
				}
			}
			else
			{
				throw new ZipException("Cannot have an empty directory name");
			}

			return name;
		}

		/// <summary>
		/// Transform a windows file name according to the Zip file naming conventions.
		/// </summary>
		/// <param name="name">The file name to transform.</param>
		/// <returns>The transformed name.</returns>
		public string TransformFile(string name)
		{
			if (name != null)
			{
				// Put separators in the expected format.
				name = name.Replace(@"\", "/");

				// Remove the path root.
				name = PathUtils.DropPathRoot(name);

				// Drop any leading and trailing slashes.
				name = name.Trim('/');

				// Convert consecutive // characters to /
				int index = name.IndexOf("//", StringComparison.Ordinal);
				while (index >= 0)
				{
					name = name.Remove(index, 1);
					index = name.IndexOf("//", StringComparison.Ordinal);
				}
			}
			else
			{
				name = string.Empty;
			}

			return name;
		}
	}
}


namespace ICSharpCode.SharpZipLib.Zip
{
	/// <summary>
	/// This is a DeflaterOutputStream that writes the files into a zip
	/// archive one after another.  It has a special method to start a new
	/// zip entry.  The zip entries contains information about the file name
	/// size, compressed size, CRC, etc.
	///
	/// It includes support for Stored and Deflated entries.
	/// This class is not thread safe.
	/// <br/>
	/// <br/>Author of the original java version : Jochen Hoenicke
	/// </summary>
	/// <example> This sample shows how to create a zip file
	/// <code>
	///
	/// using ICSharpCode.SharpZipLib.Core;
	/// using ICSharpCode.SharpZipLib.Zip;
	///
	/// class MainClass
	/// {
	/// 	public static void Main(string[] args)
	/// 	{
	/// 		string[] filenames = Directory.GetFiles(args[0]);
	/// 		byte[] buffer = new byte[4096];
	///
	/// 		using ( ZipOutputStream s = new ZipOutputStream(File.Create(args[1])) ) {
	///
	/// 			s.SetLevel(9); // 0 - store only to 9 - means best compression
	///
	/// 			foreach (string file in filenames) {
	/// 				ZipEntry entry = new ZipEntry(file);
	/// 				s.PutNextEntry(entry);
	///
	/// 				using (FileStream fs = File.OpenRead(file)) {
	///						StreamUtils.Copy(fs, s, buffer);
	/// 				}
	/// 			}
	/// 		}
	/// 	}
	/// }
	/// </code>
	/// </example>
	public class ZipOutputStream : DeflaterOutputStream
	{
		#region Constructors

		/// <summary>
		/// Creates a new Zip output stream, writing a zip archive.
		/// </summary>
		/// <param name="baseOutputStream">
		/// The output stream to which the archive contents are written.
		/// </param>
		public ZipOutputStream(Stream baseOutputStream)
			: base(baseOutputStream, new Deflater(Deflater.DEFAULT_COMPRESSION, true))
		{
		}

		/// <summary>
		/// Creates a new Zip output stream, writing a zip archive.
		/// </summary>
		/// <param name="baseOutputStream">The output stream to which the archive contents are written.</param>
		/// <param name="bufferSize">Size of the buffer to use.</param>
		public ZipOutputStream(Stream baseOutputStream, int bufferSize)
			: base(baseOutputStream, new Deflater(Deflater.DEFAULT_COMPRESSION, true), bufferSize)
		{
		}

		#endregion Constructors

		/// <summary>
		/// Gets a flag value of true if the central header has been added for this archive; false if it has not been added.
		/// </summary>
		/// <remarks>No further entries can be added once this has been done.</remarks>
		public bool IsFinished
		{
			get
			{
				return entries == null;
			}
		}

		/// <summary>
		/// Set the zip file comment.
		/// </summary>
		/// <param name="comment">
		/// The comment text for the entire archive.
		/// </param>
		/// <exception cref="ArgumentOutOfRangeException">
		/// The converted comment is longer than 0xffff bytes.
		/// </exception>
		public void SetComment(string comment)
		{
			// TODO: Its not yet clear how to handle unicode comments here.
			byte[] commentBytes = ZipStrings.ConvertToArray(comment);
			if (commentBytes.Length > 0xffff)
			{
				throw new ArgumentOutOfRangeException(nameof(comment));
			}
			zipComment = commentBytes;
		}

		/// <summary>
		/// Sets the compression level.  The new level will be activated
		/// immediately.
		/// </summary>
		/// <param name="level">The new compression level (1 to 9).</param>
		/// <exception cref="ArgumentOutOfRangeException">
		/// Level specified is not supported.
		/// </exception>
		/// <see cref="ICSharpCode.SharpZipLib.Zip.Compression.Deflater"/>
		public void SetLevel(int level)
		{
			deflater_.SetLevel(level);
			defaultCompressionLevel = level;
		}

		/// <summary>
		/// Get the current deflater compression level
		/// </summary>
		/// <returns>The current compression level</returns>
		public int GetLevel()
		{
			return deflater_.GetLevel();
		}

		/// <summary>
		/// Get / set a value indicating how Zip64 Extension usage is determined when adding entries.
		/// </summary>
		/// <remarks>Older archivers may not understand Zip64 extensions.
		/// If backwards compatability is an issue be careful when adding <see cref="ZipEntry.Size">entries</see> to an archive.
		/// Setting this property to off is workable but less desirable as in those circumstances adding a file
		/// larger then 4GB will fail.</remarks>
		public UseZip64 UseZip64
		{
			get { return useZip64_; }
			set { useZip64_ = value; }
		}

		/// <summary>
		/// Used for transforming the names of entries added by <see cref="PutNextEntry(ZipEntry)"/>.
		/// Defaults to <see cref="PathTransformer"/>, set to null to disable transforms and use names as supplied.
		/// </summary>
		public INameTransform NameTransform { get; set; } = new PathTransformer();

		/// <summary>
		/// Get/set the password used for encryption.
		/// </summary>
		/// <remarks>When set to null or if the password is empty no encryption is performed</remarks>
		public string Password
		{
			get
			{
				return password;
			}
			set
			{
				if ((value != null) && (value.Length == 0))
				{
					password = null;
				}
				else
				{
					password = value;
				}
			}
		}

		/// <summary>
		/// Write an unsigned short in little endian byte order.
		/// </summary>
		private void WriteLeShort(int value)
		{
			unchecked
			{
				baseOutputStream_.WriteByte((byte)(value & 0xff));
				baseOutputStream_.WriteByte((byte)((value >> 8) & 0xff));
			}
		}

		/// <summary>
		/// Write an int in little endian byte order.
		/// </summary>
		private void WriteLeInt(int value)
		{
			unchecked
			{
				WriteLeShort(value);
				WriteLeShort(value >> 16);
			}
		}

		/// <summary>
		/// Write an int in little endian byte order.
		/// </summary>
		private void WriteLeLong(long value)
		{
			unchecked
			{
				WriteLeInt((int)value);
				WriteLeInt((int)(value >> 32));
			}
		}

		// Apply any configured transforms/cleaning to the name of the supplied entry.
		private void TransformEntryName(ZipEntry entry)
		{
			if (this.NameTransform != null)
			{
				if (entry.IsDirectory)
				{
					entry.Name = this.NameTransform.TransformDirectory(entry.Name);
				}
				else
				{
					entry.Name = this.NameTransform.TransformFile(entry.Name);
				}
			}
		}

		/// <summary>
		/// Starts a new Zip entry. It automatically closes the previous
		/// entry if present.
		/// All entry elements bar name are optional, but must be correct if present.
		/// If the compression method is stored and the output is not patchable
		/// the compression for that entry is automatically changed to deflate level 0
		/// </summary>
		/// <param name="entry">
		/// the entry.
		/// </param>
		/// <exception cref="System.ArgumentNullException">
		/// if entry passed is null.
		/// </exception>
		/// <exception cref="System.IO.IOException">
		/// if an I/O error occured.
		/// </exception>
		/// <exception cref="System.InvalidOperationException">
		/// if stream was finished
		/// </exception>
		/// <exception cref="ZipException">
		/// Too many entries in the Zip file<br/>
		/// Entry name is too long<br/>
		/// Finish has already been called<br/>
		/// </exception>
		/// <exception cref="System.NotImplementedException">
		/// The Compression method specified for the entry is unsupported.
		/// </exception>
		public void PutNextEntry(ZipEntry entry)
		{
			if (entry == null)
			{
				throw new ArgumentNullException(nameof(entry));
			}

			if (entries == null)
			{
				throw new InvalidOperationException("ZipOutputStream was finished");
			}

			if (curEntry != null)
			{
				CloseEntry();
			}

			if (entries.Count == int.MaxValue)
			{
				throw new ZipException("Too many entries for Zip file");
			}

			CompressionMethod method = entry.CompressionMethod;

			// Check that the compression is one that we support
			if (method != CompressionMethod.Deflated && method != CompressionMethod.Stored)
			{
				throw new NotImplementedException("Compression method not supported");
			}

			// A password must have been set in order to add AES encrypted entries
			if (entry.AESKeySize > 0 && string.IsNullOrEmpty(this.Password))
			{
				throw new InvalidOperationException("The Password property must be set before AES encrypted entries can be added");
			}

			int compressionLevel = defaultCompressionLevel;

			// Clear flags that the library manages internally
			entry.Flags &= (int)GeneralBitFlags.UnicodeText;
			patchEntryHeader = false;

			bool headerInfoAvailable;

			// No need to compress - definitely no data.
			if (entry.Size == 0)
			{
				entry.CompressedSize = entry.Size;
				entry.Crc = 0;
				method = CompressionMethod.Stored;
				headerInfoAvailable = true;
			}
			else
			{
				headerInfoAvailable = (entry.Size >= 0) && entry.HasCrc && entry.CompressedSize >= 0;

				// Switch to deflation if storing isnt possible.
				if (method == CompressionMethod.Stored)
				{
					if (!headerInfoAvailable)
					{
						if (!CanPatchEntries)
						{
							// Can't patch entries so storing is not possible.
							method = CompressionMethod.Deflated;
							compressionLevel = 0;
						}
					}
					else // entry.size must be > 0
					{
						entry.CompressedSize = entry.Size;
						headerInfoAvailable = entry.HasCrc;
					}
				}
			}

			if (headerInfoAvailable == false)
			{
				if (CanPatchEntries == false)
				{
					// Only way to record size and compressed size is to append a data descriptor
					// after compressed data.

					// Stored entries of this form have already been converted to deflating.
					entry.Flags |= 8;
				}
				else
				{
					patchEntryHeader = true;
				}
			}

			if (Password != null)
			{
				entry.IsCrypted = true;
				if (entry.Crc < 0)
				{
					// Need to append a data descriptor as the crc isnt available for use
					// with encryption, the date is used instead.  Setting the flag
					// indicates this to the decompressor.
					entry.Flags |= 8;
				}
			}

			entry.Offset = offset;
			entry.CompressionMethod = (CompressionMethod)method;

			curMethod = method;
			sizePatchPos = -1;

			if ((useZip64_ == UseZip64.On) || ((entry.Size < 0) && (useZip64_ == UseZip64.Dynamic)))
			{
				entry.ForceZip64();
			}

			// Write the local file header
			WriteLeInt(ZipConstants.LocalHeaderSignature);

			WriteLeShort(entry.Version);
			WriteLeShort(entry.Flags);
			WriteLeShort((byte)entry.CompressionMethodForHeader);
			WriteLeInt((int)entry.DosTime);

			// TODO: Refactor header writing.  Its done in several places.
			if (headerInfoAvailable)
			{
				WriteLeInt((int)entry.Crc);
				if (entry.LocalHeaderRequiresZip64)
				{
					WriteLeInt(-1);
					WriteLeInt(-1);
				}
				else
				{
					WriteLeInt((int)entry.CompressedSize + entry.EncryptionOverheadSize);
					WriteLeInt((int)entry.Size);
				}
			}
			else
			{
				if (patchEntryHeader)
				{
					crcPatchPos = baseOutputStream_.Position;
				}
				WriteLeInt(0);  // Crc

				if (patchEntryHeader)
				{
					sizePatchPos = baseOutputStream_.Position;
				}

				// For local header both sizes appear in Zip64 Extended Information
				if (entry.LocalHeaderRequiresZip64 || patchEntryHeader)
				{
					WriteLeInt(-1);
					WriteLeInt(-1);
				}
				else
				{
					WriteLeInt(0);  // Compressed size
					WriteLeInt(0);  // Uncompressed size
				}
			}

			// Apply any required transforms to the entry name, and then convert to byte array format.
			TransformEntryName(entry);
			byte[] name = ZipStrings.ConvertToArray(entry.Flags, entry.Name);

			if (name.Length > 0xFFFF)
			{
				throw new ZipException("Entry name too long.");
			}

			var ed = new ZipExtraData(entry.ExtraData);

			if (entry.LocalHeaderRequiresZip64)
			{
				ed.StartNewEntry();
				if (headerInfoAvailable)
				{
					ed.AddLeLong(entry.Size);
					ed.AddLeLong(entry.CompressedSize + entry.EncryptionOverheadSize);
				}
				else
				{
					ed.AddLeLong(-1);
					ed.AddLeLong(-1);
				}
				ed.AddNewEntry(1);

				if (!ed.Find(1))
				{
					throw new ZipException("Internal error cant find extra data");
				}

				if (patchEntryHeader)
				{
					sizePatchPos = ed.CurrentReadIndex;
				}
			}
			else
			{
				ed.Delete(1);
			}

			if (entry.AESKeySize > 0)
			{
				AddExtraDataAES(entry, ed);
			}
			byte[] extra = ed.GetEntryData();

			WriteLeShort(name.Length);
			WriteLeShort(extra.Length);

			if (name.Length > 0)
			{
				baseOutputStream_.Write(name, 0, name.Length);
			}

			if (entry.LocalHeaderRequiresZip64 && patchEntryHeader)
			{
				sizePatchPos += baseOutputStream_.Position;
			}

			if (extra.Length > 0)
			{
				baseOutputStream_.Write(extra, 0, extra.Length);
			}

			offset += ZipConstants.LocalHeaderBaseSize + name.Length + extra.Length;
			// Fix offsetOfCentraldir for AES
			if (entry.AESKeySize > 0)
				offset += entry.AESOverheadSize;

			// Activate the entry.
			curEntry = entry;
			crc.Reset();
			if (method == CompressionMethod.Deflated)
			{
				deflater_.Reset();
				deflater_.SetLevel(compressionLevel);
			}
			size = 0;

			if (entry.IsCrypted)
			{
				if (entry.AESKeySize > 0)
				{
					WriteAESHeader(entry);
				}
				else
				{
					if (entry.Crc < 0)
					{            // so testing Zip will says its ok
						WriteEncryptionHeader(entry.DosTime << 16);
					}
					else
					{
						WriteEncryptionHeader(entry.Crc);
					}
				}
			}
		}

		/// <summary>
		/// Closes the current entry, updating header and footer information as required
		/// </summary>
		/// <exception cref="ZipException">
		/// Invalid entry field values.
		/// </exception>
		/// <exception cref="System.IO.IOException">
		/// An I/O error occurs.
		/// </exception>
		/// <exception cref="System.InvalidOperationException">
		/// No entry is active.
		/// </exception>
		public void CloseEntry()
		{
			if (curEntry == null)
			{
				throw new InvalidOperationException("No open entry");
			}

			long csize = size;

			// First finish the deflater, if appropriate
			if (curMethod == CompressionMethod.Deflated)
			{
				if (size >= 0)
				{
					base.Finish();
					csize = deflater_.TotalOut;
				}
				else
				{
					deflater_.Reset();
				}
			}
			else if (curMethod == CompressionMethod.Stored)
			{
				// This is done by Finish() for Deflated entries, but we need to do it
				// ourselves for Stored ones
				base.GetAuthCodeIfAES();
			}

			// Write the AES Authentication Code (a hash of the compressed and encrypted data)
			if (curEntry.AESKeySize > 0)
			{
				baseOutputStream_.Write(AESAuthCode, 0, 10);
				// Always use 0 as CRC for AE-2 format
				curEntry.Crc = 0;
			}
			else
			{
				if (curEntry.Crc < 0)
				{
					curEntry.Crc = crc.Value;
				}
				else if (curEntry.Crc != crc.Value)
				{
					throw new ZipException($"crc was {crc.Value}, but {curEntry.Crc} was expected");
				}
			}

			if (curEntry.Size < 0)
			{
				curEntry.Size = size;
			}
			else if (curEntry.Size != size)
			{
				throw new ZipException($"size was {size}, but {curEntry.Size} was expected");
			}

			if (curEntry.CompressedSize < 0)
			{
				curEntry.CompressedSize = csize;
			}
			else if (curEntry.CompressedSize != csize)
			{
				throw new ZipException($"compressed size was {csize}, but {curEntry.CompressedSize} expected");
			}

			offset += csize;

			if (curEntry.IsCrypted)
			{
				curEntry.CompressedSize += curEntry.EncryptionOverheadSize;
			}

			// Patch the header if possible
			if (patchEntryHeader)
			{
				patchEntryHeader = false;

				long curPos = baseOutputStream_.Position;
				baseOutputStream_.Seek(crcPatchPos, SeekOrigin.Begin);
				WriteLeInt((int)curEntry.Crc);

				if (curEntry.LocalHeaderRequiresZip64)
				{
					if (sizePatchPos == -1)
					{
						throw new ZipException("Entry requires zip64 but this has been turned off");
					}

					baseOutputStream_.Seek(sizePatchPos, SeekOrigin.Begin);
					WriteLeLong(curEntry.Size);
					WriteLeLong(curEntry.CompressedSize);
				}
				else
				{
					WriteLeInt((int)curEntry.CompressedSize);
					WriteLeInt((int)curEntry.Size);
				}
				baseOutputStream_.Seek(curPos, SeekOrigin.Begin);
			}

			// Add data descriptor if flagged as required
			if ((curEntry.Flags & 8) != 0)
			{
				WriteLeInt(ZipConstants.DataDescriptorSignature);
				WriteLeInt(unchecked((int)curEntry.Crc));

				if (curEntry.LocalHeaderRequiresZip64)
				{
					WriteLeLong(curEntry.CompressedSize);
					WriteLeLong(curEntry.Size);
					offset += ZipConstants.Zip64DataDescriptorSize;
				}
				else
				{
					WriteLeInt((int)curEntry.CompressedSize);
					WriteLeInt((int)curEntry.Size);
					offset += ZipConstants.DataDescriptorSize;
				}
			}

			entries.Add(curEntry);
			curEntry = null;
		}

		/// <summary>
		/// Initializes encryption keys based on given <paramref name="password"/>.
		/// </summary>
		/// <param name="password">The password.</param>
		private void InitializePassword(string password)
		{
			var pkManaged = new PkzipClassicManaged();
			byte[] key = PkzipClassic.GenerateKeys(ZipStrings.ConvertToArray(password));
			cryptoTransform_ = pkManaged.CreateEncryptor(key, null);
		}

		/// <summary>
		/// Initializes encryption keys based on given password.
		/// </summary>
		private void InitializeAESPassword(ZipEntry entry, string rawPassword,
											out byte[] salt, out byte[] pwdVerifier)
		{
			salt = new byte[entry.AESSaltLen];

			// Salt needs to be cryptographically random, and unique per file
			_aesRnd.GetBytes(salt);

			int blockSize = entry.AESKeySize / 8;   // bits to bytes

			cryptoTransform_ = new ZipAESTransform(rawPassword, salt, blockSize, true);
			pwdVerifier = ((ZipAESTransform)cryptoTransform_).PwdVerifier;
		}

		private void WriteEncryptionHeader(long crcValue)
		{
			offset += ZipConstants.CryptoHeaderSize;

			InitializePassword(Password);

			byte[] cryptBuffer = new byte[ZipConstants.CryptoHeaderSize];
			using (var rng = new RNGCryptoServiceProvider())
			{
				rng.GetBytes(cryptBuffer);
			}

			cryptBuffer[11] = (byte)(crcValue >> 24);

			EncryptBlock(cryptBuffer, 0, cryptBuffer.Length);
			baseOutputStream_.Write(cryptBuffer, 0, cryptBuffer.Length);
		}

		private static void AddExtraDataAES(ZipEntry entry, ZipExtraData extraData)
		{
			// Vendor Version: AE-1 IS 1. AE-2 is 2. With AE-2 no CRC is required and 0 is stored.
			const int VENDOR_VERSION = 2;
			// Vendor ID is the two ASCII characters "AE".
			const int VENDOR_ID = 0x4541; //not 6965;
			extraData.StartNewEntry();
			// Pack AES extra data field see http://www.winzip.com/aes_info.htm
			//extraData.AddLeShort(7);							// Data size (currently 7)
			extraData.AddLeShort(VENDOR_VERSION);               // 2 = AE-2
			extraData.AddLeShort(VENDOR_ID);                    // "AE"
			extraData.AddData(entry.AESEncryptionStrength);     //  1 = 128, 2 = 192, 3 = 256
			extraData.AddLeShort((int)entry.CompressionMethod); // The actual compression method used to compress the file
			extraData.AddNewEntry(0x9901);
		}

		// Replaces WriteEncryptionHeader for AES
		//
		private void WriteAESHeader(ZipEntry entry)
		{
			byte[] salt;
			byte[] pwdVerifier;
			InitializeAESPassword(entry, Password, out salt, out pwdVerifier);
			// File format for AES:
			// Size (bytes)   Content
			// ------------   -------
			// Variable       Salt value
			// 2              Password verification value
			// Variable       Encrypted file data
			// 10             Authentication code
			//
			// Value in the "compressed size" fields of the local file header and the central directory entry
			// is the total size of all the items listed above. In other words, it is the total size of the
			// salt value, password verification value, encrypted data, and authentication code.
			baseOutputStream_.Write(salt, 0, salt.Length);
			baseOutputStream_.Write(pwdVerifier, 0, pwdVerifier.Length);
		}

		/// <summary>
		/// Writes the given buffer to the current entry.
		/// </summary>
		/// <param name="buffer">The buffer containing data to write.</param>
		/// <param name="offset">The offset of the first byte to write.</param>
		/// <param name="count">The number of bytes to write.</param>
		/// <exception cref="ZipException">Archive size is invalid</exception>
		/// <exception cref="System.InvalidOperationException">No entry is active.</exception>
		public override void Write(byte[] buffer, int offset, int count)
		{
			if (curEntry == null)
			{
				throw new InvalidOperationException("No open entry.");
			}

			if (buffer == null)
			{
				throw new ArgumentNullException(nameof(buffer));
			}

			if (offset < 0)
			{
				throw new ArgumentOutOfRangeException(nameof(offset), "Cannot be negative");
			}

			if (count < 0)
			{
				throw new ArgumentOutOfRangeException(nameof(count), "Cannot be negative");
			}

			if ((buffer.Length - offset) < count)
			{
				throw new ArgumentException("Invalid offset/count combination");
			}

			if (curEntry.AESKeySize == 0)
			{
				// Only update CRC if AES is not enabled
				crc.Update(new ArraySegment<byte>(buffer, offset, count));
			}

			size += count;

			switch (curMethod)
			{
				case CompressionMethod.Deflated:
					base.Write(buffer, offset, count);
					break;

				case CompressionMethod.Stored:
					if (Password != null)
					{
						CopyAndEncrypt(buffer, offset, count);
					}
					else
					{
						baseOutputStream_.Write(buffer, offset, count);
					}
					break;
			}
		}

		private void CopyAndEncrypt(byte[] buffer, int offset, int count)
		{
			const int CopyBufferSize = 4096;
			byte[] localBuffer = new byte[CopyBufferSize];
			while (count > 0)
			{
				int bufferCount = (count < CopyBufferSize) ? count : CopyBufferSize;

				Array.Copy(buffer, offset, localBuffer, 0, bufferCount);
				EncryptBlock(localBuffer, 0, bufferCount);
				baseOutputStream_.Write(localBuffer, 0, bufferCount);
				count -= bufferCount;
				offset += bufferCount;
			}
		}

		/// <summary>
		/// Finishes the stream.  This will write the central directory at the
		/// end of the zip file and flush the stream.
		/// </summary>
		/// <remarks>
		/// This is automatically called when the stream is closed.
		/// </remarks>
		/// <exception cref="System.IO.IOException">
		/// An I/O error occurs.
		/// </exception>
		/// <exception cref="ZipException">
		/// Comment exceeds the maximum length<br/>
		/// Entry name exceeds the maximum length
		/// </exception>
		public override void Finish()
		{
			if (entries == null)
			{
				return;
			}

			if (curEntry != null)
			{
				CloseEntry();
			}

			long numEntries = entries.Count;
			long sizeEntries = 0;

			foreach (ZipEntry entry in entries)
			{
				WriteLeInt(ZipConstants.CentralHeaderSignature);
				WriteLeShort((entry.HostSystem << 8) | entry.VersionMadeBy);
				WriteLeShort(entry.Version);
				WriteLeShort(entry.Flags);
				WriteLeShort((short)entry.CompressionMethodForHeader);
				WriteLeInt((int)entry.DosTime);
				WriteLeInt((int)entry.Crc);

				if (entry.IsZip64Forced() ||
					(entry.CompressedSize >= uint.MaxValue))
				{
					WriteLeInt(-1);
				}
				else
				{
					WriteLeInt((int)entry.CompressedSize);
				}

				if (entry.IsZip64Forced() ||
					(entry.Size >= uint.MaxValue))
				{
					WriteLeInt(-1);
				}
				else
				{
					WriteLeInt((int)entry.Size);
				}

				byte[] name = ZipStrings.ConvertToArray(entry.Flags, entry.Name);

				if (name.Length > 0xffff)
				{
					throw new ZipException("Name too long.");
				}

				var ed = new ZipExtraData(entry.ExtraData);

				if (entry.CentralHeaderRequiresZip64)
				{
					ed.StartNewEntry();
					if (entry.IsZip64Forced() ||
						(entry.Size >= 0xffffffff))
					{
						ed.AddLeLong(entry.Size);
					}

					if (entry.IsZip64Forced() ||
						(entry.CompressedSize >= 0xffffffff))
					{
						ed.AddLeLong(entry.CompressedSize);
					}

					if (entry.Offset >= 0xffffffff)
					{
						ed.AddLeLong(entry.Offset);
					}

					ed.AddNewEntry(1);
				}
				else
				{
					ed.Delete(1);
				}

				if (entry.AESKeySize > 0)
				{
					AddExtraDataAES(entry, ed);
				}
				byte[] extra = ed.GetEntryData();

				byte[] entryComment =
					(entry.Comment != null) ?
					ZipStrings.ConvertToArray(entry.Flags, entry.Comment) :
					Empty.Array<byte>();

				if (entryComment.Length > 0xffff)
				{
					throw new ZipException("Comment too long.");
				}

				WriteLeShort(name.Length);
				WriteLeShort(extra.Length);
				WriteLeShort(entryComment.Length);
				WriteLeShort(0);    // disk number
				WriteLeShort(0);    // internal file attributes
									// external file attributes

				if (entry.ExternalFileAttributes != -1)
				{
					WriteLeInt(entry.ExternalFileAttributes);
				}
				else
				{
					if (entry.IsDirectory)
					{                         // mark entry as directory (from nikolam.AT.perfectinfo.com)
						WriteLeInt(16);
					}
					else
					{
						WriteLeInt(0);
					}
				}

				if (entry.Offset >= uint.MaxValue)
				{
					WriteLeInt(-1);
				}
				else
				{
					WriteLeInt((int)entry.Offset);
				}

				if (name.Length > 0)
				{
					baseOutputStream_.Write(name, 0, name.Length);
				}

				if (extra.Length > 0)
				{
					baseOutputStream_.Write(extra, 0, extra.Length);
				}

				if (entryComment.Length > 0)
				{
					baseOutputStream_.Write(entryComment, 0, entryComment.Length);
				}

				sizeEntries += ZipConstants.CentralHeaderBaseSize + name.Length + extra.Length + entryComment.Length;
			}

			using (ZipHelperStream zhs = new ZipHelperStream(baseOutputStream_))
			{
				zhs.WriteEndOfCentralDirectory(numEntries, sizeEntries, offset, zipComment);
			}

			entries = null;
		}

		/// <summary>
		/// Flushes the stream by calling <see cref="DeflaterOutputStream.Flush">Flush</see> on the deflater stream unless
		/// the current compression method is <see cref="CompressionMethod.Stored"/>. Then it flushes the underlying output stream.
		/// </summary>
		public override void Flush()
		{
			if(curMethod == CompressionMethod.Stored)
			{
				baseOutputStream_.Flush();
			} 
			else
			{
				base.Flush();
			}
		}

		#region Instance Fields

		/// <summary>
		/// The entries for the archive.
		/// </summary>
		private List<ZipEntry> entries = new List<ZipEntry>();

		/// <summary>
		/// Used to track the crc of data added to entries.
		/// </summary>
		private Crc32 crc = new Crc32();

		/// <summary>
		/// The current entry being added.
		/// </summary>
		private ZipEntry curEntry;

		private int defaultCompressionLevel = Deflater.DEFAULT_COMPRESSION;

		private CompressionMethod curMethod = CompressionMethod.Deflated;

		/// <summary>
		/// Used to track the size of data for an entry during writing.
		/// </summary>
		private long size;

		/// <summary>
		/// Offset to be recorded for each entry in the central header.
		/// </summary>
		private long offset;

		/// <summary>
		/// Comment for the entire archive recorded in central header.
		/// </summary>
		private byte[] zipComment = Empty.Array<byte>();

		/// <summary>
		/// Flag indicating that header patching is required for the current entry.
		/// </summary>
		private bool patchEntryHeader;

		/// <summary>
		/// Position to patch crc
		/// </summary>
		private long crcPatchPos = -1;

		/// <summary>
		/// Position to patch size.
		/// </summary>
		private long sizePatchPos = -1;

		// Default is dynamic which is not backwards compatible and can cause problems
		// with XP's built in compression which cant read Zip64 archives.
		// However it does avoid the situation were a large file is added and cannot be completed correctly.
		// NOTE: Setting the size for entries before they are added is the best solution!
		private UseZip64 useZip64_ = UseZip64.Dynamic;

		/// <summary>
		/// The password to use when encrypting archive entries.
		/// </summary>
		private string password;

		#endregion Instance Fields

		#region Static Fields

		// Static to help ensure that multiple files within a zip will get different random salt
		private static RandomNumberGenerator _aesRnd = RandomNumberGenerator.Create();

		#endregion Static Fields
	}
}


namespace ICSharpCode.SharpZipLib.Zip
{
	/// <summary>
	/// This static class contains functions for encoding and decoding zip file strings
	/// </summary>
	public static class ZipStrings
	{
		static ZipStrings()
		{
			try
			{
				var platformCodepage = Encoding.GetEncoding(0).CodePage;
				SystemDefaultCodePage = (platformCodepage == 1 || platformCodepage == 2 || platformCodepage == 3 || platformCodepage == 42) ? FallbackCodePage : platformCodepage;
			}
			catch
			{
				SystemDefaultCodePage = FallbackCodePage;
			}
		}

		/// <summary>Code page backing field</summary>
		/// <remarks>
		/// The original Zip specification (https://pkware.cachefly.net/webdocs/casestudies/APPNOTE.TXT) states
		/// that file names should only be encoded with IBM Code Page 437 or UTF-8.
		/// In practice, most zip apps use OEM or system encoding (typically cp437 on Windows).
		/// Let's be good citizens and default to UTF-8 http://utf8everywhere.org/
		/// </remarks>
		private static int codePage = AutomaticCodePage;

		/// Automatically select codepage while opening archive
		/// see https://github.com/icsharpcode/SharpZipLib/pull/280#issuecomment-433608324
		/// 
		private const int AutomaticCodePage = -1;

		/// <summary>
		/// Encoding used for string conversion. Setting this to 65001 (UTF-8) will
		/// also set the Language encoding flag to indicate UTF-8 encoded file names.
		/// </summary>
		public static int CodePage
		{
			get
			{
				return codePage == AutomaticCodePage? Encoding.UTF8.CodePage:codePage;
			}
			set
			{
				if ((value < 0) || (value > 65535) ||
					(value == 1) || (value == 2) || (value == 3) || (value == 42))
				{
					throw new ArgumentOutOfRangeException(nameof(value));
				}

				codePage = value;
			}
		}

		private const int FallbackCodePage = 437;

		/// <summary>
		/// Attempt to get the operating system default codepage, or failing that, to
		/// the fallback code page IBM 437.
		/// </summary>
		public static int SystemDefaultCodePage { get; }

		/// <summary>
		/// Get whether the default codepage is set to UTF-8. Setting this property to false will
		/// set the <see cref="CodePage"/> to <see cref="SystemDefaultCodePage"/>
		/// </summary>
		/// <remarks>
		/// Get OEM codepage from NetFX, which parses the NLP file with culture info table etc etc.
		/// But sometimes it yields the special value of 1 which is nicknamed <c>CodePageNoOEM</c> in <see cref="Encoding"/> sources (might also mean <c>CP_OEMCP</c>, but Encoding puts it so).
		/// This was observed on Ukranian and Hindu systems.
		/// Given this value, <see cref="Encoding.GetEncoding(int)"/> throws an <see cref="ArgumentException"/>.
		/// So replace it with <see cref="FallbackCodePage"/>, (IBM 437 which is the default code page in a default Windows installation console.
		/// </remarks>
		public static bool UseUnicode
		{
			get
			{
				return codePage == Encoding.UTF8.CodePage;
			}
			set
			{
				if (value)
				{
					codePage = Encoding.UTF8.CodePage;
				}
				else
				{
					codePage = SystemDefaultCodePage;
				}
			}
		}

		/// <summary>
		/// Convert a portion of a byte array to a string using <see cref="CodePage"/>
		/// </summary>
		/// <param name="data">
		/// Data to convert to string
		/// </param>
		/// <param name="count">
		/// Number of bytes to convert starting from index 0
		/// </param>
		/// <returns>
		/// data[0]..data[count - 1] converted to a string
		/// </returns>
		public static string ConvertToString(byte[] data, int count)
			=> data == null
			? string.Empty
			: Encoding.GetEncoding(CodePage).GetString(data, 0, count);

		/// <summary>
		/// Convert a byte array to a string using <see cref="CodePage"/>
		/// </summary>
		/// <param name="data">
		/// Byte array to convert
		/// </param>
		/// <returns>
		/// <paramref name="data">data</paramref>converted to a string
		/// </returns>
		public static string ConvertToString(byte[] data)
			=> ConvertToString(data, data.Length);

		private static Encoding EncodingFromFlag(int flags)
			=> ((flags & (int)GeneralBitFlags.UnicodeText) != 0)
				? Encoding.UTF8
				: Encoding.GetEncoding(
					// if CodePage wasn't set manually and no utf flag present
					// then we must use SystemDefault (old behavior)
					// otherwise, CodePage should be preferred over SystemDefault
					// see https://github.com/icsharpcode/SharpZipLib/issues/274
					codePage == AutomaticCodePage? 
						SystemDefaultCodePage:
						codePage);

		/// <summary>
		/// Convert a byte array to a string  using <see cref="CodePage"/>
		/// </summary>
		/// <param name="flags">The applicable general purpose bits flags</param>
		/// <param name="data">
		/// Byte array to convert
		/// </param>
		/// <param name="count">The number of bytes to convert.</param>
		/// <returns>
		/// <paramref name="data">data</paramref>converted to a string
		/// </returns>
		public static string ConvertToStringExt(int flags, byte[] data, int count)
			=> (data == null)
				? string.Empty
				: EncodingFromFlag(flags).GetString(data, 0, count);

		/// <summary>
		/// Convert a byte array to a string using <see cref="CodePage"/>
		/// </summary>
		/// <param name="data">
		/// Byte array to convert
		/// </param>
		/// <param name="flags">The applicable general purpose bits flags</param>
		/// <returns>
		/// <paramref name="data">data</paramref>converted to a string
		/// </returns>
		public static string ConvertToStringExt(int flags, byte[] data)
			=> ConvertToStringExt(flags, data, data.Length);

		/// <summary>
		/// Convert a string to a byte array using <see cref="CodePage"/>
		/// </summary>
		/// <param name="str">
		/// String to convert to an array
		/// </param>
		/// <returns>Converted array</returns>
		public static byte[] ConvertToArray(string str)
			=> str == null
			? Empty.Array<byte>()
			: Encoding.GetEncoding(CodePage).GetBytes(str);

		/// <summary>
		/// Convert a string to a byte array using <see cref="CodePage"/>
		/// </summary>
		/// <param name="flags">The applicable <see cref="GeneralBitFlags">general purpose bits flags</see></param>
		/// <param name="str">
		/// String to convert to an array
		/// </param>
		/// <returns>Converted array</returns>
		public static byte[] ConvertToArray(int flags, string str)
			=> (string.IsNullOrEmpty(str))
				? Empty.Array<byte>()
				: EncodingFromFlag(flags).GetBytes(str);
	}
}


namespace ICSharpCode.SharpZipLib
{
	/// <summary>
	/// SharpZipBaseException is the base exception class for SharpZipLib.
	/// All library exceptions are derived from this.
	/// </summary>
	/// <remarks>NOTE: Not all exceptions thrown will be derived from this class.
	/// A variety of other exceptions are possible for example <see cref="ArgumentNullException"></see></remarks>
	[Serializable]
	public class SharpZipBaseException : Exception
	{
		/// <summary>
		/// Initializes a new instance of the SharpZipBaseException class.
		/// </summary>
		public SharpZipBaseException()
		{
		}

		/// <summary>
		/// Initializes a new instance of the SharpZipBaseException class with a specified error message.
		/// </summary>
		/// <param name="message">A message describing the exception.</param>
		public SharpZipBaseException(string message)
			: base(message)
		{
		}

		/// <summary>
		/// Initializes a new instance of the SharpZipBaseException class with a specified
		/// error message and a reference to the inner exception that is the cause of this exception.
		/// </summary>
		/// <param name="message">A message describing the exception.</param>
		/// <param name="innerException">The inner exception</param>
		public SharpZipBaseException(string message, Exception innerException)
			: base(message, innerException)
		{
		}

		/// <summary>
		/// Initializes a new instance of the SharpZipBaseException class with serialized data.
		/// </summary>
		/// <param name="info">
		/// The System.Runtime.Serialization.SerializationInfo that holds the serialized
		/// object data about the exception being thrown.
		/// </param>
		/// <param name="context">
		/// The System.Runtime.Serialization.StreamingContext that contains contextual information
		/// about the source or destination.
		/// </param>
		protected SharpZipBaseException(SerializationInfo info, StreamingContext context)
			: base(info, context)
		{
		}
	}
}


namespace ICSharpCode.SharpZipLib
{
	/// <summary>
	/// Indicates that an error occurred during decoding of a input stream due to corrupt
	/// data or (unintentional) library incompatibility.
	/// </summary>
	[Serializable]
	public class StreamDecodingException : SharpZipBaseException
	{
		private const string GenericMessage = "Input stream could not be decoded";

		/// <summary>
		/// Initializes a new instance of the StreamDecodingException with a generic message
		/// </summary>
		public StreamDecodingException() : base(GenericMessage) { }

		/// <summary>
		/// Initializes a new instance of the StreamDecodingException class with a specified error message.
		/// </summary>
		/// <param name="message">A message describing the exception.</param>
		public StreamDecodingException(string message) : base(message) { }

		/// <summary>
		/// Initializes a new instance of the StreamDecodingException class with a specified
		/// error message and a reference to the inner exception that is the cause of this exception.
		/// </summary>
		/// <param name="message">A message describing the exception.</param>
		/// <param name="innerException">The inner exception</param>
		public StreamDecodingException(string message, Exception innerException) : base(message, innerException) { }

		/// <summary>
		/// Initializes a new instance of the StreamDecodingException class with serialized data.
		/// </summary>
		/// <param name="info">
		/// The System.Runtime.Serialization.SerializationInfo that holds the serialized
		/// object data about the exception being thrown.
		/// </param>
		/// <param name="context">
		/// The System.Runtime.Serialization.StreamingContext that contains contextual information
		/// about the source or destination.
		/// </param>
		protected StreamDecodingException(SerializationInfo info, StreamingContext context)
			: base(info, context)
		{
		}
	}
}


namespace ICSharpCode.SharpZipLib
{
	/// <summary>
	/// Indicates that the input stream could not decoded due to known library incompability or missing features
	/// </summary>
	[Serializable]
	public class StreamUnsupportedException : StreamDecodingException
	{
		private const string GenericMessage = "Input stream is in a unsupported format";

		/// <summary>
		/// Initializes a new instance of the StreamUnsupportedException with a generic message
		/// </summary>
		public StreamUnsupportedException() : base(GenericMessage) { }

		/// <summary>
		/// Initializes a new instance of the StreamUnsupportedException class with a specified error message.
		/// </summary>
		/// <param name="message">A message describing the exception.</param>
		public StreamUnsupportedException(string message) : base(message) { }

		/// <summary>
		/// Initializes a new instance of the StreamUnsupportedException class with a specified
		/// error message and a reference to the inner exception that is the cause of this exception.
		/// </summary>
		/// <param name="message">A message describing the exception.</param>
		/// <param name="innerException">The inner exception</param>
		public StreamUnsupportedException(string message, Exception innerException) : base(message, innerException) { }

		/// <summary>
		/// Initializes a new instance of the StreamUnsupportedException class with serialized data.
		/// </summary>
		/// <param name="info">
		/// The System.Runtime.Serialization.SerializationInfo that holds the serialized
		/// object data about the exception being thrown.
		/// </param>
		/// <param name="context">
		/// The System.Runtime.Serialization.StreamingContext that contains contextual information
		/// about the source or destination.
		/// </param>
		protected StreamUnsupportedException(SerializationInfo info, StreamingContext context)
			: base(info, context)
		{
		}
	}
}


namespace ICSharpCode.SharpZipLib
{
	/// <summary>
	/// Indicates that the input stream could not decoded due to the stream ending before enough data had been provided
	/// </summary>
	[Serializable]
	public class UnexpectedEndOfStreamException : StreamDecodingException
	{
		private const string GenericMessage = "Input stream ended unexpectedly";

		/// <summary>
		/// Initializes a new instance of the UnexpectedEndOfStreamException with a generic message
		/// </summary>
		public UnexpectedEndOfStreamException() : base(GenericMessage) { }

		/// <summary>
		/// Initializes a new instance of the UnexpectedEndOfStreamException class with a specified error message.
		/// </summary>
		/// <param name="message">A message describing the exception.</param>
		public UnexpectedEndOfStreamException(string message) : base(message) { }

		/// <summary>
		/// Initializes a new instance of the UnexpectedEndOfStreamException class with a specified
		/// error message and a reference to the inner exception that is the cause of this exception.
		/// </summary>
		/// <param name="message">A message describing the exception.</param>
		/// <param name="innerException">The inner exception</param>
		public UnexpectedEndOfStreamException(string message, Exception innerException) : base(message, innerException) { }

		/// <summary>
		/// Initializes a new instance of the UnexpectedEndOfStreamException class with serialized data.
		/// </summary>
		/// <param name="info">
		/// The System.Runtime.Serialization.SerializationInfo that holds the serialized
		/// object data about the exception being thrown.
		/// </param>
		/// <param name="context">
		/// The System.Runtime.Serialization.StreamingContext that contains contextual information
		/// about the source or destination.
		/// </param>
		protected UnexpectedEndOfStreamException(SerializationInfo info, StreamingContext context)
			: base(info, context)
		{
		}
	}
}


namespace ICSharpCode.SharpZipLib
{
	/// <summary>
	/// Indicates that a value was outside of the expected range when decoding an input stream
	/// </summary>
	[Serializable]
	public class ValueOutOfRangeException : StreamDecodingException
	{
		/// <summary>
		/// Initializes a new instance of the ValueOutOfRangeException class naming the causing variable
		/// </summary>
		/// <param name="nameOfValue">Name of the variable, use: nameof()</param>
		public ValueOutOfRangeException(string nameOfValue)
			: base($"{nameOfValue} out of range") { }

		/// <summary>
		/// Initializes a new instance of the ValueOutOfRangeException class naming the causing variable,
		/// it's current value and expected range.
		/// </summary>
		/// <param name="nameOfValue">Name of the variable, use: nameof()</param>
		/// <param name="value">The invalid value</param>
		/// <param name="maxValue">Expected maximum value</param>
		/// <param name="minValue">Expected minimum value</param>
		public ValueOutOfRangeException(string nameOfValue, long value, long maxValue, long minValue = 0)
			: this(nameOfValue, value.ToString(), maxValue.ToString(), minValue.ToString()) { }

		/// <summary>
		/// Initializes a new instance of the ValueOutOfRangeException class naming the causing variable,
		/// it's current value and expected range.
		/// </summary>
		/// <param name="nameOfValue">Name of the variable, use: nameof()</param>
		/// <param name="value">The invalid value</param>
		/// <param name="maxValue">Expected maximum value</param>
		/// <param name="minValue">Expected minimum value</param>
		public ValueOutOfRangeException(string nameOfValue, string value, string maxValue, string minValue = "0") :
			base($"{nameOfValue} out of range: {value}, should be {minValue}..{maxValue}")
		{ }

		private ValueOutOfRangeException()
		{
		}

		private ValueOutOfRangeException(string message, Exception innerException) : base(message, innerException)
		{
		}

		/// <summary>
		/// Initializes a new instance of the ValueOutOfRangeException class with serialized data.
		/// </summary>
		/// <param name="info">
		/// The System.Runtime.Serialization.SerializationInfo that holds the serialized
		/// object data about the exception being thrown.
		/// </param>
		/// <param name="context">
		/// The System.Runtime.Serialization.StreamingContext that contains contextual information
		/// about the source or destination.
		/// </param>
		protected ValueOutOfRangeException(SerializationInfo info, StreamingContext context)
			: base(info, context)
		{
		}
	}
}


namespace ICSharpCode.SharpZipLib.Zip.Compression
{
	/// <summary>
	/// This is the Deflater class.  The deflater class compresses input
	/// with the deflate algorithm described in RFC 1951.  It has several
	/// compression levels and three different strategies described below.
	///
	/// This class is <i>not</i> thread safe.  This is inherent in the API, due
	/// to the split of deflate and setInput.
	///
	/// author of the original java version : Jochen Hoenicke
	/// </summary>
	public class Deflater
	{
		#region Deflater Documentation

		/*
		* The Deflater can do the following state transitions:
		*
		* (1) -> INIT_STATE   ----> INIT_FINISHING_STATE ---.
		*        /  | (2)      (5)                          |
		*       /   v          (5)                          |
		*   (3)| SETDICT_STATE ---> SETDICT_FINISHING_STATE |(3)
		*       \   | (3)                 |        ,--------'
		*        |  |                     | (3)   /
		*        v  v          (5)        v      v
		* (1) -> BUSY_STATE   ----> FINISHING_STATE
		*                                | (6)
		*                                v
		*                           FINISHED_STATE
		*    \_____________________________________/
		*                    | (7)
		*                    v
		*               CLOSED_STATE
		*
		* (1) If we should produce a header we start in INIT_STATE, otherwise
		*     we start in BUSY_STATE.
		* (2) A dictionary may be set only when we are in INIT_STATE, then
		*     we change the state as indicated.
		* (3) Whether a dictionary is set or not, on the first call of deflate
		*     we change to BUSY_STATE.
		* (4) -- intentionally left blank -- :)
		* (5) FINISHING_STATE is entered, when flush() is called to indicate that
		*     there is no more INPUT.  There are also states indicating, that
		*     the header wasn't written yet.
		* (6) FINISHED_STATE is entered, when everything has been flushed to the
		*     internal pending output buffer.
		* (7) At any time (7)
		*
		*/

		#endregion Deflater Documentation

		#region Public Constants

		/// <summary>
		/// The best and slowest compression level.  This tries to find very
		/// long and distant string repetitions.
		/// </summary>
		public const int BEST_COMPRESSION = 9;

		/// <summary>
		/// The worst but fastest compression level.
		/// </summary>
		public const int BEST_SPEED = 1;

		/// <summary>
		/// The default compression level.
		/// </summary>
		public const int DEFAULT_COMPRESSION = -1;

		/// <summary>
		/// This level won't compress at all but output uncompressed blocks.
		/// </summary>
		public const int NO_COMPRESSION = 0;

		/// <summary>
		/// The compression method.  This is the only method supported so far.
		/// There is no need to use this constant at all.
		/// </summary>
		public const int DEFLATED = 8;

		#endregion Public Constants

		#region Public Enum

		/// <summary>
		/// Compression Level as an enum for safer use
		/// </summary>
		public enum CompressionLevel
		{
			/// <summary>
			/// The best and slowest compression level.  This tries to find very
			/// long and distant string repetitions.
			/// </summary>
			BEST_COMPRESSION = Deflater.BEST_COMPRESSION,

			/// <summary>
			/// The worst but fastest compression level.
			/// </summary>
			BEST_SPEED = Deflater.BEST_SPEED,

			/// <summary>
			/// The default compression level.
			/// </summary>
			DEFAULT_COMPRESSION = Deflater.DEFAULT_COMPRESSION,

			/// <summary>
			/// This level won't compress at all but output uncompressed blocks.
			/// </summary>
			NO_COMPRESSION = Deflater.NO_COMPRESSION,

			/// <summary>
			/// The compression method.  This is the only method supported so far.
			/// There is no need to use this constant at all.
			/// </summary>
			DEFLATED = Deflater.DEFLATED
		}

		#endregion Public Enum

		#region Local Constants

		private const int IS_SETDICT = 0x01;
		private const int IS_FLUSHING = 0x04;
		private const int IS_FINISHING = 0x08;

		private const int INIT_STATE = 0x00;
		private const int SETDICT_STATE = 0x01;

		//		private static  int INIT_FINISHING_STATE    = 0x08;
		//		private static  int SETDICT_FINISHING_STATE = 0x09;
		private const int BUSY_STATE = 0x10;

		private const int FLUSHING_STATE = 0x14;
		private const int FINISHING_STATE = 0x1c;
		private const int FINISHED_STATE = 0x1e;
		private const int CLOSED_STATE = 0x7f;

		#endregion Local Constants

		#region Constructors

		/// <summary>
		/// Creates a new deflater with default compression level.
		/// </summary>
		public Deflater() : this(DEFAULT_COMPRESSION, false)
		{
		}

		/// <summary>
		/// Creates a new deflater with given compression level.
		/// </summary>
		/// <param name="level">
		/// the compression level, a value between NO_COMPRESSION
		/// and BEST_COMPRESSION, or DEFAULT_COMPRESSION.
		/// </param>
		/// <exception cref="System.ArgumentOutOfRangeException">if lvl is out of range.</exception>
		public Deflater(int level) : this(level, false)
		{
		}

		/// <summary>
		/// Creates a new deflater with given compression level.
		/// </summary>
		/// <param name="level">
		/// the compression level, a value between NO_COMPRESSION
		/// and BEST_COMPRESSION.
		/// </param>
		/// <param name="noZlibHeaderOrFooter">
		/// true, if we should suppress the Zlib/RFC1950 header at the
		/// beginning and the adler checksum at the end of the output.  This is
		/// useful for the GZIP/PKZIP formats.
		/// </param>
		/// <exception cref="System.ArgumentOutOfRangeException">if lvl is out of range.</exception>
		public Deflater(int level, bool noZlibHeaderOrFooter)
		{
			if (level == DEFAULT_COMPRESSION)
			{
				level = 6;
			}
			else if (level < NO_COMPRESSION || level > BEST_COMPRESSION)
			{
				throw new ArgumentOutOfRangeException(nameof(level));
			}

			pending = new DeflaterPending();
			engine = new DeflaterEngine(pending, noZlibHeaderOrFooter);
			this.noZlibHeaderOrFooter = noZlibHeaderOrFooter;
			SetStrategy(DeflateStrategy.Default);
			SetLevel(level);
			Reset();
		}

		#endregion Constructors

		/// <summary>
		/// Resets the deflater.  The deflater acts afterwards as if it was
		/// just created with the same compression level and strategy as it
		/// had before.
		/// </summary>
		public void Reset()
		{
			state = (noZlibHeaderOrFooter ? BUSY_STATE : INIT_STATE);
			totalOut = 0;
			pending.Reset();
			engine.Reset();
		}

		/// <summary>
		/// Gets the current adler checksum of the data that was processed so far.
		/// </summary>
		public int Adler
		{
			get
			{
				return engine.Adler;
			}
		}

		/// <summary>
		/// Gets the number of input bytes processed so far.
		/// </summary>
		public long TotalIn
		{
			get
			{
				return engine.TotalIn;
			}
		}

		/// <summary>
		/// Gets the number of output bytes so far.
		/// </summary>
		public long TotalOut
		{
			get
			{
				return totalOut;
			}
		}

		/// <summary>
		/// Flushes the current input block.  Further calls to deflate() will
		/// produce enough output to inflate everything in the current input
		/// block.  This is not part of Sun's JDK so I have made it package
		/// private.  It is used by DeflaterOutputStream to implement
		/// flush().
		/// </summary>
		public void Flush()
		{
			state |= IS_FLUSHING;
		}

		/// <summary>
		/// Finishes the deflater with the current input block.  It is an error
		/// to give more input after this method was called.  This method must
		/// be called to force all bytes to be flushed.
		/// </summary>
		public void Finish()
		{
			state |= (IS_FLUSHING | IS_FINISHING);
		}

		/// <summary>
		/// Returns true if the stream was finished and no more output bytes
		/// are available.
		/// </summary>
		public bool IsFinished
		{
			get
			{
				return (state == FINISHED_STATE) && pending.IsFlushed;
			}
		}

		/// <summary>
		/// Returns true, if the input buffer is empty.
		/// You should then call setInput().
		/// NOTE: This method can also return true when the stream
		/// was finished.
		/// </summary>
		public bool IsNeedingInput
		{
			get
			{
				return engine.NeedsInput();
			}
		}

		/// <summary>
		/// Sets the data which should be compressed next.  This should be only
		/// called when needsInput indicates that more input is needed.
		/// If you call setInput when needsInput() returns false, the
		/// previous input that is still pending will be thrown away.
		/// The given byte array should not be changed, before needsInput() returns
		/// true again.
		/// This call is equivalent to <code>setInput(input, 0, input.length)</code>.
		/// </summary>
		/// <param name="input">
		/// the buffer containing the input data.
		/// </param>
		/// <exception cref="System.InvalidOperationException">
		/// if the buffer was finished() or ended().
		/// </exception>
		public void SetInput(byte[] input)
		{
			SetInput(input, 0, input.Length);
		}

		/// <summary>
		/// Sets the data which should be compressed next.  This should be
		/// only called when needsInput indicates that more input is needed.
		/// The given byte array should not be changed, before needsInput() returns
		/// true again.
		/// </summary>
		/// <param name="input">
		/// the buffer containing the input data.
		/// </param>
		/// <param name="offset">
		/// the start of the data.
		/// </param>
		/// <param name="count">
		/// the number of data bytes of input.
		/// </param>
		/// <exception cref="System.InvalidOperationException">
		/// if the buffer was Finish()ed or if previous input is still pending.
		/// </exception>
		public void SetInput(byte[] input, int offset, int count)
		{
			if ((state & IS_FINISHING) != 0)
			{
				throw new InvalidOperationException("Finish() already called");
			}
			engine.SetInput(input, offset, count);
		}

		/// <summary>
		/// Sets the compression level.  There is no guarantee of the exact
		/// position of the change, but if you call this when needsInput is
		/// true the change of compression level will occur somewhere near
		/// before the end of the so far given input.
		/// </summary>
		/// <param name="level">
		/// the new compression level.
		/// </param>
		public void SetLevel(int level)
		{
			if (level == DEFAULT_COMPRESSION)
			{
				level = 6;
			}
			else if (level < NO_COMPRESSION || level > BEST_COMPRESSION)
			{
				throw new ArgumentOutOfRangeException(nameof(level));
			}

			if (this.level != level)
			{
				this.level = level;
				engine.SetLevel(level);
			}
		}

		/// <summary>
		/// Get current compression level
		/// </summary>
		/// <returns>Returns the current compression level</returns>
		public int GetLevel()
		{
			return level;
		}

		/// <summary>
		/// Sets the compression strategy. Strategy is one of
		/// DEFAULT_STRATEGY, HUFFMAN_ONLY and FILTERED.  For the exact
		/// position where the strategy is changed, the same as for
		/// SetLevel() applies.
		/// </summary>
		/// <param name="strategy">
		/// The new compression strategy.
		/// </param>
		public void SetStrategy(DeflateStrategy strategy)
		{
			engine.Strategy = strategy;
		}

		/// <summary>
		/// Deflates the current input block with to the given array.
		/// </summary>
		/// <param name="output">
		/// The buffer where compressed data is stored
		/// </param>
		/// <returns>
		/// The number of compressed bytes added to the output, or 0 if either
		/// IsNeedingInput() or IsFinished returns true or length is zero.
		/// </returns>
		public int Deflate(byte[] output)
		{
			return Deflate(output, 0, output.Length);
		}

		/// <summary>
		/// Deflates the current input block to the given array.
		/// </summary>
		/// <param name="output">
		/// Buffer to store the compressed data.
		/// </param>
		/// <param name="offset">
		/// Offset into the output array.
		/// </param>
		/// <param name="length">
		/// The maximum number of bytes that may be stored.
		/// </param>
		/// <returns>
		/// The number of compressed bytes added to the output, or 0 if either
		/// needsInput() or finished() returns true or length is zero.
		/// </returns>
		/// <exception cref="System.InvalidOperationException">
		/// If Finish() was previously called.
		/// </exception>
		/// <exception cref="System.ArgumentOutOfRangeException">
		/// If offset or length don't match the array length.
		/// </exception>
		public int Deflate(byte[] output, int offset, int length)
		{
			int origLength = length;

			if (state == CLOSED_STATE)
			{
				throw new InvalidOperationException("Deflater closed");
			}

			if (state < BUSY_STATE)
			{
				// output header
				int header = (DEFLATED +
					((DeflaterConstants.MAX_WBITS - 8) << 4)) << 8;
				int level_flags = (level - 1) >> 1;
				if (level_flags < 0 || level_flags > 3)
				{
					level_flags = 3;
				}
				header |= level_flags << 6;
				if ((state & IS_SETDICT) != 0)
				{
					// Dictionary was set
					header |= DeflaterConstants.PRESET_DICT;
				}
				header += 31 - (header % 31);

				pending.WriteShortMSB(header);
				if ((state & IS_SETDICT) != 0)
				{
					int chksum = engine.Adler;
					engine.ResetAdler();
					pending.WriteShortMSB(chksum >> 16);
					pending.WriteShortMSB(chksum & 0xffff);
				}

				state = BUSY_STATE | (state & (IS_FLUSHING | IS_FINISHING));
			}

			for (; ; )
			{
				int count = pending.Flush(output, offset, length);
				offset += count;
				totalOut += count;
				length -= count;

				if (length == 0 || state == FINISHED_STATE)
				{
					break;
				}

				if (!engine.Deflate((state & IS_FLUSHING) != 0, (state & IS_FINISHING) != 0))
				{
					switch (state)
					{
						case BUSY_STATE:
							// We need more input now
							return origLength - length;

						case FLUSHING_STATE:
							if (level != NO_COMPRESSION)
							{
								/* We have to supply some lookahead.  8 bit lookahead
								 * is needed by the zlib inflater, and we must fill
								 * the next byte, so that all bits are flushed.
								 */
								int neededbits = 8 + ((-pending.BitCount) & 7);
								while (neededbits > 0)
								{
									/* write a static tree block consisting solely of
									 * an EOF:
									 */
									pending.WriteBits(2, 10);
									neededbits -= 10;
								}
							}
							state = BUSY_STATE;
							break;

						case FINISHING_STATE:
							pending.AlignToByte();

							// Compressed data is complete.  Write footer information if required.
							if (!noZlibHeaderOrFooter)
							{
								int adler = engine.Adler;
								pending.WriteShortMSB(adler >> 16);
								pending.WriteShortMSB(adler & 0xffff);
							}
							state = FINISHED_STATE;
							break;
					}
				}
			}
			return origLength - length;
		}

		/// <summary>
		/// Sets the dictionary which should be used in the deflate process.
		/// This call is equivalent to <code>setDictionary(dict, 0, dict.Length)</code>.
		/// </summary>
		/// <param name="dictionary">
		/// the dictionary.
		/// </param>
		/// <exception cref="System.InvalidOperationException">
		/// if SetInput () or Deflate () were already called or another dictionary was already set.
		/// </exception>
		public void SetDictionary(byte[] dictionary)
		{
			SetDictionary(dictionary, 0, dictionary.Length);
		}

		/// <summary>
		/// Sets the dictionary which should be used in the deflate process.
		/// The dictionary is a byte array containing strings that are
		/// likely to occur in the data which should be compressed.  The
		/// dictionary is not stored in the compressed output, only a
		/// checksum.  To decompress the output you need to supply the same
		/// dictionary again.
		/// </summary>
		/// <param name="dictionary">
		/// The dictionary data
		/// </param>
		/// <param name="index">
		/// The index where dictionary information commences.
		/// </param>
		/// <param name="count">
		/// The number of bytes in the dictionary.
		/// </param>
		/// <exception cref="System.InvalidOperationException">
		/// If SetInput () or Deflate() were already called or another dictionary was already set.
		/// </exception>
		public void SetDictionary(byte[] dictionary, int index, int count)
		{
			if (state != INIT_STATE)
			{
				throw new InvalidOperationException();
			}

			state = SETDICT_STATE;
			engine.SetDictionary(dictionary, index, count);
		}

		#region Instance Fields

		/// <summary>
		/// Compression level.
		/// </summary>
		private int level;

		/// <summary>
		/// If true no Zlib/RFC1950 headers or footers are generated
		/// </summary>
		private bool noZlibHeaderOrFooter;

		/// <summary>
		/// The current state.
		/// </summary>
		private int state;

		/// <summary>
		/// The total bytes of output written.
		/// </summary>
		private long totalOut;

		/// <summary>
		/// The pending output.
		/// </summary>
		private DeflaterPending pending;

		/// <summary>
		/// The deflater engine.
		/// </summary>
		private DeflaterEngine engine;

		#endregion Instance Fields
	}
}


namespace ICSharpCode.SharpZipLib.Zip.Compression
{
	/// <summary>
	/// This class contains constants used for deflation.
	/// </summary>
	[System.Diagnostics.CodeAnalysis.SuppressMessage("Naming", "CA1707:Identifiers should not contain underscores", Justification = "kept for backwards compatibility")]
	public static class DeflaterConstants
	{
		/// <summary>
		/// Set to true to enable debugging
		/// </summary>
		public const bool DEBUGGING = false;

		/// <summary>
		/// Written to Zip file to identify a stored block
		/// </summary>
		public const int STORED_BLOCK = 0;

		/// <summary>
		/// Identifies static tree in Zip file
		/// </summary>
		public const int STATIC_TREES = 1;

		/// <summary>
		/// Identifies dynamic tree in Zip file
		/// </summary>
		public const int DYN_TREES = 2;

		/// <summary>
		/// Header flag indicating a preset dictionary for deflation
		/// </summary>
		public const int PRESET_DICT = 0x20;

		/// <summary>
		/// Sets internal buffer sizes for Huffman encoding
		/// </summary>
		public const int DEFAULT_MEM_LEVEL = 8;

		/// <summary>
		/// Internal compression engine constant
		/// </summary>
		public const int MAX_MATCH = 258;

		/// <summary>
		/// Internal compression engine constant
		/// </summary>
		public const int MIN_MATCH = 3;

		/// <summary>
		/// Internal compression engine constant
		/// </summary>
		public const int MAX_WBITS = 15;

		/// <summary>
		/// Internal compression engine constant
		/// </summary>
		public const int WSIZE = 1 << MAX_WBITS;

		/// <summary>
		/// Internal compression engine constant
		/// </summary>
		public const int WMASK = WSIZE - 1;

		/// <summary>
		/// Internal compression engine constant
		/// </summary>
		public const int HASH_BITS = DEFAULT_MEM_LEVEL + 7;

		/// <summary>
		/// Internal compression engine constant
		/// </summary>
		public const int HASH_SIZE = 1 << HASH_BITS;

		/// <summary>
		/// Internal compression engine constant
		/// </summary>
		public const int HASH_MASK = HASH_SIZE - 1;

		/// <summary>
		/// Internal compression engine constant
		/// </summary>
		public const int HASH_SHIFT = (HASH_BITS + MIN_MATCH - 1) / MIN_MATCH;

		/// <summary>
		/// Internal compression engine constant
		/// </summary>
		public const int MIN_LOOKAHEAD = MAX_MATCH + MIN_MATCH + 1;

		/// <summary>
		/// Internal compression engine constant
		/// </summary>
		public const int MAX_DIST = WSIZE - MIN_LOOKAHEAD;

		/// <summary>
		/// Internal compression engine constant
		/// </summary>
		public const int PENDING_BUF_SIZE = 1 << (DEFAULT_MEM_LEVEL + 8);

		/// <summary>
		/// Internal compression engine constant
		/// </summary>
		public static int MAX_BLOCK_SIZE = Math.Min(65535, PENDING_BUF_SIZE - 5);

		/// <summary>
		/// Internal compression engine constant
		/// </summary>
		public const int DEFLATE_STORED = 0;

		/// <summary>
		/// Internal compression engine constant
		/// </summary>
		public const int DEFLATE_FAST = 1;

		/// <summary>
		/// Internal compression engine constant
		/// </summary>
		public const int DEFLATE_SLOW = 2;

		/// <summary>
		/// Internal compression engine constant
		/// </summary>
		public static int[] GOOD_LENGTH = { 0, 4, 4, 4, 4, 8, 8, 8, 32, 32 };

		/// <summary>
		/// Internal compression engine constant
		/// </summary>
		public static int[] MAX_LAZY = { 0, 4, 5, 6, 4, 16, 16, 32, 128, 258 };

		/// <summary>
		/// Internal compression engine constant
		/// </summary>
		public static int[] NICE_LENGTH = { 0, 8, 16, 32, 16, 32, 128, 128, 258, 258 };

		/// <summary>
		/// Internal compression engine constant
		/// </summary>
		public static int[] MAX_CHAIN = { 0, 4, 8, 32, 16, 32, 128, 256, 1024, 4096 };

		/// <summary>
		/// Internal compression engine constant
		/// </summary>
		public static int[] COMPR_FUNC = { 0, 1, 1, 1, 1, 2, 2, 2, 2, 2 };
	}
}

namespace ICSharpCode.SharpZipLib.Zip.Compression
{
	/// <summary>
	/// Strategies for deflater
	/// </summary>
	public enum DeflateStrategy
	{
		/// <summary>
		/// The default strategy
		/// </summary>
		Default = 0,

		/// <summary>
		/// This strategy will only allow longer string repetitions.  It is
		/// useful for random data with a small character set.
		/// </summary>
		Filtered = 1,

		/// <summary>
		/// This strategy will not look for string repetitions at all.  It
		/// only encodes with Huffman trees (which means, that more common
		/// characters get a smaller encoding.
		/// </summary>
		HuffmanOnly = 2
	}

	// DEFLATE ALGORITHM:
	//
	// The uncompressed stream is inserted into the window array.  When
	// the window array is full the first half is thrown away and the
	// second half is copied to the beginning.
	//
	// The head array is a hash table.  Three characters build a hash value
	// and they the value points to the corresponding index in window of
	// the last string with this hash.  The prev array implements a
	// linked list of matches with the same hash: prev[index & WMASK] points
	// to the previous index with the same hash.
	//

	/// <summary>
	/// Low level compression engine for deflate algorithm which uses a 32K sliding window
	/// with secondary compression from Huffman/Shannon-Fano codes.
	/// </summary>
	public class DeflaterEngine
	{
		#region Constants

		private const int TooFar = 4096;

		#endregion Constants

		#region Constructors

		/// <summary>
		/// Construct instance with pending buffer
		/// Adler calculation will be performed
		/// </summary>
		/// <param name="pending">
		/// Pending buffer to use
		/// </param>
		public DeflaterEngine(DeflaterPending pending)
			: this (pending, false)
		{
		}



		/// <summary>
		/// Construct instance with pending buffer
		/// </summary>
		/// <param name="pending">
		/// Pending buffer to use
		/// </param>
		/// <param name="noAdlerCalculation">
		/// If no adler calculation should be performed
		/// </param>
		public DeflaterEngine(DeflaterPending pending, bool noAdlerCalculation)
		{
			this.pending = pending;
			huffman = new DeflaterHuffman(pending);
			if (!noAdlerCalculation)
				adler = new Adler32();

			window = new byte[2 * DeflaterConstants.WSIZE];
			head = new short[DeflaterConstants.HASH_SIZE];
			prev = new short[DeflaterConstants.WSIZE];

			// We start at index 1, to avoid an implementation deficiency, that
			// we cannot build a repeat pattern at index 0.
			blockStart = strstart = 1;
		}

		#endregion Constructors

		/// <summary>
		/// Deflate drives actual compression of data
		/// </summary>
		/// <param name="flush">True to flush input buffers</param>
		/// <param name="finish">Finish deflation with the current input.</param>
		/// <returns>Returns true if progress has been made.</returns>
		public bool Deflate(bool flush, bool finish)
		{
			bool progress;
			do
			{
				FillWindow();
				bool canFlush = flush && (inputOff == inputEnd);

#if DebugDeflation
				if (DeflaterConstants.DEBUGGING) {
					Console.WriteLine("window: [" + blockStart + "," + strstart + ","
								+ lookahead + "], " + compressionFunction + "," + canFlush);
				}
#endif
				switch (compressionFunction)
				{
					case DeflaterConstants.DEFLATE_STORED:
						progress = DeflateStored(canFlush, finish);
						break;

					case DeflaterConstants.DEFLATE_FAST:
						progress = DeflateFast(canFlush, finish);
						break;

					case DeflaterConstants.DEFLATE_SLOW:
						progress = DeflateSlow(canFlush, finish);
						break;

					default:
						throw new InvalidOperationException("unknown compressionFunction");
				}
			} while (pending.IsFlushed && progress); // repeat while we have no pending output and progress was made
			return progress;
		}

		/// <summary>
		/// Sets input data to be deflated.  Should only be called when <code>NeedsInput()</code>
		/// returns true
		/// </summary>
		/// <param name="buffer">The buffer containing input data.</param>
		/// <param name="offset">The offset of the first byte of data.</param>
		/// <param name="count">The number of bytes of data to use as input.</param>
		public void SetInput(byte[] buffer, int offset, int count)
		{
			if (buffer == null)
			{
				throw new ArgumentNullException(nameof(buffer));
			}

			if (offset < 0)
			{
				throw new ArgumentOutOfRangeException(nameof(offset));
			}

			if (count < 0)
			{
				throw new ArgumentOutOfRangeException(nameof(count));
			}

			if (inputOff < inputEnd)
			{
				throw new InvalidOperationException("Old input was not completely processed");
			}

			int end = offset + count;

			/* We want to throw an ArrayIndexOutOfBoundsException early.  The
			* check is very tricky: it also handles integer wrap around.
			*/
			if ((offset > end) || (end > buffer.Length))
			{
				throw new ArgumentOutOfRangeException(nameof(count));
			}

			inputBuf = buffer;
			inputOff = offset;
			inputEnd = end;
		}

		/// <summary>
		/// Determines if more <see cref="SetInput">input</see> is needed.
		/// </summary>
		/// <returns>Return true if input is needed via <see cref="SetInput">SetInput</see></returns>
		public bool NeedsInput()
		{
			return (inputEnd == inputOff);
		}

		/// <summary>
		/// Set compression dictionary
		/// </summary>
		/// <param name="buffer">The buffer containing the dictionary data</param>
		/// <param name="offset">The offset in the buffer for the first byte of data</param>
		/// <param name="length">The length of the dictionary data.</param>
		public void SetDictionary(byte[] buffer, int offset, int length)
		{
#if DebugDeflation
			if (DeflaterConstants.DEBUGGING && (strstart != 1) )
			{
				throw new InvalidOperationException("strstart not 1");
			}
#endif
			adler?.Update(new ArraySegment<byte>(buffer, offset, length));
			if (length < DeflaterConstants.MIN_MATCH)
			{
				return;
			}

			if (length > DeflaterConstants.MAX_DIST)
			{
				offset += length - DeflaterConstants.MAX_DIST;
				length = DeflaterConstants.MAX_DIST;
			}

			System.Array.Copy(buffer, offset, window, strstart, length);

			UpdateHash();
			--length;
			while (--length > 0)
			{
				InsertString();
				strstart++;
			}
			strstart += 2;
			blockStart = strstart;
		}

		/// <summary>
		/// Reset internal state
		/// </summary>
		public void Reset()
		{
			huffman.Reset();
			adler?.Reset();
			blockStart = strstart = 1;
			lookahead = 0;
			totalIn = 0;
			prevAvailable = false;
			matchLen = DeflaterConstants.MIN_MATCH - 1;

			for (int i = 0; i < DeflaterConstants.HASH_SIZE; i++)
			{
				head[i] = 0;
			}

			for (int i = 0; i < DeflaterConstants.WSIZE; i++)
			{
				prev[i] = 0;
			}
		}

		/// <summary>
		/// Reset Adler checksum
		/// </summary>
		public void ResetAdler()
		{
			adler?.Reset();
		}

		/// <summary>
		/// Get current value of Adler checksum
		/// </summary>
		public int Adler
		{
			get
			{
				return (adler != null) ? unchecked((int)adler.Value) : 0;
			}
		}

		/// <summary>
		/// Total data processed
		/// </summary>
		public long TotalIn
		{
			get
			{
				return totalIn;
			}
		}

		/// <summary>
		/// Get/set the <see cref="DeflateStrategy">deflate strategy</see>
		/// </summary>
		public DeflateStrategy Strategy
		{
			get
			{
				return strategy;
			}
			set
			{
				strategy = value;
			}
		}

		/// <summary>
		/// Set the deflate level (0-9)
		/// </summary>
		/// <param name="level">The value to set the level to.</param>
		public void SetLevel(int level)
		{
			if ((level < 0) || (level > 9))
			{
				throw new ArgumentOutOfRangeException(nameof(level));
			}

			goodLength = DeflaterConstants.GOOD_LENGTH[level];
			max_lazy = DeflaterConstants.MAX_LAZY[level];
			niceLength = DeflaterConstants.NICE_LENGTH[level];
			max_chain = DeflaterConstants.MAX_CHAIN[level];

			if (DeflaterConstants.COMPR_FUNC[level] != compressionFunction)
			{
#if DebugDeflation
				if (DeflaterConstants.DEBUGGING) {
				   Console.WriteLine("Change from " + compressionFunction + " to "
										  + DeflaterConstants.COMPR_FUNC[level]);
				}
#endif
				switch (compressionFunction)
				{
					case DeflaterConstants.DEFLATE_STORED:
						if (strstart > blockStart)
						{
							huffman.FlushStoredBlock(window, blockStart,
								strstart - blockStart, false);
							blockStart = strstart;
						}
						UpdateHash();
						break;

					case DeflaterConstants.DEFLATE_FAST:
						if (strstart > blockStart)
						{
							huffman.FlushBlock(window, blockStart, strstart - blockStart,
								false);
							blockStart = strstart;
						}
						break;

					case DeflaterConstants.DEFLATE_SLOW:
						if (prevAvailable)
						{
							huffman.TallyLit(window[strstart - 1] & 0xff);
						}
						if (strstart > blockStart)
						{
							huffman.FlushBlock(window, blockStart, strstart - blockStart, false);
							blockStart = strstart;
						}
						prevAvailable = false;
						matchLen = DeflaterConstants.MIN_MATCH - 1;
						break;
				}
				compressionFunction = DeflaterConstants.COMPR_FUNC[level];
			}
		}

		/// <summary>
		/// Fill the window
		/// </summary>
		public void FillWindow()
		{
			/* If the window is almost full and there is insufficient lookahead,
			 * move the upper half to the lower one to make room in the upper half.
			 */
			if (strstart >= DeflaterConstants.WSIZE + DeflaterConstants.MAX_DIST)
			{
				SlideWindow();
			}

			/* If there is not enough lookahead, but still some input left,
			 * read in the input
			 */
			if (lookahead < DeflaterConstants.MIN_LOOKAHEAD && inputOff < inputEnd)
			{
				int more = 2 * DeflaterConstants.WSIZE - lookahead - strstart;

				if (more > inputEnd - inputOff)
				{
					more = inputEnd - inputOff;
				}

				System.Array.Copy(inputBuf, inputOff, window, strstart + lookahead, more);
				adler?.Update(new ArraySegment<byte>(inputBuf, inputOff, more));

				inputOff += more;
				totalIn += more;
				lookahead += more;
			}

			if (lookahead >= DeflaterConstants.MIN_MATCH)
			{
				UpdateHash();
			}
		}

		private void UpdateHash()
		{
			/*
						if (DEBUGGING) {
							Console.WriteLine("updateHash: "+strstart);
						}
			*/
			ins_h = (window[strstart] << DeflaterConstants.HASH_SHIFT) ^ window[strstart + 1];
		}

		/// <summary>
		/// Inserts the current string in the head hash and returns the previous
		/// value for this hash.
		/// </summary>
		/// <returns>The previous hash value</returns>
		private int InsertString()
		{
			short match;
			int hash = ((ins_h << DeflaterConstants.HASH_SHIFT) ^ window[strstart + (DeflaterConstants.MIN_MATCH - 1)]) & DeflaterConstants.HASH_MASK;

#if DebugDeflation
			if (DeflaterConstants.DEBUGGING)
			{
				if (hash != (((window[strstart] << (2*HASH_SHIFT)) ^
								  (window[strstart + 1] << HASH_SHIFT) ^
								  (window[strstart + 2])) & HASH_MASK)) {
						throw new SharpZipBaseException("hash inconsistent: " + hash + "/"
												+window[strstart] + ","
												+window[strstart + 1] + ","
												+window[strstart + 2] + "," + HASH_SHIFT);
					}
			}
#endif
			prev[strstart & DeflaterConstants.WMASK] = match = head[hash];
			head[hash] = unchecked((short)strstart);
			ins_h = hash;
			return match & 0xffff;
		}

		private void SlideWindow()
		{
			Array.Copy(window, DeflaterConstants.WSIZE, window, 0, DeflaterConstants.WSIZE);
			matchStart -= DeflaterConstants.WSIZE;
			strstart -= DeflaterConstants.WSIZE;
			blockStart -= DeflaterConstants.WSIZE;

			// Slide the hash table (could be avoided with 32 bit values
			// at the expense of memory usage).
			for (int i = 0; i < DeflaterConstants.HASH_SIZE; ++i)
			{
				int m = head[i] & 0xffff;
				head[i] = (short)(m >= DeflaterConstants.WSIZE ? (m - DeflaterConstants.WSIZE) : 0);
			}

			// Slide the prev table.
			for (int i = 0; i < DeflaterConstants.WSIZE; i++)
			{
				int m = prev[i] & 0xffff;
				prev[i] = (short)(m >= DeflaterConstants.WSIZE ? (m - DeflaterConstants.WSIZE) : 0);
			}
		}

		/// <summary>
		/// Find the best (longest) string in the window matching the
		/// string starting at strstart.
		///
		/// Preconditions:
		/// <code>
		/// strstart + DeflaterConstants.MAX_MATCH &lt;= window.length.</code>
		/// </summary>
		/// <param name="curMatch"></param>
		/// <returns>True if a match greater than the minimum length is found</returns>
		private bool FindLongestMatch(int curMatch)
		{
			int match;
			int scan = strstart;
			// scanMax is the highest position that we can look at
			int scanMax = scan + Math.Min(DeflaterConstants.MAX_MATCH, lookahead) - 1;
			int limit = Math.Max(scan - DeflaterConstants.MAX_DIST, 0);

			byte[] window = this.window;
			short[] prev = this.prev;
			int chainLength = this.max_chain;
			int niceLength = Math.Min(this.niceLength, lookahead);

			matchLen = Math.Max(matchLen, DeflaterConstants.MIN_MATCH - 1);

			if (scan + matchLen > scanMax) return false;

			byte scan_end1 = window[scan + matchLen - 1];
			byte scan_end = window[scan + matchLen];

			// Do not waste too much time if we already have a good match:
			if (matchLen >= this.goodLength) chainLength >>= 2;

			do
			{
				match = curMatch;
				scan = strstart;

				if (window[match + matchLen] != scan_end
				 || window[match + matchLen - 1] != scan_end1
				 || window[match] != window[scan]
				 || window[++match] != window[++scan])
				{
					continue;
				}

				// scan is set to strstart+1 and the comparison passed, so
				// scanMax - scan is the maximum number of bytes we can compare.
				// below we compare 8 bytes at a time, so first we compare
				// (scanMax - scan) % 8 bytes, so the remainder is a multiple of 8

				switch ((scanMax - scan) % 8)
				{
					case 1:
						if (window[++scan] == window[++match]) break;
						break;

					case 2:
						if (window[++scan] == window[++match]
				  && window[++scan] == window[++match]) break;
						break;

					case 3:
						if (window[++scan] == window[++match]
				  && window[++scan] == window[++match]
				  && window[++scan] == window[++match]) break;
						break;

					case 4:
						if (window[++scan] == window[++match]
				  && window[++scan] == window[++match]
				  && window[++scan] == window[++match]
				  && window[++scan] == window[++match]) break;
						break;

					case 5:
						if (window[++scan] == window[++match]
				  && window[++scan] == window[++match]
				  && window[++scan] == window[++match]
				  && window[++scan] == window[++match]
				  && window[++scan] == window[++match]) break;
						break;

					case 6:
						if (window[++scan] == window[++match]
				  && window[++scan] == window[++match]
				  && window[++scan] == window[++match]
				  && window[++scan] == window[++match]
				  && window[++scan] == window[++match]
				  && window[++scan] == window[++match]) break;
						break;

					case 7:
						if (window[++scan] == window[++match]
				  && window[++scan] == window[++match]
				  && window[++scan] == window[++match]
				  && window[++scan] == window[++match]
				  && window[++scan] == window[++match]
				  && window[++scan] == window[++match]
				  && window[++scan] == window[++match]) break;
						break;
				}

				if (window[scan] == window[match])
				{
					/* We check for insufficient lookahead only every 8th comparison;
					 * the 256th check will be made at strstart + 258 unless lookahead is
					 * exhausted first.
					 */
					do
					{
						if (scan == scanMax)
						{
							++scan;     // advance to first position not matched
							++match;

							break;
						}
					}
					while (window[++scan] == window[++match]
						&& window[++scan] == window[++match]
						&& window[++scan] == window[++match]
						&& window[++scan] == window[++match]
						&& window[++scan] == window[++match]
						&& window[++scan] == window[++match]
						&& window[++scan] == window[++match]
						&& window[++scan] == window[++match]);
				}

				if (scan - strstart > matchLen)
				{
#if DebugDeflation
              if (DeflaterConstants.DEBUGGING && (ins_h == 0) )
              Console.Error.WriteLine("Found match: " + curMatch + "-" + (scan - strstart));
#endif

					matchStart = curMatch;
					matchLen = scan - strstart;

					if (matchLen >= niceLength)
						break;

					scan_end1 = window[scan - 1];
					scan_end = window[scan];
				}
			} while ((curMatch = (prev[curMatch & DeflaterConstants.WMASK] & 0xffff)) > limit && 0 != --chainLength);

			return matchLen >= DeflaterConstants.MIN_MATCH;
		}

		private bool DeflateStored(bool flush, bool finish)
		{
			if (!flush && (lookahead == 0))
			{
				return false;
			}

			strstart += lookahead;
			lookahead = 0;

			int storedLength = strstart - blockStart;

			if ((storedLength >= DeflaterConstants.MAX_BLOCK_SIZE) || // Block is full
				(blockStart < DeflaterConstants.WSIZE && storedLength >= DeflaterConstants.MAX_DIST) ||   // Block may move out of window
				flush)
			{
				bool lastBlock = finish;
				if (storedLength > DeflaterConstants.MAX_BLOCK_SIZE)
				{
					storedLength = DeflaterConstants.MAX_BLOCK_SIZE;
					lastBlock = false;
				}

#if DebugDeflation
				if (DeflaterConstants.DEBUGGING)
				{
				   Console.WriteLine("storedBlock[" + storedLength + "," + lastBlock + "]");
				}
#endif

				huffman.FlushStoredBlock(window, blockStart, storedLength, lastBlock);
				blockStart += storedLength;
				return !(lastBlock || storedLength == 0);
			}
			return true;
		}

		private bool DeflateFast(bool flush, bool finish)
		{
			if (lookahead < DeflaterConstants.MIN_LOOKAHEAD && !flush)
			{
				return false;
			}

			while (lookahead >= DeflaterConstants.MIN_LOOKAHEAD || flush)
			{
				if (lookahead == 0)
				{
					// We are flushing everything
					huffman.FlushBlock(window, blockStart, strstart - blockStart, finish);
					blockStart = strstart;
					return false;
				}

				if (strstart > 2 * DeflaterConstants.WSIZE - DeflaterConstants.MIN_LOOKAHEAD)
				{
					/* slide window, as FindLongestMatch needs this.
					 * This should only happen when flushing and the window
					 * is almost full.
					 */
					SlideWindow();
				}

				int hashHead;
				if (lookahead >= DeflaterConstants.MIN_MATCH &&
					(hashHead = InsertString()) != 0 &&
					strategy != DeflateStrategy.HuffmanOnly &&
					strstart - hashHead <= DeflaterConstants.MAX_DIST &&
					FindLongestMatch(hashHead))
				{
					// longestMatch sets matchStart and matchLen
#if DebugDeflation
					if (DeflaterConstants.DEBUGGING)
					{
						for (int i = 0 ; i < matchLen; i++) {
							if (window[strstart + i] != window[matchStart + i]) {
								throw new SharpZipBaseException("Match failure");
							}
						}
					}
#endif

					bool full = huffman.TallyDist(strstart - matchStart, matchLen);

					lookahead -= matchLen;
					if (matchLen <= max_lazy && lookahead >= DeflaterConstants.MIN_MATCH)
					{
						while (--matchLen > 0)
						{
							++strstart;
							InsertString();
						}
						++strstart;
					}
					else
					{
						strstart += matchLen;
						if (lookahead >= DeflaterConstants.MIN_MATCH - 1)
						{
							UpdateHash();
						}
					}
					matchLen = DeflaterConstants.MIN_MATCH - 1;
					if (!full)
					{
						continue;
					}
				}
				else
				{
					// No match found
					huffman.TallyLit(window[strstart] & 0xff);
					++strstart;
					--lookahead;
				}

				if (huffman.IsFull())
				{
					bool lastBlock = finish && (lookahead == 0);
					huffman.FlushBlock(window, blockStart, strstart - blockStart, lastBlock);
					blockStart = strstart;
					return !lastBlock;
				}
			}
			return true;
		}

		private bool DeflateSlow(bool flush, bool finish)
		{
			if (lookahead < DeflaterConstants.MIN_LOOKAHEAD && !flush)
			{
				return false;
			}

			while (lookahead >= DeflaterConstants.MIN_LOOKAHEAD || flush)
			{
				if (lookahead == 0)
				{
					if (prevAvailable)
					{
						huffman.TallyLit(window[strstart - 1] & 0xff);
					}
					prevAvailable = false;

					// We are flushing everything
#if DebugDeflation
					if (DeflaterConstants.DEBUGGING && !flush)
					{
						throw new SharpZipBaseException("Not flushing, but no lookahead");
					}
#endif
					huffman.FlushBlock(window, blockStart, strstart - blockStart,
						finish);
					blockStart = strstart;
					return false;
				}

				if (strstart >= 2 * DeflaterConstants.WSIZE - DeflaterConstants.MIN_LOOKAHEAD)
				{
					/* slide window, as FindLongestMatch needs this.
					 * This should only happen when flushing and the window
					 * is almost full.
					 */
					SlideWindow();
				}

				int prevMatch = matchStart;
				int prevLen = matchLen;
				if (lookahead >= DeflaterConstants.MIN_MATCH)
				{
					int hashHead = InsertString();

					if (strategy != DeflateStrategy.HuffmanOnly &&
						hashHead != 0 &&
						strstart - hashHead <= DeflaterConstants.MAX_DIST &&
						FindLongestMatch(hashHead))
					{
						// longestMatch sets matchStart and matchLen

						// Discard match if too small and too far away
						if (matchLen <= 5 && (strategy == DeflateStrategy.Filtered || (matchLen == DeflaterConstants.MIN_MATCH && strstart - matchStart > TooFar)))
						{
							matchLen = DeflaterConstants.MIN_MATCH - 1;
						}
					}
				}

				// previous match was better
				if ((prevLen >= DeflaterConstants.MIN_MATCH) && (matchLen <= prevLen))
				{
#if DebugDeflation
					if (DeflaterConstants.DEBUGGING)
					{
					   for (int i = 0 ; i < matchLen; i++) {
						  if (window[strstart-1+i] != window[prevMatch + i])
							 throw new SharpZipBaseException();
						}
					}
#endif
					huffman.TallyDist(strstart - 1 - prevMatch, prevLen);
					prevLen -= 2;
					do
					{
						strstart++;
						lookahead--;
						if (lookahead >= DeflaterConstants.MIN_MATCH)
						{
							InsertString();
						}
					} while (--prevLen > 0);

					strstart++;
					lookahead--;
					prevAvailable = false;
					matchLen = DeflaterConstants.MIN_MATCH - 1;
				}
				else
				{
					if (prevAvailable)
					{
						huffman.TallyLit(window[strstart - 1] & 0xff);
					}
					prevAvailable = true;
					strstart++;
					lookahead--;
				}

				if (huffman.IsFull())
				{
					int len = strstart - blockStart;
					if (prevAvailable)
					{
						len--;
					}
					bool lastBlock = (finish && (lookahead == 0) && !prevAvailable);
					huffman.FlushBlock(window, blockStart, len, lastBlock);
					blockStart += len;
					return !lastBlock;
				}
			}
			return true;
		}

		#region Instance Fields

		// Hash index of string to be inserted
		private int ins_h;

		/// <summary>
		/// Hashtable, hashing three characters to an index for window, so
		/// that window[index]..window[index+2] have this hash code.
		/// Note that the array should really be unsigned short, so you need
		/// to and the values with 0xffff.
		/// </summary>
		private short[] head;

		/// <summary>
		/// <code>prev[index &amp; WMASK]</code> points to the previous index that has the
		/// same hash code as the string starting at index.  This way
		/// entries with the same hash code are in a linked list.
		/// Note that the array should really be unsigned short, so you need
		/// to and the values with 0xffff.
		/// </summary>
		private short[] prev;

		private int matchStart;

		// Length of best match
		private int matchLen;

		// Set if previous match exists
		private bool prevAvailable;

		private int blockStart;

		/// <summary>
		/// Points to the current character in the window.
		/// </summary>
		private int strstart;

		/// <summary>
		/// lookahead is the number of characters starting at strstart in
		/// window that are valid.
		/// So window[strstart] until window[strstart+lookahead-1] are valid
		/// characters.
		/// </summary>
		private int lookahead;

		/// <summary>
		/// This array contains the part of the uncompressed stream that
		/// is of relevance.  The current character is indexed by strstart.
		/// </summary>
		private byte[] window;

		private DeflateStrategy strategy;
		private int max_chain, max_lazy, niceLength, goodLength;

		/// <summary>
		/// The current compression function.
		/// </summary>
		private int compressionFunction;

		/// <summary>
		/// The input data for compression.
		/// </summary>
		private byte[] inputBuf;

		/// <summary>
		/// The total bytes of input read.
		/// </summary>
		private long totalIn;

		/// <summary>
		/// The offset into inputBuf, where input data starts.
		/// </summary>
		private int inputOff;

		/// <summary>
		/// The end offset of the input data.
		/// </summary>
		private int inputEnd;

		private DeflaterPending pending;
		private DeflaterHuffman huffman;

		/// <summary>
		/// The adler checksum
		/// </summary>
		private Adler32 adler;

		#endregion Instance Fields
	}
}


namespace ICSharpCode.SharpZipLib.Zip.Compression
{
	/// <summary>
	/// This is the DeflaterHuffman class.
	///
	/// This class is <i>not</i> thread safe.  This is inherent in the API, due
	/// to the split of Deflate and SetInput.
	///
	/// author of the original java version : Jochen Hoenicke
	/// </summary>
	public class DeflaterHuffman
	{
		private const int BUFSIZE = 1 << (DeflaterConstants.DEFAULT_MEM_LEVEL + 6);
		private const int LITERAL_NUM = 286;

		// Number of distance codes
		private const int DIST_NUM = 30;

		// Number of codes used to transfer bit lengths
		private const int BITLEN_NUM = 19;

		// repeat previous bit length 3-6 times (2 bits of repeat count)
		private const int REP_3_6 = 16;

		// repeat a zero length 3-10 times  (3 bits of repeat count)
		private const int REP_3_10 = 17;

		// repeat a zero length 11-138 times  (7 bits of repeat count)
		private const int REP_11_138 = 18;

		private const int EOF_SYMBOL = 256;

		// The lengths of the bit length codes are sent in order of decreasing
		// probability, to avoid transmitting the lengths for unused bit length codes.
		private static readonly int[] BL_ORDER = { 16, 17, 18, 0, 8, 7, 9, 6, 10, 5, 11, 4, 12, 3, 13, 2, 14, 1, 15 };

		private static readonly byte[] bit4Reverse = {
			0,
			8,
			4,
			12,
			2,
			10,
			6,
			14,
			1,
			9,
			5,
			13,
			3,
			11,
			7,
			15
		};

		private static short[] staticLCodes;
		private static byte[] staticLLength;
		private static short[] staticDCodes;
		private static byte[] staticDLength;

		private class Tree
		{
			#region Instance Fields

			public short[] freqs;

			public byte[] length;

			public int minNumCodes;

			public int numCodes;

			private short[] codes;
			private readonly int[] bl_counts;
			private readonly int maxLength;
			private DeflaterHuffman dh;

			#endregion Instance Fields

			#region Constructors

			public Tree(DeflaterHuffman dh, int elems, int minCodes, int maxLength)
			{
				this.dh = dh;
				this.minNumCodes = minCodes;
				this.maxLength = maxLength;
				freqs = new short[elems];
				bl_counts = new int[maxLength];
			}

			#endregion Constructors

			/// <summary>
			/// Resets the internal state of the tree
			/// </summary>
			public void Reset()
			{
				for (int i = 0; i < freqs.Length; i++)
				{
					freqs[i] = 0;
				}
				codes = null;
				length = null;
			}

			public void WriteSymbol(int code)
			{
				//				if (DeflaterConstants.DEBUGGING) {
				//					freqs[code]--;
				//					//  	  Console.Write("writeSymbol("+freqs.length+","+code+"): ");
				//				}
				dh.pending.WriteBits(codes[code] & 0xffff, length[code]);
			}

			/// <summary>
			/// Check that all frequencies are zero
			/// </summary>
			/// <exception cref="SharpZipBaseException">
			/// At least one frequency is non-zero
			/// </exception>
			public void CheckEmpty()
			{
				bool empty = true;
				for (int i = 0; i < freqs.Length; i++)
				{
					empty &= freqs[i] == 0;
				}

				if (!empty)
				{
					throw new SharpZipBaseException("!Empty");
				}
			}

			/// <summary>
			/// Set static codes and length
			/// </summary>
			/// <param name="staticCodes">new codes</param>
			/// <param name="staticLengths">length for new codes</param>
			public void SetStaticCodes(short[] staticCodes, byte[] staticLengths)
			{
				codes = staticCodes;
				length = staticLengths;
			}

			/// <summary>
			/// Build dynamic codes and lengths
			/// </summary>
			public void BuildCodes()
			{
				int numSymbols = freqs.Length;
				int[] nextCode = new int[maxLength];
				int code = 0;

				codes = new short[freqs.Length];

				//				if (DeflaterConstants.DEBUGGING) {
				//					//Console.WriteLine("buildCodes: "+freqs.Length);
				//				}

				for (int bits = 0; bits < maxLength; bits++)
				{
					nextCode[bits] = code;
					code += bl_counts[bits] << (15 - bits);

					//					if (DeflaterConstants.DEBUGGING) {
					//						//Console.WriteLine("bits: " + ( bits + 1) + " count: " + bl_counts[bits]
					//						                  +" nextCode: "+code);
					//					}
				}

#if DebugDeflation
				if ( DeflaterConstants.DEBUGGING && (code != 65536) )
				{
					throw new SharpZipBaseException("Inconsistent bl_counts!");
				}
#endif
				for (int i = 0; i < numCodes; i++)
				{
					int bits = length[i];
					if (bits > 0)
					{
						//						if (DeflaterConstants.DEBUGGING) {
						//								//Console.WriteLine("codes["+i+"] = rev(" + nextCode[bits-1]+"),
						//								                  +bits);
						//						}

						codes[i] = BitReverse(nextCode[bits - 1]);
						nextCode[bits - 1] += 1 << (16 - bits);
					}
				}
			}

			public void BuildTree()
			{
				int numSymbols = freqs.Length;

				/* heap is a priority queue, sorted by frequency, least frequent
				* nodes first.  The heap is a binary tree, with the property, that
				* the parent node is smaller than both child nodes.  This assures
				* that the smallest node is the first parent.
				*
				* The binary tree is encoded in an array:  0 is root node and
				* the nodes 2*n+1, 2*n+2 are the child nodes of node n.
				*/
				int[] heap = new int[numSymbols];
				int heapLen = 0;
				int maxCode = 0;
				for (int n = 0; n < numSymbols; n++)
				{
					int freq = freqs[n];
					if (freq != 0)
					{
						// Insert n into heap
						int pos = heapLen++;
						int ppos;
						while (pos > 0 && freqs[heap[ppos = (pos - 1) / 2]] > freq)
						{
							heap[pos] = heap[ppos];
							pos = ppos;
						}
						heap[pos] = n;

						maxCode = n;
					}
				}

				/* We could encode a single literal with 0 bits but then we
				* don't see the literals.  Therefore we force at least two
				* literals to avoid this case.  We don't care about order in
				* this case, both literals get a 1 bit code.
				*/
				while (heapLen < 2)
				{
					int node = maxCode < 2 ? ++maxCode : 0;
					heap[heapLen++] = node;
				}

				numCodes = Math.Max(maxCode + 1, minNumCodes);

				int numLeafs = heapLen;
				int[] childs = new int[4 * heapLen - 2];
				int[] values = new int[2 * heapLen - 1];
				int numNodes = numLeafs;
				for (int i = 0; i < heapLen; i++)
				{
					int node = heap[i];
					childs[2 * i] = node;
					childs[2 * i + 1] = -1;
					values[i] = freqs[node] << 8;
					heap[i] = i;
				}

				/* Construct the Huffman tree by repeatedly combining the least two
				* frequent nodes.
				*/
				do
				{
					int first = heap[0];
					int last = heap[--heapLen];

					// Propagate the hole to the leafs of the heap
					int ppos = 0;
					int path = 1;

					while (path < heapLen)
					{
						if (path + 1 < heapLen && values[heap[path]] > values[heap[path + 1]])
						{
							path++;
						}

						heap[ppos] = heap[path];
						ppos = path;
						path = path * 2 + 1;
					}

					/* Now propagate the last element down along path.  Normally
					* it shouldn't go too deep.
					*/
					int lastVal = values[last];
					while ((path = ppos) > 0 && values[heap[ppos = (path - 1) / 2]] > lastVal)
					{
						heap[path] = heap[ppos];
					}
					heap[path] = last;

					int second = heap[0];

					// Create a new node father of first and second
					last = numNodes++;
					childs[2 * last] = first;
					childs[2 * last + 1] = second;
					int mindepth = Math.Min(values[first] & 0xff, values[second] & 0xff);
					values[last] = lastVal = values[first] + values[second] - mindepth + 1;

					// Again, propagate the hole to the leafs
					ppos = 0;
					path = 1;

					while (path < heapLen)
					{
						if (path + 1 < heapLen && values[heap[path]] > values[heap[path + 1]])
						{
							path++;
						}

						heap[ppos] = heap[path];
						ppos = path;
						path = ppos * 2 + 1;
					}

					// Now propagate the new element down along path
					while ((path = ppos) > 0 && values[heap[ppos = (path - 1) / 2]] > lastVal)
					{
						heap[path] = heap[ppos];
					}
					heap[path] = last;
				} while (heapLen > 1);

				if (heap[0] != childs.Length / 2 - 1)
				{
					throw new SharpZipBaseException("Heap invariant violated");
				}

				BuildLength(childs);
			}

			/// <summary>
			/// Get encoded length
			/// </summary>
			/// <returns>Encoded length, the sum of frequencies * lengths</returns>
			public int GetEncodedLength()
			{
				int len = 0;
				for (int i = 0; i < freqs.Length; i++)
				{
					len += freqs[i] * length[i];
				}
				return len;
			}

			/// <summary>
			/// Scan a literal or distance tree to determine the frequencies of the codes
			/// in the bit length tree.
			/// </summary>
			public void CalcBLFreq(Tree blTree)
			{
				int max_count;               /* max repeat count */
				int min_count;               /* min repeat count */
				int count;                   /* repeat count of the current code */
				int curlen = -1;             /* length of current code */

				int i = 0;
				while (i < numCodes)
				{
					count = 1;
					int nextlen = length[i];
					if (nextlen == 0)
					{
						max_count = 138;
						min_count = 3;
					}
					else
					{
						max_count = 6;
						min_count = 3;
						if (curlen != nextlen)
						{
							blTree.freqs[nextlen]++;
							count = 0;
						}
					}
					curlen = nextlen;
					i++;

					while (i < numCodes && curlen == length[i])
					{
						i++;
						if (++count >= max_count)
						{
							break;
						}
					}

					if (count < min_count)
					{
						blTree.freqs[curlen] += (short)count;
					}
					else if (curlen != 0)
					{
						blTree.freqs[REP_3_6]++;
					}
					else if (count <= 10)
					{
						blTree.freqs[REP_3_10]++;
					}
					else
					{
						blTree.freqs[REP_11_138]++;
					}
				}
			}

			/// <summary>
			/// Write tree values
			/// </summary>
			/// <param name="blTree">Tree to write</param>
			public void WriteTree(Tree blTree)
			{
				int max_count;               // max repeat count
				int min_count;               // min repeat count
				int count;                   // repeat count of the current code
				int curlen = -1;             // length of current code

				int i = 0;
				while (i < numCodes)
				{
					count = 1;
					int nextlen = length[i];
					if (nextlen == 0)
					{
						max_count = 138;
						min_count = 3;
					}
					else
					{
						max_count = 6;
						min_count = 3;
						if (curlen != nextlen)
						{
							blTree.WriteSymbol(nextlen);
							count = 0;
						}
					}
					curlen = nextlen;
					i++;

					while (i < numCodes && curlen == length[i])
					{
						i++;
						if (++count >= max_count)
						{
							break;
						}
					}

					if (count < min_count)
					{
						while (count-- > 0)
						{
							blTree.WriteSymbol(curlen);
						}
					}
					else if (curlen != 0)
					{
						blTree.WriteSymbol(REP_3_6);
						dh.pending.WriteBits(count - 3, 2);
					}
					else if (count <= 10)
					{
						blTree.WriteSymbol(REP_3_10);
						dh.pending.WriteBits(count - 3, 3);
					}
					else
					{
						blTree.WriteSymbol(REP_11_138);
						dh.pending.WriteBits(count - 11, 7);
					}
				}
			}

			private void BuildLength(int[] childs)
			{
				this.length = new byte[freqs.Length];
				int numNodes = childs.Length / 2;
				int numLeafs = (numNodes + 1) / 2;
				int overflow = 0;

				for (int i = 0; i < maxLength; i++)
				{
					bl_counts[i] = 0;
				}

				// First calculate optimal bit lengths
				int[] lengths = new int[numNodes];
				lengths[numNodes - 1] = 0;

				for (int i = numNodes - 1; i >= 0; i--)
				{
					if (childs[2 * i + 1] != -1)
					{
						int bitLength = lengths[i] + 1;
						if (bitLength > maxLength)
						{
							bitLength = maxLength;
							overflow++;
						}
						lengths[childs[2 * i]] = lengths[childs[2 * i + 1]] = bitLength;
					}
					else
					{
						// A leaf node
						int bitLength = lengths[i];
						bl_counts[bitLength - 1]++;
						this.length[childs[2 * i]] = (byte)lengths[i];
					}
				}

				//				if (DeflaterConstants.DEBUGGING) {
				//					//Console.WriteLine("Tree "+freqs.Length+" lengths:");
				//					for (int i=0; i < numLeafs; i++) {
				//						//Console.WriteLine("Node "+childs[2*i]+" freq: "+freqs[childs[2*i]]
				//						                  + " len: "+length[childs[2*i]]);
				//					}
				//				}

				if (overflow == 0)
				{
					return;
				}

				int incrBitLen = maxLength - 1;
				do
				{
					// Find the first bit length which could increase:
					while (bl_counts[--incrBitLen] == 0)
					{
					}

					// Move this node one down and remove a corresponding
					// number of overflow nodes.
					do
					{
						bl_counts[incrBitLen]--;
						bl_counts[++incrBitLen]++;
						overflow -= 1 << (maxLength - 1 - incrBitLen);
					} while (overflow > 0 && incrBitLen < maxLength - 1);
				} while (overflow > 0);

				/* We may have overshot above.  Move some nodes from maxLength to
				* maxLength-1 in that case.
				*/
				bl_counts[maxLength - 1] += overflow;
				bl_counts[maxLength - 2] -= overflow;

				/* Now recompute all bit lengths, scanning in increasing
				* frequency.  It is simpler to reconstruct all lengths instead of
				* fixing only the wrong ones. This idea is taken from 'ar'
				* written by Haruhiko Okumura.
				*
				* The nodes were inserted with decreasing frequency into the childs
				* array.
				*/
				int nodePtr = 2 * numLeafs;
				for (int bits = maxLength; bits != 0; bits--)
				{
					int n = bl_counts[bits - 1];
					while (n > 0)
					{
						int childPtr = 2 * childs[nodePtr++];
						if (childs[childPtr + 1] == -1)
						{
							// We found another leaf
							length[childs[childPtr]] = (byte)bits;
							n--;
						}
					}
				}
				//				if (DeflaterConstants.DEBUGGING) {
				//					//Console.WriteLine("*** After overflow elimination. ***");
				//					for (int i=0; i < numLeafs; i++) {
				//						//Console.WriteLine("Node "+childs[2*i]+" freq: "+freqs[childs[2*i]]
				//						                  + " len: "+length[childs[2*i]]);
				//					}
				//				}
			}
		}

		#region Instance Fields

		/// <summary>
		/// Pending buffer to use
		/// </summary>
		public DeflaterPending pending;

		private Tree literalTree;
		private Tree distTree;
		private Tree blTree;

		// Buffer for distances
		private short[] d_buf;

		private byte[] l_buf;
		private int last_lit;
		private int extra_bits;

		#endregion Instance Fields

		static DeflaterHuffman()
		{
			// See RFC 1951 3.2.6
			// Literal codes
			staticLCodes = new short[LITERAL_NUM];
			staticLLength = new byte[LITERAL_NUM];

			int i = 0;
			while (i < 144)
			{
				staticLCodes[i] = BitReverse((0x030 + i) << 8);
				staticLLength[i++] = 8;
			}

			while (i < 256)
			{
				staticLCodes[i] = BitReverse((0x190 - 144 + i) << 7);
				staticLLength[i++] = 9;
			}

			while (i < 280)
			{
				staticLCodes[i] = BitReverse((0x000 - 256 + i) << 9);
				staticLLength[i++] = 7;
			}

			while (i < LITERAL_NUM)
			{
				staticLCodes[i] = BitReverse((0x0c0 - 280 + i) << 8);
				staticLLength[i++] = 8;
			}

			// Distance codes
			staticDCodes = new short[DIST_NUM];
			staticDLength = new byte[DIST_NUM];
			for (i = 0; i < DIST_NUM; i++)
			{
				staticDCodes[i] = BitReverse(i << 11);
				staticDLength[i] = 5;
			}
		}

		/// <summary>
		/// Construct instance with pending buffer
		/// </summary>
		/// <param name="pending">Pending buffer to use</param>
		public DeflaterHuffman(DeflaterPending pending)
		{
			this.pending = pending;

			literalTree = new Tree(this, LITERAL_NUM, 257, 15);
			distTree = new Tree(this, DIST_NUM, 1, 15);
			blTree = new Tree(this, BITLEN_NUM, 4, 7);

			d_buf = new short[BUFSIZE];
			l_buf = new byte[BUFSIZE];
		}

		/// <summary>
		/// Reset internal state
		/// </summary>
		public void Reset()
		{
			last_lit = 0;
			extra_bits = 0;
			literalTree.Reset();
			distTree.Reset();
			blTree.Reset();
		}

		/// <summary>
		/// Write all trees to pending buffer
		/// </summary>
		/// <param name="blTreeCodes">The number/rank of treecodes to send.</param>
		public void SendAllTrees(int blTreeCodes)
		{
			blTree.BuildCodes();
			literalTree.BuildCodes();
			distTree.BuildCodes();
			pending.WriteBits(literalTree.numCodes - 257, 5);
			pending.WriteBits(distTree.numCodes - 1, 5);
			pending.WriteBits(blTreeCodes - 4, 4);
			for (int rank = 0; rank < blTreeCodes; rank++)
			{
				pending.WriteBits(blTree.length[BL_ORDER[rank]], 3);
			}
			literalTree.WriteTree(blTree);
			distTree.WriteTree(blTree);

#if DebugDeflation
			if (DeflaterConstants.DEBUGGING) {
				blTree.CheckEmpty();
			}
#endif
		}

		/// <summary>
		/// Compress current buffer writing data to pending buffer
		/// </summary>
		public void CompressBlock()
		{
			for (int i = 0; i < last_lit; i++)
			{
				int litlen = l_buf[i] & 0xff;
				int dist = d_buf[i];
				if (dist-- != 0)
				{
					//					if (DeflaterConstants.DEBUGGING) {
					//						Console.Write("["+(dist+1)+","+(litlen+3)+"]: ");
					//					}

					int lc = Lcode(litlen);
					literalTree.WriteSymbol(lc);

					int bits = (lc - 261) / 4;
					if (bits > 0 && bits <= 5)
					{
						pending.WriteBits(litlen & ((1 << bits) - 1), bits);
					}

					int dc = Dcode(dist);
					distTree.WriteSymbol(dc);

					bits = dc / 2 - 1;
					if (bits > 0)
					{
						pending.WriteBits(dist & ((1 << bits) - 1), bits);
					}
				}
				else
				{
					//					if (DeflaterConstants.DEBUGGING) {
					//						if (litlen > 32 && litlen < 127) {
					//							Console.Write("("+(char)litlen+"): ");
					//						} else {
					//							Console.Write("{"+litlen+"}: ");
					//						}
					//					}
					literalTree.WriteSymbol(litlen);
				}
			}

#if DebugDeflation
			if (DeflaterConstants.DEBUGGING) {
				Console.Write("EOF: ");
			}
#endif
			literalTree.WriteSymbol(EOF_SYMBOL);

#if DebugDeflation
			if (DeflaterConstants.DEBUGGING) {
				literalTree.CheckEmpty();
				distTree.CheckEmpty();
			}
#endif
		}

		/// <summary>
		/// Flush block to output with no compression
		/// </summary>
		/// <param name="stored">Data to write</param>
		/// <param name="storedOffset">Index of first byte to write</param>
		/// <param name="storedLength">Count of bytes to write</param>
		/// <param name="lastBlock">True if this is the last block</param>
		public void FlushStoredBlock(byte[] stored, int storedOffset, int storedLength, bool lastBlock)
		{
#if DebugDeflation
			//			if (DeflaterConstants.DEBUGGING) {
			//				//Console.WriteLine("Flushing stored block "+ storedLength);
			//			}
#endif
			pending.WriteBits((DeflaterConstants.STORED_BLOCK << 1) + (lastBlock ? 1 : 0), 3);
			pending.AlignToByte();
			pending.WriteShort(storedLength);
			pending.WriteShort(~storedLength);
			pending.WriteBlock(stored, storedOffset, storedLength);
			Reset();
		}

		/// <summary>
		/// Flush block to output with compression
		/// </summary>
		/// <param name="stored">Data to flush</param>
		/// <param name="storedOffset">Index of first byte to flush</param>
		/// <param name="storedLength">Count of bytes to flush</param>
		/// <param name="lastBlock">True if this is the last block</param>
		public void FlushBlock(byte[] stored, int storedOffset, int storedLength, bool lastBlock)
		{
			literalTree.freqs[EOF_SYMBOL]++;

			// Build trees
			literalTree.BuildTree();
			distTree.BuildTree();

			// Calculate bitlen frequency
			literalTree.CalcBLFreq(blTree);
			distTree.CalcBLFreq(blTree);

			// Build bitlen tree
			blTree.BuildTree();

			int blTreeCodes = 4;
			for (int i = 18; i > blTreeCodes; i--)
			{
				if (blTree.length[BL_ORDER[i]] > 0)
				{
					blTreeCodes = i + 1;
				}
			}
			int opt_len = 14 + blTreeCodes * 3 + blTree.GetEncodedLength() +
				literalTree.GetEncodedLength() + distTree.GetEncodedLength() +
				extra_bits;

			int static_len = extra_bits;
			for (int i = 0; i < LITERAL_NUM; i++)
			{
				static_len += literalTree.freqs[i] * staticLLength[i];
			}
			for (int i = 0; i < DIST_NUM; i++)
			{
				static_len += distTree.freqs[i] * staticDLength[i];
			}
			if (opt_len >= static_len)
			{
				// Force static trees
				opt_len = static_len;
			}

			if (storedOffset >= 0 && storedLength + 4 < opt_len >> 3)
			{
				// Store Block

				//				if (DeflaterConstants.DEBUGGING) {
				//					//Console.WriteLine("Storing, since " + storedLength + " < " + opt_len
				//					                  + " <= " + static_len);
				//				}
				FlushStoredBlock(stored, storedOffset, storedLength, lastBlock);
			}
			else if (opt_len == static_len)
			{
				// Encode with static tree
				pending.WriteBits((DeflaterConstants.STATIC_TREES << 1) + (lastBlock ? 1 : 0), 3);
				literalTree.SetStaticCodes(staticLCodes, staticLLength);
				distTree.SetStaticCodes(staticDCodes, staticDLength);
				CompressBlock();
				Reset();
			}
			else
			{
				// Encode with dynamic tree
				pending.WriteBits((DeflaterConstants.DYN_TREES << 1) + (lastBlock ? 1 : 0), 3);
				SendAllTrees(blTreeCodes);
				CompressBlock();
				Reset();
			}
		}

		/// <summary>
		/// Get value indicating if internal buffer is full
		/// </summary>
		/// <returns>true if buffer is full</returns>
		public bool IsFull()
		{
			return last_lit >= BUFSIZE;
		}

		/// <summary>
		/// Add literal to buffer
		/// </summary>
		/// <param name="literal">Literal value to add to buffer.</param>
		/// <returns>Value indicating internal buffer is full</returns>
		public bool TallyLit(int literal)
		{
			//			if (DeflaterConstants.DEBUGGING) {
			//				if (lit > 32 && lit < 127) {
			//					//Console.WriteLine("("+(char)lit+")");
			//				} else {
			//					//Console.WriteLine("{"+lit+"}");
			//				}
			//			}
			d_buf[last_lit] = 0;
			l_buf[last_lit++] = (byte)literal;
			literalTree.freqs[literal]++;
			return IsFull();
		}

		/// <summary>
		/// Add distance code and length to literal and distance trees
		/// </summary>
		/// <param name="distance">Distance code</param>
		/// <param name="length">Length</param>
		/// <returns>Value indicating if internal buffer is full</returns>
		public bool TallyDist(int distance, int length)
		{
			//			if (DeflaterConstants.DEBUGGING) {
			//				//Console.WriteLine("[" + distance + "," + length + "]");
			//			}

			d_buf[last_lit] = (short)distance;
			l_buf[last_lit++] = (byte)(length - 3);

			int lc = Lcode(length - 3);
			literalTree.freqs[lc]++;
			if (lc >= 265 && lc < 285)
			{
				extra_bits += (lc - 261) / 4;
			}

			int dc = Dcode(distance - 1);
			distTree.freqs[dc]++;
			if (dc >= 4)
			{
				extra_bits += dc / 2 - 1;
			}
			return IsFull();
		}

		/// <summary>
		/// Reverse the bits of a 16 bit value.
		/// </summary>
		/// <param name="toReverse">Value to reverse bits</param>
		/// <returns>Value with bits reversed</returns>
		public static short BitReverse(int toReverse)
		{
			return (short)(bit4Reverse[toReverse & 0xF] << 12 |
							bit4Reverse[(toReverse >> 4) & 0xF] << 8 |
							bit4Reverse[(toReverse >> 8) & 0xF] << 4 |
							bit4Reverse[toReverse >> 12]);
		}

		private static int Lcode(int length)
		{
			if (length == 255)
			{
				return 285;
			}

			int code = 257;
			while (length >= 8)
			{
				code += 4;
				length >>= 1;
			}
			return code + length;
		}

		private static int Dcode(int distance)
		{
			int code = 0;
			while (distance >= 4)
			{
				code += 2;
				distance >>= 1;
			}
			return code + distance;
		}
	}
}

namespace ICSharpCode.SharpZipLib.Zip.Compression
{
	/// <summary>
	/// This class stores the pending output of the Deflater.
	///
	/// author of the original java version : Jochen Hoenicke
	/// </summary>
	public class DeflaterPending : PendingBuffer
	{
		/// <summary>
		/// Construct instance with default buffer size
		/// </summary>
		public DeflaterPending() : base(DeflaterConstants.PENDING_BUF_SIZE)
		{
		}
	}
}

namespace ICSharpCode.SharpZipLib.Zip.Compression
{
	/// <summary>
	/// Inflater is used to decompress data that has been compressed according
	/// to the "deflate" standard described in rfc1951.
	///
	/// By default Zlib (rfc1950) headers and footers are expected in the input.
	/// You can use constructor <code> public Inflater(bool noHeader)</code> passing true
	/// if there is no Zlib header information
	///
	/// The usage is as following.  First you have to set some input with
	/// <code>SetInput()</code>, then Inflate() it.  If inflate doesn't
	/// inflate any bytes there may be three reasons:
	/// <ul>
	/// <li>IsNeedingInput() returns true because the input buffer is empty.
	/// You have to provide more input with <code>SetInput()</code>.
	/// NOTE: IsNeedingInput() also returns true when, the stream is finished.
	/// </li>
	/// <li>IsNeedingDictionary() returns true, you have to provide a preset
	///    dictionary with <code>SetDictionary()</code>.</li>
	/// <li>IsFinished returns true, the inflater has finished.</li>
	/// </ul>
	/// Once the first output byte is produced, a dictionary will not be
	/// needed at a later stage.
	///
	/// author of the original java version : John Leuner, Jochen Hoenicke
	/// </summary>
	public class Inflater
	{
		#region Constants/Readonly

		/// <summary>
		/// Copy lengths for literal codes 257..285
		/// </summary>
		private static readonly int[] CPLENS = {
								  3, 4, 5, 6, 7, 8, 9, 10, 11, 13, 15, 17, 19, 23, 27, 31,
								  35, 43, 51, 59, 67, 83, 99, 115, 131, 163, 195, 227, 258
							  };

		/// <summary>
		/// Extra bits for literal codes 257..285
		/// </summary>
		private static readonly int[] CPLEXT = {
								  0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 2, 2, 2, 2,
								  3, 3, 3, 3, 4, 4, 4, 4, 5, 5, 5, 5, 0
							  };

		/// <summary>
		/// Copy offsets for distance codes 0..29
		/// </summary>
		private static readonly int[] CPDIST = {
								1, 2, 3, 4, 5, 7, 9, 13, 17, 25, 33, 49, 65, 97, 129, 193,
								257, 385, 513, 769, 1025, 1537, 2049, 3073, 4097, 6145,
								8193, 12289, 16385, 24577
							  };

		/// <summary>
		/// Extra bits for distance codes
		/// </summary>
		private static readonly int[] CPDEXT = {
								0, 0, 0, 0, 1, 1, 2, 2, 3, 3, 4, 4, 5, 5, 6, 6,
								7, 7, 8, 8, 9, 9, 10, 10, 11, 11,
								12, 12, 13, 13
							  };

		/// <summary>
		/// These are the possible states for an inflater
		/// </summary>
		private const int DECODE_HEADER = 0;

		private const int DECODE_DICT = 1;
		private const int DECODE_BLOCKS = 2;
		private const int DECODE_STORED_LEN1 = 3;
		private const int DECODE_STORED_LEN2 = 4;
		private const int DECODE_STORED = 5;
		private const int DECODE_DYN_HEADER = 6;
		private const int DECODE_HUFFMAN = 7;
		private const int DECODE_HUFFMAN_LENBITS = 8;
		private const int DECODE_HUFFMAN_DIST = 9;
		private const int DECODE_HUFFMAN_DISTBITS = 10;
		private const int DECODE_CHKSUM = 11;
		private const int FINISHED = 12;

		#endregion Constants/Readonly

		#region Instance Fields

		/// <summary>
		/// This variable contains the current state.
		/// </summary>
		private int mode;

		/// <summary>
		/// The adler checksum of the dictionary or of the decompressed
		/// stream, as it is written in the header resp. footer of the
		/// compressed stream.
		/// Only valid if mode is DECODE_DICT or DECODE_CHKSUM.
		/// </summary>
		private int readAdler;

		/// <summary>
		/// The number of bits needed to complete the current state.  This
		/// is valid, if mode is DECODE_DICT, DECODE_CHKSUM,
		/// DECODE_HUFFMAN_LENBITS or DECODE_HUFFMAN_DISTBITS.
		/// </summary>
		private int neededBits;

		private int repLength;
		private int repDist;
		private int uncomprLen;

		/// <summary>
		/// True, if the last block flag was set in the last block of the
		/// inflated stream.  This means that the stream ends after the
		/// current block.
		/// </summary>
		private bool isLastBlock;

		/// <summary>
		/// The total number of inflated bytes.
		/// </summary>
		private long totalOut;

		/// <summary>
		/// The total number of bytes set with setInput().  This is not the
		/// value returned by the TotalIn property, since this also includes the
		/// unprocessed input.
		/// </summary>
		private long totalIn;

		/// <summary>
		/// This variable stores the noHeader flag that was given to the constructor.
		/// True means, that the inflated stream doesn't contain a Zlib header or
		/// footer.
		/// </summary>
		private bool noHeader;

		private readonly StreamManipulator input;
		private OutputWindow outputWindow;
		private InflaterDynHeader dynHeader;
		private InflaterHuffmanTree litlenTree, distTree;
		private Adler32 adler;

		#endregion Instance Fields

		#region Constructors

		/// <summary>
		/// Creates a new inflater or RFC1951 decompressor
		/// RFC1950/Zlib headers and footers will be expected in the input data
		/// </summary>
		public Inflater() : this(false)
		{
		}

		/// <summary>
		/// Creates a new inflater.
		/// </summary>
		/// <param name="noHeader">
		/// True if no RFC1950/Zlib header and footer fields are expected in the input data
		///
		/// This is used for GZIPed/Zipped input.
		///
		/// For compatibility with
		/// Sun JDK you should provide one byte of input more than needed in
		/// this case.
		/// </param>
		public Inflater(bool noHeader)
		{
			this.noHeader = noHeader;
			if (!noHeader)
				this.adler = new Adler32();
			input = new StreamManipulator();
			outputWindow = new OutputWindow();
			mode = noHeader ? DECODE_BLOCKS : DECODE_HEADER;
		}

		#endregion Constructors

		/// <summary>
		/// Resets the inflater so that a new stream can be decompressed.  All
		/// pending input and output will be discarded.
		/// </summary>
		public void Reset()
		{
			mode = noHeader ? DECODE_BLOCKS : DECODE_HEADER;
			totalIn = 0;
			totalOut = 0;
			input.Reset();
			outputWindow.Reset();
			dynHeader = null;
			litlenTree = null;
			distTree = null;
			isLastBlock = false;
			adler?.Reset();
		}

		/// <summary>
		/// Decodes a zlib/RFC1950 header.
		/// </summary>
		/// <returns>
		/// False if more input is needed.
		/// </returns>
		/// <exception cref="SharpZipBaseException">
		/// The header is invalid.
		/// </exception>
		private bool DecodeHeader()
		{
			int header = input.PeekBits(16);
			if (header < 0)
			{
				return false;
			}
			input.DropBits(16);

			// The header is written in "wrong" byte order
			header = ((header << 8) | (header >> 8)) & 0xffff;
			if (header % 31 != 0)
			{
				throw new SharpZipBaseException("Header checksum illegal");
			}

			if ((header & 0x0f00) != (Deflater.DEFLATED << 8))
			{
				throw new SharpZipBaseException("Compression Method unknown");
			}

			/* Maximum size of the backwards window in bits.
			* We currently ignore this, but we could use it to make the
			* inflater window more space efficient. On the other hand the
			* full window (15 bits) is needed most times, anyway.
			int max_wbits = ((header & 0x7000) >> 12) + 8;
			*/

			if ((header & 0x0020) == 0)
			{ // Dictionary flag?
				mode = DECODE_BLOCKS;
			}
			else
			{
				mode = DECODE_DICT;
				neededBits = 32;
			}
			return true;
		}

		/// <summary>
		/// Decodes the dictionary checksum after the deflate header.
		/// </summary>
		/// <returns>
		/// False if more input is needed.
		/// </returns>
		private bool DecodeDict()
		{
			while (neededBits > 0)
			{
				int dictByte = input.PeekBits(8);
				if (dictByte < 0)
				{
					return false;
				}
				input.DropBits(8);
				readAdler = (readAdler << 8) | dictByte;
				neededBits -= 8;
			}
			return false;
		}

		/// <summary>
		/// Decodes the huffman encoded symbols in the input stream.
		/// </summary>
		/// <returns>
		/// false if more input is needed, true if output window is
		/// full or the current block ends.
		/// </returns>
		/// <exception cref="SharpZipBaseException">
		/// if deflated stream is invalid.
		/// </exception>
		private bool DecodeHuffman()
		{
			int free = outputWindow.GetFreeSpace();
			while (free >= 258)
			{
				int symbol;
				switch (mode)
				{
					case DECODE_HUFFMAN:
						// This is the inner loop so it is optimized a bit
						while (((symbol = litlenTree.GetSymbol(input)) & ~0xff) == 0)
						{
							outputWindow.Write(symbol);
							if (--free < 258)
							{
								return true;
							}
						}

						if (symbol < 257)
						{
							if (symbol < 0)
							{
								return false;
							}
							else
							{
								// symbol == 256: end of block
								distTree = null;
								litlenTree = null;
								mode = DECODE_BLOCKS;
								return true;
							}
						}

						try
						{
							repLength = CPLENS[symbol - 257];
							neededBits = CPLEXT[symbol - 257];
						}
						catch (Exception)
						{
							throw new SharpZipBaseException("Illegal rep length code");
						}
						goto case DECODE_HUFFMAN_LENBITS; // fall through

					case DECODE_HUFFMAN_LENBITS:
						if (neededBits > 0)
						{
							mode = DECODE_HUFFMAN_LENBITS;
							int i = input.PeekBits(neededBits);
							if (i < 0)
							{
								return false;
							}
							input.DropBits(neededBits);
							repLength += i;
						}
						mode = DECODE_HUFFMAN_DIST;
						goto case DECODE_HUFFMAN_DIST; // fall through

					case DECODE_HUFFMAN_DIST:
						symbol = distTree.GetSymbol(input);
						if (symbol < 0)
						{
							return false;
						}

						try
						{
							repDist = CPDIST[symbol];
							neededBits = CPDEXT[symbol];
						}
						catch (Exception)
						{
							throw new SharpZipBaseException("Illegal rep dist code");
						}

						goto case DECODE_HUFFMAN_DISTBITS; // fall through

					case DECODE_HUFFMAN_DISTBITS:
						if (neededBits > 0)
						{
							mode = DECODE_HUFFMAN_DISTBITS;
							int i = input.PeekBits(neededBits);
							if (i < 0)
							{
								return false;
							}
							input.DropBits(neededBits);
							repDist += i;
						}

						outputWindow.Repeat(repLength, repDist);
						free -= repLength;
						mode = DECODE_HUFFMAN;
						break;

					default:
						throw new SharpZipBaseException("Inflater unknown mode");
				}
			}
			return true;
		}

		/// <summary>
		/// Decodes the adler checksum after the deflate stream.
		/// </summary>
		/// <returns>
		/// false if more input is needed.
		/// </returns>
		/// <exception cref="SharpZipBaseException">
		/// If checksum doesn't match.
		/// </exception>
		private bool DecodeChksum()
		{
			while (neededBits > 0)
			{
				int chkByte = input.PeekBits(8);
				if (chkByte < 0)
				{
					return false;
				}
				input.DropBits(8);
				readAdler = (readAdler << 8) | chkByte;
				neededBits -= 8;
			}

			if ((int)adler?.Value != readAdler)
			{
				throw new SharpZipBaseException("Adler chksum doesn't match: " + (int)adler?.Value + " vs. " + readAdler);
			}

			mode = FINISHED;
			return false;
		}

		/// <summary>
		/// Decodes the deflated stream.
		/// </summary>
		/// <returns>
		/// false if more input is needed, or if finished.
		/// </returns>
		/// <exception cref="SharpZipBaseException">
		/// if deflated stream is invalid.
		/// </exception>
		private bool Decode()
		{
			switch (mode)
			{
				case DECODE_HEADER:
					return DecodeHeader();

				case DECODE_DICT:
					return DecodeDict();

				case DECODE_CHKSUM:
					return DecodeChksum();

				case DECODE_BLOCKS:
					if (isLastBlock)
					{
						if (noHeader)
						{
							mode = FINISHED;
							return false;
						}
						else
						{
							input.SkipToByteBoundary();
							neededBits = 32;
							mode = DECODE_CHKSUM;
							return true;
						}
					}

					int type = input.PeekBits(3);
					if (type < 0)
					{
						return false;
					}
					input.DropBits(3);

					isLastBlock |= (type & 1) != 0;
					switch (type >> 1)
					{
						case DeflaterConstants.STORED_BLOCK:
							input.SkipToByteBoundary();
							mode = DECODE_STORED_LEN1;
							break;

						case DeflaterConstants.STATIC_TREES:
							litlenTree = InflaterHuffmanTree.defLitLenTree;
							distTree = InflaterHuffmanTree.defDistTree;
							mode = DECODE_HUFFMAN;
							break;

						case DeflaterConstants.DYN_TREES:
							dynHeader = new InflaterDynHeader(input);
							mode = DECODE_DYN_HEADER;
							break;

						default:
							throw new SharpZipBaseException("Unknown block type " + type);
					}
					return true;

				case DECODE_STORED_LEN1:
					{
						if ((uncomprLen = input.PeekBits(16)) < 0)
						{
							return false;
						}
						input.DropBits(16);
						mode = DECODE_STORED_LEN2;
					}
					goto case DECODE_STORED_LEN2; // fall through

				case DECODE_STORED_LEN2:
					{
						int nlen = input.PeekBits(16);
						if (nlen < 0)
						{
							return false;
						}
						input.DropBits(16);
						if (nlen != (uncomprLen ^ 0xffff))
						{
							throw new SharpZipBaseException("broken uncompressed block");
						}
						mode = DECODE_STORED;
					}
					goto case DECODE_STORED; // fall through

				case DECODE_STORED:
					{
						int more = outputWindow.CopyStored(input, uncomprLen);
						uncomprLen -= more;
						if (uncomprLen == 0)
						{
							mode = DECODE_BLOCKS;
							return true;
						}
						return !input.IsNeedingInput;
					}

				case DECODE_DYN_HEADER:
					if (!dynHeader.AttemptRead())
					{
						return false;
					}

					litlenTree = dynHeader.LiteralLengthTree;
					distTree = dynHeader.DistanceTree;
					mode = DECODE_HUFFMAN;
					goto case DECODE_HUFFMAN; // fall through

				case DECODE_HUFFMAN:
				case DECODE_HUFFMAN_LENBITS:
				case DECODE_HUFFMAN_DIST:
				case DECODE_HUFFMAN_DISTBITS:
					return DecodeHuffman();

				case FINISHED:
					return false;

				default:
					throw new SharpZipBaseException("Inflater.Decode unknown mode");
			}
		}

		/// <summary>
		/// Sets the preset dictionary.  This should only be called, if
		/// needsDictionary() returns true and it should set the same
		/// dictionary, that was used for deflating.  The getAdler()
		/// function returns the checksum of the dictionary needed.
		/// </summary>
		/// <param name="buffer">
		/// The dictionary.
		/// </param>
		public void SetDictionary(byte[] buffer)
		{
			SetDictionary(buffer, 0, buffer.Length);
		}

		/// <summary>
		/// Sets the preset dictionary.  This should only be called, if
		/// needsDictionary() returns true and it should set the same
		/// dictionary, that was used for deflating.  The getAdler()
		/// function returns the checksum of the dictionary needed.
		/// </summary>
		/// <param name="buffer">
		/// The dictionary.
		/// </param>
		/// <param name="index">
		/// The index into buffer where the dictionary starts.
		/// </param>
		/// <param name="count">
		/// The number of bytes in the dictionary.
		/// </param>
		/// <exception cref="System.InvalidOperationException">
		/// No dictionary is needed.
		/// </exception>
		/// <exception cref="SharpZipBaseException">
		/// The adler checksum for the buffer is invalid
		/// </exception>
		public void SetDictionary(byte[] buffer, int index, int count)
		{
			if (buffer == null)
			{
				throw new ArgumentNullException(nameof(buffer));
			}

			if (index < 0)
			{
				throw new ArgumentOutOfRangeException(nameof(index));
			}

			if (count < 0)
			{
				throw new ArgumentOutOfRangeException(nameof(count));
			}

			if (!IsNeedingDictionary)
			{
				throw new InvalidOperationException("Dictionary is not needed");
			}

			adler?.Update(new ArraySegment<byte>(buffer, index, count));

			if (adler != null && (int)adler.Value != readAdler)
			{
				throw new SharpZipBaseException("Wrong adler checksum");
			}
			adler?.Reset();
			outputWindow.CopyDict(buffer, index, count);
			mode = DECODE_BLOCKS;
		}

		/// <summary>
		/// Sets the input.  This should only be called, if needsInput()
		/// returns true.
		/// </summary>
		/// <param name="buffer">
		/// the input.
		/// </param>
		public void SetInput(byte[] buffer)
		{
			SetInput(buffer, 0, buffer.Length);
		}

		/// <summary>
		/// Sets the input.  This should only be called, if needsInput()
		/// returns true.
		/// </summary>
		/// <param name="buffer">
		/// The source of input data
		/// </param>
		/// <param name="index">
		/// The index into buffer where the input starts.
		/// </param>
		/// <param name="count">
		/// The number of bytes of input to use.
		/// </param>
		/// <exception cref="System.InvalidOperationException">
		/// No input is needed.
		/// </exception>
		/// <exception cref="System.ArgumentOutOfRangeException">
		/// The index and/or count are wrong.
		/// </exception>
		public void SetInput(byte[] buffer, int index, int count)
		{
			input.SetInput(buffer, index, count);
			totalIn += (long)count;
		}

		/// <summary>
		/// Inflates the compressed stream to the output buffer.  If this
		/// returns 0, you should check, whether IsNeedingDictionary(),
		/// IsNeedingInput() or IsFinished() returns true, to determine why no
		/// further output is produced.
		/// </summary>
		/// <param name="buffer">
		/// the output buffer.
		/// </param>
		/// <returns>
		/// The number of bytes written to the buffer, 0 if no further
		/// output can be produced.
		/// </returns>
		/// <exception cref="System.ArgumentOutOfRangeException">
		/// if buffer has length 0.
		/// </exception>
		/// <exception cref="System.FormatException">
		/// if deflated stream is invalid.
		/// </exception>
		public int Inflate(byte[] buffer)
		{
			if (buffer == null)
			{
				throw new ArgumentNullException(nameof(buffer));
			}

			return Inflate(buffer, 0, buffer.Length);
		}

		/// <summary>
		/// Inflates the compressed stream to the output buffer.  If this
		/// returns 0, you should check, whether needsDictionary(),
		/// needsInput() or finished() returns true, to determine why no
		/// further output is produced.
		/// </summary>
		/// <param name="buffer">
		/// the output buffer.
		/// </param>
		/// <param name="offset">
		/// the offset in buffer where storing starts.
		/// </param>
		/// <param name="count">
		/// the maximum number of bytes to output.
		/// </param>
		/// <returns>
		/// the number of bytes written to the buffer, 0 if no further output can be produced.
		/// </returns>
		/// <exception cref="System.ArgumentOutOfRangeException">
		/// if count is less than 0.
		/// </exception>
		/// <exception cref="System.ArgumentOutOfRangeException">
		/// if the index and / or count are wrong.
		/// </exception>
		/// <exception cref="System.FormatException">
		/// if deflated stream is invalid.
		/// </exception>
		public int Inflate(byte[] buffer, int offset, int count)
		{
			if (buffer == null)
			{
				throw new ArgumentNullException(nameof(buffer));
			}

			if (count < 0)
			{
				throw new ArgumentOutOfRangeException(nameof(count), "count cannot be negative");
			}

			if (offset < 0)
			{
				throw new ArgumentOutOfRangeException(nameof(offset), "offset cannot be negative");
			}

			if (offset + count > buffer.Length)
			{
				throw new ArgumentException("count exceeds buffer bounds");
			}

			// Special case: count may be zero
			if (count == 0)
			{
				if (!IsFinished)
				{ // -jr- 08-Nov-2003 INFLATE_BUG fix..
					Decode();
				}
				return 0;
			}

			int bytesCopied = 0;

			do
			{
				if (mode != DECODE_CHKSUM)
				{
					/* Don't give away any output, if we are waiting for the
					* checksum in the input stream.
					*
					* With this trick we have always:
					*   IsNeedingInput() and not IsFinished()
					*   implies more output can be produced.
					*/
					int more = outputWindow.CopyOutput(buffer, offset, count);
					if (more > 0)
					{
						adler?.Update(new ArraySegment<byte>(buffer, offset, more));
						offset += more;
						bytesCopied += more;
						totalOut += (long)more;
						count -= more;
						if (count == 0)
						{
							return bytesCopied;
						}
					}
				}
			} while (Decode() || ((outputWindow.GetAvailable() > 0) && (mode != DECODE_CHKSUM)));
			return bytesCopied;
		}

		/// <summary>
		/// Returns true, if the input buffer is empty.
		/// You should then call setInput().
		/// NOTE: This method also returns true when the stream is finished.
		/// </summary>
		public bool IsNeedingInput
		{
			get
			{
				return input.IsNeedingInput;
			}
		}

		/// <summary>
		/// Returns true, if a preset dictionary is needed to inflate the input.
		/// </summary>
		public bool IsNeedingDictionary
		{
			get
			{
				return mode == DECODE_DICT && neededBits == 0;
			}
		}

		/// <summary>
		/// Returns true, if the inflater has finished.  This means, that no
		/// input is needed and no output can be produced.
		/// </summary>
		public bool IsFinished
		{
			get
			{
				return mode == FINISHED && outputWindow.GetAvailable() == 0;
			}
		}

		/// <summary>
		/// Gets the adler checksum.  This is either the checksum of all
		/// uncompressed bytes returned by inflate(), or if needsDictionary()
		/// returns true (and thus no output was yet produced) this is the
		/// adler checksum of the expected dictionary.
		/// </summary>
		/// <returns>
		/// the adler checksum.
		/// </returns>
		public int Adler
		{
			get
			{
				if (IsNeedingDictionary)
				{
					return readAdler;
				}
				else if (adler != null)
				{
					return (int)adler.Value;
				}
				else
				{
					return 0;
				}
			}
		}

		/// <summary>
		/// Gets the total number of output bytes returned by Inflate().
		/// </summary>
		/// <returns>
		/// the total number of output bytes.
		/// </returns>
		public long TotalOut
		{
			get
			{
				return totalOut;
			}
		}

		/// <summary>
		/// Gets the total number of processed compressed input bytes.
		/// </summary>
		/// <returns>
		/// The total number of bytes of processed input bytes.
		/// </returns>
		public long TotalIn
		{
			get
			{
				return totalIn - (long)RemainingInput;
			}
		}

		/// <summary>
		/// Gets the number of unprocessed input bytes.  Useful, if the end of the
		/// stream is reached and you want to further process the bytes after
		/// the deflate stream.
		/// </summary>
		/// <returns>
		/// The number of bytes of the input which have not been processed.
		/// </returns>
		public int RemainingInput
		{
			// TODO: This should be a long?
			get
			{
				return input.AvailableBytes;
			}
		}
	}
}


namespace ICSharpCode.SharpZipLib.Zip.Compression
{
	internal class InflaterDynHeader
	{
		#region Constants

		// maximum number of literal/length codes
		private const int LITLEN_MAX = 286;

		// maximum number of distance codes
		private const int DIST_MAX = 30;

		// maximum data code lengths to read
		private const int CODELEN_MAX = LITLEN_MAX + DIST_MAX;

		// maximum meta code length codes to read
		private const int META_MAX = 19;

		private static readonly int[] MetaCodeLengthIndex =
			{ 16, 17, 18, 0, 8, 7, 9, 6, 10, 5, 11, 4, 12, 3, 13, 2, 14, 1, 15 };

		#endregion Constants

		/// <summary>
		/// Continue decoding header from <see cref="input"/> until more bits are needed or decoding has been completed
		/// </summary>
		/// <returns>Returns whether decoding could be completed</returns>
		public bool AttemptRead()
			=> !state.MoveNext() || state.Current;

		public InflaterDynHeader(StreamManipulator input)
		{
			this.input = input;
			stateMachine = CreateStateMachine();
			state = stateMachine.GetEnumerator();
		}

		private IEnumerable<bool> CreateStateMachine()
		{
			// Read initial code length counts from header
			while (!input.TryGetBits(5, ref litLenCodeCount, 257)) yield return false;
			while (!input.TryGetBits(5, ref distanceCodeCount, 1)) yield return false;
			while (!input.TryGetBits(4, ref metaCodeCount, 4)) yield return false;
			var dataCodeCount = litLenCodeCount + distanceCodeCount;

			if (litLenCodeCount > LITLEN_MAX) throw new ValueOutOfRangeException(nameof(litLenCodeCount));
			if (distanceCodeCount > DIST_MAX) throw new ValueOutOfRangeException(nameof(distanceCodeCount));
			if (metaCodeCount > META_MAX) throw new ValueOutOfRangeException(nameof(metaCodeCount));

			// Load code lengths for the meta tree from the header bits
			for (int i = 0; i < metaCodeCount; i++)
			{
				while (!input.TryGetBits(3, ref codeLengths, MetaCodeLengthIndex[i])) yield return false;
			}

			var metaCodeTree = new InflaterHuffmanTree(codeLengths);

			// Decompress the meta tree symbols into the data table code lengths
			int index = 0;
			while (index < dataCodeCount)
			{
				byte codeLength;
				int symbol;

				while ((symbol = metaCodeTree.GetSymbol(input)) < 0) yield return false;

				if (symbol < 16)
				{
					// append literal code length
					codeLengths[index++] = (byte)symbol;
				}
				else
				{
					int repeatCount = 0;

					if (symbol == 16) // Repeat last code length 3..6 times
					{
						if (index == 0)
							throw new StreamDecodingException("Cannot repeat previous code length when no other code length has been read");

						codeLength = codeLengths[index - 1];

						// 2 bits + 3, [3..6]
						while (!input.TryGetBits(2, ref repeatCount, 3)) yield return false;
					}
					else if (symbol == 17) // Repeat zero 3..10 times
					{
						codeLength = 0;

						// 3 bits + 3, [3..10]
						while (!input.TryGetBits(3, ref repeatCount, 3)) yield return false;
					}
					else // (symbol == 18), Repeat zero 11..138 times
					{
						codeLength = 0;

						// 7 bits + 11, [11..138]
						while (!input.TryGetBits(7, ref repeatCount, 11)) yield return false;
					}

					if (index + repeatCount > dataCodeCount)
						throw new StreamDecodingException("Cannot repeat code lengths past total number of data code lengths");

					while (repeatCount-- > 0)
						codeLengths[index++] = codeLength;
				}
			}

			if (codeLengths[256] == 0)
				throw new StreamDecodingException("Inflater dynamic header end-of-block code missing");

			litLenTree = new InflaterHuffmanTree(new ArraySegment<byte>(codeLengths, 0, litLenCodeCount));
			distTree = new InflaterHuffmanTree(new ArraySegment<byte>(codeLengths, litLenCodeCount, distanceCodeCount));

			yield return true;
		}

		/// <summary>
		/// Get literal/length huffman tree, must not be used before <see cref="AttemptRead"/> has returned true
		/// </summary>
		/// <exception cref="StreamDecodingException">If hader has not been successfully read by the state machine</exception>
		public InflaterHuffmanTree LiteralLengthTree
			=> litLenTree ?? throw new StreamDecodingException("Header properties were accessed before header had been successfully read");

		/// <summary>
		/// Get distance huffman tree, must not be used before <see cref="AttemptRead"/> has returned true
		/// </summary>
		/// <exception cref="StreamDecodingException">If hader has not been successfully read by the state machine</exception>
		public InflaterHuffmanTree DistanceTree
			=> distTree ?? throw new StreamDecodingException("Header properties were accessed before header had been successfully read");

		#region Instance Fields

		private readonly StreamManipulator input;
		private readonly IEnumerator<bool> state;
		private readonly IEnumerable<bool> stateMachine;

		private byte[] codeLengths = new byte[CODELEN_MAX];

		private InflaterHuffmanTree litLenTree;
		private InflaterHuffmanTree distTree;

		private int litLenCodeCount, distanceCodeCount, metaCodeCount;

		#endregion Instance Fields
	}
}


namespace ICSharpCode.SharpZipLib.Zip.Compression
{
	/// <summary>
	/// Huffman tree used for inflation
	/// </summary>
	public class InflaterHuffmanTree
	{
		#region Constants

		private const int MAX_BITLEN = 15;

		#endregion Constants

		#region Instance Fields

		private short[] tree;

		#endregion Instance Fields

		/// <summary>
		/// Literal length tree
		/// </summary>
		public static InflaterHuffmanTree defLitLenTree;

		/// <summary>
		/// Distance tree
		/// </summary>
		public static InflaterHuffmanTree defDistTree;

		static InflaterHuffmanTree()
		{
			try
			{
				byte[] codeLengths = new byte[288];
				int i = 0;
				while (i < 144)
				{
					codeLengths[i++] = 8;
				}
				while (i < 256)
				{
					codeLengths[i++] = 9;
				}
				while (i < 280)
				{
					codeLengths[i++] = 7;
				}
				while (i < 288)
				{
					codeLengths[i++] = 8;
				}
				defLitLenTree = new InflaterHuffmanTree(codeLengths);

				codeLengths = new byte[32];
				i = 0;
				while (i < 32)
				{
					codeLengths[i++] = 5;
				}
				defDistTree = new InflaterHuffmanTree(codeLengths);
			}
			catch (Exception)
			{
				throw new SharpZipBaseException("InflaterHuffmanTree: static tree length illegal");
			}
		}

		#region Constructors

		/// <summary>
		/// Constructs a Huffman tree from the array of code lengths.
		/// </summary>
		/// <param name = "codeLengths">
		/// the array of code lengths
		/// </param>
		public InflaterHuffmanTree(IList<byte> codeLengths)
		{
			BuildTree(codeLengths);
		}

		#endregion Constructors

		private void BuildTree(IList<byte> codeLengths)
		{
			int[] blCount = new int[MAX_BITLEN + 1];
			int[] nextCode = new int[MAX_BITLEN + 1];

			for (int i = 0; i < codeLengths.Count; i++)
			{
				int bits = codeLengths[i];
				if (bits > 0)
				{
					blCount[bits]++;
				}
			}

			int code = 0;
			int treeSize = 512;
			for (int bits = 1; bits <= MAX_BITLEN; bits++)
			{
				nextCode[bits] = code;
				code += blCount[bits] << (16 - bits);
				if (bits >= 10)
				{
					/* We need an extra table for bit lengths >= 10. */
					int start = nextCode[bits] & 0x1ff80;
					int end = code & 0x1ff80;
					treeSize += (end - start) >> (16 - bits);
				}
			}

			/* -jr comment this out! doesnt work for dynamic trees and pkzip 2.04g
						if (code != 65536)
						{
							throw new SharpZipBaseException("Code lengths don't add up properly.");
						}
			*/
			/* Now create and fill the extra tables from longest to shortest
			* bit len.  This way the sub trees will be aligned.
			*/
			tree = new short[treeSize];
			int treePtr = 512;
			for (int bits = MAX_BITLEN; bits >= 10; bits--)
			{
				int end = code & 0x1ff80;
				code -= blCount[bits] << (16 - bits);
				int start = code & 0x1ff80;
				for (int i = start; i < end; i += 1 << 7)
				{
					tree[DeflaterHuffman.BitReverse(i)] = (short)((-treePtr << 4) | bits);
					treePtr += 1 << (bits - 9);
				}
			}

			for (int i = 0; i < codeLengths.Count; i++)
			{
				int bits = codeLengths[i];
				if (bits == 0)
				{
					continue;
				}
				code = nextCode[bits];
				int revcode = DeflaterHuffman.BitReverse(code);
				if (bits <= 9)
				{
					do
					{
						tree[revcode] = (short)((i << 4) | bits);
						revcode += 1 << bits;
					} while (revcode < 512);
				}
				else
				{
					int subTree = tree[revcode & 511];
					int treeLen = 1 << (subTree & 15);
					subTree = -(subTree >> 4);
					do
					{
						tree[subTree | (revcode >> 9)] = (short)((i << 4) | bits);
						revcode += 1 << bits;
					} while (revcode < treeLen);
				}
				nextCode[bits] = code + (1 << (16 - bits));
			}
		}

		/// <summary>
		/// Reads the next symbol from input.  The symbol is encoded using the
		/// huffman tree.
		/// </summary>
		/// <param name="input">
		/// input the input source.
		/// </param>
		/// <returns>
		/// the next symbol, or -1 if not enough input is available.
		/// </returns>
		public int GetSymbol(StreamManipulator input)
		{
			int lookahead, symbol;
			if ((lookahead = input.PeekBits(9)) >= 0)
			{
                symbol = tree[lookahead];
				int bitlen = symbol & 15;

				if (symbol >= 0)
				{
                    if(bitlen == 0){
                        throw new SharpZipBaseException("Encountered invalid codelength 0");
                    } 
					input.DropBits(bitlen);
					return symbol >> 4;
				}
				int subtree = -(symbol >> 4);
				if ((lookahead = input.PeekBits(bitlen)) >= 0)
				{
					symbol = tree[subtree | (lookahead >> 9)];
					input.DropBits(symbol & 15);
					return symbol >> 4;
				}
				else
				{
					int bits = input.AvailableBits;
					lookahead = input.PeekBits(bits);
					symbol = tree[subtree | (lookahead >> 9)];
					if ((symbol & 15) <= bits)
					{
						input.DropBits(symbol & 15);
						return symbol >> 4;
					}
					else
					{
						return -1;
					}
				}
			}
			else // Less than 9 bits
			{
				int bits = input.AvailableBits;
				lookahead = input.PeekBits(bits);
				symbol = tree[lookahead];
				if (symbol >= 0 && (symbol & 15) <= bits)
				{
					input.DropBits(symbol & 15);
					return symbol >> 4;
				}
				else
				{
					return -1;
				}
			}
		}
	}
}

namespace ICSharpCode.SharpZipLib.Zip.Compression
{
	/// <summary>
	/// This class is general purpose class for writing data to a buffer.
	///
	/// It allows you to write bits as well as bytes
	/// Based on DeflaterPending.java
	///
	/// author of the original java version : Jochen Hoenicke
	/// </summary>
	public class PendingBuffer
	{
		#region Instance Fields

		/// <summary>
		/// Internal work buffer
		/// </summary>
		private readonly byte[] buffer;

		private int start;
		private int end;

		private uint bits;
		private int bitCount;

		#endregion Instance Fields

		#region Constructors

		/// <summary>
		/// construct instance using default buffer size of 4096
		/// </summary>
		public PendingBuffer() : this(4096)
		{
		}

		/// <summary>
		/// </summary>
		/// <param name="bufferSize">
		/// size to use for internal buffer
		/// </param>
		public PendingBuffer(int bufferSize)
		{
			buffer = new byte[bufferSize];
		}

		#endregion Constructors

		/// <summary>
		/// Clear internal state/buffers
		/// </summary>
		public void Reset()
		{
			start = end = bitCount = 0;
		}

		/// <summary>
		/// Write a byte to buffer
		/// </summary>
		/// <param name="value">
		/// The value to write
		/// </param>
		public void WriteByte(int value)
		{
#if DebugDeflation
			if (DeflaterConstants.DEBUGGING && (start != 0) )
			{
				throw new SharpZipBaseException("Debug check: start != 0");
			}
#endif
			buffer[end++] = unchecked((byte)value);
		}

		/// <summary>
		/// Write a short value to buffer LSB first
		/// </summary>
		/// <param name="value">
		/// The value to write.
		/// </param>
		public void WriteShort(int value)
		{
#if DebugDeflation
			if (DeflaterConstants.DEBUGGING && (start != 0) )
			{
				throw new SharpZipBaseException("Debug check: start != 0");
			}
#endif
			buffer[end++] = unchecked((byte)value);
			buffer[end++] = unchecked((byte)(value >> 8));
		}

		/// <summary>
		/// write an integer LSB first
		/// </summary>
		/// <param name="value">The value to write.</param>
		public void WriteInt(int value)
		{
#if DebugDeflation
			if (DeflaterConstants.DEBUGGING && (start != 0) )
			{
				throw new SharpZipBaseException("Debug check: start != 0");
			}
#endif
			buffer[end++] = unchecked((byte)value);
			buffer[end++] = unchecked((byte)(value >> 8));
			buffer[end++] = unchecked((byte)(value >> 16));
			buffer[end++] = unchecked((byte)(value >> 24));
		}

		/// <summary>
		/// Write a block of data to buffer
		/// </summary>
		/// <param name="block">data to write</param>
		/// <param name="offset">offset of first byte to write</param>
		/// <param name="length">number of bytes to write</param>
		public void WriteBlock(byte[] block, int offset, int length)
		{
#if DebugDeflation
			if (DeflaterConstants.DEBUGGING && (start != 0) )
			{
				throw new SharpZipBaseException("Debug check: start != 0");
			}
#endif
			System.Array.Copy(block, offset, buffer, end, length);
			end += length;
		}

		/// <summary>
		/// The number of bits written to the buffer
		/// </summary>
		public int BitCount
		{
			get
			{
				return bitCount;
			}
		}

		/// <summary>
		/// Align internal buffer on a byte boundary
		/// </summary>
		public void AlignToByte()
		{
#if DebugDeflation
			if (DeflaterConstants.DEBUGGING && (start != 0) )
			{
				throw new SharpZipBaseException("Debug check: start != 0");
			}
#endif
			if (bitCount > 0)
			{
				buffer[end++] = unchecked((byte)bits);
				if (bitCount > 8)
				{
					buffer[end++] = unchecked((byte)(bits >> 8));
				}
			}
			bits = 0;
			bitCount = 0;
		}

		/// <summary>
		/// Write bits to internal buffer
		/// </summary>
		/// <param name="b">source of bits</param>
		/// <param name="count">number of bits to write</param>
		public void WriteBits(int b, int count)
		{
#if DebugDeflation
			if (DeflaterConstants.DEBUGGING && (start != 0) )
			{
				throw new SharpZipBaseException("Debug check: start != 0");
			}

			//			if (DeflaterConstants.DEBUGGING) {
			//				//Console.WriteLine("writeBits("+b+","+count+")");
			//			}
#endif
			bits |= (uint)(b << bitCount);
			bitCount += count;
			if (bitCount >= 16)
			{
				buffer[end++] = unchecked((byte)bits);
				buffer[end++] = unchecked((byte)(bits >> 8));
				bits >>= 16;
				bitCount -= 16;
			}
		}

		/// <summary>
		/// Write a short value to internal buffer most significant byte first
		/// </summary>
		/// <param name="s">value to write</param>
		public void WriteShortMSB(int s)
		{
#if DebugDeflation
			if (DeflaterConstants.DEBUGGING && (start != 0) )
			{
				throw new SharpZipBaseException("Debug check: start != 0");
			}
#endif
			buffer[end++] = unchecked((byte)(s >> 8));
			buffer[end++] = unchecked((byte)s);
		}

		/// <summary>
		/// Indicates if buffer has been flushed
		/// </summary>
		public bool IsFlushed
		{
			get
			{
				return end == 0;
			}
		}

		/// <summary>
		/// Flushes the pending buffer into the given output array.  If the
		/// output array is to small, only a partial flush is done.
		/// </summary>
		/// <param name="output">The output array.</param>
		/// <param name="offset">The offset into output array.</param>
		/// <param name="length">The maximum number of bytes to store.</param>
		/// <returns>The number of bytes flushed.</returns>
		public int Flush(byte[] output, int offset, int length)
		{
			if (bitCount >= 8)
			{
				buffer[end++] = unchecked((byte)bits);
				bits >>= 8;
				bitCount -= 8;
			}

			if (length > end - start)
			{
				length = end - start;
				System.Array.Copy(buffer, start, output, offset, length);
				start = 0;
				end = 0;
			}
			else
			{
				System.Array.Copy(buffer, start, output, offset, length);
				start += length;
			}
			return length;
		}

		/// <summary>
		/// Convert internal buffer to byte array.
		/// Buffer is empty on completion
		/// </summary>
		/// <returns>
		/// The internal buffer contents converted to a byte array.
		/// </returns>
		public byte[] ToByteArray()
		{
			AlignToByte();

			byte[] result = new byte[end - start];
			System.Array.Copy(buffer, start, result, 0, result.Length);
			start = 0;
			end = 0;
			return result;
		}
	}
}

namespace ICSharpCode.SharpZipLib.Zip.Compression.Streams
{
	/// <summary>
	/// A special stream deflating or compressing the bytes that are
	/// written to it.  It uses a Deflater to perform actual deflating.<br/>
	/// Authors of the original java version : Tom Tromey, Jochen Hoenicke
	/// </summary>
	public class DeflaterOutputStream : Stream
	{
		#region Constructors

		/// <summary>
		/// Creates a new DeflaterOutputStream with a default Deflater and default buffer size.
		/// </summary>
		/// <param name="baseOutputStream">
		/// the output stream where deflated output should be written.
		/// </param>
		public DeflaterOutputStream(Stream baseOutputStream)
			: this(baseOutputStream, new Deflater(), 512)
		{
		}

		/// <summary>
		/// Creates a new DeflaterOutputStream with the given Deflater and
		/// default buffer size.
		/// </summary>
		/// <param name="baseOutputStream">
		/// the output stream where deflated output should be written.
		/// </param>
		/// <param name="deflater">
		/// the underlying deflater.
		/// </param>
		public DeflaterOutputStream(Stream baseOutputStream, Deflater deflater)
			: this(baseOutputStream, deflater, 512)
		{
		}

		/// <summary>
		/// Creates a new DeflaterOutputStream with the given Deflater and
		/// buffer size.
		/// </summary>
		/// <param name="baseOutputStream">
		/// The output stream where deflated output is written.
		/// </param>
		/// <param name="deflater">
		/// The underlying deflater to use
		/// </param>
		/// <param name="bufferSize">
		/// The buffer size in bytes to use when deflating (minimum value 512)
		/// </param>
		/// <exception cref="ArgumentOutOfRangeException">
		/// bufsize is less than or equal to zero.
		/// </exception>
		/// <exception cref="ArgumentException">
		/// baseOutputStream does not support writing
		/// </exception>
		/// <exception cref="ArgumentNullException">
		/// deflater instance is null
		/// </exception>
		public DeflaterOutputStream(Stream baseOutputStream, Deflater deflater, int bufferSize)
		{
			if (baseOutputStream == null)
			{
				throw new ArgumentNullException(nameof(baseOutputStream));
			}

			if (baseOutputStream.CanWrite == false)
			{
				throw new ArgumentException("Must support writing", nameof(baseOutputStream));
			}

			if (bufferSize < 512)
			{
				throw new ArgumentOutOfRangeException(nameof(bufferSize));
			}

			baseOutputStream_ = baseOutputStream;
			buffer_ = new byte[bufferSize];
			deflater_ = deflater ?? throw new ArgumentNullException(nameof(deflater));
		}

		#endregion Constructors

		#region Public API

		/// <summary>
		/// Finishes the stream by calling finish() on the deflater.
		/// </summary>
		/// <exception cref="SharpZipBaseException">
		/// Not all input is deflated
		/// </exception>
		public virtual void Finish()
		{
			deflater_.Finish();
			while (!deflater_.IsFinished)
			{
				int len = deflater_.Deflate(buffer_, 0, buffer_.Length);
				if (len <= 0)
				{
					break;
				}

				if (cryptoTransform_ != null)
				{
					EncryptBlock(buffer_, 0, len);
				}

				baseOutputStream_.Write(buffer_, 0, len);
			}

			if (!deflater_.IsFinished)
			{
				throw new SharpZipBaseException("Can't deflate all input?");
			}

			baseOutputStream_.Flush();

			if (cryptoTransform_ != null)
			{
				if (cryptoTransform_ is ZipAESTransform)
				{
					AESAuthCode = ((ZipAESTransform)cryptoTransform_).GetAuthCode();
				}
				cryptoTransform_.Dispose();
				cryptoTransform_ = null;
			}
		}

		/// <summary>
		/// Gets or sets a flag indicating ownership of underlying stream.
		/// When the flag is true <see cref="Stream.Dispose()" /> will close the underlying stream also.
		/// </summary>
		/// <remarks>The default value is true.</remarks>
		public bool IsStreamOwner { get; set; } = true;

		///	<summary>
		/// Allows client to determine if an entry can be patched after its added
		/// </summary>
		public bool CanPatchEntries
		{
			get
			{
				return baseOutputStream_.CanSeek;
			}
		}

		#endregion Public API

		#region Encryption

		/// <summary>
		/// The CryptoTransform currently being used to encrypt the compressed data.
		/// </summary>
		protected ICryptoTransform cryptoTransform_;

		/// <summary>
		/// Returns the 10 byte AUTH CODE to be appended immediately following the AES data stream.
		/// </summary>
		protected byte[] AESAuthCode;

		/// <summary>
		/// Encrypt a block of data
		/// </summary>
		/// <param name="buffer">
		/// Data to encrypt.  NOTE the original contents of the buffer are lost
		/// </param>
		/// <param name="offset">
		/// Offset of first byte in buffer to encrypt
		/// </param>
		/// <param name="length">
		/// Number of bytes in buffer to encrypt
		/// </param>
		protected void EncryptBlock(byte[] buffer, int offset, int length)
		{
			cryptoTransform_.TransformBlock(buffer, 0, length, buffer, 0);
		}

		#endregion Encryption

		#region Deflation Support

		/// <summary>
		/// Deflates everything in the input buffers.  This will call
		/// <code>def.deflate()</code> until all bytes from the input buffers
		/// are processed.
		/// </summary>
		protected void Deflate()
		{
			Deflate(false);
		}

		private void Deflate(bool flushing)
		{
			while (flushing || !deflater_.IsNeedingInput)
			{
				int deflateCount = deflater_.Deflate(buffer_, 0, buffer_.Length);

				if (deflateCount <= 0)
				{
					break;
				}
				if (cryptoTransform_ != null)
				{
					EncryptBlock(buffer_, 0, deflateCount);
				}

				baseOutputStream_.Write(buffer_, 0, deflateCount);
			}

			if (!deflater_.IsNeedingInput)
			{
				throw new SharpZipBaseException("DeflaterOutputStream can't deflate all input?");
			}
		}

		#endregion Deflation Support

		#region Stream Overrides

		/// <summary>
		/// Gets value indicating stream can be read from
		/// </summary>
		public override bool CanRead
		{
			get
			{
				return false;
			}
		}

		/// <summary>
		/// Gets a value indicating if seeking is supported for this stream
		/// This property always returns false
		/// </summary>
		public override bool CanSeek
		{
			get
			{
				return false;
			}
		}

		/// <summary>
		/// Get value indicating if this stream supports writing
		/// </summary>
		public override bool CanWrite
		{
			get
			{
				return baseOutputStream_.CanWrite;
			}
		}

		/// <summary>
		/// Get current length of stream
		/// </summary>
		public override long Length
		{
			get
			{
				return baseOutputStream_.Length;
			}
		}

		/// <summary>
		/// Gets the current position within the stream.
		/// </summary>
		/// <exception cref="NotSupportedException">Any attempt to set position</exception>
		public override long Position
		{
			get
			{
				return baseOutputStream_.Position;
			}
			set
			{
				throw new NotSupportedException("Position property not supported");
			}
		}

		/// <summary>
		/// Sets the current position of this stream to the given value. Not supported by this class!
		/// </summary>
		/// <param name="offset">The offset relative to the <paramref name="origin"/> to seek.</param>
		/// <param name="origin">The <see cref="SeekOrigin"/> to seek from.</param>
		/// <returns>The new position in the stream.</returns>
		/// <exception cref="NotSupportedException">Any access</exception>
		public override long Seek(long offset, SeekOrigin origin)
		{
			throw new NotSupportedException("DeflaterOutputStream Seek not supported");
		}

		/// <summary>
		/// Sets the length of this stream to the given value. Not supported by this class!
		/// </summary>
		/// <param name="value">The new stream length.</param>
		/// <exception cref="NotSupportedException">Any access</exception>
		public override void SetLength(long value)
		{
			throw new NotSupportedException("DeflaterOutputStream SetLength not supported");
		}

		/// <summary>
		/// Read a byte from stream advancing position by one
		/// </summary>
		/// <returns>The byte read cast to an int.  THe value is -1 if at the end of the stream.</returns>
		/// <exception cref="NotSupportedException">Any access</exception>
		public override int ReadByte()
		{
			throw new NotSupportedException("DeflaterOutputStream ReadByte not supported");
		}

		/// <summary>
		/// Read a block of bytes from stream
		/// </summary>
		/// <param name="buffer">The buffer to store read data in.</param>
		/// <param name="offset">The offset to start storing at.</param>
		/// <param name="count">The maximum number of bytes to read.</param>
		/// <returns>The actual number of bytes read.  Zero if end of stream is detected.</returns>
		/// <exception cref="NotSupportedException">Any access</exception>
		public override int Read(byte[] buffer, int offset, int count)
		{
			throw new NotSupportedException("DeflaterOutputStream Read not supported");
		}

		/// <summary>
		/// Flushes the stream by calling <see cref="Flush">Flush</see> on the deflater and then
		/// on the underlying stream.  This ensures that all bytes are flushed.
		/// </summary>
		public override void Flush()
		{
			deflater_.Flush();
			Deflate(true);
			baseOutputStream_.Flush();
		}

		/// <summary>
		/// Calls <see cref="Finish"/> and closes the underlying
		/// stream when <see cref="IsStreamOwner"></see> is true.
		/// </summary>
		protected override void Dispose(bool disposing)
		{
			if (!isClosed_)
			{
				isClosed_ = true;

				try
				{
					Finish();
					if (cryptoTransform_ != null)
					{
						GetAuthCodeIfAES();
						cryptoTransform_.Dispose();
						cryptoTransform_ = null;
					}
				}
				finally
				{
					if (IsStreamOwner)
					{
						baseOutputStream_.Dispose();
					}
				}
			}
		}

		/// <summary>
		/// Get the Auth code for AES encrypted entries
		/// </summary>
		protected void GetAuthCodeIfAES()
		{
			if (cryptoTransform_ is ZipAESTransform)
			{
				AESAuthCode = ((ZipAESTransform)cryptoTransform_).GetAuthCode();
			}
		}

		/// <summary>
		/// Writes a single byte to the compressed output stream.
		/// </summary>
		/// <param name="value">
		/// The byte value.
		/// </param>
		public override void WriteByte(byte value)
		{
			byte[] b = new byte[1];
			b[0] = value;
			Write(b, 0, 1);
		}

		/// <summary>
		/// Writes bytes from an array to the compressed stream.
		/// </summary>
		/// <param name="buffer">
		/// The byte array
		/// </param>
		/// <param name="offset">
		/// The offset into the byte array where to start.
		/// </param>
		/// <param name="count">
		/// The number of bytes to write.
		/// </param>
		public override void Write(byte[] buffer, int offset, int count)
		{
			deflater_.SetInput(buffer, offset, count);
			Deflate();
		}

		#endregion Stream Overrides

		#region Instance Fields

		/// <summary>
		/// This buffer is used temporarily to retrieve the bytes from the
		/// deflater and write them to the underlying output stream.
		/// </summary>
		private byte[] buffer_;

		/// <summary>
		/// The deflater which is used to deflate the stream.
		/// </summary>
		protected Deflater deflater_;

		/// <summary>
		/// Base stream the deflater depends on.
		/// </summary>
		protected Stream baseOutputStream_;

		private bool isClosed_;

		#endregion Instance Fields
	}
}


namespace ICSharpCode.SharpZipLib.Zip.Compression.Streams
{
	/// <summary>
	/// An input buffer customised for use by <see cref="InflaterInputStream"/>
	/// </summary>
	/// <remarks>
	/// The buffer supports decryption of incoming data.
	/// </remarks>
	public class InflaterInputBuffer
	{
		#region Constructors

		/// <summary>
		/// Initialise a new instance of <see cref="InflaterInputBuffer"/> with a default buffer size
		/// </summary>
		/// <param name="stream">The stream to buffer.</param>
		public InflaterInputBuffer(Stream stream) : this(stream, 4096)
		{
		}

		/// <summary>
		/// Initialise a new instance of <see cref="InflaterInputBuffer"/>
		/// </summary>
		/// <param name="stream">The stream to buffer.</param>
		/// <param name="bufferSize">The size to use for the buffer</param>
		/// <remarks>A minimum buffer size of 1KB is permitted.  Lower sizes are treated as 1KB.</remarks>
		public InflaterInputBuffer(Stream stream, int bufferSize)
		{
			inputStream = stream;
			if (bufferSize < 1024)
			{
				bufferSize = 1024;
			}
			rawData = new byte[bufferSize];
			clearText = rawData;
		}

		#endregion Constructors

		/// <summary>
		/// Get the length of bytes in the <see cref="RawData"/>
		/// </summary>
		public int RawLength
		{
			get
			{
				return rawLength;
			}
		}

		/// <summary>
		/// Get the contents of the raw data buffer.
		/// </summary>
		/// <remarks>This may contain encrypted data.</remarks>
		public byte[] RawData
		{
			get
			{
				return rawData;
			}
		}

		/// <summary>
		/// Get the number of useable bytes in <see cref="ClearText"/>
		/// </summary>
		public int ClearTextLength
		{
			get
			{
				return clearTextLength;
			}
		}

		/// <summary>
		/// Get the contents of the clear text buffer.
		/// </summary>
		public byte[] ClearText
		{
			get
			{
				return clearText;
			}
		}

		/// <summary>
		/// Get/set the number of bytes available
		/// </summary>
		public int Available
		{
			get { return available; }
			set { available = value; }
		}

		/// <summary>
		/// Call <see cref="Inflater.SetInput(byte[], int, int)"/> passing the current clear text buffer contents.
		/// </summary>
		/// <param name="inflater">The inflater to set input for.</param>
		public void SetInflaterInput(Inflater inflater)
		{
			if (available > 0)
			{
				inflater.SetInput(clearText, clearTextLength - available, available);
				available = 0;
			}
		}

		/// <summary>
		/// Fill the buffer from the underlying input stream.
		/// </summary>
		public void Fill()
		{
			rawLength = 0;
			int toRead = rawData.Length;

			while (toRead > 0 && inputStream.CanRead)
			{
				int count = inputStream.Read(rawData, rawLength, toRead);
				if (count <= 0)
				{
					break;
				}
				rawLength += count;
				toRead -= count;
			}

			if (cryptoTransform != null)
			{
				clearTextLength = cryptoTransform.TransformBlock(rawData, 0, rawLength, clearText, 0);
			}
			else
			{
				clearTextLength = rawLength;
			}

			available = clearTextLength;
		}

		/// <summary>
		/// Read a buffer directly from the input stream
		/// </summary>
		/// <param name="buffer">The buffer to fill</param>
		/// <returns>Returns the number of bytes read.</returns>
		public int ReadRawBuffer(byte[] buffer)
		{
			return ReadRawBuffer(buffer, 0, buffer.Length);
		}

		/// <summary>
		/// Read a buffer directly from the input stream
		/// </summary>
		/// <param name="outBuffer">The buffer to read into</param>
		/// <param name="offset">The offset to start reading data into.</param>
		/// <param name="length">The number of bytes to read.</param>
		/// <returns>Returns the number of bytes read.</returns>
		public int ReadRawBuffer(byte[] outBuffer, int offset, int length)
		{
			if (length < 0)
			{
				throw new ArgumentOutOfRangeException(nameof(length));
			}

			int currentOffset = offset;
			int currentLength = length;

			while (currentLength > 0)
			{
				if (available <= 0)
				{
					Fill();
					if (available <= 0)
					{
						return 0;
					}
				}
				int toCopy = Math.Min(currentLength, available);
				System.Array.Copy(rawData, rawLength - (int)available, outBuffer, currentOffset, toCopy);
				currentOffset += toCopy;
				currentLength -= toCopy;
				available -= toCopy;
			}
			return length;
		}

		/// <summary>
		/// Read clear text data from the input stream.
		/// </summary>
		/// <param name="outBuffer">The buffer to add data to.</param>
		/// <param name="offset">The offset to start adding data at.</param>
		/// <param name="length">The number of bytes to read.</param>
		/// <returns>Returns the number of bytes actually read.</returns>
		public int ReadClearTextBuffer(byte[] outBuffer, int offset, int length)
		{
			if (length < 0)
			{
				throw new ArgumentOutOfRangeException(nameof(length));
			}

			int currentOffset = offset;
			int currentLength = length;

			while (currentLength > 0)
			{
				if (available <= 0)
				{
					Fill();
					if (available <= 0)
					{
						return 0;
					}
				}

				int toCopy = Math.Min(currentLength, available);
				Array.Copy(clearText, clearTextLength - (int)available, outBuffer, currentOffset, toCopy);
				currentOffset += toCopy;
				currentLength -= toCopy;
				available -= toCopy;
			}
			return length;
		}

		/// <summary>
		/// Read a <see cref="byte"/> from the input stream.
		/// </summary>
		/// <returns>Returns the byte read.</returns>
		public byte ReadLeByte()
		{
			if (available <= 0)
			{
				Fill();
				if (available <= 0)
				{
					throw new ZipException("EOF in header");
				}
			}
			byte result = rawData[rawLength - available];
			available -= 1;
			return result;
		}

		/// <summary>
		/// Read an <see cref="short"/> in little endian byte order.
		/// </summary>
		/// <returns>The short value read case to an int.</returns>
		public int ReadLeShort()
		{
			return ReadLeByte() | (ReadLeByte() << 8);
		}

		/// <summary>
		/// Read an <see cref="int"/> in little endian byte order.
		/// </summary>
		/// <returns>The int value read.</returns>
		public int ReadLeInt()
		{
			return ReadLeShort() | (ReadLeShort() << 16);
		}

		/// <summary>
		/// Read a <see cref="long"/> in little endian byte order.
		/// </summary>
		/// <returns>The long value read.</returns>
		public long ReadLeLong()
		{
			return (uint)ReadLeInt() | ((long)ReadLeInt() << 32);
		}

		/// <summary>
		/// Get/set the <see cref="ICryptoTransform"/> to apply to any data.
		/// </summary>
		/// <remarks>Set this value to null to have no transform applied.</remarks>
		public ICryptoTransform CryptoTransform
		{
			set
			{
				cryptoTransform = value;
				if (cryptoTransform != null)
				{
					if (rawData == clearText)
					{
						if (internalClearText == null)
						{
							internalClearText = new byte[rawData.Length];
						}
						clearText = internalClearText;
					}
					clearTextLength = rawLength;
					if (available > 0)
					{
						cryptoTransform.TransformBlock(rawData, rawLength - available, available, clearText, rawLength - available);
					}
				}
				else
				{
					clearText = rawData;
					clearTextLength = rawLength;
				}
			}
		}

		#region Instance Fields

		private int rawLength;
		private byte[] rawData;

		private int clearTextLength;
		private byte[] clearText;
		private byte[] internalClearText;

		private int available;

		private ICryptoTransform cryptoTransform;
		private Stream inputStream;

		#endregion Instance Fields
	}

	/// <summary>
	/// This filter stream is used to decompress data compressed using the "deflate"
	/// format. The "deflate" format is described in RFC 1951.
	///
	/// This stream may form the basis for other decompression filters, such
	/// as the <see cref="ICSharpCode.SharpZipLib.GZip.GZipInputStream">GZipInputStream</see>.
	///
	/// Author of the original java version : John Leuner.
	/// </summary>
	public class InflaterInputStream : Stream
	{
		#region Constructors

		/// <summary>
		/// Create an InflaterInputStream with the default decompressor
		/// and a default buffer size of 4KB.
		/// </summary>
		/// <param name = "baseInputStream">
		/// The InputStream to read bytes from
		/// </param>
		public InflaterInputStream(Stream baseInputStream)
			: this(baseInputStream, new Inflater(), 4096)
		{
		}

		/// <summary>
		/// Create an InflaterInputStream with the specified decompressor
		/// and a default buffer size of 4KB.
		/// </summary>
		/// <param name = "baseInputStream">
		/// The source of input data
		/// </param>
		/// <param name = "inf">
		/// The decompressor used to decompress data read from baseInputStream
		/// </param>
		public InflaterInputStream(Stream baseInputStream, Inflater inf)
			: this(baseInputStream, inf, 4096)
		{
		}

		/// <summary>
		/// Create an InflaterInputStream with the specified decompressor
		/// and the specified buffer size.
		/// </summary>
		/// <param name = "baseInputStream">
		/// The InputStream to read bytes from
		/// </param>
		/// <param name = "inflater">
		/// The decompressor to use
		/// </param>
		/// <param name = "bufferSize">
		/// Size of the buffer to use
		/// </param>
		public InflaterInputStream(Stream baseInputStream, Inflater inflater, int bufferSize)
		{
			if (baseInputStream == null)
			{
				throw new ArgumentNullException(nameof(baseInputStream));
			}

			if (inflater == null)
			{
				throw new ArgumentNullException(nameof(inflater));
			}

			if (bufferSize <= 0)
			{
				throw new ArgumentOutOfRangeException(nameof(bufferSize));
			}

			this.baseInputStream = baseInputStream;
			this.inf = inflater;

			inputBuffer = new InflaterInputBuffer(baseInputStream, bufferSize);
		}

		#endregion Constructors

		/// <summary>
		/// Gets or sets a flag indicating ownership of underlying stream.
		/// When the flag is true <see cref="Stream.Dispose()" /> will close the underlying stream also.
		/// </summary>
		/// <remarks>The default value is true.</remarks>
		public bool IsStreamOwner { get; set; } = true;

		/// <summary>
		/// Skip specified number of bytes of uncompressed data
		/// </summary>
		/// <param name ="count">
		/// Number of bytes to skip
		/// </param>
		/// <returns>
		/// The number of bytes skipped, zero if the end of
		/// stream has been reached
		/// </returns>
		/// <exception cref="ArgumentOutOfRangeException">
		/// <paramref name="count">The number of bytes</paramref> to skip is less than or equal to zero.
		/// </exception>
		public long Skip(long count)
		{
			if (count <= 0)
			{
				throw new ArgumentOutOfRangeException(nameof(count));
			}

			// v0.80 Skip by seeking if underlying stream supports it...
			if (baseInputStream.CanSeek)
			{
				baseInputStream.Seek(count, SeekOrigin.Current);
				return count;
			}
			else
			{
				int length = 2048;
				if (count < length)
				{
					length = (int)count;
				}

				byte[] tmp = new byte[length];
				int readCount = 1;
				long toSkip = count;

				while ((toSkip > 0) && (readCount > 0))
				{
					if (toSkip < length)
					{
						length = (int)toSkip;
					}

					readCount = baseInputStream.Read(tmp, 0, length);
					toSkip -= readCount;
				}

				return count - toSkip;
			}
		}

		/// <summary>
		/// Clear any cryptographic state.
		/// </summary>
		protected void StopDecrypting()
		{
			inputBuffer.CryptoTransform = null;
		}

		/// <summary>
		/// Returns 0 once the end of the stream (EOF) has been reached.
		/// Otherwise returns 1.
		/// </summary>
		public virtual int Available
		{
			get
			{
				return inf.IsFinished ? 0 : 1;
			}
		}

		/// <summary>
		/// Fills the buffer with more data to decompress.
		/// </summary>
		/// <exception cref="SharpZipBaseException">
		/// Stream ends early
		/// </exception>
		protected void Fill()
		{
			// Protect against redundant calls
			if (inputBuffer.Available <= 0)
			{
				inputBuffer.Fill();
				if (inputBuffer.Available <= 0)
				{
					throw new SharpZipBaseException("Unexpected EOF");
				}
			}
			inputBuffer.SetInflaterInput(inf);
		}

		#region Stream Overrides

		/// <summary>
		/// Gets a value indicating whether the current stream supports reading
		/// </summary>
		public override bool CanRead
		{
			get
			{
				return baseInputStream.CanRead;
			}
		}

		/// <summary>
		/// Gets a value of false indicating seeking is not supported for this stream.
		/// </summary>
		public override bool CanSeek
		{
			get
			{
				return false;
			}
		}

		/// <summary>
		/// Gets a value of false indicating that this stream is not writeable.
		/// </summary>
		public override bool CanWrite
		{
			get
			{
				return false;
			}
		}

		/// <summary>
		/// A value representing the length of the stream in bytes.
		/// </summary>
		public override long Length
		{
			get
			{
				//return inputBuffer.RawLength;
				throw new NotSupportedException("InflaterInputStream Length is not supported");
			}
		}

		/// <summary>
		/// The current position within the stream.
		/// Throws a NotSupportedException when attempting to set the position
		/// </summary>
		/// <exception cref="NotSupportedException">Attempting to set the position</exception>
		public override long Position
		{
			get
			{
				return baseInputStream.Position;
			}
			set
			{
				throw new NotSupportedException("InflaterInputStream Position not supported");
			}
		}

		/// <summary>
		/// Flushes the baseInputStream
		/// </summary>
		public override void Flush()
		{
			baseInputStream.Flush();
		}

		/// <summary>
		/// Sets the position within the current stream
		/// Always throws a NotSupportedException
		/// </summary>
		/// <param name="offset">The relative offset to seek to.</param>
		/// <param name="origin">The <see cref="SeekOrigin"/> defining where to seek from.</param>
		/// <returns>The new position in the stream.</returns>
		/// <exception cref="NotSupportedException">Any access</exception>
		public override long Seek(long offset, SeekOrigin origin)
		{
			throw new NotSupportedException("Seek not supported");
		}

		/// <summary>
		/// Set the length of the current stream
		/// Always throws a NotSupportedException
		/// </summary>
		/// <param name="value">The new length value for the stream.</param>
		/// <exception cref="NotSupportedException">Any access</exception>
		public override void SetLength(long value)
		{
			throw new NotSupportedException("InflaterInputStream SetLength not supported");
		}

		/// <summary>
		/// Writes a sequence of bytes to stream and advances the current position
		/// This method always throws a NotSupportedException
		/// </summary>
		/// <param name="buffer">The buffer containing data to write.</param>
		/// <param name="offset">The offset of the first byte to write.</param>
		/// <param name="count">The number of bytes to write.</param>
		/// <exception cref="NotSupportedException">Any access</exception>
		public override void Write(byte[] buffer, int offset, int count)
		{
			throw new NotSupportedException("InflaterInputStream Write not supported");
		}

		/// <summary>
		/// Writes one byte to the current stream and advances the current position
		/// Always throws a NotSupportedException
		/// </summary>
		/// <param name="value">The byte to write.</param>
		/// <exception cref="NotSupportedException">Any access</exception>
		public override void WriteByte(byte value)
		{
			throw new NotSupportedException("InflaterInputStream WriteByte not supported");
		}

		/// <summary>
		/// Closes the input stream.  When <see cref="IsStreamOwner"></see>
		/// is true the underlying stream is also closed.
		/// </summary>
		protected override void Dispose(bool disposing)
		{
			if (!isClosed)
			{
				isClosed = true;
				if (IsStreamOwner)
				{
					baseInputStream.Dispose();
				}
			}
		}

		/// <summary>
		/// Reads decompressed data into the provided buffer byte array
		/// </summary>
		/// <param name ="buffer">
		/// The array to read and decompress data into
		/// </param>
		/// <param name ="offset">
		/// The offset indicating where the data should be placed
		/// </param>
		/// <param name ="count">
		/// The number of bytes to decompress
		/// </param>
		/// <returns>The number of bytes read.  Zero signals the end of stream</returns>
		/// <exception cref="SharpZipBaseException">
		/// Inflater needs a dictionary
		/// </exception>
		public override int Read(byte[] buffer, int offset, int count)
		{
			if (inf.IsNeedingDictionary)
			{
				throw new SharpZipBaseException("Need a dictionary");
			}

			int remainingBytes = count;
			while (true)
			{
				int bytesRead = inf.Inflate(buffer, offset, remainingBytes);
				offset += bytesRead;
				remainingBytes -= bytesRead;

				if (remainingBytes == 0 || inf.IsFinished)
				{
					break;
				}

				if (inf.IsNeedingInput)
				{
					Fill();
				}
				else if (bytesRead == 0)
				{
					throw new ZipException("Invalid input data");
				}
			}
			return count - remainingBytes;
		}

		#endregion Stream Overrides

		#region Instance Fields

		/// <summary>
		/// Decompressor for this stream
		/// </summary>
		protected Inflater inf;

		/// <summary>
		/// <see cref="InflaterInputBuffer">Input buffer</see> for this stream.
		/// </summary>
		protected InflaterInputBuffer inputBuffer;

		/// <summary>
		/// Base stream the inflater reads from.
		/// </summary>
		private Stream baseInputStream;

		/// <summary>
		/// The compressed size
		/// </summary>
		protected long csize;

		/// <summary>
		/// Flag indicating whether this instance has been closed or not.
		/// </summary>
		private bool isClosed;

		#endregion Instance Fields
	}
}


namespace ICSharpCode.SharpZipLib.Zip.Compression.Streams
{
	/// <summary>
	/// Contains the output from the Inflation process.
	/// We need to have a window so that we can refer backwards into the output stream
	/// to repeat stuff.<br/>
	/// Author of the original java version : John Leuner
	/// </summary>
	public class OutputWindow
	{
		#region Constants

		private const int WindowSize = 1 << 15;
		private const int WindowMask = WindowSize - 1;

		#endregion Constants

		#region Instance Fields

		private byte[] window = new byte[WindowSize]; //The window is 2^15 bytes
		private int windowEnd;
		private int windowFilled;

		#endregion Instance Fields

		/// <summary>
		/// Write a byte to this output window
		/// </summary>
		/// <param name="value">value to write</param>
		/// <exception cref="InvalidOperationException">
		/// if window is full
		/// </exception>
		public void Write(int value)
		{
			if (windowFilled++ == WindowSize)
			{
				throw new InvalidOperationException("Window full");
			}
			window[windowEnd++] = (byte)value;
			windowEnd &= WindowMask;
		}

		private void SlowRepeat(int repStart, int length, int distance)
		{
			while (length-- > 0)
			{
				window[windowEnd++] = window[repStart++];
				windowEnd &= WindowMask;
				repStart &= WindowMask;
			}
		}

		/// <summary>
		/// Append a byte pattern already in the window itself
		/// </summary>
		/// <param name="length">length of pattern to copy</param>
		/// <param name="distance">distance from end of window pattern occurs</param>
		/// <exception cref="InvalidOperationException">
		/// If the repeated data overflows the window
		/// </exception>
		public void Repeat(int length, int distance)
		{
			if ((windowFilled += length) > WindowSize)
			{
				throw new InvalidOperationException("Window full");
			}

			int repStart = (windowEnd - distance) & WindowMask;
			int border = WindowSize - length;
			if ((repStart <= border) && (windowEnd < border))
			{
				if (length <= distance)
				{
					System.Array.Copy(window, repStart, window, windowEnd, length);
					windowEnd += length;
				}
				else
				{
					// We have to copy manually, since the repeat pattern overlaps.
					while (length-- > 0)
					{
						window[windowEnd++] = window[repStart++];
					}
				}
			}
			else
			{
				SlowRepeat(repStart, length, distance);
			}
		}

		/// <summary>
		/// Copy from input manipulator to internal window
		/// </summary>
		/// <param name="input">source of data</param>
		/// <param name="length">length of data to copy</param>
		/// <returns>the number of bytes copied</returns>
		public int CopyStored(StreamManipulator input, int length)
		{
			length = Math.Min(Math.Min(length, WindowSize - windowFilled), input.AvailableBytes);
			int copied;

			int tailLen = WindowSize - windowEnd;
			if (length > tailLen)
			{
				copied = input.CopyBytes(window, windowEnd, tailLen);
				if (copied == tailLen)
				{
					copied += input.CopyBytes(window, 0, length - tailLen);
				}
			}
			else
			{
				copied = input.CopyBytes(window, windowEnd, length);
			}

			windowEnd = (windowEnd + copied) & WindowMask;
			windowFilled += copied;
			return copied;
		}

		/// <summary>
		/// Copy dictionary to window
		/// </summary>
		/// <param name="dictionary">source dictionary</param>
		/// <param name="offset">offset of start in source dictionary</param>
		/// <param name="length">length of dictionary</param>
		/// <exception cref="InvalidOperationException">
		/// If window isnt empty
		/// </exception>
		public void CopyDict(byte[] dictionary, int offset, int length)
		{
			if (dictionary == null)
			{
				throw new ArgumentNullException(nameof(dictionary));
			}

			if (windowFilled > 0)
			{
				throw new InvalidOperationException();
			}

			if (length > WindowSize)
			{
				offset += length - WindowSize;
				length = WindowSize;
			}
			System.Array.Copy(dictionary, offset, window, 0, length);
			windowEnd = length & WindowMask;
		}

		/// <summary>
		/// Get remaining unfilled space in window
		/// </summary>
		/// <returns>Number of bytes left in window</returns>
		public int GetFreeSpace()
		{
			return WindowSize - windowFilled;
		}

		/// <summary>
		/// Get bytes available for output in window
		/// </summary>
		/// <returns>Number of bytes filled</returns>
		public int GetAvailable()
		{
			return windowFilled;
		}

		/// <summary>
		/// Copy contents of window to output
		/// </summary>
		/// <param name="output">buffer to copy to</param>
		/// <param name="offset">offset to start at</param>
		/// <param name="len">number of bytes to count</param>
		/// <returns>The number of bytes copied</returns>
		/// <exception cref="InvalidOperationException">
		/// If a window underflow occurs
		/// </exception>
		public int CopyOutput(byte[] output, int offset, int len)
		{
			int copyEnd = windowEnd;
			if (len > windowFilled)
			{
				len = windowFilled;
			}
			else
			{
				copyEnd = (windowEnd - windowFilled + len) & WindowMask;
			}

			int copied = len;
			int tailLen = len - copyEnd;

			if (tailLen > 0)
			{
				System.Array.Copy(window, WindowSize - tailLen, output, offset, tailLen);
				offset += tailLen;
				len = copyEnd;
			}
			System.Array.Copy(window, copyEnd - len, output, offset, len);
			windowFilled -= copied;
			if (windowFilled < 0)
			{
				throw new InvalidOperationException();
			}
			return copied;
		}

		/// <summary>
		/// Reset by clearing window so <see cref="GetAvailable">GetAvailable</see> returns 0
		/// </summary>
		public void Reset()
		{
			windowFilled = windowEnd = 0;
		}
	}
}


namespace ICSharpCode.SharpZipLib.Zip.Compression.Streams
{
	/// <summary>
	/// This class allows us to retrieve a specified number of bits from
	/// the input buffer, as well as copy big byte blocks.
	///
	/// It uses an int buffer to store up to 31 bits for direct
	/// manipulation.  This guarantees that we can get at least 16 bits,
	/// but we only need at most 15, so this is all safe.
	///
	/// There are some optimizations in this class, for example, you must
	/// never peek more than 8 bits more than needed, and you must first
	/// peek bits before you may drop them.  This is not a general purpose
	/// class but optimized for the behaviour of the Inflater.
	///
	/// authors of the original java version : John Leuner, Jochen Hoenicke
	/// </summary>
	public class StreamManipulator
	{
		/// <summary>
		/// Get the next sequence of bits but don't increase input pointer.  bitCount must be
		/// less or equal 16 and if this call succeeds, you must drop
		/// at least n - 8 bits in the next call.
		/// </summary>
		/// <param name="bitCount">The number of bits to peek.</param>
		/// <returns>
		/// the value of the bits, or -1 if not enough bits available.  */
		/// </returns>
		public int PeekBits(int bitCount)
		{
			if (bitsInBuffer_ < bitCount)
			{
				if (windowStart_ == windowEnd_)
				{
					return -1; // ok
				}
				buffer_ |= (uint)((window_[windowStart_++] & 0xff |
								 (window_[windowStart_++] & 0xff) << 8) << bitsInBuffer_);
				bitsInBuffer_ += 16;
			}
			return (int)(buffer_ & ((1 << bitCount) - 1));
		}

		/// <summary>
		/// Tries to grab the next <paramref name="bitCount"/> bits from the input and
		/// sets <paramref name="output"/> to the value, adding <paramref name="outputOffset"/>.
		/// </summary>
		/// <returns>true if enough bits could be read, otherwise false</returns>
		public bool TryGetBits(int bitCount, ref int output, int outputOffset = 0)
		{
			var bits = PeekBits(bitCount);
			if (bits < 0)
			{
				return false;
			}
			output = bits + outputOffset;
			DropBits(bitCount);
			return true;
		}

		/// <summary>
		/// Tries to grab the next <paramref name="bitCount"/> bits from the input and
		/// sets <paramref name="index"/> of <paramref name="array"/> to the value.
		/// </summary>
		/// <returns>true if enough bits could be read, otherwise false</returns>
		public bool TryGetBits(int bitCount, ref byte[] array, int index)
		{
			var bits = PeekBits(bitCount);
			if (bits < 0)
			{
				return false;
			}
			array[index] = (byte)bits;
			DropBits(bitCount);
			return true;
		}

		/// <summary>
		/// Drops the next n bits from the input.  You should have called PeekBits
		/// with a bigger or equal n before, to make sure that enough bits are in
		/// the bit buffer.
		/// </summary>
		/// <param name="bitCount">The number of bits to drop.</param>
		public void DropBits(int bitCount)
		{
			buffer_ >>= bitCount;
			bitsInBuffer_ -= bitCount;
		}

		/// <summary>
		/// Gets the next n bits and increases input pointer.  This is equivalent
		/// to <see cref="PeekBits"/> followed by <see cref="DropBits"/>, except for correct error handling.
		/// </summary>
		/// <param name="bitCount">The number of bits to retrieve.</param>
		/// <returns>
		/// the value of the bits, or -1 if not enough bits available.
		/// </returns>
		public int GetBits(int bitCount)
		{
			int bits = PeekBits(bitCount);
			if (bits >= 0)
			{
				DropBits(bitCount);
			}
			return bits;
		}

		/// <summary>
		/// Gets the number of bits available in the bit buffer.  This must be
		/// only called when a previous PeekBits() returned -1.
		/// </summary>
		/// <returns>
		/// the number of bits available.
		/// </returns>
		public int AvailableBits
		{
			get
			{
				return bitsInBuffer_;
			}
		}

		/// <summary>
		/// Gets the number of bytes available.
		/// </summary>
		/// <returns>
		/// The number of bytes available.
		/// </returns>
		public int AvailableBytes
		{
			get
			{
				return windowEnd_ - windowStart_ + (bitsInBuffer_ >> 3);
			}
		}

		/// <summary>
		/// Skips to the next byte boundary.
		/// </summary>
		public void SkipToByteBoundary()
		{
			buffer_ >>= (bitsInBuffer_ & 7);
			bitsInBuffer_ &= ~7;
		}

		/// <summary>
		/// Returns true when SetInput can be called
		/// </summary>
		public bool IsNeedingInput
		{
			get
			{
				return windowStart_ == windowEnd_;
			}
		}

		/// <summary>
		/// Copies bytes from input buffer to output buffer starting
		/// at output[offset].  You have to make sure, that the buffer is
		/// byte aligned.  If not enough bytes are available, copies fewer
		/// bytes.
		/// </summary>
		/// <param name="output">
		/// The buffer to copy bytes to.
		/// </param>
		/// <param name="offset">
		/// The offset in the buffer at which copying starts
		/// </param>
		/// <param name="length">
		/// The length to copy, 0 is allowed.
		/// </param>
		/// <returns>
		/// The number of bytes copied, 0 if no bytes were available.
		/// </returns>
		/// <exception cref="ArgumentOutOfRangeException">
		/// Length is less than zero
		/// </exception>
		/// <exception cref="InvalidOperationException">
		/// Bit buffer isnt byte aligned
		/// </exception>
		public int CopyBytes(byte[] output, int offset, int length)
		{
			if (length < 0)
			{
				throw new ArgumentOutOfRangeException(nameof(length));
			}

			if ((bitsInBuffer_ & 7) != 0)
			{
				// bits_in_buffer may only be 0 or a multiple of 8
				throw new InvalidOperationException("Bit buffer is not byte aligned!");
			}

			int count = 0;
			while ((bitsInBuffer_ > 0) && (length > 0))
			{
				output[offset++] = (byte)buffer_;
				buffer_ >>= 8;
				bitsInBuffer_ -= 8;
				length--;
				count++;
			}

			if (length == 0)
			{
				return count;
			}

			int avail = windowEnd_ - windowStart_;
			if (length > avail)
			{
				length = avail;
			}
			System.Array.Copy(window_, windowStart_, output, offset, length);
			windowStart_ += length;

			if (((windowStart_ - windowEnd_) & 1) != 0)
			{
				// We always want an even number of bytes in input, see peekBits
				buffer_ = (uint)(window_[windowStart_++] & 0xff);
				bitsInBuffer_ = 8;
			}
			return count + length;
		}

		/// <summary>
		/// Resets state and empties internal buffers
		/// </summary>
		public void Reset()
		{
			buffer_ = 0;
			windowStart_ = windowEnd_ = bitsInBuffer_ = 0;
		}

		/// <summary>
		/// Add more input for consumption.
		/// Only call when IsNeedingInput returns true
		/// </summary>
		/// <param name="buffer">data to be input</param>
		/// <param name="offset">offset of first byte of input</param>
		/// <param name="count">number of bytes of input to add.</param>
		public void SetInput(byte[] buffer, int offset, int count)
		{
			if (buffer == null)
			{
				throw new ArgumentNullException(nameof(buffer));
			}

			if (offset < 0)
			{
				throw new ArgumentOutOfRangeException(nameof(offset), "Cannot be negative");
			}

			if (count < 0)
			{
				throw new ArgumentOutOfRangeException(nameof(count), "Cannot be negative");
			}

			if (windowStart_ < windowEnd_)
			{
				throw new InvalidOperationException("Old input was not completely processed");
			}

			int end = offset + count;

			// We want to throw an ArrayIndexOutOfBoundsException early.
			// Note the check also handles integer wrap around.
			if ((offset > end) || (end > buffer.Length))
			{
				throw new ArgumentOutOfRangeException(nameof(count));
			}

			if ((count & 1) != 0)
			{
				// We always want an even number of bytes in input, see PeekBits
				buffer_ |= (uint)((buffer[offset++] & 0xff) << bitsInBuffer_);
				bitsInBuffer_ += 8;
			}

			window_ = buffer;
			windowStart_ = offset;
			windowEnd_ = end;
		}

		#region Instance Fields

		private byte[] window_;
		private int windowStart_;
		private int windowEnd_;

		private uint buffer_;
		private int bitsInBuffer_;

		#endregion Instance Fields
	}
}

